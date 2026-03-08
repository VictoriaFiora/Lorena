
from __future__ import annotations

import argparse
import math
import re
import shutil
import time
import unicodedata
from dataclasses import dataclass
from datetime import datetime
from typing import Any, Dict, Optional, Tuple
from pathlib import Path
import openpyxl
from openpyxl.styles import PatternFill, Font

try:
    import requests  # optional, only used with --web
except Exception:
    requests = None

FILENAME = "BaseDatos.xlsx"
SHEET_MASTER = "Master Localizacion - Corregida"
SHEET_BD = "BD 202603"
SHEET_LOC = "Maestro Localidades - Corregido"
SHEET_REF = "ID localidad-provincia-region"
SHEET_LEGEND = "Leyenda_Colores"
SHEET_AUDIT = "Auditoria_Geografia"

# Color coding
FILL_FROM_BD = PatternFill(start_color="D9EAF7", end_color="D9EAF7", fill_type="solid")
FILL_FROM_LOC = PatternFill(start_color="E2F0D9", end_color="E2F0D9", fill_type="solid")
FILL_NORMALIZED = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
FILL_FROM_WEB = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
FILL_CONFLICT_MASTER = PatternFill(start_color="F4CCCC", end_color="F4CCCC", fill_type="solid")
FILL_CONFLICT_REF = PatternFill(start_color="E4D7FF", end_color="E4D7FF", fill_type="solid")
FILL_HEADER = PatternFill(start_color="D9EAF7", end_color="D9EAF7", fill_type="solid")

NOMINATIM_URL = "https://nominatim.openstreetmap.org/search"
USER_AGENT = "MasterLocalizacionCleaner/2.0"

MASTER_FIELDS = [
    "GLN",
    "region_pais_cod",
    "amba/interior",
    "region_cod",
    "region_descr",
    "zona_cod",
    "zona_descr",
    "cod_prov",
    "provincia",
    "cod_localidad",
    "localidad",
    "departamento",
    "cp",
    "direccion",
    "lat",
    "long",
]

BD_TO_MASTER = {
    "gln": "GLN",
    "region_pais_cod": "region_pais_cod",
    "amba/interior": "amba/interior",
    "region_cod_ccu": "region_cod",
    "region_descip_ccu": "region_descr",
    "zona_cod": "zona_cod",
    "zona_descr": "zona_descr",
    "cod_provinica": "cod_prov",
    "provincia": "provincia",
    "cod_localidad": "cod_localidad",
    "localidad": "localidad",
    "departamento_ccu": "departamento",
    "sucursal_codigo_postal": "cp",
    "Direccion": "direccion",
    "latitud": "lat",
    "longitud": "long",
}

LOC_TO_MASTER = {
    "cod_localidad": "cod_localidad",
    "localidad": "localidad",
    "departamento": "departamento",
    "cod_provincia": "cod_prov",
    "provincia": "provincia",
    "cp": "cp",
    "region_pais_cod": "region_pais_cod",
    "amba/interior": "amba/interior",
    "region_cod": "region_cod",
    "region_descr": "region_descr",
    "zona_cod": "zona_cod",
    "zona_descr": "zona_descr",
}

REF_TO_MASTER = {
    "region_pais_cod": "region_pais_cod",
    "amba/interior": "amba/interior",
    "region_cod": "region_cod",
    "region_descip": "region_descr",
    "zona_cod": "zona_cod",
    "zona_descr": "zona_descr",
    "cod_provinica": "cod_prov",
    "provincia": "provincia",
    "cod_localidad": "cod_localidad",
    "localidad": "localidad",
}

CODE_FIELDS = {"GLN", "region_pais_cod", "region_cod", "zona_cod", "cod_prov", "cod_localidad", "cp"}
TEXT_UPPER_FIELDS = {"amba/interior"}
TITLE_FIELDS = {"provincia", "localidad", "departamento", "region_descr", "zona_descr"}
LATLON_FIELDS = {"lat", "long"}


@dataclass
class Stats:
    from_bd: int = 0
    from_loc: int = 0
    normalized: int = 0
    from_web: int = 0
    rows_touched: int = 0
    conflict_rows: int = 0
    conflicts: int = 0


def strip_accents(text: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFKD", text) if not unicodedata.combining(c))


def is_blank(value: Any) -> bool:
    if value is None:
        return True
    s = str(value).strip()
    return s == "" or s.lower() in {"none", "nan", "null", "faltante", "s/d", "sin dato"}


def clean_spaces(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def normalize_key(text: Any) -> str:
    if is_blank(text):
        return ""
    s = clean_spaces(str(text))
    s = strip_accents(s).upper()
    return s


def title_case(text: str) -> str:
    small = {"de", "del", "la", "las", "los", "y"}
    words = clean_spaces(text).lower().split(" ")
    out = []
    for i, w in enumerate(words):
        if i > 0 and w in small:
            out.append(w)
        elif w in {"gba", "caba", "noa", "nea"}:
            out.append(w.upper())
        elif w == "cap":
            out.append("CAP")
        else:
            out.append(w.capitalize())
    return " ".join(out)


def parse_int_like(value: Any) -> Optional[str]:
    if is_blank(value):
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float):
        if math.isnan(value):
            return None
        if value.is_integer():
            return str(int(value))
        return str(value).replace(",", ".")
    s = clean_spaces(str(value))
    s = s.replace(".0", "") if re.fullmatch(r"\d+\.0", s) else s
    s = s.replace(",", "") if re.fullmatch(r"\d{1,3}(,\d{3})+", s) else s
    s = s.replace(" ", "")
    if re.fullmatch(r"\d+", s):
        return s
    if re.fullmatch(r"\d+\.0+", s):
        return s.split(".")[0]
    return s


def normalize_id_sucursal(value: Any) -> Optional[str]:
    if is_blank(value):
        return None
    s = clean_spaces(str(value)).lower().replace(" ", "")
    raw = s[4:] if s.startswith("suc_") else s
    raw = raw.replace(",", "")
    if re.fullmatch(r"\d+\.0+", raw):
        raw = raw.split(".")[0]
    if re.fullmatch(r"\d+", raw):
        return f"suc_{int(raw)}"
    return f"suc_{raw}"


def normalize_latlon(value: Any) -> Optional[float]:
    if is_blank(value):
        return None
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if isinstance(value, float) and math.isnan(value):
            return None
        return float(value)
    s = clean_spaces(str(value))
    if re.fullmatch(r"-?\d+,\d+", s):
        s = s.replace(",", ".")
    elif re.fullmatch(r"-?\d{1,3}(,\d{3})+(\.\d+)?", s):
        s = s.replace(",", "")
    try:
        return float(s)
    except Exception:
        return None


def normalize_field_value(field: str, value: Any) -> Any:
    if is_blank(value):
        return None
    if field == "ID Sucursal":
        return normalize_id_sucursal(value)
    if field in CODE_FIELDS:
        return parse_int_like(value)
    if field in LATLON_FIELDS:
        return normalize_latlon(value)
    s = clean_spaces(str(value))
    if field in TEXT_UPPER_FIELDS:
        return strip_accents(s).upper()
    if field in TITLE_FIELDS:
        return title_case(strip_accents(s))
    return s


def comparable_value(field: str, value: Any) -> Optional[str]:
    v = normalize_field_value(field, value)
    if v is None:
        return None
    if isinstance(v, float):
        return f"{v:.8f}"
    return str(v)


def find_headers(ws) -> Dict[str, int]:
    return {
        str(ws.cell(1, col).value): col
        for col in range(1, ws.max_column + 1)
        if ws.cell(1, col).value is not None
    }


def build_bd_index(ws_bd, bd_cols: Dict[str, int]) -> Dict[str, Dict[str, Any]]:
    index: Dict[str, Dict[str, Any]] = {}
    for row in ws_bd.iter_rows(min_row=2, values_only=True):
        row_map = {h: row[idx - 1] for h, idx in bd_cols.items()}
        key = normalize_id_sucursal(row_map.get("ID Sucursal"))
        if not key:
            continue
        cleaned = {}
        for src, dst in BD_TO_MASTER.items():
            cleaned[dst] = normalize_field_value(dst, row_map.get(src))
        score = sum(v is not None for v in cleaned.values())
        prev = index.get(key)
        prev_score = prev.get("__score__", -1) if prev else -1
        if score > prev_score:
            cleaned["__score__"] = score
            index[key] = cleaned
    return index


def build_loc_indexes(ws_loc, loc_cols: Dict[str, int]) -> Tuple[Dict[str, Dict[str, Any]], Dict[Tuple[str, str], Dict[str, Any]]]:
    by_cod: Dict[str, Dict[str, Any]] = {}
    by_prov_loc: Dict[Tuple[str, str], Dict[str, Any]] = {}
    for row in ws_loc.iter_rows(min_row=2, values_only=True):
        row_map = {h: row[idx - 1] for h, idx in loc_cols.items() if h is not None}
        cleaned = {}
        for src, dst in LOC_TO_MASTER.items():
            if src in row_map:
                cleaned[dst] = normalize_field_value(dst, row_map.get(src))
        cod_loc = cleaned.get("cod_localidad")
        if cod_loc:
            by_cod[cod_loc] = cleaned
        prov = normalize_key(cleaned.get("provincia"))
        loc = normalize_key(cleaned.get("localidad"))
        if prov and loc:
            by_prov_loc[(prov, loc)] = cleaned
    return by_cod, by_prov_loc


def build_ref_index(ws_ref, ref_cols: Dict[str, int]) -> Dict[str, Dict[str, Any]]:
    index: Dict[str, Dict[str, Any]] = {}
    for row in ws_ref.iter_rows(min_row=2, values_only=True):
        row_map = {h: row[idx - 1] for h, idx in ref_cols.items()}
        key = normalize_id_sucursal(row_map.get("Id local"))
        if not key:
            continue
        cleaned = {}
        for src, dst in REF_TO_MASTER.items():
            if src in row_map:
                cleaned[dst] = normalize_field_value(dst, row_map.get(src))
        index[key] = cleaned
    return index


def set_if_missing(cell, value, fill) -> bool:
    if value is None:
        return False
    if is_blank(cell.value):
        cell.value = value
        cell.fill = fill
        return True
    return False


def normalize_row(ws, row_num: int, cols: Dict[str, int], stats: Stats) -> bool:
    touched = False
    for field in ["ID Sucursal"] + MASTER_FIELDS:
        if field not in cols:
            continue
        cell = ws.cell(row=row_num, column=cols[field])
        new_val = normalize_field_value(field, cell.value)
        old = cell.value
        old_cmp = comparable_value(field, old)
        new_cmp = comparable_value(field, new_val)
        if old_cmp != new_cmp:
            cell.value = new_val
            if old_cmp is not None or new_cmp is not None:
                cell.fill = FILL_NORMALIZED
                stats.normalized += 1
                touched = True
    return touched


def enrich_from_web(address: str, locality: str, province: str) -> Tuple[Optional[str], Optional[str]]:
    if requests is None:
        return None, None
    params = {
        "q": f"{address}, {locality}, {province}, Argentina",
        "format": "json",
        "addressdetails": 1,
        "limit": 1,
    }
    headers = {"User-Agent": USER_AGENT}
    try:
        r = requests.get(NOMINATIM_URL, params=params, headers=headers, timeout=15)
        if r.status_code != 200:
            return None, None
        data = r.json()
        if not data:
            return None, None
        addr = data[0].get("address", {})
        cp = addr.get("postcode")
        dept = addr.get("county") or addr.get("state_district") or addr.get("city_district")
        cp = parse_int_like(cp) if cp else None
        dept = title_case(strip_accents(dept)) if dept else None
        return cp, dept
    except Exception:
        return None, None


def ensure_conflict_columns(ws_master, master_cols: Dict[str, int]) -> Dict[str, int]:
    next_col = ws_master.max_column + 1
    out = {}
    for field in REF_TO_MASTER.values():
        col_name = f"ref_{field}"
        if col_name not in master_cols:
            ws_master.cell(1, next_col).value = col_name
            ws_master.cell(1, next_col).fill = FILL_HEADER
            ws_master.cell(1, next_col).font = Font(bold=True)
            master_cols[col_name] = next_col
            next_col += 1
        out[f"ref_{field}"] = master_cols[col_name]

    summary_name = "diferencias_id_localidad_provincia_region"
    if summary_name not in master_cols:
        ws_master.cell(1, next_col).value = summary_name
        ws_master.cell(1, next_col).fill = FILL_HEADER
        ws_master.cell(1, next_col).font = Font(bold=True)
        master_cols[summary_name] = next_col
        next_col += 1
    out["summary"] = master_cols[summary_name]
    return out


def clear_conflict_cells(ws_master, row_num: int, aux_cols: Dict[str, int]) -> None:
    for key, col in aux_cols.items():
        if key == "summary":
            continue
        cell = ws_master.cell(row_num, col)
        cell.value = None
        cell.fill = PatternFill(fill_type=None)
    summary = ws_master.cell(row_num, aux_cols["summary"])
    summary.value = None
    summary.fill = PatternFill(fill_type=None)


def apply_conflicts(ws_master, row_num: int, master_cols: Dict[str, int], aux_cols: Dict[str, int], ref_row: Dict[str, Any], stats: Stats, audit_counter: Dict[str, int]) -> bool:
    row_conflicts = []
    for master_field in REF_TO_MASTER.values():
        if master_field not in master_cols:
            continue
        master_cell = ws_master.cell(row_num, master_cols[master_field])
        ref_value = ref_row.get(master_field)
        master_cmp = comparable_value(master_field, master_cell.value)
        ref_cmp = comparable_value(master_field, ref_value)

        if master_cmp is not None and ref_cmp is not None and master_cmp != ref_cmp:
            master_cell.fill = FILL_CONFLICT_MASTER
            ref_cell = ws_master.cell(row_num, aux_cols[f"ref_{master_field}"])
            ref_cell.value = ref_value
            ref_cell.fill = FILL_CONFLICT_REF
            row_conflicts.append(f"{master_field}: MASTER={master_cell.value} | REF={ref_value}")
            audit_counter[master_field] = audit_counter.get(master_field, 0) + 1
            stats.conflicts += 1

    if row_conflicts:
        summary_cell = ws_master.cell(row_num, aux_cols["summary"])
        summary_cell.value = " ; ".join(row_conflicts)
        summary_cell.fill = FILL_CONFLICT_REF
        stats.conflict_rows += 1
        return True
    return False


def recreate_sheet(wb, name: str):
    if name in wb.sheetnames:
        idx = wb.sheetnames.index(name)
        ws_old = wb[name]
        wb.remove(ws_old)
        ws_new = wb.create_sheet(name, idx)
    else:
        ws_new = wb.create_sheet(name)
    return ws_new


def crear_leyenda_colores(wb):
    ws = recreate_sheet(wb, SHEET_LEGEND)

    headers = ["Color", "Significado", "Fuente", "Qué indica"]
    ws.append(headers)

    for col in range(1, len(headers) + 1):
        ws.cell(1, col).fill = FILL_HEADER
        ws.cell(1, col).font = Font(bold=True)

    rows = [
        (FILL_FROM_BD, "Dato copiado desde BD", "BD 202603",
         "El script completó el dato desde la base de sucursales"),

        (FILL_FROM_LOC, "Dato copiado desde Maestro Localidades", "Maestro Localidades - Corregido",
         "El valor se completó usando cod_localidad o provincia + localidad"),

        (FILL_FROM_WEB, "Dato enriquecido por web", "OpenStreetMap / Nominatim",
         "El script encontró el dato online"),

        (FILL_NORMALIZED, "Dato normalizado", "Pipeline",
         "El valor fue corregido de formato"),

        (FILL_CONFLICT_MASTER, "Conflicto en valor de Master",
         "Cruce con ID localidad-provincia-region",
         "El valor actual de Master difiere de la tabla de referencia"),

        (FILL_CONFLICT_REF, "Valor alternativo de referencia",
         "ID localidad-provincia-region",
         "Se muestra el valor de la tabla de referencia en columna ref_* o resumen"),

        (PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid"),
         "Sin cambios", "-", "El valor ya estaba correcto"),
    ]

    for idx, (fill, significado, fuente, indica) in enumerate(rows, start=2):
        ws.cell(idx, 1).fill = fill
        ws.cell(idx, 2).value = significado
        ws.cell(idx, 3).value = fuente
        ws.cell(idx, 4).value = indica

    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 36
    ws.column_dimensions["C"].width = 34
    ws.column_dimensions["D"].width = 70

def crear_auditoria_geografia(wb, audit_counter: Dict[str, int], stats: Stats):
    ws = recreate_sheet(wb, SHEET_AUDIT)
    headers = ["tipo_diferencia", "cantidad"]
    ws.append(headers)
    for col in range(1, len(headers) + 1):
        ws.cell(1, col).fill = FILL_HEADER
        ws.cell(1, col).font = Font(bold=True)

    ordered_fields = [
        "region_pais_cod", "amba/interior", "region_cod", "region_descr",
        "zona_cod", "zona_descr", "cod_prov", "provincia", "cod_localidad", "localidad"
    ]
    for field in ordered_fields:
        ws.append([field, audit_counter.get(field, 0)])

    ws.append([])
    ws.append(["filas_con_conflicto", stats.conflict_rows])
    ws.append(["conflictos_totales", stats.conflicts])

    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 14


def make_backup(path: str) -> str:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    original = Path(path)

    # carpeta BackUps
    backup_dir = original.parent / "BackUps"
    backup_dir.mkdir(exist_ok=True)

    # nombre del archivo
    backup_name = original.stem + f"_backup_{timestamp}" + original.suffix
    backup_path = backup_dir / backup_name

    shutil.copy2(original, backup_path)

    return str(backup_path)

def process_workbook(filename: str, use_web: bool = False, web_limit: int = 20, save_as: Optional[str] = None) -> str:
    wb = openpyxl.load_workbook(filename)
    for required in [SHEET_MASTER, SHEET_BD, SHEET_LOC, SHEET_REF]:
        if required not in wb.sheetnames:
            raise ValueError(f"No se encontró la hoja requerida: {required}")

    ws_master = wb[SHEET_MASTER]
    ws_bd = wb[SHEET_BD]
    ws_loc = wb[SHEET_LOC]
    ws_ref = wb[SHEET_REF]

    master_cols = find_headers(ws_master)
    bd_cols = find_headers(ws_bd)
    loc_cols = find_headers(ws_loc)
    ref_cols = find_headers(ws_ref)

    required_master = ["ID Sucursal"] + MASTER_FIELDS
    missing = [c for c in required_master if c not in master_cols]
    if missing:
        raise ValueError(f"Faltan columnas en {SHEET_MASTER}: {missing}")
    if "Id local" not in ref_cols:
        raise ValueError(f"Falta la columna 'Id local' en {SHEET_REF}")

    bd_index = build_bd_index(ws_bd, bd_cols)
    loc_by_cod, loc_by_prov_loc = build_loc_indexes(ws_loc, loc_cols)
    ref_index = build_ref_index(ws_ref, ref_cols)
    aux_cols = ensure_conflict_columns(ws_master, master_cols)
    master_cols = find_headers(ws_master)

    stats = Stats()
    web_calls = 0
    audit_counter: Dict[str, int] = {}

    for row_num in range(2, ws_master.max_row + 1):
        row_touched = normalize_row(ws_master, row_num, master_cols, stats)
        key = normalize_id_sucursal(ws_master.cell(row_num, master_cols["ID Sucursal"]).value)

        bd_row = bd_index.get(key)
        if bd_row:
            for field in MASTER_FIELDS:
                if field in master_cols and set_if_missing(ws_master.cell(row_num, master_cols[field]), bd_row.get(field), FILL_FROM_BD):
                    stats.from_bd += 1
                    row_touched = True

        cod_loc = normalize_field_value("cod_localidad", ws_master.cell(row_num, master_cols["cod_localidad"]).value)
        prov = normalize_key(ws_master.cell(row_num, master_cols["provincia"]).value)
        loc = normalize_key(ws_master.cell(row_num, master_cols["localidad"]).value)
        loc_row = loc_by_cod.get(cod_loc) if cod_loc else None
        if not loc_row and prov and loc:
            loc_row = loc_by_prov_loc.get((prov, loc))
        if loc_row:
            for field in [
                "region_pais_cod", "amba/interior", "region_cod", "region_descr",
                "zona_cod", "zona_descr", "cod_prov", "provincia", "cod_localidad",
                "localidad", "departamento", "cp"
            ]:
                if field in master_cols and set_if_missing(ws_master.cell(row_num, master_cols[field]), loc_row.get(field), FILL_FROM_LOC):
                    stats.from_loc += 1
                    row_touched = True

        if use_web and web_calls < web_limit:
            cp_cell = ws_master.cell(row_num, master_cols["cp"])
            dept_cell = ws_master.cell(row_num, master_cols["departamento"])
            if is_blank(cp_cell.value) or is_blank(dept_cell.value):
                addr = ws_master.cell(row_num, master_cols["direccion"]).value
                locality = ws_master.cell(row_num, master_cols["localidad"]).value
                province = ws_master.cell(row_num, master_cols["provincia"]).value
                if not is_blank(addr) and not is_blank(locality):
                    cp, dept = enrich_from_web(str(addr), str(locality), str(province or ""))
                    if is_blank(cp_cell.value) and cp is not None:
                        cp_cell.value = cp
                        cp_cell.fill = FILL_FROM_WEB
                        stats.from_web += 1
                        row_touched = True
                    if is_blank(dept_cell.value) and dept is not None:
                        dept_cell.value = dept
                        dept_cell.fill = FILL_FROM_WEB
                        stats.from_web += 1
                        row_touched = True
                    web_calls += 1
                    time.sleep(1.1)

        if normalize_row(ws_master, row_num, master_cols, stats):
            row_touched = True

        clear_conflict_cells(ws_master, row_num, aux_cols)
        ref_row = ref_index.get(key) if key else None
        if ref_row and apply_conflicts(ws_master, row_num, master_cols, aux_cols, ref_row, stats, audit_counter):
            row_touched = True

        if row_touched:
            stats.rows_touched += 1

    crear_leyenda_colores(wb)
    crear_auditoria_geografia(wb, audit_counter, stats)

    output = save_as or filename
    wb.save(output)

    print("\n=== RESUMEN ===")
    print(f"Archivo guardado en: {output}")
    print(f"Celdas completadas desde BD 202603: {stats.from_bd}")
    print(f"Celdas completadas desde Maestro Localidades: {stats.from_loc}")
    print(f"Celdas completadas desde web: {stats.from_web}")
    print(f"Celdas normalizadas: {stats.normalized}")
    print(f"Filas tocadas: {stats.rows_touched}")
    print(f"Filas con conflicto: {stats.conflict_rows}")
    print(f"Conflictos totales: {stats.conflicts}")
    print(f"Hojas generadas/actualizadas: {SHEET_LEGEND}, {SHEET_AUDIT}")
    return output


def main() -> None:
    parser = argparse.ArgumentParser(description="Pipeline completo de Master Localizacion con conflictos y auditoría")
    parser.add_argument("--file", default=FILENAME, help="Ruta al archivo Excel")
    parser.add_argument("--save-as", default=None, help="Guardar resultado en otro archivo")
    parser.add_argument("--backup", action="store_true", help="Genera una copia de seguridad antes de guardar")
    parser.add_argument("--web", action="store_true", help="Hace enriquecimiento web para cp/departamento")
    parser.add_argument("--web-limit", type=int, default=20, help="Máximo de búsquedas web por corrida")
    args = parser.parse_args()

    if args.backup:
        backup = make_backup(args.file)
        print(f"Backup creado: {backup}")

    process_workbook(args.file, use_web=args.web, web_limit=args.web_limit, save_as=args.save_as)


if __name__ == "__main__":
    main()
