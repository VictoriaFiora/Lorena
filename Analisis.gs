// ============================================================
// ANALISIS.GS
// Detecta desvíos y duplicados en BD 202603.
// Genera: "Desvios Detectados" y "Locales Duplicados".
//
// Para correr solo análisis: ejecutar runAnalisis()
// Para correr todo junto:    ejecutar runAll() en Tablas.gs
// ============================================================

// ── CONFIGURACIÓN ───────────────────────────────────────────
// (compartida con Tablas.gs — Google Apps Script unifica el scope)

var ANALISIS_CONFIG = {

  // ── Nombres de hojas ──
  SHEET_BD:             'BD 202603',
  SHEET_LOC_REF:        'ID localidad-provincia-region',
  SHEET_RESULT_DESVIOS: 'Desvios Detectados',
  SHEET_RESULT_DUPS:    'Locales Duplicados',
  SHEET_LOCALIDADES:    'Tabla Localidades',
  SHEET_SUCURSALES:     'Tabla Sucursales Localizacion',
  SHEET_BD_CORREGIDA:   'BD 202603 - Corregida',

  // ── Columnas en BD 202603 (base 0) ──
  BD_ID:            0,   // A   ID Sucursal
  BD_GLN:           1,   // B   gln
  BD_REPOSICION:    2,   // C   reposicion
  BD_TRUCK:         3,   // D   truck
  BD_CLIENTE:       4,   // E   cliente
  BD_COD_CADENA:    5,   // F   cod_cadena
  BD_CADENA_DESC:   6,   // G   cadena_descr
  BD_TIPO_CAD_COD:  7,   // H   tipo_cadena_cod
  BD_TIPO_CAD_DESC: 8,   // I   tipo_cadena_descr
  BD_ALC_CAD_COD:   9,   // J   alcance_cadena_cod
  BD_ALC_CAD:       10,  // K   alcance_cadena
  BD_TIPO:          11,  // L   TIPO
  BD_FORMATO:       12,  // M   formato PS
  BD_REG_PAIS:      13,  // N   region_pais_cod
  BD_AMBA_INT:      14,  // O   amba/interior
  BD_SUC_NUM:       15,  // P   sucursal_numero
  BD_SUC_DESC:      16,  // Q   sucursal_descr
  //                17,  // R   (oculta)
  //                18,  // S   (oculta)
  BD_DIRECCION:     19,  // T   Direccion
  BD_REGION_COD:    20,  // U   region_cod_ccu
  BD_REGION_DESC:   21,  // V   region_descrip_ccu
  BD_ZONA_COD:      22,  // W   zona_cod
  BD_ZONA_DESC:     23,  // X   zona_descr
  BD_COD_PROV:      24,  // Y   cod_provincia
  BD_PROVINCIA:     25,  // Z   provincia
  BD_COD_LOC:       26,  // AA  cod_localidad
  BD_LOCALIDAD:     27,  // AB  localidad
  BD_DPTO:          28,  // AC  departamento_ccu
  BD_CP:            29,  // AD  sucursal_codigo_postal
  BD_LAT:           30,  // AE  latitud
  BD_LNG:           31,  // AF  longitud

  // ── Columnas en ID localidad-provincia-region (base 0) ──
  LOC_ID:           0,   // A
  LOC_REGION_COD:   3,   // D
  LOC_REGION_DESC:  4,   // E
  LOC_ZONA_COD:     5,   // F
  LOC_ZONA_DESC:    6,   // G
  LOC_COD_PROV:     7,   // H
  LOC_PROVINCIA:    8,   // I
  LOC_COD_LOC:      9,   // J
  LOC_LOCALIDAD:    10,  // K
};

// ── UTILIDADES (compartidas con Tablas.gs) ──────────────────

function getOrCreateSheet(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function norm(val) {
  return String(val).trim().toLowerCase();
}

function fill2d(rows, cols, color) {
  var result = [];
  for (var i = 0; i < rows; i++) result.push(fill1d(cols, color));
  return result;
}

function fill1d(cols, color) {
  var row = [];
  for (var c = 0; c < cols; c++) row.push(color);
  return row;
}

// ── ENTRADA ─────────────────────────────────────────────────

function runAnalisis() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log('=== INICIANDO ANÁLISIS ===');
  detectarDesvios(ss);
  detectarDuplicadosEnBD(ss);
  Logger.log('=== ANÁLISIS FINALIZADO ===');
}

// ════════════════════════════════════════════════════════════
// 1. DETECTAR DESVÍOS
//    Compara los campos de ubicación de BD 202603 contra la
//    tabla de referencia "ID localidad-provincia-region".
//    Genera: "Desvios Detectados"
// ════════════════════════════════════════════════════════════

function detectarDesvios(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  var bdSheet  = ss.getSheetByName(ANALISIS_CONFIG.SHEET_BD);
  var locSheet = ss.getSheetByName(ANALISIS_CONFIG.SHEET_LOC_REF);
  if (!bdSheet || !locSheet) { Logger.log('ERROR: Hojas no encontradas.'); return; }

  var bdData  = bdSheet.getDataRange().getValues();
  var locData = locSheet.getDataRange().getValues();

  var locMap = {};
  for (var i = 1; i < locData.length; i++) {
    var id = norm(locData[i][ANALISIS_CONFIG.LOC_ID]);
    if (id) locMap[id] = locData[i];
  }

  var campos = [
    { label: 'region_cod_ccu',    bdCol: ANALISIS_CONFIG.BD_REGION_COD,  locCol: ANALISIS_CONFIG.LOC_REGION_COD },
    { label: 'region_descrip_ccu',bdCol: ANALISIS_CONFIG.BD_REGION_DESC, locCol: ANALISIS_CONFIG.LOC_REGION_DESC },
    { label: 'zona_cod',          bdCol: ANALISIS_CONFIG.BD_ZONA_COD,    locCol: ANALISIS_CONFIG.LOC_ZONA_COD },
    { label: 'zona_descr',        bdCol: ANALISIS_CONFIG.BD_ZONA_DESC,   locCol: ANALISIS_CONFIG.LOC_ZONA_DESC },
    { label: 'cod_provinca',      bdCol: ANALISIS_CONFIG.BD_COD_PROV,    locCol: ANALISIS_CONFIG.LOC_COD_PROV },
    { label: 'provincia',         bdCol: ANALISIS_CONFIG.BD_PROVINCIA,   locCol: ANALISIS_CONFIG.LOC_PROVINCIA },
    { label: 'cod_localidad',     bdCol: ANALISIS_CONFIG.BD_COD_LOC,     locCol: ANALISIS_CONFIG.LOC_COD_LOC },
    { label: 'localidad',         bdCol: ANALISIS_CONFIG.BD_LOCALIDAD,   locCol: ANALISIS_CONFIG.LOC_LOCALIDAD },
  ];

  var headers  = ['ID Sucursal','Campo','Valor en BD 202603','Valor Correcto','Fila en BD'];
  var dataRows = [];

  for (var r = 1; r < bdData.length; r++) {
    var row   = bdData[r];
    var idSuc = norm(row[ANALISIS_CONFIG.BD_ID]);
    if (!idSuc) continue;
    var refRow = locMap[idSuc];

    for (var c = 0; c < campos.length; c++) {
      var f      = campos[c];
      var valBD  = String(row[f.bdCol]).trim();
      var valRef = refRow ? String(refRow[f.locCol]).trim() : 'ID NO ENCONTRADO';
      if (valBD !== valRef) {
        dataRows.push([row[ANALISIS_CONFIG.BD_ID], f.label, valBD, valRef, r + 1]);
      }
    }
  }

  Logger.log('Desvios: ' + dataRows.length);
  var allRows = [headers].concat(dataRows);
  var sheet   = getOrCreateSheet(ss, ANALISIS_CONFIG.SHEET_RESULT_DESVIOS);
  sheet.clearContents(); sheet.clearFormats();
  sheet.getRange(1, 1, allRows.length, headers.length).setValues(allRows);
  sheet.getRange(1, 1, 1, headers.length).setBackground('#c62828').setFontColor('#fff').setFontWeight('bold');
  if (dataRows.length > 0) {
    sheet.getRange(2, 1, dataRows.length, headers.length)
      .setBackgrounds(fill2d(dataRows.length, headers.length, '#fff3e0'));
  }
  sheet.setFrozenRows(1); sheet.autoResizeColumns(1, headers.length);
  sheet.getRange(allRows.length + 2, 1).setValue('TOTAL DESVIOS: ' + dataRows.length)
    .setFontWeight('bold').setFontColor('#c62828');
}

// ════════════════════════════════════════════════════════════
// 2. DETECTAR DUPLICADOS
//    Detecta IDs, GLNs y Trucks repetidos en BD 202603.
//    Genera: "Locales Duplicados"
//    Colores: rosa=ID | amarillo claro=GLN | mostaza=TRUCK
// ════════════════════════════════════════════════════════════

function detectarDuplicadosEnBD(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  var bdSheet = ss.getSheetByName(ANALISIS_CONFIG.SHEET_BD);
  if (!bdSheet) return;
  var data = bdSheet.getDataRange().getValues();

  var idCount = {}, idRows = {};
  for (var r = 1; r < data.length; r++) {
    var id = norm(data[r][ANALISIS_CONFIG.BD_ID]);
    if (!id) continue;
    idCount[id] = (idCount[id] || 0) + 1;
    if (!idRows[id]) idRows[id] = [];
    idRows[id].push(r + 1);
  }

  var headers  = ['Tipo','ID Sucursal','GLN','zona_cod','cod_provinca','cod_localidad','localidad','Filas'];
  var dataRows = [];

  for (var id in idCount) {
    if (idCount[id] > 1) {
      for (var r = 1; r < data.length; r++) {
        if (norm(data[r][ANALISIS_CONFIG.BD_ID]) === id) {
          dataRows.push([
            'ID REPETIDO', data[r][ANALISIS_CONFIG.BD_ID],
            data[r][ANALISIS_CONFIG.BD_GLN], data[r][ANALISIS_CONFIG.BD_ZONA_COD],
            data[r][ANALISIS_CONFIG.BD_COD_PROV], data[r][ANALISIS_CONFIG.BD_COD_LOC],
            data[r][ANALISIS_CONFIG.BD_LOCALIDAD], idRows[id].join(', ')
          ]);
        }
      }
    }
  }

  // ── GLN duplicados ──
  var glnCount = {}, glnRows = {};
  for (var r = 1; r < data.length; r++) {
    var gln = String(data[r][ANALISIS_CONFIG.BD_GLN]).trim();
    if (!gln) continue;
    glnCount[gln] = (glnCount[gln] || 0) + 1;
    if (!glnRows[gln]) glnRows[gln] = [];
    glnRows[gln].push(r + 1);
  }
  for (var gln in glnCount) {
    if (glnCount[gln] > 1) {
      dataRows.push(['GLN REPETIDO','—', gln,'—','—','—','—', glnRows[gln].join(', ')]);
    }
  }

  // ── Truck duplicados ──
  var truckCount = {}, truckRows = {};
  for (var r = 1; r < data.length; r++) {
    var truck = String(data[r][ANALISIS_CONFIG.BD_TRUCK]).trim();
    if (!truck || truck === '0' || truck.toLowerCase() === 'false') continue;
    truckCount[truck] = (truckCount[truck] || 0) + 1;
    if (!truckRows[truck]) truckRows[truck] = [];
    truckRows[truck].push(r + 1);
  }
  for (var truck in truckCount) {
    if (truckCount[truck] > 1) {
      dataRows.push(['TRUCK REPETIDO','—','—','—','—','—','—', truckRows[truck].join(', ')]);
    }
  }

  Logger.log('Duplicados: ' + dataRows.length);
  var sheet = getOrCreateSheet(ss, ANALISIS_CONFIG.SHEET_RESULT_DUPS);
  sheet.clearContents(); sheet.clearFormats();
  var allRows = [headers].concat(dataRows);
  if (dataRows.length > 0) {
    sheet.getRange(1, 1, allRows.length, headers.length).setValues(allRows);
    var bgArr = dataRows.map(function(row) {
      return fill1d(headers.length,
        row[0] === 'ID REPETIDO'    ? '#fce4ec' :
        row[0] === 'GLN REPETIDO'   ? '#fff8e1' :
        /* TRUCK REPETIDO */          '#fff3cd');
    });
    sheet.getRange(2, 1, dataRows.length, headers.length).setBackgrounds(bgArr);
  } else {
    sheet.getRange(1,1).setValue('No se encontraron duplicados.');
  }
  sheet.getRange(1,1,1,headers.length).setBackground('#e65100').setFontColor('#fff').setFontWeight('bold');
  sheet.setFrozenRows(1); sheet.autoResizeColumns(1, headers.length);
}
