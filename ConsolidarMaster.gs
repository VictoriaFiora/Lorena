/**
 * FUNCIÓN MAESTRA: Corre todo de una vez.
 * Genera Master Localizacion, Maestro Localidades y actualiza la BD Corregida.
 */
function runMasterProcesoLocalizacion() {
  Logger.log(">>> INICIANDO PROCESO DE CONSOLIDACIÓN INTERNA <<<");
  
  // 1. Consolidar datos de todas las fuentes en el Master
  generarMasterLocalizacion();
  Logger.log("1. Master Localizacion OK.");
  
  // 2. Crear el Maestro deduplicado por cod_localidad
  generarMaestroLocalidades();
  Logger.log("2. Maestro Localidades OK.");
  
  // 3. Completar jerarquías faltantes con lógica interna
  completarMaestroLocalidades();
  Logger.log("3. Completado de jerarquías OK.");
  
  Logger.log(">>> PROCESO DE TABLAS FINALIZADO <<<");
}

/**
 * Completa los datos de jerarquía (Región, Zona, AMBA) en Maestro Localidades
 * basándose únicamente en la información provincial existente y patrones internos.
 */
function completarMaestroLocalidades() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Maestro Localidades");
  if (!sheet) return;
  
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return;
  
  var headers = data[0];
  var modificado = false;
  var counts = { infer_map: 0, infer_consensus: 0, web: 0, doubt: 0 };
  var REVISION_SHEET = "REVISION_GEOREF";
  var doubtLog = []; 
  var MAX_WEB_CALLS = 100; // Límite de seguridad para evitar timeout de 6 min

  // 1. Mapa estático de metadatos por Provincia (Basado en lógica de negocio CCU)
  var PROV_METADATA = {
    "02": { rp:"1", ai:"AMBA", rc:"11", rd:"REGION I", zc:"1", zd:"CAP FED" },      // CABA
    "06": { rp:"1", ai:"AMBA", rc:"11", rd:"REGION I" },                            // Buenos Aires (AMBA default)
    "14": { rp:"2", ai:"INTERIOR", rc:"12", rd:"REGION II", zc:"12", zd:"PBA" },    // Cordoba -> PBA?
    "82": { rp:"2", ai:"INTERIOR", rc:"3", rd:"Rosario", zc:"11", zd:"Rosario" },   // Santa Fe
    "50": { rp:"2", ai:"INTERIOR", rc:"7", rd:"PBA", zc:"12", zd:"PBA" },           // Mendoza -> PBA?
    "66": { rp:"2", ai:"INTERIOR", rc:"9", rd:"NOA Sur", zc:"12", zd:"Sur" },       // Salta
    "86": { rp:"2", ai:"INTERIOR", rc:"9", rd:"NOA Sur", zc:"12", zd:"Sur" }        // Tucuman
  };

  // 2. Construir mapa de consenso (Lo que ya existe en la hoja)
  var consensusMap = {}; // prov_id -> { field_idx -> { value -> count } }
  for (var i = 1; i < data.length; i++) {
    var pId = String(data[i][3]).trim(); // cod_provincia (D)
    if (!pId) continue;
    if (!consensusMap[pId]) consensusMap[pId] = {};
    
    // Columnas a inferir: 6:rp, 7:ai, 8:rc, 9:rd, 10:zc, 11:zd
    [6,7,8,9,10,11].forEach(colIdx => {
      var val = String(data[i][colIdx] || "").trim();
      if (!val) return;
      if (!consensusMap[pId][colIdx]) consensusMap[pId][colIdx] = {};
      consensusMap[pId][colIdx][val] = (consensusMap[pId][colIdx][val] || 0) + 1;
    });
  }

  // 3. Aplicar inferencia e Intento Web
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    
    // Ignorar si ya tiene cod_localidad y jerarquía (ya está completa)
    if (row[0] && row[8] && row[10] && row[0] !== "0") continue;

    var pId = String(row[3]).trim();
    var locName = String(row[1] || "").trim();
    var provName = String(row[4] || "").trim();
    var cp = String(row[5] || "").trim();
    
    var inf = PROV_METADATA[pId];
    var changedThisRow = false;
    
    // Columnas a completar: 6:rp, 7:ai, 8:rc, 9:rd, 10:zc, 11:zd
    var fields = { 6:'rp', 7:'ai', 8:'rc', 9:'rd', 10:'zc', 11:'zd' };
    
    // A. Lógica Interna
    for (var colIdx in fields) {
      colIdx = parseInt(colIdx);
      if (String(row[colIdx] || "").trim() === "") {
        if (inf && inf[fields[colIdx]]) {
          row[colIdx] = inf[fields[colIdx]];
          changedThisRow = true;
          counts.infer_map++;
        } 
        else if (consensusMap[pId] && consensusMap[pId][colIdx]) {
          var bestVal = "";
          var maxCount = 0;
          for (var val in consensusMap[pId][colIdx]) {
            if (consensusMap[pId][colIdx][val] > maxCount) {
              maxCount = consensusMap[pId][colIdx][val];
              bestVal = val;
            }
          }
          if (bestVal) {
            row[colIdx] = bestVal;
            changedThisRow = true;
            counts.infer_consensus++;
          }
        }
      }
    }

    // B. Lógica Web (Límite aplicado aquí)
    var incompleta = !row[0] || row[0] === "0" || !row[1] || !row[3];
    if (incompleta && locName && counts.web < MAX_WEB_CALLS) {
      var searchRes = buscarEnGeoref(cp, locName, provName);
      if (searchRes) {
        if (searchRes.length === 1) {
          var best = searchRes[0];
          if (!row[0] || row[0] === "0") { row[0] = best.cod_localidad; changedThisRow = true; }
          if (!row[3]) { row[3] = best.provincia_id; changedThisRow = true; }
          if (!row[4]) { row[4] = best.provincia_nombre; changedThisRow = true; }
          if (!row[2]) { row[2] = best.departamento_nombre; changedThisRow = true; }
          counts.web++;
        } else {
          // AMBIGÜEDAD: Múltiples resultados
          doubtLog.push([row[0], locName, provName, cp, searchRes.length, JSON.stringify(searchRes.map(x => x.nombre + " (" + x.provincia_nombre + ")"))]);
          counts.doubt++;
          counts.web++; // Contamos como llamada aunque sea ambigua
        }
      }
    }

    if (changedThisRow) modificado = true;
  }
  
  if (modificado) {
    sheet.getRange(1, 1, data.length, headers.length).setValues(data);
  }

  // Manejo de dudas
  if (doubtLog.length > 0) {
    var shDoubt = getOrCreateSheet(ss, REVISION_SHEET);
    shDoubt.clearContents();
    shDoubt.getRange(1,1,1,6).setValues([["COD_LOC", "LOCALIDAD", "PROVINCIA", "CP", "RESULTADOS_WEB", "ALTERNATIVAS"]]).setFontWeight("bold");
    shDoubt.getRange(2,1,doubtLog.length, 6).setValues(doubtLog);
    shDoubt.setFrozenRows(1);
  }

  Logger.log("   --- Completado Finalizado: Map=" + counts.infer_map + ", Consensus=" + counts.infer_consensus + ", Web=" + counts.web + ", Ambiguos=" + counts.doubt);
}

/**
 * Busca en la API de Georef AR. Devuelve un array de posibles resultados.
 */
function buscarEnGeoref(cp, nombre, provincia) {
  try {
    var query = "";
    if (cp && cp.length >= 4 && cp !== "0") {
      query = "municipio=" + encodeURIComponent(cp); 
    } else {
      var n = limpiarNombre(nombre);
      var p = limpiarNombre(provincia);
      if (!n) return null;
      query = "nombre=" + encodeURIComponent(n);
      if (p) query += "&provincia=" + encodeURIComponent(p);
    }
    
    var res = UrlFetchApp.fetch("https://apis.datos.gob.ar/georef/api/localidades?campos=completo&max=5&" + query, { muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) return null;
    
    var json = JSON.parse(res.getContentText());
    if (json.localidades && json.localidades.length > 0) {
      return json.localidades.map(l => ({
        nombre: l.nombre,
        provincia_id: l.provincia.id,
        provincia_nombre: l.provincia.nombre,
        departamento_nombre: l.departamento.nombre,
        cod_localidad: l.id
      }));
    }
  } catch(e) { return null; }
  return null;
}

function getOrCreateSheet(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function generarMasterLocalizacion() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var masterMap = {}; 
  
  // Target Master (0-12): 0:reg_p, 1:amba, 2:reg_c, 3:reg_d, 4:zona_c, 5:zona_d, 6:prov_c, 7:prov_d, 8:cod_l, 9:loc, 10:dpto, 11:cp, 12:dir
  var fuentes = [
    {
      nombre: 'BD 202603',
      cp_idx: 29,
      // Índices reales (base 0): N(13):rp, O(14):amba, U(20):rc, V(21):rd, W(22):zc, X(23):zd, Y(24):pc, Z(25):pd, AA(26):lc, AB(27):ld, AC(28):dpto, AD(29):cp, T(19):dir
      mapping: { 0:13, 1:14, 2:20, 3:21, 4:22, 5:23, 6:24, 7:25, 8:26, 9:27, 10:28, 11:29, 12:19 }
    },
    {
      nombre: 'ID localidad-provincia-region',
      cp_idx: -1, 
      // Índices reales (base 0): B(1):rp, C(2):amba, D(3):rc, E(4):rd, F(5):zc, G(6):zd, H(7):pc, I(8):pd, J(9):lc, K(10):ld
      mapping: { 0:1, 1:2, 2:3, 3:4, 4:5, 5:6, 6:7, 7:8, 8:9, 9:10 }
    }
  ];

  fuentes.forEach(f => {
    var sheet = ss.getSheetByName(f.nombre);
    if (!sheet) return;
    
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var cp = f.cp_idx !== -1 ? String(row[f.cp_idx]).trim() : "";
      var cod = String(row[f.mapping[8]||""]).trim(); // cod_localidad es 8 en Master
      
      if ((!cp || cp === "" || cp === "0") && (!cod || cod === "" || cod === "0")) continue;
      
      var normRow = fill1d(13, "");
      for (var masterIdx in f.mapping) {
        var sourceIdx = f.mapping[masterIdx];
        if (sourceIdx !== -1) normRow[masterIdx] = String(row[sourceIdx] || "").trim();
      }
      
      var key = normRow.join("||");
      if (!masterMap[key]) {
        masterMap[key] = normRow;
      }
    }
  });

  var finalData = Object.values(masterMap);

  // Ordenar por CP, luego localidad, luego dirección
  finalData.sort((a, b) => {
    var na = parseInt(a[11]), nb = parseInt(b[11]);
    if (na !== nb && !isNaN(na) && !isNaN(nb)) return na - nb;
    if (a[9] !== b[9]) return a[9].localeCompare(b[9]);
    return a[12].localeCompare(b[12]);
  });

  var headers = ['region_pais_cod','amba/interior','region_cod','region_descr','zona_cod','zona_descr','cod_prov','provincia','cod_localidad','localidad','departamento','cp', 'direccion'];
  var sheetMaster = ss.getSheetByName("Master Localizacion") || ss.insertSheet("Master Localizacion");
  sheetMaster.clearContents().clearFormats();
  
  var total = [headers].concat(finalData);
  Logger.log("   --- Master Localizacion: " + (total.length - 1) + " registros procesados.");
  
  // Escritura por bloques para evitar timeout
  escribirEnBloques(sheetMaster, total);

  SpreadsheetApp.getUi().alert("Master generado con " + finalData.length + " combinaciones únicas para revisión.");
}

/**
 * Genera la tabla definitiva de referencia: Maestro Localidades.
 * Unifica por cod_localidad para no perder datos, pero mantiene el CP como dato clave.
 */
function generarMaestroLocalidades() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var masterSheet = ss.getSheetByName("Master Localizacion");
  if (!masterSheet) return alert("Primero genera el Master.");
  
  var map = {}; // Clave: cod_localidad (o name_...)
  var discLog = []; // Log de discrepancias
  
  // Fuentes en orden de prioridad: La última procesada es la que aporta el valor final (sobrescribe)
  var sheetNames = ['Master Localizacion', 'ID localidad-provincia-region'];
  
  sheetNames.forEach(sName => {
    var sh = ss.getSheetByName(sName);
    if (!sh) return;
    
    var data = sh.getDataRange().getValues();
    if (data.length < 2) return;
    
    // Identificar columnas según la hoja
    var colMap = {};
    if (sName === 'Master Localizacion') {
      // 0:rp, 1:ai, 2:rc, 3:rd, 4:zc, 5:zd, 6:pc, 7:pd, 8:lc, 9:ld, 10:dpto, 11:cp
      colMap = { rp:0, ai:1, rc:2, rd:3, zc:4, zd:5, cpv:6, prov:7, cod:8, loc:9, dpto:10, cp:11 };
    } else {
      // ID localidad (Indices 0-10): 1:rp, 2:ai, 3:rc, 4:rd, 5:zc, 6:zd, 7:pc, 8:pd, 9:lc, 10:ld
      colMap = { rp:1, ai:2, rc:3, rd:4, zc:5, zd:6, cpv:7, prov:8, cod:9, loc:10, dpto:-1, cp:-1 };
    }
    
    for (var i = 1; i < data.length; i++) {
      var r = data[i];
      var cod = String(colMap.cod !== -1 ? r[colMap.cod] : "").trim();
      var name = (limpiarNombre((colMap.loc!==-1?r[colMap.loc]:"")||"") + "|" + limpiarNombre((colMap.prov!==-1?r[colMap.prov]:"")||"")).toLowerCase();
      var key = (cod && cod !== "0") ? ("cod_" + cod) : ("name_" + name);
      
      // norm now has 12 columns: 0:cod, 1:loc, 2:dpto, 3:cpv, 4:prov, 5:cp, 6:rp, 7:ai, 8:rc, 9:rd, 10:zc, 11:zd
      var normRow = [
        cod,
        colMap.loc !== -1 ? r[colMap.loc] : "",
        colMap.dpto !== -1 ? r[colMap.dpto] : "",
        colMap.cpv !== -1 ? r[colMap.cpv] : "",
        colMap.prov !== -1 ? r[colMap.prov] : "",
        colMap.cp !== -1 ? r[colMap.cp] : "",
        colMap.rp !== -1 ? r[colMap.rp] : "",
        colMap.ai !== -1 ? r[colMap.ai] : "",
        colMap.rc !== -1 ? r[colMap.rc] : "",
        colMap.rd !== -1 ? r[colMap.rd] : "",
        colMap.zc !== -1 ? r[colMap.zc] : "",
        colMap.zd !== -1 ? r[colMap.zd] : ""
      ];

      if (!map[key]) {
        map[key] = normRow;
      } else {
        var current = map[key];
        // Auditoría
        if (normRow[1] && current[1] && normRow[1] !== current[1]) discLog.push([key, "Localidad", sName, normRow[1], "Previo", current[1]]);
        if (normRow[4] && current[4] && normRow[4] !== current[4]) discLog.push([key, "Provincia", sName, normRow[4], "Previo", current[4]]);
        
        // Merge con prioridad actual (12 columns)
        for (var c = 0; c < 12; c++) {
          if (String(normRow[c] || "").trim() !== "") current[c] = normRow[c];
        }
      }
    }
  });
  
  var finalRows = Object.values(map);
  finalRows.sort((a, b) => String(a[1]).localeCompare(String(b[1]))); 
  
  var headers = ['cod_localidad','localidad','departamento','cod_provincia','provincia','cp','region_pais_cod','amba/interior','region_cod','region_descr','zona_cod','zona_descr'];
  var sheet = ss.getSheetByName("Maestro Localidades") || ss.insertSheet("Maestro Localidades");
  sheet.clearContents().clearFormats();
  
  var total = [headers].concat(finalRows);
  
  // Escritura por bloques para evitar timeout
  escribirEnBloques(sheet, total);
  
  sheet.getRange(1, 1, 1, headers.length).setBackground("#bf360c").setFontColor("#ffffff").setFontWeight("bold");
  sheet.setFrozenRows(1);
  
  // Log de Discrepancias
  var shDisc = ss.getSheetByName("AUDITORIA_DISCREPANCIAS") || ss.insertSheet("AUDITORIA_DISCREPANCIAS");
  shDisc.clearContents();
  if (discLog.length > 0) {
    shDisc.getRange(1, 1, 1, 6).setValues([["Clave", "Campo", "Origen", "Valor Nuevo", "Contexto", "Valor Anterior"]]);
    shDisc.getRange(2, 1, discLog.length, 6).setValues(discLog);
    shDisc.getRange(1, 1, 1, 6).setBackground("#d84315").setFontColor("#ffffff").setFontWeight("bold");
  } else {
    shDisc.getRange(1, 1).setValue("No hay discrepancias detectadas entre fuentes.");
  }
  
  Logger.log("   --- Maestro: " + (total.length - 1) + " localidades únicas finales.");
}

function limpiarNombre(n) {
  if (!n) return "";
  return String(n).toLowerCase().replace(/sucursal|suc\.|barrio|b\.|zona/gi, "").trim();
}

function fill1d(c, clr) { var r=[]; for(var i=0;i<c;i++) r.push(clr); return r; }

/**
 * Escribe datos en una hoja en bloques de N filas para evitar timeouts del servicio de Spreadsheets.
 */
function escribirEnBloques(sheet, values, blockSize) {
  if (!values || values.length === 0) return;
  blockSize = blockSize || 1000;
  
  var totalRows = values.length;
  var cols = values[0].length;
  
  for (var i = 0; i < totalRows; i += blockSize) {
    var chunk = values.slice(i, i + blockSize);
    sheet.getRange(i + 1, 1, chunk.length, cols).setValues(chunk);
    SpreadsheetApp.flush(); // Forzar actualización y liberar memoria/conexión
    Logger.log("      --- Escritos bloque: de fila " + (i+1) + " a " + (i + chunk.length));
  }
}
