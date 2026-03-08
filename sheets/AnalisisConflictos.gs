/**
 * SCRIPT DE AUDITORÍA Y CONSOLIDACIÓN DEL MAESTRO
 * Reúne todas las variantes de cada cod_localidad para decisión final.
 */
function ejecutarAuditoriaMaestro() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var masterMap = {}; // Clave: cod_localidad -> { campo: [val1, val2...], counts: { val: N } }
  
  // 1. Configuración de Fuentes y Mapeos
  // 0:cod, 1:loc, 2:dpto, 3:cpv, 4:prov, 5:cp, 6:rp, 7:ai, 8:rc, 9:rd, 10:zc, 11:zd
  var FIELDS = ['cod_localidad','localidad','departamento','cod_provincia','provincia','cp','region_pais_cod','amba/interior','region_cod','region_descr','zona_cod','zona_descr'];
  
  var fuentes = [
    {
      nombre: 'ID localidad-provincia-region',
      // Indices corregidos: 9:cod, 10:loc, -1:dpto, 7:prov_cod, 8:prov, -1:cp, 1:rp, 2:ai, 3:rc, 4:rd, 5:zc, 6:zd
      mapping: { 0:9, 1:10, 2:-1, 3:7, 4:8, 5:-1, 6:1, 7:2, 8:3, 9:4, 10:5, 11:6 },
      dir_idx: -1
    },
    {
      nombre: 'Master Localizacion',
      mapping: { 0:8, 1:9, 2:10, 3:6, 4:7, 5:11, 6:0, 7:1, 8:2, 9:3, 10:4, 11:5 },
      dir_idx: 12
    },
    {
      nombre: 'BD 202603',
      mapping: { 0:26, 1:27, 2:28, 3:24, 4:25, 5:29, 6:13, 7:14, 8:20, 9:21, 10:22, 11:23 },
      dir_idx: 19
    }
  ];

  // 2. Recolección de Variantes
  fuentes.forEach(f => {
    var sh = ss.getSheetByName(f.nombre);
    if (!sh) return;
    var data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var cod = String(row[f.mapping[0]] || "").trim();
      if (!cod || cod === "0" || cod === "") continue;
      
      if (!masterMap[cod]) {
        masterMap[cod] = {};
        FIELDS.forEach((_, idx) => masterMap[cod][idx] = { vals: {}, valDir: {} });
      }
      
      var dirRel = f.dir_idx !== -1 ? String(row[f.dir_idx] || "").trim() : "";
      
      for (var targetIdx in f.mapping) {
        var sourceIdx = f.mapping[targetIdx];
        if (sourceIdx === -1 || sourceIdx === undefined) continue;
        var val = String(row[sourceIdx] || "").trim();
        if (val) {
          masterMap[cod][targetIdx].vals[val] = (masterMap[cod][targetIdx].vals[val] || 0) + 1;
          if (dirRel) {
            if (!masterMap[cod][targetIdx].valDir[val]) masterMap[cod][targetIdx].valDir[val] = [];
            if (masterMap[cod][targetIdx].valDir[val].indexOf(dirRel) === -1) {
              masterMap[cod][targetIdx].valDir[val].push(dirRel);
            }
          }
        }
      }
    }
  });

  // 3. Procesamiento de Resultados y Enriquecimiento Web
  var reportRows = [];
  var sortedCodes = Object.keys(masterMap).sort();
  var MAX_WEB = 50; // Para no saturar
  var webCalls = 0;

  // 4. Chequeo Global de CP y Candidatos a Fusión
  var cpToCodes = {}; 
  var nameCpToCodes = {}; // "nombre|cp" -> [cod1, cod2...]
  
  sortedCodes.forEach(cod => {
    var item = masterMap[cod];
    var locName = limpiarNombreLoc(Object.keys(item[1].vals)[0] || "");
    var cpVals = Object.keys(item[5].vals);
    
    cpVals.forEach(cp => {
      if (!cpToCodes[cp]) cpToCodes[cp] = [];
      if (cpToCodes[cp].indexOf(cod) === -1) cpToCodes[cp].push(cod);
      
      var key = locName + "|" + cp;
      if (locName && cp) {
        if (!nameCpToCodes[key]) nameCpToCodes[key] = [];
        if (nameCpToCodes[key].indexOf(cod) === -1) nameCpToCodes[key].push(cod);
      }
    });
  });

  // 4. Procesar y Generar Reporte
  var reportRows = [];
  var MAX_WEB = 500; 
  var webCalls = 0;

  sortedCodes.forEach(cod => {
    try {
      var item = masterMap[cod];
      var finalRow = [cod]; 
      var status = "OK";
      
      for (var i = 1; i < FIELDS.length; i++) {
        var fData = item[i];
        var uniqueVals = Object.keys(fData.vals);
        
        // 1. Caso Faltante: Buscar por ID
        if (uniqueVals.length === 0) {
          if (webCalls < MAX_WEB && (i === 1 || i === 2 || i === 4)) {
            var webRes = buscarEnGeorefAuditoria(cod, null, null);
            if (webRes) {
              var valW = (i===1)? webRes.nombre : (i===2 ? (webRes.departamento?webRes.departamento.nombre:"") : (webRes.provincia?webRes.provincia.nombre:""));
              if (valW) {
                finalRow.push("[WEB] " + valW);
                webCalls++;
                if (status === "OK" || status === "INCOMPLETO") status = "CON_SUGERENCIA_WEB";
                continue;
              }
            }
          }
          finalRow.push("[FALTANTE]");
          if (status === "OK") status = "INCOMPLETO";
        } 
        else if (uniqueVals.length === 1) {
          finalRow.push(uniqueVals[0]);
        } 
        else {
          var mostrarDirs = (i === 1 || i === 2 || i === 5);
          var details = uniqueVals.map(v => {
            var d = (mostrarDirs && fData.valDir[v]) ? " (" + fData.valDir[v].join(", ") + ")" : "";
            return v + d;
          }).join(" | ");
          finalRow.push("[" + details + "]");
          if (status !== "CON_SUGERENCIA_WEB" && status !== "CANDIDATO_FUSION") status = "CONFLICTO";
        }
      }

      // --- LOGICA DE RECUPERACIÓN PROFUNDA DE CP (Índice 5) ---
      if (webCalls < MAX_WEB && (finalRow[5] === "[FALTANTE]" || finalRow[5].indexOf("|") !== -1)) {
        var cleanLoc = limpiarNombreLocCompleto(finalRow[1]);
        var provName = String(finalRow[4]).trim();
        var dptoName = limpiarNombreLocCompleto(finalRow[2]);
        var cpEncontrado = "";

        // ESTRATEGIA 1: Por Dirección (Si existe)
        var fallbackDir = "";
        for (var k in item) {
           var vD = item[k].valDir; var kD = Object.keys(vD);
           if (kD.length > 0) { fallbackDir = vD[kD[0]][0]; break; }
        }
        if (fallbackDir) {
           var resDir = buscarEnGeorefAuditoria(null, fallbackDir, cleanLoc);
           if (resDir && resDir.codigo_postal) cpEncontrado = resDir.codigo_postal;
        }

        // ESTRATEGIA 2: Por Nombre + Provincia + Departamento
        if (!cpEncontrado && cleanLoc && provName && provName !== "[FALTANTE]") {
           var resLoc = buscarEnGeorefAuditoria(null, null, cleanLoc, provName, dptoName);
           if (resLoc && resLoc.codigo_postal) cpEncontrado = resLoc.codigo_postal;
           // Fallback: Si Georef no trae CP directo en localidad, a veces es por ambigüedad
        }

        if (cpEncontrado) {
          finalRow[5] = "[S_WEB] " + cpEncontrado + " | " + finalRow[5];
          if (status !== "CANDIDATO_FUSION") status = "CON_SUGERENCIA_WEB";
          webCalls++;
        }
      }
      
      // Alerta de Fusión
      var cpClean = String(finalRow[5]).replace("[S_WEB] ", "").split(" | ")[0].replace("[", "").replace("]", "").trim();
      var locClean = limpiarNombreLoc(finalRow[1].replace("[WEB] ", ""));
      var fKey = locClean + "|" + cpClean;
      var globalAlert = "";
      if (cpClean && nameCpToCodes[fKey] && nameCpToCodes[fKey].length > 1) {
        globalAlert = "FUSION: IDs [" + nameCpToCodes[fKey].join(", ") + "]";
        status = "CANDIDATO_FUSION";
      } else if (cpClean && cpToCodes[cpClean] && cpToCodes[cpClean].length > 1) {
        globalAlert = "CP Compartido en IDs [" + cpToCodes[cpClean].join(", ") + "]";
      }
      
      if (status !== "OK") {
        // [MAESTRO (12)] + [ALERTA (1)] + [ESTADO (1)] = 14 columnas
        reportRows.push(finalRow.concat([globalAlert, status]));
      }
    } catch(e) { Logger.log("Error en " + cod + ": " + e.message); }
  });

  // 4. Ordenar y Salida (Idéntico)
  var priority = { "CANDIDATO_FUSION": 1, "CON_SUGERENCIA_WEB": 2, "CONFLICTO": 3, "INCOMPLETO": 4 };
  reportRows.sort((a, b) => {
    var pa = priority[a[a.length-1]] || 99;
    var pb = priority[b[b.length-1]] || 99;
    if (pa !== pb) return pa - pb;
    return String(a[1]).localeCompare(String(b[1]));
  });

  var outHeaders = FIELDS.concat(['ALERTA_GLOBAL', 'ESTADO_AUDITORIA']);
  var outSheet = ss.getSheetByName("AUDITORIA_CONSOLIDADA_MAESTRO") || ss.insertSheet("AUDITORIA_CONSOLIDADA_MAESTRO");
  outSheet.clearContents().clearFormats();
  if (reportRows.length === 0) { outSheet.getRange(1,1).setValue("Todo OK"); return; }
  escribirEnBloques(outSheet, [outHeaders].concat(reportRows));
  
  // Estética
  outSheet.getRange(1, 1, 1, outHeaders.length).setBackground("#37474f").setFontColor("#ffffff").setFontWeight("bold");
  outSheet.setFrozenRows(1);
  var sCol = outHeaders.length;
  var sColL = columnToLetter(sCol);
  var rangeF = outSheet.getRange(2, 1, reportRows.length, sCol);
  outSheet.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied("=$"+sColL+"2=\"CANDIDATO_FUSION\"").setBackground("#fff59d").setRanges([rangeF]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied("=$"+sColL+"2=\"CON_SUGERENCIA_WEB\"").setBackground("#bbdefb").setRanges([rangeF]).build()
  ]);

  Logger.log("Auditoría Finalizada. " + reportRows.length + " problemas. WebHits: " + webCalls);
}

/**
 * APLICA LAS CORRECCIONES SUGERIDAS EN EL REPORTE A UNA HOJA NUEVA (MODO SEGURO)
 */
function aplicarCorreccionesAuditoria() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var shAudit = ss.getSheetByName("AUDITORIA_CONSOLIDADA_MAESTRO");
  var shMaestroOriginal = ss.getSheetByName("Maestro Localidades");
  var shBD = ss.getSheetByName("BD 202603");
  var shSucs = ss.getSheetByName("Tabla Sucursales Localizacion");
  
  if (!shAudit || !shMaestroOriginal) {
    SpreadsheetApp.getUi().alert("Faltan hojas críticas (Audit o Maestro).");
    return;
  }

  // 1. Preparar la hoja de destino (No destructivo sobre el original)
  var targetName = "Maestro Localidades - Corregido";
  var shMaestroCorregido = ss.getSheetByName(targetName);
  
  if (!shMaestroCorregido) {
    shMaestroCorregido = shMaestroOriginal.copyTo(ss).setName(targetName);
  } else {
    var confirm = SpreadsheetApp.getUi().alert("La hoja '" + targetName + "' ya existe. ¿Deseas sobreescribirla?", SpreadsheetApp.getUi().ButtonSet.YES_NO);
    if (confirm !== SpreadsheetApp.getUi().Button.YES) return;
    shMaestroCorregido.clearContents().clearFormats();
    shMaestroOriginal.getDataRange().copyTo(shMaestroCorregido.getRange(1,1));
  }
  
  var auditData = shAudit.getDataRange().getValues();
  var maestroCorregidoData = shMaestroCorregido.getDataRange().getValues();
  var headersAudit = auditData[0];
  var colEstado = headersAudit.indexOf("ESTADO_AUDITORIA");
  var colAlerta = headersAudit.indexOf("ALERTA_GLOBAL");
  
  if (auditData.length < 2) return;
  
  var fusionsToProcess = []; // [{oldId, newId}]
  var updatesCount = 0;
  
  // 2. Mapear Maestro Corregido para acceso rápido
  var maestroMap = {}; 
  for (var i = 1; i < maestroCorregidoData.length; i++) {
    maestroMap[String(maestroCorregidoData[i][0])] = i + 1; // nro de fila
  }

  // 3. Procesar Sugerencias del Audit
  for (var i = 1; i < auditData.length; i++) {
    var row = auditData[i];
    var cod = String(row[0]);
    var estado = String(row[colEstado]);
    var alerta = String(row[colAlerta]);

    // Limpiar fila Audit (quitar tags [WEB], corchetes de conflicto, etc)
    var cleanedRow = row.slice(0, 12).map(val => {
       var s = String(val);
       if (s.startsWith("[S_WEB]")) s = s.split("|")[0].replace("[S_WEB]", "").trim();
       if (s.startsWith("[WEB]")) s = s.replace("[WEB]", "").trim();
       if (s.startsWith("[")) s = s.replace("[", "").split("|")[0].split("(")[0].replace("]", "").trim(); 
       if (s === "[FALTANTE]") s = "";
       return s;
    });

    // Detectar FUSION (Pero no borrar todavía para no arruinar el mapeo)
    if (estado === "CANDIDATO_FUSION" && alerta.includes("FUSION: IDs")) {
      var idsMatch = alerta.match(/\[(.*?)\]/);
      if (idsMatch) {
         var ids = idsMatch[1].split(", ");
         var targetId = ids[0];
         if (cod !== targetId) {
           fusionsToProcess.push({ oldId: cod, newId: targetId });
           continue; 
         }
      }
    }

    // Actualizar fila en la hoja corregida
    var rowIdx = maestroMap[cod];
    if (rowIdx) {
      shMaestroCorregido.getRange(rowIdx, 1, 1, 12).setValues([cleanedRow]);
      updatesCount++;
    }
  }

  // 4. Procesar Fusiones (Search & Replace en cascada + Deletion)
  if (fusionsToProcess.length > 0) {
    var confirmFus = SpreadsheetApp.getUi().alert("Se detectaron " + fusionsToProcess.length + " unificaciones. ¿Proceder?", SpreadsheetApp.getUi().ButtonSet.YES_NO);
    if (confirmFus === SpreadsheetApp.getUi().Button.YES) {
      fusionsToProcess.forEach(f => {
        if (shBD) shBD.createTextFinder(f.oldId).matchEntireCell(true).replaceAllWith(f.newId);
        if (shSucs) shSucs.createTextFinder(f.oldId).matchEntireCell(true).replaceAllWith(f.newId);
        
        var rowToDel = maestroMap[f.oldId];
        if (rowToDel) {
          shMaestroCorregido.deleteRow(rowToDel);
          // Re-mapear (corrimiento de filas por borrado)
          for (var key in maestroMap) {
            if (maestroMap[key] > rowToDel) maestroMap[key]--;
          }
        }
      });
    }
  }

  SpreadsheetApp.getUi().alert("Limpieza finalizada. Registros: " + updatesCount);
}

/**
 * MENU PERSONALIZADO
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('📊 Auditoría & Limpieza')
      .addItem('🔍 1. Ejecutar Auditoría Maestro', 'ejecutarAuditoriaMaestro')
      .addSeparator()
      .addItem('🎯 2. Aplicar Arreglos a "Maestro Corregido"', 'aplicarCorreccionesAuditoria')
      .addToUi();
}

/**
 * Busca en Georef AR. (Modo: ID, Direccion+Loc, o Loc+Prov+Dpto)
 */
function buscarEnGeorefAuditoria(cod, direccion, localidad, provincia, departamento) {
  try {
    var url = "https://apis.datos.gob.ar/georef/api/";
    if (direccion && localidad) {
      url += "direcciones?direccion=" + encodeURIComponent(direccion) + "&localidad=" + encodeURIComponent(localidad) + "&max=1";
    } else if (localidad && provincia) {
      url += "localidades?nombre=" + encodeURIComponent(localidad) + "&provincia=" + encodeURIComponent(provincia);
      if (departamento) url += "&departamento=" + encodeURIComponent(departamento);
      url += "&max=5"; // Aumentamos para elegir la mejor coincidencia
    } else if (cod) {
      url += "localidades?id=" + cod + "&campos=completo";
    } else {
      return null;
    }

    var res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (res.getResponseCode() === 200) {
      var json = JSON.parse(res.getContentText());
      if (direccion) {
        if (json.direcciones && json.direcciones.length > 0) return json.direcciones[0];
      } else {
        if (json.localidades && json.localidades.length > 0) {
           // Si hay varios, intentamos filtrar por departamento si lo tenemos
           if (departamento && json.localidades.length > 1) {
              var best = json.localidades.find(l => l.departamento && l.departamento.nombre.toLowerCase().includes(departamento.toLowerCase()));
              if (best) return best;
           }
           return json.localidades[0];
        }
      }
    }
  } catch(e) { Logger.log("Error Georef: " + e.message); }
  return null;
}

function limpiarNombreLocCompleto(n) { 
  if (!n) return "";
  return String(n).toLowerCase()
    .replace(/\[web\]|\[s_web\]|\[faltante\]/gi, "")
    .replace(/sucursal|suc\.|barrio|b\.|zona|z\/|centro|ctro\.|estacion|est\.|delegación|deleg\./gi, "")
    .replace(/sección|secc\.|mza\.|lote|parcela|piso/gi, "")
    .replace(/\(.*?\)/g, "") // Quitar paréntesis
    .replace(/[0-9]/g, "") // Quitar números que a veces son de calles o lotes
    .replace(/-/g, " ")
    .trim(); 
}

function limpiarNombreLoc(n) { return limpiarNombreLocCompleto(n); }

function escribirEnBloques(sheet, values, blockSize) {
  if (!values || values.length === 0) return;
  blockSize = blockSize || 1000;
  var totalRows = values.length;
  var cols = values[0].length;
  for (var i = 0; i < totalRows; i += blockSize) {
    var chunk = values.slice(i, i + blockSize);
    sheet.getRange(i + 1, 1, chunk.length, cols).setValues(chunk);
    SpreadsheetApp.flush();
  }
}

function columnToLetter(column) {
  var temp, letter = "";
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}
