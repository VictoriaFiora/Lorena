// ============================================================
// DataCleaner.gs — Script de limpieza de BD 202603
// ============================================================
// FLUJO:
//   runAll()       → corre todo de una vez
//   runAnalisis()  → solo detecta errores y duplicados
//   runTablas()    → solo genera tablas editables y BD corregida
// ============================================================

var CLEANER_CONFIG = {
  SHEET_BD:             'BD 202603',
  SHEET_LOC_REF:        'ID localidad-provincia-region',
  SHEET_RESULT_DESVIOS: 'Desvios Detectados',
  SHEET_RESULT_DUPS:    'Locales Duplicados',
  SHEET_LOCALIDADES:    'Maestro Localidades - Corregido', // Nuestra nueva fuente de verdad completa y corregida
  SHEET_SUCURSALES:     'Tabla Sucursales Localizacion',
  SHEET_BD_CORREGIDA:   'BD 202603 - Corregida',

  // Columnas BD (base 0)
  BD_ID: 0, 
  BD_GLN: 1, 
  BD_TRUCK: 3, 
  BD_REG_P_COD: 13, // N
  BD_AMBA: 14,      // O
  BD_SUC_DESC: 16, 
  BD_DIRECCION: 19, // T
  BD_LAT: 29,       // AD
  BD_LNG: 30,       // AE
  // Estas son las que se sobreescriben con el Maestro
  BD_REG_COD: 20,   // U
  BD_REG_DESC: 21,  // V
  BD_ZONA_COD: 22,  // W
  BD_COD_PROV: 23,  // X
  BD_PROVINCIA: 24, // Y
  BD_COD_LOC: 25,   // Z (Key para XLOOKUP)
  BD_LOCALIDAD: 26, // AA
  BD_DPTO: 27,      // AB
  BD_CP: 28,        // AC

  // Columnas Referencia (base 0)
  LOC_ID: 0, LOC_REGION_COD: 3, LOC_REGION_DESC: 4, LOC_ZONA_COD: 5,
  LOC_ZONA_DESC: 6, LOC_COD_PROV: 7, LOC_PROVINCIA: 8, LOC_COD_LOC: 9, LOC_LOCALIDAD: 10
};

function runTablas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log('=== GENERANDO BD CORREGIDA ===');
  crearBDCorregida(ss);
  Logger.log('=== BD CORREGIDA FINALIZADA ===');
}

function crearBDCorregida(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  var bdSheet = ss.getSheetByName(CLEANER_CONFIG.SHEET_BD);
  if (!bdSheet) return;
  var lastRow = bdSheet.getLastRow(), lastCol = bdSheet.getLastColumn();
  var bdData = bdSheet.getRange(1,1,lastRow,lastCol).getValues();
  var idDisp = bdSheet.getRange(1, 1, lastRow, 1).getDisplayValues();
  for (var r=0; r<bdData.length; r++) bdData[r][CLEANER_CONFIG.BD_ID] = idDisp[r][0];

  var bdHeaders = bdData[0];
  var seenId = {}, seenGln = {}, seenDir = {}, seenTruck = {};
  var corr = [bdHeaders], eliminados = [];

  for (var r=1; r<bdData.length; r++) {
    var row = bdData[r], id = row[CLEANER_CONFIG.BD_ID], idKey = String(id).trim().toLowerCase();
    var gln = String(row[CLEANER_CONFIG.BD_GLN]).trim();
    var dir = String(row[CLEANER_CONFIG.BD_DIRECCION]).trim().toLowerCase();
    
    // normalización forzada de Truck: Debe ser el ID sin "suc_"
    var tr = idKey ? idKey.replace("suc_", "") : String(row[CLEANER_CONFIG.BD_TRUCK]).trim();
    var trOriginal = String(row[CLEANER_CONFIG.BD_TRUCK]).trim();
    
    // Si el Camión estaba repetido en la BASE ORIGINAL, lo reportamos
    if (trOriginal && trOriginal!=='0' && trOriginal.toLowerCase()!=='false' && trOriginal.toLowerCase()!=='—') {
      var existing = seenTruck[trOriginal];
      if (existing) {
        var idPrev = String(existing[CLEANER_CONFIG.BD_ID]).toLowerCase();
        var idActual = idKey;
        // Prioridad: 40xx/41xx son originales. 24xx son correcciones.
        var actualEsPrioritario = (idActual.indexOf("suc_40") === 0 || idActual.indexOf("suc_41") === 0);
        var previoEsPrioritario = (idPrev.indexOf("suc_40") === 0 || idPrev.indexOf("suc_41") === 0);
        
        if (actualEsPrioritario && !previoEsPrioritario) {
          // El nuevo es mejor. El viejo se vuelve el "corregido"
          eliminados.push({tipo:'TRUCK REPETIDO (ORIGINAL)', keeper:row, elim:existing, note:'CORREGIDO'});
          seenTruck[trOriginal] = row;
        } else {
          // El viejo es mejor o igual. El nuevo se vuelve el "corregido"
          eliminados.push({tipo:'TRUCK REPETIDO (ORIGINAL)', keeper:existing, elim:row, note:'CORREGIDO'});
        }
      } else {
        seenTruck[trOriginal] = row;
      }
    }

    row[CLEANER_CONFIG.BD_TRUCK] = tr; // Aplicamos la corrección para la BD Corregida

    // Filtros de duplicados CRÍTICOS (Estos SÍ se eliminan)
    if (id && seenId[idKey]) { 
      eliminados.push({tipo:'ID DUPLICADO', keeper:seenId[idKey], elim:row}); 
      continue; 
    }
    if (gln && seenGln[gln]) { 
      eliminados.push({tipo:'GLN DUPLICADO', keeper:seenGln[gln], elim:row}); 
      continue; 
    }
    if (dir && seenDir[dir]) { 
      eliminados.push({tipo:'DIRECCIÓN DUPLICADA', keeper:seenDir[dir], elim:row}); 
      continue; 
    }

    // Guardar para siguientes comparaciones
    seenId[idKey]=row; 
    if(gln) seenGln[gln]=row; 
    if(dir) seenDir[dir]=row;

    corr.push(row);
  }

  var sheet = getOrCreateSheet(ss, CLEANER_CONFIG.SHEET_BD_CORREGIDA);
  sheet.clearContents(); sheet.clearFormats();
  sheet.getRange(1,1,corr.length,corr[0].length).setValues(corr);

  if (corr.length > 1) {
    var formulasN = []; // Para columnas N y O (14, 15)
    var formulasT = []; // Para columnas T a AF (20 a 32)
    var SL = CLEANER_CONFIG.SHEET_SUCURSALES, TL = CLEANER_CONFIG.SHEET_LOCALIDADES;
    
    // Nuevo Mapeo Maestro (A=cod_l):
    // A=cod_l(0), B=loc(1), C=dpto(2), D=cod_p(3), E=prov(4), F=cp(5), G=r_p_c(6), H=amba(7), I=r_c(8), J=r_d(9), K=z_c(10), L=z_d(11)
    
    // Para N,O (indices 14, 15): N(reg_p)=G, O(amba)=H
    // Para T-AF (indices 20-32):
    // T (pos 0): DIR -> SL!D
    // U (pos 1): reg_c -> TL!I
    // V (pos 2): reg_d -> TL!J
    // W (pos 3): zona_c -> TL!K
    // X (pos 4): zona_d -> TL!L
    // Y (pos 5): cod_p -> TL!D
    // Z (pos 6): prov -> TL!E
    // AA (pos 7): cod_l (Calculated from SUC)
    // AB (pos 8): loc -> TL!B
    // AC (pos 9): dpto -> TL!C
    // AD (pos 10): cp -> TL!F
    // AE, AF: Lat/Lng -> SL!E, SL!F
    
    var locColsArr = ['I','J', 'K', 'D', 'E', null, 'B','C','F'];

    for (var i=1; i<corr.length; i++) {
      var r = i+1, id = '$A'+r, cl = '$Z'+r; // cl = Column Z (cod_localidad)
      
      // Columnas N,O
      formulasN.push([
        "=IFERROR(XLOOKUP("+cl+",'"+TL+"'!$A:$A,'"+TL+"'!$G:$G),\"\")", // N
        "=IFERROR(XLOOKUP("+cl+",'"+TL+"'!$A:$A,'"+TL+"'!$H:$H),\"\")"  // O
      ]);

      var rowT = [];
      // T (Dirección) -> Sucursales Col D
      rowT.push("=IFERROR(XLOOKUP("+id+",'"+SL+"'!$A:$A,'"+SL+"'!$D:$D),\"\")");
      // U-AC
      for (var j=0; j<9; j++) {
        if (j===5) {
          // Z (Cod localidad) -> Sucursales Col G
          rowT.push("=IFERROR(XLOOKUP("+id+",'"+SL+"'!$A:$A,'"+SL+"'!$G:$G),\"\")"); 
        } else {
          rowT.push("=IFERROR(XLOOKUP("+cl+",'"+TL+"'!$A:$A,'"+TL+"'!$"+locColsArr[j]+":$"+locColsArr[j]+"),\"\")");
        }
      }
      // AD, AE (Lat/Lng) -> Sucursales Col E, F
      rowT.push("=IFERROR(XLOOKUP("+id+",'"+SL+"'!$A:$A,'"+SL+"'!$E:$E),\"\")");
      rowT.push("=IFERROR(XLOOKUP("+id+",'"+SL+"'!$A:$A,'"+SL+"'!$F:$F),\"\")");
      formulasT.push(rowT);
    }
    
    if (formulasN.length > 0) {
      sheet.getRange(2, 14, formulasN.length, 2).setFormulas(formulasN); // N, O
      sheet.getRange(2, 20, formulasT.length, 12).setFormulas(formulasT); // T a AE
      sheet.getRange(2, 14, formulasN.length, 2).setBackground('#e3f2fd');
      sheet.getRange(2, 20, formulasT.length, 12).setBackground('#fff8e1');
      Logger.log("   --- BD Corregida: Fórmulas aplicadas en " + formulasT.length + " filas.");
    }
    SpreadsheetApp.flush();
  }
  sheet.getRange(1, 1, 1, corr[0].length).setBackground('#1565c0').setFontColor('#fff').setFontWeight('bold');
  sheet.getRange(1, 20, 1, 13).setBackground('#e65100');

  // Reporte de Eliminados
  var dupHeaders = ['Estado','Motivo'].concat(bdHeaders).concat(['Diferencias']);
  var dupRows = [], dupColors = [];
  eliminados.forEach(d => {
    var diff = [];
    for(var c=0; c<bdHeaders.length; c++) if(String(d.keeper[c])!==String(d.elim[c])) diff.push(bdHeaders[c]);
    var diffStr = diff.length>0 ? diff.join('|') : 'Iguales';
    
    var estado = d.note || 'ELIMINADO';
    dupRows.push(['CONSERVADO', d.tipo].concat(d.keeper).concat([diffStr]));
    dupColors.push(fill1d(dupHeaders.length, '#e8f5e9'));
    dupRows.push([estado, d.tipo].concat(d.elim).concat([diffStr]));
    dupColors.push(fill1d(dupHeaders.length, estado==='ELIMINADO' ? '#fce4ec' : '#fff9c4'));
    dupRows.push(fill1d(dupHeaders.length, ''));
    dupColors.push(fill1d(dupHeaders.length, '#f5f5f5'));
  });

  var dupSheet = getOrCreateSheet(ss, 'Duplicados');
  dupSheet.clearContents(); dupSheet.clearFormats();
  if (dupRows.length > 0) {
    var allD = [dupHeaders].concat(dupRows);
    dupSheet.getRange(1, 1, allD.length, dupHeaders.length).setValues(allD);
    dupSheet.getRange(2, 1, dupRows.length, dupHeaders.length).setBackgrounds(dupColors);
  }
}

function getOrCreateSheet(ss, name) { return ss.getSheetByName(name) || ss.insertSheet(name); }
function norm(v) { return String(v).trim().toLowerCase(); }
function alert(msg) { Logger.log(msg); }
function fill1d(c, clr) { var r=[]; for(var i=0;i<c;i++) r.push(clr); return r; }
function fill2d(r, c, clr) { var res=[]; for(var i=0;i<r;i++) res.push(fill1d(c,clr)); return res; }
