// ═══════════════════════════════════════════════════════
//  PANEL DE PISO — Google Apps Script v7
//  (v6 + Aprovisionamiento automático de CECs)
//  Reemplazá TODO tu código actual con este archivo.
// ═══════════════════════════════════════════════════════

var SS_ID = PropertiesService.getScriptProperties().getProperty('SS_ID') || '';

// ════════════════════════════════════════════════════════════════════
//  APROVISIONAMIENTO — completá estos 2 valores una sola vez
// ════════════════════════════════════════════════════════════════════
// 1) Spreadsheet que se COPIA al crear un CEC nuevo (la plantilla/estructura).
//    Usá tu Spreadsheet actual de Mar del Plata como plantilla.
var GDP_PLANTILLA_SPREADSHEET_ID = '1vH_eWl4h0UeW00BC8IeDagnu_aNLjDbOu_UikRZFejM';

// 2) Carpeta de Drive donde se guardan los Spreadsheets nuevos (opcional).
//    Vacío ('') = se crean en la raíz de tu Drive. Si querés ordenarlos en una
//    carpeta, pegá el ID de la carpeta (lo sacás de la URL de la carpeta en Drive).
var GDP_DRIVE_FOLDER_ID = '';

// Hoja (dentro del Spreadsheet principal) donde se registran los CECs creados.
var GDP_HOJA_CECS = 'CECs';

var SHEET_COLUMNS = {
  Asesor:           1,
  Lider:            2,
  AltaNP_Obj:       3,
  AltaNP_Real:      4,
  AltaPorta_Obj:    5,
  AltaPorta_Real:   6,
  FTTH_Obj:         7,
  FTTH_Real:        8,
  Terminales_Obj:   9,
  Terminales_Real:  10,
  Assurant:         11,
  CPA_Obj:          12,
  CPA_Real:         13,
  Tickets:          14,
  Operaciones:      15,
  NPS:              16,
  TMA:              17,
  Vacaciones:       18
};

// ── CORS ─────────────────────────────────────────────────
function buildResponse(result) {
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── SPREADSHEET ──────────────────────────────────────────
function getOrCreateSpreadsheet() {
  if (SS_ID) {
    try { return SpreadsheetApp.openById(SS_ID); } catch(e) {}
  }
  var files = DriveApp.getFilesByName('Panel de Piso — Registros');
  if (files.hasNext()) {
    var ss = SpreadsheetApp.open(files.next());
    SS_ID = ss.getId();
    PropertiesService.getScriptProperties().setProperty('SS_ID', SS_ID);
    return ss;
  }
  var ss = SpreadsheetApp.create('Panel de Piso — Registros');
  SS_ID = ss.getId();
  PropertiesService.getScriptProperties().setProperty('SS_ID', SS_ID);
  setupEventSheet(ss);
  return ss;
}

function setupEventSheet(ss) {
  var sh = ss.getSheetByName('Eventos') || ss.insertSheet('Eventos');
  if (sh.getLastRow() === 0) {
    sh.appendRow(['Fecha','Asesor','Evento','Hora inicio','Hora fin','Duración','Registrado a las']);
    var header = sh.getRange('1:1');
    header.setFontWeight('bold').setBackground('#c8622a').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    sh.setColumnWidth(1,100); sh.setColumnWidth(2,180); sh.setColumnWidth(3,160);
    sh.setColumnWidth(4,100); sh.setColumnWidth(5,100); sh.setColumnWidth(6,100);
    sh.setColumnWidth(7,140);
    sh.getRange('A:G').setVerticalAlignment('middle');
  }
  ['Hoja 1','Sheet1','Hoja1'].forEach(function(n) {
    try { ss.deleteSheet(ss.getSheetByName(n)); } catch(e) {}
  });
}

// ── GET ───────────────────────────────────────────────────
function doGet(e) {
  var action = (e.parameter && e.parameter.action) ? e.parameter.action : '';
  var result = {};
  try {
    if (action === 'getSheetId') {
      getOrCreateSpreadsheet();
      result = { ok: true, sheetId: SS_ID };

    } else if (action === 'validatePassword') {
      var password = e.parameter.password || '';
      result = validatePassword(password);

    } else if (action === 'getComercialData') {
      var sheetName = e.parameter.sheet || '';
      result = getComercialData(sheetName, e.parameter.cec);

    } else if (action === 'listSheets') {
      result = listSheets();

    // ── Motor de Sincronización ──
    } else if (action === 'getImportConfig') {
      result = getImportConfig();

    } else if (action === 'getIdentidad') {
      result = getIdentidad(e.parameter.cec);

    } else if (action === 'getLideres') {
      result = getLideres(e.parameter.cec);

    } else if (action === 'getConfigPanel') {
      result = getConfigPanel();

    // ── ADMINISTRACIÓN V2 (Centro de Administración) ──
    } else if (action === 'v2_listarCECs') {
      result = v2_listarCECs();

    } else if (action === 'v2_crearCEC') {
      result = v2_crearCEC(e.parameter);

    } else if (action === 'v2_editarCEC') {
      result = v2_editarCEC(e.parameter);

    } else if (action === 'v2_bajaCEC') {
      result = v2_bajaCEC(e.parameter);

    // ── Comercial por CEC (lee la pestaña/mes del Spreadsheet del CEC) ──
    } else if (action === 'v2_comercialCSV') {
      return v2_comercialCSV(e.parameter); // devuelve CSV directo, no JSON

    // ── Aprovisionamiento on-demand: la sincronización garantiza infraestructura ──
    } else if (action === 'v2_asegurarComercialCEC') {
      result = v2_asegurarComercialCEC(e.parameter);

    } else if (action === 'v2_eliminarAsesorComercial') {
      result = v2_eliminarAsesorComercial(e.parameter);

    } else {
      result = { error: 'Acción desconocida: ' + action };
    }
  } catch(err) {
    result = { error: err.message };
  }
  return buildResponse(result);
}

// ── POST ──────────────────────────────────────────────────
function doPost(e) {
  var data = {};
  try { data = JSON.parse(e.postData.contents); } catch(err) {}

  var result = { ok: true };
  try {
    if (data.action === 'logEvent') {
      logEvent(data.date, data.advisor, data.event, data.timeStart, data.timeEnd, data.duration);

    } else if (data.action === 'saveComercial') {
      result = saveComercial(data);

    } else if (data.action === 'newMonth') {
      result = createNewMonth(data);

    // ── Motor de Sincronización ──
    } else if (data.action === 'guardarIdentidad') {
      result = guardarIdentidad(data);

    } else {
      result = { error: 'Acción POST desconocida: ' + data.action };
    }
  } catch(err) {
    result = { error: err.message };
  }
  return buildResponse(result);
}

// ── AUTORIZACIÓN DE OPERACIONES ──
// Un OPERADOR/OWNER autenticado (identidad validada en el frontend contra
// Firestore) está habilitado para operar. Se mantiene Admin_Password como vía
// alternativa para mantenimiento. No se duplica el registro de usuarios: el
// backend confía en la identidad de la sesión ya validada.
function autorizarOperacion_(data){
  // 1) Identidad de sesión: rol OPERADOR u OWNER habilita la operación.
  var rol = String(data.authRol || '').toUpperCase();
  if (rol === 'OWNER' || rol === 'OPERADOR') {
    return { ok:true };
  }
  // 2) Vía alternativa (mantenimiento): Admin_Password global.
  if (data.password) {
    var valid = validatePassword(data.password);
    if (valid.ok && valid.valid) return { ok:true };
  }
  return { ok:false, error:'No autorizado: se requiere un operador autenticado.' };
}

// ── VALIDAR PASSWORD ──────────────────────────────────────
function validatePassword(password) {
  try {
    var ss = getOrCreateSpreadsheet();
    var config = ss.getSheetByName('CONFIG');
    if (!config) return { ok: false, error: 'Hoja CONFIG no encontrada' };

    var data = config.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var key = String(data[i][0]).trim().toLowerCase();
      if (key === 'admin_password') {
        var stored = String(data[i][1]).trim();
        return { ok: true, valid: (stored === String(password).trim()) };
      }
    }
    return { ok: false, error: 'Admin_Password no encontrada en CONFIG' };
  } catch(err) {
    return { ok: false, error: err.message };
  }
}

// ── LEER DATOS COMERCIALES ────────────────────────────────
function getComercialData(sheetName, cec) {
  try {
    var ss = gdpSpreadsheetDeCEC_(cec);
    var sh = ss.getSheetByName(sheetName);
    if (!sh) return { ok: false, error: 'Hoja no encontrada: ' + sheetName };

    var data = sh.getDataRange().getValues();
    if (data.length < 2) return { ok: true, headers: [], rows: [] };

    var headers = data[0].map(String);
    var rows = [];
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      var row = {};
      headers.forEach(function(h, idx) { row[h] = data[i][idx]; });
      rows.push(row);
    }
    return { ok: true, headers: headers, rows: rows };
  } catch(err) {
    return { ok: false, error: err.message };
  }
}

// ── GUARDAR CAMBIOS COMERCIALES ───────────────────────────
// Devuelve el Spreadsheet de un CEC. Mar del Plata (o sin CEC) → el global.
// Otros CECs → su Spreadsheet registrado en la hoja CECs.
function gdpSpreadsheetDeCEC_(cecNombreOId){
  if(!cecNombreOId || cecNombreOId==='*') return getOrCreateSpreadsheet();
  // Puede llegar un ID ya formado ('cec_junin') o un nombre ('Junín'). Si ya
  // empieza con 'cec_', usarlo tal cual; si no, normalizarlo a ID. Sin esto,
  // gdpCecId_ anteponía 'cec_' de nuevo → 'cec_cec_junin' → no matcheaba →
  // caía al fallback y escribía en Mar del Plata.
  var cecId = String(cecNombreOId).indexOf('cec_') === 0
            ? String(cecNombreOId).trim().toLowerCase()
            : gdpCecId_(cecNombreOId);
  if(cecId === 'cec_mar_del_plata') return getOrCreateSpreadsheet();
  var sh = gdpHojaCECs_();
  var data = sh.getDataRange().getValues();
  var headers = data[0];
  var idxId = headers.indexOf('cecId');
  var idxSS = headers.indexOf('spreadsheetId');
  for (var r = 1; r < data.length; r++) {
    if (String(data[r][idxId]) === cecId && data[r][idxSS]) {
      return SpreadsheetApp.openById(data[r][idxSS]);
    }
  }
  // No se encontró Sheet para el CEC: ERROR explícito, NO escribir en MdP.
  throw new Error('No se encontró el Spreadsheet del CEC "' + cecId + '". Verificá que el CEC exista y tenga Sheet aprovisionado.');
}

// Meses del modelo (índice 0 = ENERO). Fuente única para toda la lógica de meses.
var GDP_MESES = ['ENERO','FEBRERO','MARZO','ABRIL','MAYO','JUNIO','JULIO','AGOSTO',
                 'SEPTIEMBRE','OCTUBRE','NOVIEMBRE','DICIEMBRE'];

// Normaliza texto para comparar sin acentos/mayúsculas (ej: 'Líder' == 'Lider').
function gdpNorm_(s){
  return String(s==null?'':s).trim().toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g,'');
}

// Busca el índice de una columna por nombre, tolerando acentos (Líder/Lider) y
// mayúsculas. Sin esto, el volcado del líder podía perderse por 'Líder'≠'Lider'.
function gdpFindColIdx_(headers, name){
  var i = headers.indexOf(name);
  if(i !== -1) return i;
  var target = gdpNorm_(name);
  for(var k=0;k<headers.length;k++){ if(gdpNorm_(headers[k]) === target) return k; }
  return -1;
}

// Dado 'JULIO_2026' devuelve 'JUNIO_2026'; 'ENERO_2026' → 'DICIEMBRE_2025'.
// Devuelve null si el nombre no tiene formato MES_AAAA reconocible.
function gdpMesPrevioTab_(mesTab){
  var parts = String(mesTab||'').trim().toUpperCase().split('_');
  var idx = GDP_MESES.indexOf(parts[0]);
  var anio = parseInt(parts[1], 10);
  if(idx === -1 || isNaN(anio)) return null;
  var prevIdx = idx - 1, prevAnio = anio;
  if(prevIdx < 0){ prevIdx = 11; prevAnio = anio - 1; }
  return GDP_MESES[prevIdx] + '_' + prevAnio;
}

// Devuelve la pestaña de mes MÁS RECIENTE que exista (por año+mes, no por orden de
// pestañas). Se usa como plantilla de respaldo cuando el mes anterior exacto no está.
function gdpMesTabMasReciente_(ss){
  var hojas = ss.getSheets();
  var mejor = null, mejorKey = -1;
  for(var i=0;i<hojas.length;i++){
    var partes = String(hojas[i].getName()).trim().toUpperCase().split('_');
    var mi = GDP_MESES.indexOf(partes[0]);
    var an = parseInt(partes[1], 10);
    if(mi === -1 || isNaN(an)) continue;
    var key = an * 100 + mi;
    if(key > mejorKey){ mejorKey = key; mejor = hojas[i]; }
  }
  return mejor;
}

// Asegura que exista la pestaña del mes en un Spreadsheet. Si no existe, la crea
// heredando la estructura (encabezados + asesores) del MES INMEDIATAMENTE ANTERIOR
// del MISMO CEC (JULIO_2026 ← JUNIO_2026), con los valores reales vacíos Y el líder
// heredado limpio (los líderes se resuelven contra la config actual del CEC, no por
// arrastre de un mes viejo). Si el mes anterior exacto no está, usa el mes más
// reciente que exista. Si no hay ninguno, crea con los encabezados estándar.
// Se usa al sincronizar: al inicio de cada mes la pestaña no existe todavía.
function gdpAsegurarPestanaMes_(ss, mesTab){
  var sh = ss.getSheetByName(mesTab);
  if(sh) return sh;

  // PRIORIDAD 1: el mes inmediatamente anterior del mismo CEC. PRIORIDAD 2: el mes
  // más reciente que exista. Sin esto se tomaba la 1ª pestaña de mes en orden de
  // hojas (arbitraria) → arrastraba una "foto vieja" del equipo y líderes obsoletos.
  var prevTab = gdpMesPrevioTab_(mesTab);
  var plantilla = (prevTab ? ss.getSheetByName(prevTab) : null) || gdpMesTabMasReciente_(ss);

  var realCols = ['AltaNP_Real','AltaPorta_Real','FTTH_Real','Terminales_Real',
                  'Assurant','CPA_Real','Tickets','Operaciones','NPS','TMA','Vacaciones'];

  if(plantilla){
    // Copiar estructura (encabezados + filas de asesores), limpiar reales Y líder.
    var nueva = plantilla.copyTo(ss);
    nueva.setName(mesTab);
    var headers = nueva.getRange(1, 1, 1, nueva.getLastColumn()).getValues()[0];
    var lastRow = nueva.getLastRow();
    if(lastRow > 1){
      headers.forEach(function(h, idx){
        var limpiar = realCols.indexOf(String(h).trim()) !== -1
                   || gdpNorm_(h) === 'lider';   // limpiar líder heredado (obsoleto)
        if(limpiar){
          nueva.getRange(2, idx + 1, lastRow - 1, 1).clearContent();
        }
      });
    }
    SpreadsheetApp.flush();
    return nueva;
  }

  // No hay ninguna pestaña de mes: crear con encabezados estándar del modelo.
  var creada = ss.insertSheet(mesTab);
  creada.appendRow(['Asesor','Usuario','Lider','AltaNP_Real','AltaPorta_Real','FTTH_Real',
                    'Terminales_Real','Assurant','CPA_Real','Tickets','Operaciones',
                    'NPS','TMA','Vacaciones',
                    'AltaNP_Obj','AltaPorta_Obj','FTTH_Obj','Terminales_Obj','CPA_Obj']);
  creada.getRange('1:1').setFontWeight('bold');
  SpreadsheetApp.flush();
  return creada;
}

function saveComercial(data) {
  var sheetName  = data.sheet;
  var changes    = data.changes;
  var objectives = data.objectives;
  // Autorización por identidad de sesión (o Admin_Password para mantenimiento).
  var auth = autorizarOperacion_(data);
  if (!auth.ok) return { ok: false, error: auth.error };
  try {
    // Escribir en el Spreadsheet del CEC autenticado (Junín en su Sheet, MdP en
    // el global). Sin esto, el comercial de todos iba al Sheet de Mar del Plata.
    var ss = gdpSpreadsheetDeCEC_(data.authCEC || data.cec);
    // Crear la pestaña del mes si no existe (al inicio de cada mes no está todavía).
    // El volcado la crea con la estructura heredada y sigue normalmente.
    var sh = gdpAsegurarPestanaMes_(ss, sheetName);
    if (!sh) return { ok: false, error: 'No se pudo crear/abrir la hoja: ' + sheetName };

    var data = sh.getDataRange().getValues();
    var headers = data[0].map(function(h){ return String(h).trim(); });
    var saved = 0;
    var filasCreadas = 0;

    if (changes && changes.length > 0) {
      var asesorIdx = headers.indexOf('Asesor');
      if (asesorIdx === -1) return { ok: false, error: 'Columna Asesor no encontrada' };

      // Columna Usuario: el USERNAME manda (un usuario = una fila). Si el Sheet aún no
      // la tiene, se agrega (migración) para que las escrituras keyeen por usuario y una
      // variante del nombre no genere una segunda fila.
      var usuarioIdx = gdpFindColIdx_(headers, 'Usuario');
      if (usuarioIdx === -1) {
        usuarioIdx = headers.length;
        sh.getRange(1, usuarioIdx + 1).setValue('Usuario');
        headers.push('Usuario');
        for (var rr = 0; rr < data.length; rr++) data[rr].push('');
      }

      // Agrupar por USUARIO (si viene) o por nombre. Un grupo por identidad real.
      var grupos = {};
      changes.forEach(function(change){
        var u = change.usuario ? String(change.usuario).trim() : '';
        var key = u ? ('u:' + u.toLowerCase()) : ('n:' + String(change.asesor).trim().toLowerCase());
        if(!grupos[key]) grupos[key] = { asesor:String(change.asesor).trim(), usuario:u, cols:{} };
        grupos[key].cols[change.col] = change.value;
      });

      Object.keys(grupos).forEach(function(k){
        var g = grupos[k];
        var filaIdx = -1;
        // 1) match por USUARIO (identidad real) — el user manda.
        if (g.usuario) {
          for (var r = 1; r < data.length; r++) {
            if (String(data[r][usuarioIdx]).trim().toLowerCase() === g.usuario.toLowerCase()) { filaIdx = r; break; }
          }
        }
        // 2) fallback por nombre (Sheets viejos sin usuario) → se backfillea el usuario.
        if (filaIdx === -1) {
          for (var r = 1; r < data.length; r++) {
            if (String(data[r][asesorIdx]).trim().toLowerCase() === g.asesor.toLowerCase()) { filaIdx = r; break; }
          }
        }
        // 3) si no existe, crear la fila.
        if (filaIdx === -1) {
          var nuevaFila = [];
          for (var c = 0; c < headers.length; c++) nuevaFila.push('');
          nuevaFila[asesorIdx] = g.asesor;
          if (g.usuario) nuevaFila[usuarioIdx] = g.usuario;
          sh.appendRow(nuevaFila);
          data.push(nuevaFila);
          filaIdx = data.length - 1;
          filasCreadas++;
        } else {
          // Fila existente: fijar el nombre CANÓNICO + el usuario (el user manda sobre el
          // nombre, así se corrige una variante en vez de duplicar).
          sh.getRange(filaIdx + 1, asesorIdx + 1).setValue(g.asesor);
          data[filaIdx][asesorIdx] = g.asesor;
          if (g.usuario) { sh.getRange(filaIdx + 1, usuarioIdx + 1).setValue(g.usuario); data[filaIdx][usuarioIdx] = g.usuario; }
        }
        Object.keys(g.cols).forEach(function(col){
          var colIdx = gdpFindColIdx_(headers, col);
          if (colIdx === -1) return;
          sh.getRange(filaIdx + 1, colIdx + 1).setValue(g.cols[col]);
          saved++;
        });
      });
    }

    if (objectives && Object.keys(objectives).length > 0) {
      Object.keys(objectives).forEach(function(col) {
        var colIdx = gdpFindColIdx_(headers, col);
        if (colIdx === -1) return;
        for (var r = 1; r < data.length; r++) {
          if (data[r][0]) {
            sh.getRange(r + 1, colIdx + 1).setValue(objectives[col]);
            saved++;
          }
        }
      });
    }

    // Forzar formato NUMÉRICO (entero) en la columna Vacaciones: son unidades
    // (semanas), no fechas. Si quedó con formato fecha, el CSV publicado exporta un
    // valor de fecha en lugar del número y se lee mal. Esto lo auto-corrige en cada
    // guardado. Solo Vacaciones (no toca columnas con decimales como TMA/NPS).
    try {
      var vacIdx = gdpFindColIdx_(headers, 'Vacaciones');
      var lastR = sh.getLastRow();
      if (vacIdx !== -1 && lastR > 1) {
        sh.getRange(2, vacIdx + 1, lastR - 1, 1).setNumberFormat('0');
      }
    } catch(e) {}

    SpreadsheetApp.flush();
    return { ok: true, saved: saved, filasCreadas: filasCreadas };
  } catch(err) {
    return { ok: false, error: err.message };
  }
}

// ── CREAR NUEVO MES ───────────────────────────────────────
function createNewMonth(data) {
  var fromSheet = data.fromSheet;
  var toSheet   = data.toSheet;
  var auth = autorizarOperacion_(data);
  if (!auth.ok) return { ok: false, error: auth.error };
  try {
    var ss = gdpSpreadsheetDeCEC_(data.authCEC || data.cec);
    if (ss.getSheetByName(toSheet)) {
      return { ok: false, error: 'La hoja ' + toSheet + ' ya existe' };
    }
    var source = ss.getSheetByName(fromSheet);
    if (!source) return { ok: false, error: 'Hoja origen no encontrada: ' + fromSheet };

    var newSheet = source.copyTo(ss);
    newSheet.setName(toSheet);

    var headers = newSheet.getRange(1, 1, 1, newSheet.getLastColumn()).getValues()[0];
    var realCols = ['AltaNP_Real','AltaPorta_Real','FTTH_Real','Terminales_Real',
                    'Assurant','CPA_Real','Tickets','Operaciones','NPS','TMA','Vacaciones'];
    var lastRow = newSheet.getLastRow();
    headers.forEach(function(h, idx) {
      if (realCols.indexOf(String(h).trim()) !== -1 && lastRow > 1) {
        newSheet.getRange(2, idx + 1, lastRow - 1, 1).clearContent();
      }
    });

    SpreadsheetApp.flush();
    return { ok: true, created: toSheet };
  } catch(err) {
    return { ok: false, error: err.message };
  }
}

// ── LISTAR HOJAS DISPONIBLES ──────────────────────────────
function listSheets() {
  try {
    var ss = getOrCreateSpreadsheet();
    var sheets = ss.getSheets()
      .map(function(s){ return s.getName(); })
      .filter(function(n){ return /^[A-Z]+_\d{4}$/.test(n); })
      .sort();
    return { ok: true, sheets: sheets };
  } catch(err) {
    return { ok: false, error: err.message };
  }
}

// ── LOG DE EVENTOS ────────────────────────────────────────
function logEvent(date, advisor, event, timeStart, timeEnd, duration) {
  var ss = getOrCreateSpreadsheet();
  var sh = ss.getSheetByName('Eventos');
  if (!sh) { setupEventSheet(ss); sh = ss.getSheetByName('Eventos'); }

  var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'HH:mm:ss');
  sh.appendRow([date, advisor, event, timeStart||'', timeEnd||'', duration||'', now]);

  var row = sh.getLastRow();
  var color = '#ffffff';
  if (event && event.indexOf('almuerzo') !== -1) color = '#fdf4de';
  if (event && event.indexOf('café')    !== -1) color = '#e8f0fa';
  if (event && event.indexOf('Ausente') !== -1) color = '#fdeaea';
  if (event && event.indexOf('Volvió')  !== -1) color = '#eaf4ec';
  sh.getRange(row, 1, 1, 7).setBackground(color);
}

// ═══════════════════════════════════════════════════════════════
//  MOTOR DE SINCRONIZACIÓN — funciones de configuración
// ═══════════════════════════════════════════════════════════════

// ── Lee la hoja IMPORT_CONFIG y la devuelve como JSON ───────────
function getImportConfig() {
  try {
    var ss = getOrCreateSpreadsheet();
    var sh = ss.getSheetByName('IMPORT_CONFIG');
    if (!sh) return { ok: false, error: 'Hoja IMPORT_CONFIG no encontrada' };

    var data = sh.getDataRange().getValues();
    if (data.length < 2) return { ok: true, reglas: [] };

    var headers = data[0].map(function(h){ return String(h).trim().toUpperCase(); });
    var iCampo = headers.indexOf('CAMPO_DESTINO');
    var iKw    = headers.indexOf('FUENTE_KEYWORDS');
    var iOp    = headers.indexOf('OPERACION');
    var iAct   = headers.indexOf('ACTIVO');

    if (iCampo === -1 || iKw === -1) {
      return { ok: false, error: 'IMPORT_CONFIG debe tener columnas CAMPO_DESTINO y FUENTE_KEYWORDS' };
    }

    var reglas = [];
    for (var r = 1; r < data.length; r++) {
      if (iAct !== -1) {
        var activo = String(data[r][iAct]).trim().toUpperCase();
        if (activo === 'FALSE' || activo === 'NO' || activo === '0') continue;
      }
      var campoRaw = String(data[r][iCampo]).trim();
      var kwRaw    = String(data[r][iKw]).trim();
      if (!kwRaw) continue;

      var esIgnorar = (campoRaw === '' || campoRaw.toLowerCase() === '(ignorar)' || campoRaw.toLowerCase() === 'ignorar');
      var keywords = kwRaw.split('|').map(function(k){ return k.trim(); }).filter(String);

      reglas.push({
        campo:     esIgnorar ? null : campoRaw,
        keywords:  keywords,
        operacion: iOp !== -1 ? (String(data[r][iOp]).trim().toLowerCase() || 'suma') : 'suma'
      });
    }
    return { ok: true, reglas: reglas };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

// ── IDENTIDAD: lee la hoja IDENTIDAD → { username: {nombre, lider_id, estado} } ──
function getIdentidad(cec) {
  try {
    var ss = gdpSpreadsheetDeCEC_(cec);
    var sh = ss.getSheetByName('IDENTIDAD');
    if (!sh) {
      sh = ss.insertSheet('IDENTIDAD');
      sh.appendRow(['USERNAME','NOMBRE','LIDER_ID','ESTADO']);
      sh.getRange('1:1').setFontWeight('bold');
      return { ok: true, identidad: {} };
    }
    var data = sh.getDataRange().getValues();
    if (data.length < 2) return { ok: true, identidad: {} };

    var headers = data[0].map(function(h){ return String(h).trim().toUpperCase(); });
    var iU = headers.indexOf('USERNAME');
    var iN = headers.indexOf('NOMBRE');
    var iL = headers.indexOf('LIDER_ID');
    var iE = headers.indexOf('ESTADO');
    if (iU === -1) return { ok:false, error:'IDENTIDAD debe tener columna USERNAME' };

    var identidad = {};
    for (var r = 1; r < data.length; r++) {
      var u = String(data[r][iU]).trim().toLowerCase().replace(/\s+/g,'');
      if (!u) continue;
      identidad[u] = {
        nombre:   iN !== -1 ? String(data[r][iN]).trim() : '',
        lider_id: iL !== -1 ? String(data[r][iL]).trim() : '',
        estado:   iE !== -1 ? (String(data[r][iE]).trim().toLowerCase() || 'incorporado') : 'incorporado'
      };
    }
    return { ok: true, identidad: identidad };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

// ── LIDERES: lee la hoja LIDERES → [{id, nombre}] ──
function getLideres(cec) {
  try {
    var ss = gdpSpreadsheetDeCEC_(cec);
    var sh = ss.getSheetByName('LIDERES');
    if (!sh) {
      sh = ss.insertSheet('LIDERES');
      sh.appendRow(['ID','NOMBRE']);
      sh.getRange('1:1').setFontWeight('bold');
      return { ok: true, lideres: [] };
    }
    var data = sh.getDataRange().getValues();
    if (data.length < 2) return { ok: true, lideres: [] };
    var headers = data[0].map(function(h){ return String(h).trim().toUpperCase(); });
    var iI = headers.indexOf('ID');
    var iN = headers.indexOf('NOMBRE');
    var lideres = [];
    for (var r = 1; r < data.length; r++) {
      var id = String(data[r][iI]).trim();
      if (!id) continue;
      lideres.push({ id: id, nombre: iN !== -1 ? String(data[r][iN]).trim() : id });
    }
    return { ok: true, lideres: lideres };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

// ── CONFIG_PANEL: lee config general (CEC, etc.) → { CEC: "..." } ──
function getConfigPanel() {
  try {
    var ss = getOrCreateSpreadsheet();
    var sh = ss.getSheetByName('CONFIG_PANEL');
    if (!sh) {
      sh = ss.insertSheet('CONFIG_PANEL');
      sh.appendRow(['KEY','VALUE']);
      sh.appendRow(['CEC','Mar del Plata']);
      sh.getRange('1:1').setFontWeight('bold');
    }
    var data = sh.getDataRange().getValues();
    var config = {};
    for (var r = 1; r < data.length; r++) {
      var k = String(data[r][0]).trim();
      if (k) config[k] = String(data[r][1]).trim();
    }
    return { ok: true, config: config };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

// ── GUARDAR IDENTIDAD: incorpora/ignora usuarios nuevos (persistente) ──
function guardarIdentidad(data) {
  var incorporar = data.incorporar;
  var ignorar    = data.ignorar;
  var cec        = data.cec;
  var auth = autorizarOperacion_(data);
  if (!auth.ok) return { ok: false, error: auth.error };
  try {
    // Escribir la IDENTIDAD en el Spreadsheet del CEC (Junín en su Sheet, MdP en
    // el global). Antes usaba SIEMPRE getOrCreateSpreadsheet() (Mar del Plata),
    // por eso los asesores de otros CECs se escribían en la hoja de MdP.
    var ss = gdpSpreadsheetDeCEC_(data.authCEC || cec);
    var sh = ss.getSheetByName('IDENTIDAD');
    if (!sh) {
      sh = ss.insertSheet('IDENTIDAD');
      sh.appendRow(['USERNAME','NOMBRE','LIDER_ID','ESTADO']);
      sh.getRange('1:1').setFontWeight('bold');
    }
    var data = sh.getDataRange().getValues();
    var headers = data[0].map(function(h){ return String(h).trim().toUpperCase(); });
    var iU = headers.indexOf('USERNAME');
    var iN = headers.indexOf('NOMBRE');
    var iL = headers.indexOf('LIDER_ID');
    var iE = headers.indexOf('ESTADO');

    function buscarFila(username){
      var un = String(username).trim().toLowerCase().replace(/\s+/g,'');
      for (var r = 1; r < data.length; r++) {
        if (String(data[r][iU]).trim().toLowerCase().replace(/\s+/g,'') === un) return r + 1;
      }
      return -1;
    }

    function upsert(username, nombre, lider_id, estado){
      var fila = buscarFila(username);
      if (fila > 0) {
        sh.getRange(fila, iN + 1).setValue(nombre || '');
        if (iL !== -1) sh.getRange(fila, iL + 1).setValue(lider_id || '');
        if (iE !== -1) sh.getRange(fila, iE + 1).setValue(estado);
      } else {
        var nueva = [];
        nueva[iU] = username;
        nueva[iN] = nombre || '';
        if (iL !== -1) nueva[iL] = lider_id || '';
        if (iE !== -1) nueva[iE] = estado;
        for (var c = 0; c < headers.length; c++) if (nueva[c] === undefined) nueva[c] = '';
        sh.appendRow(nueva);
        data.push(nueva);
      }
    }

    var n = 0;
    (incorporar || []).forEach(function(u){ upsert(u.username, u.nombre, u.lider_id, 'incorporado'); n++; });
    (ignorar || []).forEach(function(u){ upsert(u.username, '', '', 'ignorado'); n++; });

    SpreadsheetApp.flush();
    return { ok: true, guardados: n };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

// ═══════════════════════════════════════════════════════════════
//  ADMINISTRACIÓN V2 — APROVISIONAMIENTO DE CECs
//  Crear un CEC = acción de negocio. El backend crea toda la
//  infraestructura. El OWNER nunca toca Sheets ni IDs.
// ═══════════════════════════════════════════════════════════════

// Normaliza un nombre de CEC a id canónico: "Junín" -> "cec_junin"
function gdpCecId_(nombre){
  var s = String(nombre || '').trim().toLowerCase();
  s = s.normalize('NFD').replace(/[\u0300-\u036f]/g, ''); // quita tildes
  s = s.replace(/\s+/g, '_');
  return 'cec_' + s;
}

// Hoja de CECs (en el Spreadsheet principal). La crea si no existe.
function gdpHojaCECs_(){
  var ss = getOrCreateSpreadsheet();
  var sh = ss.getSheetByName(GDP_HOJA_CECS);
  if(!sh){
    sh = ss.insertSheet(GDP_HOJA_CECS);
    sh.appendRow(['cecId','nombre','activo','spreadsheetId','spreadsheetUrl','codigoBodega',
                  'pais','provincia','ciudad','direccion','lider','supervisor','lat','lng','latManual','lngManual',
                  'geocoder','precision','creadoEl']);
    sh.getRange('1:1').setFontWeight('bold');
    // SEMBRAR Mar del Plata (CEC original) automáticamente. Su Spreadsheet es el
    // histórico (el mismo SS principal). Así el CEC original SIEMPRE existe sin
    // recrearlo, y su comercial se lee de las hojas de siempre. Sin esto, la
    // primera carga con el backend nuevo deja la lista vacía y rompe el landing.
    var ssId = ss.getId();
    sh.appendRow([
      'cec_mar_del_plata', 'Mar del Plata', true,
      ssId, 'https://docs.google.com/spreadsheets/d/' + ssId + '/edit',
      '', 'Argentina', 'Buenos Aires', 'Mar del Plata', 'Av. Colón 2550',
      '', '', -38.00377226264887, -57.54984229189988, '', '', '', '',
      new Date().toISOString()
    ]);
  }
  return sh;
}

// ── v2_crearCEC — APROVISIONAMIENTO COMPLETO ──
function v2_crearCEC(params){
  // Seguridad: solo con password de admin válida.
  var valid = validatePassword(params.password || '');
  if (!valid.ok || !valid.valid) return { ok:false, error:'Password incorrecta' };

  try{
    var nombre = String(params.nombre || '').trim();
    if(!nombre) return { ok:false, error:'El nombre del CEC es obligatorio.' };

    var cecId = gdpCecId_(nombre);
    var sh = gdpHojaCECs_();
    var data = sh.getDataRange().getValues();
    var headers = data[0];
    var idxId = headers.indexOf('cecId');

    // ¿Ya existe?
    for (var r = 1; r < data.length; r++) {
      if (String(data[r][idxId]) === cecId) return { ok:false, error:'Ya existe un CEC con ese nombre.' };
    }

    // 1) Crear el Spreadsheet comercial copiando la plantilla.
    var plantilla = DriveApp.getFileById(GDP_PLANTILLA_SPREADSHEET_ID);
    var nombreSS = 'GDP - ' + nombre;
    var copia = plantilla.makeCopy(nombreSS);

    // Mover a la carpeta destino si se configuró una.
    if (GDP_DRIVE_FOLDER_ID) {
      try {
        var folder = DriveApp.getFolderById(GDP_DRIVE_FOLDER_ID);
        folder.addFile(copia);
        DriveApp.getRootFolder().removeFile(copia);
      } catch(e) { /* si falla el move, queda en la raíz: no es crítico */ }
    }

    var nuevoSSId  = copia.getId();
    var nuevoSSUrl = 'https://docs.google.com/spreadsheets/d/' + nuevoSSId + '/edit';

    // 1.b) VACIAR el CEC nuevo: conserva la ESTRUCTURA (encabezados, config) pero
    //      borra los datos heredados de la plantilla. Junín nace vacío y se llena
    //      al sincronizar el comercial o por carga manual desde Administración.
    gdpVaciarSpreadsheetCEC_(nuevoSSId);

    // 2) Registrar el CEC con su spreadsheetId.
    var valores = {
      cecId: cecId, nombre: nombre,
      activo: (String(params.activo) === 'false') ? false : true,
      spreadsheetId: nuevoSSId, spreadsheetUrl: nuevoSSUrl,
      codigoBodega: params.codigoBodega || '',
      pais: params.pais || 'Argentina', provincia: params.provincia || '',
      ciudad: params.ciudad || '', direccion: params.direccion || '',
      lider: params.lider || '', supervisor: params.supervisor || '',
      lat: params.lat || '', lng: params.lng || '',
      latManual: params.latManual || '', lngManual: params.lngManual || '',
      geocoder: params.geocoder || '', precision: params.precision || '',
      creadoEl: new Date().toISOString()
    };
    var fila = [];
    headers.forEach(function(h){ fila.push(valores[h] !== undefined ? valores[h] : ''); });
    sh.appendRow(fila);

    // 3) Auditoría.
    gdpAuditar_(params.username || 'desconocido', 'crearCEC', nombre + ' (' + nombreSS + ')', cecId);

    // 4) Éxito con datos del Spreadsheet creado.
    return { ok:true, cecId:cecId, nombre:nombre,
             spreadsheet:{ id:nuevoSSId, nombre:nombreSS, url:nuevoSSUrl } };
  }catch(err){
    return { ok:false, error:'Error al aprovisionar el CEC: ' + err.message };
  }
}

// ── Vaciar un CEC nuevo: conserva estructura, borra datos heredados ──
// REGLA:
//  - Hojas de CONFIG (estructura que el CEC necesita): NO se tocan.
//      CONFIG, IMPORT_CONFIG, CONFIG_PANEL
//  - Hojas de DATOS (asesores/comercial/eventos): se borran los datos pero se
//      conservan los encabezados (fila 1). Quedan vacías y listas para llenarse.
//      Meses comerciales (MES_AÑO), IDENTIDAD, LIDERES, Eventos, Auditoria, CECs
function gdpVaciarSpreadsheetCEC_(ssId){
  try{
    var ss = SpreadsheetApp.openById(ssId);

    // Hojas que NO se tocan (son la estructura/config que el CEC necesita).
    var preservar = { 'CONFIG':true, 'IMPORT_CONFIG':true, 'CONFIG_PANEL':true };

    ss.getSheets().forEach(function(sh){
      var nombre = sh.getName();

      // La config se conserva intacta.
      if (preservar[nombre]) {
        // Excepción: en CONFIG_PANEL, el CEC nuevo NO debe heredar "Mar del Plata".
        // Se limpia el valor de CEC para que tome su identidad propia.
        if (nombre === 'CONFIG_PANEL') {
          var d = sh.getDataRange().getValues();
          for (var r = 1; r < d.length; r++) {
            if (String(d[r][0]).trim().toUpperCase() === 'CEC') {
              sh.getRange(r + 1, 2).clearContent(); // se setea luego con el nombre real
            }
          }
        }
        return;
      }

      // Hojas de datos: borrar todo menos la fila de encabezados.
      var lastRow = sh.getLastRow();
      var lastCol = sh.getLastColumn();
      if (lastRow > 1 && lastCol > 0) {
        sh.getRange(2, 1, lastRow - 1, lastCol).clearContent();
      }
    });

    SpreadsheetApp.flush();
  }catch(e){ /* si algo falla, el CEC igual queda creado; no es crítico */ }
}

// ── v2_editarCEC ──
function v2_editarCEC(params){
  var valid = validatePassword(params.password || '');
  if (!valid.ok || !valid.valid) return { ok:false, error:'Password incorrecta' };

  try{
    var cecId = params.id ? String(params.id) : gdpCecId_(params.nombre);
    var sh = gdpHojaCECs_();
    var data = sh.getDataRange().getValues();
    var headers = data[0];
    var idxId = headers.indexOf('cecId');

    var filaSheet = -1;
    for (var r = 1; r < data.length; r++) {
      if (String(data[r][idxId]) === cecId) { filaSheet = r + 1; break; }
    }
    if (filaSheet === -1) return { ok:false, error:'CEC no encontrado.' };

    function setCol(col, val){
      var c = headers.indexOf(col);
      if (c !== -1 && val !== undefined && val !== '') sh.getRange(filaSheet, c+1).setValue(val);
    }
    setCol('nombre', params.nombre);
    if (params.activo !== undefined) setCol('activo', String(params.activo) === 'false' ? false : true);
    setCol('codigoBodega', params.codigoBodega);
    setCol('pais', params.pais); setCol('provincia', params.provincia);
    setCol('ciudad', params.ciudad); setCol('direccion', params.direccion);
    if (params.lider !== undefined) setCol('lider', params.lider);
    if (params.supervisor !== undefined) setCol('supervisor', params.supervisor);
    setCol('lat', params.lat); setCol('lng', params.lng);
    setCol('latManual', params.latManual); setCol('lngManual', params.lngManual);
    setCol('geocoder', params.geocoder); setCol('precision', params.precision);

    gdpAuditar_(params.username || 'desconocido', 'editarCEC', params.nombre || cecId, cecId);
    return { ok:true, cecId:cecId };
  }catch(err){
    return { ok:false, error:err.message };
  }
}

// ── v2_bajaCEC (activar/desactivar) ──
function v2_bajaCEC(params){
  var valid = validatePassword(params.password || '');
  if (!valid.ok || !valid.valid) return { ok:false, error:'Password incorrecta' };

  try{
    var cecId = String(params.id || '');
    var sh = gdpHojaCECs_();
    var data = sh.getDataRange().getValues();
    var headers = data[0];
    var idxId = headers.indexOf('cecId');
    var idxAct = headers.indexOf('activo');

    for (var r = 1; r < data.length; r++) {
      if (String(data[r][idxId]) === cecId) {
        sh.getRange(r + 1, idxAct + 1).setValue(params.activo === 'true' || params.activo === true);
        gdpAuditar_(params.username || 'desconocido', 'bajaCEC', cecId, cecId);
        return { ok:true };
      }
    }
    return { ok:false, error:'CEC no encontrado.' };
  }catch(err){
    return { ok:false, error:err.message };
  }
}

// ── v2_listarCECs (incluye spreadsheetId) ──
function v2_listarCECs(){
  try{
    var sh = gdpHojaCECs_();
    var data = sh.getDataRange().getValues();
    var headers = data[0];
    var cecs = [];
    for (var r = 1; r < data.length; r++) {
      var o = {};
      headers.forEach(function(h, i){ o[h] = data[r][i]; });
      if (!o.cecId) continue;
      o.activo = (o.activo === true || String(o.activo).toLowerCase() === 'true');
      o.ubicacion = {
        pais:o.pais, provincia:o.provincia, ciudad:o.ciudad, direccion:o.direccion,
        lat:parseFloat(o.lat)||null, lng:parseFloat(o.lng)||null,
        latManual:(o.latManual===''?null:parseFloat(o.latManual)),
        lngManual:(o.lngManual===''?null:parseFloat(o.lngManual))
      };
      o.id = o.cecId;
      cecs.push(o);
    }
    return { ok:true, cecs:cecs };
  }catch(err){
    return { ok:false, error:err.message, cecs:[] };
  }
}

// ── v2_comercialCSV — comercial de un CEC desde SU Spreadsheet ──
function v2_comercialCSV(params){
  try{
    var cecNombre = String(params.cec || '').trim();
    var mesTab    = String(params.mes || '').trim();
    if(!cecNombre) return ContentService.createTextOutput('').setMimeType(ContentService.MimeType.CSV);

    var cecId = gdpCecId_(cecNombre);
    var sh = gdpHojaCECs_();
    var data = sh.getDataRange().getValues();
    var headers = data[0];
    var idxId = headers.indexOf('cecId');
    var idxSS = headers.indexOf('spreadsheetId');

    var ssId = null;
    for (var r = 1; r < data.length; r++) {
      if (String(data[r][idxId]) === cecId) { ssId = data[r][idxSS]; break; }
    }
    if (!ssId) return ContentService.createTextOutput('').setMimeType(ContentService.MimeType.CSV);

    var ss = SpreadsheetApp.openById(ssId);
    var hoja = ss.getSheetByName(mesTab) || ss.getSheets()[0];
    if (!hoja) return ContentService.createTextOutput('').setMimeType(ContentService.MimeType.CSV);

    var datos = hoja.getDataRange().getValues();
    var csv = datos.map(function(fila){
      return fila.map(function(celda){
        var s = String(celda).replace(/"/g, '""');
        return /[",\n]/.test(s) ? '"' + s + '"' : s;
      }).join(',');
    }).join('\n');

    return ContentService.createTextOutput(csv).setMimeType(ContentService.MimeType.CSV);
  }catch(err){
    return ContentService.createTextOutput('').setMimeType(ContentService.MimeType.CSV);
  }
}

// ── Aprovisiona el Spreadsheet de un CEC (crear copia + vaciar). Reutilizable ──
// Devuelve { id, nombre, url }. Lo usan v2_crearCEC y v2_asegurarComercialCEC.
function gdpAprovisionarSpreadsheet_(nombreCEC){
  var plantilla = DriveApp.getFileById(GDP_PLANTILLA_SPREADSHEET_ID);
  var nombreSS = 'GDP - ' + nombreCEC;
  var copia = plantilla.makeCopy(nombreSS);
  if (GDP_DRIVE_FOLDER_ID) {
    try {
      var folder = DriveApp.getFolderById(GDP_DRIVE_FOLDER_ID);
      folder.addFile(copia);
      DriveApp.getRootFolder().removeFile(copia);
    } catch(e) {}
  }
  var id = copia.getId();
  gdpVaciarSpreadsheetCEC_(id); // nace vacío: estructura sí, datos no
  return { id:id, nombre:nombreSS, url:'https://docs.google.com/spreadsheets/d/' + id + '/edit' };
}

// ── v2_asegurarComercialCEC — APROVISIONAMIENTO ON-DEMAND ──
// Lo dispara la sincronización. Verifica si el CEC ya tiene Spreadsheet:
//   - Si NO: lo crea automáticamente (copia plantilla, vacía) y guarda la relación.
//   - Si SÍ: lo usa.
// El operador nunca interviene. Autoriza por identidad de sesión.
function v2_asegurarComercialCEC(params){
  var auth = autorizarOperacion_(params);
  if (!auth.ok) return { ok:false, error:auth.error };

  try{
    var nombre = String(params.cec || '').trim();
    if(!nombre) return { ok:false, error:'Falta el nombre del CEC.' };

    var cecId = gdpCecId_(nombre);
    var sh = gdpHojaCECs_();
    var data = sh.getDataRange().getValues();
    var headers = data[0];
    var idxId = headers.indexOf('cecId');
    var idxSS = headers.indexOf('spreadsheetId');
    var idxUrl = headers.indexOf('spreadsheetUrl');
    var idxNom = headers.indexOf('nombre');

    // Buscar el CEC.
    var filaSheet = -1, ssId = null;
    for (var r = 1; r < data.length; r++) {
      if (String(data[r][idxId]) === cecId) { filaSheet = r + 1; ssId = data[r][idxSS]; break; }
    }

    // El CEC no está registrado → registrarlo y aprovisionar.
    if (filaSheet === -1) {
      var prov = gdpAprovisionarSpreadsheet_(nombre);
      var valores = {
        cecId:cecId, nombre:nombre, activo:true,
        spreadsheetId:prov.id, spreadsheetUrl:prov.url,
        codigoBodega:'', pais:'Argentina', provincia:'', ciudad:'', direccion:'',
        lat:'', lng:'', latManual:'', lngManual:'', geocoder:'', precision:'',
        creadoEl:new Date().toISOString()
      };
      var fila = [];
      headers.forEach(function(h){ fila.push(valores[h] !== undefined ? valores[h] : ''); });
      sh.appendRow(fila);
      gdpAuditar_(params.authUser || 'sync', 'aprovisionarComercial', nombre, cecId);
      return { ok:true, creado:true, spreadsheet:prov };
    }

    // El CEC existe pero sin Spreadsheet → aprovisionar y guardar.
    if (!ssId) {
      var prov2 = gdpAprovisionarSpreadsheet_(nombre);
      sh.getRange(filaSheet, idxSS + 1).setValue(prov2.id);
      if (idxUrl !== -1) sh.getRange(filaSheet, idxUrl + 1).setValue(prov2.url);
      gdpAuditar_(params.authUser || 'sync', 'aprovisionarComercial', nombre, cecId);
      return { ok:true, creado:true, spreadsheet:prov2 };
    }

    // Ya tiene infraestructura.
    return { ok:true, creado:false,
             spreadsheet:{ id:ssId, url:(idxUrl!==-1?data[filaSheet-1][idxUrl]:''), nombre:'GDP - '+nombre } };
  }catch(err){
    return { ok:false, error:'Error al asegurar el comercial del CEC: ' + err.message };
  }
}

// ── Auditoría backend (hoja "Auditoria" en el Spreadsheet principal) ──
function gdpAuditar_(quien, accion, detalle, cecId){
  try{
    var ss = getOrCreateSpreadsheet();
    var sh = ss.getSheetByName('Auditoria');
    if(!sh){ sh = ss.insertSheet('Auditoria'); sh.appendRow(['ts','quien','accion','detalle','cecId']); }
    sh.appendRow([new Date().toISOString(), quien, accion, detalle, cecId]);
  }catch(e){ /* auditoría nunca debe romper la operación */ }
}

// ── SETUP INICIAL — ejecutar UNA SOLA VEZ ────────────────
function fijarSpreadsheetId() {
  PropertiesService.getScriptProperties()
    .setProperty('SS_ID', '1vH_eWl4h0UeW00BC8IeDagnu_aNLjDbOu_UikRZFejM');
  Logger.log('✅ SS_ID fijado correctamente');
}

// ── REPARAR CECs — ejecutar UNA VEZ si la hoja CECs quedó vacía ──
// Siembra Mar del Plata (CEC original) si no está. Seguro de correr varias veces.
function repararCECs() {
  var ss = getOrCreateSpreadsheet();
  var sh = ss.getSheetByName(GDP_HOJA_CECS);
  if(!sh){
    // No existe: gdpHojaCECs_ la crea Y siembra MdP.
    gdpHojaCECs_();
    Logger.log('✅ Hoja CECs creada con Mar del Plata sembrado.');
    return;
  }
  // Existe: ¿está Mar del Plata?
  var data = sh.getDataRange().getValues();
  var headers = data[0];
  var idxId = headers.indexOf('cecId');
  for (var r = 1; r < data.length; r++) {
    if (String(data[r][idxId]) === 'cec_mar_del_plata') {
      Logger.log('✅ Mar del Plata ya existe. Nada que reparar.');
      return;
    }
  }
  // Falta: sembrarlo.
  var ssId = ss.getId();
  sh.appendRow([
    'cec_mar_del_plata', 'Mar del Plata', true,
    ssId, 'https://docs.google.com/spreadsheets/d/' + ssId + '/edit',
    '', 'Argentina', 'Buenos Aires', 'Mar del Plata', 'Av. Colón 2550',
    '', '', -38.00377226264887, -57.54984229189988, '', '', '', '',
    new Date().toISOString()
  ]);
  Logger.log('✅ Mar del Plata sembrado en la hoja CECs existente.');
}

// ── Eliminar la fila comercial de un asesor (al borrarlo desde config) ──
// Quita al asesor de TODAS las pestañas de mes del CEC (match por Usuario o por Asesor),
// para que un asesor eliminado no quede huérfano en el Sheet. Genérico para cualquier CEC.
function v2_eliminarAsesorComercial(params){
  var auth = autorizarOperacion_(params);
  if (!auth.ok) return { ok:false, error:auth.error };
  try{
    var cec     = String(params.cec || '').trim();
    var nombre  = String(params.nombre || '').trim().toLowerCase();
    var usuario = String(params.usuario || '').trim().toLowerCase();
    var mesTab  = String(params.mes || '').trim();   // opcional: MES_AAAA. Vacío = todas.
    if (!cec || (!nombre && !usuario)) return { ok:false, error:'Faltan datos (cec + nombre/usuario).' };

    var ss = gdpSpreadsheetDeCEC_(cec);
    if (!ss) return { ok:false, error:'CEC sin spreadsheet.' };

    var borradas = 0;
    ss.getSheets().forEach(function(sh){
      var nm = sh.getName();
      if (!/_\d{4}$/.test(nm)) return;          // solo pestañas de mes (MES_AAAA)
      if (mesTab && nm !== mesTab) return;       // si se pidió un mes puntual, solo ese
      var data = sh.getDataRange().getValues();
      if (data.length < 2) return;
      var headers = data[0];
      var aIdx = headers.indexOf('Asesor');
      var uIdx = gdpFindColIdx_(headers, 'Usuario');
      // De abajo hacia arriba para que deleteRow no descoloque los índices.
      for (var r = data.length - 1; r >= 1; r--) {
        var matchU = usuario && uIdx !== -1 && String(data[r][uIdx]).trim().toLowerCase() === usuario;
        var matchN = nombre  && aIdx !== -1 && String(data[r][aIdx]).trim().toLowerCase() === nombre;
        if (matchU || matchN) { sh.deleteRow(r + 1); borradas++; }
      }
    });
    SpreadsheetApp.flush();
    return { ok:true, borradas:borradas };
  } catch(err){
    return { ok:false, error: err.message };
  }
}
