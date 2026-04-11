// ═══════════════════════════════════════════════════════════════════════════
// m.kids CRM — Google Apps Script
//
// ВСТАНОВЛЕННЯ:
//   1. Відкрийте конфігураційну таблицю:
//      https://docs.google.com/spreadsheets/d/11NEIEBzaMiIDFnJB9RXqKnRqjCJjNyHVqylrX7cRZhc
//   2. Extensions → Apps Script → вставте весь цей код
//   3. Запустіть setup() один раз (кнопка Run)
//   4. Deploy → New deployment → Web app
//      Execute as: Me | Who has access: Anyone
//   5. Скопіюйте URL деплойменту → вставте в clients.html → Налаштування
// ═══════════════════════════════════════════════════════════════════════════

// ─── CONSTANTS ───────────────────────────────────────────────────────────────
var CONFIG_SHEET_ID  = '11NEIEBzaMiIDFnJB9RXqKnRqjCJjNyHVqylrX7cRZhc';
var CRM_SHEET_ID_DEFAULT = '1pA2q84BFsXWuUchIlu8um853od_PXr7KepLpTovUjLo'; // fallback
var SHEET_PAYMENTS   = 'Оплати';
var SHEET_CLIENTS    = 'Клієнти';

var MONTHS_UA      = ['вересень','жовтень','листопад','грудень','січень','лютий','березень','квітень','травень','червень','липень','серпень'];
var MONTHS_JS      = [8,9,10,11,0,1,2,3,4,5,6,7];
var MONTHS_DISPLAY = ['Вересень','Жовтень','Листопад','Грудень','Січень','Лютий','Березень','Квітень','Травень','Червень','Липень','Серпень'];

// ─── INITIAL SETUP ───────────────────────────────────────────────────────────
/**
 * Запустіть цю функцію ОДИН РАЗ після вставки коду.
 * Вона створить таблицю "m.kids CRM Data" і налаштує тригер.
 */
function setup() {
  getCRMSpreadsheet(); // creates if not exists
  createDailyTrigger();
  Logger.log('Setup done. CRM Sheet ID: ' + getProps().getProperty('CRM_SHEET_ID'));
}

// ─── PROPERTIES ──────────────────────────────────────────────────────────────
function getProps() {
  return PropertiesService.getScriptProperties();
}

// ─── CRM SPREADSHEET ─────────────────────────────────────────────────────────
function getCRMSpreadsheet() {
  var props = getProps();
  var id = props.getProperty('CRM_SHEET_ID') || CRM_SHEET_ID_DEFAULT;
  if (id) {
    try {
      var ss = SpreadsheetApp.openById(id);
      // Save to properties if it came from default
      if (!props.getProperty('CRM_SHEET_ID')) {
        props.setProperty('CRM_SHEET_ID', id);
        Logger.log('Saved CRM_SHEET_ID to Script Properties: ' + id);
      }
      ensureSheetsExist(ss);
      return ss;
    } catch(e) {
      Logger.log('Cannot open CRM sheet ' + id + ': ' + e.message);
    }
  }
  // Create new spreadsheet (fallback)
  var newSS = SpreadsheetApp.create('m.kids CRM Data');
  props.setProperty('CRM_SHEET_ID', newSS.getId());
  setupSheetsStructure(newSS);
  Logger.log('Created new CRM spreadsheet: ' + newSS.getUrl());
  return newSS;
}

function ensureSheetsExist(ss) {
  if (!ss.getSheetByName(SHEET_PAYMENTS)) {
    var s = ss.insertSheet(SHEET_PAYMENTS);
    writePaymentsHeader(s);
  }
  if (!ss.getSheetByName(SHEET_CLIENTS)) {
    var s2 = ss.insertSheet(SHEET_CLIENTS);
    writeClientsHeader(s2);
  }
}

function setupSheetsStructure(ss) {
  // Rename default sheet to Payments
  var sheets = ss.getSheets();
  sheets[0].setName(SHEET_PAYMENTS);
  writePaymentsHeader(sheets[0]);
  // Add Clients sheet
  var cli = ss.insertSheet(SHEET_CLIENTS);
  writeClientsHeader(cli);
}

function writePaymentsHeader(sheet) {
  sheet.clearContents();
  sheet.appendRow([
    'Локація','Напрямок','Тип','Група','Вихователь',"Ім'я дитини",
    'Факт навчання','Факт доп.','Факт разом','Бюджет',
    'Статус','Місяць','Оновлено'
  ]);
  sheet.setFrozenRows(1);
  sheet.getRange(1,1,1,13).setFontWeight('bold');
}

function writeClientsHeader(sheet) {
  sheet.clearContents();
  sheet.appendRow([
    'ID','ПІБ дитини','Локація','Група','Вихователь',
    'ПІБ мами','Телефон мами','ПІБ тата','Телефон тата',
    'Дата договору','Тип договору','Сума договору','Вступний внесок',
    'Статус','Нотатки',
    'Відсутності (JSON)','Графік внеску (JSON)','Зміни суми (JSON)',
    'Дата народження','Створено','Оновлено'
  ]);
  sheet.setFrozenRows(1);
  sheet.getRange(1,1,1,21).setFontWeight('bold');
}

// ─── WEB APP ENDPOINTS ───────────────────────────────────────────────────────
function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : '';
  try {
    var result;
    if      (action === 'ping')         result = {ok:true, msg:'pong', ts: new Date().toISOString()};
    else if (action === 'getLocations') result = getLocations();
    else if (action === 'getPayments')  result = getPayments();
    else if (action === 'getClients')   result = getClients();
    else if (action === 'runAggregate') result = aggregatePayments();
    else if (action === 'getCRMInfo')   result = getCRMInfo();
    else if (action === 'makePublic')   result = makeSheetPublic();
    else                                result = {ok:false, error:'Unknown action: ' + action};
    return jsonOut(result);
  } catch(err) {
    return jsonOut({ok:false, error:err.message || String(err)});
  }
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var result;
    if      (body.action === 'saveClient')   result = saveClient(body.data);
    else if (body.action === 'deleteClient') result = deleteClient(body.id);
    else if (body.action === 'saveBulk')     result = saveBulkClients(body.data);
    else                                     result = {ok:false, error:'Unknown action: ' + body.action};
    return jsonOut(result);
  } catch(err) {
    return jsonOut({ok:false, error:err.message || String(err)});
  }
}

function jsonOut(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─── GET LOCATIONS ───────────────────────────────────────────────────────────
function getLocations() {
  var configSS = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  var configSheet = configSS.getSheets()[0];
  var data = configSheet.getDataRange().getValues();
  var locs = [];
  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    var loc     = trim(row[2]);
    var sheetId = trim(row[3]);
    if (!loc || !sheetId) continue;
    locs.push({
      dir:       trim(row[0]),
      typ:       trim(row[1]),
      loc:       loc,
      sheetId:   sheetId,
      sheetName: trim(row[4]) || 'Payment'
    });
  }
  return {ok:true, data:locs};
}

// ─── GET PAYMENTS ────────────────────────────────────────────────────────────
function getPayments() {
  var ss = getCRMSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_PAYMENTS);
  if (!sheet) return {ok:false, error:'Sheet "Оплати" not found'};
  var vals = sheet.getDataRange().getValues();
  if (vals.length < 2) return {ok:true, data:[], updated:''};
  var headers = vals[0];
  var rows = [];
  for (var r = 1; r < vals.length; r++) {
    if (!vals[r].some(function(v){ return v !== ''; })) continue;
    var obj = {};
    for (var c = 0; c < headers.length; c++) {
      obj[String(headers[c])] = vals[r][c];
    }
    rows.push(obj);
  }
  // Last update time from last row
  var updated = rows.length > 0 ? String(rows[rows.length-1]['Оновлено']||'') : '';
  return {ok:true, data:rows, updated:updated};
}

// ─── GET CLIENTS ─────────────────────────────────────────────────────────────
function getClients() {
  var ss = getCRMSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_CLIENTS);
  if (!sheet) return {ok:false, error:'Sheet "Клієнти" not found'};
  var vals = sheet.getDataRange().getValues();
  if (vals.length < 2) return {ok:true, data:[]};
  var headers = vals[0];
  var rows = [];
  for (var r = 1; r < vals.length; r++) {
    if (!vals[r][0]) continue; // skip rows without ID
    var obj = {};
    for (var c = 0; c < headers.length; c++) {
      obj[String(headers[c])] = vals[r][c];
    }
    rows.push(obj);
  }
  return {ok:true, data:rows};
}

// ─── SAVE CLIENT ─────────────────────────────────────────────────────────────
function saveClient(data) {
  if (!data || !data.id) return {ok:false, error:'Missing id'};
  var ss = getCRMSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_CLIENTS);
  if (!sheet) return {ok:false, error:'Sheet not found'};

  var vals = sheet.getDataRange().getValues();
  var now = formatDate(new Date());
  var createdAt = data.createdAt || now;

  var row = [
    data.id,
    data.name        || '',
    data.loc         || '',
    data.group       || '',
    data.teacher     || '',
    data.momName     || '',
    data.momPhone    || '',
    data.dadName     || '',
    data.dadPhone    || '',
    data.contractDate || '',
    data.contractType || 'standard',
    data.monthlyFee  || 0,
    data.entryFee    || 0,
    data.status      || 'active',
    data.notes       || '',
    JSON.stringify(data.absences          || []),
    JSON.stringify(data.entryFeeSchedule  || []),
    JSON.stringify(data.feeHistory        || []),
    data.bday        || '',
    createdAt,
    now
  ];

  // Find existing row by ID
  for (var r = 1; r < vals.length; r++) {
    if (String(vals[r][0]) === String(data.id)) {
      row[19] = vals[r][19] || createdAt; // preserve createdAt
      sheet.getRange(r + 1, 1, 1, row.length).setValues([row]);
      return {ok:true, action:'updated'};
    }
  }
  // New client
  sheet.appendRow(row);
  return {ok:true, action:'created'};
}

function deleteClient(id) {
  if (!id) return {ok:false, error:'Missing id'};
  var ss = getCRMSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_CLIENTS);
  if (!sheet) return {ok:false, error:'Sheet not found'};
  var vals = sheet.getDataRange().getValues();
  for (var r = vals.length - 1; r >= 1; r--) {
    if (String(vals[r][0]) === String(id)) {
      sheet.deleteRow(r + 1);
      return {ok:true};
    }
  }
  return {ok:false, error:'Client not found'};
}

function saveBulkClients(dataArr) {
  if (!Array.isArray(dataArr)) return {ok:false, error:'Expected array'};
  var saved = 0;
  dataArr.forEach(function(d){ saveClient(d); saved++; });
  return {ok:true, saved:saved};
}

// ─── AGGREGATE PAYMENTS ──────────────────────────────────────────────────────
/**
 * Головна функція: читає Payment Sheets → записує в "Оплати"
 * Запускається автоматично щодня о 06:00 Київ
 */
function aggregatePayments() {
  var configSS = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  var configSheet = configSS.getSheets()[0];
  var configData = configSheet.getDataRange().getValues();

  var crmSS = getCRMSpreadsheet();
  var paySheet = crmSS.getSheetByName(SHEET_PAYMENTS);
  if (!paySheet) {
    paySheet = crmSS.insertSheet(SHEET_PAYMENTS, 0);
    writePaymentsHeader(paySheet);
  }

  var now = new Date();
  var curJSMonth = now.getMonth();
  var monthName = getMonthDisplayName(curJSMonth);
  var updateStr = formatDate(now);

  var allRows = [];
  var errors  = [];

  for (var r = 1; r < configData.length; r++) {
    var cfgRow    = configData[r];
    var dir       = trim(cfgRow[0]);
    var typ       = trim(cfgRow[1]);
    var loc       = trim(cfgRow[2]);
    var sheetId   = trim(cfgRow[3]);
    var sheetName = trim(cfgRow[4]) || 'Payment';

    if (!loc || !sheetId) continue;

    try {
      var ss = SpreadsheetApp.openById(sheetId);
      var paymentSheet = ss.getSheetByName(sheetName);
      if (!paymentSheet) {
        paymentSheet = ss.getSheets()[0];
        Logger.log('Sheet "' + sheetName + '" not found in ' + loc + ', using first sheet: ' + paymentSheet.getName());
      }
      var data = paymentSheet.getDataRange().getValues();

      var monthCol = detectCurrentMonthCol(data, curJSMonth);
      Logger.log(loc + ': detected month col = ' + monthCol + ' for month ' + monthName);

      var groups = parsePaymentSheet(data, monthCol);

      groups.forEach(function(g) {
        g.children.forEach(function(ch) {
          var fs    = ch.factStudy  || 0;
          var fe    = ch.factExtra  || 0;
          var bud   = ch.budget     || 0;
          var total = fs + fe;
          var status;
          if (bud === 0)        status = 'unknown';
          else if (total === 0) status = 'nopay';
          else if (total > bud) status = 'over';
          else if (total === bud) status = 'paid';
          else                  status = 'debt';

          allRows.push([
            loc, dir, typ,
            g.group, g.teacher, ch.name,
            fs, fe, total, bud,
            status, monthName, updateStr
          ]);
        });
      });
    } catch(e) {
      var errMsg = loc + ': ' + e.message;
      errors.push(errMsg);
      Logger.log('ERROR ' + errMsg);
    }
  }

  // Write all rows
  paySheet.clearContents();
  writePaymentsHeader(paySheet);
  if (allRows.length > 0) {
    paySheet.getRange(2, 1, allRows.length, 13).setValues(allRows);
  }

  Logger.log('aggregatePayments done: ' + allRows.length + ' rows, ' + errors.length + ' errors');
  return {ok:true, rows:allRows.length, errors:errors, month:monthName, updated:updateStr};
}

// ─── PAYMENT SHEET PARSING ────────────────────────────────────────────────────
function detectCurrentMonthCol(rows, curJSMonth) {
  for (var r = 0; r < Math.min(7, rows.length); r++) {
    for (var c = 1; c < rows[r].length; c++) {
      var cell = String(rows[r][c] || '').toLowerCase().trim();
      for (var mi = 0; mi < MONTHS_UA.length; mi++) {
        if (cell.indexOf(MONTHS_UA[mi]) !== -1 && MONTHS_JS[mi] === curJSMonth) {
          return c;
        }
      }
    }
  }
  // Fallback: calculate from school year (Sep=col1, Oct=col4, ...)
  var schoolOrder = [8,9,10,11,0,1,2,3,4,5,6,7];
  var idx = schoolOrder.indexOf(curJSMonth);
  var col = idx < 0 ? 1 : 1 + idx * 3;
  Logger.log('Month col not detected in headers, using fallback col=' + col);
  return col;
}

function parsePaymentSheet(data, monthCol) {
  var DATA_START = 3; // sheet row 4 = array index 3
  var groups = [];
  var curGroup = null;

  for (var r = DATA_START; r < data.length; r++) {
    var row = data[r];
    var nameCell = trim(String(row[0] || ''));
    if (!nameCell) continue;

    if (isGroupHeaderRow(row, monthCol)) {
      // Extract teacher: last word after whitespace
      var parts = nameCell.split(/\s+/);
      var teacher = parts.length > 1 ? parts[parts.length - 1] : '';
      curGroup = {group: nameCell, teacher: teacher, children: []};
      groups.push(curGroup);
    } else {
      if (!curGroup) {
        curGroup = {group: '(без групи)', teacher: '', children: []};
        groups.push(curGroup);
      }
      var fs  = toNum(row[monthCol]);
      var fe  = toNum(row[monthCol + 1]);
      var bud = toNum(row[monthCol + 2]);
      curGroup.children.push({name: nameCell, factStudy: fs, factExtra: fe, budget: bud});
    }
  }
  return groups.filter(function(g){ return g.children.length > 0; });
}

function isGroupHeaderRow(row, monthCol) {
  var nameCell = trim(String(row[0] || ''));
  if (!nameCell) return false;

  // Contains digits like "4-5", "(", or keywords → likely a group
  var looksLikeGroup = /\d/.test(nameCell) || /study|kids|preschool|kid/i.test(nameCell) || nameCell.indexOf('(') !== -1;

  // Check that payment columns have no positive values
  var hasPayValue = false;
  for (var c = monthCol; c <= monthCol + 2 && c < row.length; c++) {
    var v = toNum(row[c]);
    if (v > 0) { hasPayValue = true; break; }
  }

  if (looksLikeGroup && !hasPayValue) return true;
  if (!hasPayValue) return true; // any row without payment data is treated as header
  return false;
}

// ─── DAILY TRIGGER ───────────────────────────────────────────────────────────
function createDailyTrigger() {
  // Remove existing
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'aggregatePayments') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  // Create new at 06:00 Kyiv
  ScriptApp.newTrigger('aggregatePayments')
    .timeBased()
    .everyDays(1)
    .atHour(6)
    .inTimezone('Europe/Kiev')
    .create();
  Logger.log('Daily trigger created: aggregatePayments at 06:00 Kyiv');
}

// ─── CRM INFO ────────────────────────────────────────────────────────────────
function getCRMInfo() {
  var props = getProps();
  var id = props.getProperty('CRM_SHEET_ID') || CRM_SHEET_ID_DEFAULT;
  var url = id ? 'https://docs.google.com/spreadsheets/d/' + id : '';
  return {ok:true, crmSheetId: id, crmSheetUrl: url};
}

// ─── MAKE SHEET PUBLIC (read-only for anyone with link) ──────────────────────
function makeSheetPublic() {
  var props = getProps();
  var id = props.getProperty('CRM_SHEET_ID') || CRM_SHEET_ID_DEFAULT;
  if (!id) return {ok:false, error:'CRM_SHEET_ID not set'};
  try {
    var file = DriveApp.getFileById(id);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var url = 'https://docs.google.com/spreadsheets/d/' + id;
    Logger.log('CRM sheet is now public (read-only): ' + url);
    return {ok:true, url:url, msg:'Таблиця відкрита для читання (Anyone with link → Viewer)'};
  } catch(e) {
    Logger.log('makeSheetPublic error: ' + e.message);
    return {ok:false, error:e.message};
  }
}

// ─── UTILS ───────────────────────────────────────────────────────────────────
function trim(s) {
  return String(s || '').trim();
}
function toNum(v) {
  if (v === '' || v === null || v === undefined) return 0;
  var n = parseFloat(String(v).replace(',', '.'));
  return isNaN(n) ? 0 : n;
}
function formatDate(d) {
  return Utilities.formatDate(d, 'Europe/Kiev', 'dd.MM.yyyy HH:mm');
}
function getMonthDisplayName(jsMonth) {
  var idx = MONTHS_JS.indexOf(jsMonth);
  return idx >= 0 ? MONTHS_DISPLAY[idx] : String(jsMonth + 1);
}
