// ═══════════════════════════════════════════════════════════════════════════
// m.kids CRM — Google Apps Script v5.8
// 5 колонок на місяць: навч | вступ | доп | бюджет доп | бюджет навч
// Лист "Оплати-Рік" — повні дані за 12 місяців календарного року
// ═══════════════════════════════════════════════════════════════════════════

var CONFIG_SHEET_ID  = '11NEIEBzaMiIDFnJB9RXqKnRqjCJjNyHVqylrX7cRZhc';
var SHEET_PAYMENTS   = 'Оплати';
var SHEET_YEARLY     = 'Оплати-Рік';
var SHEET_CLIENTS    = 'Клієнти';
var SHEET_ATTENDANCE = 'Табель';
var SHEET_HEALTH     = 'Здоров\'я';

var MONTHS_UA      = ['вересень','жовтень','листопад','грудень','січень','лютий','березень','квітень','травень','червень','липень','серпень'];
var MONTHS_JS      = [8,9,10,11,0,1,2,3,4,5,6,7];
var MONTHS_DISPLAY = ['Вересень','Жовтень','Листопад','Грудень','Січень','Лютий','Березень','Квітень','Травень','Червень','Липень','Серпень'];
var MONTHS_CAL = ['Січень','Лютий','Березень','Квітень','Травень','Червень','Липень','Серпень','Вересень','Жовтень','Листопад','Грудень'];

var GROUP_PATTERNS = [
  /mini.?baby/i, /^baby/i, /find/i, /study/i, /preschool/i,
  /чомус/i, /^школа$/i, /^гхзд$/i,
  /мама[\s\+]*я/i, /малюк/i, /карапуз/i, /пізнайк/i,
  /бешкетн/i, /мандрівн/i, /дослідн/i, /розумник/i,
  /^\s*\d+\s*([dDsS]\s*(клас|кл)?|класс?|кл\.?|[бвБВ])/
];

function normalizeGroupName(raw) {
  var s = trim(raw);
  if (/mini.?baby/i.test(s))  return 'miniBaby-ki';
  if (/^baby/i.test(s))       return 'Baby-ki';
  if (/find/i.test(s))        return 'Find-iki';
  if (/study/i.test(s))       return 'Study-ki';
  if (/preschool/i.test(s))   return 'Preschool';
  if (/чомус/i.test(s))       return 'Чомусики';
  if (/^гхзд$/i.test(s))      return 'ГХЗД';
  if (/мама[\s\+]*я/i.test(s))  return 'miniBaby-ki';
  if (/малюк/i.test(s))         return 'Baby-ki';
  if (/карапуз/i.test(s))       return 'Baby-ki';
  if (/пізнайк/i.test(s))       return 'Study-ki';
  if (/бешкетн/i.test(s))       return 'Find-iki';
  if (/мандрівн/i.test(s))      return 'Study-ki';
  if (/дослідн/i.test(s))       return 'Study-ki';
  if (/розумник/i.test(s))      return 'Preschool';
  if (/^\s*\d+\s*([dDsS]\s*(клас|кл)?|класс?|кл\.?|[бвБВ])/i.test(s)) return 'Школа';
  if (/^школа$/i.test(s))       return 'Школа';
  return s;
}

function isGroupHeaderRow(row, monthCol) {
  var nameCell = trim(String(row[0] || ''));
  if (!nameCell) return false;
  for (var i = 0; i < GROUP_PATTERNS.length; i++) {
    if (GROUP_PATTERNS[i].test(nameCell)) return true;
  }
  if (/вільних|місць|разом|всього|оплата за/i.test(nameCell)) return true;
  return false;
}

function setup() {
  getCRMSpreadsheet();
  createDailyTrigger();
  Logger.log('Setup done.');
}

function fixCRMSheetId() {
  var correctId = '1pA2q84BFsXWuUchIlu8um853od_PXr7KepLpTovUjLo';
  var props = PropertiesService.getScriptProperties();
  Logger.log('Поточний CRM_SHEET_ID: ' + props.getProperty('CRM_SHEET_ID'));
  props.setProperty('CRM_SHEET_ID', correctId);
  Logger.log('Новий CRM_SHEET_ID: ' + props.getProperty('CRM_SHEET_ID'));
  var ss = SpreadsheetApp.openById(correctId);
  Logger.log('SS name: ' + ss.getName());
  Logger.log('SS sheets: ' + ss.getSheets().map(function(s){ return s.getName(); }).join(', '));
}

function getProps() {
  return PropertiesService.getScriptProperties();
}

function getCRMSpreadsheet() {
  var props = getProps();
  var id = props.getProperty('CRM_SHEET_ID');
  if (id) {
    try {
      var ss = SpreadsheetApp.openById(id);
      ensureSheetsExist(ss);
      return ss;
    } catch(e) {}
  }
  var newSS = SpreadsheetApp.create('m.kids CRM Data');
  props.setProperty('CRM_SHEET_ID', newSS.getId());
  setupSheetsStructure(newSS);
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
  if (!ss.getSheetByName(SHEET_ATTENDANCE)) {
    var s3 = ss.insertSheet(SHEET_ATTENDANCE);
    writeAttendanceHeader(s3);
  }
  if (!ss.getSheetByName(SHEET_HEALTH)) {
    var s4 = ss.insertSheet(SHEET_HEALTH);
    writeHealthHeader(s4);
  }
}

function setupSheetsStructure(ss) {
  var sheets = ss.getSheets();
  sheets[0].setName(SHEET_PAYMENTS);
  writePaymentsHeader(sheets[0]);
  var cli = ss.insertSheet(SHEET_CLIENTS);
  writeClientsHeader(cli);
}

function writePaymentsHeader(sheet) {
  sheet.clearContents();
  sheet.appendRow([
    'Локація','Напрямок','Тип','Група','Вихователь',"Ім'я дитини",
    'Факт навчання','Факт вступний','Факт доп.','Факт разом',
    'Бюджет навчання','Бюджет доп.','Бюджет разом',
    'Статус','Місяць','Оновлено','Дата договору'
  ]);
  sheet.setFrozenRows(1);
}

function writeAttendanceHeader(sheet) {
  sheet.clearContents();
  sheet.appendRow(['Дата','ID дитини','Ім\'я дитини','Локація','Група','Статус','Ким','Коли']);
  sheet.setFrozenRows(1);
}

function writeHealthHeader(sheet) {
  sheet.clearContents();
  sheet.appendRow(['ID запису','ID дитини','Дата','Тип','Текст','Ким додано','Створено']);
  sheet.setFrozenRows(1);
}

function writeClientsHeader(sheet) {
  sheet.clearContents();
  sheet.appendRow([
    'ID','ПІБ дитини','Локація','Група','Вихователь','Дата народження',
    'ПІБ мами','Телефон мами','ПІБ тата','Телефон тата',
    'Дата договору','Тип договору','Сума договору','Вступний внесок',
    'Статус','Нотатки',
    'Відсутності (JSON)','Графік внеску (JSON)','Зміни суми (JSON)',
    'Номер договору','Дата адаптації','Дата розірвання','Причина розірвання',
    'Свідоцтво про народження','Місце реєстрації дитини',
    'Документ мами','РНОКПП мами','Документ тата','РНОКПП тата',
    'Створено','Оновлено'
  ]);
  sheet.setFrozenRows(1);
}

function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : '';
  try {
    var result;
    if      (action === 'ping')               result = {ok:true, msg:'pong v5.8', ts: new Date().toISOString()};
    else if (action === 'getLocations')       result = getLocations();
    else if (action === 'getPayments')        result = getPayments();
    else if (action === 'getPaymentsYearly')  result = getPaymentsYearly();
    else if (action === 'getClients')         result = getClients();
    else if (action === 'runAggregate')       result = aggregatePayments();
    else if (action === 'runAggregateYearly') result = aggregatePaymentsYearly();
    else if (action === 'runSyncBdayStatus')  result = syncBdayStatusSheet();
    else if (action === 'getRegistryUrls')    result = getRegistryUrls();
    else if (action === 'makePublic')         result = makeSheetPublic();
    else if (action === 'getAttendance')      result = getAttendance(e);
    else if (action === 'getHealthRecords')         result = getHealthRecords(e);
    else if (action === 'dryRunImportAbsences')      result = dryRunImportAbsences(e.parameter.loc || '');
    else if (action === 'importAbsencesFromPayment') result = importAbsencesFromPayment(e.parameter.loc || '');
    else                                             result = {ok:false, error:'Unknown action: ' + action};
    return jsonOut(result);
  } catch(err) {
    return jsonOut({ok:false, error:err.message || String(err)});
  }
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var result;
    if      (body.action === 'saveClient')       result = saveClient(body.data);
    else if (body.action === 'deleteClient')     result = deleteClient(body.id);
    else if (body.action === 'saveAttendance')   result = saveAttendance(body);
    else if (body.action === 'saveHealthRecord') result = saveHealthRecord(body);
    else if (body.action === 'deleteHealthRecord')    result = deleteHealthRecord(body);
    else if (body.action === 'writeAbsenceToPayment')    result = writeAbsenceToPayment(body);
    else if (body.action === 'importAbsencesFromPayment') result = importAbsencesFromPayment(body.loc || '');
    else if (body.action === 'confirmBdayMatch')          result = confirmBdayMatch(body.childId || '', body.confirmedBy || '');
    else if (body.action === 'unconfirmBdayMatch')        result = unconfirmBdayMatch(body.childId || '');
    else result = {ok:false, error:'Unknown action'};
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

function getLocations() {
  var configSS = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  var configSheet = configSS.getSheets()[0];
  var data = configSheet.getDataRange().getValues();
  var locs = [];
  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    var dir       = trim(row[0]);
    var typ       = trim(row[1]);
    var loc       = trim(row[2]);
    var sheetId   = trim(row[3]);
    var sheetName = trim(row[4]) || 'Payment';
    if (!loc || !sheetId) continue;
    locs.push({
      dir: dir, typ: typ,
      loc: loc, sheetId: sheetId,
      sheetName: sheetName,
      pw: trim(row[5]) || ''
    });
  }
  return {ok:true, data:locs};
}

function getPayments() {
  var ss = getCRMSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_PAYMENTS);
  if (!sheet) return {ok:false, error:'Sheet not found'};
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
  var updated = rows.length > 0 ? String(rows[rows.length-1]['Оновлено']||'') : '';
  return {ok:true, data:rows, updated:updated};
}

function getClients() {
  var ss = getCRMSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_CLIENTS);
  if (!sheet) return {ok:false, error:'Sheet not found'};
  var vals = sheet.getDataRange().getValues();
  if (vals.length < 2) return {ok:true, data:[]};
  var headers = vals[0];
  var rows = [];
  for (var r = 1; r < vals.length; r++) {
    if (!vals[r][0]) continue;
    var obj = {};
    for (var c = 0; c < headers.length; c++) {
      obj[String(headers[c])] = vals[r][c];
    }
    rows.push(obj);
  }
  return {ok:true, data:rows};
}

function ensureClientsHeader(sheet) {
  var EXPECTED = [
    'ID','ПІБ дитини','Локація','Група','Вихователь','Дата народження',
    'ПІБ мами','Телефон мами','ПІБ тата','Телефон тата',
    'Дата договору','Тип договору','Сума договору','Вступний внесок',
    'Статус','Нотатки',
    'Відсутності (JSON)','Графік внеску (JSON)','Зміни суми (JSON)',
    'Номер договору','Дата адаптації','Дата розірвання','Причина розірвання',
    'Свідоцтво про народження','Місце реєстрації дитини',
    'Документ мами','РНОКПП мами','Документ тата','РНОКПП тата',
    'Створено','Оновлено'
  ];
  var lastCol = sheet.getLastColumn();
  var width = Math.max(lastCol, EXPECTED.length);
  var current = sheet.getRange(1, 1, 1, width).getValues()[0];
  for (var i = 0; i < EXPECTED.length; i++) {
    if (String(current[i] || '').trim() !== EXPECTED[i]) {
      sheet.getRange(1, 1, 1, EXPECTED.length).setValues([EXPECTED]);
      sheet.setFrozenRows(1);
      return;
    }
  }
}

function saveClient(data) {
  if (!data || !data.id) return {ok:false, error:'Missing id'};
  var ss = getCRMSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_CLIENTS);
  if (!sheet) return {ok:false, error:'Sheet not found'};
  ensureClientsHeader(sheet);
  var vals = sheet.getDataRange().getValues();
  var now = formatDate(new Date());
  var row = [
    data.id, data.name||'', data.loc||'', data.group||'', data.teacher||'',
    data.bday||'', data.momName||'', data.momPhone||'', data.dadName||'', data.dadPhone||'',
    data.contractDate||'', data.contractType||'standard', data.monthlyFee||0, data.entryFee||0,
    data.status||'active', data.notes||'',
    JSON.stringify(data.absences||[]),
    JSON.stringify(data.entryFeeSchedule||[]),
    JSON.stringify(data.feeHistory||[]),
    data.contractNumber||'', data.adaptDate||'', data.terminationDate||'',
    data.terminationReason||'',
    data.birthCert||'', data.childRegAddress||'',
    data.momDoc||'', data.momRnokpp||'',
    data.dadDoc||'', data.dadRnokpp||'',
    data.createdAt||now, now
  ];
  for (var r = 1; r < vals.length; r++) {
    if (String(vals[r][0]) === String(data.id)) {
      row[29] = vals[r][29] || data.createdAt || now;
      sheet.getRange(r+1, 1, 1, row.length).setValues([row]);
      return {ok:true, action:'updated'};
    }
  }
  sheet.appendRow(row);
  return {ok:true, action:'created'};
}

function patchClientAbsences(id, absences) {
  var ss    = getCRMSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_CLIENTS);
  if (!sheet) return {ok:false, error:'Sheet not found'};

  var vals   = sheet.getDataRange().getValues();
  var hdrs   = vals[0].map(String);
  var colAbs = hdrs.indexOf('Відсутності (JSON)');
  var colUpd = hdrs.indexOf('Оновлено');
  if (colAbs < 0) return {ok:false, error:'Column "Відсутності (JSON)" not found'};

  for (var r = 1; r < vals.length; r++) {
    if (String(vals[r][0]) === String(id)) {
      sheet.getRange(r+1, colAbs+1).setValue(JSON.stringify(absences));
      if (colUpd >= 0) sheet.getRange(r+1, colUpd+1).setValue(formatDate(new Date()));
      return {ok:true, action:'patched', row: r+1};
    }
  }
  return {ok:false, error:'Client ID not found: ' + id};
}

function deleteClient(id) {
  if (!id) return {ok:false, error:'Missing id'};
  var ss = getCRMSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_CLIENTS);
  if (!sheet) return {ok:false, error:'Sheet not found'};
  var vals = sheet.getDataRange().getValues();
  for (var r = vals.length-1; r >= 1; r--) {
    if (String(vals[r][0]) === String(id)) {
      sheet.deleteRow(r+1);
      return {ok:true};
    }
  }
  return {ok:false, error:'Not found'};
}

function makeSheetPublic() {
  var ss = getCRMSpreadsheet();
  var file = DriveApp.getFileById(ss.getId());
  file.setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);
  return {ok:true, url:ss.getUrl()};
}

function aggregatePayments() {
  var configSS = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  var configSheet = configSS.getSheets()[0];
  var configData = configSheet.getDataRange().getValues();

  var crmSS = getCRMSpreadsheet();
  var paySheet = crmSS.getSheetByName(SHEET_PAYMENTS);
  if (!paySheet) paySheet = crmSS.insertSheet(SHEET_PAYMENTS, 0);

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
      if (!paymentSheet) paymentSheet = ss.getSheets()[0];
      var data = paymentSheet.getDataRange().getValues();
      var monthCol    = detectCurrentMonthCol(data, curJSMonth);
      var contractCol = detectContractDateCol(data);
      Logger.log(loc + ': monthCol=' + monthCol + ', month=' + monthName + ', contractCol=' + contractCol);
      var groups = parsePaymentSheet(data, monthCol, contractCol);
      Logger.log(loc + ': groups=' + groups.length);

      groups.forEach(function(g) {
        g.children.forEach(function(ch) {
          var fs = ch.factStudy || 0;
          var fv = ch.factEntry || 0;
          var fe = ch.factExtra || 0;
          var bd = ch.budExtra  || 0;
          var bs = ch.budStudy  || 0;
          var total = fs + fv + fe;
          var br = bs + bd;
          var totalNoEntry = fs + fe;
          var status;
          if (br === 0 && totalNoEntry === 0) status = 'unknown';
          else if (totalNoEntry === 0 && br > 0) status = 'nopay';
          else if (totalNoEntry > br)  status = 'over';
          else if (totalNoEntry >= br) status = 'paid';
          else                         status = 'debt';
          allRows.push([
            loc, dir, typ,
            g.group, g.teacher, ch.name,
            fs, fv, fe, total, bs, bd, br,
            status, monthName, updateStr,
            ch.contractDate || ''
          ]);
        });
      });
    } catch(e) {
      errors.push(loc + ': ' + e.message);
      Logger.log('ERROR ' + loc + ': ' + e.message);
    }
  }

  paySheet.clearContents();
  writePaymentsHeader(paySheet);
  if (allRows.length > 0) {
    paySheet.getRange(2, 1, allRows.length, 17).setValues(allRows);
  }
  Logger.log('Done: ' + allRows.length + ' rows, ' + errors.length + ' errors');
  return {ok:true, rows:allRows.length, errors:errors, month:monthName, updated:updateStr};
}

function detectContractDateCol(data) {
  for (var r = 0; r < Math.min(5, data.length); r++) {
    for (var c = 0; c < data[r].length; c++) {
      var cell = trim(String(data[r][c] || '')).toLowerCase();
      if (cell.indexOf('дата договору') >= 0) return c;
    }
  }
  return -1;
}

function detectCurrentMonthCol(rows, curJSMonth) {
  for (var r = 0; r < Math.min(3, rows.length); r++) {
    for (var c = 1; c < rows[r].length; c++) {
      var cell = String(rows[r][c] || '').toLowerCase().trim();
      for (var mi = 0; mi < MONTHS_UA.length; mi++) {
        if (cell === MONTHS_UA[mi] && MONTHS_JS[mi] === curJSMonth) {
          return c;
        }
      }
    }
  }
  for (var r = 0; r < Math.min(3, rows.length); r++) {
    for (var c = 1; c < rows[r].length; c++) {
      var cell = String(rows[r][c] || '').toLowerCase().trim();
      for (var mi = 0; mi < MONTHS_UA.length; mi++) {
        if (cell.indexOf(MONTHS_UA[mi]) >= 0 && MONTHS_JS[mi] === curJSMonth) {
          return c;
        }
      }
    }
  }
  var col = 1 + curJSMonth * 5;
  return col;
}

function parsePaymentSheet(data, monthCol, contractCol) {
  var DATA_START = 3;
  var groups = [];
  var curGroup = null;
  for (var r = DATA_START; r < data.length; r++) {
    var row = data[r];
    var nameCell = trim(String(row[0] || ''));
    if (!nameCell) continue;
    if (isGroupHeaderRow(row, monthCol)) {
      var firstSpace = nameCell.search(/\s/);
      var teacher = firstSpace > 0 ? nameCell.slice(firstSpace).trim() : '';
      var groupName = normalizeGroupName(nameCell);
      var groupKey = groupName + (teacher ? ' ' + teacher : '');
      curGroup = {group: groupKey, teacher: teacher, children: []};
      groups.push(curGroup);
    } else {
      if (!curGroup) {
        curGroup = {group:'(без групи)', teacher:'', children:[]};
        groups.push(curGroup);
      }
      var fs = toNum(row[monthCol]);
      var fv = toNum(row[monthCol + 1]);
      var fe = toNum(row[monthCol + 2]);
      var bd = toNum(row[monthCol + 3]);
      var bs = toNum(row[monthCol + 4]);
      var cd = (contractCol >= 0) ? parseDateDMY(row[contractCol]) : '';
      curGroup.children.push({
        name: nameCell,
        factStudy: fs, factEntry: fv, factExtra: fe,
        budExtra: bd, budStudy: bs,
        contractDate: cd
      });
    }
  }
  return groups.filter(function(g){ return g.children.length > 0; });
}

function createDailyTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    var fn = triggers[i].getHandlerFunction();
    if (fn === 'aggregatePayments' || fn === 'aggregatePaymentsYearly') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('aggregatePayments')
    .timeBased().everyDays(1).atHour(6)
    .inTimezone('Europe/Kiev').create();
  ScriptApp.newTrigger('aggregatePaymentsYearly')
    .timeBased().everyDays(1).atHour(7)
    .inTimezone('Europe/Kiev').create();
}

function trim(s) { return String(s || '').trim(); }
function toNum(v) {
  if (v === '' || v === null || v === undefined) return 0;
  var n = parseFloat(String(v).replace(',', '.'));
  return isNaN(n) ? 0 : n;
}

function parseDateDMY(v) {
  if (!v && v !== 0) return '';
  if (v instanceof Date) {
    if (isNaN(v.getTime())) return '';
    return Utilities.formatDate(v, 'Europe/Kiev', 'yyyy-MM-dd');
  }
  var s = trim(String(v));
  if (!s) return '';
  var sep = s.indexOf('.') >= 0 ? '\\.' : s.indexOf('/') >= 0 ? '\\/' : s.indexOf('|') >= 0 ? '\\|' : null;
  if (!sep) return '';
  var m3 = s.match(new RegExp('^(\\d{1,2})' + sep + '(\\d{1,2})' + sep + '(\\d{2}|\\d{4})$'));
  if (m3) {
    var day   = ('0' + m3[1]).slice(-2);
    var month = ('0' + m3[2]).slice(-2);
    var year  = m3[3].length === 2 ? '20' + m3[3] : m3[3];
    return year + '-' + month + '-' + day;
  }
  var m2 = s.match(new RegExp('^(\\d{1,2})' + sep + '(\\d{2}|\\d{4})$'));
  if (m2) {
    var month = ('0' + m2[1]).slice(-2);
    var year  = m2[2].length === 2 ? '20' + m2[2] : m2[2];
    return year + '-' + month + '-01';
  }
  return '';
}

function formatDate(d) {
  return Utilities.formatDate(d, 'Europe/Kiev', 'dd.MM.yyyy HH:mm');
}

function getMonthDisplayName(jsMonth) {
  var idx = MONTHS_JS.indexOf(jsMonth);
  return idx >= 0 ? MONTHS_DISPLAY[idx] : String(jsMonth + 1);
}

function writeYearlyHeader(sheet) {
  sheet.clearContents();
  var hdr = ['Локація','Напрямок','Тип','Група','Вихователь',"Ім'я дитини"];
  MONTHS_CAL.forEach(function(m) {
    hdr.push(m+'-Факт-навч', m+'-Факт-доп', m+'-Бюджет-навч', m+'-Бюджет-доп', m+'-Статус');
  });
  hdr.push('Факт-Рік','Бюджет-Рік','Борг-Рік','Зібрано-На-Сьогодні','Оновлено');
  sheet.appendRow(hdr);
  sheet.setFrozenRows(1);
}

function aggregatePaymentsYearly() {
  var configSS    = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  var configSheet = configSS.getSheets()[0];
  var configData  = configSheet.getDataRange().getValues();
  var crmSS    = getCRMSpreadsheet();
  var yearSheet = crmSS.getSheetByName(SHEET_YEARLY);
  if (!yearSheet) yearSheet = crmSS.insertSheet(SHEET_YEARLY);
  var now         = new Date();
  var curJSMonth  = now.getMonth();
  var updateStr   = formatDate(now);
  var allRows     = [];
  var errors      = [];
  for (var r = 1; r < configData.length; r++) {
    var cfgRow    = configData[r];
    var dir       = trim(cfgRow[0]);
    var typ       = trim(cfgRow[1]);
    var loc       = trim(cfgRow[2]);
    var sheetId   = trim(cfgRow[3]);
    var sheetName = trim(cfgRow[4]) || 'Payment';
    if (!loc || !sheetId) continue;
    try {
      var ss           = SpreadsheetApp.openById(sheetId);
      var paymentSheet = ss.getSheetByName(sheetName);
      if (!paymentSheet) paymentSheet = ss.getSheets()[0];
      var data = paymentSheet.getDataRange().getValues();
      var curMonthCol  = detectCurrentMonthCol(data, curJSMonth);
      var contractCol  = detectContractDateCol(data);
      var groups       = parsePaymentSheet(data, curMonthCol, contractCol);
      var nameToRow = {};
      for (var ri = 3; ri < data.length; ri++) {
        var nc = trim(String(data[ri][0] || ''));
        if (nc && !isGroupHeaderRow(data[ri], 1)) {
          nameToRow[nc] = ri;
        }
      }
      groups.forEach(function(g) {
        g.children.forEach(function(ch) {
          var rowIdx  = nameToRow[ch.name];
          var rowData = (rowIdx !== undefined) ? data[rowIdx] : null;
          var rowOut  = [loc, dir, typ, g.group, g.teacher, ch.name];
          var factYear  = 0;
          var budYear   = 0;
          var factToday = 0;
          for (var mi = 0; mi < 12; mi++) {
            var col = 1 + mi * 5;
            var fs  = rowData ? toNum(rowData[col])     : 0;
            var fe  = rowData ? toNum(rowData[col + 2]) : 0;
            var be  = rowData ? toNum(rowData[col + 3]) : 0;
            var bs  = rowData ? toNum(rowData[col + 4]) : 0;
            var totalNoEntry = fs + fe;
            var budget       = bs + be;
            var mStatus;
            if (budget === 0 && totalNoEntry === 0)     mStatus = 'unknown';
            else if (totalNoEntry === 0 && budget > 0)  mStatus = 'nopay';
            else if (totalNoEntry > budget)             mStatus = 'over';
            else if (totalNoEntry >= budget)            mStatus = 'paid';
            else                                        mStatus = 'debt';
            rowOut.push(fs, fe, bs, be, mStatus);
            factYear  += totalNoEntry;
            budYear   += budget;
            if (mi <= curJSMonth) factToday += totalNoEntry;
          }
          var debtYear = budYear > factYear ? budYear - factYear : 0;
          rowOut.push(factYear, budYear, debtYear, factToday, updateStr);
          allRows.push(rowOut);
        });
      });
    } catch(e) {
      errors.push(loc + ': ' + e.message);
    }
  }
  yearSheet.clearContents();
  writeYearlyHeader(yearSheet);
  var NUM_COLS = 6 + 12 * 5 + 5;
  if (allRows.length > 0) {
    yearSheet.getRange(2, 1, allRows.length, NUM_COLS).setValues(allRows);
  }
  return {ok:true, rows:allRows.length, errors:errors, updated:updateStr};
}

function getPaymentsYearly() {
  var ss    = getCRMSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_YEARLY);
  if (!sheet) return {ok:false, error:'Sheet Оплати-Рік not found. Run aggregatePaymentsYearly() first.'};
  var vals = sheet.getDataRange().getValues();
  if (vals.length < 2) return {ok:true, data:[]};
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
  return {ok:true, data:rows};
}

function getAttendance(e) {
  var params  = e ? (e.parameter || {}) : {};
  var loc     = trim(params.loc  || '');
  var from    = trim(params.from || '');
  var to      = trim(params.to   || '');
  var ss      = getCRMSpreadsheet();
  var sheet   = ss.getSheetByName(SHEET_ATTENDANCE);
  if (!sheet) return {ok:true, data:[]};
  var vals = sheet.getDataRange().getValues();
  if (vals.length < 2) return {ok:true, data:[]};
  var hdrs = vals[0].map(String);
  var rows = [];
  for (var r = 1; r < vals.length; r++) {
    var obj = {};
    for (var c = 0; c < hdrs.length; c++) obj[hdrs[c]] = String(vals[r][c] || '');
    var d = obj['Дата'] || '';
    if (!d) continue;
    if (from && d < from) continue;
    if (to   && d > to)   continue;
    if (loc  && trim(obj['Локація']) !== loc) continue;
    rows.push(obj);
  }
  return {ok:true, data:rows};
}

function saveAttendance(body) {
  var records = body.records || [];
  if (!records.length) return {ok:true, saved:0};
  var ss    = getCRMSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_ATTENDANCE);
  if (!sheet) { sheet = ss.insertSheet(SHEET_ATTENDANCE); writeAttendanceHeader(sheet); }
  var vals = sheet.getDataRange().getValues();
  var now  = formatDate(new Date());
  var saved = 0;

  records.forEach(function(rec) {
    var date    = trim(String(rec.date    || ''));
    var childId = trim(String(rec.childId || ''));
    if (!date || !childId) return;
    var row = [date, childId, rec.childName||'', rec.loc||'', rec.group||'', rec.status||'', rec.updatedBy||'', now];
    var written = false;
    for (var r = 1; r < vals.length; r++) {
      if (String(vals[r][0]) === date && String(vals[r][1]) === childId) {
        sheet.getRange(r+1, 1, 1, row.length).setValues([row]);
        vals[r] = row;
        saved++;
        written = true;
        break;
      }
    }
    if (!written) {
      sheet.appendRow(row);
      vals.push(row);
      saved++;
    }
    mirrorAttendanceToNurseSheet(rec.loc||'', rec.childName||'', date, rec.status||'');
  });

  return {ok:true, saved:saved};
}

function getHealthRecords(e) {
  var params  = e ? (e.parameter || {}) : {};
  var childId = trim(params.childId || '');
  if (!childId) return {ok:false, error:'Missing childId'};
  var ss    = getCRMSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_HEALTH);
  if (!sheet) return {ok:true, data:[]};
  var vals = sheet.getDataRange().getValues();
  if (vals.length < 2) return {ok:true, data:[]};
  var hdrs = vals[0].map(String);
  var rows = [];
  for (var r = 1; r < vals.length; r++) {
    var obj = {};
    for (var c = 0; c < hdrs.length; c++) obj[hdrs[c]] = String(vals[r][c] || '');
    if (trim(obj['ID дитини']) === childId) rows.push(obj);
  }
  return {ok:true, data:rows};
}

function saveHealthRecord(body) {
  var rec = body.record;
  if (!rec || !rec.childId) return {ok:false, error:'Missing record or childId'};
  var ss    = getCRMSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_HEALTH);
  if (!sheet) { sheet = ss.insertSheet(SHEET_HEALTH); writeHealthHeader(sheet); }
  var now = formatDate(new Date());
  var id = trim(String(rec.id || '')) || ('h_' + new Date().getTime());
  var vals = sheet.getDataRange().getValues();
  for (var r = 1; r < vals.length; r++) {
    if (String(vals[r][0]) === id) return {ok:true, id:id, action:'exists'};
  }
  sheet.appendRow([id, rec.childId, rec.date||'', rec.type||'note', rec.text||rec.desc||'', rec.createdBy||'', now]);
  return {ok:true, id:id, action:'created'};
}

function deleteHealthRecord(body) {
  var recordId = trim(String(body.recordId || ''));
  if (!recordId) return {ok:false, error:'Missing recordId'};
  var ss    = getCRMSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_HEALTH);
  if (!sheet) return {ok:false, error:'Sheet not found'};
  var vals = sheet.getDataRange().getValues();
  for (var r = vals.length-1; r >= 1; r--) {
    if (String(vals[r][0]) === recordId) {
      sheet.deleteRow(r+1);
      return {ok:true};
    }
  }
  return {ok:false, error:'Record not found'};
}

function parseAbsencePeriod(str, refYear) {
  var result = (function() {

    if (!str) return null;
    if (str instanceof Date) {
      if (isNaN(str.getTime())) return null;
      return parseAbsencePeriod(pad2(str.getMonth() + 1) + '/' + str.getFullYear(), refYear);
    }

    var s = trim(String(str)).toLowerCase();
    if (!s || s === '-' || s === 'по') return null;

    var nowMon = new Date().getMonth() + 1;

    var UA_MONTHS = {
      'січ':1,    'лют':2,    'бер':3,        'квіт':4,    'трав':5,    'черв':6,
      'лип':7,    'серп':8,   'вер':9,         'жовт':10,   'лист':11,   'груд':12,
      'січень':1, 'лютий':2,  'березень':3,    'квітень':4, 'травень':5, 'червень':6,
      'липень':7, 'серпень':8,'вересень':9,    'жовтень':10,'листопад':11,'грудень':12
    };

    function yearFor(mon) {
      return (mon >= nowMon) ? (refYear - 1) : refYear;
    }

    function syntheticWeek(yr, mon) {
      var d = new Date(yr, mon - 1, 1);
      while (d.getDay() === 0 || d.getDay() === 6) { d.setDate(d.getDate() + 1); }
      var fromD = new Date(d);
      var toD   = new Date(d); toD.setDate(toD.getDate() + 4);
      return {
        from: fromD.getFullYear() + '-' + pad2(fromD.getMonth()+1) + '-' + pad2(fromD.getDate()),
        to:   toD.getFullYear()   + '-' + pad2(toD.getMonth()+1)   + '-' + pad2(toD.getDate()),
        _synthetic: true, _originalRaw: str
      };
    }

    function findUAMonth(text) {
      var t = text.trim().toLowerCase();
      var keys = Object.keys(UA_MONTHS).sort(function(a, b) { return b.length - a.length; });
      for (var ki = 0; ki < keys.length; ki++) {
        if (t.indexOf(keys[ki]) >= 0) return UA_MONTHS[keys[ki]];
      }
      return null;
    }

    var n = s.replace(/,/g, '.').replace(/\s*по\s*/g, '-').replace(/\s+/g, '');
    var m;

    // БЛОК 1а: "01.09.2024-15.09.2024"
    m = n.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})[-–](\d{1,2})\.(\d{1,2})\.(\d{4})$/);
    if (m) {
      return {
        from: m[3] + '-' + pad2(m[2]) + '-' + pad2(m[1]),
        to:   m[6] + '-' + pad2(m[5]) + '-' + pad2(m[4])
      };
    }

    // БЛОК 1б: "02.01-08.01.2025"
    m = n.match(/^(\d{1,2})\.(\d{1,2})[-–](\d{1,2})\.(\d{1,2})\.(\d{2}|\d{4})$/);
    if (m) {
      var yr1b = m[5].length === 2 ? 2000 + (+m[5]) : +m[5];
      return {
        from: yr1b + '-' + pad2(m[2]) + '-' + pad2(m[1]),
        to:   yr1b + '-' + pad2(m[4]) + '-' + pad2(m[3])
      };
    }

    // БЛОК 2а: "09.12-15.12", "30.06-18.07"
    m = n.match(/^(\d{1,2})\.(\d{1,2})[-–](\d{1,2})\.(\d{1,2})$/);
    if (m) {
      return {
        from: yearFor(+m[2]) + '-' + pad2(m[2]) + '-' + pad2(m[1]),
        to:   yearFor(+m[4]) + '-' + pad2(m[4]) + '-' + pad2(m[3])
      };
    }

    // БЛОК 2б: "30.06-18-07" (нормалізований "30.06по 18-07")
    m = n.match(/^(\d{1,2})\.(\d{1,2})[-–](\d{1,2})[-–](\d{1,2})$/);
    if (m) {
      return {
        from: yearFor(+m[2]) + '-' + pad2(m[2]) + '-' + pad2(m[1]),
        to:   yearFor(+m[4]) + '-' + pad2(m[4]) + '-' + pad2(m[3])
      };
    }

    // БЛОК 3а: "15-20.01"
    m = n.match(/^(\d{1,2})[-–](\d{1,2})\.(\d{1,2})$/);
    if (m) {
      var mon3 = +m[3];
      return {
        from: yearFor(mon3) + '-' + pad2(m[3]) + '-' + pad2(m[1]),
        to:   yearFor(mon3) + '-' + pad2(m[3]) + '-' + pad2(m[2])
      };
    }

    // БЛОК 3б: "1-14.09.25", "1-14.09.2025"
    m = n.match(/^(\d{1,2})[-–](\d{1,2})\.(\d{1,2})\.(\d{2}|\d{4})$/);
    if (m) {
      var yr3b = m[4].length === 2 ? 2000 + (+m[4]) : +m[4];
      return {
        from: yr3b + '-' + pad2(m[3]) + '-' + pad2(m[1]),
        to:   yr3b + '-' + pad2(m[3]) + '-' + pad2(m[2])
      };
    }

    // БЛОК 4: "10/25", "10|25", "10/2025", "8.2025"
    m = n.match(/^(\d{1,2})[\/|.](\d{2}|\d{4})$/);
    if (m) {
      var mon4 = +m[1];
      var yr4  = m[2].length === 2 ? 2000 + (+m[2]) : +m[2];
      if (mon4 >= 1 && mon4 <= 12) {
        return syntheticWeek(yr4, mon4);
      }
    }

    // БЛОК 5: "жовт25", "лист25", "серп2025"
    m = s.match(/^([а-яіїє']+?)\.?\s*(\d{2}|\d{4})$/);
    if (m) {
      var mon5 = UA_MONTHS[m[1].trim()];
      if (mon5) {
        var yr5 = m[2].length === 2 ? 2000 + (+m[2]) : +m[2];
        return syntheticWeek(yr5, mon5);
      }
    }

    // БЛОК 6: "жовтень 2025"
    m = s.match(/^([а-яіїє']+)\s+(\d{4})$/);
    if (m) {
      var mon6 = UA_MONTHS[m[1].trim()];
      if (mon6) {
        return syntheticWeek(+m[2], mon6);
      }
    }

    // БЛОК 7: "10 днів серпня", "5 днів жовтня"
    m = s.match(/^(\d+)\s*дн[іияь\.]+\s*([а-яіїє'\s]+)/);
    if (m) {
      var days7 = +m[1];
      var mon7  = findUAMonth(m[2]);
      if (mon7) {
        var yr7  = yearFor(mon7);
        var d7   = new Date(yr7, mon7 - 1, 1);
        while (d7.getDay() === 0 || d7.getDay() === 6) { d7.setDate(d7.getDate() + 1); }
        var toD7 = new Date(d7); toD7.setDate(toD7.getDate() + days7 - 1);
        return {
          from: d7.getFullYear()   + '-' + pad2(d7.getMonth()+1)   + '-' + pad2(d7.getDate()),
          to:   toD7.getFullYear() + '-' + pad2(toD7.getMonth()+1) + '-' + pad2(toD7.getDate()),
          _synthetic: true, _originalRaw: str
        };
      }
    }

    // БЛОК 8: "1 т серпень", "2 т липні", "3 тижн вересень"
    m = s.match(/^(\d+)\s*т[иі]?[жщ]?[нь]?\.?\s*([а-яіїє'\s]+)/);
    if (m) {
      var weeks8 = +m[1];
      var mon8   = findUAMonth(m[2]);
      if (mon8) {
        var yr8  = yearFor(mon8);
        var d8   = new Date(yr8, mon8 - 1, 1);
        while (d8.getDay() === 0 || d8.getDay() === 6) { d8.setDate(d8.getDate() + 1); }
        var toD8 = new Date(d8); toD8.setDate(toD8.getDate() + weeks8 * 5 - 1);
        return {
          from: d8.getFullYear()   + '-' + pad2(d8.getMonth()+1)   + '-' + pad2(d8.getDate()),
          to:   toD8.getFullYear() + '-' + pad2(toD8.getMonth()+1) + '-' + pad2(toD8.getDate()),
          _synthetic: true, _originalRaw: str
        };
      }
    }

    // БЛОК 9: тільки UA місяць "жовтень", "серп", "КВІТЕНЬ"
    var mon9 = UA_MONTHS[s.trim()];
    if (mon9) {
      return syntheticWeek(yearFor(mon9), mon9);
    }

    // БЛОК 10: вільний текст що містить назву місяця
    var mon10 = findUAMonth(s);
    if (mon10) {
      return syntheticWeek(yearFor(mon10), mon10);
    }

    return null;

  })();

  if (!result && str && !(str instanceof Date) && String(str).trim() !== '-') {
    var raw        = String(str);
    var normalized = raw.toLowerCase().replace(/,/g, '.').replace(/\s*по\s*/g, '-').replace(/\s+/g, '');
    Logger.log('[parseAbs] FAIL str="' + raw + '" normalized="' + normalized + '"');
  }

  return result;
}

function pad2(n) { return ('0' + n).slice(-2); }

function writeAbsenceToPayment(body) {
  try {
    var childName = trim(body.childName || '');
    var group     = trim(body.group    || '');
    var loc       = trim(body.loc      || '');
    var slots     = body.slots || [];
    if (!childName || !loc || !slots.length) return {ok:false, error:'Missing params'};
    var configSS    = SpreadsheetApp.openById(CONFIG_SHEET_ID);
    var configSheet = configSS.getSheets()[0];
    var configData  = configSheet.getDataRange().getValues();
    var sheetId = null; var sheetName = 'Payment';
    for (var r = 1; r < configData.length; r++) {
      if (trim(String(configData[r][2] || '')) === loc) {
        sheetId   = trim(String(configData[r][3] || ''));
        sheetName = trim(String(configData[r][4] || '')) || 'Payment';
        break;
      }
    }
    if (!sheetId) return {ok:false, error:'Location not found: ' + loc};
    var paymentSS    = SpreadsheetApp.openById(sheetId);
    var paymentSheet = paymentSS.getSheetByName(sheetName) || paymentSS.getSheets()[0];
    var data         = paymentSheet.getDataRange().getValues();
    var absCols = detectAbsenceCols(data);
    if (absCols[0] === null) return {ok:false, error:'Absence columns not found: ' + loc};
    var norm = function(s){ return String(s||'').trim().toLowerCase().replace(/\s+/g,' '); };
    var normName   = norm(childName);
    var nameColIdx = 0;  // ПІБ завжди у першій колонці Payment-файлу

    // Збираємо всі рядки з точним збігом ПІБ (для виявлення тезок)
    var matchRows = [];
    for (var row = 3; row < data.length; row++) {
      var rowName = norm(data[row][nameColIdx]);
      if (rowName === normName) matchRows.push(row);
    }
    if (matchRows.length === 0) return {ok:false, error:'Child not found: ' + childName};
    if (matchRows.length > 1) {
      Logger.log('writeAbsenceToPayment WARN: знайдено ' + matchRows.length +
        ' рядків з ім\'ям "' + childName + '" у ' + loc +
        ' (рядки: ' + matchRows.map(function(r){ return r+1; }).join(', ') + ') — беремо перший');
    }
    var targetRow = matchRows[0];
    var writtenTo = [];
    var slotIdx   = 0;
    for (var ci = 0; ci < absCols.length && slotIdx < slots.length; ci++) {
      if (absCols[ci] === null) continue;
      var existing = trim(String(data[targetRow][absCols[ci]] || ''));
      if (!existing) {
        paymentSheet.getRange(targetRow + 1, absCols[ci] + 1).setValue(slots[slotIdx]);
        writtenTo.push({weekNum: ci + 1, value: slots[slotIdx]});
        slotIdx++;
      }
    }
    if (slotIdx === 0) return {ok:false, error:'All absence slots already filled for ' + childName};
    return {ok:true, writtenTo: writtenTo};
  } catch(err) {
    return {ok:false, error: err.message || String(err)};
  }
}

function detectAbsenceCols(data) {
  var labels = ['1 тиждень', '2 тиждень', '3 тиждень', '4 тиждень'];
  var cols = [null, null, null, null];
  for (var r = 0; r < Math.min(5, data.length); r++) {
    for (var c = 0; c < data[r].length; c++) {
      var cell = trim(String(data[r][c] || '')).toLowerCase();
      for (var li = 0; li < labels.length; li++) {
        if (cols[li] === null && cell.indexOf(labels[li].toLowerCase()) >= 0) {
          cols[li] = c;
        }
      }
    }
  }
  return cols;
}

// ── DRY-RUN ІМПОРТ ВІДПУСТОК З PAYMENT → CRM ────────────────────────────────
// ТІЛЬКИ читає і логує — нічого не пише!
// locFilter — необов'язковий рядок. Якщо порожній — обробляє всі локації крім шкіл.
//
// refYear: береться з поточної дати (new Date().getFullYear()).
//   parseAbsencePeriod сам вирішує рік: якщо місяць у рядку >= поточного → рік-1,
//   інакше → поточний рік. При запуску у квітні 2026:
//     "27.10-31.10" → 2025-10-27/2025-10-31  (жовтень > квітень → 2025)
//     "20.04-24.04" → 2026-04-20/2026-04-24  (квітень == квітень → 2026)
//
// Dedupe: завантажуємо всіх CRM-клієнтів один раз, будуємо map
//   norm(name)+'|'+norm(loc) → [{from,to}].
//   Якщо parsed.from+parsed.to вже є у CRM — рахуємо як duplicate, не імпортуємо.
//
// Запуск з IDE: _runDryRunImport() або _runDryRunImportOsokory()
// ── СПІЛЬНИЙ ХЕЛПЕР: завантажує CRM-клієнтів одним разом ─────────────────────
// Повертає map: norm(name)+'|'+norm(loc) →
//   {id, absences:[{from,to,...}], group, teacher, contractDate, monthlyFee, ...}
// Якщо два рядки з однаковим name+loc — перший виграє (WARN у лозі).
function _loadCRMClientsMap(norm) {
  var crmSS    = getCRMSpreadsheet();
  var crmSheet = crmSS.getSheetByName(SHEET_CLIENTS);
  var map      = {};

  Logger.log('[loadCRMMap] SS id: ' + crmSS.getId() + ' | SS name: ' + crmSS.getName());
  Logger.log('[loadCRMMap] Sheet "Клієнти" found: ' + !!crmSheet);

  if (!crmSheet) return map;
  var crmData = crmSheet.getDataRange().getValues();
  Logger.log('[loadCRMMap] sheet rows (incl header): ' + crmData.length);
  if (crmData.length < 2) return map;

  var hdrs    = crmData[0].map(String);
  Logger.log('[loadCRMMap] headers: ' + JSON.stringify(hdrs));

  var colId   = hdrs.indexOf('ID');              if (colId   < 0) colId   = 0;
  var colName = hdrs.indexOf('ПІБ дитини');      if (colName < 0) colName = 1;
  var colLoc  = hdrs.indexOf('Локація');         if (colLoc  < 0) colLoc  = 2;
  var colGrp  = hdrs.indexOf('Група');           if (colGrp  < 0) colGrp  = 3;
  var colTch  = hdrs.indexOf('Вихователь');      if (colTch  < 0) colTch  = 4;
  var colCD   = hdrs.indexOf('Дата договору');   if (colCD   < 0) colCD   = 10;
  var colCT   = hdrs.indexOf('Тип договору');    if (colCT   < 0) colCT   = 11;
  var colFee  = hdrs.indexOf('Сума договору');   if (colFee  < 0) colFee  = 12;
  var colAbs  = hdrs.indexOf('Відсутності (JSON)');
  var colNot  = hdrs.indexOf('Нотатки');
  Logger.log('[loadCRMMap] colId=' + colId + ' colName=' + colName + ' colLoc=' + colLoc + ' colAbs=' + colAbs);
  Logger.log('[loadCRMMap] row[1] raw: ' + JSON.stringify(crmData[1]));

  for (var ri = 1; ri < crmData.length; ri++) {
    var rName = norm(crmData[ri][colName] || '');
    var rLoc  = norm(crmData[ri][colLoc]  || '');

    if (String(crmData[ri][colName] || '').indexOf('Рибак') !== -1) {
      Logger.log('[loadCRMMap] Рибак row ri=' + ri + ': ' + JSON.stringify(crmData[ri]));
      Logger.log('[loadCRMMap] Рибак rName="' + rName + '" rLoc="' + rLoc + '"');
    }

    if (!rName) continue;
    var key = rName + '|' + rLoc;
    if (map[key]) {
      Logger.log('_loadCRMClientsMap WARN: тезка "' + rName + '" у "' + rLoc + '" — ігноруємо рядок ' + (ri+1));
      continue;
    }
    var absArr = [];
    if (colAbs >= 0) { try { absArr = JSON.parse(String(crmData[ri][colAbs] || '[]')); } catch(e2) {} }
    map[key] = {
      id:           String(crmData[ri][colId]  || ''),
      name:         String(crmData[ri][colName] || ''),
      loc:          String(crmData[ri][colLoc]  || ''),
      group:        String(crmData[ri][colGrp]  || ''),
      teacher:      String(crmData[ri][colTch]  || ''),
      contractDate: String(crmData[ri][colCD]   || ''),
      contractType: String(crmData[ri][colCT]   || 'standard'),
      monthlyFee:   toNum(crmData[ri][colFee]),
      notes:        colNot >= 0 ? String(crmData[ri][colNot] || '') : '',
      absences:     absArr
    };
  }

  Logger.log('[loadCRMMap] map keys count: ' + Object.keys(map).length);
  Logger.log('[loadCRMMap] map keys: ' + JSON.stringify(Object.keys(map)));
  return map;
}

// ── ХЕЛПЕР: рахує робочі дні між двома YYYY-MM-DD ──────────────────────────
function _countWorkDays(fromStr, toStr) {
  if (!fromStr || !toStr) return 0;
  var f = new Date(fromStr); var t = new Date(toStr);
  if (isNaN(f.getTime()) || isNaN(t.getTime()) || t < f) return 0;
  var n = 0; var cur = new Date(f.getTime());
  while (cur <= t) { var d = cur.getDay(); if (d !== 0 && d !== 6) n++; cur.setDate(cur.getDate()+1); }
  return n;
}

// ── ХЕЛПЕР: будує об'єкт відпустки для імпорту ─────────────────────────────
function _makeImportAbsence(parsed, rawSlot) {
  var iso = new Date().toISOString();
  var id  = 'abs_import_' + Date.now() + '_' + Math.random().toString(36).slice(2,7);
  if (parsed) {
    var wd   = _countWorkDays(parsed.from, parsed.to);
    var w    = Math.min(4, Math.ceil(Math.max(0, wd) / 5));
    var note = parsed._synthetic
      ? 'імпорт з Payment: "' + (parsed._originalRaw || rawSlot) + '" (1 тиждень у цьому місяці, точні дати не збережено)'
      : 'імпорт з Payment';
    return {
      id: id, type: 'vacation', from: parsed.from, to: parsed.to,
      workDays: wd, weeks: w,
      monthsBreakdown: [], totalPct: 0, totalAmount: 0,
      status: 'done',
      statusHistory: [{status:'done', at:iso, by:'import'}],
      rejectReason: '', note: note,
      createdBy: 'import', createdAt: iso
    };
  } else {
    // placeholder — формат зовсім не розпізнано
    return {
      id: id, type: 'vacation', from: null, to: null,
      workDays: 5, weeks: 1,
      monthsBreakdown: [], totalPct: 0, totalAmount: 0,
      status: 'done',
      statusHistory: [{status:'done', at:iso, by:'import'}],
      rejectReason: '',
      note: 'імпорт з Payment: "' + rawSlot + '" (формат не розпізнано, прийнято як 1 тиждень)',
      createdBy: 'import', createdAt: iso
    };
  }
}

// ── DRY-RUN: читає, рахує, нічого не пише ────────────────────────────────────
function dryRunImportAbsences(locFilter) {
  var SCHOOL_LOCS_SKIP = ['Школа Осокорки', 'Школа 228', 'Онлайн школа'];
  var refYear = new Date().getFullYear();
  var norm    = function(s){ return String(s||'').trim().toLowerCase().replace(/\s+/g,' '); };

  var crmMap = _loadCRMClientsMap(norm);
  Logger.log('CRM: завантажено ' + Object.keys(crmMap).length + ' клієнтів для dedupe');

  var configSS    = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  var configSheet = configSS.getSheets()[0];
  var configData  = configSheet.getDataRange().getValues();

  var totalStats = {
    locations:0, totalSlotsProcessed:0,
    wouldCreate:0, wouldCreateExact:0, wouldCreateSynthetic:0,
    wouldPlaceholder:0, duplicates:0,
    wouldCreateNewClient:0, wouldSkipNoAbsence:0
  };
  var byLocation     = {};
  var unparsedCounts = {};

  Logger.log('=== dryRunImportAbsences START refYear=' + refYear +
    (locFilter ? ' loc=' + locFilter : ' (всі локації)') + ' ===');

  for (var r = 1; r < configData.length; r++) {
    var loc       = trim(configData[r][2]);
    var sheetId   = trim(configData[r][3]);
    var sheetName = trim(configData[r][4]) || 'Payment';
    if (!loc || !sheetId) continue;
    if (SCHOOL_LOCS_SKIP.indexOf(loc) >= 0) continue;
    if (locFilter && loc !== locFilter) continue;

    totalStats.locations++;
    var locStat = {created:0, placeholder:0, duplicates:0, newClients:0, skipped:0};

    try {
      var ss           = SpreadsheetApp.openById(sheetId);
      var paymentSheet = ss.getSheetByName(sheetName) || ss.getSheets()[0];
      var data         = paymentSheet.getDataRange().getValues();

      var absCols = detectAbsenceCols(data);
      if (absCols[0] === null) {
        Logger.log(loc + ': WARN — колонки BL/BM/BN/BO не знайдено, пропускаємо');
        byLocation[loc] = locStat;
        continue;
      }

      var DATA_START = 3;
      for (var row = DATA_START; row < data.length; row++) {
        var nameCell = trim(String(data[row][0] || ''));
        if (!nameCell) continue;
        if (isGroupHeaderRow(data[row], 1)) continue;

        // Є хоча б один непорожній слот BL-BO?
        var hasAnySlot = false;
        for (var si = 0; si < absCols.length; si++) {
          if (absCols[si] !== null && trim(String(data[row][absCols[si]] || ''))) { hasAnySlot = true; break; }
        }
        if (!hasAnySlot) { locStat.skipped++; totalStats.wouldSkipNoAbsence++; continue; }

        // Шукаємо у CRM. Якщо немає → нова картка, всі слоти йдуть як нові
        var crmKey  = norm(nameCell) + '|' + norm(loc);
        var isNew   = !crmMap.hasOwnProperty(crmKey);
        var existingPairs = {};
        if (isNew) {
          locStat.newClients++;
          totalStats.wouldCreateNewClient++;
        } else {
          crmMap[crmKey].absences.forEach(function(a){
            if (a.from && a.to) existingPairs[a.from + '|' + a.to] = true;
          });
        }

        // Перебираємо 4 слоти
        for (var si2 = 0; si2 < absCols.length; si2++) {
          if (absCols[si2] === null) continue;
          var slot = trim(String(data[row][absCols[si2]] || ''));
          if (!slot) continue;
          totalStats.totalSlotsProcessed++;
          var parsed = parseAbsencePeriod(slot, refYear);
          if (parsed) {
            var pairKey = parsed.from + '|' + parsed.to;
            if (!isNew && existingPairs[pairKey]) {
              locStat.duplicates++; totalStats.duplicates++;
            } else {
              locStat.created++; totalStats.wouldCreate++;
              if (parsed._synthetic) { totalStats.wouldCreateSynthetic++; }
              else                   { totalStats.wouldCreateExact++;      }
              existingPairs[pairKey] = true;
            }
          } else {
            unparsedCounts[slot] = (unparsedCounts[slot] || 0) + 1;
            locStat.placeholder++; totalStats.wouldPlaceholder++;
          }
        }
      }

      Logger.log(loc + ': wouldCreate=' + locStat.created +
        ', placeholder=' + locStat.placeholder +
        ', duplicates='  + locStat.duplicates +
        ', newClients='  + locStat.newClients +
        ', skipped='     + locStat.skipped);

    } catch(err) {
      Logger.log(loc + ': ERROR — ' + (err.message || String(err)));
    }

    byLocation[loc] = {
      created: locStat.created, placeholder: locStat.placeholder,
      duplicates: locStat.duplicates, newClients: locStat.newClients
    };
  }

  var unparsedKeys = Object.keys(unparsedCounts);
  unparsedKeys.sort(function(a,b){ return unparsedCounts[b]-unparsedCounts[a]; });
  var unparsedSamples = unparsedKeys.slice(0,20);

  Logger.log('');
  Logger.log('=== ПІДСУМОК ===');
  Logger.log('Локацій:              ' + totalStats.locations);
  Logger.log('Слотів оброблено:     ' + totalStats.totalSlotsProcessed);
  Logger.log('wouldCreate:          ' + totalStats.wouldCreate +
    ' (exact=' + totalStats.wouldCreateExact + ', synthetic=' + totalStats.wouldCreateSynthetic + ')');
  Logger.log('wouldPlaceholder:     ' + totalStats.wouldPlaceholder);
  Logger.log('duplicates:           ' + totalStats.duplicates);
  Logger.log('wouldCreateNewClient: ' + totalStats.wouldCreateNewClient);
  Logger.log('wouldSkipNoAbsence:   ' + totalStats.wouldSkipNoAbsence);
  if (unparsedSamples.length) {
    Logger.log('TOP unparsed патерни:');
    unparsedSamples.forEach(function(k){ Logger.log('  "' + k + '" \xd7 ' + unparsedCounts[k]); });
  }
  Logger.log('=== dryRunImportAbsences END ===');

  return {ok:true, stats:totalStats, byLocation:byLocation, unparsedSamples:unparsedSamples};
}

function _runDryRunImport()        { dryRunImportAbsences(''); }
function _runDryRunImportOsokory() { dryRunImportAbsences('Осокорки'); }

// ── РЕАЛЬНИЙ ІМПОРТ ВІДПУСТОК З PAYMENT → CRM ────────────────────────────────
// Лінива ініціалізація: картки створюються тільки якщо є записи у BL-BO.
// Один виклик saveClient() на дитину: картка + відпустки разом.
// Тезки: перший рядок CRM виграє, WARN у лозі.
function importAbsencesFromPayment(locFilter) {
  var SCHOOL_LOCS_SKIP = ['Школа Осокорки', 'Школа 228', 'Онлайн школа'];
  var refYear = new Date().getFullYear();
  var norm    = function(s){ return String(s||'').trim().toLowerCase().replace(/\s+/g,' '); };
  var nowISO  = new Date().toISOString();
  var todayUA = Utilities.formatDate(new Date(), 'Europe/Kiev', 'dd.MM.yyyy');

  // Завантажуємо CRM-клієнтів один раз
  var crmMap = _loadCRMClientsMap(norm);
  Logger.log('CRM: завантажено ' + Object.keys(crmMap).length + ' клієнтів');

  var configSS    = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  var configSheet = configSS.getSheets()[0];
  var configData  = configSheet.getDataRange().getValues();

  var stats = {
    locationsProcessed:0,
    newClientsCreated:0, existingClientsUpdated:0,
    absencesAdded:0, absencesPlaceholder:0, absencesDuplicates:0,
    errors:[]
  };

  Logger.log('=== importAbsencesFromPayment START' +
    (locFilter ? ' loc=' + locFilter : ' (всі локації)') + ' ===');

  for (var r = 1; r < configData.length; r++) {
    var loc       = trim(configData[r][2]);
    var sheetId   = trim(configData[r][3]);
    var sheetName = trim(configData[r][4]) || 'Payment';
    if (!loc || !sheetId) continue;
    if (SCHOOL_LOCS_SKIP.indexOf(loc) >= 0) continue;
    if (locFilter && loc !== locFilter) continue;

    stats.locationsProcessed++;

    try {
      var ss           = SpreadsheetApp.openById(sheetId);
      var paymentSheet = ss.getSheetByName(sheetName) || ss.getSheets()[0];
      var data         = paymentSheet.getDataRange().getValues();

      var absCols = detectAbsenceCols(data);
      if (absCols[0] === null) {
        Logger.log(loc + ': WARN — колонки BL/BM/BN/BO не знайдено, пропускаємо');
        continue;
      }

      // Контекст поточної групи — відстежуємо як у parsePaymentSheet
      var curGroup = '(без групи)'; var curGroupType = ''; var curTeacher = '';

      var DATA_START = 3;
      for (var row = DATA_START; row < data.length; row++) {
        var nameCell = trim(String(data[row][0] || ''));
        if (!nameCell) continue;

        // Рядок-заголовок групи → оновлюємо контекст
        if (isGroupHeaderRow(data[row], 1)) {
          var firstSpace = nameCell.search(/\s/);
          curTeacher    = firstSpace > 0 ? nameCell.slice(firstSpace).trim() : '';
          curGroupType  = normalizeGroupName(nameCell);
          curGroup      = curGroupType + (curTeacher ? ' ' + curTeacher : '');
          continue;
        }

        // Є хоча б один слот?
        var hasAnySlot = false;
        for (var si = 0; si < absCols.length; si++) {
          if (absCols[si] !== null && trim(String(data[row][absCols[si]] || ''))) { hasAnySlot = true; break; }
        }
        if (!hasAnySlot) continue;  // дитина без відпусток — не чіпаємо

        try {
          var crmKey  = norm(nameCell) + '|' + norm(loc);
          var isNew   = !crmMap.hasOwnProperty(crmKey);
          var existingAbsences = isNew ? [] : crmMap[crmKey].absences.slice();

          // Set from+to для dedupe
          var existingPairs = {};
          existingAbsences.forEach(function(a){ if(a.from&&a.to) existingPairs[a.from+'|'+a.to]=true; });

          // Збираємо нові відпустки з 4 слотів
          var newAbsences = [];
          for (var si2 = 0; si2 < absCols.length; si2++) {
            if (absCols[si2] === null) continue;
            var slot = trim(String(data[row][absCols[si2]] || ''));
            if (!slot) continue;

            var parsed = parseAbsencePeriod(slot, refYear);
            if (parsed) {
              var pairKey = parsed.from + '|' + parsed.to;
              if (existingPairs[pairKey]) {
                stats.absencesDuplicates++;
              } else {
                var absObj = _makeImportAbsence(parsed, slot);
                newAbsences.push(absObj);
                existingPairs[pairKey] = true;
                Logger.log('  ADDED absence: ' + nameCell + ' [' + parsed.from + ' → ' + parsed.to + ']');
              }
            } else {
              var absPlaceholder = _makeImportAbsence(null, slot);
              newAbsences.push(absPlaceholder);
              stats.absencesPlaceholder++;
              Logger.log('  PLACEHOLDER absence: ' + nameCell + ' ["' + slot + '"]');
            }
          }

          if (newAbsences.length === 0) continue;  // нічого нового — не чіпаємо

          // Будуємо повний об'єкт клієнта і зберігаємо одним викликом
          var allAbsences = existingAbsences.concat(newAbsences);
          var clientData;
          if (isNew) {
            clientData = {
              id:               'c_' + nameCell.trim().slice(0,20) + '_' + curGroupType.slice(0,8) + '_' + loc.slice(0,8),
              name:             nameCell,
              loc:              loc,
              group:            curGroup,
              teacher:          curTeacher,
              bday: '', momName: '', momPhone: '', dadName: '', dadPhone: '',
              contractDate: '', contractType: 'standard',
              monthlyFee: 0, entryFee: 0,
              status:           'active',
              notes:            'Створено автоматично при імпорті відпусток з Payment ' + todayUA,
              absences:         allAbsences,
              entryFeeSchedule: [],
              feeHistory:       []
            };
            Logger.log('CREATED client: ' + nameCell + ' / ' + loc + ' / ' + curGroup);
            var saveResult = saveClient(clientData);
          } else {
            var existing = crmMap[crmKey];
            var saveResult = patchClientAbsences(existing.id, allAbsences);
          }

          if (!saveResult.ok) {
            stats.errors.push({loc:loc, child:nameCell, error: saveResult.error || 'saveClient failed'});
            continue;
          }

          // Оновлюємо crmMap щоб наступна ітерація бачила оновлені дані
          if (isNew) {
            crmMap[crmKey] = {
              id: clientData.id, name: clientData.name, loc: clientData.loc,
              group: clientData.group, teacher: clientData.teacher,
              contractDate: clientData.contractDate, contractType: clientData.contractType,
              monthlyFee: clientData.monthlyFee, notes: clientData.notes,
              absences: allAbsences
            };
          } else {
            crmMap[crmKey].absences = allAbsences;
          }

          if (isNew) { stats.newClientsCreated++; }
          else       { stats.existingClientsUpdated++; }
          stats.absencesAdded += newAbsences.filter(function(a){ return a.from; }).length;

        } catch(childErr) {
          stats.errors.push({loc:loc, child:nameCell, error: childErr.message || String(childErr)});
          Logger.log('  ERROR child ' + nameCell + ': ' + (childErr.message || String(childErr)));
        }
      }

      Logger.log(loc + ': DONE');

    } catch(locErr) {
      Logger.log(loc + ': ERROR — ' + (locErr.message || String(locErr)));
      stats.errors.push({loc:loc, child:'', error: locErr.message || String(locErr)});
    }
  }

  Logger.log('');
  Logger.log('=== ПІДСУМОК ІМПОРТУ ===');
  Logger.log('Локацій оброблено:       ' + stats.locationsProcessed);
  Logger.log('Нових карток створено:   ' + stats.newClientsCreated);
  Logger.log('Існуючих оновлено:       ' + stats.existingClientsUpdated);
  Logger.log('Відпусток додано:        ' + stats.absencesAdded);
  Logger.log('Placeholder-відпусток:   ' + stats.absencesPlaceholder);
  Logger.log('Дублікатів пропущено:    ' + stats.absencesDuplicates);
  Logger.log('Помилок:                 ' + stats.errors.length);
  if (stats.errors.length) {
    stats.errors.slice(0,10).forEach(function(e){
      Logger.log('  ERR [' + e.loc + '] ' + e.child + ': ' + e.error);
    });
  }
  Logger.log('=== importAbsencesFromPayment END ===');

  return {ok:true, stats:stats};
}

function _runImportOsokory() {
  var r = importAbsencesFromPayment('Осокорки');
  Logger.log(JSON.stringify(r.stats, null, 2));
}

function _runImportAll() {
  Logger.log('=== _runImportAll START ===');
  var r = importAbsencesFromPayment('');
  Logger.log('');
  Logger.log('=== ПІДСУМОК _runImportAll ===');
  Logger.log('Локацій оброблено:       ' + r.stats.locationsProcessed);
  Logger.log('Нових карток створено:   ' + r.stats.newClientsCreated);
  Logger.log('Існуючих оновлено:       ' + r.stats.existingClientsUpdated);
  Logger.log('Відпусток додано:        ' + r.stats.absencesAdded);
  Logger.log('Placeholder-відпусток:   ' + r.stats.absencesPlaceholder);
  Logger.log('Дублікатів пропущено:    ' + r.stats.absencesDuplicates);
  Logger.log('Помилок:                 ' + r.stats.errors.length);
  if (r.stats.errors.length) {
    r.stats.errors.forEach(function(e){
      Logger.log('  ERR [' + e.loc + '] ' + e.child + ': ' + e.error);
    });
  }
  Logger.log('=== _runImportAll END ===');
}



// ── ДІАГНОСТИКА: перевіряє що саме знаходить detectAbsenceCols ───────────────
function diagnoseAbsenceCols(loc) {
  // Знаходимо sheetId локації
  var configSS    = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  var configSheet = configSS.getSheets()[0];
  var configData  = configSheet.getDataRange().getValues();
  var sheetId = null; var sheetName = 'Payment';
  for (var r = 1; r < configData.length; r++) {
    if (trim(String(configData[r][2] || '')) === loc) {
      sheetId   = trim(String(configData[r][3] || ''));
      sheetName = trim(String(configData[r][4] || '')) || 'Payment';
      break;
    }
  }
  if (!sheetId) { Logger.log('Location not found: ' + loc); return; }

  var ss    = SpreadsheetApp.openById(sheetId);
  var sheet = ss.getSheetByName(sheetName) || ss.getSheets()[0];
  var lastCol = sheet.getLastColumn();
  var lastRow = sheet.getLastRow();

  Logger.log('=== diagnoseAbsenceCols: ' + loc + ' ===');
  Logger.log('Sheet: ' + sheet.getName() + ', lastRow=' + lastRow + ', lastCol=' + lastCol);

  // Читаємо перші 5 рядків повністю (до col 90 або lastCol)
  var scanCols = Math.min(lastCol, 90);
  var scanRows = Math.min(5, lastRow);
  var headerRange = sheet.getRange(1, 1, scanRows, scanCols);
  var headerVals  = headerRange.getValues();

  // ── 1. Заголовки перших 80 колонок (рядки 1-3) ──────────────────────────────
  Logger.log('--- Заголовки рядків 1-3 (перші 80 колонок) ---');
  for (var rr = 0; rr < Math.min(3, scanRows); rr++) {
    var parts = [];
    for (var cc = 0; cc < Math.min(80, scanCols); cc++) {
      var v = headerVals[rr][cc];
      if (v !== '' && v !== null && v !== undefined) {
        // Літерне позначення колонки (A=0, B=1, ... Z=25, AA=26 ...)
        var colLetter = '';
        var n = cc;
        do { colLetter = String.fromCharCode(65 + (n % 26)) + colLetter; n = Math.floor(n / 26) - 1; } while (n >= 0);
        parts.push(colLetter + '(' + (cc+1) + ')="' + String(v).slice(0,30) + '"');
      }
    }
    Logger.log('Row ' + (rr+1) + ': ' + (parts.length ? parts.join(' | ') : '(порожній)'));
  }

  // ── 2. Що знайшов detectAbsenceCols ─────────────────────────────────────────
  var allData = sheet.getRange(1, 1, Math.min(lastRow, 10), scanCols).getValues();
  var absCols = detectAbsenceCols(allData);
  Logger.log('--- detectAbsenceCols результат ---');
  var labels = ['1 тиждень', '2 тиждень', '3 тиждень', '4 тиждень'];
  for (var li = 0; li < absCols.length; li++) {
    var ci = absCols[li];
    if (ci === null) {
      Logger.log('  ' + labels[li] + ' → NOT FOUND');
    } else {
      var colLetter2 = '';
      var n2 = ci;
      do { colLetter2 = String.fromCharCode(65 + (n2 % 26)) + colLetter2; n2 = Math.floor(n2 / 26) - 1; } while (n2 >= 0);
      Logger.log('  ' + labels[li] + ' → col index=' + ci + ' (' + colLetter2 + ')');
    }
  }

  // ── 3. Вміст знайдених колонок у рядках 4-6 (перші дані після заголовків) ───
  Logger.log('--- Вміст знайдених колонок у рядках 4-6 ---');
  if (lastRow >= 4) {
    var dataRange = sheet.getRange(4, 1, Math.min(3, lastRow - 3), scanCols);
    var dataVals  = dataRange.getValues();
    for (var dr = 0; dr < dataVals.length; dr++) {
      var rowNum = dr + 4;
      var name   = trim(String(dataVals[dr][0] || '(порожньо)'));
      var slots  = absCols.map(function(ci2, idx) {
        if (ci2 === null) return labels[idx] + '=null';
        var v2 = dataVals[dr][ci2];
        return labels[idx] + '=[' + String(v2) + ']';
      });
      Logger.log('  Row ' + rowNum + ' name="' + name + '" | ' + slots.join(' | '));
    }
  }

  Logger.log('=== END diagnoseAbsenceCols ===');
}

function _runDiagColsOsokory() { diagnoseAbsenceCols('Осокорки'); }

// ── ДІАГНОСТИКА: читає лист "Клієнти" і показує що в ньому ──────────────────
function diagnoseCRMClients() {
  var ss = getCRMSpreadsheet();
  Logger.log('=== diagnoseCRMClients ===');
  Logger.log('SHEET_CLIENTS constant: "' + SHEET_CLIENTS + '"');

  // Усі листи в таблиці
  Logger.log('Усі листи в SS: ' + ss.getSheets().map(function(s){ return '"' + s.getName() + '"'; }).join(', '));

  var sheet = ss.getSheetByName(SHEET_CLIENTS);
  Logger.log('sheet found: ' + (sheet ? 'YES' : 'NO'));

  if (sheet) {
    Logger.log('sheet name: "' + sheet.getName() + '"');
    Logger.log('lastRow: ' + sheet.getLastRow());
    Logger.log('lastCol: ' + sheet.getLastColumn());

    var vals = sheet.getDataRange().getValues();
    Logger.log('vals.length (getDataRange): ' + vals.length);

    for (var r = 0; r < Math.min(10, vals.length); r++) {
      Logger.log('row ' + r + ': ' + JSON.stringify(vals[r].slice(0, 5)));
    }
  }

  // Додатково: спробуємо getSheets()[0] і getSheets()[1] — раптом лист не той
  var sheets = ss.getSheets();
  for (var i = 0; i < Math.min(5, sheets.length); i++) {
    var sh = sheets[i];
    Logger.log('Sheet[' + i + '] "' + sh.getName() + '": lastRow=' + sh.getLastRow() + ', lastCol=' + sh.getLastColumn());
  }

  Logger.log('=== END diagnoseCRMClients ===');
}

function _runDiagCRM() { diagnoseCRMClients(); }

// ═══════════════════════════════════════════════════════════════════════════
// MIRROR: CRM Табель → старі таблиці медсестер (17 локацій)
// ═══════════════════════════════════════════════════════════════════════════

var NURSE_SHEET_TAB = 'табель медсестри';

var UA_MONTH_NAMES = [
  'січень','лютий','березень','квітень','травень','червень',
  'липень','серпень','вересень','жовтень','листопад','грудень'
];

var _nurseCache = null;

function norm(s) {
  return String(s || '').trim().toLowerCase()
    .replace(/[’ʼ′`]/g, "'");
}

function loadNurseSheetMap() {
  if (_nurseCache && _nurseCache.map) return _nurseCache.map;
  var map = {};
  try {
    var configSS = SpreadsheetApp.openById(CONFIG_SHEET_ID);
    var nurseSheet = null;
    var sheets = configSS.getSheets();
    for (var i = 0; i < sheets.length; i++) {
      if (norm(sheets[i].getName()) === norm(NURSE_SHEET_TAB)) {
        nurseSheet = sheets[i];
        break;
      }
    }
    if (!nurseSheet) {
      Logger.log('mirror: WARN — лист "' + NURSE_SHEET_TAB + '" не знайдено у Config');
    } else {
      var data = nurseSheet.getDataRange().getValues();
      for (var r = 1; r < data.length; r++) {
        var loc = String(data[r][2] || '').trim();
        var sid = String(data[r][3] || '').trim();
        if (loc && sid) map[norm(loc)] = sid;
      }
    }
  } catch (e) {
    Logger.log('mirror: ERROR loadNurseSheetMap — ' + (e.message || e));
  }
  _nurseCache = { map: map, ss: {}, sheet: {} };
  return map;
}

function findMonthTab(spreadsheet, isoDate) {
  var m = String(isoDate).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return null;
  var monName = UA_MONTH_NAMES[+m[2] - 1];
  var yFull = m[1];
  var yy    = yFull.slice(-2);

  var sheets = spreadsheet.getSheets();

  for (var i = 0; i < sheets.length; i++) {
    var n1 = sheets[i].getName().toLowerCase();
    if (n1.indexOf(monName) >= 0 && n1.indexOf(yFull) >= 0) return sheets[i];
  }
  for (var j = 0; j < sheets.length; j++) {
    var n2 = sheets[j].getName().toLowerCase();
    if (n2.indexOf(monName) < 0) continue;
    var nums = n2.match(/\d+/g) || [];
    for (var k = 0; k < nums.length; k++) {
      if (nums[k] === yy || nums[k] === yFull) return sheets[j];
    }
  }
  return null;
}

function findChildRow(sheet, childName) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return -1;

  var values = sheet.getRange(3, 1, lastRow - 2, 1).getValues();
  var target = norm(childName);
  if (!target) return -1;

  for (var i = 0; i < values.length; i++) {
    if (norm(values[i][0]) === target) return i + 3;
  }
  return -1;
}

function findDateColumn(sheet, isoDate) {
  var m = String(isoDate).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return -1;
  var target = m[3] + '/' + m[2];
  var lastCol = sheet.getLastColumn();
  if (lastCol < 1) return -1;
  var row2 = sheet.getRange(2, 1, 1, lastCol).getValues()[0];
  for (var c = 0; c < lastCol; c++) {
    var raw = row2[c];
    var s;
    if (raw instanceof Date) {
      s = pad2(raw.getDate()) + '/' + pad2(raw.getMonth() + 1);
    } else {
      s = String(raw || '').trim();
      if (!s) continue;
      var dm = s.replace(/[.\-]/g, '/').match(/^(\d{1,2})\/(\d{1,2})(?:\/\d{2,4})?$/);
      if (dm) s = pad2(+dm[1]) + '/' + pad2(+dm[2]);
    }
    if (s === target) return c + 1;
  }
  return -1;
}

function mapStatusToNurseCell(status) {
  if (status === 'present') return '1';
  if (status === 'sick' || status === 'vacation' || status === 'absent' || status === '') return '';
  return null;
}

function mirrorAttendanceToNurseSheet(loc, childName, isoDate, status) {
  try {
    var newVal = mapStatusToNurseCell(status);
    if (newVal === null) {
      Logger.log('mirror: skip unknown status "' + status + '" for ' + childName + ' / ' + isoDate);
      return;
    }

    var map = loadNurseSheetMap();
    var sid = map[norm(loc)];
    if (!sid) {
      Logger.log('mirror: WARN — нема spreadsheetId для локації "' + loc + '" (norm="' + norm(loc) + '")');
      return;
    }

    var ss = _nurseCache.ss[sid];
    if (!ss) {
      ss = SpreadsheetApp.openById(sid);
      _nurseCache.ss[sid] = ss;
    }

    var monthKey = sid + '|' + isoDate.slice(0, 7);
    var sheet = _nurseCache.sheet[monthKey];
    if (sheet === undefined) {
      sheet = findMonthTab(ss, isoDate) || false;
      _nurseCache.sheet[monthKey] = sheet;
    }
    if (!sheet) {
      Logger.log('mirror: WARN — таб місяця не знайдено у [' + loc + '] для дати ' + isoDate);
      return;
    }

    var rowNum = findChildRow(sheet, childName);
    if (rowNum < 1) {
      Logger.log('mirror: WARN — дитину "' + childName + '" не знайдено у [' + loc + '] / таб "' + sheet.getName() + '"');
      return;
    }
    var colNum = findDateColumn(sheet, isoDate);
    if (colNum < 1) {
      Logger.log('mirror: WARN — колонку дати ' + isoDate + ' не знайдено у [' + loc + '] / таб "' + sheet.getName() + '"');
      return;
    }

    var cell = sheet.getRange(rowNum, colNum);
    var existing = String(cell.getValue() || '').trim();
    if (existing.toUpperCase() === 'А' || existing.toUpperCase() === 'A') {
      Logger.log('mirror: skip "А" (адаптація) для ' + childName + ' / ' + isoDate);
      return;
    }
    if (existing === newVal) return;
    cell.setValue(newVal);
  } catch (e) {
    Logger.log('mirror: ERROR ' + (e.message || e) + ' для ' + childName + ' / ' + loc + ' / ' + isoDate);
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// MIGRATE: дати народження дітей з реєстрів договорів → CRM Клієнти
// ═══════════════════════════════════════════════════════════════════════════

var REGISTRY_TAB_NAME = 'реєстр';

// Читає лист "реєстр" у CONFIG_SHEET_ID. Повертає [{direction, type, location, sheetId, listName}].
// Пропускає header (рядок 1) і рядки де колонка D (sheetId) порожня.
function getRegistries() {
  var configSS = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  var sheet = configSS.getSheetByName(REGISTRY_TAB_NAME);
  if (!sheet) {
    Logger.log('getRegistries: WARN — лист "' + REGISTRY_TAB_NAME + '" не знайдено у CONFIG');
    return [];
  }
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var data = sheet.getRange(1, 1, lastRow, 5).getValues();
  var out = [];
  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    var sheetId = String(row[3] || '').trim();
    if (!sheetId) continue;
    out.push({
      direction: String(row[0] || '').trim(),
      type:      String(row[1] || '').trim(),
      location:  String(row[2] || '').trim(),
      sheetId:   sheetId,
      listName:  String(row[4] || '').trim() || '2025'
    });
  }
  return out;
}

// Повертає map {location: spreadsheetUrl} для клікабельних бейджів реєстру у CRM.
// Якщо одна локація має кілька записів — перший по порядку перемагає.
function getRegistryUrls() {
  var regs = getRegistries();
  var map = {};
  for (var i = 0; i < regs.length; i++) {
    var r = regs[i];
    if (!r.location || !r.sheetId) continue;
    if (map[r.location]) continue;
    map[r.location] = 'https://docs.google.com/spreadsheets/d/' + r.sheetId + '/edit';
  }
  return {ok:true, data:map};
}

// Парсить дату з реєстру в ISO YYYY-MM-DD. Підтримує:
// - Date object (sheets API повертає коли клітинка тип Date)
// - "DD.MM.YYYY" або "DD/MM/YYYY"
// - "YYYY-MM-DD"
// Повертає null для непарсабельних значень.
function parseRegistryBday(val) {
  if (val === null || val === undefined || val === '') return null;
  if (val instanceof Date) {
    if (isNaN(val.getTime())) return null;
    var y = val.getFullYear(), m = val.getMonth() + 1, d = val.getDate();
    if (y < 1900 || y > 2030) return null;
    return y + '-' + (m < 10 ? '0' + m : m) + '-' + (d < 10 ? '0' + d : d);
  }
  // Excel serial number (число > 10000)
  if (typeof val === 'number' && val > 10000) {
    var excelEpoch = new Date(Date.UTC(1899, 11, 30));
    var dt = new Date(excelEpoch.getTime() + val * 86400000);
    var ye = dt.getUTCFullYear(), me = dt.getUTCMonth() + 1, de = dt.getUTCDate();
    if (ye < 1900 || ye > 2030) return null;
    return ye + '-' + (me < 10 ? '0' + me : me) + '-' + (de < 10 ? '0' + de : de);
  }
  if (typeof val === 'string') {
    // Нормалізуємо роздільники: пробіл / дефіс / слеш → крапка
    var s = val.trim().replace(/[\s\-\/]+/g, '.');
    // DD.MM.YYYY
    var m1 = s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);
    if (m1) {
      var dd = +m1[1], mm = +m1[2], yyyy = +m1[3];
      if (yyyy < 1900 || yyyy > 2030 || mm < 1 || mm > 12 || dd < 1 || dd > 31) return null;
      return yyyy + '-' + (mm < 10 ? '0' + mm : mm) + '-' + (dd < 10 ? '0' + dd : dd);
    }
    // YYYY.MM.DD (після нормалізації роздільників — раніше було YYYY-MM-DD)
    var m2 = s.match(/^(\d{4})\.(\d{2})\.(\d{2})$/);
    if (m2) {
      var yyyy2 = +m2[1], mm2 = +m2[2], dd2 = +m2[3];
      if (yyyy2 < 1900 || yyyy2 > 2030 || mm2 < 1 || mm2 > 12 || dd2 < 1 || dd2 > 31) return null;
      return yyyy2 + '-' + (mm2 < 10 ? '0' + mm2 : mm2) + '-' + (dd2 < 10 ? '0' + dd2 : dd2);
    }
  }
  return null;
}

function _normChildName(s) {
  return String(s || '')
    .replace(/\([^)]*\)/g, '')      // прибрати (...)
    .replace(/\[[^\]]*\]/g, '')     // прибрати [...]
    .trim()
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .replace(/[’ʼ′`'']/g, "'")  // нормалізувати різні апострофи
    .replace(/і/g, 'и')
    .replace(/ї/g, 'и')
    .replace(/є/g, 'е');
}

// Точна копія childId з clients.html:847 — для синхронізації lookup-ключів
function _childId(name, group, loc) {
  return 'c_' + String(name||'').trim().slice(0,20) +
         '_' + String(group||'').slice(0,8) +
         '_' + String(loc||'').slice(0,8);
}

function _commonPrefixLen(a, b) {
  var n = Math.min(a.length, b.length), i = 0;
  while (i < n && a.charCodeAt(i) === b.charCodeAt(i)) i++;
  return i;
}

// Обходить усі реєстри, парсить ПІБ + дату нар., оновлює пусті bday у CRM Клієнти.
// Повертає об'єкт-звіт.
function migrateChildrenBdays() {
  var regs = getRegistries();
  Logger.log('Реєстрів: ' + regs.length);

  var crmSS = getCRMSpreadsheet();
  var crmSheet = crmSS.getSheetByName(SHEET_CLIENTS);
  if (!crmSheet) {
    Logger.log('CRM Клієнти sheet not found');
    return { regsScanned: 0, regsErrored: 0, errors: ['CRM Клієнти sheet not found'],
             totalRowsInRegistries: 0, matched: 0, updated: 0,
             alreadyHadBday: 0, unparseable: 0,
             notFoundInCRMCount: 0, notFoundSamples: [] };
  }

  ensureClientsHeader(crmSheet);
  var crmData = crmSheet.getDataRange().getValues();
  if (crmData.length < 2) {
    Logger.log('CRM Клієнти is empty');
    return { regsScanned: regs.length, regsErrored: 0, errors: [],
             totalRowsInRegistries: 0, matched: 0, updated: 0,
             alreadyHadBday: 0, unparseable: 0,
             notFoundInCRMCount: 0, notFoundSamples: [] };
  }
  var crmHdr = crmData[0].map(function(h){ return String(h || ''); });
  var crmNameI = crmHdr.indexOf('ПІБ дитини');
  var crmBdayI = crmHdr.indexOf('Дата народження');
  if (crmNameI < 0 || crmBdayI < 0) {
    var msg = 'CRM headers missing: ПІБ дитини=' + crmNameI + ', Дата народження=' + crmBdayI;
    Logger.log(msg);
    return { regsScanned: regs.length, regsErrored: 0, errors: [msg],
             totalRowsInRegistries: 0, matched: 0, updated: 0,
             alreadyHadBday: 0, unparseable: 0,
             notFoundInCRMCount: 0, notFoundSamples: [] };
  }

  Logger.log('CRM diag: lastRow=' + crmSheet.getLastRow() + ', getDataRange rows=' + crmData.length);
  var nonEmptyChildRows = 0;
  for (var ci = 1; ci < crmData.length; ci++) {
    var rawN = String(crmData[ci][crmNameI] || '').trim();
    if (rawN) nonEmptyChildRows++;
  }
  Logger.log('CRM diag: rows with non-empty ПІБ дитини = ' + nonEmptyChildRows);

  // Будуємо мапу нормалізованого ПІБ → {row (1-based), currentBday}
  var crmMap = {};
  var dupeCount = 0;
  for (var i = 1; i < crmData.length; i++) {
    var nm = _normChildName(crmData[i][crmNameI]);
    if (!nm) continue;
    if (!crmMap[nm]) {
      crmMap[nm] = { row: i + 1, currentBday: String(crmData[i][crmBdayI] || '').trim() };
    } else {
      dupeCount++;
    }
  }
  Logger.log('CRM map size: ' + Object.keys(crmMap).length + ' (dupes after norm: ' + dupeCount + ')');

  var updates = [];
  var alreadyHadBday = 0;
  var unparseable = 0;
  var notFoundInCRMCount = 0;
  var notFoundSamples = [];
  var totalRowsInRegistries = 0;
  var errors = [];

  for (var ri = 0; ri < regs.length; ri++) {
    var reg = regs[ri];
    Logger.log('--- [' + (ri+1) + '/' + regs.length + '] ' + reg.location + ' (sheetId=' + reg.sheetId + ', list=' + reg.listName + ')');
    try {
      var ss = SpreadsheetApp.openById(reg.sheetId);
      var sh = ss.getSheetByName(reg.listName);
      if (!sh) sh = ss.getSheets()[0];
      if (!sh) { errors.push(reg.location + ': лист не знайдено'); continue; }

      var data = sh.getDataRange().getValues();
      if (data.length < 2) { Logger.log('  empty'); continue; }
      var hdr = data[0];

      var childCol = -1, bdayCol = -1;
      for (var hi = 0; hi < hdr.length; hi++) {
        var hLower = String(hdr[hi] || '').toLowerCase();
        if (childCol < 0 && hLower.indexOf('піб дитини') >= 0) childCol = hi;
        if (bdayCol < 0 && hLower.indexOf('дата народження') >= 0) bdayCol = hi;
      }
      if (childCol < 0 || bdayCol < 0) {
        errors.push(reg.location + ': не знайдено колонок ПІБ дитини=' + childCol + ' / Дата народження=' + bdayCol);
        Logger.log('  headers: ' + JSON.stringify(hdr));
        continue;
      }
      Logger.log('  childCol=' + childCol + ', bdayCol=' + bdayCol + ', rows=' + (data.length - 1));

      for (var dr = 1; dr < data.length; dr++) {
        var name = String(data[dr][childCol] || '').trim();
        var bdayRaw = data[dr][bdayCol];
        if (!name || (bdayRaw === '' || bdayRaw === null || bdayRaw === undefined)) continue;
        totalRowsInRegistries++;

        var bdayISO = parseRegistryBday(bdayRaw);
        if (!bdayISO) { unparseable++; continue; }

        var lookup = crmMap[_normChildName(name)];
        if (!lookup) {
          notFoundInCRMCount++;
          if (notFoundSamples.length < 20) notFoundSamples.push(name + ' [' + reg.location + ']');
          continue;
        }
        if (lookup.currentBday) { alreadyHadBday++; continue; }
        updates.push({ row: lookup.row, bday: bdayISO });
        lookup.currentBday = bdayISO; // щоб не оновити двічі
      }
    } catch (e) {
      var em = reg.location + ': ' + (e.message || e);
      errors.push(em);
      Logger.log('  ERROR ' + em);
    }
  }

  // Виконуємо batch updates
  for (var ui = 0; ui < updates.length; ui++) {
    var u = updates[ui];
    crmSheet.getRange(u.row, crmBdayI + 1).setValue(u.bday);
  }

  var report = {
    regsScanned: regs.length,
    regsErrored: errors.length,
    errors: errors,
    totalRowsInRegistries: totalRowsInRegistries,
    matched: updates.length + alreadyHadBday,
    updated: updates.length,
    alreadyHadBday: alreadyHadBday,
    unparseable: unparseable,
    notFoundInCRMCount: notFoundInCRMCount,
    notFoundSamples: notFoundSamples
  };
  Logger.log('=== REPORT ===');
  Logger.log(JSON.stringify(report, null, 2));
  return report;
}

function runBdayMigration() {
  var r = migrateChildrenBdays();
  var summary =
    'Реєстрів: ' + r.regsScanned + '\n' +
    'Помилок: ' + r.regsErrored + '\n' +
    'Рядків переглянуто: ' + r.totalRowsInRegistries + '\n' +
    'Знайдено в CRM: ' + r.matched + '\n' +
    'ЗАПОВНЕНО bday: ' + r.updated + '\n' +
    'Вже мали bday: ' + r.alreadyHadBday + '\n' +
    'Непарсабельні: ' + r.unparseable + '\n' +
    'Не знайдено в CRM: ' + r.notFoundInCRMCount + '\n\n' +
    (r.errors.length ? 'Помилки:\n' + r.errors.join('\n') + '\n\n' : '') +
    (r.notFoundSamples.length ? 'Перші 20 не знайдених:\n' + r.notFoundSamples.join('\n') : '');
  try {
    SpreadsheetApp.getUi().alert('Міграція bday', summary, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    Logger.log('UI alert недоступний (запущено не з sheet context). Звіт у логах.');
    Logger.log(summary);
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// SYNC BDAYS: реєстри договорів → колонка "Дата народження" у листі "Оплати"
// ═══════════════════════════════════════════════════════════════════════════
//
// Не чіпає aggregatePayments. Викликати ОКРЕМО після кожного перерахунку
// (бо aggregatePayments робить clearContents() і затирає колонку bday).
// Або використовувати runFullAggregateAndSync() — комбінований wrapper.

// @deprecated — використовуйте syncBdayStatusSheet().
// Зберігається тимчасово; колонка "Дата народження" в Оплати стирається при aggregatePayments.
function syncBdaysToPayments() {
  var crmSS = getCRMSpreadsheet();
  var paySheet = crmSS.getSheetByName(SHEET_PAYMENTS);
  if (!paySheet) {
    Logger.log('syncBdaysToPayments: ERROR — лист "' + SHEET_PAYMENTS + '" не знайдено');
    return { ok: false, error: 'Sheet not found', errors: [] };
  }

  var data = paySheet.getDataRange().getValues();
  if (data.length < 2) {
    Logger.log('syncBdaysToPayments: лист "' + SHEET_PAYMENTS + '" порожній');
    return { ok: true, total: 0, withBday: 0, withoutBday: 0, errors: [] };
  }

  var hdr = data[0].map(function(h){ return String(h || ''); });
  var locI = hdr.indexOf('Локація');
  var nameI = hdr.indexOf("Ім'я дитини");
  if (locI < 0 || nameI < 0) {
    var msg = 'Не знайдено колонки: Локація=' + locI + ', Ім\'я дитини=' + nameI;
    Logger.log('syncBdaysToPayments: ERROR — ' + msg);
    return { ok: false, error: msg, errors: [] };
  }

  // Якщо колонки "Дата народження" немає — додаємо в кінець
  var bdayI = hdr.indexOf('Дата народження');
  if (bdayI < 0) {
    bdayI = hdr.length;
    paySheet.getRange(1, bdayI + 1).setValue('Дата народження');
    Logger.log('syncBdaysToPayments: додано колонку "Дата народження" на позиції ' + (bdayI + 1));
    // Перечитуємо щоб мати актуальний шар колонок
    data = paySheet.getDataRange().getValues();
  }

  // Будуємо мапу bday з усіх реєстрів
  var regs = getRegistries();
  Logger.log('syncBdaysToPayments: реєстрів — ' + regs.length);
  var bdayMap = {};
  var errors = [];
  var totalRegRows = 0;

  for (var ri = 0; ri < regs.length; ri++) {
    var reg = regs[ri];
    try {
      var ss = SpreadsheetApp.openById(reg.sheetId);
      var sh = ss.getSheetByName(reg.listName) || ss.getSheets()[0];
      if (!sh) { errors.push(reg.location + ': лист не знайдено'); continue; }
      var rData = sh.getDataRange().getValues();
      if (rData.length < 2) continue;
      var rHdr = rData[0];
      var childCol = -1, bdayColReg = -1;
      for (var hi = 0; hi < rHdr.length; hi++) {
        var hLower = String(rHdr[hi] || '').toLowerCase();
        if (childCol < 0 && hLower.indexOf('піб дитини') >= 0) childCol = hi;
        if (bdayColReg < 0 && hLower.indexOf('дата народження') >= 0) bdayColReg = hi;
      }
      if (childCol < 0 || bdayColReg < 0) {
        errors.push(reg.location + ': не знайдено колонок ПІБ дитини=' + childCol + ' / Дата народження=' + bdayColReg);
        continue;
      }
      for (var dr = 1; dr < rData.length; dr++) {
        var nm = String(rData[dr][childCol] || '').trim();
        var bdayRaw = rData[dr][bdayColReg];
        if (!nm || bdayRaw === '' || bdayRaw === null || bdayRaw === undefined) continue;
        totalRegRows++;
        var bdayISO = parseRegistryBday(bdayRaw);
        if (!bdayISO) continue;
        var key = _normChildName(nm) + '|' + reg.location;
        bdayMap[key] = bdayISO;
      }
    } catch (e) {
      errors.push(reg.location + ': ' + (e.message || e));
    }
  }
  Logger.log('syncBdaysToPayments: bdayMap size — ' + Object.keys(bdayMap).length + ' з ' + totalRegRows + ' рядків реєстрів');

  // [DEBUG] Reverse index: per-loc список нормалізованих імен з реєстрів
  var bdayKeysByLoc = {};
  Object.keys(bdayMap).forEach(function(k){
    var pipe = k.lastIndexOf('|');
    if (pipe < 0) return;
    var nm = k.slice(0, pipe);
    var lc = k.slice(pipe + 1);
    if (!bdayKeysByLoc[lc]) bdayKeysByLoc[lc] = [];
    bdayKeysByLoc[lc].push(nm);
  });

  // Будуємо нову колонку bday для Оплати: повний масив (длина = data.length - 1)
  var newCol = [];
  var withBday = 0;
  var withoutBday = 0;
  var exactMatched = 0;
  var fuzzyMatched = 0;
  var notMatchedSamples = []; // [DEBUG]
  for (var r = 1; r < data.length; r++) {
    var pName = String(data[r][nameI] || '').trim();
    var pLoc  = String(data[r][locI]  || '').trim();
    var existing = String(data[r][bdayI] || '').trim();
    if (existing) {
      newCol.push([existing]);
      withBday++;
      continue;
    }
    var norm = _normChildName(pName);
    var exactKey = norm + '|' + pLoc;
    var bday = bdayMap[exactKey] || '';
    if (bday) {
      exactMatched++;
    } else {
      // Fuzzy fallback: префікс прізвище + перші 3 літери імені
      var parts = norm.split(' ').filter(function(p){ return p; });
      if (parts.length >= 2) {
        var surname = parts[0];
        var firstStart = parts[1].substring(0, 3);
        var prefix = surname + ' ' + firstStart;
        var candidates = [];
        for (var k in bdayMap) {
          if (k.indexOf('|' + pLoc) === k.length - pLoc.length - 1 && k.indexOf(prefix) === 0) {
            candidates.push(k);
          }
        }
        if (candidates.length === 1) {
          bday = bdayMap[candidates[0]];
          fuzzyMatched++;
        }
      }
    }
    newCol.push([bday]);
    if (bday) {
      withBday++;
    } else {
      withoutBday++;
      if (notMatchedSamples.length < 50) {
        notMatchedSamples.push({name: pName, loc: pLoc, normKey: exactKey});
      }
    }
  }
  // [DEBUG] загальний список не-матчів
  Logger.log('[DEBUG] exactMatched=' + exactMatched + ' fuzzyMatched=' + fuzzyMatched + ' withoutBday=' + withoutBday);
  Logger.log('[DEBUG] notMatchedSamples (' + notMatchedSamples.length + '/50):');
  notMatchedSamples.forEach(function(s, i){
    Logger.log('  [' + (i+1) + '] "' + s.name + '" / "' + s.loc + '" → key="' + s.normKey + '"');
  });

  // Один batch-запис у всю колонку
  if (newCol.length > 0) {
    paySheet.getRange(2, bdayI + 1, newCol.length, 1).setValues(newCol);
  }

  var report = {
    ok: true,
    total: data.length - 1,
    withBday: withBday,
    withoutBday: withoutBday,
    exactMatched: exactMatched,
    fuzzyMatched: fuzzyMatched,
    bdayMapSize: Object.keys(bdayMap).length,
    regsScanned: regs.length,
    regsErrored: errors.length,
    errors: errors
  };
  Logger.log('=== syncBdaysToPayments REPORT ===');
  Logger.log(JSON.stringify(report, null, 2));
  return report;
}

// @deprecated — використовуйте runSyncBdayStatus().
function runSyncBdaysToPayments() {
  var r = syncBdaysToPayments();
  var summary = !r.ok
    ? 'Помилка: ' + (r.error || 'unknown')
    : 'Рядків Оплати: ' + r.total + '\n' +
      'Заповнено bday: ' + r.withBday + '\n' +
      'Без bday: ' + r.withoutBday + '\n' +
      'Реєстрів: ' + r.regsScanned + ' (помилок: ' + r.regsErrored + ')\n\n' +
      (r.errors && r.errors.length ? 'Помилки:\n' + r.errors.join('\n') : '');
  try {
    SpreadsheetApp.getUi().alert('Синхронізація bday', summary, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    Logger.log('UI alert недоступний. Звіт у логах.');
    Logger.log(summary);
  }
}

// Use this instead of direct aggregatePayments call.
// aggregatePayments оновлює лист "Оплати"; syncBdayStatusSheet після нього
// перебудовує bday_sync_status (нові діти попадають з ПІБ та локацією).
function runFullAggregateAndSync() {
  try {
    aggregatePayments();        // існуюча, не чіпаємо
    syncBdayStatusSheet();      // НОВА (замість deprecated syncBdaysToPayments)
    try {
      SpreadsheetApp.getUi().alert(
        '✅ Перерахунок виконано',
        'Оплати оновлені + bday_sync_status оновлено',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    } catch (e2) {
      Logger.log('runFullAggregateAndSync: успіх (UI alert недоступний)');
    }
  } catch (e) {
    Logger.log('runFullAggregateAndSync error: ' + e);
    try {
      SpreadsheetApp.getUi().alert('Помилка', String(e),
        SpreadsheetApp.getUi().ButtonSet.OK);
    } catch (e2) {}
  }
}

// ── DEBUG: повна категоризація всіх не-зіставлених дітей ─────────────────────
function _debugAllMissingBdays() {
  Logger.log('=== ПОВНА ДІАГНОСТИКА 737 НЕ ЗІСТАВЛЕНИХ ===');

  // 1. Будуємо bdayMap (як syncBdaysToPayments)
  // АЛЕ також зберігаємо ВСІ варіанти (для виявлення колізій)
  const regs = getRegistries();
  const bdayMap = {};                  // key → bday
  const bdayMapAllOccurrences = {};    // key → [{bday, regLocation, fullName}]
  const allRegistryNames = {};         // location → [name1, name2, ...]
  const allRegistryNormNames = {};     // normName → [{loc, fullName, bday}]

  for (const reg of regs) {
    allRegistryNames[reg.location] = [];
    try {
      const ss = SpreadsheetApp.openById(reg.sheetId);
      const sh = ss.getSheetByName(reg.listName) || ss.getSheets()[0];
      const data = sh.getDataRange().getValues();
      const hdr = data[0];
      const childCol = hdr.findIndex(h =>
        String(h).toLowerCase().includes('піб дитини'));
      const bdayColReg = hdr.findIndex(h =>
        String(h).toLowerCase().includes('дата народження'));
      if (childCol < 0 || bdayColReg < 0) continue;

      for (let i = 1; i < data.length; i++) {
        const name = String(data[i][childCol] || '').trim();
        const bdayRaw = data[i][bdayColReg];
        if (!name) continue;
        allRegistryNames[reg.location].push(name);
        if (!bdayRaw) continue;
        const bdayISO = parseRegistryBday(bdayRaw);
        if (!bdayISO) continue;

        const norm = _normChildName(name);
        const key = norm + '|' + reg.location;

        if (!bdayMapAllOccurrences[key]) {
          bdayMapAllOccurrences[key] = [];
        }
        bdayMapAllOccurrences[key].push({
          bday: bdayISO, fullName: name, regLocation: reg.location
        });

        if (!allRegistryNormNames[norm]) {
          allRegistryNormNames[norm] = [];
        }
        allRegistryNormNames[norm].push({
          loc: reg.location, fullName: name, bday: bdayISO
        });

        bdayMap[key] = bdayISO;
      }
    } catch (e) {
      Logger.log('ERROR ' + reg.location + ': ' + e);
    }
  }

  // 2. Обходимо Оплати
  const ss = getCRMSpreadsheet();
  const paySheet = ss.getSheetByName('Оплати');
  const data = paySheet.getDataRange().getValues();
  const header = data[0];
  const nameCol = header.findIndex(h =>
    String(h).toLowerCase().includes("ім'я дитини"));
  const locCol = header.findIndex(h =>
    String(h).toLowerCase().includes('локація'));

  // 3. Категоризація 737 не зіставлених
  const categoryA = [];  // в реєстрі є, в bdayMap є — БАГ скрипта
  const categoryB = [];  // в реєстрі НЕМАЄ
  const categoryC = [];  // однофамільці (колізія)
  const categoryD = [];  // є в реєстрі іншої локації
  const categoryE = [];  // дата в реєстрі є, але не парсається

  for (let i = 1; i < data.length; i++) {
    const name = String(data[i][nameCol] || '').trim();
    const loc = String(data[i][locCol] || '').trim();
    if (!name || !loc) continue;

    const norm = _normChildName(name);
    const lookupKey = norm + '|' + loc;

    if (bdayMap[lookupKey]) continue; // зіставлено

    // НЕ зіставлено — категоризуємо
    const occurrences = bdayMapAllOccurrences[lookupKey];
    const allNormMatches = allRegistryNormNames[norm] || [];
    const inSameLoc = allRegistryNames[loc] &&
      allRegistryNames[loc].some(n => _normChildName(n) === norm);

    if (occurrences && occurrences.length > 1) {
      categoryC.push({name, loc, count: occurrences.length});
    } else if (inSameLoc) {
      // в реєстрі є, в bdayMap немає — або бад дата
      categoryE.push({name, loc});
    } else if (allNormMatches.length > 0) {
      // в іншій локації є
      categoryD.push({
        name, loc,
        foundIn: allNormMatches.map(m => m.loc).join(',')
      });
    } else {
      categoryB.push({name, loc});
    }
  }

  // 4. Звіт
  Logger.log('--- ЗВІТ ---');
  Logger.log('bdayMap size: ' + Object.keys(bdayMap).length);
  Logger.log('');
  Logger.log('A. Колізії (в map переписано): ' + categoryC.length);
  Logger.log('B. Немає в жодному реєстрі: ' + categoryB.length);
  Logger.log('C. Є в реєстрі іншої локації: ' + categoryD.length);
  Logger.log('D. Є в реєстрі цієї локації, але дата не парсається: '
    + categoryE.length);

  Logger.log('');
  Logger.log('--- A. КОЛІЗІЇ (перші 10) ---');
  categoryC.slice(0, 10).forEach(x =>
    Logger.log('  ' + x.name + ' / ' + x.loc + ' (× ' + x.count + ')'));

  Logger.log('');
  Logger.log('--- B. НЕМАЄ В РЕЄСТРАХ (перші 20) ---');
  categoryB.slice(0, 20).forEach(x =>
    Logger.log('  ' + x.name + ' / ' + x.loc));

  Logger.log('');
  Logger.log('--- C. В ІНШИХ ЛОКАЦІЯХ (перші 20) ---');
  categoryD.slice(0, 20).forEach(x =>
    Logger.log('  ' + x.name + ' / ' + x.loc + ' → знайдено в: ' + x.foundIn));

  Logger.log('');
  Logger.log('--- D. ДАТА НЕ ПАРСАЄТЬСЯ (перші 10) ---');
  categoryE.slice(0, 10).forEach(x =>
    Logger.log('  ' + x.name + ' / ' + x.loc));

  Logger.log('');
  Logger.log('=== РЕКОМЕНДАЦІЇ ===');
  Logger.log('A. Колізії = виправити кодом (додати другий ключ - група або номер договору)');
  Logger.log('B. = окрема задача для HR (внести в реєстр)');
  Logger.log('C. = виправити кодом (lookup без локації або по locFromOplata→locInRegistry)');
  Logger.log('D. = виправити в реєстрі (дата написана дивно)');
}

// ═══════════════════════════════════════════════════════════════════════════
// SYNC bday_sync_status: окремий лист зі статусами match'у дітей з реєстрами
// ═══════════════════════════════════════════════════════════════════════════
//
// Не модифікує лист "Оплати". Створює/перезаписує окремий лист bday_sync_status.
// Frontend читає з нього bday + contractNumber + status badge.

var BDAY_STATUS_SHEET = 'bday_sync_status';

function syncBdayStatusSheet() {
  var crmSS = getCRMSpreadsheet();
  var statusSheet = crmSS.getSheetByName(BDAY_STATUS_SHEET);
  if (!statusSheet) {
    statusSheet = crmSS.insertSheet(BDAY_STATUS_SHEET);
    Logger.log('syncBdayStatusSheet: створено новий лист "' + BDAY_STATUS_SHEET + '"');
  }
  var HEADER = ['ChildID','Name','Loc','Bday','ContractNumber','Status','MatchedRegName','UpdatedAt','ConfirmedBy','ConfirmedAt'];

  // 0. Знімок confirmed-рядків — їх НЕ перезаписуємо
  var confirmedById = {};
  var existingLastRow = statusSheet.getLastRow();
  if (existingLastRow >= 2) {
    var existingHdr = statusSheet.getRange(1, 1, 1, statusSheet.getLastColumn()).getValues()[0];
    var idIdx        = existingHdr.indexOf('ChildID');
    var statusIdx    = existingHdr.indexOf('Status');
    if (idIdx >= 0 && statusIdx >= 0) {
      var existingData = statusSheet.getRange(2, 1, existingLastRow - 1, statusSheet.getLastColumn()).getValues();
      for (var ei = 0; ei < existingData.length; ei++) {
        var erow = existingData[ei];
        if (String(erow[statusIdx]||'').trim() === 'confirmed') {
          var eid = String(erow[idIdx]||'').trim();
          if (eid) confirmedById[eid] = erow;
        }
      }
    }
    Logger.log('syncBdayStatusSheet: збережено ' + Object.keys(confirmedById).length + ' confirmed-рядків');
  }

  // 1. Читаємо лист "Оплати" (read-only) — список дітей
  var paySheet = crmSS.getSheetByName(SHEET_PAYMENTS);
  if (!paySheet) {
    Logger.log('syncBdayStatusSheet: ERROR — лист "' + SHEET_PAYMENTS + '" не знайдено');
    return { ok: false, error: 'Оплати not found' };
  }
  var payData = paySheet.getDataRange().getValues();
  if (payData.length < 2) {
    Logger.log('syncBdayStatusSheet: лист "' + SHEET_PAYMENTS + '" порожній');
    return { ok: true, total: 0 };
  }
  var payHdr = payData[0].map(function(h){ return String(h || ''); });
  var locI   = payHdr.indexOf('Локація');
  var grpI   = payHdr.indexOf('Група');
  var nameI  = payHdr.indexOf("Ім'я дитини");
  if (locI < 0 || grpI < 0 || nameI < 0) {
    Logger.log('syncBdayStatusSheet: ERROR — не знайдено колонки Оплати: locI=' + locI + ' grpI=' + grpI + ' nameI=' + nameI);
    return { ok: false, error: 'Оплати headers missing' };
  }

  // 2. Будуємо candidatesByLoc з реєстрів
  var regs = getRegistries();
  var candidatesByLoc = {};
  var errors = [];
  var registriesScanned = 0;
  for (var ri = 0; ri < regs.length; ri++) {
    var reg = regs[ri];
    try {
      var ss = SpreadsheetApp.openById(reg.sheetId);
      var sh = ss.getSheetByName(reg.listName) || ss.getSheets()[0];
      if (!sh) { errors.push(reg.location + ': лист не знайдено'); continue; }
      var rData = sh.getDataRange().getValues();
      var rDataDisplay = sh.getDataRange().getDisplayValues();
      if (rData.length < 2) continue;
      var rHdr = rData[0];
      var rChildCol = -1, rBdayCol = -1, rCnCol = -1;
      for (var hi = 0; hi < rHdr.length; hi++) {
        var hLow = String(rHdr[hi] || '').toLowerCase();
        if (rChildCol < 0 && hLow.indexOf('піб дитини') >= 0) rChildCol = hi;
        if (rBdayCol < 0 && hLow.indexOf('дата народження') >= 0) rBdayCol = hi;
        if (rCnCol < 0 && hLow.indexOf('номер договору') >= 0) rCnCol = hi;
      }
      if (rChildCol < 0) {
        errors.push(reg.location + ': не знайдено колонки "ПІБ дитини"');
        continue;
      }
      registriesScanned++;
      if (!candidatesByLoc[reg.location]) candidatesByLoc[reg.location] = [];
      for (var dr = 1; dr < rData.length; dr++) {
        var fullName = String(rData[dr][rChildCol] || '').trim();
        if (!fullName) continue;
        var normName = _normChildName(fullName);
        var parts = normName.split(' ').filter(function(p){ return p; });
        if (parts.length < 1) continue;
        var normSurname = parts[0];
        var normFirstName = parts.slice(1).join(' ');
        var bday = rBdayCol >= 0 ? (parseRegistryBday(rData[dr][rBdayCol]) || '') : '';
        var cn = rCnCol >= 0 ? String(rDataDisplay[dr][rCnCol] || '').trim() : '';
        candidatesByLoc[reg.location].push({
          fullName: fullName,
          normName: normName,
          normSurname: normSurname,
          normFirstName: normFirstName,
          bday: bday,
          contractNumber: cn
        });
      }
    } catch (e) {
      errors.push(reg.location + ': ' + (e.message || e));
    }
  }
  Logger.log('syncBdayStatusSheet: candidates by loc — ' + Object.keys(candidatesByLoc).map(function(l){
    return l + ':' + candidatesByLoc[l].length;
  }).join(', '));

  // 3. Smart match для кожної дитини з Оплати
  var stats = { exact: 0, smart_check: 0, ambiguous: 0, school_no_contract: 0, not_found: 0, name_mismatch: 0, confirmed: 0 };
  var rowsOut = [];
  var nowStr = formatDate(new Date());

  for (var pr = 1; pr < payData.length; pr++) {
    var name = String(payData[pr][nameI] || '').trim();
    var loc  = String(payData[pr][locI]  || '').trim();
    var grp  = String(payData[pr][grpI]  || '').trim();
    if (!name || !loc) continue;

    var id = _childId(name, grp, loc);
    var norm = _normChildName(name);
    var nameParts = norm.split(' ').filter(function(p){ return p; });
    var status, bdayOut = '', cnOut = '', matchedRegOut = '';

    if (nameParts.length < 2) {
      status = 'not_found';
    } else {
      var surname = nameParts[0];
      var firstName = nameParts.slice(1).join(' ');
      var locCands = candidatesByLoc[loc] || [];
      var candidates = locCands.filter(function(c){ return c.normSurname === surname; });

      if (candidates.length === 0) {
        if (loc.toLowerCase().indexOf('школа') >= 0) {
          status = 'school_no_contract';
        } else {
          status = 'not_found';
        }
      } else if (candidates.length === 1) {
        var c = candidates[0];
        if (norm === c.normName) {
          status = 'exact';
          bdayOut = c.bday; cnOut = c.contractNumber; matchedRegOut = c.fullName;
        } else {
          var minLen = Math.min(firstName.length, c.normFirstName.length);
          var prefMatch = firstName.indexOf(c.normFirstName) === 0 || c.normFirstName.indexOf(firstName) === 0;
          if (minLen >= 3 && prefMatch) {
            status = 'smart_check';
            bdayOut = c.bday; cnOut = c.contractNumber; matchedRegOut = c.fullName;
          } else {
            status = 'name_mismatch';
            matchedRegOut = c.fullName;
          }
        }
      } else {
        // 2+ кандидатів — сортуємо за довжиною спільного префіксу імен
        var scored = candidates.map(function(c){
          return { c: c, score: _commonPrefixLen(c.normFirstName, firstName) };
        }).sort(function(a, b){ return b.score - a.score; });
        var best = scored[0];
        var second = scored[1];
        if (best.score - second.score <= 1) {
          status = 'ambiguous';
          var uniqBdays = {};
          candidates.forEach(function(c){ if (c.bday) uniqBdays[c.bday] = true; });
          var ub = Object.keys(uniqBdays);
          bdayOut = ub.length === 1 ? ub[0] : '';
          // contractNumber — НЕ записуємо при ambiguous
          matchedRegOut = candidates.map(function(c){ return c.fullName; }).join(' | ');
        } else {
          status = 'smart_check';
          bdayOut = best.c.bday;
          cnOut = best.c.contractNumber;
          matchedRegOut = best.c.fullName;
        }
      }
    }

    if (confirmedById[id]) {
      var preserved = confirmedById[id];
      var padded = preserved.slice(0, HEADER.length);
      while (padded.length < HEADER.length) padded.push('');
      rowsOut.push(padded);
      stats.confirmed = (stats.confirmed || 0) + 1;
      continue;
    }

    stats[status] = (stats[status] || 0) + 1;
    rowsOut.push([id, name, loc, bdayOut, cnOut, status, matchedRegOut, nowStr, '', '']);
  }

  // 4. Запис у sheet (clear + write)
  statusSheet.clearContents();
  statusSheet.getRange(1, 1, 1, HEADER.length).setValues([HEADER]);
  statusSheet.setFrozenRows(1);
  statusSheet.getRange(2, 5, Math.max(rowsOut.length, 1), 1).setNumberFormat('@');
  if (rowsOut.length > 0) {
    statusSheet.getRange(2, 1, rowsOut.length, HEADER.length).setValues(rowsOut);
  }

  // 5. Звіт
  var report = {
    ok: true,
    total: rowsOut.length,
    exact: stats.exact || 0,
    smart_check: stats.smart_check || 0,
    ambiguous: stats.ambiguous || 0,
    school_no_contract: stats.school_no_contract || 0,
    not_found: stats.not_found || 0,
    name_mismatch: stats.name_mismatch || 0,
    confirmed: stats.confirmed || 0,
    registriesScanned: registriesScanned,
    errors: errors
  };
  Logger.log('=== syncBdayStatusSheet REPORT ===');
  Logger.log(JSON.stringify(report, null, 2));
  return report;
}

function runSyncBdayStatus() {
  var r = syncBdayStatusSheet();
  var summary = !r.ok
    ? 'Помилка: ' + (r.error || 'unknown')
    : 'Всього дітей в Оплати: ' + r.total + '\n' +
      '─────────────────────\n' +
      '✅ exact:              ' + r.exact + '\n' +
      '⚠️ smart_check:        ' + r.smart_check + '\n' +
      '⚠️ ambiguous:          ' + r.ambiguous + '\n' +
      '⚠️ name_mismatch:      ' + r.name_mismatch + '\n' +
      '✅ confirmed:          ' + r.confirmed + '\n' +
      '📝 school_no_contract: ' + r.school_no_contract + '\n' +
      '❌ not_found:          ' + r.not_found + '\n' +
      '─────────────────────\n' +
      'Реєстрів просканено: ' + r.registriesScanned + '\n' +
      (r.errors && r.errors.length ? '\nПомилки:\n' + r.errors.join('\n') : '');
  try {
    SpreadsheetApp.getUi().alert('bday_sync_status', summary, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    Logger.log('UI alert недоступний. Звіт у логах.');
    Logger.log(summary);
  }
}

// Знаходить рядок у bday_sync_status за ChildID. Повертає {sheet, rowIndex, header}
// або null якщо не знайдено. rowIndex — 1-based (для setValues).
function _findBdayStatusRow(childId) {
  if (!childId) return null;
  var crmSS = getCRMSpreadsheet();
  var sh = crmSS.getSheetByName(BDAY_STATUS_SHEET);
  if (!sh) return null;
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return null;
  var lastCol = sh.getLastColumn();
  var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var idIdx = hdr.indexOf('ChildID');
  if (idIdx < 0) return null;
  var ids = sh.getRange(2, idIdx + 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]||'').trim() === String(childId).trim()) {
      return { sheet: sh, rowIndex: i + 2, header: hdr };
    }
  }
  return null;
}

// Підтверджує збіг дитини з реєстром (статус → 'confirmed').
function confirmBdayMatch(childId, confirmedBy) {
  var loc = _findBdayStatusRow(childId);
  if (!loc) return { ok: false, error: 'ChildID не знайдено в ' + BDAY_STATUS_SHEET };
  var hdr = loc.header;
  var statusIdx      = hdr.indexOf('Status');
  var updatedAtIdx   = hdr.indexOf('UpdatedAt');
  var confirmedByIdx = hdr.indexOf('ConfirmedBy');
  var confirmedAtIdx = hdr.indexOf('ConfirmedAt');
  if (statusIdx < 0 || confirmedByIdx < 0 || confirmedAtIdx < 0) {
    return { ok: false, error: 'Немає колонок Status/ConfirmedBy/ConfirmedAt — потрібен повний sync' };
  }
  var nowStr = formatDate(new Date());
  loc.sheet.getRange(loc.rowIndex, statusIdx + 1).setValue('confirmed');
  loc.sheet.getRange(loc.rowIndex, confirmedByIdx + 1).setValue(String(confirmedBy || ''));
  loc.sheet.getRange(loc.rowIndex, confirmedAtIdx + 1).setValue(nowStr);
  if (updatedAtIdx >= 0) loc.sheet.getRange(loc.rowIndex, updatedAtIdx + 1).setValue(nowStr);
  return { ok: true, status: 'confirmed', confirmedBy: confirmedBy, confirmedAt: nowStr };
}

// Скасовує підтвердження. Статус відновлюється з MatchedRegName:
//   - "A | B" → ambiguous; інакше → smart_check (catch-all для smart_check/name_mismatch)
//   - наступний повний sync уточнить статус.
function unconfirmBdayMatch(childId) {
  var loc = _findBdayStatusRow(childId);
  if (!loc) return { ok: false, error: 'ChildID не знайдено в ' + BDAY_STATUS_SHEET };
  var hdr = loc.header;
  var statusIdx        = hdr.indexOf('Status');
  var matchedRegIdx    = hdr.indexOf('MatchedRegName');
  var updatedAtIdx     = hdr.indexOf('UpdatedAt');
  var confirmedByIdx   = hdr.indexOf('ConfirmedBy');
  var confirmedAtIdx   = hdr.indexOf('ConfirmedAt');
  if (statusIdx < 0) return { ok: false, error: 'Немає колонки Status' };
  var matchedReg = matchedRegIdx >= 0
    ? String(loc.sheet.getRange(loc.rowIndex, matchedRegIdx + 1).getValue() || '').trim()
    : '';
  var newStatus = matchedReg.indexOf(' | ') >= 0 ? 'ambiguous' : 'smart_check';
  var nowStr = formatDate(new Date());
  loc.sheet.getRange(loc.rowIndex, statusIdx + 1).setValue(newStatus);
  if (confirmedByIdx >= 0) loc.sheet.getRange(loc.rowIndex, confirmedByIdx + 1).setValue('');
  if (confirmedAtIdx >= 0) loc.sheet.getRange(loc.rowIndex, confirmedAtIdx + 1).setValue('');
  if (updatedAtIdx >= 0)   loc.sheet.getRange(loc.rowIndex, updatedAtIdx + 1).setValue(nowStr);
  return { ok: true, status: newStatus };
}

// ═══════════════════════════════════════════════════════════════════════════
// CLEANUP: видалити timestamp/дати з колонки "Номер договору" в CRM Клієнти
// ═══════════════════════════════════════════════════════════════════════════
//
// Сміттєві дані потрапили туди від старого migrateChildrenBdays та інших
// скриптів. Реальні номери (формат "75-30-08") живуть у реєстрах і
// підтягуються через bday_sync_status. Очистка ХОВАЄ старі timestamp.

function cleanContractNumberGarbage(dryRun) {
  if (dryRun === undefined) dryRun = true;
  var ss = getCRMSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_CLIENTS);
  if (!sheet) {
    Logger.log('cleanContractNumberGarbage: ERROR — sheet "' + SHEET_CLIENTS + '" not found');
    return { ok: false, error: 'Sheet not found' };
  }

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    Logger.log('cleanContractNumberGarbage: empty');
    return { ok: true, total: 0, garbage: 0, real: 0, empty: 0 };
  }

  var hdr = data[0].map(function(h){ return String(h || ''); });
  var cnCol = hdr.indexOf('Номер договору');
  if (cnCol < 0) {
    Logger.log('cleanContractNumberGarbage: ERROR — колонка "Номер договору" не знайдена');
    return { ok: false, error: 'Column "Номер договору" not found' };
  }
  Logger.log('cleanContractNumberGarbage: колонка "Номер договору" — index ' + cnCol + ' (col ' + (cnCol + 1) + ')');

  // Регекси для виявлення сміття у колонці "Номер договору"
  var GARBAGE_PATTERNS = [
    /^\d{4}-\d{2}-\d{2}/,                          // ISO: 2026-04-20 (з можливим хвостом)
    /^\d{1,2}\.\d{1,2}\.\d{4}/                     // D.M.YYYY або DD.MM.YYYY (з timestamp або без)
  ];

  function isGarbage(val) {
    if (val === null || val === undefined || val === '') return false;
    if (val instanceof Date) return true;          // Excel timestamp / Date object
    var s = String(val).trim();
    if (!s) return false;
    for (var i = 0; i < GARBAGE_PATTERNS.length; i++) {
      if (GARBAGE_PATTERNS[i].test(s)) return true;
    }
    return false;
  }

  var stats = { total: 0, garbage: 0, real: 0, empty: 0 };
  var garbageRows = []; // [{row, value}]
  var realSamples = []; // приклади що залишимо (до 10)

  for (var r = 1; r < data.length; r++) {
    stats.total++;
    var val = data[r][cnCol];
    if (val === null || val === undefined || String(val).trim() === '') {
      stats.empty++;
      continue;
    }
    if (isGarbage(val)) {
      stats.garbage++;
      if (garbageRows.length < 1000) garbageRows.push({ row: r + 1, value: String(val) });
    } else {
      stats.real++;
      if (realSamples.length < 10) realSamples.push(String(val));
    }
  }

  Logger.log('=== cleanContractNumberGarbage ' + (dryRun ? 'DRY RUN' : 'REAL') + ' ===');
  Logger.log('Total rows:                  ' + stats.total);
  Logger.log('Garbage (timestamps/dates):  ' + stats.garbage);
  Logger.log('Real numbers (kept):         ' + stats.real);
  Logger.log('Empty (kept):                ' + stats.empty);
  Logger.log('');
  Logger.log('Sample REAL numbers (first 10) — НЕ чіпатимемо:');
  realSamples.forEach(function(s){ Logger.log('  "' + s + '"'); });
  Logger.log('');
  Logger.log('Sample GARBAGE values (first 10) — будемо очищати:');
  garbageRows.slice(0, 10).forEach(function(g){
    Logger.log('  row ' + g.row + ': "' + g.value + '"');
  });

  if (dryRun) {
    Logger.log('');
    Logger.log('DRY RUN — нічого не змінено. Для реальної очистки запустіть runCleanGarbageReal().');
    return { ok: true, dryRun: true, total: stats.total, garbage: stats.garbage, real: stats.real, empty: stats.empty };
  }

  // REAL: пишемо очищену колонку одним batch setValues
  if (garbageRows.length > 0) {
    var newCol = [];
    for (var rr = 1; rr < data.length; rr++) {
      var v = data[rr][cnCol];
      if (isGarbage(v)) {
        newCol.push(['']);
      } else {
        newCol.push([(v === null || v === undefined) ? '' : v]);
      }
    }
    sheet.getRange(2, cnCol + 1, newCol.length, 1).setValues(newCol);
    Logger.log('REAL — очищено ' + garbageRows.length + ' клітинок у колонці "Номер договору"');
  } else {
    Logger.log('REAL — нічого очищати, сміття не знайдено');
  }

  return { ok: true, dryRun: false, cleaned: garbageRows.length, total: stats.total, garbage: stats.garbage, real: stats.real, empty: stats.empty };
}

function runCleanGarbageDryRun() {
  var r = cleanContractNumberGarbage(true);
  var summary = !r.ok
    ? 'Помилка: ' + (r.error || 'unknown')
    : 'DRY RUN — нічого не змінено\n' +
      '─────────────────────\n' +
      'Всього рядків:       ' + r.total + '\n' +
      'Сміття (timestamp):  ' + r.garbage + '\n' +
      'Справжні номери:     ' + r.real + '\n' +
      'Порожні:             ' + r.empty + '\n' +
      '─────────────────────\n' +
      'Деталі (приклади) — у Logger.\n' +
      'Для реальної очистки: runCleanGarbageReal()';
  try {
    SpreadsheetApp.getUi().alert('Очистка сміття (DRY RUN)', summary, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    Logger.log('UI alert недоступний.');
    Logger.log(summary);
  }
}

function runCleanGarbageReal() {
  var confirmed = false;
  try {
    var ui = SpreadsheetApp.getUi();
    var resp = ui.alert(
      '⚠️ Очистка сміття (РЕАЛЬНО)',
      'УВАГА! Спершу зробіть копію CRM Sheets:\n' +
      'File → Make a copy.\n\n' +
      'Ця операція очистить колонку "Номер договору" у листі "Клієнти" ' +
      'від timestamp/дат (наприклад "25.04.2026 13:03", "2026-04-20"). ' +
      'Справжні номери (формат "75-30-08", з літерами тощо) — залишаються.\n\n' +
      'Продовжити?',
      ui.ButtonSet.YES_NO
    );
    confirmed = (resp === ui.Button.YES);
  } catch (e) {
    // Запущено з editor — без UI підтвердження. Логуємо WARNING.
    Logger.log('runCleanGarbageReal: UI недоступний (editor) — продовжуємо без підтвердження');
    confirmed = true;
  }
  if (!confirmed) {
    Logger.log('runCleanGarbageReal: користувач відхилив');
    return;
  }

  var r = cleanContractNumberGarbage(false);
  var summary = !r.ok
    ? 'Помилка: ' + (r.error || 'unknown')
    : '✅ Очистку завершено\n' +
      '─────────────────────\n' +
      'Видалено timestamp:        ' + (r.cleaned || 0) + '\n' +
      'Справжні номери збережено: ' + r.real + '\n' +
      'Порожні без змін:          ' + r.empty;
  try {
    SpreadsheetApp.getUi().alert('Очистка завершена', summary, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    Logger.log(summary);
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// TRIGGERS: автоматична щоденна синхронізація bday_sync_status
// ═══════════════════════════════════════════════════════════════════════════

function installSyncTrigger() {
  var existing = ScriptApp.getProjectTriggers();
  var removed = 0;
  existing.forEach(function(t){
    if (t.getHandlerFunction() === 'runSyncBdayStatus') {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  });
  ScriptApp.newTrigger('runSyncBdayStatus')
    .timeBased()
    .everyDays(1)
    .atHour(4)
    .inTimezone('Europe/Kiev')
    .create();
  Logger.log('✅ Тригер встановлено: щодня о 04:00 Europe/Kiev (видалено старих: ' + removed + ')');
  try {
    SpreadsheetApp.getUi().alert(
      'Тригер встановлено',
      '✅ runSyncBdayStatus запускатиметься щодня о 04:00 (Київ).\n' +
      'Видалено старих тригерів: ' + removed,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch(e) {}
}

function removeSyncTrigger() {
  var existing = ScriptApp.getProjectTriggers();
  var removed = 0;
  existing.forEach(function(t){
    if (t.getHandlerFunction() === 'runSyncBdayStatus') {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  });
  Logger.log('🗑️ Тригери runSyncBdayStatus видалено: ' + removed);
  try {
    SpreadsheetApp.getUi().alert('Тригер видалено',
      'Видалено тригерів: ' + removed,
      SpreadsheetApp.getUi().ButtonSet.OK);
  } catch(e) {}
}

function listAllTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  Logger.log('=== Усі тригери проєкту (' + triggers.length + ') ===');
  triggers.forEach(function(t, i){
    Logger.log('[' + (i+1) + '] handler="' + t.getHandlerFunction() +
      '" eventType=' + t.getEventType() +
      ' uniqueId=' + t.getUniqueId());
  });
  return triggers.length;
}
