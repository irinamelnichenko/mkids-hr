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

function saveClient(data) {
  if (!data || !data.id) return {ok:false, error:'Missing id'};
  var ss = getCRMSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_CLIENTS);
  if (!sheet) return {ok:false, error:'Sheet not found'};
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
    data.createdAt||now, now
  ];
  for (var r = 1; r < vals.length; r++) {
    if (String(vals[r][0]) === String(data.id)) {
      row[19] = vals[r][19] || data.createdAt || now;
      sheet.getRange(r+1, 1, 1, row.length).setValues([row]);
      return {ok:true, action:'updated'};
    }
  }
  sheet.appendRow(row);
  return {ok:true, action:'created'};
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
    for (var r = 1; r < vals.length; r++) {
      if (String(vals[r][0]) === date && String(vals[r][1]) === childId) {
        sheet.getRange(r+1, 1, 1, row.length).setValues([row]);
        vals[r] = row;
        saved++;
        return;
      }
    }
    sheet.appendRow(row);
    vals.push(row);
    saved++;
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
  if (!str) return null;

  // Date-об'єкт з Google Sheets — трактуємо як MM/YYYY (перший тиждень місяця)
  if (str instanceof Date) {
    if (isNaN(str.getTime())) return null;
    return parseAbsencePeriod(pad2(str.getMonth() + 1) + '/' + str.getFullYear(), refYear);
  }

  var s = trim(String(str));
  if (!s || s === '-' || s.toLowerCase() === 'по') return null;

  // Формат 1: "01.09.2024-15.09.2024"
  var m1 = s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})\s*[-–]\s*(\d{1,2})\.(\d{1,2})\.(\d{4})$/);
  if (m1) {
    return {
      from: m1[3] + '-' + pad2(m1[2]) + '-' + pad2(m1[1]),
      to:   m1[6] + '-' + pad2(m1[5]) + '-' + pad2(m1[4])
    };
  }

  // Формат 2: "09.12-15.12"
  var m2 = s.match(/^(\d{1,2})\.(\d{1,2})\s*[-–]\s*(\d{1,2})\.(\d{1,2})$/);
  if (m2) {
    var fromMon = +m2[2], toMon = +m2[4];
    var nowMon = new Date().getMonth() + 1;
    var fy = (fromMon >= nowMon) ? (refYear - 1) : refYear;
    var ty = (toMon   >= nowMon) ? (refYear - 1) : refYear;
    return {
      from: fy + '-' + pad2(m2[2]) + '-' + pad2(m2[1]),
      to:   ty + '-' + pad2(m2[4]) + '-' + pad2(m2[3])
    };
  }

  // Формат 3: "15-20.01"
  var m3 = s.match(/^(\d{1,2})\s*[-–]\s*(\d{1,2})\.(\d{1,2})$/);
  if (m3) {
    var mon = +m3[3];
    var nowMon = new Date().getMonth() + 1;
    var yr = (mon >= nowMon) ? (refYear - 1) : refYear;
    return {
      from: yr + '-' + pad2(m3[3]) + '-' + pad2(m3[1]),
      to:   yr + '-' + pad2(m3[3]) + '-' + pad2(m3[2])
    };
  }

  // Формат 4: "MM/YY", "MM|YY", "MM/YYYY", "MM|YYYY"
  // "10/25" → 1 тиждень у жовтні 2025: перший повний робочий тиждень місяця
  var m4 = s.match(/^(\d{1,2})\s*[\/|]\s*(\d{2}|\d{4})$/);
  if (m4) {
    var mon4 = +m4[1];
    var yr4  = m4[2].length === 2 ? 2000 + (+m4[2]) : +m4[2];
    if (mon4 >= 1 && mon4 <= 12) {
      // Знаходимо перший робочий день місяця
      var d = new Date(yr4, mon4 - 1, 1);
      while (d.getDay() === 0 || d.getDay() === 6) { d.setDate(d.getDate() + 1); }
      var fromD = new Date(d);
      var toD   = new Date(d); toD.setDate(toD.getDate() + 4);
      return {
        from: fromD.getFullYear() + '-' + pad2(fromD.getMonth()+1) + '-' + pad2(fromD.getDate()),
        to:   toD.getFullYear()   + '-' + pad2(toD.getMonth()+1)   + '-' + pad2(toD.getDate()),
        _synthetic:    true,  // дати умовні — точний тиждень невідомий
        _originalRaw:  str    // оригінальний рядок зі слоту
      };
    }
  }

  return null;
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
  if (!crmSheet) return map;
  var crmData = crmSheet.getDataRange().getValues();
  if (crmData.length < 2) return map;
  var hdrs    = crmData[0].map(String);
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
  for (var ri = 1; ri < crmData.length; ri++) {
    var rName = norm(crmData[ri][colName] || '');
    var rLoc  = norm(crmData[ri][colLoc]  || '');
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
      var curGroup = '(без групи)'; var curTeacher = '';

      var DATA_START = 3;
      for (var row = DATA_START; row < data.length; row++) {
        var nameCell = trim(String(data[row][0] || ''));
        if (!nameCell) continue;

        // Рядок-заголовок групи → оновлюємо контекст
        if (isGroupHeaderRow(data[row], 1)) {
          var firstSpace = nameCell.search(/\s/);
          curTeacher = firstSpace > 0 ? nameCell.slice(firstSpace).trim() : '';
          curGroup   = normalizeGroupName(nameCell) + (curTeacher ? ' ' + curTeacher : '');
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
              id:               'import_' + Date.now() + '_' + Math.random().toString(36).slice(2,8),
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
          } else {
            var existing = crmMap[crmKey];
            clientData = {
              id:               existing.id,
              name:             existing.name,
              loc:              existing.loc,
              group:            existing.group,
              teacher:          existing.teacher,
              contractDate:     existing.contractDate,
              contractType:     existing.contractType,
              monthlyFee:       existing.monthlyFee,
              entryFee:         0,
              status:           'active',
              notes:            existing.notes,
              bday: '', momName: '', momPhone: '', dadName: '', dadPhone: '',
              absences:         allAbsences,
              entryFeeSchedule: [],
              feeHistory:       []
            };
          }

          var saveResult = saveClient(clientData);
          if (!saveResult.ok) {
            stats.errors.push({loc:loc, child:nameCell, error: saveResult.error || 'saveClient failed'});
            continue;
          }

          // Оновлюємо crmMap щоб наступна ітерація бачила оновлені дані
          crmMap[crmKey] = {
            id: clientData.id, name: clientData.name, loc: clientData.loc,
            group: clientData.group, teacher: clientData.teacher,
            contractDate: clientData.contractDate, contractType: clientData.contractType,
            monthlyFee: clientData.monthlyFee, notes: clientData.notes,
            absences: allAbsences
          };

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
