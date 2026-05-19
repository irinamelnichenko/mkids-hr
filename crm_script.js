// ═══════════════════════════════════════════════════════════════════════════
// m.kids CRM — Google Apps Script v6.7
// v6.7: міграція директорів і медсестер у єдиний лист "Користувачі";
//        паролі SHA-256; addAllDirectorsAndNurses() — разова утиліта
// v6.6: Задачник — управління задачами в команді; листи "Задачі" +
//        "Задачі_Активність"; email-нагадування; daily-тригер о 09:00
// v6.5: вчителі-предметники — групові заняття у вартості навчання;
//        ЗП = ставка × унікальні (група+дата); пише ТІЛЬКИ у Salary
// v6.1: автоекспорт пише У ФАЙЛИ ЛОКАЦІЙ (а не в CRM-зведення)
//        — Пейменти: "Голосієво Payment" / "Осокорки Payment" тощо
//          Колонка "Бюджет доп" місяця = monthCol + 3
//          Місяць експорту = місяць_відмітки + 1 (травень → бюджет червня)
//        — Salary: файл локації, лист Salary, Budget-колонка місяця+1
// v6.2: толерантний матч імен (_normName: lowercase + без whitespace/NBSP)
// v6.3: розумне перезаписування через лист "Експорт_Журнал"
//        base = currentCell - lastWritten; new = base + newSum
// v6.4: точковий запис у Payment/Salary (НЕ setValues на стовпець) —
//        не затирає формули у підсумкових рядках; seedActivitiesCatalog
// 5 колонок на місяць: Факт навч | Факт вступ | Факт доп | Бюджет доп | Бюджет навч
// ═══════════════════════════════════════════════════════════════════════════

// Зручна обгортка для запуску dry-run з Apps Script editor (Run не передає
// аргументи). Запускає матчинг предметників → Salary за травень 2026.
function runDryRunMay2026() {
  return exportAllPredmetnyToSalary_DRY_RUN(5, 2026);
}

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
    if      (action === 'ping')               result = {ok:true, msg:'pong v6.5', ts: new Date().toISOString()};
    else if (action === 'getLocations')       result = getLocations();
    else if (action === 'getPayments')        result = getPayments();
    else if (action === 'getPaymentsYearly')  result = getPaymentsYearly();
    else if (action === 'getClients')         result = getClients();
    else if (action === 'runAggregate')       result = aggregatePayments();
    else if (action === 'syncPayments')        result = syncPayments();
    else if (action === 'runAggregateYearly') result = aggregatePaymentsYearly();
    else if (action === 'runSyncBdayStatus')  result = syncBdayStatusSheet();
    else if (action === 'getRegistryUrls')    result = getRegistryUrls();
    else if (action === 'makePublic')         result = makeSheetPublic();
    else if (action === 'getAttendance')      result = getAttendance(e);
    else if (action === 'getHealthRecords')         result = getHealthRecords(e);
    else if (action === 'dryRunImportAbsences')      result = dryRunImportAbsences(e.parameter.loc || '');
    else if (action === 'importAbsencesFromPayment') result = importAbsencesFromPayment(e.parameter.loc || '');
    else if (action === 'getOpexData')               result = getOpexData(e.parameter.loc || '', e.parameter.year || '');
    else if (action === 'getOpexOverview')           result = getOpexOverview(e.parameter.year || '');
    else if (action === 'getCategoryAnalytics')      result = getCategoryAnalytics(e.parameter.year || '', e.parameter.month || '');
    else if (action === 'getSalaryData')             result = getSalaryData(e.parameter.loc || '', e.parameter.year || '');
    else if (action === 'getSalaryOverview')         result = getSalaryOverview(e.parameter.year || '');
    else if (action === 'getOverviewAnalytics')      result = getOverviewAnalytics(e.parameter.year || '', e.parameter.month || '');
    else if (action === 'getUsers')                  result = getUsers();
    else if (action === 'getGroupNorms')             result = getGroupNorms();
    else if (action === 'getActivitiesCatalog')      result = getActivitiesCatalog(e.parameter && e.parameter.loc || '');
    else if (action === 'getAttendanceMarks')         result = getAttendanceMarks(e.parameter || {});
    else if (action === 'getPredmetnyCatalog')        result = getPredmetnyCatalog(e.parameter && e.parameter.loc || '');
    else if (action === 'getPredmetnyMarks')          result = getPredmetnyMarks(e.parameter || {});
    else if (action === 'getTasks')                   result = getTasks(e.parameter || {});
    else if (action === 'getTaskActivity')            result = getTaskActivity(e.parameter && e.parameter.taskId || 0);
    else if (action === 'getDashboardNotifications')  result = getDashboardNotifications(e.parameter && e.parameter.userId || 0, e.parameter && e.parameter.role || '');
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
    else if (body.action === 'authenticate')              result = authenticate(body.login || '', body.password || '');
    else if (body.action === 'updatePassword')            result = updatePassword(body.userId || 0, body.newPassword || '');
    else if (body.action === 'addUser')                   result = addUser(body.data || {});
    else if (body.action === 'deactivateUser')            result = deactivateUser(body.userId || 0);
    else if (body.action === 'activateUser')              result = activateUser(body.userId || 0);
    else if (body.action === 'syncPayments')              result = syncPayments();
    else if (body.action === 'addActivity')               result = addActivity(body.data || {});
    else if (body.action === 'updateActivity')            result = updateActivity(body.id || 0, body.data || {});
    else if (body.action === 'deleteActivity')            result = deleteActivity(body.id || 0);
    else if (body.action === 'copyActivitiesFromLocation') result = copyActivitiesFromLocation(body.fromLoc || '', body.toLoc || '');
    else if (body.action === 'addAttendanceMark')         result = addAttendanceMark(body.data || {});
    else if (body.action === 'removeAttendanceMark')      result = removeAttendanceMark(body.id || 0);
    else if (body.action === 'exportAttendanceToPayments') result = exportAttendanceToPayments(body || {});
    else if (body.action === 'exportToSalaryExtras')      result = exportToSalaryExtras(body || {});
    else if (body.action === 'exportAttendance')          result = exportAttendance(body || {});
    else if (body.action === 'addPredmetny')              result = addPredmetny(body.data || {});
    else if (body.action === 'updatePredmetny')           result = updatePredmetny(body.id || 0, body.data || {});
    else if (body.action === 'deletePredmetny')           result = deletePredmetny(body.id || 0);
    else if (body.action === 'addPredmetnyMark')          result = addPredmetnyMark(body.data || {});
    else if (body.action === 'removePredmetnyMark')       result = removePredmetnyMark(body.id || 0);
    else if (body.action === 'getPredmetnyMarks')         result = getPredmetnyMarks(body || {});
    else if (body.action === 'exportPredmetnyToSalary')   result = exportPredmetnyToSalary(body || {});
    else if (body.action === 'createTask')                result = createTask(body.data || {});
    else if (body.action === 'updateTaskStatus')          result = updateTaskStatus(body.taskId || 0, body.status || '', body.actorId || 0);
    else if (body.action === 'updateTask')                result = updateTask(body.taskId || 0, body.data || {}, body.actorId || 0);
    else if (body.action === 'addTaskComment')            result = addTaskComment(body.taskId || 0, body.comment || '', body.fileUrl || '', body.actorId || 0);
    else if (body.action === 'deleteTask')                result = deleteTask(body.taskId || 0, body.actorId || 0);
    else if (body.action === 'setUserPassword')           result = setUserPassword(body.username || '', body.newPassword || '', body.actorId || 0);
    else if (body.action === 'resetAllLocationPasswords') result = resetAllLocationPasswords(body.actorId || 0);
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

    m = n.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})[-–](\d{1,2})\.(\d{1,2})\.(\d{4})$/);
    if (m) {
      return {
        from: m[3] + '-' + pad2(m[2]) + '-' + pad2(m[1]),
        to:   m[6] + '-' + pad2(m[5]) + '-' + pad2(m[4])
      };
    }

    m = n.match(/^(\d{1,2})\.(\d{1,2})[-–](\d{1,2})\.(\d{1,2})\.(\d{2}|\d{4})$/);
    if (m) {
      var yr1b = m[5].length === 2 ? 2000 + (+m[5]) : +m[5];
      return {
        from: yr1b + '-' + pad2(m[2]) + '-' + pad2(m[1]),
        to:   yr1b + '-' + pad2(m[4]) + '-' + pad2(m[3])
      };
    }

    m = n.match(/^(\d{1,2})\.(\d{1,2})[-–](\d{1,2})\.(\d{1,2})$/);
    if (m) {
      return {
        from: yearFor(+m[2]) + '-' + pad2(m[2]) + '-' + pad2(m[1]),
        to:   yearFor(+m[4]) + '-' + pad2(m[4]) + '-' + pad2(m[3])
      };
    }

    m = n.match(/^(\d{1,2})\.(\d{1,2})[-–](\d{1,2})[-–](\d{1,2})$/);
    if (m) {
      return {
        from: yearFor(+m[2]) + '-' + pad2(m[2]) + '-' + pad2(m[1]),
        to:   yearFor(+m[4]) + '-' + pad2(m[4]) + '-' + pad2(m[3])
      };
    }

    m = n.match(/^(\d{1,2})[-–](\d{1,2})\.(\d{1,2})$/);
    if (m) {
      var mon3 = +m[3];
      return {
        from: yearFor(mon3) + '-' + pad2(m[3]) + '-' + pad2(m[1]),
        to:   yearFor(mon3) + '-' + pad2(m[3]) + '-' + pad2(m[2])
      };
    }

    m = n.match(/^(\d{1,2})[-–](\d{1,2})\.(\d{1,2})\.(\d{2}|\d{4})$/);
    if (m) {
      var yr3b = m[4].length === 2 ? 2000 + (+m[4]) : +m[4];
      return {
        from: yr3b + '-' + pad2(m[3]) + '-' + pad2(m[1]),
        to:   yr3b + '-' + pad2(m[3]) + '-' + pad2(m[2])
      };
    }

    m = n.match(/^(\d{1,2})[\/|.](\d{2}|\d{4})$/);
    if (m) {
      var mon4 = +m[1];
      var yr4  = m[2].length === 2 ? 2000 + (+m[2]) : +m[2];
      if (mon4 >= 1 && mon4 <= 12) {
        return syntheticWeek(yr4, mon4);
      }
    }

    m = s.match(/^([а-яіїє']+?)\.?\s*(\d{2}|\d{4})$/);
    if (m) {
      var mon5 = UA_MONTHS[m[1].trim()];
      if (mon5) {
        var yr5 = m[2].length === 2 ? 2000 + (+m[2]) : +m[2];
        return syntheticWeek(yr5, mon5);
      }
    }

    m = s.match(/^([а-яіїє']+)\s+(\d{4})$/);
    if (m) {
      var mon6 = UA_MONTHS[m[1].trim()];
      if (mon6) {
        return syntheticWeek(+m[2], mon6);
      }
    }

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

    var mon9 = UA_MONTHS[s.trim()];
    if (mon9) {
      return syntheticWeek(yearFor(mon9), mon9);
    }

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
    var nameColIdx = 0;

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

function _countWorkDays(fromStr, toStr) {
  if (!fromStr || !toStr) return 0;
  var f = new Date(fromStr); var t = new Date(toStr);
  if (isNaN(f.getTime()) || isNaN(t.getTime()) || t < f) return 0;
  var n = 0; var cur = new Date(f.getTime());
  while (cur <= t) { var d = cur.getDay(); if (d !== 0 && d !== 6) n++; cur.setDate(cur.getDate()+1); }
  return n;
}

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
        byLocation[loc] = locStat;
        continue;
      }

      var DATA_START = 3;
      for (var row = DATA_START; row < data.length; row++) {
        var nameCell = trim(String(data[row][0] || ''));
        if (!nameCell) continue;
        if (isGroupHeaderRow(data[row], 1)) continue;

        var hasAnySlot = false;
        for (var si = 0; si < absCols.length; si++) {
          if (absCols[si] !== null && trim(String(data[row][absCols[si]] || ''))) { hasAnySlot = true; break; }
        }
        if (!hasAnySlot) { locStat.skipped++; totalStats.wouldSkipNoAbsence++; continue; }

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

  return {ok:true, stats:totalStats, byLocation:byLocation, unparsedSamples:unparsedSamples};
}

function importAbsencesFromPayment(locFilter) {
  var SCHOOL_LOCS_SKIP = ['Школа Осокорки', 'Школа 228', 'Онлайн школа'];
  var refYear = new Date().getFullYear();
  var norm    = function(s){ return String(s||'').trim().toLowerCase().replace(/\s+/g,' '); };
  var nowISO  = new Date().toISOString();
  var todayUA = Utilities.formatDate(new Date(), 'Europe/Kiev', 'dd.MM.yyyy');

  var crmMap = _loadCRMClientsMap(norm);

  var configSS    = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  var configSheet = configSS.getSheets()[0];
  var configData  = configSheet.getDataRange().getValues();

  var stats = {
    locationsProcessed:0,
    newClientsCreated:0, existingClientsUpdated:0,
    absencesAdded:0, absencesPlaceholder:0, absencesDuplicates:0,
    errors:[]
  };

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
      if (absCols[0] === null) continue;

      var curGroup = '(без групи)'; var curGroupType = ''; var curTeacher = '';

      var DATA_START = 3;
      for (var row = DATA_START; row < data.length; row++) {
        var nameCell = trim(String(data[row][0] || ''));
        if (!nameCell) continue;

        if (isGroupHeaderRow(data[row], 1)) {
          var firstSpace = nameCell.search(/\s/);
          curTeacher    = firstSpace > 0 ? nameCell.slice(firstSpace).trim() : '';
          curGroupType  = normalizeGroupName(nameCell);
          curGroup      = curGroupType + (curTeacher ? ' ' + curTeacher : '');
          continue;
        }

        var hasAnySlot = false;
        for (var si = 0; si < absCols.length; si++) {
          if (absCols[si] !== null && trim(String(data[row][absCols[si]] || ''))) { hasAnySlot = true; break; }
        }
        if (!hasAnySlot) continue;

        try {
          var crmKey  = norm(nameCell) + '|' + norm(loc);
          var isNew   = !crmMap.hasOwnProperty(crmKey);
          var existingAbsences = isNew ? [] : crmMap[crmKey].absences.slice();

          var existingPairs = {};
          existingAbsences.forEach(function(a){ if(a.from&&a.to) existingPairs[a.from+'|'+a.to]=true; });

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
              }
            } else {
              var absPlaceholder = _makeImportAbsence(null, slot);
              newAbsences.push(absPlaceholder);
              stats.absencesPlaceholder++;
            }
          }

          if (newAbsences.length === 0) continue;

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
            var saveResult = saveClient(clientData);
          } else {
            var existing = crmMap[crmKey];
            var saveResult = patchClientAbsences(existing.id, allAbsences);
          }

          if (!saveResult.ok) {
            stats.errors.push({loc:loc, child:nameCell, error: saveResult.error || 'saveClient failed'});
            continue;
          }

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
        }
      }

    } catch(locErr) {
      stats.errors.push({loc:loc, child:'', error: locErr.message || String(locErr)});
    }
  }

  return {ok:true, stats:stats};
}

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
    if (newVal === null) return;

    var map = loadNurseSheetMap();
    var sid = map[norm(loc)];
    if (!sid) return;

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
    if (!sheet) return;

    var rowNum = findChildRow(sheet, childName);
    if (rowNum < 1) return;
    var colNum = findDateColumn(sheet, isoDate);
    if (colNum < 1) return;

    var cell = sheet.getRange(rowNum, colNum);
    var existing = String(cell.getValue() || '').trim();
    if (existing.toUpperCase() === 'А' || existing.toUpperCase() === 'A') return;
    if (existing === newVal) return;
    cell.setValue(newVal);
  } catch (e) {
    Logger.log('mirror: ERROR ' + (e.message || e));
  }
}

var REGISTRY_TAB_NAME = 'реєстр';

function getRegistries() {
  var configSS = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  var sheet = configSS.getSheetByName(REGISTRY_TAB_NAME);
  if (!sheet) return [];
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

function parseRegistryBday(val) {
  if (val === null || val === undefined || val === '') return null;
  if (val instanceof Date) {
    if (isNaN(val.getTime())) return null;
    var y = val.getFullYear(), m = val.getMonth() + 1, d = val.getDate();
    if (y < 1900 || y > 2030) return null;
    return y + '-' + (m < 10 ? '0' + m : m) + '-' + (d < 10 ? '0' + d : d);
  }
  if (typeof val === 'number' && val > 10000) {
    var excelEpoch = new Date(Date.UTC(1899, 11, 30));
    var dt = new Date(excelEpoch.getTime() + val * 86400000);
    var ye = dt.getUTCFullYear(), me = dt.getUTCMonth() + 1, de = dt.getUTCDate();
    if (ye < 1900 || ye > 2030) return null;
    return ye + '-' + (me < 10 ? '0' + me : me) + '-' + (de < 10 ? '0' + de : de);
  }
  if (typeof val === 'string') {
    var s = val.trim().replace(/[\s\-\/]+/g, '.');
    var m1 = s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);
    if (m1) {
      var dd = +m1[1], mm = +m1[2], yyyy = +m1[3];
      if (yyyy < 1900 || yyyy > 2030 || mm < 1 || mm > 12 || dd < 1 || dd > 31) return null;
      return yyyy + '-' + (mm < 10 ? '0' + mm : mm) + '-' + (dd < 10 ? '0' + dd : dd);
    }
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
    .replace(/\([^)]*\)/g, '')
    .replace(/\[[^\]]*\]/g, '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .replace(/[’ʼ′`'']/g, "'")
    .replace(/і/g, 'и')
    .replace(/ї/g, 'и')
    .replace(/є/g, 'е');
}

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

var BDAY_STATUS_SHEET = 'bday_sync_status';

function syncBdayStatusSheet() {
  var crmSS = getCRMSpreadsheet();
  var statusSheet = crmSS.getSheetByName(BDAY_STATUS_SHEET);
  if (!statusSheet) {
    statusSheet = crmSS.insertSheet(BDAY_STATUS_SHEET);
  }
  var HEADER = ['ChildID','Name','Loc','Bday','ContractNumber','Status','MatchedRegName','UpdatedAt','ConfirmedBy','ConfirmedAt'];

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
  }

  var paySheet = crmSS.getSheetByName(SHEET_PAYMENTS);
  if (!paySheet) return { ok: false, error: 'Оплати not found' };
  var payData = paySheet.getDataRange().getValues();
  if (payData.length < 2) return { ok: true, total: 0 };
  var payHdr = payData[0].map(function(h){ return String(h || ''); });
  var locI   = payHdr.indexOf('Локація');
  var grpI   = payHdr.indexOf('Група');
  var nameI  = payHdr.indexOf("Ім'я дитини");
  if (locI < 0 || grpI < 0 || nameI < 0) {
    return { ok: false, error: 'Оплати headers missing' };
  }

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

  statusSheet.clearContents();
  statusSheet.getRange(1, 1, 1, HEADER.length).setValues([HEADER]);
  statusSheet.setFrozenRows(1);
  statusSheet.getRange(2, 5, Math.max(rowsOut.length, 1), 1).setNumberFormat('@');
  if (rowsOut.length > 0) {
    statusSheet.getRange(2, 1, rowsOut.length, HEADER.length).setValues(rowsOut);
  }

  return {
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
}

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

function _opexIsSkippedCategory(name) {
  var normalized = String(name || '').trim().toLowerCase();
  if (!normalized) return true;
  var SKIP_NAMES = [
    'витрати',
    'підсумок',
    'знижки',
    'кількість дітей',
    'кількість груп',
    'кількість основного персоналу'
  ];
  if (SKIP_NAMES.indexOf(normalized) !== -1) return true;
  if (/^[\d\s.,]+$/.test(normalized)) return true;
  return false;
}

function _opexNormalizeCategoryName(name) {
  var s = String(name || '').trim();
  if (/^госп[\.\s]*товари/i.test(s)) return 'ХОЗ.ТОВАРИ';
  if (/^сніданки/i.test(s))          return 'СНІДАНКИ';
  return s;
}

function _opexNum(v) {
  if (typeof v === 'number' && isFinite(v)) return v;
  if (typeof v === 'string') {
    var n = parseFloat(v.replace(/\s+/g, '').replace(',', '.'));
    return isFinite(n) ? n : 0;
  }
  return 0;
}

function getOpexData(loc, year) {
  loc = String(loc || '').trim();
  if (!loc) return {ok:false, error:'Missing loc'};

  var configSS = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  var regSheet = configSS.getSheetByName('OPEX');
  if (!regSheet) return {ok:false, error:'OPEX registry tab not found in CONFIG'};

  var regData = regSheet.getDataRange().getValues();
  var sheetId = '', listName = 'OPEX';
  for (var i = 1; i < regData.length; i++) {
    if (String(regData[i][2] || '').trim() === loc) {
      sheetId  = String(regData[i][3] || '').trim();
      listName = String(regData[i][4] || '').trim() || 'OPEX';
      break;
    }
  }
  if (!sheetId) return {ok:false, error:'Location not found'};

  var locSS = SpreadsheetApp.openById(sheetId);
  var opex  = locSS.getSheetByName(listName);
  if (!opex) return {ok:false, error:'OPEX sheet not found in location file'};

  var lastRow = Math.max(opex.getLastRow(), 30);
  var lastCol = Math.max(opex.getLastColumn(), 37);
  var data = opex.getRange(1, 1, lastRow, lastCol).getValues();
  var width = lastCol;

  var categories = [];
  for (var rowNum = 3; rowNum <= 30; rowNum++) {
    var idx = rowNum - 1;
    if (idx >= data.length) break;
    var rowArr = data[idx] || [];
    var rawName = String(rowArr[0] || '').trim();
    if (_opexIsSkippedCategory(rawName)) continue;
    var name = _opexNormalizeCategoryName(rawName);

    var months = [];
    var totalFact = 0, totalBudget = 0;
    for (var m = 1; m <= 12; m++) {
      var fIdx = (m - 1) * 3 + 1;
      var bIdx = (m - 1) * 3 + 2;
      var fact   = fIdx < width ? _opexNum(rowArr[fIdx]) : 0;
      var budget = bIdx < width ? _opexNum(rowArr[bIdx]) : 0;
      months.push({month: m, fact: fact, budget: budget});
      totalFact   += fact;
      totalBudget += budget;
    }
    categories.push({
      name: name,
      row: rowNum,
      months: months,
      totalFact: totalFact,
      totalBudget: totalBudget
    });
  }

  return {
    ok: true,
    loc: loc,
    year: year ? Number(year) || year : '',
    categories: categories
  };
}

function getOpexOverview(year) {
  var configSS = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  var regSheet = configSS.getSheetByName('OPEX');
  if (!regSheet) return {ok:false, error:'OPEX registry tab not found in CONFIG'};

  var regData = regSheet.getDataRange().getValues();
  var locations = [];
  var errors = [];

  for (var i = 1; i < regData.length; i++) {
    var loc      = String(regData[i][2] || '').trim();
    var typ      = String(regData[i][1] || '').trim();
    var sheetId  = String(regData[i][3] || '').trim();
    var listName = String(regData[i][4] || '').trim() || 'OPEX';
    if (!loc || !sheetId) continue;

    try {
      var locSS = SpreadsheetApp.openById(sheetId);
      var opex  = locSS.getSheetByName(listName);
      if (!opex) {
        errors.push({loc: loc, error: 'OPEX sheet not found'});
        continue;
      }

      var lastRow = Math.max(opex.getLastRow(), 30);
      var lastCol = Math.max(opex.getLastColumn(), 37);
      var data    = opex.getRange(1, 1, lastRow, lastCol).getValues();
      var width   = lastCol;

      var catIdxs = [];
      for (var rowNum = 3; rowNum <= 30; rowNum++) {
        var idx = rowNum - 1;
        if (idx >= data.length) break;
        var rawName = String((data[idx] || [])[0] || '').trim();
        if (_opexIsSkippedCategory(rawName)) continue;
        catIdxs.push(idx);
      }

      var monthsTotals = [];
      var yearFact = 0, yearBudget = 0;

      for (var m = 1; m <= 12; m++) {
        var fIdx = (m - 1) * 3 + 1;
        var bIdx = (m - 1) * 3 + 2;
        var monthFact = 0, monthBudget = 0;

        for (var k = 0; k < catIdxs.length; k++) {
          var rowArr = data[catIdxs[k]] || [];
          monthFact   += fIdx < width ? _opexNum(rowArr[fIdx]) : 0;
          monthBudget += bIdx < width ? _opexNum(rowArr[bIdx]) : 0;
        }

        monthsTotals.push({month: m, fact: monthFact, budget: monthBudget});
        yearFact   += monthFact;
        yearBudget += monthBudget;
      }

      locations.push({
        loc: loc,
        type: typ,
        monthsTotals: monthsTotals,
        yearFact: yearFact,
        yearBudget: yearBudget
      });
    } catch (e) {
      errors.push({loc: loc, error: (e && e.message) ? e.message : String(e)});
    }
  }

  return {
    ok: true,
    year: year ? Number(year) || year : '',
    locations: locations,
    errors: errors
  };
}

var _OPEX_NORM_MAP = {
  'ОРЕНДА':                            'absolute',
  'КОМУНАЛЬНІ ПОСЛУГИ':                'absolute',
  'ПОВЕРНЕННЯ':                        'absolute',
  'Обслуговування':                    'absolute',
  'Маркетинг':                         'absolute',
  'ЗАХОДИ':                            'absolute',
  'КАП':                               'absolute',
  'СНІДАНКИ':                          'child',
  'КУХНЯ':                             'child',
  'Вода':                              'child',
  'ПРАЛЬНЯ':                           'child',
  'Підручники':                        'child',
  'КАНЦТОВАРИ':                        'group',
  'ІГРАШКИ':                           'group',
  'ХОЗ.ТОВАРИ':                        'group',
  'ТЕХНІКА':                           'group',
  'МЕБЛІ':                             'group',
  'РЕМОНТ':                            'group',
  'Постільне/коври/пуфи/форма':        'group',
  'НЕЗАПЛАНОВАНІ ВИТРАТИ':             'group',
  'Методична частина':                 'group',
  'Для персоналу':                     'staff',
  'ПОДАТОК':                           'staff',
  'Персонал':                          'staff',
  'Обслуговування ФОП':                'staff'
};

function _opexNormType(name) {
  return _OPEX_NORM_MAP[name] || 'unknown';
}

function _opexNormalize(value, normType, denoms) {
  if (normType === 'absolute' || normType === 'unknown') return value;
  var d;
  if      (normType === 'child') d = denoms.children;
  else if (normType === 'group') d = denoms.groups;
  else if (normType === 'staff') d = denoms.staff;
  else return value;
  if (!d || d <= 0) return null;
  return value / d;
}

function getCategoryAnalytics(year, month) {
  var m = parseInt(month, 10);
  if (!m || m < 1 || m > 12) return {ok:false, error:'Invalid month (must be 1-12)'};

  var configSS = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  var regSheet = configSS.getSheetByName('OPEX');
  if (!regSheet) return {ok:false, error:'OPEX registry tab not found in CONFIG'};

  var regData = regSheet.getDataRange().getValues();
  var fIdx    = (m - 1) * 3 + 1;
  var bIdx    = (m - 1) * 3 + 2;
  var metaIdx = 2 + (m - 1) * 3;

  var catMap = {};
  var errors = [];

  for (var i = 1; i < regData.length; i++) {
    var loc      = String(regData[i][2] || '').trim();
    var typ      = String(regData[i][1] || '').trim();
    var sheetId  = String(regData[i][3] || '').trim();
    var listName = String(regData[i][4] || '').trim() || 'OPEX';
    if (!loc || !sheetId) continue;

    try {
      var locSS = SpreadsheetApp.openById(sheetId);
      var opex  = locSS.getSheetByName(listName);
      if (!opex) {
        errors.push({loc: loc, error: 'OPEX sheet not found'});
        continue;
      }

      var lastRow = Math.max(opex.getLastRow(), 40);
      var lastCol = Math.max(opex.getLastColumn(), 37);
      var data    = opex.getRange(1, 1, lastRow, lastCol).getValues();
      var width   = lastCol;

      var children = (35 < data.length && metaIdx < width) ? _opexNum(data[35][metaIdx]) : 0;
      var groups   = (36 < data.length && metaIdx < width) ? _opexNum(data[36][metaIdx]) : 0;
      var staff    = (37 < data.length && metaIdx < width) ? _opexNum(data[37][metaIdx]) : 0;
      var denoms   = {children: children, groups: groups, staff: staff};

      for (var rowNum = 3; rowNum <= 30; rowNum++) {
        var idx = rowNum - 1;
        if (idx >= data.length) break;
        var rowArr  = data[idx] || [];
        var rawName = String(rowArr[0] || '').trim();
        if (_opexIsSkippedCategory(rawName)) continue;
        var name = _opexNormalizeCategoryName(rawName);

        var fact     = fIdx < width ? _opexNum(rowArr[fIdx]) : 0;
        var budget   = bIdx < width ? _opexNum(rowArr[bIdx]) : 0;
        var normType = _opexNormType(name);
        var normFact   = _opexNormalize(fact,   normType, denoms);
        var normBudget = _opexNormalize(budget, normType, denoms);

        if (!catMap[name]) {
          catMap[name] = {
            name: name,
            normType: normType,
            locations: [],
            totalFact: 0,
            totalBudget: 0
          };
        }
        var bucket = catMap[name];
        bucket.locations.push({
          loc:        loc,
          type:       typ,
          fact:       fact,
          budget:     budget,
          children:   children,
          groups:     groups,
          staff:      staff,
          normFact:   normFact,
          normBudget: normBudget
        });
        bucket.totalFact   += fact;
        bucket.totalBudget += budget;
      }
    } catch (e) {
      errors.push({loc: loc, error: (e && e.message) ? e.message : String(e)});
    }
  }

  var categories = Object.keys(catMap).map(function(name){
    var c = catMap[name];
    var sumF = 0, cntF = 0, sumB = 0, cntB = 0;
    c.locations.forEach(function(L){
      if (L.normFact !== null && L.normFact !== undefined && isFinite(L.normFact)){
        sumF += L.normFact; cntF++;
      }
      if (L.normBudget !== null && L.normBudget !== undefined && isFinite(L.normBudget)){
        sumB += L.normBudget; cntB++;
      }
    });
    c.avgNormFact   = cntF > 0 ? sumF / cntF : null;
    c.avgNormBudget = cntB > 0 ? sumB / cntB : null;
    return c;
  });

  return {
    ok: true,
    year:  year ? (Number(year) || year) : '',
    month: m,
    categories: categories,
    errors: errors
  };
}

function _salaryIsSkippedRow(name) {
  var s = String(name || '').trim();
  if (!s) return true;
  var lower = s.toLowerCase();
  if (lower === 'зарплата' || lower === 'персонал') return true;
  if (/^[\d\s.,]+$/.test(s)) return true;
  return false;
}

function _salaryIsSubtotalRow(name) {
  var lower = String(name || '').trim().toLowerCase();
  if (!lower) return false;
  if (lower.indexOf('додаткові заняття') !== -1) return true;
  if (lower.indexOf('день народження') !== -1) return true;
  return false;
}

function _salaryGetRegistry() {
  var configSS = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  var regSheet = configSS.getSheetByName('Salary');
  if (!regSheet) return {ok:false, error:'Salary registry tab not found in CONFIG'};

  var data = regSheet.getDataRange().getValues();
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var loc      = String(data[i][2] || '').trim();
    var typ      = String(data[i][1] || '').trim();
    var sheetId  = String(data[i][3] || '').trim();
    var listName = String(data[i][4] || '').trim() || 'Salary';
    if (!loc || !sheetId) continue;
    rows.push({
      typ:      typ,
      loc:      loc,
      sheetId:  sheetId,
      listName: listName
    });
  }
  return {ok:true, rows:rows};
}

function getSalaryData(loc, year) {
  loc = String(loc || '').trim();
  if (!loc) return {ok:false, error:'Missing loc'};

  var reg = _salaryGetRegistry();
  if (!reg.ok) return reg;

  var entry = null;
  for (var i = 0; i < reg.rows.length; i++) {
    if (reg.rows[i].loc === loc) { entry = reg.rows[i]; break; }
  }
  if (!entry) return {ok:false, error:'Location not found in Salary registry'};

  var locSS = SpreadsheetApp.openById(entry.sheetId);
  var sheet = locSS.getSheetByName(entry.listName);
  if (!sheet) return {ok:false, error:'Salary sheet not found in location file'};

  var lastRow = Math.max(sheet.getLastRow(), 80);
  var lastCol = Math.max(sheet.getLastColumn(), 37);
  var data    = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  var width   = lastCol;

  var rows = [];
  for (var rowNum = 4; rowNum <= data.length; rowNum++) {
    var idx = rowNum - 1;
    var rowArr = data[idx] || [];
    var rawName = String(rowArr[0] || '').trim();
    if (_salaryIsSkippedRow(rawName))  continue;
    if (_salaryIsSubtotalRow(rawName)) continue;

    var months = [];
    var totalFact = 0, totalBudget = 0;
    for (var m = 1; m <= 12; m++) {
      var fIdx = (m - 1) * 3 + 1;
      var bIdx = (m - 1) * 3 + 2;
      var fact   = fIdx < width ? _opexNum(rowArr[fIdx]) : 0;
      var budget = bIdx < width ? _opexNum(rowArr[bIdx]) : 0;
      months.push({month: m, fact: fact, budget: budget});
      totalFact   += fact;
      totalBudget += budget;
    }
    rows.push({
      row:         rowNum,
      name:        rawName,
      months:      months,
      totalFact:   totalFact,
      totalBudget: totalBudget
    });
  }

  return {
    ok:   true,
    loc:  loc,
    year: year ? Number(year) || year : '',
    rows: rows
  };
}

function getSalaryOverview(year) {
  var reg = _salaryGetRegistry();
  if (!reg.ok) return reg;

  var locations = [];
  var errors    = [];

  reg.rows.forEach(function(entry) {
    try {
      var locSS = SpreadsheetApp.openById(entry.sheetId);
      var sheet = locSS.getSheetByName(entry.listName);
      if (!sheet) {
        errors.push({loc: entry.loc, error: 'Salary sheet not found'});
        return;
      }

      var lastRow = Math.max(sheet.getLastRow(), 80);
      var lastCol = Math.max(sheet.getLastColumn(), 37);
      var data    = sheet.getRange(1, 1, lastRow, lastCol).getValues();
      var width   = lastCol;

      var rowIdxs = [];
      for (var rowNum = 4; rowNum <= data.length; rowNum++) {
        var idx = rowNum - 1;
        var rowArr = data[idx] || [];
        var rawName = String(rowArr[0] || '').trim();
        if (_salaryIsSkippedRow(rawName))  continue;
        if (_salaryIsSubtotalRow(rawName)) continue;
        rowIdxs.push(idx);
      }

      var monthsTotals = [];
      var yearFact = 0, yearBudget = 0;

      for (var m = 1; m <= 12; m++) {
        var fIdx = (m - 1) * 3 + 1;
        var bIdx = (m - 1) * 3 + 2;
        var monthFact = 0, monthBudget = 0;

        for (var k = 0; k < rowIdxs.length; k++) {
          var rowArr = data[rowIdxs[k]] || [];
          monthFact   += fIdx < width ? _opexNum(rowArr[fIdx]) : 0;
          monthBudget += bIdx < width ? _opexNum(rowArr[bIdx]) : 0;
        }

        monthsTotals.push({month: m, fact: monthFact, budget: monthBudget});
        yearFact   += monthFact;
        yearBudget += monthBudget;
      }

      locations.push({
        loc:          entry.loc,
        type:         entry.typ,
        monthsTotals: monthsTotals,
        yearFact:     yearFact,
        yearBudget:   yearBudget
      });
    } catch (e) {
      errors.push({loc: entry.loc, error: (e && e.message) ? e.message : String(e)});
    }
  });

  return {
    ok:        true,
    year:      year ? Number(year) || year : '',
    locations: locations,
    errors:    errors
  };
}

function _ovaIsSubjectTeacherName(name) {
  var s = String(name || '');
  var lower = s.toLowerCase();
  var SUBJECT_KEYWORDS = [
    'англійська','англійський',
    'логопед','муз.керівник','муз керівник',
    'хореограф','фітнес','психолог',
    'підготовка до школи','чомусики',
    'архітектура','смм','speaking','информатика'
  ];
  for (var i = 0; i < SUBJECT_KEYWORDS.length; i++){
    if (lower.indexOf(SUBJECT_KEYWORDS[i]) !== -1) return true;
  }
  if (/грн\s*\/?\s*(урок|занят)|за\s+занят/i.test(s)) return true;
  if (/\b1[0-4]\d{2}\b/.test(s) || /\b[1-9]\d{2}\b/.test(s)) return true;
  return false;
}

function _ovaClassifySalaryRow(name, currentSection) {
  var lower = String(name || '').toLowerCase();
  if (currentSection === 'extras')                                    return 'extras';
  if (lower.indexOf('директор')   !== -1)                             return 'director';
  if (lower.indexOf('вихователь') !== -1)                             return 'teacher';
  if (lower.indexOf('помічник')   !== -1)                             return 'assistant';
  if (lower.indexOf('медсестра')  !== -1)                             return 'nurse';
  if (lower.indexOf('охорон')     !== -1 || lower.indexOf('охран') !== -1) return 'guard';
  if (lower.indexOf('прибиральн') !== -1)                             return 'cleaner';
  if (lower.indexOf('тьютор')     !== -1 || lower.indexOf('тімлід') !== -1) return 'tutor';
  if (_ovaIsSubjectTeacherName(name))                                  return 'subject';
  return null;
}

function getOverviewAnalytics(year, month) {
  var m = parseInt(month, 10);
  if (!m || m < 1 || m > 12) return {ok:false, error:'Invalid month (1-12)'};

  var configSS = SpreadsheetApp.openById(CONFIG_SHEET_ID);

  var opexReg = configSS.getSheetByName('OPEX');
  if (!opexReg) return {ok:false, error:'OPEX registry tab not found in CONFIG'};
  var opexRegRows = opexReg.getDataRange().getValues();

  var salByLoc = {};
  var salReg = configSS.getSheetByName('Salary');
  if (salReg){
    var salRows = salReg.getDataRange().getValues();
    for (var si = 1; si < salRows.length; si++){
      var sLoc = String(salRows[si][2] || '').trim();
      var sId  = String(salRows[si][3] || '').trim();
      var sLst = String(salRows[si][4] || '').trim() || 'Salary';
      if (sLoc && sId) salByLoc[sLoc] = {sheetId: sId, listName: sLst};
    }
  }

  var allClients = [];
  try {
    var clRes = getClients();
    if (clRes && clRes.ok && clRes.data) allClients = clRes.data;
  } catch (e) {}

  var metaIdx = (m - 1) * 3 + 2;
  var fIdx    = (m - 1) * 3 + 1;
  var bIdx    = (m - 1) * 3 + 2;

  var locations = [];
  var errors = [];

  for (var i = 1; i < opexRegRows.length; i++) {
    var loc      = String(opexRegRows[i][2] || '').trim();
    var typ      = String(opexRegRows[i][1] || '').trim();
    var sheetId  = String(opexRegRows[i][3] || '').trim();
    var listName = String(opexRegRows[i][4] || '').trim() || 'OPEX';
    if (!loc || !sheetId) continue;
    if (typ === 'Управління') continue;

    var entry = {
      loc:                   loc,
      type:                  typ,
      childrenCount:         0,
      groupsCount:           0,
      mainStaffCount:        0,
      childrenByGroup:       {},
      childrenTotalFromList: 0,
      staffCounts: {
        director: 0, teacher: 0, assistant: 0, nurse: 0,
        guard: 0, cleaner: 0, tutor: 0, subject: 0, extras: 0
      },
      mainStaffFromSalary:   0,
      salaryFact:            0,
      salaryBudget:          0
    };

    try {
      var locOpexSS = SpreadsheetApp.openById(sheetId);
      var opexSh    = locOpexSS.getSheetByName(listName);
      if (!opexSh) {
        errors.push({loc: loc, source: 'opex', error: 'OPEX sheet not found'});
      } else {
        var opLastRow = Math.max(opexSh.getLastRow(), 40);
        var opLastCol = Math.max(opexSh.getLastColumn(), 37);
        var opData    = opexSh.getRange(1, 1, opLastRow, opLastCol).getValues();
        if (35 < opData.length && metaIdx < opLastCol) entry.childrenCount  = _opexNum(opData[35][metaIdx]);
        if (36 < opData.length && metaIdx < opLastCol) entry.groupsCount    = _opexNum(opData[36][metaIdx]);
        if (37 < opData.length && metaIdx < opLastCol) entry.mainStaffCount = _opexNum(opData[37][metaIdx]);
      }
    } catch (e) {
      errors.push({loc: loc, source: 'opex', error: (e && e.message) ? e.message : String(e)});
    }

    try {
      var byGroup = {};
      var totalActive = 0;
      for (var ci = 0; ci < allClients.length; ci++) {
        var c = allClients[ci];
        if (String(c['Локація'] || '').trim() !== loc) continue;
        var termRaw = c['Дата розірвання'];
        var hasTermDate = (termRaw instanceof Date) ||
                          (termRaw !== null && termRaw !== undefined && String(termRaw).trim() !== '');
        if (hasTermDate) continue;
        var statusLower = String(c['Статус'] || '').toLowerCase();
        if (statusLower.indexOf('розірв') !== -1) continue;

        var grp = String(c['Група'] || '').trim() || '(без групи)';
        byGroup[grp] = (byGroup[grp] || 0) + 1;
        totalActive++;
      }
      entry.childrenByGroup = byGroup;
      entry.childrenTotalFromList = totalActive;
    } catch (e) {
      errors.push({loc: loc, source: 'clients', error: (e && e.message) ? e.message : String(e)});
    }

    try {
      var salEntry = salByLoc[loc];
      if (!salEntry){
        errors.push({loc: loc, source: 'salary', error: 'Location not in Salary registry'});
      } else {
        var locSalSS = SpreadsheetApp.openById(salEntry.sheetId);
        var salSh    = locSalSS.getSheetByName(salEntry.listName);
        if (!salSh) {
          errors.push({loc: loc, source: 'salary', error: 'Salary sheet not found'});
        } else {
          var slastRow = Math.max(salSh.getLastRow(), 80);
          var slastCol = Math.max(salSh.getLastColumn(), 37);
          var salData  = salSh.getRange(1, 1, slastRow, slastCol).getValues();
          var salWidth = slastCol;

          var section = 'main';

          for (var rowNum = 4; rowNum <= salData.length; rowNum++) {
            var idx = rowNum - 1;
            var rowArr = salData[idx] || [];
            var rawName = String(rowArr[0] || '').trim();
            if (_salaryIsSkippedRow(rawName)) continue;

            if (_salaryIsSubtotalRow(rawName)) {
              if (/додаткові заняття/i.test(rawName)) section = 'extras';
              continue;
            }

            var category = _ovaClassifySalaryRow(rawName, section);
            if (category === 'subject' && section === 'main') section = 'subjects';
            if (category){
              entry.staffCounts[category] = (entry.staffCounts[category] || 0) + 1;
            }

            var fact   = fIdx < salWidth ? _opexNum(rowArr[fIdx]) : 0;
            var budget = bIdx < salWidth ? _opexNum(rowArr[bIdx]) : 0;
            entry.salaryFact   += fact;
            entry.salaryBudget += budget;
          }

          var sc = entry.staffCounts;
          entry.mainStaffFromSalary =
            sc.director + sc.teacher + sc.assistant + sc.nurse +
            sc.guard    + sc.cleaner + sc.tutor;
        }
      }
    } catch (e) {
      errors.push({loc: loc, source: 'salary', error: (e && e.message) ? e.message : String(e)});
    }

    locations.push(entry);
  }

  return {
    ok:        true,
    year:      year ? (Number(year) || year) : '',
    month:     m,
    locations: locations,
    errors:    errors
  };
}

var USERS_SHEET_NAME = 'Користувачі';
var VALID_USER_ROLES = ['cfo','ceo','cco','coo','rnd_director','hr_trainer','legal','cmo'];

function _getUsersSheet() {
  var ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  var sh = ss.getSheetByName(USERS_SHEET_NAME);
  if (!sh) throw new Error('Users sheet "' + USERS_SHEET_NAME + '" not found in CONFIG');
  return sh;
}

function _parseUserRow(row) {
  return {
    id:        Number(row[0]) || 0,
    name:      String(row[1] || '').trim(),
    login:     String(row[2] || '').trim(),
    password:  String(row[3] == null ? '' : row[3]),
    role:      String(row[4] || '').trim(),
    loc:       String(row[5] || '').trim(),
    email:     String(row[6] || '').trim(),
    active:    row[7] === true || /^(true|так|y|1|active|активний)$/i.test(String(row[7])),
    lastLogin: row[8]
                 ? (row[8] instanceof Date ? row[8].toISOString() : String(row[8]))
                 : ''
  };
}

function getUsers() {
  var sh = _getUsersSheet();
  var data = sh.getDataRange().getValues();
  var users = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    users.push(_parseUserRow(data[i]));
  }
  return {ok: true, users: users};
}

function authenticate(login, password) {
  login    = String(login    || '').trim();
  password = String(password == null ? '' : password);
  if (!login || !password) return {ok: false, error: 'Введіть логін і пароль'};

  var sh = _getUsersSheet();
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    var u = _parseUserRow(data[i]);
    if (u.login !== login) continue;
    if (!u.active) return {ok: false, error: 'Користувача деактивовано'};
    // v6.7: паролі — SHA-256. Plaintext-гілка лишена як страховка на
    // час міграції (після addAllDirectorsAndNurses() всі рядки хешовані).
    var stored = String(u.password == null ? '' : u.password);
    if (stored !== _sha256(password) && stored !== password)
      return {ok: false, error: 'Невірний пароль'};
    sh.getRange(i + 1, 9).setValue(new Date());
    delete u.password;
    return {ok: true, user: u};
  }
  return {ok: false, error: 'Користувача не знайдено'};
}

function updatePassword(userId, newPassword) {
  var id = Number(userId);
  newPassword = String(newPassword == null ? '' : newPassword);
  if (!id || !newPassword) return {ok: false, error: 'Missing userId or newPassword'};

  var sh = _getUsersSheet();
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (Number(data[i][0]) === id) {
      sh.getRange(i + 1, 4).setValue(newPassword);
      return {ok: true};
    }
  }
  return {ok: false, error: 'Користувача не знайдено'};
}

function addUser(payload) {
  if (!payload || !payload.name || !payload.login || !payload.password || !payload.role) {
    return {ok: false, error: 'Missing required fields (name, login, password, role)'};
  }
  if (VALID_USER_ROLES.indexOf(payload.role) === -1) {
    return {ok: false, error: 'Invalid role: ' + payload.role};
  }

  var sh = _getUsersSheet();
  var data = sh.getDataRange().getValues();

  var newLogin = String(payload.login).trim();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][2] || '').trim() === newLogin) {
      return {ok: false, error: 'Логін "' + newLogin + '" вже існує'};
    }
  }

  var maxId = 0;
  for (var i = 1; i < data.length; i++) {
    var n = Number(data[i][0]);
    if (n > maxId) maxId = n;
  }
  var newId = maxId + 1;

  sh.appendRow([
    newId,
    String(payload.name).trim(),
    newLogin,
    String(payload.password),
    payload.role,
    String(payload.loc || 'Менеджмент').trim(),
    String(payload.email || '').trim(),
    true,
    ''
  ]);
  return {ok: true, id: newId};
}

function deactivateUser(userId) { return _setUserActive(userId, false); }
function activateUser(userId)   { return _setUserActive(userId, true); }

function _setUserActive(userId, active) {
  var id = Number(userId);
  if (!id) return {ok: false, error: 'Missing userId'};

  var sh = _getUsersSheet();
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (Number(data[i][0]) === id) {
      sh.getRange(i + 1, 8).setValue(active);
      return {ok: true};
    }
  }
  return {ok: false, error: 'Користувача не знайдено'};
}

var GROUP_NORMS_SHEET_NAME = 'Норми груп';

function getGroupNorms() {
  try {
    var ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
    var sh = ss.getSheetByName(GROUP_NORMS_SHEET_NAME);
    if (!sh) return {ok: false, error: 'Sheet "' + GROUP_NORMS_SHEET_NAME + '" not found'};

    var lastRow = sh.getLastRow();
    if (lastRow < 2) return {ok: true, norms: {}, list: []};

    var rows = sh.getRange(2, 1, lastRow - 1, 2).getValues();
    var norms = {};
    var list = [];
    for (var i = 0; i < rows.length; i++) {
      var name = String(rows[i][0] || '').trim();
      var raw  = rows[i][1];
      if (!name) continue;
      var n = Number(raw);
      if (!isFinite(n) || n <= 0) continue;
      norms[name] = n;
      list.push({group: name, norm: n});
    }
    return {ok: true, norms: norms, list: list};
  } catch (e) {
    return {ok: false, error: String(e && e.message || e)};
  }
}

function syncPayments() {
  if (typeof aggregatePayments !== 'function') {
    return {
      ok: false,
      error: 'Функція синхронізації не знайдена в Apps Script.'
    };
  }
  try {
    var startedAt = new Date();
    var res = aggregatePayments() || {};
    var finishedAt = new Date();
    return {
      ok: !!res.ok,
      syncedAt:    finishedAt.toISOString(),
      durationSec: Math.round((finishedAt - startedAt) / 1000),
      rowsCount:   res.rows || 0,
      errors:      res.errors || [],
      month:       res.month || '',
      updated:     res.updated || ''
    };
  } catch (e) {
    return {ok: false, error: String(e && e.message || e)};
  }
}

var ACTIVITIES_SHEET_NAME = 'Додаткові_Каталог';
var ACTIVITIES_HEADER = [
  'id','Локація','Заняття','Ціна_клієнту','Модель_ЗП_викладача',
  'Ставка_викладача','Викладач','Активне'
];

function _getActivitiesSheet(createIfMissing){
  var ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  var sh = ss.getSheetByName(ACTIVITIES_SHEET_NAME);
  if (!sh && createIfMissing){
    sh = ss.insertSheet(ACTIVITIES_SHEET_NAME);
    sh.getRange(1, 1, 1, ACTIVITIES_HEADER.length).setValues([ACTIVITIES_HEADER]);
    sh.setFrozenRows(1);
  }
  if (!sh) throw new Error('Sheet "' + ACTIVITIES_SHEET_NAME + '" не знайдено. Створіть лист з колонками: ' + ACTIVITIES_HEADER.join(', '));
  return sh;
}

function _parseActivityRow(row){
  return {
    id:           Number(row[0]) || 0,
    loc:          String(row[1] || '').trim(),
    name:         String(row[2] || '').trim(),
    clientPrice:  Number(row[3]) || 0,
    teacherModel: String(row[4] || '').trim(),
    teacherRate:  Number(row[5]) || 0,
    teacher:      String(row[6] || '').trim(),
    active:       row[7] === true ||
                  /^(true|так|y|1|active|активне|✅)$/i.test(String(row[7] || '').trim())
  };
}

function getActivitiesCatalog(loc){
  try {
    var sh = _getActivitiesSheet(false);
    var data = sh.getDataRange().getValues();
    if (data.length < 2) return {ok: true, items: []};
    var items = [];
    var filterLoc = String(loc || '').trim();
    for (var i = 1; i < data.length; i++){
      var row = data[i];
      if (!row[2]) continue;
      var rec = _parseActivityRow(row);
      if (filterLoc && rec.loc !== filterLoc) continue;
      items.push(rec);
    }
    return {ok: true, items: items};
  } catch(e){
    return {ok: false, error: String(e && e.message || e)};
  }
}

// ───────────────────────────────────────────────────────────────────────────
// Seed каталогу ДОДАТКОВИХ занять — усі 11 локацій (Голосієво + 10 інших)
// у лист "Додаткові_Каталог".
//   seedActivitiesCatalog()       — м'який режим: якщо канонічні локації
//                                   вже мають заняття, НІЧОГО не чіпає
//                                   (захист ручних правок з UI).
//   seedActivitiesCatalogForce()  — force: перезаписує всі 11 локацій
//                                   канонічними цифрами. Рядки локацій поза
//                                   списком зберігаються незмінними.
// Запускати ВРУЧНУ з Apps Script editor.
// ───────────────────────────────────────────────────────────────────────────
function seedActivitiesCatalog(force){
  var sh = _getActivitiesSheet(true); // створить лист із шапкою якщо нема
  var data = sh.getDataRange().getValues();

  // Канонічні каталоги всіх 11 локацій. Порядок фіксований — Голосієво
  // перший, тож його id лишаються 1..10 (як було) і не ламають відмітки.
  // Запис: [Заняття, Ціна_клієнту, Модель_ЗП_викладача, Ставка_викладача]
  var CANON = [
    ['Голосієво', [
      ['Лего', 280, 'За дитину', 150],
      ['Арт', 280, 'За заняття', 400],
      ['m.Dance', 230, 'За заняття', 400],
      ['Логопед', 520, 'За дитину', 350],
      ['Гончарство', 280, 'За дитину', 150],
      ['Айкідо', 300, 'За заняття', 1000],
      ['Робототехніка', 420, 'За дитину', 250],
      ['Англійська групові', 280, 'За заняття', 280],
      ['Вокал індивід', 600, 'За дитину', 300],
      ['Вокал група', 300, 'За заняття', 500]
    ]],
    ['Бігова', [
      ['Лего', 260, 'За заняття', 500],
      ['Логопед', 450, 'За дитину', 250],
      ['Футбол', 270, 'За заняття', 600]
    ]],
    ['Борщагівка', [
      ['Лего', 280, 'За дитину', 150],
      ['Арт', 220, 'За заняття', 350],
      ['Театральна студія', 230, 'За заняття', 500],
      ['Вокал', 300, 'За заняття', 500],
      ['Робототехніка', 350, 'За дитину', 250],
      ['Індивідуальні з логопедом', 650, 'За дитину', 450],
      ['Гончарство', 350, 'За заняття', 600],
      ['М.Денс', 250, 'За заняття', 400]
    ]],
    ['Бровари', [
      ['Лего', 280, 'За дитину', 150],
      ['Арт Studio', 280, 'За заняття', 350],
      ['m.Dance', 230, 'За дитину', 80],
      ['Логопед', 500, 'За дитину', 250],
      ['Робототехніка', 430, 'За дитину', 250],
      ['Планетарій', 450, 'За дитину', 100],
      ['Карате', 180, 'За дитину', 85]
    ]],
    ["Кар'єрна", [
      ['Лего', 280, 'За дитину', 150],
      ['Арт', 230, 'За заняття', 400],
      ['m.Dance', 230, 'За заняття', 400],
      ['Логопед', 600, 'За дитину', 400],
      ['Айкідо', 280, 'За заняття', 550],
      ['Робототехніка', 350, 'За дитину', 250],
      ['Шахи', 230, 'За дитину', 115],
      ['Speaking Club', 500, 'За заняття', 1000]
    ]],
    ['Кругла', [
      ['Лего', 370, 'За дитину', 220],
      ['Логопед', 450, 'За дитину', 250],
      ['Футбол', 280, 'За заняття', 600],
      ['Чирлідинг', 200, 'За заняття', 400]
    ]],
    ['Оранж', [
      ['Арт', 230, 'За дитину', 120],
      ['Логопед', 500, 'За дитину', 250],
      ['Гончарство', 280, 'За заняття', 600],
      ['Робототехніка', 420, 'За дитину', 240],
      ['Англійська групові', 300, 'За дитину', 150],
      ['Карате', 330, 'За заняття', 500],
      ['Гімнастика', 380, 'За заняття', 350]
    ]],
    ['Осокорки', [
      ['Лего', 280, 'За заняття', 300],
      ['Арт', 280, 'За дитину', 125],
      ['m.Dance', 280, 'За заняття', 450],
      ['Логопед', 600, 'За дитину', 270],
      ['Гончарство', 300, 'За дитину', 125],
      ['Робототехніка', 420, 'За дитину', 250],
      ['Англійська групові', 550, 'За заняття', 500],
      ['Футбол', 350, 'За дитину', 150],
      ['Фортепіано', 500, 'За дитину', 250],
      ['Нейрофітнес', 280, 'За заняття', 300],
      ['Капоейро', 300, 'За дитину', 150]
    ]],
    ['Позняки', [
      ['Лего', 280, 'За дитину', 150],
      ['Арт', 280, 'За заняття', 450],
      ['m.Dance', 250, 'За заняття', 400],
      ['Логопед', 500, 'За дитину', 300],
      ['Гончарство', 300, 'За дитину', 140],
      ['Айкідо', 280, 'За заняття', 550],
      ['Робототехніка', 400, 'За дитину', 200],
      ['Англійська групові', 250, 'За заняття', 350]
    ]],
    ['Пуща', [
      ['Лего', 280, 'За дитину', 150],
      ['Арт', 300, 'За дитину', 150],
      ['Логопед', 500, 'За дитину', 250],
      ['Робототехніка', 450, 'За дитину', 250],
      ['Англійська групові', 450, 'За дитину', 250],
      ['Гімнастика', 280, 'За дитину', 150]
    ]],
    ['Тичини', [
      ['Лего', 280, 'За дитину', 150],
      ['Арт', 280, 'За заняття', 400],
      ['m.Dance', 310, 'За заняття', 350],
      ['Логопед', 520, 'За дитину', 350],
      ['Гончарство', 300, 'За дитину', 150],
      ['Айкідо', 280, 'За заняття', 550],
      ['Вокал індивід', 500, 'За дитину', 300],
      ['Англійська групові', 300, 'За заняття', 500]
    ]]
  ];

  // Множина канонічних локацій — для розподілу наявних рядків.
  var CANON_LOC = {};
  CANON.forEach(function(pair){ CANON_LOC[pair[0]] = true; });

  // Розділяємо наявні рядки: канонічні локації vs усі інші (зберігаємо як є).
  var canonExisting = 0, otherRows = [];
  for (var r = 1; r < data.length; r++){
    var row = data[r];
    if (!row[2]) continue; // нема назви заняття — порожній рядок
    var rowLoc = String(row[1] || '').trim();
    if (CANON_LOC[rowLoc]) canonExisting++;
    else                   otherRows.push(_normCatalogRow(row));
  }

  if (canonExisting > 0 && !force){
    Logger.log('[seedActivitiesCatalog] Канонічні локації вже мають %s занять. Запусти seedActivitiesCatalogForce() щоб перезаписати.', canonExisting);
    return {ok: true, skipped: true, existingRows: canonExisting};
  }

  // Канонічні рядки з наскрізними id (Голосієво перший → id 1..10).
  // Колонки: id | Локація | Заняття | Ціна_клієнту | Модель_ЗП_викладача |
  //          Ставка_викладача | Викладач | Активне
  var canonRows = [], id = 1;
  CANON.forEach(function(pair){
    var loc = pair[0];
    pair[1].forEach(function(a){
      canonRows.push([id++, loc, a[0], a[1], a[2], a[3], '', true]);
    });
  });

  var allRows = otherRows.concat(canonRows);

  // Очищаємо область даних і пишемо заново (шапка лишається у рядку 1).
  var lastRow = sh.getLastRow();
  if (lastRow > 1){
    sh.getRange(2, 1, lastRow - 1, ACTIVITIES_HEADER.length).clearContent();
  }
  if (allRows.length){
    sh.getRange(2, 1, allRows.length, ACTIVITIES_HEADER.length).setValues(allRows);
  }

  Logger.log('[seedActivitiesCatalog] Залито %s занять у %s локацій; інших рядків збережено: %s (force=%s)', canonRows.length, CANON.length, otherRows.length, !!force);
  return {ok: true, seeded: canonRows.length, locations: CANON.length, keptOtherRows: otherRows.length, force: !!force};
}

// Apps Script editor не дозволяє передавати аргументи у Run — окремий wrapper.
function seedActivitiesCatalogForce(){
  return seedActivitiesCatalog(true);
}

// Нормалізує рядок каталогу до рівно ACTIVITIES_HEADER.length колонок.
function _normCatalogRow(row){
  var out = [];
  for (var c = 0; c < ACTIVITIES_HEADER.length; c++){
    out.push(row[c] !== undefined ? row[c] : '');
  }
  return out;
}

function _nextActivityId(sh){
  var data = sh.getDataRange().getValues();
  var max = 0;
  for (var i = 1; i < data.length; i++){
    var n = Number(data[i][0]) || 0;
    if (n > max) max = n;
  }
  return max + 1;
}

function addActivity(data){
  try {
    var sh = _getActivitiesSheet(true);
    var id = _nextActivityId(sh);
    var row = [
      id,
      String(data.loc  || '').trim(),
      String(data.name || '').trim(),
      Number(data.clientPrice) || 0,
      String(data.teacherModel || '').trim(),
      Number(data.teacherRate) || 0,
      String(data.teacher || '').trim(),
      data.active !== false
    ];
    if (!row[1]) return {ok: false, error: 'Поле "Локація" обовʼязкове'};
    if (!row[2]) return {ok: false, error: 'Поле "Назва заняття" обовʼязкове'};
    sh.appendRow(row);
    return {ok: true, id: id};
  } catch(e){
    return {ok: false, error: String(e && e.message || e)};
  }
}

function updateActivity(id, data){
  try {
    var nid = Number(id);
    if (!nid) return {ok: false, error: 'Missing id'};
    var sh = _getActivitiesSheet(false);
    var rows = sh.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++){
      if (Number(rows[i][0]) !== nid) continue;
      var r1 = i + 1;
      if ('loc'          in data) sh.getRange(r1, 2).setValue(String(data.loc  || '').trim());
      if ('name'         in data) sh.getRange(r1, 3).setValue(String(data.name || '').trim());
      if ('clientPrice'  in data) sh.getRange(r1, 4).setValue(Number(data.clientPrice) || 0);
      if ('teacherModel' in data) sh.getRange(r1, 5).setValue(String(data.teacherModel || '').trim());
      if ('teacherRate'  in data) sh.getRange(r1, 6).setValue(Number(data.teacherRate) || 0);
      if ('teacher'      in data) sh.getRange(r1, 7).setValue(String(data.teacher || '').trim());
      if ('active'       in data) sh.getRange(r1, 8).setValue(data.active !== false);
      return {ok: true};
    }
    return {ok: false, error: 'Заняття не знайдено'};
  } catch(e){
    return {ok: false, error: String(e && e.message || e)};
  }
}

function deleteActivity(id){
  return updateActivity(id, {active: false});
}

function copyActivitiesFromLocation(fromLoc, toLoc){
  try {
    var fLoc = String(fromLoc || '').trim();
    var tLoc = String(toLoc   || '').trim();
    if (!fLoc) return {ok: false, error: 'Не вказано локацію-джерело'};
    if (!tLoc) return {ok: false, error: 'Не вказано локацію-приймач'};
    if (fLoc === tLoc) return {ok: false, error: 'Локація-джерело і приймач однакові'};

    var sh = _getActivitiesSheet(true);
    var data = sh.getDataRange().getValues();
    var src = [];
    for (var i = 1; i < data.length; i++){
      var rec = _parseActivityRow(data[i]);
      if (rec.loc === fLoc && rec.active) src.push(rec);
    }
    if (!src.length) return {ok: false, error: 'У локації "' + fLoc + '" немає активних занять'};

    var existsInTo = {};
    for (var j = 1; j < data.length; j++){
      var rec2 = _parseActivityRow(data[j]);
      if (rec2.loc === tLoc) existsInTo[rec2.name.toLowerCase()] = true;
    }

    var idCounter = _nextActivityId(sh);
    var toAppend = [];
    var skipped = 0;
    src.forEach(function(rec){
      if (existsInTo[rec.name.toLowerCase()]){ skipped++; return; }
      toAppend.push([
        idCounter++, tLoc, rec.name, rec.clientPrice,
        rec.teacherModel, 0,
        '',
        true
      ]);
    });

    if (toAppend.length){
      var startRow = sh.getLastRow() + 1;
      sh.getRange(startRow, 1, toAppend.length, ACTIVITIES_HEADER.length).setValues(toAppend);
    }

    return {
      ok: true,
      copied:  toAppend.length,
      skipped: skipped,
      total:   src.length
    };
  } catch(e){
    return {ok: false, error: String(e && e.message || e)};
  }
}

var ATTENDANCE_SHEET_NAME = 'Додаткові_Відвідуваність';
var ATTENDANCE_HEADER = [
  'id','Дата','Локація','Група','Дитина',
  'id_заняття','Назва_заняття','Ціна','Відмітив','Час_відмітки'
];

function _getAttendanceSheet(createIfMissing){
  var ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  var sh = ss.getSheetByName(ATTENDANCE_SHEET_NAME);
  if (!sh && createIfMissing){
    sh = ss.insertSheet(ATTENDANCE_SHEET_NAME);
    sh.getRange(1, 1, 1, ATTENDANCE_HEADER.length).setValues([ATTENDANCE_HEADER]);
    sh.setFrozenRows(1);
  }
  if (!sh) throw new Error('Sheet "' + ATTENDANCE_SHEET_NAME + '" не знайдено. Створіть лист з колонками: ' + ATTENDANCE_HEADER.join(', '));
  return sh;
}

function _parseAttendanceRow(row){
  var d = row[1];
  var dateStr;
  if (d instanceof Date){
    var y = d.getFullYear(), m = d.getMonth() + 1, dd = d.getDate();
    dateStr = y + '-' + (m < 10 ? '0' + m : m) + '-' + (dd < 10 ? '0' + dd : dd);
  } else {
    dateStr = String(d || '').trim();
  }
  return {
    id:           Number(row[0]) || 0,
    date:         dateStr,
    loc:          String(row[2] || '').trim(),
    group:        String(row[3] || '').trim(),
    child:        String(row[4] || '').trim(),
    activityId:   Number(row[5]) || 0,
    activityName: String(row[6] || '').trim(),
    price:        Number(row[7]) || 0,
    markedBy:     String(row[8] || '').trim(),
    markedAt:     row[9] instanceof Date ? row[9].toISOString() : String(row[9] || '')
  };
}

function getAttendanceMarks(filters){
  try {
    filters = filters || {};
    var sh = _getAttendanceSheet(false);
    var data = sh.getDataRange().getValues();
    if (data.length < 2) return {ok: true, items: []};
    var items = [];
    for (var i = 1; i < data.length; i++){
      if (!data[i][0] && !data[i][4]) continue;
      var m = _parseAttendanceRow(data[i]);
      if (filters.date  && m.date !== String(filters.date)) continue;
      if (filters.loc   && m.loc !== String(filters.loc)) continue;
      if (filters.group && m.group !== String(filters.group)) continue;
      if (filters.child && m.child !== String(filters.child)) continue;
      if (filters.activityId && m.activityId !== Number(filters.activityId)) continue;
      items.push(m);
    }
    return {ok: true, items: items};
  } catch(e){
    return {ok: false, error: String(e && e.message || e)};
  }
}

function _nextAttendanceId(sh){
  var data = sh.getDataRange().getValues();
  var max = 0;
  for (var i = 1; i < data.length; i++){
    var n = Number(data[i][0]) || 0;
    if (n > max) max = n;
  }
  return max + 1;
}

function addAttendanceMark(data){
  try {
    var sh = _getAttendanceSheet(true);
    var id = _nextAttendanceId(sh);
    var row = [
      id,
      String(data.date  || '').trim(),
      String(data.loc   || '').trim(),
      String(data.group || '').trim(),
      String(data.child || '').trim(),
      Number(data.activityId) || 0,
      String(data.activityName || '').trim(),
      Number(data.price) || 0,
      String(data.markedBy || '').trim(),
      new Date()
    ];
    if (!row[1] || !row[4] || !row[5]){
      return {ok: false, error: 'Поля Дата / Дитина / id_заняття обовʼязкові'};
    }
    sh.appendRow(row);
    return {ok: true, id: id};
  } catch(e){
    return {ok: false, error: String(e && e.message || e)};
  }
}

function removeAttendanceMark(id){
  try {
    var nid = Number(id);
    if (!nid) return {ok: false, error: 'Missing id'};
    var sh = _getAttendanceSheet(false);
    var data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++){
      if (Number(data[i][0]) === nid){
        sh.deleteRow(i + 1);
        return {ok: true};
      }
    }
    return {ok: false, error: 'Відмітку не знайдено'};
  } catch(e){
    return {ok: false, error: String(e && e.message || e)};
  }
}

var MONTHS_CAL_UA = [
  'Січень','Лютий','Березень','Квітень','Травень','Червень',
  'Липень','Серпень','Вересень','Жовтень','Листопад','Грудень'
];


// ═══════════════════════════════════════════════════════════════════════════
// AUTO-EXPORT v6.4 — пише у файл локації (Голосієво Payment, Осокорки Payment...)
// 5 колонок на місяць: Факт навч | Факт вступ | Факт доп | Бюджет доп | Бюджет навч
// Відмітка за місяць N → накопичена сума → Бюджет доп місяця N+1
// (бо діти ходять у травні, а оплата виставляється у червні)
// ═══════════════════════════════════════════════════════════════════════════

// Допоміжна — повертає (наступний_місяць, рік) для (місяць, рік)
function _nextMonth(month, year){
  var nm = month + 1, ny = year;
  if (nm > 12){ nm = 1; ny++; }
  return {month: nm, year: ny};
}

// Знаходить sheetId+listName файлу локації з CONFIG-реєстру локацій
function _getLocationPaymentRegistry(loc){
  var configSS    = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  var configSheet = configSS.getSheets()[0]; // перший лист = реєстр локацій
  var data = configSheet.getDataRange().getValues();
  for (var r = 1; r < data.length; r++){
    if (trim(data[r][2]) === loc){
      return {
        sheetId:   trim(data[r][3]),
        sheetName: trim(data[r][4]) || 'Payment'
      };
    }
  }
  return null;
}

// ───────────────────────────────────────────────────────────────────────────
// EXPORT JOURNAL — лист "Експорт_Журнал" у CONFIG-таблиці.
// Зберігає що ми попередньо записали по (loc, kind, name, targetYear, targetMonth).
// Дозволяє підрахувати baseValue = currentCell - lastWritten (тобто ручні
// поправки фінансиста: борг, переплата) і писати baseValue + newSum.
// kind: 'payment' (Бюджет доп у Payment) | 'salary' (Budget викладача у Salary).
// ───────────────────────────────────────────────────────────────────────────
var EXPORT_JOURNAL_SHEET = 'Експорт_Журнал';

function _getExportJournalSheet(){
  var configSS = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  var sh = configSS.getSheetByName(EXPORT_JOURNAL_SHEET);
  if (!sh){
    sh = configSS.insertSheet(EXPORT_JOURNAL_SHEET);
    sh.getRange(1, 1, 1, 7).setValues([['loc','kind','name','year','month','last_written_sum','last_written_at']]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function _journalNormName(s){
  return String(s || '').replace(/[\s\u00A0]+/g, '').toLowerCase();
}

// Зчитує усі записи журналу для (loc, kind, year, month). Ключ — норм-ім'я.
function _readJournalForTarget(loc, kind, year, month){
  var sh = _getExportJournalSheet();
  var data = sh.getDataRange().getValues();
  var byNormName = {};
  for (var r = 1; r < data.length; r++){
    var row = data[r];
    if (trim(String(row[0])) !== loc) continue;
    if (trim(String(row[1])) !== kind) continue;
    if (Number(row[3]) !== year) continue;
    if (Number(row[4]) !== month) continue;
    var nk = _journalNormName(row[2]);
    byNormName[nk] = {
      sum: Number(row[5]) || 0,
      at:  row[6],
      row: r + 1,            // 1-based — для adresування у setValues
      nameInJournal: String(row[2] || '')
    };
  }
  return {sheet: sh, byNormName: byNormName};
}

// Batch-апсерт записів у журналі. ops: [{nk, loc, kind, name, year, month, newSum}].
function _commitJournalUpdates(journal, ops){
  if (!ops || !ops.length) return;
  var now = new Date();
  var appends = [];
  ops.forEach(function(op){
    if (journal.byNormName.hasOwnProperty(op.nk)){
      var entry = journal.byNormName[op.nk];
      // last_written_sum (col 6) + last_written_at (col 7)
      journal.sheet.getRange(entry.row, 6, 1, 2).setValues([[op.newSum, now]]);
      entry.sum = op.newSum;
      entry.at  = now;
    } else {
      appends.push([op.loc, op.kind, op.name, op.year, op.month, op.newSum, now]);
    }
  });
  if (appends.length){
    var startRow = journal.sheet.getLastRow() + 1;
    journal.sheet.getRange(startRow, 1, appends.length, 7).setValues(appends);
    // Оновлюємо in-memory мапу, щоб подальші виклики бачили нові рядки.
    appends.forEach(function(a, i){
      journal.byNormName[_journalNormName(a[2])] = {
        sum: a[5], at: a[6], row: startRow + i, nameInJournal: a[2]
      };
    });
  }
}

function exportAttendanceToPayments(params){
  try {
    var loc = String(params.loc || '').trim();
    var month = Number(params.month);
    var year = Number(params.year) || new Date().getFullYear();
    if (!loc) return {ok: false, error: 'Параметр loc обовʼязковий'};
    if (!month || month < 1 || month > 12) return {ok: false, error: 'month має бути 1-12'};

    Logger.log('[exportAttendanceToPayments] START loc="%s" month=%s year=%s', loc, month, year);

    var monthName = MONTHS_CAL_UA[month - 1];

    // === 1. Підраховуємо суму додаткових для кожної дитини за вказаний місяць ===
    var attSh = _getAttendanceSheet(false);
    var attData = attSh.getDataRange().getValues();
    var mm = month < 10 ? '0' + month : String(month);
    var dateFrom = year + '-' + mm + '-01';
    var nextM = _nextMonth(month, year);
    var nmm = nextM.month < 10 ? '0' + nextM.month : String(nextM.month);
    var dateTo = nextM.year + '-' + nmm + '-01';
    Logger.log('[exportAttendanceToPayments] фільтр дат: [%s .. %s)  attData.length=%s', dateFrom, dateTo, attData.length);

    var sumPerChild = {};
    var passedRecords = 0;
    var skippedByLoc = 0;
    var skippedByDate = 0;
    for (var i = 1; i < attData.length; i++){
      var rec = _parseAttendanceRow(attData[i]);
      if (rec.loc !== loc){ skippedByLoc++; continue; }
      if (rec.date < dateFrom || rec.date >= dateTo){
        skippedByDate++;
        Logger.log('[exportAttendanceToPayments] skip-by-date: child="%s" date="%s" (поза [%s..%s))', rec.child, rec.date, dateFrom, dateTo);
        continue;
      }
      passedRecords++;
      sumPerChild[rec.child] = (sumPerChild[rec.child] || 0) + (rec.price || 0);
    }
    Logger.log('[exportAttendanceToPayments] attendance: passed=%s, skippedByLoc=%s, skippedByDate=%s', passedRecords, skippedByLoc, skippedByDate);
    Logger.log('[exportAttendanceToPayments] sumPerChild keys = %s', JSON.stringify(Object.keys(sumPerChild)));
    Logger.log('[exportAttendanceToPayments] sumPerChild = %s', JSON.stringify(sumPerChild));

    // === 2. Відкриваємо файл локації через CONFIG ===
    var reg = _getLocationPaymentRegistry(loc);
    if (!reg || !reg.sheetId){
      Logger.log('[exportAttendanceToPayments] ERROR: локацію "%s" не знайдено в CONFIG', loc);
      return {ok: false, error: 'Локацію "' + loc + '" не знайдено в CONFIG-реєстрі'};
    }
    Logger.log('[exportAttendanceToPayments] registry: sheetId="%s" sheetName="%s"', reg.sheetId, reg.sheetName);
    var paymentSS = SpreadsheetApp.openById(reg.sheetId);
    var paySh = paymentSS.getSheetByName(reg.sheetName) || paymentSS.getSheets()[0];
    if (!paySh){
      Logger.log('[exportAttendanceToPayments] ERROR: лист "%s" не знайдено', reg.sheetName);
      return {ok: false, error: 'Лист "' + reg.sheetName + '" не знайдено у файлі локації'};
    }

    var data = paySh.getDataRange().getValues();
    Logger.log('[exportAttendanceToPayments] Payment-лист "%s" відкрито, rows=%s', paySh.getName(), data.length);

    // === 3. Обчислюємо колонку місяця-ТАРГЕТУ (місяць + 1) ===
    var targetMonthIdx = nextM.month - 1;
    var monthStartCol0 = 1 + targetMonthIdx * 5;
    var budgetDopColIdx = monthStartCol0 + 3;
    var budgetDopCol1 = budgetDopColIdx + 1;
    var targetMonthName = MONTHS_CAL_UA[targetMonthIdx];
    Logger.log('[exportAttendanceToPayments] target month: idx=%s (%s), Бюджет доп col=%s (1-based)', targetMonthIdx, targetMonthName, budgetDopCol1);

    // === 4. Толерантний матч імен (lowercase + видалити ВСІ whitespace — пробіли, NBSP, табуляції) ===
    // "Волков Матвій" / "волков  матвій" / "Волков\u00A0Матвій" / "ВОЛКОВ\tМАТВІЙ" → "волковматвій"
    function _normName(s){
      return String(s || '').replace(/[\s\u00A0]+/g, '').toLowerCase();
    }
    var DATA_START = 3;
    var matchedRows = {};   // {childName_з_attendance: rowIdx (0-based) у Payment}
    var paymentByNorm = {}; // {normName: rowIdx (0-based)} — індекс по Payment-листу
    var paymentNames = [];  // діагностика: усі імена, що ми побачили в Payment
    for (var r = DATA_START; r < data.length; r++){
      var nameCell = trim(String(data[r][0] || ''));
      if (!nameCell) continue;
      if (isGroupHeaderRow(data[r], 1)) continue;
      paymentNames.push(nameCell);
      var nk = _normName(nameCell);
      // ПЕРШЕ співпадіння імені (якщо тезки в різних групах — пише в першу)
      if (!paymentByNorm.hasOwnProperty(nk)) paymentByNorm[nk] = r;
    }
    Logger.log('[exportAttendanceToPayments] paymentNames.length=%s, перші 10: %s', paymentNames.length, JSON.stringify(paymentNames.slice(0, 10)));

    Object.keys(sumPerChild).forEach(function(childName){
      var nk = _normName(childName);
      if (paymentByNorm.hasOwnProperty(nk)){
        matchedRows[childName] = paymentByNorm[nk];
        Logger.log('[exportAttendanceToPayments] MATCH: "%s" (norm="%s") → row %s', childName, nk, paymentByNorm[nk] + 1);
      } else {
        Logger.log('[exportAttendanceToPayments] NO-MATCH: "%s" (norm="%s") — серед %s імен Payment-листа збігу немає', childName, nk, paymentNames.length);
      }
    });

    // === 5. РОЗУМНЕ ПЕРЕЗАПИСУВАННЯ через журнал ===
    // Для кожного child-рядка:
    //   currentValue = поточне значення клітинки
    //   lastWritten  = скільки ми поклали туди минулого разу (з журналу)
    //   baseValue    = currentValue - lastWritten   (ручні поправки фінансиста)
    //   newSum       = сума з sumPerChild або 0 якщо галочки зняті
    //   newCell      = baseValue + newSum
    // Журнал → нова newSum (lastWritten для наступного запуску).
    //
    // ⚠️ ПЕРШИЙ запуск після впровадження журналу: lastWritten=0, baseValue
    //    дорівнює поточному значенню → подвоєння попередніх код-записаних сум.
    //    Якщо колонка не порожня — обнули її руками одноразово перед запуском.
    var lastSheetRow = paySh.getLastRow();
    var colValues   = paySh.getRange(1, budgetDopCol1, lastSheetRow, 1).getValues();
    var colFormulas = paySh.getRange(1, budgetDopCol1, lastSheetRow, 1).getFormulas();
    var journal = _readJournalForTarget(loc, 'payment', nextM.year, nextM.month);
    Logger.log('[exportAttendanceToPayments] journal: %s записів для (%s, payment, %s/%s)', Object.keys(journal.byNormName).length, loc, nextM.year, nextM.month);

    var sumByNorm = {};
    Object.keys(sumPerChild).forEach(function(childName){
      sumByNorm[_normName(childName)] = {name: childName, sum: sumPerChild[childName]};
    });

    var updated = 0;
    var totalAmount = 0;
    var details = [];
    var matchedChildren = {};
    var journalOps = [];
    var cellsWritten = 0;
    var formulaRowsSkipped = 0;

    // ⚠️ ТОЧКОВИЙ запис: НЕ використовуємо setValues на весь стовпець —
    // це затирало б формули у підсумкових рядках (=SUM(AI24:AI39) → Cashflow).
    // Пишемо setValue() лише у child-рядки, і лише якщо значення змінилось.
    // Рядки з формулами не чіпаємо за жодних обставин.
    Object.keys(paymentByNorm).forEach(function(nk){
      var rowIdx0 = paymentByNorm[nk];
      if (colFormulas[rowIdx0] && colFormulas[rowIdx0][0]){
        formulaRowsSkipped++;
        Logger.log('[exportAttendanceToPayments] skipped formula row %s: %s', rowIdx0 + 1, colFormulas[rowIdx0][0]);
        return;
      }
      var paymentName  = trim(String(data[rowIdx0][0] || ''));
      var currentValue = Number(colValues[rowIdx0][0]) || 0;
      var je           = journal.byNormName[nk];
      var lastWritten  = je ? je.sum : 0;
      var baseValue    = currentValue - lastWritten;
      var match        = sumByNorm[nk];
      var newSum       = match ? match.sum : 0;
      var newValue     = baseValue + newSum;

      // Точковий запис лише змінених клітинок.
      if (newValue !== currentValue){
        paySh.getRange(rowIdx0 + 1, budgetDopCol1).setValue(newValue);
        cellsWritten++;
      }

      if (newSum !== lastWritten){
        journalOps.push({
          nk: nk, loc: loc, kind: 'payment', name: paymentName,
          year: nextM.year, month: nextM.month, newSum: newSum
        });
      }

      if (match){
        matchedRows[match.name] = rowIdx0;
        matchedChildren[match.name] = true;
        updated++;
        totalAmount += newSum;
        details.push({
          child: match.name, sum: newSum,
          currentBefore: currentValue, lastWritten: lastWritten,
          baseValue: baseValue, newCell: newValue,
          row: rowIdx0 + 1, status: 'updated'
        });
        Logger.log('[exportAttendanceToPayments] WRITE row=%s "%s" cur=%s last=%s base=%s newSum=%s → %s', rowIdx0 + 1, paymentName, currentValue, lastWritten, baseValue, newSum, newValue);
      } else if (lastWritten !== 0){
        details.push({
          child: paymentName, sum: 0,
          currentBefore: currentValue, lastWritten: lastWritten,
          baseValue: baseValue, newCell: newValue,
          row: rowIdx0 + 1, status: 'cleared'
        });
        Logger.log('[exportAttendanceToPayments] CLEAR row=%s "%s" cur=%s last=%s base=%s → %s (відмітки зняті)', rowIdx0 + 1, paymentName, currentValue, lastWritten, baseValue, newValue);
      }
    });

    Logger.log('[exportAttendanceToPayments] точковий запис: %s клітинок змінено, %s формульних рядків пропущено', cellsWritten, formulaRowsSkipped);

    _commitJournalUpdates(journal, journalOps);
    Logger.log('[exportAttendanceToPayments] journal upsert: %s op(s)', journalOps.length);

    var notFound = Object.keys(sumPerChild).filter(function(n){ return !matchedChildren[n]; });
    notFound.forEach(function(n){
      details.push({child: n, sum: sumPerChild[n], status: 'not-found-in-payment'});
    });
    Logger.log('[exportAttendanceToPayments] DONE: updated=%s, totalAmount=%s, notFound=%s', updated, totalAmount, JSON.stringify(notFound));

    return {
      ok: true,
      updated: updated,
      totalAmount: totalAmount,
      notFound: notFound,
      loc: loc,
      sourceMonth: monthName,        // травень (місяць відвідувань)
      targetMonth: targetMonthName,  // червень (куди записали бюджет)
      targetCol: budgetDopCol1,
      // Діагностика — щоб з фронту видно було, що саме знайшли:
      attendanceKeys: Object.keys(sumPerChild),
      paymentNamesCount: paymentNames.length,
      paymentNamesSample: paymentNames.slice(0, 50),
      details: details
    };
  } catch(e){
    Logger.log('[exportAttendanceToPayments] EXCEPTION: %s\n%s', e && e.message, e && e.stack);
    return {ok: false, error: String(e && e.message || e)};
  }
}

function exportToPayments(params){ return exportAttendanceToPayments(params); }

// ───────────────────────────────────────────────────────────────────────────
// Тест: запусти вручну з Apps Script editor (View → Executions → дивись Логи).
// Викликає exportAttendanceToPayments({loc:"Голосієво", month:6, year:2026})
// і логує JSON-результат + ключові поля окремо.
// ───────────────────────────────────────────────────────────────────────────
function testExportVolkov(){
  var result = exportAttendanceToPayments({loc: 'Голосієво', month: 6, year: 2026});
  Logger.log('[testExportVolkov] result JSON: %s', JSON.stringify(result, null, 2));
  Logger.log('[testExportVolkov] error             = %s', result && result.error);
  Logger.log('[testExportVolkov] updated           = %s', result && result.updated);
  Logger.log('[testExportVolkov] totalAmount       = %s', result && result.totalAmount);
  Logger.log('[testExportVolkov] attendanceKeys    = %s', JSON.stringify(result && result.attendanceKeys));
  Logger.log('[testExportVolkov] paymentNamesCount = %s', result && result.paymentNamesCount);
  Logger.log('[testExportVolkov] paymentNamesSample= %s', JSON.stringify(result && result.paymentNamesSample));
  Logger.log('[testExportVolkov] notFound          = %s', JSON.stringify(result && result.notFound));
  Logger.log('[testExportVolkov] details           = %s', JSON.stringify(result && result.details));
  return result;
}

// ═══════════════════════════════════════════════════════════════════════════
// SALARY EXTRAS v6.4 — пише у файл локації (Голосієво Salary тощо)
// Структура Salary: A=ПІБ, потім по 3 колонки на місяць: Fact | Budget | ?
// Викладач якого місяця працював → Budget колонка наступного місяця
// (зарплату нараховують у наступному місяці після того як викладач відпрацював)
// ═══════════════════════════════════════════════════════════════════════════
function exportToSalaryExtras(params){
  try {
    var loc = String(params.loc || '').trim();
    var month = Number(params.month);
    var year = Number(params.year) || new Date().getFullYear();
    if (!loc) return {ok: false, error: 'Параметр loc обовʼязковий'};
    if (!month || month < 1 || month > 12) return {ok: false, error: 'month має бути 1-12'};

    var monthName = MONTHS_CAL_UA[month - 1];

    // === 1. Каталог занять для локації ===
    var catRes = getActivitiesCatalog(loc);
    if (!catRes.ok) return catRes;
    var allActive = (catRes.items || []).filter(function(a){ return a.active; });
    var withRate = allActive.filter(function(a){ return a.teacherRate > 0 && a.teacherModel; });
    var skipped = allActive.filter(function(a){ return !(a.teacherRate > 0 && a.teacherModel); })
                           .map(function(a){ return a.name; });

    // === 2. Підрахунок ЗП кожного викладача за місяць ===
    var attSh = _getAttendanceSheet(false);
    var attData = attSh.getDataRange().getValues();
    var mm = month < 10 ? '0' + month : String(month);
    var dateFrom = year + '-' + mm + '-01';
    var nextM = _nextMonth(month, year);
    var nmm = nextM.month < 10 ? '0' + nextM.month : String(nextM.month);
    var dateTo = nextM.year + '-' + nmm + '-01';

    Logger.log('[exportToSalaryExtras] START loc="%s" month=%s year=%s; фільтр [%s..%s)', loc, month, year, dateFrom, dateTo);
    Logger.log('[exportToSalaryExtras] каталог: allActive=%s, withRate=%s, skipped=%s', allActive.length, withRate.length, JSON.stringify(skipped));

    var byActId = {};
    for (var i = 1; i < attData.length; i++){
      var rec = _parseAttendanceRow(attData[i]);
      if (rec.loc !== loc) continue;
      if (rec.date < dateFrom || rec.date >= dateTo) continue;
      if (!byActId[rec.activityId]) byActId[rec.activityId] = {count: 0, dates: {}};
      byActId[rec.activityId].count++;
      byActId[rec.activityId].dates[rec.date] = true;
    }

    // Резолвимо activityId → назва (для діагностики) з каталогу.
    var idToName = {};
    allActive.forEach(function(a){ idToName[a.id] = a.name; });
    var attendanceKeys = Object.keys(byActId).map(function(id){
      return (idToName[id] || ('id=' + id)) + ' ×' + byActId[id].count + ' (днів ' + Object.keys(byActId[id].dates).length + ')';
    });
    Logger.log('[exportToSalaryExtras] byActId (заняття з відмітками): %s', JSON.stringify(attendanceKeys));

    var factByName = {};
    withRate.forEach(function(a){
      var stat = byActId[a.id] || {count: 0, dates: {}};
      var fact = 0;
      if (a.teacherModel === 'За дитину'){
        fact = stat.count * a.teacherRate;
      } else if (a.teacherModel === 'За заняття'){
        fact = Object.keys(stat.dates).length * a.teacherRate;
      }
      // Ключ — нормалізована назва (lowercase + без whitespace), як у Payment.
      factByName[_journalNormName(a.name)] = {fact: fact, name: a.name, hasMarks: stat.count > 0};
    });

    var factByActivity = {};
    Object.keys(factByName).forEach(function(ln){
      factByActivity[factByName[ln].name] = factByName[ln].fact;
    });
    Logger.log('[exportToSalaryExtras] factByActivity (ЗП по заняттях): %s', JSON.stringify(factByActivity));

    // === 3. Відкриваємо Salary файл локації ===
    var reg = _salaryGetRegistry();
    if (!reg.ok) return reg;
    var entry = null;
    for (var j = 0; j < reg.rows.length; j++){
      if (reg.rows[j].loc === loc){ entry = reg.rows[j]; break; }
    }
    if (!entry) return {ok: false, error: 'Локація "' + loc + '" не знайдена у Salary-реєстрі'};

    var locSS = SpreadsheetApp.openById(entry.sheetId);
    var sheet = locSS.getSheetByName(entry.listName);
    if (!sheet) return {ok: false, error: 'Salary sheet "' + entry.listName + '" не знайдено'};

    var lastRow = Math.max(sheet.getLastRow(), 80);
    var names = sheet.getRange(1, 1, lastRow, 1).getValues();
    // Salary: 3 колонки на місяць (A=name, B=Fact_січень, C=Budget_січень, ...)
    // Викладач відпрацював у місяці N → ЗП у Budget місяця N+1
    var targetMonth = nextM.month; // 1-12
    var budgetCol = (targetMonth - 1) * 3 + 3; // 1-based: для січня = 3, для лютого = 6 ...
    var targetMonthName = MONTHS_CAL_UA[targetMonth - 1];
    Logger.log('[exportToSalaryExtras] Salary файл "%s" → лист "%s", lastRow=%s', entry.sheetId, entry.listName, lastRow);
    Logger.log('[exportToSalaryExtras] targetMonth=%s (%s), budgetCol=%s (1-based)', targetMonth, targetMonthName, budgetCol);

    // Діагностика: перші 20 рядків колонки A Salary-листа (з номерами рядків).
    var salaryRowNames = [];
    for (var sn = 0; sn < names.length && sn < 20; sn++){
      salaryRowNames.push('row' + (sn + 1) + ': "' + String(names[sn][0] || '') + '"');
    }
    Logger.log('[exportToSalaryExtras] salaryRowNames (перші 20): %s', JSON.stringify(salaryRowNames));

    var updated = 0, totalFact = 0;
    var notFound = [];
    var details = [];

    // Карта нормалізована-назва → 1-based row у Salary-листі (прохід по A-колонці).
    // Нормалізація толерантна: "m. Dance" і "m.Dance" → "m.dance".
    var actRowByLname = {};
    for (var k = 3; k < names.length; k++){
      var rname = _journalNormName(names[k][0]);
      if (rname && !actRowByLname.hasOwnProperty(rname)){
        actRowByLname[rname] = k + 1;
      }
    }
    Logger.log('[exportToSalaryExtras] actRowByLname (рядки занять, від row4): %s', JSON.stringify(actRowByLname));

    // === РОЗУМНЕ ПЕРЕЗАПИСУВАННЯ через журнал (kind=salary) ===
    // baseValue = currentCell - lastWritten; newCell = baseValue + newFact.
    // Дивись коментар у exportAttendanceToPayments — та сама логіка.
    var budgetColValues   = sheet.getRange(1, budgetCol, lastRow, 1).getValues();
    var budgetColFormulas = sheet.getRange(1, budgetCol, lastRow, 1).getFormulas();
    var journal = _readJournalForTarget(loc, 'salary', nextM.year, nextM.month);
    Logger.log('[exportToSalaryExtras] journal: %s записів для (%s, salary, %s/%s)', Object.keys(journal.byNormName).length, loc, nextM.year, nextM.month);

    var journalOps = [];
    var cellsWritten = 0;
    var formulaRowsSkipped = 0;

    // ⚠️ ТОЧКОВИЙ запис: НЕ setValues на весь стовпець (затирало б формули
    // у підсумкових рядках). Пишемо setValue() лише у рядки занять, і лише
    // якщо значення змінилось. Формульні рядки не чіпаємо за жодних обставин.
    //
    // Ідемо по всіх АКТИВНИХ заняттях каталогу (а не лише по withRate) —
    // інакше якщо викладача прибрали з активних, попередня сума не очиститься.
    allActive.forEach(function(a){
      var lname = _journalNormName(a.name);
      var rowFound = actRowByLname[lname] || -1;
      if (rowFound <= 0){
        // Активність є у каталозі, але рядка у Salary-листі для неї нема.
        // Записувати ні куди. Лиш діагностика.
        if (factByName.hasOwnProperty(lname)){
          notFound.push(factByName[lname].name);
          details.push({activity: factByName[lname].name, fact: factByName[lname].fact, status: 'not-in-salary'});
        }
        return;
      }
      var rowIdx0 = rowFound - 1;
      if (budgetColFormulas[rowIdx0] && budgetColFormulas[rowIdx0][0]){
        formulaRowsSkipped++;
        Logger.log('[exportToSalaryExtras] skipped formula row %s: %s', rowFound, budgetColFormulas[rowIdx0][0]);
        return;
      }
      var nk           = lname; // нормалізована назва — спільний ключ
      var currentValue = Number(budgetColValues[rowIdx0][0]) || 0;
      var je           = journal.byNormName[nk];
      var lastWritten  = je ? je.sum : 0;
      var baseValue    = currentValue - lastWritten;
      var info         = factByName[lname]; // може бути undefined якщо у заняття немає rate
      var newFact      = info ? info.fact : 0;
      var newValue     = baseValue + newFact;

      // Точковий запис лише змінених клітинок.
      if (newValue !== currentValue){
        sheet.getRange(rowFound, budgetCol).setValue(newValue);
        cellsWritten++;
      }

      if (newFact !== lastWritten){
        journalOps.push({
          nk: nk, loc: loc, kind: 'salary', name: a.name,
          year: nextM.year, month: nextM.month, newSum: newFact
        });
      }

      if (info){
        updated++;
        totalFact += newFact;
        details.push({
          activity: info.name, fact: newFact,
          currentBefore: currentValue, lastWritten: lastWritten,
          baseValue: baseValue, newCell: newValue,
          row: rowFound, status: 'updated'
        });
        Logger.log('[exportToSalaryExtras] WRITE row=%s "%s" cur=%s last=%s base=%s newFact=%s → %s', rowFound, a.name, currentValue, lastWritten, baseValue, newFact, newValue);
      } else if (lastWritten !== 0){
        details.push({
          activity: a.name, fact: 0,
          currentBefore: currentValue, lastWritten: lastWritten,
          baseValue: baseValue, newCell: newValue,
          row: rowFound, status: 'cleared'
        });
        Logger.log('[exportToSalaryExtras] CLEAR row=%s "%s" cur=%s last=%s base=%s → %s', rowFound, a.name, currentValue, lastWritten, baseValue, newValue);
      }
    });

    Logger.log('[exportToSalaryExtras] точковий запис: %s клітинок змінено, %s формульних рядків пропущено', cellsWritten, formulaRowsSkipped);

    _commitJournalUpdates(journal, journalOps);
    Logger.log('[exportToSalaryExtras] journal upsert: %s op(s)', journalOps.length);

    Logger.log('[exportToSalaryExtras] DONE: updated=%s, totalFact=%s, notFound=%s', updated, totalFact, JSON.stringify(notFound));

    return {
      ok: true,
      updated:  updated,
      totalFact: totalFact,
      notFound: notFound,
      skipped:  skipped,
      details:  details,
      loc: loc,
      sourceMonth: monthName,       // травень
      targetMonth: targetMonthName, // червень
      // === діагностика ===
      targetMonthNum: targetMonth,
      budgetCol: budgetCol,
      attendanceKeys: attendanceKeys,
      factByActivity: factByActivity,
      salaryRowNames: salaryRowNames,
      actRowByLname: actRowByLname,
      allActiveCount: allActive.length,
      withRateCount: withRate.length
    };
  } catch(e){
    Logger.log('[exportToSalaryExtras] EXCEPTION: %s\n%s', e && e.message, e && e.stack);
    return {ok: false, error: String(e && e.message || e)};
  }
}

// ───────────────────────────────────────────────────────────────────────────
// Тест: запусти вручну з Apps Script editor (View → Executions → дивись Логи).
// Викликає exportToSalaryExtras({loc:"Голосієво", month:6, year:2026}) і логує
// усі ключові поля окремо.
// ───────────────────────────────────────────────────────────────────────────
function testExportSalaryVolkov(){
  var result = exportToSalaryExtras({loc: 'Голосієво', month: 6, year: 2026});
  Logger.log('[testExportSalaryVolkov] result JSON: %s', JSON.stringify(result, null, 2));
  Logger.log('[testExportSalaryVolkov] error          = %s', result && result.error);
  Logger.log('[testExportSalaryVolkov] updated        = %s', result && result.updated);
  Logger.log('[testExportSalaryVolkov] totalFact      = %s', result && result.totalFact);
  Logger.log('[testExportSalaryVolkov] targetMonth    = %s (num %s)', result && result.targetMonth, result && result.targetMonthNum);
  Logger.log('[testExportSalaryVolkov] budgetCol      = %s', result && result.budgetCol);
  Logger.log('[testExportSalaryVolkov] attendanceKeys = %s', JSON.stringify(result && result.attendanceKeys));
  Logger.log('[testExportSalaryVolkov] factByActivity = %s', JSON.stringify(result && result.factByActivity));
  Logger.log('[testExportSalaryVolkov] salaryRowNames = %s', JSON.stringify(result && result.salaryRowNames));
  Logger.log('[testExportSalaryVolkov] actRowByLname  = %s', JSON.stringify(result && result.actRowByLname));
  Logger.log('[testExportSalaryVolkov] allActiveCount = %s, withRateCount = %s', result && result.allActiveCount, result && result.withRateCount);
  Logger.log('[testExportSalaryVolkov] notFound       = %s', JSON.stringify(result && result.notFound));
  Logger.log('[testExportSalaryVolkov] skipped        = %s', JSON.stringify(result && result.skipped));
  Logger.log('[testExportSalaryVolkov] details        = %s', JSON.stringify(result && result.details));
  return result;
}

function exportAttendance(params){
  var p = exportToPayments(params);
  var s = exportToSalaryExtras(params);
  return {
    ok: !!(p && p.ok && s && s.ok),
    payments: p,
    salary:   s,
    loc:      params && params.loc || '',
    month:    params && params.month || 0,
    year:     params && params.year || 0
  };
}

// ═══════════════════════════════════════════════════════════════════════════
// ВЧИТЕЛІ-ПРЕДМЕТНИКИ v6.5
// Групові заняття включені у вартість навчання (НЕ в Бюджет ДОП Payment).
// ЗП викладача = ставка × к-сть унікальних пар (група + дата), де відмічено
// ≥1 дитину. Дитячі галочки — лише статистика присутності.
// Пише ТІЛЬКИ у Salary-файл локації, у рядки "<предмет> <ставка>".
// ═══════════════════════════════════════════════════════════════════════════
var PREDMETNY_CATALOG_SHEET  = 'Предметники_Каталог';
var PREDMETNY_ATT_SHEET      = 'Предметники_Відвідуваність';
var PREDMETNY_CATALOG_HEADER = ['id','Локація','Предмет','Ставка_за_заняття','Викладач','Активне'];
var PREDMETNY_ATT_HEADER     = ['id','Дата','Локація','Група','Дитина','id_предмета','Назва_предмета','Ставка','Відмітив','Час_відмітки'];

function _getPredmetnyCatalogSheet(createIfMissing){
  var ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  var sh = ss.getSheetByName(PREDMETNY_CATALOG_SHEET);
  if (!sh && createIfMissing){
    sh = ss.insertSheet(PREDMETNY_CATALOG_SHEET);
    sh.getRange(1, 1, 1, PREDMETNY_CATALOG_HEADER.length).setValues([PREDMETNY_CATALOG_HEADER]);
    sh.setFrozenRows(1);
  }
  if (!sh) throw new Error('Sheet "' + PREDMETNY_CATALOG_SHEET + '" не знайдено.');
  return sh;
}

function _getPredmetnyAttSheet(createIfMissing){
  var ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  var sh = ss.getSheetByName(PREDMETNY_ATT_SHEET);
  if (!sh && createIfMissing){
    sh = ss.insertSheet(PREDMETNY_ATT_SHEET);
    sh.getRange(1, 1, 1, PREDMETNY_ATT_HEADER.length).setValues([PREDMETNY_ATT_HEADER]);
    sh.setFrozenRows(1);
  }
  if (!sh) throw new Error('Sheet "' + PREDMETNY_ATT_SHEET + '" не знайдено.');
  return sh;
}

function _parsePredmetnyCatRow(row){
  return {
    id:      Number(row[0]) || 0,
    loc:     String(row[1] || '').trim(),
    subject: String(row[2] || '').trim(),
    rate:    Number(row[3]) || 0,
    teacher: String(row[4] || '').trim(),
    active:  row[5] === true ||
             /^(true|так|y|1|active|активне|✅)$/i.test(String(row[5] || '').trim())
  };
}

function _parsePredmetnyAttRow(row){
  var d = row[1], dateStr;
  if (d instanceof Date){
    // Канонічна TZ проєкту — Europe/Kiev (як formatDate вище). Без явної
    // TZ getDate() рахує у TZ скрипта і дата зсувається на ±1 день, через
    // що точний фільтр getPredmetnyMarks не знаходить збережену відмітку.
    dateStr = Utilities.formatDate(d, 'Europe/Kiev', 'yyyy-MM-dd');
  } else {
    dateStr = String(d || '').trim();
  }
  return {
    id:          Number(row[0]) || 0,
    date:        dateStr,
    loc:         String(row[2] || '').trim(),
    group:       String(row[3] || '').trim(),
    child:       String(row[4] || '').trim(),
    subjectId:   Number(row[5]) || 0,
    subjectName: String(row[6] || '').trim(),
    rate:        Number(row[7]) || 0,
    markedBy:    String(row[8] || '').trim(),
    markedAt:    row[9] instanceof Date ? row[9].toISOString() : String(row[9] || '')
  };
}

function _nextPredmetnyRowId(sh){
  var data = sh.getDataRange().getValues();
  var max = 0;
  for (var i = 1; i < data.length; i++){
    var n = Number(data[i][0]) || 0;
    if (n > max) max = n;
  }
  return max + 1;
}

// ── Каталог CRUD ──────────────────────────────────────────────────────────
function getPredmetnyCatalog(loc){
  try {
    var sh = _getPredmetnyCatalogSheet(true);
    var data = sh.getDataRange().getValues();
    if (data.length < 2) return {ok: true, items: []};
    var items = [];
    var filterLoc = String(loc || '').trim();
    for (var i = 1; i < data.length; i++){
      if (!data[i][2]) continue;
      var rec = _parsePredmetnyCatRow(data[i]);
      if (filterLoc && rec.loc !== filterLoc) continue;
      items.push(rec);
    }
    return {ok: true, items: items};
  } catch(e){
    return {ok: false, error: String(e && e.message || e)};
  }
}

function addPredmetny(data){
  try {
    var sh = _getPredmetnyCatalogSheet(true);
    var id = _nextPredmetnyRowId(sh);
    var row = [
      id,
      String(data.loc || '').trim(),
      String(data.subject || '').trim(),
      Number(data.rate) || 0,
      String(data.teacher || '').trim(),
      data.active !== false
    ];
    if (!row[1]) return {ok: false, error: 'Поле "Локація" обовʼязкове'};
    if (!row[2]) return {ok: false, error: 'Поле "Предмет" обовʼязкове'};
    sh.appendRow(row);
    return {ok: true, id: id};
  } catch(e){
    return {ok: false, error: String(e && e.message || e)};
  }
}

function updatePredmetny(id, data){
  try {
    var nid = Number(id);
    if (!nid) return {ok: false, error: 'Missing id'};
    var sh = _getPredmetnyCatalogSheet(false);
    var rows = sh.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++){
      if (Number(rows[i][0]) !== nid) continue;
      var r1 = i + 1;
      if ('loc'     in data) sh.getRange(r1, 2).setValue(String(data.loc || '').trim());
      if ('subject' in data) sh.getRange(r1, 3).setValue(String(data.subject || '').trim());
      if ('rate'    in data) sh.getRange(r1, 4).setValue(Number(data.rate) || 0);
      if ('teacher' in data) sh.getRange(r1, 5).setValue(String(data.teacher || '').trim());
      if ('active'  in data) sh.getRange(r1, 6).setValue(data.active !== false);
      return {ok: true};
    }
    return {ok: false, error: 'Предмет не знайдено'};
  } catch(e){
    return {ok: false, error: String(e && e.message || e)};
  }
}

function deletePredmetny(id){
  try {
    var nid = Number(id);
    if (!nid) return {ok: false, error: 'Missing id'};
    var sh = _getPredmetnyCatalogSheet(false);
    var rows = sh.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++){
      if (Number(rows[i][0]) === nid){
        sh.deleteRow(i + 1);
        return {ok: true};
      }
    }
    return {ok: false, error: 'Предмет не знайдено'};
  } catch(e){
    return {ok: false, error: String(e && e.message || e)};
  }
}

// ── Відвідуваність ────────────────────────────────────────────────────────
function addPredmetnyMark(data){
  try {
    var sh = _getPredmetnyAttSheet(true);
    // Колонка Дата (B) — примусово текстова. Інакше Sheets конвертує
    // рядок '2026-05-18' у date-serial, і round-trip дати стає залежним
    // від TZ — точний фільтр getPredmetnyMarks втрачає відмітку.
    sh.getRange(1, 2, sh.getMaxRows(), 1).setNumberFormat('@');
    var id = _nextPredmetnyRowId(sh);
    var row = [
      id,
      String(data.date || '').trim(),
      String(data.loc || '').trim(),
      String(data.group || '').trim(),
      String(data.child || '').trim(),
      Number(data.subjectId) || 0,
      String(data.subjectName || '').trim(),
      Number(data.rate) || 0,
      String(data.markedBy || '').trim(),
      new Date()
    ];
    if (!row[1] || !row[4] || !row[5]){
      return {ok: false, error: 'Поля Дата / Дитина / id_предмета обовʼязкові'};
    }
    sh.appendRow(row);
    return {ok: true, id: id};
  } catch(e){
    return {ok: false, error: String(e && e.message || e)};
  }
}

function removePredmetnyMark(id){
  try {
    var nid = Number(id);
    if (!nid) return {ok: false, error: 'Missing id'};
    var sh = _getPredmetnyAttSheet(false);
    var data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++){
      if (Number(data[i][0]) === nid){
        sh.deleteRow(i + 1);
        return {ok: true};
      }
    }
    return {ok: false, error: 'Відмітку не знайдено'};
  } catch(e){
    return {ok: false, error: String(e && e.message || e)};
  }
}

function getPredmetnyMarks(filters){
  try {
    filters = filters || {};
    var sh = _getPredmetnyAttSheet(true);
    var data = sh.getDataRange().getValues();
    if (data.length < 2) return {ok: true, items: []};
    var items = [];
    for (var i = 1; i < data.length; i++){
      if (!data[i][0] && !data[i][4]) continue;
      var m = _parsePredmetnyAttRow(data[i]);
      if (filters.date      && m.date !== String(filters.date)) continue;
      if (filters.loc       && m.loc !== String(filters.loc)) continue;
      if (filters.group     && m.group !== String(filters.group)) continue;
      if (filters.child     && m.child !== String(filters.child)) continue;
      if (filters.subjectId && m.subjectId !== Number(filters.subjectId)) continue;
      items.push(m);
    }
    return {ok: true, items: items};
  } catch(e){
    return {ok: false, error: String(e && e.message || e)};
  }
}

// ── Seed каталогу предметників ────────────────────────────────────────────
function _normPredmetnyCatRow(row){
  var out = [];
  for (var c = 0; c < PREDMETNY_CATALOG_HEADER.length; c++){
    out.push(row[c] !== undefined ? row[c] : '');
  }
  return out;
}

// ───────────────────────────────────────────────────────────────────────────
// Seed каталогу ПРЕДМЕТНИКІВ — усі 11 локацій (Голосієво + 10 інших)
// у лист "Предметники_Каталог".
//   seedPredmetnyCatalog()       — м'який режим: якщо канонічні локації
//                                  вже мають предмети, НІЧОГО не чіпає.
//   seedPredmetnyCatalogForce()  — force: перезаписує всі 11 локацій.
// Рядки локацій поза списком зберігаються незмінними.
// ───────────────────────────────────────────────────────────────────────────
function seedPredmetnyCatalog(force){
  var sh = _getPredmetnyCatalogSheet(true);
  var data = sh.getDataRange().getValues();

  // Канонічні каталоги всіх 11 локацій. Голосієво перший → id 1..4 як було.
  // Запис: [Предмет, Ставка_за_заняття]
  var CANON = [
    ['Голосієво', [
      ['Англійська мова', 280],
      ['Логопед', 300],
      ['Муз.керівник', 300],
      ['Хореограф', 270]
    ]],
    ['Бігова', [
      ['Англійська мова', 230],
      ['Логопед', 300],
      ['Муз.керівник', 230],
      ['Хореограф', 220],
      ['Чомусики', 350]
    ]],
    ['Борщагівка', [
      ['Англійська мова', 280],
      ['Логопед', 300],
      ['Муз.керівник', 300],
      ['Хореограф', 300]
    ]],
    ['Бровари', [
      ['Англійська мова', 250],
      ['Логопед', 250],
      ['Муз.керівник', 250],
      ['Хореограф', 230]
    ]],
    ["Кар'єрна", [
      ['Англійська мова', 280],
      ['Логопед', 350],
      ['Муз.керівник', 300],
      ['Хореограф', 300]
    ]],
    ['Кругла', [
      ['Англійська мова', 230],
      ['Логопед', 300],
      ['Муз.керівник', 230],
      ['Хореограф', 220],
      ['Чомусики', 350]
    ]],
    ['Оранж', [
      ['Англійська мова', 250],
      ['Логопед', 250],
      ['Муз.керівник', 300],
      ['Хореограф', 230]
    ]],
    ['Осокорки', [
      ['Англійська мова', 250],
      ['Логопед', 300],
      ['Муз.керівник', 250],
      ['Хореограф', 350],
      ['Підготовка до школи', 450]
    ]],
    ['Позняки', [
      ['Англійська мова', 250],
      ['Логопед', 350],
      ['Муз.керівник', 250],
      ['Хореограф', 250]
    ]],
    ['Пуща', [
      ['Англійська мова', 280],
      ['Логопед', 300],
      ['Муз.керівник', 250],
      ['Хореограф', 300]
    ]],
    ['Тичини', [
      ['Англійська мова', 280],
      ['Логопед', 300],
      ['Муз.керівник', 250],
      ['Хореограф', 250]
    ]]
  ];

  var CANON_LOC = {};
  CANON.forEach(function(pair){ CANON_LOC[pair[0]] = true; });

  // Розділяємо наявні рядки: канонічні локації vs усі інші (зберігаємо як є).
  var canonExisting = 0, otherRows = [];
  for (var r = 1; r < data.length; r++){
    var row = data[r];
    if (!row[2]) continue;
    var rowLoc = String(row[1] || '').trim();
    if (CANON_LOC[rowLoc]) canonExisting++;
    else                   otherRows.push(_normPredmetnyCatRow(row));
  }

  if (canonExisting > 0 && !force){
    Logger.log('[seedPredmetnyCatalog] Канонічні локації вже мають %s предметів. seedPredmetnyCatalogForce() щоб перезаписати.', canonExisting);
    return {ok: true, skipped: true, existingRows: canonExisting};
  }

  // Канонічні рядки з наскрізними id (Голосієво перший → id 1..4).
  // Колонки: id | Локація | Предмет | Ставка_за_заняття | Викладач | Активне
  var canonRows = [], id = 1;
  CANON.forEach(function(pair){
    var loc = pair[0];
    pair[1].forEach(function(p){
      canonRows.push([id++, loc, p[0], p[1], '', true]);
    });
  });

  var allRows = otherRows.concat(canonRows);
  var lastRow = sh.getLastRow();
  if (lastRow > 1){
    sh.getRange(2, 1, lastRow - 1, PREDMETNY_CATALOG_HEADER.length).clearContent();
  }
  if (allRows.length){
    sh.getRange(2, 1, allRows.length, PREDMETNY_CATALOG_HEADER.length).setValues(allRows);
  }
  Logger.log('[seedPredmetnyCatalog] Залито %s предметів у %s локацій; інших рядків: %s (force=%s)', canonRows.length, CANON.length, otherRows.length, !!force);
  return {ok: true, seeded: canonRows.length, locations: CANON.length, keptOtherRows: otherRows.length, force: !!force};
}

function seedPredmetnyCatalogForce(){ return seedPredmetnyCatalog(true); }

// Перезаписує ОБИДВА каталоги (додаткові + предметники) для всіх 11 локацій.
// Запускати ВРУЧНУ з Apps Script editor.
function seedAllCatalogsForce(){
  var activities = seedActivitiesCatalog(true);
  var predmetny  = seedPredmetnyCatalog(true);
  Logger.log('[seedAllCatalogsForce] activities=%s | predmetny=%s',
    JSON.stringify(activities), JSON.stringify(predmetny));
  return {ok: true, activities: activities, predmetny: predmetny};
}

// ── Експорт у Salary ──────────────────────────────────────────────────────
// Архітектура як exportToSalaryExtras: журнал kind='predmetnyky', розумне
// перезаписування, толерантний матч рядків, точковий запис (skip формул).
// Знаходить рядок Salary для предмета каталогу за пріоритетами P1–P4.
//   salaryRows      — [{row, raw, norm}] непорожні рядки колонки A.
//   subject, rate   — предмет і ставка з каталогу предметників.
//   addSubjNorms    — {нормНазваДодатковогоЗаняття: true} для цієї локації.
// Повертає {row, matchedAs, priority:'P1'..'P4'} або null (P5).
function _findPredmetnySalaryRow(salaryRows, subject, rate, addSubjNorms){
  var subjNorm  = _journalNormName(subject);
  var exactNorm = _journalNormName(subject + ' ' + rate);   // напр. "логопед250"
  if (!subjNorm) return null;
  var i;

  // P1 — точний збіг "предмет ставка".
  for (i = 0; i < salaryRows.length; i++){
    if (salaryRows[i].norm === exactNorm)
      return {row: salaryRows[i].row, matchedAs: salaryRows[i].raw, priority: 'P1'};
  }
  // P2 — префікс "предмет ставка ..." (далі — не цифра, тобто ставка та сама).
  for (i = 0; i < salaryRows.length; i++){
    var n2 = salaryRows[i].norm;
    if (n2.length > exactNorm.length && n2.indexOf(exactNorm) === 0
        && !/[0-9]/.test(n2.charAt(exactNorm.length)))
      return {row: salaryRows[i].row, matchedAs: salaryRows[i].raw, priority: 'P2'};
  }
  // P3 — той самий предмет з БУДЬ-ЯКОЮ ставкою; беремо найближчу.
  var best = null, bestDiff = Infinity;
  for (i = 0; i < salaryRows.length; i++){
    var n3 = salaryRows[i].norm;
    if (n3.indexOf(subjNorm) !== 0) continue;
    var m = n3.slice(subjNorm.length).match(/^([0-9]+)/);
    if (!m) continue;
    var diff = Math.abs(Number(m[1]) - Number(rate));
    if (diff < bestDiff){ bestDiff = diff; best = salaryRows[i]; }
  }
  if (best) return {row: best.row, matchedAs: best.raw, priority: 'P3'};
  // P4 — назва без ставки; лише якщо у Додаткові_Каталог локації такого
  // заняття НЕМАЄ (інакше рядок без ставки належить додатковому).
  if (!addSubjNorms[subjNorm]){
    for (i = 0; i < salaryRows.length; i++){
      if (salaryRows[i].norm === subjNorm)
        return {row: salaryRows[i].row, matchedAs: salaryRows[i].raw, priority: 'P4'};
    }
  }
  return null;   // P5
}

function exportPredmetnyToSalary(params){
  try {
    var loc = String(params.loc || '').trim();
    var month = Number(params.month);
    var year = Number(params.year) || new Date().getFullYear();
    if (!loc) return {ok: false, error: 'Параметр loc обовʼязковий'};
    if (!month || month < 1 || month > 12) return {ok: false, error: 'month має бути 1-12'};
    Logger.log('[exportPredmetnyToSalary] START loc="%s" month=%s year=%s', loc, month, year);
    var monthName = MONTHS_CAL_UA[month - 1];

    // 1. Каталог предметників локації.
    var catRes = getPredmetnyCatalog(loc);
    if (!catRes.ok) return catRes;
    var withRate = (catRes.items || []).filter(function(a){ return a.active && a.rate > 0; });

    // 1b. Каталог ДОДАТКОВИХ занять — для гейту P4 (рядок без ставки).
    var addSubjNorms = {};
    var actRes = getActivitiesCatalog(loc);
    (((actRes && actRes.items) || [])).forEach(function(x){
      var nn = _journalNormName(x.name);
      if (nn) addSubjNorms[nn] = true;
    });

    // 2. Період + відвідуваність → унікальні (група+дата) на кожен предмет.
    var attSh = _getPredmetnyAttSheet(true);
    var attData = attSh.getDataRange().getValues();
    var mm = month < 10 ? '0' + month : String(month);
    var dateFrom = year + '-' + mm + '-01';
    var nextM = _nextMonth(month, year);
    var nmm = nextM.month < 10 ? '0' + nextM.month : String(nextM.month);
    var dateTo = nextM.year + '-' + nmm + '-01';

    var lessonsBySubj = {}; // subjectId -> {"group|date": true}
    for (var i = 1; i < attData.length; i++){
      var rec = _parsePredmetnyAttRow(attData[i]);
      if (rec.loc !== loc) continue;
      if (rec.date < dateFrom || rec.date >= dateTo) continue;
      if (!lessonsBySubj[rec.subjectId]) lessonsBySubj[rec.subjectId] = {};
      lessonsBySubj[rec.subjectId][rec.group + '|' + rec.date] = true;
    }

    // 3. Salary-файл локації.
    var reg = _salaryGetRegistry();
    if (!reg.ok) return reg;
    var entry = null;
    for (var j = 0; j < reg.rows.length; j++){
      if (reg.rows[j].loc === loc){ entry = reg.rows[j]; break; }
    }
    if (!entry) return {ok: false, error: 'Локація "' + loc + '" не знайдена у Salary-реєстрі'};
    var locSS = SpreadsheetApp.openById(entry.sheetId);
    var sheet = locSS.getSheetByName(entry.listName);
    if (!sheet) return {ok: false, error: 'Salary sheet "' + entry.listName + '" не знайдено'};

    var lastRow = Math.max(sheet.getLastRow(), 80);
    var names = sheet.getRange(1, 1, lastRow, 1).getValues();
    var targetMonth = nextM.month;
    var budgetCol = (targetMonth - 1) * 3 + 3;
    var targetMonthName = MONTHS_CAL_UA[targetMonth - 1];
    Logger.log('[exportPredmetnyToSalary] targetMonth=%s (%s), budgetCol=%s', targetMonth, targetMonthName, budgetCol);

    // Непорожні рядки колонки A (з рядка 4 — вище шапка/мета).
    var salaryRows = [];
    for (var k = 3; k < names.length; k++){
      var raw = String(names[k][0] == null ? '' : names[k][0]).trim();
      if (!raw) continue;
      salaryRows.push({row: k + 1, raw: raw, norm: _journalNormName(raw)});
    }

    var budgetColValues   = sheet.getRange(1, budgetCol, lastRow, 1).getValues();
    var budgetColFormulas = sheet.getRange(1, budgetCol, lastRow, 1).getFormulas();
    var journal = _readJournalForTarget(loc, 'predmetnyky', nextM.year, nextM.month);

    var journalOps = [];
    var updated = 0, totalFact = 0, cellsWritten = 0, formulaRowsSkipped = 0;
    var notFound = [], details = [];
    var stats = {attempts: 0, p1: 0, p2: 0, p3: 0, p4: 0, p5: 0};

    // 4. Матчинг кожного предмета каталогу → рядок Salary (P1–P5).
    withRate.forEach(function(a){
      var uniq = lessonsBySubj[a.id] ? Object.keys(lessonsBySubj[a.id]).length : 0;
      var fact = uniq * a.rate;
      var catName = a.subject + ' ' + a.rate;       // ключ журналу — з каталогу
      var nk = _journalNormName(catName);
      stats.attempts++;

      var found = _findPredmetnySalaryRow(salaryRows, a.subject, a.rate, addSubjNorms);
      if (!found){
        stats.p5++;
        notFound.push(catName);
        details.push({subject: catName, fact: fact, lessons: uniq, status: 'not-found'});
        Logger.log('[%s] %s → P5 НЕ ЗНАЙДЕНО у Salary', loc, catName);
        return;
      }
      stats['p' + found.priority.slice(1)]++;
      Logger.log('[%s] %s → matched %s на A%s "%s"', loc, catName, found.priority, found.row, found.matchedAs);

      var rowIdx0 = found.row - 1;
      if (budgetColFormulas[rowIdx0] && budgetColFormulas[rowIdx0][0]){
        formulaRowsSkipped++;
        Logger.log('[%s] %s → пропущено: формула у рядку %s', loc, catName, found.row);
        return;
      }
      var currentValue = Number(budgetColValues[rowIdx0][0]) || 0;
      var je = journal.byNormName[nk];
      var lastWritten = je ? je.sum : 0;
      var baseValue = currentValue - lastWritten;
      var newValue = baseValue + fact;
      if (newValue !== currentValue){
        sheet.getRange(found.row, budgetCol).setValue(newValue);
        cellsWritten++;
      }
      if (fact !== lastWritten){
        journalOps.push({nk: nk, loc: loc, kind: 'predmetnyky', name: catName,
          year: nextM.year, month: nextM.month, newSum: fact});
      }
      updated++;
      totalFact += fact;
      details.push({subject: catName, matchedAs: found.matchedAs, priority: found.priority,
        fact: fact, lessons: uniq, row: found.row, currentBefore: currentValue,
        lastWritten: lastWritten, newCell: newValue, status: 'updated'});
    });

    _commitJournalUpdates(journal, journalOps);
    Logger.log('[%s] СВОДКА: спроб=%s | P1=%s P2=%s P3=%s P4=%s P5=%s | записано клітинок=%s формул-пропущено=%s',
      loc, stats.attempts, stats.p1, stats.p2, stats.p3, stats.p4, stats.p5, cellsWritten, formulaRowsSkipped);

    return {
      ok: true,
      updated: updated,
      totalFact: totalFact,
      notFound: notFound,
      details: details,
      matchStats: stats,
      loc: loc,
      sourceMonth: monthName,
      targetMonth: targetMonthName,
      budgetCol: budgetCol,
      cellsWritten: cellsWritten,
      formulaRowsSkipped: formulaRowsSkipped
    };
  } catch(e){
    Logger.log('[exportPredmetnyToSalary] EXCEPTION: %s\n%s', e && e.message, e && e.stack);
    return {ok: false, error: String(e && e.message || e)};
  }
}

// Експорт предметників у Salary для всіх 11 локацій + зведена таблиця.
// УВАГА: реально пише у Salary-файли. Запускати ВРУЧНУ.
function exportAllPredmetnyToSalary(month, year){
  var LOCS = ['Голосієво','Бігова','Борщагівка','Бровари',"Кар'єрна",'Кругла',
              'Оранж','Осокорки','Позняки','Пуща','Тичини'];
  var lines = [];
  LOCS.forEach(function(loc){
    var r = exportPredmetnyToSalary({loc: loc, month: month, year: year});
    var s = (r && r.matchStats) || {attempts:0,p1:0,p2:0,p3:0,p4:0,p5:0};
    lines.push(loc + ' | ' + s.attempts + ' | ' + s.p1 + ' | ' + s.p2 + ' | ' +
      s.p3 + ' | ' + s.p4 + ' | ' + s.p5 + ((r && r.ok) ? '' : '  ❌ ' + (r && r.error)));
  });
  Logger.log('\n════════ SUMMARY предметники → Salary (місяць %s/%s) ════════', month, year);
  Logger.log('Локація | Спроб | P1 точн | P2 префікс | P3 інша ст | P4 без ст | P5 не зн');
  lines.forEach(function(x){ Logger.log('  ' + x); });
  return {ok: true};
}

// Тест: запусти вручну з Apps Script editor.
function testExportPredmetny(){
  var result = exportPredmetnyToSalary({loc: 'Голосієво', month: 6, year: 2026});
  Logger.log('[testExportPredmetny] result: %s', JSON.stringify(result, null, 2));
  return result;
}

// ═══════════════════════════════════════════════════════════════════════════
// ЗАДАЧНИК (v6.6) — управління задачами в команді.
// Листи CONFIG: "Задачі" + "Задачі_Активність" (журнал коментарів/подій).
// Активність задач — окремий лист, НЕ пов'язаний з Експорт_Журналом.
// ═══════════════════════════════════════════════════════════════════════════
var TASKS_SHEET_NAME     = 'Задачі';
var TASKS_ACT_SHEET_NAME = 'Задачі_Активність';
var TASKS_HEADER = ['id','created_at','author','assignee','title','description',
                    'priority','deadline','location','status','group_id','parent_id'];
var TASKS_ACT_HEADER = ['id','task_id','author','type','content','file_url','created_at'];

// Ролі-менеджмент: ставлять задачі (tasks-create). director — лише отримує.
var TASK_MGMT_ROLES = ['cfo','ceo','cco','coo','hr','cmo','rnd_director','hr_trainer','legal'];

function _getTasksSheet(create){
  var ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  var sh = ss.getSheetByName(TASKS_SHEET_NAME);
  if (!sh && create){
    sh = ss.insertSheet(TASKS_SHEET_NAME);
    sh.getRange(1,1,1,TASKS_HEADER.length).setValues([TASKS_HEADER]);
    sh.setFrozenRows(1);
  }
  if (!sh) throw new Error('Лист "'+TASKS_SHEET_NAME+'" не знайдено.');
  // created_at(2) / deadline(8) — текстові, щоб дати не плавали по TZ.
  sh.getRange(1,2,sh.getMaxRows(),1).setNumberFormat('@');
  sh.getRange(1,8,sh.getMaxRows(),1).setNumberFormat('@');
  return sh;
}
function _getTaskActSheet(create){
  var ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  var sh = ss.getSheetByName(TASKS_ACT_SHEET_NAME);
  if (!sh && create){
    sh = ss.insertSheet(TASKS_ACT_SHEET_NAME);
    sh.getRange(1,1,1,TASKS_ACT_HEADER.length).setValues([TASKS_ACT_HEADER]);
    sh.setFrozenRows(1);
  }
  if (!sh) throw new Error('Лист "'+TASKS_ACT_SHEET_NAME+'" не знайдено.');
  sh.getRange(1,7,sh.getMaxRows(),1).setNumberFormat('@');
  return sh;
}

function _taskDateStr(v){
  if (v instanceof Date) return Utilities.formatDate(v,'Europe/Kiev','yyyy-MM-dd');
  return String(v == null ? '' : v).trim();
}
function _taskNow(){   return Utilities.formatDate(new Date(),'Europe/Kiev','yyyy-MM-dd HH:mm'); }
function _taskToday(){ return Utilities.formatDate(new Date(),'Europe/Kiev','yyyy-MM-dd'); }

function _parseTaskRow(row){
  return {
    id:          Number(row[0]) || 0,
    createdAt:   String(row[1] || '').trim(),
    author:      Number(row[2]) || 0,
    assignee:    Number(row[3]) || 0,
    title:       String(row[4] || '').trim(),
    description: String(row[5] || '').trim(),
    priority:    String(row[6] || 'normal').trim() || 'normal',
    deadline:    _taskDateStr(row[7]),
    location:    String(row[8] || '').trim(),
    status:      String(row[9] || 'new').trim() || 'new',
    groupId:     String(row[10] || '').trim(),
    parentId:    Number(row[11]) || 0
  };
}
function _parseTaskActRow(row){
  return {
    id:        Number(row[0]) || 0,
    taskId:    Number(row[1]) || 0,
    author:    Number(row[2]) || 0,
    type:      String(row[3] || '').trim(),
    content:   String(row[4] || '').trim(),
    fileUrl:   String(row[5] || '').trim(),
    createdAt: String(row[6] || '').trim()
  };
}
function _nextTaskRowId(sh){
  var data = sh.getDataRange().getValues();
  var max = 0;
  for (var i=1;i<data.length;i++){ var n=Number(data[i][0])||0; if(n>max)max=n; }
  return max+1;
}
// Карта користувачів {id: {name, role, email, ...}} — для резолву імен/email.
function _taskUserMap(){
  var sh = _getUsersSheet();
  var data = sh.getDataRange().getValues();
  var map = {};
  for (var i=1;i<data.length;i++){
    if (!data[i][0]) continue;
    var u = _parseUserRow(data[i]);
    map[u.id] = u;
  }
  return map;
}
function _taskMail(toEmail, subject, body){
  if (!toEmail) return;
  try { MailApp.sendEmail(toEmail, subject, body); }
  catch(e){ Logger.log('[tasks] mail fail to %s: %s', toEmail, e); }
}

// ── createTask ─────────────────────────────────────────────────────────────
// assignee може бути id користувача АБО 'ALL_DIRECTORS' / 'ALL_MANAGEMENT' —
// тоді створюється N копій задачі з одним group_id (групова задача).
// Резолв виконавців: assignee може бути id, спец-рядком
// ('ALL_DIRECTORS' / 'ALL_MANAGEMENT') АБО масивом будь-чого з цього
// (мультивибір людей / вибір за локацією на фронті). Повертає унікальні id.
function _resolveAssignees(raw, users){
  var items = (raw && Array.isArray(raw)) ? raw : [raw];
  var out = [], seen = {};
  function add(n){ n = Number(n) || 0; if (n && !seen[n]){ seen[n] = 1; out.push(n); } }
  items.forEach(function(it){
    if (it === 'ALL_DIRECTORS'){
      Object.keys(users).forEach(function(id){
        if (users[id].active && users[id].role === 'director') add(id);
      });
    } else if (it === 'ALL_MANAGEMENT'){
      Object.keys(users).forEach(function(id){
        if (users[id].active && TASK_MGMT_ROLES.indexOf(users[id].role) !== -1) add(id);
      });
    } else {
      add(it);
    }
  });
  return out;
}

function createTask(params){
  try {
    params = params || {};
    var sh = _getTasksSheet(true);
    var users = _taskUserMap();
    var author = Number(params.author) || 0;
    var title  = String(params.title || '').trim();
    if (!author) return {ok:false, error:'Не вказано автора'};
    if (!title)  return {ok:false, error:'Вкажіть назву задачі'};

    var assigneeList = _resolveAssignees(params.assignee, users);
    if (!assigneeList.length) return {ok:false, error:'Не вказано виконавця'};

    var isGroup  = assigneeList.length > 1;
    var groupId  = isGroup ? ('g' + Date.now()) : '';
    var now      = _taskNow();
    var priority = String(params.priority || 'normal').trim() || 'normal';
    var deadline = String(params.deadline || '').trim();
    var location = String(params.location || '').trim();
    var descr    = String(params.description || '').trim();
    var parentId = Number(params.parentId) || '';
    var actSh    = _getTaskActSheet(true);
    var ids = [];

    assigneeList.forEach(function(aid){
      var id = _nextTaskRowId(sh);
      sh.appendRow([id, now, author, aid, title, descr, priority,
                    deadline, location, 'new', groupId, parentId]);
      actSh.appendRow([_nextTaskRowId(actSh), id, author, 'created',
                       'Задачу створено', '', now]);
      ids.push(id);
      var au = users[aid];
      if (au && au.email){
        _taskMail(au.email, 'Нова задача: ' + title,
          'Вам поставлено задачу "' + title + '".\n' +
          (descr ? descr + '\n' : '') +
          (deadline ? 'Дедлайн: ' + deadline + '\n' : '') +
          'Пріоритет: ' + priority + '\n\nВідкрийте Задачник у системі m.kids.');
      }
    });
    return {ok:true, ids:ids, groupId:groupId, count:ids.length};
  } catch(e){
    return {ok:false, error:String(e && e.message || e)};
  }
}

// ── getTasks ───────────────────────────────────────────────────────────────
function getTasks(filters){
  try {
    filters = filters || {};
    var viewerId   = Number(filters.viewerId) || 0;
    var viewerRole = String(filters.viewerRole || '').trim().toLowerCase();
    var isMgmt = (viewerRole === 'cfo' || viewerRole === 'hr' ||
                  TASK_MGMT_ROLES.indexOf(viewerRole) !== -1);
    var scope = String(filters.scope || 'mine_assigned');
    if (!isMgmt) scope = 'mine_assigned'; // директор локації — лише свої

    var sh = _getTasksSheet(true);
    var data = sh.getDataRange().getValues();
    var users = _taskUserMap();
    var all = [], groupStat = {};

    for (var i=1;i<data.length;i++){
      if (!data[i][0]) continue;
      var t = _parseTaskRow(data[i]);
      if (t.status === 'deleted') continue;
      all.push(t);
      if (t.groupId){
        if (!groupStat[t.groupId]) groupStat[t.groupId] = {total:0, done:0};
        groupStat[t.groupId].total++;
        if (t.status === 'done') groupStat[t.groupId].done++;
      }
    }

    var tasks = [];
    all.forEach(function(t){
      if (scope === 'mine_assigned' && t.assignee !== viewerId) return;
      if (scope === 'mine_authored' && t.author   !== viewerId) return;
      if (filters.status   && t.status   !== filters.status)   return;
      if (filters.priority && t.priority !== filters.priority) return;
      if (filters.location && t.location !== filters.location) return;
      if (filters.search){
        var q = String(filters.search).toLowerCase();
        if ((t.title+' '+t.description).toLowerCase().indexOf(q) === -1) return;
      }
      var au = users[t.author], as = users[t.assignee];
      tasks.push({
        id:t.id, createdAt:t.createdAt, title:t.title, description:t.description,
        priority:t.priority, deadline:t.deadline, location:t.location,
        status:t.status, groupId:t.groupId, parentId:t.parentId,
        author:t.author, authorName: au ? au.name : ('#'+t.author),
        assignee:t.assignee, assigneeName: as ? as.name : ('#'+t.assignee),
        group: t.groupId ? groupStat[t.groupId] : null
      });
    });
    return {ok:true, tasks:tasks};
  } catch(e){
    return {ok:false, error:String(e && e.message || e)};
  }
}

// ── updateTaskStatus ───────────────────────────────────────────────────────
function updateTaskStatus(taskId, newStatus, actorId){
  try {
    var nid = Number(taskId) || 0;
    var VALID = ['new','in_progress','done'];
    if (!nid) return {ok:false, error:'Missing taskId'};
    if (VALID.indexOf(newStatus) === -1) return {ok:false, error:'Невірний статус'};
    var sh = _getTasksSheet(true);
    var data = sh.getDataRange().getValues();
    var users = _taskUserMap();
    for (var i=1;i<data.length;i++){
      if (Number(data[i][0]) === nid){
        var t = _parseTaskRow(data[i]);
        sh.getRange(i+1, 10).setValue(newStatus); // колонка status
        var now = _taskNow();
        _getTaskActSheet(true).appendRow(
          [_nextTaskRowId(_getTaskActSheet(true)), nid, Number(actorId)||0,
           'status_change', newStatus, '', now]);
        var author = users[t.author];
        if (newStatus === 'done' && author && author.email){
          _taskMail(author.email, 'Задача виконана: '+t.title,
            'Виконавець позначив задачу "'+t.title+'" як виконану.');
        }
        return {ok:true, status:newStatus};
      }
    }
    return {ok:false, error:'Задачу не знайдено'};
  } catch(e){
    return {ok:false, error:String(e && e.message || e)};
  }
}

// ── updateTask (редагування — лише автор) ─────────────────────────────────
function updateTask(taskId, data, actorId){
  try {
    var nid = Number(taskId) || 0;
    if (!nid) return {ok:false, error:'Missing taskId'};
    data = data || {};
    var actor = Number(actorId) || 0;
    var title = String(data.title || '').trim();
    if (!title) return {ok:false, error:'Вкажіть назву задачі'};
    var sh = _getTasksSheet(true);
    var rows = sh.getDataRange().getValues();
    for (var i=1;i<rows.length;i++){
      if (Number(rows[i][0]) === nid){
        var t = _parseTaskRow(rows[i]);
        if (t.author !== actor) return {ok:false, error:'Редагувати може лише автор'};
        sh.getRange(i+1, 5).setValue(title);                                            // title
        sh.getRange(i+1, 6).setValue(String(data.description || '').trim());             // description
        sh.getRange(i+1, 7).setValue(String(data.priority || 'normal').trim()||'normal');// priority
        sh.getRange(i+1, 8).setValue(String(data.deadline || '').trim());                // deadline
        sh.getRange(i+1, 9).setValue(String(data.location || '').trim());                // location
        var actSh = _getTaskActSheet(true);
        actSh.appendRow([_nextTaskRowId(actSh), nid, actor, 'comment',
                         '✏️ Задачу відредаговано', '', _taskNow()]);
        return {ok:true};
      }
    }
    return {ok:false, error:'Задачу не знайдено'};
  } catch(e){
    return {ok:false, error:String(e && e.message || e)};
  }
}

// ── addTaskComment ─────────────────────────────────────────────────────────
function addTaskComment(taskId, comment, fileUrl, actorId){
  try {
    var nid = Number(taskId) || 0;
    if (!nid) return {ok:false, error:'Missing taskId'};
    comment = String(comment || '').trim();
    fileUrl = String(fileUrl || '').trim();
    if (!comment && !fileUrl) return {ok:false, error:'Порожній коментар'};
    var sh = _getTasksSheet(true);
    var data = sh.getDataRange().getValues();
    var users = _taskUserMap();
    var task = null;
    for (var i=1;i<data.length;i++){
      if (Number(data[i][0]) === nid){ task = _parseTaskRow(data[i]); break; }
    }
    if (!task) return {ok:false, error:'Задачу не знайдено'};
    var actSh = _getTaskActSheet(true);
    var aId = Number(actorId) || 0;
    actSh.appendRow([_nextTaskRowId(actSh), nid, aId,
                     fileUrl ? 'file' : 'comment', comment, fileUrl, _taskNow()]);
    var otherId = (aId === task.author) ? task.assignee : task.author;
    var other = users[otherId];
    if (other && other.email){
      _taskMail(other.email, 'Новий коментар у задачі: '+task.title,
        (users[aId] ? users[aId].name : 'Користувач') + ': ' +
        (comment || '[файл]') + (fileUrl ? '\nФайл: '+fileUrl : ''));
    }
    return {ok:true};
  } catch(e){
    return {ok:false, error:String(e && e.message || e)};
  }
}

// ── deleteTask (м'яке видалення — лише автор) ──────────────────────────────
function deleteTask(taskId, actorId){
  try {
    var nid = Number(taskId) || 0;
    if (!nid) return {ok:false, error:'Missing taskId'};
    var actor = Number(actorId) || 0;
    var sh = _getTasksSheet(true);
    var data = sh.getDataRange().getValues();

    // Знаходимо задачу.
    var found = null;
    for (var i=1;i<data.length;i++){
      if (Number(data[i][0]) === nid){ found = _parseTaskRow(data[i]); break; }
    }
    if (!found) return {ok:false, error:'Задачу не знайдено'};
    if (found.author !== actor) return {ok:false, error:'Видалити може лише автор'};

    // Групова задача (є group_id) — видаляємо ВСІ копії групи, інакше лише цю.
    var actSh = _getTaskActSheet(true);
    var deleted = 0;
    for (var j=1;j<data.length;j++){
      var t = _parseTaskRow(data[j]);
      if (t.status === 'deleted') continue;
      var inScope = found.groupId ? (t.groupId === found.groupId) : (t.id === nid);
      if (!inScope || t.author !== actor) continue;
      sh.getRange(j+1, 10).setValue('deleted');   // колонка status
      actSh.appendRow([_nextTaskRowId(actSh), t.id, actor,
                       'status_change', 'deleted', '', _taskNow()]);
      deleted++;
    }
    return {ok:true, deleted:deleted};
  } catch(e){
    return {ok:false, error:String(e && e.message || e)};
  }
}

// ── getTaskActivity ────────────────────────────────────────────────────────
function getTaskActivity(taskId){
  try {
    var nid = Number(taskId) || 0;
    if (!nid) return {ok:false, error:'Missing taskId'};
    var sh = _getTaskActSheet(true);
    var data = sh.getDataRange().getValues();
    var users = _taskUserMap();
    var items = [];
    for (var i=1;i<data.length;i++){
      if (!data[i][0]) continue;
      var a = _parseTaskActRow(data[i]);
      if (a.taskId !== nid) continue;
      var u = users[a.author];
      items.push({
        id:a.id, type:a.type, content:a.content, fileUrl:a.fileUrl,
        createdAt:a.createdAt, author:a.author,
        authorName: u ? u.name : ('#'+a.author)
      });
    }
    items.sort(function(x,y){ return x.id - y.id; });
    return {ok:true, items:items};
  } catch(e){
    return {ok:false, error:String(e && e.message || e)};
  }
}

// ── getDashboardNotifications ──────────────────────────────────────────────
// Бейдж = нові призначені задачі + свіжі коментарі/зміни (останні 2 дні).
function getDashboardNotifications(userId, role){
  try {
    var uid = Number(userId) || 0;
    var sh = _getTasksSheet(true);
    var data = sh.getDataRange().getValues();
    var users = _taskUserMap();
    var tomorrow = Utilities.formatDate(new Date(Date.now()+86400000),'Europe/Kiev','yyyy-MM-dd');

    var myTasks = {}, newTasks = 0, overdueDueTomorrow = 0;
    for (var i=1;i<data.length;i++){
      if (!data[i][0]) continue;
      var t = _parseTaskRow(data[i]);
      if (t.status === 'deleted') continue;
      if (t.assignee !== uid && t.author !== uid) continue;
      myTasks[t.id] = t;
      if (t.assignee === uid && t.status === 'new') newTasks++;
      if (t.assignee === uid && t.status !== 'done' && t.deadline &&
          t.deadline <= tomorrow) overdueDueTomorrow++;
    }
    var actSh = _getTaskActSheet(true);
    var adata = actSh.getDataRange().getValues();
    var cutoff = Utilities.formatDate(new Date(Date.now()-2*86400000),
                                      'Europe/Kiev','yyyy-MM-dd HH:mm');
    var events = [], comments = 0;
    for (var j=adata.length-1;j>=1;j--){
      if (!adata[j][0]) continue;
      var a = _parseTaskActRow(adata[j]);
      if (!myTasks[a.taskId]) continue;
      if (a.author === uid) continue;
      if (a.createdAt < cutoff) continue;
      if (a.type === 'comment' || a.type === 'file') comments++;
      if (events.length < 10){
        var u = users[a.author];
        events.push({
          taskId:a.taskId, taskTitle: myTasks[a.taskId].title,
          type:a.type, content:a.content,
          who: u ? u.name : ('#'+a.author), at:a.createdAt
        });
      }
    }
    return {ok:true, newTasks:newTasks, overdueDueTomorrow:overdueDueTomorrow,
            comments:comments, badge:(newTasks + comments), events:events};
  } catch(e){
    return {ok:false, error:String(e && e.message || e)};
  }
}

// ── Time-trigger: щоденні нагадування о 09:00 ──────────────────────────────
// setupTaskReminders() запустити ВРУЧНУ один раз із Apps Script editor.
function setupTaskReminders(){
  var triggers = ScriptApp.getProjectTriggers();
  for (var i=0;i<triggers.length;i++){
    if (triggers[i].getHandlerFunction() === 'dailyTaskReminders')
      ScriptApp.deleteTrigger(triggers[i]);
  }
  ScriptApp.newTrigger('dailyTaskReminders')
    .timeBased().atHour(9).everyDays(1)
    .inTimezone('Europe/Kiev').create();
  Logger.log('[setupTaskReminders] daily 09:00 trigger створено');
  return {ok:true};
}
function dailyTaskReminders(){
  var sh = _getTasksSheet(true);
  var data = sh.getDataRange().getValues();
  var users = _taskUserMap();
  var today = _taskToday();
  var tomorrow = Utilities.formatDate(new Date(Date.now()+86400000),'Europe/Kiev','yyyy-MM-dd');
  var sent = 0;
  for (var i=1;i<data.length;i++){
    if (!data[i][0]) continue;
    var t = _parseTaskRow(data[i]);
    if (t.status === 'done' || t.status === 'deleted' || !t.deadline) continue;
    var assignee = users[t.assignee], author = users[t.author];
    if (t.deadline < today){
      if (assignee && assignee.email){
        _taskMail(assignee.email,'Задача прострочена: '+t.title,
          'Задача "'+t.title+'" прострочена (дедлайн '+t.deadline+').'); sent++; }
      if (author && author.email && t.author !== t.assignee){
        _taskMail(author.email,'Задача прострочена: '+t.title,
          'Задача "'+t.title+'" (виконавець '+(assignee?assignee.name:'?')+') прострочена.'); sent++; }
    } else if (t.deadline === today){
      if (assignee && assignee.email){
        _taskMail(assignee.email,'У вас задача на сьогодні: '+t.title,
          'Сьогодні дедлайн задачі "'+t.title+'".'); sent++; }
    } else if (t.deadline === tomorrow){
      if (assignee && assignee.email){
        _taskMail(assignee.email,'Завтра дедлайн: '+t.title,
          'Завтра дедлайн задачі "'+t.title+'".'); sent++; }
    }
  }
  Logger.log('[dailyTaskReminders] листів надіслано: %s', sent);
  return {ok:true, sent:sent};
}

// ═══════════════════════════════════════════════════════════════════════════
// МІГРАЦІЯ КОРИСТУВАЧІВ (v6.7) — директори і медсестри усіх локацій
// переносяться у єдиний лист "Користувачі". Паролі — SHA-256.
// Запуск: addAllDirectorsAndNurses() ВРУЧНУ з Apps Script editor (один раз).
// ═══════════════════════════════════════════════════════════════════════════

// SHA-256 → hex lowercase.
function _sha256(str){
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,
                                      String(str == null ? '' : str),
                                      Utilities.Charset.UTF_8);
  var hex = '';
  for (var i = 0; i < bytes.length; i++){
    var b = (bytes[i] + 256) % 256;
    hex += (b < 16 ? '0' : '') + b.toString(16);
  }
  return hex;
}

// Транслітерація кирилиці → латиниця lowercase, без пробілів/розділових.
// Точна копія translitUA з index.html — щоб slug (а отже й пароль) збігались.
function _translitUA(str){
  var T = {'а':'a','б':'b','в':'v','г':'h','ґ':'g','д':'d','е':'e','є':'ye',
    'ж':'zh','з':'z','и':'y','і':'i','ї':'i','й':'y','к':'k','л':'l','м':'m',
    'н':'n','о':'o','п':'p','р':'r','с':'s','т':'t','у':'u','ф':'f','х':'kh',
    'ц':'ts','ч':'ch','ш':'sh','щ':'shch','ь':'','ю':'yu','я':'ya'};
  var s = String(str || '').toLowerCase(), out = '';
  for (var i = 0; i < s.length; i++){
    var c = s[i];
    if (T[c] !== undefined) out += T[c];
    else if (/[a-z0-9]/.test(c)) out += c;
  }
  return out;
}

// Пароль локації — копія genLocPw з index.html: <slug>2025 (директор),
// <slug>2025n (медсестра), <slug>2025v (вихователь).
function _locPassword(loc, role){
  var base = _translitUA(loc) + '2026';
  if (role === 'nurse')     return base + 'n';
  if (role === 'vyhovatel') return base + 'v';
  return base;
}

// Перехешовує усі НЕ-хешовані паролі у листі "Користувачі" (значення не
// змінюється — лише формат: plaintext → SHA-256). Ідемпотентно.
function _rehashManagementPasswords(){
  var sh = _getUsersSheet();
  var data = sh.getDataRange().getValues();
  var n = 0;
  for (var i = 1; i < data.length; i++){
    if (!data[i][0]) continue;
    var pw = String(data[i][3] == null ? '' : data[i][3]);
    if (!pw) continue;
    if (/^[0-9a-f]{64}$/i.test(pw)) continue;   // вже SHA-256
    sh.getRange(i + 1, 4).setValue(_sha256(pw));
    n++;
  }
  return {ok: true, rehashed: n};
}

// Активні локації (паролі генеруються алгоритмом _locPassword — окремого
// листа "Налаштування Паролі Локацій" у системі немає). Назви точно
// відповідають LOGIN_LOCATIONS у index.html / реєстру Sheets.
var LOCATION_USER_LOCS = [
  'Осокорки','Позняки','Тичини',"Кар'єрна",'Голосієво','Пуща','Оранж',
  'Борщагівка','Бровари','Кругла','Бігова',
  'Школа Осокорки','Школа 228',
  'Житомир','Нац.Гвардії (Благо)','Манхетен (Благо)',
  'Кухня Київ','Кухня Львів','Іва-Франківськ кухня'
];

// Створює рядки директорів / медсестер / вихователів для всіх локацій.
// roles — необовʼязковий масив (напр. ['vyhovatel']); за замовч. усі три.
// Якщо логін уже існує — пропускає (не перезаписує). Повертає лічильники.
function migrateAllLocationUsers(roles){
  try {
    var ALL = [['director','Директор'], ['nurse','Медсестра'], ['vyhovatel','Вихователь']];
    var pick = (roles && roles.length)
      ? ALL.filter(function(r){ return roles.indexOf(r[0]) !== -1; })
      : ALL;
    var sh = _getUsersSheet();
    var data = sh.getDataRange().getValues();
    var existing = {}, maxId = 0;
    for (var i = 1; i < data.length; i++){
      if (!data[i][0]) continue;
      existing[String(data[i][2] || '').trim().toLowerCase()] = true;
      var n = Number(data[i][0]) || 0; if (n > maxId) maxId = n;
    }
    var nextId = maxId + 1;
    var counts = {director:0, nurse:0, vyhovatel:0, skipped:0}, rows = [];
    LOCATION_USER_LOCS.forEach(function(loc){
      var slug = _translitUA(loc);
      pick.forEach(function(rd){
        var role = rd[0], lbl = rd[1];
        var login = role + '.' + slug;
        if (existing[login]){ counts.skipped++; return; }
        var pwHash = _sha256(_locPassword(loc, role));
        // Колонки: id | name | login | password | role | loc | email | active | lastLogin
        rows.push([nextId++, lbl + ' ' + loc, login, pwHash, role, loc, '', true, '']);
        existing[login] = true;
        counts[role]++;
      });
    });
    if (rows.length){
      sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    }
    Logger.log('[migrateAllLocationUsers] Створено %s директорів, %s медсестер, %s вихователів, %s пропущено',
      counts.director, counts.nurse, counts.vyhovatel, counts.skipped);
    return {ok: true, directors: counts.director, nurses: counts.nurse,
            vyhovateli: counts.vyhovatel, skipped: counts.skipped};
  } catch(e){
    return {ok: false, error: String(e && e.message || e)};
  }
}

// Разова утиліта — запустити ВРУЧНУ з Apps Script editor.
// Перехешовує наявні паролі + створює директорів і медсестер.
function addAllDirectorsAndNurses(){
  var rehash = _rehashManagementPasswords();
  var mig = migrateAllLocationUsers(['director','nurse']);
  if (!mig.ok){
    Logger.log('[addAllDirectorsAndNurses] ПОМИЛКА: %s', mig.error);
    return mig;
  }
  Logger.log('[addAllDirectorsAndNurses] Паролів перехешовано: %s | ' +
             'Створено %s директорів, %s медсестер, %s пропущено (вже існує)',
    rehash.rehashed, mig.directors, mig.nurses, mig.skipped);
  return {ok: true, rehashed: rehash.rehashed,
          directors: mig.directors, nurses: mig.nurses, skipped: mig.skipped};
}

// Разова утиліта — додає ЛИШЕ вихователів (vyhovatel.<slug>) для всіх
// локацій. Запустити ВРУЧНУ з Apps Script editor.
function addAllVyhovateli(){
  var mig = migrateAllLocationUsers(['vyhovatel']);
  if (!mig.ok){
    Logger.log('[addAllVyhovateli] ПОМИЛКА: %s', mig.error);
    return mig;
  }
  Logger.log('[addAllVyhovateli] Створено %s вихователів, %s пропущено (вже існує)',
    mig.vyhovateli, mig.skipped);
  return {ok: true, vyhovateli: mig.vyhovateli, skipped: mig.skipped};
}

// ── Керування паролями локаційних користувачів (v6.7) ──────────────────────
// Пише реальні SHA-256 паролі у лист "Користувачі". Раніше сторінка
// налаштувань редагувала лише localStorage — пароль входу не змінювався.
var PW_ADMIN_ROLES = ['cfo','hr','ceo','cco','coo'];

// Перевірка ролі через session: actorId — id поточного користувача;
// роль перечитується з листа "Користувачі" (не довіряємо клієнту).
function _isPasswordAdmin(actorId){
  var id = Number(actorId) || 0;
  if (!id) return false;
  var sh = _getUsersSheet();
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++){
    if (Number(data[i][0]) === id){
      var role = String(data[i][4] || '').toLowerCase().trim();
      return PW_ADMIN_ROLES.indexOf(role) !== -1;
    }
  }
  return false;
}

// Записує SHA-256(newPassword) у рядок листа "Користувачі" з логіном username.
function setUserPassword(username, newPassword, actorId){
  try {
    if (!_isPasswordAdmin(actorId))
      return {ok:false, error:'Лише CFO/CEO/CCO/COO можуть міняти паролі'};
    username    = String(username || '').trim();
    newPassword = String(newPassword == null ? '' : newPassword);
    if (!username)    return {ok:false, error:'Не вказано логін'};
    if (!newPassword) return {ok:false, error:'Порожній пароль'};
    var sh = _getUsersSheet();
    var data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++){
      if (String(data[i][2] || '').trim() === username){
        sh.getRange(i + 1, 4).setValue(_sha256(newPassword));
        return {ok:true, username:username};
      }
    }
    return {ok:false, error:'Користувача "' + username + '" не знайдено'};
  } catch(e){
    return {ok:false, error:String(e && e.message || e)};
  }
}

// Скидає паролі всіх локаційних користувачів (director/nurse/vyhovatel
// × LOCATION_USER_LOCS) до стандартних _locPassword. Лише адмін.
// Ядро: переписує паролі всіх локаційних користувачів до стандартних
// _locPassword. Повертає {updated, missing}. Без перевірки ролі.
function _reseedAllLocationPasswords(){
  var sh = _getUsersSheet();
  var data = sh.getDataRange().getValues();
  var byLogin = {};
  for (var i = 1; i < data.length; i++){
    if (!data[i][0]) continue;
    byLogin[String(data[i][2] || '').trim()] = i;
  }
  var roles = ['director','nurse','vyhovatel'];
  var updated = 0, missing = [];
  LOCATION_USER_LOCS.forEach(function(loc){
    var slug = _translitUA(loc);
    roles.forEach(function(role){
      var login = role + '.' + slug;
      var idx = byLogin[login];
      if (idx === undefined){ missing.push(login); return; }
      sh.getRange(idx + 1, 4).setValue(_sha256(_locPassword(loc, role)));
      updated++;
    });
  });
  return {updated:updated, missing:missing};
}

function resetAllLocationPasswords(actorId){
  try {
    if (!_isPasswordAdmin(actorId))
      return {ok:false, error:'Лише CFO/CEO/CCO/COO можуть міняти паролі'};
    var r = _reseedAllLocationPasswords();
    Logger.log('[resetAllLocationPasswords] оновлено %s, відсутні: %s',
      r.updated, JSON.stringify(r.missing));
    return {ok:true, updated:r.updated, missing:r.missing};
  } catch(e){
    return {ok:false, error:String(e && e.message || e)};
  }
}

// Разова утиліта — запустити ВРУЧНУ з Apps Script editor після зміни
// року у _locPassword (2025 → 2026). Переписує паролі всіх 19 локацій
// × 3 ролі на нові 2026-паролі. Старі 2025-паролі перестануть діяти.
function reseedPasswordsFor2026(){
  var r = _reseedAllLocationPasswords();
  Logger.log('[reseedPasswordsFor2026] Оновлено %s паролів%s', r.updated,
    r.missing.length ? (' · відсутні логіни: ' + JSON.stringify(r.missing)) : '');
  return {ok:true, updated:r.updated, missing:r.missing};
}

// ── ДІАГНОСТИКА: предметники у Salary-файлах ───────────────────────────────
// auditPredmetnyInSalary() — для всіх 11 локацій порівнює каталог предметників
// із секцією "Вчителі-предметники" у Salary-файлі локації. Запускати ВРУЧНУ
// з Apps Script editor; звіт — у Execution log.
function auditPredmetnyInSalary(){
  var LOCS = ['Голосієво','Бігова','Борщагівка','Бровари',"Кар'єрна",'Кругла',
              'Оранж','Осокорки','Позняки','Пуща','Тичини'];
  var reg = _salaryGetRegistry();
  if (!reg.ok){ Logger.log('Salary-реєстр: %s', reg.error); return reg; }
  var regByLoc = {};
  reg.rows.forEach(function(r){ regByLoc[r.loc] = r; });
  var summary = [];

  LOCS.forEach(function(loc){
    Logger.log('\n════════════════ %s ════════════════', loc);
    var entry = regByLoc[loc];
    if (!entry){
      Logger.log('  ❌ локацію не знайдено у Salary-реєстрі CONFIG');
      summary.push(loc + ' | — | — | — | нема в реєстрі'); return;
    }
    var catItems = ((getPredmetnyCatalog(loc) || {}).items || [])
      .filter(function(a){ return a.active; });

    var sheet;
    try {
      sheet = SpreadsheetApp.openById(entry.sheetId).getSheetByName(entry.listName);
    } catch(e){
      Logger.log('  ❌ доступ до Salary-файлу: %s', e && e.message || e);
      summary.push(loc + ' | ' + catItems.length + ' | — | — | помилка доступу'); return;
    }
    if (!sheet){
      Logger.log('  ❌ лист "%s" не знайдено', entry.listName);
      summary.push(loc + ' | ' + catItems.length + ' | — | — | нема листа'); return;
    }

    var lastRow = Math.max(sheet.getLastRow(), 60);
    var colA = sheet.getRange(1, 1, lastRow, 1).getValues();

    // Розбір колонки A на секції за заголовками.
    var section = 'header';
    var sec = {predmetny:[], dodatkovi:[], other:[]};
    for (var i = 0; i < colA.length; i++){
      var raw = String(colA[i][0] == null ? '' : colA[i][0]).trim();
      if (!raw) continue;
      var low = raw.toLowerCase();
      if (low.indexOf('предметник') !== -1 || low.indexOf('вчител') !== -1){ section = 'predmetny'; continue; }
      if (low.indexOf('додатков') !== -1){ section = 'dodatkovi'; continue; }
      var bucket = (section === 'predmetny') ? sec.predmetny
                 : (section === 'dodatkovi') ? sec.dodatkovi : sec.other;
      bucket.push({row: i + 1, raw: raw, norm: _journalNormName(raw)});
    }

    Logger.log('  Salary: %s [%s]', entry.sheetId, entry.listName);
    Logger.log('  ── Секція "Вчителі-предметники" — %s рядків:', sec.predmetny.length);
    if (!sec.predmetny.length){
      Logger.log('     (секцію не знайдено — повний дамп колонки A:)');
      sec.other.concat(sec.dodatkovi).forEach(function(r){
        Logger.log('     A%s: "%s"', r.row, r.raw);
      });
    } else {
      sec.predmetny.forEach(function(r){ Logger.log('     A%s: "%s"', r.row, r.raw); });
    }

    // Порівняння каталог ↔ секція "Вчителі-предметники".
    var predNorm = {};
    sec.predmetny.forEach(function(r){ predNorm[r.norm] = r; });
    var ok = [], missing = [], wrongRate = [], escaped = [];
    catItems.forEach(function(a){
      var want = _journalNormName(a.subject + ' ' + a.rate);
      var subjNorm = _journalNormName(a.subject);
      if (predNorm[want]){ ok.push(a.subject + ' ' + a.rate); return; }
      var wr = null;
      sec.predmetny.forEach(function(r){ if (r.norm.indexOf(subjNorm) === 0) wr = r; });
      if (wr){ wrongRate.push(a.subject + ': каталог=' + a.rate + ' / Salary="' + wr.raw + '"'); return; }
      var esc = null;
      sec.dodatkovi.forEach(function(r){
        if (r.norm === want || r.norm.indexOf(subjNorm) === 0) esc = r;
      });
      if (esc){ escaped.push(a.subject + ' ' + a.rate + ' → "Додаткові" A' + esc.row + ': "' + esc.raw + '"'); return; }
      missing.push(a.subject + ' ' + a.rate);
    });

    Logger.log('  ✅ Збігається (%s/%s): %s', ok.length, catItems.length, ok.join(', ') || '—');
    if (wrongRate.length) Logger.log('  ⚠️ Інша ставка: %s', wrongRate.join('  |  '));
    if (escaped.length)   Logger.log('  📍 Лежить у "Додаткові": %s', escaped.join('  |  '));
    if (missing.length)   Logger.log('  ❌ Відсутні у Salary: %s', missing.join(', '));

    summary.push(loc + ' | ' + catItems.length + ' | ' + ok.length +
      ' | ' + missing.length + ' | ' + escaped.length);
  });

  Logger.log('\n════════════════ ЗВЕДЕННЯ ════════════════');
  Logger.log('Локація | Каталог | OK | Відсутні | Не там (у Додаткових)');
  summary.forEach(function(s){ Logger.log('  ' + s); });
  return {ok: true};
}

// ── DRY-RUN: матчинг предметників → Salary без запису ──────────────────────
// exportAllPredmetnyToSalary_DRY_RUN(month, year) — проганяє _findPredmetnySalaryRow
// для всіх 11 локацій, НІЧОГО не пише у Salary. Звіт — у Execution log.
function exportAllPredmetnyToSalary_DRY_RUN(month, year){
  month = Number(month); year = Number(year) || new Date().getFullYear();
  if (!month || month < 1 || month > 12){ Logger.log('month має бути 1-12'); return; }
  var LOCS = ['Голосієво','Бігова','Борщагівка','Бровари',"Кар'єрна",'Кругла',
              'Оранж','Осокорки','Позняки','Пуща','Тичини'];
  var ICON = {P1:'✅', P2:'🔵', P3:'🟡', P4:'🟢', P5:'❌'};

  var reg = _salaryGetRegistry();
  if (!reg.ok){ Logger.log('Salary-реєстр: %s', reg.error); return; }
  var regByLoc = {};
  reg.rows.forEach(function(r){ regByLoc[r.loc] = r; });

  var mm = month < 10 ? '0' + month : String(month);
  var dateFrom = year + '-' + mm + '-01';
  var nextM = _nextMonth(month, year);
  var nmm = nextM.month < 10 ? '0' + nextM.month : String(nextM.month);
  var dateTo = nextM.year + '-' + nmm + '-01';
  var attData = _getPredmetnyAttSheet(true).getDataRange().getValues();

  var summary = [], doubtful = [], gateLines = [];
  Logger.log('═══════════ DRY-RUN: предметники → Salary | місяць %s/%s ═══════════', month, year);

  LOCS.forEach(function(loc){
    Logger.log('\n[ЛОКАЦІЯ %s]', loc);
    var catItems = ((getPredmetnyCatalog(loc) || {}).items || [])
      .filter(function(a){ return a.active && a.rate > 0; });

    // Додаткові_Каталог локації — для P4-гейту.
    var addItems = ((getActivitiesCatalog(loc) || {}).items || []);
    var addNames = addItems.map(function(x){ return x.name; });
    var addSubjNorms = {};
    addItems.forEach(function(x){
      var nn = _journalNormName(x.name); if (nn) addSubjNorms[nn] = true;
    });

    var entry = regByLoc[loc];
    if (!entry){
      Logger.log('  ❌ локацію не знайдено у Salary-реєстрі');
      summary.push(loc + ' | ' + catItems.length + ' | — нема в реєстрі —');
      return;
    }
    var sheet;
    try { sheet = SpreadsheetApp.openById(entry.sheetId).getSheetByName(entry.listName); }
    catch(e){
      Logger.log('  ❌ доступ до Salary: %s', e && e.message || e);
      summary.push(loc + ' | ' + catItems.length + ' | — помилка доступу —'); return;
    }
    if (!sheet){
      Logger.log('  ❌ лист "%s" не знайдено', entry.listName);
      summary.push(loc + ' | ' + catItems.length + ' | — нема листа —'); return;
    }

    var lastRow = Math.max(sheet.getLastRow(), 80);
    var names = sheet.getRange(1, 1, lastRow, 1).getValues();
    var salaryRows = [];
    for (var k = 3; k < names.length; k++){
      var raw = String(names[k][0] == null ? '' : names[k][0]).trim();
      if (raw) salaryRows.push({row: k + 1, raw: raw, norm: _journalNormName(raw)});
    }

    var lessonsBySubj = {};
    for (var i = 1; i < attData.length; i++){
      var rec = _parsePredmetnyAttRow(attData[i]);
      if (rec.loc !== loc) continue;
      if (rec.date < dateFrom || rec.date >= dateTo) continue;
      if (!lessonsBySubj[rec.subjectId]) lessonsBySubj[rec.subjectId] = {};
      lessonsBySubj[rec.subjectId][rec.group + '|' + rec.date] = true;
    }

    var st = {p1:0,p2:0,p3:0,p4:0,p5:0}, sumFact = 0;
    catItems.forEach(function(a){
      var uniq = lessonsBySubj[a.id] ? Object.keys(lessonsBySubj[a.id]).length : 0;
      var fact = uniq * a.rate;
      var catName = a.subject + ' ' + a.rate;
      var found = _findPredmetnySalaryRow(salaryRows, a.subject, a.rate, addSubjNorms);
      if (!found){
        st.p5++;
        Logger.log('  %s P5: %s → НЕ ЗНАЙДЕНО (fact=%s₴)', ICON.P5, catName, fact);
        doubtful.push('❌ [' + loc + '] ' + catName + ' → P5 НЕ ЗНАЙДЕНО · fact=' + fact + '₴ (нікуди записати)');
        return;
      }
      st['p' + found.priority.slice(1)]++;
      sumFact += fact;
      var note = found.priority === 'P3' ? ' (інша ставка — найближча)'
               : found.priority === 'P4' ? ' (без ставки — P4-гейт пройдено)'
               : found.priority === 'P2' ? ' (префікс)' : '';
      Logger.log('  %s %s: %s → A%s "%s"%s', ICON[found.priority], found.priority,
        catName, found.row, found.matchedAs, note);
      if (found.priority === 'P3' || found.priority === 'P4'){
        doubtful.push((found.priority === 'P3' ? '🟡' : '🟢') + ' [' + loc + '] ' +
          catName + ' → ' + found.priority + ' A' + found.row + ' "' + found.matchedAs +
          '" · fact=' + fact + '₴');
      }
    });

    Logger.log('  ── каталог=%s | P1=%s P2=%s P3=%s P4=%s P5=%s | сума fact=%s₴',
      catItems.length, st.p1, st.p2, st.p3, st.p4, st.p5, sumFact);
    summary.push(loc + ' | ' + catItems.length + ' | ' + st.p1 + ' | ' + st.p2 +
      ' | ' + st.p3 + ' | ' + st.p4 + ' | ' + st.p5 + ' | ' + sumFact);

    var blocked = [], allowed = [];
    catItems.forEach(function(a){
      if (addSubjNorms[_journalNormName(a.subject)]) blocked.push(a.subject);
      else allowed.push(a.subject);
    });
    gateLines.push('[' + loc + '] Додаткові_Каталог: ' + JSON.stringify(addNames));
    gateLines.push('   → P4 ЗАБЛОКОВАНО: ' + (blocked.join(', ') || '—') +
      '  |  P4 дозволено: ' + (allowed.join(', ') || '—'));
  });

  Logger.log('\n═══════════ ЗВЕДЕНА ТАБЛИЦЯ ═══════════');
  Logger.log('Локація | Каталог | P1 | P2 | P3 | P4 | P5 | Сума fact (грн)');
  summary.forEach(function(s){ Logger.log('  ' + s); });

  Logger.log('\n═══════════ СУМНІВНІ МАТЧІ — перевірити вручну (P3/P4/P5) ═══════════');
  if (!doubtful.length) Logger.log('  (немає — усі матчі точні P1/P2)');
  doubtful.forEach(function(d){ Logger.log('  ' + d); });

  Logger.log('\n═══════════ P4-ГЕЙТ: Додаткові_Каталог по локаціях ═══════════');
  gateLines.forEach(function(g){ Logger.log('  ' + g); });

  return {ok: true};
}
