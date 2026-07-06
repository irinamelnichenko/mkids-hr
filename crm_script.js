// ═══════════════════════════════════════════════════════════════════════════
// m.kids CRM — Google Apps Script v6.11
// v6.11: вчителі-предметники v2 — норми по регіонах × предмет × group_type,
//        журнал занять (Predmetnyky_Lessons), endpoints + тести
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

// БОЙОВИЙ експорт предметників → Salary за травень 2026 (реально пише
// у Salary-файли). Запускати ВРУЧНУ після перевірки dry-run.
function runRealExportMay2026() {
  return exportAllPredmetnyToSalary(5, 2026);
}

var CONFIG_SHEET_ID  = '11NEIEBzaMiIDFnJB9RXqKnRqjCJjNyHVqylrX7cRZhc';
var SHEET_PAYMENTS   = 'Оплати';
var SHEET_YEARLY     = 'Оплати-Рік';
var SHEET_CLIENTS    = 'Клієнти';
var SHEET_ATTENDANCE = 'Табель';
var SHEET_HEALTH     = 'Здоров\'я';

// HR-таблиця (картка співробітника, v6.9). Окремий Google Sheet.
// Структура: A=dir B=typ C=loc D=grp E=last F=first G=phone H=pos I=stat
//            J=bday K=bmon L=bdate M=wday N=hired O=fired
//            P=formula DATEDIF(M;O) — не чіпати при update
//            Q=reserved (порожнє) — не чіпати при update
//            R=email
var HR_SHEET_ID       = '1KeSelcGyp8ijUQOmzUjFlX78TOScMWAFm7Hjf33Doyo';
var HR_TAB_NAME       = '2025-2026';
// v6.44: A:O(1-15) дані, P(16)=формула life-cycle, Q(17)=reserved, R(18)=email,
//   S(19)=Паспорт, T(20)=Ставка ЗП, U(21)=Умови роботи, V(22)=Модель розрахунку.
//   W(23)=Оцінка (JSON), X(24)=Дата виходу з декрету.
//   Append-only: існуючі колонки не зсуваються.
var HR_COLS           = 24;   // v6.59: +X='Дата виходу з декрету' (планована дата повернення)
var HR_AUDIT_SHEET    = 'HR_Audit';
var HR_AUDIT_HEADER   = ['ts','actorId','actorName','action','rowNum','before_json','after_json'];

// Усе lowercase: helper-функції нижче порівнюють case-insensitive
// (бо у Користувачі ролі написані змішано: 'CFO', 'CEO', 'RnD_director',
// 'HR_trainer', 'Legal', а director / nurse / vyhovatel — lowercase).
var EMP_MGMT_ROLES    = ['cfo','ceo','cco','coo','cmo','hr','hr_trainer','rnd_director','legal'];
var EMP_DIR_ROLES     = ['director'];
var EMP_VIEW_ROLES    = ['nurse','vyhovatel'];
var EMP_DELETE_ROLES  = ['cfo','ceo','coo','hr','hr_trainer'];

var MONTHS_UA      = ['вересень','жовтень','листопад','грудень','січень','лютий','березень','квітень','травень','червень','липень','серпень'];
var MONTHS_JS      = [8,9,10,11,0,1,2,3,4,5,6,7];
var MONTHS_DISPLAY = ['Вересень','Жовтень','Листопад','Грудень','Січень','Лютий','Березень','Квітень','Травень','Червень','Липень','Серпень'];
var MONTHS_CAL = ['Січень','Лютий','Березень','Квітень','Травень','Червень','Липень','Серпень','Вересень','Жовтень','Листопад','Грудень'];

// ═══ ЄДИНИЙ ЗОННИЙ ПОРЯДОК ЛОКАЦІЙ (single source of truth) — v6.50.4 ═══
// Дзеркало LOCATION_ORDER у фронт-файлах. CFO-порядок:
//   садочки → школи → управління → онлайн-школа → кухні.
// Невідомі назви (поза списком) → у кінець (rank 999), алфавітно.
var LOCATION_ORDER = [
  'Осокорки','Позняки','Тичини',"Кар'єрна",'Голосієво','Пуща',
  'Оранж','Борщагівка','Бровари','Кругла','Бігова',
  'Школа Осокорки','Школа 228','Житомир',
  'Нац.Гвардії (Благо)','Манхетен (Благо)',
  'Онлайн школа','Кухня Київ','Кухня Львів','Іва-Франківськ кухня'
];
function locationRank(name){
  var i = LOCATION_ORDER.indexOf(String(name == null ? '' : name).trim());
  return i === -1 ? 999 : i;
}
function sortByLocationOrder(arr, getName){
  return (arr || []).slice().sort(function(a, b){
    var na = getName ? getName(a) : a, nb = getName ? getName(b) : b;
    var ra = locationRank(na), rb = locationRank(nb);
    return ra - rb || String(na == null ? '' : na).localeCompare(String(nb == null ? '' : nb), 'uk');
  });
}

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
    else if (action === 'getLocationCards')    result = getLocationCards();
    else if (action === 'getLocationCapacity') result = getLocationCapacity();
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
    else if (action === 'getDopMerges')               result = getDopMerges(e.parameter || {});
    else if (action === 'getPredMerges')              result = getPredMerges(e.parameter || {});
    else if (action === 'backupClients')              result = backupClientsAbsences();
    else if (action === 'mergeSplitVacationRows')     result = mergeSplitVacationRows(!(e.parameter && (e.parameter.dryRun === '0' || e.parameter.dryRun === 'false')));
    else if (action === 'getChomusykyMarks')          result = getChomusykyMarks(e.parameter || {});
    else if (action === 'getChomusykyReport')         result = getChomusykyReport(e.parameter || {});
    else if (action === 'getPredmetnyCatalog')        result = getPredmetnyCatalog(e.parameter && e.parameter.loc || '');
    else if (action === 'getPredmetnyMarks')          result = getPredmetnyMarks(e.parameter || {});
    else if (action === 'getTasks')                   result = getTasks(e.parameter || {});
    else if (action === 'getTaskActivity')            result = getTaskActivity(e.parameter && e.parameter.taskId || 0);
    else if (action === 'getDashboardNotifications')  result = getDashboardNotifications(e.parameter && e.parameter.userId || 0, e.parameter && e.parameter.role || '');
    else if (action === 'getEmployees')               result = getEmployees(Number(e.parameter && e.parameter.actorId || 0), e.parameter && e.parameter.loc || '');
    else if (action === 'getPredmetnyky')              result = getPredmetnyky(Number(e.parameter && e.parameter.actorId || 0));
    else if (action === 'getInvoiceListData')          result = getInvoiceListData(e.parameter || {});
    else if (action === 'generateInvoicePDF')          result = generateInvoicePDF(e.parameter || {}); // v6.50
    else if (action === 'getInvoiceStatusReport')      result = getInvoiceStatusReport(e.parameter || {}); // v6.50.3
    else if (action === 'getFillStatus')               result = getFillStatus(e.parameter || {});           // v6.51
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
    else if (body.action === 'exportVacationDiscountToPayments') result = exportVacationDiscountToPayments(body || {});
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
    else if (body.action === 'seedActivityTeachersInHR')   result = _seedActivityTeachersInHR(Number(body.actorId || 1));
    else if (body.action === 'addAttendanceMark')         result = addAttendanceMark(body.data || {});
    else if (body.action === 'removeAttendanceMark')      result = removeAttendanceMark(body.id || 0);
    else if (body.action === 'bulkAttendanceMarks')       result = bulkAttendanceMarks(body || {});
    else if (body.action === 'bulkRemoveAttendanceMarks') result = bulkRemoveAttendanceMarks(body || {});
    else if (body.action === 'exportAttendanceToPayments') result = exportAttendanceToPayments(body || {});
    else if (body.action === 'reconcilePreview')            result = reconcilePreview(body || {}); // v6.51 ФАЗА 1
    else if (body.action === 'reconcileApply')              result = reconcileApply(body || {});   // v6.51.3 ФАЗА 2
    else if (body.action === 'exportToSalaryExtras')      result = exportToSalaryExtras(body || {});
    else if (body.action === 'saveDopMerge')              result = saveDopMerge(body || {});
    else if (body.action === 'deleteDopMerge')            result = deleteDopMerge(body || {});
    else if (body.action === 'savePredMerge')             result = savePredMerge(body || {});
    else if (body.action === 'deletePredMerge')           result = deletePredMerge(body || {});
    else if (body.action === 'addChomusykyMark')          result = addChomusykyMark(body.data || body || {});
    else if (body.action === 'removeChomusykyMark')       result = removeChomusykyMark(body || {});
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
    else if (body.action === 'saveEmployee')              result = saveEmployee(Number(body.actorId || 0), body.payload || {}, body.rowNum || null);
    else if (body.action === 'saveLocationCard')          result = saveLocationCard(Number(body.actorId || 0), body.payload || {});
    else if (body.action === 'deleteEmployee')            result = deleteEmployee(Number(body.actorId || 0), body.rowNum || 0);
    else if (body.action === 'savePredmetnykyLesson')     result = savePredmetnykyLesson(Number(body.actorId || 0), body.lesson || {});
    else if (body.action === 'deletePredmetnykyLesson')   result = deletePredmetnykyLesson(Number(body.actorId || 0), Number(body.lessonId || 0));
    else if (body.action === 'savePredmetnykyAssignment')   result = savePredmetnykyAssignment(Number(body.actorId || 0), body.payload || body.data || {});
    else if (body.action === 'deletePredmetnykyAssignment') result = deletePredmetnykyAssignment(Number(body.actorId || 0), Number(body.id || 0));
    else if (body.action === 'runPredmetnykyHrSeed')        result = _seedPredmetnykyAssignmentsFromHR();
    else if (body.action === 'clearAllPredmetnykyLessons')  result = clearAllPredmetnykyLessons(Number(body.actorId || 0), body.location || body.loc || '');
    else if (body.action === 'exportPredmetnykyToSalary')   result = exportPredmetnykyToSalary(body || {});
    else if (body.action === 'generateInvoicePDF')          result = generateInvoicePDF(body || {});   // v6.50
    else if (body.action === 'invoicePdfLink')              result = invoicePdfLink(body || {});       // v6.72 Viber link
    else if (body.action === 'invoiceViberMessage')         result = invoiceViberMessage(body || {});   // v6.72 Viber text
    else if (body.action === 'sendInvoiceEmail')            result = sendInvoiceEmail(body || {});      // v6.50
    else if (body.action === 'bulkSendInvoices')            result = bulkSendInvoices(body || {});      // v6.50
    else if (body.action === 'logViberSent')                result = logViberSent(body || {});          // v6.50.3
    else if (body.action === 'getInvoiceStatusReport')      result = getInvoiceStatusReport(body || {}); // v6.50.3
    else if (body.action === 'refreshYearlyAggregate')      result = refreshYearlyAggregate();           // лише агрегат (швидко)
    else if (body.action === 'reexportLocation')            result = reexportLocationFull(body.loc || '', body.month, body.year); // кнопка "Перерахувати локацію": Payment+Salary+агрегат
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
  locs = sortByLocationOrder(locs, function(l){ return l.loc; });   // v6.50.4: єдиний зонний порядок
  return {ok:true, data:locs};
}

// ─── v6.45: КАРТКА ЛОКАЦІЇ (новий аркуш «Локації» у CONFIG-таблиці) ───────────
// Окремий named-аркуш; getLocations (реєстр, getSheets()[0]) НЕ чіпаємо.
var LOCATIONS_TAB = 'Локації';
// v6.45.4: + 'documents' (JSON: договір/ліцензія/сертифікати/поліси + власні типи).
var LOC_COLS = ['loc','fullName','format','license','legalAddr','factAddr','phone','email',
  'area','rooms','features','playground','shelter','wifi','capacity','docsNote','updatedBy','updatedAt','documents'];

// Idempotent: створює аркуш + гарантує ВСІ заголовки LOC_COLS (відсутні додає в кінець,
// існуючі колонки/дані не зсуває). Тому нова колонка 'documents' з'явиться автоматично.
function ensureLocationsSheet(){
  var ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  var sh = ss.getSheetByName(LOCATIONS_TAB);
  if (!sh){
    sh = ss.insertSheet(LOCATIONS_TAB);
    sh.getRange(1, 1, 1, LOC_COLS.length).setValues([LOC_COLS]);
    sh.setFrozenRows(1);
    return sh;
  }
  var lastCol = sh.getLastColumn();
  var hdr = lastCol > 0 ? sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String) : [];
  var changed = false;
  LOC_COLS.forEach(function(c){ if (hdr.indexOf(c) === -1){ hdr.push(c); changed = true; } });
  if (changed) sh.getRange(1, 1, 1, hdr.length).setValues([hdr]);
  return sh;
}

// Читання всіх карток → {loc: {field:val,...}}
function getLocationCards(){
  var sh = ensureLocationsSheet();
  var data = sh.getDataRange().getValues();
  if (data.length < 2) return {ok:true, data:{}};
  var hdr = data[0].map(String), out = {};
  for (var r = 1; r < data.length; r++){
    var loc = trim(String(data[r][0] || '')); if (!loc) continue;
    var o = {};
    for (var c = 0; c < hdr.length; c++) o[hdr[c]] = String(data[r][c] == null ? '' : data[r][c]);
    out[loc] = o;
  }
  return {ok:true, data:out};
}

// Легка вибірка лише capacity для дашборду → {loc: number}
function getLocationCapacity(){
  var cards = getLocationCards().data, cap = {};
  Object.keys(cards).forEach(function(l){ cap[l] = Number(cards[l].capacity) || 0; });
  return {ok:true, data:cap};
}

// Запис картки (OVERWRITE по loc). Права: mgmt — будь-яка, director — лише своя
// (дзеркало _canEditEmployee).
function saveLocationCard(actorId, payload){
  try {
    var actor = _getActor(actorId);
    var loc = trim(String((payload && payload.loc) || ''));
    if (!loc) return {ok:false, error:'Missing loc', code:'VALIDATION'};
    if (!_canEditEmployee(actor, loc))
      return {ok:false, error:'Permission denied for location "' + loc + '"', code:'PERM_DENIED'};

    var lock = LockService.getDocumentLock();
    if (!lock.tryLock(10000)) return {ok:false, error:'Could not acquire lock', code:'LOCK_TIMEOUT'};
    try {
      var sh = ensureLocationsSheet();
      var data = sh.getDataRange().getValues();
      var hdr = (data[0] || LOC_COLS).map(String);   // header-driven: рядок будуємо за фактичним заголовком
      var locIdx = hdr.indexOf('loc'); if (locIdx < 0) locIdx = 0;
      var now = formatDate(new Date());
      var row = hdr.map(function(k){
        if (k === 'updatedBy') return actor ? (actor.name || '') : '';
        if (k === 'updatedAt') return now;
        return payload[k] != null ? payload[k] : '';
      });
      for (var r = 1; r < data.length; r++){
        if (trim(String(data[r][locIdx] || '')) === loc){
          sh.getRange(r + 1, 1, 1, row.length).setValues([row]);
          return {ok:true, action:'updated', loc:loc};
        }
      }
      sh.appendRow(row);
      return {ok:true, action:'created', loc:loc};
    } finally { lock.releaseLock(); }
  } catch(e){
    return {ok:false, error: e.message || String(e)};
  }
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
    'Створено','Оновлено',
    // v6.26 Етап 1A — поля для PDF-рахунків (email + підписант + дод. договір).
    // Append-only: index 29 ('Створено') не зсувається — saveClient row[29] лишається коректним.
    'Email мами','Email тата','Підписант договору','Номер додаткового договору',
    // v6.46 — здоров'я (JSON): measurements 3 точки навч.року + allergies/chronic/notes.
    // Append-only: index 29 ('Створено') не зсувається.
    'Здоров\'я (JSON)',
    // v6.47 — особистісно-соціальний розвиток (JSON): 46 критеріїв × 3 точки навч.року.
    'Розвиток (JSON)'
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
    data.createdAt||now, now,
    // v6.26 Етап 1A — append-only поля для PDF-рахунків.
    data.momEmail||'', data.dadEmail||'',
    data.signerParent||'mom', data.additionalContractNumber||'',
    // v6.46 — здоров'я (append-only, не чіпає інші поля)
    JSON.stringify(data.health||{}),
    // v6.47 — розвиток (append-only)
    JSON.stringify(data.development||{})
  ];
  // Primary: точний збіг id → оновлюємо той самий рядок (звичайне редагування).
  for (var r = 1; r < vals.length; r++) {
    if (String(vals[r][0]) === String(data.id)) {
      row[29] = vals[r][29] || data.createdAt || now;
      sheet.getRange(r+1, 1, 1, row.length).setValues([row]);
      return {ok:true, action:'updated'};
    }
  }
  // v7.24 АНТИ-ДУБЛЬ: id містить ГРУПУ (c_<ПІБ>_<група>_<локація>) → зміна групи
  // дає НОВИЙ id, і без цієї перевірки saveClient дописав би ДРУГИЙ рядок (дубль,
  // через який договір і відпустки розʼїжджались по різних рядках). Якщо існує
  // рядок ТІЄЇ Ж дитини (ПІБ+локація; дата народження або збігається, або
  // відсутня — щоб не злити тезок) → ОНОВЛЮЄМО його (переносимо на нову групу),
  // а absences ОБʼЄДНУЄМО (union), щоб не втратити відпустки при переводі.
  var _nn = function(s){ return String(s || '').trim().toLowerCase().replace(/\s+/g, ' '); };
  var wantName = _nn(data.name), wantLoc = _nn(data.loc), wantBday = String(data.bday || '').trim();
  var cand = -1, candCount = 0;
  for (var r2 = 1; r2 < vals.length; r2++) {
    if (_nn(vals[r2][1]) !== wantName) continue;
    if (_nn(vals[r2][2]) !== wantLoc)  continue;
    var rowBday = String(vals[r2][5] || '').trim();
    if (wantBday && rowBday && wantBday !== rowBday) continue;   // різні дати народження = тезки, не чіпаємо
    candCount++;
    if (cand < 0) cand = r2;
  }
  if (cand >= 0) {
    if (candCount > 1) Logger.log('[saveClient] АНТИ-ДУБЛЬ: "%s" (%s) має %s рядків-кандидатів — оновлюю перший (row %s)', data.name, data.loc, candCount, cand + 1);
    var existAbs = [];
    try { existAbs = JSON.parse(String(vals[cand][16] || '[]')); } catch(e){}
    var mergedAbs = _mergeAbsencesUnion(existAbs, data.absences || []);
    row[16] = JSON.stringify(mergedAbs);
    row[29] = vals[cand][29] || data.createdAt || now;
    sheet.getRange(cand + 1, 1, 1, row.length).setValues([row]);
    return {ok:true, action:'updated-moved', mergedAbsences: mergedAbs.length};
  }
  sheet.appendRow(row);
  return {ok:true, action:'created'};
}

// Обʼєднання двох списків absences без втрат і без дублів. Ключ: id, інакше
// from|to|type|note (щоб плейсхолдери з from/to=null теж дедуплікувались за текстом).
function _mergeAbsencesUnion(a, b){
  var out = [], seen = {};
  function key(x){
    if (x && x.id) return 'id:' + x.id;
    return 'ft:' + String((x && x.from) || '') + '|' + String((x && x.to) || '') +
           '|'  + String((x && x.type) || '') + '|' + String((x && x.note) || '');
  }
  (a || []).concat(b || []).forEach(function(x){
    if (!x) return;
    var k = key(x);
    if (seen[k]) return;
    seen[k] = true;
    out.push(x);
  });
  return out;
}

// БЕКАП: повна копія листа Клієнти (з усіма absences) у новий лист з таймстампом.
// Не змінює джерело. Запускати ПЕРЕД будь-яким злиттям рядків. Ідемпотентно-безпечно
// (щоразу новий лист). Викликати через ?action=backupClients.
function backupClientsAbsences(){
  try {
    var ss = getCRMSpreadsheet();
    var src = ss.getSheetByName(SHEET_CLIENTS);
    if (!src) return {ok:false, error:'Лист "' + SHEET_CLIENTS + '" не знайдено'};
    var stamp = Utilities.formatDate(new Date(), 'Europe/Kiev', 'yyyy-MM-dd_HH-mm-ss');
    var name  = 'Клієнти_BACKUP_' + stamp;
    var copy  = src.copyTo(ss);
    copy.setName(name);
    ss.setActiveSheet(copy);
    ss.moveActiveSheet(ss.getNumSheets());
    var rows = Math.max(0, src.getLastRow() - 1);
    Logger.log('[backupClientsAbsences] створено "%s", рядків=%s', name, rows);
    return {ok:true, backupSheet:name, rows:rows, spreadsheet:ss.getId()};
  } catch(e){
    return {ok:false, error:String(e && e.message || e)};
  }
}

// ЗЛИТТЯ рядків-дублів після переводу групи: відпустки зі старого рядка (без
// договору) → на новий (з договором), старий рядок видаляється. Тільки 6 названих
// дітей. absences ОБʼЄДНУЮТЬСЯ (union, без втрат). Ідемпотентно (повторний запуск —
// джерел уже нема → noop). Безпека: якщо у дитини ≠1 рядка з договором → ПРОПУСК.
// dryRun=true (за замовч.) — лише звіт, БЕЗ запису. Перед реальним — БЕКАП обовʼязково.
var MERGE_SPLIT_VAC_KIDS = [
  {name:'Горбачевський Дамір', loc:'Бровари'},
  {name:'Попов Назарій',       loc:'Бровари'},
  {name:'Ральков Макар',       loc:'Бровари'},
  {name:'Свердлик Ясміна',     loc:'Бровари'},
  {name:'Таркаєва Аліса',      loc:'Бровари'},
  {name:'Ванієв Дамір',        loc:'Осокорки'}
];

function mergeSplitVacationRows(dryRun){
  dryRun = (dryRun === undefined) ? true : !!dryRun;   // дефолт — безпечний dryRun
  var lock = LockService.getScriptLock();
  try { lock.waitLock(30000); }
  catch(e){ return {ok:false, error:'LOCK_TIMEOUT: ' + (e && e.message || e)}; }
  try {
    var ss = getCRMSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_CLIENTS);
    if (!sheet) return {ok:false, error:'Лист "' + SHEET_CLIENTS + '" не знайдено'};
    var vals = sheet.getDataRange().getValues();
    var hdrs = vals[0].map(String);
    var colName = hdrs.indexOf('ПІБ дитини');        if (colName < 0) colName = 1;
    var colLoc  = hdrs.indexOf('Локація');           if (colLoc  < 0) colLoc  = 2;
    var colGrp  = hdrs.indexOf('Група');             if (colGrp  < 0) colGrp  = 3;
    var colCD   = hdrs.indexOf('Дата договору');     if (colCD   < 0) colCD   = 10;
    var colAbs  = hdrs.indexOf('Відсутності (JSON)'); if (colAbs  < 0) colAbs  = 16;
    var _nn = function(s){ return String(s || '').trim().toLowerCase().replace(/\s+/g, ' '); };
    var _vc = function(arr){ return (arr || []).filter(function(a){ return a && a.type === 'vacation'; }).length; };

    var report = [], toDelete = [], mergedCount = 0, skipped = 0, vacMoved = 0;

    MERGE_SPLIT_VAC_KIDS.forEach(function(kid){
      var wantN = _nn(kid.name), wantL = _nn(kid.loc);
      var rows = [];
      for (var i = 1; i < vals.length; i++){
        if (_nn(vals[i][colName]) !== wantN) continue;
        if (_nn(vals[i][colLoc])  !== wantL) continue;
        var abs = []; try { abs = JSON.parse(String(vals[i][colAbs] || '[]')); } catch(e){}
        rows.push({sheetRow: i + 1, contract: String(vals[i][colCD] || '').trim(), abs: abs, group: String(vals[i][colGrp] || '').trim()});
      }
      var contractRows = rows.filter(function(x){ return x.contract; });
      var sourceRows   = rows.filter(function(x){ return !x.contract && _vc(x.abs) > 0; });
      var e = {name: kid.name, loc: kid.loc, rowsFound: rows.length};

      if (contractRows.length !== 1){
        e.action = 'skip';
        e.reason = contractRows.length === 0 ? 'нема рядка з договором' : (contractRows.length + ' рядків з договором — неоднозначно');
        skipped++; report.push(e); return;
      }
      var tgt = contractRows[0];
      if (sourceRows.length === 0){
        e.action = 'noop'; e.reason = 'джерел нема (вже злито)';
        e.targetGroup = tgt.group; e.vacInTarget = _vc(tgt.abs);
        report.push(e); return;
      }
      var srcAbs = sourceRows.reduce(function(acc, s){ return acc.concat(s.abs); }, []);
      var merged = _mergeAbsencesUnion(tgt.abs, srcAbs);
      e.action = 'merge';
      e.targetGroup = tgt.group; e.targetContract = tgt.contract;
      e.sourceGroups = sourceRows.map(function(s){ return s.group; });
      e.vacBefore = _vc(tgt.abs); e.vacMoved = _vc(srcAbs); e.vacAfter = _vc(merged);
      e.absBefore = tgt.abs.length; e.absAfter = merged.length;
      report.push(e);
      mergedCount++; vacMoved += _vc(srcAbs);

      if (!dryRun){
        sheet.getRange(tgt.sheetRow, colAbs + 1).setValue(JSON.stringify(merged));
        sourceRows.forEach(function(s){ toDelete.push(s.sheetRow); });
      }
    });

    if (!dryRun && toDelete.length){
      // видаляємо ЗНИЗУ ВГОРУ, щоб номери рядків не зсувались
      toDelete.sort(function(a, b){ return b - a; }).forEach(function(rn){ sheet.deleteRow(rn); });
    }
    Logger.log('[mergeSplitVacationRows] dryRun=%s | merge=%s skip=%s | відпусток=%s | рядків видалено=%s',
      dryRun, mergedCount, skipped, vacMoved, dryRun ? 0 : toDelete.length);
    return {ok:true, dryRun:dryRun, merged:mergedCount, skipped:skipped,
            vacMoved:vacMoved, rowsDeleted:(dryRun ? 0 : toDelete.length), report:report};
  } catch(e){
    return {ok:false, error:String(e && e.message || e)};
  } finally {
    try { lock.releaseLock(); } catch(_){}
  }
}

// v6.26 Етап 1A — міграція: додає 4 нові колонки у Клієнти-аркуш якщо відсутні.
// Запускати ВРУЧНУ з Apps Script editor (без routing у doPost — це one-shot).
// Безпечно повторно: якщо колонки вже існують, нічого не змінює.
function migrateClientsAddInvoiceFields(){
  var ss = getCRMSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_CLIENTS);
  if (!sheet){
    Logger.log('[migrateClientsInvoice] ❌ sheet "%s" не знайдено', SHEET_CLIENTS);
    return {ok:false, error:'Sheet not found'};
  }
  var newCols = ['Email мами','Email тата','Підписант договору','Номер додаткового договору'];
  var lastCol = sheet.getLastColumn();
  var hdr = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  var missing = newCols.filter(function(c){ return hdr.indexOf(c) === -1; });
  if (!missing.length){
    Logger.log('[migrateClientsInvoice] ✓ усі %s колонок вже існують — no-op', newCols.length);
    return {ok:true, action:'no-op', alreadyHave: newCols};
  }
  ensureClientsHeader(sheet);
  Logger.log('[migrateClientsInvoice] ✓ додано колонок: %s', JSON.stringify(missing));
  Logger.log('[migrateClientsInvoice] existing rows лишилися — нові колонки порожні для них');
  return {ok:true, action:'added', added: missing};
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

// Перезапуск агрегації Оплати-Рік після дедупу/перерахунку додаткових — обгортка з логом.
// aggregatePaymentsYearly() ІДЕМПОТЕНТНА: clearContents + повний rebuild з per-location
// Payment-файлів (одним викликом по всіх локаціях). Повторний запуск дає той самий
// результат. Запускати ВРУЧНУ з редактора (читає всі файли локацій — довго, не через
// веб-апку, щоб не впертись у ~6-хв ліміт). Після запуску перевір errors:[] — якщо є,
// діти тих локацій випали з агрегату (бо clearContents стер усе).
function reaggregateYearlyAfterDedup(){
  Logger.log('═══ ПЕРЕЗАПУСК АГРЕГАЦІЇ Оплати-Рік ═══');
  var res = aggregatePaymentsYearly();
  if (!res || !res.ok){
    Logger.log('❌ Агрегація НЕ виконана: %s', res && res.error);
    return res || {ok:false, error:'no result'};
  }
  Logger.log('✅ Оплати-Рік перебудовано: рядків=%s | мітка оновлення=%s', res.rows, res.updated);
  if (res.errors && res.errors.length){
    Logger.log('⚠️ ПОМИЛКИ по локаціях (%s) — діти цих локацій могли ВИПАСТИ з агрегату, перевір/перезапусти:', res.errors.length);
    res.errors.forEach(function(e){ Logger.log('   • %s', e); });
  } else {
    Logger.log('✓ Помилок по локаціях немає — усі локації агреговані.');
  }
  return res;
}

// POST-екшен для кнопки "🔄 Оновити дані" на сторінці рахунків (invoices.html).
// Перебудовує Оплати-Рік (aggregatePaymentsYearly) і повертає {ok, rows, errors}.
function refreshYearlyAggregate(){
  var res = aggregatePaymentsYearly();
  return {ok: !!(res && res.ok), rows: (res && res.rows) || 0, errors: (res && res.errors) || []};
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

// v6.30.2: нормалізує клітинку «Дата» (Date-об'єкт АБО рядок) → 'yyyy-MM-dd'.
// Корінь бага: appendRow('2026-06-01') авто-конвертувався Sheets у Date, і
// String(Date)='Mon Jun 01 2026...' ламав строкове порівняння from/to у
// getAttendance (всі рядки відкидались) + матчинг у saveAttendance (дублі).
function _attDateIso(cell, tz) {
  if (cell instanceof Date) {
    if (isNaN(cell.getTime())) return '';
    return Utilities.formatDate(cell, tz, 'yyyy-MM-dd');
  }
  var s = trim(String(cell == null ? '' : cell));
  if (!s) return '';
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;                      // вже ISO
  var m = /^(\d{1,2})\.(\d{1,2})\.(\d{4})$/.exec(s);                // DD.MM.YYYY
  if (m) return m[3] + '-' + ('0'+m[2]).slice(-2) + '-' + ('0'+m[1]).slice(-2);
  var d = new Date(s);                                              // довгий date-string
  if (!isNaN(d.getTime())) return Utilities.formatDate(d, tz, 'yyyy-MM-dd');
  return s;
}

function getAttendance(e) {
  var params  = e ? (e.parameter || {}) : {};
  var loc     = trim(params.loc  || '');
  var from    = trim(params.from || '');
  var to      = trim(params.to   || '');
  var ss      = getCRMSpreadsheet();
  var tz      = ss.getSpreadsheetTimeZone() || 'Europe/Kiev';
  var sheet   = ss.getSheetByName(SHEET_ATTENDANCE);
  if (!sheet) return {ok:true, data:[]};
  var vals = sheet.getDataRange().getValues();
  if (vals.length < 2) return {ok:true, data:[]};
  var hdrs = vals[0].map(String);
  var iDate = hdrs.indexOf('Дата');     if (iDate < 0) iDate = 0;
  var iCid  = hdrs.indexOf('ID дитини'); if (iCid  < 0) iCid  = 1;
  var iLoc  = hdrs.indexOf('Локація');   if (iLoc  < 0) iLoc  = 3;

  // dedupe: остання відмітка per (дата, дитина) — як і мерж на фронті.
  var byKey = {}, order = [];
  for (var r = 1; r < vals.length; r++) {
    var iso = _attDateIso(vals[r][iDate], tz);
    if (!iso) continue;
    if (from && iso < from) continue;
    if (to   && iso > to)   continue;
    if (loc  && trim(String(vals[r][iLoc] || '')) !== loc) continue;
    var obj = {};
    for (var c = 0; c < hdrs.length; c++) obj[hdrs[c]] = String(vals[r][c] == null ? '' : vals[r][c]);
    obj['Дата'] = iso;                                              // нормалізована ISO-дата
    var key = iso + '|' + trim(String(vals[r][iCid] || ''));
    if (!(key in byKey)) order.push(key);
    byKey[key] = obj;                                              // last wins
  }
  var rows = order.map(function(k){ return byKey[k]; });
  return {ok:true, data:rows};
}

function saveAttendance(body) {
  var records = body.records || [];
  if (!records.length) return {ok:true, saved:0};
  var ss    = getCRMSpreadsheet();
  var tz    = ss.getSpreadsheetTimeZone() || 'Europe/Kiev';
  var sheet = ss.getSheetByName(SHEET_ATTENDANCE);
  if (!sheet) { sheet = ss.insertSheet(SHEET_ATTENDANCE); writeAttendanceHeader(sheet); }
  var vals = sheet.getDataRange().getValues();
  var now  = formatDate(new Date());
  var saved = 0;

  records.forEach(function(rec) {
    var date    = _attDateIso(rec.date, tz);              // нормалізуємо вхідну дату
    var childId = trim(String(rec.childId || ''));
    if (!date || !childId) return;
    var row = [date, childId, rec.childName||'', rec.loc||'', rec.group||'', rec.status||'', rec.updatedBy||'', now];
    // Оновлюємо ОСТАННІЙ існуючий рядок (read бере last) — матчинг через
    // нормалізовану дату, тож більше НЕ створюємо дублі.
    var lastIdx = -1;
    for (var r = 1; r < vals.length; r++) {
      if (_attDateIso(vals[r][0], tz) === date && trim(String(vals[r][1] || '')) === childId) lastIdx = r;
    }
    if (lastIdx >= 0) {
      sheet.getRange(lastIdx+1, 1, 1, row.length).setValues([row]);
      vals[lastIdx] = row;
    } else {
      sheet.appendRow(row);
      vals.push(row);
    }
    saved++;
    mirrorAttendanceToNurseSheet(rec.loc||'', rec.childName||'', date, rec.status||'');
  });

  return {ok:true, saved:saved};
}

// v6.30.2 editor-only: одноразова дедуплікація аркуша «Табель» — лишає
// ОСТАННІЙ запис per (дата, дитина), нормалізує дати в ISO-текст.
// Запускати з Apps Script editor → Run → dedupAttendance.
function dedupAttendance() {
  var ss    = getCRMSpreadsheet();
  var tz    = ss.getSpreadsheetTimeZone() || 'Europe/Kiev';
  var sheet = ss.getSheetByName(SHEET_ATTENDANCE);
  if (!sheet) return {ok:false, error:'no sheet'};
  var vals = sheet.getDataRange().getValues();
  if (vals.length < 2) return {ok:true, kept:0, removed:0};
  var hdr = vals[0];
  var byKey = {}, order = [];
  for (var r = 1; r < vals.length; r++) {
    var iso = _attDateIso(vals[r][0], tz);
    var cid = trim(String(vals[r][1] || ''));
    if (!iso || !cid) continue;
    var key = iso + '|' + cid;
    var row = vals[r].slice();
    row[0] = iso;                                          // ISO-текст
    if (!(key in byKey)) order.push(key);
    byKey[key] = row;                                      // last wins
  }
  var out = order.map(function(k){ return byKey[k]; });
  sheet.clearContents();
  sheet.getRange(1, 1, 1, hdr.length).setValues([hdr]);
  if (out.length) {
    sheet.getRange(2, 1, out.length, 1).setNumberFormat('@');   // дата як текст → без авто-Date
    sheet.getRange(2, 1, out.length, hdr.length).setValues(out);
  }
  sheet.setFrozenRows(1);
  var res = {ok:true, kept:out.length, removed:(vals.length-1)-out.length};
  Logger.log('[dedupAttendance] ' + JSON.stringify(res));
  return res;
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
      // v7.25: Sheets інколи конвертує текст відпустки у Date-клітинку.
      // День=1 → маркер місяця (синтетичний тиждень місяця). День>1 → конкретна
      // дата → робочий тиждень від неї (5 роб.днів). Раніше day губився (→ MM/YYYY).
      var _dd = str.getDate(), _mm = str.getMonth() + 1, _yy = str.getFullYear();
      if (_dd === 1) return syntheticWeek(_yy, _mm);
      var _d0 = new Date(_yy, _mm - 1, _dd);
      while (_d0.getDay() === 0 || _d0.getDay() === 6) _d0.setDate(_d0.getDate() + 1);
      var _d1 = new Date(_d0); _d1.setDate(_d1.getDate() + 4);
      return {
        from: _d0.getFullYear() + '-' + pad2(_d0.getMonth() + 1) + '-' + pad2(_d0.getDate()),
        to:   _d1.getFullYear() + '-' + pad2(_d1.getMonth() + 1) + '-' + pad2(_d1.getDate()),
        _synthetic: true, _originalRaw: String(str)
      };
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

    // v7.25: діапазон з роком з ОБОХ боків, рік 2- АБО 4-цифровий:
    // "12.05.26-22.05.26", "09.06.26-23.06.26", "01.02.2026-05.02.2026".
    m = n.match(/^(\d{1,2})\.(\d{1,2})\.(\d{2}|\d{4})[-–](\d{1,2})\.(\d{1,2})\.(\d{2}|\d{4})$/);
    if (m) {
      var _y1 = m[3].length === 2 ? 2000 + (+m[3]) : +m[3];
      var _y2 = m[6].length === 2 ? 2000 + (+m[6]) : +m[6];
      return {
        from: _y1 + '-' + pad2(m[2]) + '-' + pad2(m[1]),
        to:   _y2 + '-' + pad2(m[5]) + '-' + pad2(m[4])
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

    // v7.25: запасний регекс на JS-формат дати (клітинка-Date, що вже стала рядком):
    // "Sun Mar 01 2026 10:00:00 GMT+0200 (…)". День=1 → місяць; день>1 → тиждень від дати.
    m = s.match(/^[a-z]{3}\s+([a-z]{3})\s+(\d{1,2})\s+(\d{4})\b/);
    if (m) {
      var EN_MON = {jan:1,feb:2,mar:3,apr:4,may:5,jun:6,jul:7,aug:8,sep:9,oct:10,nov:11,dec:12};
      var monJ = EN_MON[m[1]];
      if (monJ) {
        var dayJ = +m[2], yrJ = +m[3];
        if (dayJ === 1) return syntheticWeek(yrJ, monJ);
        var dJ0 = new Date(yrJ, monJ - 1, dayJ);
        while (dJ0.getDay() === 0 || dJ0.getDay() === 6) dJ0.setDate(dJ0.getDate() + 1);
        var dJ1 = new Date(dJ0); dJ1.setDate(dJ1.getDate() + 4);
        return {
          from: dJ0.getFullYear() + '-' + pad2(dJ0.getMonth() + 1) + '-' + pad2(dJ0.getDate()),
          to:   dJ1.getFullYear() + '-' + pad2(dJ1.getMonth() + 1) + '-' + pad2(dJ1.getDate()),
          _synthetic: true, _originalRaw: str
        };
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
          var rawCell = data[row][absCols[si2]];                       // v7.25: сира клітинка
          var slot = trim(String(rawCell || ''));
          if (!slot) continue;
          totalStats.totalSlotsProcessed++;
          // Date-клітинку передаємо ОБʼЄКТОМ (спрацює Date-гілка), інакше рядком.
          var parsed = parseAbsencePeriod((rawCell instanceof Date) ? rawCell : slot, refYear);
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
          // v7.25: скидаємо старі VACATION-плейсхолдери без дат (from/to=null) — вони
          // не несуть інформації і будуть перепарсені з Payment цим же імпортом
          // (тепер парсер читає Date-формат). Так немає дублю (старий null + новий з
          // датою) і повторний імпорт лишається ідемпотентним.
          var existingAbsences = isNew ? [] : crmMap[crmKey].absences.slice().filter(function(a){
            return !(a && a.type === 'vacation' && !a.from && !a.to);
          });

          var existingPairs = {};
          existingAbsences.forEach(function(a){ if(a.from&&a.to) existingPairs[a.from+'|'+a.to]=true; });

          var newAbsences = [];
          for (var si2 = 0; si2 < absCols.length; si2++) {
            if (absCols[si2] === null) continue;
            var rawCell = data[row][absCols[si2]];                     // v7.25: сира клітинка
            var slot = trim(String(rawCell || ''));
            if (!slot) continue;

            var parsed = parseAbsencePeriod((rawCell instanceof Date) ? rawCell : slot, refYear);
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

  // fallback-джерело дати: картка CRM (Клієнти), якщо в реєстрі порожньо
  var crmBdayByKey = {};
  try {
    var _cliRes = getClients();
    if (_cliRes && _cliRes.ok) {
      (_cliRes.data || []).forEach(function(o){
        var _cn = String(o['ПІБ дитини'] || '').trim();
        var _cloc = String(o['Локація'] || '').trim();
        if (!_cn || !_cloc) return;
        var _b = parseRegistryBday(o['Дата народження']) || '';
        if (_b) crmBdayByKey[_normChildName(_cn) + '|' + _cloc] = _b;
      });
    }
  } catch (e) {}

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
    if (!bdayOut) { var _fb = crmBdayByKey[norm + '|' + loc]; if (_fb) bdayOut = _fb; }
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
  // "День народження" БІЛЬШЕ не вважаємо subtotal-рядком — це звичайний extras
  // (в новій моделі state machine просто falls through у поточний state).
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

  // ── SECTION-BASED класифікація (v6.8) ─────────────────────
  // Збираємо raw-рядки у source-order, ВКЛЮЧАЮЧИ subtotal-рядки
  // ("Додаткові заняття" / "День народження"), бо вони перемикають
  // state у _classifyAllSalaryRows. Фронт отримує rows із полями
  // _section / _category замість того щоб класифікувати самостійно.
  var rawRows = [];
  for (var rowNum = 4; rowNum <= data.length; rowNum++) {
    var idx = rowNum - 1;
    var rowArr = data[idx] || [];
    var rawName = String(rowArr[0] || '').trim();
    if (_salaryIsSkippedRow(rawName)) continue;  // empty / pure-numeric only

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
    rawRows.push({
      row:         rowNum,
      name:        rawName,
      months:      months,
      totalFact:   totalFact,
      totalBudget: totalBudget,
      // для state machine — totals використовуються як індикатор hasNoSum
      fact:        totalFact,
      budget:      totalBudget
    });
  }

  var classified = _classifyAllSalaryRows(rawRows);

  // DEBUG: перші 5 класифікованих рядків.
  for (var d = 0; d < Math.min(5, classified.length); d++){
    Logger.log('[getSalaryData:%s] "%s" → cat=%s sec=%s',
      loc, classified[d].name, classified[d]._category, classified[d]._section);
  }

  // Section-header-рядки — мета, у відповідь не йдуть. Group-header лишається
  // (фронт їх відрисує як підзаголовки груп всередині 'main').
  var rows = classified.filter(function(r){ return r._category !== 'section_header'; });

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

// Чи починається rowNorm з назви каталогу (далі — пробіл/цифра/кінець).
function _startsWithCatalogName(rowNorm, catNorm){
  if (!catNorm || rowNorm.indexOf(catNorm) !== 0) return false;
  if (rowNorm.length === catNorm.length) return true;
  var c = rowNorm.charAt(catNorm.length);
  return c === ' ' || (c >= '0' && c <= '9');
}

// ═══════════════════════════════════════════════════════════════════════════
//  SECTION-BASED CLASSIFIER (v6.8 — гібрид логіки b080103 + 3-digit override)
// ═══════════════════════════════════════════════════════════════════════════
// На відміну від попередньої isolated-per-row класифікації, тут обхід Salary-
// рядків ЗВЕРХУ ВНИЗ зі state machine. State міняють sheet-секції ("Додаткові
// заняття", "День народження"); MAIN-staff keywords та 3-digit override
// перебивають state. Каталоги CONFIG не використовуються — keyword + ставка
// дають надійний матч і працюють для локацій без записів у каталогах
// (Школа Осокорки, Школа 228, Управління-локації).
//
// Вхід: rows = [{name, fact?, budget?, ...}]  у source-order.
// Вихід: новий масив [{...row, _category, _section, _inSchoolGroup?}].
// Рядки "День народження" та empty/numeric-only — повністю пропускаються
// (не з'являються в out, state не міняється).
//
//   _section:  'main' | 'subjects' | 'extras'
//   _category: 'director' | 'teacher' | 'assistant' | 'nurse' | 'guard' |
//              'cleaner' | 'tutor' | 'mentor' | 'duty' | 'meal_extra' |
//              'substitute' | 'subject' | 'school_subject' | 'extras' |
//              'section_header' | 'group_header' | null
function _classifyAllSalaryRows(rows){
  var SUBJECT_KEYWORDS = [
    'англійська','англійський',
    'логопед','муз.керівник','муз керівник',
    'хореограф','фітнес','психолог','спорт',
    'підготовка до школи','чомусики',
    'архітектура','смм','speaking','информатика'
  ];

  function staffCategory(lower){
    if (lower.indexOf('директор')   !== -1)                             return 'director';
    if (lower.indexOf('вихователь') !== -1)                             return 'teacher';
    if (lower.indexOf('вчитель')    !== -1)                             return 'teacher';
    if (lower.indexOf('помічник')   !== -1)                             return 'assistant';
    if (lower.indexOf('медсестра')  !== -1)                             return 'nurse';
    if (lower.indexOf('охорон') !== -1 || lower.indexOf('охран') !== -1) return 'guard';
    if (lower.indexOf('прибиральн') !== -1)                             return 'cleaner';
    if (lower.indexOf('тьютор') !== -1 || lower.indexOf('тімлід') !== -1) return 'tutor';
    if (lower.indexOf('ментор')     !== -1)                             return 'mentor';
    if (lower.indexOf('чергуван')   !== -1)                             return 'duty';
    if (lower.indexOf('млинц')      !== -1)                             return 'meal_extra';
    if (lower.indexOf('замін')      !== -1)                             return 'substitute';
    return null;
  }

  // Стандартний садкочковий предметник — список ключових назв предметів, що
  // зі ставкою вважаються штатним предметництвом ('subject'). Усе інше зі
  // ставкою всередині group_header "Школа" — це шкільні позиції ('school_subject')
  // на кшталт "Фото 200".
  var STANDARD_SADOCHOK_SUBJECTS = [
    'англійська мова', 'логопед', 'муз.керівник', 'муз керівник',
    'хореограф', 'підготовка до школи'
  ];
  function isStandardSadochokSubject(lower){
    for (var i = 0; i < STANDARD_SADOCHOK_SUBJECTS.length; i++){
      if (lower.indexOf(STANDARD_SADOCHOK_SUBJECTS[i]) !== -1) return true;
    }
    return false;
  }

  function subjectKwMatch(lower){
    for (var i = 0; i < SUBJECT_KEYWORDS.length; i++){
      if (lower.indexOf(SUBJECT_KEYWORDS[i]) !== -1) return true;
    }
    return false;
  }

  function emit(r, cat, sec){
    var copy = {};
    for (var k in r) if (r.hasOwnProperty(k)) copy[k] = r[k];
    copy._category = cat;
    copy._section  = sec;
    // Підказка для UI: location-wide штат всередині школи не виноситься в
    // noGroupMain, а лишається під заголовком "Школа".
    if (inSchoolGroup) copy._inSchoolGroup = true;
    return copy;
  }

  var out = [];
  var state = 'main';            // 'main' | 'subjects' | 'extras'
  var inSchoolGroup = false;     // КАР'ЄРНА: всередині group_header "Школа"
                                 // subject keywords дають 'school_subject' у main,
                                 // НЕ переключаючи state на subjects.

  for (var i = 0; i < rows.length; i++){
    var r      = rows[i];
    var name   = String(r.name || '').trim();
    var lower  = name.toLowerCase();
    var budget = Number(r.budget) || 0;
    var fact   = Number(r.fact)   || 0;

    if (!name){ out.push(emit(r, null, state)); continue; }

    // ── 1. Sheet-section headers — міняють state, скидають school-context ─
    //   "школа" БІЛЬШЕ НЕ є section header — це group всередині main (FIX 3).
    if (lower === 'персонал' || lower === 'зарплата'){
      state = 'main';
      inSchoolGroup = false;
      out.push(emit(r, 'section_header', 'main'));
      continue;
    }
    if (lower.indexOf('додаткові заняття') !== -1 || lower === 'додаткові'){
      state = 'extras';
      inSchoolGroup = false;
      out.push(emit(r, 'section_header', 'extras'));
      continue;
    }
    // ── 2. Group-headers всередині main (Findики/.../Школа) ─────
    //   "Школа" — group у Кар'єрній (FIX 3). Інші — за префіксом.
    //   "mini baby-ki" / "mini-baby" з пробілом/дефісом ловить mini\s*-?\s*baby.
    var hasNoSum   = budget === 0 && fact === 0;
    var isSchoolGp = /^школа$/i.test(name);
    var isGroupNm  = isSchoolGp ||
                     /^(find|baby|mini\s*-?\s*baby|minibaby|preschool|study|стаді|студі)/i.test(name);
    if (state === 'main' && hasNoSum && isGroupNm){
      inSchoolGroup = isSchoolGp;   // ON для "Школа", OFF для будь-якої іншої групи
      out.push(emit(r, 'group_header', 'main'));
      continue;
    }

    // ── 3. Main-staff keywords — ЗАВЖДИ 'main'. НЕ скидає inSchoolGroup ─
    //   У КАР'ЄРНІЙ всередині group "Школа" є СВОЯ прибиральниця/медсестра/
    //   охорона тощо. Тому location-wide keywords не виводять зі school-
    //   контексту. _inSchoolGroup-флаг (через emit) допомагає UI рендерити
    //   їх під заголовком "Школа", а не у "noGroupMain" bucket.
    var staffCat = staffCategory(lower);
    if (staffCat){
      out.push(emit(r, staffCat, 'main'));
      continue;
    }

    // ── 4. 3-DIGIT OVERRIDE → 'subject' (з нюансом для школи) ────
    //   Ставка типу 250/280/300 у назві → штатний предметник, навіть якщо
    //   state=extras. Це лікує Кар'єрна/Борщагівка/Бровари (де "Логопед 250"
    //   фізично під "Додатковими заняттями").
    //
    //   Всередині group "Школа" ставка переводить у 'subject' (вихід зі
    //   школи) ТІЛЬКИ якщо назва містить стандартний садочковий предметник
    //   ('Англійська мова 280', 'Логопед 350' тощо). Інакше — це шкільна
    //   стаття ("Фото 200"): cat='school_subject', section='main', школа НЕ
    //   виходить.
    if (/\b[1-9]\d{2}\b/.test(name)){
      if (inSchoolGroup && !isStandardSadochokSubject(lower)){
        out.push(emit(r, 'school_subject', 'main'));
        continue;
      }
      inSchoolGroup = false;
      out.push(emit(r, 'subject', 'subjects'));
      if (state === 'main') state = 'subjects';
      continue;
    }

    // ── 5. SUBJECT_KEYWORDS у state=main → 'subject' (або 'school_subject') ─
    //   "Логопед" без ставки у main-зоні → 'subject', state→subjects.
    //   Усередині group_header "Школа" — 'school_subject', state ЛИШАЄТЬСЯ main
    //   (шкільні предмети, не садкові предметники).
    //   У state=extras БЕЗ ставки те ж саме слово → 'extras' (разова послуга).
    if (state === 'main' && subjectKwMatch(lower)){
      if (inSchoolGroup){
        out.push(emit(r, 'school_subject', 'main'));
      } else {
        out.push(emit(r, 'subject', 'subjects'));
        state = 'subjects';
      }
      continue;
    }

    // ── 6. Default → поточний state ─────────────────────────────
    if (state === 'extras')        out.push(emit(r, 'extras',  'extras'));
    else if (state === 'subjects') out.push(emit(r, 'subject', 'subjects'));
    else                           out.push(emit(r, null,      'main'));
  }

  return out;
}

// ═══════════════════════════════════════════════════════════════════════════
// DEPRECATED — попередній (catalog-driven, isolated-per-row) класифікатор.
// Залишено для historical reference; не використовується. Видалити в наступній
// чистці. Замість нього — _classifyAllSalaryRows (section-based state machine).
// ═══════════════════════════════════════════════════════════════════════════
function _classifySalaryRowByCatalog_OLD(name, predmetnySubjects, addActivities){
  var norm = _softNorm(name);
  if (!norm) return 'main';
  if (/директор|вихователь|медсестра|охрана|охорона|чергування|підміна|прибиральниц|кухар|повар|кухня|техперсонал|психолог/.test(norm))
    return 'main';
  var hasRate = /\b[1-9]\d{2}\b/.test(name);
  if (hasRate && predmetnySubjects){
    for (var i = 0; i < predmetnySubjects.length; i++){
      if (_startsWithCatalogName(norm, _softNorm(predmetnySubjects[i]))) return 'subjects';
    }
  }
  if (addActivities){
    for (var j = 0; j < addActivities.length; j++){
      if (_startsWithCatalogName(norm, _softNorm(addActivities[j]))) return 'extras';
    }
  }
  if (hasRate) return 'subjects';
  return 'main';
}

// DEPRECATED — гранулярний тип штату ізольовано. Зараз цю логіку містить
// staffCategory() у _classifyAllSalaryRows. Залишено для historical reference.
function _ovaGranularStaff_OLD(name){
  var lower = String(name || '').toLowerCase();
  if (lower.indexOf('директор')   !== -1)                             return 'director';
  if (lower.indexOf('вихователь') !== -1)                             return 'teacher';
  if (lower.indexOf('помічник')   !== -1)                             return 'assistant';
  if (lower.indexOf('медсестра')  !== -1)                             return 'nurse';
  if (lower.indexOf('охорон')     !== -1 || lower.indexOf('охран') !== -1) return 'guard';
  if (lower.indexOf('прибиральн') !== -1)                             return 'cleaner';
  if (lower.indexOf('тьютор')     !== -1 || lower.indexOf('тімлід') !== -1) return 'tutor';
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

  // Каталоги предметників і додаткових — групуємо по локації одним проходом
  // (щоб не відкривати CONFIG-таблицю на кожну з 11 локацій).
  var predByLoc = {}, actByLoc = {};
  try {
    var pAll = (getPredmetnyCatalog('').items || []);
    pAll.forEach(function(x){
      (predByLoc[x.loc] = predByLoc[x.loc] || []).push(x.subject);
    });
  } catch(e){}
  try {
    var aAll = (getActivitiesCatalog('').items || []);
    aAll.forEach(function(x){
      (actByLoc[x.loc] = actByLoc[x.loc] || []).push(x.name);
    });
  } catch(e){}

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

          // ── SECTION-BASED класифікація (v6.8) ─────────────────────
          // Збираємо raw-рядки у source-order, ВКЛЮЧАЮЧИ subtotal-рядки
          // ("Додаткові заняття" / "День народження"), бо саме вони
          // перемикають state machine у _classifyAllSalaryRows.
          var rawRows = [];
          for (var rowNum = 4; rowNum <= salData.length; rowNum++) {
            var idx = rowNum - 1;
            var rowArr = salData[idx] || [];
            var rawName = String(rowArr[0] || '').trim();
            if (_salaryIsSkippedRow(rawName)) continue;  // empty / pure-numeric
            var rFact   = fIdx < salWidth ? _opexNum(rowArr[fIdx]) : 0;
            var rBudget = bIdx < salWidth ? _opexNum(rowArr[bIdx]) : 0;
            rawRows.push({name: rawName, fact: rFact, budget: rBudget});
          }

          var classified = _classifyAllSalaryRows(rawRows);

          // DEBUG: перші 5 класифікованих рядків.
          for (var d = 0; d < Math.min(5, classified.length); d++){
            Logger.log('[ova:%s] "%s" → cat=%s sec=%s',
              loc, classified[d].name, classified[d]._category, classified[d]._section);
          }

          classified.forEach(function(cr){
            // Meta-рядки не рахуємо ні в staffCounts, ні в salary-сумах
            // (вони сервісні маркери секцій / груп, без власних грошей).
            if (cr._category === 'section_header') return;
            if (cr._category === 'group_header')   return;

            // staffCounts — бамп гранулярного типу.
            if (cr._category && entry.staffCounts.hasOwnProperty(cr._category)){
              entry.staffCounts[cr._category]++;
            }

            entry.salaryFact   += Number(cr.fact)   || 0;
            entry.salaryBudget += Number(cr.budget) || 0;
          });

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
  'Ставка_викладача','Викладач','Активне','Розрахунок'
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
                  /^(true|так|y|1|active|активне|✅)$/i.test(String(row[7] || '').trim()),
    payType:      String(row[8] || '').trim()
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
      data.active !== false,
      String(data.payType || '').trim()
    ];
    if (!row[1]) return {ok: false, error: 'Поле "Локація" обовʼязкове'};
    if (!row[2]) return {ok: false, error: 'Поле "Назва заняття" обовʼязкове'};
    var _tz = sh.getParent().getSpreadsheetTimeZone() || 'Europe/Kiev';
    var _key = _attDupKey(row[1], row[4], row[5], _tz);
    var _ex = sh.getDataRange().getValues();
    for (var _i = 1; _i < _ex.length; _i++){
      if (_attDupKey(_ex[_i][1], _ex[_i][4], _ex[_i][5], _tz) === _key){
        return {ok: true, id: _ex[_i][0], dup: true};
      }
    }
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
      if ('payType'      in data) sh.getRange(r1, 9).setValue(String(data.payType || '').trim());
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

// ───────────────────────────────────────────────────────────────────────────
// v6.21 ONE-TIME MIGRATION: додає колонку "Розрахунок" (I) до Додаткові_Каталог.
// Запускати ВРУЧНУ з Apps Script editor після оновлення коду, ПЕРЕД новим
// Deploy. Безпечно повторювати — якщо колонка вже існує, нічого не робить.
// ───────────────────────────────────────────────────────────────────────────
function migrateActivitiesAddPayType(){
  try {
    var sh = _getActivitiesSheet(false);
    var lastCol = sh.getLastColumn();
    var header = sh.getRange(1, 1, 1, Math.max(lastCol, 9)).getValues()[0];
    if (String(header[8] || '').trim() === 'Розрахунок'){
      Logger.log('[migrate] Колонка "Розрахунок" вже існує — skip');
      return {ok: true, alreadyMigrated: true};
    }
    sh.getRange(1, 9).setValue('Розрахунок');
    Logger.log('[migrate] ✅ Додано колонку "Розрахунок" в I1');
    return {ok: true, migrated: true};
  } catch(e){
    Logger.log('[migrate] ❌ ' + (e && e.message || e));
    return {ok: false, error: String(e && e.message || e)};
  }
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

// ── v6.13: попередній матчинг catalog.teacher → HR.row ───────────────────
// Для кожного catalog entry з заповненим полем "Викладач" перевіряє чи в HR
// існує співробітник з тим же ПІБ у тій же локації. Якщо НІ — створює його
// в HR через saveEmployee (typ='Викладач додаткових', pos=назва активності,
// phone/email порожні — заповнить director/CFO потім через UI).
// Ідемпотентний — повторні виклики нічого не змінюють.
// Запуск: з Apps Script editor → Run → _seedActivityTeachersInHR
//         або POST {action:'seedActivityTeachersInHR', actorId}.
function _seedActivityTeachersInHR(actorId){
  try {
    actorId = actorId || 1;   // default: CFO Мельніченко Ірина
    var catRes = getActivitiesCatalog('');
    if (!catRes.ok) return catRes;
    var catalog = (catRes.items || []).filter(function(a){
      return a.active && a.teacher && String(a.teacher).trim().length > 0;
    });

    function normName(s){
      return String(s || '').trim().toLowerCase().replace(/\s+/g, ' ');
    }

    // Завантажуємо всі активні HR-рядки
    var hrSh = _getHrSheet();
    var hrLastRow = hrSh.getLastRow();
    var hrTeachers = [];
    if (hrLastRow >= 2){
      var hrData = hrSh.getRange(2, 1, hrLastRow - 1, HR_COLS).getValues();
      for (var i = 0; i < hrData.length; i++){
        var emp = _parseEmpRow(hrData[i], i + 2);
        if (emp.archived) continue;
        if (!emp.last && !emp.first) continue;
        hrTeachers.push(emp);
      }
    }

    // Індексуємо за (loc|normalized_name) — обидва варіанти "last first" та "first last"
    var existsByKey = {};
    hrTeachers.forEach(function(t){
      var n1 = normName(t.last + ' ' + t.first);
      var n2 = normName(t.first + ' ' + t.last);
      existsByKey[t.loc + '|' + n1] = t;
      if (n2 !== n1) existsByKey[t.loc + '|' + n2] = t;
    });

    var stats = {catalog: catalog.length, matched: 0, created: 0, errors: 0, skippedBadName: 0};
    var unmatched = [];

    catalog.forEach(function(a){
      var teacherName = String(a.teacher || '').trim();
      var key = a.loc + '|' + normName(teacherName);
      if (existsByKey[key]){ stats.matched++; return; }

      var parts = teacherName.split(/\s+/);
      if (parts.length < 2){
        stats.skippedBadName++;
        unmatched.push({loc:a.loc, activity:a.name, teacher:teacherName, reason:'name needs ≥2 words'});
        return;
      }
      var last  = parts[0];
      var first = parts.slice(1).join(' ');

      var saveRes = saveEmployee(actorId, {
        last:  last,
        first: first,
        loc:   a.loc,
        pos:   a.name,                  // = назва активності (як у предметників)
        typ:   'Викладач додаткових',
        phone: '',
        email: '',
        hired: _fmtDateDmy(new Date())
      }, null);

      if (saveRes && saveRes.ok){
        stats.created++;
        existsByKey[a.loc + '|' + normName(last + ' ' + first)] = saveRes.employee;
      } else if (saveRes && saveRes.code === 'DUPLICATE'){
        // Дубль за last+first+phone+loc → вважаємо matched
        stats.matched++;
      } else {
        stats.errors++;
        unmatched.push({loc:a.loc, activity:a.name, teacher:teacherName,
                        err: (saveRes && saveRes.error) || 'unknown'});
      }
    });

    Logger.log('[_seedActivityTeachersInHR] catalog=%s matched=%s created=%s errors=%s skippedBadName=%s',
      stats.catalog, stats.matched, stats.created, stats.errors, stats.skippedBadName);
    if (unmatched.length){
      Logger.log('[_seedActivityTeachersInHR] unmatched: %s', JSON.stringify(unmatched).slice(0, 1500));
    }
    return {ok: true, stats: stats, unmatched: unmatched};
  } catch(e){
    return {ok: false, error: String(e && e.message || e)};
  }
}

// Wrapper для прямого запуску з Apps Script editor (без параметрів).
function seedActivityTeachersInHR(){
  return _seedActivityTeachersInHR(1);
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
    var _tz = sh.getParent().getSpreadsheetTimeZone() || 'Europe/Kiev';
    // ФІКС B: фільтр діапазону дат (month+year АБО from/to) — застосовуємо на читанні,
    // ДО повного парсингу рядка. Кратно зменшує payload і CPU для помісячної сітки.
    var monthPrefix = '';
    if (filters.year && filters.month){
      monthPrefix = String(Number(filters.year)) + '-' + ('0' + Number(filters.month)).slice(-2) + '-';
    }
    var fromISO = filters.from ? String(filters.from) : '';
    var toISO   = filters.to   ? String(filters.to)   : '';
    var hasDateRange = !!(monthPrefix || fromISO || toISO);
    var items = [];
    for (var i = 1; i < data.length; i++){
      if (!data[i][0] && !data[i][4]) continue;
      // швидкий фільтр по сирій даті ДО _parseAttendanceRow
      if (hasDateRange){
        var _d = _attDateISO(data[i][1], _tz);
        if (monthPrefix && _d.indexOf(monthPrefix) !== 0) continue;
        if (fromISO && _d < fromISO) continue;
        if (toISO   && _d > toISO)   continue;
      }
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
  // v6.84: O(1) — читаємо лише останній рядок (ID лише дописуються).
  // Раніше читався весь аркуш — тримало замок довго під навантаженням.
  var last = sh.getLastRow();
  if (last < 2) return 1;
  var lastId = Number(sh.getRange(last, 1).getValue()) || 0;
  if (lastId > 0) return lastId + 1;
  var ids = sh.getRange(2, 1, last - 1, 1).getValues();
  var max = 0;
  for (var i = 0; i < ids.length; i++){
    var n = Number(ids[i][0]) || 0;
    if (n > max) max = n;
  }
  return max + 1;
}

// ═══════════════════════════════════════════════════════════════════════════
// ДОДАТКОВІ · ОБʼЄДНАННЯ ГРУП (v7.08)
// Викладач «За заняття»/«За захід» інколи веде кілька груп РАЗОМ одного дня —
// це ОДНЕ заняття, не N. Аркуш "Додаткові_Обʼєднання" зберігає такі схлопування:
//   id | Локація | id_заняття | Назва_заняття | Дата | Групи(JSON) | Ким | Коли
// Групи(JSON) — масив назв груп, що велись разом того дня (сирі назви, як у сітці).
// Один запис = один (loc, id_заняття, дата). У exportToSalaryExtras merge-набір
// схлопує свої групи в 1 сесію (session-key = група×дата).
//
// ⚠️ Нормалізація груп — _dopNormGroup: лише trim+lowercase+collapse-spaces.
// НЕ через normalizeGroupName (та мапить "Preschool 2"→"Preschool" — злила б
// різні групи в одну). "Preschool" і "Preschool 2" мусять лишатись РІЗНИМИ.
// ═══════════════════════════════════════════════════════════════════════════
var DOP_MERGES_SHEET_NAME = 'Додаткові_Обʼєднання';
var DOP_MERGES_HEADER = ['id','Локація','id_заняття','Назва_заняття','Дата','Групи(JSON)','Ким','Коли'];

function _getDopMergesSheet(createIfMissing){
  var ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  var sh = ss.getSheetByName(DOP_MERGES_SHEET_NAME);
  if (!sh && createIfMissing){
    sh = ss.insertSheet(DOP_MERGES_SHEET_NAME);
    sh.getRange(1, 1, 1, DOP_MERGES_HEADER.length).setValues([DOP_MERGES_HEADER]);
    sh.setFrozenRows(1);
  }
  if (!sh) throw new Error('Sheet "' + DOP_MERGES_SHEET_NAME + '" не знайдено. Створіть лист з колонками: ' + DOP_MERGES_HEADER.join(', '));
  return sh;
}

// Легка нормалізація групи для співставлення (зберігає цифру-суфікс!).
function _dopNormGroup(s){
  return String(s || '').trim().replace(/\s+/g, ' ').toLowerCase();
}

function _dopDateISO(v){
  if (v instanceof Date){
    var y = v.getFullYear(), m = v.getMonth() + 1, dd = v.getDate();
    return y + '-' + (m < 10 ? '0' + m : m) + '-' + (dd < 10 ? '0' + dd : dd);
  }
  return String(v || '').trim();
}

function _parseDopMergeGroups(cell){
  var raw = String(cell || '').trim();
  if (!raw) return [];
  try {
    var arr = JSON.parse(raw);
    if (Array.isArray(arr)) return arr.map(function(g){ return String(g || '').trim(); }).filter(Boolean);
  } catch(e){}
  // fallback: comma-separated
  return raw.split(/[;,|]/).map(function(g){ return g.trim(); }).filter(Boolean);
}

function _parseDopMergeRow(row){
  return {
    id:           Number(row[0]) || 0,
    loc:          String(row[1] || '').trim(),
    activityId:   Number(row[2]) || 0,
    activityName: String(row[3] || '').trim(),
    date:         _dopDateISO(row[4]),
    groups:       _parseDopMergeGroups(row[5]),
    by:           String(row[6] || '').trim(),
    at:           row[7] instanceof Date ? row[7].toISOString() : String(row[7] || '')
  };
}

function _nextDopMergeId(sh){
  var last = sh.getLastRow();
  if (last < 2) return 1;
  var lastId = Number(sh.getRange(last, 1).getValue()) || 0;
  if (lastId > 0) return lastId + 1;
  var ids = sh.getRange(2, 1, last - 1, 1).getValues();
  var max = 0;
  for (var i = 0; i < ids.length; i++){
    var n = Number(ids[i][0]) || 0;
    if (n > max) max = n;
  }
  return max + 1;
}

// GET: обʼєднання за фільтром (loc обовʼязковий; опц. activityId / date / month+year / from-to).
function getDopMerges(params){
  try {
    params = params || {};
    var sh;
    try { sh = _getDopMergesSheet(false); }
    catch(e){ return {ok: true, items: []}; }   // листа ще нема → порожньо
    var data = sh.getDataRange().getValues();
    if (data.length < 2) return {ok: true, items: []};

    var fLoc  = String(params.loc || '').trim();
    var fAct  = Number(params.activityId) || 0;
    var fDate = String(params.date || '').trim();
    var monthPrefix = '';
    if (params.year && params.month){
      monthPrefix = String(Number(params.year)) + '-' + ('0' + Number(params.month)).slice(-2) + '-';
    }
    var fromISO = params.from ? String(params.from) : '';
    var toISO   = params.to   ? String(params.to)   : '';

    var items = [];
    for (var i = 1; i < data.length; i++){
      if (!data[i][0] && !data[i][1]) continue;
      var m = _parseDopMergeRow(data[i]);
      if (fLoc  && m.loc !== fLoc) continue;
      if (fAct  && m.activityId !== fAct) continue;
      if (fDate && m.date !== fDate) continue;
      if (monthPrefix && m.date.indexOf(monthPrefix) !== 0) continue;
      if (fromISO && m.date < fromISO) continue;
      if (toISO   && m.date > toISO)   continue;
      if (m.groups.length < 2) continue;   // одна група = не обʼєднання
      items.push(m);
    }
    return {ok: true, items: items};
  } catch(e){
    return {ok: false, error: String(e && e.message || e)};
  }
}

// SAVE (upsert по loc+activityId+date). groups<2 → знімаємо обʼєднання (delete).
function saveDopMerge(body){
  var lock = LockService.getScriptLock();
  try { lock.waitLock(30000); }
  catch(e){ return {ok: false, error: 'LOCK_TIMEOUT: ' + (e && e.message || e)}; }
  try {
    body = body || {};
    var loc   = String(body.loc || '').trim();
    var actId = Number(body.activityId) || 0;
    var date  = String(body.date || '').trim();
    if (!loc || !actId || !date){
      return {ok: false, error: 'Поля loc / activityId / date обовʼязкові'};
    }
    // Унікальні групи (по нормалізованому ключу, зберігаємо перше сире написання).
    var seen = {};
    var groups = [];
    (Array.isArray(body.groups) ? body.groups : []).forEach(function(g){
      var raw = String(g || '').trim();
      if (!raw) return;
      var nk = _dopNormGroup(raw);
      if (seen[nk]) return;
      seen[nk] = true;
      groups.push(raw);
    });

    var sh = _getDopMergesSheet(true);
    var data = sh.getDataRange().getValues();
    var foundRow = -1;
    for (var i = 1; i < data.length; i++){
      var r = _parseDopMergeRow(data[i]);
      if (r.loc === loc && r.activityId === actId && r.date === date){ foundRow = i + 1; break; }
    }

    // <2 груп → це не обʼєднання: якщо був запис — видаляємо (розʼєднання).
    if (groups.length < 2){
      if (foundRow > 0){ sh.deleteRow(foundRow); return {ok: true, removed: true, unmerged: true}; }
      return {ok: true, removed: false, unmerged: true};
    }

    var actName = String(body.activityName || (foundRow > 0 ? _parseDopMergeRow(data[foundRow - 1]).activityName : '')).trim();
    var by      = String(body.by || body.markedBy || '').trim();
    var now     = new Date();

    if (foundRow > 0){
      // upsert: перезаписуємо групи/назву/ким/коли, id лишаємо.
      var id = Number(data[foundRow - 1][0]) || _nextDopMergeId(sh);
      sh.getRange(foundRow, 1, 1, DOP_MERGES_HEADER.length).setValues([[
        id, loc, actId, actName, date, JSON.stringify(groups), by, now
      ]]);
      return {ok: true, id: id, updated: true, groups: groups};
    } else {
      var newId = _nextDopMergeId(sh);
      sh.appendRow([newId, loc, actId, actName, date, JSON.stringify(groups), by, now]);
      return {ok: true, id: newId, created: true, groups: groups};
    }
  } catch(e){
    return {ok: false, error: String(e && e.message || e)};
  } finally {
    try { lock.releaseLock(); } catch(_){}
  }
}

// DELETE: по id АБО по (loc, activityId, date).
function deleteDopMerge(body){
  var lock = LockService.getScriptLock();
  try { lock.waitLock(30000); }
  catch(e){ return {ok: false, error: 'LOCK_TIMEOUT: ' + (e && e.message || e)}; }
  try {
    body = body || {};
    var id    = Number(body.id) || 0;
    var loc   = String(body.loc || '').trim();
    var actId = Number(body.activityId) || 0;
    var date  = String(body.date || '').trim();
    if (!id && !(loc && actId && date)){
      return {ok: false, error: 'Треба id АБО (loc + activityId + date)'};
    }
    var sh;
    try { sh = _getDopMergesSheet(false); }
    catch(e){ return {ok: true, removed: false}; }
    var data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++){
      var r = _parseDopMergeRow(data[i]);
      var hit = id ? (r.id === id) : (r.loc === loc && r.activityId === actId && r.date === date);
      if (hit){ sh.deleteRow(i + 1); return {ok: true, removed: true}; }
    }
    return {ok: true, removed: false};
  } catch(e){
    return {ok: false, error: String(e && e.message || e)};
  } finally {
    try { lock.releaseLock(); } catch(_){}
  }
}

// Завантажує merges локації у діапазоні [dateFrom, dateTo) → мапа для експорту:
//   { activityId: { 'YYYY-MM-DD': [ [normGroup,...], ... ] } }
function _loadDopMergesMap(loc, dateFrom, dateTo){
  var map = {};
  var sh;
  try { sh = _getDopMergesSheet(false); }
  catch(e){ return map; }
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++){
    if (!data[i][1]) continue;
    var m = _parseDopMergeRow(data[i]);
    if (m.loc !== loc) continue;
    if (dateFrom && m.date < dateFrom) continue;
    if (dateTo   && m.date >= dateTo)  continue;
    var normSet = m.groups.map(_dopNormGroup).filter(Boolean);
    if (normSet.length < 2) continue;
    if (!map[m.activityId]) map[m.activityId] = {};
    if (!map[m.activityId][m.date]) map[m.activityId][m.date] = [];
    map[m.activityId][m.date].push(normSet);
  }
  return map;
}

// Рахує кількість СЕСІЙ (session-key = група×дата) з урахуванням обʼєднань.
//   groupsByDate: { 'YYYY-MM-DD': { normGroup: true } }
//   mergesForAct: { 'YYYY-MM-DD': [ [normGroup,...], ... ] }
// Merge-набір, що має ≥1 присутню групу того дня, схлопується в 1 сесію.
function _dopCountSessions(groupsByDate, mergesForAct){
  var total = 0;
  Object.keys(groupsByDate).forEach(function(date){
    var present = groupsByDate[date];            // {normGroup:true}
    var mergeSets = (mergesForAct && mergesForAct[date]) || [];
    var covered = {};
    var sessions = 0;
    mergeSets.forEach(function(ms){
      var any = false;
      ms.forEach(function(g){ if (present[g]){ any = true; covered[g] = true; } });
      if (any) sessions++;
    });
    Object.keys(present).forEach(function(g){ if (!covered[g]) sessions++; });
    total += sessions;
  });
  return total;
}

// ═══════════════════════════════════════════════════════════════════════════
// ПРЕДМЕТНИКИ · ОБʼЄДНАННЯ ГРУП (v7.20) — той самий session-key (група×дата),
// що й у допівців, але ключ обʼєднання = (Локація, Предмет, Дата) замість
// (Локація, id_заняття, Дата). Викладач-предметник інколи веде кілька груп
// РАЗОМ одного дня (напр. Baby-ki + Baby-ki 2 звели в одне заняття) — це ОДНЕ
// заняття у ЗП, не N. Аркуш "Predmetnyky_Обʼєднання":
//   id | Локація | Предмет | Дата | Групи(JSON) | Ким | Коли
// Дата зберігається ISO (YYYY-MM-DD), як у допівців. Нормалізація груп —
// _dopNormGroup (зберігає цифру-суфікс: Baby-ki ≠ Baby-ki 2). Рахунок сесій —
// _dopCountSessions (generic, спільна з допівцями).
// ═══════════════════════════════════════════════════════════════════════════
var PRED_MERGES_SHEET_NAME = 'Predmetnyky_Обʼєднання';
var PRED_MERGES_HEADER = ['id','Локація','Предмет','Дата','Групи(JSON)','Ким','Коли'];

function _getPredMergesSheet(createIfMissing){
  var ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  var sh = ss.getSheetByName(PRED_MERGES_SHEET_NAME);
  if (!sh && createIfMissing){
    sh = ss.insertSheet(PRED_MERGES_SHEET_NAME);
    sh.getRange(1, 1, 1, PRED_MERGES_HEADER.length).setValues([PRED_MERGES_HEADER]);
    sh.setFrozenRows(1);
  }
  if (!sh) throw new Error('Sheet "' + PRED_MERGES_SHEET_NAME + '" не знайдено. Створіть лист з колонками: ' + PRED_MERGES_HEADER.join(', '));
  return sh;
}

function _parsePredMergeRow(row){
  return {
    id:      Number(row[0]) || 0,
    loc:     String(row[1] || '').trim(),
    subject: String(row[2] || '').trim(),
    date:    _dopDateISO(row[3]),                 // спільний хелпер із допівцями
    groups:  _parseDopMergeGroups(row[4]),        // спільний парсер JSON/CSV
    by:      String(row[5] || '').trim(),
    at:      row[6] instanceof Date ? row[6].toISOString() : String(row[6] || '')
  };
}

function _nextPredMergeId(sh){
  var last = sh.getLastRow();
  if (last < 2) return 1;
  var lastId = Number(sh.getRange(last, 1).getValue()) || 0;
  if (lastId > 0) return lastId + 1;
  var ids = sh.getRange(2, 1, last - 1, 1).getValues();
  var max = 0;
  for (var i = 0; i < ids.length; i++){ var n = Number(ids[i][0]) || 0; if (n > max) max = n; }
  return max + 1;
}

// GET: обʼєднання за фільтром (loc обовʼязковий; опц. subject / date / month+year / from-to).
function getPredMerges(params){
  try {
    params = params || {};
    var sh;
    try { sh = _getPredMergesSheet(false); }
    catch(e){ return {ok: true, items: []}; }
    var data = sh.getDataRange().getValues();
    if (data.length < 2) return {ok: true, items: []};

    var fLoc  = String(params.loc || '').trim();
    var fSubj = String(params.subject || '').trim();
    var fDate = String(params.date || '').trim();
    var monthPrefix = '';
    if (params.year && params.month){
      monthPrefix = String(Number(params.year)) + '-' + ('0' + Number(params.month)).slice(-2) + '-';
    }
    var fromISO = params.from ? String(params.from) : '';
    var toISO   = params.to   ? String(params.to)   : '';

    var items = [];
    for (var i = 1; i < data.length; i++){
      if (!data[i][0] && !data[i][1]) continue;
      var m = _parsePredMergeRow(data[i]);
      if (fLoc  && m.loc !== fLoc) continue;
      if (fSubj && m.subject !== fSubj) continue;
      if (fDate && m.date !== fDate) continue;
      if (monthPrefix && m.date.indexOf(monthPrefix) !== 0) continue;
      if (fromISO && m.date < fromISO) continue;
      if (toISO   && m.date > toISO)   continue;
      if (m.groups.length < 2) continue;
      items.push(m);
    }
    return {ok: true, items: items};
  } catch(e){
    return {ok: false, error: String(e && e.message || e)};
  }
}

// SAVE (upsert по loc+subject+date). groups<2 → знімаємо обʼєднання (delete).
function savePredMerge(body){
  var lock = LockService.getScriptLock();
  try { lock.waitLock(30000); }
  catch(e){ return {ok: false, error: 'LOCK_TIMEOUT: ' + (e && e.message || e)}; }
  try {
    body = body || {};
    var loc  = String(body.loc || '').trim();
    var subj = String(body.subject || '').trim();
    var date = String(body.date || '').trim();
    if (!loc || !subj || !date){
      return {ok: false, error: 'Поля loc / subject / date обовʼязкові'};
    }
    var seen = {}, groups = [];
    (Array.isArray(body.groups) ? body.groups : []).forEach(function(g){
      var raw = String(g || '').trim();
      if (!raw) return;
      var nk = _dopNormGroup(raw);
      if (seen[nk]) return;
      seen[nk] = true;
      groups.push(raw);
    });

    var sh = _getPredMergesSheet(true);
    var data = sh.getDataRange().getValues();
    var foundRow = -1;
    for (var i = 1; i < data.length; i++){
      var r = _parsePredMergeRow(data[i]);
      if (r.loc === loc && r.subject === subj && r.date === date){ foundRow = i + 1; break; }
    }

    if (groups.length < 2){
      if (foundRow > 0){ sh.deleteRow(foundRow); return {ok: true, removed: true, unmerged: true}; }
      return {ok: true, removed: false, unmerged: true};
    }

    var by  = String(body.by || body.markedBy || '').trim();
    var now = new Date();
    if (foundRow > 0){
      var id = Number(data[foundRow - 1][0]) || _nextPredMergeId(sh);
      sh.getRange(foundRow, 1, 1, PRED_MERGES_HEADER.length).setValues([[
        id, loc, subj, date, JSON.stringify(groups), by, now
      ]]);
      return {ok: true, id: id, updated: true, groups: groups};
    } else {
      var newId = _nextPredMergeId(sh);
      sh.appendRow([newId, loc, subj, date, JSON.stringify(groups), by, now]);
      return {ok: true, id: newId, created: true, groups: groups};
    }
  } catch(e){
    return {ok: false, error: String(e && e.message || e)};
  } finally {
    try { lock.releaseLock(); } catch(_){}
  }
}

// DELETE: по id АБО по (loc, subject, date).
function deletePredMerge(body){
  var lock = LockService.getScriptLock();
  try { lock.waitLock(30000); }
  catch(e){ return {ok: false, error: 'LOCK_TIMEOUT: ' + (e && e.message || e)}; }
  try {
    body = body || {};
    var id   = Number(body.id) || 0;
    var loc  = String(body.loc || '').trim();
    var subj = String(body.subject || '').trim();
    var date = String(body.date || '').trim();
    if (!id && !(loc && subj && date)){
      return {ok: false, error: 'Треба id АБО (loc + subject + date)'};
    }
    var sh;
    try { sh = _getPredMergesSheet(false); }
    catch(e){ return {ok: true, removed: false}; }
    var data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++){
      var r = _parsePredMergeRow(data[i]);
      var hit = id ? (r.id === id) : (r.loc === loc && r.subject === subj && r.date === date);
      if (hit){ sh.deleteRow(i + 1); return {ok: true, removed: true}; }
    }
    return {ok: true, removed: false};
  } catch(e){
    return {ok: false, error: String(e && e.message || e)};
  } finally {
    try { lock.releaseLock(); } catch(_){}
  }
}

// Завантажує merges локації у [dateFrom, dateTo) → { subjNormKey: { 'YYYY-MM-DD': [ [normGroup,...], ... ] } }
// subjNormKey = _dopNormGroup(subject) — стійке співставлення з catalog.subject_norm/lessons.subject.
function _loadPredMergesMap(loc, dateFrom, dateTo){
  var map = {};
  var sh;
  try { sh = _getPredMergesSheet(false); }
  catch(e){ return map; }
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++){
    if (!data[i][1]) continue;
    var m = _parsePredMergeRow(data[i]);
    if (m.loc !== loc) continue;
    if (dateFrom && m.date < dateFrom) continue;
    if (dateTo   && m.date >= dateTo)  continue;
    var normSet = m.groups.map(_dopNormGroup).filter(Boolean);
    if (normSet.length < 2) continue;
    var sk = _dopNormGroup(m.subject);
    if (!map[sk]) map[sk] = {};
    if (!map[sk][m.date]) map[sk][m.date] = [];
    map[sk][m.date].push(normSet);
  }
  return map;
}

// Плоский список merges для фронту (сітка вантажить обʼєднання одним getPredmetnyky).
// locFilter falsy → всі локації (мережеві ролі); рядок → лише ця локація.
function _loadPredMergesList(locFilter){
  var out = [];
  var sh;
  try { sh = _getPredMergesSheet(false); }
  catch(e){ return out; }
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++){
    if (!data[i][1]) continue;
    var m = _parsePredMergeRow(data[i]);
    if (locFilter && m.loc !== locFilter) continue;
    if (m.groups.length < 2) continue;
    out.push({id: m.id, loc: m.loc, subject: m.subject, date: m.date, groups: m.groups});
  }
  return out;
}

// DMY ("01.06.2026") або Date → ISO ("2026-06-01"). Для співставлення дат уроків з merges.
function _predDateToISO(dateInput){
  if (dateInput instanceof Date) return _dopDateISO(dateInput);
  var s = String(dateInput || '').trim();
  var dmy = /^(\d{1,2})\.(\d{1,2})\.(\d{4})$/.exec(s);
  if (dmy) return dmy[3] + '-' + ('0' + dmy[2]).slice(-2) + '-' + ('0' + dmy[1]).slice(-2);
  var iso = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(s);
  if (iso) return iso[1] + '-' + ('0' + iso[2]).slice(-2) + '-' + ('0' + iso[3]).slice(-2);
  return s;
}

// ═══════════════════════════════════════════════════════════════════════════
// ЧОМУСИКИ · АБОНЕМЕНТ (v7.23) — розвиваючі заняття за абонементом.
// На відміну від решти предметників (група×дата, ЗП викладача, клієнта не білять),
// Чомусики: облік ПО ДІТЯХ (дитина×дата, як у допівців) + клієнт платить ФІКСОВАНУ
// суму за абонемент на N відвідувань/місяць (не за візит).
// Лист "Чомусики_Відвідуваність": id | Дата | Локація | Дитина | Відмітив | Час.
// Один рядок = одна дитина відвідала один день.
//
// ⚠️ ДЕФОЛТИ бізнес-логіки (Іра скоригує — позначено у відповіді):
//   • price = 1800 ₴/міс за абонемент, visits = 8 відвідувань включено.
//   • Білимо ФІКСОВАНО: дитина з ≥1 візитом цього місяця = 1 абонемент (price).
//     (менше 8 → сума та сама; >8 → поки теж та сама, стоп на абонементі).
//   • Клієнтський Payment поки НЕ пишемо автоматично (потрібне підтвердження ціни):
//     getChomusykyReport дає ПРЕВʼЮ суми; запис у Payment — наступний крок.
// ═══════════════════════════════════════════════════════════════════════════
var CHOMUSYKY_ATT_SHEET_NAME = 'Чомусики_Відвідуваність';
var CHOMUSYKY_ATT_HEADER = ['id','Дата','Локація','Дитина','Відмітив','Час'];
var CHOMUSYKY_SUBSCRIPTION = { price: 1800, visits: 8 };   // ДЕФОЛТ — скоригувати

function _getChomusykyAttSheet(createIfMissing){
  var ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  var sh = ss.getSheetByName(CHOMUSYKY_ATT_SHEET_NAME);
  if (!sh && createIfMissing){
    sh = ss.insertSheet(CHOMUSYKY_ATT_SHEET_NAME);
    sh.getRange(1, 1, 1, CHOMUSYKY_ATT_HEADER.length).setValues([CHOMUSYKY_ATT_HEADER]);
    sh.setFrozenRows(1);
  }
  if (!sh) throw new Error('Sheet "' + CHOMUSYKY_ATT_SHEET_NAME + '" не знайдено');
  return sh;
}

function _parseChomusykyRow(row){
  return {
    id:    Number(row[0]) || 0,
    date:  _dopDateISO(row[1]),
    loc:   String(row[2] || '').trim(),
    child: String(row[3] || '').trim(),
    by:    String(row[4] || '').trim(),
    at:    row[5] instanceof Date ? row[5].toISOString() : String(row[5] || '')
  };
}

function _nextChomusykyId(sh){
  var last = sh.getLastRow();
  if (last < 2) return 1;
  var lastId = Number(sh.getRange(last, 1).getValue()) || 0;
  if (lastId > 0) return lastId + 1;
  var ids = sh.getRange(2, 1, last - 1, 1).getValues();
  var max = 0;
  for (var i = 0; i < ids.length; i++){ var n = Number(ids[i][0]) || 0; if (n > max) max = n; }
  return max + 1;
}

// GET: відмітки Чомусиків (loc обовʼязковий; опц. year+month / date / child).
function getChomusykyMarks(filters){
  try {
    filters = filters || {};
    var sh;
    try { sh = _getChomusykyAttSheet(false); }
    catch(e){ return {ok: true, items: [], subscription: CHOMUSYKY_SUBSCRIPTION}; }
    var data = sh.getDataRange().getValues();
    if (data.length < 2) return {ok: true, items: [], subscription: CHOMUSYKY_SUBSCRIPTION};

    var fLoc = String(filters.loc || '').trim();
    var fChild = String(filters.child || '').trim();
    var fDate = String(filters.date || '').trim();
    var monthPrefix = '';
    if (filters.year && filters.month){
      monthPrefix = String(Number(filters.year)) + '-' + ('0' + Number(filters.month)).slice(-2) + '-';
    }
    var items = [];
    for (var i = 1; i < data.length; i++){
      if (!data[i][0] && !data[i][3]) continue;
      var m = _parseChomusykyRow(data[i]);
      if (fLoc && m.loc !== fLoc) continue;
      if (fChild && m.child !== fChild) continue;
      if (fDate && m.date !== fDate) continue;
      if (monthPrefix && m.date.indexOf(monthPrefix) !== 0) continue;
      items.push(m);
    }
    return {ok: true, items: items, subscription: CHOMUSYKY_SUBSCRIPTION};
  } catch(e){
    return {ok: false, error: String(e && e.message || e)};
  }
}

// POST: додати візит (dedup по loc+child+date). Повертає {ok,id} або {ok,dup}.
function addChomusykyMark(data){
  data = data || {};
  var date  = _dopDateISO(String(data.date || '').trim());
  var loc   = String(data.loc || '').trim();
  var child = String(data.child || '').trim();
  if (!date || !loc || !child) return {ok: false, error: 'Поля date / loc / child обовʼязкові'};
  var lock = LockService.getScriptLock();
  try { lock.waitLock(30000); }
  catch(e){ return {ok: false, error: 'LOCK_TIMEOUT: ' + (e && e.message || e)}; }
  try {
    var sh = _getChomusykyAttSheet(true);
    var vals = sh.getDataRange().getValues();
    for (var i = 1; i < vals.length; i++){
      var r = _parseChomusykyRow(vals[i]);
      if (r.loc === loc && r.child === child && r.date === date){
        return {ok: true, dup: true, id: r.id};
      }
    }
    var id = _nextChomusykyId(sh);
    sh.appendRow([id, date, loc, child, String(data.by || data.markedBy || '').trim(), new Date()]);
    return {ok: true, id: id};
  } catch(e){
    return {ok: false, error: String(e && e.message || e)};
  } finally {
    try { lock.releaseLock(); } catch(_){}
  }
}

// POST: прибрати візит — по id АБО по (loc, child, date).
function removeChomusykyMark(body){
  body = body || {};
  var id    = Number(body.id) || 0;
  var loc   = String(body.loc || '').trim();
  var child = String(body.child || '').trim();
  var date  = _dopDateISO(String(body.date || '').trim());
  if (!id && !(loc && child && date)) return {ok: false, error: 'Треба id АБО (loc + child + date)'};
  var lock = LockService.getScriptLock();
  try { lock.waitLock(30000); }
  catch(e){ return {ok: false, error: 'LOCK_TIMEOUT: ' + (e && e.message || e)}; }
  try {
    var sh;
    try { sh = _getChomusykyAttSheet(false); }
    catch(e){ return {ok: true, removed: false}; }
    var data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++){
      var r = _parseChomusykyRow(data[i]);
      var hit = id ? (r.id === id) : (r.loc === loc && r.child === child && r.date === date);
      if (hit){ sh.deleteRow(i + 1); return {ok: true, removed: true}; }
    }
    return {ok: true, removed: false};
  } catch(e){
    return {ok: false, error: String(e && e.message || e)};
  } finally {
    try { lock.releaseLock(); } catch(_){}
  }
}

// Звіт по абонементах за місяць: по кожній дитині — к-сть візитів (X з N) + сума
// абонемента (фіксовано за дитину з ≥1 візитом). ПРЕВʼЮ (без запису в Payment).
function getChomusykyReport(params){
  try {
    params = params || {};
    var loc = String(params.loc || '').trim();
    var mk = getChomusykyMarks({loc: loc, year: params.year, month: params.month});
    if (!mk.ok) return mk;
    var byChild = {};
    mk.items.forEach(function(m){
      if (!byChild[m.child]) byChild[m.child] = 0;
      byChild[m.child]++;
    });
    var sub = CHOMUSYKY_SUBSCRIPTION;
    var children = Object.keys(byChild).sort(function(a,b){ return a.localeCompare(b,'uk'); })
      .map(function(c){ return {child: c, visits: byChild[c], included: sub.visits, over: Math.max(0, byChild[c]-sub.visits)}; });
    var subscribers = children.length;              // дитина з ≥1 візитом = абонент
    var revenue = subscribers * sub.price;
    return {ok: true, loc: loc, year: params.year, month: params.month,
            subscription: sub, subscribers: subscribers, revenue: revenue, children: children};
  } catch(e){
    return {ok: false, error: String(e && e.message || e)};
  }
}

// ───────────────────────────────────────────────────────────────────────────
// v6.23 ДІАГНОСТИКА: запускати ВРУЧНУ з Apps Script editor щоб побачити що
// реально записано в "Додаткові_Відвідуваність" за останні N годин.
// Параметр HOURS (рядок 4) можна змінювати у функції — за замовч. 1 година.
// Дає швидку відповідь на питання: "Frontend каже 13 marks, скільки реально
// у Sheets?". Якщо у Sheets <13 — є втрачені у offline-черзі (фронт-side).
// ───────────────────────────────────────────────────────────────────────────
function diagAttendanceMarksRecent(){
  var HOURS = 1;
  var since = new Date(Date.now() - HOURS * 3600 * 1000);
  try {
    var sh = _getAttendanceSheet(false);
    var data = sh.getDataRange().getValues();
    if (data.length < 2){
      Logger.log('[diag] Лист порожній.');
      return {ok: true, total: 0, byChild: {}};
    }
    var byChild = {};
    var byActivity = {};
    var byLoc = {};
    var total = 0;
    for (var i = 1; i < data.length; i++){
      if (!data[i][0] && !data[i][4]) continue;
      var rec = _parseAttendanceRow(data[i]);
      var rawMarkedAt = data[i][9];
      if (!(rawMarkedAt instanceof Date)) continue;
      if (rawMarkedAt < since) continue;
      total++;
      if (!byChild[rec.child]) byChild[rec.child] = [];
      byChild[rec.child].push({
        markedAt: rawMarkedAt.toISOString(),
        date: rec.date, loc: rec.loc, group: rec.group,
        activity: rec.activityName, activityId: rec.activityId,
        price: rec.price, markedBy: rec.markedBy
      });
      byActivity[rec.activityName] = (byActivity[rec.activityName] || 0) + 1;
      byLoc[rec.loc] = (byLoc[rec.loc] || 0) + 1;
    }
    Logger.log('═══ DIAG: останні %s год ═══', HOURS);
    Logger.log('Всього marks: %s', total);
    Logger.log('Унікальних дітей: %s', Object.keys(byChild).length);
    Logger.log('По локаціях: %s', JSON.stringify(byLoc));
    Logger.log('По заняттях: %s', JSON.stringify(byActivity));
    Logger.log('── Per child ──');
    Object.keys(byChild).sort().forEach(function(c){
      var arr = byChild[c];
      Logger.log('  %s (%s mark%s):', c, arr.length, arr.length === 1 ? '' : 's');
      arr.forEach(function(m){
        Logger.log('    %s · %s · %s · %s · %s ₴ · by %s',
          m.markedAt, m.date, m.loc, m.activity, m.price, m.markedBy);
      });
    });
    Logger.log('═══════════════════════════════');
    return {ok: true, total: total, byChild: byChild, byActivity: byActivity, byLoc: byLoc, hours: HOURS};
  } catch(e){
    Logger.log('[diag] ❌ ' + (e && e.message || e));
    return {ok: false, error: String(e && e.message || e)};
  }
}

// v6.65: dedup guard helpers for extras attendance (prevents double rows -> double pay)
function _attDateISO(v, tz){
  if (v instanceof Date) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  var sx = String(v || '').trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(sx)) return sx;
  var d = new Date(sx);
  if (!isNaN(d.getTime())) return Utilities.formatDate(d, tz, 'yyyy-MM-dd');
  return sx;
}
function _attDupKey(date, child, actId, tz){
  return _attDateISO(date, tz) + '|' + String(child || '').trim().toLowerCase().replace(/\s+/g,' ') + '|' + (Number(actId) || 0);
}

function addAttendanceMark(data){
  // Валідація ДО локу (швидкий фейл).
  var date  = String(data.date  || '').trim();
  var child = String(data.child || '').trim();
  var actId = Number(data.activityId) || 0;
  if (!date || !child || !actId){
    return {ok: false, error: 'Поля Дата / Дитина / id_заняття обовʼязкові'};
  }
  // DEDUP-GUARD (дзеркало bulkAttendanceMarks): під LockService перевіряємо
  // _attDupKey (Дата + Дитина норм + id_заняття) проти існуючих рядків — якщо
  // такий запис уже є, НЕ дублюємо (повертаємо існуючий id, dup:true). Лок також
  // робить _nextAttendanceId атомарним → прибирає старий race із неунікальними id.
  var lock = LockService.getScriptLock();
  try { lock.waitLock(30000); }
  catch(e){ return {ok: false, error: 'LOCK_TIMEOUT: ' + (e && e.message || e)}; }
  try {
    var sh  = _getAttendanceSheet(true);
    var _tz = sh.getParent().getSpreadsheetTimeZone() || 'Europe/Kiev';
    var _k  = _attDupKey(date, child, actId, _tz);
    var _exVals = sh.getDataRange().getValues();
    for (var _e = 1; _e < _exVals.length; _e++){
      if (_attDupKey(_exVals[_e][1], _exVals[_e][4], _exVals[_e][5], _tz) === _k){
        return {ok: true, dup: true, id: Number(_exVals[_e][0]) || 0};   // повтор — тихо ігноруємо
      }
    }
    var id = _nextAttendanceId(sh);
    var row = [
      id, date,
      String(data.loc   || '').trim(),
      String(data.group || '').trim(),
      child, actId,
      String(data.activityName || '').trim(),
      Number(data.price) || 0,
      String(data.markedBy || '').trim(),
      new Date()
    ];
    sh.appendRow(row);
    return {ok: true, id: id};
  } catch(e){
    return {ok: false, error: String(e && e.message || e)};
  } finally {
    try { lock.releaseLock(); } catch(_){}
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

// v6.24: BULK — додає масив відміток за один LockService + один setValues.
// Замість 14× (HTTP→lock→appendRow) = 1× (HTTP→lock→batch write).
// Усуває race condition між паралельними auto-export запитами.
function bulkAttendanceMarks(body){
  var lock = LockService.getScriptLock();
  try { lock.waitLock(30000); }
  catch(e){ return {ok: false, error: 'LOCK_TIMEOUT: ' + (e && e.message || e)}; }
  try {
    var items = (body && body.items) || [];
    if (!items.length) return {ok: false, error: 'No items'};
    var sh = _getAttendanceSheet(true);
    var nextId = _nextAttendanceId(sh);
    var results = [];
    var rows = [];
    var now = new Date();
    var _tz = sh.getParent().getSpreadsheetTimeZone() || 'Europe/Kiev';
    var _seen = {};
    var _exVals = sh.getDataRange().getValues();
    for (var _e = 1; _e < _exVals.length; _e++){
      _seen[_attDupKey(_exVals[_e][1], _exVals[_e][4], _exVals[_e][5], _tz)] = true;
    }
    for (var i = 0; i < items.length; i++){
      var d = items[i] || {};
      var date  = String(d.date  || '').trim();
      var child = String(d.child || '').trim();
      var actId = Number(d.activityId) || 0;
      if (!date || !child || !actId){
        results.push({ok: false, error: 'Поля Дата/Дитина/id_заняття обовʼязкові'});
        continue;
      }
      var _k = _attDupKey(date, child, actId, _tz);
      if (_seen[_k]){ results.push({ok: true, dup: true}); continue; }
      _seen[_k] = true;
      var id = nextId++;
      rows.push([
        id, date,
        String(d.loc   || '').trim(),
        String(d.group || '').trim(),
        child, actId,
        String(d.activityName || '').trim(),
        Number(d.price) || 0,
        String(d.markedBy || '').trim(),
        now
      ]);
      results.push({ok: true, id: id});
    }
    if (rows.length){
      var startRow = sh.getLastRow() + 1;
      sh.getRange(startRow, 1, rows.length, ATTENDANCE_HEADER.length).setValues(rows);
    }
    return {ok: true, results: results, count: rows.length};
  } catch(e){
    return {ok: false, error: String(e && e.message || e)};
  } finally {
    try { lock.releaseLock(); } catch(_){}
  }
}

// v6.24: BULK — видаляє масив відміток за один lock + один getValues.
// Видалення рядків у зворотньому порядку щоб індекси не зсувались.
function bulkRemoveAttendanceMarks(body){
  var lock = LockService.getScriptLock();
  try { lock.waitLock(30000); }
  catch(e){ return {ok: false, error: 'LOCK_TIMEOUT: ' + (e && e.message || e)}; }
  try {
    var ids = (body && body.ids) || [];
    if (!ids.length) return {ok: false, error: 'No ids'};
    var idSet = {};
    for (var k = 0; k < ids.length; k++){
      var n = Number(ids[k]) || 0;
      if (n) idSet[n] = true;
    }
    var sh = _getAttendanceSheet(false);
    var data = sh.getDataRange().getValues();
    var rowsToDelete = [];
    var found = {};
    for (var i = 1; i < data.length; i++){
      var rid = Number(data[i][0]) || 0;
      if (rid && idSet[rid]){
        rowsToDelete.push(i + 1);
        found[rid] = true;
      }
    }
    rowsToDelete.sort(function(a, b){ return b - a; });
    for (var j = 0; j < rowsToDelete.length; j++){
      sh.deleteRow(rowsToDelete[j]);
    }
    var results = [];
    for (var m = 0; m < ids.length; m++){
      var nn = Number(ids[m]) || 0;
      if (nn && found[nn]) results.push({ok: true, id: nn});
      else                 results.push({ok: false, id: nn, error: 'Відмітку не знайдено'});
    }
    return {ok: true, results: results, removed: rowsToDelete.length};
  } catch(e){
    return {ok: false, error: String(e && e.message || e)};
  } finally {
    try { lock.releaseLock(); } catch(_){}
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// ЧИСТКА ДУБЛІВ ДОДАТКОВИХ ВІДМІТОК (Додаткові_Відвідуваність).
// Дубль = той самий Дата + Дитина(норм) + id_заняття записаний 2+ разів (роздуває
// "Факт доп" у рахунках). У кожній групі лишаємо keeper = НАЙРАНІШИЙ Час_відмітки,
// решту (пізніші) видаляємо.
// ⚠️ Видалення ТІЛЬКИ ПО ПОЗИЦІЇ РЯДКА (не по id!) — id у цьому аркуші НЕ унікальні
// (race у _nextAttendanceId без локу), тож removeAttendanceMark(id) знесе чужий рядок.
// Запускати ВРУЧНУ з Apps Script editor. Порядок: backup → Dry → Apply.
// ═══════════════════════════════════════════════════════════════════════════
var _DEDUP_EXTRAS_BACKUP_PREFIX = 'BACKUP_extras_attendance_';

// Час_відмітки → мілісекунди для порівняння (інвалід → +Infinity, щоб НЕ став keeper).
function _attMarkedAtMs(v){
  if (v instanceof Date) return v.getTime();
  var s = String(v || '').trim();
  if (!s) return Number.POSITIVE_INFINITY;
  var t = new Date(s).getTime();
  return isNaN(t) ? Number.POSITIVE_INFINITY : t;
}

// 1) БЕКАП: повна копія аркуша Додаткові_Відвідуваність у BACKUP_extras_attendance_<stamp>.
function backupAttendanceExtras(){
  var ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  var src = ss.getSheetByName(ATTENDANCE_SHEET_NAME);
  if (!src){
    Logger.log('❌ Аркуш "%s" не знайдено — бекап не створено', ATTENDANCE_SHEET_NAME);
    return {ok: false, error: 'Sheet "' + ATTENDANCE_SHEET_NAME + '" не знайдено'};
  }
  var stamp = Utilities.formatDate(new Date(), 'GMT+3', 'yyyyMMdd_HHmmss');
  var bName = _DEDUP_EXTRAS_BACKUP_PREFIX + stamp;
  var copy = src.copyTo(ss).setName(bName);
  var rows = Math.max(0, copy.getLastRow() - 1);
  Logger.log('✅ Бекап створено: аркуш "%s", рядків даних=%s', bName, rows);
  return {ok: true, backupSheet: bName, dataRows: rows};
}

// Спільний сканер: повертає {groups:[{loc,child,date,actId,actName,keeperRow,keeperAt,
// delRows:[...], delPrice}], totalDelRows, totalOver}. Працює на «сирих» рядках аркуша.
function _dedupExtrasScan(){
  var ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  var sh = ss.getSheetByName(ATTENDANCE_SHEET_NAME);
  if (!sh) return {ok: false, error: 'Sheet "' + ATTENDANCE_SHEET_NAME + '" не знайдено'};
  var tz = ss.getSpreadsheetTimeZone() || 'Europe/Kiev';
  var data = sh.getDataRange().getValues();
  var buckets = {};   // dupKey → [{row(1-based), child, date, actId, actName, price, atMs}]
  for (var i = 1; i < data.length; i++){
    var row = data[i];
    if (!row[0] && !row[4]) continue;                 // порожній рядок
    var child = String(row[4] || '').trim();
    var actId = Number(row[5]) || 0;
    if (!child || !actId) continue;                   // неповний — не чіпаємо
    var k = _attDupKey(row[1], child, actId, tz);
    (buckets[k] = buckets[k] || []).push({
      row: i + 1,                                     // 1-based позиція у аркуші
      child: child,
      date: _attDateISO(row[1], tz),
      actId: actId,
      actName: String(row[6] || '').trim(),
      price: Number(row[7]) || 0,
      atMs: _attMarkedAtMs(row[9])
    });
  }
  var groups = [], totalDelRows = 0, totalOver = 0;
  Object.keys(buckets).forEach(function(k){
    var arr = buckets[k];
    if (arr.length < 2) return;
    // keeper = найраніший Час_відмітки; тай-брейк — менший номер рядка (раніший у аркуші)
    arr.sort(function(a, b){ return (a.atMs - b.atMs) || (a.row - b.row); });
    var keeper = arr[0];
    var del = arr.slice(1);
    var delPrice = del.reduce(function(s, x){ return s + x.price; }, 0);
    groups.push({
      loc: String(data[keeper.row - 1][2] || '').trim() || '(порожньо)',
      child: keeper.child, date: keeper.date, actId: keeper.actId, actName: keeper.actName,
      keeperRow: keeper.row, keeperAtMs: keeper.atMs,
      delRows: del.map(function(x){ return x.row; }),
      delPrice: delPrice
    });
    totalDelRows += del.length;
    totalOver += delPrice;
  });
  return {ok: true, groups: groups, totalDelRows: totalDelRows, totalOver: totalOver, tz: tz};
}

// 2) DRY-RUN: лише Logger.log, НІЧОГО не видаляє.
function dedupExtrasAttendanceDry(){
  var scan = _dedupExtrasScan();
  if (!scan.ok){ Logger.log('❌ %s', scan.error); return scan; }
  Logger.log('═══ DRY-RUN: чистка дублів додаткових — НІЧОГО НЕ ВИДАЛЯЄМО ═══');
  Logger.log('Формат: дитина | дата | заняття | на видалення | keeper-рядок | рядки-на-видалення (ПО ПОЗИЦІЇ, не id)');

  var byLoc = {};
  scan.groups.forEach(function(g){
    if (!byLoc[g.loc]) byLoc[g.loc] = {groups: 0, delRows: 0, over: 0, items: []};
    byLoc[g.loc].groups++; byLoc[g.loc].delRows += g.delRows.length; byLoc[g.loc].over += g.delPrice;
    byLoc[g.loc].items.push(g);
  });

  Object.keys(byLoc).sort(function(a, b){ return byLoc[b].over - byLoc[a].over; }).forEach(function(loc){
    var L = byLoc[loc];
    Logger.log('\n━━━ %s ━━━ груп=%s | видалиться рядків=%s | прибереться завищення=%s грн', loc, L.groups, L.delRows, L.over);
    L.items.sort(function(a, b){ return b.delPrice - a.delPrice || b.delRows.length - a.delRows.length; }).forEach(function(g){
      Logger.log('   %s | %s | %s (id%s) | видалити ×%s | keeper-рядок=%s | видалити-рядки[%s]',
        g.child, g.date, g.actName, g.actId, g.delRows.length, g.keeperRow, g.delRows.join(','));
    });
  });
  Logger.log('\n─── ПІДСУМОК DRY: груп-дублів=%s | рядків на видалення=%s | прибереться завищення=%s грн ───',
    scan.groups.length, scan.totalDelRows, scan.totalOver);
  return {ok: true, dupGroups: scan.groups.length, rowsToDelete: scan.totalDelRows, overchargeRemoved: scan.totalOver};
}

// 3) APPLY: РЕАЛЬНО видаляє пізніші дублі ПО ПОЗИЦІЇ РЯДКА (знизу вгору).
// Гарантія безпеки: НЕ запуститься, поки немає бекап-аркуша (backupAttendanceExtras).
function dedupExtrasAttendanceApply(){
  var ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  // safety-guard: вимагаємо хоча б один бекап-аркуш
  var hasBackup = ss.getSheets().some(function(s){ return s.getName().indexOf(_DEDUP_EXTRAS_BACKUP_PREFIX) === 0; });
  if (!hasBackup){
    Logger.log('❌ Бекап не знайдено. Спершу запусти backupAttendanceExtras().');
    return {ok: false, error: 'Спершу backupAttendanceExtras() — бекап-аркуш відсутній'};
  }
  var lock = LockService.getScriptLock();
  try { lock.waitLock(30000); }
  catch(e){ return {ok: false, error: 'LOCK_TIMEOUT: ' + (e && e.message || e)}; }
  try {
    // СВІЖИЙ скан під локом (щоб позиції рядків були актуальні на момент видалення)
    var scan = _dedupExtrasScan();
    if (!scan.ok){ Logger.log('❌ %s', scan.error); return scan; }
    var sh = ss.getSheetByName(ATTENDANCE_SHEET_NAME);

    // зібрати всі позиції на видалення, відсортувати ЗНИЗУ ВГОРУ (щоб індекси не зсувались)
    var rowsToDelete = [];
    scan.groups.forEach(function(g){ g.delRows.forEach(function(r){ rowsToDelete.push(r); }); });
    rowsToDelete.sort(function(a, b){ return b - a; });

    Logger.log('═══ APPLY: видаляю %s дублюючих рядків по позиції (знизу вгору) ═══', rowsToDelete.length);
    var deleted = 0;
    for (var i = 0; i < rowsToDelete.length; i++){
      sh.deleteRow(rowsToDelete[i]);
      deleted++;
    }
    Logger.log('─── ГОТОВО: видалено рядків=%s | прибрано завищення=%s грн | груп оброблено=%s ───',
      deleted, scan.totalOver, scan.groups.length);
    return {ok: true, deleted: deleted, overchargeRemoved: scan.totalOver, dupGroups: scan.groups.length};
  } catch(e){
    Logger.log('[dedupExtrasApply] EXCEPTION: %s\n%s', e && e.message, e && e.stack);
    return {ok: false, error: String(e && e.message || e)};
  } finally {
    try { lock.releaseLock(); } catch(_){}
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// ВИДАЛЕННЯ ХИБНИХ Логопед-відміток (Кругла) — 3 дітям, що НЕ ходять на логопеда.
// READ-ONLY ФАЙНДЕР: показує точні №рядків у аркуші для звірки ПЕРЕД видаленням.
// Видалення лише по ПОЗИЦІЇ рядка (id не унікальні). Запускати ВРУЧНУ з редактора.
// ═══════════════════════════════════════════════════════════════════════════
var _LOGOPED_FIX_LOC   = 'Кругла';
var _LOGOPED_FIX_NAMES = ['Голуб Еней', 'Станіславська Софія', 'Неруш Мирон'];

function findLogopedKruglaToDelete(){
  var ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  var sh = ss.getSheetByName(ATTENDANCE_SHEET_NAME);
  if (!sh){ Logger.log('❌ аркуш "%s" не знайдено', ATTENDANCE_SHEET_NAME); return {ok:false, error:'no sheet'}; }
  var tz = ss.getSpreadsheetTimeZone() || 'Europe/Kiev';
  var nameSet = _LOGOPED_FIX_NAMES.map(function(n){ return _normForMatch(n); });
  var data = sh.getDataRange().getValues();
  var hits = [];
  for (var i = 1; i < data.length; i++){
    var row = data[i];
    if (!row[0] && !row[4]) continue;                              // порожній рядок
    if (trim(String(row[2])) !== _LOGOPED_FIX_LOC) continue;       // лише Кругла
    if (!/логопед/i.test(String(row[6] || ''))) continue;          // ЛИШЕ Логопед (Футбол не чіпаємо)
    if (nameSet.indexOf(_normForMatch(row[4])) === -1) continue;   // лише 3 дитини
    hits.push({
      rowPos: i + 1,                                               // позиція в аркуші (1-based)
      id: row[0],
      date: _attDateISO(row[1], tz),
      child: String(row[4] || '').trim(),
      activity: String(row[6] || '').trim(),
      price: Number(row[7]) || 0,
      markedBy: String(row[8] || '').trim(),
      markedAt: row[9] instanceof Date ? row[9].toISOString() : String(row[9] || '')
    });
  }
  Logger.log('═══ READ-ONLY: Логопед на видалення (loc=%s) — знайдено %s ═══', _LOGOPED_FIX_LOC, hits.length);
  Logger.log('Формат: №рядка | id | дата | дитина | заняття | ціна | відмітив | час');
  hits.forEach(function(h){
    Logger.log('  рядок %s | id=%s | %s | %s | %s | %s грн | %s | %s',
      h.rowPos, h.id, h.date, h.child, h.activity, h.price, h.markedBy, h.markedAt);
  });
  var rowsDesc = hits.map(function(h){ return h.rowPos; }).sort(function(a, b){ return b - a; });
  Logger.log('─── Позиції на видалення (знизу вгору): [%s] | прибереться: %s грн | НІЧОГО НЕ ВИДАЛЕНО (read-only) ───',
    rowsDesc.join(','), hits.reduce(function(s, h){ return s + h.price; }, 0));
  return {ok:true, count:hits.length, rows:rowsDesc, hits:hits};
}

// APPLY: видаляє знайдені Логопед-рядки ПО ПОЗИЦІЇ (знизу вгору, id не унікальні).
// Guard: без бекап-аркуша (backupAttendanceExtras) не запускається. LockService +
// СВІЖИЙ скан під локом (findLogopedKruglaToDelete) — позиції актуальні на момент видалення.
function deleteLogopedKruglaApply(){
  var ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  var hasBackup = ss.getSheets().some(function(s){ return s.getName().indexOf(_DEDUP_EXTRAS_BACKUP_PREFIX) === 0; });
  if (!hasBackup){
    Logger.log('❌ Бекап не знайдено. Спершу запусти backupAttendanceExtras().');
    return {ok:false, error:'Спершу backupAttendanceExtras() — бекап-аркуш відсутній'};
  }
  var lock = LockService.getScriptLock();
  try { lock.waitLock(30000); }
  catch(e){ return {ok:false, error:'LOCK_TIMEOUT: ' + (e && e.message || e)}; }
  try {
    var scan = findLogopedKruglaToDelete();        // СВІЖИЙ скан під локом (read-only)
    if (!scan.ok) return scan;
    if (!scan.rows || !scan.rows.length){
      Logger.log('ℹ️ Нічого видаляти — Логопед-рядків не знайдено.');
      return {ok:true, deleted:0, removedAmount:0};
    }
    var sh = ss.getSheetByName(ATTENDANCE_SHEET_NAME);
    var rowsDesc = scan.rows.slice().sort(function(a, b){ return b - a; });   // ЗНИЗУ ВГОРУ
    Logger.log('═══ APPLY: видаляю %s Логопед-рядків по позиції (знизу вгору): [%s] ═══', rowsDesc.length, rowsDesc.join(','));
    var deleted = 0;
    for (var i = 0; i < rowsDesc.length; i++){ sh.deleteRow(rowsDesc[i]); deleted++; }
    var removed = scan.hits.reduce(function(s, h){ return s + h.price; }, 0);
    Logger.log('─── ГОТОВО: видалено рядків=%s | прибрано=%s грн ───', deleted, removed);
    return {ok:true, deleted:deleted, removedAmount:removed, rows:rowsDesc, hits:scan.hits};
  } catch(e){
    Logger.log('[deleteLogopedKrugla] EXCEPTION: %s\n%s', e && e.message, e && e.stack);
    return {ok:false, error:String(e && e.message || e)};
  } finally {
    try { lock.releaseLock(); } catch(_){}
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

function _prevMonth(month, year){
  return (Number(month) <= 1) ? {month: 12, year: Number(year) - 1} : {month: Number(month) - 1, year: Number(year)};
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
    var formulaConverted = 0;

    // ⚠️ ТОЧКОВИЙ запис: setValue() ЛИШЕ у child-рядки (paymentByNorm побудований без
    // group-headers), і лише якщо значення змінилось.
    // v6.x: формульні child-рядки БІЛЬШЕ НЕ пропускаємо (дзеркало exportVacationDiscountToPayments):
    // getValues() дає ОБЧИСЛЕНЕ значення формули (напр. ручне =-50 → -50) — воно стає базою,
    // setValue замінює формулу числом, зверху додаються додаткові (−50 + 1350 = 1300).
    // Підсумкові =SUM-рядки НЕ чіпаються: вони або group-header (не в paymentByNorm), або
    // не матчать дитину → newSum=0, lastWritten=0 → newValue===currentValue → запису немає.
    Object.keys(paymentByNorm).forEach(function(nk){
      var rowIdx0 = paymentByNorm[nk];
      var wasFormula   = !!(colFormulas[rowIdx0] && colFormulas[rowIdx0][0]);
      var paymentName  = trim(String(data[rowIdx0][0] || ''));
      var currentValue = Number(colValues[rowIdx0][0]) || 0;   // результат формули (getValues)
      var je           = journal.byNormName[nk];
      var lastWritten  = je ? je.sum : 0;
      var baseValue    = currentValue - lastWritten;
      var match        = sumByNorm[nk];
      var newSum       = match ? match.sum : 0;
      var newValue     = baseValue + newSum;

      // Точковий запис лише змінених клітинок (формула → число відбувається тут же).
      if (newValue !== currentValue){
        paySh.getRange(rowIdx0 + 1, budgetDopCol1).setValue(newValue);
        cellsWritten++;
        if (wasFormula) formulaConverted++;
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

    Logger.log('[exportAttendanceToPayments] точковий запис: %s клітинок змінено, %s формул конвертовано в число', cellsWritten, formulaConverted);

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
      formulaCellsConverted: formulaConverted,   // v6.x: формул→число у child-рядках
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

// ═══════════════════════════════════════════════════════════════════════════
// v6.51 — АВТО-ЗВІРКА ПЛАТЕЖІВ (виписка ПриватБанку → Payment «Факт навч»/«Факт доп»).
// ФАЗА 1 (read-only): reconcilePreview — резолв локації/категорії за IBAN,
//   витяг платника, матч платник→дитина(діти) у межах локації. БЕЗ ЗАПИСІВ.
//   Парсинг файлу виписки — на фронті (SheetJS); сюди приходить готовий JSON.
// ═══════════════════════════════════════════════════════════════════════════

// 0-based індекс колонки → літера (26→AA). Для діагностики у прев'ю.
function _colLetter(idx0){
  var s = '', n = idx0 + 1;
  while (n > 0){ var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

// Парс дати платежу: dd.MM.yyyy | dd/MM/yyyy | yyyy-MM-dd | інше → Date|null.
function _recParseDate(v){
  if (v instanceof Date) return v;
  var s = String(v || '').trim(); if (!s) return null;
  var m = s.match(/^(\d{2})[.\/](\d{2})[.\/](\d{4})/); if (m) return new Date(+m[3], +m[2] - 1, +m[1]);
  m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);              if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
  var d = new Date(s); return isNaN(d) ? null : d;
}

// IBAN (з виписки) → {loc, type, ...} з Реквізити_Локацій (CONFIG). type: studies|extras.
function _resolveAccountByIban(iban){
  var norm = String(iban || '').replace(/\s+/g, '').toUpperCase();
  if (!norm) return null;
  var sh = SpreadsheetApp.openById(CONFIG_SHEET_ID).getSheetByName('Реквізити_Локацій');
  if (!sh) return null;
  var data = sh.getDataRange().getValues();
  for (var r = 1; r < data.length; r++){
    var e = String(data[r][4] || '').replace(/\s+/g, '').toUpperCase();   // E = IBAN
    if (e && e === norm){
      var typB = String(data[r][1] || '').trim().toLowerCase();           // B = Тип
      var type = (typB.indexOf('додатков') >= 0) ? 'extras' : 'studies';
      return { loc: trim(data[r][0]), type: type, typeLabel: trim(data[r][1]), orgName: trim(data[r][2]) };
    }
  }
  return null;
}

// Header-driven колонка місяця (узагальнення detectCurrentMonthCol під будь-який jsMonth 0-11).
// Повертає 0-based індекс СТАРТУ блоку місяця (= «Факт навч»), або -1.
function _detectMonthColByHeader(rows, jsMonth){
  for (var pass = 0; pass < 2; pass++){
    for (var r = 0; r < Math.min(3, rows.length); r++){
      for (var c = 1; c < rows[r].length; c++){
        var cell = String(rows[r][c] || '').toLowerCase().trim();
        for (var mi = 0; mi < MONTHS_UA.length; mi++){
          var hit = pass === 0 ? (cell === MONTHS_UA[mi]) : (cell.indexOf(MONTHS_UA[mi]) >= 0);
          if (hit && MONTHS_JS[mi] === jsMonth) return c;
        }
      }
    }
  }
  return -1;
}
// Факт-колонка за категорією: studies → старт блоку (Факт навч); extras → +2 (Факт доп). 0-based.
function _factColForType(monthStartCol0, type){
  return type === 'extras' ? (monthStartCol0 + 2) : (type === 'vstup' ? (monthStartCol0 + 1) : monthStartCol0);
}

// Детект колонки «сторона договору» у Payment-листі по ШАПЦІ (нинішня BJ, але без хардкоду).
function _detectContractPartyCol(rows){
  for (var r = 0; r < Math.min(5, rows.length); r++){
    for (var c = 0; c < rows[r].length; c++){
      var cell = String(rows[r][c] || '').toLowerCase();
      if (cell.indexOf('сторона договору') >= 0 || cell.indexOf('платник') >= 0) return c;
    }
  }
  return -1;
}

// ГОМОГЛІФИ: ПриватБанк пише ПІБ латинськими двійниками ("Сахащiк" — латинська 'i').
// Зводимо латиницю до кирилиці — застосовуємо ОДНАКОВО з обох боків (платник + індекс).
function _lat2cyr(lowerStr){
  var map = {a:'а', c:'с', e:'е', i:'і', o:'о', p:'р', x:'х', y:'у', k:'к', m:'м'};
  return String(lowerStr || '').replace(/[aceiopxykm]/g, function(ch){ return map[ch] || ch; });
}
// Нормалізація ПІБ (повна): апострофи знімаємо, lower, гомогліфи→кирилиця, схлопуємо пробіли.
function _normPayerName(s){
  var t = String(s || '').replace(/[`'’ʼ]/g, '').toLowerCase().trim().replace(/\s+/g, ' ');
  return _lat2cyr(t);
}
// КЛЮЧ МАТЧИНГУ = ПРІЗВИЩЕ (перший токен). Платник-батько часто має інше ІМ'Я,
// ніж parent у BJ (мама) — тож матчимо лише по прізвищу. Складене прізвище через
// дефіс лишається цілим токеном.
function _surnameKey(s){
  var t = _normPayerName(s);
  if (!t) return '';
  return t.split(' ')[0];
}

// Витяг платника за дискримінатором коду контрагента:
//   10 цифр → фізособа (платник = «Назва контрагента»);
//    8 цифр → юр/банк → шукаємо «Платник <ПІБ>» у призначенні (тягнемо повне ПІБ).
function _fixHomoglyph(s){
  return String(s || '').replace(/i/g, 'і').replace(/I/g, 'І');
}
function _extractPayer(rec){
  var code = String(rec.edrpou || '').replace(/\D/g, '');
  var cp = _fixHomoglyph(rec.counterparty);
  if (code.length === 10) return { raw: trim(cp), via: 'individual' };
  var purpose = _fixHomoglyph(String(rec.purpose || ''));
  // Прізвище (з великої) + до 2 наступних токенів (ім'я/по-батькові/ініціали з крапками).
  var m = purpose.match(/платник[\s:]*([А-ЯІЇЄҐ][а-яіїєґ'’ʼ`\-]+(?:\s+[А-ЯІЇЄҐ][а-яіїєґ'’ʼ`.]*){0,2})/i);
  if (m) return { raw: m[1].trim(), via: 'bank' };
  var v = purpose.match(/[Вв]ід\s+([А-ЯІЇЄҐ][а-яіїєґ'’ʼ`\-]+(?:\s+[А-ЯІЇЄҐ][а-яіїєґ'’ʼ`.]*){0,2})/);
  if (v) return { raw: v[1].trim(), via: 'bank-vid' };
  return { raw: trim(cp), via: 'bank-unparsed' };
}

// Індекс «ПРІЗВИЩЕ → [діти]» у межах локації. Три джерела прізвища:
//   (A) ПІБ ДИТИНИ (кол. A Payment) — прізвище дитини часто = прізвище платника;
//   (BJ) «сторона договору» у Payment;
//   (картка) батьки з CRM (мама/тато/підписант).
// Один платник може мати кілька дітей (брати-сестри) → значення = СПИСОК (dedup по childName).
// Повертає {index, diag:{bjColDetected, fromChild, fromBJ, fromCard}}.
function _buildPayerIndex(loc){
  var idx = {};   // surnameKey -> [{childName, row(0-based; -1 якщо лише з CRM), group}]
  var diag = { bjColDetected: false, fromChild: 0, fromBJ: 0, fromCard: 0 };
  // add: childName унікальний у межах ключа-прізвища (одну дитину з різних джерел не дублюємо)
  function add(surnameSrc, child){
    var sk = _surnameKey(surnameSrc); if (!sk || !child.childName) return false;
    var list = (idx[sk] = idx[sk] || []);
    for (var i = 0; i < list.length; i++) if (list[i].childName === child.childName) return false;
    list.push(child); return true;
  }
  // 1) Payment-файл: ПІБ дитини (A) [джерело A] + «сторона договору» (по шапці) [джерело BJ]
  var payRows = [];
  var reg = _getLocationPaymentRegistry(loc);
  if (reg && reg.sheetId){
    var ss = SpreadsheetApp.openById(reg.sheetId);
    var sh = ss.getSheetByName(reg.sheetName) || ss.getSheets()[0];
    var data = sh.getDataRange().getValues();
    var partyCol = _detectContractPartyCol(data);
    diag.bjColDetected = (partyCol >= 0);
    var curGroup = '';
    for (var r = 3; r < data.length; r++){
      var name = trim(data[r][0]);
      if (!name) continue;
      if (isGroupHeaderRow(data[r], 1)){ curGroup = normalizeGroupName(name); continue; }
      var child = { childName: name, row: r, group: curGroup };
      payRows.push(child);
      if (add(name, child)) diag.fromChild++;                              // прізвище ДИТИНИ
      if (partyCol >= 0){ var party = trim(data[r][partyCol]); if (party && add(party, child)) diag.fromBJ++; }
    }
  }
  // 2) CRM-картки: батьки (мама/тато/підписант) для дітей цієї локації → той самий payRow по імені
  var byNorm = {}; payRows.forEach(function(c){ byNorm[_normNameVac(c.childName)] = c; });
  var cli = getClients();
  if (cli && cli.ok){
    (cli.data || []).forEach(function(o){
      if (trim(o['Локація']) !== loc) return;
      var cn = trim(o['ПІБ дитини']); if (!cn) return;
      var child = byNorm[_normNameVac(cn)] || { childName: cn, row: -1, group: trim(o['Група']) };
      ['ПІБ мами', 'ПІБ тата', 'Підписант договору'].forEach(function(f){
        var p = trim(o[f]); if (p && add(p, child)) diag.fromCard++;
      });
    });
  }
  return { index: idx, diag: diag, roster: payRows.map(function(c){ return { name: c.childName, row: c.row, group: c.group }; }) };
}

// ── ФАЗА 1: PREVIEW (read-only). Вхід: {iban, payments:[{date,amount,purpose,edrpou,counterparty,ref,...}]}.
// v6.74 — фолбек: якщо по платнику дитину не знайдено — шукаємо прізвища дітей у призначенні (транзитні рахунки, платник ≠ сторона договору).


// v6.74 — пошук дітей по ВСЬОМУ тексту виписки (призначення + контрагент + платник + усі поля).
function _recSearchBlob(rec){
  var blob = '';
  if (rec && typeof rec === 'object'){
    for (var k in rec){ if (rec.hasOwnProperty(k) && typeof rec[k] === 'string') blob += ' ' + rec[k]; }
  }
  if (typeof _fixHomoglyph === 'function') blob = _fixHomoglyph(blob);
  return blob;
}
function _scanRecordForChildren(rec, idx){
  var out = [], seen = {};
  var toks = _recSearchBlob(rec).match(/[А-ЯІЇЄҐ][а-яіїєґ'’ʼ`\-]{2,}/g) || [];
  for (var i = 0; i < toks.length; i++){
    var sk = _surnameKey(toks[i]); var list = idx[sk];
    if (!list) continue;
    for (var j = 0; j < list.length; j++){
      if (seen[list[j].childName]) continue;
      seen[list[j].childName] = true; out.push(list[j]);
    }
  }
  return out;
}

function reconcilePreview(body){
  try{
    body = body || {};
    var iban = trim(body.iban);
    var payments = body.payments || [];
    var acct = _resolveAccountByIban(iban);
    if (!acct || !acct.loc) return {ok:false, error:'Рахунок "' + iban + '" не знайдено в Реквізити_Локацій'};
    var loc = acct.loc, type = acct.type;

    var reg = _getLocationPaymentRegistry(loc);
    if (!reg || !reg.sheetId) return {ok:false, error:'Локацію "' + loc + '" не знайдено в CONFIG-реєстрі'};
    var ss = SpreadsheetApp.openById(reg.sheetId);
    var paySh = ss.getSheetByName(reg.sheetName) || ss.getSheets()[0];
    var data = paySh.getDataRange().getValues();
    var built = _buildPayerIndex(loc);
    var idx = built.index;

    // Дати, що вже звірялися (для цього loc+iban) — м'яке попередження «день уже оброблено».
    var processed = {};
    var ibanN = iban.replace(/\s+/g, '').toUpperCase();
    var logSh0 = _getReconcileLogSheet();
    var lv0 = logSh0.getDataRange().getValues();
    for (var lr = 1; lr < lv0.length; lr++){
      if (trim(lv0[lr][2]) === loc && String(lv0[lr][3] || '').replace(/\s+/g, '').toUpperCase() === ibanN){
        var dd = trim(lv0[lr][4]); if (dd) processed[dd] = true;
      }
    }

    var rows = payments.map(function(rec, i){
      var payer = _extractPayer(rec);
      var sk = _surnameKey(payer.raw);               // КЛЮЧ = ПРІЗВИЩЕ платника
      var uniq = (idx[sk] || []).slice();
      var matchVia = uniq.length ? 'payer' : '';
      // _scanRecordForChildren сканує ВЕСЬ blob запису (включно з purpose),
      // тож окремий purpose-фолбек зайвий. Текстовий матч → мітка 'text'.
      var _pm = _scanRecordForChildren(rec, idx);
      if (_pm.length){
        var _seen = {}; uniq.forEach(function(c){ _seen[c.childName] = 1; });
        _pm.forEach(function(c){ if (!_seen[c.childName]){ _seen[c.childName] = 1; uniq.push(c); } });
        if (!matchVia) matchVia = 'text';
      }            // вже dedup по childName в індексі

      var d = _recParseDate(rec.date);
      var jsMonth = d ? d.getMonth() : -1;
      var monthCol0 = (jsMonth >= 0) ? _detectMonthColByHeader(data, jsMonth) : -1;
      var factCol0  = (monthCol0 >= 0) ? _factColForType(monthCol0, type) : -1;

      // Непорожній Факт — НОРМА (батьки платять частинами → сума ДОДАСТЬСЯ у Фазі 2).
      // Тому "дубля" по непорожній клітинці НЕ існує; справжній дубль = той самий
      // Референс уже у Звірки_Платежів — а у Фазі 1 (без записів) дублів не буває.
      var existing = 0;
      if (uniq.length === 1 && uniq[0].row >= 0 && factCol0 >= 0){
        existing = Number(toNum((data[uniq[0].row] || [])[factCol0])) || 0;
      }
      var status = uniq.length === 0 ? 'none' : (uniq.length === 1 ? 'auto' : 'multi');

      return {
        i: i, date: String(rec.date || ''), amount: Number(rec.amount) || 0, ref: String(rec.ref || ''),
        payerRaw: payer.raw, payerVia: payer.via, payerSurname: sk, matchVia: matchVia,
        month: (jsMonth + 1), status: status, candidates: uniq, existing: existing,
        factCol: factCol0 >= 0 ? _colLetter(factCol0) : ''
      };
    });

    return {ok:true, loc:loc, type:type, typeLabel:acct.typeLabel, orgName:acct.orgName,
            iban:iban, count:rows.length, rows:rows, diag:built.diag,
            roster:(built.roster||[]),
            processedDates:Object.keys(processed)};
  } catch(e){ return {ok:false, error: e.message || String(e)}; }
}

// ── Лог-аркуш звірок (CONFIG). Колонка «Ключ дедупу» (15) — для ідемпотентності по Референсу.
var RECONCILE_LOG_SHEET = 'Звірки_Платежів';
function _getReconcileLogSheet(){
  var ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  var sh = ss.getSheetByName(RECONCILE_LOG_SHEET);
  if (!sh){
    sh = ss.insertSheet(RECONCILE_LOG_SHEET);
    sh.getRange(1, 1, 1, 15).setValues([['Коли','Ким','Локація','IBAN','Дата платежу','Референс',
      'Платник','Дитина','Рядок','Місяць','Колонка','Сума','Було','Стало','Ключ дедупу']]);
    sh.setFrozenRows(1);
  }
  return sh;
}

// ── ФАЗА 2: ЗАПИС. Вхід: {iban, by, items:[{childRow(0-based), childName, amount, date, ref, payerRaw}]}.
//   Накопичувально (Факт += сума). Колонка — ТІЛЬКИ по шапці (−1 → помилка, без offset).
//   Ідемпотентність по Референсу (порожній → composite). Skip formula. LockService.
function reconcileApply(body){
  var lock = LockService.getScriptLock();
  try { lock.waitLock(30000); }
  catch(e){ return {ok:false, error:'Система зайнята (інша звірка триває) — спробуйте ще раз'}; }
  try {
    body = body || {};
    var iban = trim(body.iban);
    var by   = trim(body.by) || '?';
    var items = body.items || [];
    var acct = _resolveAccountByIban(iban);
    if (!acct || !acct.loc) return {ok:false, error:'Рахунок "' + iban + '" не знайдено в Реквізити_Локацій'};
    var loc = acct.loc, type = acct.type;
    var reg = _getLocationPaymentRegistry(loc);
    if (!reg || !reg.sheetId) return {ok:false, error:'Локацію "' + loc + '" не знайдено в CONFIG-реєстрі'};
    var ss = SpreadsheetApp.openById(reg.sheetId);
    var paySh = ss.getSheetByName(reg.sheetName) || ss.getSheets()[0];
    var data = paySh.getDataRange().getValues();

    var logSh = _getReconcileLogSheet();
    var lv = logSh.getDataRange().getValues();
    var applied = {};                              // ключ дедупу -> true (вже застосовано)
    for (var lr = 1; lr < lv.length; lr++){ var k = trim(lv[lr][14]); if (k) applied[k] = true; }

    var ibanN = iban.replace(/\s+/g, '').toUpperCase();
    var written = 0, skipped = 0, errors = 0, details = [], logAppend = [], overlay = {};

    items.forEach(function(it){
      var childRow = Number(it.childRow);          // 0-based у Payment
      var childName = trim(it.childName);
      var amount = Number(it.amount) || 0;
      var ref = trim(it.ref);
      var dateStr = String(it.date || '');
      var dupKey = ref ? ('REF|' + ref) : ('CMP|' + ibanN + '|' + dateStr + '|' + childRow + '|' + amount);

      if (applied[dupKey]){ skipped++; details.push({childName:childName, status:'skipped-dup', ref:ref}); return; }
      if (!(childRow >= 3) || childRow >= data.length){ errors++; details.push({childName:childName, status:'error', msg:'рядок поза межами'}); return; }

      var d = _recParseDate(dateStr);
      var jsMonth = d ? d.getMonth() : -1;
      var monthCol0 = (jsMonth >= 0) ? _detectMonthColByHeader(data, jsMonth) : -1;
      if (monthCol0 < 0){ errors++; details.push({childName:childName, status:'error', msg:'місяць не знайдено в шапці — НЕ записано'}); return; }
      var factCol0 = _factColForType(monthCol0, it.type || type);

      var cell = paySh.getRange(childRow + 1, factCol0 + 1);
      var okey = childRow + ',' + factCol0;
      // якщо формула — беремо її обчислене значення і замінюємо числом (значення + платіж)
      var prev = overlay.hasOwnProperty(okey) ? overlay[okey] : (Number(toNum(cell.getValue())) || 0);
      var nv = prev + amount;
      cell.setValue(nv);
      overlay[okey] = nv;
      applied[dupKey] = true;                      // анти-дубль і в межах батчу
      written++;
      logAppend.push([new Date(), by, loc, iban, dateStr, ref, trim(it.payerRaw), childName,
                      childRow + 1, (jsMonth + 1), _colLetter(factCol0), amount, prev, nv, dupKey]);
      details.push({childName:childName, status:'written', prev:prev, now:nv, col:_colLetter(factCol0)});
    });

    if (logAppend.length) logSh.getRange(logSh.getLastRow() + 1, 1, logAppend.length, 15).setValues(logAppend);
    return {ok:true, loc:loc, written:written, skipped:skipped, errors:errors, details:details};
  } catch(e){ return {ok:false, error: e.message || String(e)}; }
  finally { try { lock.releaseLock(); } catch(e){} }
}

// ═══════════════════════════════════════════════════════════════════════════
// v6.40 — ЗНИЖКА ВІДПУСТКИ → Платежі (колонка "Бюджет навч" таргет-місяця).
// Дзеркалить exportAttendanceToPayments: журнал (kind='vacation'), baseValue,
// точковий запис, skip формул/group-header. Знижка = ВІД'ЄМНИЙ доданок.
// Логіка перерахунку — точна копія v6.36 (clients.html), плюс ОКРУГЛЕННЯ ВГОРУ
// до тижня для списання ліміту (ceil(днів/5)*5) — узгоджено з карткою.
// ═══════════════════════════════════════════════════════════════════════════
function _pad2v(n){ return n < 10 ? '0' + n : '' + n; }

function _vacParseISO(s){
  if (s instanceof Date) return new Date(s.getFullYear(), s.getMonth(), s.getDate());
  var m = String(s || '').slice(0, 10).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  return m ? new Date(+m[1], +m[2] - 1, +m[3]) : null;
}
function _vacISOof(d){ return d.getFullYear() + '-' + _pad2v(d.getMonth() + 1) + '-' + _pad2v(d.getDate()); }

var _VAC_HOL_CACHE = {};
function _vacHolidaySet(year){
  if (_VAC_HOL_CACHE[year]) return _VAC_HOL_CACHE[year];
  var fixed = ['01-01','01-07','03-08','05-01','05-09','06-28','08-24','10-14','12-25'];
  var movable = {2025:['05-05','06-23'], 2026:['04-12','05-31']};   // Великдень, Трійця
  var s = {};
  fixed.forEach(function(dm){ s[year + '-' + dm] = 1; });
  (movable[year] || []).forEach(function(dm){ s[year + '-' + dm] = 1; });
  _VAC_HOL_CACHE[year] = s; return s;
}
function _vacIsWorkday(d){
  var wd = d.getDay();
  return wd !== 0 && wd !== 6 && !_vacHolidaySet(d.getFullYear())[_vacISOof(d)];
}
function _vacWorkDaysInMonth(year, month){
  var days = new Date(year, month, 0).getDate(), n = 0;
  for (var d = 1; d <= days; d++) if (_vacIsWorkday(new Date(year, month - 1, d))) n++;
  return n;
}
function _vacCountWorkDays(fromISO, toISO){
  var f = _vacParseISO(fromISO), t = _vacParseISO(toISO);
  if (!f || !t || t < f) return 0;
  var n = 0, cur = new Date(f.getTime());
  while (cur <= t){ if (_vacIsWorkday(cur)) n++; cur.setDate(cur.getDate() + 1); }
  return n;
}
function _vacContractYear(contractISO, refISO){
  var cd = _vacParseISO(contractISO); if (!cd) return null;
  var td = _vacParseISO(refISO) || new Date();
  var diffY = td.getFullYear() - cd.getFullYear();
  for (var off = diffY - 1; off <= diffY + 1; off++){
    var f = new Date(cd.getFullYear() + off, cd.getMonth(), cd.getDate());
    var t = new Date(cd.getFullYear() + off + 1, cd.getMonth(), cd.getDate() - 1);
    if (td >= f && td <= t) return {from: _vacISOof(f), to: _vacISOof(t)};
  }
  var f0 = new Date(cd.getFullYear() + diffY, cd.getMonth(), cd.getDate());
  var t0 = new Date(cd.getFullYear() + diffY + 1, cd.getMonth(), cd.getDate() - 1);
  return {from: _vacISOof(f0), to: _vacISOof(t0)};
}
// v6.40: списання ліміту ОКРУГЛЯЄТЬСЯ ВГОРУ до тижня для кожної відпустки
// (ceil(днів/5)*5). Різниця «згорає» (грошей не торкається). Узгоджено з карткою.
function _vacUsedInContractYear(absences, contractISO, refISO, excludeId){
  var w = _vacContractYear(contractISO, refISO); if (!w) return 0;
  var used = 0;
  (absences || []).forEach(function(a){
    if (!a || a.type !== 'vacation') return;
    if (excludeId && a.id === excludeId) return;
    if (a.status === 'rejected' || a.status === 'cancelled') return;
    if (!a.from || !a.to) return;
    var s = (a.from < w.from) ? w.from : a.from;
    var e = (a.to   > w.to)   ? w.to   : a.to;
    if (e < s) return;
    var wd = _vacCountWorkDays(s, e);
    used += Math.ceil(wd / 5) * 5;          // ⬅️ округлення вгору до тижня
  });
  return used;
}
function _vacLimitRemaining(absences, contractISO, refISO, excludeId){
  return Math.max(0, 20 - _vacUsedInContractYear(absences, contractISO, refISO, excludeId));
}
// Знижка по місяцях для ОДНОГО періоду — копія calcAbsMonthBreakdown(type='vacation').
// amount(discount) = round(fee × eligible_days / month_workdays); понад ліміт → 0.
function _vacMonthBreakdown(fromISO, toISO, fee, allAbsences, contractISO, selfId){
  var f = _vacParseISO(fromISO), t = _vacParseISO(toISO);
  if (!f || !t || t < f) return [];
  var mmap = {}, cur = new Date(f.getTime());
  while (cur <= t){
    if (_vacIsWorkday(cur)){
      var y = cur.getFullYear(), mo = cur.getMonth() + 1, mk = y + '-' + _pad2v(mo);
      if (!mmap[mk]) mmap[mk] = {ym: mk, y: y, m: mo, vacDays: 0};
      mmap[mk].vacDays++;
    }
    cur.setDate(cur.getDate() + 1);
  }
  var arr = Object.keys(mmap).sort().map(function(k){ return mmap[k]; });
  var remaining = _vacLimitRemaining(allAbsences, contractISO, fromISO, selfId);
  return arr.map(function(mi){
    var mwd = _vacWorkDaysInMonth(mi.y, mi.m);
    var eligible = Math.min(mi.vacDays, Math.max(0, remaining));
    remaining -= eligible;
    var discount = (fee > 0 && mwd > 0) ? Math.round(fee * eligible / mwd) : 0;
    return {ym: mi.ym, y: mi.y, m: mi.m, vacDays: mi.vacDays, monthWorkDays: mwd,
            eligibleDays: eligible, overLimitDays: mi.vacDays - eligible, discount: discount};
  });
}
// v6.58: ВИБІР ШКАЛИ хвороби за локацією + місяцем ПОЧАТКУ лікарняного (from).
// Перехід: усі локації — нова шкала з ЛИПНЯ 2026 (NEW_SICK_FROM); Бровари — вже з ЧЕРВНЯ.
// Нова шкала (роб.дні поспіль): <5=0%, 5-10=5%, 11-15=10%, 16-20=15%, 21+=20%.
// Стара: [0,5,10,15,20][min(4,ceil(днів/5))]. Стара гілка НЕ змінена (нульовий ризик).
var NEW_SICK_FROM = 202607;
var NEW_SICK_LOC_EXCEPTIONS = { 'Бровари': 202606 };
function _sickUseNew(loc, y, m){
  var ym = Number(y) * 100 + Number(m);
  if (ym >= NEW_SICK_FROM) return true;
  var ex = NEW_SICK_LOC_EXCEPTIONS[String(loc || '').trim()];
  return !!(ex && ym >= ex);
}
function _sickNewScalePct(days){
  if (days < 5)   return 0;
  if (days <= 10) return 5;
  if (days <= 15) return 10;
  if (days <= 20) return 15;
  return 20;
}

// ХВОРОБА: розбивка знижки по місяцях — дзеркало calcAbsMonthBreakdown(type='sick').
// Роб.дні Пн-Пт (свята НЕ виключаються). Вимір PER-RECORD (один запис поспіль).
// Шкала обирається _sickUseNew(loc, місяць_from). Стара гілка = попередня логіка 1-в-1.
function _sickMonthBreakdown(fromISO, toISO, fee, loc){
  var f = _vacParseISO(fromISO), t = _vacParseISO(toISO);
  if (!f || !t || t < f || !(fee > 0)) return [];
  var mmap = {}, cur = new Date(f.getTime());
  while (cur <= t){
    var wd = cur.getDay();
    if (wd !== 0 && wd !== 6){
      var y = cur.getFullYear(), mo = cur.getMonth() + 1, mk = y + '-' + _pad2v(mo);
      if (!mmap[mk]) mmap[mk] = {ym: mk, y: y, m: mo, workDays: 0, weeks: 0, discount: 0};
      mmap[mk].workDays++;
    }
    cur.setDate(cur.getDate() + 1);
  }
  var arr = Object.keys(mmap).sort().map(function(k){ return mmap[k]; });
  var totalDays = arr.reduce(function(s, x){ return s + x.workDays; }, 0);
  var useNew = _sickUseNew(loc, f.getFullYear(), f.getMonth() + 1);

  if (!useNew){
    // СТАРА шкала — без змін (розподіл тижнів «більший місяць першим»)
    var totalWeeks = Math.min(4, Math.ceil(totalDays / 5));
    var PCT = [0, 5, 10, 15, 20];
    var sorted = arr.slice().sort(function(a, b){ return b.workDays - a.workDays; });
    var weeksLeft = totalWeeks;
    for (var i = 0; i < sorted.length; i++){
      var mi = sorted[i];
      if (i === sorted.length - 1) mi.weeks = Math.max(0, weeksLeft);
      else { mi.weeks = Math.min(Math.ceil(mi.workDays / 5), weeksLeft); weeksLeft -= mi.weeks; }
      mi.discount = Math.round(fee * PCT[Math.min(4, mi.weeks)] / 100);
    }
  } else {
    // НОВА шкала — % від сумарних днів поспіль; discount рознесено по місяцях пропорційно
    // роб.дням (та сама механіка «більший місяць першим», залишок — останньому місяцю).
    var totalPct = _sickNewScalePct(totalDays);
    var totalDisc = Math.round(fee * totalPct / 100);
    var sortedN = arr.slice().sort(function(a, b){ return b.workDays - a.workDays; });
    var acc = 0;
    for (var j = 0; j < sortedN.length; j++){
      var mj = sortedN[j];
      mj.weeks = Math.min(4, Math.ceil(mj.workDays / 5));   // лише для відображення
      if (j === sortedN.length - 1) mj.discount = Math.max(0, totalDisc - acc);
      else { mj.discount = (totalDays > 0) ? Math.round(totalDisc * mj.workDays / totalDays) : 0; acc += mj.discount; }
    }
  }
  return arr;   // [{ym, y, m, workDays, weeks, discount}]
}
var _VAC_SCHOOL_LOCS = ['Школа Осокорки','Школа 228','Онлайн школа'];
// Дзеркало фронту (clients.html getContractType3): діти з договором ≥01.10.2025,
// яким зберігаємо відпустку як 'standard'. Матч по нормалізованому ПІБ (_normForMatch).
var _VAC_EXCEPTIONS = ['андреєва ангеліна','тандиряк северин','гаркуша богдан','мельничук дарина','скоріна аліса','городний яким',"щуров мар'ян"];
// preschool-відпустка зараховується як standard ЛИШЕ якщо весь період у літі
// (місяці from і to в межах 06–08) — дзеркало saveAbsencePeriod у clients.html.
function _vacIsSummerPeriod(fromISO, toISO){
  var mf = parseInt(String(fromISO || '').substr(5, 2), 10);
  var mt = parseInt(String(toISO   || '').substr(5, 2), 10);
  return mf >= 6 && mf <= 8 && mt >= 6 && mt <= 8;
}
function _vacContractType(contractISO, group, loc, name){
  if (name && _VAC_EXCEPTIONS.indexOf(_normForMatch(name)) !== -1) return 'standard'; // виняток — перед датою
  if (_VAC_SCHOOL_LOCS.indexOf(loc || '') !== -1) return 'school-no-absence';
  if (!contractISO) return 'standard';
  if (contractISO >= '2025-10-01') return 'new';                          // лише хвороба
  if (/preschool|розумник/i.test(String(group || ''))) return 'preschool'; // без перерахунків
  return 'standard';
}
function _vacContractISO(v){
  if (v instanceof Date) return _vacISOof(new Date(v.getFullYear(), v.getMonth(), v.getDate()));
  var s = String(v || '').slice(0, 10);
  return /^\d{4}-\d{2}-\d{2}$/.test(s) ? s : '';
}
function _vacParseAbsences(json){
  if (Array.isArray(json)) return json;
  try { return JSON.parse(String(json || '[]')) || []; } catch(e){ return []; }
}
function _normNameVac(s){ return String(s || '').replace(/[\s ]+/g, '').toLowerCase(); }
function _vacEntryDateISO(ts){
  if (ts instanceof Date) return _vacISOof(new Date(ts.getFullYear(), ts.getMonth(), ts.getDate()));
  var s = String(ts || '');
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
  var d = new Date(s);
  return isNaN(d.getTime()) ? '' : _vacISOof(d);
}
// Куди йде знижка за місяць відпустки: timestamp внесення < 1-ше число місяця
// відпустки → той самий місяць; >= 1-ше → наступний (рахунок уже виставлений).
function determineTargetMonth(vacationMonth, vacationYear, entryTimestamp){
  var entryISO = _vacEntryDateISO(entryTimestamp);
  if (!entryISO) return {month: vacationMonth, year: vacationYear, reason: 'no-timestamp→той самий місяць'};
  // v6.55: порівняння МІСЯЦІВ, не днів. Знижка йде в місяць відсутності, якщо внесено
  // у ТОЙ САМИЙ місяць або раніше (рахунок того місяця ще не виставлений). Якщо внесено
  // пізніше (наступний+ календарний місяць) — рахунок уже пішов → наступний місяць.
  // Фікс бага: раніше «entryISO < 1-ше число» кидало все, внесене 1-го числа, у наступний.
  var absKey   = vacationYear + '-' + _pad2v(vacationMonth);
  var entryKey = entryISO.slice(0, 7);
  if (entryKey <= absKey)  return {month: vacationMonth, year: vacationYear, reason: 'внесено ≤ місяця відсутності→той самий місяць'};
  var nm = _nextMonth(vacationMonth, vacationYear);
  return {month: nm.month, year: nm.year, reason: 'внесено пізніше→наступний місяць'};
}

function exportVacationDiscountToPayments(params){
  try {
    var loc   = String(params.loc || '').trim();
    var month = Number(params.month);
    var year  = Number(params.year) || new Date().getFullYear();
    if (!loc) return {ok: false, error: 'Параметр loc обовʼязковий'};
    if (!month || month < 1 || month > 12) return {ok: false, error: 'month має бути 1-12'};
    var monthName = MONTHS_CAL_UA[month - 1];

    // === 1. Клієнти локації → знижка, що РОУТИТЬСЯ в таргет (month,year) ===
    var cliRes = getClients();
    if (!cliRes.ok) return cliRes;
    var clients = (cliRes.data || []).filter(function(o){ return String(o['Локація'] || '').trim() === loc; });

    var discByNorm = {};   // normName → {name, discount(+)}
    var routing = [];
    clients.forEach(function(o){
      var name = trim(String(o['ПІБ дитини'] || ''));
      if (!name) return;
      var contractISO = _vacContractISO(o['Дата договору']);
      var fee = Number(o['Сума договору']) || 0;
      var ct  = _vacContractType(contractISO, o['Група'], loc, name);
      if (ct === 'school-no-absence') return;                     // школи — жодних знижок
      var allowVacation = (ct === 'standard' || ct === 'preschool'); // відпустка — лише standard/preschool
      var isPreschool   = (ct === 'preschool');
      // ХВОРОБА застосовна до ВСІХ типів окрім school (включно з 'new': договір ≥01.10.2025 = лише хвороба).
      var abs = _vacParseAbsences(o['Відсутності (JSON)']);

      // Спільний акумулятор: відпустка + хвороба сумуються в discByNorm[nk] →
      // одна клітинка «Бюджет навч», одна журнальна доріжка kind='vacation' (без зтирань).
      function _accrue(mb, kind, a){
        if (mb.discount <= 0) return;
        var tgt = determineTargetMonth(mb.m, mb.y, a.createdAt);
        if (tgt.year !== year || tgt.month !== month) return;
        var nk = _normNameVac(name);
        (discByNorm[nk] = discByNorm[nk] || {name: name, discount: 0}).discount += mb.discount;
        routing.push({child: name, kind: kind, vacMonth: mb.ym, discount: mb.discount,
                      routedTo: tgt.year + '-' + _pad2v(tgt.month), reason: tgt.reason});
      }

      abs.forEach(function(a){
        if (!a || (a.type !== 'vacation' && a.type !== 'sick')) return;
        if (a.status === 'rejected' || a.status === 'cancelled') return; // п.8
        if (!a.from || !a.to) return;
        if (a.type === 'vacation'){
          if (!allowVacation) return;                                       // new → відпустку не чіпаємо
          if (isPreschool && !_vacIsSummerPeriod(a.from, a.to)) return;      // preschool — лише літо
          _vacMonthBreakdown(a.from, a.to, fee, abs, contractISO, a.id).forEach(function(mb){ _accrue(mb, 'vacation', a); });
        } else { // sick — будь-який тип окрім school (вже відсіяли вище)
          _sickMonthBreakdown(a.from, a.to, fee, loc).forEach(function(mb){ _accrue(mb, 'sick', a); });
        }
      });
    });

    // === 2. Payment-файл локації ===
    var reg = _getLocationPaymentRegistry(loc);
    if (!reg || !reg.sheetId) return {ok: false, error: 'Локацію "' + loc + '" не знайдено в CONFIG-реєстрі'};
    var ss    = SpreadsheetApp.openById(reg.sheetId);
    var paySh = ss.getSheetByName(reg.sheetName) || ss.getSheets()[0];
    var data  = paySh.getDataRange().getValues();

    // === 3. Колонка "Бюджет навч" таргет-місяця + НАСТУПНОГО місяця (v6.41) ===
    var monthStartCol0  = 1 + (month - 1) * 5;
    var budgetNavchCol1 = monthStartCol0 + 4 + 1;   // 1-based (січень = F = 6)
    var nextNavchCol1   = budgetNavchCol1 + 5;       // той самий рядок, +1 місяць
    // Захист N+1 можливий лише в межах того ж річного листа (міс. 1-11).
    // Грудень → січень наступного року = інший лист/файл → захист пропускаємо.
    var canProtectNext  = (month <= 11);

    // === 4. Журнал kind='vacation' + значення/формули колонок ===
    var lastRow      = paySh.getLastRow();
    var colValues    = paySh.getRange(1, budgetNavchCol1, lastRow, 1).getValues();
    var colFormulas  = paySh.getRange(1, budgetNavchCol1, lastRow, 1).getFormulas();
    var nextFormulas = canProtectNext ? paySh.getRange(1, nextNavchCol1, lastRow, 1).getFormulas() : [];
    var journal      = _readJournalForTarget(loc, 'vacation', year, month);

    var DATA_START = 3, rowByNorm = {};
    for (var r = DATA_START; r < data.length; r++){
      var nm = trim(String(data[r][0] || ''));
      if (!nm || isGroupHeaderRow(data[r], 1)) continue;          // п.12 group-header
      var nk0 = _normNameVac(nm);
      if (!rowByNorm.hasOwnProperty(nk0)) rowByNorm[nk0] = r;
    }

    // === 5. Точковий запис через журнал (знижка = ВІД'ЄМНИЙ доданок) ===
    var journalOps = [], details = [], notFound = [];
    var cellsWritten = 0, formulaConverted = 0, nextProtected = 0, updated = 0, totalDisc = 0;

    Object.keys(rowByNorm).forEach(function(nk){
      var rowIdx0 = rowByNorm[nk];
      // v6.41: формульні рядки БІЛЬШЕ НЕ пропускаємо. getValues() дає обчислене
      // значення формули (напр. 21000), яке стає базою; setValue замінює формулу числом.
      var wasFormula   = !!(colFormulas[rowIdx0] && colFormulas[rowIdx0][0]);
      var payName      = trim(String(data[rowIdx0][0] || ''));
      var currentValue = Number(colValues[rowIdx0][0]) || 0;  // результат формули
      var je           = journal.byNormName[nk];
      var lastWritten  = je ? je.sum : 0;                  // попередня знижка (від'ємна)
      var baseValue    = currentValue - lastWritten;       // недисконтована база
      var disc         = discByNorm[nk] ? Math.round(discByNorm[nk].discount) : 0;
      var newSum       = disc > 0 ? -disc : 0;             // ⬅️ ЗМЕНШЕННЯ бюджету
      var newValue     = baseValue + newSum;

      if (newValue !== currentValue){
        paySh.getRange(rowIdx0 + 1, budgetNavchCol1).setValue(newValue);  // замінює формулу числом
        cellsWritten++;
        if (wasFormula) formulaConverted++;
      }

      // v6.41: ЗАХИСТ НАСТУПНОГО МІСЯЦЯ. Якщо пишемо знижку (disc>0) і N+1 містить
      // формулу (=N) — заморожуємо N+1 = baseValue (недисконтована, напр. 21000),
      // інакше N+1 успадкував би дисконтоване число. У журнал НЕ пишемо (технічна дія).
      if (disc > 0 && canProtectNext && nextFormulas[rowIdx0] && nextFormulas[rowIdx0][0]){
        paySh.getRange(rowIdx0 + 1, nextNavchCol1).setValue(baseValue);
        nextProtected++;
        Logger.log('[vacDiscount] PROTECT next-month row=%s "%s": формула %s → %s',
          rowIdx0 + 1, payName, nextFormulas[rowIdx0][0], baseValue);
      }

      if (newSum !== lastWritten){
        journalOps.push({nk: nk, loc: loc, kind: 'vacation', name: payName,
                         year: year, month: month, newSum: newSum});
      }
      if (disc > 0){
        updated++; totalDisc += disc;
        details.push({child: payName, discount: disc, currentBefore: currentValue,
                      wasFormula: wasFormula, base: baseValue, newCell: newValue,
                      row: rowIdx0 + 1, status: 'discounted'});
      } else if (lastWritten !== 0){                        // п.8 скасування → повертаємо бюджет
        details.push({child: payName, discount: 0, currentBefore: currentValue,
                      lastWritten: lastWritten, base: baseValue, newCell: newValue,
                      row: rowIdx0 + 1, status: 'restored'});
      }
    });

    Object.keys(discByNorm).forEach(function(nk){
      if (!rowByNorm.hasOwnProperty(nk)) notFound.push(discByNorm[nk].name);
    });

    _commitJournalUpdates(journal, journalOps);
    Logger.log('[vacDiscount] loc=%s %s/%s: updated=%s totalDisc=%s cells=%s formulaConv=%s nextProtected=%s notFound=%s',
      loc, month, year, updated, totalDisc, cellsWritten, formulaConverted, nextProtected, JSON.stringify(notFound));

    return {
      ok: true, loc: loc, targetMonth: monthName, targetMonthNum: month, year: year,
      budgetNavchCol: budgetNavchCol1, updated: updated, totalDiscount: totalDisc,
      cellsWritten: cellsWritten,
      formulaCellsConverted: formulaConverted,    // v6.41: формул→число у таргеті
      nextMonthProtected: nextProtected,          // v6.41: скільки N+1 заморожено
      nextMonthProtectable: canProtectNext,       // false для грудня
      notFound: notFound, details: details, routing: routing
    };
  } catch(e){
    Logger.log('[exportVacationDiscount] EXCEPTION: %s\n%s', e && e.message, e && e.stack);
    return {ok: false, error: String(e && e.message || e)};
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// ONE-SHOT: перенос знижок Борщагівки з СЕРПНЯ 2026 у ЛИПЕНЬ 2026 після фікса
// determineTargetMonth (баг 1-го числа). Крок 1 — ре-експорт серпня (нова логіка
// віднесе липневі відпустки до липня → серпень обнулиться/відновиться). Крок 2 —
// ре-експорт липня (знижки ляжуть у липень). Ідемпотентно: журнал тримає
// last_written_sum на (loc,kind,рік,місяць) → повторний запуск не подвоює.
// Run у редакторі БЕЗ параметрів. Не чіпає інші локації.
// ═══════════════════════════════════════════════════════════════════════════
function reexportBorshchahivka0708(){
  var LOC = 'Борщагівка', YEAR = 2026;
  var WATCH = ['Щуров Марʼян', 'Петрів Олександр', 'Савчук Марія'];

  Logger.log('═══ РЕ-ЕКСПОРТ %s: серпень→липень %s ═══', LOC, YEAR);

  Logger.log('--- Крок 1: ре-експорт СЕРПНЯ (8) — обнулення/відновлення ---');
  var aug = exportVacationDiscountToPayments({ loc: LOC, month: 8, year: YEAR });
  Logger.log('серпень: ok=%s updated=%s totalDiscount=%s cells=%s notFound=%s',
    aug.ok, aug.updated, aug.totalDiscount, aug.cellsWritten, JSON.stringify(aug.notFound));

  Logger.log('--- Крок 2: ре-експорт ЛИПНЯ (7) — запис знижок ---');
  var jul = exportVacationDiscountToPayments({ loc: LOC, month: 7, year: YEAR });
  Logger.log('липень: ok=%s updated=%s totalDiscount=%s cells=%s notFound=%s',
    jul.ok, jul.updated, jul.totalDiscount, jul.cellsWritten, JSON.stringify(jul.notFound));

  // --- Звірка по журналу для 3 цільових дітей ---
  Logger.log('--- ЗВІРКА журналу (Експорт_Журнал, kind=vacation) ---');
  var jAug = _readJournalForTarget(LOC, 'vacation', YEAR, 8);
  var jJul = _readJournalForTarget(LOC, 'vacation', YEAR, 7);
  WATCH.forEach(function(nm){
    var nk = _journalNormName(nm);
    var a = jAug.byNormName[nk], j = jJul.byNormName[nk];
    Logger.log('  %s → серпень=%s | липень=%s  %s',
      nm,
      a ? a.sum : '(нема)',
      j ? j.sum : '(нема)',
      (a && a.sum === 0 && j && j.sum < 0) ? '✅ перенесено' : '⚠️ перевір');
  });
  Logger.log('Очікувано: серпень=0, липень<0 (Щуров −9130, Петрів −4565, Савчук −9130).');

  return { ok: (aug.ok && jul.ok), august: aug, july: jul };
}

// ═══════════════════════════════════════════════════════════════════════════
// ONE-SHOT: Бровари — застосувати НОВУ шкалу хвороби до ЧЕРВНЕВИХ лікарняних
// (виняток переходу). Run у редакторі БЕЗ параметрів. Ідемпотентно (журнал).
// Прогоняє цільові місяці 7 і 6 2026 (куди маршрутизуються червневі знижки),
// оновлює агрегат і друкує порівняння стара→нова по кожному червневому лікарняному.
// ═══════════════════════════════════════════════════════════════════════════
function reexportBrovaryNewSickScale(){
  var LOC = 'Бровари', YEAR = 2026;
  Logger.log('═══ РЕ-ЕКСПОРТ Бровари: нова шкала хвороби (черв.+лип. %s) ═══', YEAR);

  var jul = exportVacationDiscountToPayments({ loc: LOC, month: 7, year: YEAR });
  var jun = exportVacationDiscountToPayments({ loc: LOC, month: 6, year: YEAR });
  Logger.log('липень: ok=%s updated=%s totalDiscount=%s', jul.ok, jul.updated, jul.totalDiscount);
  Logger.log('червень: ok=%s updated=%s totalDiscount=%s', jun.ok, jun.updated, jun.totalDiscount);

  // застосовані знижки-хвороби з routing обох прогонів
  Logger.log('--- застосовані ЛІКАРНЯНІ (routing kind=sick) ---');
  [{ r: jul, tgt: 7 }, { r: jun, tgt: 6 }].forEach(function(o){
    (o.r.routing || []).filter(function(x){ return x.kind === 'sick'; }).forEach(function(x){
      Logger.log('  [ціль %s] %s | хвороба міс.%s | знижка=%s → %s', o.tgt, x.child, x.vacMonth, x.discount, x.routedTo);
    });
  });

  // порівняння стара→нова по ЧЕРВНЕВИХ лікарняних Бровар
  Logger.log('--- ПОРІВНЯННЯ стара→нова (червневі лікарняні) ---');
  function _oldSickPct(d){ return [0, 5, 10, 15, 20][Math.min(4, Math.ceil(d / 5))]; }
  function _wdCount(fromISO, toISO){
    var f = _vacParseISO(fromISO), t = _vacParseISO(toISO);
    if (!f || !t || t < f) return 0;
    var n = 0, c = new Date(f.getTime());
    while (c <= t){ var w = c.getDay(); if (w !== 0 && w !== 6) n++; c.setDate(c.getDate() + 1); }
    return n;
  }
  var cliRes = getClients();
  var clients = (cliRes.data || []).filter(function(o){ return String(o['Локація'] || '').trim() === LOC; });
  var cnt = 0;
  clients.forEach(function(o){
    var name = trim(String(o['ПІБ дитини'] || '')), fee = Number(o['Сума договору']) || 0;
    _vacParseAbsences(o['Відсутності (JSON)']).forEach(function(a){
      if (!a || a.type !== 'sick' || a.status === 'rejected' || a.status === 'cancelled' || !a.from || !a.to) return;
      if (String(a.from).slice(0, 7) !== '2026-06') return;                 // лише червневі
      var d = _wdCount(a.from, a.to), op = _oldSickPct(d), np = _sickNewScalePct(d);
      cnt++;
      Logger.log('  %s | %s..%s | роб.дн=%s | стара %s%% (%s) → нова %s%% (%s)',
        name, a.from, a.to, d, op, Math.round(fee * op / 100), np, Math.round(fee * np / 100));
    });
  });
  if (!cnt) Logger.log('  (червневих лікарняних не знайдено)');

  var agg = aggregatePaymentsYearly();
  Logger.log('aggregatePaymentsYearly: ok=%s', agg && agg.ok);
  Logger.log('═══ Готово. Перевір «Оплати» Бровар. ═══');
  return { ok: (jul.ok && jun.ok), july: jul, june: jun };
}

// ═══════════════════════════════════════════════════════════════════════════
// ФІКС зламаної бази «Бюджет навч» липня 2026 у 3 дітей Борщагівки.
// Причина: знижка коректна, але база (повна ціна) у липневій комірці була
// порожня (Савчук → −9130) або не оновилась/перекрита (Дандикіна ×2 → лишилось
// 18900 без знижки). Ставимо ФІНАЛЬНЕ значення (ціна − знижка) напряму.
// Узгоджено з журналом: майбутній ре-експорт рахує база = комірка − остання_знижка
// (11870−(−9130)=21000; 10683−(−8217)=18900) → значення лишається стабільним.
//
// СПОЧАТКУ запусти diagBorshchahivkaJuly() (тільки читання) — покаже рядки,
// поточні значення, формули й ДУБЛІ. Потім fixBorshchahivkaJulyBudget().
// ═══════════════════════════════════════════════════════════════════════════
function _borshJulyBudgetCol1(){ var monthStartCol0 = 1 + (7 - 1) * 5; return monthStartCol0 + 4 + 1; } // =36
var _BORSH_JULY_TARGETS = [
  { name: 'Савчук Марія',     value: 11870 },  // 21000 − 9130
  { name: 'Дандикіна Ніколь', value: 10683 },  // 18900 − 8217
  { name: 'Дандикіна Емілія', value: 10683 }   // 18900 − 8217
];

function diagBorshchahivkaJuly(){
  var reg = _getLocationPaymentRegistry('Борщагівка');
  if (!reg || !reg.sheetId) { Logger.log('❌ Борщагівка не знайдена в CONFIG'); return {ok:false}; }
  var ss = SpreadsheetApp.openById(reg.sheetId);
  var sh = ss.getSheetByName(reg.sheetName) || ss.getSheets()[0];
  var data = sh.getDataRange().getValues();
  var col1 = _borshJulyBudgetCol1();
  var formulas = sh.getRange(1, col1, sh.getLastRow(), 1).getFormulas();
  var want = {}; _BORSH_JULY_TARGETS.forEach(function(t){ want[_normNameVac(t.name)] = t; });
  Logger.log('═══ ДІАГНОСТИКА Борщагівка ЛИПЕНЬ «Бюджет навч» (кол.%s) — тільки читання ═══', col1);
  var found = {};
  for (var r = 2; r < data.length; r++){
    var nm = String(data[r][0] || '').trim();
    if (!nm) continue;
    var nk = _normNameVac(nm);
    if (!want[nk]) continue;
    found[nk] = (found[nk] || 0) + 1;
    Logger.log('  row %s | "%s" | значення=%s | формула=%s | має бути=%s%s',
      r + 1, nm, JSON.stringify(data[r][col1 - 1]),
      (formulas[r] && formulas[r][0]) ? formulas[r][0] : '(число)',
      want[nk].value, found[nk] > 1 ? '  ⚠️ ДУБЛЬ-РЯДОК' : '');
  }
  _BORSH_JULY_TARGETS.forEach(function(t){ if (!found[_normNameVac(t.name)]) Logger.log('  ❌ НЕ ЗНАЙДЕНО: %s', t.name); });
  Logger.log('Далі: fixBorshchahivkaJulyBudget()');
  return { ok: true, found: found };
}

function fixBorshchahivkaJulyBudget(){
  var reg = _getLocationPaymentRegistry('Борщагівка');
  if (!reg || !reg.sheetId) { Logger.log('❌ Борщагівка не знайдена в CONFIG'); return {ok:false}; }
  var ss = SpreadsheetApp.openById(reg.sheetId);
  var sh = ss.getSheetByName(reg.sheetName) || ss.getSheets()[0];
  var data = sh.getDataRange().getValues();
  var col1 = _borshJulyBudgetCol1();
  var want = {}; _BORSH_JULY_TARGETS.forEach(function(t){ want[_normNameVac(t.name)] = t; });
  var seen = {}, changed = [];
  Logger.log('═══ ФІКС Борщагівка ЛИПЕНЬ «Бюджет навч» (кол.%s) ═══', col1);
  for (var r = 2; r < data.length; r++){
    var nm = String(data[r][0] || '').trim();
    if (!nm) continue;
    var nk = _normNameVac(nm);
    if (!want[nk]) continue;
    seen[nk] = (seen[nk] || 0) + 1;
    if (seen[nk] > 1){
      Logger.log('  ⚠️ ДУБЛЬ "%s" row %s — НЕ чіпаю (перевір вручну)', nm, r + 1);
      continue;                                          // тільки перший (авторитетний) рядок
    }
    var before = data[r][col1 - 1];
    if (Number(before) === want[nk].value){
      Logger.log('  = "%s" row %s вже %s — пропуск', nm, r + 1, want[nk].value);
    } else {
      sh.getRange(r + 1, col1).setValue(want[nk].value);  // число, замінює формулу якщо була
      changed.push({ name: nm, row: r + 1, before: before, after: want[nk].value });
      Logger.log('  ✔ "%s" row %s: %s → %s', nm, r + 1, JSON.stringify(before), want[nk].value);
    }
  }
  _BORSH_JULY_TARGETS.forEach(function(t){ if (!seen[_normNameVac(t.name)]) Logger.log('  ❌ НЕ ЗНАЙДЕНО (не виправлено): %s', t.name); });
  Logger.log('═══ Готово: змінено %s комірок. Прогони aggregatePaymentsYearly() для оновлення «Оплати». ═══', changed.length);
  return { ok: true, changed: changed, seen: seen };
}

// ═══════════════════════════════════════════════════════════════════════════
// УНІВЕРСАЛЬНІ РУЧНІ КОРИГУВАННЯ переплат/боргів за додаткові — «Бюджет доп».
// Директори дають коригування по своїх локаціях; вписуй усі рядки в DOP_OVERPAY_LIST
// (або передай масив у applyDopOverpay(list)) і запускай РАЗ.
//   amount ВІД'ЄМНЕ  = переплата (менше до сплати),  amount ДОДАТНЕ = борг (більше).
//   child — написання ПІБ ЯК У Payment-файлі локації.
// Механізм (як Мандзюватий -50): amount входить у БАЗУ комірки «Бюджет доп» місяця.
// exportAttendanceToPayments рахує база = комірка - остання_знижка(kind='payment') і додає
// допи ЗВЕРХУ → коригування зберігається при кожному перерахунку допів.
// Ідемпотентно + РЕДАГОВАНО: журнал kind='dop-manual' тримає застосований amount. Повтор
// із тим самим amount → нічого; зі зміненим → застосує лише ДЕЛЬТУ (amount - попередній).
// (exportAttendanceToPayments читає лише kind='payment', тож dop-manual його не зачіпає.)
// Спершу diagDopOverpay() (тільки читання) → потім applyDopOverpay().
// ═══════════════════════════════════════════════════════════════════════════
var DOP_ADJ_YEAR = 2026, DOP_ADJ_MONTH = 7;    // місяць коригувань (за потреби зміни)
var DOP_OVERPAY_LIST = [
  { loc: 'Оранж', child: 'Пашаев Емир', amount: -500 },
  { loc: 'Оранж', child: 'Міщенко Лев', amount: -400 }
  // ↓ додавай інші локації сюди, напр.:
  // { loc: 'Позняки', child: 'ПІБ як у Payment', amount: -300 },
  // { loc: 'Бровари', child: 'ПІБ як у Payment', amount:  250 },
];
function _dopNorm(s){ return String(s || '').replace(/[\s ]+/g, '').toLowerCase(); }
function _dopBudgetCol1(month){ var monthStartCol0 = 1 + (month - 1) * 5; return monthStartCol0 + 3 + 1; }

function _dopOverpayRun(list, apply){
  list = list || [];
  var col1 = _dopBudgetCol1(DOP_ADJ_MONTH);
  var monthName = MONTHS_CAL_UA[DOP_ADJ_MONTH - 1];
  Logger.log('=== %s коригувань «Бюджет доп» %s %s (кол.%s) — %s рядків ===',
    apply ? 'ФІКС' : 'ДІАГ (тільки читання)', monthName, DOP_ADJ_YEAR, col1, list.length);

  // групуємо по локації — кожен Payment-файл відкриваємо один раз
  var byLoc = {};
  list.forEach(function(it){ var L = String(it.loc || '').trim(); (byLoc[L] = byLoc[L] || []).push(it); });

  var totalChanged = 0, results = [];
  Object.keys(byLoc).forEach(function(loc){
    var reg = _getLocationPaymentRegistry(loc);
    if (!reg || !reg.sheetId){ Logger.log('  [X] локацію "%s" не знайдено в CONFIG — пропуск', loc); return; }
    var ss = SpreadsheetApp.openById(reg.sheetId);
    var sh = ss.getSheetByName(reg.sheetName) || ss.getSheets()[0];
    var data = sh.getDataRange().getValues();
    var formulas = sh.getRange(1, col1, sh.getLastRow(), 1).getFormulas();
    var journal = _readJournalForTarget(loc, 'dop-manual', DOP_ADJ_YEAR, DOP_ADJ_MONTH);

    // індекс рядків по нормалізованому імені (перший збіг, як exportAttendanceToPayments)
    var rowByNorm = {};
    for (var r = 2; r < data.length; r++){
      var nm0 = trim(String(data[r][0] || ''));
      if (!nm0 || isGroupHeaderRow(data[r], 1)) continue;
      var k0 = _dopNorm(nm0);
      if (!rowByNorm.hasOwnProperty(k0)) rowByNorm[k0] = r;
    }

    var ops = [];
    byLoc[loc].forEach(function(it){
      var child  = String(it.child || '').trim();
      var amount = Number(it.amount) || 0;
      var nk = _dopNorm(child);
      if (!rowByNorm.hasOwnProperty(nk)){
        Logger.log('  [X] [%s] НЕ ЗНАЙДЕНО у Payment: "%s"', loc, child);
        results.push({ loc: loc, child: child, status: 'not-found' }); return;
      }
      var r = rowByNorm[nk];
      var cur = Number(data[r][col1 - 1]) || 0;
      var fml = (formulas[r] && formulas[r][0]) ? formulas[r][0] : '(число)';
      var jkey = _journalNormName(child);
      var prev = journal.byNormName[jkey] ? journal.byNormName[jkey].sum : 0;   // раніше застосований amount
      var delta = amount - prev;
      if (delta === 0){
        Logger.log('  [=] [%s] "%s" row %s: amount=%s вже застосовано — пропуск (комірка=%s)', loc, data[r][0], r + 1, amount, cur);
        results.push({ loc: loc, child: child, row: r + 1, current: cur, amount: amount, status: 'unchanged' }); return;
      }
      var next = cur + delta;
      Logger.log('  %s [%s] "%s" row %s: комірка=%s (формула=%s) | було=%s -> нове=%s (delta %s) -> комірка=%s',
        apply ? '[v]' : '[.]', loc, data[r][0], r + 1, cur, fml, prev, amount, delta, next);
      if (apply){
        sh.getRange(r + 1, col1).setValue(next);
        ops.push({ nk: jkey, loc: loc, kind: 'dop-manual', name: String(data[r][0]).trim(),
                   year: DOP_ADJ_YEAR, month: DOP_ADJ_MONTH, newSum: amount });
        totalChanged++;
      }
      results.push({ loc: loc, child: child, row: r + 1, before: cur, after: next, amount: amount, prevAdj: prev, status: apply ? 'applied' : 'preview' });
    });
    if (apply && ops.length){ _commitJournalUpdates(journal, ops); Logger.log('  [%s] журнал dop-manual оновлено (%s)', loc, ops.length); }
  });

  Logger.log('=== Готово%s. Далі: aggregatePaymentsYearly() для оновлення «Оплати». ===', apply ? (' — змінено ' + totalChanged) : ' (тільки читання)');
  return { ok: true, changed: totalChanged, results: results };
}
function diagDopOverpay(list){ return _dopOverpayRun(list || DOP_OVERPAY_LIST, false); }
function applyDopOverpay(list){ return _dopOverpayRun(list || DOP_OVERPAY_LIST, true); }

// ═══════════════════════════════════════════════════════════════════════════
// ДО-ЕКСПОРТ ЗНИЖОК ВІДПУСТОК — ЛИПЕНЬ 2026 (one-shot backfill).
// Причина: markAbsDone (погодження) НЕ тригерить експорт знижки — він іде лише
// при saveAbsencePeriod (додавання). Тому вже погоджені (status='done') липневі
// відпустки могли не потрапити в Payment. Ці функції перераховують їх через
// ВИПРАВЛЕНИЙ _vacContractType (винятки + preschool-літо) і дозаписують.
// Запускати ВРУЧНУ з Apps Script editor (View → Executions → Логи).
// ═══════════════════════════════════════════════════════════════════════════
var _REEXPORT_VAC_YEAR  = 2026;
var _REEXPORT_VAC_MONTH = 7;

// Сканує ВСІХ дітей із vacation-відсутністю status='done' у липні 2026.
// Повертає [{name, loc, group, fee, ct, julyDiscount(+)}] — знижка, що РОУТИТЬСЯ
// в липень 2026 (з урахуванням determineTargetMonth) лише по done-відпустках.
function _reexportVacJulyScan(){
  var Y = _REEXPORT_VAC_YEAR, M = _REEXPORT_VAC_MONTH;
  var first = Y + '-' + _pad2v(M) + '-01';
  var last  = Y + '-' + _pad2v(M) + '-31';
  var cliRes = getClients();
  if (!cliRes.ok) return [];
  var out = [];
  (cliRes.data || []).forEach(function(o){
    var name = trim(String(o['ПІБ дитини'] || ''));
    if (!name) return;
    var loc  = trim(String(o['Локація'] || ''));
    var abs  = _vacParseAbsences(o['Відсутності (JSON)']);
    // тільки ПОГОДЖЕНІ (done) відпустки з періодом, що зачіпає липень 2026
    var doneJuly = abs.filter(function(a){
      return a && a.type === 'vacation' && a.status === 'done'
        && a.from && a.to && a.from <= last && a.to >= first;
    });
    if (!doneJuly.length) return;

    var contractISO = _vacContractISO(o['Дата договору']);
    var fee = Number(o['Сума договору']) || 0;
    var ct  = _vacContractType(contractISO, o['Група'], loc, name);

    var julyDisc = 0;
    if (ct === 'standard' || ct === 'preschool'){
      var isPre = (ct === 'preschool');
      doneJuly.forEach(function(a){
        if (isPre && !_vacIsSummerPeriod(a.from, a.to)) return;   // preschool — лише літо
        _vacMonthBreakdown(a.from, a.to, fee, abs, contractISO, a.id).forEach(function(mb){
          if (mb.discount <= 0) return;
          var tgt = determineTargetMonth(mb.m, mb.y, a.createdAt);
          if (tgt.year === Y && tgt.month === M) julyDisc += mb.discount;   // роут саме в липень
        });
      });
    }
    out.push({name: name, loc: loc, group: o['Група'] || '', fee: fee,
              ct: ct, julyDiscount: Math.round(julyDisc)});
  });
  return out;
}

// DRY-RUN: лише Logger.log, НІЧОГО не пише.
function reexportVacDiscountsJulyDry(){
  var Y = _REEXPORT_VAC_YEAR, M = _REEXPORT_VAC_MONTH;
  var list = _reexportVacJulyScan();
  Logger.log('═══ DRY-RUN: до-експорт знижок відпусток %s/%s — НІЧОГО НЕ ПИШЕМО ═══', M, Y);
  Logger.log('Формат: ПІБ | локація | знижка | вже в Payment? | дія');

  var journalCache = {};
  var willWrite = 0, already = 0, skipFee0 = 0, skipZero = 0, totalNew = 0;
  list.forEach(function(r){
    if (!journalCache.hasOwnProperty(r.loc)){
      journalCache[r.loc] = _readJournalForTarget(r.loc, 'vacation', Y, M).byNormName || {};
    }
    var je = journalCache[r.loc][_normNameVac(r.name)];
    var inPayment = !!(je && je.sum);                      // sum < 0 якщо знижка вже записана

    var action;
    if (r.julyDiscount <= 0){
      if (r.fee <= 0){ action = 'ПРОПУСК (fee=0)'; skipFee0++; }
      else           { action = 'ПРОПУСК (знижка=0)'; skipZero++; }
    } else if (inPayment){
      action = 'ВЖЕ Є (journal ' + je.sum + ') — не дублюю'; already++;
    } else {
      action = 'ЗАПИШЕТЬСЯ −' + r.julyDiscount; willWrite++; totalNew += r.julyDiscount;
    }
    Logger.log('%s | %s | %s | %s | %s', r.name, r.loc, r.julyDiscount, inPayment ? 'так' : 'ні', action);
  });
  Logger.log('─── ПІДСУМОК DRY: дітей=%s | запишеться=%s (−%s грн) | вже є=%s | пропуск fee=0=%s | пропуск знижка=0=%s ───',
    list.length, willWrite, totalNew, already, skipFee0, skipZero);
  return {ok: true, year: Y, month: M, total: list.length, willWrite: willWrite,
          totalNewDiscount: totalNew, alreadyInPayment: already, skippedFee0: skipFee0, skippedZero: skipZero};
}

// APPLY: РЕАЛЬНО пише знижку в Payment 'Бюджет навч' липня.
// Анти-дублювання: делегує в exportVacationDiscountToPayments(loc,7,2026) — той
// рахує дельту через журнал (kind='vacation'): baseValue = поточне − попередньо_записане,
// тож повторний запуск ідемпотентний (вже записане НЕ дублюється).
function reexportVacDiscountsJulyApply(){
  var Y = _REEXPORT_VAC_YEAR, M = _REEXPORT_VAC_MONTH;
  var list = _reexportVacJulyScan();

  // унікальні локації, де є що писати (знижка > 0)
  var locSet = {};
  list.forEach(function(r){ if (r.julyDiscount > 0) locSet[r.loc] = true; });
  var locs = Object.keys(locSet);
  Logger.log('═══ APPLY: до-експорт знижок відпусток %s/%s — локацій=%s ═══', M, Y, locs.length);
  Logger.log('Формат: ПІБ | локація | знижка | дія');

  // знімок журналу ДО — щоб відрізнити «вже було» від «щойно записано»
  var before = {};
  locs.forEach(function(loc){
    before[loc] = _readJournalForTarget(loc, 'vacation', Y, M).byNormName || {};
  });

  var written = 0, already = 0, restored = 0, totalWritten = 0, errs = [];
  locs.forEach(function(loc){
    var res = exportVacationDiscountToPayments({loc: loc, month: M, year: Y});
    if (!res || !res.ok){
      errs.push(loc + ': ' + (res && res.error || '?'));
      Logger.log('❌ %s — помилка експорту: %s', loc, res && res.error);
      return;
    }
    (res.details || []).forEach(function(d){
      var wasThere = !!(before[loc][_normNameVac(d.child)] || {}).sum;
      if (d.status === 'discounted'){
        if (wasThere){ already++; Logger.log('%s | %s | −%s | ВЖЕ БУЛО (без змін)', d.child, loc, d.discount); }
        else { written++; totalWritten += d.discount; Logger.log('%s | %s | −%s | ✅ ЗАПИСАНО', d.child, loc, d.discount); }
      } else if (d.status === 'restored'){
        restored++; Logger.log('%s | %s | бюджет ВІДНОВЛЕНО (знижку знято)', d.child, loc);
      }
    });
    (res.notFound || []).forEach(function(nm){ Logger.log('⚠️ %s | %s | НЕ ЗНАЙДЕНО рядок у Payment-файлі', nm, loc); });
  });
  Logger.log('─── ПІДСУМОК APPLY: записано нових=%s (−%s грн) | вже було=%s | відновлено=%s | помилок=%s ───',
    written, totalWritten, already, restored, errs.length);
  return {ok: errs.length === 0, year: Y, month: M, locations: locs,
          written: written, totalWritten: totalWritten, alreadyInPayment: already,
          restored: restored, errors: errs};
}

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
// v6.66: safe salary activity<->row matcher (does NOT touch _journalNormName / journal keys)
function _salaryFold2(x){
  return String(x || '')
    .replace(/['\u02BC`\u2018\u2019]/g, '')
    .replace(/\u0456/g, '\u0438')
    .replace(/\u0457/g, '\u0438')
    .replace(/\u0454/g, '\u0435')
    .replace(/\u0451/g, '\u0435');
}
var SALARY_ALIASES = {
  '\u0430\u043d\u0433\u043b\u0456\u0439\u0441\u044c\u043a\u0430\u0433\u0440\u0443\u043f\u043e\u0432\u0456': ['\u0430\u043d\u0433\u043b\u0456\u0439\u0441\u044c\u043a\u0430', '\u0430\u043d\u0433\u043b\u0456\u0439\u0441\u044c\u043a\u0430\u043c\u043e\u0432\u0430', 'englishtoschool', 'english'],
  '\u0456\u043d\u0434\u0438\u0432\u0456\u0434\u0443\u0430\u043b\u044c\u043d\u0456\u0437\u043b\u043e\u0433\u043e\u043f\u0435\u0434\u043e\u043c': ['\u043b\u043e\u0433\u043e\u043f\u0435\u0434']
};
function _salaryMatchRow(lname, actRowByLname, foldedRowMap){
  if (actRowByLname[lname] > 0) return actRowByLname[lname];
  var f = _salaryFold2(lname);
  if (foldedRowMap[f] > 0) return foldedRowMap[f];
  var aliases = SALARY_ALIASES[lname];
  if (aliases){
    for (var i = 0; i < aliases.length; i++){
      var al = aliases[i];
      if (actRowByLname[al] > 0) return actRowByLname[al];
      var af = _salaryFold2(al);
      if (foldedRowMap[af] > 0) return foldedRowMap[af];
    }
  }
  return -1;
}

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

    // v7.08: обʼєднання груп (session-key = група×дата). Мапа loc→act→date→[merge-набори].
    var mergesMap = _loadDopMergesMap(loc, dateFrom, dateTo);

    var byActId = {};
    for (var i = 1; i < attData.length; i++){
      var rec = _parseAttendanceRow(attData[i]);
      if (rec.loc !== loc) continue;
      if (rec.date < dateFrom || rec.date >= dateTo) continue;
      if (!byActId[rec.activityId]) byActId[rec.activityId] = {count: 0, dates: {}, groupsByDate: {}};
      byActId[rec.activityId].count++;
      byActId[rec.activityId].dates[rec.date] = true;
      // session-key: множина нормалізованих груп у кожну дату (порожня група → '' — 1 сесія/дату, як раніше)
      var _ng = _dopNormGroup(rec.group);
      if (!byActId[rec.activityId].groupsByDate[rec.date]) byActId[rec.activityId].groupsByDate[rec.date] = {};
      byActId[rec.activityId].groupsByDate[rec.date][_ng] = true;
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
      var stat = byActId[a.id] || {count: 0, dates: {}, groupsByDate: {}};
      var fact = 0;
      if (a.teacherModel === 'За дитину'){
        fact = stat.count * a.teacherRate;
      } else if (a.teacherModel === 'За заняття' || a.teacherModel === 'За захід'){
        // v7.08: "За заняття"/"За захід" — фіксована сума за кожну СЕСІЮ.
        // Сесія = унікальна (група × дата); обʼєднані одного дня групи схлопуються
        // у 1 сесію (мапа mergesMap). Раніше рахувались лише унікальні дати — це
        // недоплачувало викладачам, що ведуть кілька РІЗНИХ груп одного дня.
        // "За захід" рахується ідентично (Театр/Майстер-клас/вистави).
        fact = _dopCountSessions(stat.groupsByDate, mergesMap[a.id] || {}) * a.teacherRate;
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

    // v6.25: Карта нормалізована-назва → 1-based row будується ТІЛЬКИ з рядків,
    // у яких _section === 'extras'. Інакше "Логопед" у секції "Вчителі-предметники"
    // (рядок ~12) переможе "Логопед" у "Додаткові заняття" (рядок ~50), і ставка
    // з Додаткові_Каталог запишеться у predmetnyky-рядок. Класифікація через
    // ту саму state machine, що й у getSalaryData / diagSalaryExtrasSections.
    var lastCol = Math.max(sheet.getLastColumn(), 37);
    var fullData = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    var rawRowsForClassify = [];
    for (var rN = 4; rN <= fullData.length; rN++){
      var rIdx = rN - 1;
      var rArr = fullData[rIdx] || [];
      var rName = String(rArr[0] || '').trim();
      if (_salaryIsSkippedRow(rName)) continue;
      var rFact = 0, rBudget = 0;
      for (var mN = 1; mN <= 12; mN++){
        var fI = (mN - 1) * 3 + 1;
        var bI = (mN - 1) * 3 + 2;
        if (fI < lastCol) rFact   += _opexNum(rArr[fI]);
        if (bI < lastCol) rBudget += _opexNum(rArr[bI]);
      }
      rawRowsForClassify.push({row: rN, name: rName, fact: rFact, budget: rBudget});
    }
    var classifiedRows = _classifyAllSalaryRows(rawRowsForClassify);

    var actRowByLname = {};
    classifiedRows.forEach(function(cr){
      if (cr._section !== 'extras') return;
      if (cr._category === 'section_header' || cr._category === 'group_header') return;
      var rname = _journalNormName(cr.name);
      if (rname && !actRowByLname.hasOwnProperty(rname)){
        actRowByLname[rname] = cr.row;
      }
    });
    Logger.log('[exportToSalaryExtras] actRowByLname (тільки extras-секція): %s', JSON.stringify(actRowByLname));
    var _salaryFoldedRowMap = {};
    Object.keys(actRowByLname).forEach(function(k){
      var ff = _salaryFold2(k);
      if (!_salaryFoldedRowMap.hasOwnProperty(ff)) _salaryFoldedRowMap[ff] = actRowByLname[k];
    });

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
      var rowFound = _salaryMatchRow(lname, actRowByLname, _salaryFoldedRowMap);
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
      var info         = factByName[lname]; // може бути undefined якщо у заняття немає rate
      var newFact      = info ? info.fact : 0;
      // v6.13: OVERWRITE замість delta. Клітинка Salary "<Activity>" належить
      // виключно activities — забруднених baseValues нема (як і в предметниках).
      var newValue     = newFact;

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
          newCell: newValue, row: rowFound, status: 'updated'
        });
        Logger.log('[exportToSalaryExtras] WRITE row=%s "%s" cur=%s last=%s newFact=%s → %s', rowFound, a.name, currentValue, lastWritten, newFact, newValue);
      } else if (lastWritten !== 0){
        details.push({
          activity: a.name, fact: 0,
          currentBefore: currentValue, lastWritten: lastWritten,
          newCell: newValue, row: rowFound, status: 'cleared'
        });
        Logger.log('[exportToSalaryExtras] CLEAR row=%s "%s" cur=%s last=%s → %s', rowFound, a.name, currentValue, lastWritten, newValue);
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

// ═══════════════════════════════════════════════════════════════════════════
// ПЕРЕРАХУНОК ДОДАТКОВИХ ЗА ЧЕРВЕНЬ 2026 ПІСЛЯ ЧИСТКИ ДУБЛІВ — обгортки.
// Дублі (143 рядки) видалені з Додаткові_Відвідуваність; суми в Payment/Salary
// лишились завищені. Перезапуск експортів перепише їх через журнал (дельта
// baseValue = поточне − останнє_записане) → стають чисті, без подвоєння.
// Запускати ВРУЧНУ з Apps Script editor. Спершу Payments, перевірити, потім Salary.
// ═══════════════════════════════════════════════════════════════════════════
var _REEXPORT_EXTRAS_LOCS = ["Кар'єрна", 'Бровари', 'Пуща', 'Позняки', 'Оранж', 'Осокорки', 'Голосієво'];

// 1) Payment: «Бюджет доп» липня (= N+1 від червня) для 7 локацій.
function reexportExtrasJune2026Payments(){
  Logger.log('═══ ПЕРЕРАХУНОК ДОДАТКОВИХ → PAYMENT (червень 2026) — локацій=%s ═══', _REEXPORT_EXTRAS_LOCS.length);
  var summary = [], okCount = 0, errCount = 0;
  _REEXPORT_EXTRAS_LOCS.forEach(function(loc){
    var res = exportAttendanceToPayments({loc: loc, month: 6, year: 2026});
    if (res && res.ok){
      okCount++;
      Logger.log('✅ %s | оновлено=%s | сума=%s грн | notFound=%s',
        loc, res.updated, res.totalAmount, JSON.stringify(res.notFound || []));
      summary.push({loc: loc, ok: true, updated: res.updated, total: res.totalAmount, notFound: res.notFound || []});
    } else {
      errCount++;
      Logger.log('❌ %s | помилка: %s', loc, res && res.error);
      summary.push({loc: loc, ok: false, error: res && res.error});
    }
  });
  Logger.log('─── ПІДСУМОК PAYMENT: успішно=%s | помилок=%s ───', okCount, errCount);
  return {ok: errCount === 0, kind: 'payment', month: 6, year: 2026, okCount: okCount, errCount: errCount, summary: summary};
}

// 2) Salary: ЗП викладачів за червень 2026 (count відміток по заняттю) для 7 локацій.
function reexportExtrasJune2026Salary(){
  Logger.log('═══ ПЕРЕРАХУНОК ДОДАТКОВИХ → SALARY (червень 2026) — локацій=%s ═══', _REEXPORT_EXTRAS_LOCS.length);
  var summary = [], okCount = 0, errCount = 0;
  _REEXPORT_EXTRAS_LOCS.forEach(function(loc){
    var res = exportToSalaryExtras({loc: loc, month: 6, year: 2026});
    if (res && res.ok){
      okCount++;
      Logger.log('✅ %s | оновлено=%s | сума ЗП=%s грн | notFound=%s | пропущено(без ставки)=%s',
        loc, res.updated, res.totalFact, JSON.stringify(res.notFound || []), JSON.stringify(res.skipped || []));
      summary.push({loc: loc, ok: true, updated: res.updated, totalFact: res.totalFact, notFound: res.notFound || [], skipped: res.skipped || []});
    } else {
      errCount++;
      Logger.log('❌ %s | помилка: %s', loc, res && res.error);
      summary.push({loc: loc, ok: false, error: res && res.error});
    }
  });
  Logger.log('─── ПІДСУМОК SALARY: успішно=%s | помилок=%s ───', okCount, errCount);
  return {ok: errCount === 0, kind: 'salary', month: 6, year: 2026, okCount: okCount, errCount: errCount, summary: summary};
}

// ═══════════════════════════════════════════════════════════════════════════
// ОРАНЖ — червень 2026: ЗП допівців із урахуванням обʼєднань груп + ДРУК по
// кожному «За заняття»/«За захід»: назва | груп×дат сирих | після обʼєднань |
// ставка | ЗП. Спершу read-only розклад (діагностика), потім реальний запис у
// Salary через exportToSalaryExtras (журнальна дельта — ідемпотентно). Обгортка
// БЕЗ параметрів — запускати ВРУЧНУ з Apps Script editor.
// «сирі» = сума розрізнених груп×дат до обʼєднань; «після обʼєднань» = сесії, де
// обʼєднані одного дня групи схлопуються в 1 (_dopCountSessions з mergesMap).
// ═══════════════════════════════════════════════════════════════════════════
function reexportSalaryExtrasOrange(){
  var loc = 'Оранж', month = 6, year = 2026;
  Logger.log('═══ ОРАНЖ ЗП ДОПІВЦІВ (%s/%s) з обʼєднаннями — РОЗКЛАД + ЗАПИС ═══', month, year);

  // === 1. Каталог занять локації → лише «За заняття»/«За захід» зі ставкою ===
  var catRes = getActivitiesCatalog(loc);
  if (!catRes.ok){ Logger.log('❌ каталог: %s', catRes.error); return catRes; }
  var perLesson = (catRes.items || []).filter(function(a){
    return a.active && a.teacherRate > 0 &&
      (a.teacherModel === 'За заняття' || a.teacherModel === 'За захід');
  });

  // === 2. Відмітки червня → groupsByDate по кожному activityId ===
  var mm = month < 10 ? '0' + month : String(month);
  var dateFrom = year + '-' + mm + '-01';
  var nextM = _nextMonth(month, year);
  var nmm = nextM.month < 10 ? '0' + nextM.month : String(nextM.month);
  var dateTo = nextM.year + '-' + nmm + '-01';

  var attSh = _getAttendanceSheet(false);
  var attData = attSh.getDataRange().getValues();
  var byActId = {};
  for (var i = 1; i < attData.length; i++){
    var rec = _parseAttendanceRow(attData[i]);
    if (rec.loc !== loc) continue;
    if (rec.date < dateFrom || rec.date >= dateTo) continue;
    if (!byActId[rec.activityId]) byActId[rec.activityId] = {};
    var ng = _dopNormGroup(rec.group);
    if (!byActId[rec.activityId][rec.date]) byActId[rec.activityId][rec.date] = {};
    byActId[rec.activityId][rec.date][ng] = true;
  }

  // === 3. Обʼєднання груп (session-key схлопування) ===
  var mergesMap = _loadDopMergesMap(loc, dateFrom, dateTo);

  // === 4. ДРУК по кожному «за заняття» ===
  Logger.log('─── назва | груп×дат сирих | після обʼєднань | ставка | ЗП ───');
  var rows = [], totalRaw = 0, totalMerged = 0, totalZP = 0;
  perLesson.forEach(function(a){
    var gbd    = byActId[a.id] || {};
    var raw    = _dopCountSessions(gbd, {});                     // до обʼєднань
    var merged = _dopCountSessions(gbd, mergesMap[a.id] || {});  // після обʼєднань
    var zp     = merged * a.teacherRate;
    totalRaw += raw; totalMerged += merged; totalZP += zp;
    Logger.log('%s | %s | %s | %s | %s', a.name, raw, merged, a.teacherRate, zp);
    rows.push({name: a.name, model: a.teacherModel, rawSessions: raw, mergedSessions: merged, rate: a.teacherRate, zp: zp});
  });
  Logger.log('─── РАЗОМ: сирих=%s | після обʼєднань=%s | ЗП=%s грн (занять=%s) ───',
    totalRaw, totalMerged, totalZP, rows.length);

  // === 5. Реальний запис у Salary (журнальна дельта, ідемпотентно) ===
  var exp = exportToSalaryExtras({loc: loc, month: month, year: year});
  if (exp && exp.ok){
    Logger.log('✅ ЗАПИС Salary: оновлено=%s | сума ЗП=%s грн | notFound=%s | пропущено(без ставки)=%s',
      exp.updated, exp.totalFact, JSON.stringify(exp.notFound || []), JSON.stringify(exp.skipped || []));
  } else {
    Logger.log('❌ ЗАПИС Salary: %s', exp && exp.error);
  }

  return {
    ok: !!(exp && exp.ok),
    loc: loc, month: month, year: year,
    rows: rows,
    totalRawSessions: totalRaw,
    totalMergedSessions: totalMerged,
    totalZP: totalZP,
    write: exp
  };
}

// ═══════════════════════════════════════════════════════════════════════════
// МАСОВИЙ ПЕРЕРАХУНОК SALARY (Додаткові) для ВСІХ локацій Salary-реєстру.
// Використати після впровадження обʼєднань груп (v7.08) або будь-якої зміни
// логіки ЗП «За заняття»/«За захід», щоб перезаписати суми в усіх файлах.
// month/year — місяць ВІДПРАЦЮВАННЯ (ЗП летить у Budget місяця N+1). Дефолт —
// поточний місяць. Ідемпотентно (журнальна дельта). Запускати ВРУЧНУ з редактора.
// ═══════════════════════════════════════════════════════════════════════════
// БЕЗ параметрів → ЧЕРВЕНЬ 2026 (місяць відпрацювання; ЗП летить у Budget липня).
// Перераховує ЗП допівців по ВСІХ локаціях Salary-реєстру з НОВОЮ логікою
// session-key (група×дата, кожна група окремо; обʼєднання ще не проставлені —
// _dopCountSessions застосує mergesMap коли директор їх позначить).
// Друк по кожній локації: локація | скільки допівців | сума ЗП.
// «Допівців» = к-сть занять «За заняття»/«За дитину»/«За захід» з нарахованою
// сумою > 0 (реально відпрацьовані цього місяця). Ідемпотентно (журнальна
// дельта в exportToSalaryExtras) — можна перезапускати після проставлення
// обʼєднань, суми лише скоригуються, не подвояться. Запускати ВРУЧНУ з редактора.
// Параметри опційні (для повторного використання іншими місяцями).
function reexportSalaryExtrasAllLocations(month, year){
  month = Number(month) || 6;      // дефолт: червень 2026
  year  = Number(year)  || 2026;
  var reg = _salaryGetRegistry();
  if (!reg.ok){ Logger.log('❌ %s', reg.error); return reg; }
  var locs = [];
  reg.rows.forEach(function(r){ if (locs.indexOf(r.loc) < 0) locs.push(r.loc); });
  Logger.log('═══ МАСОВИЙ ПЕРЕРАХУНОК SALARY ДОДАТКОВИХ (%s/%s) — локацій=%s ═══', month, year, locs.length);
  Logger.log('─── локація | допівців | сума ЗП ───');
  var summary = [], okCount = 0, errCount = 0, grandTotal = 0, grandDop = 0;
  locs.forEach(function(loc){
    var res = exportToSalaryExtras({loc: loc, month: month, year: year});
    if (res && res.ok){
      okCount++;
      // Скільки допівців реально нараховано (fact>0) з details експорту.
      var dopCount = (res.details || []).filter(function(d){
        return d && d.status === 'updated' && Number(d.fact) > 0;
      }).length;
      grandTotal += Number(res.totalFact) || 0;
      grandDop   += dopCount;
      Logger.log('✅ %s | допівців=%s | ЗП=%s грн | notFound=%s | пропущено=%s',
        loc, dopCount, res.totalFact, JSON.stringify(res.notFound || []), JSON.stringify(res.skipped || []));
      summary.push({loc: loc, ok: true, dopCount: dopCount, updated: res.updated, totalFact: res.totalFact, notFound: res.notFound || [], skipped: res.skipped || []});
    } else {
      errCount++;
      Logger.log('❌ %s | помилка: %s', loc, res && res.error);
      summary.push({loc: loc, ok: false, error: res && res.error});
    }
  });
  Logger.log('─── РАЗОМ: допівців=%s | ЗП=%s грн | локацій успішно=%s | помилок=%s ───',
    grandDop, grandTotal, okCount, errCount);
  return {ok: errCount === 0, kind: 'salary', month: month, year: year,
          okCount: okCount, errCount: errCount, totalDop: grandDop, totalZP: grandTotal, summary: summary};
}

// ═══════════════════════════════════════════════════════════════════════════
// ПОВНИЙ ПЕРЕРАХУНОК ЗП МЕРЕЖІ ЗА ЧЕРВЕНЬ 2026 — і ДОПІВЦІ, і ПРЕДМЕТНИКИ, по
// ВСІХ локаціях Salary-реєстру, з урахуванням проставлених ОБʼЄДНАНЬ (merges
// схлопують група×дата в обох потоках — exportToSalaryExtras/_dopCountSessions
// для допівців та exportPredmetnykyToSalary/_loadPredMergesMap для предметників).
// Друк по локації: локація | допівці ЗП | предметники ЗП | разом + фінал РАЗОМ
// по мережі. Ідемпотентно (журнальна дельта в обох експортах) — можна
// перезапускати після того як директори проставлять ще обʼєднання.
// БЕЗ параметрів. Запускати ВРУЧНУ з Apps Script editor.
// ═══════════════════════════════════════════════════════════════════════════
function reexportAllSalaryJune(){
  var month = 6, year = 2026;
  var reg = _salaryGetRegistry();
  if (!reg.ok){ Logger.log('❌ %s', reg.error); return reg; }
  var locs = [];
  reg.rows.forEach(function(r){ if (locs.indexOf(r.loc) < 0) locs.push(r.loc); });
  Logger.log('═══ ПОВНИЙ ПЕРЕРАХУНОК ЗП МЕРЕЖІ (%s/%s) — локацій=%s ═══', month, year, locs.length);
  Logger.log('─── локація | допівці ЗП | предметники ЗП | разом ───');

  var summary = [], gDop = 0, gPred = 0, errs = [];
  locs.forEach(function(loc){
    // 1) Допівці «За заняття»/«За захід» (session-key група×дата + обʼєднання).
    var dop  = exportToSalaryExtras({loc: loc, month: month, year: year});
    // 2) Предметники (унікальні група×дата × ставку + обʼєднання схлопують).
    var pred = exportPredmetnykyToSalary({loc: loc, month: month, year: year});

    var dopZP  = (dop  && dop.ok)  ? (Number(dop.totalFact)  || 0) : 0;
    var predZP = (pred && pred.ok) ? (Number(pred.totalFact) || 0) : 0;
    var sum = dopZP + predZP;
    gDop += dopZP; gPred += predZP;
    if (!(dop  && dop.ok))  errs.push(loc + ' · допівці: '     + (dop  && dop.error));
    if (!(pred && pred.ok)) errs.push(loc + ' · предметники: ' + (pred && pred.error));

    Logger.log('%s | %s | %s | %s', loc, dopZP, predZP, sum);
    summary.push({
      loc: loc, dopZP: dopZP, predZP: predZP, total: sum,
      dopOk: !!(dop && dop.ok), predOk: !!(pred && pred.ok),
      dopError: dop && dop.error, predError: pred && pred.error
    });
  });

  Logger.log('─── РАЗОМ ПО МЕРЕЖІ: допівці=%s | предметники=%s | ВСЬОГО=%s грн ───',
    gDop, gPred, gDop + gPred);
  if (errs.length) Logger.log('⚠️ Помилки (%s): %s', errs.length, JSON.stringify(errs));

  return {ok: errs.length === 0, kind: 'salary_full', month: month, year: year,
          locations: locs.length, totalDop: gDop, totalPred: gPred,
          grandTotal: gDop + gPred, summary: summary, errors: errs};
}

// ═══════════════════════════════════════════════════════════════════════════
// ПОВНИЙ ПЕРЕРАХУНОК ОДНІЄЇ ЛОКАЦІЇ після виправлення відміток (червень 2026).
// 3 кроки: Payment (exportAttendanceToPayments) → Salary (exportToSalaryExtras) →
// Агрегат Оплати-Рік (aggregatePaymentsYearly). Один запуск = і рахунки, і ЗП, і
// таблиця рахунків стають коректні. Усі кроки ідемпотентні (журнальна дельта +
// повний rebuild агрегату). Запускати ВРУЧНУ з редактора (читає файли локацій — довго).
// ⚠️ aggregatePaymentsYearly() — ГЛОБАЛЬНА (перебудовує весь Оплати-Рік по всіх локаціях),
//    не per-loc; тут вона завершальним кроком підтягує виправлену локацію в агрегат.
// ═══════════════════════════════════════════════════════════════════════════
// month/year — опційні (дефолт червень 2026, щоб старі виклики reexportKruglaFull /
// reexportAllDedupedLocations не змінились). Кнопка на рахунках передає вибраний місяць.
function reexportLocationFull(loc, month, year){
  loc = String(loc || '').trim();
  if (!loc){ Logger.log('❌ loc обовʼязковий'); return {ok:false, error:'loc обовʼязковий'}; }
  month = Number(month) || 6;
  year  = Number(year)  || 2026;
  Logger.log('═══ ПОВНИЙ ПЕРЕРАХУНОК ЛОКАЦІЇ "%s" (%s/%s) ═══', loc, month, year);

  Logger.log('▶ Крок 1/3: Payment (exportAttendanceToPayments)…');
  var pay = exportAttendanceToPayments({loc: loc, month: month, year: year});
  if (pay && pay.ok) Logger.log('  ✅ Payment: оновлено=%s | сума=%s грн | notFound=%s', pay.updated, pay.totalAmount, JSON.stringify(pay.notFound || []));
  else               Logger.log('  ❌ Payment: %s', pay && pay.error);

  Logger.log('▶ Крок 2/3: Salary (exportToSalaryExtras)…');
  var sal = exportToSalaryExtras({loc: loc, month: month, year: year});
  if (sal && sal.ok) Logger.log('  ✅ Salary: оновлено=%s | сума ЗП=%s грн | notFound=%s', sal.updated, sal.totalFact, JSON.stringify(sal.notFound || []));
  else               Logger.log('  ❌ Salary: %s', sal && sal.error);

  Logger.log('▶ Крок 3/3: Агрегація Оплати-Рік (aggregatePaymentsYearly, глобальна)…');
  var agg = aggregatePaymentsYearly();
  if (agg && agg.ok) Logger.log('  ✅ Агрегат: рядків=%s | помилок=%s', agg.rows, (agg.errors || []).length);
  else               Logger.log('  ❌ Агрегат: %s', agg && agg.error);

  var allOk = !!(pay && pay.ok && sal && sal.ok && agg && agg.ok);
  Logger.log('─── ПОВНИЙ ПЕРЕРАХУНОК "%s" завершено (ok=%s) ───', loc, allOk);
  return {ok: allOk, loc: loc, month: month, year: year, payment: pay, salary: sal, aggregate: agg};
}

// Обгортка для зручного запуску з редактора (без аргументів).
function reexportKruglaFull(){ return reexportLocationFull('Кругла'); }

// ТИМЧАСОВА обгортка (one-shot, після звірки 2026-06-30): перерахунок 5 проблемних
// локацій. ⚠️ reexportLocationFull НЕ пише знижки (Бюджет навч) — лише extras+salary+
// агрегат. Тому тут ОКРЕМО кличемо exportVacationDiscountToPayments (знижки, ТАРГЕТ-
// місяць = 7), а extras/salary — за місяць ВІДМІТОК = 6. Один фінальний агрегат у кінці.
// Видалити після використання.
function reexportProblemLocationsJuly(){
  var LOCS = ["Кругла", "Бровари", "Пуща", "Кар'єрна", "Бігова"];
  Logger.log('═══ ПЕРЕРАХУНОК ПРОБЛЕМНИХ ЛОКАЦІЙ (відмітки червня → бюджет липня) ═══');
  var summary = [];
  LOCS.forEach(function(loc){
    Logger.log('\n━━━ %s ━━━', loc);
    var pay = exportAttendanceToPayments({loc: loc, month: 6, year: 2026});       // додаткові: черв → лип «Бюджет доп»
    Logger.log(pay && pay.ok ? '  ✅ Доп: оновлено=' + pay.updated + ' сума=' + pay.totalAmount + ' notFound=' + JSON.stringify(pay.notFound || []) : '  ❌ Доп: ' + (pay && pay.error));
    var sal = exportToSalaryExtras({loc: loc, month: 6, year: 2026});             // ЗП викладачів
    Logger.log(sal && sal.ok ? '  ✅ ЗП: оновлено=' + sal.updated + ' сума=' + sal.totalFact : '  ❌ ЗП: ' + (sal && sal.error));
    var dis = exportVacationDiscountToPayments({loc: loc, month: 7, year: 2026}); // ЗНИЖКИ (відпустка+хвороба): ТАРГЕТ липень
    Logger.log(dis && dis.ok ? '  ✅ Знижки: оновлено=' + dis.updated + ' сума=-' + dis.totalDiscount + ' notFound=' + JSON.stringify(dis.notFound || []) : '  ❌ Знижки: ' + (dis && dis.error));
    summary.push({loc: loc, payTotal: pay && pay.totalAmount, salTotal: sal && sal.totalFact,
                  disTotal: dis && dis.totalDiscount, disNotFound: (dis && dis.notFound) || []});
  });
  Logger.log('\n═══ ФІНАЛ: aggregatePaymentsYearly() (глобальна, один раз) ═══');
  var agg = aggregatePaymentsYearly();
  Logger.log(agg && agg.ok ? '  ✅ Агрегат: рядків=' + agg.rows + ' помилок=' + (agg.errors || []).length : '  ❌ Агрегат: ' + (agg && agg.error));
  Logger.log('\n─── ГОТОВО (локацій=%s) ───', LOCS.length);
  return {ok: !!(agg && agg.ok), locations: LOCS, summary: summary, aggregate: agg};
}

// READ-ONLY скан: знаходить формульні клітинки у колонці «Бюджет доп» липня (col 35)
// по ВСІХ локаціях-реєстру — щоб знати, кого зачіпав баг пропуску формул. Нічого не пише.
// Запускати ВРУЧНУ з редактора. month=7 (таргет липня).
function scanDopFormulasJuly(){
  var month = 7;
  var col1 = 1 + (month - 1) * 5 + 3 + 1;   // = 35 (budgetDopCol1)
  var configSS = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  var cfg = configSS.getSheets()[0].getDataRange().getValues();
  var found = [], locsWith = {};
  Logger.log('═══ READ-ONLY: формули в «Бюджет доп» липня (col %s) по локаціях ═══', col1);
  for (var r = 1; r < cfg.length; r++){
    var loc = trim(String(cfg[r][2] || '')), sheetId = trim(String(cfg[r][3] || ''));
    var sheetName = trim(String(cfg[r][4] || '')) || 'Payment';
    if (!loc || !sheetId) continue;
    try {
      var sh = SpreadsheetApp.openById(sheetId).getSheetByName(sheetName) || SpreadsheetApp.openById(sheetId).getSheets()[0];
      var last = sh.getLastRow();
      if (last < 3) continue;
      var fs = sh.getRange(1, col1, last, 1).getFormulas();
      var vs = sh.getRange(1, col1, last, 1).getValues();
      var names = sh.getRange(1, 1, last, 1).getValues();
      for (var i = 2; i < last; i++){
        if (fs[i][0]){
          var nm = trim(String(names[i][0] || ''));
          if (!nm || isGroupHeaderRow([names[i][0]], 1)) continue;   // підсумкові/group — ігноруємо
          found.push({loc: loc, row: i + 1, name: nm, formula: fs[i][0], value: vs[i][0]});
          locsWith[loc] = (locsWith[loc] || 0) + 1;
          Logger.log('  %s | row %s | "%s" | %s = %s', loc, i + 1, nm, fs[i][0], vs[i][0]);
        }
      }
    } catch(e){ Logger.log('  ⚠️ %s: %s', loc, e && e.message); }
  }
  Logger.log('─── Формул-рядків: %s | локацій з формулами: %s ───', found.length, JSON.stringify(locsWith));
  return {ok: true, col: col1, totalFormulaRows: found.length, byLoc: locsWith, rows: found};
}

// Перерахунок УСІХ почищених локацій за один запуск: по кожній — Payment + Salary
// (per-loc), а агрегацію Оплати-Рік робимо ОДИН РАЗ у кінці (глобальна, поза циклом,
// щоб не ганяти 8 разів). Усі кроки ідемпотентні. Запускати ВРУЧНУ з редактора
// (читає файли локацій — довго, не через веб-апку).
var _REEXPORT_DEDUPED_LOCS = ["Кар'єрна", 'Пуща', 'Бровари', 'Осокорки', 'Позняки', 'Голосієво', 'Борщагівка', 'Кругла'];

function reexportAllDedupedLocations(){
  Logger.log('═══ ПЕРЕРАХУНОК ПОЧИЩЕНИХ ЛОКАЦІЙ (червень 2026) — локацій=%s ═══', _REEXPORT_DEDUPED_LOCS.length);
  var perLoc = [], payOk = 0, salOk = 0, errs = [];

  _REEXPORT_DEDUPED_LOCS.forEach(function(loc){
    Logger.log('\n━━━ %s ━━━', loc);

    Logger.log('  ▶ Payment (exportAttendanceToPayments)…');
    var pay = exportAttendanceToPayments({loc: loc, month: 6, year: 2026});
    if (pay && pay.ok){ payOk++; Logger.log('    ✅ Payment: оновлено=%s | сума=%s грн | notFound=%s', pay.updated, pay.totalAmount, JSON.stringify(pay.notFound || [])); }
    else { errs.push(loc + ' Payment: ' + (pay && pay.error)); Logger.log('    ❌ Payment: %s', pay && pay.error); }

    Logger.log('  ▶ Salary (exportToSalaryExtras)…');
    var sal = exportToSalaryExtras({loc: loc, month: 6, year: 2026});
    if (sal && sal.ok){ salOk++; Logger.log('    ✅ Salary: оновлено=%s | сума ЗП=%s грн | notFound=%s', sal.updated, sal.totalFact, JSON.stringify(sal.notFound || [])); }
    else { errs.push(loc + ' Salary: ' + (sal && sal.error)); Logger.log('    ❌ Salary: %s', sal && sal.error); }

    perLoc.push({loc: loc, payment: pay, salary: sal});
  });

  // === ФІНАЛЬНА АГРЕГАЦІЯ — один раз, поза циклом ===
  Logger.log('\n═══ ФІНАЛ: aggregatePaymentsYearly() (глобальна, один раз) ═══');
  var agg = aggregatePaymentsYearly();
  if (agg && agg.ok) Logger.log('  ✅ Агрегат Оплати-Рік: рядків=%s | помилок=%s', agg.rows, (agg.errors || []).length);
  else { errs.push('Aggregate: ' + (agg && agg.error)); Logger.log('  ❌ Агрегат: %s', agg && agg.error); }

  Logger.log('\n─── ПІДСУМОК: Payment ok=%s/%s | Salary ok=%s/%s | агрегат=%s | помилок=%s ───',
    payOk, _REEXPORT_DEDUPED_LOCS.length, salOk, _REEXPORT_DEDUPED_LOCS.length, (agg && agg.ok) ? 'ok' : 'FAIL', errs.length);
  if (errs.length) errs.forEach(function(e){ Logger.log('   • %s', e); });

  return {ok: errs.length === 0, locations: _REEXPORT_DEDUPED_LOCS, payOk: payOk, salOk: salOk,
          aggregate: agg, perLoc: perLoc, errors: errs};
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

// ───────────────────────────────────────────────────────────────────────────
// v6.25 ДІАГНОСТИКА: для кожної локації у Salary-реєстрі прогоняє вміст
// Salary-файла через _classifyAllSalaryRows і рахує:
//   • чи є рядок-заголовок "Додаткові заняття" (section_header у секції extras)
//   • скільки рядків з _section === 'extras' (без заголовків/group-headers)
//   • перелік назв extras-рядків (rowNumber + name)
// Запускати ВРУЧНУ з Apps Script editor. Логи у View → Executions.
//
// Мета — перед фіксом exportToSalaryExtras (фільтр по _section === 'extras')
// переконатись, що у кожній з ~19 локацій є секція extras з ≥1 рядком.
// Інакше extras-фільтр поверне 0 і всі заняття полетять у notFound.
// ───────────────────────────────────────────────────────────────────────────
function diagSalaryExtrasSections(){
  var reg = _salaryGetRegistry();
  if (!reg.ok){
    Logger.log('[diagSalaryExtras] ❌ %s', reg.error);
    return reg;
  }
  Logger.log('[diagSalaryExtras] START — %s локацій у реєстрі', reg.rows.length);

  var results = [];
  var locsOk = 0, locsNoHeader = 0, locsEmpty = 0, locsErr = 0;

  reg.rows.forEach(function(entry){
    var loc = entry.loc;
    var rep = {loc: loc, typ: entry.typ, hasHeader: false, extrasCount: 0, items: [], error: null};
    try {
      var locSS = SpreadsheetApp.openById(entry.sheetId);
      var sheet = locSS.getSheetByName(entry.listName);
      if (!sheet){
        rep.error = 'sheet "' + entry.listName + '" не знайдено';
        results.push(rep); locsErr++;
        return;
      }
      var lastRow = Math.max(sheet.getLastRow(), 80);
      var lastCol = Math.max(sheet.getLastColumn(), 37);
      var data = sheet.getRange(1, 1, lastRow, lastCol).getValues();

      // Збираємо raw-rows як у getSalaryData (з рядка 4) — БЕЗ skip subtotal,
      // бо саме subtotal-рядок "Додаткові заняття" перемикає state у extras.
      var rawRows = [];
      for (var rowNum = 4; rowNum <= data.length; rowNum++){
        var idx = rowNum - 1;
        var rowArr = data[idx] || [];
        var rawName = String(rowArr[0] || '').trim();
        if (_salaryIsSkippedRow(rawName)) continue;
        var totalFact = 0, totalBudget = 0;
        for (var m = 1; m <= 12; m++){
          var fIdx = (m - 1) * 3 + 1;
          var bIdx = (m - 1) * 3 + 2;
          if (fIdx < lastCol) totalFact   += _opexNum(rowArr[fIdx]);
          if (bIdx < lastCol) totalBudget += _opexNum(rowArr[bIdx]);
        }
        rawRows.push({row: rowNum, name: rawName, fact: totalFact, budget: totalBudget});
      }

      var classified = _classifyAllSalaryRows(rawRows);
      classified.forEach(function(r){
        if (r._section !== 'extras') return;
        if (r._category === 'section_header'){ rep.hasHeader = true; return; }
        if (r._category === 'group_header') return;
        rep.extrasCount++;
        if (rep.items.length < 30){
          rep.items.push('A' + r.row + ': "' + r.name + '"');
        }
      });

      if (!rep.hasHeader && rep.extrasCount === 0){ locsEmpty++; }
      else if (!rep.hasHeader){ locsNoHeader++; }
      else { locsOk++; }
    } catch(e){
      rep.error = String(e && e.message || e);
      locsErr++;
    }
    results.push(rep);
  });

  // Звіт як таблиця у логах.
  Logger.log('[diagSalaryExtras] ═══════════════════════════════════════');
  Logger.log('[diagSalaryExtras] LOC                  | TYP    | HDR | EXTRAS | NOTE');
  Logger.log('[diagSalaryExtras] ─────────────────────┼────────┼─────┼────────┼─────────────────────');
  results.forEach(function(r){
    var locP = (r.loc + '                    ').slice(0, 20);
    var typP = ((r.typ || '') + '       ').slice(0, 6);
    var hdr  = r.hasHeader ? ' ✓ ' : ' ✗ ';
    var ext  = ('     ' + r.extrasCount).slice(-6);
    var note = r.error ? ('❌ ' + r.error) :
               (r.hasHeader && r.extrasCount > 0 ? 'OK' :
                (r.hasHeader ? '⚠ header є, але 0 extras-рядків' :
                 (r.extrasCount > 0 ? '⚠ extras-рядки є БЕЗ header (state-machine не активувала extras)' :
                  '⚠ ні header, ні extras-рядків — секція відсутня')));
    Logger.log('[diagSalaryExtras] %s | %s | %s | %s | %s', locP, typP, hdr, ext, note);
  });
  Logger.log('[diagSalaryExtras] ═══════════════════════════════════════');
  Logger.log('[diagSalaryExtras] SUMMARY: OK=%s, NO_HEADER=%s, EMPTY=%s, ERR=%s (total=%s)',
    locsOk, locsNoHeader, locsEmpty, locsErr, reg.rows.length);

  // Деталі по кожній локації — назви extras-рядків (для перевірки чи "Логопед"
  // справді є у списку extras).
  Logger.log('[diagSalaryExtras] ─── DETAIL: extras-рядки по локаціях ───');
  results.forEach(function(r){
    if (r.error) return;
    if (!r.items.length){
      Logger.log('[diagSalaryExtras] %s: (немає extras-рядків)', r.loc);
      return;
    }
    Logger.log('[diagSalaryExtras] %s (%s рядків): %s',
      r.loc, r.extrasCount, r.items.join(', '));
  });

  return {ok: true, results: results,
    summary: {ok: locsOk, noHeader: locsNoHeader, empty: locsEmpty, err: locsErr,
              total: reg.rows.length}};
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
// v6.26 Етап 1B — Дані для вкладки "📧 Рахунки" (UI-список дітей з сумами).
// Повертає для локації + (місяць навчання, місяць додаткових):
//   • paymentSum — з "Оплати-Рік" → "Бюджет-навч" обраного місяця
//   • extrasSum + breakdown — з "Додаткові_Відвідуваність" × ціни занять
// Фільтр: status IN ('active','adaptation') AND (paymentSum > 0 OR extrasSum > 0).
// Підписант (signerName/Phone/Email) — з картки клієнта (поле "Підписант договору").
// Той самий endpoint буде reused у 1C при генерації PDF.
// ═══════════════════════════════════════════════════════════════════════════






function diagDopPicture(loc){
  loc = loc || 'Осокорки';
  var r = getInvoiceListData({loc:loc, payMonth:6, payYear:2026, extMonth:6, extYear:2026});
  var actBy = {};
  if (r && r.ok){ (r.children||[]).forEach(function(c){
    var lines = (c.extrasBreakdown||[]).map(function(b){ return b.name+'×'+b.count+'='+b.total; }).join(', ');
    actBy[c.name] = {sum: Number(c.extrasBreakdownSum)||0, lines: lines};
  }); }
  var sh = getCRMSpreadsheet().getSheetByName(SHEET_YEARLY);
  var v = sh.getDataRange().getValues();
  var h = v[0].map(function(x){ return String(x); });
  var nameI=h.indexOf("Ім'я дитини"), locI=h.indexOf('Локація');
  var cb=h.indexOf('Червень-Бюджет-доп'), cf=h.indexOf('Червень-Факт-доп');
  var lb=h.indexOf('Липень-Бюджет-доп'), lf=h.indexOf('Липень-Факт-доп');
  Logger.log('=== %s | заняття червня + ДОП Червень→Липень ===', loc);
  var n=0;
  for (var i=1;i<v.length;i++){
    if (String(v[i][locI]).trim()!==loc) continue;
    var nm=String(v[i][nameI]).trim(); if(!nm) continue;
    var cbV=Number(v[i][cb])||0, cfV=Number(v[i][cf])||0, lbV=Number(v[i][lb])||0, lfV=Number(v[i][lf])||0;
    var act=actBy[nm]||{sum:0,lines:''};
    if (cbV===0&&cfV===0&&lbV===0&&lfV===0&&act.sum===0) continue;
    n++;
    var debtCh=cbV-cfV; var flag=debtCh>0.5?('  ← борг '+debtCh):'';
    Logger.log('%s | заняття чер=%s [%s] | Чер: Б=%s Ф=%s | Лип: Б=%s Ф=%s%s', nm, act.sum, act.lines, cbV, cfV, lbV, lfV, flag);
  }
  Logger.log('=== всього: %s ===', n);
}

function diagNavchPicture(loc){
  loc = loc || 'Осокорки';
  var sh = getCRMSpreadsheet().getSheetByName(SHEET_YEARLY);
  var v = sh.getDataRange().getValues();
  var h = v[0].map(function(x){ return String(x); });
  var nameI=h.indexOf("Ім'я дитини"), locI=h.indexOf('Локація');
  var cb=h.indexOf('Червень-Бюджет-навч'), cf=h.indexOf('Червень-Факт-навч');
  var lb=h.indexOf('Липень-Бюджет-навч'), lf=h.indexOf('Липень-Факт-навч');
  if (cb<0){ Logger.log('⚠ колонки навч не знайдено (Червень-Бюджет-навч=%s)', cb); return; }
  Logger.log('=== %s | НАВЧАННЯ Червень→Липень ===', loc);
  var n=0, debtors=0;
  for (var i=1;i<v.length;i++){
    if (String(v[i][locI]).trim()!==loc) continue;
    var nm=String(v[i][nameI]).trim(); if(!nm) continue;
    var cbV=Number(v[i][cb])||0, cfV=Number(v[i][cf])||0, lbV=Number(v[i][lb])||0, lfV=Number(v[i][lf])||0;
    if (cbV===0&&cfV===0&&lbV===0&&lfV===0) continue;
    n++; var debt=cbV-cfV; var flag=debt>0.5?('  ← борг '+debt):(debt<-0.5?('  ← перепл '+(-debt)):'');
    if(debt>0.5) debtors++;
    Logger.log('%s | Чер: Б=%s Ф=%s | Лип: Б=%s Ф=%s%s', nm, cbV, cfV, lbV, lfV, flag);
  }
  Logger.log('=== всього: %s | боржників (Чер Б>Ф): %s ===', n, debtors);
}



function diagNewInvoice(loc){
  loc = loc || 'Осокорки';
  var sh = getCRMSpreadsheet().getSheetByName(SHEET_YEARLY);
  var v = sh.getDataRange().getValues();
  var h = v[0].map(function(x){ return String(x); });
  var nameI=h.indexOf("Ім'я дитини"), locI=h.indexOf('Локація');
  function col(c){ return h.indexOf(c); }
  var cBn=col('Червень-Бюджет-навч'), cFn=col('Червень-Факт-навч');
  var cBd=col('Червень-Бюджет-доп'), cFd=col('Червень-Факт-доп');
  var lBn=col('Липень-Бюджет-навч'), lBd=col('Липень-Бюджет-доп');
  Logger.log('=== ПРОЕКТ РАХУНКІВ | локація=%s ===', loc);
  function num(r,i){ return i>=0 ? (Number(v[r][i])||0) : 0; }
  function ds(x){ return x>0.5?('+борг '+x):(x<-0.5?('перепл '+(-x)):'0'); }
  var n=0;
  for (var i=1;i<v.length;i++){
    if (String(v[i][locI]).trim()!==loc) continue;
    var nm=String(v[i][nameI]).trim(); if(!nm) continue;
    var navchCur=num(i,lBn), navchDebt=num(i,cBn)-num(i,cFn);
    var dopCur=num(i,lBd), dopDebt=num(i,cBd)-num(i,cFd);
    var total=navchCur+navchDebt+dopCur+dopDebt;
    if (navchCur===0 && dopCur===0 && navchDebt===0 && dopDebt===0) continue;
    n++;
    Logger.log('[%s] %s (рядок %s) | навч: %s (борг %s) | доп: %s (борг %s) | РАЗОМ=%s', loc, nm, (i+1), navchCur, ds(navchDebt), dopCur, ds(dopDebt), total);
  }
  Logger.log('=== всього: %s ===', n);
}

function diagDupAbsences(){
  var ss=getCRMSpreadsheet();
  var sheet=ss.getSheetByName(SHEET_CLIENTS);
  var vals=sheet.getDataRange().getValues();
  var hdrs=vals[0].map(String);
  var colAbs=hdrs.indexOf('Відсутності (JSON)');
  var colName=-1;
  for(var k=0;k<hdrs.length;k++){ if(hdrs[k].indexOf("м'я")>=0||hdrs[k].indexOf('ПІБ')>=0){ colName=k; break; } }
  if(colName<0) colName=1;
  var colLoc=hdrs.indexOf('Локація'); if(colLoc<0) colLoc=2;
  if(colAbs<0){ Logger.log('⚠ нема колонки Відсутності (JSON)'); return; }
  var totalDup=0, kidsDup=0;
  for(var r=1;r<vals.length;r++){
    var raw=vals[r][colAbs]; if(!raw) continue;
    var arr; try{ arr=JSON.parse(raw); }catch(e){ continue; }
    if(!arr||!arr.length) continue;
    var seen={}, dups=[];
    arr.forEach(function(a){
      if(!a||a.status==='cancelled'||a.status==='rejected') return;
      var key=(a.type||'')+'|'+(a.from||'')+'|'+(a.to||'');
      if(seen[key]){ dups.push(a); } else seen[key]=a;
    });
    if(dups.length){
      kidsDup++; totalDup+=dups.length;
      Logger.log('[%s] %s — дублів: %s | %s', vals[r][colLoc], vals[r][colName], dups.length, dups.map(function(d){return (d.type||'')+' '+(d.from||'')+'→'+(d.to||'')+' '+(d.totalPct||'')+'%';}).join(' ; '));
    }
  }
  Logger.log('=== дітей з дублями: %s | всього зайвих записів: %s ===', kidsDup, totalDup);
}

function diagChildAbsences(namePart){
  namePart = (namePart||'Заяні').toLowerCase();
  var ss=getCRMSpreadsheet();
  var sheet=ss.getSheetByName(SHEET_CLIENTS);
  var vals=sheet.getDataRange().getValues();
  var hdrs=vals[0].map(String);
  var colAbs=hdrs.indexOf('Відсутності (JSON)');
  var colName=-1; for(var k=0;k<hdrs.length;k++){ if(hdrs[k].indexOf("м'я")>=0||hdrs[k].indexOf('ПІБ')>=0){ colName=k; break; } } if(colName<0) colName=1;
  var colLoc=hdrs.indexOf('Локація'); if(colLoc<0) colLoc=2;
  var colGrp=hdrs.indexOf('Група'); if(colGrp<0) colGrp=3;
  var n=0;
  for(var r=1;r<vals.length;r++){
    var nm=String(vals[r][colName]||'');
    if(nm.toLowerCase().indexOf(namePart)<0) continue;
    n++;
    var id=String(vals[r][0]);
    var raw=vals[r][colAbs]; var arr=[]; try{ arr=JSON.parse(raw)||[]; }catch(e){}
    Logger.log('—— Рядок %s | id=%s | лок=%s | група=%s | відсутностей: %s', r+1, id, vals[r][colLoc], vals[r][colGrp], arr.length);
    arr.forEach(function(a,i){ if(!a) return; Logger.log('     [%s] %s %s→%s | %s | %s%% | createdBy=%s createdAt=%s id=%s', i, a.type, a.from, a.to, a.status, a.totalPct, a.createdBy, (a.createdAt||'').slice(0,16), a.id); });
  }
  Logger.log('=== рядків з "%s": %s ===', namePart, n);
}

function diagOranzhDates(loc){
  loc = loc || 'Оранж';
  function dkey(v){
    if(v instanceof Date){ return Utilities.formatDate(v, 'Europe/Kiev', 'yyyy-MM-dd'); }
    var ss=String(v||'').trim();
    if(/^\d{4}-\d{2}-\d{2}/.test(ss)) return ss.slice(0,10);
    var m=ss.match(/^(\d{1,2})[.\/](\d{1,2})[.\/](\d{4})/);
    if(m) return m[3]+'-'+('0'+m[2]).slice(-2)+'-'+('0'+m[1]).slice(-2);
    return ss;
  }
  var from='2026-06-01', to='2026-06-30';
  var tab={}, dop={};
  var crm=getCRMSpreadsheet().getSheetByName(SHEET_ATTENDANCE);
  if(crm){
    var tv=crm.getDataRange().getValues();
    for(var i=1;i<tv.length;i++){
      if(String(tv[i][3]||'').trim()!==loc) continue;
      var d=dkey(tv[i][0]); if(d<from||d>to) continue;
      tab[d]=(tab[d]||0)+1;
    }
  }
  var dsh=_getAttendanceSheet(false);
  if(dsh){
    var dv=dsh.getDataRange().getValues();
    for(var j=1;j<dv.length;j++){
      var rec=_parseAttendanceRow(dv[j]);
      if(String(rec.loc||'').trim()!==loc) continue;
      var dd=dkey(rec.date); if(dd<from||dd>to) continue;
      dop[dd]=(dop[dd]||0)+1;
    }
  }
  Logger.log('=== %s | табель + допи по днях ЧЕРВЕНЬ ===', loc);
  var wd=['Нд','Пн','Вт','Ср','Чт','Пт','Сб'];
  var tabTot=0, dopTot=0;
  for(var dnum=1; dnum<=30; dnum++){
    var key='2026-06-'+('0'+dnum).slice(-2);
    var dt=new Date(2026,5,dnum); var w=wd[dt.getDay()];
    var t=tab[key]||0, dd2=dop[key]||0;
    tabTot+=t; dopTot+=dd2;
    var isWeekend=(dt.getDay()===0||dt.getDay()===6);
    var flag=(!isWeekend && t===0 && dd2===0)?'  ← ПОРОЖНІЙ будній':'';
    Logger.log('%s %s | табель: %s | допи: %s%s', w, key, t, dd2, flag);
  }
  Logger.log('=== РАЗОМ червень | табель: %s | допи: %s ===', tabTot, dopTot);
}

function diagDopByWeekday(){
  function dkey(v){
    if(v instanceof Date){ return Utilities.formatDate(v, 'Europe/Kiev', 'yyyy-MM-dd'); }
    var ss=String(v||'').trim();
    if(/^\d{4}-\d{2}-\d{2}/.test(ss)) return ss.slice(0,10);
    var m=ss.match(/^(\d{1,2})[.\/](\d{1,2})[.\/](\d{4})/);
    if(m) return m[3]+'-'+('0'+m[2]).slice(-2)+'-'+('0'+m[1]).slice(-2);
    return ss;
  }
  var wd=['Нд','Пн','Вт','Ср','Чт','Пт','Сб'];
  var dsh=_getAttendanceSheet(false);
  var dv=dsh.getDataRange().getValues();
  var byLoc={};
  for(var j=1;j<dv.length;j++){
    var rec=_parseAttendanceRow(dv[j]);
    var dd=dkey(rec.date); if(dd<'2026-06-01'||dd>'2026-06-30') continue;
    var loc=String(rec.loc||'?').trim();
    var parts=dd.split('-'); var dt=new Date(+parts[0], +parts[1]-1, +parts[2]); var w=dt.getDay();
    if(!byLoc[loc]) byLoc[loc]=[0,0,0,0,0,0,0];
    byLoc[loc][w]++;
  }
  Logger.log('=== ДОПИ по днях тижня (червень), по локаціях ===');
  Object.keys(byLoc).sort().forEach(function(loc){
    var c=byLoc[loc];
    Logger.log('%s | Пн=%s Вт=%s Ср=%s Чт=%s Пт=%s | (Сб=%s Нд=%s)', loc, c[1], c[2], c[3], c[4], c[5], c[6], c[0]);
  });
}

function diagOranzhDopDetail(loc){
  loc = loc || 'Оранж';
  function dkey(v){
    if(v instanceof Date){ return Utilities.formatDate(v, 'Europe/Kiev', 'yyyy-MM-dd'); }
    var ss=String(v||'').trim();
    if(/^\d{4}-\d{2}-\d{2}/.test(ss)) return ss.slice(0,10);
    var m=ss.match(/^(\d{1,2})[.\/](\d{1,2})[.\/](\d{4})/);
    if(m) return m[3]+'-'+('0'+m[2]).slice(-2)+'-'+('0'+m[1]).slice(-2);
    return ss;
  }
  var wd=['Нд','Пн','Вт','Ср','Чт','Пт','Сб'];
  var dsh=_getAttendanceSheet(false);
  var dv=dsh.getDataRange().getValues();
  var byAct={};
  for(var j=1;j<dv.length;j++){
    var rec=_parseAttendanceRow(dv[j]);
    if(String(rec.loc||'').trim()!==loc) continue;
    var dd=dkey(rec.date); if(dd<'2026-06-01'||dd>'2026-06-30') continue;
    var act=String(rec.activityName||'?').trim();
    if(!byAct[act]) byAct[act]={};
    byAct[act][dd]=(byAct[act][dd]||0)+1;
  }
  Logger.log('=== %s | заняття × дати (червень) ===', loc);
  Object.keys(byAct).sort().forEach(function(act){
    var dates=Object.keys(byAct[act]).sort();
    var parts=dates.map(function(d){ var pp=d.split('-'); var dt=new Date(+pp[0],+pp[1]-1,+pp[2]); return wd[dt.getDay()]+' '+d.slice(8)+'.06('+byAct[act][d]+')'; });
    Logger.log('%s: %s', act, parts.join('  '));
  });
}

function diagCatalog(loc){
  loc = loc || 'Оранж';
  Logger.log('=== %s | КАТАЛОГ (джерело правди) ===', loc);
  var cat=getActivitiesCatalog(loc);
  if(cat&&cat.items){
    cat.items.forEach(function(a){
      Logger.log('  %s | ціна=%s | ставка=%s | %s | %s', a.name, a.clientPrice, a.teacherRate, a.payType, (a.active?'активне':'—'));
    });
    Logger.log('  (всього занять: %s)', cat.items.length);
  } else { Logger.log('  помилка каталогу'); }
  function dkey(v){ if(v instanceof Date){ return Utilities.formatDate(v,'Europe/Kiev','yyyy-MM-dd'); } var ss=String(v||'').trim(); if(/^\d{4}-\d{2}-\d{2}/.test(ss)) return ss.slice(0,10); var m=ss.match(/^(\d{1,2})[.\/](\d{1,2})[.\/](\d{4})/); if(m) return m[3]+'-'+('0'+m[2]).slice(-2)+'-'+('0'+m[1]).slice(-2); return ss; }
  Logger.log('=== %s | ЦІНИ В ПОЗНАЧКАХ (червень) — по них рахується ===', loc);
  var dsh=_getAttendanceSheet(false); var dv=dsh.getDataRange().getValues();
  var byAct={};
  for(var j=1;j<dv.length;j++){
    var rec=_parseAttendanceRow(dv[j]);
    if(String(rec.loc||'').trim()!==loc) continue;
    var dd=dkey(rec.date); if(dd<'2026-06-01'||dd>'2026-06-30') continue;
    var act=String(rec.activityName||'?').trim(); var pr=Number(rec.price)||0;
    if(!byAct[act]) byAct[act]={};
    byAct[act][pr]=(byAct[act][pr]||0)+1;
  }
  Object.keys(byAct).sort().forEach(function(act){
    var prices=Object.keys(byAct[act]).map(function(pr){ return pr+'₴×'+byAct[act][pr]; }).join('  ');
    Logger.log('  %s: %s', act, prices);
  });
}

function diagRequisites(){
  var sh=SpreadsheetApp.openById(CONFIG_SHEET_ID).getSheetByName('Реквізити_Локацій');
  if(!sh){ Logger.log('немає аркуша'); return; }
  var v=sh.getDataRange().getValues();
  var hdr=v[0].map(function(x,i){ return i+':'+String(x); });
  Logger.log('=== ЗАГОЛОВКИ ===');
  Logger.log(hdr.join(' | '));
  var linkCol=-1;
  for(var i=0;i<v[0].length;i++){ if(String(v[0][i]).indexOf('Посилання')>=0){ linkCol=i; break; } }
  Logger.log('=== колонка Посилання_на_оплату = індекс %s (реадер очікує 6) ===', linkCol);
  Logger.log('=== ЛОКАЦІЯ | ТИП | ПОСИЛАННЯ ===');
  for(var r=1;r<v.length;r++){
    var loc=String(v[r][0]||'').trim(); if(!loc) continue;
    var typ=String(v[r][1]||'').trim();
    var link=linkCol>=0?String(v[r][linkCol]||'').trim():'';
    Logger.log('%s | %s | %s', loc, typ, link||'(порожньо)');
  }
}

function diagReqLookup(){
  ['Осокорки','Оранж'].forEach(function(loc){
    ['studies','extras'].forEach(function(t){
      var r=_getInvoiceRequisites(loc,t);
      if(r&&r.ok){ Logger.log('%s/%s → OK | ЮО=%s | IBAN=%s | link=%s', loc, t, r.name, r.iban, r.payLink||'(порожньо)'); }
      else { Logger.log('%s/%s → ПОМИЛКА: %s', loc, t, r&&r.error); }
    });
  });
}

function diagSigner(namePart){
  namePart=(namePart||'Огієнко').toLowerCase();
  var gc=getClients(); if(!gc.ok){ Logger.log('getClients fail'); return; }
  var list=gc.data||[];
  var found=0;
  for(var i=0;i<list.length;i++){
    var c=list[i];
    var nm=String(c['ПІБ дитини']||'');
    if(nm.toLowerCase().indexOf(namePart)<0) continue;
    found++;
    Logger.log('—— %s | лок=%s', nm, c['Локація']);
    Logger.log('   Підписант договору = [%s]', c['Підписант договору']);
    Logger.log('   ПІБ мами = [%s]', c['ПІБ мами']);
    Logger.log('   ПІБ тата = [%s]', c['ПІБ тата']);
    Logger.log('   Телефон мами = [%s]', c['Телефон мами']);
    var rel=Object.keys(c).filter(function(k){ var kl=k.toLowerCase(); return kl.indexOf('мам')>=0||kl.indexOf('тат')>=0||kl.indexOf('підпис')>=0||kl.indexOf('піб')>=0; });
    Logger.log('   всі релевантні ключі: %s', JSON.stringify(rel));
    rel.forEach(function(k){ var v=String(c[k]||'').trim(); if(v) Logger.log('      [%s] = [%s]', k, v); });
    var d=_invoiceClientData(nm, String(c['Локація']||''), 'studies');
    Logger.log('   ЧИТАЧ рахунку: signerParent=[%s] signerName=[%s] phone=[%s]', d.signerParent, d.signerName, d.signerPhone);
  }
  if(!found) Logger.log('не знайдено: %s', namePart);
}

// v6.57: оцінка «заповненості» картки дитини — щоб при дублях лишати РЕАЛЬНУ картку
// (з договором/сумою/підписантом), а не порожній плейсхолдер. Більший бал = лишаємо.
function _invoiceCardScore(c){
  var s = 0;
  if (String(c['Дата договору']   || '').trim()) s += 8;   // є дата договору
  if (Number(c['Сума договору']) > 0)             s += 4;   // є сума
  if (String(c['Номер договору']  || '').trim()) s += 2;   // є номер договору
  var signer = String(c['Підписант договору'] || '').trim();
  if (signer) s += 1;
  var sn = signer === 'dad' ? c['ПІБ тата'] : c['ПІБ мами'];
  if (String(sn || '').trim()) s += 1;                      // у підписанта є ПІБ
  var st = String(c['Статус'] || '').trim();
  if (st === 'active' || st === 'adaptation') s += 1;       // активний > graduated
  return s;
}

// v6.57: ДЕДУП дубль-карток у межах локації (той самий нормалізований ПІБ). Борг даних:
// частина дітей має по 2 картки → getInvoiceListData давав 2 однакові рядки рахунку
// (paymentSum по імені → однаковий). Лишаємо ОДНУ — з найбільшим _invoiceCardScore.
// Діти БЕЗ дублів (більшість) не зачіпаються (група розміром 1 → повертається як є).
function _dedupInvoiceClients(list){
  var byKey = {};
  list.forEach(function(c){
    var k = _normForMatch(c['ПІБ дитини']);                 // loc уже єдина (list відфільтрований)
    (byKey[k] = byKey[k] || []).push(c);
  });
  var out = [], dropped = [];
  Object.keys(byKey).forEach(function(k){
    var arr = byKey[k];
    if (arr.length === 1){ out.push(arr[0]); return; }       // немає дубля — не чіпаємо
    var best = arr[0], bestScore = _invoiceCardScore(arr[0]);
    for (var i = 1; i < arr.length; i++){
      var sc = _invoiceCardScore(arr[i]);
      if (sc > bestScore){ best = arr[i]; bestScore = sc; }  // строге > → при рівності лишається перший
    }
    out.push(best);
    arr.forEach(function(c){ if (c !== best) dropped.push(String(c['ПІБ дитини'] || '') + '(id=' + String(c['ID'] || '') + ')'); });
  });
  if (dropped.length) Logger.log('[getInvoiceListData] дедуп: прибрано %s дубль-карток: %s', dropped.length, JSON.stringify(dropped));
  return out;
}

// ONE-SHOT ДІАГНОСТИКА (тільки читання): усі дубль-картки по мережі — id/договір/fee/
// статус/у Payment/score. Запусти в редакторі → View → Executions → Logs.
function diagDuplicateClients(){
  var res = getClients(); if (!res.ok){ Logger.log('❌ getClients fail'); return res; }
  var rows = res.data || [];
  var present = {};
  var paySheet = getCRMSpreadsheet().getSheetByName(SHEET_YEARLY);
  if (paySheet){
    var pv = paySheet.getDataRange().getValues();
    if (pv.length >= 2){
      var h = pv[0].map(String), ni = h.indexOf('Ім\'я дитини'), li = h.indexOf('Локація');
      if (ni >= 0 && li >= 0) for (var r = 1; r < pv.length; r++){ present[_normForMatch(pv[r][ni]) + '|' + _normForMatch(pv[r][li])] = true; }
    }
  }
  var map = {};
  rows.forEach(function(c){
    var nm = String(c['ПІБ дитини'] || '').trim(); if (!nm) return;
    var k = _normForMatch(nm) + '|' + _normForMatch(c['Локація']);
    (map[k] = map[k] || []).push(c);
  });
  var groups = Object.keys(map).filter(function(k){ return map[k].length > 1; }).sort();
  Logger.log('═══ ДУБЛЬ-КАРТКИ КЛІЄНТІВ: %s дітей ═══', groups.length);
  groups.forEach(function(k){
    var arr = map[k];
    Logger.log('▸ %s | %s | ×%s | у Payment=%s', String(arr[0]['ПІБ дитини'] || ''), String(arr[0]['Локація'] || '').trim(), arr.length, present[k] ? 'так' : 'ні');
    var best = arr[0], bs = _invoiceCardScore(arr[0]);
    arr.forEach(function(c){ var sc = _invoiceCardScore(c); if (sc > bs){ best = c; bs = sc; } });
    arr.forEach(function(c){
      Logger.log('    %s id=%s | договір=%s | fee=%s | статус=%s | №дог=%s | оновл=%s | score=%s',
        (c === best ? '✅ЛИШИТИ' : '🗑видалити'), String(c['ID'] || ''),
        String(c['Дата договору'] || '').slice(0, 10) || '—', Number(c['Сума договору']) || 0,
        String(c['Статус'] || ''), String(c['Номер договору'] || '') || '—',
        String(c['Оновлено'] || '') || '—', _invoiceCardScore(c));
    });
  });
  Logger.log('✅ЛИШИТИ = найбільший score (заповнена). 🗑 = видалити через deleteClient(id).');
  return { ok: true, groups: groups.length };
}

function getInvoiceListData(params){
  try {
    var loc      = String(params.loc      || '').trim();
    var payMonth = Number(params.payMonth || 0);
    var payYear  = Number(params.payYear  || 0);
    var extMonth = Number(params.extMonth || 0);
    var extYear  = Number(params.extYear  || 0);
    if (!loc) return {ok:false, error:'Параметр loc обовʼязковий'};
    if (payMonth < 1 || payMonth > 12) return {ok:false, error:'payMonth має бути 1-12'};
    if (extMonth < 1 || extMonth > 12) return {ok:false, error:'extMonth має бути 1-12'};

    // === 1. Клієнти (фільтр: loc + status IN active/adaptation) ===
    var crmRes = getClients();
    if (!crmRes.ok) return crmRes;
    var clientsAll = crmRes.data || [];
    var clients = clientsAll.filter(function(c){
      if (String(c['Локація'] || '').trim() !== loc) return false;
      var st = String(c['Статус'] || '').trim();
      // Fix A: graduated теж кандидат — у циклі нижче лишаємо лише тих, хто має
      // додаткові за extMonth (extrasBreakdownSum > 0). "Голі" випускники відсіються.
      return st === 'active' || st === 'adaptation' || st === 'graduated';
    });

    // v6.57: дедуп дубль-карток (той самий ПІБ у цій loc) → 1 дитина = 1 рядок рахунку.
    // Лишаємо найбільш заповнену картку; діти без дублів не зачіпаються.
    clients = _dedupInvoiceClients(clients);

    // === 2. Payment per-child з "Оплати-Рік" (поля "<Місяць>-Бюджет-навч" + "-доп") ===
    var paymentByName = {};
    var extrasByPayment = {};   // v6.11.16: "<Місяць>-Бюджет-доп" — загальна сума додаткових ДО СПЛАТИ
    var paySheet = getCRMSpreadsheet().getSheetByName(SHEET_YEARLY);
    var budNavchCol = MONTHS_CAL[payMonth - 1] + '-Бюджет-навч';
    var budDopCol   = MONTHS_CAL[payMonth - 1] + '-Бюджет-доп';
    if (paySheet){
      var pvals = paySheet.getDataRange().getValues();
      if (pvals.length >= 2){
        var phdrs = pvals[0].map(function(h){ return String(h); });
        var nameIdx = phdrs.indexOf('Ім\'я дитини');
        var locIdx  = phdrs.indexOf('Локація');
        var budIdx  = phdrs.indexOf(budNavchCol);
        var dopIdx  = phdrs.indexOf(budDopCol);
        if (nameIdx < 0 || locIdx < 0 || budIdx < 0){
          Logger.log('[getInvoiceListData] Оплати-Рік headers missing: name=%s loc=%s bud=%s (col="%s")',
            nameIdx, locIdx, budIdx, budNavchCol);
        } else {
          if (dopIdx < 0) Logger.log('[getInvoiceListData] ⚠ колонку "%s" не знайдено — extrasSum=0 для всіх', budDopCol);
          for (var pr = 1; pr < pvals.length; pr++){
            if (String(pvals[pr][locIdx]).trim() !== loc) continue;
            var pname = String(pvals[pr][nameIdx]).trim();
            if (!pname) continue;
            paymentByName[pname]   = Number(pvals[pr][budIdx]) || 0;
            extrasByPayment[pname] = dopIdx >= 0 ? (Number(pvals[pr][dopIdx]) || 0) : 0;
          }
        }
      }
    } else {
      Logger.log('[getInvoiceListData] Sheet "%s" не знайдено — paymentSum буде 0 для всіх', SHEET_YEARLY);
    }

    // === 3. Extras per-child з "Додаткові_Відвідуваність" (фільтр extMonth/extYear) ===
    var attSh = _getAttendanceSheet(false);
    var attData = attSh.getDataRange().getValues();
    var emm = extMonth < 10 ? '0' + extMonth : String(extMonth);
    var dateFrom = extYear + '-' + emm + '-01';
    var nextE = _nextMonth(extMonth, extYear);
    var enmm = nextE.month < 10 ? '0' + nextE.month : String(nextE.month);
    var dateTo = nextE.year + '-' + enmm + '-01';

    var extrasByChild = {};
    for (var i = 1; i < attData.length; i++){
      var rec = _parseAttendanceRow(attData[i]);
      if (rec.loc !== loc) continue;
      if (rec.date < dateFrom || rec.date >= dateTo) continue;
      if (!extrasByChild[rec.child]){
        extrasByChild[rec.child] = {sum: 0, breakdown: {}};
      }
      var b = extrasByChild[rec.child];
      b.sum += (rec.price || 0);
      if (!b.breakdown[rec.activityName]){
        b.breakdown[rec.activityName] = {name: rec.activityName, count: 0, price: rec.price, total: 0};
      }
      b.breakdown[rec.activityName].count++;
      b.breakdown[rec.activityName].total += (rec.price || 0);
    }

    // === 4. Збираємо response — фільтр paymentSum > 0 OR extrasSum > 0 ===
    var children = [];
    clients.forEach(function(c){
      var name = String(c['ПІБ дитини'] || '').trim();
      if (!name) return;
      // v6.11.15: показуємо ВСІХ matched у Платежах (active+adaptation), навіть з
      // paymentSum=0 (відпустка/переплата) — щоб точно бачити дітей; рахунок 0 →
      // "—" на фронті. Прибрано вимогу paymentSum>0. Критерій = є в Оплати-Рік
      // для цієї loc (АБО має додаткові — щоб ніхто з extras не зник).
      var inPayments = Object.prototype.hasOwnProperty.call(paymentByName, name);
      var paymentSum = paymentByName[name] || 0;
      var extras = extrasByChild[name] || {sum: 0, breakdown: {}};
      var extrasBreakdownSum = extras.sum;                    // сума занять з відмічань (розписка)
      var extrasSum = extrasByPayment[name] || 0;             // v6.11.16: до сплати = "<Місяць>-Бюджет-доп"
      var extrasAdjustment = extrasSum - extrasBreakdownSum;  // борг(+) / переплата(-) / 0
      // Fix A: graduated (випускник) показуємо ЛИШЕ якщо є додаткові за extMonth;
      // "голі" випускники без занять — не показуємо (рахунок тільки на гуртки).
      if (String(c['Статус'] || '').trim() === 'graduated' && extrasBreakdownSum <= 0) return;
      // фільтр НЕ змінено: показуємо matched у Платежах АБО з реальними відмічаннями
      if (!inPayments && extrasBreakdownSum <= 0) return;

      var breakdownArr = Object.keys(extras.breakdown).map(function(k){
        return extras.breakdown[k];
      }).sort(function(a,b){ return a.name.localeCompare(b.name, 'uk'); });

      var signer = String(c['Підписант договору'] || 'mom').trim() || 'mom';
      var signerName  = signer === 'dad' ? String(c['ПІБ тата']    || '') : String(c['ПІБ мами']    || '');
      var signerPhone = signer === 'dad' ? String(c['Телефон тата']|| '') : String(c['Телефон мами']|| '');
      var signerEmail = signer === 'dad' ? String(c['Email тата']  || '') : String(c['Email мами']  || '');

      children.push({
        id: String(c['ID'] || ''),
        name: name,
        group: String(c['Група'] || '').trim(),
        status: String(c['Статус'] || '').trim(),
        paymentSum: paymentSum,
        extrasSum: extrasSum,
        extrasBreakdown: breakdownArr,
        extrasBreakdownSum: extrasBreakdownSum,
        extrasAdjustment: extrasAdjustment,
        signerParent: signer,
        signerName:  signerName.trim(),
        signerPhone: signerPhone.trim(),
        signerEmail: signerEmail.trim(),
        contractNumber:           _contractStr(c['Номер договору']),            // v6.11.25 read-guard
        additionalContractNumber: _contractStr(c['Номер додаткового договору'])
      });
    });

    children.sort(function(a,b){ return a.name.localeCompare(b.name, 'uk'); });

    Logger.log('[getInvoiceListData] loc=%s pay=%s/%s ext=%s/%s | active+adapt=%s, shown=%s',
      loc, payMonth, payYear, extMonth, extYear, clients.length, children.length);

    return {
      ok: true,
      loc: loc,
      payMonth: payMonth, payYear: payYear,
      extMonth: extMonth, extYear: extYear,
      totalClients:  clients.length,
      shownChildren: children.length,
      children: children
    };
  } catch(e){
    Logger.log('[getInvoiceListData] EXCEPTION: %s\n%s', e && e.message, e && e.stack);
    return {ok: false, error: String(e && e.message || e)};
  }
}

// Тест wrapper для Apps Script editor (View → Executions → Logs).
function testGetInvoiceListData(){
  var result = getInvoiceListData({loc: 'Бровари', payMonth: 6, payYear: 2026, extMonth: 5, extYear: 2026});
  Logger.log('[testGetInvoiceListData] ok=%s, totalClients=%s, shownChildren=%s',
    result.ok, result.totalClients, result.shownChildren);
  if (result.error) Logger.log('[testGetInvoiceListData] ERROR: %s', result.error);
  Logger.log('[testGetInvoiceListData] children (first 5): %s',
    JSON.stringify((result.children || []).slice(0, 5), null, 2));
  return result;
}

// ═══════════════════════════════════════════════════════════════════════════
// v6.11.17 ЕТАП 1C.1 — ГЕНЕРАЦІЯ PDF-РАХУНКІВ (тільки backend, тільки навчання)
// generateInvoicePDF(opts) → {ok, pdfBase64, pdfFilename, invoiceNumber, sum}
//   opts: {childName, loc, type:'studies'|'extras', month, year, invoiceDate}
// Реквізити — з наявного аркуша "Реквізити_Локацій" (A:Локація B:Тип C:Назва ЮО
//   D:ЄДРПОУ E:IBAN F:Банк G:Посилання H:Logo_URL). Маппінг типу:
//   'studies'→"Навчання", 'extras'→"Додаткові заняття". isFOP = назва з "ФОП".
// Нумерація — аркуш "Лічильники_Рахунків" (per ЄДРПОУ, з 1, атомарно LockService).
// PDF: Utilities.newBlob(html,'text/html').getAs('application/pdf'). Лого — поки
//   CSS-текст "m.kids" (оранж), або <img> якщо в H є Logo_URL.
// КРОК 1C.2 (потім): рядки Борг/Переплата для extras. 1C.3+: фронт/email/viber.
// ═══════════════════════════════════════════════════════════════════════════
function _fmtUahSrv(n){
  n = Math.round(Number(n) || 0); var neg = n < 0; var d = String(Math.abs(n)); var o = '';
  for (var i = 0; i < d.length; i++){ if (i > 0 && (d.length - i) % 3 === 0) o += ' '; o += d.charAt(i); }
  return (neg ? '-' : '') + o;
}

function _invoicePurposeTitle(type, childName, m, y){
  var mn = (m >= 1 && m <= 12) ? MONTHS_CAL[m-1].toLowerCase() : '';
  var base = (type === 'extras')
    ? 'Оплата за організацію освітніх послуг та додаткових занять (гуртків) '
    : 'Оплата за навчання ';
  return base + childName + ', ' + mn + ' ' + y;
}

function _firstName(full){
  var parts = String(full || '').trim().split(/\s+/);
  return parts.length >= 2 ? parts[1] : (parts[0] || '');
}

function _invoicePurposeTitle(type, childName, m, y){
  var mn = (m >= 1 && m <= 12) ? MONTHS_CAL[m-1].toLowerCase() : '';
  var base = (type === 'extras')
    ? 'Оплата за організацію освітніх послуг та додаткових занять (гуртків) '
    : 'Оплата за навчання ';
  return base + childName + ', ' + mn + ' ' + y;
}

function _vocativeUa(name){
  name = String(name||'').trim(); if (!name) return name;
  var low=name.toLowerCase(), last=low.slice(-1), last2=low.slice(-2);
  if (last2==='ія') return name.slice(0,-1)+'є';
  if (last==='а')  return name.slice(0,-1)+'о';
  if (last==='я')  return name.slice(0,-1)+'е';
  if (last2==='ій') return name.slice(0,-1)+'ю';
  if (last==='о')  return name.slice(0,-1)+'е';
  if ('бвгґджзклмнпрстфхцчшщ'.indexOf(last)>=0) return name+'е';
  return name;
}

function invoiceViberMessage(opts){
  opts = opts || {};
  var childName = String(opts.childName || '').trim();
  var loc = String(opts.loc || '').trim();
  if (!childName || !loc) return {ok:false, error:'childName/loc'};
  var types = [];
  if (opts.sendStudies) types.push({t:'studies', m:Number(opts.payMonth || opts.month || 0), y:Number(opts.payYear || opts.year || 0)});
  if (opts.sendExtras)  types.push({t:'extras',  m:Number(opts.extMonth || opts.month || 0), y:Number(opts.extYear || opts.year || 0)});
  if (!types.length) return {ok:false, error:'Немає сум'};
  var messages = [], errs = [], grand = 0;
  for (var i = 0; i < types.length; i++){
    var ty = types[i];
    // v6.92: generateInvoicePDF напряму — без створення файлу на Drive (PDF-лінк у Viber не потрібен).
    var r = generateInvoicePDF({childName:childName, loc:loc, type:ty.t, month:ty.m, year:ty.y, invoiceDate:opts.invoiceDate, dataOnly:true});
    if (!r || !r.ok){ errs.push((ty.t === 'extras' ? 'Додаткові' : 'Навчання') + ': ' + ((r && r.error) || '?')); continue; }
    grand += Number(r.sum) || 0;
    var fn = _firstName(r.buyerName);
    var greet = fn ? ('Доброго дня, ' + _vocativeUa(fn) + '! 🌞') : 'Доброго дня! 🌞';
    var title = _invoicePurposeTitle(ty.t, childName, ty.m, ty.y);
    var L = [];
    L.push('*m.kids ' + loc + '*');
    L.push('');
    L.push(greet);
    L.push('Надсилаємо рахунок для оплати.');
    L.push('');
    L.push(title);
    L.push('');
    if (ty.t === 'extras' && r.lines && r.lines.length){
      r.lines.forEach(function(ln){
        if ((Number(ln.sum) || 0) === 0) return;
        L.push('• ' + ln.name + ((Number(ln.qty) || 0) > 1 ? ' ×' + ln.qty : '') + ' — ' + _fmtUahSrv(ln.sum) + ' грн');
      });
      L.push('');
    }
    L.push('Сума до сплати: ' + _fmtUahSrv(r.sum) + ' грн');
    L.push('');
    if (r.payLink){ L.push('Оплатити карткою (натисніть):'); L.push(r.payLink); L.push(''); }
    L.push('Реквізити для оплати:');
    if (r.juName) L.push('Отримувач: ' + r.juName);
    if (r.iban)   L.push('IBAN: ' + r.iban);
    if (r.edrpou) L.push('ЄДРПОУ/РНОКПП: ' + r.edrpou);
    L.push('');
    // v6.92 ТИМЧАСОВО вимкнено PDF у Viber-повідомленні (повернути: SEND_PDF=true)
    var SEND_PDF = false;
    if (SEND_PDF) { L.push('Рахунок (PDF): ' + r.url); }
    L.push('');
    L.push('Дякуємо! 🍊');
    messages.push({type:ty.t, title:title, text:L.join('\n'), sum:r.sum});
  }
  if (!messages.length) return {ok:false, error:(errs.join('; ') || 'Не вдалось')};
  return {ok:true, messages:messages, sum:grand, errors:errs};
}

function testViberMessage(){
  var r = invoiceViberMessage({childName:'Сапогов Рінат', loc:'Осокорки', sendStudies:true, sendExtras:true, payMonth:6, payYear:2026, extMonth:6, extYear:2026, invoiceDate:'01.06.2026'});
  Logger.log('ok=%s | повідомлень=%s', r.ok, r.ok ? r.messages.length : 0);
  if (r.ok) r.messages.forEach(function(m){ Logger.log('\n===== %s =====\n%s\n=====', m.type, m.text); });
  else Logger.log('ERROR: %s', r.error);
}





function _getInvoiceDriveFolder(){
  var name = 'm.kids Рахунки (Viber)';
  var it = DriveApp.getFoldersByName(name);
  return it.hasNext() ? it.next() : DriveApp.createFolder(name);
}

function invoicePdfLink(opts){
  var res = generateInvoicePDF(opts || {});
  if (!res || !res.ok) return {ok:false, error:(res && res.error) || 'PDF не згенеровано'};
  try {
    var folder = _getInvoiceDriveFolder();
    var fname = res.pdfFilename || ('Рахунок.pdf');
    var existing = folder.getFilesByName(fname);
    while (existing.hasNext()) { try { existing.next().setTrashed(true); } catch(e){} }
    var bytes = Utilities.base64Decode(res.pdfBase64);
    var blob = Utilities.newBlob(bytes, 'application/pdf', fname);
    var file = folder.createFile(blob);
    try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(e){}
    return {ok:true, url:file.getUrl(), filename:fname, sum:res.sum, juName:res.juName, edrpou:res.edrpou, iban:res.iban, bank:res.bank, payLink:res.payLink, lines:res.lines, buyerName:res.buyerName};
  } catch(e){
    return {ok:false, error:'Drive: ' + (e.message || e)};
  }
}

function testInvoicePdfLink(){
  var r = invoicePdfLink({childName:'Матущенко Сара', loc:'Осокорки', type:'studies', month:6, year:2026, invoiceDate:'01.06.2026'});
  Logger.log('[testInvoicePdfLink] %s', JSON.stringify(r));
}

function generateInvoicePDF(opts){
  opts = opts || {};
  var childName = String(opts.childName || '').trim();
  var loc       = String(opts.loc || '').trim();
  var type      = (opts.type === 'extras') ? 'extras' : 'studies';
  var month     = Number(opts.month || 0);
  var year      = Number(opts.year  || 0);
  if (!childName) return {ok:false, error:'childName обовʼязковий'};
  if (!loc)       return {ok:false, error:'loc обовʼязковий'};
  if (month < 1 || month > 12) return {ok:false, error:'month має бути 1-12'};

  var invoiceDate = String(opts.invoiceDate || '').trim() ||
    Utilities.formatDate(new Date(), 'Europe/Kiev', 'dd.MM.yyyy');

  var _ts=Date.now(), _tp=_ts;   // TIMING (dataOnly)
  var req = _getInvoiceRequisites(loc, type);
  if(opts.dataOnly){ Logger.log('[TIMING] _getInvoiceRequisites: %s ms', Date.now()-_tp); _tp=Date.now(); }
  if (!req.ok){ Logger.log('[invoicePDF] ❌ %s', req.error); return {ok:false, error:req.error}; }

  var LABEL_DEBT    = 'Борг попереднього періоду';     // v6.11.24: фіксовані лейбли (легко змінити)
  var LABEL_OVERPAY = 'Переплата минулого періоду';

  var client = _invoiceClientData(childName, loc, type);
  if(opts.dataOnly){ Logger.log('[TIMING] _invoiceClientData: %s ms', Date.now()-_tp); _tp=Date.now(); }
  // v6.11.20: підписант ОБОВʼЯЗКОВИЙ (домовленість з директорами) — без fallback.
  if (!client.found){
    return {ok:false, error:'Дитину "' + childName + '" не знайдено в Клієнти для локації "' + loc + '".'};
  }
  if (!client.signerParent){
    return {ok:false, error:'Для дитини "' + childName + '" не заповнено підписанта договору. Заповни в картці перш ніж формувати рахунок.'};
  }
  if (!client.signerName){
    var whoRod = client.signerParent === 'dad' ? 'тата' : 'мами';
    var whoCap = client.signerParent === 'dad' ? 'Тато' : 'Мама';
    return {ok:false, error:'Для дитини "' + childName + '" вказано підписанта (' + whoCap + '), але ПІБ ' + whoRod + ' порожнє. Заповни ПІБ ' + whoRod + ' у картці.'};
  }
  var buyerDisplay = client.signerName;

  // v6.11.24: позиції таблиці (lines) + загальна сума (total).
  var lines = [], total = 0;
  if (type === 'extras'){
    // ФІКС асиметрії місяця (таблиця↔Viber/PDF/email): extras-сума (extrasSum) має читатись
    // з тієї САМОЇ колонки Оплати-Рік, що й таблиця рахунків. Таблиця бере <payMonth>-Бюджет-доп,
    // де payMonth = місяць ВИСТАВЛЕННЯ рахунку. Класи місяця M експортуються в «Бюджет доп» M+1
    // (exportAttendanceToPayments, N+1) → правильна колонка = extMonth+1.
    //   • payMonth беремо з opts (фронт передає той самий, що в таблиці) — пріоритет;
    //   • якщо не передано — дефолт = extMonth+1 (= дефолтний payMonth таблиці);
    //   • extMonth лишаємо = month: • рядки-послуги рахуються наживо з Додаткові_Відвідуваність.
    var _bm = _nextMonth(month, year);
    var billMonth = Number(opts.payMonth) || _bm.month;
    var billYear  = Number(opts.payYear)  || _bm.year;
    var inv = getInvoiceListData({loc: loc, payMonth: billMonth, payYear: billYear, extMonth: month, extYear: year});
    if(opts.dataOnly){ Logger.log('[TIMING] getInvoiceListData(extras): %s ms', Date.now()-_tp); _tp=Date.now(); }
    if (!inv.ok) return {ok:false, error: inv.error};
    var ch = null;
    (inv.children || []).forEach(function(c){ if (String(c.name).trim() === childName) ch = c; });
    if (!ch || (Number(ch.extrasSum) || 0) === 0 || !ch.extrasBreakdown || !ch.extrasBreakdown.length){
      return {ok:false, error:'У дитини "' + childName + '" немає додаткових занять за ' + MONTHS_CAL[month - 1] + ' ' + year + '. Рахунок не формується.'};
    }
    ch.extrasBreakdown.forEach(function(b){
      // v6.50.1: НЕ дублюємо ПІБ дитини в назві послуги (вона вгорі у "Замовник").
      lines.push({name: b.name, qty: b.count, price: b.price, sum: b.total});
    });
    var adj = Number(ch.extrasAdjustment) || 0;
    if (adj > 0)      lines.push({name: LABEL_DEBT,    qty: 1, price: adj, sum: adj});
    else if (adj < 0) lines.push({name: LABEL_OVERPAY, qty: 1, price: adj, sum: adj});
    total = Number(ch.extrasSum) || 0;
  } else {
    var sum = _invoiceSumFromYearly(childName, loc, month, type);
    if(opts.dataOnly){ Logger.log('[TIMING] _invoiceSumFromYearly: %s ms', Date.now()-_tp); _tp=Date.now(); }
    lines.push({name: 'Оплата за навчання', qty: 1, price: sum, sum: sum});
    total = sum;
    // v6.11.26: вступний внесок — план у картці = джерело правди. Додаємо рядки, де
    // paid!==true і дата платежу <= останній день поточного місяця (вкл. прострочені).
    var efSched = client.entryFeeSchedule || [];
    var cutoffISO = _lastDayOfMonth(month, year);
    efSched.forEach(function(p){
      if (!p || p.paid === true) return;
      var iso = _toISO(p.date);
      if (!iso || iso > cutoffISO) return;            // невалідна або майбутня дата → пропустити
      var amt = Number(p.amount) || 0;
      if (amt <= 0) return;
      lines.push({name: 'Вступний внесок', qty: 1, price: amt, sum: amt});
      total += amt;
    });
  }

  // v6.93: dataOnly — повертаємо лише дані для Viber-повідомлення, БЕЗ номера рахунку
  // (_getNextInvoiceNumber інкрементує лічильник) та БЕЗ важкої побудови PDF (HTML→PDF).
  if (opts.dataOnly){
    Logger.log('[TIMING] ВСЬОГО dataOnly (від _getInvoiceRequisites): %s ms', Date.now()-_ts);
    return {ok:true, sum:total, juName:req.name, edrpou:req.edrpou, iban:req.iban,
            bank:req.bank, payLink:req.payLink, lines:lines, buyerName:buyerDisplay,
            sumWords:_numberToUkrainianWords(total)};
  }

  var invoiceNumber = _getNextInvoiceNumber(req.edrpou, req.name);

  var _mLabel = (MONTHS_CAL[month - 1] || '').toLowerCase() + ' ' + year;
  var _title = (type === 'extras')
    ? ('Оплата за організацію освітніх послуг та додаткових занять (гуртків) ' + childName + ', ' + _mLabel)
    : ('Оплата за навчання ' + childName + ', ' + _mLabel);
  var _dueText = String(opts.dueText || '').trim() || 'Оплата до 5 числа поточного місяця';

  var html = _buildInvoiceHtml({
    req: req,
    title: _title,
    dueText: _dueText,
    invoiceNumber: invoiceNumber,
    invoiceDate: invoiceDate,
    buyerName: buyerDisplay,
    contractNumber: client.contractNumber,
    lines: lines,
    total: total,
    sumWords: _numberToUkrainianWords(total)
  });

  var safeName = childName.replace(/[\\/:*?"<>|]/g, '_');
  var pdfFilename = 'Рахунок_' + invoiceNumber + '_' + safeName + '_' + month + '-' + year + '.pdf';
  var blob = Utilities.newBlob(html, 'text/html', pdfFilename).getAs('application/pdf');
  blob.setName(pdfFilename);

  Logger.log('[invoicePDF] OK childName="%s" loc="%s" type=%s №%s сума=%s ЮО="%s"',
    childName, loc, type, invoiceNumber, total, req.name);

  return {
    ok: true,
    pdfBase64: Utilities.base64Encode(blob.getBytes()),
    pdfFilename: pdfFilename,
    invoiceNumber: invoiceNumber,
    sum: total,
    juName: req.name,
    edrpou: req.edrpou,
    iban: req.iban,
    bank: req.bank,
    payLink: req.payLink,
    lines: lines,
    buyerName: buyerDisplay,
    sumWords: _numberToUkrainianWords(total)
  };
}

// ═══════════════════════════════════════════════════════════════════════════
// v6.50 — EMAIL-РОЗСИЛКА РАХУНКІВ (PDF у вкладенні) + ЛОГ "Розсилки_Рахунків"
//   sendInvoiceEmail(opts) — 1 лист підписанту, до 2 PDF (навчання+додаткові,
//     різні ЮО). bulkSendInvoices(opts) — серверний цикл по локації.
//   Лог — аркуш "Розсилки_Рахунків" у CONFIG_SHEET: 1 рядок на КОЖЕН рахунок/ЮО.
// ═══════════════════════════════════════════════════════════════════════════
var INVOICE_LOG_TAB    = 'Розсилки_Рахунків';
var INVOICE_LOG_HEADER = ['Дата+час','Дитина','Локація','Тип','Юрособа',
  'Номер рахунку','Email','Сума','Статус','Хто надіслав','Помилка'];

// Аркуш логу у CONFIG_SHEET (вся фін-інфра разом). Авто-створення з шапкою.
function _getInvoiceLogSheet(){
  var ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  var sh = ss.getSheetByName(INVOICE_LOG_TAB);
  if (!sh){
    sh = ss.insertSheet(INVOICE_LOG_TAB);
    sh.getRange(1, 1, 1, INVOICE_LOG_HEADER.length).setValues([INVOICE_LOG_HEADER]);
    sh.setFrozenRows(1);
  }
  return sh;
}

// row: {child, loc, type, ju, invoiceNumber, email, sum, status, actor, error}
function _logInvoiceSend(row){
  try {
    var sh = _getInvoiceLogSheet();
    var now = Utilities.formatDate(new Date(), 'Europe/Kiev', 'dd.MM.yyyy HH:mm:ss');
    sh.appendRow([
      now,
      String(row.child || ''),
      String(row.loc || ''),
      String(row.type || ''),
      String(row.ju || ''),
      String(row.invoiceNumber || ''),
      String(row.email || ''),
      Number(row.sum || 0),
      String(row.status || ''),
      String(row.actor || ''),
      String(row.error || '')
    ]);
  } catch(e){
    Logger.log('[invLog] append fail: %s', e && e.message);
  }
}

function _invTypeUk(type){ return type === 'extras' ? 'Додаткові' : 'Навчання'; }

// Короткий HTML-лист (перелік рахунків з ЮО/сумою/№).
function _buildInvoiceEmailHtml(childName, monthLabel, invoices){
  var rows = invoices.map(function(inv){
    return '<li><b>' + inv.typeUk + '</b> — ' + _fmtUah(inv.sum) + ' грн' +
      ' <span style="color:#666">(' + inv.juName + ', ЄДРПОУ ' + inv.edrpou +
      ', рахунок №' + inv.invoiceNumber + ')</span></li>';
  }).join('');
  return [
    '<div style="font-family:Arial,sans-serif;font-size:14px;color:#222;line-height:1.6">',
    '<p>Доброго дня!</p>',
    '<p>Надсилаємо рахунки <b>m.kids</b> за <b>' + monthLabel + '</b> для <b>' + childName + '</b>:</p>',
    '<ul>' + rows + '</ul>',
    '<p>Деталі та банківські реквізити — у вкладених PDF-файлах.</p>',
    '<p style="color:#888;font-size:12px">Лист сформовано автоматично системою m.kids.</p>',
    '</div>'
  ].join('');
}

// Один лист підписанту з 1-2 PDF. Логує КОЖЕН рахунок (1 рядок/ЮО).
// opts: {childName, loc, invoiceDate, actorName,
//        payMonth, payYear, extMonth, extYear, sendStudies, sendExtras}
function sendInvoiceEmail(opts){
  opts = opts || {};
  var childName = String(opts.childName || '').trim();
  var loc       = String(opts.loc || '').trim();
  var actor     = String(opts.actorName || '').trim();
  if (!childName) return {ok:false, error:'childName обовʼязковий'};
  if (!loc)       return {ok:false, error:'loc обовʼязковий'};

  var sendStudies = !!opts.sendStudies;
  var sendExtras  = !!opts.sendExtras;
  if (!sendStudies && !sendExtras) return {ok:false, error:'Не вказано що надсилати (sendStudies/sendExtras)'};

  // email підписанта
  var cd = _invoiceClientData(childName, loc, 'studies');
  if (!cd.found) return {ok:false, error:'Дитину "' + childName + '" не знайдено в Клієнти для локації "' + loc + '".'};
  var email = String(cd.signerEmail || '').trim();
  if (!email){
    _logInvoiceSend({child:childName, loc:loc, type:'—', ju:'—', invoiceNumber:'—',
      email:'', sum:0, status:'помилка', actor:actor, error:'нема email підписанта'});
    return {ok:false, error:'У підписанта дитини "' + childName + '" немає email.', noEmail:true};
  }

  var payMonth = Number(opts.payMonth || 0), payYear = Number(opts.payYear || 0);
  var extMonth = Number(opts.extMonth || 0), extYear = Number(opts.extYear || 0);
  var invoiceDate = String(opts.invoiceDate || '').trim();

  var attachments = [], invoices = [], errors = [];
  function buildOne(type, month, year){
    var res = generateInvoicePDF({childName:childName, loc:loc, type:type,
      month:month, year:year, invoiceDate:invoiceDate});
    if (!res.ok){
      errors.push(_invTypeUk(type) + ': ' + res.error);
      _logInvoiceSend({child:childName, loc:loc, type:_invTypeUk(type), ju:'—',
        invoiceNumber:'—', email:email, sum:0, status:'помилка', actor:actor, error:res.error});
      return;
    }
    attachments.push(Utilities.newBlob(Utilities.base64Decode(res.pdfBase64),
      'application/pdf', res.pdfFilename));
    invoices.push({type:type, typeUk:_invTypeUk(type), invoiceNumber:res.invoiceNumber,
      sum:res.sum, juName:res.juName, edrpou:res.edrpou});
  }
  if (sendStudies) buildOne('studies', payMonth, payYear);
  if (sendExtras)  buildOne('extras',  extMonth, extYear);

  if (!attachments.length){
    return {ok:false, error: errors.join('; ') || 'жодного PDF не сформовано'};
  }

  var monthLabel = (payMonth >= 1 && payMonth <= 12) ? (MONTHS_CAL[payMonth-1] + ' ' + payYear)
                 : (extMonth >= 1 && extMonth <= 12) ? (MONTHS_CAL[extMonth-1] + ' ' + extYear) : '';
  var subject = 'Рахунок m.kids — ' + monthLabel + ' — ' + childName;
  var html    = _buildInvoiceEmailHtml(childName, monthLabel, invoices);
  var plain   = 'Доброго дня! Надсилаємо рахунки m.kids за ' + monthLabel + ' для ' + childName +
                '. Деталі — у вкладених PDF.';

  try {
    MailApp.sendEmail(email, subject, plain, {htmlBody: html, attachments: attachments, name: 'm.kids'});
  } catch(e){
    var em = String(e && e.message || e);
    invoices.forEach(function(inv){
      _logInvoiceSend({child:childName, loc:loc, type:inv.typeUk, ju:inv.juName,
        invoiceNumber:inv.invoiceNumber, email:email, sum:inv.sum,
        status:'помилка', actor:actor, error:'MailApp: ' + em});
    });
    return {ok:false, error:'MailApp: ' + em};
  }

  var partial = errors.length ? ('частково: ' + errors.join('; ')) : '';
  invoices.forEach(function(inv){
    _logInvoiceSend({child:childName, loc:loc, type:inv.typeUk, ju:inv.juName,
      invoiceNumber:inv.invoiceNumber, email:email, sum:inv.sum,
      status:'надіслано', actor:actor, error:partial});
  });

  Logger.log('[sendInvoiceEmail] OK "%s" → %s | %s рахунків%s',
    childName, email, invoices.length, errors.length ? (' (частково: ' + errors.join('; ') + ')') : '');
  return {ok:true, email:email, sent:invoices.length, invoices:invoices,
    partialErrors: errors.length ? errors : undefined};
}

// Серверний цикл по локації (для крон/одноразового запуску). Фронт «Надіслати всім»
// робить власний цикл (для live-прогресу), теж через sendInvoiceEmail.
// opts: {loc, payMonth, payYear, extMonth, extYear, invoiceDate, actorName}
function bulkSendInvoices(opts){
  opts = opts || {};
  var loc      = String(opts.loc || '').trim();
  var payMonth = Number(opts.payMonth || 0), payYear = Number(opts.payYear || 0);
  var extMonth = Number(opts.extMonth || 0), extYear = Number(opts.extYear || 0);
  var invoiceDate = String(opts.invoiceDate || '').trim();
  var actor    = String(opts.actorName || '').trim();
  if (!loc) return {ok:false, error:'loc обовʼязковий'};

  var inv = getInvoiceListData({loc:loc, payMonth:payMonth, payYear:payYear,
    extMonth:extMonth, extYear:extYear});
  if (!inv.ok) return inv;

  var sent = 0, failed = 0, noEmail = 0, details = [];
  (inv.children || []).forEach(function(c){
    var hasStudies = (Number(c.paymentSum) || 0) > 0;
    var hasExtras  = (Number(c.extrasSum)  || 0) > 0;
    if (!hasStudies && !hasExtras) return;
    if (!c.signerEmail){ noEmail++; details.push({name:c.name, status:'без email'}); return; }
    var r = sendInvoiceEmail({childName:c.name, loc:loc, invoiceDate:invoiceDate, actorName:actor,
      payMonth:payMonth, payYear:payYear, extMonth:extMonth, extYear:extYear,
      sendStudies:hasStudies, sendExtras:hasExtras});
    if (r.ok){ sent++; details.push({name:c.name, status:'надіслано', sent:r.sent}); }
    else     { failed++; details.push({name:c.name, status:'помилка', error:r.error}); }
  });

  Logger.log('[bulkSendInvoices] loc=%s | sent=%s failed=%s noEmail=%s', loc, sent, failed, noEmail);
  return {ok:true, loc:loc, sent:sent, failed:failed, noEmail:noEmail, details:details};
}

// ТЕСТ: один реальний лист (редактор Apps Script → Run → перевір Inbox + лог-аркуш).
// Заміни childName/loc на свою тестову дитину з email-підписантом перед запуском.
function testSendInvoiceEmailOne(){
  var res = sendInvoiceEmail({
    childName: 'Матущенко Сара', loc: 'Осокорки',
    payMonth: 6, payYear: 2026, extMonth: 5, extYear: 2026,
    invoiceDate: '01.06.2026', actorName: 'TEST',
    sendStudies: true, sendExtras: false
  });
  Logger.log('[testSendInvoiceEmailOne] %s', JSON.stringify(res));
  return res;
}

// ═══════════════════════════════════════════════════════════════════════════
// v6.50.3 — СТАТУСИ РОЗСИЛКИ (бейджі) + ЗВІТ ДЛЯ CFO
//   logViberSent(opts) — авто-лог кліку 💬 у новий аркуш "Viber_Розсилки".
//   getInvoiceStatusReport(opts) — per-child статуси (для бейджів) + агрегат
//     по локаціях (для CFO-звіту). Email-статус з "Розсилки_Рахунків",
//     Viber-статус з "Viber_Розсилки". Фільтр по місяцю надсилання.
// ═══════════════════════════════════════════════════════════════════════════
var VIBER_LOG_TAB    = 'Viber_Розсилки';
var VIBER_LOG_HEADER = ['Дата+час','Дитина','Локація','Телефон','Директор'];

function _getViberLogSheet(){
  var ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  var sh = ss.getSheetByName(VIBER_LOG_TAB);
  if (!sh){
    sh = ss.insertSheet(VIBER_LOG_TAB);
    sh.getRange(1, 1, 1, VIBER_LOG_HEADER.length).setValues([VIBER_LOG_HEADER]);
    sh.setFrozenRows(1);
  }
  return sh;
}

// Авто-позначка Viber при кліку 💬 (фронт шле POST). opts: {childName, loc, phone, actorName}
function logViberSent(opts){
  opts = opts || {};
  var child = String(opts.childName || '').trim();
  var loc   = String(opts.loc || '').trim();
  if (!child || !loc) return {ok:false, error:'childName+loc обовʼязкові'};
  try {
    var sh = _getViberLogSheet();
    var now = Utilities.formatDate(new Date(), 'Europe/Kiev', 'dd.MM.yyyy HH:mm:ss');
    sh.appendRow([now, child, loc, String(opts.phone || ''), String(opts.actorName || '')]);
    return {ok:true, ts:now};
  } catch(e){
    return {ok:false, error:String(e && e.message || e)};
  }
}

// Статуси + агрегат. opts: {loc?, month?(1-12), year?} — month/year фільтрують по ДАТІ надсилання.
function getInvoiceStatusReport(opts){
  try {
    opts = opts || {};
    var locFilter = String(opts.loc || '').trim();
    var month = Number(opts.month || 0);
    var year  = Number(opts.year  || 0);

    function inPeriod(tsStr){
      if (!month && !year) return true;
      var m = String(tsStr || '').match(/^(\d{2})\.(\d{2})\.(\d{4})/);
      if (!m) return false;
      if (month && Number(m[2]) !== month) return false;
      if (year  && Number(m[3]) !== year)  return false;
      return true;
    }

    // ── EMAIL-лог ("Розсилки_Рахунків") ──
    var emailByKey = {};   // "child||loc" → {status, ts}
    var errorsByLoc = {};
    var ve = _getInvoiceLogSheet().getDataRange().getValues();
    for (var r = 1; r < ve.length; r++){
      var ts = String(ve[r][0] || ''), child = String(ve[r][1] || '').trim();
      var loc = String(ve[r][2] || '').trim(), status = String(ve[r][8] || '').trim();
      if (!child || !inPeriod(ts)) continue;
      var key = child + '||' + loc;
      if (status === 'надіслано'){
        emailByKey[key] = {status:'надіслано', ts:ts};   // останній надісланий «перебиває» помилку
      } else if (status === 'помилка'){
        errorsByLoc[loc] = (errorsByLoc[loc] || 0) + 1;
        if (!emailByKey[key]) emailByKey[key] = {status:'помилка', ts:ts};
      }
    }

    // ── VIBER-лог ("Viber_Розсилки") ──
    var viberByKey = {};
    var vv = _getViberLogSheet().getDataRange().getValues();
    for (var r2 = 1; r2 < vv.length; r2++){
      var ts2 = String(vv[r2][0] || ''), child2 = String(vv[r2][1] || '').trim();
      var loc2 = String(vv[r2][2] || '').trim();
      if (!child2 || !inPeriod(ts2)) continue;
      viberByKey[child2 + '||' + loc2] = {ts:ts2};
    }

    // ── per-child для поточної локації (бейджі на сторінці Рахунків) ──
    var byChild = {};
    if (locFilter){
      Object.keys(emailByKey).forEach(function(k){
        var p = k.split('||'); if (p[1] !== locFilter) return;
        (byChild[p[0]] = byChild[p[0]] || {}).email = emailByKey[k];
      });
      Object.keys(viberByKey).forEach(function(k){
        var p = k.split('||'); if (p[1] !== locFilter) return;
        (byChild[p[0]] = byChild[p[0]] || {}).viber = viberByKey[k];
      });
    }

    // ── кількість активних дітей по локаціях ──
    var kidsByLoc = {};
    var gc = getClients();
    if (gc.ok){
      (gc.data || []).forEach(function(c){
        var st = String(c['Статус'] || '').trim();
        if (st !== 'active' && st !== 'adaptation') return;
        var loc = String(c['Локація'] || '').trim();
        if (loc) kidsByLoc[loc] = (kidsByLoc[loc] || 0) + 1;
      });
    }

    // ── distinct надіслано по локаціях ──
    var emailSentByLoc = {}, viberSentByLoc = {};
    Object.keys(emailByKey).forEach(function(k){
      if (emailByKey[k].status !== 'надіслано') return;
      var loc = k.split('||')[1]; emailSentByLoc[loc] = (emailSentByLoc[loc] || 0) + 1;
    });
    Object.keys(viberByKey).forEach(function(k){
      var loc = k.split('||')[1]; viberSentByLoc[loc] = (viberSentByLoc[loc] || 0) + 1;
    });

    var locSet = {};
    [kidsByLoc, emailSentByLoc, viberSentByLoc, errorsByLoc].forEach(function(o){
      Object.keys(o).forEach(function(l){ locSet[l] = 1; });
    });
    var report = Object.keys(locSet).map(function(l){
      var kids = kidsByLoc[l] || 0, es = emailSentByLoc[l] || 0, vs = viberSentByLoc[l] || 0;
      return {loc:l, kids:kids, emailSent:es, viberSent:vs, errors:errorsByLoc[l] || 0,
        emailPct: kids ? Math.round(es * 100 / kids) : 0,
        viberPct: kids ? Math.round(vs * 100 / kids) : 0};
    }).sort(function(a,b){ return a.loc.localeCompare(b.loc, 'uk'); });

    var tot = {loc:'УСЯ МЕРЕЖА', kids:0, emailSent:0, viberSent:0, errors:0};
    report.forEach(function(x){ tot.kids += x.kids; tot.emailSent += x.emailSent; tot.viberSent += x.viberSent; tot.errors += x.errors; });
    tot.emailPct = tot.kids ? Math.round(tot.emailSent * 100 / tot.kids) : 0;
    tot.viberPct = tot.kids ? Math.round(tot.viberSent * 100 / tot.kids) : 0;

    return {ok:true, loc:locFilter, month:month, year:year, byChild:byChild, report:report, total:tot};
  } catch(e){
    return {ok:false, error:String(e && e.message || e)};
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// v6.51 — КОНТРОЛЬ ЗАПОВНЕННЯ (CFO): за дату показує по ВСІХ локаціях, де
//   заповнено відвідуваність дітей (по групах) і табель співробітників.
//   ОДНЕ читання аркуша "Табель". Дату нормалізуємо самі (tz-safe), фільтр на
//   сервері. Повертаємо компактний JSON (весь Табель НЕ віддаємо).
// ═══════════════════════════════════════════════════════════════════════════
// Порт isTeacherTypeForTabel (index.html) — виключення предметників зі штату.
var _FILL_TABEL_CORE_RE = /вихователь|помічник|директор|дирек|медсестр|психолог|тьютор|тімлі|охорон|водій|кухар|прибиральниц|посудомийниц|фінансист|юрист|методист|заступ|зам |сео|смм|маркет|тренер|операці|академічн|раннього розвитку|вчитель/;
var _FILL_TABEL_SUBJ_RE = /англійськ|польськ|іспанськ|французськ|німецьк|музик|фітнес|логопед|танц|айкідо|робототехнік|інформатик|хореограф|вокал|малюванн|шах/;
var _FILL_TABEL_SUBJ_EXACT = ['арт','лего','театр','карате','каратэ'];
function _fillIsTeacher(typ, pos){
  var t = String(typ || '').trim().toLowerCase();
  if (/предметник/.test(t) || /додаткових/.test(t) || t === 'subject_teacher' || t === 'extras_teacher') return true;
  var p = String(pos || '').trim().toLowerCase();
  if (!p) return false;
  var toks = p.split(/[\\\/,]+/).map(function(s){ return s.trim(); }).filter(Boolean);
  if (toks.some(function(x){ return _FILL_TABEL_CORE_RE.test(x); })) return false;   // основна роль → лишаємо
  return toks.some(function(x){ return _FILL_TABEL_SUBJ_RE.test(x) || _FILL_TABEL_SUBJ_EXACT.indexOf(x) !== -1; });
}
// Активний співробітник (порт isActive/isFired/декрет).
function _fillEmpActive(e){
  if (e.fired && String(e.fired).trim()) return false;
  if (e.stat && /звіл/i.test(e.stat)) return false;
  if (e.stat && /декрет/i.test(e.stat)) return false;
  return true;
}
// Нормалізація дати в ISO БЕЗ зсуву таймзони: Date-обʼєкти форматуємо у tz таблиці.
function _fillNormD(v, tz){
  if (v instanceof Date) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  var s = String(v == null ? '' : v).trim();
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
  var m = s.match(/^(\d{2})\.(\d{2})\.(\d{4})/);
  if (m) return m[3] + '-' + m[2] + '-' + m[1];
  var d = new Date(s);
  return isNaN(d) ? '' : Utilities.formatDate(d, tz, 'yyyy-MM-dd');
}
// Вихідний / держ.свято (фіксовані дати; рухомі — Великдень/Трійця — не враховуємо).
var _FILL_UA_HOLIDAYS = {
  '01-01':'Новий рік','03-08':'8 Березня','05-01':'День праці','05-09':'День памʼяті',
  '06-28':'День Конституції','08-24':'День Незалежності','10-01':'День захисників','12-25':'Різдво'
};
function _fillNonWorking(dateISO){
  var p = dateISO.split('-');
  var dw = new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2])).getDay();  // weekday TZ-незалежний
  if (dw === 0 || dw === 6) return {nonWorking:true, reason:'Вихідний'};
  var md = p[1] + '-' + p[2];
  if (_FILL_UA_HOLIDAYS[md]) return {nonWorking:true, reason:_FILL_UA_HOLIDAYS[md]};
  return {nonWorking:false, reason:''};
}

// Очікувані групи/діти/персонал (спільне для single-day і range).
// expGroups[loc]={group->activeKids}; expKids[loc]=DISTINCT діти; expStaff[loc]=персонал(−предметники).
// v6.64.2: швидке читання Табеля — лише останні ~50k рядків (cols A-H),
// з відкатом на повне читання, якщо запит старший за вікно.
function _fillReadAttendance(attSh, tz, neededFrom){
  if (!attSh) return [];
  var lastRow = attSh.getLastRow();
  if (lastRow < 2) return [];
  var cap = 50000;
  var startRow = Math.max(2, lastRow - cap + 1);
  var av = attSh.getRange(startRow, 1, lastRow - startRow + 1, 8).getValues();
  if (startRow > 2 && neededFrom){
    var earliest = null;
    for (var k = 0; k < av.length; k++){
      var dd = _fillNormD(av[k][0], tz);
      if (dd && (earliest === null || dd < earliest)) earliest = dd;
    }
    if (earliest && earliest > neededFrom){
      av = attSh.getRange(2, 1, lastRow - 1, 8).getValues();
    }
  }
  return av;
}

function _fillExpected(){
  var expGroups = {}, expKids = {}, expStaff = {};
// v6.64.1: читаємо лише потрібні колонки клієнтів (без JSON Розвиток/Здоров'я) — інакше getFillStatus читає мегабайти JSON і тормозить.
  var _cs = getCRMSpreadsheet().getSheetByName(SHEET_CLIENTS);
  if (_cs && _cs.getLastRow() >= 2){
    var _hdr = _cs.getRange(1, 1, 1, _cs.getLastColumn()).getValues()[0];
    var _ci = function(n){ for (var i = 0; i < _hdr.length; i++){ if (String(_hdr[i]).trim() === n) return i; } return -1; };
    var _iId = _ci('ID'), _iSt = _ci('Статус'), _iLoc = _ci('Локація'), _iGr = _ci('Група');
    if (_iSt >= 0 && _iLoc >= 0){
      var _maxC = Math.max(_iId, _iSt, _iLoc, _iGr) + 1;
      var _cv = _cs.getRange(1, 1, _cs.getLastRow(), _maxC).getValues();
      for (var _r = 1; _r < _cv.length; _r++){
        if (_iId >= 0 && !_cv[_r][_iId]) continue;
        var st = String(_cv[_r][_iSt] || '').trim();
        if (st !== 'active' && st !== 'adaptation') continue;
        var loc = String(_cv[_r][_iLoc] || '').trim();
        if (!loc) continue;
        var grp = (_iGr >= 0 ? String(_cv[_r][_iGr] || '').trim() : '') || '—';
        (expGroups[loc] = expGroups[loc] || {});
        expGroups[loc][grp] = (expGroups[loc][grp] || 0) + 1;
        expKids[loc] = (expKids[loc] || 0) + 1;
      }
    }
  }
  var hrSh = SpreadsheetApp.openById(HR_SHEET_ID).getSheetByName(HR_TAB_NAME);
  if (hrSh){
    var hr = hrSh.getDataRange().getValues();
    for (var i = 1; i < hr.length; i++){
      if (!hr[i] || (!hr[i][4] && !hr[i][5])) continue;   // нема ПІБ
      var e = _parseEmpRow(hr[i], i + 1);
      if (!_fillEmpActive(e)) continue;
      if (_fillIsTeacher(e.typ, e.pos)) continue;
      if (e.loc) expStaff[e.loc] = (expStaff[e.loc] || 0) + 1;
    }
  }
  return {expGroups:expGroups, expKids:expKids, expStaff:expStaff};
}

// «Коли» формату 'dd.MM.yyyy HH:mm' → сортувальний ключ 'yyyyMMddHHmm' (для «найпізніший час»).
// Невідомий/порожній формат → '' (програє будь-якому реальному timestamp).
function _fillAtKey(at){
  var m = String(at || '').match(/^(\d{2})\.(\d{2})\.(\d{4})\s+(\d{2}):(\d{2})/);
  return m ? (m[3] + m[2] + m[1] + m[4] + m[5]) : '';
}

// Список ISO-дат from..to включно (TZ-незалежний крок по компонентах).
function _fillDateList(from, to){
  var list = [];
  var s = from.split('-'), e = to.split('-');
  var d   = new Date(Number(s[0]), Number(s[1]) - 1, Number(s[2]));
  var end = new Date(Number(e[0]), Number(e[1]) - 1, Number(e[2]));
  var guard = 0;
  while (d <= end && guard++ < 400){
    var y = d.getFullYear(), m = d.getMonth() + 1, dd = d.getDate();
    list.push(y + '-' + (m < 10 ? '0' + m : m) + '-' + (dd < 10 ? '0' + dd : dd));
    d.setDate(d.getDate() + 1);
  }
  return list;
}

// v6.51.3: режим діапазону (Тиждень/Місяць). ОДНЕ читання "Табель" на from..to
// (без циклу по днях). По кожній локації окремо діти/табель: daysFilled(🟢),
// daysPartial(🟡), daysEmpty(🔴), workingDays, byDay[{date,childrenStatus,timesheetStatus}].
function _getFillStatusRange(from, to){
  var ss = getCRMSpreadsheet();
  var tz = ss.getSpreadsheetTimeZone() || 'Europe/Kiev';
  var exp = _fillExpected();
  var dayList = _fillDateList(from, to);

  // Одне читання Табель → byDate[date][loc] = {kidIds:{}, staffIds:{}} (distinct id).
  var byDate = {};
  var attSh = ss.getSheetByName(SHEET_ATTENDANCE);
  if (attSh){
    var av = _fillReadAttendance(attSh, tz, from);
    for (var r = 0; r < av.length; r++){
      var row = av[r];
      var d = _fillNormD(row[0], tz);
      if (!d || d < from || d > to) continue;          // ISO-рядки → лексикографічне порівняння
      var id = String(row[1] || '').trim();
      var loc = String(row[3] || '').trim();
      var status = String(row[5] || '').trim();
      if (!loc || !status) continue;                   // «заповнено» = непорожній Статус
      var dd = (byDate[d] = byDate[d] || {});
      var cell = (dd[loc] = dd[loc] || {kidIds:{}, staffIds:{}});
      if (id.indexOf('STAFF::') === 0) cell.staffIds[id] = true;
      else cell.kidIds[id] = true;
    }
  }

  // Локації: union очікуваних + побачених у Табелі, зонний порядок.
  var allLocs = {};
  Object.keys(exp.expGroups).forEach(function(l){ allLocs[l] = 1; });
  Object.keys(exp.expStaff).forEach(function(l){ allLocs[l] = 1; });
  Object.keys(byDate).forEach(function(d){ Object.keys(byDate[d]).forEach(function(l){ allLocs[l] = 1; }); });
  var locList = sortByLocationOrder(Object.keys(allLocs));

  // Робочих днів у діапазоні (однаково для всіх локацій).
  var _todayISO = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  var workingDays = 0;
  dayList.forEach(function(d){ if (d <= _todayISO && !_fillNonWorking(d).nonWorking) workingDays++; });

  function rs(filled, working, partial){
    if (working === 0) return 'gray';
    if (filled >= working) return 'green';
    if (filled === 0 && partial === 0) return 'red';
    return 'yellow';
  }

  var summary = {mode:'range', childrenGreen:0, childrenTotal:0,
                 timesheetGreen:0, timesheetTotal:0, red:0, workingDays:workingDays};
  var locations = locList.map(function(loc){
    var cF = 0, cP = 0, cE = 0, tF = 0, tP = 0, tE = 0;
    var ckExp = exp.expKids[loc] || 0, sExp = exp.expStaff[loc] || 0;
    var byDay = dayList.map(function(d){
      if (d > _todayISO) return {date:d, future:true, childrenStatus:'gray', timesheetStatus:'gray'};
      if (_fillNonWorking(d).nonWorking) return {date:d, nw:true, childrenStatus:'gray', timesheetStatus:'gray'};
      var cell = byDate[d] && byDate[d][loc];
      var ckMarked = cell ? Object.keys(cell.kidIds).length : 0;
      var cs = ckMarked === 0 ? 'red' : 'green';
      if (cs === 'green') cF++; else if (cs === 'yellow') cP++; else cE++;
      var smk = cell ? Object.keys(cell.staffIds).length : 0;
      var ts = smk === 0 ? 'red' : 'green';
      if (ts === 'green') tF++; else if (ts === 'yellow') tP++; else tE++;
      return {date:d, nw:false, childrenStatus:cs, timesheetStatus:ts};
    });
    var cStatus = rs(cF, workingDays, cP), tStatus = rs(tF, workingDays, tP);
    summary.childrenTotal++; summary.timesheetTotal++;
    if (cStatus === 'green') summary.childrenGreen++;
    if (tStatus === 'green') summary.timesheetGreen++;
    if (cStatus === 'red') summary.red++;
    if (tStatus === 'red') summary.red++;
    return {loc:loc,
      children:{status:cStatus, daysFilled:cF, daysPartial:cP, daysEmpty:cE, workingDays:workingDays, expected:ckExp},
      timesheet:{status:tStatus, daysFilled:tF, daysPartial:tP, daysEmpty:tE, workingDays:workingDays, expected:sExp},
      byDay:byDay};
  });

  return {ok:true, mode:'range', from:from, to:to, nonWorking:false,
          workingDays:workingDays, summary:summary, locations:locations};
}

// params: {date:'YYYY-MM-DD'} (День) АБО {from, to} (Тиждень/Місяць → range-режим).
function getFillStatus(params){
  try {
    params = params || {};
    var date = String(params.date || '').trim();
    var from = String(params.from || '').trim();
    var to   = String(params.to   || '').trim();

    // v6.51.3: якщо передано валідні from&to → режим діапазону.
    if (/^\d{4}-\d{2}-\d{2}$/.test(from) && /^\d{4}-\d{2}-\d{2}$/.test(to)){
      return _getFillStatusRange(from, to);
    }

    if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) return {ok:false, error:'param date=YYYY-MM-DD обовʼязковий'};

    var ss = getCRMSpreadsheet();
    var tz = ss.getSpreadsheetTimeZone() || 'Europe/Kiev';

    // 1+2) Очікувані групи/діти/персонал (спільний хелпер з range-режимом).
    var _exp = _fillExpected();
    var expGroups = _exp.expGroups, expKids = _exp.expKids, expStaff = _exp.expStaff;

    // 3) Одне читання "Табель" → marked діти/персонал за дату (дедуп last-wins по id)
    var attSh = ss.getSheetByName(SHEET_ATTENDANCE);
    var kidMarked = {};      // loc -> group -> {ids:{}, by, at}
    var markedKidIds = {};   // v6.51.1: loc -> {childId:true} (distinct, для location-level)
    var kidWho = {};         // v6.51.4: loc -> {set:{}, cnt, bestKey, bestBy, bestAt} (хто+коли відмітив дітей)
    var staffMarked = {};    // loc -> {ids:{}, by, at}
    if (attSh){
      var av = _fillReadAttendance(attSh, tz, date);
      for (var r = 0; r < av.length; r++){
        var row = av[r];
        if (_fillNormD(row[0], tz) !== date) continue;
        var id = String(row[1] || '').trim();
        var loc = String(row[3] || '').trim();
        var status = String(row[5] || '').trim();
        if (!loc || !status) continue;                 // «заповнено» = непорожній Статус
        var by = String(row[6] || '').trim(), at = String(row[7] || '').trim();
        if (id.indexOf('STAFF::') === 0){
          var so = (staffMarked[loc] = staffMarked[loc] || {ids:{}, by:'', at:''});
          so.ids[id] = true; if (by) so.by = by; if (at) so.at = at;
        } else {
          var g = String(row[4] || '').trim() || '—';
          var kl = (kidMarked[loc] = kidMarked[loc] || {});
          var kg = (kl[g] = kl[g] || {ids:{}, by:'', at:''});
          kg.ids[id] = true; if (by) kg.by = by; if (at) kg.at = at;
          (markedKidIds[loc] = markedKidIds[loc] || {})[id] = true;   // v6.51.1 distinct по локації
          // v6.51.4: location-level «хто + коли» по реальних дитячих рядках (поле Ким).
          if (by){
            var kw = (kidWho[loc] = kidWho[loc] || {set:{}, cnt:0, bestKey:null, bestBy:'', bestAt:''});
            if (!kw.set[by]){ kw.set[by] = true; kw.cnt++; }
            var atk = _fillAtKey(at);
            if (kw.bestKey === null || atk >= kw.bestKey){ kw.bestKey = atk; kw.bestBy = by; kw.bestAt = at; }
          }
        }
      }
    }

    // 4) Збираємо по локаціях у LOCATION_ORDER
    var allLocs = {};
    [expGroups, expStaff, kidMarked, staffMarked].forEach(function(o){
      Object.keys(o).forEach(function(l){ allLocs[l] = 1; });
    });
    var locList = sortByLocationOrder(Object.keys(allLocs));

    var summary = {childrenGreen:0, childrenTotal:0, groupsGreen:0, groupsTotal:0,
                   timesheetGreen:0, timesheetTotal:0, red:0};
    var locations = locList.map(function(loc){
      // v6.51.1: location-level відвідуваність — DISTINCT childId (не сума по групах).
      var ckExp = expKids[loc] || 0;
      var ckMarked = markedKidIds[loc] ? Object.keys(markedKidIds[loc]).length : 0;
      var ckStatus = ckMarked === 0 ? 'red' : 'green';
      summary.childrenTotal++;
      if (ckStatus === 'green') summary.childrenGreen++;

      var grpNames = {};
      Object.keys(expGroups[loc] || {}).forEach(function(g){ grpNames[g] = 1; });
      Object.keys(kidMarked[loc]  || {}).forEach(function(g){ grpNames[g] = 1; });
      var groups = Object.keys(grpNames).sort(function(a,b){ return a.localeCompare(b, 'uk'); }).map(function(g){
        var expected = (expGroups[loc] && expGroups[loc][g]) || 0;
        var mk = kidMarked[loc] && kidMarked[loc][g];
        var marked = mk ? Object.keys(mk.ids).length : 0;
        var status = marked === 0 ? 'red' : 'green';
        summary.groupsTotal++;
        if (status === 'green') summary.groupsGreen++;
        if (status === 'red') summary.red++;
        return {group:g, teacher:(mk && mk.by) || '', status:status, marked:marked,
                expected:expected, by:(mk && mk.by) || '', at:(mk && mk.at) || ''};
      });
      var expS = expStaff[loc] || 0;
      var sm = staffMarked[loc];
      var smk = sm ? Object.keys(sm.ids).length : 0;
      var tStatus = smk === 0 ? 'red' : 'green';
      summary.timesheetTotal++;
      if (tStatus === 'green') summary.timesheetGreen++;
      if (tStatus === 'red') summary.red++;
      var kw = kidWho[loc];
      return {loc:loc,
              children:{status:ckStatus, marked:ckMarked, expected:ckExp,             // v6.51.1 location-level
                        by:(kw && kw.bestBy) || '', at:(kw && kw.bestAt) || '',       // v6.51.4 хто+коли (найпізніший)
                        byCount:(kw && kw.cnt) || 0},
              groups:groups,
              timesheet:{status:tStatus, marked:smk, expected:expS, by:(sm && sm.by) || '', at:(sm && sm.at) || ''}};
    });

    var nw = _fillNonWorking(date);
    return {ok:true, date:date, nonWorking:nw.nonWorking, nonWorkingReason:nw.reason,
            summary:summary, locations:locations};
  } catch(e){
    return {ok:false, error:String(e && e.message || e)};
  }
}

// ТЕСТ (Apps Script editor → Run): сьогоднішня дата. Перевір, що _normD не ріже
// дати через таймзону (порівняй із реальними відмітками у листі "Табель").
function testGetFillStatus(){
  var d = Utilities.formatDate(new Date(), 'Europe/Kiev', 'yyyy-MM-dd');
  var r = getFillStatus({date: d});
  Logger.log('[testGetFillStatus] date=%s ok=%s locations=%s summary=%s',
    d, r.ok, r.locations && r.locations.length, JSON.stringify(r.summary));
  return r;
}

// v6.11.25 read-guard: договір-поля могли зберегтись як Date (Sheets auto-coerce
// дато-схожого тексту). Захист на ВИВОДІ: Date → ''; інакше String(v).trim().
// Дані в БД НЕ змінює — лише не дає Date потрапити у рахунок/UI.
function _contractStr(v){
  if (v instanceof Date) return '';
  return String(v == null ? '' : v).trim();
}

// Реквізити по (Локація × Тип). Маппінг 'studies'→"Навчання", 'extras'→"Додаткові заняття".
function _getInvoiceRequisites(loc, type){
  var typeUk = (type === 'extras') ? 'Додаткові заняття' : 'Навчання';
  // v6.11.18 fix: Реквізити_Локацій живуть у CONFIG_SHEET (поруч з OPEX/Salary), НЕ в CRM_SHEET.
  var sh = SpreadsheetApp.openById(CONFIG_SHEET_ID).getSheetByName('Реквізити_Локацій');
  if (!sh) return {ok:false, error:'Аркуш "Реквізити_Локацій" не знайдено в CONFIG_SHEET'};
  var vals = sh.getDataRange().getValues();
  for (var r = 0; r < vals.length; r++){
    if (_normForMatch(vals[r][0]) !== _normForMatch(loc)) continue;
    var _tv = String(vals[r][1] || '').trim().toLowerCase();
    var _isExtras = _tv.indexOf('гуртк') >= 0 || _tv.indexOf('додаткових занять') >= 0 || _tv === 'додаткові заняття';
    var _isStudies = !_isExtras && _tv.indexOf('навчання') >= 0;
    if ((type === 'extras') ? !_isExtras : !_isStudies) continue;   // гнучке зіставлення типу (старі/нові назви)
    var name = String(vals[r][2] || '').trim();
    return {
      ok: true,
      loc: loc, type: typeUk,
      name: name,
      edrpou: String(vals[r][3] || '').trim(),
      iban:   String(vals[r][4] || '').trim(),
      bank:   String(vals[r][5] || '').trim(),
      payLink: String(vals[r][6] || '').trim(),
      logoUrl: String(vals[r][7] || '').trim(),
      isFOP: name.toUpperCase().indexOf('ФОП') === 0
    };
  }
  return {ok:false, error:'Не знайдено реквізити для loc="' + loc + '" тип="' + typeUk + '"'};
}

// Підписант (mom/dad → ПІБ) + номер договору з картки Клієнти.
function _invoiceClientData(childName, loc, type){
  var res = {found:false, signerParent:'', signerName:'', signerEmail:'', signerPhone:'', contractNumber:'', entryFeeSchedule:[]};
  var gc = getClients();
  if (!gc.ok) return res;
  var nn = String(childName).trim(), ll = String(loc).trim();
  var list = gc.data || [];
  // v6.85: серед можливих ДУБЛІВ картки обираємо ту, де заповнений ПІБ підписанта
  // (порожні чернетки авто-синку з Оплати-Рік не блокують рахунок).
  var picked = null, firstMatch = null;
  for (var i = 0; i < list.length; i++){
    var cc = list[i];
    if (_normForMatch(cc['ПІБ дитини']) !== _normForMatch(nn)) continue;
    if (_normForMatch(cc['Локація'])    !== _normForMatch(ll)) continue;
    if (!firstMatch) firstMatch = cc;
    var sg = String(cc['Підписант договору'] || '').trim();
    var snm = (sg === 'dad') ? String(cc['ПІБ тата'] || '').trim()
            : (sg === 'mom') ? String(cc['ПІБ мами'] || '').trim() : '';
    if (snm){ picked = cc; break; }
  }
  var c = picked || firstMatch;
  if (!c) return res;
  res.found = true;
  var signer = String(c['Підписант договору'] || '').trim();
  res.signerParent = signer;
  res.signerName = (signer === 'dad') ? String(c['ПІБ тата'] || '').trim()
                 : (signer === 'mom') ? String(c['ПІБ мами'] || '').trim()
                 : '';
  res.signerEmail = (signer === 'dad') ? String(c['Email тата']   || '').trim()
                  : (signer === 'mom') ? String(c['Email мами']   || '').trim() : '';
  res.signerPhone = (signer === 'dad') ? String(c['Телефон тата'] || '').trim()
                  : (signer === 'mom') ? String(c['Телефон мами'] || '').trim() : '';
  res.contractNumber = (type === 'extras')
    ? _contractStr(c['Номер додаткового договору'])
    : _contractStr(c['Номер договору']);
  try { var efs = JSON.parse(c['Графік внеску (JSON)'] || '[]'); res.entryFeeSchedule = Array.isArray(efs) ? efs : []; }
  catch(e){ res.entryFeeSchedule = []; }
  return res;
}

// Сума з Оплати-Рік: studies → "<Місяць>-Бюджет-навч", extras → "-Бюджет-доп".
function _invoiceSumFromYearly(childName, loc, month, type){
  var sh = getCRMSpreadsheet().getSheetByName(SHEET_YEARLY);
  if (!sh) return 0;
  var vals = sh.getDataRange().getValues();
  if (vals.length < 2) return 0;
  var h = vals[0].map(function(x){ return String(x); });
  var nameIdx = h.indexOf("Ім'я дитини");
  var locIdx  = h.indexOf('Локація');
  var col = (type === 'extras' ? '-Бюджет-доп' : '-Бюджет-навч');
  var budIdx = h.indexOf(MONTHS_CAL[month - 1] + col);
  if (nameIdx < 0 || locIdx < 0 || budIdx < 0){
    Logger.log('[invoicePDF] ⚠ Оплати-Рік: name=%s loc=%s bud=%s (col="%s")',
      nameIdx, locIdx, budIdx, MONTHS_CAL[month - 1] + col);
    return 0;
  }
  var nn = String(childName).trim(), ll = String(loc).trim();
  for (var r = 1; r < vals.length; r++){
    if (String(vals[r][locIdx]).trim() !== ll) continue;
    if (String(vals[r][nameIdx]).trim() !== nn) continue;
    return Number(vals[r][budIdx]) || 0;
  }
  return 0;
}

// Атомарний наступний номер per ЄДРПОУ з аркуша "Лічильники_Рахунків". З 1.
function _getNextInvoiceNumber(edrpou, juName){
  edrpou = String(edrpou || '').trim();
  var lock = LockService.getScriptLock();
  try { lock.waitLock(30000); } catch(e){ Logger.log('[invoiceNum] lock fail: %s', e); }
  try {
    // v6.11.18 fix: лічильник тримаємо у CONFIG_SHEET (все фінансове разом).
    var ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
    var sh = ss.getSheetByName('Лічильники_Рахунків');
    if (!sh){
      sh = ss.insertSheet('Лічильники_Рахунків');
      sh.appendRow(['ЄДРПОУ', 'Назва ЮО', 'Останній_номер', 'Дата_останнього']);
      sh.setFrozenRows(1);
    }
    var vals = sh.getDataRange().getValues();
    var now = Utilities.formatDate(new Date(), 'Europe/Kiev', 'dd.MM.yyyy HH:mm');
    for (var r = 1; r < vals.length; r++){
      if (String(vals[r][0]).trim() === edrpou){
        var next = (Number(vals[r][2]) || 0) + 1;
        sh.getRange(r + 1, 3).setValue(next);
        sh.getRange(r + 1, 4).setValue(now);
        return next;
      }
    }
    sh.appendRow([edrpou, juName, 1, now]);
    return 1;
  } finally {
    try { lock.releaseLock(); } catch(e){}
  }
}

// v6.11.26 хелпери вступного: нормалізація дати в ISO / показ DD.MM.YYYY / кінець місяця.
function _toISO(d){
  d = String(d == null ? '' : d).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(d)) return d;
  return parseDateDMY(d);   // DD.MM.YYYY → YYYY-MM-DD ('' якщо не розпізнано)
}
function _isoToDMY(iso){
  var p = String(iso || '').split('-');
  return p.length === 3 ? (p[2] + '.' + p[1] + '.' + p[0]) : String(iso || '');
}
function _lastDayOfMonth(month, year){
  var last = new Date(year, month, 0).getDate();   // month 1-based → останній день місяця
  return year + '-' + ('0' + month).slice(-2) + '-' + ('0' + last).slice(-2);
}

// Формат грошей: 1234.5 → "1 234,50" (nbsp як роздільник тисяч).
function _fmtUah(n){
  n = Number(n) || 0;
  var s = n.toFixed(2).split('.');
  s[0] = s[0].replace(/\B(?=(\d{3})+(?!\d))/g, ' ');
  return s[0] + ',' + s[1];
}

// HTML-шаблон рахунку (inline-CSS — зовнішній CSS у newBlob→PDF ненадійний).
function _buildInvoiceHtml(d){
  var req = d.req;
  var logoHtml = req.logoUrl
    ? '<img src="' + req.logoUrl + '" style="height:48px"/>'
    : '<div style="display:inline-block;background:#FF6A00;color:#fff;font-weight:700;font-size:22px;padding:8px 16px;border-radius:8px;letter-spacing:.5px;">m.kids</div>';
  var taxLine = req.isFOP ? '' : '<div>Не є платником податку на прибуток на загальних підставах</div>';
  var contractLine = d.contractNumber ? '<div class="party"><b>Договір:</b> ' + d.contractNumber + '</div>' : '';
  var rowsHtml = (d.lines || []).map(function(ln, i){
    return '    <tr><td class="c">' + (i + 1) + '</td><td>' + ln.name + '</td><td class="c">' + ln.qty +
           '</td><td class="num">' + _fmtUah(ln.price) + '</td><td class="num">' + _fmtUah(ln.sum) + '</td></tr>';
  }).join('\n');
  var totalStr = _fmtUah(d.total);
  var lineCount = (d.lines || []).length;

  return [
'<!doctype html><html><head><meta charset="utf-8"><style>',
'  * { font-family: Arial, Helvetica, sans-serif; }',
'  body { color:#111; font-size:12px; margin:24px; }',
'  .head { display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:14px; }',
'  .req-box { border:1px solid #999; padding:8px 12px; font-size:11px; line-height:1.5; max-width:62%; }',
'  .req-box .title { font-weight:700; margin-bottom:4px; }',
'  h1 { font-size:16px; border-bottom:2px solid #111; padding-bottom:6px; margin:14px 0 10px; }',
'  .party { margin:6px 0; line-height:1.5; }',
'  table { width:100%; border-collapse:collapse; margin-top:12px; font-size:11px; }',
'  th,td { border:1px solid #555; padding:6px 8px; }',
'  th { background:#f0f0f0; text-align:center; }',
'  td.num { text-align:right; } td.c { text-align:center; }',
'  .total { text-align:right; font-weight:700; font-size:13px; margin-top:10px; }',
'  .summary { margin-top:8px; line-height:1.6; }',
'</style></head><body>',
'  <h1>' + d.title + '</h1>',
'  <div class="party">',
'    <div><b>Постачальник:</b> ' + req.name + '</div>',
'    <div>ЄДРПОУ: ' + req.edrpou + '</div>',
'    <div>IBAN: ' + req.iban + '</div>',
'    <div>Банк: ' + req.bank + '</div>',
'    ' + taxLine,
'  </div>',
'  <div class="party"><b>Замовник:</b> ' + d.buyerName + '</div>',
'  ' + contractLine,
'  <table>',
'    <tr><th>№</th><th>Товари (роботи, послуги)</th><th>Кіл-ть</th><th>Ціна</th><th>Сума</th></tr>',
rowsHtml,
'  </table>',
'  <div class="total">Разом: ' + totalStr + ' грн.</div>',
'  <div class="summary">Всього найменувань ' + lineCount + ', на суму ' + totalStr + ' грн.</div>',
'  <div class="summary"><b>Сума прописом:</b> ' + d.sumWords + '</div>',
(d.dueText ? '  <div class="summary" style="margin-top:12px"><b>Термін оплати:</b> ' + d.dueText + '</div>' : ''),
'</body></html>'
  ].join('\n');
}

// Сума прописом українською (грн жін.рід + копійки). Підтримує до мільйонів.
function _numberToUkrainianWords(amount){
  amount = Number(amount) || 0;
  var grn = Math.floor(amount + 1e-9);
  var kop = Math.round((amount - grn) * 100);
  if (kop === 100){ grn += 1; kop = 0; }

  var ones   = ["", "один","два","три","чотири","п'ять","шість","сім","вісім","дев'ять"];
  var onesF  = ["", "одна","дві","три","чотири","п'ять","шість","сім","вісім","дев'ять"];
  var teens  = ["десять","одинадцять","дванадцять","тринадцять","чотирнадцять","п'ятнадцять","шістнадцять","сімнадцять","вісімнадцять","дев'ятнадцять"];
  var tens   = ["","","двадцять","тридцять","сорок","п'ятдесят","шістдесят","сімдесят","вісімдесят","дев'яносто"];
  var hund   = ["","сто","двісті","триста","чотириста","п'ятсот","шістсот","сімсот","вісімсот","дев'ятсот"];

  function triad(num, fem){
    var w = [], h = Math.floor(num / 100), t = Math.floor((num % 100) / 10), o = num % 10;
    if (h) w.push(hund[h]);
    if (t === 1){ w.push(teens[o]); }
    else { if (t) w.push(tens[t]); if (o) w.push(fem ? onesF[o] : ones[o]); }
    return w.join(' ');
  }
  function plural(num, forms){
    var n = num % 100;
    if (n >= 11 && n <= 14) return forms[2];
    var dd = num % 10;
    if (dd === 1) return forms[0];
    if (dd >= 2 && dd <= 4) return forms[1];
    return forms[2];
  }

  var parts = [];
  var mil = Math.floor(grn / 1000000);
  var thou = Math.floor((grn % 1000000) / 1000);
  var rest = grn % 1000;
  if (mil){  parts.push(triad(mil, false)); parts.push(plural(mil, ["мільйон","мільйони","мільйонів"])); }
  if (thou){ parts.push(triad(thou, true)); parts.push(plural(thou, ["тисяча","тисячі","тисяч"])); }
  if (rest || (!mil && !thou)) parts.push(triad(rest, true));

  var words = parts.join(' ').replace(/\s+/g, ' ').trim();
  if (!words) words = 'нуль';
  words = words.charAt(0).toUpperCase() + words.slice(1);
  var grnLabel = plural(grn, ["гривня","гривні","гривень"]);
  var kopStr = (kop < 10 ? '0' + kop : '' + kop);
  var kopLabel = plural(kop, ["копійка","копійки","копійок"]);
  return words + ' ' + grnLabel + ' ' + kopStr + ' ' + kopLabel;
}

// ТЕСТ: Гайдай Коля / Осокорки / червень 2026 / навчання → PDF у Drive (root), URL у лог.
function testGenerateInvoiceStudies(){
  var res = generateInvoicePDF({childName: 'Матущенко Сара', loc: 'Осокорки', type: 'studies', month: 6, year: 2026, invoiceDate: '01.06.2026'});
  if (!res.ok){ Logger.log('[testInvoice] ❌ %s', res.error); return res; }
  var bytes = Utilities.base64Decode(res.pdfBase64);
  var blob = Utilities.newBlob(bytes, 'application/pdf', res.pdfFilename);
  var file = DriveApp.createFile(blob);
  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(e){}
  Logger.log('[testInvoice] ✅ №%s | сума=%s (%s) | ЮО="%s" | покупець="%s"',
    res.invoiceNumber, res.sum, res.sumWords, res.juName, res.buyerName);
  Logger.log('[testInvoice] PDF URL:     %s', file.getUrl());
  Logger.log('[testInvoice] download:    %s', 'https://drive.google.com/uc?export=download&id=' + file.getId());
  return res;
}

// ───────────────────────────────────────────────────────────────────────────
// v6.11.26 ТЕСТ вступного: записує план Матущенко Сарі (2×5000₴, 15.06 + 15.07.2026,
// paid:false). Бекап попереднього "Графік внеску (JSON)" у PropertiesService для
// відкату. Дати в ISO (як у картці). Запобіжник dryRun+confirm.
// ───────────────────────────────────────────────────────────────────────────
var TEST_ENTRYFEE_BACKUP_KEY = 'TEST_ENTRYFEE_BACKUP';

function createTestEntryFeePlan(opts){
  opts = opts || {};
  var dryRun = (opts.dryRun !== false);
  if (!dryRun && opts.confirm !== 'YES_WRITE'){
    Logger.log('[testEF] ⚠ REAL заблоковано: бракує confirm="YES_WRITE". Працюю як DRY-RUN.');
    dryRun = true;
  }
  var name = 'Матущенко Сара', loc = 'Осокорки';
  var plan = [
    {date:'2026-06-15', amount:5000, note:'тест вступний', paid:false},
    {date:'2026-07-15', amount:5000, note:'тест вступний', paid:false}
  ];

  var sh = getCRMSpreadsheet().getSheetByName(SHEET_CLIENTS);
  if (!sh){ Logger.log('[testEF] ❌ "%s" не знайдено', SHEET_CLIENTS); return {ok:false}; }
  var vals = sh.getDataRange().getValues();
  var hdr = vals[0].map(function(x){ return String(x); });
  var nameIdx = hdr.indexOf('ПІБ дитини'), locIdx = hdr.indexOf('Локація');
  var efIdx = hdr.indexOf('Графік внеску (JSON)');
  if (efIdx < 0){ Logger.log('[testEF] ❌ колонка "Графік внеску (JSON)" не знайдена'); return {ok:false}; }

  var rowNum = -1, prev = '';
  for (var r = 1; r < vals.length; r++){
    if (String(vals[r][nameIdx]).trim() === name && String(vals[r][locIdx]).trim() === loc){
      rowNum = r + 1; prev = String(vals[r][efIdx] || ''); break;
    }
  }
  if (rowNum < 0){ Logger.log('[testEF] ❌ "%s" (%s) не знайдено в Клієнти', name, loc); return {ok:false}; }

  Logger.log('[testEF] ═══════════════════════════════════════');
  Logger.log('[testEF] dryRun=%s | row %s | поточний графік: %s', dryRun, rowNum, prev || '(порожньо)');
  Logger.log('[testEF] НОВИЙ план: %s', JSON.stringify(plan));

  if (dryRun){
    Logger.log('[testEF] DRY-RUN — нічого не записано. Реально: runCreateTestEntryFeePlan');
    Logger.log('[testEF] ═══════════════════════════════════════');
    return {ok:true, dryRun:true, willWrite:plan};
  }
  PropertiesService.getScriptProperties().setProperty(TEST_ENTRYFEE_BACKUP_KEY, prev);
  sh.getRange(rowNum, efIdx + 1).setValue(JSON.stringify(plan));
  Logger.log('[testEF] ✅ Записано план (бекап збережено). Далі: testGenerateInvoiceStudies');
  Logger.log('[testEF] ═══════════════════════════════════════');
  return {ok:true, dryRun:false, row:rowNum, backup:prev, wrote:plan};
}

function rollbackTestEntryFeePlan(opts){
  opts = opts || {};
  var dryRun = (opts.dryRun !== false);
  if (!dryRun && opts.confirm !== 'YES_WRITE'){
    Logger.log('[rbEF] ⚠ REAL заблоковано: бракує confirm="YES_WRITE". Працюю як DRY-RUN.');
    dryRun = true;
  }
  var name = 'Матущенко Сара', loc = 'Осокорки';
  var sh = getCRMSpreadsheet().getSheetByName(SHEET_CLIENTS);
  if (!sh){ Logger.log('[rbEF] ❌ "%s" не знайдено', SHEET_CLIENTS); return {ok:false}; }
  var vals = sh.getDataRange().getValues();
  var hdr = vals[0].map(function(x){ return String(x); });
  var nameIdx = hdr.indexOf('ПІБ дитини'), locIdx = hdr.indexOf('Локація');
  var efIdx = hdr.indexOf('Графік внеску (JSON)');
  var rowNum = -1;
  for (var r = 1; r < vals.length; r++){
    if (String(vals[r][nameIdx]).trim() === name && String(vals[r][locIdx]).trim() === loc){ rowNum = r + 1; break; }
  }
  if (rowNum < 0){ Logger.log('[rbEF] ❌ "%s" не знайдено', name); return {ok:false}; }

  var backup = PropertiesService.getScriptProperties().getProperty(TEST_ENTRYFEE_BACKUP_KEY);
  var restore = (backup == null) ? '' : backup;
  Logger.log('[rbEF] ═══════════════════════════════════════');
  Logger.log('[rbEF] dryRun=%s | row %s | відновити на: %s', dryRun, rowNum, restore || '(порожньо)');
  if (dryRun){
    Logger.log('[rbEF] DRY-RUN — нічого не змінено. Реально: runRollbackTestEntryFeePlan');
    Logger.log('[rbEF] ═══════════════════════════════════════');
    return {ok:true, dryRun:true, willRestore:restore};
  }
  sh.getRange(rowNum, efIdx + 1).setValue(restore);
  PropertiesService.getScriptProperties().deleteProperty(TEST_ENTRYFEE_BACKUP_KEY);
  Logger.log('[rbEF] ✅ Відновлено попередній стан графіку');
  Logger.log('[rbEF] ═══════════════════════════════════════');
  return {ok:true, dryRun:false, restored:restore};
}

function runCreateTestEntryFeePlan(){ return createTestEntryFeePlan({dryRun:false, confirm:'YES_WRITE'}); }
function runRollbackTestEntryFeePlan(){ return rollbackTestEntryFeePlan({dryRun:false, confirm:'YES_WRITE'}); }

// ───────────────────────────────────────────────────────────────────────────
// v6.11.22 ДІАГНОСТИКА (КРОК 1C.2): шукає дітей локації з додатковими за місяць,
// придатних для тесту extras-рахунку. Read-only. Для кожної дитини з extras>0
// перевіряє підписника через _invoiceClientData (як реальна валідація
// generateInvoicePDF) → позначає ✅ готова / ⚠ нема підписника.
// Запуск: testDiagFindExtrasCandidates().
// ───────────────────────────────────────────────────────────────────────────
function diagFindExtrasCandidates(loc, extMonth, extYear){
  loc = String(loc || 'Осокорки').trim();
  extMonth = Number(extMonth || 5);
  extYear  = Number(extYear  || 2026);

  Logger.log('[diagExtras] ═══════════════════════════════════════');
  Logger.log('[diagExtras] loc="%s" extras %s/%s', loc, extMonth, extYear);

  var inv = getInvoiceListData({loc: loc, payMonth: extMonth, payYear: extYear, extMonth: extMonth, extYear: extYear});
  if (!inv.ok){ Logger.log('[diagExtras] ❌ getInvoiceListData: %s', inv.error); return inv; }

  var withExtras = (inv.children || []).filter(function(c){
    return (Number(c.extrasSum) || 0) > 0 || (Number(c.extrasBreakdownSum) || 0) > 0;
  });
  Logger.log('[diagExtras] Дітей з extras (extrasSum>0 або breakdownSum>0): %s з %s показаних',
    withExtras.length, inv.shownChildren);
  Logger.log('[diagExtras] ─────────────────────────────────────');

  var ready = [];
  withExtras.forEach(function(c){
    var cd = _invoiceClientData(c.name, loc, 'extras');
    var signerOk = !!(cd.signerParent && cd.signerName);
    if (signerOk) ready.push(c.name);
    var bd = (c.extrasBreakdown || []).map(function(b){
      return b.name + ' ' + b.count + '×' + b.price + '=' + b.total;
    }).join('; ');
    Logger.log('[diagExtras] %s "%s"', signerOk ? '✅' : '⚠', c.name);
    Logger.log('   extrasSum=%s | breakdownSum=%s | adjustment=%s',
      c.extrasSum, c.extrasBreakdownSum, c.extrasAdjustment);
    Logger.log('   заняття: %s', bd || '(порожньо)');
    Logger.log('   підписант: %s | дод.договір="%s"',
      signerOk ? (cd.signerParent + ' / ' + cd.signerName)
               : ('⚠ НЕ заповнено (signerParent="' + cd.signerParent + '" name="' + cd.signerName + '")'),
      cd.contractNumber);
  });

  Logger.log('[diagExtras] ─────────────────────────────────────');
  Logger.log('[diagExtras] ГОТОВІ до extras-рахунку (extras>0 + підписант): %s', ready.length);
  ready.forEach(function(n){ Logger.log('   • %s', n); });
  Logger.log('[diagExtras] ═══════════════════════════════════════');

  return {ok:true, totalWithExtras: withExtras.length, ready: ready};
}

function testDiagFindExtrasCandidates(){
  return diagFindExtrasCandidates('Осокорки', 5, 2026);
}

// ───────────────────────────────────────────────────────────────────────────
// v6.11.23 ДІАГНОСТИКА БЛОКЕРА: чому extrasBreakdown порожній у getInvoiceListData.
// Перевіряє гіпотези: A) дані по іншому місяцю; B) name+loc не матчиться;
// C) схема аркуша змінилась; D) ціна 0. Read-only. Дзеркалить точну логіку
// extrasByChild з getInvoiceListData (точний match child===ПІБ, loc, вікно дат).
// Запуск: testDiagExtrasBreakdown().
// ───────────────────────────────────────────────────────────────────────────
function diagExtrasBreakdownForChild(name, loc, month, year){
  name  = String(name || '').trim();
  loc   = String(loc  || '').trim();
  month = Number(month || 5);
  year  = Number(year  || 2026);
  var nameNorm = name.replace(/\s+/g, ' ').toLowerCase();

  Logger.log('[diagBrk] ═══════════════════════════════════════');
  Logger.log('[diagBrk] "%s" | loc="%s" | %s/%s', name, loc, month, year);

  // 1) Дитина в Клієнти (точне написання ПІБ)
  var gc = getClients();
  var clientName = '', clientLoc = '';
  if (gc.ok){
    (gc.data || []).forEach(function(c){
      var n = String(c['ПІБ дитини'] || '');
      if (n.replace(/\s+/g, ' ').toLowerCase().indexOf(nameNorm) !== -1 && String(c['Локація'] || '').trim() === loc){
        clientName = n; clientLoc = String(c['Локація'] || '').trim();
      }
    });
  }
  Logger.log('[diagBrk] 1) Клієнти: name="%s" (len=%s) loc="%s"', clientName, clientName.length, clientLoc);
  if (!clientName) Logger.log('[diagBrk]   ⚠ дитину не знайдено в Клієнти для loc="%s"', loc);

  // 2) Аркуш Додаткові_Відвідуваність + вікно дат (точно як getInvoiceListData)
  var sh;
  try { sh = _getAttendanceSheet(false); } catch(e){ Logger.log('[diagBrk] ❌ %s', e); return {ok:false, error:String(e)}; }
  var data = sh.getDataRange().getValues();
  Logger.log('[diagBrk] 2) Аркуш "%s": %s рядків (з шапкою)', ATTENDANCE_SHEET_NAME, data.length);
  Logger.log('[diagBrk]   Заголовки: %s', JSON.stringify((data[0] || []).map(function(x){ return String(x); })));
  var emm = month < 10 ? '0' + month : '' + month;
  var dateFrom = year + '-' + emm + '-01';
  var nx = _nextMonth(month, year);
  var nmm = nx.month < 10 ? '0' + nx.month : '' + nx.month;
  var dateTo = nx.year + '-' + nmm + '-01';
  Logger.log('[diagBrk]   Вікно дат: [%s ... %s)', dateFrom, dateTo);

  // 3) ВСІ рядки по дитині (за підрядком імені) — будь-яка дата/локація
  var allForChild = [];
  for (var i = 1; i < data.length; i++){
    var rec = _parseAttendanceRow(data[i]);
    if (String(rec.child || '').replace(/\s+/g, ' ').toLowerCase().indexOf(nameNorm) === -1) continue;
    allForChild.push(rec);
  }
  Logger.log('[diagBrk] 3) Рядків по підрядку імені: %s', allForChild.length);
  allForChild.slice(0, 60).forEach(function(rec){
    var inLoc = String(rec.loc).trim() === loc;
    var inWin = rec.date >= dateFrom && rec.date < dateTo;
    Logger.log('   date=%s | child="%s"(len=%s) | loc="%s"%s | %s | ціна=%s%s',
      rec.date, rec.child, String(rec.child).length, rec.loc, inLoc ? '' : ' ←loc≠',
      rec.activityName, rec.price, inWin ? ' ✓у вікні' : ' ✗поза вікном');
  });
  if (!allForChild.length) Logger.log('[diagBrk]   ⚠ ЖОДНОГО рядка по цій дитині в усьому аркуші — перевір ім\'я/аркуш');

  // 4) Реплікація extrasByChild (точний match child===clientName, loc, вікно) — як у getInvoiceListData
  var matchName = clientName || name;
  var sum = 0, breakdown = {}, rowsInWindow = 0;
  for (var r = 1; r < data.length; r++){
    var rc = _parseAttendanceRow(data[r]);
    if (String(rc.loc).trim() !== loc) continue;
    if (rc.date < dateFrom || rc.date >= dateTo) continue;
    if (String(rc.child) !== matchName) continue;   // ТОЧНИЙ матч (extrasByChild[rec.child] → потім [name])
    rowsInWindow++;
    sum += (rc.price || 0);
    if (!breakdown[rc.activityName]) breakdown[rc.activityName] = {name: rc.activityName, count: 0, price: rc.price, total: 0};
    breakdown[rc.activityName].count++;
    breakdown[rc.activityName].total += (rc.price || 0);
  }
  Logger.log('[diagBrk] 4) extrasByChild["%s"] (точний матч як getInvoiceListData): рядків=%s sum=%s', matchName, rowsInWindow, sum);
  Object.keys(breakdown).forEach(function(k){
    var b = breakdown[k];
    Logger.log('   %s ×%s @%s = %s', b.name, b.count, b.price, b.total);
  });
  if (!rowsInWindow) Logger.log('[diagBrk]   ⚠ 0 рядків при точному матчі → САМЕ ТОМУ breakdown порожній у getInvoiceListData');

  // 5) Розподіл відмічань дитини по місяцях (будь-яка локація) — щоб бачити де реально дані
  var byMonth = {};
  allForChild.forEach(function(rec){ var ym = String(rec.date).slice(0, 7); byMonth[ym] = (byMonth[ym] || 0) + 1; });
  Logger.log('[diagBrk] 5) Розподіл по місяцях (YYYY-MM):');
  Object.keys(byMonth).sort().forEach(function(ym){ Logger.log('   %s = %s', ym, byMonth[ym]); });

  Logger.log('[diagBrk] ═══════════════════════════════════════');
  return {ok:true, clientName: clientName, allRows: allForChild.length, windowRows: rowsInWindow, sum: sum, byMonth: byMonth};
}

function testDiagExtrasBreakdown(){
  return diagExtrasBreakdownForChild('Матущенко Сара', 'Осокорки', 5, 2026);
}

// ───────────────────────────────────────────────────────────────────────────
// v6.11.24 ТЕСТОВІ ВІДМІЧАННЯ (Варіант Б): додає 5 рядків у Додаткові_Відвідуваність
// для Матущенко Сара (Осокорки, травень 2026), щоб побачити повний extras-PDF.
// Усі рядки мітяться у полі "Відмітив" = 'ТЕСТ_EXTRAS_PDF' → rollbackTestExtras
// видалить їх по мітці. Запобіжник: dryRun=true default, confirm:'YES_WRITE'.
// breakdown: Лего 280×2 + Айкідо 500×2 + Логопед 600 = 2160₴.
// ───────────────────────────────────────────────────────────────────────────
var TEST_EXTRAS_MARKER = 'ТЕСТ_EXTRAS_PDF';

function createTestExtrasAttendance(opts){
  opts = opts || {};
  var dryRun = (opts.dryRun !== false);
  if (!dryRun && opts.confirm !== 'YES_WRITE'){
    Logger.log('[testExtras] ⚠ REAL заблоковано: бракує confirm="YES_WRITE". Працюю як DRY-RUN.');
    dryRun = true;
  }
  var loc = 'Осокорки', group = 'Preschool Юля', child = 'Матущенко Сара';
  var rows = [
    {date:'2026-05-15', act:'Лего',    price:280},
    {date:'2026-05-17', act:'Лего',    price:280},
    {date:'2026-05-20', act:'Айкідо',  price:500},
    {date:'2026-05-22', act:'Айкідо',  price:500},
    {date:'2026-05-24', act:'Логопед', price:600}
  ];
  var bsum = rows.reduce(function(s, r){ return s + r.price; }, 0);

  Logger.log('[testExtras] ═══════════════════════════════════════');
  Logger.log('[testExtras] dryRun=%s | %s рядків для "%s" (%s) | breakdownSum=%s', dryRun, rows.length, child, loc, bsum);

  var sh = _getAttendanceSheet(true);
  var data = sh.getDataRange().getValues();
  var maxId = 0;
  for (var r = 1; r < data.length; r++){ var n = Number(data[r][0]) || 0; if (n > maxId) maxId = n; }

  var now = Utilities.formatDate(new Date(), 'Europe/Kiev', 'yyyy-MM-dd HH:mm');
  // Порядок колонок як ATTENDANCE_HEADER: id|Дата|Локація|Група|Дитина|id_заняття|Назва_заняття|Ціна|Відмітив|Час
  var toWrite = rows.map(function(rr){ return [++maxId, rr.date, loc, group, child, 0, rr.act, rr.price, TEST_EXTRAS_MARKER, now]; });
  toWrite.forEach(function(w){ Logger.log('   %s | %s | %s | %s | %s₴ | мітка=%s', w[1], w[2], w[4], w[6], w[7], w[8]); });

  if (dryRun){
    Logger.log('[testExtras] DRY-RUN — нічого не записано. %s рядків БУДЕ додано при confirm="YES_WRITE".', toWrite.length);
    Logger.log('[testExtras] ═══════════════════════════════════════');
    return {ok:true, dryRun:true, willAdd:toWrite.length, breakdownSum:bsum};
  }
  sh.getRange(sh.getLastRow() + 1, 1, toWrite.length, toWrite[0].length).setValues(toWrite);
  Logger.log('[testExtras] ✅ ДОДАНО %s тестових рядків (мітка "%s")', toWrite.length, TEST_EXTRAS_MARKER);
  Logger.log('[testExtras] ═══════════════════════════════════════');
  return {ok:true, dryRun:false, added:toWrite.length, breakdownSum:bsum};
}

function testCreateTestExtrasDryRun(){ return createTestExtrasAttendance({dryRun:true}); }
function runCreateTestExtras(){ return createTestExtrasAttendance({dryRun:false, confirm:'YES_WRITE'}); }

// Відкат тестових відмічань по мітці. dryRun=true default; confirm + count-safety (5).
function rollbackTestExtras(opts){
  opts = opts || {};
  var dryRun = (opts.dryRun !== false);
  if (!dryRun && opts.confirm !== 'YES_WRITE'){
    Logger.log('[rbExtras] ⚠ REAL заблоковано: бракує confirm="YES_WRITE". Працюю як DRY-RUN.');
    dryRun = true;
  }
  var expected = (opts.expectedCount != null) ? Number(opts.expectedCount) : 5;

  var sh = _getAttendanceSheet(false);
  var data = sh.getDataRange().getValues();
  var markIdx = 8;   // 'Відмітив'
  var hits = [];     // sheet row numbers (1-based)
  for (var r = 1; r < data.length; r++){
    if (String(data[r][markIdx]).trim() === TEST_EXTRAS_MARKER) hits.push(r + 1);
  }
  Logger.log('[rbExtras] ═══════════════════════════════════════');
  Logger.log('[rbExtras] dryRun=%s | знайдено %s рядків з міткою "%s" (очікувано %s)', dryRun, hits.length, TEST_EXTRAS_MARKER, expected);
  hits.forEach(function(rn){ var row = data[rn - 1]; Logger.log('   row %s | %s | %s | %s | %s₴', rn, row[1], row[4], row[6], row[7]); });

  if (dryRun){
    Logger.log('[rbExtras] DRY-RUN — нічого не видалено. %s рядків БУДЕ видалено при confirm="YES_WRITE".', hits.length);
    Logger.log('[rbExtras] ═══════════════════════════════════════');
    return {ok:true, dryRun:true, found:hits.length};
  }
  if (hits.length !== expected){
    Logger.log('[rbExtras] ⛔ ЗАБЛОКОВАНО: знайдено %s != expected %s. Видалення скасовано. Виклич {dryRun:false, confirm:"YES_WRITE", expectedCount:%s} якщо свідомо.', hits.length, expected, hits.length);
    Logger.log('[rbExtras] ═══════════════════════════════════════');
    return {ok:false, blocked:true, found:hits.length, expected:expected};
  }
  hits.sort(function(a, b){ return b - a; });   // знизу вгору, щоб індекси не зсувались
  hits.forEach(function(rn){ sh.deleteRow(rn); });
  Logger.log('[rbExtras] ✅ ВИДАЛЕНО %s тестових рядків', hits.length);
  Logger.log('[rbExtras] ═══════════════════════════════════════');
  return {ok:true, dryRun:false, deleted:hits.length};
}

function testRollbackTestExtrasDryRun(){ return rollbackTestExtras({dryRun:true}); }
function runRollbackTestExtras(){ return rollbackTestExtras({dryRun:false, confirm:'YES_WRITE'}); }

// ТЕСТ extras-рахунку: Матущенко Сара / Осокорки / травень 2026 → PDF у Drive, URL у лог.
function testGenerateInvoiceExtras(){
  var res = generateInvoicePDF({childName: 'Матущенко Сара', loc: 'Осокорки', type: 'extras', month: 5, year: 2026, invoiceDate: '01.06.2026'});
  if (!res.ok){ Logger.log('[testInvoiceExtras] ❌ %s', res.error); return res; }
  var bytes = Utilities.base64Decode(res.pdfBase64);
  var blob = Utilities.newBlob(bytes, 'application/pdf', res.pdfFilename);
  var file = DriveApp.createFile(blob);
  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(e){}
  Logger.log('[testInvoiceExtras] ✅ №%s | сума=%s (%s) | ЮО="%s" | замовник="%s"',
    res.invoiceNumber, res.sum, res.sumWords, res.juName, res.buyerName);
  Logger.log('[testInvoiceExtras] PDF URL:  %s', file.getUrl());
  Logger.log('[testInvoiceExtras] download: %s', 'https://drive.google.com/uc?export=download&id=' + file.getId());
  return res;
}

// ───────────────────────────────────────────────────────────────────────────
// v6.11.25 ДІАГНОСТИКА (КРОК B): скан усіх Клієнти на договір-поля, що збереглись
// як Date (Sheets auto-coerce). Read-only. Перевіряє "Номер договору" та
// "Номер додаткового договору" через instanceof Date. Запуск: testDiagContractDateBug().
// ───────────────────────────────────────────────────────────────────────────
function diagContractDateBug(){
  var sh = getCRMSpreadsheet().getSheetByName(SHEET_CLIENTS);
  if (!sh){ Logger.log('[diagContractDate] ❌ "%s" не знайдено', SHEET_CLIENTS); return {ok:false}; }
  var vals = sh.getDataRange().getValues();
  var hdr = vals[0].map(function(x){ return String(x); });
  var nameIdx = hdr.indexOf('ПІБ дитини');
  var locIdx  = hdr.indexOf('Локація');
  var c1 = hdr.indexOf('Номер договору');
  var c2 = hdr.indexOf('Номер додаткового договору');

  Logger.log('[diagContractDate] ═══════════════════════════════════════');
  Logger.log('[diagContractDate] Клієнти: %s рядків даних | "Номер договору"@%s | "Номер дод.договору"@%s',
    vals.length - 1, c1, c2);
  Logger.log('[diagContractDate] ─────────────────────────────────────');

  var only1 = 0, only2 = 0, both = 0;
  for (var r = 1; r < vals.length; r++){
    var name = String(vals[r][nameIdx] || '').trim();
    if (!name) continue;
    var loc = String(vals[r][locIdx] || '').trim();
    var v1 = c1 >= 0 ? vals[r][c1] : '';
    var v2 = c2 >= 0 ? vals[r][c2] : '';
    var d1 = v1 instanceof Date;
    var d2 = v2 instanceof Date;
    if (d1){ only1++; Logger.log('[diagContractDate] ⚠ %s | %s | звичайний договір | %s', name, loc, v1); }
    if (d2){ only2++; Logger.log('[diagContractDate] ⚠ %s | %s | додатковий договір | %s', name, loc, v2); }
    if (d1 && d2) both++;
  }

  var affected = only1 + only2 - both;
  Logger.log('[diagContractDate] ─────────────────────────────────────');
  Logger.log('[diagContractDate] ПІДСУМОК: карток з Date у договорі: %s (звичайний: %s, додатковий: %s, обидва поля: %s)',
    affected, only1, only2, both);
  Logger.log('[diagContractDate] ═══════════════════════════════════════');
  return {ok:true, affectedCards: affected, contract: only1, additional: only2, both: both};
}

function testDiagContractDateBug(){ return diagContractDateBug(); }

// ───────────────────────────────────────────────────────────────────────────
// v6.26.1 ДІАГНОСТИКА: чому деякі діти Бровари відсутні в invoices.html
// попри наявність "Бюджет-навч" у Оплати-Рік. Гіпотези:
//   A) Назва колонки не співпадає (наприклад "Червень-Бюджет-навч" відсутня).
//   B) Match по ПІБ ламається — зайвий пробіл, регістр, emoji статусу 🏖️.
//
// Запускати ВРУЧНУ з Apps Script editor через testDiagInvoiceListDataMatch().
// Логи у View → Executions → клацай на запуск → Logs.
// ───────────────────────────────────────────────────────────────────────────
function diagInvoiceListDataMatch(loc, payMonth, payYear){
  loc = String(loc || 'Бровари').trim();
  payMonth = Number(payMonth || 6);
  payYear = Number(payYear || 2026);

  Logger.log('[diagInvMatch] ═══════════════════════════════════════');
  Logger.log('[diagInvMatch] START loc="%s" payMonth=%s payYear=%s', loc, payMonth, payYear);
  Logger.log('[diagInvMatch] ═══════════════════════════════════════');

  // === 1. Клієнти ===
  var crmRes = getClients();
  if (!crmRes.ok){
    Logger.log('[diagInvMatch] ❌ getClients failed: %s', crmRes.error);
    return crmRes;
  }
  var clientsAll = crmRes.data || [];

  // Завжди показуємо ВСІ унікальні локації з Клієнти — щоб ловити розбіжності
  // у назві (наприклад 'Осокорки' vs 'Осокорки сад' vs 'Осокорки садок').
  var clientLocations = {};
  clientsAll.forEach(function(c){
    var l = String(c['Локація'] || '').trim();
    if (l) clientLocations[l] = (clientLocations[l] || 0) + 1;
  });
  Logger.log('[diagInvMatch] УСІ локації в Клієнти (унікальні): %s', Object.keys(clientLocations).length);
  Object.keys(clientLocations).sort().forEach(function(l){
    var marker = l === loc ? ' ← ШУКАЄМО' : '';
    Logger.log('  · "%s" (%s клієнтів)%s', l, clientLocations[l], marker);
  });
  Logger.log('[diagInvMatch] ─────────────────────────────────────');

  var clients = clientsAll.filter(function(c){
    if (String(c['Локація'] || '').trim() !== loc) return false;
    var st = String(c['Статус'] || '').trim();
    return st === 'active' || st === 'adaptation';
  });

  Logger.log('[diagInvMatch] КЛІЄНТИ %s (active+adaptation): %s', loc, clients.length);
  if (clients.length === 0){
    Logger.log('[diagInvMatch] ⚠ ЖОДНОГО клієнта з loc="%s" статусу active/adaptation', loc);
    Logger.log('[diagInvMatch] Перевір список вище — можливо локація має іншу назву у Клієнти');
  }
  clients.forEach(function(c, i){
    var name = String(c['ПІБ дитини'] || '');
    Logger.log('  %s. "%s" | len=%s | група="%s" | статус=%s',
      i+1, name, name.length, c['Група'], c['Статус']);
  });
  Logger.log('[diagInvMatch] ─────────────────────────────────────');

  // === 2. Оплати-Рік ===
  var paySheet = getCRMSpreadsheet().getSheetByName(SHEET_YEARLY);
  if (!paySheet){
    Logger.log('[diagInvMatch] ❌ Sheet "%s" не знайдено!', SHEET_YEARLY);
    return {ok:false, error:'Sheet not found'};
  }
  var pvals = paySheet.getDataRange().getValues();
  if (pvals.length < 2){
    Logger.log('[diagInvMatch] ❌ Оплати-Рік порожній (lastRow=%s)', pvals.length);
    return {ok:false, error:'Empty'};
  }

  var phdrs = pvals[0].map(function(h){ return String(h); });
  var budNavchCol = MONTHS_CAL[payMonth - 1] + '-Бюджет-навч';

  Logger.log('[diagInvMatch] ОПЛАТИ-РІК: %s колонок, %s рядків', phdrs.length, pvals.length - 1);
  Logger.log('[diagInvMatch] Headers (перші 12): %s', JSON.stringify(phdrs.slice(0, 12)));
  Logger.log('[diagInvMatch] Шукаємо колонку: "%s"', budNavchCol);

  var nameIdx = phdrs.indexOf("Ім'я дитини");
  var locIdx  = phdrs.indexOf('Локація');
  var budIdx  = phdrs.indexOf(budNavchCol);
  Logger.log('[diagInvMatch] indices: name=%s loc=%s bud=%s', nameIdx, locIdx, budIdx);

  if (budIdx < 0){
    var candidates = phdrs.filter(function(h){ return String(h).indexOf('Бюджет-навч') !== -1; });
    Logger.log('[diagInvMatch] ❌ Колонку "%s" НЕ знайдено!', budNavchCol);
    Logger.log('[diagInvMatch] Усі колонки з "Бюджет-навч": %s', JSON.stringify(candidates));
    return {ok:false, error:'Column "' + budNavchCol + '" not found'};
  }
  if (nameIdx < 0 || locIdx < 0){
    Logger.log('[diagInvMatch] ❌ name/loc індекси не знайдено: name=%s loc=%s', nameIdx, locIdx);
    return {ok:false, error:'name/loc cols missing'};
  }

  // Завжди показуємо ВСІ унікальні локації з Оплати-Рік (для порівняння з Клієнти)
  var paymentLocations = {};
  for (var prL = 1; prL < pvals.length; prL++){
    var lP = String(pvals[prL][locIdx] || '').trim();
    if (lP) paymentLocations[lP] = (paymentLocations[lP] || 0) + 1;
  }
  Logger.log('[diagInvMatch] УСІ локації в Оплати-Рік (унікальні): %s', Object.keys(paymentLocations).length);
  Object.keys(paymentLocations).sort().forEach(function(l){
    var marker = l === loc ? ' ← ШУКАЄМО' : '';
    Logger.log('  · "%s" (%s рядків)%s', l, paymentLocations[l], marker);
  });
  Logger.log('[diagInvMatch] ─────────────────────────────────────');

  // Рядки Оплати-Рік для loc
  var paymentRows = [];
  for (var pr = 1; pr < pvals.length; pr++){
    if (String(pvals[pr][locIdx]).trim() !== loc) continue;
    var rawName = String(pvals[pr][nameIdx] || '');
    paymentRows.push({
      name: rawName,
      nameTrim: rawName.trim(),
      bud: Number(pvals[pr][budIdx]) || 0,
      rowNum: pr + 1
    });
  }
  Logger.log('[diagInvMatch] ОПЛАТИ-РІК (%s) — %s рядків:', loc, paymentRows.length);
  paymentRows.forEach(function(p, i){
    // hex dump перших 60 chars (для виявлення невидимих символів)
    var hexHint = '';
    for (var ch = 0; ch < p.name.length && ch < 30; ch++){
      var code = p.name.charCodeAt(ch);
      if (code < 32 || (code >= 127 && code < 160) || code === 0xFEFF || code === 0x200B){
        hexHint += ' [pos' + ch + '=U+' + code.toString(16).toUpperCase() + ']';
      }
    }
    Logger.log('  %s. row%s | "%s" | len=%s | бюджет-навч=%s%s',
      i+1, p.rowNum, p.name, p.nameTrim.length, p.bud, hexHint);
  });
  Logger.log('[diagInvMatch] ─────────────────────────────────────');

  // === 3. Map для матчингу: nameTrim → bud (як у getInvoiceListData) ===
  var paymentByName = {};
  paymentRows.forEach(function(p){ paymentByName[p.nameTrim] = p.bud; });

  // === 4. Match table ===
  Logger.log('[diagInvMatch] MATCH TABLE (✓ = знайдено, ✗ = не знайдено):');
  var matched = 0, missing = 0, zero = 0;
  var missingNames = [];
  clients.forEach(function(c){
    var cname = String(c['ПІБ дитини'] || '').trim();
    var bud = paymentByName[cname];
    if (bud !== undefined){
      matched++;
      if (bud > 0) Logger.log('  ✓ %s → %s ₴', cname, bud);
      else { zero++; Logger.log('  ⚠ %s → 0 ₴ (буде відкинутий фільтром paymentSum>0)', cname); }
    } else {
      missing++;
      missingNames.push(cname);
      // Шукаємо близькі матчі — за substring (case-insensitive)
      var lc = cname.toLowerCase();
      var close = paymentRows.filter(function(p){
        var pl = p.nameTrim.toLowerCase();
        return pl.indexOf(lc) !== -1 || (lc.length > 4 && lc.indexOf(pl) !== -1);
      });
      var closeStr = close.length
        ? ' | close: ' + close.map(function(p){ return '"' + p.name + '"=' + p.bud + ' (row' + p.rowNum + ')'; }).join('; ')
        : ' | (close-match не знайдено)';
      Logger.log('  ✗ %s — НЕ МАТЧИТЬСЯ%s', cname, closeStr);
    }
  });

  Logger.log('[diagInvMatch] ─────────────────────────────────────');
  Logger.log('[diagInvMatch] РЕЗУЛЬТАТ matched=%s, missing=%s, zero-bud=%s (з %s клієнтів)',
    matched, missing, zero, clients.length);

  // === 5. Payment-only (є в Оплати-Рік, нема в Клієнти active+adaptation) ===
  var clientsByName = {};
  clients.forEach(function(c){ clientsByName[String(c['ПІБ дитини']||'').trim()] = true; });
  var paymentOnly = paymentRows.filter(function(p){ return !clientsByName[p.nameTrim]; });
  Logger.log('[diagInvMatch] Payment-only (%s) — є в Оплати-Рік, нема в Клієнти active+adapt:', paymentOnly.length);
  paymentOnly.forEach(function(p){
    Logger.log('  · "%s" | row%s | бюджет=%s', p.name, p.rowNum, p.bud);
  });

  Logger.log('[diagInvMatch] ═══════════════════════════════════════');
  Logger.log('[diagInvMatch] DONE. Якщо missing>0 — дивись секцію MATCH TABLE: close-match покаже як');
  Logger.log('[diagInvMatch] Клієнти запис відрізняється від Оплати-Рік (емодзі, пробіл, регістр).');

  return {
    ok: true,
    matched: matched, missing: missing, zero: zero,
    totalClients: clients.length, totalPayment: paymentRows.length,
    missingNames: missingNames,
    paymentOnlyCount: paymentOnly.length
  };
}

function testDiagInvoiceListDataMatch(){
  return diagInvoiceListDataMatch('Осокорки', 6, 2026);
}

// ───────────────────────────────────────────────────────────────────────────
// v6.26.1 ТОЧКОВА ДІАГНОСТИКА: чому конкретні діти (Гайдай Коля та ін.) мають
// картку в Клієнти, але getClients/getInvoiceListData їх не бачить.
// Читає аркуш Клієнти НАПРЯМУ (getDataRange), без фільтрів, і для шуканих імен
// показує hex-dump Локації/Статусу/Групи + чи є ID (getClients пропускає рядки
// з порожнім ID через `if (!vals[r][0]) continue`).
// Запускати ВРУЧНУ: testDiagFindSpecificClients(). Read-only, без deploy.
// ───────────────────────────────────────────────────────────────────────────
function diagFindSpecificClients(){
  var targets = ['Гайдай', 'Нікітіна Ніка', 'Лінник Ксенія', 'Мордачова'];

  var ss = getCRMSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_CLIENTS);
  if (!sheet){ Logger.log('[diagFind] ❌ Sheet "%s" не знайдено', SHEET_CLIENTS); return {ok:false}; }
  var vals = sheet.getDataRange().getValues();
  var headers = vals[0].map(function(h){ return String(h); });

  var idIdx     = headers.indexOf('ID');
  var nameIdx   = headers.indexOf('ПІБ дитини');
  var locIdx    = headers.indexOf('Локація');
  var statusIdx = headers.indexOf('Статус');
  var groupIdx  = headers.indexOf('Група');

  Logger.log('[diagFind] ═══════════════════════════════════════');
  Logger.log('[diagFind] Аркуш "%s": рядків з шапкою=%s, даних=%s', SHEET_CLIENTS, vals.length, vals.length - 1);
  Logger.log('[diagFind] Колонки: ID@%s name@%s loc@%s status@%s group@%s', idIdx, nameIdx, locIdx, statusIdx, groupIdx);

  // Скільки getClients реально читає (фільтр !vals[r][0] = порожній ID)
  var getClientsCount = 0, noIdCount = 0;
  for (var r0 = 1; r0 < vals.length; r0++){
    if (!vals[r0][0]){ noIdCount++; continue; }
    getClientsCount++;
  }
  Logger.log('[diagFind] getClients прочитає %s рядків; рядків з ПОРОЖНІМ ID (пропущених) = %s', getClientsCount, noIdCount);
  Logger.log('[diagFind] ─────────────────────────────────────');

  function hexDump(s){
    s = String(s == null ? '' : s);
    var out = '';
    for (var i = 0; i < s.length; i++){
      var code = s.charCodeAt(i);
      // control / non-breaking space (A0) / zero-width (200B) / BOM (FEFF)
      if (code < 32 || (code >= 127 && code < 160) || code === 0xFEFF || code === 0x200B || code === 0xA0){
        out += ' [pos' + i + '=U+' + ('0000' + code.toString(16).toUpperCase()).slice(-4) + ']';
      }
    }
    return out || ' (чисто)';
  }

  targets.forEach(function(t){
    var found = [];
    for (var r = 1; r < vals.length; r++){
      var nm = String(vals[r][nameIdx] || '');
      if (nm.indexOf(t) !== -1) found.push(r);
    }
    if (!found.length){
      Logger.log('[diagFind] ✗ NOT FOUND: "%s"', t);
      return;
    }
    found.forEach(function(r){
      var row = vals[r];
      var idVal     = idIdx     >= 0 ? row[idIdx]            : '(no col)';
      var nameVal   = String(row[nameIdx]   || '');
      var locVal    = String(row[locIdx]    || '');
      var statusVal = String(row[statusIdx] || '');
      var groupVal  = String(row[groupIdx]  || '');
      Logger.log('[diagFind] ✓ "%s" → sheet row %s:', t, r + 1);
      Logger.log('    ID      = "%s"%s', idVal, (!idVal ? '  ← ⚠ ПОРОЖНІЙ! getClients ПРОПУСКАЄ цей рядок' : ''));
      Logger.log('    name    = "%s" len=%s |%s', nameVal,   nameVal.length,   hexDump(nameVal));
      Logger.log('    Локація = "%s" len=%s |%s', locVal,    locVal.length,    hexDump(locVal));
      Logger.log('    Статус  = "%s" len=%s |%s', statusVal, statusVal.length, hexDump(statusVal));
      Logger.log('    Група   = "%s" len=%s |%s', groupVal,  groupVal.length,  hexDump(groupVal));
    });
  });

  Logger.log('[diagFind] ═══════════════════════════════════════');
  Logger.log('[diagFind] DONE. Якщо у дитини ID порожній → це причина (getClients skip).');
  Logger.log('[diagFind] Якщо ID є, але Локація/Статус мають [pos..=U+00A0/U+200B] → прихований символ.');

  return {ok:true, totalRows: vals.length - 1, getClientsReads: getClientsCount, emptyIdRows: noIdCount};
}

function testDiagFindSpecificClients(){
  return diagFindSpecificClients();
}

// ───────────────────────────────────────────────────────────────────────────
// v6.11.13 ТОЧКОВА ДІАГНОСТИКА ОДНІЄЇ ДИТИНИ: чому конкретна дитина не зʼявляється
// в invoices.html. Проходить весь ланцюг getInvoiceListData:
//   1) є в Клієнти? ID не порожній (інакше getClients SKIP)? статус active/adapt?
//      група точно очікувана? (+ hex-dump прихованих символів)
//   2) getClients() реально віддає цю дитину для loc?
//   3) Оплати-Рік: рядок(и) дитини + значення "<Місяць>-Бюджет-навч" + УСІ
//      Бюджет-колонки рядка (щоб бачити, чи бюджет стоїть в іншому місяці).
//   4) getInvoiceListData(loc, місяць) бачить її в children[]? з якими сумами?
//   5) Вердикт: яка саме умова відсікає.
// КЛЮЧОВЕ: payment-матч у getInvoiceListData — ПО ТОЧНОМУ ПІБ (trim), НЕ по групі.
// Тобто перейменування групи НЕ впливає на появу в рахунках.
// Read-only, без запису. Запуск: testDiagSapogov().
// ───────────────────────────────────────────────────────────────────────────
function diagSpecificChild(name, loc, payMonth, payYear){
  name     = String(name || '').trim();
  loc      = String(loc  || '').trim();
  payMonth = Number(payMonth || 6);
  payYear  = Number(payYear  || 2026);
  var nameNorm = name.replace(/\s+/g,' ').toLowerCase();

  function hexDump(s){
    s = String(s == null ? '' : s);
    var out = '';
    for (var i = 0; i < s.length; i++){
      var code = s.charCodeAt(i);
      if (code < 32 || (code >= 127 && code < 160) || code === 0xFEFF || code === 0x200B || code === 0xA0){
        out += ' [pos' + i + '=U+' + ('0000' + code.toString(16).toUpperCase()).slice(-4) + ']';
      }
    }
    return out || ' (чисто)';
  }

  Logger.log('[diagChild] ═══════════════════════════════════════');
  Logger.log('[diagChild] ШУКАЮ "%s" | loc="%s" | %s/%s', name, loc, payMonth, payYear);
  Logger.log('[diagChild] ═══════════════════════════════════════');

  // === 1. Клієнти НАПРЯМУ (getDataRange, без фільтрів) ===
  var cSheet = getCRMSpreadsheet().getSheetByName(SHEET_CLIENTS);
  if (!cSheet){ Logger.log('[diagChild] ❌ "%s" не знайдено', SHEET_CLIENTS); return {ok:false}; }
  var cvals = cSheet.getDataRange().getValues();
  var chdrs = cvals[0].map(function(h){ return String(h); });
  var cIdIdx     = chdrs.indexOf('ID');
  var cNameIdx   = chdrs.indexOf('ПІБ дитини');
  var cLocIdx    = chdrs.indexOf('Локація');
  var cStatusIdx = chdrs.indexOf('Статус');
  var cGroupIdx  = chdrs.indexOf('Група');

  var hits = [];
  for (var r = 1; r < cvals.length; r++){
    var nm = String(cvals[r][cNameIdx] || '');
    if (nm.replace(/\s+/g,' ').toLowerCase().indexOf(nameNorm) !== -1) hits.push(r);
  }
  Logger.log('[diagChild] 1) КЛІЄНТИ: %s рядків з підрядком імені', hits.length);
  var clientActiveOk = false, clientGroup = '';
  hits.forEach(function(r){
    var row = cvals[r];
    var idVal     = String(row[cIdIdx]     || '');
    var nameVal   = String(row[cNameIdx]   || '');
    var locVal    = String(row[cLocIdx]    || '');
    var statusVal = String(row[cStatusIdx] || '');
    var groupVal  = String(row[cGroupIdx]  || '');
    clientGroup = groupVal;
    var stTrim = statusVal.trim();
    if ((stTrim === 'active' || stTrim === 'adaptation') && idVal) clientActiveOk = true;
    Logger.log('[diagChild]   sheet row %s:', r + 1);
    Logger.log('[diagChild]     ID      = "%s"%s', idVal, (!idVal ? '  ← ⚠ ПОРОЖНІЙ! getClients ПРОПУСКАЄ цей рядок' : ''));
    Logger.log('[diagChild]     name    = "%s" len=%s |%s', nameVal,   nameVal.length,   hexDump(nameVal));
    Logger.log('[diagChild]     Локація = "%s" len=%s |%s', locVal,    locVal.length,    hexDump(locVal));
    Logger.log('[diagChild]     Статус  = "%s"%s', statusVal, (stTrim==='active'||stTrim==='adaptation' ? ' ✓' : ' ← ⚠ НЕ active/adaptation'));
    Logger.log('[diagChild]     Група   = "%s" len=%s |%s', groupVal,  groupVal.length,  hexDump(groupVal));
  });

  // === 2. getClients() — чи реально віддає цю дитину для loc ===
  var gc = getClients();
  var inGetClients = false, gcName = '';
  if (gc.ok){
    (gc.data || []).forEach(function(c){
      var n = String(c['ПІБ дитини'] || '');
      if (n.replace(/\s+/g,' ').toLowerCase().indexOf(nameNorm) !== -1 &&
          String(c['Локація']||'').trim() === loc){
        inGetClients = true; gcName = n;
      }
    });
  }
  Logger.log('[diagChild] 2) getClients() віддає дитину для loc="%s": %s', loc, inGetClients ? 'ТАК ("'+gcName+'")' : 'НІ');

  // === 3. Оплати-Рік — рядок(и) дитини + усі "Бюджет"-колонки ===
  var paySheet = getCRMSpreadsheet().getSheetByName(SHEET_YEARLY);
  var budNavchCol = MONTHS_CAL[payMonth - 1] + '-Бюджет-навч';
  var paymentSumByName = 0, foundInPay = false;
  if (paySheet){
    var pvals = paySheet.getDataRange().getValues();
    var phdrs = pvals[0].map(function(h){ return String(h); });
    var pNameIdx = phdrs.indexOf("Ім'я дитини");
    var pLocIdx  = phdrs.indexOf('Локація');
    var pBudIdx  = phdrs.indexOf(budNavchCol);
    var budColIdxs = [];
    phdrs.forEach(function(h, idx){ if (String(h).indexOf('Бюджет') !== -1) budColIdxs.push(idx); });
    Logger.log('[diagChild] 3) ОПЛАТИ-РІК: колонка "%s" → idx=%s %s', budNavchCol, pBudIdx, (pBudIdx<0?'← ⚠ КОЛОНКИ НЕМАЄ':''));
    for (var pr = 1; pr < pvals.length; pr++){
      var pn = String(pvals[pr][pNameIdx] || '');
      if (pn.replace(/\s+/g,' ').toLowerCase().indexOf(nameNorm) === -1) continue;
      foundInPay = true;
      var pl = String(pvals[pr][pLocIdx] || '');
      var budVal = pBudIdx >= 0 ? (Number(pvals[pr][pBudIdx]) || 0) : 0;
      Logger.log('[diagChild]   row %s | name="%s"%s | loc="%s"%s | %s=%s',
        pr + 1, pn, hexDump(pn), pl, (pl.trim()===loc?'':' ← ⚠ loc≠'+loc), budNavchCol, budVal);
      var budDump = budColIdxs.map(function(idx){ return phdrs[idx] + '=' + (Number(pvals[pr][idx])||0); }).join(' | ');
      Logger.log('[diagChild]     ВСІ Бюджет-колонки рядка: %s', budDump);
    }
    // paymentSum РІВНО як рахує getInvoiceListData: точний trim-матч ПІБ + фільтр loc
    if (pNameIdx >= 0 && pLocIdx >= 0 && pBudIdx >= 0){
      for (var pr2 = 1; pr2 < pvals.length; pr2++){
        if (String(pvals[pr2][pLocIdx]).trim() !== loc) continue;
        if (String(pvals[pr2][pNameIdx]).trim() !== name) continue;
        paymentSumByName = Number(pvals[pr2][pBudIdx]) || 0;
      }
    }
  } else {
    Logger.log('[diagChild] 3) ❌ "%s" не знайдено', SHEET_YEARLY);
  }
  Logger.log('[diagChild]   → paymentSum (точний trim-матч ПІБ+loc, як getInvoiceListData) = %s', paymentSumByName);
  if (!foundInPay) Logger.log('[diagChild]   ⚠ Дитину НЕ знайдено в Оплати-Рік взагалі (за підрядком імені)');

  // === 4. getInvoiceListData — чи реально в списку ===
  var inv = getInvoiceListData({loc: loc, payMonth: payMonth, payYear: payYear, extMonth: payMonth, extYear: payYear});
  var inInvoice = false, invChild = null;
  if (inv.ok){
    (inv.children || []).forEach(function(ch){
      if (String(ch.name||'').replace(/\s+/g,' ').toLowerCase().indexOf(nameNorm) !== -1){
        inInvoice = true; invChild = ch;
      }
    });
  }
  Logger.log('[diagChild] 4) getInvoiceListData(loc=%s pay=%s/%s ext=%s/%s): ok=%s shown=%s',
    loc, payMonth, payYear, payMonth, payYear, inv.ok, inv.shownChildren);
  Logger.log('[diagChild]   → дитина у children[]: %s', inInvoice
    ? ('ТАК ' + JSON.stringify({paymentSum: invChild.paymentSum, extrasSum: invChild.extrasSum, group: invChild.group}))
    : 'НІ');

  // === 5. ВЕРДИКТ ===
  Logger.log('[diagChild] ─────────────────────────────────────');
  Logger.log('[diagChild] ВЕРДИКТ:');
  Logger.log('[diagChild]   • у Клієнти active/adaptation з непорожнім ID: %s', clientActiveOk ? 'ТАК' : 'НІ ← ПРИЧИНА');
  Logger.log('[diagChild]   • getClients() віддає для loc: %s', inGetClients ? 'ТАК' : 'НІ ← ПРИЧИНА');
  Logger.log('[diagChild]   • paymentSum (%s) > 0: %s (=%s)', budNavchCol, paymentSumByName > 0 ? 'ТАК' : 'НІ', paymentSumByName);
  Logger.log('[diagChild]   • показується в getInvoiceListData: %s', inInvoice ? 'ТАК' : 'НІ');
  if (!inInvoice && clientActiveOk && inGetClients && paymentSumByName <= 0){
    Logger.log('[diagChild]   ⇒ ПРИЧИНА: paymentSum=0 (колонка "%s" порожня/0) і extrasSum=0 → фільтр "paymentSum<=0 && extrasSum<=0" відкидає. Група НЕ впливає (матч по ПІБ). САМЕ ЦЕ виправить нова логіка фільтра.', budNavchCol);
  }
  Logger.log('[diagChild] ═══════════════════════════════════════');

  return {
    ok: true,
    clientActiveOk: clientActiveOk, clientGroup: clientGroup,
    inGetClients: inGetClients,
    foundInPay: foundInPay, paymentSum: paymentSumByName,
    inInvoice: inInvoice, invoiceShown: inv.shownChildren
  };
}

function testDiagSapogov(){
  return diagSpecificChild('Сапогов Рінат', 'Осокорки', 6, 2026);
}

// ───────────────────────────────────────────────────────────────────────────
// v6.11.14 ДІАГНОСТИКА ПОРОЖНІХ СТАТУСІВ: знаходить картки в Клієнти, у яких
// поле "Статус" порожнє — їх відсікає фільтр active/adaptation у
// getInvoiceListData (як сталося з Сапоговим Рінатом). Read-only.
//   1) рахує розподіл статусів (active / adaptation / інші / порожні);
//   2) для порожніх — групує по локаціях, показує name|loc|група + чи є бюджет
//      у Оплати-Рік (по ключу name+loc, нормалізовано);
//   3) позначає походження картки за Нотатками: "синхр." (Чернетка —
//      автосинхронізація), "груп.оновл." (Група оновлена з Платежів), "стара".
// Запуск: testDiagClientsWithoutStatus().
// ───────────────────────────────────────────────────────────────────────────
function diagClientsWithoutStatus(){
  function normKey(name, loc){
    var n = String(name||'').trim().replace(/\s+/g,' ').toLowerCase();
    var l = String(loc ||'').trim().replace(/\s+/g,' ').toLowerCase();
    return n + '|' + l;
  }

  Logger.log('[diagNoStatus] ═══════════════════════════════════════');

  // === Клієнти НАПРЯМУ (getDataRange — бачимо й рядки з порожнім ID) ===
  var cSheet = getCRMSpreadsheet().getSheetByName(SHEET_CLIENTS);
  if (!cSheet){ Logger.log('[diagNoStatus] ❌ "%s" не знайдено', SHEET_CLIENTS); return {ok:false}; }
  var cvals = cSheet.getDataRange().getValues();
  var chdrs = cvals[0].map(function(h){ return String(h); });
  var cIdIdx     = chdrs.indexOf('ID');
  var cNameIdx   = chdrs.indexOf('ПІБ дитини');
  var cLocIdx    = chdrs.indexOf('Локація');
  var cGroupIdx  = chdrs.indexOf('Група');
  var cStatusIdx = chdrs.indexOf('Статус');
  var cNotesIdx  = chdrs.indexOf('Нотатки');

  // === Бюджет-мапа з Оплати-Рік (name+loc → є бюджет?) ===
  var hasBudget = {};
  var paySheet = getCRMSpreadsheet().getSheetByName(SHEET_YEARLY);
  if (paySheet){
    var pvals = paySheet.getDataRange().getValues();
    if (pvals.length >= 2){
      var phdrs = pvals[0].map(function(h){ return String(h); });
      var pNameIdx = phdrs.indexOf("Ім'я дитини");
      var pLocIdx  = phdrs.indexOf('Локація');
      var budColIdxs = [];
      phdrs.forEach(function(h, idx){ if (String(h).indexOf('Бюджет') !== -1) budColIdxs.push(idx); });
      if (pNameIdx >= 0 && pLocIdx >= 0){
        for (var pr = 1; pr < pvals.length; pr++){
          var pn = String(pvals[pr][pNameIdx] || '').trim();
          if (!pn) continue;
          var any = false;
          for (var bi = 0; bi < budColIdxs.length; bi++){
            if ((Number(pvals[pr][budColIdxs[bi]]) || 0) > 0){ any = true; break; }
          }
          if (any) hasBudget[normKey(pn, pvals[pr][pLocIdx])] = true;
        }
      }
    }
  }

  // === Прохід по Клієнти ===
  var cnt = {active:0, adaptation:0, otherNonEmpty:0, empty:0};
  var otherStatuses = {};       // статус → кількість
  var emptyByLoc = {};          // loc → [ {name, group, hasBud, origin} ]
  for (var r = 1; r < cvals.length; r++){
    var row = cvals[r];
    var name = String(row[cNameIdx] || '').trim();
    if (!name) continue;        // зовсім порожній рядок — пропускаємо
    var st = String(row[cStatusIdx] || '').trim();
    if (st === 'active'){ cnt.active++; continue; }
    if (st === 'adaptation'){ cnt.adaptation++; continue; }
    if (st){ cnt.otherNonEmpty++; otherStatuses[st] = (otherStatuses[st]||0) + 1; continue; }

    // ПОРОЖНІЙ статус
    cnt.empty++;
    var loc   = String(row[cLocIdx]   || '').trim();
    var group = String(row[cGroupIdx] || '').trim();
    var notes = String(row[cNotesIdx] || '');
    var origin = 'стара';
    if (notes.indexOf('Чернетка — автосинхронізація') !== -1) origin = 'синхр.';
    else if (notes.indexOf('Група оновлена з Платежів') !== -1) origin = 'груп.оновл.';
    var bud = hasBudget[normKey(name, loc)] ? 'так' : 'ні';
    (emptyByLoc[loc] = emptyByLoc[loc] || []).push({
      name: name, group: group, hasBud: bud, origin: origin, idEmpty: !row[cIdIdx]
    });
  }

  // === Звіт ===
  Logger.log('[diagNoStatus] РОЗПОДІЛ СТАТУСІВ (з %s рядків даних):', cvals.length - 1);
  Logger.log('[diagNoStatus]   active      = %s', cnt.active);
  Logger.log('[diagNoStatus]   adaptation  = %s', cnt.adaptation);
  Logger.log('[diagNoStatus]   інші непорожні = %s', cnt.otherNonEmpty);
  Object.keys(otherStatuses).sort().forEach(function(s){
    Logger.log('[diagNoStatus]       · "%s" = %s', s, otherStatuses[s]);
  });
  Logger.log('[diagNoStatus]   ⚠ ПОРОЖНІ    = %s', cnt.empty);
  Logger.log('[diagNoStatus] ─────────────────────────────────────');

  var emptyLocs = Object.keys(emptyByLoc).sort();
  Logger.log('[diagNoStatus] ПОРОЖНІЙ СТАТУС по локаціях (%s локацій):', emptyLocs.length);
  emptyLocs.forEach(function(loc){
    var list = emptyByLoc[loc];
    Logger.log('[diagNoStatus] ▼ %s (%s)', loc || '(порожня локація)', list.length);
    list.forEach(function(c){
      Logger.log('   %s | %s | бюджет=%s | %s%s',
        c.name, c.group || '(без групи)', c.hasBud, c.origin, c.idEmpty ? ' | ⚠ ID порожній' : '');
    });
  });
  Logger.log('[diagNoStatus] ─────────────────────────────────────');
  Logger.log('[diagNoStatus] ПІДСУМОК: %s карток без статусу в %s локаціях', cnt.empty, emptyLocs.length);
  Logger.log('[diagNoStatus] ═══════════════════════════════════════');

  return {
    ok: true,
    counts: cnt,
    otherStatuses: otherStatuses,
    emptyCount: cnt.empty,
    emptyLocations: emptyLocs.length,
    byLoc: emptyLocs.map(function(l){ return {loc:l, n:emptyByLoc[l].length}; })
  };
}

function testDiagClientsWithoutStatus(){
  return diagClientsWithoutStatus();
}

// ───────────────────────────────────────────────────────────────────────────
// v6.26.1 СИНХРОНІЗАТОР: створює чернетки карток у Клієнти для дітей, що є в
// Оплати-Рік, але відсутні в Клієнти (по ключу name+локація, нормалізовано).
//
// ДВА РЕЖИМИ:
//   opts.dryRun = true (default) → НІЧОГО не пише, лише логує що БУДЕ створено
//   opts.dryRun = false          → реально append рядки в Клієнти (НЕ перезаписує)
//
// Ключ порівняння нормалізований: trim + collapse spaces + lowercase, щоб
// "Гайдай Коля " == "Гайдай Коля". Оригінальне написання зберігається для запису.
// ID генерується за тим самим патерном що clients.html childId():
//   'c_' + name.trim().slice(0,20) + '_' + group.slice(0,8) + '_' + loc.slice(0,8)
// ───────────────────────────────────────────────────────────────────────────
function syncMissingClientsFromPayments(opts){
  opts = opts || {};
  // default dryRun=true: лише opts.dryRun === false вмикає реальний режим.
  var dryRun = (opts.dryRun !== false);
  // v6.26.1 ЗАПОБІЖНИК (після інциденту 832): реальний запис вимагає ще й
  // opts.confirm === 'YES_WRITE'. Випадковий виклик з {dryRun:false} без
  // confirm → залишається dry-run, нічого не пише.
  if (!dryRun && opts.confirm !== 'YES_WRITE'){
    Logger.log('[syncMissing] ⚠ REAL-режим заблоковано: бракує opts.confirm="YES_WRITE". Працюю як DRY-RUN.');
    dryRun = true;
  }

  Logger.log('[syncMissing] ═══════════════════════════════════════');
  Logger.log('[syncMissing] START dryRun=%s', dryRun);

  // === 1. Оплати-Рік ===
  var paySheet = getCRMSpreadsheet().getSheetByName(SHEET_YEARLY);
  if (!paySheet){ Logger.log('[syncMissing] ❌ "%s" не знайдено', SHEET_YEARLY); return {ok:false, error:'no payment sheet'}; }
  var pvals = paySheet.getDataRange().getValues();
  if (pvals.length < 2){ Logger.log('[syncMissing] ❌ Оплати-Рік порожній'); return {ok:false, error:'empty'}; }
  var phdrs = pvals[0].map(String);
  var pLocIdx   = phdrs.indexOf('Локація');
  var pGroupIdx = phdrs.indexOf('Група');
  var pTeachIdx = phdrs.indexOf('Вихователь');
  var pNameIdx  = phdrs.indexOf("Ім'я дитини");
  // v6.26.1: річні підсумки для критерію «реальна дитина = має суму»
  var pBudRikIdx  = phdrs.indexOf('Бюджет-Рік');
  var pFaktRikIdx = phdrs.indexOf('Факт-Рік');
  if (pNameIdx < 0 || pLocIdx < 0){
    Logger.log('[syncMissing] ❌ Оплати-Рік: name/loc cols missing (name=%s loc=%s)', pNameIdx, pLocIdx);
    return {ok:false, error:'payment cols missing'};
  }
  if (pBudRikIdx < 0 || pFaktRikIdx < 0){
    Logger.log('[syncMissing] ⚠ Бюджет-Рік=%s Факт-Рік=%s — критерій суми не спрацює коректно. Запусти aggregatePaymentsYearly.', pBudRikIdx, pFaktRikIdx);
  }

  // === 2. Клієнти (НАПРЯМУ — всі рядки, навіть з порожнім ID) ===
  var cSheet = getCRMSpreadsheet().getSheetByName(SHEET_CLIENTS);
  if (!cSheet){ Logger.log('[syncMissing] ❌ "%s" не знайдено', SHEET_CLIENTS); return {ok:false, error:'no clients sheet'}; }
  ensureClientsHeader(cSheet);
  var cvals = cSheet.getDataRange().getValues();
  var chdrs = cvals[0].map(String);
  var cNameIdx = chdrs.indexOf('ПІБ дитини');
  var cLocIdx  = chdrs.indexOf('Локація');

  function normKey(name, loc){
    var n = String(name||'').trim().replace(/\s+/g,' ').toLowerCase();
    var l = String(loc ||'').trim().replace(/\s+/g,' ').toLowerCase();
    return n + '|' + l;
  }
  function isServiceRow(name){
    var n = String(name||'').trim();
    if (!n) return true;
    if (n.toLowerCase().indexOf('тест') === 0) return true;  // "ТЕСТ Документи" тощо
    return false;
  }
  function genChildId(name, group, loc){
    return 'c_' + String(name||'').trim().slice(0,20) + '_' + String(group||'').slice(0,8) + '_' + String(loc||'').slice(0,8);
  }

  var existing = {};
  for (var cr = 1; cr < cvals.length; cr++){
    var cn = cvals[cr][cNameIdx];
    if (!String(cn||'').trim()) continue;
    existing[normKey(cn, cvals[cr][cLocIdx])] = true;
  }
  Logger.log('[syncMissing] Клієнти: %s рядків даних, %s унікальних name+loc ключів',
    cvals.length - 1, Object.keys(existing).length);

  // === 3. Знайти missing — фільтр як у робочих модулях + критерій «має суму» ===
  // Реальна дитина проходить ВСІ умови:
  //   1) НЕ isGroupHeaderRow (групи + вільних/місць/разом/всього/оплата за)
  //   2) НЕ суто числове ім'я (/^\d+$/) — відсіює службові слоти "9","10","11"
  //   3) НЕ "тест*" і не порожнє
  //   4) Бюджет-Рік > 0 OR Факт-Рік > 0 (фінансова активність)
  var missing = [];
  var seenNew = {};
  var skip = {header:0, numeric:0, test:0, zeroSum:0, noLoc:0, existing:0, dupInPay:0};
  var skipNumericSamples = [];
  for (var pr = 1; pr < pvals.length; pr++){
    var prow    = pvals[pr];
    var name    = String(prow[pNameIdx] || '').trim();
    var loc     = String(prow[pLocIdx]  || '').trim();
    var group   = pGroupIdx >= 0 ? String(prow[pGroupIdx] || '').trim() : '';
    var teacher = pTeachIdx >= 0 ? String(prow[pTeachIdx] || '').trim() : '';
    var budRik  = pBudRikIdx  >= 0 ? (Number(prow[pBudRikIdx])  || 0) : 0;
    var faktRik = pFaktRikIdx >= 0 ? (Number(prow[pFaktRikIdx]) || 0) : 0;

    // (1) службовий заголовок / групова назва
    if (isGroupHeaderRow(prow, 1)){ skip.header++; continue; }
    // (2) суто числове ім'я — слоти "9","10"
    if (/^\d+$/.test(name)){ skip.numeric++; if (skipNumericSamples.length < 10) skipNumericSamples.push(name); continue; }
    // (3) тест / порожнє
    if (isServiceRow(name)){ skip.test++; continue; }
    // локація обовʼязкова
    if (!loc){ skip.noLoc++; continue; }
    // (4) має фінансову активність за рік
    if (budRik <= 0 && faktRik <= 0){ skip.zeroSum++; continue; }

    var key = normKey(name, loc);
    if (existing[key]){ skip.existing++; continue; }   // вже є в Клієнти
    if (seenNew[key]){ skip.dupInPay++; continue; }     // дубль у самій Оплати-Рік
    seenNew[key] = true;
    missing.push({name:name, loc:loc, group:group, teacher:teacher,
      budRik:budRik, faktRik:faktRik, id:genChildId(name, group, loc)});
  }

  // === 4. Лог відсіяних + по локаціях + перші 50 ===
  Logger.log('[syncMissing] ─────────────────────────────────────');
  Logger.log('[syncMissing] ВІДСІЯНО службових/нерелевантних:');
  Logger.log('  · header (групи/вільних/разом):  %s', skip.header);
  Logger.log('  · числове ім\'я ("9","10"...):    %s  %s', skip.numeric,
    skipNumericSamples.length ? '(приклади: ' + JSON.stringify(skipNumericSamples) + ')' : '');
  Logger.log('  · тест/порожнє:                  %s', skip.test);
  Logger.log('  · без локації:                   %s', skip.noLoc);
  Logger.log('  · нульова сума (Бюджет+Факт=0):   %s', skip.zeroSum);
  Logger.log('  · вже є в Клієнти:               %s', skip.existing);
  Logger.log('  · дубль у Оплати-Рік:            %s', skip.dupInPay);
  Logger.log('[syncMissing] ─────────────────────────────────────');

  var byLoc = {};
  missing.forEach(function(m){ byLoc[m.loc] = (byLoc[m.loc]||0) + 1; });
  Logger.log('[syncMissing] РЕАЛЬНИХ missing (після фільтра 1+2+3+4): %s', missing.length);
  Logger.log('[syncMissing] По локаціях:');
  Object.keys(byLoc).sort().forEach(function(l){ Logger.log('  · %s = %s', l, byLoc[l]); });
  Logger.log('[syncMissing] ─────────────────────────────────────');
  Logger.log('[syncMissing] Перші 50 (name | локація | група | бюджет-рік | future ID):');
  missing.slice(0, 50).forEach(function(m, i){
    Logger.log('  %s. "%s" | %s | %s | %s ₴ | %s', i+1, m.name, m.loc, m.group, m.budRik, m.id);
  });

  // === 5. Запис (тільки якщо НЕ dryRun) ===
  if (dryRun){
    Logger.log('[syncMissing] ─────────────────────────────────────');
    Logger.log('[syncMissing] DRY-RUN — нічого не записано. %s рядків БУДЕ створено при dryRun=false', missing.length);
    Logger.log('[syncMissing] ═══════════════════════════════════════');
    return {ok:true, dryRun:true, missingCount:missing.length, byLoc:byLoc, skip:skip};
  }

  if (!missing.length){
    Logger.log('[syncMissing] Нема чого створювати — 0 missing');
    Logger.log('[syncMissing] ═══════════════════════════════════════');
    return {ok:true, dryRun:false, created:0};
  }

  var now = formatDate(new Date());
  // Порядок колонок ТОЧНО як ensureClientsHeader (35 колонок).
  var newRows = missing.map(function(m){
    return [
      m.id, m.name, m.loc, m.group, m.teacher,
      '',                          // Дата народження
      '', '', '', '',              // ПІБ/Тел мами, ПІБ/Тел тата
      '', 'standard', 0, 0,        // Дата договору, Тип, Сума, Вступний
      'active',                    // Статус
      'Чернетка — автосинхронізація з Оплати-Рік ' + now,  // Нотатки
      '[]', '[]', '[]',            // Відсутності/Графік/Зміни JSON
      '', '', '', '',              // Номер договору, Дата адапт, Розірв, Причина
      '', '',                      // Свідоцтво, Місце реєстрації
      '', '', '', '',              // Документи/РНОКПП мами і тата
      now, now,                    // Створено, Оновлено
      '', '',                      // Email мами, Email тата
      'mom', ''                    // Підписант договору, Номер дод. договору
    ];
  });
  var startRow = cSheet.getLastRow() + 1;
  cSheet.getRange(startRow, 1, newRows.length, newRows[0].length).setValues(newRows);
  Logger.log('[syncMissing] ✅ СТВОРЕНО %s рядків у Клієнти (від row %s)', newRows.length, startRow);
  Logger.log('[syncMissing] ═══════════════════════════════════════');
  return {ok:true, dryRun:false, created:newRows.length, byLoc:byLoc};
}

function testSyncMissingClientsDryRun(){
  return syncMissingClientsFromPayments({dryRun: true});
}
// v6.26.1: НЕ запускає real напряму (після інциденту 832). Для реального
// запису виклич ВРУЧНУ: syncMissingClientsFromPayments({dryRun:false, confirm:'YES_WRITE'})
function testSyncMissingClientsREAL(){
  Logger.log('[syncMissing] ⛔ Цей wrapper НЕ робить real-запис.');
  Logger.log('[syncMissing] Для реального запису обери у dropdown: runSyncMissingREAL_792');
  return {ok:false, blocked:true, hint:'use runSyncMissingREAL_792 in dropdown'};
}

// v6.26.1 ONE-SHOT: реальний запис ~792 карток (свідомо підтверджено користувачем).
// Безпечно при повторному запуску: вже заведені діти підуть у skip.existing,
// missing стане 0 — нічого не задублюється. Після завершення синку — видалити
// цю функцію разом з іншими діагностичними при чистці перед 1C.
function runSyncMissingREAL_792(){
  return syncMissingClientsFromPayments({dryRun: false, confirm: 'YES_WRITE'});
}

// ═══════════════════════════════════════════════════════════════════════════
// v6.11.12: syncGroupsFromPayments — Платежі (Оплати-Рік) = ДЖЕРЕЛО ПРАВДИ для
// поля "Група". У Платежах групи стоять правильно (кожна дитина у своїй групі).
// Старі ~270 карток у Клієнти мають застарілі/неповні назви груп; нові 792
// синхронізовані картки вже коректні (бо бралися з Платежів). Беремо назву
// групи з Платежів по ключу name+локація і ТОЧКОВО оновлюємо "Група" в Клієнти.
// ЗАПОБІЖНИК (як у sync): dryRun=true default; реальний запис лише при
// opts.confirm === 'YES_WRITE'. Чіпає ТІЛЬКИ "Група" + дописує мітку в
// "Нотатки" ("Група оновлена з Платежів DD.MM.YYYY") — інші поля не перезаписує.
// ═══════════════════════════════════════════════════════════════════════════
function syncGroupsFromPayments(opts){
  opts = opts || {};
  var dryRun = (opts.dryRun !== false);
  if (!dryRun && opts.confirm !== 'YES_WRITE'){
    Logger.log('[syncGroups] ⚠ REAL-режим заблоковано: бракує opts.confirm="YES_WRITE". Працюю як DRY-RUN.');
    dryRun = true;
  }

  Logger.log('[syncGroups] ═══════════════════════════════════════');
  Logger.log('[syncGroups] START dryRun=%s', dryRun);

  function normKey(name, loc){
    var n = String(name||'').trim().replace(/\s+/g,' ').toLowerCase();
    var l = String(loc ||'').trim().replace(/\s+/g,' ').toLowerCase();
    return n + '|' + l;
  }
  function normGroup(g){
    return String(g||'').trim().replace(/\s+/g,' ').toLowerCase();
  }
  function isServiceRow(name){
    var n = String(name||'').trim();
    if (!n) return true;
    if (n.toLowerCase().indexOf('тест') === 0) return true;  // "ТЕСТ Документи" тощо
    return false;
  }

  // === 1. Платежі (Оплати-Рік) → map ключ(name+loc) → група з Платежів ===
  var paySheet = getCRMSpreadsheet().getSheetByName(SHEET_YEARLY);
  if (!paySheet){ Logger.log('[syncGroups] ❌ "%s" не знайдено', SHEET_YEARLY); return {ok:false, error:'no payment sheet'}; }
  var pvals = paySheet.getDataRange().getValues();
  if (pvals.length < 2){ Logger.log('[syncGroups] ❌ Оплати-Рік порожній'); return {ok:false, error:'empty'}; }
  var phdrs = pvals[0].map(String);
  var pLocIdx   = phdrs.indexOf('Локація');
  var pGroupIdx = phdrs.indexOf('Група');
  var pNameIdx  = phdrs.indexOf("Ім'я дитини");
  if (pNameIdx < 0 || pLocIdx < 0 || pGroupIdx < 0){
    Logger.log('[syncGroups] ❌ Оплати-Рік: cols missing (name=%s loc=%s group=%s)', pNameIdx, pLocIdx, pGroupIdx);
    return {ok:false, error:'payment cols missing'};
  }

  var payGroup = {};      // key → група з Платежів (точна стрічка, перша непорожня)
  var paySeen  = {};      // key → true: дитина присутня в Платежах (навіть з порожньою групою)
  var payConflicts = {};  // key → [групи]: у Платежах кілька різних груп для одної дитини
  for (var pr = 1; pr < pvals.length; pr++){
    var prow = pvals[pr];
    if (isGroupHeaderRow(prow, 1)) continue;
    var pname = String(prow[pNameIdx] || '').trim();
    if (!pname || /^\d+$/.test(pname) || isServiceRow(pname)) continue;
    var ploc  = String(prow[pLocIdx] || '').trim();
    if (!ploc) continue;
    var key = normKey(pname, ploc);
    paySeen[key] = true;
    var pgrp = String(prow[pGroupIdx] || '').trim();
    if (!pgrp) continue;                       // порожня група в Платежах → не джерело правди
    if (!payGroup[key]){
      payGroup[key] = pgrp;
    } else if (normGroup(payGroup[key]) !== normGroup(pgrp)){
      (payConflicts[key] = payConflicts[key] || [payGroup[key]]).push(pgrp);
    }
  }
  Logger.log('[syncGroups] Платежі: %s дітей присутні, %s з непорожньою групою',
    Object.keys(paySeen).length, Object.keys(payGroup).length);
  var conflictKeys = Object.keys(payConflicts);
  if (conflictKeys.length){
    Logger.log('[syncGroups] ⚠ У Платежах %s дітей мають >1 різну групу (беру ПЕРШУ). Приклади:', conflictKeys.length);
    conflictKeys.slice(0,10).forEach(function(k){ Logger.log('   · %s → %s', k, JSON.stringify(payConflicts[k])); });
  }

  // === 2. Клієнти ===
  var cSheet = getCRMSpreadsheet().getSheetByName(SHEET_CLIENTS);
  if (!cSheet){ Logger.log('[syncGroups] ❌ "%s" не знайдено', SHEET_CLIENTS); return {ok:false, error:'no clients sheet'}; }
  ensureClientsHeader(cSheet);
  var cvals = cSheet.getDataRange().getValues();
  var chdrs = cvals[0].map(String);
  var cNameIdx  = chdrs.indexOf('ПІБ дитини');
  var cLocIdx   = chdrs.indexOf('Локація');
  var cGroupIdx = chdrs.indexOf('Група');
  var cNotesIdx = chdrs.indexOf('Нотатки');
  if (cNameIdx < 0 || cLocIdx < 0 || cGroupIdx < 0){
    Logger.log('[syncGroups] ❌ Клієнти: cols missing (name=%s loc=%s group=%s)', cNameIdx, cLocIdx, cGroupIdx);
    return {ok:false, error:'clients cols missing'};
  }

  // === 3. Знайти кандидатів на оновлення групи ===
  var candidates = [];   // {row, name, loc, oldGroup, newGroup}
  var skip = {notInPay:0, payEmpty:0, same:0, test:0, noName:0};
  for (var cr = 1; cr < cvals.length; cr++){
    var crow = cvals[cr];
    var cname = String(crow[cNameIdx] || '').trim();
    if (!cname){ skip.noName++; continue; }
    if (isServiceRow(cname)){ skip.test++; continue; }        // "ТЕСТ Документи" → не чіпати
    var cloc = String(crow[cLocIdx] || '').trim();
    var key = normKey(cname, cloc);
    var newGrp = payGroup[key];
    if (newGrp === undefined){
      if (paySeen[key]) skip.payEmpty++;                       // є в Платежах, але група там порожня → не чіпати
      else              skip.notInPay++;                       // дитини нема в Платежах → не чіпати
      continue;
    }
    var oldGrp = String(crow[cGroupIdx] || '').trim();
    if (normGroup(oldGrp) === normGroup(newGrp)){ skip.same++; continue; }  // однакова → не чіпати
    candidates.push({row: cr + 1, name: cname, loc: cloc, oldGroup: oldGrp, newGroup: newGrp});
  }

  // === 4. Лог кандидатів (ПОВНИЙ список, згрупований по локаціях) ===
  var byLoc = {};
  candidates.forEach(function(c){ (byLoc[c.loc] = byLoc[c.loc] || []).push(c); });
  var locNames = Object.keys(byLoc).sort();
  Logger.log('[syncGroups] ─────────────────────────────────────');
  Logger.log('[syncGroups] ПРОПУЩЕНО: нема в Платежах=%s | група в Платежах порожня=%s | однакова=%s | тест=%s | без імені=%s',
    skip.notInPay, skip.payEmpty, skip.same, skip.test, skip.noName);
  Logger.log('[syncGroups] ─────────────────────────────────────');
  Logger.log('[syncGroups] КАНДИДАТИ на оновлення групи (повний список):');
  locNames.forEach(function(loc){
    var list = byLoc[loc];
    Logger.log('[syncGroups] ▼ %s (%s)', loc, list.length);
    list.forEach(function(c){
      Logger.log('   %s | %s | Клієнти=%s → Платежі=%s',
        c.name, c.loc, JSON.stringify(c.oldGroup), JSON.stringify(c.newGroup));
    });
  });
  Logger.log('[syncGroups] ─────────────────────────────────────');
  Logger.log('[syncGroups] Знайдено %s карток для оновлення в %s локаціях', candidates.length, locNames.length);

  // === 5. DRY-RUN — вихід без запису ===
  if (dryRun){
    Logger.log('[syncGroups] DRY-RUN — нічого не записано. %s карток БУДЕ оновлено при {dryRun:false, confirm:"YES_WRITE"}', candidates.length);
    Logger.log('[syncGroups] ═══════════════════════════════════════');
    return {ok:true, dryRun:true, count:candidates.length, locations:locNames.length,
      skip:skip, byLoc:locNames.map(function(l){ return {loc:l, n:byLoc[l].length}; })};
  }

  // === 6. РЕАЛЬНИЙ ЗАПИС — точково "Група" + дописати мітку в "Нотатки" ===
  if (!candidates.length){
    Logger.log('[syncGroups] Нема чого оновлювати — 0 кандидатів');
    Logger.log('[syncGroups] ═══════════════════════════════════════');
    return {ok:true, dryRun:false, updated:0};
  }
  var stamp = Utilities.formatDate(new Date(), 'Europe/Kiev', 'dd.MM.yyyy');
  var note = 'Група оновлена з Платежів ' + stamp;
  candidates.forEach(function(c){
    // 6a. Група — одна комірка (точково, інші поля недоторкані)
    cSheet.getRange(c.row, cGroupIdx + 1).setValue(c.newGroup);
    // 6b. Нотатки — ДОПИСАТИ мітку для можливого відкату, не перезаписуючи наявний текст
    if (cNotesIdx >= 0){
      var prev = String(cvals[c.row - 1][cNotesIdx] || '').trim();
      cSheet.getRange(c.row, cNotesIdx + 1).setValue(prev ? (prev + ' | ' + note) : note);
    }
  });
  Logger.log('[syncGroups] ✅ ОНОВЛЕНО групу у %s картках. Мітка в Нотатках: "%s"', candidates.length, note);
  Logger.log('[syncGroups] ═══════════════════════════════════════');
  return {ok:true, dryRun:false, updated:candidates.length, locations:locNames.length};
}

function testSyncGroupsFromPaymentsDryRun(){
  return syncGroupsFromPayments({dryRun: true});
}
// Для реального запису виклич ВРУЧНУ (ТІЛЬКИ після спільної перевірки dry-run):
//   syncGroupsFromPayments({dryRun:false, confirm:'YES_WRITE'})
function runSyncGroupsFromPaymentsREAL(){
  return syncGroupsFromPayments({dryRun: false, confirm: 'YES_WRITE'});
}

// ───────────────────────────────────────────────────────────────────────────
// v6.26.1 ROLLBACK: видаляє рядки, створені syncMissingClientsFromPayments.
// Безпечно — по МІТЦІ в полі Нотатки ("Чернетка — автосинхронізація"), а не
// по номеру рядка (тому навіть якщо нумерація зсунулась — видалить лише наше).
//
// ДВА РЕЖИМИ:
//   opts.dryRun = true (default) → лише показує що БУДЕ видалено
//   opts.dryRun = false          → реально видаляє
//
// SAFETY: real-видалення блокується якщо знайдена кількість != opts.expectedCount
// (default 832). Якщо реальна кількість інша — виклич з явним expectedCount.
// ───────────────────────────────────────────────────────────────────────────
function rollbackSyncedClients(opts){
  opts = opts || {};
  var dryRun = opts.dryRun !== false;  // default true
  var expectedCount = opts.expectedCount != null ? Number(opts.expectedCount) : 832;
  var MARKER = 'Чернетка — автосинхронізація';

  Logger.log('[rollback] ═══════════════════════════════════════');
  Logger.log('[rollback] START dryRun=%s expectedCount=%s', dryRun, expectedCount);
  Logger.log('[rollback] Marker (у полі Нотатки): "%s"', MARKER);

  var cSheet = getCRMSpreadsheet().getSheetByName(SHEET_CLIENTS);
  if (!cSheet){ Logger.log('[rollback] ❌ "%s" не знайдено', SHEET_CLIENTS); return {ok:false, error:'no sheet'}; }
  var vals = cSheet.getDataRange().getValues();
  var headers = vals[0].map(String);
  var notesIdx = headers.indexOf('Нотатки');
  var nameIdx  = headers.indexOf('ПІБ дитини');
  if (notesIdx < 0){ Logger.log('[rollback] ❌ Колонка "Нотатки" не знайдена'); return {ok:false, error:'no notes col'}; }

  Logger.log('[rollback] Клієнти зараз: %s рядків даних (%s з шапкою)', vals.length - 1, vals.length);

  // Знайти рядки з міткою
  var markedRows = [];  // 1-based sheet rows
  for (var r = 1; r < vals.length; r++){
    if (String(vals[r][notesIdx] || '').indexOf(MARKER) !== -1){
      markedRows.push({sheetRow: r + 1, name: String(vals[r][nameIdx] || '')});
    }
  }

  var count = markedRows.length;
  var minRow = count ? markedRows[0].sheetRow : 0;
  var maxRow = count ? markedRows[count-1].sheetRow : 0;
  var contiguous = count > 0 && (maxRow - minRow + 1 === count);

  // Перевірка: скільки рядків БЕЗ мітки (реальні картки — НЕ чіпаємо).
  // Очікувано 270 (рядки 2..271 у sheet). Якщо marked починаються з row 272 —
  // це підтверджує що реальні дані недоторкані.
  var cleanCount = (vals.length - 1) - count;
  var firstMarkedRow = count ? markedRows[0].sheetRow : 0;

  Logger.log('[rollback] ─────────────────────────────────────');
  Logger.log('[rollback] Знайдено рядків з міткою: %s', count);
  Logger.log('[rollback] Діапазон marked: row %s … %s | contiguous=%s', minRow, maxRow, contiguous);
  Logger.log('[rollback] Рядків БЕЗ мітки (реальні картки, НЕ чіпаємо): %s', cleanCount);
  Logger.log('[rollback] Перший marked рядок: %s (очікувано 272 — реальні 1..271 цілі)', firstMarkedRow);
  if (count){
    Logger.log('[rollback] Перші 10 marked (sheetRow | name):');
    markedRows.slice(0, 10).forEach(function(m){ Logger.log('    row %s | "%s"', m.sheetRow, m.name); });
    Logger.log('[rollback] Останні 10 marked (sheetRow | name):');
    markedRows.slice(-10).forEach(function(m){ Logger.log('    row %s | "%s"', m.sheetRow, m.name); });
  }
  Logger.log('[rollback] Після видалення залишиться: %s рядків даних', cleanCount);

  // DRY-RUN: показуємо все, але нічого не чіпаємо
  if (dryRun){
    if (count !== expectedCount){
      Logger.log('[rollback] ⚠ УВАГА: count=%s != expectedCount=%s — real-режим БУДЕ ЗАБЛОКОВАНО', count, expectedCount);
      Logger.log('[rollback] Якщо %s — правильна кількість, виклич real з opts.expectedCount=%s', count, count);
    } else {
      Logger.log('[rollback] ✓ count=%s == expectedCount — real-режим дозволено', count);
    }
    Logger.log('[rollback] DRY-RUN — нічого не видалено.');
    Logger.log('[rollback] ═══════════════════════════════════════');
    return {ok:true, dryRun:true, wouldDelete:count, minRow:minRow, maxRow:maxRow,
            contiguous:contiguous, remainAfter:(vals.length-1)-count, safetyPass:(count===expectedCount)};
  }

  // REAL — SAFETY CHECK
  if (count !== expectedCount){
    Logger.log('[rollback] ⚠⚠⚠ STOP: знайдено %s рядків, очікувалось %s. НІЧОГО НЕ ВИДАЛЕНО.', count, expectedCount);
    Logger.log('[rollback] Якщо %s — справді правильна кількість, виклич з opts.expectedCount=%s', count, count);
    Logger.log('[rollback] ═══════════════════════════════════════');
    return {ok:false, error:'count mismatch', found:count, expected:expectedCount};
  }
  if (count === 0){
    Logger.log('[rollback] Нема рядків з міткою — нічого видаляти');
    return {ok:true, dryRun:false, deleted:0};
  }

  if (contiguous){
    cSheet.deleteRows(minRow, count);
    Logger.log('[rollback] ✅ Видалено %s послідовних рядків: deleteRows(%s, %s)', count, minRow, count);
  } else {
    // fallback: по одному у ЗВОРОТНОМУ порядку (щоб індекси не зсувались)
    for (var i = markedRows.length - 1; i >= 0; i--){
      cSheet.deleteRow(markedRows[i].sheetRow);
    }
    Logger.log('[rollback] ✅ Видалено %s рядків по одному (зворотній порядок)', count);
  }

  var after = cSheet.getDataRange().getValues().length - 1;
  Logger.log('[rollback] Тепер у Клієнти: %s рядків даних', after);
  Logger.log('[rollback] ═══════════════════════════════════════');
  return {ok:true, dryRun:false, deleted:count, remainNow:after};
}

function testRollbackDryRun(){
  return rollbackSyncedClients({dryRun: true});
}
function testRollbackREAL(){
  return rollbackSyncedClients({dryRun: false});
}

// ───────────────────────────────────────────────────────────────────────────
// v6.26.1 ДІАГНОСТИКА форматів груп: виявляє "розколоті" групи, де одна реальна
// група представлена двома форматами — "Baby-ki" (стара картка) і "Baby-ki
// Наталія" (нова синхронізована). Через це invoices.html групує їх окремо.
// Стара/нова визначається по мітці "Чернетка — автосинхронізація" в Нотатках.
// Read-only. Запуск: testDiagGroupFormats().
// ───────────────────────────────────────────────────────────────────────────
function diagGroupFormats(loc){
  loc = String(loc || 'Осокорки').trim();
  var MARKER = 'Чернетка — автосинхронізація';

  var crmRes = getClients();
  if (!crmRes.ok){ Logger.log('[diagGroups] ❌ getClients failed: %s', crmRes.error); return crmRes; }
  var all = crmRes.data || [];

  Logger.log('[diagGroups] ═══════════════════════════════════════');
  Logger.log('[diagGroups] Деталізація для локації: %s', loc);

  // === Детальна частина для loc ===
  var locClients = all.filter(function(c){
    if (String(c['Локація']||'').trim() !== loc) return false;
    var st = String(c['Статус']||'').trim();
    return st === 'active' || st === 'adaptation';
  });

  var byGroup = {};  // group → {total, old, new, names:[]}
  locClients.forEach(function(c){
    var g = String(c['Група']||'').trim() || '(порожня)';
    var isNew = String(c['Нотатки']||'').indexOf(MARKER) !== -1;
    if (!byGroup[g]) byGroup[g] = {total:0, old:0, new:0, names:[]};
    byGroup[g].total++;
    if (isNew) byGroup[g].new++; else byGroup[g].old++;
    if (byGroup[g].names.length < 12) byGroup[g].names.push(String(c['ПІБ дитини']||''));
  });

  var groups = Object.keys(byGroup).sort(function(a,b){ return a.localeCompare(b,'uk'); });
  Logger.log('[diagGroups] %s: %s дітей (active+adapt), %s унікальних форматів груп', loc, locClients.length, groups.length);
  Logger.log('[diagGroups] ─────────────────────────────────────');
  groups.forEach(function(g){
    var d = byGroup[g];
    Logger.log('  "%s" → %s дітей (стара: %s, нова: %s) — [%s%s]',
      g, d.total, d.old, d.new, d.names.slice(0,8).join(', '), d.names.length > 8 ? ', …' : '');
  });

  // === Пари-розколи для loc ===
  Logger.log('[diagGroups] ─────────────────────────────────────');
  Logger.log('[diagGroups] ПАРИ-РОЗКОЛИ для %s (base ↔ base+вихователь):', loc);
  var pairsFound = 0;
  groups.forEach(function(a){
    groups.forEach(function(b){
      if (a !== b && b.indexOf(a + ' ') === 0){
        Logger.log('  "%s" (%s дітей) ↔ "%s" (%s дітей) — ймовірно одна група',
          a, byGroup[a].total, b, byGroup[b].total);
        pairsFound++;
      }
    });
  });
  if (!pairsFound) Logger.log('  (розколів не знайдено)');

  // === Пункт 5: розколи по ВСІХ локаціях ===
  Logger.log('[diagGroups] ═══════════════════════════════════════');
  Logger.log('[diagGroups] РОЗКОЛОТІ ГРУПИ ПО ВСІХ ЛОКАЦІЯХ:');
  var byLocGroups = {};  // loc → {group → count}
  all.forEach(function(c){
    var st = String(c['Статус']||'').trim();
    if (st !== 'active' && st !== 'adaptation') return;
    var l = String(c['Локація']||'').trim();
    var g = String(c['Група']||'').trim();
    if (!l || !g) return;
    if (!byLocGroups[l]) byLocGroups[l] = {};
    byLocGroups[l][g] = (byLocGroups[l][g]||0) + 1;
  });
  var totalSplits = 0;
  Object.keys(byLocGroups).sort().forEach(function(l){
    var gs = Object.keys(byLocGroups[l]);
    var splitsHere = [];
    gs.forEach(function(a){
      gs.forEach(function(b){
        if (a !== b && b.indexOf(a + ' ') === 0){
          splitsHere.push('"' + a + '" (' + byLocGroups[l][a] + ') ↔ "' + b + '" (' + byLocGroups[l][b] + ')');
          totalSplits++;
        }
      });
    });
    if (splitsHere.length){
      Logger.log('  %s: %s розкол(ів)', l, splitsHere.length);
      splitsHere.forEach(function(s){ Logger.log('     %s', s); });
    }
  });
  Logger.log('[diagGroups] ─────────────────────────────────────');
  Logger.log('[diagGroups] ВСЬОГО розколотих пар по мережі: %s', totalSplits);
  Logger.log('[diagGroups] ═══════════════════════════════════════');

  return {ok:true, loc:loc, locGroupFormats:groups.length, totalSplits:totalSplits};
}

function testDiagGroupFormats(){
  return diagGroupFormats('Осокорки');
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

// ── PHANTOM-CATALOG CLEANUP (v6.28.3, editor-only) ─────────────────────────
// Видаляє рядки Предметники_Каталог, чия Локація НЕ існує у getLocations
// (CONFIG sheet — джерело правди). Напр. дубль «Осокорки сад» (5 рядків,
// ідентичних «Осокорки»). dryRun за замовчуванням; реальне видалення —
// лише з opts.confirm='YES_DELETE'. Зразок: syncGroupsFromPayments.
//   testCleanupPhantomPredCatalog()         → dry-run (нічого не пише)
//   runCleanupPhantomPredCatalogREAL()       → реальне видалення
function cleanupPhantomPredCatalog(opts){
  opts = opts || {};
  var dryRun = (opts.dryRun !== false);
  if (!dryRun && opts.confirm !== 'YES_DELETE'){
    Logger.log('[cleanupPhantomCatalog] ⚠ REAL-режим заблоковано: бракує confirm="YES_DELETE". Працюю як DRY-RUN.');
    dryRun = true;
  }

  // 1. Канонічні локації з CONFIG.
  var locRes = getLocations();
  if (!locRes || !locRes.ok) return {ok:false, error:'getLocations failed'};
  var canonical = {};
  locRes.data.forEach(function(l){ if (l && l.loc) canonical[String(l.loc).trim()] = true; });

  // 2. Скан каталогу → фантомні рядки (loc не в canonical).
  var sh = _getPredmetnyCatalogSheet(false);
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return {ok:true, dryRun:dryRun, phantomRows:0, deleted:0, byLoc:{}};
  var data = sh.getRange(2, 1, lastRow - 1, PREDMETNY_CATALOG_HEADER.length).getValues();

  var phantom = [];               // {rowNum, loc, subject, rate, id}
  var byLoc = {};
  for (var i = 0; i < data.length; i++){
    var rec = _parsePredmetnyCatRow(data[i]);
    if (!rec.loc) continue;
    if (canonical[rec.loc]) continue;             // легітимна локація — пропускаємо
    phantom.push({rowNum: i + 2, loc: rec.loc, subject: rec.subject, rate: rec.rate, id: rec.id});
    byLoc[rec.loc] = (byLoc[rec.loc] || 0) + 1;
  }

  Logger.log('[cleanupPhantomCatalog] ─────────────────────────────');
  Logger.log('[cleanupPhantomCatalog] dryRun=%s | фантомних рядків=%s', dryRun, phantom.length);
  Object.keys(byLoc).sort().forEach(function(loc){
    Logger.log('   ▼ %s (%s)', loc, byLoc[loc]);
  });
  phantom.forEach(function(p){
    Logger.log('     row %s | %s | %s | ставка=%s | id=%s', p.rowNum, p.loc, p.subject, p.rate, p.id);
  });

  if (dryRun){
    Logger.log('[cleanupPhantomCatalog] DRY-RUN — нічого не видалено. Для видалення: runCleanupPhantomPredCatalogREAL()');
    return {ok:true, dryRun:true, phantomRows:phantom.length, byLoc:byLoc,
            rows:phantom};
  }

  // 3. Реальне видалення — з НИЗУ догори (щоб індекси не зсувались).
  var rowNums = phantom.map(function(p){ return p.rowNum; }).sort(function(a, b){ return b - a; });
  for (var j = 0; j < rowNums.length; j++){
    sh.deleteRow(rowNums[j]);
  }
  Logger.log('[cleanupPhantomCatalog] ✅ Видалено %s рядків.', rowNums.length);
  return {ok:true, dryRun:false, phantomRows:phantom.length, deleted:rowNums.length,
          byLoc:byLoc, rows:phantom};
}

function testCleanupPhantomPredCatalog(){
  var res = cleanupPhantomPredCatalog({dryRun:true});
  Logger.log('RESULT: ' + JSON.stringify(res));
  return res;
}
function runCleanupPhantomPredCatalogREAL(){
  var res = cleanupPhantomPredCatalog({dryRun:false, confirm:'YES_DELETE'});
  Logger.log('RESULT: ' + JSON.stringify(res));
  return res;
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
      ['Фітнес', 250]
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
// "М'яка" нормалізація — lowercase, крапки/пробіли/NBSP → один пробіл.
// Зберігає межі слів (на відміну від _journalNormName, що склеює все).
function _softNorm(s){
  return String(s == null ? '' : s).toLowerCase()
    .replace(/[.\s ]+/g, ' ').trim();
}
// Чи містить haystack рядок needle як окреме СЛОВО (межі — край рядка
// або не літера/цифра). Літери — будь-якого алфавіту (\p{L}).
function _softContainsWord(haystack, needle){
  if (!needle) return false;
  var from = 0;
  while (from <= haystack.length){
    var pos = haystack.indexOf(needle, from);
    if (pos === -1) return false;
    var before = pos > 0 ? haystack.charAt(pos - 1) : '';
    var after  = haystack.charAt(pos + needle.length);
    var bw = before !== '' && /[\p{L}\p{N}]/u.test(before);
    var aw = after  !== '' && /[\p{L}\p{N}]/u.test(after);
    if (!bw && !aw) return true;
    from = pos + 1;
  }
  return false;
}

// Знаходить рядок Salary для предмета каталогу за пріоритетами P1–P6.
//   salaryRows — [{row, raw, norm, soft}] непорожні рядки колонки A.
//   subject, rate — предмет і ставка з каталогу предметників.
// Повертає {row, matchedAs, priority:'P1'..'P6'} або null (P7 — рядка нема).
function _findPredmetnySalaryRow(salaryRows, subject, rate){
  var subjNorm  = _journalNormName(subject);              // "логопед"
  var exactNorm = _journalNormName(subject + ' ' + rate); // "логопед250"
  var subjSoft  = _softNorm(subject);                     // "муз керівник"
  if (!subjNorm || !subjSoft) return null;
  var rateStr = String(rate);
  var i, n, best, bestDiff, nums, x;
  function hit(r, p){ return {row: r.row, matchedAs: r.raw, priority: p}; }

  // P1 — точний збіг "предмет ставка".
  for (i = 0; i < salaryRows.length; i++)
    if (salaryRows[i].norm === exactNorm) return hit(salaryRows[i], 'P1');

  // P2 — префікс "предмет ставка ..." (далі не цифра — ставка та сама).
  for (i = 0; i < salaryRows.length; i++){
    n = salaryRows[i].norm;
    if (n.length > exactNorm.length && n.indexOf(exactNorm) === 0
        && !/[0-9]/.test(n.charAt(exactNorm.length)))
      return hit(salaryRows[i], 'P2');
  }

  // P3 — предмет НА ПОЧАТКУ + будь-яка ставка одразу після; найближча.
  best = null; bestDiff = Infinity;
  for (i = 0; i < salaryRows.length; i++){
    n = salaryRows[i].norm;
    if (n.indexOf(subjNorm) !== 0) continue;
    var m3 = n.slice(subjNorm.length).match(/^([0-9]+)/);
    if (!m3) continue;
    var d3 = Math.abs(Number(m3[1]) - Number(rate));
    if (d3 < bestDiff){ bestDiff = d3; best = salaryRows[i]; }
  }
  if (best) return hit(best, 'P3');

  // P4 — предмет як СЛОВО-підрядок + у рядку є число, що === ставці каталогу.
  for (i = 0; i < salaryRows.length; i++){
    if (!_softContainsWord(salaryRows[i].soft, subjSoft)) continue;
    nums = salaryRows[i].soft.match(/[0-9]+/g);
    if (nums && nums.indexOf(rateStr) !== -1) return hit(salaryRows[i], 'P4');
  }

  // P5 — предмет як СЛОВО-підрядок + будь-яке число; найближче до ставки.
  best = null; bestDiff = Infinity;
  for (i = 0; i < salaryRows.length; i++){
    if (!_softContainsWord(salaryRows[i].soft, subjSoft)) continue;
    nums = salaryRows[i].soft.match(/[0-9]+/g);
    if (!nums) continue;
    for (x = 0; x < nums.length; x++){
      var d5 = Math.abs(Number(nums[x]) - Number(rate));
      if (d5 < bestDiff){ bestDiff = d5; best = salaryRows[i]; }
    }
  }
  if (best) return hit(best, 'P5');

  // P6 — предмет як СЛОВО-підрядок, БЕЗ жодного числа. Гейту немає:
  // якщо дійшли сюди — рядка предметника зі ставкою у Salary нема,
  // отже рядок без ставки належить предметнику (каталог підтверджує).
  for (i = 0; i < salaryRows.length; i++){
    if (_softContainsWord(salaryRows[i].soft, subjSoft)
        && !/[0-9]/.test(salaryRows[i].soft))
      return hit(salaryRows[i], 'P6');
  }

  return null;   // P7 — рядка немає, треба додати
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
      salaryRows.push({row: k + 1, raw: raw,
        norm: _journalNormName(raw), soft: _softNorm(raw)});
    }

    var budgetColValues   = sheet.getRange(1, budgetCol, lastRow, 1).getValues();
    var budgetColFormulas = sheet.getRange(1, budgetCol, lastRow, 1).getFormulas();
    var journal = _readJournalForTarget(loc, 'predmetnyky', nextM.year, nextM.month);

    var journalOps = [];
    var updated = 0, totalFact = 0, cellsWritten = 0, formulaRowsSkipped = 0;
    var notFound = [], details = [], p7queue = [], maxMatchedRow = 0;
    var stats = {attempts: 0, p1: 0, p2: 0, p3: 0, p4: 0, p5: 0, p6: 0, p7: 0};

    // 4. Матчинг кожного предмета каталогу → рядок Salary (P1–P7).
    withRate.forEach(function(a){
      var uniq = lessonsBySubj[a.id] ? Object.keys(lessonsBySubj[a.id]).length : 0;
      var fact = uniq * a.rate;
      var catName = a.subject + ' ' + a.rate;       // ключ журналу — з каталогу
      var nk = _journalNormName(catName);
      stats.attempts++;

      var found = _findPredmetnySalaryRow(salaryRows, a.subject, a.rate);
      if (!found){
        stats.p7++;
        p7queue.push({subject: a.subject, rate: a.rate, fact: fact, lessons: uniq,
          catName: catName, nk: nk});
        Logger.log('[%s] %s → P7 рядка немає — буде ДОДАНО у Salary', loc, catName);
        return;
      }
      stats['p' + found.priority.slice(1)]++;
      Logger.log('[%s] %s → matched %s на A%s "%s"', loc, catName, found.priority, found.row, found.matchedAs);
      if (found.row > maxMatchedRow) maxMatchedRow = found.row;

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

    // 5. P7 — рядків немає: додаємо нові у Salary після останнього
    // зіставленого рядка предметників (якщо жоден не зіставився — у кінець
    // листа). Вставки — ПІСЛЯ всіх записів P1–P6, тож позиції не зсуваються.
    p7queue.forEach(function(p){
      var newRow;
      if (maxMatchedRow > 0){
        sheet.insertRowsAfter(maxMatchedRow, 1);
        newRow = maxMatchedRow + 1;
        maxMatchedRow = newRow;            // наступний P7 — нижче цього
      } else {
        newRow = sheet.getLastRow() + 1;   // у локації жоден не зіставився
      }
      sheet.getRange(newRow, 1).setValue(p.subject + ' ' + p.rate);
      sheet.getRange(newRow, budgetCol).setValue(p.fact);
      cellsWritten++;
      journalOps.push({nk: p.nk, loc: loc, kind: 'predmetnyky', name: p.catName,
        year: nextM.year, month: nextM.month, newSum: p.fact});
      updated++;
      totalFact += p.fact;
      details.push({subject: p.catName, fact: p.fact, lessons: p.lessons,
        priority: 'P7', row: newRow, status: 'row-added'});
      Logger.log('[%s] %s → P7 ДОДАНО рядок A%s = %s₴', loc, p.catName, newRow, p.fact);
    });

    _commitJournalUpdates(journal, journalOps);
    Logger.log('[%s] СВОДКА: спроб=%s | P1=%s P2=%s P3=%s P4=%s P5=%s P6=%s P7=%s | клітинок=%s формул-пропущено=%s',
      loc, stats.attempts, stats.p1, stats.p2, stats.p3, stats.p4, stats.p5, stats.p6, stats.p7,
      cellsWritten, formulaRowsSkipped);

    return {
      ok: true,
      updated: updated,
      totalFact: totalFact,
      notFound: notFound,
      rowsAdded: stats.p7,
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
  var LOCS = sortByLocationOrder(['Голосієво','Бігова','Борщагівка','Бровари',"Кар'єрна",'Кругла',
              'Оранж','Осокорки','Позняки','Пуща','Тичини']);   // v6.50.4 зонний порядок
  var lines = [];
  LOCS.forEach(function(loc){
    var r = exportPredmetnyToSalary({loc: loc, month: month, year: year});
    var s = (r && r.matchStats) || {attempts:0,p1:0,p2:0,p3:0,p4:0,p5:0,p6:0,p7:0};
    lines.push(loc + ' | ' + s.attempts + ' | ' + s.p1 + ' | ' + s.p2 + ' | ' +
      s.p3 + ' | ' + s.p4 + ' | ' + s.p5 + ' | ' + s.p6 + ' | ' + s.p7 +
      ((r && r.ok) ? '' : '  ❌ ' + (r && r.error)));
  });
  Logger.log('\n════════ SUMMARY предметники → Salary (місяць %s/%s) ════════', month, year);
  Logger.log('Локація | Спроб | P1 точн | P2 преф | P3 інша ст | P4 підр+ст | P5 підр+чис | P6 без ст | P7 додано');
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
function fixPolyakova(){
  var sh = _getUsersSheet();
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++){
    if (String(data[i][1] || '').indexOf('Поляков') >= 0){
      Logger.log('Рядок %s | ПІБ=%s | логін=%s | роль=%s | активний=%s', i+1, data[i][1], data[i][2], data[i][4], data[i][6]);
      var r = setUserPassword(String(data[i][2]).trim(), 'Mkids2026Rnd', 1);
      Logger.log('Скидання пароля: %s', JSON.stringify(r));
      return;
    }
  }
  Logger.log('Полякову не знайдено в листі Користувачі');
}

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

// ── РАЗОВА УТИЛІТА: почистити баги у CONFIG → Salary ─────────────────────
// 1) trim() для всіх значень у колонках A (Напрямок), B (Тип), C (Локація),
//    E (Назва листа) — щоб "Школа " і "Школа" були однаковими.
// 2) Для рядків де Локація ∈ {Житомир, Нац.Гвардії (Благо), Манхетен (Благо)}
//    виставити Напрямок = "Управління" (зараз там некоректно "Садочок").
// Запускати ВРУЧНУ з Apps Script editor — БЕЗ автотригерів.
// Логує кожну зміну в Execution log.
function _fixSalaryConfigBugs(){
  var configSS = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  var sh = configSS.getSheetByName('Salary');
  if (!sh){ Logger.log('!!! CONFIG → Salary tab not found'); return {ok:false}; }

  var rng  = sh.getDataRange();
  var data = rng.getValues();
  if (data.length < 2){ Logger.log('Salary registry порожній'); return {ok:false}; }

  var MGMT_LOCS = ['Житомир', 'Нац.Гвардії (Благо)', 'Манхетен (Благо)'];
  var trimCols  = [0, 1, 2, 4];   // A, B, C, E (D = Spreadsheet ID — не чіпаємо)

  var changes = [];

  for (var i = 1; i < data.length; i++){
    var rowNum = i + 1; // 1-based sheet row

    // 1) trim усіх 4 колонок.
    trimCols.forEach(function(c){
      var v = data[i][c];
      if (v == null) return;
      var str  = String(v);
      var trim = str.trim();
      if (trim !== str){
        changes.push({row: rowNum, col: c, from: JSON.stringify(str), to: JSON.stringify(trim), kind: 'trim'});
        data[i][c] = trim;
      }
    });

    // 2) Управління: фікс Напрямок (A) якщо Локація (C) у списку.
    var loc = String(data[i][2] || '').trim();
    if (MGMT_LOCS.indexOf(loc) !== -1){
      var oldNapr = String(data[i][0] || '').trim();
      if (oldNapr !== 'Управління'){
        changes.push({row: rowNum, col: 0, from: JSON.stringify(oldNapr), to: '"Управління"', kind: 'mgmt-napr', loc: loc});
        data[i][0] = 'Управління';
      }
    }
  }

  if (!changes.length){
    Logger.log('[_fixSalaryConfigBugs] нічого виправляти — CONFIG чистий.');
    return {ok:true, changed:0};
  }

  // Лог кожної зміни.
  changes.forEach(function(ch){
    Logger.log('  row %s col %s [%s] %s → %s%s',
      ch.row, ch.col, ch.kind, ch.from, ch.to,
      ch.loc ? (' (loc=' + ch.loc + ')') : '');
  });
  Logger.log('[_fixSalaryConfigBugs] усього змін: %s', changes.length);

  // Один батч-запис на весь діапазон.
  rng.setValues(data);
  return {ok:true, changed: changes.length, details: changes};
}

// ── ДІАГНОСТИКА: предметники у Salary-файлах ───────────────────────────────
// auditPredmetnyInSalary() — для всіх 11 локацій порівнює каталог предметників
// із секцією "Вчителі-предметники" у Salary-файлі локації. Запускати ВРУЧНУ
// з Apps Script editor; звіт — у Execution log.
function auditPredmetnyInSalary(){
  var LOCS = sortByLocationOrder(['Голосієво','Бігова','Борщагівка','Бровари',"Кар'єрна",'Кругла',
              'Оранж','Осокорки','Позняки','Пуща','Тичини']);   // v6.50.4 зонний порядок
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
  var LOCS = sortByLocationOrder(['Голосієво','Бігова','Борщагівка','Бровари',"Кар'єрна",'Кругла',
              'Оранж','Осокорки','Позняки','Пуща','Тичини']);   // v6.50.4 зонний порядок
  var ICON = {P1:'✅', P2:'🔵', P3:'🟡', P4:'🟠', P5:'🟣', P6:'🟢', P7:'❌'};
  var NOTE = {
    P2: ' (префікс зі ставкою)',
    P3: ' (та сама назва, інша ставка — найближча)',
    P4: ' (назва-підрядок + ставка збігається)',
    P5: ' (назва-підрядок + інша ставка — найближча)',
    P6: ' (назва без ставки)'
  };

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

  var summary = [], doubtful = [];
  Logger.log('═══════════ DRY-RUN: предметники → Salary | місяць %s/%s ═══════════', month, year);

  LOCS.forEach(function(loc){
    Logger.log('\n[ЛОКАЦІЯ %s]', loc);
    var catItems = ((getPredmetnyCatalog(loc) || {}).items || [])
      .filter(function(a){ return a.active && a.rate > 0; });

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
      if (raw) salaryRows.push({row: k + 1, raw: raw,
        norm: _journalNormName(raw), soft: _softNorm(raw)});
    }

    var lessonsBySubj = {};
    for (var i = 1; i < attData.length; i++){
      var rec = _parsePredmetnyAttRow(attData[i]);
      if (rec.loc !== loc) continue;
      if (rec.date < dateFrom || rec.date >= dateTo) continue;
      if (!lessonsBySubj[rec.subjectId]) lessonsBySubj[rec.subjectId] = {};
      lessonsBySubj[rec.subjectId][rec.group + '|' + rec.date] = true;
    }

    var st = {p1:0,p2:0,p3:0,p4:0,p5:0,p6:0,p7:0}, sumFact = 0;
    catItems.forEach(function(a){
      var uniq = lessonsBySubj[a.id] ? Object.keys(lessonsBySubj[a.id]).length : 0;
      var fact = uniq * a.rate;
      var catName = a.subject + ' ' + a.rate;
      sumFact += fact;
      var found = _findPredmetnySalaryRow(salaryRows, a.subject, a.rate);
      if (!found){
        st.p7++;
        Logger.log('  %s P7: %s → НЕ ЗНАЙДЕНО — ДОДАТИ рядок до Salary (fact=%s₴)', ICON.P7, catName, fact);
        doubtful.push('❌ [' + loc + '] ' + catName + ' → P7 ДОДАТИ рядок · fact=' + fact + '₴');
        return;
      }
      st['p' + found.priority.slice(1)]++;
      Logger.log('  %s %s: %s → A%s "%s"%s', ICON[found.priority], found.priority,
        catName, found.row, found.matchedAs, (NOTE[found.priority] || ''));
      if (found.priority !== 'P1' && found.priority !== 'P2'){
        doubtful.push(ICON[found.priority] + ' [' + loc + '] ' + catName + ' → ' +
          found.priority + ' A' + found.row + ' "' + found.matchedAs + '" · fact=' + fact + '₴');
      }
    });

    Logger.log('  ── каталог=%s | P1=%s P2=%s P3=%s P4=%s P5=%s P6=%s P7=%s | сума fact=%s₴',
      catItems.length, st.p1, st.p2, st.p3, st.p4, st.p5, st.p6, st.p7, sumFact);
    summary.push(loc + ' | ' + catItems.length + ' | ' + st.p1 + ' | ' + st.p2 +
      ' | ' + st.p3 + ' | ' + st.p4 + ' | ' + st.p5 + ' | ' + st.p6 + ' | ' + st.p7 +
      ' | ' + sumFact);
  });

  Logger.log('\n═══════════ ЗВЕДЕНА ТАБЛИЦЯ ═══════════');
  Logger.log('Локація | Каталог | P1 | P2 | P3 | P4 | P5 | P6 | P7 | Сума fact (грн)');
  summary.forEach(function(s){ Logger.log('  ' + s); });

  Logger.log('\n═══════════ СУМНІВНІ МАТЧІ — перевірити вручну (P3–P7) ═══════════');
  Logger.log('  P3 інша ставка · P4 підрядок+ставка · P5 підрядок+інша ст · P6 без ставки · P7 буде додано');
  if (!doubtful.length) Logger.log('  (немає — усі матчі точні P1/P2)');
  doubtful.forEach(function(d){ Logger.log('  ' + d); });

  return {ok: true};
}

// ── ДІАГНОСТИКА section-based класифікатора по локаціях ─────────────────
// READ-ONLY — не пише нічого. Запускати ВРУЧНУ з Apps Script editor.
//
//   testClassifyAllLocations()  → всі 16 локацій з Salary registry
//   testClassifyRealLocations() → 3 локації (швидкий smoke-test)
//
// Для кожної: підсумки по 3 секціях (main / subjects / extras) + список
// потенційних проблем. У кінці — зведена таблиця зі статусами:
//   ✅ ok       — нерозпізнаних нема
//   ⚠️ warn    — null-category або subject без сум
//   ❌ fail    — порожня локація або помилка
function testClassifyAllLocations(){
  return _runClassifyDiagnostic(null);
}

function testClassifyRealLocations(){
  return _runClassifyDiagnostic(["Кар'єрна", 'Голосієво', 'Осокорки']);
}

function _runClassifyDiagnostic(locFilter){
  var YEAR  = 2026;
  var MONTH = 5;   // травень
  var fIdx  = (MONTH - 1) * 3 + 1;
  var bIdx  = (MONTH - 1) * 3 + 2;

  var configSS = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  var salReg   = configSS.getSheetByName('Salary');
  if (!salReg){ Logger.log('!!! CONFIG → Salary tab not found'); return {ok:false}; }

  // Реєстр у source-order — щоб зведена таблиця збереглася як у CONFIG.
  // Колонка 1 (Тип) використовується для фільтра проблем: повний "ПОТЕНЦІЙНІ
  // ПРОБЛЕМИ" блок показується тільки для Тип='Садочок' (Школи/Управління —
  // у summary, але без розгорнутого аналізу).
  var entries = [];
  var regRows = salReg.getDataRange().getValues();
  for (var i = 1; i < regRows.length; i++){
    var sTyp = String(regRows[i][1] || '').trim();
    var sLoc = String(regRows[i][2] || '').trim();
    var sId  = String(regRows[i][3] || '').trim();
    var sLst = String(regRows[i][4] || '').trim() || 'Salary';
    if (!sLoc || !sId) continue;
    if (locFilter && locFilter.indexOf(sLoc) === -1) continue;
    entries.push({typ: sTyp, loc: sLoc, sheetId: sId, listName: sLst});
  }

  function padR(s, n){ s = String(s == null ? '' : s); return s + new Array(Math.max(1, n - s.length + 1)).join(' '); }
  function padL(s, n){ s = String(s == null ? '' : s); return new Array(Math.max(1, n - s.length + 1)).join(' ') + s; }

  var summary = [];

  entries.forEach(function(ent){
    var isSadochok = ent.typ === 'Садочок';
    Logger.log('\n══════════════ %s [%s] ══════════════', ent.loc, ent.typ || '?');
    var stat = {
      loc: ent.loc, typ: ent.typ,
      main: 0, subjects: 0, extras: 0,
      bM: 0, bS: 0, bE: 0,
      fM: 0, fS: 0, fE: 0,
      status: '✅', issues: 0
    };

    try {
      var sh = SpreadsheetApp.openById(ent.sheetId).getSheetByName(ent.listName);
      if (!sh){ Logger.log('  ❌ немає листа "%s"', ent.listName); stat.status = '❌'; summary.push(stat); return; }

      var lastRow = Math.max(sh.getLastRow(), 80);
      var lastCol = Math.max(sh.getLastColumn(), 37);
      var data    = sh.getRange(1, 1, lastRow, lastCol).getValues();
      var width   = lastCol;

      var raw = [];
      for (var r = 4; r <= data.length; r++){
        var idx = r - 1;
        var rowArr = data[idx] || [];
        var name = String(rowArr[0] || '').trim();
        if (_salaryIsSkippedRow(name)) continue;
        var fact   = fIdx < width ? _opexNum(rowArr[fIdx]) : 0;
        var budget = bIdx < width ? _opexNum(rowArr[bIdx]) : 0;
        raw.push({row: r, name: name, fact: fact, budget: budget});
      }

      if (!raw.length){
        Logger.log('  (порожньо — у Salary файлі немає даних з рядка 4)');
        stat.status = '❌';
        summary.push(stat);
        return;
      }

      var classified = _classifyAllSalaryRows(raw);

      // Підрахунки.
      classified.forEach(function(cr){
        if (cr._category === 'section_header' || cr._category === 'group_header') return;
        var s = cr._section;
        var fF = Number(cr.fact) || 0, bB = Number(cr.budget) || 0;
        if (s === 'main')          { stat.main++;     stat.bM += bB; stat.fM += fF; }
        else if (s === 'subjects') { stat.subjects++; stat.bS += bB; stat.fS += fF; }
        else if (s === 'extras')   { stat.extras++;   stat.bE += bB; stat.fE += fF; }
      });

      Logger.log('  ── ПІДСУМКИ за травень %s ──', YEAR);
      Logger.log('  main     : %s рядків · бюджет=%s₴ · факт=%s₴', padL(stat.main, 3),     stat.bM, stat.fM);
      Logger.log('  subjects : %s рядків · бюджет=%s₴ · факт=%s₴', padL(stat.subjects, 3), stat.bS, stat.fS);
      Logger.log('  extras   : %s рядків · бюджет=%s₴ · факт=%s₴', padL(stat.extras, 3),   stat.bE, stat.fE);

      // ── ПОТЕНЦІЙНІ ПРОБЛЕМИ — тільки для Тип='Садочок' ──
      // Для Школа/Управління блок проблем приховано — у них своя структура
      // Salary-файлів, класифікатор поки заточений під 11 садочків.
      if (isSadochok){
        var problems = [];

        // 1) null-category у main — нерозпізнаний штат.
        classified.forEach(function(cr){
          if (cr._category === null && cr._section === 'main' && cr.name){
            problems.push({sev:'warn',
              text: 'row ' + cr.row + ' "' + cr.name + '" → НЕРОЗПІЗНАНО (main+null)'});
          }
        });

        // 2) subject без бюджету і без факту — підозра на помилку.
        classified.forEach(function(cr){
          if (cr._category === 'subject' && Number(cr.fact) === 0 && Number(cr.budget) === 0){
            problems.push({sev:'warn',
              text: 'row ' + cr.row + ' "' + cr.name + '" → subject але 0 бюджет + 0 факт'});
          }
        });

        // 3) Підсекція "Школа" — group_header + усе до наступного group/section header.
        var school = null;
        for (var k = 0; k < classified.length; k++){
          var cr = classified[k];
          var nm = String(cr.name || '').trim();
          if (cr._category === 'group_header' && /^школа$/i.test(nm)){
            school = {start: cr.row, rows: []};
            continue;
          }
          if (school){
            if (cr._category === 'group_header' || cr._category === 'section_header') break;
            school.rows.push(cr);
          }
        }
        if (school){
          problems.push({sev:'info',
            text: 'Підсекція "Школа" (з row ' + school.start + ', ' + school.rows.length + ' рядків):'});
          school.rows.forEach(function(cr){
            problems.push({sev:'info',
              text: '    row ' + cr.row + ' "' + cr.name + '" → cat=' + cr._category + ' sec=' + cr._section});
          });
        }

        if (problems.length){
          Logger.log('\n  ── ПОТЕНЦІЙНІ ПРОБЛЕМИ ──');
          problems.forEach(function(p){ Logger.log('  - ' + p.text); });
          stat.issues = problems.filter(function(p){ return p.sev === 'warn'; }).length;
        }
        if (stat.issues > 0) stat.status = '⚠️';
      } else {
        // Школа/Управління — статус нейтральний (не аналізуємо).
        Logger.log('  ── (Школа/Управління — діагностика проблем поки не включена) ──');
        stat.status = '➖';
      }

      if (!stat.main && !stat.subjects && !stat.extras) stat.status = '❌';

    } catch (e){
      Logger.log('  ❌ ERROR: %s', (e && e.message) ? e.message : String(e));
      stat.status = '❌';
    }

    summary.push(stat);
  });

  // ── ЗВЕДЕНА ТАБЛИЦЯ ──
  Logger.log('\n\n═════════════════════════════ ЗВЕДЕНА ТАБЛИЦЯ ═════════════════════════════');
  Logger.log('┌────────────────────────┬────────────┬──────┬─────┬──────┬────────┐');
  Logger.log('│ Локація                │ Тип        │ main │ sub │ extr │ status │');
  Logger.log('├────────────────────────┼────────────┼──────┼─────┼──────┼────────┤');
  summary.forEach(function(s){
    Logger.log('│ ' + padR(s.loc, 22) + ' │ ' +
               padR(s.typ || '?', 10) + ' │ ' +
               padL(s.main, 4) + ' │ ' +
               padL(s.subjects, 3) + ' │ ' +
               padL(s.extras, 4) + ' │ ' +
               padR(s.status, 6) + ' │');
  });
  Logger.log('└────────────────────────┴────────────┴──────┴─────┴──────┴────────┘');

  var counts = {ok:0, warn:0, fail:0, neutral:0};
  summary.forEach(function(s){
    if      (s.status === '✅') counts.ok++;
    else if (s.status === '⚠️') counts.warn++;
    else if (s.status === '➖') counts.neutral++;
    else                        counts.fail++;
  });
  Logger.log('\n  ✅ ok=%s  ⚠️ warn=%s  ➖ neutral=%s  ❌ fail=%s  усього=%s',
    counts.ok, counts.warn, counts.neutral, counts.fail, summary.length);

  return {ok: true, summary: summary};
}

// ═══════════════════════════════════════════════════════════════════════════
//  HR — КАРТКА СПІВРОБІТНИКА (v6.9)
// ═══════════════════════════════════════════════════════════════════════════
// Backend для додавання/редагування/soft-delete співробітника.
// Permission-модель:
//   director (own loc) | mgmt (any loc): CRUD; nurse/vyhovatel (own loc): read-only
//   delete (soft, fired=today): тільки cfo/ceo/hr/coo
// Audit log: CONFIG.HR_Audit [ts, actorId, actorName, action, rowNum, before, after]
// Колонки P (formula) і Q (reserved) не чіпаються при update.

// ── Sheets ──────────────────────────────────────────────────────────
function _getHrSheet(){
  var ss = SpreadsheetApp.openById(HR_SHEET_ID);
  var sh = ss.getSheetByName(HR_TAB_NAME);
  if (!sh) throw new Error('HR sheet tab "' + HR_TAB_NAME + '" not found');
  return sh;
}

// ═══════════════════════════════════════════════════════════════════
// UNIFY LVIV LOCATION NAMES (editor-only) — лише HR Співробітники.
// «Львів Кругла»→«Кругла», «Львів Бігова»→«Бігова» (синхрон з CONFIG).
// Торкаємось ВИКЛЮЧНО колонки C (Локація) листа HR '2025-2026'.
//   diagUnifyLvivLocationNames()      — dry-run (нічого не пише)
//   runUnifyLvivLocationNamesREAL()   — реальна заміна + backup у Props
//   runRollbackUnifyLvivLocationNames() — відкат з backup
// ═══════════════════════════════════════════════════════════════════
var UNIFY_LVIV_MAP        = { 'Львів Кругла':'Кругла', 'Львів Бігова':'Бігова' };
var UNIFY_LVIV_LOC_COL    = 3;                      // колонка C = Локація (1-based)
var UNIFY_LVIV_BACKUP_KEY = 'UNIFY_LVIV_LOC_BACKUP';

function diagUnifyLvivLocationNames(){
  var sh = _getHrSheet();
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return {ok:true, dryRun:true, total:0, toChange:0, byName:{}};
  var locs = sh.getRange(2, UNIFY_LVIV_LOC_COL, lastRow - 1, 1).getValues();
  var byName = {}, rows = [];
  for (var i = 0; i < locs.length; i++){
    var v = String(locs[i][0] == null ? '' : locs[i][0]).trim();
    if (UNIFY_LVIV_MAP.hasOwnProperty(v)){
      byName[v] = (byName[v] || 0) + 1;
      rows.push({row:i + 2, from:v, to:UNIFY_LVIV_MAP[v]});
    }
  }
  Logger.log('[diagUnifyLviv] DRY-RUN | total=%s toChange=%s | %s',
    lastRow - 1, rows.length, JSON.stringify(byName));
  rows.forEach(function(r){ Logger.log('   row %s: «%s» → «%s»', r.row, r.from, r.to); });
  return {ok:true, dryRun:true, total:lastRow - 1, toChange:rows.length, byName:byName, rows:rows};
}

function runUnifyLvivLocationNamesREAL(){
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = _getHrSheet();
    var lastRow = sh.getLastRow();
    if (lastRow < 2) return {ok:true, changed:0, msg:'порожній лист'};
    var range = sh.getRange(2, UNIFY_LVIV_LOC_COL, lastRow - 1, 1);
    var locs = range.getValues();
    var backup = [], byName = {}, changed = 0;
    for (var i = 0; i < locs.length; i++){
      var v = String(locs[i][0] == null ? '' : locs[i][0]).trim();
      if (UNIFY_LVIV_MAP.hasOwnProperty(v)){
        backup.push({row:i + 2, old:locs[i][0]});      // зберігаємо ОРИГІНАЛ
        locs[i][0] = UNIFY_LVIV_MAP[v];
        byName[v] = (byName[v] || 0) + 1;
        changed++;
      }
    }
    if (!changed) return {ok:true, changed:0, msg:'нема що змінювати'};
    var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    // backup ПЕРЕД записом (для відкату). Лише якщо ще нема бекапу — щоб
    // повторний запуск не затер оригінали.
    var props = PropertiesService.getScriptProperties();
    if (!props.getProperty(UNIFY_LVIV_BACKUP_KEY)){
      props.setProperty(UNIFY_LVIV_BACKUP_KEY, JSON.stringify(
        {ts:stamp, sheetId:HR_SHEET_ID, tab:HR_TAB_NAME, col:UNIFY_LVIV_LOC_COL, items:backup}));
    }
    range.setValues(locs);
    var res = {ok:true, changed:changed, byName:byName, backupTs:stamp};
    Logger.log('[runUnifyLvivREAL] ' + JSON.stringify(res));
    return res;
  } catch(e){
    Logger.log('[runUnifyLvivREAL] EXCEPTION: %s', e && e.message);
    return {ok:false, error: e && e.message || String(e)};
  } finally {
    try { lock.releaseLock(); } catch(_e){}
  }
}

function runRollbackUnifyLvivLocationNames(){
  var raw = PropertiesService.getScriptProperties().getProperty(UNIFY_LVIV_BACKUP_KEY);
  if (!raw) return {ok:false, error:'Немає бекапу (runUnifyLvivLocationNamesREAL ще не запускався)'};
  var b;
  try { b = JSON.parse(raw); } catch(e){ return {ok:false, error:'Бекап пошкоджений'}; }
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = _getHrSheet();
    var col = b.col || UNIFY_LVIV_LOC_COL;
    var restored = 0;
    (b.items || []).forEach(function(it){
      sh.getRange(it.row, col).setValue(it.old);
      restored++;
    });
    // прибираємо бекап після успішного відкату
    PropertiesService.getScriptProperties().deleteProperty(UNIFY_LVIV_BACKUP_KEY);
    var res = {ok:true, restored:restored, fromTs:b.ts};
    Logger.log('[rollbackUnifyLviv] ' + JSON.stringify(res));
    return res;
  } catch(e){
    return {ok:false, error: e && e.message || String(e)};
  } finally {
    try { lock.releaseLock(); } catch(_e){}
  }
}

function _getHrAuditSheet(){
  var ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  var sh = ss.getSheetByName(HR_AUDIT_SHEET);
  if (!sh){
    sh = ss.insertSheet(HR_AUDIT_SHEET);
    sh.getRange(1, 1, 1, HR_AUDIT_HEADER.length).setValues([HR_AUDIT_HEADER]);
    sh.setFrozenRows(1);
  }
  return sh;
}

// ── Actor / permissions ────────────────────────────────────────────
function _getActor(actorId){
  if (!actorId) throw new Error('actorId required');
  var sh = _getUsersSheet();
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++){
    if (Number(data[i][0]) === Number(actorId)){
      var u = _parseUserRow(data[i]);
      if (!u.active) throw new Error('Actor user is inactive');
      return u;
    }
  }
  throw new Error('Actor not found (id=' + actorId + ')');
}

// Case-insensitive порівняння ролі — бо в Користувачі є 'CFO'/'CEO'/'Legal',
// а наші константи lowercase.
function _roleKey(r){ return String(r == null ? '' : r).trim().toLowerCase(); }

function _empHasMgmtRole(role){ return EMP_MGMT_ROLES.indexOf(_roleKey(role)) !== -1; }
function _empHasDirRole(role){  return EMP_DIR_ROLES.indexOf(_roleKey(role))  !== -1; }
function _empHasViewRole(role){ return EMP_VIEW_ROLES.indexOf(_roleKey(role)) !== -1; }

function _canViewEmployees(actor){
  return _empHasMgmtRole(actor.role) || _empHasDirRole(actor.role) || _empHasViewRole(actor.role);
}

function _canEditEmployee(actor, targetLoc){
  if (_empHasMgmtRole(actor.role)) return true;
  if (_empHasDirRole(actor.role)){
    if (!targetLoc) return false;
    return String(actor.loc || '').trim() === String(targetLoc).trim();
  }
  return false;
}

function _canDeleteEmployee(actor){
  return EMP_DELETE_ROLES.indexOf(_roleKey(actor.role)) !== -1;
}

// director/nurse/vyhovatel → їх loc (read-scope); mgmt → null (full access)
function _empLocScope(actor){
  if (_empHasMgmtRole(actor.role)) return null;
  return String(actor.loc || '').trim() || null;
}

// ── Helpers ─────────────────────────────────────────────────────────
function _normalizePhone(phone){
  var p = String(phone || '').trim().replace(/\s/g, '');
  if (!p) return '';
  if (!p.startsWith('0') && /^\d{9}$/.test(p)) return '0' + p;
  return p;
}

function _fmtDateDmy(v){
  if (!v) return '';
  if (v instanceof Date)
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'dd.MM.yyyy');
  return String(v).trim();
}

// "15.03.2024" / "2024-03-15" / Date → Date (для запису у sheet).
function _parseDateInput(s){
  if (s == null || s === '') return '';
  if (s instanceof Date) return s;
  var str = String(s).trim();
  if (!str) return '';
  var dmy = /^(\d{1,2})\.(\d{1,2})\.(\d{4})$/.exec(str);
  if (dmy) return new Date(Number(dmy[3]), Number(dmy[2]) - 1, Number(dmy[1]));
  var iso = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(str);
  if (iso) return new Date(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3]));
  return str;  // fallback — sheet validate сам
}

// row[18] → Employee object. rowNum — 1-based sheet row (stable ID).
function _parseEmpRow(row, rowNum){
  function s(v){ return String(v == null ? '' : v).trim(); }
  var fired = _fmtDateDmy(row[14]);
  return {
    rowNum:   rowNum,
    dir:      s(row[0]),
    typ:      s(row[1]),
    loc:      s(row[2]),
    grp:      s(row[3]),
    last:     s(row[4]),
    first:    s(row[5]),
    phone:    s(row[6]),
    pos:      s(row[7]),
    stat:     s(row[8]),
    bday:     s(row[9]),
    bmon:     s(row[10]),
    bdate:    _fmtDateDmy(row[11]),
    wday:     _fmtDateDmy(row[12]),
    hired:    _fmtDateDmy(row[13]),
    fired:    fired,
    // P (idx 15) — formula, не повертаємо. Q (idx 16) — reserved.
    email:    s(row[17]),
    // v6.44 — нові поля картки (append-only).
    passport:  s(row[18]),   // S
    rate:      s(row[19]),   // T — Ставка ЗП
    workCond:  s(row[20]),   // U — Умови роботи (Договір/ФОП/ЦПХ/ГПД…)
    calcModel: s(row[21]),   // V — Модель розрахунку (За заняття/За дитину/За захід)
    assessment: s(row[22]),  // W — Оцінка (JSON) v6.48
    matReturn: _fmtDateDmy(row[23]),  // X — Дата виходу з декрету v6.59
    archived: fired !== ''
  };
}

// Payload → 15 елементів A..O (для setValues).
function _payloadToAtoO(p){
  return [
    String(p.dir   || '').trim(),         // A
    String(p.typ   || '').trim(),         // B
    String(p.loc   || '').trim(),         // C
    String(p.grp   || '').trim(),         // D
    String(p.last  || '').trim(),         // E
    String(p.first || '').trim(),         // F
    _normalizePhone(p.phone),             // G
    String(p.pos   || '').trim(),         // H
    String(p.stat  || '').trim(),         // I
    String(p.bday  || '').trim(),         // J — день (рядок)
    String(p.bmon  || '').trim(),         // K — місяць (рядок)
    _parseDateInput(p.bdate),             // L
    _parseDateInput(p.wday),              // M
    _parseDateInput(p.hired),             // N
    _parseDateInput(p.fired)              // O
  ];
}

function _validateEmpPayload(p){
  if (!p || typeof p !== 'object') return 'Empty payload';
  if (!String(p.first || '').trim()) return 'Field "first" (Ім\'я) is required';
  if (!String(p.last  || '').trim()) return 'Field "last" (Прізвище) is required';
  if (!String(p.loc   || '').trim()) return 'Field "loc" (Локація) is required';
  if (!String(p.pos   || '').trim()) return 'Field "pos" (Посада) is required';
  return null;
}

// Дублікат = той самий ПІБ+phone+loc у активних (не archived). excludeRowNum
// дозволяє безпечно update без хибного дубль-конфлікту "сам із собою".
function _findEmpDuplicate(allEmps, payload, excludeRowNum){
  function key(o){
    return (String(o.last  || '').trim() + '|' +
            String(o.first || '').trim() + '|' +
            _normalizePhone(o.phone)     + '|' +
            String(o.loc   || '').trim()
           ).toLowerCase();
  }
  var target = key(payload);
  for (var i = 0; i < allEmps.length; i++){
    if (excludeRowNum && allEmps[i].rowNum === excludeRowNum) continue;
    if (allEmps[i].archived) continue;
    if (key(allEmps[i]) === target) return allEmps[i];
  }
  return null;
}

// ── Audit ──────────────────────────────────────────────────────────
function _writeHrAudit(actor, action, rowNum, before, after){
  try {
    var sh = _getHrAuditSheet();
    sh.appendRow([
      new Date(),
      actor ? actor.id : 0,
      actor ? actor.name : '',
      action || '',
      rowNum || 0,
      JSON.stringify(before || null),
      JSON.stringify(after || null)
    ]);
  } catch(e){
    Logger.log('[_writeHrAudit] error: ' + (e && e.message));
  }
}

// ── PUBLIC API ──────────────────────────────────────────────────────

// getEmployees(actorId, locFilter?) — read список з permission-filter.
//   director/nurse/vyhovatel: автоматично scope=own loc (locFilter ігнор).
//   mgmt: всі (або locFilter якщо передано).
function getEmployees(actorId, locFilter){
  try {
    var actor = _getActor(actorId);
    if (!_canViewEmployees(actor))
      return {ok:false, error:'Permission denied', code:'PERM_DENIED'};

    var scope  = _empLocScope(actor);
    var filter = scope || (String(locFilter || '').trim() || null);

    var sh = _getHrSheet();
    var lastRow = sh.getLastRow();
    if (lastRow < 2) return {ok:true, items:[], totalCount:0, scope: filter || 'all'};

    var data  = sh.getRange(2, 1, lastRow - 1, HR_COLS).getValues();
    var items = [];
    for (var i = 0; i < data.length; i++){
      var row = data[i];
      var hasAny = false;
      for (var c = 0; c < row.length; c++){
        if (row[c] !== '' && row[c] !== null){ hasAny = true; break; }
      }
      if (!hasAny) continue;
      var emp = _parseEmpRow(row, i + 2);
      if (!emp.last && !emp.first) continue;
      if (filter && emp.loc !== filter) continue;
      items.push(emp);
    }
    return {ok:true, items:items, totalCount:items.length, scope: filter || 'all'};
  } catch(e){
    return {ok:false, error: e.message || String(e)};
  }
}

// saveEmployee(actorId, payload, rowNum?) — create або update.
//   rowNum=null/0  → create (appendRow + copy P formula з row 2)
//   rowNum>0       → update (setValues A:O і R; P/Q не чіпаємо)
function saveEmployee(actorId, payload, rowNum){
  try {
    var actor = _getActor(actorId);

    var vErr = _validateEmpPayload(payload);
    if (vErr) return {ok:false, error:vErr, code:'VALIDATION'};

    if (!_canEditEmployee(actor, payload.loc))
      return {ok:false, error:'Permission denied for location "' + payload.loc + '"', code:'PERM_DENIED'};

    var lock = LockService.getDocumentLock();
    if (!lock.tryLock(10000))
      return {ok:false, error:'Could not acquire lock (try again)', code:'LOCK_TIMEOUT'};

    try {
      var sh = _getHrSheet();

      // Усі співробітники у scope актора (для дубль-чеку).
      var allRes = getEmployees(actorId);
      if (!allRes.ok) return allRes;
      var allEmps = allRes.items;

      var newAtoO = _payloadToAtoO(payload);
      var email   = String(payload.email || '').trim();
      // v6.44 — нові поля картки (append-only S:V). Зберігаємо як текст.
      var passport  = String(payload.passport  || '').trim();   // S
      var rate      = String(payload.rate      || '').trim();   // T
      var workCond  = String(payload.workCond  || '').trim();   // U
      var calcModel = String(payload.calcModel || '').trim();   // V
      // v6.48 — assessment (W): об'єкт {templateId:{period:{...}}} → JSON-рядок.
      var assessment = (payload.assessment == null) ? null
                     : (typeof payload.assessment === 'string' ? payload.assessment : JSON.stringify(payload.assessment));
      // v6.59 — X (Дата виходу з декрету). null = ключ не переданий (реактивація/інші
      //   часткові payload) → НЕ затираємо. '' = поле очищене у модалці → пишемо порожнє.
      var matReturn = (payload.matReturn == null) ? null : _parseDateInput(payload.matReturn);

      // ── UPDATE ─────────────────────────────────────────
      if (rowNum){
        rowNum = Number(rowNum) || 0;
        if (!rowNum) return {ok:false, error:'Invalid rowNum', code:'VALIDATION'};

        var existingRow = sh.getRange(rowNum, 1, 1, HR_COLS).getValues()[0];
        var existing    = _parseEmpRow(existingRow, rowNum);

        // Director не може переносити співробітника у іншу локацію
        if (_empHasDirRole(actor.role) && existing.loc && existing.loc !== payload.loc)
          return {ok:false, error:'Director cannot move employee to another location', code:'PERM_DENIED'};

        var dup = _findEmpDuplicate(allEmps, payload, rowNum);
        if (dup) return {ok:false, error:'Duplicate employee in same location (row ' + dup.rowNum + ')', code:'DUPLICATE'};

        // Пишемо A:O (15 cols) і R (1 col). P (formula) + Q (reserved) — НЕ чіпаємо.
        sh.getRange(rowNum, 1,  1, 15).setValues([newAtoO]);
        sh.getRange(rowNum, 18, 1, 1).setValue(email);
        // v6.44 — S:V (Паспорт, Ставка, Умови, Модель). Append-only, поза P/Q.
        sh.getRange(rowNum, 19, 1, 4).setValues([[passport, rate, workCond, calcModel]]);
        // v6.48 — W (assessment) пишемо ЛИШЕ якщо передано (модалка/реактивація не шлють → не затираємо).
        if (assessment != null) sh.getRange(rowNum, 23, 1, 1).setValue(assessment);
        // v6.59 — X (Дата виходу з декрету) пишемо ЛИШЕ якщо ключ переданий.
        if (matReturn != null) sh.getRange(rowNum, 24, 1, 1).setValue(matReturn);

        var updated = _parseEmpRow(sh.getRange(rowNum, 1, 1, HR_COLS).getValues()[0], rowNum);
        _writeHrAudit(actor, 'update', rowNum, existing, updated);
        return {ok:true, rowNum:rowNum, employee:updated};
      }

      // ── CREATE ─────────────────────────────────────────
      var dupNew = _findEmpDuplicate(allEmps, payload, null);
      if (dupNew) return {ok:false, error:'Duplicate employee in same location (row ' + dupNew.rowNum + ')', code:'DUPLICATE'};

      // appendRow з 24 cols: A-O (15) + ['',''] (P,Q) + R (email) + S:V (v6.44) + W (v6.48) + X (v6.59).
      // P-формула — порожня тут, копіюється з row 2 одразу після append.
      var fullRow = newAtoO.concat(['', '', email, passport, rate, workCond, calcModel,
        (assessment != null ? assessment : ''), (matReturn != null ? matReturn : '')]);
      sh.appendRow(fullRow);
      var newRowNum = sh.getLastRow();

      // P (life-cycle formula) — копіюємо з row 2 (relative refs автоматично переставляться).
      try {
        sh.getRange(2, 16).copyTo(
          sh.getRange(newRowNum, 16),
          SpreadsheetApp.CopyPasteType.PASTE_FORMULA,
          false  // transposed
        );
      } catch(e){
        Logger.log('[saveEmployee] copyTo P formula failed: ' + (e && e.message));
      }

      var created = _parseEmpRow(sh.getRange(newRowNum, 1, 1, HR_COLS).getValues()[0], newRowNum);
      _writeHrAudit(actor, 'create', newRowNum, null, created);
      return {ok:true, rowNum:newRowNum, employee:created};
    } finally {
      lock.releaseLock();
    }
  } catch(e){
    return {ok:false, error: e.message || String(e)};
  }
}

// v6.44 — one-shot міграція: підписує заголовки нових колонок S:V у HR-аркуші.
//   Запускати ВРУЧНУ з Apps Script editor (без routing у doPost). Безпечно
//   повторно: пише лише порожні клітинки заголовка, дані рядків не чіпає.
//   saveEmployee і так пише дані у S:V навіть без заголовків — це лише підписи.
function migrateHrAddCardFields(){
  var sh = _getHrSheet();
  var HEAD = { 19:'Паспорт', 20:'Ставка ЗП', 21:'Умови роботи', 22:'Модель розрахунку', 23:'Оцінка (JSON)', 24:'Дата виходу з декрету' };
  var done = [];
  Object.keys(HEAD).forEach(function(col){
    col = Number(col);
    var cell = sh.getRange(1, col);
    if (String(cell.getValue() || '').trim() === ''){
      cell.setValue(HEAD[col]);
      done.push(HEAD[col]);
    }
  });
  Logger.log('[migrateHrAddCardFields] ✓ підписано заголовків: %s', done.length ? JSON.stringify(done) : 'none (вже існували)');
  return {ok:true, added:done};
}

// inspectDevSources() — РОЗВІДНИК для майбутнього імпорту розвитку дітей.
//   Запускати ВРУЧНУ з редактора Apps Script (як migrateHrAddCardFields).
//   НЕ webapp-екшен. Нічого НЕ пише — лише Logger.log.
//   Майстер: A=Локація, B=Група, C=Spreadsheet ID. Бере перші 2 рядки з непорожнім C,
//   відкриває кожен файл і показує аркуші + превʼю першого аркуша (12×26).
function inspectDevSources(){
  var MASTER_ID = '1od1nd818xMEcszMX_WCFdciL63x4X2pSQpd6LMqGDAc';
  Logger.log('═══ inspectDevSources ═══ master=%s', MASTER_ID);

  var master, msheet, data;
  try {
    master = SpreadsheetApp.openById(MASTER_ID);
    msheet = master.getSheets()[0];
    data   = msheet.getDataRange().getValues();
  } catch(e){
    Logger.log('✗ Не вдалося відкрити майстер: %s', e && e.message || e);
    return;
  }
  Logger.log('Майстер: "%s" · аркуш "%s" · рядків=%s', master.getName(), msheet.getName(), data.length);
  Logger.log('Шапка (A/B/C): %s', JSON.stringify((data[0] || []).slice(0, 3)));

  // Перші 2 рядки з непорожнім C (Spreadsheet ID).
  var picked = [];
  for (var i = 1; i < data.length && picked.length < 2; i++){
    var c = String(data[i][2] == null ? '' : data[i][2]).trim();
    if (c) picked.push({ row: i + 1, loc: String(data[i][0] || '').trim(), grp: String(data[i][1] || '').trim(), id: c });
  }
  Logger.log('Знайдено рядків з ID: показую %s', picked.length);

  picked.forEach(function(p, n){
    Logger.log('───── [%s] рядок %s · Локація="%s" · Група="%s" · ID=%s', n + 1, p.row, p.loc, p.grp, p.id);

    // 1) mimeType — підтвердити, чи це нативний Sheet, чи .xlsx у Drive.
    var mime;
    try {
      mime = DriveApp.getFileById(p.id).getMimeType();
      Logger.log('  mimeType: %s', mime);
    } catch(e){
      Logger.log('  ✗ getFileById/getMimeType: %s', e && e.message || e);
      return;
    }

    // 2) Якщо НЕ Google Sheet — конвертуємо у тимчасовий Sheet через Advanced Drive Service.
    var ssId = null, isTemp = false;
    if (mime === MimeType.GOOGLE_SHEETS){
      ssId = p.id;
    } else {
      if (typeof Drive === 'undefined'){
        Logger.log('  ⚠ Advanced Drive Service «Drive» недоступний. Увімкни в редакторі: Services (+) → Drive API, потім перезапусти. Пропускаю цей файл.');
        return;
      }
      try {
        var blob = DriveApp.getFileById(p.id).getBlob();
        var meta = { name: 'tmp_dev_' + new Date().getTime(), mimeType: MimeType.GOOGLE_SHEETS };
        var created;
        if (Drive.Files && typeof Drive.Files.create === 'function'){
          created = Drive.Files.create(meta, blob);                                              // Drive API v3
        } else if (Drive.Files && typeof Drive.Files.insert === 'function'){
          created = Drive.Files.insert({ title: meta.name, mimeType: meta.mimeType }, blob, { convert: true }); // v2
        } else {
          Logger.log('  ⚠ Drive.Files.create/insert не знайдено — невідома версія Drive API. Пропускаю.');
          return;
        }
        ssId = created && (created.id || created.getId && created.getId());
        if (!ssId){ Logger.log('  ✗ Конвертація: не отримав id тимчасового файлу.'); return; }
        isTemp = true;
        Logger.log('  ✓ .xlsx сконвертовано у тимчасовий Sheet id=%s', ssId);
      } catch(e){
        Logger.log('  ✗ Конвертація не вдалася: %s', e && e.message || e);
        return;
      }
    }

    // 3) Читаємо аркуші + превʼю першого (12 × до 26). Тимчасовий файл прибираємо у finally.
    try {
      var ss = SpreadsheetApp.openById(ssId);
      Logger.log('  Sheet: "%s"', ss.getName());
      var sheets = ss.getSheets();
      sheets.forEach(function(sh){
        Logger.log('    • аркуш "%s" [%s рядків x %s колонок]', sh.getName(), sh.getMaxRows(), sh.getMaxColumns());
      });
      var first = sheets[0];
      var rows = Math.min(12, first.getMaxRows());
      var cols = Math.min(26, first.getMaxColumns());
      if (rows && cols){
        var vals = first.getRange(1, 1, rows, cols).getValues();
        Logger.log('  Превʼю першого аркуша "%s" (%s×%s):', first.getName(), rows, cols);
        vals.forEach(function(r, ri){
          Logger.log('    r%s: %s', ri + 1, JSON.stringify(r));
        });
      } else {
        Logger.log('  (перший аркуш порожній)');
      }
    } catch(e){
      Logger.log('  ✗ Читання Sheet не вдалося: %s', e && e.message || e);
    } finally {
      if (isTemp && ssId){
        try {
          DriveApp.getFileById(ssId).setTrashed(true);
          Logger.log('  🗑 тимчасовий файл переміщено у кошик');
        } catch(e){
          Logger.log('  ⚠ не вдалося видалити тимчасовий %s: %s', ssId, e && e.message || e);
        }
      }
    }
  });

  Logger.log('═══ inspectDevSources done ═══');
}

// inspectDevTemplates() — РОЗВІДНИК шаблонів критеріїв по вікових групах.
//   Запускати ВРУЧНУ з редактора Apps Script (як inspectDevSources). НЕ webapp.
//   Нічого НЕ пише — лише Logger.log. Потребує Advanced Drive Service (Drive).
//   По одному файлу на кожну УНІКАЛЬНУ нормалізовану групу (preschool I/II → preschool),
//   макс 6. Для кожного: перший аркуш (шаблон) — колонка A + жирність (секції vs критерії).
function inspectDevTemplates(){
  var MASTER_ID = '1od1nd818xMEcszMX_WCFdciL63x4X2pSQpd6LMqGDAc';
  Logger.log('═══ inspectDevTemplates ═══ master=%s', MASTER_ID);

  // Конвертація .xlsx → тимчасовий Google Sheet. Повертає {ssId, isTemp} або null.
  function toSheet(id){
    var mime;
    try { mime = DriveApp.getFileById(id).getMimeType(); }
    catch(e){ Logger.log('  ✗ getMimeType: %s', e && e.message || e); return null; }
    if (mime === MimeType.GOOGLE_SHEETS) return { ssId: id, isTemp: false };
    if (typeof Drive === 'undefined'){
      Logger.log('  ⚠ Advanced Drive Service «Drive» недоступний. Services (+) → Drive API.');
      return null;
    }
    try {
      var blob = DriveApp.getFileById(id).getBlob();
      var meta = { name: 'tmp_devtpl_' + new Date().getTime(), mimeType: MimeType.GOOGLE_SHEETS };
      var created;
      if (Drive.Files && typeof Drive.Files.create === 'function')      created = Drive.Files.create(meta, blob);                                              // v3
      else if (Drive.Files && typeof Drive.Files.insert === 'function') created = Drive.Files.insert({ title: meta.name, mimeType: meta.mimeType }, blob, { convert: true }); // v2
      else { Logger.log('  ⚠ Drive.Files.create/insert не знайдено.'); return null; }
      var ssId = created && (created.id || created.getId && created.getId());
      if (!ssId){ Logger.log('  ✗ Конвертація: немає id.'); return null; }
      return { ssId: ssId, isTemp: true };
    } catch(e){ Logger.log('  ✗ Конвертація: %s', e && e.message || e); return null; }
  }

  // Нормалізація групи: lowercase + trim + прибрати хвіст " i"/" ii"/" iii"/" 1"/" 2".
  function normGroup(g){
    return String(g == null ? '' : g).toLowerCase().trim().replace(/\s+(i{1,3}|\d+)$/, '').trim();
  }

  var master, data;
  try {
    master = SpreadsheetApp.openById(MASTER_ID);
    data   = master.getSheets()[0].getDataRange().getValues();
  } catch(e){
    Logger.log('✗ Не вдалося відкрити майстер: %s', e && e.message || e);
    return;
  }

  // По одному файлу на унікальну нормалізовану групу (макс 6).
  var seen = {}, picked = [];
  for (var i = 1; i < data.length && picked.length < 6; i++){
    var id = String(data[i][2] == null ? '' : data[i][2]).trim();
    if (!id) continue;
    var loc = String(data[i][0] || '').trim();
    var grp = String(data[i][1] || '').trim();
    var norm = normGroup(grp);
    if (!norm || seen[norm]) continue;
    seen[norm] = true;
    picked.push({ norm: norm, loc: loc, grp: grp, id: id });
  }
  Logger.log('Унікальних груп до перевірки: %s', picked.length);

  picked.forEach(function(p, n){
    Logger.log('=== ГРУПА: %s · %s/%s · %s ===', p.norm, p.loc, p.grp, p.id);
    var conv = toSheet(p.id);
    if (!conv) return;
    try {
      var ss = SpreadsheetApp.openById(conv.ssId);
      var sheets = ss.getSheets();
      var tpl = sheets[0];
      Logger.log('  Шаблон-аркуш: "%s" [%s рядків x %s колонок]', tpl.getName(), tpl.getMaxRows(), tpl.getMaxColumns());

      // Рядок 1: де періоди (B/C/D).
      var hdr = tpl.getRange(1, 1, 1, Math.min(4, tpl.getMaxColumns())).getValues()[0];
      Logger.log('  Шапка: A="%s" B="%s" C="%s" D="%s"', hdr[0] || '', hdr[1] || '', hdr[2] || '', hdr[3] || '');

      // Колонка A, рядки 1..120: текст + жирність (секції=BOLD, критерії=normal).
      var rows = Math.min(120, tpl.getMaxRows());
      var aVals    = tpl.getRange(1, 1, rows, 1).getValues();
      var aWeights = tpl.getRange(1, 1, rows, 1).getFontWeights();
      Logger.log('  Колонка A (непорожні, %s рядків скановано):', rows);
      for (var r = 0; r < rows; r++){
        var txt = String(aVals[r][0] == null ? '' : aVals[r][0]).trim();
        if (!txt) continue;
        var bold = String(aWeights[r][0]).toLowerCase() === 'bold';
        Logger.log('    R%s [%s] "%s"', r + 1, bold ? 'BOLD' : 'normal', txt);
      }

      // ДОДАТКОВО лише для ПЕРШОГО файлу: sheets[1] (перша дитина) — формат оцінок.
      if (n === 0){
        var kid = sheets[1];
        if (kid){
          Logger.log('  ── ПЕРША ДИТИНА: аркуш "%s" — рядки 1..30, колонки A–D ──', kid.getName());
          var kr = Math.min(30, kid.getMaxRows());
          var kc = Math.min(4, kid.getMaxColumns());
          var kvals = kid.getRange(1, 1, kr, kc).getValues();
          kvals.forEach(function(row, ri){
            Logger.log('    r%s: %s', ri + 1, JSON.stringify(row));
          });
        } else {
          Logger.log('  (другого аркуша (дитини) немає)');
        }
      }
    } catch(e){
      Logger.log('  ✗ Читання не вдалося: %s', e && e.message || e);
    } finally {
      if (conv.isTemp && conv.ssId){
        try { DriveApp.getFileById(conv.ssId).setTrashed(true); Logger.log('  🗑 тимчасовий файл — у кошик'); }
        catch(e){ Logger.log('  ⚠ не видалив temp %s: %s', conv.ssId, e && e.message || e); }
      }
    }
  });

  Logger.log('═══ inspectDevTemplates done ═══');
}

// buildDevTemplatesJS() — генерує готовий JS-обʼєкт DEVELOPMENT_TEMPLATES (вікові шаблони)
//   із файлів-джерел. Запускати ВРУЧНУ з редактора Apps Script. НЕ webapp.
//   Нічого НЕ пише — лише Logger.log. Потребує Advanced Drive Service (Drive).
//   Секції розпізнаються ПО канонічному списку (не по bold). Якщо канонічна секція
//   трапляється вдруге в межах файлу — парсинг цього файлу зупиняється (сміттєвий дубль).
function buildDevTemplatesJS(){
  var MASTER_ID = '1od1nd818xMEcszMX_WCFdciL63x4X2pSQpd6LMqGDAc';
  Logger.log('═══ buildDevTemplatesJS ═══ master=%s', MASTER_ID);

  var SECTIONS = [
    'Емоційний інтелект',
    'Соціальна адаптація',
    'Уміння вирішувати конфлікти',
    'Самостійність у прийнятті рішень',
    'Мовленнєвий розвиток',
    'Сенсорно-пізнавальний розвиток',
    'Ігрова діяльність',
    'Предметно-практична діяльність',
    'Художньо-естетичний розвиток'
  ];
  var SKIP = [
    'Маса тіла','Довжина тіла (зріст)','Обхват голови','Обхват грудної клітки',
    'Особистісно-соціальний розвиток:','Столбец 1'
  ];

  // Нормалізація для порівняння: lowercase + стиснути пробіли + прибрати хвостову пунктуацію.
  function norm(s){
    return String(s == null ? '' : s).toLowerCase().replace(/\s+/g, ' ').trim().replace(/[:;.,\-\s]+$/, '');
  }
  var SECTION_MAP = {};                 // norm → канонічний текст
  SECTIONS.forEach(function(s){ SECTION_MAP[norm(s)] = s; });
  var SKIP_SET = {};
  SKIP.forEach(function(s){ SKIP_SET[norm(s)] = true; });
  function isSkip(nt){ return SKIP_SET[nt] || nt.indexOf('столбец') === 0; }

  // Норм-група → ключ шаблону (preschool I/II → preschool; пробіли прибираємо).
  function groupKey(g){
    return String(g == null ? '' : g).toLowerCase().trim()
      .replace(/\s+(i{1,3}|\d+)$/, '').replace(/\s+/g, '');
  }

  // Конвертація .xlsx → тимчасовий Sheet. Повертає {ssId, isTemp} або null.
  function toSheet(id){
    var mime;
    try { mime = DriveApp.getFileById(id).getMimeType(); }
    catch(e){ Logger.log('  ✗ getMimeType: %s', e && e.message || e); return null; }
    if (mime === MimeType.GOOGLE_SHEETS) return { ssId: id, isTemp: false };
    if (typeof Drive === 'undefined'){ Logger.log('  ⚠ Drive advanced service недоступний.'); return null; }
    try {
      var blob = DriveApp.getFileById(id).getBlob();
      var meta = { name: 'tmp_devbld_' + new Date().getTime(), mimeType: MimeType.GOOGLE_SHEETS };
      var created;
      if (Drive.Files && typeof Drive.Files.create === 'function')      created = Drive.Files.create(meta, blob);
      else if (Drive.Files && typeof Drive.Files.insert === 'function') created = Drive.Files.insert({ title: meta.name, mimeType: meta.mimeType }, blob, { convert: true });
      else { Logger.log('  ⚠ Drive.Files.create/insert не знайдено.'); return null; }
      var ssId = created && (created.id || created.getId && created.getId());
      if (!ssId){ Logger.log('  ✗ Конвертація: немає id.'); return null; }
      return { ssId: ssId, isTemp: true };
    } catch(e){ Logger.log('  ✗ Конвертація: %s', e && e.message || e); return null; }
  }

  var master, data;
  try {
    master = SpreadsheetApp.openById(MASTER_ID);
    data   = master.getSheets()[0].getDataRange().getValues();
  } catch(e){ Logger.log('✗ Майстер не відкрився: %s', e && e.message || e); return; }

  // По одному файлу на унікальний ключ групи.
  var seen = {}, picked = [];
  for (var i = 1; i < data.length; i++){
    var id = String(data[i][2] == null ? '' : data[i][2]).trim();
    if (!id) continue;
    var key = groupKey(data[i][1]);
    if (!key || seen[key]) continue;
    seen[key] = true;
    picked.push({ key: key, loc: String(data[i][0] || '').trim(), grp: String(data[i][1] || '').trim(), id: id });
  }
  Logger.log('Унікальних груп: %s → [%s]', picked.length, picked.map(function(p){ return p.key; }).join(', '));

  var result = {}, summary = [];

  picked.forEach(function(p){
    var conv = toSheet(p.id);
    if (!conv) { summary.push(p.key + ': ПРОПУЩЕНО (конвертація)'); return; }
    try {
      var tpl = SpreadsheetApp.openById(conv.ssId).getSheets()[0];
      var rows = Math.min(120, tpl.getMaxRows());
      var aVals = tpl.getRange(1, 1, rows, 1).getValues();

      var sections = [];          // [{title, items:[]}]
      var openedNorm = {};        // norm секції → вже відкривалась?
      var curr = null, lastText = '', stopped = false;

      for (var r = 0; r < rows && !stopped; r++){
        var raw = String(aVals[r][0] == null ? '' : aVals[r][0]).trim();
        if (!raw) continue;
        var nt = norm(raw);
        if (SECTION_MAP[nt]){
          if (openedNorm[nt]){ stopped = true; break; }   // секція вдруге → сміттєвий дубль
          openedNorm[nt] = true;
          curr = { title: SECTION_MAP[nt], items: [] };
          sections.push(curr);
          lastText = '';
          continue;
        }
        if (isSkip(nt)) continue;
        if (!curr) continue;                              // критерій до першої секції — ігноруємо
        if (nt === lastText) continue;                    // послідовний дубль тексту
        lastText = nt;
        curr.items.push(raw);
      }

      // Зібрати у render-сумісний обʼєкт.
      var tplObj = {}, critCount = 0;
      sections.forEach(function(sec, si){
        var idx = si + 1;
        var items = sec.items.map(function(text, ci){
          critCount++;
          return { id: idx + '.' + (ci + 1), text: text };
        });
        tplObj[String(idx)] = { title: sec.title, items: items };
      });
      result[p.key] = tplObj;
      summary.push(p.key + ': ' + sections.length + ' секцій, ' + critCount + ' критеріїв' + (stopped ? ' (стоп на дублі)' : ''));
    } catch(e){
      Logger.log('  ✗ [%s] парсинг: %s', p.key, e && e.message || e);
      summary.push(p.key + ': ПОМИЛКА');
    } finally {
      if (conv.isTemp && conv.ssId){
        try { DriveApp.getFileById(conv.ssId).setTrashed(true); }
        catch(e){ Logger.log('  ⚠ temp %s не видалено: %s', conv.ssId, e && e.message || e); }
      }
    }
  });

  // Вивід — ОКРЕМИЙ Logger.log на кожну групу (щоб довгий JSON не обрізався).
  Logger.log('───── ШАБЛОНИ (по групах) ─────');
  Object.keys(result).forEach(function(key){
    Logger.log('=== ШАБЛОН: ' + key + ' ===');
    Logger.log(JSON.stringify(result[key], null, 2));
  });
  Logger.log('───── ПІДСУМОК ───── ' + summary.join(' | '));
  Logger.log('═══ buildDevTemplatesJS done ═══');
}

// deleteEmployee(actorId, rowNum) — soft-delete (O = today, формула P
// автоматично переробить "life-cycle" з активного на "X років Y місяців").
function deleteEmployee(actorId, rowNum){
  try {
    var actor = _getActor(actorId);
    if (!_canDeleteEmployee(actor))
      return {ok:false, error:'Permission denied (delete restricted to top management)', code:'PERM_DENIED'};
    rowNum = Number(rowNum) || 0;
    if (!rowNum) return {ok:false, error:'Invalid rowNum', code:'VALIDATION'};

    var lock = LockService.getDocumentLock();
    if (!lock.tryLock(10000))
      return {ok:false, error:'Could not acquire lock', code:'LOCK_TIMEOUT'};

    try {
      var sh          = _getHrSheet();
      var existingRow = sh.getRange(rowNum, 1, 1, HR_COLS).getValues()[0];
      var existing    = _parseEmpRow(existingRow, rowNum);
      if (!existing.last && !existing.first)
        return {ok:false, error:'Row ' + rowNum + ' is empty', code:'NOT_FOUND'};

      // O = column 15 (1-based). Date об'єкт → sheet форматнe як DD.MM.YYYY.
      sh.getRange(rowNum, 15).setValue(new Date());

      var updated = _parseEmpRow(sh.getRange(rowNum, 1, 1, HR_COLS).getValues()[0], rowNum);
      _writeHrAudit(actor, 'soft_delete', rowNum, existing, updated);
      return {ok:true, rowNum:rowNum, firedAt: updated.fired};
    } finally {
      lock.releaseLock();
    }
  } catch(e){
    return {ok:false, error: e.message || String(e)};
  }
}

// ═══════════════════════════════════════════════════════════════════════════
//  PREDMETNYKY — ВЧИТЕЛІ-ПРЕДМЕТНИКИ (v6.11, per-group fix v6.28.2)
// ═══════════════════════════════════════════════════════════════════════════
// Облік проведених занять: 1 рядок Predmetnyky_Lessons = 1 заняття.
// Норма = (region × subject × group_TYPE) на місяць — матриця НЕ змінюється.
// v6.28.2: норма ТИПУ застосовується до КОЖНОЇ реальної групи окремо.
//   Напр. Осокорки × Англійська × Find = 10. Три реальні Find-групи
//   (Find-iki Вікторія / Find-iki 2 Марія / Find-iki Маша Вовк) проводять
//   ПО 10 занять кожна → ліміт рахується по ПОВНІЙ назві групи (L.group).
// "Чомусики" — без ліміту (норма 0 у всіх клітинках = сигнал unlimited).
// Не торкаємось v6.5 (Предметники_Каталог читаємо тільки).
//
// Дві таблиці (у CRM_SHEET, auto-create):
//   • Predmetnyky_Norms    — норми по group_type + idempotent seed (12 рядків).
//   • Predmetnyky_Lessons  — журнал занять.
//
// Permission-модель:
//   view :  mgmt (всі ролі) | director/nurse/vyhovatel (own loc)
//   edit :  cfo/ceo/coo/cco (any loc) | director (own loc)
//   read-only mgmt :  cmo, hr, hr_trainer, rnd_director, legal
// Audit реюзає HR_Audit (action: 'pred_save_lesson' / 'pred_delete_lesson').

// ── Constants ────────────────────────────────────────────────────
var PRED_NORMS_TAB        = 'Predmetnyky_Norms';     // CRM_SHEET
var PRED_LESSONS_TAB      = 'Predmetnyky_Lessons';   // CRM_SHEET
var PRED_NORMS_HEADER     = ['Регіон','Предмет','miniBaby','Baby','Find','Study','Preschool'];
var PRED_LESSONS_HEADER   = ['ID','EmpKey','Location','Group','Subject','Date','CreatedAt','CreatedBy'];

// Нормалізовані назви предметів для системи норм.
var PRED_SUBJECTS         = ['Англійська','Музика','Хореограф','Логопед','Психолог','Чомусики'];
var PRED_UNLIMITED_SUBJ   = 'Чомусики';   // норму НЕ перевіряємо
var PRED_GROUP_TYPES      = ['miniBaby','Baby','Find','Study','Preschool'];
var PRED_LVIV_LOCATIONS   = ['Кругла','Бігова'];   // все інше → 'Київ'

// Маппінг "сирих" назв з Предметники_Каталог → PRED_SUBJECTS.
// null → не в системі норм (frontend ховає).
var CATALOG_TO_NORM_MAP = {
  'Англійська мова':       'Англійська',
  'Логопед':               'Логопед',
  'Муз.керівник':          'Музика',
  'Хореограф':             'Хореограф',
  'Психолог':              'Психолог',
  'Чомусики':              'Чомусики',
  'Підготовка до школи':   null
};

var PRED_EDIT_ROLES_ANY = ['cfo','ceo','coo','cco'];   // будь-яка локація
// director (EMP_DIR_ROLES) — тільки своя локація

// Seed для Predmetnyky_Norms (12 рядків — 2 регіони × 6 предметів).
var PRED_NORMS_SEED = [
  ['Київ',  'Англійська', 8, 8, 10, 15, 15],
  ['Київ',  'Музика',     6, 6,  6,  6,  6],
  ['Київ',  'Хореограф',  6, 6,  6,  6,  6],
  ['Київ',  'Логопед',    0, 4,  4,  4,  4],
  ['Київ',  'Психолог',   0, 2,  2,  2,  2],
  ['Київ',  'Чомусики',   0, 0,  0,  0,  0],   // unlimited
  ['Львів', 'Англійська', 8, 8, 10, 15, 15],
  ['Львів', 'Музика',     6, 8,  8,  8,  8],
  ['Львів', 'Хореограф',  6, 6,  6,  6,  6],
  ['Львів', 'Логопед',    0, 4,  4,  4,  4],
  ['Львів', 'Психолог',   0, 0,  0,  0,  0],
  ['Львів', 'Чомусики',   0, 0,  0,  0,  0]    // unlimited
];

// ── Normalizers ──────────────────────────────────────────────────
// HR.D (назва групи) → один з PRED_GROUP_TYPES або null.
// Regexes case-insensitive і толерантні до "-ki"/" ki"/empty suffix —
// HR-дані бувають з різними написаннями.
function _normalizeGroupType(grp){
  var s = String(grp || '').trim();
  if (!s) return null;
  if (/mini[\s\-]?baby/i.test(s))             return 'miniBaby';
  if (/^baby[\-\s]?ki|^baby$/i.test(s))       return 'Baby';
  if (/find[\-\s]?iki/i.test(s))              return 'Find';
  if (/study[\-\s]?ki/i.test(s))              return 'Study';
  if (/preschool|шкільн|підготов/i.test(s))   return 'Preschool';
  return null;
}

// Catalog subject АБО HR.pos → один з PRED_SUBJECTS або null.
//   1) Якщо точно у CATALOG_TO_NORM_MAP — повертаємо значення (включно з null).
//   2) Інакше — contains-match по ключових словах (для HR pos: "Викладач
//      англійської", "Музкерівник", "Логопед-психолог" тощо).
function _normalizeSubject(raw){
  var t = String(raw || '').trim();
  if (!t) return null;
  if (CATALOG_TO_NORM_MAP.hasOwnProperty(t)) return CATALOG_TO_NORM_MAP[t];
  var low = t.toLowerCase();
  if (low.indexOf('англ')    !== -1) return 'Англійська';
  if (low.indexOf('хорео')   !== -1) return 'Хореограф';
  if (low.indexOf('муз')     !== -1) return 'Музика';
  if (low.indexOf('логопед') !== -1) return 'Логопед';
  if (low.indexOf('психол')  !== -1) return 'Психолог';
  if (low.indexOf('чомус')   !== -1) return 'Чомусики';
  return null;
}

// ── Sheets (Norms — auto-create + seed) ──────────────────────────
function _getPredNormsSheet(seedIfMissing){
  var ss = getCRMSpreadsheet();
  var sh = ss.getSheetByName(PRED_NORMS_TAB);
  if (!sh){
    sh = ss.insertSheet(PRED_NORMS_TAB);
    sh.getRange(1, 1, 1, PRED_NORMS_HEADER.length).setValues([PRED_NORMS_HEADER]);
    sh.setFrozenRows(1);
    if (seedIfMissing !== false){
      sh.getRange(2, 1, PRED_NORMS_SEED.length, PRED_NORMS_HEADER.length)
        .setValues(PRED_NORMS_SEED);
    }
  }
  return sh;
}

// Ідемпотентний seed — окремий public-call, можна запускати руками з editor.
function _seedPredmetnykyNorms(){
  var sh = _getPredNormsSheet(false);
  var lastRow = sh.getLastRow();
  if (lastRow > 1) return {ok:true, skipped:true, msg:'Norms already exist ('+(lastRow-1)+' rows)'};
  sh.getRange(2, 1, PRED_NORMS_SEED.length, PRED_NORMS_HEADER.length)
    .setValues(PRED_NORMS_SEED);
  return {ok:true, seeded: PRED_NORMS_SEED.length};
}

// ── Loaders (Norms + Catalog) ────────────────────────────────────
// Norms sheet → {'Київ': {'Англійська': {miniBaby:8,...}, ...}, 'Львів': {...}}
// Норма — по group_TYPE (matriця не змінюється).
function _predNormsUseOld(targetYmd){
  var d = targetYmd || Utilities.formatDate(new Date(), 'Europe/Kiev', 'yyyy-MM-dd');
  return String(d) < '2026-09-01';
}
function _predPickNorm(row, useOld, oldIdx, newIdx){
  if (!useOld) return Number(row[newIdx]) || 0;
  var raw = row[oldIdx]; var v = Number(raw);
  return (raw !== '' && raw !== null && isFinite(v)) ? v : (Number(row[newIdx]) || 0);
}
function _loadPredNorms(targetYmd){
  var sh = _getPredNormsSheet(true);
  var data = sh.getDataRange().getValues();
  var useOld = _predNormsUseOld(targetYmd);
  var out = {};
  for (var i = 1; i < data.length; i++){
    var row = data[i];
    var region  = String(row[0] || '').trim();
    var subject = String(row[1] || '').trim();
    if (!region || !subject) continue;
    if (!out[region]) out[region] = {};
    out[region][subject] = {
      miniBaby:  _predPickNorm(row, useOld, 8, 2),
      Baby:      _predPickNorm(row, useOld, 9, 3),
      Find:      _predPickNorm(row, useOld, 10, 4),
      Study:     _predPickNorm(row, useOld, 11, 5),
      Preschool: _predPickNorm(row, useOld, 12, 6)
    };
  }
  return out;
}

// Норма ТИПУ для (loc, subject, повна_назва_групи): нормалізуємо назву
// групи → group_type, дивимось матрицю по region. 0 = не передбачено
// (або тип не розпізнано). Чомусики не має ліміту (обробляється у caller).
function _getPredNormByType(loc, subject, group){
  var gt = _normalizeGroupType(group);
  if (!gt) return 0;                              // нестандартна назва → без ліміту
  var region = _predLocToRegion(loc);
  var norms = _loadPredNorms();
  return (norms[region] && norms[region][subject] && norms[region][subject][gt]) || 0;
}

// Реальні групи локацій з листа Клієнти (distinct непорожні Група).
// locFilter='' → усі локації. → {'Голосієво': ['Baby-ki', 'Find-iki', ...]}
function _loadRealGroups(locFilter){
  var out = {};
  var sh = getCRMSpreadsheet().getSheetByName(SHEET_CLIENTS);
  if (!sh || sh.getLastRow() < 2) return out;
  var data = sh.getDataRange().getValues();
  var hdrs = data[0].map(String);
  var colLoc = hdrs.indexOf('Локація');  if (colLoc < 0) colLoc = 2;
  var colGrp = hdrs.indexOf('Група');    if (colGrp < 0) colGrp = 3;
  var filter = String(locFilter || '').trim();
  var seen = {};                                   // "loc|group" → true
  for (var i = 1; i < data.length; i++){
    var loc = String(data[i][colLoc] || '').trim();
    var grp = String(data[i][colGrp] || '').trim();
    if (!loc || !grp) continue;
    if (filter && loc !== filter) continue;
    var k = loc + '|' + grp;
    if (seen[k]) continue;
    seen[k] = true;
    if (!out[loc]) out[loc] = [];
    out[loc].push(grp);
  }
  Object.keys(out).forEach(function(l){ out[l].sort(); });
  return out;
}

// Існуючий Предметники_Каталог → [{loc, subject_raw, subject_norm, rate, teacher, active}].
// read-only; не торкаємось схеми. locFilter='' → усі локації.
function _loadPredCatalog(locFilter){
  var sh = _getPredmetnyCatalogSheet(false);
  if (!sh || sh.getLastRow() < 2) return [];
  var data = sh.getRange(2, 1, sh.getLastRow() - 1, PREDMETNY_CATALOG_HEADER.length).getValues();
  var out = [];
  var filter = String(locFilter || '').trim();
  for (var i = 0; i < data.length; i++){
    var rec = _parsePredmetnyCatRow(data[i]);
    if (!rec.loc || !rec.subject) continue;
    if (filter && rec.loc !== filter) continue;
    out.push({
      loc:          rec.loc,
      subject_raw:  rec.subject,
      subject_norm: _normalizeSubject(rec.subject),   // може бути null
      rate:         rec.rate || null,
      teacher:      rec.teacher,
      active:       rec.active
    });
  }
  return out;
}

// ── Sheets (Lessons) ─────────────────────────────────────────────
function _getPredLessonsSheet(){
  var ss = getCRMSpreadsheet();
  var sh = ss.getSheetByName(PRED_LESSONS_TAB);
  if (!sh){
    sh = ss.insertSheet(PRED_LESSONS_TAB);
    sh.getRange(1, 1, 1, PRED_LESSONS_HEADER.length).setValues([PRED_LESSONS_HEADER]);
    sh.setFrozenRows(1);
  }
  return sh;
}

// ── Endpoint helpers ─────────────────────────────────────────────
function _predLocToRegion(loc){
  var s = String(loc || '').trim();
  return PRED_LVIV_LOCATIONS.indexOf(s) !== -1 ? 'Львів' : 'Київ';
}

// Має співпадати з frontend mkKey(last, first, wday||hired).
function _mkEmpKey(last, first, wdayOrHired){
  function s(x){ return String(x == null ? '' : x).trim().replace(/\s+/g, '').slice(0, 25); }
  return 'e5_' + s(last) + '_' + s(first) + '_' + s(wdayOrHired);
}

// 'DD.MM.YYYY' / 'YYYY-MM-DD' / Date → {y, m}. m — 1-based. null якщо не парситься.
function _lessonYearMonth(dateInput){
  if (dateInput instanceof Date)
    return {y: dateInput.getFullYear(), m: dateInput.getMonth() + 1};
  var s = String(dateInput || '').trim();
  var dmy = /^(\d{1,2})\.(\d{1,2})\.(\d{4})$/.exec(s);
  if (dmy) return {y: Number(dmy[3]), m: Number(dmy[2])};
  var iso = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(s);
  if (iso) return {y: Number(iso[1]), m: Number(iso[2])};
  return null;
}

function _fmtLessonDate(d){
  if (d instanceof Date)
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd.MM.yyyy');
  return String(d || '').trim();
}

function _nextPredLessonId(sh){
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return 1;
  var ids = sh.getRange(2, 1, lastRow - 1, 1).getValues();
  var max = 0;
  for (var i = 0; i < ids.length; i++){
    var n = Number(ids[i][0]);
    if (n > max) max = n;
  }
  return max + 1;
}

// ── Permission helpers ───────────────────────────────────────────
function _canEditPredmetnyky(actor, targetLoc){
  if (!actor) return false;
  if (PRED_EDIT_ROLES_ANY.indexOf(_roleKey(actor.role)) !== -1) return true;
  if (_empHasDirRole(actor.role)){
    if (!targetLoc) return false;
    return String(actor.loc || '').trim() === String(targetLoc).trim();
  }
  return false;
}

// view scope: mgmt → null (всі), director/nurse/vyhovatel → own loc, інші → throw
function _predViewScope(actor){
  if (_empHasMgmtRole(actor.role)) return null;
  if (_empHasDirRole(actor.role) || _empHasViewRole(actor.role))
    return String(actor.loc || '').trim() || null;
  throw new Error('PERM_DENIED');
}

// ── Data loaders (teachers + lessons) ────────────────────────────
// HR.emps → teachers. Filter: emp.pos має зматчити PRED_SUBJECTS через _normalizeSubject.
function _loadPredTeachers(locFilter){
  var sh = _getHrSheet();
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  var data = sh.getRange(2, 1, lastRow - 1, HR_COLS).getValues();
  var teachers = [];
  for (var i = 0; i < data.length; i++){
    var emp = _parseEmpRow(data[i], i + 2);
    if (emp.archived) continue;
    if (!emp.last && !emp.first) continue;
    if (locFilter && emp.loc !== locFilter) continue;
    var subj = _normalizeSubject(emp.pos);
    if (!subj) continue;
    teachers.push({
      empId:     emp.rowNum,
      empKey:    _mkEmpKey(emp.last, emp.first, emp.wday || emp.hired),
      name:      (emp.last + ' ' + emp.first).trim(),
      position:  emp.pos,
      subject:   subj,             // нормалізований (один з PRED_SUBJECTS)
      locations: [emp.loc],
      phone:     emp.phone || '',
      email:     emp.email || ''
    });
  }
  return teachers;
}

// Lessons sheet → [{id, empKey, loc, group, subject, date}].
function _loadPredLessons(locFilter){
  var sh = _getPredLessonsSheet();
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  var data = sh.getRange(2, 1, lastRow - 1, PRED_LESSONS_HEADER.length).getValues();
  var out = [];
  for (var i = 0; i < data.length; i++){
    var row = data[i];
    var id  = Number(row[0]);
    if (!id) continue;
    var loc = String(row[2] || '').trim();
    if (locFilter && loc !== locFilter) continue;
    out.push({
      id:      id,
      empKey:  String(row[1] || '').trim(),
      loc:     loc,
      group:   String(row[3] || '').trim(),
      subject: String(row[4] || '').trim(),
      date:    _fmtLessonDate(row[5])
    });
  }
  return out;
}

// ── PUBLIC API ───────────────────────────────────────────────────

// GET ?action=getPredmetnyky&actorId=N
// {ok, teachers, norms (по group_type, всі), groups (реальні групи локацій,
//  scoped), catalog (scoped), lessons (scoped), assignments, scope}
function getPredmetnyky(actorId){
  try {
    var actor = _getActor(actorId);
    var scope;
    try { scope = _predViewScope(actor); }
    catch(e){ return {ok:false, error:'Permission denied', code:'PERM_DENIED'}; }

    return {
      ok:           true,
      teachers:     _loadPredTeachers(scope),
      norms:        _loadPredNorms(),               // матриця по group_type
      groups:       _loadRealGroups(scope),         // реальні групи локацій
      catalog:      _loadPredCatalog(scope),
      lessons:      _loadPredLessons(scope),
      assignments:  _loadPredAssignments(scope),
      predMerges:   _loadPredMergesList(scope),       // v7.20: обʼєднання груп (для сітки)
      scope:        scope || 'all'
    };
  } catch(e){
    return {ok:false, error: e.message || String(e)};
  }
}

// POST {action:'savePredmetnykyLesson', actorId, lesson:{empKey, location, group, subject, date}}
// Success:  {ok:true, id, current, norm}  (norm=0 → ліміту немає / unlimited)
// Errors:   {ok:false, code:'BAD_SUBJECT'|'BAD_GROUP'|'NORM_REACHED'|'PERM_DENIED'|..., error, ...}
// v6.28.2: ліміт рахується по ПОВНІЙ назві групи (кожна реальна група веде
//   свою норму свого ТИПУ). Норма береться з матриці по group_type. BAD_GROUP
//   тепер означає "групи немає в локації" (а не "тип не розпізнано").
function savePredmetnykyLesson(actorId, lesson){
  try {
    var actor = _getActor(actorId);
    lesson = lesson || {};

    // ── валідація ──
    var empKey   = String(lesson.empKey   || '').trim();
    var location = String(lesson.location || '').trim();
    var group    = String(lesson.group    || '').trim();
    var subject  = String(lesson.subject  || '').trim();
    var dateStr  = String(lesson.date     || '').trim();
    if (!empKey)   return {ok:false, error:'empKey is required'};
    if (!location) return {ok:false, error:'location is required'};
    if (!group)    return {ok:false, error:'group is required'};
    if (!subject)  return {ok:false, error:'subject is required'};
    if (!dateStr)  return {ok:false, error:'date is required'};
    if (PRED_SUBJECTS.indexOf(subject) === -1)
      return {ok:false, code:'BAD_SUBJECT', error:'Unknown subject: ' + subject};
    var ym = _lessonYearMonth(dateStr);
    if (!ym) return {ok:false, error:'Bad date format: ' + dateStr};

    // ── валідація групи: чи існує така група в локації (lenient) ──
    // Реальні групи беремо з Клієнтів. Якщо для локації список груп НЕ
    // порожній і групи в ньому немає — відмова. Якщо груп нема взагалі
    // (порожні Клієнти) — пропускаємо перевірку, не блокуємо.
    var realGroups = (_loadRealGroups(location)[location]) || [];
    if (realGroups.length && realGroups.indexOf(group) === -1)
      return {ok:false, code:'BAD_GROUP',
              error:'Групи "' + group + '" немає в локації ' + location, group:group};

    // ── permission ──
    if (!_canEditPredmetnyky(actor, location))
      return {ok:false, code:'PERM_DENIED', error:'Permission denied'};

    var lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      // ── norm check (норма по ТИПУ, ліміт по ПОВНІЙ назві групи) ──
      // norm > 0 → enforced місячний ліміт по (location, ПОВНА група, subject).
      // Чомусики / нестандартна назва / тип з нормою 0 → ліміту немає.
      var current = 0;
      var norm = (subject === PRED_UNLIMITED_SUBJ)
        ? 0
        : _getPredNormByType(location, subject, group);
      // v6.57: мінімальний місячний ліміт 15 занять для дозволених типів (norm>0) —
      // дзеркало getNormFor у predmetnyky.html. 0 = заборонено (не піднімаємо).
      if (norm > 0) norm = Math.max(norm, 15);
      if (norm > 0){
        var existing = _loadPredLessons(location);
        for (var i = 0; i < existing.length; i++){
          var L = existing[i];
          if (L.group   !== group)   continue;        // ← ПОВНА назва групи
          if (L.subject !== subject) continue;
          var lym = _lessonYearMonth(L.date);
          if (!lym || lym.y !== ym.y || lym.m !== ym.m) continue;
          current++;
        }
        if (current >= norm){
          return {ok:false, code:'NORM_REACHED', error:'norm_reached',
                  current:current, norm:norm, group:group, subject:subject};
        }
      }

      // ── insert ──
      var sh = _getPredLessonsSheet();
      var id = _nextPredLessonId(sh);
      var dateVal = _parseDateInput(dateStr);
      sh.appendRow([
        id,
        empKey,
        location,
        group,
        subject,
        dateVal instanceof Date ? dateVal : dateStr,
        new Date(),
        actor.id
      ]);
      _writeHrAudit(actor, 'pred_save_lesson', id, null,
                    {empKey:empKey, location:location, group:group, subject:subject, date:dateStr});
      return {ok:true, id:id, current: current + 1, norm: norm};
    } finally {
      lock.releaseLock();
    }
  } catch(e){
    return {ok:false, error: e.message || String(e)};
  }
}

// POST {action:'deletePredmetnykyLesson', actorId, lessonId}
function deletePredmetnykyLesson(actorId, lessonId){
  try {
    var actor = _getActor(actorId);
    var id = Number(lessonId);
    if (!id) return {ok:false, error:'lessonId is required'};

    var lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      var sh = _getPredLessonsSheet();
      var lastRow = sh.getLastRow();
      if (lastRow < 2) return {ok:false, code:'NOT_FOUND', error:'Lesson not found'};
      var data = sh.getRange(2, 1, lastRow - 1, PRED_LESSONS_HEADER.length).getValues();
      for (var i = 0; i < data.length; i++){
        if (Number(data[i][0]) !== id) continue;
        var rowNum = i + 2;
        var location = String(data[i][2] || '').trim();
        if (!_canEditPredmetnyky(actor, location))
          return {ok:false, code:'PERM_DENIED', error:'Permission denied'};
        var before = {
          empKey:   String(data[i][1] || '').trim(),
          location: location,
          group:    String(data[i][3] || '').trim(),
          subject:  String(data[i][4] || '').trim(),
          date:     _fmtLessonDate(data[i][5])
        };
        sh.deleteRow(rowNum);
        _writeHrAudit(actor, 'pred_delete_lesson', id, before, null);
        return {ok:true};
      }
      return {ok:false, code:'NOT_FOUND', error:'Lesson not found'};
    } finally {
      lock.releaseLock();
    }
  } catch(e){
    return {ok:false, error: e.message || String(e)};
  }
}

// ═══════════════════════════════════════════════════════════════════
// PREDMETNYKY_ASSIGNMENTS (v6.11.4) — призначення викладача на
// (Location, Group, Subject). Один запис = один блок у frontend.
// Sheet auto-створюється + idempotent seed з Предметники_Каталог.
// ═══════════════════════════════════════════════════════════════════
var PRED_ASSIGN_TAB    = 'Predmetnyky_Assignments';        // CRM_SHEET
// v6.11.9: схема — один викладач per (loc, subject), без колонки Group.
var PRED_ASSIGN_HEADER = ['ID','Location','Subject','EmpKey','CreatedAt','CreatedBy'];

var PRED_GROUP_NAME_BY_TYPE = {
  miniBaby:'miniBaby-ki', Baby:'Baby-ki', Find:'Find-iki',
  Study:'Study-ki', Preschool:'Preschool'
};

function _getPredAssignSheet(){
  var ss = getCRMSpreadsheet();
  var sh = ss.getSheetByName(PRED_ASSIGN_TAB);
  if (!sh){
    sh = ss.insertSheet(PRED_ASSIGN_TAB);
    sh.getRange(1, 1, 1, PRED_ASSIGN_HEADER.length).setValues([PRED_ASSIGN_HEADER]);
    sh.setFrozenRows(1);
    return sh;
  }
  // Detect old schema (з колонкою Group) → migrate in-place.
  var headerRow = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 7)).getValues()[0];
  if (String(headerRow[2] || '').trim().toLowerCase() === 'group'){
    _migratePredAssignSchema(sh);
  }
  return sh;
}

// Конвертує старий PRED_ASSIGN (7 колонок з Group) у новий (6 без Group).
// Дедуп: на (loc, subject) лишаємо ПЕРШИЙ empKey (за rowNum).
function _migratePredAssignSchema(sh){
  var lastRow = sh.getLastRow();
  var oldData = lastRow > 1
    ? sh.getRange(2, 1, lastRow - 1, 7).getValues()
    : [];
  var seen = {};                                   // "loc|subject" → true
  var newRows = [];
  var nextId = 1;
  for (var i = 0; i < oldData.length; i++){
    var loc     = String(oldData[i][1] || '').trim();
    var subject = String(oldData[i][3] || '').trim();
    var empKey  = String(oldData[i][4] || '').trim();
    var created = oldData[i][5] || new Date();
    var by      = String(oldData[i][6] || '').trim() || 'migrated';
    if (!loc || !subject || !empKey) continue;
    var k = loc + '|' + subject;
    if (seen[k]) continue;
    seen[k] = true;
    newRows.push([nextId++, loc, subject, empKey, created, by]);
  }
  // Rewrite sheet — clear old data + put new header
  sh.clearContents();
  sh.getRange(1, 1, 1, PRED_ASSIGN_HEADER.length).setValues([PRED_ASSIGN_HEADER]);
  if (newRows.length){
    sh.getRange(2, 1, newRows.length, PRED_ASSIGN_HEADER.length).setValues(newRows);
  }
  sh.setFrozenRows(1);
  Logger.log('[migratePredAssignSchema] old rows=%s → new rows=%s (dedup by loc+subject)',
    oldData.length, newRows.length);
}

function _nextPredAssignId(sh){
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return 1;
  var ids = sh.getRange(2, 1, lastRow - 1, 1).getValues();
  var max = 0;
  for (var i = 0; i < ids.length; i++){
    var n = Number(ids[i][0]);
    if (n > max) max = n;
  }
  return max + 1;
}

// Assignments sheet → [{id, loc, subject, empKey}].
// Один викладач per (loc, subject) (v6.11.9).
// Auto-seeds (католог → HR fallback) при першому виклику на порожньому листі.
function _loadPredAssignments(locFilter){
  var sh = _getPredAssignSheet();
  if (sh.getLastRow() < 2){
    try {
      var catRes = _seedPredmetnykyAssignmentsFromCatalog();
      Logger.log('[loadPredAssignments] catalog-seed: %s', JSON.stringify(catRes));
      var hrRes = _seedPredmetnykyAssignmentsFromHR();
      Logger.log('[loadPredAssignments] hr-seed: %s', JSON.stringify(hrRes));
    } catch(e){
      Logger.log('[loadPredAssignments] seed failed: %s', e && e.message);
    }
  }
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  var data = sh.getRange(2, 1, lastRow - 1, PRED_ASSIGN_HEADER.length).getValues();
  var out = [];
  var filter = String(locFilter || '').trim();
  for (var i = 0; i < data.length; i++){
    var id  = Number(data[i][0]); if (!id) continue;
    var loc = String(data[i][1] || '').trim();
    if (filter && loc !== filter) continue;
    out.push({
      id:      id,
      loc:     loc,
      subject: String(data[i][2] || '').trim(),
      empKey:  String(data[i][3] || '').trim()
    });
  }
  return out;
}

// Нормалізує ПІБ для матчингу catalog.teacher → HR.name.
// Whitespace + lowercase. (Без strip-diacritics — українська збігається.)
function _predNormalizeName(name){
  return String(name || '').trim().toLowerCase().replace(/\s+/g, ' ');
}

// Ідемпотентний seed assignments з Предметники_Каталог.
// Один assignment per (loc, subject_norm). Створюємо тільки якщо
// принаймні одна group має norm>0 (або subject = Чомусики).
// Матчинг catalog.teacher → HR.empKey за нормалізованим ПІБ.
function _seedPredmetnykyAssignmentsFromCatalog(){
  var sh = _getPredAssignSheet();
  if (sh.getLastRow() > 1) return {ok:true, skipped:true, msg:'Assignments already populated'};

  var catalog    = _loadPredCatalog('');
  var norms      = _loadPredNorms();
  var hrTeachers = _loadPredTeachers(null);

  // ПІБ-індекс (включаючи альтернативу "First Last")
  var byName = {};
  hrTeachers.forEach(function(t){
    var k1 = _predNormalizeName(t.name);
    if (k1 && !byName[k1]) byName[k1] = t.empKey;
    var parts = String(t.name || '').trim().split(/\s+/);
    if (parts.length >= 2){
      var k2 = _predNormalizeName(parts[1] + ' ' + parts[0]);
      if (k2 && !byName[k2]) byName[k2] = t.empKey;
    }
  });

  var now = new Date();
  var rows = [];
  var nextId = 1;
  var seen = {};                                       // "loc|subject" → true (dedup)
  var stats = {catalog:0, withTeacher:0, matched:0, unmatched:[], generated:0};

  catalog.forEach(function(a){
    stats.catalog++;
    if (!a.active)       return;
    if (!a.teacher)      return;
    if (!a.subject_norm) return;
    stats.withTeacher++;

    var key = a.loc + '|' + a.subject_norm;
    if (seen[key]) return;                             // вже додано для цієї пари

    var empKey = byName[_predNormalizeName(a.teacher)] || null;
    if (!empKey){
      stats.unmatched.push(a.loc + '/' + a.subject_raw + '/«' + a.teacher + '»');
      return;
    }

    // Перевірка: хоча б одна group має norm>0 (або subject unlimited)
    var region = _predLocToRegion(a.loc);
    var isUnlimited = (a.subject_norm === PRED_UNLIMITED_SUBJ);
    var hasNorm = false;
    if (isUnlimited){
      hasNorm = true;
    } else {
      for (var gi = 0; gi < PRED_GROUP_TYPES.length; gi++){
        var n = (norms[region] && norms[region][a.subject_norm] &&
                 norms[region][a.subject_norm][PRED_GROUP_TYPES[gi]]) || 0;
        if (n > 0){ hasNorm = true; break; }
      }
    }
    if (!hasNorm) return;

    seen[key] = true;
    stats.matched++;
    rows.push([nextId++, a.loc, a.subject_norm, empKey, now, 'seed']);
    stats.generated++;
  });

  if (rows.length){
    sh.getRange(2, 1, rows.length, PRED_ASSIGN_HEADER.length).setValues(rows);
  }
  Logger.log('[seedAssignments] cat=%s teachersInCat=%s matched=%s generated=%s unmatched=%s',
    stats.catalog, stats.withTeacher, stats.matched, stats.generated, stats.unmatched.length);
  if (stats.unmatched.length){
    Logger.log('[seedAssignments] unmatched: %s', stats.unmatched.join(' | '));
  }
  return {ok:true, seeded:rows.length, stats:stats};
}

// Idempotent fallback seed: для кожної пари (loc, subject) де в HR
// РІВНО ОДИН викладач — авто-створюємо ОДИН assignment (per loc+subject).
// Існуючі assignments не чіпаємо. Multi-candidate випадки лишаємо для
// ручного вибору через UI ✏️.
function _seedPredmetnykyAssignmentsFromHR(){
  var sh = _getPredAssignSheet();

  // map існуючих assignments: "loc|subject" → true
  var existing = {};
  if (sh.getLastRow() > 1){
    var data = sh.getRange(2, 1, sh.getLastRow() - 1, PRED_ASSIGN_HEADER.length).getValues();
    for (var i = 0; i < data.length; i++){
      var k = String(data[i][1]||'').trim() + '|' + String(data[i][2]||'').trim();
      existing[k] = true;
    }
  }

  var hrTeachers = _loadPredTeachers(null);
  var norms = _loadPredNorms();

  // Групуємо HR-викладачів по (loc, subject)
  var byLocSubj = {};
  hrTeachers.forEach(function(t){
    var loc = t.locations && t.locations[0];
    if (!loc || !t.subject) return;
    var key = loc + '|' + t.subject;
    if (!byLocSubj[key]) byLocSubj[key] = [];
    byLocSubj[key].push(t);
  });

  var now = new Date();
  var nextId = _nextPredAssignId(sh);
  var rowsToAppend = [];
  var stats = {pairs:0, single:0, multi:0, skippedExisting:0, generated:0,
               multiList:[]};

  Object.keys(byLocSubj).forEach(function(key){
    stats.pairs++;
    var list = byLocSubj[key];
    if (list.length !== 1){
      stats.multi++;
      stats.multiList.push(key + ' [' + list.length + ' candidates]');
      return;
    }
    stats.single++;
    var t = list[0];
    var loc = t.locations[0];
    var subj = t.subject;
    if (existing[key]){ stats.skippedExisting++; return; }

    // Перевірка: хоча б одна group має norm>0 (або subject unlimited)
    var region = _predLocToRegion(loc);
    var isUnlimited = (subj === PRED_UNLIMITED_SUBJ);
    var hasNorm = isUnlimited;
    if (!hasNorm){
      for (var gi = 0; gi < PRED_GROUP_TYPES.length; gi++){
        var n = (norms[region] && norms[region][subj] &&
                 norms[region][subj][PRED_GROUP_TYPES[gi]]) || 0;
        if (n > 0){ hasNorm = true; break; }
      }
    }
    if (!hasNorm) return;

    rowsToAppend.push([nextId++, loc, subj, t.empKey, now, 'hr-seed']);
    existing[key] = true;
    stats.generated++;
  });

  if (rowsToAppend.length){
    var startRow = sh.getLastRow() + 1;
    sh.getRange(startRow, 1, rowsToAppend.length, PRED_ASSIGN_HEADER.length)
      .setValues(rowsToAppend);
  }
  Logger.log('[seedFromHR] pairs=%s single=%s multi=%s skippedExisting=%s generated=%s',
    stats.pairs, stats.single, stats.multi, stats.skippedExisting, stats.generated);
  if (stats.multiList.length){
    Logger.log('[seedFromHR] multi-candidate pairs (manual assign needed): %s',
      stats.multiList.join(' | '));
  }
  return {ok:true, seeded:rowsToAppend.length, stats:stats};
}

// Upsert ставки у Предметники_Каталог для (loc, subject_norm).
// Якщо catalog row існує (active, _normalizeSubject(row.subject) === subject_norm)
// — оновлюємо rate; інакше — створюємо новий рядок з subject_raw = subject_norm.
// Повертає {ok, id, created} або {ok:false, error}.
function _upsertCatalogRate(loc, subject_norm, rate, actor){
  var rateNum = Number(rate);
  if (!isFinite(rateNum) || rateNum < 0) return {ok:false, error:'Invalid rate'};

  var sh = _getPredmetnyCatalogSheet(true);
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++){
    var rec = _parsePredmetnyCatRow(data[i]);
    if (rec.loc !== loc) continue;
    if (_normalizeSubject(rec.subject) !== subject_norm) continue;
    var oldRate = Number(rec.rate) || 0;
    sh.getRange(i + 1, 4).setValue(rateNum);                     // col D = Ставка
    if (!rec.active){ sh.getRange(i + 1, 6).setValue(true); }    // re-activate якщо було off
    _writeHrAudit(actor, 'pred_catalog_rate', rec.id,
      {rate:oldRate, active:rec.active},
      {rate:rateNum, loc:loc, subject:rec.subject});
    return {ok:true, id:rec.id, updated:true, oldRate:oldRate, newRate:rateNum};
  }
  // Не знайдено — створюємо новий catalog row.
  var newId = _nextPredmetnyRowId(sh);
  sh.appendRow([newId, loc, subject_norm, rateNum, '', true]);
  _writeHrAudit(actor, 'pred_catalog_rate', newId, null,
    {loc:loc, subject:subject_norm, rate:rateNum, created:true});
  return {ok:true, id:newId, created:true, newRate:rateNum};
}

// POST {action:'savePredmetnykyAssignment', actorId,
//        payload:{loc, subject, empKey?, rate?}}
// Принаймні одне з empKey або rate має бути присутнє.
// empKey: upsert assignment row by (loc, subject).
// rate:   upsert catalog rate by (loc, subject_norm).
function savePredmetnykyAssignment(actorId, payload){
  try {
    var actor = _getActor(actorId);
    payload = payload || {};
    var loc     = String(payload.loc     || payload.location || '').trim();
    var subject = String(payload.subject || '').trim();
    var empKey  = String(payload.empKey  || '').trim();
    var hasRate = payload.rate !== undefined && payload.rate !== null && payload.rate !== '';
    var rate    = hasRate ? Number(payload.rate) : null;

    if (!loc)     return {ok:false, error:'loc required'};
    if (!subject) return {ok:false, error:'subject required'};
    if (!empKey && !hasRate)
      return {ok:false, error:'Provide empKey or rate (or both)'};

    if (PRED_SUBJECTS.indexOf(subject) === -1)
      return {ok:false, code:'BAD_SUBJECT', error:'Unknown subject: ' + subject};
    if (hasRate && (!isFinite(rate) || rate < 0))
      return {ok:false, error:'Invalid rate: ' + payload.rate};

    if (!_canEditPredmetnyky(actor, loc))
      return {ok:false, code:'PERM_DENIED', error:'Permission denied'};

    var lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      var result = {ok:true, loc:loc, subject:subject};

      // ── 1. Catalog rate (якщо передано) ──
      if (hasRate){
        var catRes = _upsertCatalogRate(loc, subject, rate, actor);
        if (!catRes.ok) return {ok:false, error:'Catalog: ' + catRes.error};
        result.rate = rate;
        result.catalogId = catRes.id;
        result.catalogCreated = !!catRes.created;
      }

      // ── 2. Assignment (якщо empKey передано) ──
      if (empKey){
        var sh = _getPredAssignSheet();
        var lastRow = sh.getLastRow();
        var data = lastRow > 1
          ? sh.getRange(2, 1, lastRow - 1, PRED_ASSIGN_HEADER.length).getValues()
          : [];

        var foundExisting = false;
        for (var i = 0; i < data.length; i++){
          if (String(data[i][1] || '').trim() !== loc)     continue;
          if (String(data[i][2] || '').trim() !== subject) continue;
          var rowNum = i + 2;
          var existingId = Number(data[i][0]) || 0;
          var beforeEmpKey = String(data[i][3] || '').trim();
          if (beforeEmpKey === empKey){
            result.id = existingId;
            result.empKey = empKey;
            result.updated = false;
          } else {
            sh.getRange(rowNum, 4).setValue(empKey);
            _writeHrAudit(actor, 'pred_save_assign', existingId,
              {empKey:beforeEmpKey},
              {empKey:empKey, loc:loc, subject:subject});
            result.id = existingId;
            result.empKey = empKey;
            result.updated = true;
          }
          foundExisting = true;
          break;
        }
        if (!foundExisting){
          var newId = _nextPredAssignId(sh);
          sh.appendRow([newId, loc, subject, empKey, new Date(), actor.id]);
          _writeHrAudit(actor, 'pred_save_assign', newId, null,
            {empKey:empKey, loc:loc, subject:subject});
          result.id = newId;
          result.empKey = empKey;
          result.created = true;
        }
      }

      return result;
    } finally {
      lock.releaseLock();
    }
  } catch(e){
    return {ok:false, error: e.message || String(e)};
  }
}

// POST {action:'deletePredmetnykyAssignment', actorId, id}
function deletePredmetnykyAssignment(actorId, id){
  try {
    var actor = _getActor(actorId);
    id = Number(id);
    if (!id) return {ok:false, error:'id required'};

    var lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      var sh = _getPredAssignSheet();
      var lastRow = sh.getLastRow();
      if (lastRow < 2) return {ok:false, code:'NOT_FOUND', error:'Assignment not found'};
      var data = sh.getRange(2, 1, lastRow - 1, PRED_ASSIGN_HEADER.length).getValues();
      for (var i = 0; i < data.length; i++){
        if (Number(data[i][0]) !== id) continue;
        var loc = String(data[i][1] || '').trim();
        if (!_canEditPredmetnyky(actor, loc))
          return {ok:false, code:'PERM_DENIED', error:'Permission denied'};
        var before = {
          loc:     loc,
          subject: String(data[i][2] || '').trim(),
          empKey:  String(data[i][3] || '').trim()
        };
        sh.deleteRow(i + 2);
        _writeHrAudit(actor, 'pred_delete_assign', id, before, null);
        return {ok:true, id:id};
      }
      return {ok:false, code:'NOT_FOUND', error:'Assignment not found'};
    } finally {
      lock.releaseLock();
    }
  } catch(e){
    return {ok:false, error: e.message || String(e)};
  }
}

// POST {action:'clearAllPredmetnykyLessons', actorId, location}
// DESTRUCTIVE: видаляє ВСІ заняття для конкретної локації (по всіх
// місяцях/групах/предметах). Корисно для тестових скидань.
// Permission: cfo/ceo/coo/cco або director у власній локації.
function clearAllPredmetnykyLessons(actorId, location){
  try {
    var actor = _getActor(actorId);
    var loc = String(location || '').trim();
    if (!loc) return {ok:false, error:'location required'};
    if (!_canEditPredmetnyky(actor, loc))
      return {ok:false, code:'PERM_DENIED', error:'Permission denied'};

    var lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      var sh = _getPredLessonsSheet();
      var lastRow = sh.getLastRow();
      if (lastRow < 2) return {ok:true, deleted:0, loc:loc};

      var data = sh.getRange(2, 1, lastRow - 1, PRED_LESSONS_HEADER.length).getValues();
      var rowsToDelete = [];
      var deletedIds = [];
      for (var i = 0; i < data.length; i++){
        if (String(data[i][2] || '').trim() !== loc) continue;
        rowsToDelete.push(i + 2);
        deletedIds.push(Number(data[i][0]) || 0);
      }
      // Видалення з низу — щоб індекси не зсувались.
      rowsToDelete.sort(function(a, b){ return b - a; });
      for (var j = 0; j < rowsToDelete.length; j++){
        sh.deleteRow(rowsToDelete[j]);
      }
      _writeHrAudit(actor, 'pred_clear_lessons', 0,
        {loc:loc, count:deletedIds.length}, null);
      Logger.log('[clearAllPredmetnykyLessons] loc="%s" deleted=%s actor=%s',
        loc, rowsToDelete.length, actor.id);
      return {ok:true, deleted:rowsToDelete.length, loc:loc, ids:deletedIds};
    } finally {
      lock.releaseLock();
    }
  } catch(e){
    return {ok:false, error: e.message || String(e)};
  }
}

// Apps Script editor wrapper: _clearAllPredmetnykyLessons('Голосієво')
// Викликає clearAllPredmetnykyLessons як CFO (actorId=1).
function _clearAllPredmetnykyLessons(location){
  var res = clearAllPredmetnykyLessons(1, location);
  Logger.log('RESULT: ' + JSON.stringify(res));
  return res;
}

// ═══════════════════════════════════════════════════════════════════
// Salary export з Predmetnyky_Lessons (v6.11.4). Алгоритм такий
// як v6.5 exportPredmetnyToSalary, але джерело — лист lessons:
//   1) Кожен active catalog (loc, subject, rate) → count = унікальні
//      (group, date) з PRED_LESSONS_TAB у (loc, subject_norm, year, month).
//   2) fact = count × rate.
//   3) Salary row matching по "<subject_raw> <rate>" (P1-P7).
//   4) Дельта через Експорт_Журнал (kind='predmetnyky').
// Target month — наступний (зарплата за лекції місяця N виплачується N+1).
// ═══════════════════════════════════════════════════════════════════
function exportPredmetnykyToSalary(params){
  params = params || {};

  // Batch mode: locations: [...] → виклик по кожній локації окремо.
  if (Array.isArray(params.locations)){
    Logger.log('[exportPredmetnykyToSalary] BATCH locs=%s month=%s/%s',
      params.locations.length, params.month, params.year);
    var batchResults = [], totFact = 0, totUpd = 0, totCells = 0;
    for (var li = 0; li < params.locations.length; li++){
      var sub = exportPredmetnykyToSalary({
        actorId: params.actorId,
        loc:     params.locations[li],
        year:    params.year,
        month:   params.month
      });
      batchResults.push(sub);
      if (sub && sub.ok){
        totFact  += (sub.totalFact   || 0);
        totUpd   += (sub.updated     || 0);
        totCells += (sub.cellsWritten|| 0);
      }
    }
    return {ok:true, batch:true, locations: batchResults,
            totalFact: totFact, totalUpdated: totUpd, totalCellsWritten: totCells};
  }

  try {
    var actorId = Number(params.actorId || 0);
    var loc     = String(params.loc || '').trim();
    var month   = Number(params.month);
    var year    = Number(params.year) || new Date().getFullYear();

    if (!loc) return {ok:false, error:'loc обовʼязковий'};
    if (!month || month < 1 || month > 12) return {ok:false, error:'month має бути 1-12'};

    if (actorId){
      var actor;
      try { actor = _getActor(actorId); } catch(e){
        return {ok:false, code:'PERM_DENIED', error:'Actor not found'};
      }
      if (!_canEditPredmetnyky(actor, loc))
        return {ok:false, code:'PERM_DENIED', error:'Permission denied'};
    }

    Logger.log('[exportPredmetnykyToSalary] START loc="%s" month=%s year=%s', loc, month, year);

    // 1. Catalog для локації — лише active + rate>0 + subject_norm
    var withRate = _loadPredCatalog(loc).filter(function(a){
      return a.active && a.rate > 0 && a.subject_norm;
    });
    if (!withRate.length){
      Logger.log('[exportPredmetnykyToSalary] no active rated catalog entries for ' + loc);
      return {ok:true, updated:0, totalFact:0, cellsWritten:0,
              info:'No active catalog entries with rate for ' + loc, loc:loc};
    }

    // 2. Lessons → groupsByDate per subject (session-key = група×дата).
    // v7.20: структура { subjNormKey: { 'YYYY-MM-DD': { normGroup: true } } } —
    // рахунок через _dopCountSessions (обʼєднання схлопують групи одного дня в 1).
    // subjNormKey = _dopNormGroup(subject) (стійке співставлення з subject_norm).
    var lessons = _loadPredLessons(loc);
    var gbdBySubj = {};
    for (var i = 0; i < lessons.length; i++){
      var L = lessons[i];
      var ym = _lessonYearMonth(L.date);
      if (!ym || ym.y !== year || ym.m !== month) continue;
      var sk   = _dopNormGroup(L.subject);
      var dIso = _predDateToISO(L.date);
      var ng   = _dopNormGroup(L.group);
      if (!gbdBySubj[sk]) gbdBySubj[sk] = {};
      if (!gbdBySubj[sk][dIso]) gbdBySubj[sk][dIso] = {};
      gbdBySubj[sk][dIso][ng] = true;
    }
    // Обʼєднання предметників (session-key схлопування) за цей місяць.
    var mm2 = month < 10 ? '0' + month : String(month);
    var predDateFrom = year + '-' + mm2 + '-01';
    var predNextM = _nextMonth(month, year);
    var pnmm = predNextM.month < 10 ? '0' + predNextM.month : String(predNextM.month);
    var predDateTo = predNextM.year + '-' + pnmm + '-01';
    var predMergesMap = _loadPredMergesMap(loc, predDateFrom, predDateTo);

    // 3. Salary registry → файл локації
    var reg = _salaryGetRegistry();
    if (!reg.ok) return reg;
    var entry = null;
    for (var j = 0; j < reg.rows.length; j++){
      if (reg.rows[j].loc === loc){ entry = reg.rows[j]; break; }
    }
    if (!entry) return {ok:false, error:'Локація "' + loc + '" не знайдена у Salary-реєстрі'};
    var locSS = SpreadsheetApp.openById(entry.sheetId);
    var sheet = locSS.getSheetByName(entry.listName);
    if (!sheet) return {ok:false, error:'Salary sheet "' + entry.listName + '" не знайдено'};

    var nextM = _nextMonth(month, year);
    var lastRow = Math.max(sheet.getLastRow(), 80);
    var names = sheet.getRange(1, 1, lastRow, 1).getValues();
    var targetMonth = nextM.month;
    var budgetCol = (targetMonth - 1) * 3 + 3;
    var targetMonthName = MONTHS_CAL_UA[targetMonth - 1];
    var sourceMonthName = MONTHS_CAL_UA[month - 1];

    var salaryRows = [];
    for (var k = 3; k < names.length; k++){
      var raw = String(names[k][0] == null ? '' : names[k][0]).trim();
      if (!raw) continue;
      salaryRows.push({row:k+1, raw:raw,
        norm:_journalNormName(raw), soft:_softNorm(raw)});
    }

    var budgetColValues   = sheet.getRange(1, budgetCol, lastRow, 1).getValues();
    var budgetColFormulas = sheet.getRange(1, budgetCol, lastRow, 1).getFormulas();
    var journal = _readJournalForTarget(loc, 'predmetnyky', nextM.year, nextM.month);

    var journalOps = [];
    var updated = 0, totalFact = 0, cellsWritten = 0, formulaRowsSkipped = 0;
    var p7queue = [], maxMatchedRow = 0, details = [];
    var stats = {attempts:0, p1:0, p2:0, p3:0, p4:0, p5:0, p6:0, p7:0};

    // 4. Матчинг кожного catalog entry → Salary row.
    // OVERWRITE-логіка: клітинка "<Subject> <Rate>" у Salary належить
    // ВИКЛЮЧНО predmetnyky (ніхто інший туди не пише), тому пишемо fact
    // напряму без додавання до попереднього значення. Це фіксить баг де
    // ручні дані / старий v6.5 експорт залишили "забруднену" клітинку
    // і дельта-логіка (currentValue - 0 + fact) накопичувала суму.
    withRate.forEach(function(a){
      var sk   = _dopNormGroup(a.subject_norm);
      var gbd  = gbdBySubj[sk] || {};
      // v7.20: група×дата з урахуванням обʼєднань (схлопують групи одного дня в 1).
      var uniq = _dopCountSessions(gbd, predMergesMap[sk] || {});
      var fact = uniq * a.rate;
      var catName = a.subject_raw + ' ' + a.rate;
      var nk = _journalNormName(catName);
      stats.attempts++;

      var found = _findPredmetnySalaryRow(salaryRows, a.subject_raw, a.rate);
      if (!found){
        stats.p7++;
        p7queue.push({subject:a.subject_raw, rate:a.rate, fact:fact,
                      lessons:uniq, catName:catName, nk:nk});
        Logger.log('[%s] %s → P7 (рядок не знайдено) | lessons=%s × %s = %s',
          loc, catName, uniq, a.rate, fact);
        return;
      }
      stats['p' + found.priority.slice(1)]++;
      if (found.row > maxMatchedRow) maxMatchedRow = found.row;

      var rowIdx0 = found.row - 1;
      if (budgetColFormulas[rowIdx0] && budgetColFormulas[rowIdx0][0]){
        formulaRowsSkipped++;
        Logger.log('[%s] %s → SKIP (формула у Salary row %s)', loc, catName, found.row);
        return;
      }
      var currentValue = Number(budgetColValues[rowIdx0][0]) || 0;
      var newValue    = fact;                  // ← OVERWRITE (не дельта)
      if (newValue !== currentValue){
        sheet.getRange(found.row, budgetCol).setValue(newValue);
        cellsWritten++;
      }
      // Журнал лишається для аудиту (last_written_sum + timestamp).
      var je = journal.byNormName[nk];
      var lastWritten = je ? je.sum : 0;
      if (fact !== lastWritten){
        journalOps.push({nk:nk, loc:loc, kind:'predmetnyky', name:catName,
          year:nextM.year, month:nextM.month, newSum:fact});
      }
      updated++;
      totalFact += fact;
      details.push({subject:catName, matchedAs:found.matchedAs, priority:found.priority,
        fact:fact, lessons:uniq, row:found.row, oldValue:currentValue, newValue:newValue});
      Logger.log('[%s] %s → %s | lessons=%s × %s = %s | Salary row=%s col=%s | %s → %s',
        loc, catName, found.priority, uniq, a.rate, fact,
        found.row, budgetCol, currentValue, newValue);
    });

    // 5. P7 — додаємо нові рядки після останнього зматченого
    p7queue.forEach(function(p){
      var newRow;
      if (maxMatchedRow > 0){
        sheet.insertRowsAfter(maxMatchedRow, 1);
        newRow = maxMatchedRow + 1;
        maxMatchedRow = newRow;
      } else {
        newRow = sheet.getLastRow() + 1;
      }
      sheet.getRange(newRow, 1).setValue(p.subject + ' ' + p.rate);
      sheet.getRange(newRow, budgetCol).setValue(p.fact);
      cellsWritten++;
      journalOps.push({nk:p.nk, loc:loc, kind:'predmetnyky', name:p.catName,
        year:nextM.year, month:nextM.month, newSum:p.fact});
      updated++;
      totalFact += p.fact;
      details.push({subject:p.catName, fact:p.fact, lessons:p.lessons,
        priority:'P7', row:newRow, status:'row-added'});
    });

    _commitJournalUpdates(journal, journalOps);
    Logger.log('[%s] СВОДКА: catalog=%s | P1=%s P2=%s P3=%s P4=%s P5=%s P6=%s P7=%s | клітинок=%s формул-пропущено=%s',
      loc, stats.attempts, stats.p1, stats.p2, stats.p3, stats.p4, stats.p5, stats.p6, stats.p7,
      cellsWritten, formulaRowsSkipped);

    return {
      ok: true,
      loc: loc,
      sourceMonth: sourceMonthName,
      targetMonth: targetMonthName,
      budgetCol: budgetCol,
      updated: updated,
      totalFact: totalFact,
      cellsWritten: cellsWritten,
      formulaRowsSkipped: formulaRowsSkipped,
      rowsAdded: stats.p7,
      matchStats: stats,
      details: details
    };
  } catch(e){
    Logger.log('[exportPredmetnykyToSalary] EXCEPTION: %s\n%s', e && e.message, e && e.stack);
    return {ok:false, error: String(e && e.message || e)};
  }
}

// ── TESTS (v6.11, per-group fix v6.28.2) ──────────────────────────
// Запускати з Apps Script editor → Run → _testPredmetnykyBackend.
// Actor: CFO Мельніченко Ірина (ID=1), testDate=15.01.2199 (далеке майбутнє
// щоб не колізіювати з реальними lessons). Тест НЕ чіпає норми (матриця).
// Для ліміту обирає РЕАЛЬНУ групу локації з найменшою ненульовою нормою
// (норма береться по типу). Side-effects: створює/видаляє рядки у Lessons.
function _testPredmetnykyBackend(){
  var actorId  = 1;                  // CFO Мельніченко Ірина
  var testDate = '15.01.2199';
  var createdIds = [];
  var pass = 0, fail = 0;

  Logger.log('━━━ Predmetnyky backend tests ━━━');
  Logger.log('  actorId=' + actorId + ', testDate=' + testDate);

  function ok(name, cond, info){
    if (cond){ pass++; Logger.log('  ✅ ' + name + (info ? ' ('+info+')' : '')); }
    else     { fail++; Logger.log('  ❌ ' + name + (info ? ' — '+info : '')); }
  }
  function cleanup(){
    var n = 0;
    for (var i = 0; i < createdIds.length; i++){
      try {
        var r = deletePredmetnykyLesson(actorId, createdIds[i]);
        if (r && r.ok) n++;
      } catch(e){}
    }
    Logger.log('  cleanup: removed ' + n + '/' + createdIds.length + ' test rows');
  }

  // ── 0. seed norms (idempotent) ──
  var seedRes = _seedPredmetnykyNorms();
  ok('0. seed idempotent', seedRes && seedRes.ok === true,
     seedRes && (seedRes.skipped ? 'already seeded' : 'seeded=' + seedRes.seeded));

  // ── 1. getPredmetnyky ──
  var got = getPredmetnyky(actorId);
  ok('1a. getPredmetnyky.ok', got && got.ok === true, got && got.error);
  if (!got || !got.ok){
    Logger.log('━━━ ABORT (cannot continue without getPredmetnyky) ━━━');
    return {ok:false, pass:pass, fail:fail+9};
  }
  ok('1b. teachers ≥ 5', got.teachers.length >= 5, 'got=' + got.teachers.length);
  ok('1c. norms[Київ][Англійська]',  !!(got.norms.Київ  && got.norms.Київ['Англійська']));
  ok('1d. groups is object', got.groups && typeof got.groups === 'object');
  ok('1e. lessons is array', Array.isArray(got.lessons));
  ok('1f. catalog is array', Array.isArray(got.catalog));

  // Test fixture: реальна Київ-локація (не-Lviv) з перших teachers; fallback 'Голосієво'.
  var testLoc = 'Голосієво';
  for (var t = 0; t < got.teachers.length; t++){
    var l = got.teachers[t].locations[0];
    if (l && PRED_LVIV_LOCATIONS.indexOf(l) === -1){ testLoc = l; break; }
  }
  var testEmpKey = (got.teachers[0] && got.teachers[0].empKey) || 'e5_TEST_TEST_TEST';
  var region     = _predLocToRegion(testLoc);
  var groupsHere = (got.groups && got.groups[testLoc]) || [];
  Logger.log('  fixture: loc=' + testLoc + ', region=' + region +
             ', empKey=' + testEmpKey + ', realGroups=' + groupsHere.length);

  // Обираємо РЕАЛЬНУ групу + предмет з найменшою ненульовою нормою (по типу).
  var best = null;   // {group, subject, norm}
  for (var gi = 0; gi < groupsHere.length; gi++){
    var gt = _normalizeGroupType(groupsHere[gi]);
    if (!gt) continue;
    for (var si = 0; si < PRED_SUBJECTS.length; si++){
      var subj = PRED_SUBJECTS[si];
      if (subj === PRED_UNLIMITED_SUBJ) continue;
      var nv = (got.norms[region] && got.norms[region][subj] &&
                got.norms[region][subj][gt]) || 0;
      if (nv > 0 && (!best || nv < best.norm))
        best = {group:groupsHere[gi], subject:subj, norm:nv};
    }
  }

  if (!best){
    // Нема реальних груп з ненульовою нормою → ліміт-тести пропускаємо.
    Logger.log('  ⚠ Нема реальної групи з нормою>0 у ' + testLoc + ' — тести 2-5 SKIP');
    ok('2-5. limit tests (skipped — no real group with norm)', true, 'skipped');
  } else {
    Logger.log('  limit fixture: group="' + best.group + '" subj=' + best.subject + ' norm=' + best.norm);
    var L = {empKey:testEmpKey, location:testLoc, group:best.group, subject:best.subject, date:testDate};

    // ── 2. save first lesson → current=1, norm=best.norm ──
    var r2 = savePredmetnykyLesson(actorId, L);
    ok('2a. save first lesson ok', r2 && r2.ok === true, r2 && r2.error);
    ok('2b. current=1, norm=' + best.norm, r2 && r2.current === 1 && r2.norm === best.norm,
       r2 ? ('current=' + r2.current + ', norm=' + r2.norm) : 'no response');
    if (r2 && r2.id) createdIds.push(r2.id);

    // ── 3. fill to norm ──
    var safety = 0;
    while (createdIds.length < best.norm && safety++ < 60){
      var rf = savePredmetnykyLesson(actorId, L);
      if (!rf || !rf.ok) break;
      if (rf.id) createdIds.push(rf.id);
    }
    ok('3. filled to norm (' + createdIds.length + '/' + best.norm + ')',
       createdIds.length === best.norm, 'created=' + createdIds.length);

    // ── 4. overflow → NORM_REACHED ──
    var r4 = savePredmetnykyLesson(actorId, L);
    ok('4. overflow → NORM_REACHED',
       r4 && r4.ok === false && r4.code === 'NORM_REACHED',
       r4 && (r4.code + ' current=' + r4.current + ' norm=' + r4.norm));

    // ── 5. ІНША реальна група того ж типу веде СВОЮ норму (не аґреговано) ──
    var other = null;
    for (var gj = 0; gj < groupsHere.length; gj++){
      if (groupsHere[gj] === best.group) continue;
      if (_normalizeGroupType(groupsHere[gj]) === _normalizeGroupType(best.group)){ other = groupsHere[gj]; break; }
    }
    if (other){
      var r5 = savePredmetnykyLesson(actorId,
        {empKey:testEmpKey, location:testLoc, group:other, subject:best.subject, date:testDate});
      ok('5. інша група того ж типу → current=1 (своя норма)',
         r5 && r5.ok === true && r5.current === 1 && r5.norm === best.norm,
         r5 ? ('group=' + other + ' current=' + r5.current + ' norm=' + r5.norm) : 'no response');
      if (r5 && r5.id) createdIds.push(r5.id);
    } else {
      ok('5. інша група того ж типу (skipped — нема другої)', true, 'skipped');
    }
  }

  // Група для валідаційних тестів: реальна, якщо є; інакше синтетична
  // (коли груп нема — existence-перевірка лояльна й не блокує).
  var anyGroup = groupsHere.length ? groupsHere[0] : 'Find-iki ТЕСТ';

  // ── 6. empty fields → error ──
  var r6a = savePredmetnykyLesson(actorId,
    {empKey:'', location:testLoc, group:anyGroup, subject:'Англійська', date:testDate});
  ok('6a. empty empKey → error', r6a && r6a.ok === false && /empKey/i.test(r6a.error || ''), r6a && r6a.error);
  var r6b = savePredmetnykyLesson(actorId,
    {empKey:testEmpKey, location:testLoc, group:'', subject:'Англійська', date:testDate});
  ok('6b. empty group → error', r6b && r6b.ok === false && /group/i.test(r6b.error || ''), r6b && r6b.error);
  var r6c = savePredmetnykyLesson(actorId,
    {empKey:testEmpKey, location:testLoc, group:anyGroup, subject:'', date:testDate});
  ok('6c. empty subject → error', r6c && r6c.ok === false && /subject/i.test(r6c.error || ''), r6c && r6c.error);

  // ── 7. unknown subject → BAD_SUBJECT ──
  var r7 = savePredmetnykyLesson(actorId,
    {empKey:testEmpKey, location:testLoc, group:anyGroup, subject:'НеіснуючийПредмет', date:testDate});
  ok('7. unknown subject → BAD_SUBJECT', r7 && r7.ok === false && r7.code === 'BAD_SUBJECT',
     r7 && (r7.code + ' / ' + r7.error));

  // ── 8. неіснуюча група в локації → BAD_GROUP (лише якщо є реальні групи) ──
  if (groupsHere.length){
    var r8 = savePredmetnykyLesson(actorId,
      {empKey:testEmpKey, location:testLoc, group:'НЕ-ІСНУЄ-zzz-9000', subject:'Англійська', date:testDate});
    ok('8. неіснуюча група → BAD_GROUP', r8 && r8.ok === false && r8.code === 'BAD_GROUP',
       r8 && (r8.code + ' / ' + r8.error));
  } else {
    ok('8. неіснуюча група → BAD_GROUP (skipped — нема реальних груп)', true, 'skipped');
  }

  // ── 9. Чомусики (unlimited) на реальній групі → ok без ліміту ──
  var r9 = savePredmetnykyLesson(actorId,
    {empKey:testEmpKey, location:testLoc, group:anyGroup, subject:'Чомусики', date:testDate});
  ok('9a. Чомусики → ok, norm=0', r9 && r9.ok === true && r9.norm === 0,
     r9 ? ('ok=' + r9.ok + ' norm=' + r9.norm) : 'no response');
  if (r9 && r9.id) createdIds.push(r9.id);
  var r9b = savePredmetnykyLesson(actorId,
    {empKey:testEmpKey, location:testLoc, group:anyGroup, subject:'Чомусики', date:testDate});
  ok('9b. Чомусики 2nd save still ok (unlimited)', r9b && r9b.ok === true, r9b && r9b.error);
  if (r9b && r9b.id) createdIds.push(r9b.id);

  // ── 10. delete створеного → ok ──
  if (createdIds.length){
    var idDel = createdIds[0];
    var r10 = deletePredmetnykyLesson(actorId, idDel);
    ok('10. delete created → ok', r10 && r10.ok === true,
       r10 ? ('id=' + idDel + ', ' + (r10.error || 'ok')) : 'no response');
    if (r10 && r10.ok) createdIds.shift();
  } else {
    ok('10. delete created → ok', false, 'no created lessons to delete');
  }

  // ── 11. delete фейкового id → NOT_FOUND ──
  var r11 = deletePredmetnykyLesson(actorId, 999999999);
  ok('11. delete missing → NOT_FOUND', r11 && r11.ok === false && r11.code === 'NOT_FOUND', r11 && r11.code);

  // ── cleanup ──
  cleanup();

  Logger.log('━━━ tests: ' + pass + ' pass, ' + fail + ' fail ━━━');
  return {ok: fail === 0, pass:pass, fail:fail};
}

function testExportHolosievo() {
  var result = exportPredmetnykyToSalary({
    actorId: 1,
    locations: ['Голосієво'],
    year: 2026,
    month: 6
  });
  Logger.log('RESULT: ' + JSON.stringify(result));
  return result;
}


// ІМПОРТ РОЗВИТКУ ДІТЕЙ з Excel-файлів майстер-таблиці у CRM.
//   dryRunDevImport() — ПРЕВ'Ю (нічого не пише, лише Logger.log).
//   runDevImport()    — ЗАПИС у CRM (Розвиток JSON + Здоров'я JSON, точкові комірки).
// Запускати ВРУЧНУ з редактора Apps Script. Потребує Advanced Drive Service (Drive).
function dryRunDevImport(){ return _devImportRun(false); }
function runDevImport(){ return _devImportRun(true); }

function _devImportRun(doWrite){
  var MASTER_ID = '1od1nd818xMEcszMX_WCFdciL63x4X2pSQpd6LMqGDAc';
  var SY = 2025;
  var PERIODS = { '09':SY+'-09', '01':(SY+1)+'-01', '06':(SY+1)+'-06' };
  Logger.log('=== %s DEV IMPORT === master=%s', doWrite ? 'RUN(write)' : 'DRY-RUN', MASTER_ID);

  var SECTIONS = ['Емоційний інтелект','Соціальна адаптація','Уміння вирішувати конфлікти',
    'Самостійність у прийнятті рішень','Мовленнєвий розвиток','Сенсорно-пізнавальний розвиток',
    'Ігрова діяльність','Предметно-практична діяльність','Художньо-естетичний розвиток'];
  var MEAS = [['Маса тіла','weight'],['Довжина тіла (зріст)','height'],
    ['Обхват голови','headCirc'],['Обхват грудної клітки','chestCirc']];
  var SKIP = ['Особистісно-соціальний розвиток:','Столбец 1'];

  function norm(s){ return String(s==null?'':s).toLowerCase().replace(/\s+/g,' ').trim().replace(/[:;.,\-\s]+$/,''); }
  function nkey(s){ return String(s==null?'':s).toLowerCase().replace(/ /g,' ').replace(/\s+/g,' ').trim(); }
  function nameTokens(s){
    var t=nkey(s).replace(/[()\.,_]/g,' ').replace(/\s+/g,' ').trim();
    var parts=t.split(' '), out=[];
    for(var i=0;i<parts.length;i++){ var p=parts[i]; if(p.length>=2 && !/^\d+$/.test(p)) out.push(p); }
    return out;
  }
  function editLE1(a,b){
    if(a===b) return true;
    var la=a.length, lb=b.length;
    if(Math.abs(la-lb)>1) return false;
    if(la===lb){ var d=0; for(var i=0;i<la;i++){ if(a[i]!==b[i]){ if(++d>1) return false; } } return true; }
    var sh=la<lb?a:b, lo=la<lb?b:a, i2=0,j=0,sk=0;
    while(i2<sh.length && j<lo.length){ if(sh[i2]===lo[j]){ i2++; j++; } else { if(++sk>1) return false; j++; } }
    return true;
  }
  function tokMatch(S,C){
    var usedS={},usedC={},exact=0;
    for(var i=0;i<S.length;i++){ for(var j=0;j<C.length;j++){
      if(!usedS['s'+i] && !usedC['c'+j] && S[i]===C[j]){ usedS['s'+i]=1; usedC['c'+j]=1; exact++; break; } }}
    if(exact>=2) return true;
    if(exact>=1){
      for(var i3=0;i3<S.length;i3++){ if(usedS['s'+i3])continue;
        for(var j3=0;j3<C.length;j3++){ if(usedC['c'+j3])continue;
          if(S[i3].length>=4 && C[j3].length>=4 && editLE1(S[i3],C[j3])) return true; }}
    }
    return false;
  }
  var SECTION_MAP={}; SECTIONS.forEach(function(s){ SECTION_MAP[norm(s)]=s; });
  var MEAS_MAP={}; MEAS.forEach(function(m){ MEAS_MAP[norm(m[0])]=m[1]; });
  var SKIP_SET={}; SKIP.forEach(function(s){ SKIP_SET[norm(s)]=true; });
  function isSkip(nt){ return SKIP_SET[nt] || nt.indexOf('столбец')===0; }
  function score(v){ var t=String(v==null?'':v).trim(); if(t==='*')return '+'; if(t==='-'||t==='–'||t==='—'||t==='−')return '−'; return null; }

  function toSheet(id){
    var mime; try{ mime=DriveApp.getFileById(id).getMimeType(); }catch(e){ return {err:'getMimeType: '+(e&&e.message||e)}; }
    if(mime===MimeType.GOOGLE_SHEETS) return {ssId:id,isTemp:false};
    if(typeof Drive==='undefined') return {err:'Drive advanced service unavailable'};
    try{
      var blob=DriveApp.getFileById(id).getBlob();
      var meta={name:'tmp_devimp_'+new Date().getTime(),mimeType:MimeType.GOOGLE_SHEETS};
      var created;
      if(Drive.Files&&typeof Drive.Files.create==='function') created=Drive.Files.create(meta,blob);
      else if(Drive.Files&&typeof Drive.Files.insert==='function') created=Drive.Files.insert({title:meta.name,mimeType:meta.mimeType},blob,{convert:true});
      else return {err:'Drive.Files.create/insert not found'};
      var ssId=created&&(created.id||(created.getId&&created.getId()));
      if(!ssId) return {err:'convert: no id'};
      return {ssId:ssId,isTemp:true};
    }catch(e){ return {err:'convert: '+(e&&e.message||e)}; }
  }

  var ss=getCRMSpreadsheet(), clSheet=ss.getSheetByName(SHEET_CLIENTS);
  if(!clSheet){ Logger.log('NO Clients sheet'); return {ok:false}; }
  var clVals=clSheet.getDataRange().getValues(), clHead=clVals[0];
  function colIdx(name){ for(var c=0;c<clHead.length;c++) if(String(clHead[c]).trim()===name) return c; return -1; }
  var ciName=colIdx('ПІБ дитини'), ciLoc=colIdx('Локація'), ciDev=colIdx('Розвиток (JSON)'), ciHea=colIdx("Здоров'я (JSON)");
  if(ciName<0||ciLoc<0||ciDev<0||ciHea<0){ Logger.log('cols not found: name=%s loc=%s dev=%s health=%s',ciName,ciLoc,ciDev,ciHea); return {ok:false}; }

  var refs=[], tokIndex={};
  for(var r=1;r<clVals.length;r++){
    if(!clVals[r][0]) continue;
    var nm=String(clVals[r][ciName]||'').trim(); if(!nm) continue;
    var toks=nameTokens(nm);
    var idx=refs.length; refs.push({rowIndex:r,name:nm,loc:nkey(clVals[r][ciLoc]),tokens:toks});
    var seen={};
    for(var ti=0;ti<toks.length;ti++){ var tk=toks[ti]; if(seen[tk])continue; seen[tk]=1; (tokIndex[tk]=tokIndex[tk]||[]).push(idx); }
  }

  var master,mdata;
  try{ master=SpreadsheetApp.openById(MASTER_ID); mdata=master.getSheets()[0].getDataRange().getValues(); }
  catch(e){ Logger.log('master: %s',e&&e.message||e); return {ok:false}; }

  var stat={files:0,filesSkipped:0,tabs:0,matched:0,unmatched:0,conflict:0,emptyTabs:0,critWritten:0,measWritten:0,rowsToWrite:{}};
  var unmatchedList=[], conflictList=[], matchedFuzzy=[];
  var pendingDev={}, pendingHea={};

  for(var mi=1; mi<mdata.length; mi++){
    var loc=String(mdata[mi][0]||'').trim(), grp=String(mdata[mi][1]||'').trim(), fid=String(mdata[mi][2]==null?'':mdata[mi][2]).trim();
    if(!fid) continue;
    var conv=toSheet(fid);
    if(conv.err){ stat.filesSkipped++; Logger.log('  x [%s/%s] %s',loc,grp,conv.err); continue; }
    stat.files++;
    try{
      var sheets=SpreadsheetApp.openById(conv.ssId).getSheets();
      for(var si=0; si<sheets.length; si++){
        var sh=sheets[si], pib=String(sh.getName()||'').trim(), npib=nkey(pib);
        if(npib==='піб'||npib==='пиб'||npib==='шаблон'||npib.indexOf('столбец')===0||!npib) continue;
        var lastR=sh.getLastRow(), lastC=Math.min(sh.getLastColumn(),8);
        if(lastR<2||lastC<2) continue;
        var rows=Math.min(lastR,140);
        var vals=sh.getRange(1,1,rows,lastC).getValues();
        stat.tabs++;

        var pcols={};
        for(var c=1;c<lastC;c++){ var h=norm(vals[0][c]);
          if(h.indexOf('верес')>=0) pcols[c]=PERIODS['09'];
          else if(h.indexOf('січ')>=0||h.indexOf('сiч')>=0) pcols[c]=PERIODS['01'];
          else if(h.indexOf('червен')>=0) pcols[c]=PERIODS['06']; }
        if(!Object.keys(pcols).length){ pcols[1]=PERIODS['09']; pcols[2]=PERIODS['01']; pcols[3]=PERIODS['06']; }

        var dev={}, meas={}, secIdx=0, itemIdx=0, lastText='', curOpened={}, stopped=false, nCrit=0, nMeasCells=0;
        for(var rr=0; rr<rows && !stopped; rr++){
          var raw=String(vals[rr][0]==null?'':vals[rr][0]).trim();
          if(!raw) continue;
          var nt=norm(raw);
          if(MEAS_MAP[nt]){ var field=MEAS_MAP[nt];
            for(var pc in pcols){ var mv=String(vals[rr][pc]==null?'':vals[rr][pc]).trim();
              if(mv){ var pk=pcols[pc]; (meas[pk]=meas[pk]||{})[field]=mv; nMeasCells++; } }
            continue; }
          if(SECTION_MAP[nt]){ if(curOpened[nt]){ stopped=true; break; } curOpened[nt]=true; secIdx++; itemIdx=0; lastText=''; continue; }
          if(isSkip(nt)) continue;
          if(!secIdx) continue;
          if(nt===lastText) continue;
          lastText=nt; itemIdx++;
          var critId=secIdx+'.'+itemIdx;
          for(var pc2 in pcols){ var sv=score(vals[rr][pc2]);
            if(sv){ var pk2=pcols[pc2]; (dev[pk2]=dev[pk2]||{criteria:{}}).criteria[critId]=sv; nCrit++; } }
        }

        if(nCrit===0 && nMeasCells===0){ stat.emptyTabs++; continue; }

        var stoks=nameTokens(pib), matches=[], seenC={}, exactHit=false;
        if(stoks.length>=2){
          for(var s2=0;s2<stoks.length;s2++){ var lst=tokIndex[stoks[s2]]; if(!lst)continue;
            for(var li=0;li<lst.length;li++){ var ci=lst[li]; if(seenC[ci])continue; seenC[ci]=1;
              if(tokMatch(stoks, refs[ci].tokens)){ matches.push(refs[ci]); if(nkey(refs[ci].name)===npib) exactHit=true; } } }
        }
        var pick=null, conflict=false;
        if(matches.length===1) pick=matches[0];
        else if(matches.length>1){ var byLoc=matches.filter(function(x){return x.loc===nkey(loc);});
          if(byLoc.length===1) pick=byLoc[0]; else conflict=true; }

        if(!pick){
          if(conflict){ stat.conflict++; conflictList.push(pib+' @ '+loc+' ('+grp+') cand:'+matches.length); }
          else { stat.unmatched++; unmatchedList.push(pib+' @ '+loc+' ('+grp+')'); }
          continue;
        }
        if(!exactHit) matchedFuzzy.push(pib+' -> '+pick.name+' @ '+loc);
        stat.matched++;
        stat.critWritten+=nCrit; stat.measWritten+=nMeasCells;
        stat.rowsToWrite[pick.rowIndex]=true;

        if(nCrit>0){
          var exD=pendingDev[pick.rowIndex];
          if(!exD){ try{ exD=JSON.parse(clVals[pick.rowIndex][ciDev]||'{}'); }catch(e){ exD={}; } pendingDev[pick.rowIndex]=exD; }
          for(var pk3 in dev){ if(!exD[pk3]) exD[pk3]={criteria:{}}; if(!exD[pk3].criteria) exD[pk3].criteria={};
            for(var cid in dev[pk3].criteria) exD[pk3].criteria[cid]=dev[pk3].criteria[cid];
            exD[pk3].by='Імпорт'; exD[pk3].at=formatDate(new Date()); }
        }
        if(nMeasCells>0){
          var exH=pendingHea[pick.rowIndex];
          if(!exH){ try{ exH=JSON.parse(clVals[pick.rowIndex][ciHea]||'{}'); }catch(e){ exH={}; } pendingHea[pick.rowIndex]=exH; }
          if(!exH.measurements) exH.measurements={};
          for(var pk4 in meas){ var rec=exH.measurements[pk4]||{};
            for(var f in meas[pk4]) rec[f]=meas[pk4][f]; rec.by='Імпорт'; rec.at=formatDate(new Date());
            exH.measurements[pk4]=rec; }
        }
      }
    }catch(e){ Logger.log('  x [%s/%s] %s',loc,grp,e&&e.message||e); }
    finally{ if(conv.isTemp&&conv.ssId){ try{ DriveApp.getFileById(conv.ssId).setTrashed(true); }catch(e){} } }
  }

  var written=0;
  if(doWrite){
    for(var ri in pendingDev){ clSheet.getRange(Number(ri)+1, ciDev+1).setValue(JSON.stringify(pendingDev[ri])); written++; }
    for(var ri2 in pendingHea){ clSheet.getRange(Number(ri2)+1, ciHea+1).setValue(JSON.stringify(pendingHea[ri2])); }
  }

  Logger.log('----- SUMMARY -----');
  Logger.log('Files: %s (skipped: %s)', stat.files, stat.filesSkipped);
  Logger.log('Child tabs: %s | matched: %s | NOT found: %s | conflicts: %s | empty: %s',
    stat.tabs, stat.matched, stat.unmatched, stat.conflict, stat.emptyTabs);
  Logger.log('Scores to write: %s | measures: %s | cards: %s', stat.critWritten, stat.measWritten, Object.keys(stat.rowsToWrite).length);
  if(matchedFuzzy.length){ Logger.log('-- matched via smart (review, %s): --', matchedFuzzy.length); matchedFuzzy.forEach(function(s){ Logger.log('  ~ '+s); }); }
  if(unmatchedList.length){ Logger.log('-- NOT found in CRM (%s): --', unmatchedList.length); unmatchedList.forEach(function(s){ Logger.log('  - '+s); }); }
  if(conflictList.length){ Logger.log('-- CONFLICTS same name (%s): --', conflictList.length); conflictList.forEach(function(s){ Logger.log('  - '+s); }); }
  Logger.log(doWrite ? ('WROTE cards: '+written) : 'DRY-RUN - nothing written. Review list above, then run runDevImport.');
  Logger.log('=== DEV IMPORT done ===');
  return {ok:true, write:doWrite, matched:stat.matched, unmatched:stat.unmatched, conflict:stat.conflict, crit:stat.critWritten, cards:Object.keys(stat.rowsToWrite).length};
}


// v6.64.2: діагностика швидкості — запустити вручну, дивитись Execution log.
function diagFillPerf(){
  var ss = getCRMSpreadsheet();
  var tz = ss.getSpreadsheetTimeZone() || 'Europe/Kiev';
  var attSh = ss.getSheetByName(SHEET_ATTENDANCE);
  Logger.log('Tabel rows=%s cols=%s', attSh ? attSh.getLastRow() : 0, attSh ? attSh.getLastColumn() : 0);
  var t0 = new Date().getTime();
  _fillExpected();
  var t1 = new Date().getTime();
  Logger.log('_fillExpected (HR+clients): %s ms', t1 - t0);
  var today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  var av = _fillReadAttendance(attSh, tz, today);
  var t2 = new Date().getTime();
  Logger.log('attendance CAPPED read: %s rows, %s ms', av.length, t2 - t1);
  var full = attSh ? attSh.getDataRange().getValues() : [];
  var t3 = new Date().getTime();
  Logger.log('attendance FULL read: %s rows, %s ms', full.length, t3 - t2);
  Logger.log('TOTAL: %s ms', t3 - t0);
}

function diagBdayBrovary(){
  var cli = getClients();
  Logger.log('clients ok: ' + (cli && cli.ok));
  var n=0;
  (cli.data||[]).forEach(function(o){
    var loc=String(o['Локація']||'');
    if(loc.toLowerCase().indexOf('бровар')<0) return;
    n++;
    if(n<=12) Logger.log('CRM | loc="'+loc+'" | name="'+String(o['ПІБ дитини']||'')+'" | bdayRaw="'+String(o['Дата народження']||'')+'" | parsed="'+(parseRegistryBday(o['Дата народження'])||'')+'"');
  });
  Logger.log('=== Бровари у CRM-картках: '+n+' дітей ===');
  var crmSS=getCRMSpreadsheet();
  var payData=crmSS.getSheetByName(SHEET_PAYMENTS).getDataRange().getValues();
  var payHdr=payData[0].map(function(h){return String(h||'');});
  var locI=payHdr.indexOf('Локація');
  var locs={};
  for(var i=1;i<payData.length;i++){ var l=String(payData[i][locI]||''); if(l.toLowerCase().indexOf('бровар')>=0) locs[l]=(locs[l]||0)+1; }
  Logger.log('=== Оплати Бровари — назви локацій: '+JSON.stringify(locs)+' ===');
}

function diagInvoiceRequisites(){
  var sh = SpreadsheetApp.openById(CONFIG_SHEET_ID).getSheetByName('Реквізити_Локацій');
  if (!sh) { Logger.log('❌ Аркуш "Реквізити_Локацій" не знайдено'); return; }
  var vals = sh.getDataRange().getValues();
  Logger.log('=== Реквізити_Локацій: %s рядків ===', vals.length);
  for (var r = 0; r < vals.length; r++){
    var a = JSON.stringify(String(vals[r][0] == null ? '' : vals[r][0]));
    var b = JSON.stringify(String(vals[r][1] == null ? '' : vals[r][1]));
    var nm = String(vals[r][2] || '');
    Logger.log('row %s | loc=%s | type=%s | name="%s"', r, a, b, nm);
  }
}

function backupAbsences(){
  var ss = getCRMSpreadsheet();
  var sh = ss.getSheetByName(SHEET_CLIENTS);
  var vals = sh.getDataRange().getValues();
  var hdr = vals[0];
  var iAbs = hdr.indexOf('Відсутності (JSON)');
  var iName = hdr.indexOf('ПІБ дитини');
  if(iAbs<0){ Logger.log('❌ колонка не знайдена'); return; }
  var stamp = Utilities.formatDate(new Date(), 'GMT+3', 'yyyyMMdd_HHmmss');
  var bName = 'BACKUP_absences_'+stamp;
  var bk = ss.insertSheet(bName);
  var out = [['row','name','absences_json']];
  for(var r=1;r<vals.length;r++){
    out.push([r+1, String(vals[r][iName]||''), String(vals[r][iAbs]||'')]);
  }
  bk.getRange(1,1,out.length,3).setValues(out);
  Logger.log('✅ Бекап створено: аркуш "'+bName+'", рядків='+(out.length-1));
}

function cleanImportDups(APPLY){
  var ss = getCRMSpreadsheet();
  var sh = ss.getSheetByName(SHEET_CLIENTS);
  var vals = sh.getDataRange().getValues();
  var hdr = vals[0];
  var iAbs = hdr.indexOf('Відсутності (JSON)');
  var iName = hdr.indexOf('ПІБ дитини');
  var mode = APPLY ? '🔴 РЕАЛЬНЕ ВИДАЛЕННЯ' : '🟡 DRY-RUN (тільки показ)';
  Logger.log('=== '+mode+' ===');
  var totDel=0, totRows=0;
  for(var r=1;r<vals.length;r++){
    var raw=String(vals[r][iAbs]||'');
    if(raw.length<5) continue;
    var arr; try{ arr=JSON.parse(raw); }catch(e){ continue; }
    if(!arr||!arr.length) continue;
    var byKey={};
    arr.forEach(function(a){
      if(a.type!=='vacation'){ return; }
      if(a.status==='cancelled'||a.status==='rejected'){ return; }
      var key='';
      var m=String(a.note||'').match(/Payment:\s*"([^"]+)"/);
      if(m){ key='imp:'+m[1].toLowerCase().replace(/\s+/g,''); }
      else if(a.from){ key='rng:'+a.from+'_'+a.to; }
      else { return; }
      (byKey[key]=byKey[key]||[]).push(a);
    });
    var toDelete={};
    Object.keys(byKey).forEach(function(k){
      var g=byKey[k];
      if(g.length<2) return;
      var withDates=g.filter(function(a){ return !!a.from; });
      var noDates=g.filter(function(a){ return !a.from; });
      if(withDates.length>0 && noDates.length>0){
        noDates.forEach(function(a){ toDelete[a.id]=true; });
      }
    });
    var delIds=Object.keys(toDelete);
    if(delIds.length){
      totRows++;
      totDel+=delIds.length;
      Logger.log('РЯДОК '+(r+1)+' | '+String(vals[r][iName]||'')+' → видалити '+delIds.length+' дубль(ів):');
      arr.forEach(function(a){
        if(toDelete[a.id]) Logger.log('    ❌ '+a.id+' from='+a.from+' note='+String(a.note||'').substring(0,45));
      });
      if(APPLY){
        var kept=arr.filter(function(a){ return !toDelete[a.id]; });
        sh.getRange(r+1, iAbs+1).setValue(JSON.stringify(kept));
      }
    }
  }
  Logger.log('=== ПІДСУМОК: рядків='+totRows+' | видалити записів='+totDel+' | режим='+(APPLY?'ЗАСТОСОВАНО':'тільки показ')+' ===');
}
function cleanImportDupsDryRun(){ cleanImportDups(false); }
function cleanImportDupsApply(){ cleanImportDups(true); }

function _normForMatch(s){
  return String(s == null ? '' : s)
    .replace(/[ ​‌‍﻿‎‏]/g, '')
    .replace(/[‐‑‒–—−]/g, '-')
    .replace(/[‘’ʼ′]/g, "'")
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

function diagMatchAll(){
  var ss = getCRMSpreadsheet();
  var cli = getClients();
  var clientSet = {};
  var iN, iL;
  (function(){
    var sh = ss.getSheetByName(SHEET_CLIENTS);
    var v = sh.getDataRange().getValues();
    var h = v[0];
    iN = h.indexOf('ПІБ дитини'); iL = h.indexOf('Локація');
    for(var r=1;r<v.length;r++){
      var key = _normForMatch(v[r][iN]) + '|||' + _normForMatch(v[r][iL]);
      clientSet[key] = true;
    }
  })();
  var at = ss.getSheetByName(SHEET_ATTENDANCE);
  var av = at.getDataRange().getValues();
  var ah = av[0];
  var aN = ah.indexOf('ПІБ дитини'); if(aN<0) aN = ah.indexOf("Ім'я");
  var aL = ah.indexOf('Локація');
  Logger.log('=== Табель: name col='+aN+' loc col='+aL+' | Клієнти: name='+iN+' loc='+iL+' ===');
  var miss=0, strictMiss=0, byLoc={};
  var clientStrict={};
  (function(){
    var sh = ss.getSheetByName(SHEET_CLIENTS);
    var v = sh.getDataRange().getValues();
    for(var r=1;r<v.length;r++){
      clientStrict[String(v[r][iN]||'').trim()+'|||'+String(v[r][iL]||'').trim()] = true;
    }
  })();
  for(var r=1;r<av.length;r++){
    var nm = String(av[r][aN]||'').trim();
    var lc = String(av[r][aL]||'').trim();
    if(!nm) continue;
    var normKey = _normForMatch(nm)+'|||'+_normForMatch(lc);
    var strictKey = nm+'|||'+lc;
    var nFound = !!clientSet[normKey];
    var sFound = !!clientStrict[strictKey];
    if(!sFound){
      strictMiss++;
      byLoc[lc] = (byLoc[lc]||0)+1;
      if(!nFound){
        miss++;
        Logger.log('❌ НЕ ЗНАЙДЕНО НАВІТЬ ПІСЛЯ НОРМ: "'+nm+'" | "'+lc+'"');
      } else {
        Logger.log('🟡 строго НІ, після норм ТАК: "'+nm+'" | "'+lc+'"');
      }
    }
  }
  Logger.log('=== ПІДСУМОК ===');
  Logger.log('Строго не знаходиться: '+strictMiss+' дітей');
  Logger.log('З них норм-фікс рятує: '+(strictMiss-miss));
  Logger.log('Лишається проблемних після норм: '+miss);
  Logger.log('По локаціях (строгий міс): '+JSON.stringify(byLoc));
}


function diagShysh(){
  var ss = getCRMSpreadsheet();
  var sh = ss.getSheetByName(SHEET_CLIENTS);
  var v = sh.getDataRange().getValues();
  var h = v[0];
  var iN = h.indexOf('ПІБ дитини');
  var iL = h.indexOf('Локація');
  var iD = h.indexOf('Дата договору');
  var iA = h.indexOf('Відсутності (JSON)');
  Logger.log('=== Пошук "Шиш" ===');
  for(var r=1;r<v.length;r++){
    var nm=String(v[r][iN]||'');
    if(nm.toLowerCase().indexOf('шиш')<0) continue;
    Logger.log('РЯДОК '+(r+1)+': "'+nm+'" | лок='+v[r][iL]+' | договір='+v[r][iD]);
    var raw=String(v[r][iA]||'');
    Logger.log('RAW absences: '+raw.substring(0,2000));
    try{
      var arr=JSON.parse(raw||'[]');
      Logger.log('Записів: '+arr.length);
      arr.forEach(function(a,i){
        Logger.log('  ['+i+'] type='+a.type+' from='+a.from+' to='+a.to+' days='+(a.workDays||a.days)+' weeks='+a.weeks+' status='+a.status+' note='+String(a.note||'').substring(0,50));
      });
    }catch(e){ Logger.log('parse err: '+e); }
  }
  Logger.log('=== кінець ===');
}

function diagShyshCalc(){
  var contract = new Date(2025,0,20);
  var todayD = new Date(2026,5,29);
  function ymd(d){ return Utilities.formatDate(d,'GMT+2','yyyy-MM-dd'); }
  var diffY = todayD.getFullYear()-contract.getFullYear();
  Logger.log('Договір=2025-01-20, сьогодні=2026-06-29, diffY='+diffY);
  for(var off=diffY-1; off<=diffY+1; off++){
    var from=new Date(contract.getFullYear()+off, contract.getMonth(), contract.getDate());
    var to=new Date(contract.getFullYear()+off+1, contract.getMonth(), contract.getDate());
    to.setDate(to.getDate()-1);
    var contains = (todayD>=from && todayD<=to);
    Logger.log('  off='+off+': рік ['+ymd(from)+' .. '+ymd(to)+'] містить сьогодні? '+contains);
  }
  var abs=[
    {from:'2025-05-26',to:'2025-05-30'},
    {from:'2025-06-02',to:'2025-06-06'},
    {from:'2025-10-06',to:'2025-10-10'},
    {from:'2025-12-29',to:'2026-01-02'}
  ];
  var yFrom=new Date(2026,0,20), yTo=new Date(2027,0,19);
  Logger.log('=== Поточний договірний рік: 2026-01-20 .. 2027-01-19 ===');
  var used=0;
  abs.forEach(function(a){
    var af=new Date(a.from), at=new Date(a.to);
    var inYear = (at>=yFrom && af<=yTo);
    Logger.log('  відпустка '+a.from+'..'+a.to+' → у поточному році? '+inYear);
  });
  Logger.log('=== Якщо всі 4 показують FALSE — ліміт має бути 20 (чистий). Якщо хоч одна TRUE — баг меж року ===');
}

function diagFindVacExc(){
  var targets = ['Кузьмін','Андрєєв','Андреєв','Вірьовк','Середа','Голодн','Тандиряк','Гаркуш','Мельничук'];
  var ss = getCRMSpreadsheet();
  var sh = ss.getSheetByName(SHEET_CLIENTS);
  var v = sh.getDataRange().getValues();
  var h = v[0];
  var iN = h.indexOf('ПІБ дитини');
  var iL = h.indexOf('Локація');
  var iD = h.indexOf('Дата договору');
  Logger.log('=== Пошук дітей-винятків ===');
  targets.forEach(function(p){
    var found=false;
    for(var r=1;r<v.length;r++){
      var nm=String(v[r][iN]||'');
      if(nm.toLowerCase().indexOf(p.toLowerCase())>=0){
        found=true;
        Logger.log('"'+nm+'" | лок='+v[r][iL]+' | договір='+v[r][iD]);
      }
    }
    if(!found) Logger.log('❌ "'+p+'" не знайдено');
  });
  Logger.log('=== кінець ===');
}


function diagFind2(){
  var targets = ['Скорін','Скоріна','Голодн','Голодний','Голодько'];
  var ss = getCRMSpreadsheet();
  var sh = ss.getSheetByName(SHEET_CLIENTS);
  var v = sh.getDataRange().getValues();
  var h = v[0];
  var iN = h.indexOf('ПІБ дитини');
  var iL = h.indexOf('Локація');
  var iD = h.indexOf('Дата договору');
  Logger.log('=== Пошук Скоріна + Голодний ===');
  targets.forEach(function(p){
    var found=false;
    for(var r=1;r<v.length;r++){
      var nm=String(v[r][iN]||'');
      if(nm.toLowerCase().indexOf(p.toLowerCase())>=0){
        found=true;
        Logger.log('"'+nm+'" | лок='+v[r][iL]+' | договір='+v[r][iD]);
      }
    }
    if(!found) Logger.log('нема: "'+p+'"');
  });
  Logger.log('=== кінець ===');
}


function diagFind3(){
  var targets = ['Городн','Городній','Городний'];
  var ss = getCRMSpreadsheet();
  var sh = ss.getSheetByName(SHEET_CLIENTS);
  var v = sh.getDataRange().getValues();
  var h = v[0];
  var iN = h.indexOf('ПІБ дитини');
  var iL = h.indexOf('Локація');
  var iD = h.indexOf('Дата договору');
  Logger.log('=== Пошук Городний ===');
  targets.forEach(function(p){
    var found=false;
    for(var r=1;r<v.length;r++){
      var nm=String(v[r][iN]||'');
      if(nm.toLowerCase().indexOf(p.toLowerCase())>=0){
        found=true;
        Logger.log('"'+nm+'" | лок='+v[r][iL]+' | договір='+v[r][iD]);
      }
    }
    if(!found) Logger.log('нема: "'+p+'"');
  });
  Logger.log('=== кінець ===');
}
