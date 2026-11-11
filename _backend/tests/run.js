#!/usr/bin/env node
// Тести бекенду перед деплоєм (v7.530). Запуск: node _backend/tests/run.js [шлях до crm_script.js]
// Беруть ЧИСТІ функції прямо з crm_script.js (без Google-сервісів) і проганяють реальні випадки 02.11.2026:
//   рахунки — «борг / переплата / коригування ⚠️» (_invPrevBalances → _invExtrasCalc / _invStudyCalc → _invPrecheckItems),
//   знижка відпустки — _vacCellPlan. Будь-яка розбіжність → код виходу 1, і predeploy.sh не копіює бекенд у буфер.
'use strict';
const fs = require('fs'), path = require('path'), crypto = require('crypto');
const SRC = process.argv[2] || path.join(__dirname, '..', 'crm_script.js');
const src = fs.readFileSync(SRC, 'utf8');
const FX = JSON.parse(fs.readFileSync(path.join(__dirname, 'fixtures_invoices.json'), 'utf8'));

function fn(name){
  const i = src.indexOf('\nfunction ' + name + '(');
  if (i < 0) throw new Error('функцію ' + name + ' не знайдено в ' + SRC);
  let d = 0;
  for (let k = src.indexOf('{', i); k < src.length; k++){
    if (src[k] === '{') d++;
    else if (src[k] === '}' && --d === 0) return src.slice(i, k + 1);
  }
  throw new Error('не вдалося виділити ' + name);
}
function varBlock(name){   // багаторядкове оголошення var NAME = {...};
  const i = src.indexOf('\nvar ' + name + ' ');
  if (i < 0) throw new Error('змінну ' + name + ' не знайдено');
  const j = src.indexOf(';\n', i);
  return src.slice(i + 1, j + 1);
}
function varLine(name){
  const m = new RegExp('\\nvar ' + name + '\\s*=\\s*([^\\n]*);').exec(src);
  if (!m) throw new Error('змінну ' + name + ' не знайдено');
  return 'var ' + name + ' = ' + m[1] + ';';
}

// ── мінімальне оточення Apps Script ──
let JOURNAL = [['loc','kind','name','year','month','last_written_sum','last_written_at']];
const env = {
  trim: (x) => String(x == null ? '' : x).trim(),
  _journalNormName: (x) => String(x == null ? '' : x).replace(/[\s ]+/g, '').replace(/[’ʼ`']/g, "'").toLowerCase(),
  _yearlyDepartedIdx: () => -1, _isYearlyDeparted: () => false,
  Logger: {log(){}},
  _getExportJournalSheet: () => ({getDataRange: () => ({getValues: () => JOURNAL})}),
  Utilities: {DigestAlgorithm: {MD5: 'md5'}, Charset: {UTF_8: 'utf8'},
    computeDigest: (_a, s) => Array.from(crypto.createHash('md5').update(String(s), 'utf8').digest()).map((b) => (b > 127 ? b - 256 : b))},
};
const code = [varLine('MONTHS_CAL'), varLine('PROMO_FORMULA_RE'),
  ...['_invR2','_prevYM','_invSplit','_invExtrasCalc','_invStudyCalc','_invPrevBalances','_invPrecheckItems','_invPrecheckHash','_vacCellPlan']
     .concat(src.indexOf('\nfunction _dopSubPlan(') >= 0 ? ['_dopSubPlan'] : []).map(fn)
     .concat(src.indexOf('\nfunction _predGroupAllowed(') >= 0
       ? [varLine('PRED_GROUP_TYPES'), varBlock('_GT_TRANSLIT'), varLine('_GT_BABY')]
           .concat(['_groupLetters','_predGroupKey','_canonGroupType','_normalizeGroupType','_predGroupAllowed'].map(fn)) : []),
  'return {_invPrevBalances, _invExtrasCalc, _invStudyCalc, _invPrecheckItems, _invPrecheckHash, _vacCellPlan,' +
  ' _dopSubPlan: (typeof _dopSubPlan === "function" ? _dopSubPlan : null),' +
  ' _predGroupAllowed: (typeof _predGroupAllowed === "function" ? _predGroupAllowed : null),' +
  ' _normalizeGroupType: (typeof _normalizeGroupType === "function" ? _normalizeGroupType : null)};'].join('\n');
const F = new Function(...Object.keys(env), code)(...Object.values(env));

let pass = 0, fail = 0;
function check(id, title, got, exp){
  const bad = Object.keys(exp).filter((k) => JSON.stringify(got[k]) !== JSON.stringify(exp[k]));
  if (bad.length){ fail++; console.log('✗ ' + id + ' — ' + title); bad.forEach((k) => console.log('    ' + k + ': очікували ' + JSON.stringify(exp[k]) + ', вийшло ' + JSON.stringify(got[k]))); }
  else { pass++; console.log('✓ ' + id); }
}

const children = [];
FX.cases.forEach((c) => {
  JOURNAL = [JOURNAL[0]].concat(c.journal || []);
  const head = Object.keys(c.rows[0]);
  const vals = [head].concat(c.rows.map((r) => head.map((k) => r[k])));
  const pb = F._invPrevBalances(c.loc, FX.payMonth, vals, FX.payYear)[c.name];
  if (!pb){ fail++; console.log('✗ ' + c.id + ' — дитину не знайдено в рядках'); return; }
  const sum = (col) => c.rows.reduce((a, r) => a + (Number(r[col]) || 0), 0);
  const ch = {name: c.name, group: '', studyBudget: sum('Листопад-Бюджет-навч'), extrasPayment: sum('Листопад-Бюджет-доп'),
              prevNavch: pb.monN, prevDop: pb.monD};
  if (c.extras){
    const r = F._invExtrasCalc(c.extras.marks, ch.extrasPayment, 0, false, pb);
    check(c.id + ' / допи', c.title, {debt: r.debt, adjust: r.adjust, due: Math.max(0, r.due)}, c.extras.expect);
    Object.assign(ch, {extrasBreakdownSum: c.extras.marks, extrasAdjustment: r.debt, extrasPaymentDiff: r.adjust, extrasOnlySum: Math.max(0, r.due)});
  }
  if (c.study){
    const r = F._invStudyCalc(ch.studyBudget, c.study.net, c.study.netPrev, pb);
    check(c.id + ' / навчання', c.title, {debt: r.debt, adjust: r.adjust, due: Math.max(0, r.due)}, c.study.expect);
    Object.assign(ch, {studyNet: c.study.net, studyCarry: r.debt, studyChange: r.adjust, paymentSum: Math.max(0, r.due)});
  }
  ch._listed = c.listed;
  children.push(ch);
});

// перевірка перед розсилкою: у списку саме ті, де є борг / переплата / коригування; чиста дитина — ні; хеш реагує на суми
const clean = {name: 'Чиста Дитина', studyBudget: 21500, paymentSum: 21500, studyCarry: 0, studyChange: 0, extrasPayment: 3000,
               extrasBreakdownSum: 3000, extrasAdjustment: 0, extrasPaymentDiff: 0, extrasOnlySum: 3000, _listed: false};
const all = children.concat([clean]);
const items = F._invPrecheckItems(all), names = items.map((x) => x.name);
all.forEach((c) => check('перевірка: ' + c.name, 'у списку перед розсилкою = ' + c._listed, {listed: names.indexOf(c.name) >= 0}, {listed: c._listed}));
const tk = items.find((x) => x.name === 'Ткаленко Анна');
check('перевірка: розшифровка Ткаленко', 'допи: бюджет 8320, відмітки 8870, коригування −550', tk && tk.extras ? {budget: tk.extras.budget, marks: tk.extras.marks, change: tk.extras.change} : {},
      {budget: 8320, marks: 8870, change: -550});
const ml = items.find((x) => x.name === 'Малпетті Даніелє');
check('перевірка: розшифровка Малпетті', 'навчання: місяць з залишком — жовтень 0 / 21500', ml && ml.study ? {months: ml.study.months} : {},
      {months: [{m: 10, label: 'Жовтень', budget: 0, fact: 21500, diff: -21500}]});
const h1 = F._invPrecheckHash(items), items2 = JSON.parse(JSON.stringify(items));
items2[0].study ? (items2[0].study.due += 1) : (items2[0].extras.due += 1);
check('перевірка: хеш', 'хеш стабільний і змінюється від суми', {same: F._invPrecheckHash(F._invPrecheckItems(all)) === h1, changed: F._invPrecheckHash(items2) !== h1}, {same: true, changed: true});

FX.vacation.forEach((v) => {
  const r = F._vacCellPlan.apply(null, v.args), got = {};
  Object.keys(v.expect).forEach((k) => { got[k] = r[k]; });
  check(v.id, v.title, got, v.expect);
});

// v7.531: абонемент для батьків — ціну місяця несе перша відмітка дитини, решта 0 (реальні відмітки Нац.Гвардії, жовтень 2026)
if (FX.subscription){
  if (!F._dopSubPlan){ fail++; console.log('✗ _dopSubPlan відсутній у ' + SRC); }
  else FX.subscription.cases.forEach((c) => {
    const sp = {}; Object.keys(FX.subscription.price).forEach((k) => { sp[Number(k)] = FX.subscription.price[k]; });
    const r = F._dopSubPlan(c.marks.map((m, i) => Object.assign({row: i + 2}, m)), sp);
    const groups = r.groups.map((g) => ({child: g.child, ym: g.ym, marks: g.marks, before: g.before, after: g.after}));
    check(c.id + ' / групи', c.title, {groups}, {groups: c.expectGroups});
    check(c.id + ' / зміни', c.title, {changes: r.changes.length}, {changes: c.expectChanges});
    if (c.untouchedIds) check(c.id + ' / інші заняття', 'не чіпаємо', {touched: r.changes.filter((x) => c.untouchedIds.indexOf(x.id) >= 0).length}, {touched: 0});
    if (c.expectFirstTo){ const hit = r.changes.find((x) => x.id === c.expectFirstTo.id); check(c.id + ' / перенесення ціни', c.title, {to: hit ? hit.to : null}, {to: c.expectFirstTo.to}); }
    const total = c.marks.filter((m) => sp.hasOwnProperty(m.actId)).reduce((a, m) => { const ch = r.changes.find((x) => x.id === m.id); return a + (ch ? ch.to : m.price); }, 0);
    check(c.id + ' / сума', 'після нормалізації сума = ціна × абонементів', {total}, {total: c.expectGroups.reduce((a, g) => a + g.after, 0)});
  });
}

// v7.532: предметники — позиція каталогу лише для певних груп (Осокорки, «Підготовка до школи» → Preschool)
if (FX.predGroups){
  if (!F._predGroupAllowed){ fail++; console.log('✗ _predGroupAllowed відсутній у ' + SRC); }
  else FX.predGroups.cases.forEach((c) => {
    const got = {}; Object.keys(c.expect).forEach((g) => { got[g] = F._predGroupAllowed(c.allowed, g, F._normalizeGroupType(g)); });
    check(c.id, 'дозволені групи: ' + (c.allowed.join(', ') || 'усі'), got, c.expect);
  });
}

// v7.533: запобіжник Salary-рядка — ім'я співробітника не робить рядок чужим (Школа Осокорки, Бєлобровець ↔ картка «Ольга Павленко»)
if (src.indexOf('\nfunction _salaryRowGuard(') >= 0){
  const G = new Function(['_salaryRowTokens','_normEmpNm','_lat2cyr','_levSalary','_surnameSim','_salaryRowGuard'].map(fn).join('\n') +
                         '\nreturn _salaryRowGuard;')();
  const row = {raw: 'Директорка  Ольга Бєлобровець'};
  const pavl = {rowNum: 712, last: 'Ольга', first: 'Павленко'}, shk = {rowNum: 692, last: 'Шкіньова', first: 'Ольга'};
  const bel = {rowNum: 693, last: 'Бєлобровець', first: 'Ольга'}, belOld = {rowNum: 693, last: 'Белобровец', first: 'Ольга'};
  const pick = (g) => ({ok: g.ok, why: g.why});
  check('guard: Бєлобровець — свій рядок', 'прізвище з банку', pick(G(row, bel, [bel, pavl, shk], true)), {ok: true, why: 'own'});
  check('guard: Белобровец (старе написання) — не «чужий»', 'картка last=«Ольга» не власник', pick(G(row, belOld, [belOld, pavl, shk], true)), {ok: true, why: 'anon'});
  check('guard: Павленко Вікторія — чужий рядок', 'запобіжник лишається', pick(G({raw: 'Вчитель Павленко Вікторія'}, shk, [shk, {rowNum: 714, last: 'Павленко', first: 'Вікторія'}], true)), {ok: false, why: 'foreign'});
}

// v7.534: бот рахунків — (1) зіставлення заявок з випискою (реальні заявки й платежі 28.10–10.11.2026),
//   (2) «оплачено» бота не чіпає виписку / журнал розпису, (3) «✅ Оплачено» лише в банківські дні о 09:00.
if (src.indexOf('\nfunction _invMatchCore(') >= 0){
  const FXM = JSON.parse(fs.readFileSync(path.join(__dirname, 'fixtures_invmatch.json'), 'utf8'));
  const kyiv = (d) => { const o = {}; new Intl.DateTimeFormat('en-GB', {timeZone: 'Europe/Kiev', year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', weekday: 'short', hourCycle: 'h23'})
    .formatToParts(d).forEach((p) => { o[p.type] = p.value; }); return o; };
  const U = {
    formatDate: (d, tz, f) => { const o = kyiv(d); const wd = {Mon: 1, Tue: 2, Wed: 3, Thu: 4, Fri: 5, Sat: 6, Sun: 7}[o.weekday];
      return f === 'u' ? String(wd) : f === 'H' ? String(Number(o.hour)) : f === 'yyyy-MM-dd' ? o.year + '-' + o.month + '-' + o.day : ''; },
    parseDate: (str, tz, f) => { const m = /^(\d{4})-(\d{2})-(\d{2}) (\d{2}):(\d{2})$/.exec(str);
      const guess = Date.UTC(+m[1], m[2] - 1, +m[3], +m[4], +m[5]); const o = kyiv(new Date(guess));
      const off = Date.UTC(+o.year, o.month - 1, +o.day, +o.hour, +o.minute) - guess; return new Date(guess - off); }
  };
  const vars = ['INV_AMOUNT_EPS','INV_DAYS_BEFORE','INV_DAYS_AFTER','INV_DATE_PLAUSIBLE','INV_CREATED_BEFORE','INV_CREATED_AFTER','INV_VAT_RATE','INV_DUP_DAYS','INV_QUIET_UNTIL_HOUR'].map(varLine);
  const fns = ['_invDateMs','_invNumKey','_invPurposeHasNumber','_invEdrMatch','_invEdr7','_invPurposeHasNumberX','_invPurposeInvNumbers',
               '_invPurposeOtherNumber','_invAmtEq','_invNumsCompatible','_invWindowOk','_invIsDup','_invMatchCore','_invRefBase','_invBankDayKyiv','_invNextFlushAt'].map(fn);
  const B = new Function('Utilities', vars.concat(fns).join('\n') + '\nreturn {_invMatchCore, _invRefBase, _invBankDayKyiv, _invNextFlushAt};')(U);
  const dms = (s) => { const m = /^(\d{1,2})\.(\d{1,2})\.(\d{4})/.exec(s || ''); return m ? new Date(+m[3], m[2] - 1, +m[1]).getTime() : 0; };
  const mk = (x) => Object.assign({}, x, {invDateMs: dms(x.invDate), createdMs: dms(x.created)});
  FXM.cases.forEach((c) => {
    const r = B._invMatchCore(c.open.map(mk), c.pays, c.paid.map(mk), {});
    const got = {};
    Object.keys(c.expect).forEach((id) => {
      const g = r.good.find((x) => x.invs.some((i) => i.id === id));
      got[id] = g ? g.kind : (r.partial.some((q) => q.inv.id === id) ? 'partial' : null);
    });
    check('рахунки: ' + c.id, c.title, got, c.expect);
  });
  // (2) запис «оплачено» — лише лист заявок; журнал розпису виписки й OPEX/Salary бот не пише
  const body = fn('matchInvoicesToPayments') + fn('_invAutoMatch') + fn('_bankNightly__impl');
  check('рахунки: бот не розписує виписку', 'у записі «оплачено» немає OPEX / Salary / журналу виписки',
        {touches: /opexAddExpenses|salaryAddExtrasPayments|OPEX_EXP_LOG_SHEET|reconcileApply|_getReconcileLogSheet/.test(body)}, {touches: false});
  check('рахунки: «розписано» — з журналу виписки', 'частини платежу (-split-N, -salsplit-N) — той самий платіж',
        {a: B._invRefBase('JBKLQB6OF4KBZN.1-split-2'), b: B._invRefBase('JBKLQB6OF4KBZN.1-salsplit-0'), c: B._invRefBase('JBKLQB6OF4KELE.1')},
        {a: 'JBKLQB6OF4KBZN.1', b: 'JBKLQB6OF4KBZN.1', c: 'JBKLQB6OF4KELE.1'});
  // (3) банківські дні: субота/неділя — ні; черга «✅ Оплачено» з суботи → понеділок 09:00 (Київ)
  const iso = (d) => U.formatDate(d, 'Europe/Kiev', 'yyyy-MM-dd') + ' ' + U.formatDate(d, 'Europe/Kiev', 'H') + 'h';
  check('рахунки: банківський день', 'пт так, сб/нд ні, пн так',
        {fri: B._invBankDayKyiv(new Date('2026-11-06T10:00:00Z')), sat: B._invBankDayKyiv(new Date('2026-11-07T10:00:00Z')),
         sun: B._invBankDayKyiv(new Date('2026-11-08T10:00:00Z')), mon: B._invBankDayKyiv(new Date('2026-11-09T10:00:00Z'))},
        {fri: true, sat: false, sun: false, mon: true});
  check('рахунки: «✅ Оплачено» після вихідних', 'збіг у суботу 05:00 → понеділок 09:00; пт 05:00 → пт 09:00',
        {sat: iso(B._invNextFlushAt(new Date('2026-11-07T03:00:00Z'))), fri: iso(B._invNextFlushAt(new Date('2026-11-06T03:00:00Z'))),
         sun: iso(B._invNextFlushAt(new Date('2026-11-08T12:00:00Z')))},
        {sat: '2026-11-09 9h', fri: '2026-11-06 9h', sun: '2026-11-09 9h'});
}

// v7.535: особистий звіт CFO «оплачено, ще не розписано» — лише в особистий чат, у групу директорів нічого
if (src.indexOf('\nfunction _invCfoReportTexts(') >= 0){
  const C = new Function([varLine('INV_TG_MAX')].concat(['_htmlEsc','_invMoneyUa','_invCfoReportTexts','_invCfoTarget'].map(fn)).join('\n') +
                         '\nreturn {_invCfoReportTexts, _invCfoTarget};')();
  const GROUP = '-5142518705';
  check('звіт CFO: адресат', 'лише особистий чат; група директорів / будь-яка група / порожньо — ні',
        {priv: C._invCfoTarget('123456789', GROUP), group: C._invCfoTarget(GROUP, GROUP), other: C._invCfoTarget('-100200300', GROUP), none: C._invCfoTarget('', GROUP)},
        {priv: '123456789', group: '', other: '', none: ''});
  check('звіт CFO: порожній список', 'не надсилати', {n: C._invCfoReportTexts([], '11.11.2026').length}, {n: 0});
  const items = [{payDate: '10.11.2026', loc: 'Кругла', supplier: 'TJOB', amount: 4275, ids: ['230']},
                 {payDate: '06.11.2026', loc: 'Пуща', supplier: 'ТОВ "ШЕРИФ БІЗНЕС"', amount: 945, ids: ['213', '212']},
                 {payDate: '10.11.2026', loc: 'Бігова', supplier: 'ФОП Грень І. С.', amount: 2000, ids: ['269']}];
  const t = C._invCfoReportTexts(items, '11.11.2026');
  check('звіт CFO: рядки й разом', 'дата · локація · постачальник · сума · №, за датою; «Разом» унизу',
        {parts: t.length, first: t[0].split('\n')[1], dup: /#213, #212/.test(t[0]), total: t[0].trim().split('\n').pop()},
        {parts: 1, first: '06.11 · Пуща · ТОВ &quot;ШЕРИФ БІЗНЕС&quot; · <b>945,00 ₴</b> · #213, #212'.replace('&quot;', '"').replace('&quot;', '"'), dup: true, total: '<b>Разом: 7 220,00 ₴</b>'});
  const many = Array.from({length: 120}, (_, i) => ({payDate: '10.11.2026', loc: 'Локація ' + i, supplier: 'Постачальник з довгою назвою ' + i, amount: 100, ids: [String(i)]}));
  const tm = C._invCfoReportTexts(many, '');
  check('звіт CFO: довгий список', 'частини ≤ ліміту Telegram, «Разом» лише в останній',
        {ok: tm.length > 1 && tm.every((x) => x.length <= 4096), totalOnce: tm.filter((x) => /Разом:/.test(x)).length, last: /Разом: 12 000,00/.test(tm[tm.length - 1])},
        {ok: true, totalOnce: 1, last: true});
  // у групу (заявки, черга 09:00, повтор) бот пише лише «✅ Оплачено …» — «розписано» туди не потрапляє
  const sends = [];
  const re = /_invSend\(([^,]+),\s*([^;]*?)\)/g; let m;
  while ((m = re.exec(src))){
    const fstart = src.lastIndexOf('\nfunction ', m.index), fname = /function (\w+)/.exec(src.slice(fstart, fstart + 80))[1];
    sends.push({fname, to: m[1].trim(), text: m[2]});
  }
  const cfoFns = ['_invCfoReport__impl', '_invCfoLinkFromMessage'];
  const leak = sends.filter((x) => cfoFns.indexOf(x.fname) < 0 && /розпис|texts|_invCfo/.test(x.text));
  const cfoTo = sends.filter((x) => cfoFns.indexOf(x.fname) >= 0).every((x) => x.to === 'to' || x.to === 'id');
  const queueOnly = (src.match(/(?<!function )_invQueuePaid\([^;]*\)/g) || []).every((x) => /✅ Оплачено/.test(x));
  check('звіт CFO: група директорів', 'у групу — лише «✅ Оплачено», звіт іде тільки на адресата з _invCfoTarget',
        {leak: leak.length, cfoTo: cfoTo, queueOnly: queueOnly, privateHook: /chat\.type === 'private'\) return _invCfoLinkFromMessage/.test(src)},
        {leak: 0, cfoTo: true, queueOnly: true, privateHook: true});
}

// v7.536: групи «Банк — <локація>» — ізоляція, боржники (Крошкін, Абраменко), вступний частинами, заявка понад бюджет
if (src.indexOf('\nfunction _bankChatFor(') >= 0){
  const Bv = ['_APOSTROPHES','INV_TG_MAX','BANK_DEBT_MIN','BANK_DEBT_DAYS','BANK_GROUPS_PROP'].map(varLine).concat([varBlock('BANK_INV_CAT_RULES')]);
  const Bf = ['_lat2cyr','_nameFold','_nameFoldTight','_journalNormName','_bankWords','_dopAnKid','_levSalary','_surnameSim','_htmlEsc','_invMoneyUa','_bankGroupsParse','_bankChatFor','_bankLocOfChat','_bankMoney',
              '_bankNormCat','_bankInvCat','_bankDebtDayFor','_bankDebtorsCalc','_bankEntryCalc','_bankOpexCalc','_bankInvOver','_bankOverNew','_bankChunks',
              '_bankPaidTexts','_bankDebtorsTexts','_bankOpexTexts','_bankInvOverText','_bankOverTexts','_bankCommentTarget','_bankSuspectDebts','_bankHoldFilter'].map(fn);
  const K = new Function(Bv.concat(Bf).join('\n') + '\nreturn {' + ['_bankChatFor','_bankLocOfChat','_bankInvCat','_bankDebtDayFor','_bankDebtorsCalc','_bankEntryCalc',
    '_bankOpexCalc','_bankInvOver','_bankOverNew','_bankPaidTexts','_bankDebtorsTexts','_bankInvOverText','_bankCommentTarget','_dopAnKid','_bankSuspectDebts','_bankHoldFilter'].join(',') + '};')();
  const INVG = '-5142518705', CFOC = '123456789';
  // ізоляція: чат належить рівно одній локації; група рахунків / особистий чат / спільний чат двох локацій — ні
  const G = {'Позняки': '-1001', 'Пуща': '-1002'};
  check('банк: ізоляція — свій чат', 'Позняки → свій, Пуща → свій, чужої локації немає',
        {p: K._bankChatFor(G, 'Позняки', INVG, CFOC), u: K._bankChatFor(G, 'Пуща', INVG, CFOC), x: K._bankChatFor(G, 'Оранж', INVG, CFOC),
         byChat: K._bankLocOfChat(G, '-1002')}, {p: '-1001', u: '-1002', x: '', byChat: 'Пуща'});
  check('банк: ізоляція — спільний чат', 'одна група на дві локації → нікому не пишемо',
        {p: K._bankChatFor({'Позняки': '-1001', 'Пуща': '-1001'}, 'Позняки', INVG, CFOC), u: K._bankChatFor({'Позняки': '-1001', 'Пуща': '-1001'}, 'Пуща', INVG, CFOC),
         byChat: K._bankLocOfChat({'Позняки': '-1001', 'Пуща': '-1001'}, '-1001')}, {p: '', u: '', byChat: ''});
  check('банк: ізоляція — група рахунків / CFO', 'не можуть бути групою «Банк»',
        {inv: K._bankChatFor({'Позняки': INVG}, 'Позняки', INVG, CFOC), cfo: K._bankChatFor({'Позняки': CFOC}, 'Позняки', INVG, CFOC)}, {inv: '', cfo: ''});
  // боржники
  const pick = (r) => r.map((x) => ({name: x.name, navch: x.navch, dop: x.dop}));
  check('банк: Крошкін (Позняки) листопад', 'бюджет 0 → боргу немає',
        {list: pick(K._bankDebtorsCalc([{name: 'Крошкін Олексій', bs: 0, fs: 0, be: 0, fe: 0, net: 10750, marks: 4370}], {}))}, {list: []});
  check('банк: Абраменко (Позняки) листопад', 'навч. 27 500, з них перенесено 6 000; допи 10 940, з них перенесено 4 340',
        {list: pick(K._bankDebtorsCalc([{name: 'Абраменко Тимофій', bs: 27500, fs: 0, be: 10940, fe: 0, net: 21500, marks: 6600}], {}))},
        {list: [{name: 'Абраменко Тимофій', navch: {debt: 27500, carried: 6000}, dop: {debt: 10940, carried: 4340}}]});
  check('банк: Абраменко — навчання оплачено', 'хто оплатив — зникає (факт навч 27 500 на 11.11)',
        {list: pick(K._bankDebtorsCalc([{name: 'Абраменко Тимофій', bs: 27500, fs: 27500, be: 10940, fe: 0, net: 21500, marks: 6600}], {}))},
        {list: [{name: 'Абраменко Тимофій', navch: null, dop: {debt: 10940, carried: 4340}}]});
  const dt = K._bankDebtorsTexts('Позняки', 'листопад', K._bankDebtorsCalc([{name: 'Абраменко Тимофій', bs: 27500, fs: 0, be: 10940, fe: 0, net: 21500, marks: 6600}],
        {[K._dopAnKid('Абраменко Тимофій')]: 'обіцяли до 15-го — Ірина (11.11)'}), [])[0];
  check('банк: текст боржника з коментарем', 'рядок і 💬', {line: dt.split('\n')[1], cmt: /💬 обіцяли до 15-го/.test(dt)},
        {line: '• Абраменко Тимофій — навч. 27 500 (з них перенесено 6 000); допи 10 940 (з них перенесено 4 340)', cmt: true});
  check('банк: дні списку боржників', 'до 5-го — ні; 6, 8 (вихідні після 6-го) → 6; 31 → 30', {d5: K._bankDebtDayFor(5), d6: K._bankDebtDayFor(6), d8: K._bankDebtDayFor(8), d31: K._bankDebtDayFor(31)},
        {d5: 0, d6: 6, d8: 6, d31: 30});
  // вступний частинами (реальні «Факт вступ» Позняків 2026)
  const F = {}; [['Шагурі Єва', [0,0,0,10150,10250,100,0,0,0,0,0,0]], ['Драгунов Іван', [0,0,0,0,0,15000,0,0,20500,0,0,0]], ['Дністянський Миколай', [0,0,0,0,0,0,0,0,15000,0,0,0]]].forEach((x) => { F[K._dopAnKid(x[0])] = {name: x[0], months: x[1]}; });
  check('банк: вступний частинами', 'Σ «Факт вступ» від дати договору; до договору не рахуємо; інше написання прізвища',
        {list: K._bankEntryCalc([{name: 'Шагурі Єва', fee: 20500, contract: '2026-04-01'}, {name: 'Шагурі Єва', fee: 25000, contract: '2026-04-01'},
          {name: 'Драгунов Іван', fee: 20500, contract: '2026-08-31'}, {name: 'Драгунов Іван', fee: 35500, contract: '2026-08-31'},
          {name: 'Дністранський Миколай', fee: 15000, contract: '2026-09-28'}], F, 2026, '2026-11-11')},
        {list: [{name: 'Драгунов Іван', paid: 20500, fee: 35500, contract: '31.08.2026'}, {name: 'Шагурі Єва', paid: 20500, fee: 25000, contract: '01.04.2026'}]});
  // заявка понад бюджет: ХОЗ.ТОВАРИ Позняків листопад — бюджет 5 720, витрачено 5 412; нова заявка 1 055,20
  check('банк: заявка понад бюджет', 'витрачено + черга + заявка − бюджет; текст попередження',
        {over: K._bankInvOver(5720, 5412, 0, 1055.2), under: K._bankInvOver(5720, 0, 640, 1000), text: K._bankInvOverText('290', 'ХОЗ.ТОВАРИ', 747.2, 5720),
         cat: [K._bankInvCat('госп.товари', ['ХОЗ.ТОВАРИ', 'ПРАЛЬНЯ']), K._bankInvCat('Пральня 760 +', ['ХОЗ.ТОВАРИ', 'ПРАЛЬНЯ']), K._bankInvCat('Вітаю', ['ХОЗ.ТОВАРИ'])]},
        {over: 747.2, under: 0, text: '⚠️ Рахунок #290, ХОЗ.ТОВАРИ: після оплати перевищення <b>747,20 ₴</b> з 5 720. Потрібне погодження.', cat: ['ХОЗ.ТОВАРИ', 'ПРАЛЬНЯ', '']});
  const sweep = fn('_bankBudgetSweep'), bsend = fn('_bankSend');
  check('банк: попередження лише в групу «Банк»', 'тік заявок пише тільки через _bankSend, а той — лише в чат локації з _bankChatFor',
        {viaBank: /_bankSend\(loc, _bankInvOverText/.test(sweep), noInvGroup: !/_invSend\(|INVOICE_CHAT_ID|_invChatId\(\)\s*,\s*'/.test(sweep), guarded: /_bankChatFor\(_bankGroups\(\), loc,/.test(bsend)},
        {viaBank: true, noInvGroup: true, guarded: true});
  check('банк: перевищення статті', 'нове або глибше — повідомляємо; те саме — ні',
        {first: K._bankOverNew([{name: 'ХОЗ.ТОВАРИ', budget: 5720, spent: 6467}], {}, '2026-11').items.length,
         same: K._bankOverNew([{name: 'ХОЗ.ТОВАРИ', budget: 5720, spent: 6467}], {'2026-11|ХОЗ.ТОВАРИ': 747}, '2026-11').items.length,
         deeper: K._bankOverNew([{name: 'ХОЗ.ТОВАРИ', budget: 5720, spent: 7000}], {'2026-11|ХОЗ.ТОВАРИ': 747}, '2026-11').items.length}, {first: 1, same: 0, deeper: 1});
  check('банк: коментар до боргу', 'прізвище в reply → дитина; один у списку — він; незрозуміло — ні',
        {two: K._bankCommentTarget(['Абраменко Тимофій', 'Шагурі Єва'], 'Абраменко: обіцяли до 15-го'), one: K._bankCommentTarget(['Шагурі Єва'], 'обіцяли до 15-го'),
         none: K._bankCommentTarget(['Абраменко Тимофій', 'Шагурі Єва'], 'обіцяли до 15-го')},
        {two: {child: 'Абраменко Тимофій', text: 'обіцяли до 15-го'}, one: {child: 'Шагурі Єва', text: 'обіцяли до 15-го'}, none: null});
  check('банк: можливо вигаданий борг', 'Капітонов (Осокорки): вересень сплачено 25 500 при бюджеті 0, листопад 25 500 не сплачено → у списку; без переплати — ні',
        {list: K._bankSuspectDebts([{name: 'Капітонов Ян', bs: 25500, fs: 0, be: 0, fe: 0, overN: 25500, overD: 0, overMonths: {n: ['вересень +25 500'], d: []}},
                                    {name: 'Звичайний', bs: 25500, fs: 0, be: 0, fe: 0, overN: 0, overD: 0}]).map((x) => x.name + ':' + x.items.map((i) => i.kind + ' ' + i.prepaid).join(','))},
        {list: ['Капітонов Ян:навч. 25500']});
  const sendSrc = fn('_bankSend'), dailySrc = fn('_bankDaily__impl'), sweepSrc = fn('_bankBudgetSweep'), linkSrc = fn('_bankLinkFromMessage');
  check('банк: вимикач до перегляду', 'не ввімкнена група — нічого не надсилаємо (відправка, 16:00, ⚠️, підтвердження підключення)',
        {send: /if \(!_bankLive\(loc\)\) return/.test(sendSrc), daily: /_bankLive\(l\)/.test(dailySrc), sweep: /!_bankLive\(loc\)/.test(sweepSrc), link: /if \(_bankLive\(loc\)\) _invSend/.test(linkSrc)},
        {send: true, daily: true, sweep: true, link: true});
  check('банк: відкладений боржник', 'Капітонов (Осокорки) — не в списку до рішення CFO; інші лишаються',
        {list: K._bankHoldFilter([{name: 'Капітонов Ян'}, {name: 'Шевчук Саша'}], ['Капітонов  Ян']).map((x) => x.name)}, {list: ['Шевчук Саша']});
  check('банк: оплатили', 'порожньо — не пишемо; разом унизу',
        {empty: K._bankPaidTexts('Позняки', [], '11.11').length,
         last: K._bankPaidTexts('Позняки', [{child: 'Хмельницький Іван', amount: 21500, kind: 'навчання', date: '10.11.2026'}, {child: 'Драгунов Іван', amount: 2250, kind: 'допи', date: '10.11.2026'}], '11.11')[0].split('\n').pop()},
        {empty: 0, last: '<b>Разом: 23 750 ₴</b>'});
}

// v7.538: сума й стаття заявки — підпис директора головний (#196, #101, #288 Осокорки), комунальні — «До сплати»
if (src.indexOf('\nfunction _invPickAmount(') >= 0){
  const Cv = ['INV_CAP_DIFF','INV_AMOUNT_EPS','INV_DAYS_BEFORE','INV_DAYS_AFTER','INV_DATE_PLAUSIBLE','INV_CREATED_BEFORE','INV_CREATED_AFTER','INV_VAT_RATE','INV_DUP_DAYS'].map(varLine)
    .concat([varLine('INV_UTIL_RE'), varBlock('INV_CAT_KEYWORDS'), varLine('INV_CAP_NOISE_RE'), varLine('INV_CAP_PLACE_RE')]);
  const C = new Function(Cv.concat(['trim','_invNum','_invSplitNum','_invPickAmount','_invCaptionClean','_invCaptionAmount','_invCaptionNums','_invAmountDecide','_invCaptionCategory',
    '_invDateMs','_invNumKey','_invPurposeHasNumber','_invEdrMatch','_invEdr7','_invPurposeHasNumberX','_invPurposeInvNumbers','_invPurposeOtherNumber',
    '_invAmtEq','_invNumsCompatible','_invWindowOk','_invIsDup','_invMatchCore'].map(fn)).join('\n') +
    '\nreturn {_invPickAmount, _invCaptionAmount, _invCaptionNums, _invAmountDecide, _invCaptionCategory, _invMatchCore};')();
  const yasno = 'ТОВ «ЯСНО+» ЄДРПОУ 45179093\nФактичний рахунок № Y100001749331 за вересень 2026\nОсобовий рахунок 77765531012\n' +
                'Спожито електроенергії 2 136 кВт·год\nЗаборгованість на початок періоду: 202 630,09\nНараховано за період: 9 852,94\nДо сплати: 9 852,94 грн';
  check('заявки: комунальний — «До сплати»', 'ЯСНО: не найбільше число (202 630,09), а «До сплати» 9 852,94; без «до сплати» — 0',
        {a: C._invPickAmount(yasno).amount, b: C._invPickAmount('Рахунок за електроенергію\nСпожито 2 136 кВт·год\nЗаборгованість 202 630,09').amount},
        {a: 9852.94, b: 0});
  check('заявки: звичайний рахунок', '«Всього до сплати» має пріоритет над «Всього» (без ПДВ)',
        {a: C._invPickAmount('Рахунок № 188851\nВсього: 1 664,83\nПДВ: 332,97\nВсього до сплати: 1 997,80').amount}, {a: 1997.8});
  check('заявки: сума з підпису', '«КАнцелярія 1998 грн» → 1998; «796грн + 260грн» → 1056; без суми → 0',
        {a: C._invCaptionAmount('Осокорки, садочок\nПапір А4 та конверти\nКАнцелярія 1998 грн', []),
         b: C._invCaptionAmount('Школа Осокорки\nГосптовари 796грн + канцтовари 260грн', []),
         c: C._invCaptionAmount('Осокорки, садочок\nЕлектроенергія, жовтень', [])}, {a: 1998, b: 1056, c: 0});
  check('заявки: підпис головний', '#196: 1998 vs OCR 31 450 → 1998 і питання; 1056 vs 1055,20 → точна OCR; без підпису → OCR',
        {a: C._invAmountDecide(1998, 31450), b: C._invAmountDecide(1056, 1055.2), c: C._invAmountDecide(0, 9852.94), d: C._invAmountDecide(1000, 1040)},
        {a: {amount: 1998, ask: true, src: 'підпис'}, b: {amount: 1055.2, ask: false, src: 'ocr≈підпис'}, c: {amount: 9852.94, ask: false, src: 'ocr'}, d: {amount: 1000, ask: false, src: 'підпис'}});
  const YS = JSON.parse(fs.readFileSync(path.join(__dirname, 'fixtures_yasno.json'), 'utf8'));
  check('заявки: справжні PDF ЯСНО', '«До сплати по договору» → прописом → рядок «сума · ПДВ · разом»; не склеєні дати «2026 30.09»',
        {a: C._invPickAmount(YS.yasno101).amount, b: C._invPickAmount(YS.yasno288).amount}, {a: 9852.94, b: 16215.45});
  const LL = ['Школа 228', 'Осокорки', 'Школа Осокорки', 'Теремки', 'Тичини'];
  check('заявки: хибні суми підпису (аудит з 01.10)', 'локація «228», «Теремки2», дата «20.10» — не сума; «8 088, 65 грн» → 8088.65',
        {a: C._invCaptionAmount('Їдальня 228', [{name: 'Їдальня', amount: 228}], LL), b: C._invCaptionAmount('Теремки2 / Господарські', [], LL),
         c: C._invCaptionAmount('Осокорки, садочок / Театр / Вистава відбулася 20.10', [{name: 'Вистава відбулася', amount: 20.1}], LL),
         d: C._invCaptionAmount('Осокорки, садочок\nГосподарські товари 8 088, 65 грн', [], LL)}, {a: 0, b: 0, c: 0, d: 8088.65});
  check('заявки: число в підписі = OCR', '«Ашан. 2697 (Сніданки 2046, …)»: розбивка неповна, але 2697 = OCR 2697,01 → без питання',
        {a: C._invAmountDecide(2557, 2697.01, C._invCaptionNums('Ашан. 2697 (Сніданки 2046, персонал 245, незапл.витр. 266). Тичини.', LL))},
        {a: {amount: 2697.01, ask: false, src: 'ocr=підпис'}});
  check('заявки: Кругла #198 і довгі номери', 'Σ «120 000 + 6350 грн» = OCR 126 350 → без питання; номер рахунку з «грн» — не сума',
        {a: C._invAmountDecide(6350, 126350, C._invCaptionNums('Львів Кругла / оренда 120 000 за жовтень / комунальні 6350 грн за вересень', LL)),
         b: C._invCaptionAmount('рахунки за жовтень 2026!! Утримання UA807020100000051800 грн', [], LL)},
        {a: {amount: 126350, ask: false, src: 'ocr=підпис'}, b: 0});
  const OS = ['ОРЕНДА','КОМУНАЛЬНІ ПОСЛУГИ','СНІДАНКИ','КУХНЯ','Вода','КАНЦТОВАРИ','ІГРАШКИ','ХОЗ.ТОВАРИ','ПРАЛЬНЯ','РЕМОНТ'], LOCS = ['Осокорки','Школа Осокорки','Позняки'];
  check('заявки: стаття з підпису', '«Добрий день!», локація, «садочок» — ігнор; електроенергія → КОМУНАЛЬНІ ПОСЛУГИ; папір → КАНЦТОВАРИ',
        {a: C._invCaptionCategory('Добрий день!\nОсокорки, садочок\nЕлектроенергія', OS, LOCS), b: C._invCaptionCategory('Осокорки, садочок\nЕлектроенергія, жовтень', OS, LOCS),
         c: C._invCaptionCategory('Осокорки, садочок\nПапір А4 та конверти\nКАнцелярія 1998 грн', OS, LOCS), d: C._invCaptionCategory('Доброго ранку!\nОсокорки', OS, LOCS)},
        {a: 'КОМУНАЛЬНІ ПОСЛУГИ', b: 'КОМУНАЛЬНІ ПОСЛУГИ', c: 'КАНЦТОВАРИ', d: ''});
  const r196 = C._invMatchCore([{id: '196', amount: 1998, edrpou: '36265773', number: '188851', loc: 'Осокорки', invDateMs: new Date(2026, 10, 5).getTime(), createdMs: new Date(2026, 10, 5).getTime()}],
    [{date: '06.11.2026', amount: 1997.8, edrpou: '36265773', purpose: 'ОПЛАТА ЗА ТОВАР ЗГІДНО РАХ. № 188851 від 05.11.2026 р. Р. У сумі 1664.83 грн., ПДВ - 20 % 332.97 грн.', ref: 'JBKLQB6OF4KBEP.1'}], [], {});
  const rNo = C._invMatchCore([{id: 'x', amount: 1998, edrpou: '36265773', number: '188999', invDateMs: new Date(2026, 10, 5).getTime(), createdMs: new Date(2026, 10, 5).getTime()}],
    [{date: '06.11.2026', amount: 1997.8, edrpou: '36265773', purpose: 'ОПЛАТА ЗА ТОВАР', ref: 'R2'}], [], {});
  check('заявки: округлена сума з підпису', '#196: 1998 ↔ платіж 1 997,80 з № 188851 → оплачено; без № у призначенні — ні',
        {a: r196.good.length, b: rNo.good.length}, {a: 1, b: 0});
}

// v7.542: «🔎 Звірка надходжень» у звіті CFO — виписка vs Payment, платник ≠ підписант, не та колонка, зайве
if (src.indexOf('\nfunction _recDayCheck(') >= 0){
  const R = new Function(['_APOSTROPHES','INV_TG_MAX'].map(varLine).concat(['trim','_lat2cyr','_nameFold','_nameFoldTight','_journalNormName','_dopAnKid','_levSalary','_surnameSim',
    '_htmlEsc','_invMoneyUa','_bankMoney','_bankWords','_recIsTransit','_recNotClient','_recPayerName','_recSamePerson','_recSurnameStem','_recSameFamily','_recPurposeKids','_recDayCheck','_recDayText'].map(fn)).join('\n') +
    '\nreturn {_recPayerName, _recSamePerson, _recDayCheck, _recDayText, _recNotClient, _recSameFamily};')();
  const IB = 'UA813052990000026006016208600', L = {[IB]: 'Осокорки'};
  // реальні платежі 09.11 (Осокорки): Боднік Віка 19 500 і 5 000 «за навчання», транзитний рахунок еквайрингу
  const p1 = {iban: IB, date: '09.11.2026', amount: 19500, ref: 'D3K4QB9AADHJ7E.P', purpose: 'Оплата за навчання, дог.№ 75-32-74, Осокорки садочок, Боднік Сергій Віталійович'};
  const p2 = {iban: IB, date: '09.11.2026', amount: 5000, ref: 'D3P6QB9AADI1EC.P', purpose: 'Оплата за навчання, дог.№ 75-32-74, Осокорки садочок, Боднік Сергій Віталійович'};
  [p1, p2].forEach((p) => { p.payer = R._recPayerName({counterparty: 'K3_транз.рахунок платежi BP 3504039', purpose: p.purpose}); });
  check('звірка: платник з транзитного платежу', 'з кінця призначення', {payer: p1.payer}, {payer: 'Боднік Сергій Віталійович'});
  const bod = (ref, amt) => ({iban: IB, loc: 'Осокорки', date: '09.11.2026', ref: ref, child: 'Боднік Віка', amount: amt, kind: 'навчання', budNavch: 26300, budDop: 350,
                              signer: 'Боднік Ірен Сергіївна', entryFee: 19500, entryPaid: 15000, contract: 24300});
  const rB = R._recDayCheck([p1, p2], [bod(p1.ref, 19500), bod(p2.ref, 5000)], L)[0];
  check('звірка: Боднік Віка — вступний у навчанні', '19 500 = вступному з картки; 5 000 — вступний не доплачено (15 000 з 19 500)',
        {col: rB.column.map((c) => c.amount + ':' + (/= вступному/.test(c.why) ? 'вступний' : /можливо вступний/.test(c.why) ? 'можливо' : '?'))}, {col: ['19500:вступний', '5000:можливо']});
  check('звірка: платник з родини — не ⚠️', 'платив тато (Боднік Сергій), підписант — мама (Боднік Ірен): прізвище те саме', {n: rB.payer.length}, {n: 0});
  check('звірка: суми зійшлися', 'виписка = Payment, нічого не бракує', {s: rB.stmtSum, l: rB.logSum, miss: rB.missing.length, extra: rB.extra.length}, {s: 24500, l: 24500, miss: 0, extra: 0});
  const sig = R._recSamePerson('Новак Вероніка Василівна', 'Новак Вероніка'), sig2 = R._recSamePerson('Іваненко О.', 'Іваненко Ольга Петрівна');
  check('звірка: той самий платник', 'прізвище + ім\'я (або ініціал) — ок', {a: sig, b: sig2}, {a: true, b: true});
  // не та колонка: сума = бюджету навчання, а внесено в допи; немає в Payment; зайве; двічі
  const pA = {iban: IB, date: '10.11.2026', amount: 25500, ref: 'A1', payer: 'Курінна Олена', purpose: 'за навчання'};
  const pB = {iban: IB, date: '10.11.2026', amount: 5500, ref: 'B1', payer: 'Іванова Марія', purpose: 'гуртки'};
  const pC = {iban: IB, date: '10.11.2026', amount: 3000, ref: 'C1', payer: 'Пархоменко Ольга', purpose: 'за навчання'};
  const eA = {iban: IB, loc: 'Осокорки', date: '10.11.2026', ref: 'A1', child: 'Курінний Михайло', amount: 25500, kind: 'допи', budNavch: 25500, budDop: 3810, signer: 'Курінна Олена', entryFee: 0};
  const eC1 = {iban: IB, loc: 'Осокорки', date: '10.11.2026', ref: 'C1', child: 'Пархоменко Єва', amount: 3000, kind: 'навчання', budNavch: 25500, budDop: 0, signer: 'Пархоменко Ольга', entryFee: 0};
  const eX = {iban: IB, loc: 'Осокорки', date: '10.11.2026', ref: 'ZZ', child: 'Шевчук Саша', amount: 21200, kind: 'навчання', budNavch: 21200, budDop: 0, signer: '', entryFee: 0};
  const r2 = R._recDayCheck([pA, pB, pC], [eA, eC1, Object.assign({}, eC1), eX], L)[0];
  check('звірка: не та колонка', '25 500 = бюджету навчання, внесено в допи', {col: r2.column.map((c) => c.child)}, {col: ['Курінний Михайло']});
  check('звірка: немає в Payment', 'платіж Іванової 5 500 не рознесено', {miss: r2.missing.map((m) => m.payer + ' ' + m.amount)}, {miss: ['Іванова Марія 5500']});
  check('звірка: зайве і двічі', 'референсу ZZ немає у виписці; C1 рознесено двічі (6 000 з 3 000)',
        {extra: r2.extra.map((x) => x.child + ' ' + x.amount).sort()}, {extra: ['Пархоменко Єва 3000', 'Шевчук Саша 21200']});
  const pK = {iban: IB, date: '10.11.2026', amount: 31000, ref: 'K1', payer: 'Акименко Віра', purpose: 'за навчання'};
  const eK = {iban: IB, loc: 'Школа Осокорки', date: '10.11.2026', ref: 'K1', child: 'Акименко Віра', amount: 31000, kind: 'навчання', budNavch: 31000, budDop: 0, signer: 'Акименко Катерина', entryFee: 0};
  check('звірка: платник = дитина', 'ім\'я дитини з призначення — не «платник ≠ підписант»; повернення від постачальника — не клієнт',
        {payer: R._recDayCheck([pK], [eK], L)[0].payer.length, refund: R._recNotClient({purpose: 'Повернення надлишк. коштів за товар згід.листа б/н від 10.11.2026р.'})}, {payer: 0, refund: true});
  const fam = (payer, signer, child) => R._recDayCheck([{iban: IB, date: '10.11.2026', amount: 100, ref: 'F', payer: payer, purpose: ''}],
    [{iban: IB, loc: 'Кругла', date: '10.11.2026', ref: 'F', child: child, amount: 100, kind: 'навчання', budNavch: 0, budDop: 0, signer: signer, entryFee: 0}], L)[0].payer.length;
  check('звірка: прізвище платника (реальні 10.11)', 'Тітов/Тітова, Нагачевський/Нагачевська, Мельничук, Сук (як у дитини) — без ⚠️; Курко ↔ Цьома, Гусєв ↔ Біканасова — ⚠️',
        {titov: fam('Тітов Денис Олександрович', 'Тітова Марина', 'Тітов Кирило'), nag: fam('Нагачевський Євген Богданович', 'Нагачевська Галина Іванівна', 'Нагачевська Софія (Preschool)'),
         mel: fam('Мельничук Вікторія Романівна', 'Мельничук Леся Сергіївна', 'Мельничук Дарина'), suk: fam('Сук Святослав Анатолійович', 'Венедиктова Ольга', 'Сук Лера'),
         kurko: fam('Курко Галина Андріївна', 'Цьома Андріана Станіславівна', 'Цьома Захар (Findi-Ki)'), gus: fam('ГУСЄВ НІКІТА ВАЛЕРІЙОВИЧ', 'Біканасова Жанна', 'Біканасова Аріна (4 клас)')},
        {titov: 0, nag: 0, mel: 0, suk: 0, kurko: 1, gus: 1});
  // v7.543: дитина в призначенні ↔ дитина, на яку рознесено; без дитини — платник ↔ підписант / прізвище дитини
  const KIDS = ['Тітов Кирило', 'Русанов Олександр', 'Боднік Віка', 'Боднік Кирил', 'Цьома Захар', 'Сук Лера', 'Комаренко Саша', 'Комаренко Єгор'];
  const chk = (purpose, payer, child, signer) => R._recDayCheck([{iban: IB, date: '11.11.2026', amount: 100, ref: 'W', payer: payer, purpose: purpose}],
    [{iban: IB, loc: 'Кар\'єрна', date: '11.11.2026', ref: 'W', child: child, amount: 100, kind: 'навчання', budNavch: 0, budDop: 0, signer: signer, entryFee: 0, kids: KIDS}], L)[0];
  const w1 = chk('Оплата за навчання, 61-04-09, Тітов Кирило Денисович, листопад 2026', 'Тітов Денис', 'Тітов Кирило', 'Тітова Марина');
  const w2 = chk('Оплата за навчання, 61-04-09, Тітов Кирило Денисович, листопад 2026', 'Тітов Денис', 'Русанов Олександр', 'Русанова Ірина');
  const w3 = chk('Оплата за організацію додаткових занять (гуртків) Боднік Віка, жовтень 2026', 'Курко Галина', 'Боднік Віка', 'Боднік Ірен');
  const w4 = chk('Комаренко Єгор', 'Комаренко Антоніна', 'Комаренко Саша', 'Комаренко Антоніна');
  const w5 = chk('Оплата за навчання', 'Курко Галина Андріївна', 'Цьома Захар (Findi-Ki)', 'Цьома Андріана Станіславівна');
  const w6 = chk('Оплата за навчання', 'Нагачевський Євген', 'Нагачевська Софія (Preschool)', 'Нагачевська Галина');
  const w7 = chk('Оплата за навчання Цьома Захар', 'Курко Галина', 'Цьома Захар (Findi-Ki)', 'Цьома Андріана');
  check('звірка: дитина в призначенні', 'збіг — ок і платника не перевіряємо; інша дитина (у т.ч. брат/сестра) — ⚠️ «не на ту дитину»',
        {ok: w1.wrongKid.length + w1.payer.length, wrong: w2.wrongKid.map((x) => x.named), payerSkipped: w3.payer.length + w3.wrongKid.length, sibling: w4.wrongKid.map((x) => x.named), paren: w7.wrongKid.length},
        {ok: 0, wrong: ['Тітов Кирило'], payerSkipped: 0, sibling: ['Комаренко Єгор'], paren: 0});
  check('звірка: дитини в призначенні немає', 'платник з іншим прізвищем — ⚠️ «платник не підписант»; Нагачевський/Нагачевська — ні',
        {kurko: w5.payer.length, nag: w6.payer.length, text: /платник не підписант/.test(R._recDayText([w5], '11.11'))}, {kurko: 1, nag: 0, text: true});
  const ok = R._recDayCheck([pC], [eC1], L);
  check('звірка: усе зійшлося', 'один рядок ✅', {t: R._recDayText(ok, '10.11')}, {t: '✅ Надходження за 10.11 зійшлися'});
}

console.log('\n' + pass + ' пройдено, ' + fail + ' з помилкою — ' + path.basename(SRC));
process.exit(fail ? 1 : 0);
