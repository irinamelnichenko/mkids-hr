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
  const queueOnly = (src.match(/(?<!function )_invQueuePaid\([^;]*\)/g) || []).every((x) => /✅ Оплачено|_invPaidText(Full)?\(/.test(x));   // v7.569: текст через _invPaidText
  check('звіт CFO: група директорів', 'у групу — лише «✅ Оплачено», звіт іде тільки на адресата з _invCfoTarget',
        {leak: leak.length, cfoTo: cfoTo, queueOnly: queueOnly, privateHook: /chat\.type === 'private'\) return _invCfoLinkFromMessage/.test(src)},
        {leak: 0, cfoTo: true, queueOnly: true, privateHook: true});
}

// v7.536: групи «Банк — <локація>» — ізоляція, боржники (Крошкін, Абраменко), вступний частинами, заявка понад бюджет
if (src.indexOf('\nfunction _bankChatFor(') >= 0){
  const Bv = ['_APOSTROPHES','INV_TG_MAX','BANK_DEBT_MIN','BANK_DEBT_DAYS','BANK_GROUPS_PROP'].concat(src.indexOf('\nvar BANK_DEBT_LIST_MIN') >= 0 ? ['BANK_DEBT_LIST_MIN'] : []).concat(src.indexOf('\nvar BANK_SHARED') >= 0 ? ['BANK_SHARED','BANK_LOC_TITLE'] : []).map(varLine)
    .concat([varBlock('BANK_INV_CAT_RULES')]).concat(src.indexOf('\nvar BANK_CAT_SHORT ') >= 0 ? [varBlock('BANK_CAT_SHORT')] : []);
  const Bf = ['_lat2cyr','_nameFold','_nameFoldTight','_journalNormName','_bankWords','_dopAnKid','_levSalary','_surnameSim','_htmlEsc','_invMoneyUa','_bankGroupsParse',...(src.indexOf('\nfunction _bankOwnersOk(') >= 0 ? ['_bankSharedSet','_bankOwnersOk','_bankOwners','_bankLocsOfChat','_bankTopicOwner','_bankLocTitle'] : []),'_bankChatFor','_bankLocOfChat','_bankMoney',
              '_bankNormCat','_bankInvCat','_bankDebtDayFor','_bankDebtorsCalc','_bankEntryCalc','_bankOpexCalc','_bankInvOver','_bankOverNew','_bankChunks',
              ...(src.indexOf('\nfunction _bankInt(') >= 0 ? ['_bankInt','_bankPadR','_bankPadL','_bankCode','_bankKids','_bankShortCat','_bankDebtGroups','_bankCarried'] : []),
              '_bankPaidTexts','_bankDebtorsTexts','_bankOpexTexts','_bankInvOverText','_bankOverTexts','_bankCommentTarget','_bankSuspectDebts','_bankHoldFilter'].map(fn);
  const K = new Function(Bv.concat(Bf).join('\n') + '\nreturn {' + ['_bankChatFor','_bankLocOfChat','_bankInvCat','_bankDebtDayFor','_bankDebtorsCalc','_bankEntryCalc',
    '_bankOpexCalc','_bankInvOver','_bankOverNew','_bankPaidTexts','_bankDebtorsTexts','_bankOpexTexts','_bankInvOverText','_bankCommentTarget','_dopAnKid','_bankSuspectDebts','_bankHoldFilter'].concat(src.indexOf('\nfunction _bankOwnersOk(') >= 0 ? ['_bankLocsOfChat','_bankTopicOwner','_bankLocTitle'] : []).join(',') + '};')();
  const INVG = '-5142518705', CFOC = '123456789';
  {
    const pk = K._dopAnKid('Пархоменко Єва'), fx = {}; fx[pk] = {name: 'Пархоменко Єва', months: [0, 0, 0, 0, 0, 0, 0, 0, 0, 3000, 0, 0]};
    const cards = [{name: 'Пархоменко Єва', fee: 25500, contract: '2026-09-30'}], adj = {}; adj[pk] = -3000;
    check('банк: поправка CFO до вступного (Пархоменко, еталон 18.11)', '3 000 у «Факт вступ» жовтня — оплата навчання: без поправки борг 22 500, з поправкою −3 000 — 25 500; Payment не змінюється',
          {before: K._bankEntryCalc(cards, fx, 2026, '2026-11-18').map((x) => x.fee - x.paid), after: K._bankEntryCalc(cards, fx, 2026, '2026-11-18', adj).map((x) => x.fee - x.paid),
           held: K._bankHoldFilter(K._bankEntryCalc(cards, fx, 2026, '2026-11-18', adj), ['Пархоменко Єва']).length},
          {before: [22500], after: [25500], held: 0});
  }
  if (K._bankLocsOfChat){
    const SG = {"Кар'єрна": '-1001', "Школа Кар'єрна": '-1001', 'Позняки': '-1002', 'Тичини': '-1003', 'Пуща': '-1003'};
    check('банк: спільна група Кар\'єрної', 'садок і школа — одна група, обидві отримують; теми спільні (ключ — садок); заголовки «Кар\'єрна, садок / школа»; інші пари — ні',
          {kar: K._bankChatFor(SG, "Кар'єрна", INVG, CFOC), sch: K._bankChatFor(SG, "Школа Кар'єрна", INVG, CFOC), loc: K._bankLocOfChat(SG, '-1001'), locs: K._bankLocsOfChat(SG, '-1001'),
           topic: K._bankTopicOwner(SG, "Школа Кар'єрна"), title: [K._bankLocTitle("Кар'єрна"), K._bankLocTitle("Школа Кар'єрна"), K._bankLocTitle('Позняки')],
           other: [K._bankChatFor(SG, 'Тичини', INVG, CFOC), K._bankLocOfChat(SG, '-1003')], pozn: K._bankChatFor(SG, 'Позняки', INVG, CFOC)},
          {kar: '-1001', sch: '-1001', loc: "Кар'єрна", locs: ["Кар'єрна", "Школа Кар'єрна"], topic: "Кар'єрна", title: ["Кар'єрна, садок", "Кар'єрна, школа", 'Позняки'],
           other: ['', ''], pozn: '-1002'});
  }
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
  check('банк: текст боржника з коментарем', 'кілька позицій: дитина й сума, під нею навчання · допи (без «за жовтень»); 💬',
        {line: dt.split('\n').filter((l) => /^<code>/.test(l)).slice(0, 3), cmt: /💬 <i>обіцяли до 15-го/.test(dt)},
        {line: ['<code>Абраменко Тимофій  38 440</code>', '<code>  навчання         27 500</code>', '<code>  допи             10 940</code>'], cmt: true});
  // v7.546: групування й блоки (реальні Осокорки 16.11 — форма)
  const dg = K._bankDebtorsTexts('Осокорки', 'листопад', [
      {name: 'Боднік Віка', navch: {debt: 25500, carried: 25500}, dop: null, total: 25500},
      {name: 'Шевчук Саша', navch: {debt: 4200, carried: 0}, dop: {debt: 3810, carried: 1000}, total: 8010},
      {name: 'Іванова Марія', navch: null, dop: {debt: 7, carried: 0}, total: 7},
      {name: 'Курінний Михайло', navch: {debt: 31000, carried: 6000}, dop: null, total: 31000}],
    [{name: 'Боднік Віка', paid: 15000, fee: 19500}, {name: 'Драгунов Іван', paid: 20500, fee: 35500}], '16.11', 'жовтень')[0].split('\n');
  check('банк: боржники по дітях', 'кожна дитина раз; блоки в порядку; за сумою; без «(… за жовтень)» (v7.553); дрібні суми теж; вступний «борг  (сплачено / сума)»; кілька позицій — рядок на позицію; разом',
        {heads: dg.filter((l) => /^<b>(Винні|Тільки)/.test(l)), sub: dg[1], kur: dg.find((l) => /Курінний/.test(l)), ent: dg.find((l) => /Драгунов/.test(l)),
         small: dg.some((l) => /Іванова Марія\s+7<\/code>/.test(l)), tot: dg.find((l) => /Разом боргу/.test(l)), bod: dg.filter((l) => /Боднік/.test(l)).length, bodPos: dg.slice(dg.findIndex((l) => /Боднік/.test(l)) + 1, dg.findIndex((l) => /Боднік/.test(l)) + 3)},
        {heads: ['<b>Винні за кілька позицій</b> · 2 дитини · 38 010', '<b>Тільки навчання</b> · 1 дитина · 31 000', '<b>Тільки допи</b> · 1 дитина · 7', '<b>Тільки вступний</b> · 1 дитина · 15 000'],
         sub: '<i>на 16.11 · оплата до 5-го</i>', kur: '<code>Курінний Михайло  31 000</code>', ent: '<code>Драгунов Іван     15 000  (20 500 / 35 500)</code>',
         small: true, tot: '<b>Разом боргу: 84 017 ₴ · 5 дітей</b>', bod: 1,
         bodPos: ['<code>  навчання        25 500</code>', '<code>  вступний         4 500  (15 000 / 19 500)</code>']});
  const ox = K._bankOpexTexts('Осокорки', 'листопад', K._bankOpexCalc([{name: 'КАНЦТОВАРИ', budget: 2000, spent: 1998}, {name: 'РЕМОНТ', budget: 4000, spent: 3000},
      {name: 'НЕЗАПЛАНОВАНІ ВИТРАТИ', budget: 3000, spent: 900}, {name: 'СНІДАНКИ', budget: 15000, spent: 6000}, {name: 'ІГРАШКИ', budget: 1500, spent: 0}, {name: 'Вода', budget: 800, spent: 0}],
      {'РЕМОНТ': 2200, '': 6362.51}), '16.11', 53.3)[0].split('\n');
  check('банк: OPEX новий вигляд', 'черга = витрачено; 🔴 −сума, 🟡 > 80%, 🟢 за % спадно; ⚪ одним рядком; разом і без статті',
        {rows: ox.filter((l) => /^<code>/.test(l)), none: ox.find((l) => /Без витрат/.test(l)), sub: ox[1], tot: ox.find((l) => /Разом/.test(l)), nc: ox[ox.length - 1]},
        {rows: ['<code>Ремонт      5 200 /  4 000  −1 200</code>', '<code>Канцтовари  1 998 /  2 000    100%</code>', '<code>Сніданки    6 000 / 15 000     40%</code>', '<code>Незаплан.     900 /  3 000     30%</code>'],
         none: '⚪ <b>Без витрат:</b> Іграшки, Вода', sub: '<i>на 16.11 · минуло 53% місяця</i>', tot: '<b>Разом: витрачено 14 098 з 26 300 · вільно 12 202 ₴</b>', nc: '⚠️ 6 363 ₴ у рахунках без статті'});
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

// v7.547: тестовий режим групи «Банк» і зведення відгуків
if (src.indexOf('\nfunction _bankTestText(') >= 0){
  const T = new Function(['INV_TG_MAX','BANK_TEST_LINE','BANK_TEST_TYPES'].map(varLine).concat(['_htmlEsc','_bankChunks','_bankTestText','_bankFeedbackTexts'].map(fn)).join('\n') +
    '\nreturn {_bankTestText, _bankFeedbackTexts};')();
  check('банк: тестовий режим', 'оплатили, боржники, ліміти й фонд ЗП — з рядком-позначкою; перевищення, ⚠️, питання про статтю — не надсилаємо; повний режим без змін',
        {paid: T._bankTestText('test', 'paid', '💰 x').split('\n')[0], debt: /^🧪/.test(T._bankTestText('test', 'debt', 'y')),
         blocked: ['over', 'invOver', 'catQ'].map((k) => T._bankTestText('test', k, 'z')), full: T._bankTestText('full', 'opex', 'z'),
         weekly: ['opex', 'salary'].map((k) => /^🧪/.test(T._bankTestText('test', k, 'z')))},
        {paid: '🧪 Тестовий режим. Перевірте список і напишіть у відповідь, що не так.', debt: true, blocked: [null, null, null], full: 'z', weekly: [true, true]});
  const fb = T._bankFeedbackTexts([{loc: 'Позняки', child: 'Абраменко Тимофій', who: 'Олена Директор', text: 'оплатили готівкою 12.11', kind: '🔴 Боржники'},
    {loc: 'Позняки', child: '', who: 'Олена Директор', text: 'список ок', kind: '💰 Оплатили'}, {loc: 'Пуща', child: '', who: 'Ірина', text: '<немає Шевчук>', kind: '🔴 Боржники'}], '16.11');
  check('банк: зведення відгуків о 18:00', 'по локаціях; дитина — якщо вказано; автор; текст екрановано; порожньо — нічого',
        {head: fb[0].split('\n')[0], lines: fb[0].split('\n').slice(1), empty: T._bankFeedbackTexts([], '16.11').length},
        {head: '🧪 <b>Відгуки тестових груп «Банк»</b> — 16.11 · 3', lines: ['', '<b>Позняки</b> · 2', '• <b>Абраменко Тимофій</b> · Олена Директор <i>(🔴 Боржники)</i>: оплатили готівкою 12.11',
          '• Олена Директор <i>(💰 Оплатили)</i>: список ок', '', '<b>Пуща</b> · 1', '• Ірина <i>(🔴 Боржники)</i>: &lt;немає Шевчук&gt;'], empty: 0});
}

// v7.548: теми (Topics) у групах «Банк» і понеділковий «💼 Фонд ЗП»
if (src.indexOf('\nfunction _bankTopicFor(') >= 0){
  const P = new Function(['INV_TG_MAX','BANK_TOPICS','BANK_TOPIC_OF'].map(varLine).concat(['_htmlEsc','_bankChunks','_bankInt','_bankPadR','_bankPadL','_bankCode',
    '_bankTopicFor','_bankTopicMissing','_bankTopicIcon','_bankRealReply','_bankThreadOf','_bankReplyExtra','_bankThreadOk',
    '_salaryIsSkippedRow','_salaryIsSubtotalRow','_classifyAllSalaryRows','_bankSalaryCalc','_bankSalaryTexts'].map(fn)).join('\n') +
    '\nreturn {BANK_TOPICS, _bankTopicIcon, _bankTopicFor, _bankTopicMissing, _bankRealReply, _bankReplyExtra, _bankThreadOk, _bankSalaryCalc, _bankSalaryTexts};')();
  const tp = {chat: '-1003', paid: 11, debt: 12, opex: 13, salary: 14};
  check('банк: тема для типу', 'оплатили → Оплати; боржники → Боржники; OPEX, ⚠️, 🔴, питання статті → Ліміти; ЗП → Salary; без тем / невідомий тип — загальний чат (0)',
        {by: ['paid', 'debt', 'opex', 'invOver', 'over', 'catQ', 'salary', 'other'].map((k) => P._bankTopicFor(tp, k)), none: P._bankTopicFor(null, 'paid')},
        {by: [11, 12, 13, 13, 13, 13, 14, 0], none: 0});
  check('банк: чого бракує для тем', 'теми вимкнені; бот без права «Керування темами»; теми вже створені — право не потрібне; власник групи — можна',
        {off: /теми/.test(P._bankTopicMissing({is_forum: false}, null, true)),
         noRight: /Керування темами/.test(P._bankTopicMissing({is_forum: true}, {status: 'administrator', can_manage_topics: false}, true)),
         member: /Керування темами/.test(P._bankTopicMissing({is_forum: true}, {status: 'member'}, true)),
         haveAll: P._bankTopicMissing({is_forum: true}, {status: 'member'}, false),
         ok: P._bankTopicMissing({is_forum: true}, {status: 'administrator', can_manage_topics: true}, true)},
        {off: true, noRight: true, member: true, haveAll: '', ok: ''});
  const stk = [{emoji: '💵', custom_emoji_id: '501'}, {emoji: '💰', custom_emoji_id: '502'}, {emoji: '❗️', custom_emoji_id: '503'}, {emoji: '📈', custom_emoji_id: '504'}];
  check('банк: назви й іконки тем', 'назви без емодзі; іконка — стікер з найближчим емодзі (💰 раніше 💵; ❗ з варіаційним селектором; 📊 немає → 📈); підходящого немає → \'\' (лише колір)',
        {names: P.BANK_TOPICS.map((x) => x[1]), icons: P.BANK_TOPICS.map((x) => P._bankTopicIcon(stk, x[2])), colors: P.BANK_TOPICS.every((x) => [7322096, 16766590, 13338331, 9367192, 16749490, 16478047].includes(x[3]))},
        {names: ['Оплати', 'Боржники', 'Ліміти', 'Salary'], icons: ['502', '503', '504', ''], colors: true});
  const root = {message_id: 12, from: {is_bot: true}, forum_topic_created: {name: '🔴 Боржники'}};
  const botMsg = {message_id: 500, from: {is_bot: true}};
  check('банк: відповідь у темі', 'звичайне повідомлення в темі — не reply (формально відповідає на початок теми); справжній reply — так; відповідь бота — у ту саму тему',
        {plain: P._bankRealReply({message_id: 600, is_topic_message: true, message_thread_id: 12, reply_to_message: root}),
         plain2: P._bankRealReply({message_id: 601, is_topic_message: true, message_thread_id: 12, reply_to_message: {message_id: 12, from: {is_bot: true}}}),
         real: P._bankRealReply({message_id: 602, is_topic_message: true, message_thread_id: 12, reply_to_message: botMsg}).message_id,
         general: P._bankRealReply({message_id: 603, reply_to_message: botMsg}).message_id,
         extra: P._bankReplyExtra({message_id: 602, is_topic_message: true, message_thread_id: 12}), extraGen: P._bankReplyExtra({message_id: 603})},
        {plain: null, plain2: null, real: 500, general: 500, extra: {reply_to_message_id: 602, message_thread_id: 12}, extraGen: {reply_to_message_id: 603}});
  check('банк: коментарі лише в «Боржниках»', 'з темами — коментар у Боржниках так, в Оплатах / Лімітах / Загальній — ні; група без тем — як раніше',
        {debt: P._bankThreadOk(tp, 12, ['debt']), paid: P._bankThreadOk(tp, 11, ['debt']), opex: P._bankThreadOk(tp, 13, ['debt']), general: P._bankThreadOk(tp, 0, ['debt']),
         test: P._bankThreadOk(tp, 11, ['paid', 'debt']), noTopics: P._bankThreadOk(null, 0, ['debt'])},
        {debt: true, paid: false, opex: false, general: false, test: true, noTopics: true});
  const sal = [
    {row: 4, name: 'Зарплата', fact: 0, budget: 0, mf: 0, mb: 0},
    {row: 5, name: 'Директор', fact: 300000, budget: 300000, mf: 30000, mb: 30000},
    {row: 6, name: 'Find-iki', fact: 0, budget: 0, mf: 0, mb: 0},
    {row: 7, name: 'Вихователь Find-iki', fact: 250000, budget: 250000, mf: 12000, mb: 25000},
    {row: 8, name: 'Англійська мова 280', fact: 80000, budget: 80000, mf: 0, mb: 8000},
    {row: 9, name: 'Логопед 300', fact: 60000, budget: 60000, mf: 3000, mb: 6000},
    {row: 10, name: 'Додаткові заняття', fact: 0, budget: 0, mf: 0, mb: 0},
    {row: 11, name: 'Хореографія', fact: 40000, budget: 40000, mf: 4500, mb: 4000}];
  const S = P._bankSalaryCalc(sal), ST = P._bankSalaryTexts('Осокорки сад', 'листопад', S, '16.11');
  check('банк: фонд ЗП по розділах', 'основні / предметники (ставка в назві) / додаткові (після «Додаткові заняття»); заголовки груп і розділів не рахуються',
        {S}, {S: {main: {plan: 55000, paid: 42000}, subjects: {plan: 14000, paid: 3000}, extras: {plan: 4000, paid: 4500}}});
  check('банк: «💼 Фонд ЗП» — вигляд', 'шапка й «на ДД.ММ»; колонки в <code> без копійок; залишок може бути з мінусом; внизу «Разом: залишок»; без імен співробітників',
        {lines: ST[0].split('\n'), names: /Директор|Вихователь|Хореограф/.test(ST[0]), empty: P._bankSalaryTexts('X', 'листопад', {main: {plan: 0, paid: 0}}, '16.11').length},
        {lines: ['💼 <b>Фонд ЗП — Осокорки сад, листопад</b>', '<i>на 16.11</i>', '',
          '<code>Розділ         План  Виплачено  Залишок</code>',
          '<code>Основні      55 000     42 000   13 000</code>',
          '<code>Предметники  14 000      3 000   11 000</code>',
          '<code>Додаткові     4 000      4 500     −500</code>', '', '<b>Разом: залишок 23 500 ₴</b>'], names: false, empty: 0});
}

// v7.559: еталон боргів CFO (Осокорки, 18.11) — цілісність файлу; звірка з живим списком — node _backend/tests/bank_etalon.js
{
  const EOS = JSON.parse(fs.readFileSync(path.join(__dirname, 'fixtures_bank_etalon_osokorky.json'), 'utf8'));
  check('банк: еталон боржників Осокорків 18.11', '23 дитини, разом 233 110',
        {n: EOS.kids.length, total: EOS.kids.reduce((a, k) => a + (k.navch || 0) + (k.dop || 0) + (k.entry || 0), 0)}, {n: 23, total: 233110});
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

// v7.544: розумна стаття OPEX — правило → історія ≥ 80% → ШІ → питання
if (src.indexOf('\nfunction _catDecide(') >= 0){
  const K = new Function(['OPEX_SKIP_CAT','CAT_HIST_SHARE','CAT_HIST_MIN_N','CAT_CONF_MIN','CAT_LLM_TEXT_MAX','CAT_OTHER','INV_TG_MAX'].map(varLine).concat(
    ['trim','_htmlEsc','_invMoneyUa','_bankMoney','_invEdrMatch','_catNorm','_catInList','_catItemKey','_catItemsFromCell','_catHistIndex','_catHistPick','_catRulePick',
     '_catLlmPayload','_catLlmParse','_catDecide','_catQuestionText','_catAnswerParse'].map(fn)).join('\n') +
    '\nreturn {_catItemKey, _catHistIndex, _catHistPick, _catRulePick, _catLlmPayload, _catLlmParse, _catDecide, _catQuestionText, _catAnswerParse};')();
  const OS = ['ОРЕНДА','КОМУНАЛЬНІ ПОСЛУГИ','СНІДАНКИ','КУХНЯ','Вода','КАНЦТОВАРИ','ІГРАШКИ','ХОЗ.ТОВАРИ','ПРАЛЬНЯ','РЕМОНТ'];
  const H = ['Локація','IBAN','Дата платежу','Референс','ЄДРПОУ','Сума','Категорія','Місяць','Коли','Ким','Статус','Причина'];
  const lg = (loc, edr, cat, st, ref) => [loc, 'UA', '01.11.2026', ref || 'R', edr, 100, cat, 11, '', 'cfo', st || 'внесено', ''];
  const IDX = K._catHistIndex([H,
    lg('Осокорки', '36265773', 'КАНЦТОВАРИ'), lg('Осокорки', '36265773', 'КАНЦТОВАРИ'), lg('Осокорки', '36265773', 'КАНЦТОВАРИ'), lg('Осокорки', '36265773', 'ІГРАШКИ'),
    lg('Осокорки', '32490244', 'ХОЗ.ТОВАРИ'), lg('Осокорки', '32490244', 'ХОЗ.ТОВАРИ'), lg('Позняки', '32490244', 'РЕМОНТ'), lg('Осокорки', '32490244', 'РЕМОНТ'),
    lg('Позняки', '41981829', 'Комунальні послуги'), lg('Позняки', '41981829', 'КОМУНАЛЬНІ ПОСЛУГИ'), lg('Тичини', '41981829', 'КОМУНАЛЬНІ ПОСЛУГИ'),
    lg('Осокорки', '11111111', '→Salary: Іваненко'), lg('Осокорки', '11111111', 'КУХНЯ', 'пропущено'), lg('Осокорки', '22222222', 'ВОДА', '', 'ONE')]);
  const hK = K._catHistPick(IDX, '36265773', 'Осокорки', OS), hE = K._catHistPick(IDX, '32490244', 'Осокорки', OS), hY = K._catHistPick(IDX, '419818226543', 'Осокорки', OS);
  check('стаття: історія виписки', 'канцтовари 3 з 4 = 75% — мало; змішаний (Епіцентр у локації 2/3) — без статті, але з підказкою; ЯСНО — мережа 3/3 (ІПН ↔ ЄДРПОУ)',
        {k: hK.cat + ' ' + hK.share, e: hE.cat + '|' + hE.top.map((t) => t.cat + t.n).join(','), y: hY.cat + ' ' + hY.scope + ' ' + hY.n},
        {k: ' 0.75', e: '|ХОЗ.ТОВАРИ2,РЕМОНТ1', y: 'КОМУНАЛЬНІ ПОСЛУГИ мережа 3'});
  const IDX2 = K._catHistIndex([H, lg('Осокорки', '36265773', 'КАНЦТОВАРИ'), lg('Осокорки', '36265773', 'КАНЦТОВАРИ', '', 'SELF'), lg('Осокорки', '36265773', 'Канцтовари')]);
  check('стаття: ≥ 80% і без підглядання', '3/3 → КАНЦТОВАРИ; мінус платіж самої заявки → 2/2 ще так; 1 платіж — замало; Salary і пропущені — не історія',
        {a: K._catHistPick(IDX2, '36265773', 'Осокорки', OS).cat, b: K._catHistPick(IDX2, '36265773', 'Осокорки', OS, [{loc: 'Осокорки', cat: 'КАНЦТОВАРИ'}]).cat,
         c: K._catHistPick(IDX, '22222222', 'Осокорки', OS).cat, d: K._catHistPick(IDX, '11111111', 'Осокорки', OS)}, {a: 'КАНЦТОВАРИ', b: 'КАНЦТОВАРИ', c: '', d: null});
  const RULES = [{type: 'постачальник', key: '32490244', loc: '', cat: 'ХОЗ.ТОВАРИ', src: 'контрагенти'},
                 {type: 'позиція', key: K._catItemKey('Фарба інтер\'єрна біла 10 л'), loc: 'Осокорки', cat: 'РЕМОНТ', src: 'CFO'},
                 {type: 'постачальник', key: '32490244', loc: 'Осокорки', cat: 'ХОЗ.ТОВАРИ', src: 'CFO'}];
  check('стаття: правило людини', 'позиція «фарба» → РЕМОНТ важливіша за постачальника; інші позиції — правило постачальника; стаття не з цієї локації — не правило',
        {item: K._catRulePick(RULES, '32490244', 'Осокорки', ['Фарба інтерʼєрна біла, 10л'], OS).cat, sup: K._catRulePick(RULES, '32490244', 'Осокорки', ['Пакети для сміття 120 л'], OS).src,
         none: K._catRulePick([{type: 'постачальник', key: '36265773', loc: '', cat: 'Розвиток', src: ''}], '36265773', 'Осокорки', [], OS), key: K._catItemKey('Папір А4 офісний 500 арк.')},
        {item: 'РЕМОНТ', sup: 'правило: постачальник (CFO)', none: null, key: 'папір офісний'});
  const pl = K._catLlmPayload('claude-haiku-5-5', OS, {loc: 'Осокорки', caption: 'Осокорки, садочок / Папір', supplier: 'ТОВ «Буквоїд»', edrpou: '36265773', text: 'Папір А4 — 5 пач.', hint: 'КАНЦТОВАРИ ×1'});
  const plDoc = K._catLlmPayload('claude-haiku-5-5', OS, {loc: 'Осокорки', doc: {type: 'document'}});
  check('стаття: запит до ШІ', 'Haiku, статті — enum, текст рахунку в промпті; без тексту — сам PDF вкладенням',
        {m: pl.model, en: pl.output_config.format.schema.properties.category.enum.length, txt: /Папір А4/.test(pl.messages[0].content) && /КАНЦТОВАРИ ×1/.test(pl.messages[0].content),
         doc: Array.isArray(plDoc.messages[0].content) && plDoc.messages[0].content[0].type}, {m: 'claude-haiku-5-5', en: 10, txt: true, doc: 'document'});
  const resp = (o) => ({content: [{type: 'text', text: JSON.stringify(o)}], usage: {input_tokens: 900, output_tokens: 40}});
  const L1 = K._catLlmParse(resp({category: 'канцтовари', alternative: 'ІГРАШКИ', confidence: 0.92, items: ['Папір А4', 'Конверти']}), OS);
  const L2 = K._catLlmParse(resp({category: 'ХОЗ.ТОВАРИ', alternative: 'РЕМОНТ', confidence: 0.55, items: ['Фарба біла', 'Валик']}), OS);
  check('стаття: відповідь ШІ', 'назва статті — як у файлі; поза списком / відмова — помилка (без статті)',
        {a: L1.cat + ' ' + L1.conf, bad: K._catLlmParse(resp({category: 'Розвиток', alternative: '', confidence: 1, items: []}), OS).ok, ref: K._catLlmParse({stop_reason: 'refusal', content: []}, OS).error},
        {a: 'КАНЦТОВАРИ 0.92', bad: false, ref: 'ШІ відмовив'});
  const dQ = K._catDecide(null, null, L2, ''), dE = K._catDecide(null, null, {ok: false, error: 'нема ANTHROPIC_API_KEY'}, 'КАНЦТОВАРИ');
  check('стаття: рішення', 'правило → історія → ШІ ≥ 0,75; 0,55 — питання «X чи Y»; без ШІ — з підпису',
        {r: K._catDecide({cat: 'РЕМОНТ', src: 'правило: позиція'}, {cat: 'ХОЗ.ТОВАРИ', share: 1, n: 5, scope: 'локація'}, L1, '').cat,
         h: K._catDecide(null, {cat: 'ХОЗ.ТОВАРИ', share: 0.83, n: 6, scope: 'локація'}, null, '').src, l: K._catDecide(null, null, L1, '').cat,
         q: dQ.cat + '|' + dQ.ask.join(' / '), e: dE.cat + ' · ' + dE.src},
        {r: 'РЕМОНТ', h: 'історія 83% з 6 (локація)', l: 'КАНЦТОВАРИ', q: '|ХОЗ.ТОВАРИ / РЕМОНТ', e: 'КАНЦТОВАРИ · підпис (нема ANTHROPIC_API_KEY)'});
  check('стаття: питання в групі «Банк»', 'текст «[позиції] — X чи Y?»; відповідь «2», «ремонт», «канц», «так» на одну статтю; незрозуміле — порожньо',
        {t: K._catQuestionText('301', 'ТОВ «Епіцентр К»', 2450, L2.items, dQ.ask).split('\n')[1],
         a2: K._catAnswerParse('2', dQ.ask, OS), ar: K._catAnswerParse('Ремонт.', dQ.ask, OS), ak: K._catAnswerParse('канц', dQ.ask, OS),
         y: K._catAnswerParse('так', ['КАНЦТОВАРИ'], OS), y2: K._catAnswerParse('так', dQ.ask, OS), x: K._catAnswerParse('не знаю', dQ.ask, OS)},
        {t: '<b>Фарба біла, Валик</b> — ХОЗ.ТОВАРИ чи РЕМОНТ?', a2: 'РЕМОНТ', ar: 'РЕМОНТ', ak: 'КАНЦТОВАРИ', y: 'КАНЦТОВАРИ', y2: '', x: ''});
}

// ── v7.563: експорт допів під локом + перерахунок після видалення (Голосієво, дубль Лего 28.10, +340 у трьох дітей) ──
{
  const st = {held: false, waits: 0, flushed: 0, fail: false};
  const LS = {getScriptLock: () => ({waitLock(){ if (st.fail) throw new Error('timeout'); if (st.held) throw new Error('deadlock'); st.held = true; st.waits++; },
                                      releaseLock(){ st.held = false; }})};
  const SA = {flush(){ st.flushed++; }};
  const W = new Function('LockService', 'SpreadsheetApp', ['var _DOP_EXP_LOCKED = false;', fn('_dopExpLocked'), fn('_dopRecalcTargets')].join('\n') +
    '\nreturn {_dopExpLocked, _dopRecalcTargets};')(LS, SA);
  // модель експорту: клітинка = (клітинка − журнал) + сума відміток; журнал = сума. Читання й запис — усередині core.
  const world = {cell: 2810 + 340, journal: 2810 + 340, marks: 2810};   // стан, коли дубль уже знято, а клітинка/журнал бачили його
  const core = (p) => { const held = st.held, base = world.cell - world.journal;
    if (!p.dryRun){ world.cell = base + world.marks; world.journal = world.marks; }
    if (p.nested) W._dopExpLocked({}, (q) => ({ok: true, held: st.held}));
    return {ok: true, held: held}; };
  const r1 = W._dopExpLocked({loc: 'Голосієво'}, core);
  const waits1 = st.waits;
  const rN = W._dopExpLocked({nested: true}, core);
  const rD = W._dopExpLocked({dryRun: true}, core);
  st.fail = true; const rT = W._dopExpLocked({}, core); st.fail = false;
  check('допи: експорт під локом', 'реальний прогін — під ScriptLock (вкладений Payment+Salary не бере лок вдруге), dryRun без локу, зайнято — LOCK_TIMEOUT; після прогону клітинка = сума відміток',
        {held: r1.held, waits: waits1, nestedWaits: st.waits - waits1, dryHeld: rD.held, timeout: rT.code, released: st.held, cell: world.cell},
        {held: true, waits: 1, nestedWaits: 1, dryHeld: false, timeout: 'LOCK_TIMEOUT', released: false, cell: 2810});
  const snaps = [{loc: 'Голосієво', date: '2026-10-28', child: 'Бучковський Лука'}, {loc: 'Голосієво', date: '2026-10-28', child: 'Благовісна Аліса'},
                 {loc: 'Голосієво', date: '2026-10-28', child: 'Ковальчук Соломія'}, {loc: 'Голосієво', date: '2026-11-02'}, {loc: 'Позняки', date: '2026-10-30'}, {loc: '', date: '2026-10-01'}];
  check('допи: що перерахувати після видалення', 'локація × місяць відмітки, без повторів (три видалені Лего 28.10 — один перерахунок жовтня)',
        {t: W._dopRecalcTargets(snaps).map((x) => x.loc + ' ' + x.year + '-' + x.month)},
        {t: ['Голосієво 2026-10', 'Голосієво 2026-11', 'Позняки 2026-10']});
}

// ── v7.564: дубль відмітки допів — той самий ключ попри регістр, пробіли, апострофи й формат дати ──
{
  const D = new Function([varLine('_APOSTROPHES'), fn('_nameFold'), fn('_nameFoldTight'), fn('_journalNormName'), fn('_attDateFast'), fn('_attDateISO'), fn('_attDupKey')].join('\n') +
    '\nreturn {_attDupKey};')();
  const k = (d, c, a, l) => D._attDupKey(d, c, a, 'Europe/Kiev', l);
  const base = k('2026-10-28', 'Ковальчук Соломія', 1, 'Голосієво');
  check('допи: ключ дубля відмітки', 'дата + локація + дитина + заняття; «Сук  Лера» / «сук лера» / нерозривний пробіл / Кар’єрна — те саме; інший доп, дата чи садок — ні',
        {same: [k('2026-10-28', ' ковальчук  Соломія ', 1, 'Голосієво'), k('2026-10-28', 'Ковальчук Соломія', '1', 'Голосієво')].every((x) => x === base),
         suk: k('2026-07-07', 'Сук  Лера', 7, "Кар'єрна") === k('2026-07-07', 'сук лера', 7, 'Кар’єрна'),
         otherAct: k('2026-10-28', 'Ковальчук Соломія', 2, 'Голосієво') === base,
         otherDate: k('2026-10-21', 'Ковальчук Соломія', 1, 'Голосієво') === base,
         otherLoc: k('2026-10-28', 'Ковальчук Соломія', 1, 'Позняки') === base},
        {same: true, suk: true, otherAct: false, otherDate: false, otherLoc: false});
}

// ── v7.565: унікальні № відміток, видалення з перевіркою, заморожені червень–вересень 2026 ──
{
  const R = new Function([varLine('_APOSTROPHES'), fn('_nameFold'), fn('_nameFoldTight'), fn('_journalNormName'), fn('_attDateFast'), fn('_attDateISO'),
    fn('_attNextIdFrom'), fn('_attRenumberPlan'), fn('_attDelMismatch')].join('\n') + '\nreturn {_attNextIdFrom, _attRenumberPlan, _attDelMismatch};')();
  // реальні повтори: №165 у трьох дітей Борщагівки, №166 у двох; останній рядок аркуша — старий №2 (звідси й повтори)
  const col = [[1], [2], [165], [165], [165], [166], [166], [500], [''], [2]];
  const plan = R._attRenumberPlan(col, 0);
  check('відмітки: перенумерація повторів', 'перший рядок лишає №, решта — max+1…; порожній № теж; новий № = max колонки, а не № останнього рядка + 1',
        {plan: plan.map((q) => q.from + '→' + q.to), next: R._attNextIdFrom(col, 0), hwm: R._attNextIdFrom(col, 900)},
        {plan: ['165→501', '165→502', '166→503', '0→504', '2→505'], next: 501, hwm: 901});
  const row = [19525, '2026-10-28', 'Голосієво', 'Preschool', 'Бучковський Лука', 1, 'Лего', 340, 'vyhovatel', ''];
  const tz = 'Europe/Kiev';
  check('відмітки: видалення лише № + перевірка', 'збіг дитини, дати й заняття — видаляємо; інша дитина / дата / заняття чи без перевірки — ні',
        {ok: R._attDelMismatch(row, {child: 'бучковський  лука', date: '2026-10-28', activityId: 1}, tz),
         kid: /дитина «Бучковський Лука», а не «Благовісна Аліса»/.test(R._attDelMismatch(row, {child: 'Благовісна Аліса', date: '2026-10-28', activityId: 1}, tz)),
         date: /дата 2026-10-28, а не 2026-10-21/.test(R._attDelMismatch(row, {child: 'Бучковський Лука', date: '2026-10-21', activityId: 1}, tz)),
         act: /заняття «Лего» \(#1\), а не #7/.test(R._attDelMismatch(row, {child: 'Бучковський Лука', date: '2026-10-28', activityId: 7}, tz)),
         none: /оновіть сторінку/.test(R._attDelMismatch(row, {}, tz))},
        {ok: '', kid: true, date: true, act: true, none: true});
  const C = new Function('SpreadsheetApp', ['var _CLOSED_SET_MEMO = null;', varBlock('FROZEN_MONTHS'), fn('_isFrozenMonth'), fn('_closedMonthsSet'), fn('_isMonthClosed'),
    fn('_closedMonthError'), 'function _getClosedMonthsSheet(){ return null; }'].join('\n') + '\nreturn {_isMonthClosed, _closedMonthError};')({});
  check('заморожені місяці', 'червень–вересень 2026 закриті завжди (навіть без рядка в «Закриті_Місяці»), травень і жовтень — ні; помилка FROZEN_MONTH',
        {m: [5, 6, 7, 8, 9, 10].map((m) => C._isMonthClosed(2026, m)), y27: C._isMonthClosed(2027, 6), code: C._closedMonthError(2026, 9).code},
        {m: [false, true, true, true, true, false], y27: false, code: 'FROZEN_MONTH'});
}

// ── v7.566: звірка, «вже внесено» — на кого внесено (Школа 228, К3_транз.рахунок, листопад) ──
{
  const Ut = {formatDate: (d) => '19.11.2026 08:13'};
  const Z = new Function('Utilities', 'trim', [fn('_rcAppliedByRef'), fn('_rcColKind')].join('\n') + '\nreturn {_rcAppliedByRef, _rcColKind};')(Ut, (x) => String(x == null ? '' : x).trim());
  const H = ['Коли','Ким','Локація','IBAN','Дата платежу','Референс','Платник','Дитина','Рядок','Місяць','Колонка','Сума','Було','Стало','Ключ дедупу'];
  const L = (ref, child, row, col, amt, key) => [new Date(), 'Мельніченко Ірина (CFO)', 'Школа 228', 'UA72…925', '02.11.2026', ref, 'K3_транз.рахунок платежі BP 3500306', child, row, 11, col, amt, 0, amt, key === undefined ? 'REF|' + ref : key];
  const A = Z._rcAppliedByRef([H,
    L('D3K1QB2A9Y1E3B.P', 'Сухина Міна (2 клас)', 120, 'BA', 17000),            // кандидатів за платником 0 — раніше просто «вже внесено»
    L('D3K5QB3A9ZYZFC.P', 'Шихов Ігор', 88, 'BA', 2100),                        // двоє Шихових у кандидатах
    L('SPLIT1', 'Терещук Максим', 90, 'BA', 9000), L('SPLIT1', 'Терещук Лука', 91, 'BA', 8000),   // один платіж на двох дітей
    L('NOKID', '', 0, '', 500)]);
  check('звірка: на кого вже внесено', 'дитина, рядок, місяць, колонка з «Звірки_Платежів» за референсом; кілька дітей — усі; без дитини видно; колонка → навч / вступ / доп',
        {suh: A['D3K1QB2A9Y1E3B.P'].map((e) => e.child + ' р.' + e.row + ' міс.' + e.month + ' ' + e.col),
         split: A['SPLIT1'].map((e) => e.child + ' ' + e.amount), nokid: A['NOKID'].length + ':' + A['NOKID'][0].child,
         kinds: [Z._rcColKind(52, 52), Z._rcColKind(53, 52), Z._rcColKind(54, 52), Z._rcColKind(55, 52), Z._rcColKind(10, -1)]},
        {suh: ['Сухина Міна (2 клас) р.120 міс.11 BA'], split: ['Терещук Максим 9000', 'Терещук Лука 8000'], nokid: '1:',
         kinds: ['навч', 'вступ', 'доп', '', '']});
}

// ── v7.567: примітки-розшифровки (Payment / OPEX / Salary) ──
{
  const N = new Function([varLine('NOTE_TAG'), varLine('NOTE_USER_SEP'), fn('_noteRound'), fn('_noteMoney'), fn('_noteDay'), fn('_noteCompose'),
    fn('_noteMerge'), fn('_noteDopLines'), fn('_noteIban4'), fn('_noteJoin'), fn('_recParseDate'), fn('_notesDirtyKeys')].join('\n') +
    '\nreturn {_noteMoney, _noteDay, _noteCompose, _noteMerge, _noteDopLines, _noteIban4, _noteJoin, _notesDirtyKeys};')();
  // Ткаченко Вікторія (Позняки): «Бюджет доп» листопада = відмітки жовтня 6 030
  const tk = [['Арт', 370, '2026-10-01'], ['Логопед', 650, '2026-10-02'], ['Логопед', 650, '2026-10-06'], ['Лего', 370, '2026-10-08'], ['Театр', 500, '2026-10-19'],
    ['Гончарство', 400, '2026-10-20'], ['Логопед', 650, '2026-10-20'], ['Логопед', 650, '2026-10-22'], ['Лего', 370, '2026-10-22'], ['Логопед', 650, '2026-10-27'],
    ['Гончарство', 400, '2026-10-27'], ['Лего', 370, '2026-10-29']].map((x) => ({activityName: x[0], price: x[1], date: x[2]}));
  const dl = N._noteDopLines(tk);
  const ok = N._noteCompose('Ткаченко Вікторія, бюджет допів, листопад', dl, 6030);
  const man = N._noteCompose('Ткаченко Вікторія, бюджет допів, листопад', dl, 5380);
  check('примітки: бюджет допів', 'заняття × ціна з датами; сума сходиться — без «вручну», не сходиться — останній рядок «− вручну 650»',
        {head: ok.split('\n')[0], lego: dl.find((l) => /^Лего/.test(l.text)).text, logo: dl.find((l) => /^Логопед/.test(l.text)).text,
         okLast: ok.split('\n').slice(-1)[0], manLast: man.split('\n').slice(-1)[0]},
        {head: 'ⓘ Ткаченко Вікторія, бюджет допів, листопад: 6 030', lego: 'Лего: 3 × 370 = 1 110 (08, 22, 29.10)', logo: 'Логопед: 5 × 650 = 3 250 (02, 06, 20, 22, 27.10)',
         okLast: 'Гончарство: 2 × 400 = 800 (20, 27.10)', manLast: '− вручну 650'});
  const pay = N._noteCompose('Бучковський Лука, оплата допів, листопад', [{text: N._noteJoin([N._noteDay('02.11.2026'), N._noteMoney(2810), 'Бучковський Георгій Валерійович', N._noteIban4('UA72 3052 9900 0002 6000 0462 0592 5')]), amount: 2810}], 2810);
  const sal = N._noteCompose('Іванова Олена, нараховано, листопад', [{text: 'ставка за карткою: 18000', amount: 0, noSum: true}, {text: 'допи: 8 занять, 40 відміток → 4 000', amount: 4000},
    {text: 'Виплачено:', noSum: true, tail: true}, {text: '  05.11 · 10 000 · карта …1234 · відомість №7', amount: 10000, noSum: true, tail: true}], 22000);
  check('примітки: оплата, Salary, нуль', 'оплата — дата · сума · платник · рахунок; Salary — «вручну» перед блоком «Виплачено»; порожня клітинка без рядків — без примітки',
        {pay: pay.split('\n')[1], sal: sal.split('\n'), empty: N._noteCompose('x', [], 0), plus: N._noteCompose('x', [{text: 'a', amount: 100}], 150.5).split('\n').pop()},
        {pay: '02.11 · 2 810 · Бучковський Георгій Валерійович · рах. …5925', empty: '', plus: '+ вручну 50,50',
         sal: ['ⓘ Іванова Олена, нараховано, листопад: 22 000', 'ставка за карткою: 18000', 'допи: 8 занять, 40 відміток → 4 000', '+ вручну 18 000', 'Виплачено:', '  05.11 · 10 000 · карта …1234 · відомість №7']});
  const ours1 = 'ⓘ Лего, листопад: 340', ours2 = 'ⓘ Лего, листопад: 680';
  check('примітки: ручна примітка не губиться', 'наша частина першою, ручна — після «✎»; оновлення міняє лише нашу; прибрали нашу — лишається ручна',
        {first: N._noteMerge('домовились з мамою', ours1), upd: N._noteMerge(N._noteMerge('домовились з мамою', ours1), ours2), del: N._noteMerge(ours1 + '\n\n✎ домовились з мамою', ''), plain: N._noteMerge(ours1, ours2)},
        {first: 'ⓘ Лего, листопад: 340\n\n✎ домовились з мамою', upd: 'ⓘ Лего, листопад: 680\n\n✎ домовились з мамою', del: 'домовились з мамою', plain: ours2});
  check('примітки: що оновити після запису', 'звірка → Payment місяця платежу; експорт допів → Payment і Salary наступного місяця; OPEX — по рядках; dryRun і помилка — нічого',
        {rec: N._notesDirtyKeys('reconcileApply', {items: [{date: '02.11.2026'}, {month: 10}]}, {ok: true, loc: 'Позняки'}),
         exp: N._notesDirtyKeys('exportAttendance', {loc: 'Голосієво', month: 11, year: 2026}, {ok: true}),
         opex: N._notesDirtyKeys('opexAddExpenses', {loc: 'Тичини', month: 11, year: 2026, items: [{category: 'КАНЦТОВАРИ'}, {skip: true}, {loc: 'Оранж', month: 10}]}, {ok: true}),
         dry: N._notesDirtyKeys('addLocPaymentDop', {loc: 'Позняки', month: 11, dryRun: true}, {ok: true}).length, bad: N._notesDirtyKeys('cashAdd', {}, {ok: false}).length},
        {rec: ['payment|Позняки|2026|11', 'payment|Позняки|2026|10'], exp: ['payment|Голосієво|2026|12', 'salary|Голосієво|2026|12'],
         opex: ['opex|Тичини|2026|11', 'opex|Оранж|2026|10'], dry: 0, bad: 0});
}

// ── v7.569: бот рахунків — супергрупа, гілки, «✅ Оплачено ДД.ММ — сума» ──
{
  const I = new Function([fn('_invChatResolve'), fn('_invPaidText'), fn('_invPaidParams'), fn('_invTopicLocOf')].join('\n') +
    '\nreturn {_invChatResolve, _invPaidText, _invPaidParams, _invTopicLocOf};')();
  const LOCS = ['Тичини', 'Осокорки', 'Позняки', "Кар'єрна", 'Голосієво', 'Пуща', 'Оранж', 'Борщагівка', 'Бровари', 'Кругла', 'Бігова',
    "Школа Кар'єрна", 'Школа Осокорки', 'Школа 228', 'Житомир', 'Манхетен (Благо)', 'Нац.Гвардії (Благо)', 'Хіміків (Благо)'];
  const L = (n) => I._invTopicLocOf(n, LOCS);
  check('рахунки: стара група → супергрупа', 'заявки зі старим chat_id (-5142518705) ідуть у нову групу; відповідь у гілку, на повідомлення з рахунком, зникло — без відповіді',
        {chat: I._invChatResolve('-5142518705', {'-5142518705': '-1004472189929'}), same: I._invChatResolve('-1004472189929', {'-5142518705': '-1004472189929'}),
         p: I._invPaidParams('377', '12'), g: I._invPaidParams(120, '')},
        {chat: '-1004472189929', same: '-1004472189929',
         p: {reply_parameters: {message_id: 377, allow_sending_without_reply: true}, message_thread_id: 12}, g: {reply_parameters: {message_id: 120, allow_sending_without_reply: true}}});
  check('рахунки: текст «оплачено»', '«✅ Оплачено ДД.ММ — сума»: ціле без копійок, копійки через кому',
        {a: I._invPaidText('18.11.2026', 5500), b: I._invPaidText('19.11.2026', 11595.83), c: I._invPaidText('17.11.2026', 481), z: I._invPaidText('17.11.2026', 0)},
        {a: '✅ Оплачено 18.11 — 5 500', b: '✅ Оплачено 19.11 — 11 595,83', c: '✅ Оплачено 17.11 — 481', z: '✅ Оплачено 17.11'});
  check('рахунки: гілка → локація', 'однозначна назва → локація; «Школа Осокорки» не плутається з «Осокорки»; неоднозначне чи незнайоме — без локації',
        {t: L('Тичини').loc, s: L('Школа Осокорки').loc, o: L('Осокорки').loc, k: L("Кар'єрна садок").loc, m: L('Манхетен').loc,
         amb: L('Осокорки сад і школа').loc, gen: L('Загальна').loc, ofc: L('Офіс').why},
        {t: 'Тичини', s: 'Школа Осокорки', o: 'Осокорки', k: "Кар'єрна", m: 'Манхетен (Благо)', amb: '', gen: '', ofc: 'не знайдено локації'});
}

// ── v7.570: «оплачено» — повний текст і гілка локації ──
{
  const Q = new Function([fn('_invPaidTextFull'), fn('_invLocThreads'), fn('_invPaidRoute')].join('\n') + '\nreturn {_invPaidTextFull, _invLocThreads, _invPaidRoute};')();
  const full = Q._invPaidTextFull({dmy: '18.11.2026', loc: 'Тичини', supplier: 'ТОВ «Епіцентр К»', amount: 5500, caption: 'Тичини\nканцтовари  для групи', number: '301'});
  const map = {'2': {name: 'Осокорки сад', loc: 'Осокорки', why: 'назва містить локацію'}, '4': {name: 'Тичини', loc: 'Тичини', why: 'назва = локація'},
    '5': {name: 'Карʼєрна', loc: "Кар'єрна", why: 'назва = локація'}, '14': {name: 'Кухня Карʼєрна', loc: "Кар'єрна", why: 'назва містить локацію'}, '12': {name: 'Осокорки школа', loc: '', why: 'неоднозначно'}};
  const lt = Q._invLocThreads(map);
  check('рахунки: повне «оплачено»', '«✅ Оплачено ДД.ММ — Локація — Контрагент — сума ₴ — опис / № рахунку»; без опису й номера — видно, чого бракує',
        {full: full, bare: Q._invPaidTextFull({dmy: '17.11.2026', amount: 481.5})},
        {full: '✅ Оплачено 18.11 — Тичини — ТОВ «Епіцентр К» — 5 500 ₴ — Тичини канцтовари для групи / № 301',
         bare: '✅ Оплачено 17.11 — локація ? — контрагент ? — 481,50 ₴ — без опису'});
  check('рахунки: куди писати «оплачено»', 'у гілці — відповіддю в ній; із загальної (до гілок) — у гілку локації без відповіді; Кар\'єрна — основна гілка, не «Кухня»; без гілки локації — загальна відповіддю',
        {lt: lt, inThread: Q._invPaidRoute({messageId: 160, threadId: '4', loc: 'Тичини'}, lt), general: Q._invPaidRoute({messageId: 596, threadId: '', loc: 'Осокорки'}, lt),
         kar: Q._invPaidRoute({messageId: 601, threadId: '', loc: "Кар'єрна"}, lt), none: Q._invPaidRoute({messageId: 794, threadId: '', loc: 'Школа 228'}, lt)},
        {lt: {'Осокорки': '2', 'Тичини': '4', "Кар'єрна": '5'}, inThread: {thread: 4, reply: 160}, general: {thread: 2, reply: 0},
         kar: {thread: 5, reply: 0}, none: {thread: 0, reply: 794}});
}

// ── v7.571: «оплачено» — локація (гілка → відправник → рахунок оплати) і контрагент із виписки ──
{
  const W = new Function([fn('_invSenderLocs'), fn('_invPickLoc'), fn('_invSupplierFromRefs'), fn('_invTopicLocFromMap'), fn('_invSupGeneric')].join('\n') +
    '\nreturn {_invSenderLocs, _invPickLoc, _invSupplierFromRefs, _invTopicLocFromMap, _invSupGeneric};')();
  const H = ['invoice_id', 'від кого', 'локація'];
  const v = [H, [1, '@mkidsBrovary', 'Бровари'], [2, '@mkidsBrovary', 'Бровари'], [3, '@mkidsBrovary', ''], [4, '@Ira_Melnichenk0', 'Тичини'], [5, '@Ira_Melnichenk0', 'Пуща'],
    [6, 'm.Kids Пуща-Водиця', 'Пуща'], [7, 'Ольга', 'Позняки'], [8, 'Ольга', 'Позняки'], [9, 'Ольга', 'Кухня Карʼєрна (не локація)']];
  const SL = W._invSenderLocs(v, H);
  const IB = {'UA093052990000026000010124708': ['Бровари'], 'UA723052990000026000046205925': ['Школа 228'], 'UA111': ['Осокорки', 'Школа Осокорки']};
  const P = (x) => W._invPickLoc(x, SL, IB);
  check('рахунки: локація для «оплачено»', 'гілка → відправник (усі його заявки — одна локація, ≥ 2) → рахунок оплати (IBAN однієї локації); CFO і спільний рахунок — «?»',
        {sl: SL, th: P({threadLoc: 'Тичини', sender: '@mkidsBrovary'}), snd: P({sender: '@mkidsBrovary', iban: 'UA723052990000026000046205925'}),
         iban: P({sender: '@Ira_Melnichenk0', iban: 'UA72 3052 9900 0002 6000 0462 0592 5'}), shared: P({sender: '@Ira_Melnichenk0', iban: 'UA111'}), one: P({sender: 'm.Kids Пуща-Водиця'})},
        {sl: {'@mkidsBrovary': 'Бровари', 'Ольга': 'Позняки'}, th: {loc: 'Тичини', src: 'гілка'}, snd: {loc: 'Бровари', src: 'відправник'},
         iban: {loc: 'Школа 228', src: 'рахунок оплати'}, shared: {loc: '', src: ''}, one: {loc: '', src: ''}});
  const BP = {'JBKLQBJOF7VNX2.1': {name: 'ФОП  Ковальчук   Ігор'}, 'B2': {name: ''}};
  check('рахунки: контрагент із виписки', 'отримувач за matched_ref; кілька референсів — перший відомий; гілка кухні лишається міткою кухні',
        {one: W._invSupplierFromRefs('JBKLQBJOF7VNX2.1', BP), multi: W._invSupplierFromRefs('B2+JBKLQBJOF7VNX2.1', BP), none: W._invSupplierFromRefs('', BP),
         kitchen: W._invTopicLocFromMap({'14': {loc: 'Кухня Карʼєрна (не локація)'}}, '14.0')},
        {one: 'ФОП Ковальчук Ігор', multi: 'ФОП Ковальчук Ігор', none: '', kitchen: 'Кухня Карʼєрна (не локація)'});
  check('рахунки: загальне слово замість контрагента', '«Замовник», «Одержувач», голе «ФІЗИЧНА ОСОБА ПІДПРИЕМЕЦЬ» / «Товариство з обмеженою відповідальністю» — беремо отримувача з виписки; справжня назва лишається',
        {g: ['Замовник', 'Одержувач', 'ФІЗИЧНА ОСОБА ПІДПРИЕМЕЦЬ', 'Товариство з обмеженою відповідальністю', ''].map(W._invSupGeneric),
         real: ['ФОП Чаус Я.О.', 'ТОВ "Броварський ЗБК"', 'Назва ФОП Мельник Діана Олегівна'].map(W._invSupGeneric)},
        {g: [true, true, true, true, true], real: [false, false, false]});
}

// ── v7.572: примітки OPEX — читабельний формат ──
{
  const vv = (n) => /var [A-Z_]+ *=[\s\S]*?;\n/.exec(src.slice(src.indexOf('var ' + n + ' ')))[0];
  const W = new Function(['NOTE_TAG', 'NOTE_CYR', 'NOTE_LET', 'NOTE_LOC_SRC'].map(vv).join('\n') +
    ['_noteRound', '_noteDay', '_noteUah', '_noteWordRe', '_noteTitleWord', '_noteShortParty', '_notePartyOk', '_noteCleanCaption', '_noteComposeOpex', '_noteSortKey'].map(fn).join('\n') +
    '\nreturn {_noteShortParty, _notePartyOk, _noteCleanCaption, _noteComposeOpex};')();
  check('OPEX примітка: отримувач коротко', 'без ТОВ/ФОП/лапок/п/р/IBAN; ФОП-особа → Прізвище І.Б.; уламок OCR — не отримувач',
        {a: W._noteShortParty('ТОВ "Сервіс Про"'), b: W._noteShortParty('ФОП Григоришина Світлана Олександрівна'),
         c: W._noteShortParty('ТОВАРИСТВО З ОБМЕЖЕНОЮ ВІДПОВІДАЛЬНІСТЮ "СЕРВІС ПРО"'), d: W._noteShortParty('Товариство з обмеженою відповідальністю "Гал Еко Трейд" п/р UA273006140000026000000000001'),
         e: W._noteShortParty('Фізична особа-підприємець Златова Ірина Володимирівна'), f: W._noteShortParty('ТОВ «СНС КЕЙТЕРИНГ»'),
         g: W._notePartyOk(W._noteShortParty('а, самовивозом, за наявності')), h: W._notePartyOk('Замовник')},
        {a: 'Сервіс Про', b: 'Григоришина С.О.', c: 'Сервіс Про', d: 'Гал Еко Трейд', e: 'Златова І.В.', f: 'СНС Кейтеринг', g: false, h: false});
  check('OPEX примітка: опис директора', 'без локації, @згадок, привітань, «прошу оплатити», назв файлів',
        {a: W._noteCleanCaption('Добрий день \nЖурнали брокеражу \nПуща \nДякую', 'Пуща'),
         b: W._noteCleanCaption('Дуже прошу оплатити цей рахунок @Ira_Melnichenk0 \nПослуга надана 20.10.', 'Тичини'),
         c: W._noteCleanCaption("Аптечка для укриття, Кар'єрна. З непередбачуваних", "Кар'єрна"),
         d: W._noteCleanCaption('Рахунок ТО_вікон_2a9fd8ed-686d-4197-b6f3-c0f7f015a9c4.pdf', 'Оранж'),
         e: W._noteCleanCaption('Сніданки 228', 'Школа 228'), f: W._noteCleanCaption('', 'Оранж'), g: W._noteCleanCaption('Осокорки, садочок\nСніданки\n@vldawq', 'Осокорки')},
        {a: 'Журнали брокеражу', b: 'Послуга надана 20.10', c: 'Аптечка для укриття, З непередбачуваних', d: '', e: 'Сніданки', f: '', g: 'Сніданки'});
  check('OPEX примітка: формат', 'СТАТТЯ · місяць · сума ₴; витрата — два рядки; цілі гривні; без опису — один рядок; «+ вручну» лише коли не сходиться',
        W._noteComposeOpex('Господарські товари', 'листопад', [
          {day: '12.11', sort: '1112', who: 'Сервіс Про', amount: 3200.4, desc: 'Госп. товари'},
          {day: '03.11', sort: '1103', who: 'Григоришина С.О.', amount: 1250.55, desc: ''}], 5450.95),
        'ⓘ ГОСПОДАРСЬКІ ТОВАРИ · листопад · 5 451 ₴\n\n03.11  Григоришина С.О. — 1 251 ₴\n12.11  Сервіс Про — 3 200 ₴\n       Госп. товари\n\n+ вручну 1 000 ₴');
  check('OPEX примітка: суми сходяться', 'різниця в копійках — без «вручну»; порожня стаття з нулем — без примітки',
        {a: W._noteComposeOpex('Вода', 'жовтень', [{day: '01.10', sort: '1001', who: 'Діана', amount: 700.2}], 700.2), b: W._noteComposeOpex('Вода', 'жовтень', [], 0)},
        {a: 'ⓘ ВОДА · жовтень · 700 ₴\n\n01.10  Діана — 700 ₴', b: ''});
}

// ── v7.573: назва з виписки «ФОРМА, НАЗВА» ──
{
  const vv = (n) => /var [A-Z_]+ *=[\s\S]*?;\n/.exec(src.slice(src.indexOf('var ' + n + ' ')))[0];
  const W = new Function(['NOTE_CYR', 'NOTE_LET'].map(vv).join('\n') + ['_noteWordRe', '_noteTitleWord', '_noteShortParty', '_notePartyOk'].map(fn).join('\n') +
    '\nreturn {_noteShortParty, _notePartyOk};')();
  check('OPEX примітка: форма власності перед комою', '«ОВВ, …» не дає «Овв»; «ТОВ, " БРОВАРСЬКИЙ ЗБК "» не дає порожньо → ЄДРПОУ',
        {a: W._noteShortParty('ОВВ, ГОЛОВНЕ УПРАВЛІННЯ ДЕРЖАВНОІ КАЗНАЧЕЙСЬКОІ СЛУЖБИ'), b: W._noteShortParty('ТОВ, " БРОВАРСЬКИЙ ЗБК "'),
         c: W._noteShortParty('ПРАТ, ФАРЛЕП-ІНВЕСТ'), d: W._noteShortParty('ТОВ, "КОМПАНІЯ ІННОВАЦІЙНИХ РІШЕНЬ"'), e: W._noteShortParty('Прокуда Н. Г. ФОП'), f: W._noteShortParty("ТОВ  'НОВА ПОШТА'"), g: W._noteShortParty("ТОВ 'УПРАВЛЯЮЧА КОМПАНІЯ 'КРІСТАЛ АВЕНЮ'"), h: W._notePartyOk(W._noteShortParty('ТОВ СТ')), i: W._noteShortParty("ФОП Кар'єрна Ірина")},
        {a: 'Головне Управління Державноі Казначейськоі Служби', b: 'Броварський ЗБК', c: 'Фарлеп-Інвест', d: 'Компанія Інноваційних Рішень', e: 'Прокуда Н. Г.', f: 'Нова Пошта', g: 'Управляюча Компанія Крістал Авеню', h: true, i: "Кар'єрна Ірина"});
}

// ── v7.575: податки до основного платежу ──
{
  const vv = (n) => /var [A-Z_]+ *=[\s\S]*?;\n/.exec(src.slice(src.indexOf('var ' + n + ' ')))[0];
  const names = ['_opexTaxFold', '_opexTaxDayNum', '_opexTaxInfo', '_opexTaxLinks', '_opexTaxInfoBase', '_opexTaxBaseCat', '_opexTaxLabel'];
  const W = new Function(['OPEX_TAX_RATES', 'OPEX_TAX_MONTHS'].map(vv).join('\n') + names.map(fn).join('\n') + '\nreturn {' + names.join(',') + '};')();
  const T = 'ОВВ, ГОЛОВНЕ УПРАВЛІННЯ ДЕРЖАВНОІ КАЗНАЧЕЙСЬКОІ СЛУЖБИ';
  const pays = [
    {ref: 'm1', date: '05.10.2026', amount: 11550, purpose: 'Оплата за оренду приміщення за жовтень', counterparty: 'Андрєєва Ірина Анатоліївна', edrpou: '2911111111'},
    {ref: 't1', date: '05.10.2026', amount: 2700, purpose: '101 ПОДАТОК З ДОХОДІВ ІНШИХ НІЖ ЗАРОБІТНА ПЛАТА. ОРЕНДА', counterparty: T, edrpou: '37993783'},
    {ref: 't2', date: '05.10.2026', amount: 750, purpose: '101 ВІЙСЬКОВИЙ ЗБІР З ОРЕНДИ', counterparty: T, edrpou: '37993783'},
    {ref: 's1', date: '22.10.2026', amount: 5783.9, purpose: '101 ПДФО за першу половину жовтня 2026р.', counterparty: T, edrpou: '37993783'},
    {ref: 's2', date: '06.11.2026', amount: 3802.51, purpose: 'Сплата ПДФО за жовтень 2026', counterparty: T, edrpou: '37993783'},
    {ref: 'o1', date: '14.10.2026', amount: 27559.2, purpose: '101 ЄДИНИЙ ПОДАТОК З ФІЗИЧНИХ ОСІБ ЗА ВЕРЕСЕНЬ  2026Р.', counterparty: T, edrpou: '37993783'},
    {ref: 'c1', date: '04.11.2026', amount: 6046.04, purpose: 'Компенсація вартості освітніх послуг за серпень 2026р', counterparty: 'КЛАКОВ ПАВЛО АНДРІЙОВИЧ', edrpou: '3470711820'},
    {ref: 'c2', date: '04.11.2026', amount: 6046.04, purpose: 'Компенсація вартості освітніх послуг за серпень 2026р', counterparty: 'ФІЗОР ВІТАЛІНА СЕРГІЇВНА', edrpou: '3458912000'},
    {ref: 'c3', date: '04.11.2026', amount: 6046.04, purpose: 'Компенсація вартості освітніх послуг за вересень 2026р', counterparty: 'ФІЗОР ВІТАЛІНА СЕРГІЇВНА', edrpou: '3458912000'},
    {ref: 'ct1', date: '04.11.2026', amount: 1413.36, purpose: 'ПДФО з доходів по відшкодуванню компенсації вартості освітніх послуг за Кучанська Роксолана Олександрівна 3470711820.', counterparty: T, edrpou: '37993783'},
    {ref: 'ct2', date: '04.11.2026', amount: 1413.36, purpose: 'ПДФО з доходів по відшкодуванню компенсації вартості освітніх послуг за Фізор Віталіна Сергіївна 3458912000.', counterparty: T, edrpou: '37993783'},
    {ref: 'ct3', date: '04.11.2026', amount: 1413.36, purpose: 'ПДФО з доходів по відшкодуванню компенсації вартості освітніх послуг за Фізор Віталіна Сергіївна 3458912000.', counterparty: T, edrpou: '37993783'},
    {ref: 'u1', date: '04.11.2026', amount: 392.6, purpose: 'ВІЙСЬКОВИЙ ЗБІР з доходів по відшкодуванню компенсації вартості освітніх послуг за Тонченко Сергій Анатолійович 3203915810', counterparty: T, edrpou: '37993783'},
    {ref: 'm2', date: '10.11.2026', amount: 7700, purpose: 'Послуги згідно акту', counterparty: 'Петренко Олег Іванович', edrpou: '3011111111'},
    {ref: 'b1', date: '11.11.2026', amount: 1800, purpose: '101 ПДФО', counterparty: T, edrpou: '37993783'},
    {ref: 'b2', date: '11.11.2026', amount: 500, purpose: '101 Військовий збір', counterparty: T, edrpou: '37993783'}];
  const r = W._opexTaxLinks(pays), L = r.links;
  check('OPEX податки: звʼязок з основним платежем', 'оренда за призначенням; компенсація за ІПН (навіть якщо отримувач з іншим прізвищем); два однакові платежі особі — по одному податку; без опису — фізособа ±1 день і ставка; ЗП / єдиний — не чіпаємо; не знайдено — у списку',
        {t1: L.t1 && L.t1.mainRef, t2: L.t2 && L.t2.mainRef, s: [!!L.s1, !!L.s2, !!L.o1], ct1: L.ct1 && L.ct1.mainRef, ct23: [L.ct2 && L.ct2.mainRef, L.ct3 && L.ct3.mainRef].sort(),
         b: [L.b1 && L.b1.mainRef, L.b2 && L.b2.mainRef], unl: r.unlinked, lbl: W._opexTaxLabel(['ВЗ', 'ПДФО'], 'оренда'), cat: W._opexTaxBaseCat('оренда')},
        {t1: 'm1', t2: 'm1', s: [false, false, false], ct1: 'c1', ct23: ['c2', 'c3'], b: ['m2', 'm2'], unl: ['u1'], lbl: 'ПДФО + ВЗ з оренди', cat: 'ОРЕНДА'});
  // той самий код у звірці (reconcile.html)
  const hp = path.join(__dirname, '..', '..', 'reconcile.html');
  if (fs.existsSync(hp)){
    const h = fs.readFileSync(hp, 'utf8');
    const hf = (name) => { const i = h.indexOf('function ' + name + '('); let d = 0;
      for (let k = h.indexOf('{', i); k < h.length; k++){ if (h[k] === '{') d++; else if (h[k] === '}' && --d === 0) return h.slice(i, k + 1); } };
    check('OPEX податки: звірка = бекенд', 'функції звʼязку податку в reconcile.html і crm_script.js однакові',
          names.filter((n) => hf(n) !== fn(n)), []);
  }
  // примітка: податок рядком під основним
  const N = new Function(['NOTE_TAG'].map(vv).join('\n') + ['_noteRound', '_noteUah', '_noteComposeOpex', '_opexTaxLabel'].map(fn).join('\n') + '\nreturn {_noteComposeOpex};')();
  check('OPEX примітка: податок під основним платежем', 'сума статті = основний + податок; «вручну» не зʼявляється',
        N._noteComposeOpex('Оренда', 'жовтень', [{day: '05.10', sort: '1005', who: 'Андрєєва І.А.', amount: 11550, tax: {kinds: ['ПДФО', 'ВЗ'], base: 'оренда', amount: 3450}}], 15000),
        'ⓘ ОРЕНДА · жовтень · 15 000 ₴\n\n05.10  Андрєєва І.А. — 11 550 ₴\n       ПДФО + ВЗ з оренди — 3 450 ₴');
}

// ── v7.562: Предметники (фронт) — «Не передбачено для цієї групи» лише за справжньою нормою 0 ──
{
  const hp = path.join(__dirname, '..', '..', 'predmetnyky.html');
  if (fs.existsSync(hp)){
    const h = fs.readFileSync(hp, 'utf8');
    const hf = (name) => { const i = h.indexOf('function ' + name + '('); let d = 0;
      for (let k = h.indexOf('{', i); k < h.length; k++){ if (h[k] === '{') d++; else if (h[k] === '}' && --d === 0) return h.slice(i, k + 1); } };
    const hv = (name) => /var [A-Z_]+ *=[\s\S]*?;\n/.exec(h.slice(h.indexOf('var ' + name + ' ')))[0];
    const code = ['PRED_LVIV_LOCATIONS', 'GT_TRANSLIT', 'GT_BABY', 'PRED_GROUP_TYPES', 'PRED_UNLIMITED_SUBJ'].map(hv).join('\n') +
      ['groupLetters', 'groupKey', 'canonGroupType', 'normalizeGroupType', 'predLocToRegion', 'isSchoolLoc', 'locNormFor',
       'getNormFor', 'catGroupAllowed', 'regionPlansSubject'].map(hf).join('\n');
    const z = {miniBaby: 0, Baby: 0, Find: 0, Study: 0, Preschool: 0};
    const S = {locations: [], catalog: [],
      norms: {'Київ': {'Фітнес': z, 'Логопед': {miniBaby: 0, Baby: 4, Find: 4, Study: 4, Preschool: 4}}},
      locNorms: {'Нац.Гвардії (Благо)': {'Польська': {byName: {beshketnyky: 0}, byType: {Find: 8}, any: null}}}};
    const P = new Function('S', code + ';return {normalizeGroupType, locNormFor, getNormFor, catGroupAllowed, regionPlansSubject, isSchoolLoc};')(S);
    const forb = (loc, subj, g) => { const gt = P.normalizeGroupType(g), ln = P.locNormFor(loc, subj, g, gt);
      const unl = (!gt && ln === null) || (ln === null && !P.isSchoolLoc(loc) && !P.regionPlansSubject(loc, subj));
      return (!unl && (unl ? 0 : P.getNormFor(loc, subj, gt, g)) <= 0) || !P.catGroupAllowed(loc, subj, g, gt); };
    check('предметники: Фітнес Хіміків', 'рядок нулів у матриці = норму не задано → можна відмічати; справжній 0 лишається забороною',
          {fit: ['Карапузи', 'Малюки', 'Розумники'].map((g) => forb('Хіміків (Благо)', 'Фітнес', g)),
           log: [forb('Бровари', 'Логопед', 'miniBaby-ki'), forb('Бровари', 'Логопед', 'Baby-ki')],
           pol: forb('Нац.Гвардії (Благо)', 'Польська', 'Бешкетники')},
          {fit: [false, false, false], log: [true, false], pol: true});
  }
}

console.log('\n' + pass + ' пройдено, ' + fail + ' з помилкою — ' + path.basename(SRC));
process.exit(fail ? 1 : 0);
