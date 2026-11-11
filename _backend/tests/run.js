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

console.log('\n' + pass + ' пройдено, ' + fail + ' з помилкою — ' + path.basename(SRC));
process.exit(fail ? 1 : 0);
