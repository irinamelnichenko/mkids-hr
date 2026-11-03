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

console.log('\n' + pass + ' пройдено, ' + fail + ' з помилкою — ' + path.basename(SRC));
process.exit(fail ? 1 : 0);
