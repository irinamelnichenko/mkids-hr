#!/usr/bin/env node
// v7.559: звірка живого списку боржників бота з еталоном CFO (fixtures_bank_etalon_*.json).
// Запуск: curl …?action=bankGroupsPreview&loc=Осокорки&kinds=debt | node _backend/tests/bank_etalon.js [еталон.json]
// Вихід 1, якщо є розбіжності (дитина, позиція або сума).
'use strict';
const fs = require('fs'), path = require('path');
// v7.606: --offline — без сервера: знімок входів (fixtures_bank_input_osokorky_1127.json: рядки «Оплати-Рік», картки, відкладені, поправки CFO)
//   проганяється через ТІ САМІ функції з crm_script.js (_bankKidsFromRows → _bankDebtorsCalc / _bankEntryCalc → _bankDebtorsTexts) і
//   звіряється з тим, що бот показав 27.11 (fixtures_bank_etalon_osokorky_1127.json). Так predeploy ловить будь-яку зміну логіки боржників.
//   Еталон CFO 18.11 (fixtures_bank_etalon_osokorky.json) — для живої звірки: з 18.11 батьки доплатили (Пархоменко, Пєший, Герасимчук, Лінник…).
const OFF = process.argv.indexOf('--offline') >= 0, args = process.argv.slice(2).filter((a) => a !== '--offline');
const ET = JSON.parse(fs.readFileSync(args[0] || path.join(__dirname, OFF ? 'fixtures_bank_etalon_osokorky_1127.json' : 'fixtures_bank_etalon_osokorky.json'), 'utf8'));
const pv = OFF ? (function(){
  const B = require('./_bankload.js')(args[1]);
  const I = JSON.parse(fs.readFileSync(path.join(__dirname, 'fixtures_bank_input_osokorky_1127.json'), 'utf8'));
  const KF = B._bankKidsFromRows(I.rows, I.month), kids = Object.keys(KF.by).map((k) => KF.by[k]);
  I.net.forEach((x) => { const k = KF.by[B._dopAnKid(x.name)]; if (k){ k.net = x.net; k.marks = x.marks; } });
  const debtors = B._bankHoldFilter(B._bankDebtorsCalc(kids, {}), I.hold);
  const entries = B._bankHoldFilter(B._bankEntryCalc(I.cards, KF.facts, I.year, I.today, I.entryAdj), I.hold);
  const texts = B._bankDebtorsTexts(ET.loc, I.monthName, debtors, entries, I.dayLabel, I.prevMonth);
  return {locs: {[ET.loc]: {texts: {debt: texts}}}};
})() : JSON.parse(fs.readFileSync(0, 'utf8'));
const txt = ((pv.locs || {})[ET.loc] || {}).texts.debt.join('\n');
const num = (s) => Number(String(s).replace(/[^\d]/g, ''));
const html = (s) => s.replace(/<[^>]+>/g, '').replace(/&#39;|&apos;/g, "'").replace(/&amp;/g, '&');
const bot = {}; let cur = null, block = '';
txt.split('\n').map(html).forEach((l) => {
  if (/^Винні за кілька/.test(l)) block = 'multi'; else if (/^Тільки навчання/.test(l)) block = 'navch'; else if (/^Тільки допи/.test(l)) block = 'dop'; else if (/^Тільки вступний/.test(l)) block = 'entry';
  const pos = /^\s{2}(навчання|допи|вступний)\s+([\d ]+)/.exec(l);
  if (pos && cur){ cur[{навчання: 'navch', допи: 'dop', вступний: 'entry'}[pos[1]]] = num(pos[2]); return; }
  const kid = /^(\S.*?\S)\s{2,}([\d ]+?)(?:\s{2}\(.*\))?$/.exec(l);
  if (kid && block){ cur = bot[kid[1]] = {name: kid[1]}; if (block !== 'multi') cur[block] = num(kid[2]); }
});
const key = (n) => n.toLowerCase().replace(/[’ʼ`']/g, "'");
const E = {}; ET.kids.forEach((k) => { E[key(k.name)] = k; });
const B = {}; Object.values(bot).forEach((k) => { B[key(k.name)] = k; });
const diffs = [];
new Set(Object.keys(E).concat(Object.keys(B))).forEach((k) => {
  const e = E[k] || {}, b = B[k] || {};
  ['navch', 'dop', 'entry'].forEach((f) => { if ((e[f] || 0) !== (b[f] || 0)) diffs.push({name: (E[k] || B[k]).name, item: f, etalon: e[f] || 0, bot: b[f] || 0}); });
});
const tot = Object.values(B).reduce((a, k) => a + (k.navch || 0) + (k.dop || 0) + (k.entry || 0), 0);
console.log(JSON.stringify({etalonTotal: ET.total, botTotal: tot, diffs}, null, 1));
process.exit(diffs.length ? 1 : 0);
