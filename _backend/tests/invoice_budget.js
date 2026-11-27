#!/usr/bin/env node
// v7.606: рахунки грудня 2026 = «Бюджет» грудня в Payment по всіх локаціях (навч і доп окремо) — 0 розбіжностей.
//   Знімок «Оплати-Рік» 27.11 (повний перезбір з Payment) → _invPayByName (та сама функція, що в getInvoiceListData) → сума рахунку;
//   Payment → бюджет; _invVerifyDiff (та сама звірка, що перед розсилкою). Плюс контроль: стара сума Голосієва (3 950 при 4 300) ловиться.
'use strict';
const fs = require('fs'), path = require('path');
const B = require('./_bankload.js')(process.argv[2]);
const FX = JSON.parse(fs.readFileSync(path.join(__dirname, 'fixtures_invoices_dec_2026.json'), 'utf8'));
const H = ['Локація', 'Напрямок', 'Тип', 'Група', 'Вихователь', "Ім'я дитини"];
B.MONTHS_CAL.forEach((m) => { ['Факт-навч', 'Факт-доп', 'Факт-вступ', 'Бюджет-навч', 'Бюджет-доп', 'Статус'].forEach((x) => H.push(m + '-' + x)); });
H.push('Факт-Рік', 'Бюджет-Рік', 'Борг-Рік', 'Зібрано-На-Сьогодні', 'Оновлено', 'Вибув');
const iBs = H.indexOf('Грудень-Бюджет-навч'), iBe = H.indexOf('Грудень-Бюджет-доп'), iDep = H.indexOf('Вибув');
let bad = 0, total = 0; const out = [];
Object.keys(FX.locs).forEach((loc) => {
  const L = FX.locs[loc];
  const pvals = [H].concat(L.rows.map((r) => { const a = H.map(() => ''); a[0] = loc; a[5] = r[0]; a[iBs] = r[1]; a[iBe] = r[2]; a[iDep] = r[3]; return a; }));
  const P = B._invPayByName(pvals, loc, FX.payMonth);
  const live = {}; L.rows.forEach((r) => { if (r[3] === 'так' || !String(r[0]).trim()) return; const k = B._dopAnKid(r[0]), x = live[k] || (live[k] = {bs: 0, be: 0}); x.bs += Number(r[1]) || 0; x.be += Number(r[2]) || 0; });
  const kids = L.children.map((c) => ({name: c.name, prepayCovered: c.cover ? 'так' : '', paymentSum: c.cover ? 0 : (P.pay[c.name] || 0), extrasOnlySum: P.dop[c.name] || 0}));
  const diff = B._invVerifyDiff(kids, live), n = Object.keys(diff).length;
  bad += n; total += kids.length; out.push(loc + ': ' + kids.length + ' дітей, розбіжностей ' + n + (n ? ' — ' + JSON.stringify(diff) : ''));
});
const G = FX.locs['Голосієво'], bl = G.children.find((c) => c.name === 'Благовісна Аліса');
const ctl = B._invVerifyDiff([{name: bl.name, paymentSum: 23000, extrasOnlySum: 3950}], {[B._dopAnKid(bl.name)]: {bs: 23000, be: 4300}});
const ctlOk = /допи: рахунок 3950 ≠ Payment 4300/.test(ctl[bl.name] || '');
console.log(out.join('\n') + '\nУсього ' + total + ' дітей, розбіжностей ' + bad + ' · контроль (стара сума 3 950 при Payment 4 300 ловиться): ' + (ctlOk ? 'так' : 'НІ'));
process.exit(bad || !ctlOk ? 1 : 0);
