// v7.606: чисті функції груп «Банк» і рахунків прямо з crm_script.js (без Google-сервісів) — для bank_etalon.js і invoice_budget.js
'use strict';
const fs = require('fs'), path = require('path');
module.exports = function(srcPath){
  const src = fs.readFileSync(srcPath || path.join(__dirname, '..', 'crm_script.js'), 'utf8');
  const fn = (name) => { const i = src.indexOf('\nfunction ' + name + '('); if (i < 0) throw new Error('функцію ' + name + ' не знайдено');
    let d = 0; for (let k = src.indexOf('{', i); k < src.length; k++){ if (src[k] === '{') d++; else if (src[k] === '}' && --d === 0) return src.slice(i, k + 1); } };
  const varLine = (name) => { const m = new RegExp('\\nvar ' + name + '\\s*=\\s*([^\\n]*);').exec(src); if (!m) throw new Error('змінну ' + name + ' не знайдено'); return 'var ' + name + ' = ' + m[1] + ';'; };
  const V = ['_APOSTROPHES', 'INV_TG_MAX', 'BANK_DEBT_MIN', 'BANK_DEBT_LIST_MIN', 'MONTHS_CAL'].map(varLine);
  const F = ['toNum', '_lat2cyr', '_nameFold', '_nameFoldTight', '_journalNormName', '_bankWords', '_dopAnKid', '_levSalary', '_surnameSim', '_htmlEsc', '_invMoneyUa', '_bankMoney',
    '_bankChunks', '_bankInt', '_bankPadR', '_bankPadL', '_bankCode', '_bankKids', '_bankDebtGroups', '_bankCarried', '_bankDebtorsTexts',
    '_bankKidsFromRows', '_bankDebtorsCalc', '_bankEntryCalc', '_bankSuspectDebts', '_bankHoldFilter', '_invR2', '_prevYM', '_invBankView', '_invVerifyDiff',
    '_invCellErr', '_invPrecheckBlock', '_invPrecheckItems', '_yearlyDepartedIdx', '_isYearlyDeparted', '_invPayByName'].map(fn);
  const names = ['_bankKidsFromRows', '_bankDebtorsCalc', '_bankEntryCalc', '_bankSuspectDebts', '_bankHoldFilter', '_bankDebtorsTexts', '_invBankView', '_invVerifyDiff', '_invPrecheckItems', '_invPayByName', '_dopAnKid', 'MONTHS_CAL'];
  return new Function('Logger', V.concat(F).join('\n') + '\nreturn {' + names.join(',') + '};')({log(){}});
};
