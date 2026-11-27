#!/bin/bash
# Перед кожним деплоєм бекенду: синтаксис + тести на реальних випадках; лише тоді crm_script.js іде в буфер.
# Запуск: bash _backend/predeploy.sh   (з кореня репозиторію або будь-звідки)
set -e
DIR="$(cd "$(dirname "$0")" && pwd)"
F="$DIR/crm_script.js"
node -e "new Function(require('fs').readFileSync(process.argv[1],'utf8'))" "$F" || { echo "✗ синтаксична помилка в crm_script.js — у буфер НЕ скопійовано"; exit 1; }
node "$DIR/tests/run.js" "$F" || { echo "✗ тести не пройшли — у буфер НЕ скопійовано"; exit 1; }
node "$DIR/tests/bank_etalon.js" --offline "" "$F" > /dev/null || { node "$DIR/tests/bank_etalon.js" --offline "" "$F"; echo "✗ еталон боржників груп «Банк» розійшовся — у буфер НЕ скопійовано"; exit 1; }   # v7.606
node "$DIR/tests/invoice_budget.js" "$F" | tail -1 || true
node "$DIR/tests/invoice_budget.js" "$F" > /dev/null || { node "$DIR/tests/invoice_budget.js" "$F"; echo "✗ рахунки ≠ бюджет Payment — у буфер НЕ скопійовано"; exit 1; }   # v7.606
pbcopy < "$F"
echo "✓ $(grep -m1 -o "pong v[0-9.]*" "$F" | sed 's/pong //') у буфері · md5 файлу $(md5 -q "$F") · буфера $(pbpaste | md5 -q)"
