# TODO — mkids HR CRM

## Пріоритет 1 — Відпустки

- [ ] **Задача 1 (clients.html)** — UI редагування/скасування відпустки
  - Кнопки ✏️ / 🚫 в `renderAbsEntry` для HR на done-записах
  - Нові функції: `editAbsencePeriod`, `cancelAbsencePeriod`
  - Статус `cancelled` виключається з `recalcDiscount`

- [ ] **Задача 2 (crm_script.js)** — `writeAbsenceToPayment` / `removeAbsenceFromPayment`
  - Писати в BL-BO колонки Payment-файлів при переході pending→done
  - `detectAbsenceCols` вже є
  - Тільки після завершення Задачі 1

- [ ] **Задача 3** — `importPaymentAbsencesToCRM`
  - 300+ дітей, спочатку `dryRun`
  - Тільки після Задач 1 + 2

## Пріоритет 2 — Синхронізація

- [ ] Зменшити `CLI_CACHE` TTL з 15хв до 1-2хв
- [ ] Верифікувати зворотню синхронізацію (директор змінює → HR бачить)
- [ ] Мігрувати localStorage інших користувачів (відпустки можуть бути в `entryFeeSchedule`)
  - Console-скрипт: перенести записи з `type:'vacation'|'sick'` з `entryFeeSchedule` → `absences`

## Пріоритет 3 — Інше

- [ ] Медсестра: табель → синхронізація з Google Sheets
- [ ] Імпорт відпусток Осокорки з Payment-файлів (~95 відпусток, `dryRunImportAbsences` готовий)
