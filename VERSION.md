# Джерело правди

**`~/Desktop/BOT/HR mkids_bot/crm_script.js`** — єдине джерело правди для бекенду.

| | |
|---|---|
| Версія у репозиторії | **v7.168** |
| md5 | `8ed4858c8507811a58b64066c2ebeae2` |
| Розмір | 1382043 байт |
| Зафіксовано | 2026-08-19 |
| Версія у проді | **v7.160** — v7.168 ЩЕ НЕ ЗАДЕПЛОЄНО |
| Перевірка | `action=ping` → має стати `pong v7.168` після деплою |

## Як це працює

Бекенд **не деплоїться з git**. Файл вставляється вручну в редактор Apps Script.
Тому git — це реєстр того, що *має* бути в проді, а не те, що там є автоматично.

Після будь-якої правки в редакторі: скопіювати код назад сюди, звірити md5, закомітити.
Розбіжність md5 = локальна копія і прод розійшлись.

Перевірка версії проду:
```
curl -sL "https://script.google.com/macros/s/AKfycbyTSUVlaN4-PpXe47zCSmhVs0Qxy1FDXG_XsB4zcKNpqBxdhDtS9ibM4YFGkGjmPQDFWQ/exec?action=ping"
```

## Застарілі копії — НЕ використовувати

- `~/Downloads/crm_script.js` — v6.11, травень 2026
- `crm_script.js.bak_16152` — v7.26, не відстежується git
- `~/Desktop/BOT/_backend_snapshots/`, `_RESCUE_19-08/` — архів аварії 19.08
- `/private/tmp/mk_recon`, `/private/tmp/mk_fresh2` — тимчасові, /tmp чиститься при перезавантаженні

## Разові дії після деплою v7.153

1. `IMPORT_LEADS_HISTORY_DRYRUN()` → перевірити лог → `IMPORT_LEADS_HISTORY_APPLY()`
   (створює аркуш `Ліди_Історія`; повторний запуск перебудовує його, не дублює).
2. `SETUP_MIRROR_SPREADSHEET_DRYRUN()` → `SETUP_MIRROR_SPREADSHEET()` — структура
   в окремій таблиці 1RLdzmff… + Script Property `LEADS_MIRROR_ID`.
   Потім `MIRROR_RESYNC_DRYRUN()` → `MIRROR_RESYNC_ALL()`.
3. `SEED_REPORT_RECIPIENTS()` → `installOwnerReportTrigger()` — тижневий звіт
   власнику щопонеділка о 9:00. Спершу написати боту в особисті, інакше
   Telegram не дасть йому почати діалог.
3. `installLeadsDigestTrigger()` — щоденне зведення по лідах о 9:00.
3. `clearAuthLog()` — з v7.151, ще не виконано.
4. `SEED_DIRECTORS_ME_DRYRUN()` → `SEED_DIRECTORS_ME()` — вписати
   @Ira_Melnichenk0 на всі локації для тестів (потім замінити на директорок).
5. `REPAIR_LEAD_CONFLICTS()` → `..._APPLY()` — почистити картки, де вже
   стоять одночасно екскурсія і передзвін (правило діє лише на нові дії).

## Ремап сиріт у Табелі (після деплою v7.163)

`REMAP_ORPHAN_ATTENDANCE_DRYRUN()` → звірити лог → `REMAP_ORPHAN_ATTENDANCE_APPLY()`.
Оригінали зачеплених рядків лягають у JSON-файл `BACKUP_Табель_ремап_<stamp>.json`
НА DRIVE (не в книгу — вона на межі 10 млн комірок). Саме він, а не lastRow,
є точкою відкату, бо ремап ПРАВИТЬ і ВИДАЛЯЄ рядки, а не тільки дописує.

`DELETE_MIGBKP_DRYRUN()` → `DELETE_MIGBKP_APPLY()` — прибрати липневі MIGBKP
(кожен аркуш спершу вигружається в JSON у теку `MIGBKP_архів_<stamp>` на Drive).

`CREATE_MISSING_CARDS_DRYRUN()` → `..._APPLY()` — 3 картки для дітей,
чиї відмітки нема куди перенести. Потім ремап підхопить їх за точним ПІБ.

`AUDIT_SHEET_CELLS()` — скільки комірок з'їдає кожен аркуш і скільки звільнить
прибирання старих бекапів. Тільки читання.

## Прапорці

`TG_TEST` — константа в коді (рядок 627), **не** Script Property.
Поточне значення: `false` (робочі пороги нагадувань).

`AUTH_ENFORCE` — Script Property, **вимкнено** (`authEnforce:false`).
Фронтенд до вмикання підготовлено (v7.151), але сам прапорець свідомо не чіпали.
Перед вмиканням: очистити `Авторизація_Лог` через `clearAuthLog()` і подивитись 2–3 доби.

`WEB_LEAD_SECRET` — Script Property, **не задано**. Поки його немає, ендпойнт
форми з сайту відмовляє на будь-який виклик. Опис контракту — у `LEADS_WEB_FORM.md`.
