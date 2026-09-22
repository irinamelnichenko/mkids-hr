/**
 * m.kids · Telegram → Apps Script РЕЛЕЙ (Cloudflare Worker)
 * ═════════════════════════════════════════════════════════
 * НАВІЩО. Apps Script web app на POST відповідає 302-редіректом на
 * script.googleusercontent.com. Telegram за редіректами НЕ йде: він бачить
 * «Wrong response from the webhook: 302 Found», вважає доставку невдалою і
 * ретраїть той самий апдейт. Саме це показав getWebhookInfo бота рахунків
 * 22.09.2026. Воркер приймає POST від Telegram, ходить на exec САМ (fetch
 * слідує редіректу) і повертає Telegram чистий 200.
 *
 * ШВИДКА ВІДПОВІДЬ. Telegram чекає відповідь ~60 c, а холодний Apps Script
 * буває повільним. Тому воркер відповідає 200 ОДРАЗУ, а запит до exec
 * доопрацьовує у waitUntil — апдейт не ретраїться і не губиться.
 *
 * МАРШРУТИ (шлях вибирає бота, секрет у query лишається як є):
 *   POST /invoice?s=<INVOICE_WEBHOOK_SECRET>  → exec?action=tgInvoiceWebhook&s=…
 *   POST /leads?s=<TG_WEBHOOK_SECRET>         → exec?action=tgWebhook&s=…
 *
 * ЗМІННІ (Cloudflare → Worker → Settings → Variables):
 *   EXEC_URL — https://script.google.com/macros/s/<deployment-id>/exec
 *
 * ДЕПЛОЙ: Cloudflare dashboard → Workers & Pages → Create → Worker →
 * вставити цей файл → Deploy. URL виду https://<name>.<subdomain>.workers.dev.
 * Далі в Apps Script:
 *   POST {action:'tgInvoiceSetWebhook', hookUrl:'https://<worker>/invoice?s=<секрет>'}
 * Секрет — той самий INVOICE_WEBHOOK_SECRET зі Script Properties.
 */

const ROUTES = {
  '/invoice': 'tgInvoiceWebhook',   // бот рахунків
  '/leads':   'tgWebhook',          // бот лідів
};

export default {
  async fetch(request, env, ctx) {
    const url = new URL(request.url);

    if (request.method === 'GET') {
      // Проба живучості: /health → 200. Нічого не проксює.
      return new Response(url.pathname === '/health' ? 'ok' : 'm.kids tg relay', { status: 200 });
    }
    if (request.method !== 'POST') return new Response('method not allowed', { status: 405 });

    const action = ROUTES[url.pathname];
    if (!action) return new Response('unknown route', { status: 404 });
    if (!env.EXEC_URL) return new Response('EXEC_URL not configured', { status: 500 });

    // Тіло читаємо ТУТ: після повернення відповіді стрім запиту вже недоступний.
    const body = await request.text();
    // Секрет: з query (?s=…) або із заголовка, який Telegram шле сам — ми передаємо
    // його в setWebhook як secret_token. Заголовок рятує, коли вебхук зареєстрували
    // без ?s= : Apps Script звіряє саме e.parameter.s і без нього відповідає
    // «bad secret», мовчки відкидаючи апдейт (доставка при цьому виглядає успішною).
    const secret = url.searchParams.get('s')
      || request.headers.get('x-telegram-bot-api-secret-token')
      || '';
    const target = `${env.EXEC_URL}?action=${action}&s=${encodeURIComponent(secret)}`;

    // Telegram отримує 200 негайно; Apps Script доопрацьовує у фоні.
    ctx.waitUntil(
      fetch(target, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body,
        redirect: 'follow',          // ← те, чого не вміє Telegram
      }).catch(() => {})             // збій exec не має перетворюватись на ретрай Telegram
    );

    return new Response('ok', { status: 200 });
  },
};
