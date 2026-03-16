/**
 * Vercel Serverless Function — /api/gas
 * Proxy entre o gestor.html e o Google Apps Script (GAS).
 *
 * ⚙️  Configure a variável de ambiente no painel da Vercel:
 *     GAS_URL = https://script.google.com/macros/s/AKfycbxDcJJK7Qaji5Bwi_ZwOI4TIFcqNGBZDX_HtHT2zaafrWNXyu8kQ8BWk7pTOq4eI36x/exec
 *
 * Suporta GET (com query params) e POST (com body JSON).
 *
 * ── Por que o redirect manual? ──────────────────────────────────────────────
 * O Google Apps Script SEMPRE responde com um redirect 302 antes de retornar
 * o JSON final. O fetch nativo do Node.js segue o redirect automaticamente em
 * GET, mas em POST converte para GET (comportamento padrão HTTP). Além disso,
 * alguns ambientes Vercel bloqueiam redirects cross-origin silenciosamente.
 * A solução mais robusta é desabilitar o redirect automático (redirect:'manual'),
 * capturar o Location header e fazer uma segunda requisição GET para a URL final.
 */

const GAS_URL = process.env.GAS_URL;

export default async function handler(req, res) {
  // ── CORS ───────────────────────────────────────────────────────────────────
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  // ── Valida configuração ────────────────────────────────────────────────────
  if (!GAS_URL) {
    return res.status(500).json({
      ok: false,
      error: 'Variável de ambiente GAS_URL não configurada no Vercel.'
    });
  }

  try {
    const action = (req.query.action || '').toLowerCase();

    // ── Monta a URL alvo com todos os query params ─────────────────────────
    const params = new URLSearchParams(req.query);
    const targetUrl = `${GAS_URL}?${params.toString()}`;

    console.log(`[gas] ${req.method} action=${action} → ${targetUrl}`);

    // ── Primeira requisição — NÃO segue redirect automaticamente ──────────
    // O GAS retorna 302 → Location: <url_real_do_json>
    // Precisamos capturar esse Location e fazer uma segunda chamada GET.
    let firstRes;

    if (req.method === 'POST') {
      firstRes = await fetch(targetUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(req.body || {}),
        redirect: 'manual',
      });
    } else {
      firstRes = await fetch(targetUrl, {
        method: 'GET',
        redirect: 'manual',
      });
    }

    console.log(`[gas] primeira resposta: status=${firstRes.status}`);

    // ── Segue redirect manualmente (302 / 301 / 307 / 308) ────────────────
    let finalText;

    if (firstRes.status >= 300 && firstRes.status < 400) {
      const location = firstRes.headers.get('location');
      console.log(`[gas] redirect → ${location}`);

      if (!location) {
        return res.status(502).json({
          ok: false,
          error: `GAS retornou ${firstRes.status} sem header Location.`
        });
      }

      // Segunda chamada: sempre GET para a URL final
      const secondRes = await fetch(location, {
        method: 'GET',
        redirect: 'follow',
      });

      console.log(`[gas] segunda resposta: status=${secondRes.status}`);
      finalText = await secondRes.text();

    } else {
      // Não houve redirect — lê a resposta direto
      finalText = await firstRes.text();
    }

    console.log(`[gas] resposta final (primeiros 200 chars): ${finalText.slice(0, 200)}`);

    // ── Parseia e retorna ──────────────────────────────────────────────────
    let json;
    try {
      json = JSON.parse(finalText);
    } catch {
      console.error('[gas] resposta não-JSON:', finalText.slice(0, 500));
      return res.status(502).json({
        ok: false,
        error: 'GAS retornou resposta não-JSON. Verifique se o script está publicado corretamente.',
        raw: finalText.slice(0, 300)
      });
    }

    return res.status(200).json(json);

  } catch (err) {
    console.error('[gas] erro:', err);
    return res.status(502).json({
      ok: false,
      error: 'Falha ao contactar o Google Apps Script: ' + String(err.message || err)
    });
  }
}
