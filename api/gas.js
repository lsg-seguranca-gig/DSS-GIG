/**
 * Vercel Serverless Function — /api/gas
 * Proxy entre o gestor.html e o Google Apps Script (GAS).
 *
 * ⚙️  Configure a variável de ambiente no painel da Vercel:
 *     GAS_URL = https://script.google.com/macros/s/AKfycbxDcJJK7Qaji5Bwi_ZwOI4TIFcqNGBZDX_HtHT2zaafrWNXyu8kQ8BWk7pTOq4eI36x/exec
 *
 * Suporta GET (com query params) e POST (com body JSON).
 * Adiciona os headers CORS necessários.
 */

const GAS_URL = process.env.GAS_URL;

export default async function handler(req, res) {
  // ── CORS ────────────────────────────────────────────────────────────────────
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  // ── Valida configuração ─────────────────────────────────────────────────────
  if (!GAS_URL) {
    return res.status(500).json({
      ok: false,
      error: 'Variável de ambiente GAS_URL não configurada no Vercel.'
    });
  }

  try {
    let gasResponse;

    if (req.method === 'POST') {
      // ── POST: repassa action como query param + body como JSON ──────────────
      const action = req.query.action || '';
      const targetUrl = `${GAS_URL}?action=${encodeURIComponent(action)}`;

      gasResponse = await fetch(targetUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(req.body || {}),
        redirect: 'follow',   // GAS redireciona para a URL final — importante!
      });

    } else {
      // ── GET: repassa todos os query params para o GAS ──────────────────────
      const params = new URLSearchParams(req.query);
      const targetUrl = `${GAS_URL}?${params.toString()}`;

      gasResponse = await fetch(targetUrl, {
        method: 'GET',
        redirect: 'follow',   // GAS redireciona para a URL final — importante!
      });
    }

    // ── Lê e retorna a resposta do GAS ─────────────────────────────────────
    const text = await gasResponse.text();

    // Tenta parsear como JSON; se falhar, retorna erro descritivo
    let json;
    try {
      json = JSON.parse(text);
    } catch {
      console.error('[gas proxy] Resposta não-JSON do GAS:', text.slice(0, 500));
      return res.status(502).json({
        ok: false,
        error: 'GAS retornou resposta não-JSON. Verifique se o script está publicado corretamente.',
        raw: text.slice(0, 300)
      });
    }

    return res.status(200).json(json);

  } catch (err) {
    console.error('[gas proxy] Erro ao chamar GAS:', err);
    return res.status(502).json({
      ok: false,
      error: 'Falha ao contactar o Google Apps Script: ' + String(err.message || err)
    });
  }
}
