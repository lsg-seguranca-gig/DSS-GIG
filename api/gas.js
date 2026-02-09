// api/gas.js — Vercel Function
// Mantém proxy para o Apps Script e acrescenta endpoints locais:
//  - action=registros&matricula=...            (lê a aba Registros via Google Sheets API)
//  - action=treinamentos_nao_assistidos&matricula=... (compõe registros + treinamentos e devolve apenas não assistidos)

const { google } = require('googleapis');

module.exports = async (req, res) => {
  try {
    const action = String(req.query?.action || '').trim();

    // --- Handler local: REGISTROS ---
    if (action === 'registros') {
      const matricula = String(req.query?.matricula || '').trim();

      // validação básica
      if (!/^\d{1,15}$/.test(matricula)) {
        return res.status(400).json({ ok: false, error: 'Matrícula inválida' });
      }

      const SPREADSHEET_ID = process.env.SPREADSHEET_ID;
      if (!SPREADSHEET_ID) {
        return res.status(500).json({ ok: false, error: 'SPREADSHEET_ID não configurado' });
      }

      // autenticação Service Account
      const auth = new google.auth.JWT(
        process.env.GS_CLIENT_EMAIL,
        null,
        (process.env.GS_PRIVATE_KEY || '').replace(/\\n/g, '\n'),
        ['https://www.googleapis.com/auth/spreadsheets.readonly']
      );

      const sheets = google.sheets({ version: 'v4', auth });
      const RANGE_REGISTROS = process.env.RANGE_REGISTROS || 'Registros!A1:I';

      // Esperado: cabeçalho na linha 1: Timestamp, Matricula, Nome, Setor, SemanaISO, TituloVideo, URLVideo, AssinaturaPNG, DeviceInfo
      const resp = await sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: RANGE_REGISTROS
      });

      const rows = resp.data.values || [];
      if (!rows.length) {
        return res.status(200).json({ ok: true, data: [] });
      }

      const [header, ...data] = rows;
      const idx = Object.fromEntries(header.map((h, i) => [h, i]));

      // filtra por matrícula e retorna só os campos necessários para comparar
      const filtered = data
        .filter(r => (r[idx['Matricula']] || '').trim() === matricula)
        .map(r => ({
          Timestamp: r[idx['Timestamp']] || '',
          Matricula: r[idx['Matricula']] || '',
          SemanaISO: r[idx['SemanaISO']] || '',
          TituloVideo: r[idx['TituloVideo']] || '',
          URLVideo: r[idx['URLVideo']] || ''
        }));

      return res.status(200).json({ ok: true, data: filtered });
    }

    // --- Handler composto (opcional): TREINAMENTOS_NAO_ASSISTIDOS ---
    if (action === 'treinamentos_nao_assistidos') {
      const matricula = String(req.query?.matricula || '').trim();
      if (!/^\d{1,15}$/.test(matricula)) {
        return res.status(400).json({ ok: false, error: 'Matrícula inválida' });
      }

      // 1) Busca registros localmente (mesma lógica do handler acima)
      const registrosResp = await callSelf(req, 'registros');
      if (!registrosResp.ok) {
        return res.status(registrosResp.status || 500).json(registrosResp.body);
      }
      const registros = registrosResp.body.data || [];
      const watchedByKey = new Set(registros.map(r => `${r.SemanaISO}@@${r.TituloVideo}`));

      // 2) Busca treinamentos no GAS (proxy)
      const tResp = await proxyToGAS(req, res, 'treinamentos', { returnResponse: true });
      if (!tResp.ok) {
        return res.status(tResp.status || 500).json(tResp.body);
      }
      const todosTreinamentos = tResp.body?.data || [];

      // 3) Filtra e devolve só os não assistidos (por SemanaISO+Titulo)
      const naoAssistidos = todosTreinamentos.filter(t => !watchedByKey.has(`${t.SemanaISO}@@${t.Titulo}`));

      return res.status(200).json({ ok: true, data: naoAssistidos });
    }

    // --- PROXY padrão para Apps Script (todas as outras ações) ---
    return await proxyToGAS(req, res);

  } catch (err) {
    return res.status(500).json({ ok: false, error: String(err) });
  }
};

// ------------------------------------------------------------------
// Utilitários
// ------------------------------------------------------------------

// Reaproveita o próprio endpoint (para encadear a leitura de 'registros')
async function callSelf(req, actionName) {
  const baseUrl =
    process.env.VERCEL_URL
      ? `https://${process.env.VERCEL_URL}`
      : `${req.headers['x-forwarded-proto'] || 'https'}://${req.headers.host}`;

  const url = new URL(`${baseUrl}${req.url.split('?')[0]}`);
  const qs = new URLSearchParams(req.query);
  qs.set('action', actionName);
  url.search = qs.toString();

  const r = await fetch(url.toString(), { method: 'GET' });
  const contentType = r.headers.get('content-type') || '';
  const body = contentType.includes('application/json') ? await r.json() : { ok: false, error: 'Non-JSON' };
  return { ok: r.ok, status: r.status, body };
}

// Proxy para o Apps Script (mantém seu comportamento original)
async function proxyToGAS(req, res, forceAction, opts = {}) {
  const GAS_BASE = 'https://script.google.com/macros/s/AKfycbyUbzgwsF1OFYFl1I341Umfx96Iz9F5XNqutcDECxyR7NsLgJidk4mkHklTHF_EeEJh/exec';
  const qs = new URLSearchParams(req.query);
  if (forceAction) qs.set('action', forceAction);
  const target = qs.toString() ? `${GAS_BASE}?${qs.toString()}` : GAS_BASE;

  const options = { method: req.method };
  if (req.method === 'POST') {
    options.headers = { 'Content-Type': 'application/json' };
    options.body = JSON.stringify(req.body || {});
  }

  const upstream = await fetch(target, options);
  const status = upstream.status;
  const contentType = upstream.headers.get('content-type') || '';

  if (opts.returnResponse) {
    if (contentType.includes('application/json')) {
      const data = await upstream.json();
      return { ok: upstream.ok, status, body: data };
    } else {
      const text = await upstream.text();
      return { ok: false, status, body: { ok: false, error: 'Upstream returned non-JSON', status, raw: text?.slice(0, 800) } };
    }
  }

  if (contentType.includes('application/json')) {
    const data = await upstream.json();
    return res.status(status).json(data);
  } else {
    const text = await upstream.text();
    return res
      .status(status)
      .json({ ok: false, error: 'Upstream returned non-JSON', status, raw: text?.slice(0, 800) });
  }
}
