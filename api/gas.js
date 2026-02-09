// api/gas.js — Proxy serverless na Vercel apontando para seu Apps Script
module.exports = async (req, res) => {
  try {
    const GAS_BASE = 'https://script.google.com/macros/s/AKfycbyUbzgwsF1OFYFl1I341Umfx96Iz9F5XNqutcDECxyR7NsLgJidk4mkHklTHF_EeEJh/exec';
    const qs = new URLSearchParams(req.query).toString();
    const target = qs ? `${GAS_BASE}?${qs}` : GAS_BASE;

    const options = { method: req.method };
    if (req.method === 'POST') {
      options.headers = { 'Content-Type': 'application/json' };
      options.body = JSON.stringify(req.body || {});
    }

    const upstream = await fetch(target, options);
    const status = upstream.status;
    const contentType = upstream.headers.get('content-type') || '';

    if (contentType.includes('application/json')) {
      const data = await upstream.json();
      return res.status(status).json(data);
    } else {
      const text = await upstream.text();
      return res.status(status).json({ ok:false, error:'Upstream returned non-JSON', status, raw: text?.slice(0,800) });
    }
  } catch (err) {
    return res.status(500).json({ ok:false, error:String(err) });
  }
};
