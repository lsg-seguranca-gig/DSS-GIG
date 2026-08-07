// A API fica na raiz do projeto (/api/gas), independentemente da pasta
// onde o HTML está hospedado (ex.: /gestor/). Por isso usamos caminho absoluto.
const API_BASE = "/api/gas";

  // Controlo de cache manual e chamadas de API
  function apiFetch(params, options){
    const url = API_BASE + '?' + params + '&_t=' + Date.now();
    return fetch(url, Object.assign({ cache: 'no-store', headers: { 'Cache-Control': 'no-cache' } }, options || {}));
  }

  const LOGO_SRC = "/logo lsg.png";
  let LOGO_DATAURL = "";
  let registros = [];
  let registrosFiltrados = [];
  let feriasServidor = []; // Guarda as informações da planilha "Ferias" do Google Sheets

  /* ====================================================
   * LSG SKY CHEFS — LOGO EMBUTIDO (SVG) COMO FALLBACK GARANTIDO
   * Usado no PDF se lsg sky chefs logo.png não estiver acessível
   * ==================================================== */

  /* ====================================================
   * INSTRUTORES DO TREINAMENTO
   * As células de assinatura são desenhadas como campos em branco
   * pelo jsPDF (linha clássica de assinatura)
   * ==================================================== */
  const PDF_INSTRUTORES = [
    {
      nome: 'Vitor Hugo Teixeira da Silva',
      cargo: 'Supervisor de Segurança / LSA AVSEC / Ramp Safety Owner',
      matricula: '15623',
      assinatura: null
    },
    {
      nome: 'Francinele Ribeiro Machado',
      cargo: 'Assistente Administrativo / DLSA AVSEC / Ramp Safety Deputy',
      matricula: '16977',
      assinatura: null
  },
];

// Pré-carregamento do logo assim que a página abre (para o PDF não precisar esperar)
let _logoLoadPromise = null;
function loadLogoAsDataURL(){
  if (_logoLoadPromise) return _logoLoadPromise;
  LOGO_DATAURL = window._LOGO_FALLBACK_B64;
  _logoLoadPromise = Promise.resolve(LOGO_DATAURL);
  fetch(LOGO_SRC, { cache: 'force-cache' })
    .then(r => r.ok ? r.blob() : Promise.reject())
    .then(blob => new Promise((res, rej) => {
      const fr = new FileReader();
      fr.onload = () => res(fr.result);
      fr.onerror = rej;
      fr.readAsDataURL(blob);
    }))
    .then(dataUrl => { if (dataUrl) LOGO_DATAURL = dataUrl; })
    .catch(() => { /* logo embutido já está em LOGO_DATAURL */ });
  return _logoLoadPromise;
}
loadLogoAsDataURL();

document.getElementById('btnBuscar').addEventListener('click', buscar);
document.getElementById('btnPDF').addEventListener('click', gerarPDF);
document.getElementById('btnXLS').addEventListener('click', gerarXLS);
document.getElementById('btnAddFunc').addEventListener('click', novoFuncionarioViaPrompt);
document.getElementById('btnDelFunc').addEventListener('click', excluirFuncionarioViaPrompt);

['fMat','fNome','fDataInicio','fDataFinal','fSemanaTitulo'].forEach(id=>{
  const el = document.getElementById(id); 
  if(!el) return; 
  el.addEventListener('keydown', e=>{ if(e.key === 'Enter'){ e.preventDefault(); buscar(); } });
});

// Configuração dos Inputs de Data com suporte a Calendário Nativo
function setupDatePicker(textId, nativeId, calBtnId){
  const txt    = document.getElementById(textId);
  const native = document.getElementById(nativeId);
  const btn    = document.getElementById(calBtnId);
  if (!txt || !native || !btn) return;

  txt.addEventListener('input', function(){
    let v = this.value.replace(/\D/g, '');
    if (v.length > 8) v = v.slice(0, 8);
    if      (v.length >= 5) this.value = v.slice(0, 2) + '/' + v.slice(2, 4) + '/' + v.slice(4);
    else if (v.length >= 3) this.value = v.slice(0, 2) + '/' + v.slice(2);
    else                    this.value = v;
  });

  btn.addEventListener('click', function(){
    const m = txt.value.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
    if (m) native.value = `${m[3]}-${m[2]}-${m[1]}`;
    native.showPicker ? native.showPicker() : native.click();
  });

  native.addEventListener('change', function(){
    if (!this.value) return;
    const [y, mo, d] = this.value.split('-');
    txt.value = `${d}/${mo}/${y}`;
  });
}

setupDatePicker('fDataInicio', 'fDataInicioNative', 'btnCalInicio');
setupDatePicker('fDataFinal',  'fDataFinalNative',  'btnCalFinal');

['fDataInicio','fDataFinal'].forEach(id => {
  document.getElementById(id)?.addEventListener('keydown', e => { if(e.key === 'Enter'){ e.preventDefault(); buscar(); } });
});

// Gestão de chamadas paralelas de API em cache para evitar sobrecarga no servidor
const _apiCache = {};
function _cachedFetch(action){
  if (!_apiCache[action]) {
    _apiCache[action] = apiFetch('action=' + action)
      .then(r => r.ok ? r.json().catch(() => null) : null)
      .catch(() => null);
  }
  return _apiCache[action];
}

const _apiCacheComStatus = {};
function _cachedFetchComStatus(action){
  if (!_apiCacheComStatus[action]) {
    _apiCacheComStatus[action] = apiFetch('action=' + action)
      .then(async r => {
        if (!r.ok) return { ok: false, json: null };
        const j = await r.json().catch(() => null);
        return { ok: !!(j && j.ok !== false), json: j };
      })
      .catch(() => ({ ok: false, json: null }));
  }
  return _apiCacheComStatus[action];
}

function _parseRows(json){
  if (Array.isArray(json))         return json;
  if (Array.isArray(json?.data))   return json.data;
  if (Array.isArray(json?.result)) return json.result;
  if (typeof json === 'object' && json){
    const v = Object.values(json).find(v => Array.isArray(v)); 
    if (v) return v;
  }
  return [];
}

function _buildTituloOpts(mapaISO, prefixOpt){
  return prefixOpt + [...mapaISO.entries()]
    .sort((a,b)=>{ 
      const pa = dashParseSemanaISO(a[1]), pb = dashParseSemanaISO(b[1]); 
      if(pa.year !== pb.year) return pb.year - pa.year; 
      return pb.week - pa.week; 
    })
    .map(([t]) => `<option value="${escapeHtml(t)}">${escapeHtml(t)}</option>`).join('');
}

let _catalogoCarregado = false;
let _catalogoComplementado = false;

function _popularSelectsCatalogo(mapaISO, isoSet){
  if (mapaISO.size) {
    const opts = _buildTituloOpts(mapaISO, '');

    const selP = document.getElementById('fSemanaTitulo');
    const selAtualP = selP.value;
    selP.innerHTML = '<option value="">Selecione uma semana</option>' + opts;
    if (selAtualP && [...selP.options].some(o => o.value === selAtualP)) selP.value = selAtualP;

    const selD = document.getElementById('dashTitulo');
    const selAtualD = selD.value;
    selD.innerHTML = '<option value="">Selecione um título...</option>' + opts;
    if (selAtualD && [...selD.options].some(o => o.value === selAtualD)) selD.value = selAtualD;
  }

  if (isoSet.size) {
    const isoList = [...isoSet].sort(dashCompareSemanaISODesc);
    const selIso = document.getElementById('dashSemanaIso');
    if (selIso) selIso.innerHTML = '<option value="">Selecione...</option>' + isoList.map(sem => `<option value="${escapeHtml(sem)}">${escapeHtml(sem)}</option>`).join('');
  }
}

function _extrairTituloISO(rows, campoTitulo, mapaISO, isoSet){
  rows.forEach(r => {
    const t = String((campoTitulo ? r[campoTitulo] : (r.TituloVideo ?? r.Titulo ?? r.titulo)) ?? '').trim();
    const iso = String(r.SemanaISO ?? '').trim();
    if (t && !mapaISO.has(t)) mapaISO.set(t, iso);
    else if (t && iso && !mapaISO.get(t)) mapaISO.set(t, iso);
    if (iso) isoSet.add(iso);
  });
}

async function carregarCatalogoSemanas(){
  if (_catalogoCarregado) return;

  const normF = s => String(s??'').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'');
  const exactos = ['Titulo','titulo','TITULO','Title','title','TituloVideo','titulovideo'];
  const mapaISO = new Map();
  const isoSet = new Set();

  try {
    const jTrein = await _cachedFetch('treinamentos');
    const rows1 = _parseRows(jTrein);
    const chaves = Object.keys(rows1[0]||{});
    const campo = exactos.find(c => chaves.includes(c)) ?? chaves.find(k => normF(k).includes('titul')) ?? null;
    if (campo) _extrairTituloISO(rows1, campo, mapaISO, isoSet);
  } catch(e){}

  _popularSelectsCatalogo(mapaISO, isoSet);
  _catalogoCarregado = mapaISO.size > 0;

  if (_catalogoComplementado) return;
  _catalogoComplementado = true;
  try {
    const jRegs = await _cachedFetch('registros');
    const antesDoTamanho = mapaISO.size;
    _extrairTituloISO(_parseRows(jRegs), null, mapaISO, isoSet);
    if (mapaISO.size !== antesDoTamanho) _popularSelectsCatalogo(mapaISO, isoSet);
  } catch(e){}
}

carregarCatalogoSemanas();

(function initFiltroTitulo(){
  const selP = document.getElementById('fSemanaTitulo');
  const selD = document.getElementById('dashTitulo');
  [selP, selD].forEach(sel => {
    if (!sel) return;
    sel.addEventListener('focus', carregarCatalogoSemanas, { once: true });
    sel.addEventListener('mousedown', carregarCatalogoSemanas, { once: true });
  });
})();

// Lógica principal de pesquisa de registos
async function buscar(){
  const status = document.getElementById('status');
  const btnBuscar = document.getElementById('btnBuscar');
  const textoOriginalBuscar = btnBuscar ? btnBuscar.innerHTML : '';
  status.innerHTML = '<div class="text-sky-600 bg-sky-50 px-4 py-2.5 rounded-lg border border-sky-100">A pesquisar...</div>';
  if (btnBuscar) { btnBuscar.disabled = true; btnBuscar.innerHTML = '<svg class="w-4 h-4 animate-spin" fill="none" viewBox="0 0 24 24"><circle class="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" stroke-width="4"></circle><path class="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8v8z"></path></svg> A pesquisar...'; }

  const fMat = document.getElementById('fMat').value.trim();
  const fNome = document.getElementById('fNome').value.trim();
  const fTitulo = document.getElementById('fSemanaTitulo').value.trim();
  const fDataI = document.getElementById('fDataInicio').value;
  const fDataF = document.getElementById('fDataFinal').value;

  const qs = new URLSearchParams({ action: 'registros' });
  if (fMat) qs.append('matricula', fMat);
  if (fNome) qs.append('nome', fNome);
  if (fTitulo) qs.append('titulo', fTitulo);
  if (fDataI) qs.append('dataInicial', fDataI);
  if (fDataF) qs.append('dataFinal', fDataF);

  try {
    const res = await apiFetch(qs.toString());
    let data; 
    try { data = await res.json(); } catch { data = { ok: false, error: 'Resposta não-JSON do proxy' }; }
    if (!data.ok) throw new Error(data.error || 'Erro na pesquisa');
    registros = Array.isArray(data.data) ? data.data : [];

    const di = normalizarDataInput(fDataI);
    const df = fimDoDia(normalizarDataInput(fDataF));

    let lista = registros.filter(r => {
      const ts = parseTimestamp(r.Timestamp);
      if (!ts) return false;
      if (di && ts < di) return false;
      if (df && ts > df) return false;
      if (fTitulo && (r.TituloVideo ?? '') !== fTitulo) return false;
      return true;
    });

    const mapa = new Map();
    for (const r of lista){
      const chave = `${r.Matricula ?? ''}__${r.SemanaISO ?? ''}__${r.TituloVideo ?? ''}`;
      const atual = mapa.get(chave);
      const novoTS = parseTimestamp(r.Timestamp);
      if (!atual) mapa.set(chave, r);
      else { 
        const antigoTS = parseTimestamp(atual.Timestamp); 
        if (novoTS && antigoTS && novoTS < antigoTS) mapa.set(chave, r); 
      }
    }

    registrosFiltrados = Array.from(mapa.values()).sort((a,b) => String(a.Nome ?? '').localeCompare(String(b.Nome ?? ''), 'pt-PT', { sensitivity: 'base' }));
    renderTabela(registrosFiltrados);
    status.innerHTML = '<div class="text-emerald-600 font-semibold bg-emerald-50 px-4 py-2.5 rounded-lg border border-emerald-100">Pesquisa concluída: ' + registrosFiltrados.length + ' registo(s) encontrado(s).</div>';

    try {
      const temas = await buscarAssuntosPorSemana(registrosFiltrados);
      renderTemasAbordados('temasAbordados', temas);
    } catch(e) {
      renderTemasAbordados('temasAbordados', []);
    }
  } catch(err){ 
    status.innerHTML = '<div class="text-rose-600 font-semibold bg-rose-50 px-4 py-2.5 rounded-lg border border-rose-100">Falha ao processar a consulta: ' + (err.message || '') + '</div>'; 
    renderTemasAbordados('temasAbordados', []);
  } finally {
    if (btnBuscar) { btnBuscar.disabled = false; btnBuscar.innerHTML = textoOriginalBuscar; }
  }
}

// Renderização da tabela de registos no ecrã principal
function renderTabela(list){
  const tbody = document.getElementById('tbody');
  tbody.innerHTML = '';
  if (!list || list.length === 0){ 
    tbody.innerHTML = '<tr><td colspan="6" class="px-6 py-10 text-center text-slate-400">Sem registos encontrados para os filtros aplicados.</td></tr>'; 
    return; 
  }
  tbody.innerHTML = list.map(r => {
    const dataPart = formatTimestamp(r.Timestamp);
    return `
      <tr class="hover:bg-slate-50 transition-colors">
        <td class="px-6 py-4 whitespace-nowrap font-medium text-slate-900">${escapeHtml(r.Matricula ?? '')}</td>
        <td class="px-6 py-4 whitespace-nowrap">${escapeHtml(r.Nome ?? '')}</td>
        <td class="px-6 py-4 whitespace-nowrap">${escapeHtml(r.Setor ?? '')}</td>
        <td style="display:none">${escapeHtml(r.SemanaISO ?? '')}</td>
        <td class="px-6 py-4 max-w-xs truncate" title="${escapeHtml(r.TituloVideo ?? '')}">${escapeHtml(r.TituloVideo ?? '')}</td>
        <td class="px-6 py-4 whitespace-nowrap">${escapeHtml(dataPart)}</td>
      </tr>`;
  }).join('');
}

function _loadScript(src){
  return new Promise((resolve, reject) => {
    if (document.querySelector(`script[src="${src}"]`)) { resolve(); return; }
    const s = document.createElement('script');
    s.src = src;
    s.onload = resolve;
    s.onerror = () => reject(new Error('Falha ao carregar: ' + src));
    document.head.appendChild(s);
  });
}

async function _ensureXLSX(){
  if (window.XLSX) return;
  await _loadScript('https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js');
}

let _b64Loaded = false;

async function _ensureJsPDF(){
  if (!window.jspdf) {
    await _loadScript('https://cdn.jsdelivr.net/npm/jspdf@2.5.1/dist/jspdf.umd.min.js');
  }

  const testDoc = window.jspdf && new window.jspdf.jsPDF();
  if (!testDoc || typeof testDoc.autoTable !== 'function') {
    await _loadScript('https://cdn.jsdelivr.net/npm/jspdf-autotable@3.8.2/dist/jspdf.plugin.autotable.min.js');
  }

  if (!_b64Loaded) {
    await Promise.all([
      _loadScriptForce('/gestor/logo_b64.js'),
      _loadScriptForce('/gestor/assin_b64.js'),
    ]);
    _b64Loaded = !!(window._LOGO_FALLBACK_B64 && window._ASSINATURAS_IMG_B64);
  }
}

function _loadScriptForce(src){
  return new Promise((resolve, reject) => {
    const existing = document.querySelector(`script[src="${src}"]`);
    if (existing) existing.remove();
    const s = document.createElement('script');
    s.src = src;
    s.onload = resolve;
    s.onerror = () => reject(new Error('Falha ao carregar: ' + src));
    document.head.appendChild(s);
  });
}

// Exportação em Excel (SheetJS)
async function gerarXLS(){
  const base = (registrosFiltrados && registrosFiltrados.length) ? registrosFiltrados : registros;
  if (!base || !base.length) { alert('Efetue uma pesquisa primeiro.'); return; }
  try { await _ensureXLSX(); } catch(e) { alert('Erro ao carregar biblioteca Excel: ' + e.message); return; }
  const ordenada = base === registros ? [...base].sort((a,b) => String(a.Nome ?? '').localeCompare(String(b.Nome ?? ''), 'pt-PT', { sensitivity: 'base' })) : base;
  const linhas = ordenada.map(r => ({
    'Matrícula': r.Matricula ?? '',
    'Colaborador': r.Nome ?? '',
    'Setor': r.Setor ?? '',
    'Semana (ISO)': r.SemanaISO ?? '',
    'Título do Vídeo': r.TituloVideo ?? '',
    'Data de Participação': formatTimestamp(r.Timestamp),
  }));
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(linhas);
  XLSX.utils.book_append_sheet(wb, ws, 'Pesquisa');
  XLSX.writeFile(wb, nomeArquivo('xlsx'));
}

async function fetchTreinamentos(){
  try {
    const data = await _cachedFetch('treinamentos');
    if (!data || !data.ok) return [];
    return _parseRows(data);
  } catch(e){
    return [];
  }
}

async function buscarAssuntosPorSemana(lista){
  if (!Array.isArray(lista) || !lista.length) return [];
  const titulos = [...new Set(lista.map(r => r.TituloVideo).filter(Boolean))];
  if (!titulos.length) return [];

  let tList = [];
  try { tList = await fetchTreinamentos(); } catch(e) { tList = []; }

  const blocks = titulos.map(tit => {
    const regsDoGrupo = lista.filter(r => (r.TituloVideo ?? '') === tit);
    const semanaISO = String(regsDoGrupo.find(r => r.SemanaISO)?.SemanaISO ?? '');
    let hit = (tList||[]).find(t => String(t['Titulo']).trim() === tit);
    if (!hit && semanaISO) hit = (tList||[]).find(t => normalizarSemanaISO(String(t['SemanaISO'])) === normalizarSemanaISO(semanaISO));
    const assuntos = hit && hit['Assuntos'] ? String(hit['Assuntos']) : '';
    return { titulo: tit, semanaISO, assuntos };
  }).filter(b => b.assuntos);

  blocks.sort((a, b) => dashCompareSemanaISODesc(a.semanaISO, b.semanaISO));
  return blocks;
}

function renderTemasAbordados(containerId, blocks){
  const el = document.getElementById(containerId);
  if (!el) return;
  if (!blocks || !blocks.length){
    el.classList.add('hidden');
    el.innerHTML = '';
    return;
  }
  el.innerHTML = blocks.map(b => {
    const linhas = String(b.assuntos).replace(/\r\n/g,'\n').replace(/\r/g,'\n').split('\n').map(s=>s.trim()).filter(Boolean);
    return `
      <div class="mb-4 last:mb-0">
        <p class="text-xs font-bold text-brand-700 uppercase tracking-wider mb-1.5">
          Temas Abordados — Semana ${escapeHtml(b.semanaISO || '')}${b.titulo ? ' · ' + escapeHtml(b.titulo) : ''}
        </p>
        <ul class="list-disc list-inside space-y-0.5 text-sm text-slate-700">
          ${linhas.map(l => `<li>${escapeHtml(l.replace(/^-\s*/, ''))}</li>`).join('')}
        </ul>
      </div>`;
  }).join('<hr class="my-3 border-brand-100">');
  el.classList.remove('hidden');
}

const __dataUrlCache = new Map();
function isDataURLImage(v){ return (typeof v === 'string' && /^data:image\/(png|jpeg|jpg);base64,/i.test(v.trim())); }
async function urlToDataURL(url){ 
  try { 
    if(__dataUrlCache.has(url)) return __dataUrlCache.get(url); 
    const res = await fetch(url, { mode: 'cors' }); 
    const blob = await res.blob(); 
    const dataUrl = await new Promise((resolve) => { 
      const fr = new FileReader(); 
      fr.onload = () => resolve(fr.result); 
      fr.onerror = () => resolve(''); 
      fr.readAsDataURL(blob);
    }); 
    __dataUrlCache.set(url, dataUrl || ''); 
    return dataUrl || ''; 
  } catch(e){ 
    return ''; 
  } 
}

async function ensureDataURLImage(v){ 
  if(!v) return ''; 
  if(isDataURLImage(v)) return v; 
  return await urlToDataURL(v); 
}

async function resolveRowSignatures(rows){ 
  const out = []; 
  for(const r of rows){ 
    const sig = await ensureDataURLImage(r._sig || r.AssinaturaPNG || ''); 
    out.push({ ...r, _sig: sig }); 
  } 
  return out; 
}

// Geração Inteligente do Relatório de Presenças em formato PDF (A4)
async function gerarPDF(){
  const btnPDF = document.getElementById('btnPDF');
  const status = document.getElementById('status');
  const textoOriginal = btnPDF ? btnPDF.innerHTML : '';
  try {
    if (btnPDF) { btnPDF.disabled = true; btnPDF.innerHTML = '<svg class="w-4 h-4 animate-spin" fill="none" viewBox="0 0 24 24"><circle class="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" stroke-width="4"></circle><path class="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8v8z"></path></svg> A gerar...'; }
    status.innerHTML = '<div class="text-amber-600 bg-amber-50 px-4 py-2 rounded-lg border border-amber-100">A carregar bibliotecas e recursos...</div>';
    await Promise.all([_ensureJsPDF(), loadLogoAsDataURL()]);
    const baseSource = (Array.isArray(registrosFiltrados) && registrosFiltrados.length) ? registrosFiltrados : registros;
    if (!baseSource || !baseSource.length){ alert('Sem dados disponíveis para gerar PDF.'); return; }
    const base = [...baseSource].sort((a,b) => String(a.Nome ?? '').localeCompare(String(b.Nome ?? ''), 'pt-PT', { sensitivity: 'base' }));

    try {
      const matriculasBase = [...new Set(base.map(r => r.Matricula).filter(v => v !== null && v !== undefined && v !== ''))];
      if (matriculasBase.length && matriculasBase.length <= 200) {
        const qsSig = new URLSearchParams({ action: 'registros', comAssinatura: '1' });
        qsSig.set('matriculas', matriculasBase.join(','));

        const resSig = await apiFetch(qsSig.toString());
        const dataSig = await resSig.json().catch(() => null);
        if (dataSig && dataSig.ok && Array.isArray(dataSig.data)) {
          const mapaSig = new Map();
          dataSig.data.forEach(r => {
            const chave = `${r.Matricula ?? ''}__${r.SemanaISO ?? ''}__${r.TituloVideo ?? ''}`;
            if (r.AssinaturaPNG) mapaSig.set(chave, r.AssinaturaPNG);
          });
          base.forEach(r => {
            const chave = `${r.Matricula ?? ''}__${r.SemanaISO ?? ''}__${r.TituloVideo ?? ''}`;
            if (mapaSig.has(chave)) r.AssinaturaPNG = mapaSig.get(chave);
          });
        }
      }
    } catch(e) { console.warn('Não foi possível carregar assinaturas para o PDF:', e); }
    const titulos = [...new Set(base.map(r => r.TituloVideo).filter(Boolean))];
    const tituloSemana = titulos.length ? titulos.join('; ') : '-';

    const { jsPDF } = window.jspdf;
    const M_TOP = 28.35, M_BOTTOM = 28.35, M_LEFT = 14.17, M_RIGHT = 14.17;
    const doc = new jsPDF({ orientation: 'portrait', unit: 'pt', format: 'a4', compress: true });
    const pageWidth = doc.internal.pageSize.getWidth();
    const pageHeight = doc.internal.pageSize.getHeight();
    const usableWidth = pageWidth - M_LEFT - M_RIGHT;
    const LOGO_W = 120, LOGO_H = 52, L_H = 14;

    let semanaBlocks = [];
    try {
      const tList = await fetchTreinamentos();
      const gruposTitulo = titulos.length ? titulos : [null];
      semanaBlocks = gruposTitulo.map(tit => {
        const regsDoGrupo = tit ? base.filter(r => (r.TituloVideo ?? '') === tit) : base;
        const semanaISO = String(regsDoGrupo.find(r => r.SemanaISO)?.SemanaISO ?? '');
        let hit = null;
        if (tit) hit = (tList||[]).find(t => String(t['Titulo']).trim() === tit);
        if (!hit && semanaISO) hit = (tList||[]).find(t => normalizarSemanaISO(String(t['SemanaISO'])) === normalizarSemanaISO(semanaISO));
        const assuntos = hit && hit['Assuntos'] ? String(hit['Assuntos']) : '';
        return { titulo: tit || tituloSemana, semanaISO, assuntos };
      });
      semanaBlocks.sort((a, b) => dashCompareSemanaISODesc(a.semanaISO, b.semanaISO));
    } catch(e){
      semanaBlocks = [{ titulo: tituloSemana, semanaISO: '', assuntos: '' }];
    }

    semanaBlocks.forEach(b => {
      b.assuntosLines = [];
      if (b.assuntos) {
        const raw = b.assuntos.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
        const partes = raw.split('\n').map(s => s.trim()).filter(Boolean);
        b.assuntosLines = partes.flatMap(linha => doc.splitTextToSize(linha, usableWidth - 8));
      }
    });

    function drawHeader(pageNumber){
      if (LOGO_DATAURL && LOGO_DATAURL.startsWith('data:image')) {
        try {
          const fmt = LOGO_DATAURL.includes('data:image/svg') ? 'SVG'
                    : LOGO_DATAURL.includes('data:image/png') ? 'PNG' : 'JPEG';
          doc.addImage(LOGO_DATAURL, fmt, pageWidth - M_RIGHT - LOGO_W, M_TOP, LOGO_W, LOGO_H);
        } catch(e) { console.warn('Logo PDF:', e.message); }
      }
      doc.setFont('helvetica', 'bold'); doc.setFontSize(18); doc.text('DIÁLOGO SEMANAL DE SEGURANÇA', M_LEFT, M_TOP + 20);
      if (pageNumber === 1){
        let y = M_TOP + 20 + 2 * L_H;
        semanaBlocks.forEach(b => {
          doc.setFont('helvetica', 'bold'); doc.setFontSize(14);
          doc.text('Semana: ' + (b.titulo || '-'), M_LEFT, y);
          y += 2 * L_H;
          if (b.assuntosLines.length){
            doc.setFont('helvetica', 'bold'); doc.setFontSize(11);
            doc.text('Temas Abordados:', M_LEFT, y);
            doc.setFont('helvetica', 'normal'); doc.setFontSize(11);
            b.assuntosLines.forEach((linha, i) => {
              doc.text(linha, M_LEFT + 8, y + (i + 1) * L_H);
            });
            y += L_H + b.assuntosLines.length * L_H;
          }
          y += L_H;
        });
      }
    }

    function drawFooter(pageNumber, totalPagesText){
      const footerY = pageHeight - M_BOTTOM; 
      const generatedAt = new Date().toLocaleString('pt-PT');
      doc.setFont('helvetica', 'italic'); doc.setFontSize(9); doc.text('Gerado em: ' + generatedAt, M_LEFT, footerY - 4 * L_H);
      doc.setFont('helvetica', 'normal'); doc.setFontSize(9);
      doc.text('Caterair Serviços de Bordo e Hotelaria LTDA - Base GIG', M_LEFT, footerY - 3 * L_H);
      doc.text('CNPJ 33.375.601/0001-38', M_LEFT, footerY - 2 * L_H);
      doc.text('Rua P, S/N, Área de Apoio do Aeroporto Internacional do Rio de Janeiro - Ilha do Governador - RJ', M_LEFT, footerY - 1 * L_H);
      const pagText = `Página ${pageNumber} de ${totalPagesText}`.trim();
      doc.text(pagText, pageWidth / 2, footerY, { align: 'center' });
    }

    const columns = [
      { header: 'Matrícula', dataKey: 'Matricula' },
      { header: 'Colaborador', dataKey: 'Nome' },
      { header: 'Setor', dataKey: 'Setor' },
      { header: 'Data de Participação', dataKey: 'DataFmt' },
      { header: 'Assinatura', dataKey: '_sig' },
    ];

    status.innerHTML = '<div class="text-amber-600 bg-amber-50 px-4 py-2 rounded-lg border border-amber-100">A processar assinaturas dos colaboradores...</div>';
    let rows = base.map(r => ({ Matricula: r.Matricula??'', Nome: r.Nome??'', Setor: r.Setor??'', DataFmt: formatTimestamp(r.Timestamp)??'', _sig: r.AssinaturaPNG ?? '' }));
    rows = await resolveRowSignatures(rows);

    const totalPagesExp = '{total_pages_count_string}';

    let alturaBlocosSemana = 0;
    semanaBlocks.forEach(b => {
      alturaBlocosSemana += 2 * L_H;
      if (b.assuntosLines.length){
        alturaBlocosSemana += L_H + b.assuntosLines.length * L_H;
      }
      alturaBlocosSemana += L_H;
    });

    let yStartFirstPage = M_TOP + 20 + 2 * L_H + alturaBlocosSemana + L_H;
    const yStartFirstPageMin = M_TOP + 20 + 6 * L_H; 
    if (yStartFirstPage < yStartFirstPageMin) yStartFirstPage = yStartFirstPageMin;
    const yStartOtherPages = M_TOP + LOGO_H + 5 * L_H;

    doc.autoTable({
      startY: yStartFirstPage,
      margin: { left: M_LEFT, right: M_RIGHT, top: yStartOtherPages, bottom: M_BOTTOM + 80 },
      pageBreak: 'auto',
      tableWidth: usableWidth,
      columns,
      body: rows,
      styles: { font: 'helvetica', fontSize: 10, cellPadding: 4, valign: 'middle', overflow: 'linebreak', minCellHeight: 52 },
      headStyles: { fillColor: [33, 150, 243], textColor: 255, halign: 'center' },
      columnStyles: {
        Matricula: { halign: 'left', cellWidth: 60 },
        Nome:      { halign: 'left' },
        Setor:     { halign: 'left', cellWidth: 110 },
        DataFmt:   { halign: 'left', cellWidth: 90 },
        _sig:      { halign: 'center', cellWidth: 130 }
      },
      didParseCell: (data) => {
        if (data.section === 'body' && data.column.dataKey === '_sig'){
          data.cell.text = [''];
          if (data.cell && data.cell.styles){ data.cell.styles.lineWidth = 0; }
        }
      },
      didDrawPage: (data) => {
        drawHeader(data.pageNumber);
        drawFooter(data.pageNumber, totalPagesExp);
        if (data.pageNumber > 1 && data.cursor && data.cursor.y < yStartOtherPages){ data.cursor.y = yStartOtherPages; }
      },
      didDrawCell: (data) => {
        if (data.section === 'body' && data.column && data.column.dataKey === '_sig'){
          const val = data.row?.raw?._sig || '';
          if (typeof val === 'string' && /^data:image\/(png|jpeg|jpg);base64,/i.test(val)){
            try {
              const cleanVal = val.replace(/\s/g, '');
              let finalImg = cleanVal;
              try {
                const canvas = document.createElement('canvas');
                const imgEl  = new Image();
                imgEl.src    = cleanVal;
                const nW = imgEl.naturalWidth  || 400;
                const nH = imgEl.naturalHeight || 160;
                canvas.width  = nW;
                canvas.height = nH;
                const ctx = canvas.getContext('2d');
                ctx.fillStyle = '#ffffff';
                ctx.fillRect(0, 0, nW, nH);
                ctx.drawImage(imgEl, 0, 0);
                const imgData = ctx.getImageData(0, 0, nW, nH);
                const d = imgData.data;
                for (let i = 0; i < d.length; i += 4){
                  const brightness = (d[i] + d[i+1] + d[i+2]) / 3;
                  if (brightness < 180){
                    d[i]   = Math.max(0, d[i]   - 40);
                    d[i+1] = Math.max(0, d[i+1] - 40);
                    d[i+2] = Math.max(0, d[i+2] - 40);
                  } else {
                    d[i] = d[i+1] = d[i+2] = 255;
                  }
                  d[i+3] = 255;
                }
                ctx.putImageData(imgData, 0, 0);
                finalImg = canvas.toDataURL('image/png');
              } catch(canvasErr){ finalImg = cleanVal; }

              const props = doc.getImageProperties(finalImg);
              const pad  = 3;
              const maxW = data.cell.width  - pad * 2;
              const maxH = data.cell.height - pad * 2;
              const scaleW = maxW / (props.width  * 72 / 96);
              const scaleH = maxH / (props.height * 72 / 96);
              const scale  = Math.min(scaleW, scaleH);
              const w = (props.width  * 72 / 96) * scale;
              const h = (props.height * 72 / 96) * scale;
              const x = data.cell.x + (data.cell.width  - w) / 2;
              const y = data.cell.y + (data.cell.height - h) / 2;
              doc.addImage(finalImg, 'PNG', x, y, w, h);
            } catch(e){ console.warn('Sig render:', e); }
          }
        }
      },
    });

    const totalPages = doc.internal.getNumberOfPages();
    doc.setPage(totalPages);

    if (window._ASSINATURAS_IMG_B64 && window._ASSINATURAS_IMG_B64.startsWith('data:image')) {
      try {
        const yAfterTable = doc.lastAutoTable.finalY ?? (pageHeight - M_BOTTOM - 140);
        const footerAreaH = M_BOTTOM + 4 * L_H + 20;

        let highResImg = window._ASSINATURAS_IMG_B64;
        try {
          const SCALE = 2;
          const tmpImg = new Image();
          tmpImg.src   = window._ASSINATURAS_IMG_B64;
          const srcW = tmpImg.naturalWidth  || 1200;
          const srcH = tmpImg.naturalHeight || 400;
          const canvas2 = document.createElement('canvas');
          canvas2.width  = srcW * SCALE;
          canvas2.height = srcH * SCALE;
          const ctx2 = canvas2.getContext('2d');
          ctx2.imageSmoothingEnabled  = true;
          ctx2.imageSmoothingQuality  = 'high';
          ctx2.fillStyle = '#ffffff';
          ctx2.fillRect(0, 0, canvas2.width, canvas2.height);
          ctx2.drawImage(tmpImg, 0, 0, canvas2.width, canvas2.height);
          highResImg = canvas2.toDataURL('image/png');
        } catch(e) { console.warn('Canvas upscale falhou, usando original:', e); }

        const props = doc.getImageProperties(window._ASSINATURAS_IMG_B64);
        const ratio = props.height / props.width;
        const imgW  = usableWidth;
        const imgH  = imgW * ratio;

        let yImg = yAfterTable + 16;
        if (yImg + imgH > pageHeight - footerAreaH) {
          doc.addPage();
          const newPageNum = doc.internal.getNumberOfPages();
          doc.setPage(newPageNum);
          drawHeader(newPageNum);
          drawFooter(newPageNum, totalPagesExp);
          yImg = yStartOtherPages + 10;
        }

        doc.addImage(highResImg, 'PNG', M_LEFT, yImg, imgW, imgH);
      } catch(e) {
        console.warn('Erro ao renderizar imagem de assinaturas:', e);
      }
    }

    if (typeof doc.putTotalPages === 'function'){ doc.putTotalPages(totalPagesExp); }
    doc.save('Relatorio_Dialogo_Semanal_de_Seguranca.pdf');
    status.innerHTML = '<div class="text-emerald-600 bg-emerald-50 px-4 py-2 rounded-lg border border-emerald-100">PDF gerado com sucesso!</div>';
  } catch(err) {
    console.error('Erro ao gerar PDF:', err);
    status.innerHTML = '<div class="text-rose-600 bg-rose-50 px-4 py-2 rounded-lg border border-rose-100">Erro ao gerar PDF: ' + (err.message || 'Erro desconhecido') + '</div>';
  } finally {
    if (btnPDF) { btnPDF.disabled = false; btnPDF.innerHTML = textoOriginal; }
  }
}

// Prompt administrativo para cadastrar novos funcionários (pt-PT)
async function novoFuncionarioViaPrompt(){
  const status = document.getElementById('status');
  try {
    let mRaw = prompt('Nova matrícula (5 algarismos, apenas números):');
    if (mRaw === null) return; 
    mRaw = (mRaw ?? '').replace(/\D/g, '');
    while (!/^\d{5}$/.test(mRaw)){
      mRaw = prompt('Número de registo inválido. Forneça exatamente 5 dígitos (ex.: 12345):', mRaw);
      if (mRaw === null) return; 
      mRaw = (mRaw ?? '').replace(/\D/g, '');
    }
    const matricula = mRaw;
    let nome = prompt('Nome completo:'); 
    if (nome === null) return; 
    nome = (nome ?? '').trim();
    if (!nome){ alert('O nome do colaborador é obrigatório.'); return; }

    let setor = prompt('Setor (opcional):'); 
    if (setor === null) setor = ''; 
    setor = (setor ?? '').trim();

    const ok = confirm(`Confirmar cadastro do novo colaborador?\n\nMatrícula: ${matricula}\nNome: ${nome}\nSetor: ${setor ?? '-'}\nAtivo: SIM`);
    if (!ok){ status.innerHTML = '<div class="text-amber-600 bg-amber-50 px-4 py-2 rounded-lg border border-amber-100">Operação cancelada.</div>'; return; }

    const payload = { matricula, nome, setor, ativo: true };
    status.innerHTML = '<div class="text-amber-500 bg-amber-50 px-4 py-2 rounded-lg border border-amber-100">A submeter novo registo ao servidor...</div>';

    const res = await fetch(API_BASE + '?action=addFuncionario', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(payload) });
    let data; 
    try { data = await res.json(); } catch { data = { ok: false, error: 'Erro de formatação na resposta' }; }
    if (!data.ok) throw new Error(data.error ?? 'Falha ao incluir colaborador');
    status.innerHTML = '<div class="text-emerald-600 bg-emerald-50 px-4 py-2 rounded-lg border border-emerald-100">Colaborador inserido com sucesso!</div>';
  } catch(err){ 
    status.innerHTML = '<div class="text-rose-600 bg-rose-50 px-4 py-2 rounded-lg border border-rose-100">' + (err && err.message ? err.message : 'Erro ao processar criação de colaborador') + '</div>'; 
  }
}

// Prompt administrativo para excluir/deletar funcionário do banco de dados definitivamente
async function excluirFuncionarioViaPrompt(){
  const status = document.getElementById('status');
  try {
    let mRaw = prompt('Digite a matrícula (5 dígitos) do colaborador que deseja EXCLUIR definitivamente do banco de dados:');
    if (mRaw === null) return;
    mRaw = mRaw.replace(/\D/g, '');
    while (!/^\d{5}$/.test(mRaw)){
      mRaw = prompt('Matrícula inválida. Digite exatamente 5 números (ex.: 12345):', mRaw);
      if (mRaw === null) return;
      mRaw = mRaw.replace(/\D/g, '');
    }
    const matricula = mRaw;

    await dashEnsureData();
    const colab = (DASH_funcionarios || []).find(f => afastNormMat(f.Matricula) === matricula);

    let confirmMsg = `Deseja realmente apagar em definitivo o colaborador com a matrícula ${matricula}?`;
    if (colab && colab.Nome) {
      confirmMsg = `Deseja realmente EXCLUIR em definitivo o colaborador abaixo do banco de dados?\n\nNome: ${colab.Nome}\nMatrícula: ${matricula}\nSetor: ${colab.Setor ?? '-'}\n\nNota: Esta ação irá apagar o registo do colaborador do Sheets.`;
    }

    const ok = confirm(confirmMsg);
    if (!ok){ status.innerHTML = '<div class="text-amber-600 bg-amber-50 px-4 py-2 rounded-lg border border-amber-100">Exclusão cancelada.</div>'; return; }

    status.innerHTML = '<div class="text-rose-500 bg-rose-50 px-4 py-2 rounded-lg border border-rose-100">A remover registo do Google Sheets...</div>';

    let success = false;
    let errorMsg = '';

    try {
      const res = await fetch(API_BASE + '?action=excluirFuncionario', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ matricula })
      });
      const data = await res.json();
      if (data && data.ok) {
        success = true;
      } else {
        errorMsg = data.error || '';
      }
    } catch(e) {
      errorMsg = e.message;
    }

    if (!success) {
      try {
        const res = await fetch(API_BASE + '?action=deleteFuncionario', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ matricula })
        });
        const data = await res.json();
        if (data && data.ok) {
          success = true;
        } else {
          errorMsg = data.error || errorMsg;
        }
      } catch(e) {
        errorMsg = e.message || errorMsg;
      }
    }

    if (!success) {
      try {
        const res = await fetch(API_BASE + '?action=excluirColaborador', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ matricula })
        });
        const data = await res.json();
        if (data && data.ok) {
          success = true;
        } else {
          errorMsg = data.error || errorMsg;
        }
      } catch(e) {
        errorMsg = e.message || errorMsg;
      }
    }

    if (!success) {
      throw new Error(errorMsg || 'Erro desconhecido na remoção do Sheets');
    }

    status.innerHTML = '<div class="text-emerald-600 bg-emerald-50 px-4 py-2 rounded-lg border border-emerald-100">Colaborador removido em definitivo do Sheets!</div>';

    delete _apiCache['funcionarios'];
    DASH_funcionarios = null;
  } catch(err){
    status.innerHTML = '<div class="text-rose-600 bg-rose-50 px-4 py-2 rounded-lg border border-rose-100">' + (err && err.message ? err.message : 'Erro ao excluir colaborador') + '</div>';
  }
}

function escapeHtml(str){
  return String(str ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/\x22/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function nomeArquivo(ext){ 
  const hoje = new Date(); 
  const pad = n => String(n).padStart(2, '0'); 
  return `DSS_GIG_Relatorio_${hoje.getFullYear()}-${pad(hoje.getMonth() + 1)}-${pad(hoje.getDate())}.${ext}`; 
}

function normalizarDataInput(v){
  if (!v) return null;
  v = String(v).trim();
  const m = v.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) {
    const [, dd, mm, yyyy] = m;
    return new Date(Date.UTC(+yyyy, +mm - 1, +dd, 0, 0, 0));
  }
  const m2 = v.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m2) {
    const [, yyyy, mm, dd] = m2;
    return new Date(Date.UTC(+yyyy, +mm - 1, +dd, 0, 0, 0));
  }
  return null;
}

function fimDoDia(dIniUTC){ 
  if(!dIniUTC) return null; 
  return new Date(dIniUTC.getTime() + 24 * 60 * 60 * 1000 - 1); 
}

function parseTimestamp(valor){ 
  if(!valor) return null; 
  let d = new Date(valor); 
  if(!Number.isNaN(d.getTime())) return d; 
  const rx = /(^\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/; 
  const m = String(valor).trim().match(rx); 
  if(m){ 
    const dd = +m[1], mm = +m[2], yy = +m[3]; 
    const hh = +(m[4]??0), mi = +(m[5]??0), ss = +(m[6]??0); 
    return new Date(yy, mm-1, dd, hh, mi, ss); 
  } 
  return null; 
}

function formatTimestamp(ts){ 
  const d = parseTimestamp(ts); 
  if(!d) return String(ts ?? ''); 
  const pad = n => String(n).padStart(2, '0'); 
  return `${pad(d.getDate())}/${pad(d.getMonth() + 1)}/${d.getFullYear()} ${pad(d.getHours())}:${pad(d.getMinutes())}`; 
}

/* ====== LÓGICA DO PAINEL DASHBOARD ====== */
let DASH_registrosAll = null, DASH_funcionarios = null, DASH_treinamentos = null;
let DASH_participantes = [], DASH_naoParticipantes = [];

/* ====== GESTÃO DE AUSÊNCIAS / FÉRIAS ====== */
let AFAST_set = new Set(); 

function afastSalvar(){ 
  try { localStorage.setItem('dss_afast', JSON.stringify([...AFAST_set])); } catch(e){} 
}

function dashRender(naoParticipantes, participantes, dispensados){
  const tbodyNP = document.getElementById('tbodyNP');
  const tbodyP  = document.getElementById('tbodyP');
  const tbodyD  = document.getElementById('tbodyDispensados');

  if (tbodyNP) {
    if (!naoParticipantes || !naoParticipantes.length)
      tbodyNP.innerHTML = '<tr><td colspan="3" class="px-4 py-8 text-center text-emerald-600 font-semibold bg-emerald-50">✅ Todos os colaboradores elegíveis participaram!</td></tr>';
    else
      tbodyNP.innerHTML = naoParticipantes.map(n => `
        <tr class="hover:bg-rose-50/40 transition-colors">
          <td class="px-4 py-3 font-semibold text-slate-900">${escapeHtml(n.Matricula??'')}</td>
          <td class="px-4 py-3">${escapeHtml(n.Nome??'')}</td>
          <td class="px-4 py-3">${escapeHtml(n.Setor??'')}</td>
        </tr>`).join('');
  }

  if (tbodyP) {
    if (!participantes || !participantes.length)
      tbodyP.innerHTML = '<tr><td colspan="4" class="px-4 py-8 text-center text-slate-400">Nenhum registo de participação processado.</td></tr>';
    else
      tbodyP.innerHTML = participantes.map(p => `
        <tr class="hover:bg-emerald-50/40 transition-colors">
          <td class="px-4 py-3 font-semibold text-slate-900">${escapeHtml(p.Matricula??'')}</td>
          <td class="px-4 py-3">${escapeHtml(p.Nome??'')}</td>
          <td class="px-4 py-3">${escapeHtml(p.Setor??'')}</td>
          <td class="px-4 py-3">${escapeHtml(formatTimestamp(p.Timestamp))}</td>
        </tr>`).join('');
  }

  if (tbodyD) {
    if (!dispensados || !dispensados.length)
      tbodyD.innerHTML = '<tr><td colspan="5" class="px-4 py-8 text-center text-slate-400">Nenhum colaborador dispensado nesta semana.</td></tr>';
    else {
      const fmtPeriodo = s => {
        if (!s) return '';
        return s.replace(/(\d{4})-(\d{2})-(\d{2})T[^\s]*/g, (_, y, m, d) => `${d}/${m}/${y}`)
                .replace(/(\d{4})-(\d{2})-(\d{2})/g, (_, y, m, d) => `${d}/${m}/${y}`);
      };
      tbodyD.innerHTML = dispensados.map(d => `
        <tr class="hover:bg-sky-50/40 transition-colors">
          <td class="px-4 py-3 font-semibold text-slate-900">${escapeHtml(d.Matricula??'')}</td>
          <td class="px-4 py-3">${escapeHtml(d.Nome??'')}</td>
          <td class="px-4 py-3">${escapeHtml(d.Setor??'')}</td>
          <td class="px-4 py-3 font-medium text-sky-700">${escapeHtml(d.Motivo??'')}</td>
          <td class="px-4 py-3 text-slate-500 text-xs">${escapeHtml(fmtPeriodo(d.Periodo??''))}</td>
        </tr>`).join('');
    }
  }
}

function dashKPIs({ total, part, nPart, nDisp, semana, titulo, registrosAll, funcAtivos }){
  const nDispSafe = nDisp ?? 0;
  const totalAll  = total + nDispSafe;
  const pct    = total > 0 ? Math.round((part   / total) * 100) : 0;
  const pctNP  = total > 0 ? Math.round((nPart  / total) * 100) : 0;
  const pctD   = totalAll > 0 ? Math.round((nDispSafe / totalAll) * 100) : 0;

  document.getElementById('kpiTotalAtivos').textContent     = String(total ?? '-');
  document.getElementById('kpiParticiparam').textContent    = String(part  ?? '-');
  document.getElementById('kpiNaoParticiparam').textContent = String(nPart ?? '-');
  document.getElementById('kpiParticiparamPct').textContent = total ? `${pct}%` : '-';
  document.getElementById('kpiPctNum').textContent          = total ? `${pct}%` : '-';
  document.getElementById('kpiSemanaSel').textContent       = String(semana ?? '-');
  document.getElementById('kpiTituloSel').textContent       = String(titulo ?? '-');

  const kpiDisp = document.getElementById('kpiDispensados');
  if (kpiDisp) kpiDisp.textContent = String(nDispSafe);

  document.getElementById('barPart').style.width    = total ? `${pct}%`   : '0%';
  document.getElementById('barNaoPart').style.width = total ? `${pctNP}%` : '0%';
  const barDisp = document.getElementById('barDisp');
  if (barDisp) {
    barDisp.style.width = totalAll ? `${pctD}%` : '0%';
    barDisp.style.backgroundColor = '#38bdf8';
  }

  const canvas = document.getElementById('kpiDonut');
  if (canvas) {
    const ctx = canvas.getContext('2d');
    const cx = canvas.width / 2, cy = canvas.height / 2, r = 54, thick = 16;
    ctx.clearRect(0, 0, canvas.width, canvas.height);

    const arc = (s, e, color) => {
      ctx.beginPath();
      ctx.arc(cx, cy, r, (s / 100) * 2 * Math.PI - Math.PI / 2, (e / 100) * 2 * Math.PI - Math.PI / 2);
      ctx.strokeStyle = color;
      ctx.lineWidth = thick;
      ctx.lineCap = 'round';
      ctx.stroke();
    };

    ctx.beginPath();
    ctx.arc(cx, cy, r, 0, 2 * Math.PI);
    ctx.strokeStyle = '#f3f4f6';
    ctx.lineWidth = thick;
    ctx.stroke();

    if (totalAll > 0) {
      const pctPartAll = Math.round((part       / totalAll) * 100);
      const pctNPAll   = Math.round((nPart      / totalAll) * 100);
      const pctDAll    = Math.round((nDispSafe  / totalAll) * 100);
      let cursor = 0;
      if (pctNPAll  > 0){ arc(cursor, cursor + pctNPAll,  '#ef4444'); cursor += pctNPAll;  }
      if (pctPartAll > 0){ arc(cursor, cursor + pctPartAll, '#22c55e'); cursor += pctPartAll; }
      if (pctDAll   > 0){ arc(cursor, cursor + pctDAll,   '#38bdf8'); }
    }
  }

  const nuncaEl   = document.getElementById('kpiNuncaCount');
  const nuncaList = document.getElementById('kpiNuncaList');
  const regs = Array.isArray(registrosAll) ? registrosAll : [];
  const ativos = Array.isArray(funcAtivos) ? funcAtivos : [];

  const normMat = v => String(v??'').replace(/\D/g, '').padStart(5, '0');
  const jaParticiparam = new Set(regs.map(r => normMat(r.Matricula)));
  const nunca = ativos.filter(f =>
    !jaParticiparam.has(normMat(f.Matricula)) && !AFAST_set.has(afastNormMat(f.Matricula))
  ).sort((a,b) => String(a.Nome??'').localeCompare(String(b.Nome??''), 'pt-PT', { sensitivity: 'base' }));

  if (nuncaEl) nuncaEl.textContent = nunca.length || '0';
  if (nuncaList){
    if (!nunca.length){
      nuncaList.innerHTML = '<span class="text-xs font-semibold text-emerald-600 bg-emerald-50 px-2 py-1 rounded">✅ Todos os colaboradores ativos já participaram!</span>';
    } else {
      nuncaList.innerHTML = nunca.slice(0, 5).map(f =>
        `<div class="flex items-center gap-2 text-xs">
          <span class="w-4 h-4 rounded-full bg-rose-100 text-rose-600 flex items-center justify-center font-bold">!</span>
          <span class="truncate flex-grow text-slate-700 font-medium">${escapeHtml(f.Nome??f.Matricula??'')}</span>
          <span class="text-slate-400 font-mono">${normMat(f.Matricula)}</span>
        </div>`
      ).join('') + (nunca.length > 5 ? `<div class="text-[10px] text-slate-400 mt-1">+ ${nunca.length - 5} colaboradores na lista</div>` : '');
    }
  }

  const assiduosEl = document.getElementById('kpiAssiduosList');
  if (assiduosEl){
    const regsPorMatSemana = new Map();
    for (const r of regs){
      const mat = normMat(r.Matricula);
      const iso = String(r.SemanaISO??'');
      const ts  = parseTimestamp(r.Timestamp);
      if (!mat || !iso || !ts) continue;
      const chave = mat + '__' + iso;
      const atual = regsPorMatSemana.get(chave);
      if (!atual || ts < atual.ts) regsPorMatSemana.set(chave, { mat, iso, ts, nome: r.Nome??'' });
    }

    const contagem = new Map();
    for (const { mat, iso, ts, nome } of regsPorMatSemana.values()){
      const m = iso.match(/^(\d{4})-W\d{1,2}$/i);
      let segsDesdeInicio = 0;
      if (m){
        const ano = +m[1], sem = +m[2];
        const jan4 = new Date(ano, 0, 4);
        const segunda = new Date(jan4);
        segunda.setDate(jan4.getDate() - ((jan4.getDay() + 6) % 7) + (sem - 1) * 7);
        segsDesdeInicio = Math.max(0, (ts - segunda) / 1000);
      }
      const entry = contagem.get(mat) || { nome, totalSegs: 0, semanas: 0 };
      entry.totalSegs += segsDesdeInicio;
      entry.semanas   += 1;
      entry.nome = entry.nome || nome;
      contagem.set(mat, entry);
    }

    const ranking = [...contagem.entries()]
      .filter(([,v]) => v.semanas > 0)
      .sort((a,b) => {
        if (b[1].semanas !== a[1].semanas) return b[1].semanas - a[1].semanas;
        return (a[1].totalSegs / a[1].semanas) - (b[1].totalSegs / b[1].semanas);
      })
      .slice(0, 5);

    const rankBg = ['bg-amber-100 text-amber-800', 'bg-slate-200 text-slate-800', 'bg-orange-100 text-orange-800', 'bg-slate-100 text-slate-600', 'bg-slate-100 text-slate-600'];
    assiduosEl.innerHTML = ranking.length
      ? ranking.map(([mat, v], i) => {
          return `<div class="flex items-center gap-2 text-xs">
            <span class="w-4 h-4 rounded-full font-bold flex items-center justify-center text-[10px] ${rankBg[i]}">${i + 1}</span>
            <span class="truncate flex-grow text-slate-700 font-medium">${escapeHtml(v.nome || mat)}</span>
            <span class="text-brand-600 font-semibold bg-brand-50 px-1.5 py-0.5 rounded text-[10px]" title="${v.semanas} sessões concluídas">${v.semanas}✓</span>
          </div>`;
        }).join('')
      : '<span class="text-xs text-slate-400">Dados insuficientes para gerar classificação</span>';
  }
}

function afastCarregar(){ 
  try { 
    const d = localStorage.getItem('dss_afast'); 
    if(d){ JSON.parse(d).forEach(m => AFAST_set.add(m)); } 
  } catch(e){} 
}

function afastNormMat(v){ return String(v??'').replace(/\D/g, '').padStart(5, '0'); }

function mapFeriasRow(row) {
  const keys = Object.keys(row);
  const norm = s => s.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'');
  const findVal = (possibleKeys) => {
    const foundKey = keys.find(k => possibleKeys.includes(norm(k)));
    return foundKey ? row[foundKey] : '';
  };
  return {
    Matricula:    String(findVal(['matricula','mat','numregisto']) || '').trim(),
    Funcionario:  String(findVal(['funcionario','nome','colaborador']) || '').trim(),
    Situacao:     String(findVal(['situacao','status','tipo','motivo']) || 'Férias').trim(),
    InicioFerias: String(findVal(['inicioferias','inicio','datainicio','datadeinicio','data_inicio','inicio_ferias']) || '').trim(),
    FimFerias:    String(findVal(['fimferias','fim','datafim','datadefim','data_fim','fim_ferias']) || '').trim()
  };
}

function funcEstaAfastadoNaSemana(matricula, semanaNorm, segMs, domMs){
  const mat = afastNormMat(matricula);
  const registros = feriasServidor.filter(f => afastNormMat(f.Matricula) === mat);
  if (!registros.length) return false;

  const parseMs = s => {
    if (!s) return null;
    s = String(s).trim();
    const isoFull = s.match(/^(\d{4})-(\d{2})-(\d{2})T/);
    if (isoFull) return Date.UTC(+isoFull[1], +isoFull[2]-1, +isoFull[3]);
    const br = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (br) return Date.UTC(+br[3], +br[2]-1, +br[1]);
    const iso = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (iso) return Date.UTC(+iso[1], +iso[2]-1, +iso[3]);
    return null;
  };

  let wSeg = segMs, wDom = domMs;
  if (!wSeg || !wDom) {
    const monday = semanaISOToMonday(semanaNorm);
    if (monday) {
      wSeg = monday.getTime();
      wDom = monday.getTime() + 6 * 86400000;
    }
  }

  for (const f of registros){
    const dIniMs = parseMs(f.InicioFerias);
    const dFimMs = parseMs(f.FimFerias);

    if (!dIniMs && !dFimMs) {
      return true;
    }

    const iniOk = !dIniMs || (wDom !== undefined && dIniMs <= wDom);
    const fimOk = !dFimMs || (wSeg !== undefined && dFimMs >= wSeg);
    if (iniOk && fimOk) return true;
  }
  return false;
}

function afastRenderTags(){
  const badge  = document.getElementById('afastBadge');
  const tbody  = document.getElementById('tbodyAfastados');
  if (!tbody) return;

  const hoje = new Date();
  const hojeMs = Date.UTC(hoje.getUTCFullYear(), hoje.getUTCMonth(), hoje.getUTCDate());

  const parseMs = s => {
    if (!s) return null; s = String(s).trim();
    const isoFull = s.match(/^(\d{4})-(\d{2})-(\d{2})T/);
    if (isoFull) return Date.UTC(+isoFull[1],+isoFull[2]-1,+isoFull[3]);
    const br = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (br) return Date.UTC(+br[3],+br[2]-1,+br[1]);
    const iso = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (iso) return Date.UTC(+iso[1],+iso[2]-1,+iso[3]);
    return null;
  };

  const comStatus = feriasServidor.map((f, idx) => {
    const dIniMs = parseMs(f.InicioFerias);
    const dFimMs = parseMs(f.FimFerias);
    let ativo = false;
    if (!dIniMs && !dFimMs) ativo = true;
    else if (dIniMs && !dFimMs) ativo = hojeMs >= dIniMs;
    else if (!dIniMs && dFimMs) ativo = hojeMs <= dFimMs;
    else ativo = hojeMs >= dIniMs && hojeMs <= dFimMs;
    return { ...f, _idx: idx, _ativo: ativo };
  });

  const totalAtivos = comStatus.filter(f => f._ativo).length;
  if (badge) badge.textContent = `${totalAtivos} ativos / ${comStatus.length} total`;

  if (comStatus.length === 0){
    tbody.innerHTML = `<tr><td colspan="6" class="px-6 py-10 text-center text-slate-400 italic">
      Nenhum registo na aba <strong>Ferias</strong> do Google Sheets.<br>
      <span class="text-xs">Use o formulário acima para inserir o primeiro registo.</span>
    </td></tr>`;
    return;
  }

  const formatDate = s => {
    if (!s) return '';
    s = String(s).trim();
    const isoFull = s.match(/^(\d{4})-(\d{2})-(\d{2})T/);
    if (isoFull) return `${isoFull[3]}/${isoFull[2]}/${isoFull[1]}`;
    const isoSimple = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (isoSimple) return `${isoSimple[3]}/${isoSimple[2]}/${isoSimple[1]}`;
    if (/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(s)) return s;
    return s;
  };

  const sorted = [...comStatus].sort((a, b) => {
    if (a._ativo !== b._ativo) return a._ativo ? -1 : 1;
    const aMs = parseMs(a.InicioFerias) || 0;
    const bMs = parseMs(b.InicioFerias) || 0;
    return bMs - aMs;
  });

  tbody.innerHTML = sorted.map((f) => {
    const mat      = afastNormMat(f.Matricula);
    const nome     = escapeHtml(f.Funcionario || '-');
    const sit      = escapeHtml(f.Situacao || 'Férias');
    const sitNorm  = sit.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'');
    const isDispMed = sitNorm.includes('dispensa') || sitNorm.includes('medica') || sitNorm.includes('inss') || sitNorm.includes('afastado');
    const badgeClass = isDispMed
      ? 'bg-rose-100 text-rose-800 border-rose-200'
      : 'bg-amber-100 text-amber-800 border-amber-200';

    const dIni   = formatDate(f.InicioFerias);
    const dFim   = formatDate(f.FimFerias);
    const periodo = (dIni && dFim)
      ? `${dIni} a ${dFim}`
      : (dIni || dFim || '<em class="text-slate-400">Sem prazo definido</em>');

    const statusBadge = f._ativo
      ? '<span class="px-2 py-0.5 text-xs font-bold rounded-full bg-emerald-100 text-emerald-800 border border-emerald-200">Ativo</span>'
      : '<span class="px-2 py-0.5 text-xs font-bold rounded-full bg-slate-100 text-slate-500 border border-slate-200">Encerrado</span>';

    const deleteBtn = f._ativo
      ? `<button type="button" data-idx="${f._idx}" class="btnAfastExcluir inline-flex items-center gap-1 px-2.5 py-1.5 text-xs font-semibold text-rose-600 hover:text-white hover:bg-rose-500 border border-rose-200 hover:border-rose-500 rounded-lg transition-colors">
           <svg class="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M6 18L18 6M6 6l12 12"/></svg>
           Excluir
         </button>`
      : `<button type="button" data-idx="${f._idx}" class="btnAfastExcluir inline-flex items-center gap-1 px-2.5 py-1.5 text-xs font-semibold text-slate-400 hover:text-white hover:bg-slate-400 border border-slate-200 hover:border-slate-400 rounded-lg transition-colors" title="⚠️ Não exclua registros históricos — eles garantem que o colaborador apareça como Dispensado nas semanas passadas">
           <svg class="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M6 18L18 6M6 6l12 12"/></svg>
           ⚠️ Histórico
         </button>`;

    return `<tr class="hover:bg-slate-50 transition-colors ${f._ativo ? '' : 'opacity-70'}">
      <td class="px-4 py-3 font-mono font-semibold text-slate-800">${mat}</td>
      <td class="px-4 py-3 font-medium text-slate-800">${nome}</td>
      <td class="px-4 py-3">
        <span class="px-2.5 py-1 text-xs font-semibold rounded-full border ${badgeClass}">${sit}</span>
      </td>
      <td class="px-4 py-3 text-sm text-slate-600">${periodo}</td>
      <td class="px-4 py-3">${statusBadge}</td>
      <td class="px-4 py-3 text-center">${deleteBtn}</td>
    </tr>`;
  }).join('');

  tbody.querySelectorAll('.btnAfastExcluir').forEach(btn => {
    btn.addEventListener('click', function(){
      const idx = parseInt(this.dataset.idx);
      const reg = feriasServidor[idx];
      const isEncerrado = !sorted.find(s => s._idx === idx)?._ativo;
      if (isEncerrado) {
        if (!confirm(`⚠️ ATENÇÃO: Este registro já está encerrado.\n\nExcluir registros históricos fará o colaborador "${reg?.Funcionario || ''}" aparecer incorretamente como Ausente/Falta em relatórios passados.\n\nDeseja realmente excluir do Sheets?`)) {
          return;
        }
      }
      if (reg && reg.Matricula) {
        afastRemover(reg.Matricula, reg.Situacao);
      }
    });
  });
}

async function afastRemover(matricula, situacao){
  try {
    const res = await fetch(API_BASE + '?action=deleteFerias', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ matricula, situacao })
    });
    const data = await res.json();
    if (!data.ok) throw new Error(data.error || 'Erro ao remover ausência');
    
    feriasServidor = feriasServidor.filter(f => 
      !(afastNormMat(f.Matricula) === afastNormMat(matricula) && 
        (!situacao || String(f.Situacao).toLowerCase() === String(situacao).toLowerCase()))
    );
    afastRenderTags();
  } catch (err) {
    alert('Erro ao excluir registro de ausência: ' + err.message);
  }
}
