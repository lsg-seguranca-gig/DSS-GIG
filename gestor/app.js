// A API fica na raiz do projeto (/api/gas), independentemente da pasta
// onde o HTML está hospedado (ex.: /gestor/). Por isso usamos caminho absoluto.
const API_BASE = "/api/gas";

  // Controlo de cache manual e chamadas de API
  function apiFetch(params, options){
    const url = API_BASE + '?' + params + '&_t=' + Date.now();
    return fetch(url, Object.assign({ cache: 'no-store', headers: { 'Cache-Control': 'no-cache' } }, options || {}));
  }

  const LOGO_SRC = "logo lsg.png";
  let LOGO_DATAURL = "";
  let registros = [];
  let registrosFiltrados = [];
  let feriasServidor = []; // Guarda as informações da planilha "Ferias" do Google Sheets

  /* ====================================================
   * LSG SKY CHEFS — LOGO EMBUTIDO (SVG) COMO FALLBACK GARANTIDO
   * Usado no PDF se lsg sky chefs logo.png não estiver acessível
   * ==================================================== */
  // window._LOGO_FALLBACK_B64 é resolvido no momento do uso (após lazy load via logo_b64.js)

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

// Imagem de assinaturas dos instrutores embutida diretamente no código
// window._ASSINATURAS_IMG_B64 é resolvido no momento do uso (após lazy load via assin_b64.js)

// Pré-carregamento do logo assim que a página abre (para o PDF não precisar esperar)
let _logoLoadPromise = null;
function loadLogoAsDataURL(){
  if (_logoLoadPromise) return _logoLoadPromise;
  // Logo real já embutido como window._LOGO_FALLBACK_B64 — resolve imediatamente (sem esperar rede)
  LOGO_DATAURL = window._LOGO_FALLBACK_B64;
  _logoLoadPromise = Promise.resolve(LOGO_DATAURL);
  // Tentativa silenciosa de usar o ficheiro externo se disponível no servidor
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

  // Aplicar máscara dd/mm/aaaa dinamicamente ao digitar
  txt.addEventListener('input', function(){
    let v = this.value.replace(/\D/g, '');
    if (v.length > 8) v = v.slice(0, 8);
    if      (v.length >= 5) this.value = v.slice(0, 2) + '/' + v.slice(2, 4) + '/' + v.slice(4);
    else if (v.length >= 3) this.value = v.slice(0, 2) + '/' + v.slice(2);
    else                    this.value = v;
  });

  // Abrir o datepicker nativo invisível ao clicar no ícone do calendário
  btn.addEventListener('click', function(){
    const m = txt.value.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
    if (m) native.value = `${m[3]}-${m[2]}-${m[1]}`;
    native.showPicker ? native.showPicker() : native.click();
  });

  // Sincronizar o valor selecionado no nativo de volta para o input formatado
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

// NOTA: nenhuma chamada à API é feita automaticamente ao abrir a página.
// Os dados (treinamentos/registos/funcionários/férias) só são carregados
// quando o usuário interage com os filtros e clica em "Pesquisar"/"Atualizar".

// Descompactação das estruturas JSON retornadas pela API (GAS)
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

// Auxiliar para a geração das opções (Dropdown) dos títulos e semanas
function _buildTituloOpts(mapaISO, prefixOpt){
  return prefixOpt + [...mapaISO.entries()]
    .sort((a,b)=>{ 
      const pa = dashParseSemanaISO(a[1]), pb = dashParseSemanaISO(b[1]); 
      if(pa.year !== pb.year) return pb.year - pa.year; 
      return pb.week - pa.week; 
    })
    .map(([t]) => `<option value="${escapeHtml(t)}">${escapeHtml(t)}</option>`).join('');
}

// Carrega o catálogo de Semanas/Títulos (treinamentos + registos) e popula
// os dropdowns "Título do Vídeo" (Principal), "Tema / Vídeo" (Dashboard) e
// o select oculto de Semana ISO do Dashboard. Esta é a ÚNICA carga
// automática feita ao abrir a página — os dados pesados (resultados da
// pesquisa, KPIs do Dashboard, férias) continuam sendo carregados apenas
// quando o usuário clicar em "Pesquisar"/"Atualizar". Usamos tanto
// "treinamentos" (catálogo oficial) quanto "registos" (que pode conter
// semanas mais antigas que já não estão mais no catálogo), para que os
// dropdowns mostrem TODAS as semanas disponíveis na base de dados.
let _catalogoCarregado = false;
async function carregarCatalogoSemanas(){
  if (_catalogoCarregado) return;

  const normF = s => String(s??'').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'');
  const exactos = ['Titulo','titulo','TITULO','Title','title','TituloVideo','titulovideo'];
  const mapaISO = new Map();
  const isoSet = new Set();

  try {
    const [jTrein, jRegs] = await Promise.all([
      _cachedFetch('treinamentos'),
      _cachedFetch('registros'),
    ]);

    const rows1 = _parseRows(jTrein);
    const chaves = Object.keys(rows1[0]||{});
    const campo = exactos.find(c => chaves.includes(c)) ?? chaves.find(k => normF(k).includes('titul')) ?? null;
    if (campo) rows1.forEach(r => {
      const t = String(r[campo]??'').trim();
      const iso = String(r.SemanaISO??'').trim();
      if(t) mapaISO.set(t, iso);
      if(iso) isoSet.add(iso);
    });

    const rows2 = _parseRows(jRegs);
    rows2.forEach(r => {
      const t = String(r.TituloVideo??r.Titulo??r.titulo??'').trim();
      const iso = String(r.SemanaISO??'').trim();
      if(t && !mapaISO.has(t)) mapaISO.set(t, iso);
      else if(t && !mapaISO.get(t) && iso) mapaISO.set(t, iso);
      if(iso) isoSet.add(iso);
    });
  } catch(e){}

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

    _catalogoCarregado = true;
  }

  if (isoSet.size) {
    const isoList = [...isoSet].sort(dashCompareSemanaISODesc);
    const selIso = document.getElementById('dashSemanaIso');
    if (selIso) selIso.innerHTML = '<option value="">Selecione...</option>' + isoList.map(sem => `<option value="${escapeHtml(sem)}">${escapeHtml(sem)}</option>`).join('');
  }
}

// Carrega o catálogo imediatamente (fire-and-forget)
carregarCatalogoSemanas();

// Caso a primeira tentativa falhe (ex.: instabilidade de rede no momento do
// load), tenta novamente na primeira interação do usuário com os campos.
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

  try {
    const res = await apiFetch(qs.toString());
    let data; 
    try { data = await res.json(); } catch { data = { ok: false, error: 'Resposta não-JSON do proxy' }; }
    if (!data.ok) throw new Error(data.error || 'Erro na pesquisa');
    registros = Array.isArray(data.data) ? data.data : [];

    const di = normalizarDataInput(fDataI);
    const df = fimDoDia(normalizarDataInput(fDataF));

    // Filtro de data: compara diretamente o Timestamp de cada registo com o
    // intervalo informado. Isso garante que "Data Inicial = Data Final = 15/06/2026"
    // retorne todos os registos cujo Timestamp caia nesse dia — independentemente
    // da semana ISO a que pertençam.
    // O filtro por "Título do Vídeo" continua funcionando por semana (valor do select).

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
  } catch(err){ 
    status.innerHTML = '<div class="text-rose-600 font-semibold bg-rose-50 px-4 py-2.5 rounded-lg border border-rose-100">Falha ao processar a consulta: ' + (err.message || '') + '</div>'; 
  } finally {
    if (btnBuscar) { btnBuscar.disabled = false; btnBuscar.innerHTML = textoOriginalBuscar; }
  }
}

// Renderização da tabela de registos no ecrã principal
function renderTabela(list){
  const tbody = document.getElementById('tbody');
  tbody.innerHTML = '';
  if (!list || list.length === 0){ 
    tbody.innerHTML = '<tr><td colspan="7" class="px-6 py-10 text-center text-slate-400">Sem registos encontrados para os filtros aplicados.</td></tr>'; 
    return; 
  }
  tbody.innerHTML = list.map(r => {
    const dataPart = formatTimestamp(r.Timestamp);
    const assinatura = r.AssinaturaPNG ? '<img class="sig-img" src="' + r.AssinaturaPNG + '" alt="Assinatura" />' : '-';
    return `
      <tr class="hover:bg-slate-50 transition-colors">
        <td class="px-6 py-4 whitespace-nowrap font-medium text-slate-900">${escapeHtml(r.Matricula ?? '')}</td>
        <td class="px-6 py-4 whitespace-nowrap">${escapeHtml(r.Nome ?? '')}</td>
        <td class="px-6 py-4 whitespace-nowrap">${escapeHtml(r.Setor ?? '')}</td>
        <td style="display:none">${escapeHtml(r.SemanaISO ?? '')}</td>
        <td class="px-6 py-4 max-w-xs truncate" title="${escapeHtml(r.TituloVideo ?? '')}">${escapeHtml(r.TituloVideo ?? '')}</td>
        <td class="px-6 py-4 whitespace-nowrap">${escapeHtml(dataPart)}</td>
        <td class="px-6 py-4 whitespace-nowrap">${assinatura}</td>
      </tr>`;
  }).join('');
}

// ─── Lazy loaders: as bibliotecas de exportação só são baixadas quando ───
// ─── o usuário realmente clicar em "Exportar Excel" ou "Gerar PDF".    ───
// ─── Isso elimina ~1,4 MB de scripts bloqueantes do carregamento inicial. ─

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

async function _ensureJsPDF(){
  if (window.jspdf) return;
  // Carrega as imagens base64 (logo + assinaturas) e o jsPDF em paralelo
  await Promise.all([
    _loadScript('logo_b64.js'),
    _loadScript('assin_b64.js'),
    _loadScript('https://cdn.jsdelivr.net/npm/jspdf@2.5.1/dist/jspdf.umd.min.js'),
  ]);
  // autotable depende do jsPDF — carrega em sequência
  await _loadScript('https://cdn.jsdelivr.net/npm/jspdf-autotable@3.8.2/dist/jspdf.plugin.autotable.min.js');
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

// Auxiliares de Assinatura Digital e Parsing de URLs de Base64
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
    const titulos = [...new Set(base.map(r => r.TituloVideo).filter(Boolean))];
    const tituloSemana = titulos.length ? titulos.join('; ') : '-';

    const { jsPDF } = window.jspdf;
    const M_TOP = 28.35, M_BOTTOM = 28.35, M_LEFT = 14.17, M_RIGHT = 14.17;
    const doc = new jsPDF({ orientation: 'portrait', unit: 'pt', format: 'a4', compress: true });
    const pageWidth = doc.internal.pageSize.getWidth();
    const pageHeight = doc.internal.pageSize.getHeight();
    const usableWidth = pageWidth - M_LEFT - M_RIGHT;
    const LOGO_W = 120, LOGO_H = 52, L_H = 14;

    // Montar um bloco "Semana / Temas Abordados" para cada Título de Vídeo
    // presente nos registos selecionados. Se mais de um vídeo (mais de uma
    // semana) tiver sido selecionado, cada um aparece em seu próprio bloco,
    // com os respectivos assuntos.
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
      // Ordenar da semana mais recente para a mais antiga
      semanaBlocks.sort((a, b) => dashCompareSemanaISODesc(a.semanaISO, b.semanaISO));
    } catch(e){
      semanaBlocks = [{ titulo: tituloSemana, semanaISO: '', assuntos: '' }];
    }

    // Pré-calcular as linhas de cada bloco de assuntos (já com a quebra automática)
    semanaBlocks.forEach(b => {
      b.assuntosLines = [];
      if (b.assuntos) {
        const raw = b.assuntos.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
        const partes = raw.split('\n').map(s => s.trim()).filter(Boolean);
        b.assuntosLines = partes.flatMap(linha => doc.splitTextToSize(linha, usableWidth - 8));
      }
    });

    function drawHeader(pageNumber){
      // Logo embutido — sempre disponível, formato JPEG
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
          y += L_H; // espaçamento entre blocos de semana
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

    // Calcular a altura total ocupada pelos blocos "Semana / Temas Abordados"
    // (deve usar exatamente o mesmo cálculo de incrementos usado em drawHeader)
    let alturaBlocosSemana = 0;
    semanaBlocks.forEach(b => {
      alturaBlocosSemana += 2 * L_H; // linha "Semana: X"
      if (b.assuntosLines.length){
        alturaBlocosSemana += L_H + b.assuntosLines.length * L_H; // "Temas Abordados:" + linhas
      }
      alturaBlocosSemana += L_H; // espaçamento entre blocos
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
      styles: { font: 'helvetica', fontSize: 10, cellPadding: 4, valign: 'middle', overflow: 'linebreak', minCellHeight: 36 },
      headStyles: { fillColor: [33, 150, 243], textColor: 255, halign: 'center' },
      columnStyles: { Matricula: { halign: 'left' }, Nome: { halign: 'left' }, Setor: { halign: 'left', cellWidth: 120 }, DataFmt: { halign: 'left' }, _sig: { halign: 'center', cellWidth: 150 } },
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
              const props = doc.getImageProperties(cleanVal); 
              const wPtNat=(props.width*72)/96; 
              const hPtNat=(props.height*72)/96; 
              const pad=4; 
              const maxW=data.cell.width-pad*2; 
              const maxH=data.cell.height-pad*2; 
              const scale=Math.min(maxW/wPtNat, maxH/hPtNat, 1); 
              const w=wPtNat*scale; 
              const h=hPtNat*scale; 
              const x=data.cell.x + (data.cell.width - w)/2; 
              const y=data.cell.y + (data.cell.height - h)/2; 
              doc.addImage(cleanVal, props.fileType || 'PNG', x, y, w, h); 
            } catch(e){}
          }
        }
      },
    });

    // Garante que estamos na última página absoluta para desenhar os instrutores
    const totalPages = doc.internal.getNumberOfPages();
    doc.setPage(totalPages);

    // Bloco de Instrutores — usa a imagem assinaturas.png embutida diretamente
    if (window._ASSINATURAS_IMG_B64 && window._ASSINATURAS_IMG_B64.startsWith('data:image')) {
      try {
        const yAfterTable = doc.lastAutoTable.finalY ?? (pageHeight - M_BOTTOM - 140);
        const footerAreaH = M_BOTTOM + 4 * L_H + 20;

        // Obter dimensões naturais da imagem para calcular proporção correta
        const props = doc.getImageProperties(window._ASSINATURAS_IMG_B64);
        const natW = props.width, natH = props.height;
        const ratio = natH / natW;

        // A imagem ocupa toda a largura útil da página
        const imgW = usableWidth;
        const imgH = imgW * ratio;

        let yImg = yAfterTable + 16;
        // Se não couber na página atual, cria nova página
        if (yImg + imgH > pageHeight - footerAreaH) {
          doc.addPage();
          const newPageNum = doc.internal.getNumberOfPages();
          doc.setPage(newPageNum);
          drawHeader(newPageNum);
          drawFooter(newPageNum, totalPagesExp);
          yImg = yStartOtherPages + 10;
        }

        doc.addImage(window._ASSINATURAS_IMG_B64, 'JPEG', M_LEFT, yImg, imgW, imgH);
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

    // Buscar localmente os dados do colaborador para confirmar exclusão
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

    // Prova de exclusão primária: action=excluirFuncionario
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

    // Fallback 1: action=deleteFuncionario
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

    // Fallback 2: action=excluirColaborador
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

    // Limpar cache local para forçar recarregamento na próxima busca
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

// Funções de Gestão de cache das ausências re-definidas para evitar erros de referência (ReferenceError)
function afastSalvar(){ 
  try { localStorage.setItem('dss_afast', JSON.stringify([...AFAST_set])); } catch(e){} 
}

// Gestão de renderização das listas de presenças e faltas no ecrã Dashboard
function dashRender(naoParticipantes, participantes){
  const tbodyNP = document.getElementById('tbodyNP');
  const tbodyP = document.getElementById('tbodyP');
  if (!naoParticipantes || !naoParticipantes.length){ 
    tbodyNP.innerHTML = '<tr><td colspan="3" class="px-4 py-8 text-center text-emerald-600 font-semibold bg-emerald-50">Todos os colaboradores participaram! 🎉</td></tr>'; 
  } else { 
    tbodyNP.innerHTML = naoParticipantes.map(n => `<tr><td class="px-4 py-3 font-semibold text-slate-900">${escapeHtml(n.Matricula ?? '')}</td><td class="px-4 py-3">${escapeHtml(n.Nome ?? '')}</td><td class="px-4 py-3">${escapeHtml(n.Setor ?? '')}</td></tr>`).join(''); 
  }
  if (!participantes || !participantes.length){ 
    tbodyP.innerHTML = '<tr><td colspan="4" class="px-4 py-8 text-center text-slate-400">Nenhum registo de participação processado.</td></tr>'; 
  } else { 
    tbodyP.innerHTML = participantes.map(p => `<tr><td class="px-4 py-3 font-semibold text-slate-900">${escapeHtml(p.Matricula ?? '')}</td><td class="px-4 py-3">${escapeHtml(p.Nome ?? '')}</td><td class="px-4 py-3">${escapeHtml(p.Setor ?? '')}</td><td class="px-4 py-3">${escapeHtml(formatTimestamp(p.Timestamp))}</td></tr>`).join(''); 
  }
}

// Renderização e cálculo de dados estatísticos (KPIs) semanais
function dashKPIs({ total, part, nPart, semana, titulo, registrosAll, funcAtivos }){
  const pct   = total > 0 ? Math.round((part  / total) * 100) : 0;
  const pctNP = total > 0 ? Math.round((nPart / total) * 100) : 0;

  document.getElementById('kpiTotalAtivos').textContent     = String(total ?? '-');
  document.getElementById('kpiParticiparam').textContent    = String(part  ?? '-');
  document.getElementById('kpiNaoParticiparam').textContent = String(nPart ?? '-');
  document.getElementById('kpiParticiparamPct').textContent = total ? `${pct}%` : '-';
  document.getElementById('kpiPctNum').textContent          = total ? `${pct}%` : '-';
  document.getElementById('kpiSemanaSel').textContent       = String(semana ?? '-');
  document.getElementById('kpiTituloSel').textContent       = String(titulo ?? '-');

  document.getElementById('barPart').style.width    = total ? `${pct}%`   : '0%';
  document.getElementById('barNaoPart').style.width = total ? `${pctNP}%` : '0%';

  // Renderização dinâmica do gráfico de rosca nativo usando HTML5 Canvas
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

    if (total > 0){
      if (pctNP > 0) arc(0, pctNP, '#ef4444');
      if (pct   > 0) arc(pctNP, pctNP + pct, '#22c55e');
    }
  }

  // KPIs Secundários: Colaboradores sem nenhuma adesão (Nunca Participaram)
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

  // Algoritmo de Assiduidade e Agilidade (Top 5 mais frequentes)
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

/* ----------------------------------------------------------------
 * Verifica se um funcionário está afastado considerando a semana ISO
 * - Situação "Afastado INSS" → sempre afastado (sem período)
 * - Situação "Férias" → afastado se a semana ISO cai dentro do período
 * ---------------------------------------------------------------- */
function funcEstaAfastadoNaSemana(matricula, semanaNorm){
  const mat = afastNormMat(matricula);
  const registros = feriasServidor.filter(f => afastNormMat(f.Matricula) === mat);
  if (!registros.length) return false;

  for (const f of registros){
    const sitNorm = String(f.Situacao || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').trim();
    // Afastado INSS: sem período → sempre exclui
    if (sitNorm.includes('inss') || sitNorm.includes('afastado')){
      return true;
    }
    // Férias: verificar se a semana cai no período
    if (sitNorm.includes('ferias') || sitNorm.includes('férias')){
      if (!f.InicioFerias || !f.FimFerias) return true; // sem período → exclui sempre
      // Converter datas BR (dd/mm/yyyy) ou ISO (yyyy-mm-dd) para Date
      const parseDate = s => {
        if (!s) return null;
        s = String(s).trim();
        // ISO completo com timezone: 2026-05-04T07:00:00.000Z
        const isoFull = s.match(/^(\d{4})-(\d{2})-(\d{2})T/);
        if (isoFull) return new Date(`${isoFull[1]}-${isoFull[2]}-${isoFull[3]}`);
        // dd/mm/yyyy
        const br = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
        if (br) return new Date(`${br[3]}-${br[2].padStart(2,'0')}-${br[1].padStart(2,'0')}`);
        // yyyy-mm-dd
        const iso = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
        if (iso) return new Date(s);
        return null;
      };
      const dIni = parseDate(f.InicioFerias);
      const dFim = parseDate(f.FimFerias);
      if (!dIni || !dFim) return true;

      // Converter semana ISO para segunda e domingo
      const isoToMonday = (isoWeek) => {
        const m = isoWeek.match(/^(\d{4})-W(\d{2})$/);
        if (!m) return null;
        const year = parseInt(m[1]), week = parseInt(m[2]);
        const jan4 = new Date(year, 0, 4);
        const dayOfWeek = jan4.getDay() || 7;
        const monday = new Date(jan4);
        monday.setDate(jan4.getDate() - dayOfWeek + 1 + (week - 1) * 7);
        return monday;
      };
      const monday = isoToMonday(semanaNorm || '');
      if (!monday) return true;
      const sunday = new Date(monday); sunday.setDate(monday.getDate() + 6);
      // Sobreposição: férias começa antes de domingo E termina depois de segunda
      if (dIni <= sunday && dFim >= monday) return true;
    }
  }
  return false;
}

/* ----------------------------------------------------------------
 * Renderiza a tabela de férias/ausências com botão de exclusão
 * ---------------------------------------------------------------- */
function afastRenderTags(){
  const badge  = document.getElementById('afastBadge');
  const tbody  = document.getElementById('tbodyAfastados');
  if (!tbody) return;

  if (badge) badge.textContent = feriasServidor.length;

  if (feriasServidor.length === 0){
    tbody.innerHTML = `<tr><td colspan="5" class="px-6 py-10 text-center text-slate-400 italic">
      Nenhum registo na aba <strong>Ferias</strong> do Google Sheets.<br>
      <span class="text-xs">Use o formulário acima para inserir o primeiro registo.</span>
    </td></tr>`;
    return;
  }

  tbody.innerHTML = feriasServidor.map((f, idx) => {
    const mat      = afastNormMat(f.Matricula);
    const nome     = escapeHtml(f.Funcionario || '-');
    const sit      = escapeHtml(f.Situacao || 'Férias');
    const sitNorm  = sit.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'');
    const isINSS   = sitNorm.includes('inss') || sitNorm.includes('afastado');
    const badgeClass = isINSS
      ? 'bg-rose-100 text-rose-800 border-rose-200'
      : 'bg-amber-100 text-amber-800 border-amber-200';

    const formatDate = s => {
      if (!s) return '';
      s = String(s).trim();
      // ISO completo com timezone: 2026-05-04T07:00:00.000Z → usa apenas a parte da data
      const isoFull = s.match(/^(\d{4})-(\d{2})-(\d{2})T/);
      if (isoFull) return `${isoFull[3]}/${isoFull[2]}/${isoFull[1]}`;
      // yyyy-mm-dd simples → dd/mm/yyyy
      const isoSimple = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
      if (isoSimple) return `${isoSimple[3]}/${isoSimple[2]}/${isoSimple[1]}`;
      // já está em dd/mm/yyyy — mantém
      if (/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(s)) return s;
      return s;
    };
    const dIni  = formatDate(f.InicioFerias);
    const dFim  = formatDate(f.FimFerias);
    const periodo = isINSS ? '<em class="text-slate-400">Sem prazo definido</em>'
                           : (dIni && dFim ? `${dIni} a ${dFim}` : (dIni || dFim || '<em class="text-slate-400">—</em>'));

    return `<tr class="hover:bg-slate-50 transition-colors">
      <td class="px-4 py-3 font-mono font-semibold text-slate-800">${mat}</td>
      <td class="px-4 py-3 font-medium text-slate-800">${nome}</td>
      <td class="px-4 py-3">
        <span class="px-2.5 py-1 text-xs font-semibold rounded-full border ${badgeClass}">${sit}</span>
      </td>
      <td class="px-4 py-3 text-sm text-slate-600">${periodo}</td>
      <td class="px-4 py-3 text-center">
        <button type="button" data-idx="${idx}" class="btnAfastExcluir inline-flex items-center gap-1 px-2.5 py-1.5 text-xs font-semibold text-rose-600 hover:text-white hover:bg-rose-500 border border-rose-200 hover:border-rose-500 rounded-lg transition-colors">
          <svg class="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M6 18L18 6M6 6l12 12"/></svg>
          Excluir
        </button>
      </td>
    </tr>`;
  }).join('');

  // Delegação de eventos para botões de exclusão
  tbody.querySelectorAll('.btnAfastExcluir').forEach(btn => {
    btn.addEventListener('click', function(){
      const idx = parseInt(this.dataset.idx);
      afastExcluir(idx);
    });
  });
}

/* ----------------------------------------------------------------
 * Inserir novo registo de férias/ausência no Google Sheets
 * ---------------------------------------------------------------- */
async function afastInserir(){
  const msgEl  = document.getElementById('afastStatusMsg');
  const btnIns = document.getElementById('btnAfastInserir');
  const mat    = document.getElementById('afastInputMat').value.replace(/\D/g,'').padStart(5,'0');
  const nome   = document.getElementById('afastInputNome').value.trim();
  const sit    = document.getElementById('afastInputSit').value;
  const ini    = document.getElementById('afastInputIni').value;  // yyyy-mm-dd ou ''
  const fim    = document.getElementById('afastInputFim').value;

  if (!mat || mat === '00000'){ msgEl.innerHTML = '<span class="text-rose-600">Matrícula inválida.</span>'; return; }
  if (!nome)                  { msgEl.innerHTML = '<span class="text-rose-600">Nome é obrigatório.</span>'; return; }
  const isFerias = sit.toLowerCase().includes('f') && !sit.toLowerCase().includes('inss');
  if (isFerias && (!ini || !fim)){ msgEl.innerHTML = '<span class="text-rose-600">Férias exigem início e fim.</span>'; return; }

  // Converter datas para formato dd/mm/yyyy (padrão do Sheets)
  const fmtDate = s => {
    if (!s) return '';
    const m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    return m ? `${m[3]}/${m[2]}/${m[1]}` : s;
  };

  const payload = {
    matricula:    mat,
    funcionario:  nome,
    situacao:     sit,
    inicioFerias: isFerias ? fmtDate(ini) : '',
    fimFerias:    isFerias ? fmtDate(fim) : ''
  };

  btnIns.disabled = true;
  msgEl.innerHTML = '<span class="text-amber-600">A inserir...</span>';
  try {
    const res  = await fetch(API_BASE + '?action=addFerias', {
      method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    });
    const data = await res.json().catch(() => ({ ok: false, error: 'Resposta inválida' }));
    if (!data.ok) throw new Error(data.error || 'Falha no servidor');

    // Atualizar lista local imediatamente
    feriasServidor.push({
      Matricula: mat, Funcionario: nome, Situacao: sit,
      InicioFerias: isFerias ? fmtDate(ini) : '',
      FimFerias:    isFerias ? fmtDate(fim) : ''
    });
    AFAST_set.add(mat);
    afastRenderTags();
    msgEl.innerHTML = '<span class="text-emerald-600 font-semibold">✓ Inserido com sucesso!</span>';
    // Limpar formulário
    document.getElementById('afastInputMat').value  = '';
    document.getElementById('afastInputNome').value = '';
    document.getElementById('afastInputSit').value  = 'Férias';
    document.getElementById('afastInputIni').value  = '';
    document.getElementById('afastInputFim').value  = '';
    setTimeout(() => { msgEl.innerHTML = ''; }, 3000);
  } catch(e){
    msgEl.innerHTML = `<span class="text-rose-600">Erro: ${e.message}</span>`;
  } finally {
    btnIns.disabled = false;
  }
}

/* ----------------------------------------------------------------
 * Excluir registo de férias/ausência do Google Sheets
 * ---------------------------------------------------------------- */
async function afastExcluir(idx){
  const f = feriasServidor[idx];
  if (!f) return;
  if (!confirm(`Excluir o registo de "${f.Funcionario || f.Matricula}" (${f.Situacao})?\n\nEste colaborador voltará a ser considerado ativo no dashboard.`)) return;

  const mat   = afastNormMat(f.Matricula);
  const tbody = document.getElementById('tbodyAfastados');
  const msgEl = document.getElementById('afastStatusMsg');
  if (msgEl) msgEl.innerHTML = '<span class="text-amber-600">A excluir...</span>';

  try {
    // Tenta action=deleteFerias e action=excluirFerias como fallback
    let success = false;
    for (const action of ['deleteFerias','excluirFerias','removeFerias']){
      try {
        const res  = await fetch(API_BASE + '?action=' + action, {
          method: 'POST', headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ matricula: mat, situacao: f.Situacao })
        });
        const data = await res.json().catch(() => null);
        if (data && data.ok){ success = true; break; }
      } catch(e){ /* tenta próximo */ }
    }
    if (!success) throw new Error('O servidor não reconheceu a ação de exclusão.');

    // Atualizar lista local
    feriasServidor.splice(idx, 1);
    // Reconstruir AFAST_set
    AFAST_set.clear();
    feriasServidor.forEach(r => { const m = afastNormMat(r.Matricula); if (m !== '00000') AFAST_set.add(m); });
    afastRenderTags();
    if (msgEl) { msgEl.innerHTML = '<span class="text-emerald-600 font-semibold">✓ Excluído!</span>'; setTimeout(()=>{ msgEl.innerHTML=''; }, 3000); }
  } catch(e){
    if (msgEl) msgEl.innerHTML = `<span class="text-rose-600">Erro: ${e.message}</span>`;
  }
}

// Inicialização da área de Ausências
(function initAfast(){
  afastCarregar(); // restaura AFAST_set do localStorage (compatibilidade)
  // afastRenderTags() será chamado após dashEnsureData() no dashInit

  const panel   = document.getElementById('afastPanel');
  const chevron = document.getElementById('afastChevron');
  if (panel && chevron){
    panel.addEventListener('toggle', () => { chevron.textContent = panel.open ? '▲' : '▼'; });
  }

  // Botão inserir
  document.getElementById('btnAfastInserir')?.addEventListener('click', afastInserir);

  // Mostrar/ocultar campos de período conforme situação
  const selSit  = document.getElementById('afastInputSit');
  const wIni    = document.getElementById('afastInputIniWrap');
  const wFim    = document.getElementById('afastInputFimWrap');
  const togglePeriodo = () => {
    const isFerias = selSit.value.toLowerCase().includes('f') && !selSit.value.toLowerCase().includes('inss');
    [wIni, wFim].forEach(el => { if(el) el.style.display = isFerias ? '' : 'none'; });
  };
  selSit?.addEventListener('change', togglePeriodo);
  togglePeriodo(); // estado inicial
})();

document.getElementById('btnDashAtualizar').addEventListener('click', dashAtualizar);
document.getElementById('btnDashXLSNP').addEventListener('click', dashGerarXLS_NP);
document.getElementById('btnDashXLSP').addEventListener('click', dashGerarXLS_P);

// Diagnóstico da API — testa todos os endpoints e mostra a resposta bruta
(function initDiag(){
  const btnDiag = document.getElementById('btnDiag');
  const btnDiagClear = document.getElementById('btnDiagClear');
  const diagBox = document.getElementById('diagBox');
  if (!btnDiag || !diagBox) return;

  btnDiag.addEventListener('click', async function(){
    diagBox.classList.remove('hidden');
    btnDiagClear.classList.remove('hidden');
    btnDiag.classList.add('hidden');
    diagBox.textContent = 'A testar endpoints da API...\n';

    const actions = ['treinamentos','registros','funcionarios','ferias','Ferias','ausencias','Ausencias','ferias_ativas'];
    for (const action of actions) {
      try {
        const res = await apiFetch('action=' + action);
        const txt = await res.text();
        let parsed;
        try { parsed = JSON.parse(txt); } catch { parsed = null; }
        const ok = parsed ? (parsed.ok ? '✅ OK' : '❌ ok=false') : '⚠️  não-JSON';
        const rows = parsed ? (_parseRows(parsed).length + ' linhas') : '';
        diagBox.textContent += `[${action}] ${ok} ${rows}\n`;
        if (parsed && !parsed.ok && parsed.error) diagBox.textContent += `  erro: ${parsed.error}\n`;
      } catch(e) {
        diagBox.textContent += `[${action}] ❌ FALHA: ${e.message}\n`;
      }
    }
    diagBox.textContent += '\nDiagnóstico concluído.';
  });

  btnDiagClear.addEventListener('click', function(){
    diagBox.classList.add('hidden');
    btnDiagClear.classList.add('hidden');
    btnDiag.classList.remove('hidden');
    diagBox.textContent = '';
  });
})();

document.getElementById('dashTitulo').addEventListener('change', function(){
  const titVal = this.value.trim();
  const hit = (DASH_treinamentos||[]).find(t => String(t.Titulo).trim() === titVal);
  document.getElementById('dashSemanaIso').value = (hit && hit.SemanaISO) ? String(hit.SemanaISO) : '';
});

// Inicialização "leve" do Dashboard: NÃO faz nenhuma chamada à API.
// Os dados só são carregados quando o usuário clicar em "Atualizar".
(function dashInitLight(){
  const dashStatus = document.getElementById('dashStatus');
  if (dashStatus) {
    dashStatus.innerHTML = '<div class="text-sky-600 bg-sky-50 px-4 py-2 rounded-lg border border-sky-100">Clique em <strong>Atualizar</strong> para carregar os dados do Dashboard.</div>';
  }
  const tbodyAfastados = document.getElementById('tbodyAfastados');
  if (tbodyAfastados) {
    tbodyAfastados.innerHTML = '<tr><td colspan="5" class="px-6 py-8 text-center text-slate-400">Clique em "Atualizar" para carregar os registos do Google Sheets.</td></tr>';
  }
})();

async function dashEnsureData(){
  // Carregar tudo em paralelo (treinamentos, funcionarios, registros)
  const [jTrein, jFunc, jRegs] = await Promise.all([
    !DASH_treinamentos ? _cachedFetch('treinamentos').catch(() => null) : Promise.resolve(null),
    !DASH_funcionarios ? _cachedFetch('funcionarios').catch(() => null) : Promise.resolve(null),
    !DASH_registrosAll ? _cachedFetch('registros').catch(() => null)    : Promise.resolve(null),
  ]);
  if (!DASH_treinamentos) DASH_treinamentos = _parseRows(jTrein);
  if (!DASH_funcionarios) DASH_funcionarios = _parseRows(jFunc);
  if (!DASH_registrosAll) DASH_registrosAll = _parseRows(jRegs);

  // Carregar férias — tenta todos os nomes possíveis da aba no GAS.
  // As tentativas são feitas em PARALELO (em vez de sequenciais) para que
  // o carregamento seja o mais rápido possível.
  const feriasCandidates = ['ferias','Ferias','férias','Férias','ausencias','Ausencias','ausências','Ausências','vacation','leaves'];
  let feriasOk = false;
  const feriasResultados = await Promise.all(feriasCandidates.map(async action => {
    try {
      const res = await apiFetch('action=' + action);
      const data = await res.json().catch(() => null);
      return { action, data };
    } catch(e) {
      return { action, data: null };
    }
  }));
  for (const { action, data } of feriasResultados) {
    if (data && data.ok) {
      const rawRows = _parseRows(data);
      if (rawRows.length > 0) {
        feriasServidor = rawRows.map(mapFeriasRow).filter(f => f.Matricula && f.Matricula !== '00000');
        feriasOk = true;
        console.info('[Férias] Carregado via action=' + action + ', ' + feriasServidor.length + ' registo(s)');
        break;
      }
    }
  }

  if (!feriasOk) {
    // Fallback final: derivar dos funcionários com campo Situacao
    const norm = s => String(s||'').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').trim();
    const situacoesFerias = ['ferias','afastado','licenca','inss','afastamento','licenca medica','licenca maternidade'];
    const derived = (DASH_funcionarios || []).filter(f =>
      situacoesFerias.some(s => norm(f.Situacao || f.situacao || '').includes(s))
    ).map(f => ({
      Matricula: f.Matricula,
      Funcionario: f.Nome,
      Situacao: f.Situacao || f.situacao || 'Férias',
      InicioFerias: f.InicioFerias || f.inicioFerias || '',
      FimFerias: f.FimFerias || f.fimFerias || ''
    }));
    if (derived.length) {
      feriasServidor = derived;
      console.info('[Férias] Carregado via fallback (campo Situacao), ' + derived.length + ' registo(s)');
    } else {
      console.warn('[Férias] Nenhum dado encontrado. Verifique o nome da aba no Google Sheets (use Diagnóstico Bruto da API).');
    }
  }

  // Sincronizar conjunto local AFAST_set para o filtro do dashboard
  AFAST_set.clear();
  feriasServidor.forEach(f => {
    const m = afastNormMat(f.Matricula);
    if (m && m !== '00000') AFAST_set.add(m);
  });
}

async function dashPopularSelects(){
  const normF = s => String(s??'').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'');
  const exactos = ['Titulo','titulo','TITULO','Title','title','TituloVideo','titulovideo'];
  const mapaISO = new Map();

  const rows1 = DASH_treinamentos || [];
  const chaves = Object.keys(rows1[0]||{});
  const campo = exactos.find(c => chaves.includes(c)) ?? chaves.find(k => normF(k).includes('titul')) ?? null;
  if (campo) rows1.forEach(r => { const t = String(r[campo]??'').trim(); if(t) mapaISO.set(t, String(r.SemanaISO??'').trim()); });

  const isoSet = new Set(rows1.map(r => String(r.SemanaISO||'')).filter(Boolean));

  const titSel = document.getElementById('dashTitulo');
  if (mapaISO.size){
    titSel.innerHTML = '<option value="">Selecione um título...</option>' + _buildTituloOpts(mapaISO, '');
  }

  const rows2 = DASH_registrosAll || [];
  rows2.forEach(r => { 
    const t = String(r.TituloVideo??r.Titulo??r.titulo??'').trim(); 
    const iso = String(r.SemanaISO??'').trim(); 
    if(t && !mapaISO.has(t)) mapaISO.set(t, iso); 
    else if(t && !mapaISO.get(t) && iso) mapaISO.set(t, iso); 
    if(iso) isoSet.add(iso); 
  });

  const selAtual = titSel.value;
  titSel.innerHTML = '<option value="">Selecione um título...</option>' + _buildTituloOpts(mapaISO, '');
  if (selAtual && [...titSel.options].some(o => o.value === selAtual)) titSel.value = selAtual;

  const isoList = [...isoSet].sort(dashCompareSemanaISODesc);
  document.getElementById('dashSemanaIso').innerHTML =
    '<option value="">Selecione...</option>' + isoList.map(sem => `<option value="${escapeHtml(sem)}">${escapeHtml(sem)}</option>`).join('');
}

function dashSugerirSemanaAtivaMaisRecente(){
  const ativos = (DASH_treinamentos||[]).filter(t => { 
    const v = String(t.Ativo??t.ativo??'').toLowerCase().trim(); 
    return v === 'true' || v === '1' || v === 'sim' || v === 'yes'; 
  });
  const lista = ativos.length ? ativos : (DASH_treinamentos||[]);
  let tituloSelecionado = null;
  if (lista && lista.length){
    const hasISO = lista.filter(t => t.SemanaISO && t.Titulo);
    if (hasISO.length){
      const maisRecente = hasISO.sort((a,b) => dashCompareSemanaISODesc(String(a.SemanaISO), String(b.SemanaISO)))[0];
      tituloSelecionado = String(maisRecente.Titulo);
    }
  }
  const titSel = document.getElementById('dashTitulo');
  if (tituloSelecionado){
    const opt = [...titSel.options].find(o => o.value === tituloSelecionado);
    if (opt) titSel.value = opt.value;
  } else if (titSel.options.length > 1){
    titSel.selectedIndex = 1;
  }

  const titVal = titSel.value;
  if (titVal){
    const hit = (DASH_treinamentos||[]).find(t => String(t.Titulo).trim() === titVal);
    if (hit && hit.SemanaISO) document.getElementById('dashSemanaIso').value = String(hit.SemanaISO);
  }
}

// Atualização de dados do dashboard com recarregamento limpo das fontes de dados
async function dashAtualizar(){
  const dashStatus = document.getElementById('dashStatus');
  const btnAtualizar = document.getElementById('btnDashAtualizar');
  const textoOriginalBtn = btnAtualizar ? btnAtualizar.innerHTML : '';

  // Capturar o título selecionado ANTES de limpar o cache
  const titSelecionadoAntes = document.getElementById('dashTitulo').value.trim();

  if (btnAtualizar) {
    btnAtualizar.disabled = true;
    btnAtualizar.innerHTML = '<svg class="w-4 h-4 animate-spin" fill="none" viewBox="0 0 24 24"><circle class="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" stroke-width="4"></circle><path class="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8v8z"></path></svg> A carregar...';
  }
  if (dashStatus) dashStatus.innerHTML = '<div class="text-sky-600 bg-sky-50 px-4 py-2 rounded-lg border border-sky-100">A carregar dados atualizados...</div>';

  try {
    // Limpar cache para forçar recarregamento
    delete _apiCache['treinamentos'];
    delete _apiCache['registros'];
    delete _apiCache['funcionarios'];
    DASH_treinamentos = null;
    DASH_funcionarios = null;
    DASH_registrosAll = null;

    // Carregar todos os dados em paralelo
    await dashEnsureData();
    afastRenderTags(); // Atualizar tabela de férias após carregar

    await dashPopularSelects();

    // Tentar restaurar a seleção anterior; se não existir, sugerir a mais recente
    const sel = document.getElementById('dashTitulo');
    if (titSelecionadoAntes && [...sel.options].some(o => o.value === titSelecionadoAntes)) {
      sel.value = titSelecionadoAntes;
    } else {
      dashSugerirSemanaAtivaMaisRecente();
    }

    // Verificar se há um título selecionado para processar
    const titFinal = sel.value.trim();
    if (!titFinal) {
      if (dashStatus) dashStatus.innerHTML = '<div class="text-amber-600 bg-amber-50 px-4 py-2 rounded-lg border border-amber-100">Por favor, selecione um Tema para prosseguir.</div>';
      dashRender([], []);
      dashKPIs({ total: 0, part: 0, nPart: 0, semana: '-', titulo: '-', registrosAll: DASH_registrosAll||[], funcAtivos: (DASH_funcionarios||[]).filter(f => ativoValido(f.Ativo??f.ativo)) });
      return;
    }

    // Processar os dados com o título selecionado
    await dashProcessar(titFinal, dashStatus);
  } catch(e) {
    console.error('Erro ao atualizar dashboard:', e);
    if (dashStatus) dashStatus.innerHTML = '<div class="text-rose-600 bg-rose-50 px-4 py-2 rounded-lg border border-rose-100">Erro ao carregar dados. Tente novamente.</div>';
  } finally {
    if (btnAtualizar) {
      btnAtualizar.disabled = false;
      btnAtualizar.innerHTML = textoOriginalBtn || '<svg class="w-4 h-4" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" viewBox="0 0 24 24"><polyline points="23 4 23 10 17 10"/><path d="M20.49 15a9 9 0 1 1-2.12-9.36L23 10"/></svg> Atualizar';
    }
  }
}

// Funções auxiliares do dashboard (definidas no escopo global para reutilização)
function normalizarMatricula(v){ return String(v ?? '').replace(/\D/g, '').padStart(5, '0'); }
function normalizarSemanaISO(s){ const m = String(s||'').match(/^(\d{4})-W?(\d{1,2})$/i); if(!m) return ''; return `${m[1]}-W${String(m[2]).padStart(2, '0')}`; }
function ativoValido(v){ return ['true','1','sim','yes'].includes(String(v ?? '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim()); }

// Converte uma Semana ISO (ex: "2026-W24") na data (UTC) da segunda-feira
// correspondente. Usado pelo filtro de Data Inicial/Final em "buscar()" para
// determinar quais semanas caem dentro do período informado.
function semanaISOToMonday(isoWeek){
  const m = String(isoWeek||'').match(/^(\d{4})-W(\d{1,2})$/i);
  if (!m) return null;
  const year = parseInt(m[1], 10), week = parseInt(m[2], 10);
  const jan4 = new Date(Date.UTC(year, 0, 4));
  const dow = jan4.getUTCDay() || 7; // 1 (segunda) .. 7 (domingo)
  const monday = new Date(jan4);
  monday.setUTCDate(jan4.getUTCDate() - dow + 1 + (week - 1) * 7);
  monday.setUTCHours(0, 0, 0, 0);
  return monday;
}

// Processar e renderizar os dados do dashboard para um título selecionado
async function dashProcessar(selTit, dashStatus) {
  let semanaAlvo = '';
  if (!semanaAlvo && selTit){
    const hit = (DASH_treinamentos||[]).find(t => String(t.Titulo).trim() === selTit);
    if (hit && hit.SemanaISO) semanaAlvo = String(hit.SemanaISO);
  }

  if (!semanaAlvo){
    dashStatus.innerHTML = '<div class="text-amber-600 bg-amber-50 px-4 py-2 rounded-lg border border-amber-100">Por favor, selecione um Tema para prosseguir.</div>';
    dashRender([], []);
    dashKPIs({ total: 0, part: 0, nPart: 0, semana: '-', titulo: '-', registrosAll: DASH_registrosAll||[], funcAtivos: (DASH_funcionarios||[]).filter(f => ativoValido(f.Ativo??f.ativo)) });
    return;
  }

  const semanaNorm = normalizarSemanaISO(semanaAlvo);
  const mapaPart = new Map();

  for (const r of (DASH_registrosAll||[])){
    if (normalizarSemanaISO(r.SemanaISO) !== semanaNorm) continue;
    const key = normalizarMatricula(r.Matricula);
    const ts = parseTimestamp(r.Timestamp);

    if (!mapaPart.has(key)){
      mapaPart.set(key, r);
    } else {
      const antigo = mapaPart.get(key);
      const tsAnt = parseTimestamp(antigo.Timestamp);
      if (ts && tsAnt && ts < tsAnt){
        mapaPart.set(key, r);
      }
    }
  }

  const participantes = [...mapaPart.values()];
  const funcAtivos = (DASH_funcionarios||[]).filter(f => ativoValido(f.Ativo ?? f.ativo));
  const setMatPart = new Set(participantes.map(p => normalizarMatricula(p.Matricula)));

  // Excluir funcionários em férias/afastamento considerando o período da semana selecionada
  const naoPart = funcAtivos.filter(f =>
    !setMatPart.has(normalizarMatricula(f.Matricula)) &&
    !funcEstaAfastadoNaSemana(f.Matricula, semanaNorm)
  );

  const mapFunc = new Map(funcAtivos.map(f => [normalizarMatricula(f.Matricula), f]));

  const participantesFull = participantes.map(p => {
    const f = mapFunc.get(normalizarMatricula(p.Matricula)) || {};
    return {
      Matricula: normalizarMatricula(p.Matricula ?? f.Matricula ?? ''),
      Nome: p.Nome ?? f.Nome ?? '',
      Setor: p.Setor ?? f.Setor ?? '',
      Timestamp: p.Timestamp ?? ''
    };
  }).sort((a,b) => String(a.Nome ?? '').localeCompare(String(b.Nome ?? ''), 'pt-PT', { sensitivity: 'base' }));

  const naoParticipantesFull = naoPart.map(f => ({
    Matricula: normalizarMatricula(f.Matricula ?? ''),
    Nome: f.Nome ?? '',
    Setor: f.Setor ?? ''
  })).sort((a,b) => String(a.Nome ?? '').localeCompare(String(b.Nome ?? ''), 'pt-PT', { sensitivity: 'base' }));

  DASH_participantes = participantesFull; 
  DASH_naoParticipantes = naoParticipantesFull;

  const total = funcAtivos.filter(f => !funcEstaAfastadoNaSemana(f.Matricula, semanaNorm)).length;
  const part = participantesFull.length;
  const nPart = naoParticipantesFull.length;

  const tituloAlvo = (() => { 
    const t = (DASH_treinamentos||[]).find(t => normalizarSemanaISO(t.SemanaISO) === semanaNorm);
    return t && t.Titulo ? String(t.Titulo) : (selTit || '-');
  })();

  dashKPIs({ total, part, nPart, semana: semanaNorm, titulo: tituloAlvo, registrosAll: DASH_registrosAll || [], funcAtivos });
  dashRender(naoParticipantesFull, participantesFull);
  dashStatus.innerHTML = `<div class="text-emerald-600 bg-emerald-50 px-4 py-2 rounded-lg border border-emerald-100">Painel atualizado para a semana: ${semanaNorm}</div>`;
}

// Exportação Excel de Não Participantes do Dashboard
async function dashGerarXLS_NP(){
  const base = DASH_naoParticipantes || []; 
  if(!base.length){ alert('Não existem dados de ausentes para exportar.'); return; }
  try { await _ensureXLSX(); } catch(e) { alert('Erro ao carregar biblioteca Excel: ' + e.message); return; }
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(base.map(r => ({ 'Matrícula': r.Matricula ?? '', 'Colaborador': r.Nome ?? '', 'Setor': r.Setor ?? '' })));
  XLSX.utils.book_append_sheet(wb, ws, 'Ausentes');
  XLSX.writeFile(wb, `DSS_GIG_NaoParticipantes_${(document.getElementById('kpiSemanaSel').textContent || 'semana')}.xlsx`);
}

// Exportação Excel de Participantes do Dashboard
async function dashGerarXLS_P(){
  const base = DASH_participantes || []; 
  if(!base.length){ alert('Não existem dados de presenças para exportar.'); return; }
  try { await _ensureXLSX(); } catch(e) { alert('Erro ao carregar biblioteca Excel: ' + e.message); return; }
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(base.map(r => ({ 'Matrícula': r.Matricula ?? '', 'Colaborador': r.Nome ?? '', 'Setor': r.Setor ?? '', 'Data de Participação': formatTimestamp(r.Timestamp) })));
  XLSX.utils.book_append_sheet(wb, ws, 'Presencas');
  XLSX.writeFile(wb, `DSS_GIG_Participantes_${(document.getElementById('kpiSemanaSel').textContent || 'semana')}.xlsx`);
}

function dashCompareSemanaISODesc(a, b){ 
  const pa = dashParseSemanaISO(a), pb = dashParseSemanaISO(b); 
  if (pa.year !== pb.year) return pb.year - pa.year; 
  return pb.week - pa.week; 
}

function dashParseSemanaISO(s){ 
  const m = String(s||'').match(/^(\d{4})-W?(\d{1,2})$/i); 
  if(!m) return { year: 0, week: 0 }; 
  return { year: +m[1], week: +m[2] }; 
}

// ── Dark Mode ─────────────────────────────────────────────────
(function initDarkMode(){
  const btn       = document.getElementById('btnDarkMode');
  const iconDark  = document.getElementById('iconDark');
  const iconLight = document.getElementById('iconLight');
  const html      = document.documentElement;

  const apply = (dark) => {
    if (dark) {
      html.classList.add('dark');
      iconDark.classList.add('hidden');
      iconLight.classList.remove('hidden');
      btn.title = 'Mudar para modo claro';
    } else {
      html.classList.remove('dark');
      iconDark.classList.remove('hidden');
      iconLight.classList.add('hidden');
      btn.title = 'Mudar para modo escuro';
    }
    try { localStorage.setItem('dssgig_dark', dark ? '1' : '0'); } catch(e){}
  };

  // Preferência salva ou preferência do sistema
  let saved;
  try { saved = localStorage.getItem('dssgig_dark'); } catch(e){}
  const prefersDark = window.matchMedia?.('(prefers-color-scheme: dark)').matches;
  apply(saved !== null ? saved === '1' : prefersDark);

  btn.addEventListener('click', () => apply(!html.classList.contains('dark')));
})();

// ── Abas / Navegação ─────────────────────────────────────────
(function initTabs(){
  const btnPrincipal  = document.getElementById('tabPrincipal');
  const btnDashboard  = document.getElementById('tabDashboard');
  const secFiltros    = document.getElementById('cardFiltros');
  const secResultados = document.getElementById('cardResultados');
  const secDashboard  = document.getElementById('dashCard');

  function setActive(btnOn, btnOff){
    // Aba activa: destaque azul + sombra
    btnOn.classList.add('tab-active', 'bg-brand-500', 'text-white');
    btnOn.classList.remove('text-slate-600', 'hover:text-slate-900', 'bg-white');
    // Aba inactiva: discreta
    btnOff.classList.remove('tab-active', 'bg-brand-500', 'text-white', 'bg-white');
    btnOff.classList.add('text-slate-600', 'hover:text-slate-900');
  }

  function showPrincipal(){
    setActive(btnPrincipal, btnDashboard);
    secFiltros?.classList.remove('is-hidden');
    secResultados?.classList.remove('is-hidden');
    secDashboard?.classList.add('is-hidden');
  }

  function showDashboard(){
    setActive(btnDashboard, btnPrincipal);
    secFiltros?.classList.add('is-hidden');
    secResultados?.classList.add('is-hidden');
    secDashboard?.classList.remove('is-hidden');
  }

  showPrincipal();

  btnPrincipal.addEventListener('click', showPrincipal);
  btnDashboard.addEventListener('click', showDashboard);
  btnPrincipal.addEventListener('touchstart', e => { e.preventDefault(); showPrincipal(); }, { passive: false });
  btnDashboard.addEventListener('touchstart', e => { e.preventDefault(); showDashboard(); }, { passive: false });
})();
