// DSS GIG — Colaborador (colab.js)
// Toda a lógica de identificação, listagem de vídeos, player e registro.

const API_BASE = '/api/gas';

let funcionario = null;
let treinamentos = [];
let selectedVideo = null;
let player = null;
let signaturePad = null;
let videoEnded = false;
let ytReady = false;

// ─── Utilitários ──────────────────────────────────────────────────────────────

function extractYouTubeId(url) {
  try {
    const u = new URL(url);
    if (u.hostname.includes('youtu.be')) return u.pathname.slice(1);
    if (u.searchParams.get('v')) return u.searchParams.get('v');
    const m = url.match(/embed\/([\w\-]+)/);
    if (m) return m[1];
  } catch (e) {}
  return null;
}

function normMat(v) {
  return String(v || '').replace(/\D/g, '').padStart(5, '0');
}

function parseDateBR(s) {
  if (!s) return null;
  const m = String(s).trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (!m) return null;
  return new Date(Date.UTC(+m[3], +m[2] - 1, +m[1]));
}

function isoParaSegunda(iso) {
  const m = String(iso || '').match(/^(\d{4})-W(\d{1,2})$/i);
  if (!m) return null;
  const jan4 = new Date(Date.UTC(+m[1], 0, 4));
  const dow = jan4.getUTCDay() || 7;
  const seg = new Date(jan4);
  seg.setUTCDate(jan4.getUTCDate() - dow + 1 + (+m[2] - 1) * 7);
  return seg;
}

function isoParaDomingo(iso) {
  const seg = isoParaSegunda(iso);
  if (!seg) return null;
  const dom = new Date(seg);
  dom.setUTCDate(seg.getUTCDate() + 6);
  return dom;
}

// Retorna a semana ISO da data atual no formato "AAAA-WNN"
function getSemanaAtualISO() {
  const hoje = new Date();
  // Janeiro 4 sempre cai na semana 1 (norma ISO 8601)
  const jan4 = new Date(Date.UTC(hoje.getUTCFullYear(), 0, 4));
  const dow = jan4.getUTCDay() || 7;
  const seg1 = new Date(jan4);
  seg1.setUTCDate(jan4.getUTCDate() - dow + 1);

  const hojeUTC = new Date(Date.UTC(hoje.getUTCFullYear(), hoje.getUTCMonth(), hoje.getUTCDate()));
  const diffDias = Math.round((hojeUTC - seg1) / 86400000);
  let semana = Math.floor(diffDias / 7) + 1;
  let ano = hoje.getUTCFullYear();

  // Ajuste de ano (semana 53 pode pertencer ao próximo ano)
  if (semana < 1) { ano--; semana = 52; }
  if (semana > 52) {
    const jan4Prox = new Date(Date.UTC(ano + 1, 0, 4));
    const dow2 = jan4Prox.getUTCDay() || 7;
    const seg1Prox = new Date(jan4Prox);
    seg1Prox.setUTCDate(jan4Prox.getUTCDate() - dow2 + 1);
    if (hojeUTC >= seg1Prox) { ano++; semana = 1; }
  }

  return `${ano}-W${String(semana).padStart(2, '0')}`;
}

function semanaEmFerias(semanaISO, inicioFerias, fimFerias) {
  const seg = isoParaSegunda(semanaISO);
  const dom = isoParaDomingo(semanaISO);
  if (!seg || !dom) return true;
  const ini = parseDateBR(inicioFerias);
  const fim = parseDateBR(fimFerias);
  if (!ini && !fim) return true;
  if (ini && !fim) return dom >= ini;
  if (!ini && fim) return seg <= fim;
  return seg <= fim && dom >= ini;
}

// ─── Chamadas à API ───────────────────────────────────────────────────────────

async function apiFetch(params) {
  const res = await fetch(API_BASE + '?' + params + '&_t=' + Date.now());
  try { return await res.json(); }
  catch { return { ok: false, error: 'Resposta inválida', status: res.status }; }
}

async function fetchFuncionario(m) {
  return apiFetch('action=funcionario&matricula=' + encodeURIComponent(m));
}

async function fetchTreinamentos() {
  return apiFetch('action=treinamentos');
}

async function fetchRegistros(m) {
  return apiFetch('action=registros&matricula=' + encodeURIComponent(m));
}

async function fetchFeriasList() {
  try { return await apiFetch('action=getFeriasList'); }
  catch { return { ok: false }; }
}

// ─── UI helpers ───────────────────────────────────────────────────────────────

function showInfo(html) {
  const el = document.getElementById('funcInfo');
  el.innerHTML = html;
  el.classList.remove('hidden');
}

function hideInfo() {
  const el = document.getElementById('funcInfo');
  el.innerHTML = '';
  el.classList.add('hidden');
}

function alertCard(type, html) {
  const styles = {
    ok:   'bg-emerald-50 dark:bg-emerald-900/30 border border-emerald-200 dark:border-emerald-700 text-emerald-800 dark:text-emerald-200',
    warn: 'bg-amber-50 dark:bg-amber-900/30 border border-amber-200 dark:border-amber-700 text-amber-800 dark:text-amber-200',
    error:'bg-red-50 dark:bg-red-900/30 border border-red-200 dark:border-red-700 text-red-800 dark:text-red-200',
    block:'bg-red-50 dark:bg-red-900/40 border-2 border-red-400 dark:border-red-600 text-red-900 dark:text-red-200',
  };
  return `<div class="px-4 py-3 rounded-xl text-sm font-medium ${styles[type] || styles.warn}">${html}</div>`;
}

// ─── Dark Mode ────────────────────────────────────────────────────────────────

(function initDarkMode() {
  const html = document.documentElement;
  const sun  = document.getElementById('iconSun');
  const moon = document.getElementById('iconMoon');

  // Cores da caneta: claro = azul muito escuro, escuro = branco suave
  const PEN_LIGHT = '#0f172a';
  const PEN_DARK  = '#e2e8f0';

  function getPenColor() {
    return html.classList.contains('dark') ? PEN_DARK : PEN_LIGHT;
  }

  function setDark(dark) {
    html.classList.toggle('dark', dark);
    sun.classList.toggle('hidden', !dark);
    moon.classList.toggle('hidden', dark);
    try { localStorage.setItem('dssgig_dark', dark ? '1' : '0'); } catch (e) {}

    // Atualizar cor da caneta em tempo real, sem apagar a assinatura
    if (signaturePad) {
      signaturePad.penColor = getPenColor();
    }

    // Atualizar cor de fundo do canvas para contrastar com o tema
    const canvas = document.getElementById('signaturePad');
    if (canvas) {
      canvas.style.background = dark ? '#0f172a' : '#ffffff';
    }
  }

  let saved = '0';
  try { saved = localStorage.getItem('dssgig_dark') || '0'; } catch (e) {}
  setDark(saved === '1');

  document.getElementById('btnDarkMode').addEventListener('click', () => {
    setDark(!html.classList.contains('dark'));
  });

  // Expor getPenColor para uso na inicialização do SignaturePad
  window._getPenColor = getPenColor;
})();

// ─── YouTube API ──────────────────────────────────────────────────────────────

// Callback global chamado pelo script da YouTube IFrame API
window.onYouTubeIframeAPIReady = function () {
  ytReady = true;
};

function onPlayerStateChange(e) {
  const playBtn = document.getElementById('btnPlayOverlay');
  if (!playBtn) return;

  if (e.data === YT.PlayerState.PLAYING) {
    playBtn.classList.add('hidden');
  } else if (e.data === YT.PlayerState.PAUSED || e.data === YT.PlayerState.CUED) {
    playBtn.classList.remove('hidden');
  }

  if (e.data === YT.PlayerState.ENDED) {
    videoEnded = true;
    playBtn.classList.remove('hidden');

    // Mostrar o bloco ANTES de medir o canvas, para que offsetWidth seja correto
    const signBlock = document.getElementById('signatureBlock');
    signBlock.classList.remove('hidden');
    document.getElementById('btnRegistrar').classList.remove('hidden');

    // Inicializar o pad de assinatura com tamanho correto
    const canvas = document.getElementById('signaturePad');
    // Forçar largura explícita pelo offsetWidth (só funciona após display:block)
    const containerWidth = canvas.parentElement ? canvas.parentElement.clientWidth - 4 : 600;
    canvas.width  = containerWidth > 0 ? containerWidth : 600;
    canvas.height = 280;

    // Aguardar SignaturePad estar disponível (pode ainda não ter carregado)
    const initPad = () => {
      if (typeof SignaturePad === 'undefined') {
        setTimeout(initPad, 100);
        return;
      }
      const isDark = document.documentElement.classList.contains('dark');
      canvas.style.background = isDark ? '#0f172a' : '#ffffff';

      if (!signaturePad) {
        signaturePad = new SignaturePad(canvas, {
          minWidth: 1,
          maxWidth: 3,
          penColor: (window._getPenColor ? window._getPenColor() : (isDark ? '#e2e8f0' : '#0f172a')),
          backgroundColor: 'rgba(0,0,0,0)',
        });
      } else {
        signaturePad.penColor = window._getPenColor ? window._getPenColor() : (isDark ? '#e2e8f0' : '#0f172a');
        signaturePad.clear();
      }
    };
    initPad();

    // Rolar suavemente até a assinatura
    signBlock.scrollIntoView({ behavior: 'smooth', block: 'start' });
  }
}

// ─── Seleção e carregamento do vídeo ─────────────────────────────────────────

function selectVideo(idx) {
  selectedVideo = treinamentos[idx];
  videoEnded = false;

  document.getElementById('videoTitle').textContent = selectedVideo.Titulo;
  document.getElementById('semanaISO').textContent = selectedVideo.SemanaISO;
  document.getElementById('btnRegistrar').classList.add('hidden');
  document.getElementById('signatureBlock').classList.add('hidden');
  document.getElementById('mensagem').innerHTML = '';
  document.getElementById('btnRegistrar').disabled = false;

  const vid = extractYouTubeId(selectedVideo.URL);
  if (!vid) {
    document.getElementById('mensagem').innerHTML = alertCard('error', 'URL do vídeo inválida.');
    return;
  }

  const loadPlayer = () => {
    if (player) {
      player.loadVideoById(vid);
      return;
    }
    if (!window.YT || !YT.Player) {
      setTimeout(loadPlayer, 250);
      return;
    }
    player = new YT.Player('ytplayer', {
      height: '100%',
      width: '100%',
      videoId: vid,
      playerVars: { controls: 0, disablekb: 1, fs: 0, rel: 0, modestbranding: 1, playsinline: 1 },
      events: { onStateChange: onPlayerStateChange },
    });
  };
  loadPlayer();

  const playBtn = document.getElementById('btnPlayOverlay');
  playBtn.classList.remove('hidden');
  playBtn.onclick = () => { if (player && player.playVideo) player.playVideo(); };

  document.getElementById('secPlayer').classList.remove('hidden');
  document.getElementById('secPlayer').scrollIntoView({ behavior: 'smooth', block: 'start' });
}

// ─── Renderização da lista de vídeos ─────────────────────────────────────────

function renderListaVideos(list) {
  const c = document.getElementById('listaVideos');
  c.innerHTML = '';

  if (!list || list.length === 0) {
    c.innerHTML = alertCard('ok',
      '<span class="text-lg mr-2">✅</span> Você já está em dia! Não há vídeos pendentes para esta semana.');
    return;
  }

  list.forEach((t, idx) => {
    const card = document.createElement('div');
    card.className = [
      'flex items-center justify-between gap-4 p-4',
      'bg-slate-50 dark:bg-slate-700/50 hover:bg-slate-100 dark:hover:bg-slate-700',
      'border border-slate-200 dark:border-slate-600 rounded-xl transition-colors',
    ].join(' ');

    card.innerHTML = `
      <div class="min-w-0">
        <p class="text-xs font-semibold text-brand-600 dark:text-brand-400 mb-0.5">
          Semana ${t.SemanaISO}
        </p>
        <h3 class="font-semibold text-slate-800 dark:text-white text-sm leading-snug truncate">${t.Titulo}</h3>
      </div>
      <button data-idx="${idx}"
        class="shrink-0 bg-brand-500 hover:bg-brand-600 active:bg-brand-700 text-white font-semibold py-2 px-4 rounded-xl text-sm transition-all shadow-sm hover:shadow-md flex items-center gap-1.5 whitespace-nowrap">
        <svg class="w-4 h-4" fill="currentColor" viewBox="0 0 24 24"><polygon points="5,3 19,12 5,21"/></svg>
        Assistir
      </button>`;

    card.querySelector('button').addEventListener('click', () => selectVideo(idx));
    c.appendChild(card);
  });
}

// ─── Botão Buscar ─────────────────────────────────────────────────────────────

document.getElementById('btnBuscar').addEventListener('click', async () => {
  const m = document.getElementById('matricula').value.trim();
  hideInfo();
  document.getElementById('secVideos').classList.add('hidden');
  document.getElementById('secPlayer').classList.add('hidden');

  if (!m) {
    showInfo(alertCard('warn', 'Informe a matrícula.'));
    return;
  }

  // Estado de carregamento
  const btnBuscar = document.getElementById('btnBuscar');
  const txtOriginal = btnBuscar.innerHTML;
  btnBuscar.disabled = true;
  btnBuscar.innerHTML = `<svg class="w-4 h-4 animate-spin" fill="none" viewBox="0 0 24 24">
    <circle class="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" stroke-width="4"/>
    <path class="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8v8z"/>
  </svg> A buscar...`;

  try {
    // 1. Verificar funcionário
    const resp = await fetchFuncionario(m);
    if (!resp || !resp.ok) {
      const d = resp && resp.error ? ` (${resp.error})` : '';
      showInfo(alertCard('error', `Falha ao consultar funcionário${d}.`));
      return;
    }
    if (resp.found === false) {
      showInfo(alertCard('error', 'Funcionário não encontrado ou inativo.'));
      return;
    }

    funcionario = resp.data;
    const matNorm = normMat(funcionario.Matricula);

    // 2. Carregar férias, treinamentos e registros em paralelo
    const [fResp, tResp, rResp] = await Promise.all([
      fetchFeriasList(),
      fetchTreinamentos(),
      fetchRegistros(funcionario.Matricula),
    ]);

    // 3. Processar registros de férias/afastamento deste colaborador
    let listaFerias = [];
    if (fResp && fResp.ok && Array.isArray(fResp.data)) {
      listaFerias = fResp.data.filter(f => normMat(f.Matricula) === matNorm);
    }

    // Verificar se está de férias/afastamento HOJE — bloqueia acesso total
    const hoje = new Date();
    hoje.setUTCHours(0, 0, 0, 0);
    const feriaHoje = listaFerias.find(f => {
      const ini = parseDateBR(f.InicioFerias);
      const fim = parseDateBR(f.FimFerias);
      if (!ini && !fim) return true;           // sem período = sempre bloqueado
      if (ini && !fim) return hoje >= ini;
      if (!ini && fim) return hoje <= fim;
      return hoje >= ini && hoje <= fim;
    });

    if (feriaHoje) {
      const sit = (feriaHoje.Situacao || 'Férias / Afastamento').toUpperCase();
      // Usar as strings originais (dd/mm/aaaa) diretamente — sem converter para Date
      const ini = feriaHoje.InicioFerias || '';
      const fim = feriaHoje.FimFerias    || '';
      const periodo = (ini || fim) ? ` (${ini || '?'}${fim ? ' a ' + fim : ''})` : '';
      showInfo(`
        <div class="flex flex-col items-center gap-3 text-center py-2">
          <span class="text-4xl">🚫</span>
          <p class="font-extrabold text-red-700 dark:text-red-400 text-sm uppercase tracking-wide leading-snug">
            Você não está autorizado a acessar<br>esta plataforma neste período.
          </p>
          <p class="text-xs font-semibold text-red-600 dark:text-red-400">
            Motivo: ${sit}${periodo}
          </p>
        </div>`);
      return;
    }

    // 4. Saudação
    showInfo(alertCard('ok',
      `Olá, <strong>${funcionario.Nome}</strong> — Setor: <strong>${funcionario.Setor || '-'}</strong>`));

    if (!tResp || !tResp.ok) {
      const d = tResp && tResp.error ? ` (${tResp.error})` : '';
      showInfo(alertCard('error', `Falha ao carregar vídeos${d}.`));
      return;
    }

    let todosTreinamentos = tResp.data || [];

    // 5. Remover vídeos cujo período (semana ISO) coincide com QUALQUER
    //    registro de férias/afastamento — mesmo após o retorno do colaborador,
    //    esses vídeos nunca estarão disponíveis para ele.
    if (listaFerias.length > 0) {
      todosTreinamentos = todosTreinamentos.filter(t => {
        const cobertaPorFerias = listaFerias.some(f =>
          semanaEmFerias(t.SemanaISO, f.InicioFerias, f.FimFerias)
        );
        return !cobertaPorFerias;
      });
    }

    // 6. Exibir APENAS o vídeo da semana vigente.
    //    Vídeos de semanas passadas não ficam disponíveis — se o colaborador
    //    perdeu a semana, o vídeo não é mais acessível.
    const semanaHoje = getSemanaAtualISO();
    todosTreinamentos = todosTreinamentos.filter(t => {
      // Normalizar semana do vídeo para comparação
      const iso = String(t.SemanaISO || '').trim().toUpperCase();
      const m = iso.match(/^(\d{4})-W?(\d{1,2})$/);
      if (!m) return false;
      const isoNorm = `${m[1]}-W${String(m[2]).padStart(2, '0')}`;
      return isoNorm === semanaHoje;
    });

    // 6. Remover vídeos já assistidos e registrados
    if (!rResp || !rResp.ok) {
      showInfo(alertCard('ok',
        `Olá, <strong>${funcionario.Nome}</strong> — Setor: <strong>${funcionario.Setor || '-'}</strong>`) +
        alertCard('warn', 'Não foi possível verificar vídeos já assistidos. Mostrando todos disponíveis.'));
      treinamentos = todosTreinamentos;
    } else {
      const registros = rResp.data || [];
      // Chave única por semana ISO + título do vídeo
      const watchedKey = new Set(registros.map(r => `${r.SemanaISO}@@${r.TituloVideo}`));
      treinamentos = todosTreinamentos.filter(t => !watchedKey.has(`${t.SemanaISO}@@${t.Titulo}`));
    }

    renderListaVideos(treinamentos);
    document.getElementById('secVideos').classList.remove('hidden');
    document.getElementById('secVideos').scrollIntoView({ behavior: 'smooth', block: 'start' });

  } finally {
    btnBuscar.disabled = false;
    btnBuscar.innerHTML = txtOriginal;
  }
});

// Buscar ao pressionar Enter na matrícula
document.getElementById('matricula').addEventListener('keydown', e => {
  if (e.key === 'Enter') document.getElementById('btnBuscar').click();
});

// ─── Redimensionamento do canvas de assinatura ────────────────────────────────
// Quando o utilizador roda o telemóvel ou redimensiona a janela, o canvas
// precisa de ser ajustado — caso contrário a assinatura fica distorcida.
(function initSignatureResize() {
  const canvas = document.getElementById('signaturePad');
  if (!canvas) return;

  let resizeTimer;
  const resizeCanvas = () => {
    if (!signaturePad) return;
    // Guardar imagem atual antes de redimensionar
    const data = signaturePad.isEmpty() ? null : signaturePad.toDataURL();
    const parent = canvas.parentElement;
    const newW = parent ? parent.clientWidth - 4 : 600;
    if (newW > 0 && canvas.width !== newW) {
      canvas.width  = newW;
      canvas.height = 280;
      signaturePad.clear();
      // Restaurar imagem guardada
      if (data) {
        const img = new Image();
        img.onload = () => {
          const ctx = canvas.getContext('2d');
          ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
          signaturePad.fromDataURL(data);
        };
        img.src = data;
      }
    }
  };

  window.addEventListener('resize', () => {
    clearTimeout(resizeTimer);
    resizeTimer = setTimeout(resizeCanvas, 150);
  });
})();



document.getElementById('btnTrocarVideo').addEventListener('click', () => {
  if (player) { try { player.stopVideo(); } catch (e) {} }
  document.getElementById('secPlayer').classList.add('hidden');
  document.getElementById('btnRegistrar').classList.add('hidden');
  document.getElementById('signatureBlock').classList.add('hidden');
  document.getElementById('mensagem').innerHTML = '';
  selectedVideo = null;
  videoEnded = false;
  document.getElementById('secVideos').scrollIntoView({ behavior: 'smooth', block: 'start' });
});

// ─── Botão Limpar Assinatura ──────────────────────────────────────────────────

document.getElementById('btnLimpar').addEventListener('click', () => {
  if (signaturePad) signaturePad.clear();
});

// ─── Botão Registrar ──────────────────────────────────────────────────────────

document.getElementById('btnRegistrar').addEventListener('click', async () => {
  const msg = document.getElementById('mensagem');
  msg.innerHTML = '';

  if (!funcionario || !selectedVideo) {
    msg.innerHTML = alertCard('error', 'Dados incompletos. Recarregue a página e tente novamente.');
    return;
  }
  if (!videoEnded) {
    msg.innerHTML = alertCard('warn', 'O vídeo ainda não foi concluído. Assista até o final.');
    return;
  }
  if (!signaturePad || signaturePad.isEmpty()) {
    msg.innerHTML = alertCard('warn', 'Por favor, assine no quadro antes de registrar.');
    document.getElementById('signatureBlock').scrollIntoView({ behavior: 'smooth', block: 'start' });
    return;
  }

  const btnReg = document.getElementById('btnRegistrar');
  const txtOriginal = btnReg.innerHTML;
  btnReg.disabled = true;
  btnReg.innerHTML = `<svg class="w-4 h-4 animate-spin" fill="none" viewBox="0 0 24 24">
    <circle class="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" stroke-width="4"/>
    <path class="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8v8z"/>
  </svg> A registrar...`;

  const payload = {
    matricula: String(funcionario.Matricula),
    semanaISO: String(selectedVideo.SemanaISO),
    tituloVideo: String(selectedVideo.Titulo),
    urlVideo: String(selectedVideo.URL),
    assinaturaPNG: signaturePad.toDataURL('image/png'),
    deviceInfo: navigator.userAgent,
  };

  try {
    const res = await fetch(API_BASE + '?action=registrar', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload),
    });

    let data;
    try { data = await res.json(); }
    catch { data = { ok: false, error: 'Resposta inválida do servidor.' }; }

    if (data.ok) {
      msg.innerHTML = `
        <div class="flex flex-col items-center gap-3 text-center py-4 px-6 bg-emerald-50 dark:bg-emerald-900/30 border border-emerald-200 dark:border-emerald-700 rounded-xl">
          <span class="text-4xl">✅</span>
          <p class="font-bold text-emerald-800 dark:text-emerald-200">Registro realizado com sucesso!</p>
          <p class="text-xs text-emerald-700 dark:text-emerald-300">
            Obrigado, <strong>${funcionario.Nome}</strong>. Sua participação foi registrada.
          </p>
        </div>`;

      // Remover da lista local e re-renderizar
      treinamentos = treinamentos.filter(t =>
        !(t.SemanaISO === selectedVideo.SemanaISO && t.Titulo === selectedVideo.Titulo));
      renderListaVideos(treinamentos);
      document.getElementById('secPlayer').classList.add('hidden');
      document.getElementById('secVideos').classList.remove('hidden');

      // Rolar até a mensagem
      msg.scrollIntoView({ behavior: 'smooth', block: 'center' });

    } else {
      msg.innerHTML = alertCard('error', 'Falha ao registrar: ' + (data.error || 'Erro desconhecido.'));
      btnReg.disabled = false;
      btnReg.innerHTML = txtOriginal;
    }

  } catch (err) {
    msg.innerHTML = alertCard('error', 'Erro de conexão ao salvar o registro. Verifique sua internet e tente novamente.');
    btnReg.disabled = false;
    btnReg.innerHTML = txtOriginal;
  }
});
