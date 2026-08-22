/**
 * DSS GIG — API (Apps Script) para Treinamentos Semanais de Segurança
 *
 * Abas esperadas na planilha:
 *   Funcionarios  (Matricula, Nome, Setor, Ativo)
 *   Treinamentos  (SemanaISO, Titulo, URL, Ativo)
 *   Registros     (Timestamp, Matricula, Nome, Setor, SemanaISO, TituloVideo, URLVideo, AssinaturaPNG, DeviceInfo, TempoAssistidoSegundos, DuracaoSegundos)
 *   Ferias        (Matricula, Funcionario, Situacao, InicioFerias, FimFerias)
 *   Perguntas     (SemanaISO, Titulo, NumPergunta, Pergunta, OpcaoA, OpcaoB, OpcaoC, OpcaoD, RespostaCorreta, Ativo)
 *                 — 1 linha por pergunta. Não há limite de quantidade: o desenvolvedor
 *                 do vídeo pode cadastrar 1, 3, 10... quantas perguntas quiser para uma
 *                 mesma semana (SemanaISO + Titulo), basta adicionar mais linhas com
 *                 NumPergunta sequencial (1, 2, 3...). Até o desenvolvedor cadastrar
 *                 perguntas para uma semana, o quiz é automaticamente pulado para aquele
 *                 vídeo (retrocompatível). RespostaCorreta aceita DOIS formatos, para não
 *                 depender de decorar uma convenção: a letra ("B") OU o texto exatamente
 *                 igual ao de uma das colunas OpcaoA-D ("Resp2") — o sistema identifica a
 *                 letra correta automaticamente nesse segundo caso.
 *
 * Endpoints GET  (?action=...):
 *   funcionario    — busca 1 funcionário ativo por matrícula
 *   funcionarios   — retorna todos os funcionários (para o dashboard)
 *   treinamentos   — retorna TODOS os treinamentos ativos (mais recentes primeiro)
 *   registros      — retorna registros com filtros opcionais (matricula, nome, semana). NUNCA inclui a matrícula TESTE.
 *   ferias         — retorna todos os registros da aba Ferias
 *   getFeriasList  — alias de ferias (compatibilidade com o frontend)
 *   perguntas      — retorna as perguntas (sem a resposta correta) de um vídeo, por ?semana= e ?titulo=
 *   identificarColaborador — combinado (funcionário + férias + treinamentos + registros).
 *                 Matrícula especial "TESTE": não precisa existir na aba Funcionarios,
 *                 sempre recebe apenas o treinamento mais recente, nunca é bloqueada
 *                 por férias e nunca aparece como "já assistido" (registros sempre []).
 *
 * Endpoints POST (?action=...):
 *   registrar          — salva presença de um colaborador (valida tempo assistido e quiz no servidor)
 *   validarquiz         — valida as respostas do quiz de um vídeo, sem revelar a resposta correta
 *   addFuncionario     — inclui novo funcionário (verifica duplicidade)
 *   excluirFuncionario — remove funcionário pelo número de matrícula
 *   deleteFuncionario  — alias de excluirFuncionario
 *   excluirColaborador — alias de excluirFuncionario
 *   addFerias          — inclui registro na aba Ferias
 *   deleteFerias       — remove registro da aba Ferias pela matrícula
 *   excluirFerias      — alias de deleteFerias
 *   removeFerias       — alias de deleteFerias
 *
 * Usuário de testes "TESTE": matrícula reservada para o desenvolvedor validar a
 * exibição do vídeo mais recente quantas vezes precisar. Os registros que ela
 * gera são gravados normalmente (para não quebrar o fluxo), mas são filtrados
 * SEMPRE que a matrícula for "TESTE" na action 'registros' — por isso nunca
 * aparecem nos relatórios do DSS Gestor.
 */

// ── PLANILHA ALVO ─────────────────────────────────────────────────────────────
const SPREADSHEET_ID = '1zmEC0-JC-F2zi9oaP3Ea5lUXCNexl3MFf4KnSPMZlOs';
const SHEET_FUNC = 'Funcionarios';
const SHEET_TREI = 'Treinamentos';
const SHEET_REG  = 'Registros';
const SHEET_FER  = 'Ferias';
const SHEET_QUIZ = 'Perguntas';

// Matrícula reservada para testes do desenvolvedor (ver cabeçalho do arquivo)
const MATRICULA_TESTE = 'TESTE';

// ── UTILITÁRIOS ───────────────────────────────────────────────────────────────

function respond(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet(name) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error('Aba não encontrada: ' + name);
  return sh;
}

function getDataAsObjects(sheetName) {
  const sh  = getSheet(sheetName);
  const rng = sh.getDataRange().getValues();
  if (rng.length < 2) return [];
  const headers = rng[0];
  return rng.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[String(h).trim()] = row[i]);
    return obj;
  });
}

function appendRow(sheetName, obj) {
  const sh      = getSheet(sheetName);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const row     = headers.map(h =>
    Object.prototype.hasOwnProperty.call(obj, h) ? obj[h] : ''
  );
  sh.appendRow(row);
}

// ── CACHE (10 minutos) ────────────────────────────────────────────────────────
// Reduz drasticamente o tempo de resposta ao evitar releitura da planilha
// a cada requisição. O cache é invalidado automaticamente após escritas.

const CACHE_TTL = 600;

// `loader` pode ser o nome de uma aba (string, comportamento original) OU uma
// função que retorna os dados (usado para versões "leves" que excluem colunas
// pesadas, como AssinaturaPNG).
//
// O CacheService do Google limita cada CHAVE a 100KB. Planilhas grandes (aqui,
// Registros passa de 1700 linhas) facilmente excedem isso mesmo já sem a
// assinatura (DeviceInfo, URLVideo e TituloVideo somados já bastam). Por isso
// guardamos os dados em vários pedaços de <100KB cada, sob chaves numeradas
// (cacheKey:0, cacheKey:1, ...), e uma chave-índice (cacheKey:meta) com a
// contagem de pedaços — e remontamos tudo na leitura.
function getCachedData(cacheKey, loader) {
  const cache  = CacheService.getScriptCache();
  const cached = readCachedChunks_(cache, cacheKey);
  if (cached !== null) return cached;

  const data = (typeof loader === 'function') ? loader() : getDataAsObjects(loader);
  writeCachedChunks_(cache, cacheKey, data);
  return data;
}

function readCachedChunks_(cache, cacheKey) {
  const meta = cache.get(cacheKey + ':meta');
  if (!meta) return null;
  let n;
  try { n = JSON.parse(meta).n; } catch (e) { return null; }
  if (!n || n < 1) return null;

  const keys = [];
  for (let i = 0; i < n; i++) keys.push(cacheKey + ':' + i);
  const parts = cache.getAll(keys);

  let joined = '';
  for (let i = 0; i < n; i++) {
    const part = parts[cacheKey + ':' + i];
    if (part == null) return null; // pedaço expirado/faltando → cache inválido
    joined += part;
  }
  try { return JSON.parse(joined); } catch (e) { return null; }
}

function writeCachedChunks_(cache, cacheKey, data) {
  try {
    const serialized = JSON.stringify(data);
    const CHUNK_SIZE  = 90000; // margem de segurança abaixo do limite de 100KB
    const n = Math.max(1, Math.ceil(serialized.length / CHUNK_SIZE));
    for (let i = 0; i < n; i++) {
      cache.put(cacheKey + ':' + i, serialized.slice(i * CHUNK_SIZE, (i + 1) * CHUNK_SIZE), CACHE_TTL);
    }
    cache.put(cacheKey + ':meta', JSON.stringify({ n: n }), CACHE_TTL);
  } catch (e) {
    // Dado grande demais até para caber em pedaços — segue sem cache.
  }
}

// Igual a getDataAsObjects, mas omite as colunas indicadas em excludeHeaders
// e guarda o número da linha real da planilha em `_row` (para poder buscar
// uma coluna pesada específica depois, só para os registos filtrados).
function getDataAsObjectsLite(sheetName, excludeHeaders) {
  const sh  = getSheet(sheetName);
  const rng = sh.getDataRange().getValues();
  if (rng.length < 2) return [];
  const headers   = rng[0].map(h => String(h).trim());
  const excludeSet = new Set(excludeHeaders || []);
  return rng.slice(1).map((row, idx) => {
    const obj = { _row: idx + 2 };
    headers.forEach((h, i) => { if (!excludeSet.has(h)) obj[h] = row[i]; });
    return obj;
  });
}

function invalidateCache(cacheKey) {
  try {
    const cache = CacheService.getScriptCache();
    const meta = cache.get(cacheKey + ':meta');
    const n = meta ? (JSON.parse(meta).n || 0) : 0;
    const keys = [cacheKey, cacheKey + ':meta'];
    for (let i = 0; i < n; i++) keys.push(cacheKey + ':' + i);
    cache.removeAll(keys);
  } catch (e) {}
}

// ── NORMALIZAÇÕES ─────────────────────────────────────────────────────────────

function normalizeMatricula(v) {
  // Remove tudo que não é dígito e depois os zeros à esquerda, para que
  // "10", "010" e "00010" sejam sempre tratados como a MESMA matrícula.
  // Sem isso, a mesma pessoa pode "sumir" de uma busca só porque a aba
  // Funcionarios guarda a matrícula com um número de zeros à esquerda
  // diferente da aba Registros (comum quando os dados são digitados à mão
  // em momentos diferentes).
  const digits = String(v || '').trim().replace(/\D/g, '');
  const semZeros = digits.replace(/^0+/, '');
  return semZeros || (digits ? '0' : ''); // preserva "0" se a matrícula for só zeros
}

// Converte um valor de data/hora (Date nativo do Sheets, string ISO, ou
// "dd/mm/aaaa [hh:mm[:ss]]") para milissegundos desde epoch, ou null se não
// for possível interpretar. Usado para filtrar Registros por intervalo de
// datas diretamente no servidor.
function parseTimestampMs_(v) {
  if (!v) return null;
  if (v instanceof Date) return v.getTime();
  const s = String(v).trim();
  if (!s) return null;

  const br = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
  if (br) {
    const [, d, mo, y, hh, mm, ss] = br;
    return new Date(+y, +mo - 1, +d, +(hh || 0), +(mm || 0), +(ss || 0)).getTime();
  }

  const d2 = new Date(s);
  return isNaN(d2.getTime()) ? null : d2.getTime();
}

function normalizeSemanaISO(v) {
  v = String(v || '').toUpperCase().trim();
  const m = v.match(/^(\d{4})-W?(\d{1,2})$/);
  if (!m) return v;
  return m[1] + '-W' + ('0' + m[2]).slice(-2);
}

function isAtivo(v) {
  return ['true', '1', 'sim', 'yes'].includes(
    String(v || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim()
  );
}

// ── DOMÍNIO ───────────────────────────────────────────────────────────────────

// Retorna TODOS os treinamentos ativos ordenados do mais recente para o mais antigo.
// (sem limite de 3 — o frontend decide quantos exibir)
function getAllActiveWeeks(treinamentos) {
  return treinamentos
    .filter(t => isAtivo(t['Ativo']))
    .sort((a, b) => String(b['SemanaISO']).localeCompare(String(a['SemanaISO'])));
}

// Retorna só o treinamento mais recente (usado pelo usuário especial TESTE).
function getTreinamentoMaisRecente_(treinamentosNormalizados) {
  const ativos = getAllActiveWeeks(treinamentosNormalizados);
  return ativos.length ? ativos[0] : null;
}

// ── QUIZ (Perguntas) ────────────────────────────────────────────────────────

// Uma pergunta é considerada ativa por padrão (coluna Ativo em branco = ativa),
// só fica inativa se alguém explicitamente marcar como não/false/0.
function perguntaAtiva_(v) {
  if (v === undefined || v === null || String(v).trim() === '') return true;
  return isAtivo(v);
}

// Busca as perguntas cadastradas para um vídeo específico (SemanaISO + Titulo
// precisam bater exatamente com o que está na aba Treinamentos). Se a aba
// "Perguntas" ainda não existir na planilha, retorna lista vazia em vez de
// quebrar — assim o quiz é simplesmente pulado até ser configurado.
function getPerguntasDoVideo_(semanaISO, titulo) {
  let rows;
  try {
    rows = getCachedData('cache_perguntas', SHEET_QUIZ);
  } catch (e) {
    return [];
  }
  const sem = normalizeSemanaISO(semanaISO);
  const tit = String(titulo || '').trim();
  return rows
    .filter(r =>
      normalizeSemanaISO(r['SemanaISO']) === sem &&
      String(r['Titulo'] || '').trim() === tit &&
      perguntaAtiva_(r['Ativo'])
    )
    .sort((a, b) => Number(a['NumPergunta']) - Number(b['NumPergunta']));
}

// Confere as respostas enviadas contra a aba Perguntas. NUNCA retorna a
// resposta correta para o cliente — só se cada pergunta foi acertada ou não.
// Usada tanto pela action 'validarquiz' (feedback imediato na tela) quanto
// dentro de 'registrar' (segunda checagem, no servidor, antes de gravar).
// Descobre qual letra (A/B/C/D) é a resposta correta de uma pergunta.
// Aceita dois formatos na coluna RespostaCorreta, para não depender de quem
// preenche a planilha lembrar uma convenção específica:
//   1) A letra diretamente: "B"
//   2) O texto igual ao de uma das colunas Opcao A-D: "Resp2" (case-insensitive,
//      ignora espaços nas pontas) — nesse caso a letra é deduzida automaticamente.
function letraRespostaCorreta_(pergunta) {
  const bruto = String(pergunta['RespostaCorreta'] || '').trim();
  if (!bruto) return '';

  const comoLetra = bruto.toUpperCase();
  if (['A', 'B', 'C', 'D'].includes(comoLetra)) return comoLetra;

  // Não é uma letra — tenta casar com o TEXTO de uma das alternativas.
  const norm = s => String(s || '').trim().toLowerCase();
  const opcoes = { A: pergunta['OpcaoA'], B: pergunta['OpcaoB'], C: pergunta['OpcaoC'], D: pergunta['OpcaoD'] };
  for (const letra in opcoes) {
    if (norm(opcoes[letra]) === norm(bruto)) return letra;
  }
  return ''; // não foi possível determinar — nenhuma resposta será aceita como correta
}

function validarRespostasQuiz_(semanaISO, titulo, respostas) {
  const perguntas = getPerguntasDoVideo_(semanaISO, titulo);
  if (perguntas.length === 0) {
    // Vídeo sem quiz configurado — não bloqueia o fluxo (retrocompatível).
    // Aceita QUALQUER quantidade de perguntas cadastradas para o vídeo (1, 3, 10...),
    // não há limite fixo — quem decide é o número de linhas na aba Perguntas.
    return { temQuiz: false, aprovado: true, total: 0, corretas: 0, resultados: [] };
  }

  const respMap = {};
  (respostas || []).forEach(r => {
    respMap[String(r.numPergunta)] = String(r.resposta || '').trim().toUpperCase();
  });

  let corretas = 0;
  const resultados = perguntas.map(p => {
    const num    = String(p['NumPergunta']);
    const dada   = respMap[num] || '';
    const certa  = letraRespostaCorreta_(p);
    const acertou = !!dada && !!certa && dada === certa;
    if (acertou) corretas++;
    return { numPergunta: p['NumPergunta'], correta: acertou };
  });

  return {
    temQuiz: true,
    total: perguntas.length,
    corretas,
    aprovado: corretas === perguntas.length, // exige acertar todas — ajuste aqui se quiser um mínimo diferente
    resultados
  };
}

function findFuncionarioByMatricula(matricula) {
  matricula = normalizeMatricula(matricula);
  const list = getCachedData('cache_funcionarios', SHEET_FUNC);
  return list.find(f =>
    normalizeMatricula(f['Matricula']) === matricula && isAtivo(f['Ativo'])
  );
}

// ── doGet — leitura ───────────────────────────────────────────────────────────

function doGet(e) {
  try {
    const action = (e.parameter.action || '').toLowerCase().trim();

    // ── funcionario (singular — por matrícula) ────────────────────────────────
    if (action === 'funcionario') {
      const matricula = e.parameter.matricula;
      if (!matricula) return respond({ ok: false, error: 'Informe ?matricula=' });
      const func = findFuncionarioByMatricula(matricula);
      if (!func) return respond({ ok: true, found: false });
      return respond({ ok: true, found: true, data: func });
    }

    // ── funcionarios (plural — todos, para o dashboard) ───────────────────────
    if (action === 'funcionarios') {
      const todos = getCachedData('cache_funcionarios', SHEET_FUNC);
      return respond({ ok: true, data: todos });
    }

    // ── treinamentos ──────────────────────────────────────────────────────────
    if (action === 'treinamentos') {
      const rows = getCachedData('cache_treinamentos', SHEET_TREI);
      const normalized = rows.map(t => ({
        ...t,
        SemanaISO: normalizeSemanaISO(t['SemanaISO'])
      }));
      return respond({ ok: true, data: getAllActiveWeeks(normalized) });
    }

    // ── perguntas (quiz de encerramento do vídeo) ─────────────────────────────
    // Nunca retorna a coluna RespostaCorreta — a validação é sempre feita no
    // servidor (actions 'validarquiz' e 'registrar').
    if (action === 'perguntas') {
      const semanaISO = e.parameter.semana || e.parameter.semanaISO || '';
      const titulo    = e.parameter.titulo || '';
      if (!semanaISO || !titulo) return respond({ ok: false, error: 'Informe ?semana= e ?titulo=' });

      const perguntas = getPerguntasDoVideo_(semanaISO, titulo).map(p => ({
        NumPergunta: p['NumPergunta'],
        Pergunta   : p['Pergunta'],
        OpcaoA     : p['OpcaoA'],
        OpcaoB     : p['OpcaoB'],
        OpcaoC     : p['OpcaoC'],
        OpcaoD     : p['OpcaoD']
      }));
      return respond({ ok: true, data: perguntas });
    }

    // ── registros ─────────────────────────────────────────────────────────────
    // A coluna AssinaturaPNG (imagem em base64) é pesada o bastante para
    // impedir o cache de funcionar (limite de 100KB do CacheService), o que
    // fazia o script reler a planilha inteira (1700+ linhas) a cada chamada.
    // Por isso filtramos sempre sobre uma versão "leve" (sem assinatura, e
    // por isso cacheável/rápida) e só buscamos a assinatura, sob demanda,
    // para os poucos registos que efetivamente vão para a resposta —
    // nunca no fetch em massa sem filtros que o Dashboard usa.
    if (action === 'registros') {
      const qMat    = normalizeMatricula(e.parameter.matricula || '');
      const qNome   = String(e.parameter.nome    || '').trim().toLowerCase();
      const qSem    = normalizeSemanaISO(e.parameter.semana || '');
      const qExato  = String(e.parameter.exato   || '') === '1';
      const qTitulo = String(e.parameter.titulo  || '').trim();
      const qDataI  = parseTimestampMs_(e.parameter.dataInicial || '');
      const qDataF0 = parseTimestampMs_(e.parameter.dataFinal   || '');
      const qDataF  = qDataF0 !== null ? qDataF0 + (24 * 60 * 60 * 1000 - 1) : null; // fim do dia
      // Lista explícita de matrículas (separadas por vírgula) — usada pelo
      // PDF para pedir exatamente as assinaturas de quem já está confirmado
      // na lista exibida, sem depender de reler campos do formulário (que
      // podem ter mudado entre a pesquisa e o clique em "Gerar PDF").
      const qMatriculasSet = String(e.parameter.matriculas || '')
        .split(',').map(normalizeMatricula).filter(Boolean);
      const temListaMatriculas = qMatriculasSet.length > 0;

      const base = getCachedData('cache_registros_lite',
        () => getDataAsObjectsLite(SHEET_REG, ['AssinaturaPNG']))
        // O usuário de testes do desenvolvedor (matrícula "TESTE") NUNCA deve
        // aparecer em buscas/relatórios do DSS Gestor — filtrado incondicionalmente,
        // mesmo que algum filtro de matrícula/nome tente encontrá-lo de propósito.
        .filter(r => String(r['Matricula'] || '').trim().toUpperCase() !== MATRICULA_TESTE);

      const matched = base.filter(r => {
        const m = normalizeMatricula(r['Matricula']);
        const n = String(r['Nome']     || '').toLowerCase();
        const s = normalizeSemanaISO(r['SemanaISO']);
        const t = String(r['TituloVideo'] || '');
        const matriculaOk = temListaMatriculas
          ? qMatriculasSet.includes(m)
          : (!qMat || (qExato ? m === qMat : m.includes(qMat)));
        const tituloOk = !qTitulo || t === qTitulo;
        let dataOk = true;
        if (qDataI !== null || qDataF !== null) {
          const ts = parseTimestampMs_(r['Timestamp']);
          if (ts === null) dataOk = false;
          else {
            if (qDataI !== null && ts < qDataI) dataOk = false;
            if (qDataF !== null && ts > qDataF) dataOk = false;
          }
        }
        return matriculaOk
            && (!qNome || n.includes(qNome))
            && (!qSem  || s === qSem)
            && tituloOk
            && dataOk;
      });

      // Busca a assinatura sempre que a pesquisa tem algum filtro (matrícula,
      // nome, semana, título, lista de matrículas ou intervalo de datas) —
      // cobre tanto a busca pontual por pessoa quanto a busca por
      // semana/título da aba Principal, que é o caso mais comum e não usa
      // matrícula/nome. Sem filtro nenhum (fetch em massa do Dashboard, sem
      // imagens) continua rápido e leve.
      // A assinatura só é buscada quando o CLIENTE pede explicitamente
      // (comAssinatura=1) — usado apenas na geração do PDF, nunca na busca
      // normal da tela. Isso mantém a pesquisa da aba Principal sempre leve
      // (sem imagens em base64 na resposta nem buscas linha-a-linha na
      // planilha), já que a assinatura só é realmente necessária no relatório.
      const qComAssinatura = String(e.parameter.comAssinatura || '') === '1';
      const temFiltro = !!(qMat || qNome || qSem || qTitulo || qDataI !== null || qDataF !== null || temListaMatriculas);
      const precisaAssinatura = qComAssinatura && temFiltro && matched.length > 0 && matched.length <= 200;
      let regs = matched;
      if (precisaAssinatura) {
        const sh = getSheet(SHEET_REG);
        const headerRow = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(h => String(h).trim());
        const sigCol = headerRow.indexOf('AssinaturaPNG');
        if (sigCol !== -1) {
          regs = matched.map(r => {
            const sig = sh.getRange(r._row, sigCol + 1, 1, 1).getValue();
            return Object.assign({}, r, { AssinaturaPNG: sig });
          });
        }
      }
      regs = regs.map(r => { const { _row, ...rest } = r; return rest; });

      return respond({ ok: true, data: regs });
    }

    // ── ferias / getFeriasList ────────────────────────────────────────────────
    if (action === 'ferias' || action === 'getferiaslist') {
      try {
        const ferias = getCachedData('cache_ferias', SHEET_FER);
        return respond({ ok: true, data: ferias });
      } catch (e) {
        // Aba Ferias ainda não existe — retorna lista vazia sem erro
        return respond({ ok: true, data: [] });
      }
    }

    // ── identificarColaborador ────────────────────────────────────────────
    // Combina, numa ÚNICA chamada, tudo que o site do colaborador precisa ao
    // identificar alguém: dados do funcionário, férias/afastamentos dele,
    // catálogo de treinamentos e os registros de participação dele. Antes
    // isso eram 4 chamadas separadas (1 sequencial + 3 em paralelo), cada
    // uma pagando o custo fixo de rede/proxy — aqui é uma só.
    if (action === 'identificarcolaborador') {
      const matriculaRaw = String(e.parameter.matricula || '').trim();

      // ── Usuário especial de testes do desenvolvedor ──────────────────────
      // Não precisa existir na aba Funcionarios. Sempre recebe só o
      // treinamento mais recente e "registros" vazio de propósito, para que
      // o vídeo NUNCA seja marcado como já assistido — permitindo assistir
      // (e registrar) quantas vezes forem necessárias para testar a exibição.
      if (matriculaRaw.toUpperCase() === MATRICULA_TESTE) {
        const treinamentosTodos = getCachedData('cache_treinamentos', SHEET_TREI).map(t => ({
          ...t,
          SemanaISO: normalizeSemanaISO(t['SemanaISO'])
        }));
        const maisRecente = getTreinamentoMaisRecente_(treinamentosTodos);

        return respond({
          ok: true,
          found: true,
          funcionario: {
            Matricula: MATRICULA_TESTE,
            Nome: 'Usuário de Teste (Desenvolvedor)',
            Setor: 'TI / Desenvolvimento',
            Ativo: true
          },
          treinamentos: maisRecente ? [maisRecente] : [],
          ferias: [],
          registros: []
        });
      }

      const matricula = normalizeMatricula(matriculaRaw);
      if (!matricula) return respond({ ok: false, error: 'Informe ?matricula=' });

      const func = findFuncionarioByMatricula(matricula);
      if (!func) return respond({ ok: true, found: false });

      const treinamentos = getCachedData('cache_treinamentos', SHEET_TREI);

      let ferias = [];
      try {
        ferias = getCachedData('cache_ferias', SHEET_FER)
          .filter(f => normalizeMatricula(f['Matricula']) === matricula);
      } catch (e) { /* aba Ferias pode não existir ainda */ }

      const baseReg = getCachedData('cache_registros_lite',
        () => getDataAsObjectsLite(SHEET_REG, ['AssinaturaPNG']));
      const registros = baseReg
        .filter(r => normalizeMatricula(r['Matricula']) === matricula)
        .map(r => { const { _row, ...rest } = r; return rest; });

      return respond({ ok: true, found: true, funcionario: func, treinamentos, ferias, registros });
    }

    return respond({
      ok:  true,
      msg: 'DSS GIG API — use ?action=[funcionario|funcionarios|treinamentos|registros|ferias|identificarColaborador|perguntas]'
    });

  } catch (err) {
    return respond({ ok: false, error: String(err) });
  }
}

// ── doPost — escrita ──────────────────────────────────────────────────────────

function doPost(e) {
  try {
    const action = (e.parameter.action || '').toLowerCase().trim();
    const body   = (e.postData && e.postData.contents)
      ? JSON.parse(e.postData.contents)
      : {};

    // ── registrar ─────────────────────────────────────────────────────────────
    if (action === 'registrar') {
      let { matricula, semanaISO, tituloVideo, urlVideo, assinaturaPNG, deviceInfo,
            tempoAssistidoSegundos, duracaoSegundos, respostasQuiz } = body || {};

      const isTeste = String(matricula || '').trim().toUpperCase() === MATRICULA_TESTE;
      semanaISO = normalizeSemanaISO(semanaISO);

      if (!matricula || !semanaISO || !tituloVideo || !urlVideo || !assinaturaPNG) {
        return respond({ ok: false, error: 'Campos obrigatórios: matricula, semanaISO, tituloVideo, urlVideo, assinaturaPNG' });
      }

      // ── Usuário de testes: não existe na aba Funcionarios, então usamos um
      // registro sintético em vez de consultar a planilha.
      let func;
      if (isTeste) {
        func = { Nome: 'Usuário de Teste (Desenvolvedor)', Setor: 'TI / Desenvolvimento' };
        matricula = MATRICULA_TESTE;
      } else {
        matricula = normalizeMatricula(matricula);
        func = findFuncionarioByMatricula(matricula);
        if (!func) return respond({ ok: false, error: 'Funcionário não encontrado ou inativo.' });
      }

      // ── Validação server-side do tempo assistido ────────────────────────────
      // O frontend já bloqueia o botão "Registrar" se o vídeo não chegou ao fim
      // (flag videoEnded), mas esse flag pode ser manipulado por quem chama esta
      // API diretamente (fetch manual, DevTools). Por isso, a prova real de
      // conclusão é conferida aqui: o tempo assistido reportado precisa cobrir
      // pelo menos 90% da duração do vídeo (a tolerância de 10% cobre pequenas
      // variações de buffer/latência normais da reprodução).
      // O usuário TESTE fica isento — o objetivo dele é justamente testar a
      // exibição repetidas vezes, inclusive pulando partes do vídeo.
      const tempoAssistido = Number(tempoAssistidoSegundos) || 0;
      const duracao        = Number(duracaoSegundos) || 0;
      if (!isTeste && duracao > 0 && tempoAssistido < duracao * 0.9) {
        return respond({
          ok: false,
          error: 'O vídeo não foi assistido integralmente pela plataforma. Assista até o final antes de registrar.'
        });
      }

      // ── Validação server-side do quiz ───────────────────────────────────────
      // Segunda checagem (a primeira já aconteceu na action 'validarquiz', que
      // deu o feedback em tela) — garante que ninguém registre sem responder
      // corretamente, mesmo chamando esta API diretamente. Vídeos sem quiz
      // configurado (nenhuma linha na aba Perguntas) não são bloqueados.
      if (!isTeste) {
        const quizResultado = validarRespostasQuiz_(semanaISO, tituloVideo, respostasQuiz);
        if (quizResultado.temQuiz && !quizResultado.aprovado) {
          return respond({
            ok: false,
            error: 'Responda corretamente as perguntas do treinamento antes de registrar.'
          });
        }
      }

      appendRow(SHEET_REG, {
        'Timestamp'    : new Date(),
        'Matricula'    : matricula,
        'Nome'         : func['Nome'],
        'Setor'        : func['Setor'],
        'SemanaISO'    : semanaISO,
        'TituloVideo'  : tituloVideo,
        'URLVideo'     : urlVideo,
        'AssinaturaPNG': assinaturaPNG,
        'DeviceInfo'   : deviceInfo || '',
        // Colunas de auditoria — se ainda não existirem na planilha "Registros",
        // adicione os cabeçalhos "TempoAssistidoSegundos" e "DuracaoSegundos"
        // (appendRow ignora silenciosamente colunas que não existem no cabeçalho).
        'TempoAssistidoSegundos': tempoAssistido,
        'DuracaoSegundos'       : duracao
      });
      // Nota: o registro do usuário TESTE é gravado normalmente na planilha
      // (para não complicar o fluxo), mas a action 'registros' — usada pelo
      // DSS Gestor para gerar relatórios — filtra a matrícula TESTE sempre,
      // então ele nunca aparece em buscas/relatórios.
      invalidateCache('cache_registros_lite');
      return respond({ ok: true, message: 'Registro salvo.' });
    }

    // ── validarquiz ────────────────────────────────────────────────────────────
    // Verifica as respostas do quiz e devolve feedback (quais perguntas foram
    // acertadas) sem nunca revelar qual era a alternativa correta.
    if (action === 'validarquiz') {
      const { semanaISO, tituloVideo, respostas } = body || {};
      if (!semanaISO || !tituloVideo) {
        return respond({ ok: false, error: 'Campos obrigatórios: semanaISO, tituloVideo' });
      }
      const resultado = validarRespostasQuiz_(normalizeSemanaISO(semanaISO), tituloVideo, respostas);
      return respond(Object.assign({ ok: true }, resultado));
    }

    // ── addFuncionario ────────────────────────────────────────────────────────
    if (action === 'addfuncionario') return addFuncionario_(body);

    // ── excluirFuncionario (e aliases) ────────────────────────────────────────
    if (['excluirfuncionario', 'deletefuncionario', 'excluircolaborador'].includes(action)) {
      return excluirFuncionario_(body);
    }

    // ── addFerias ─────────────────────────────────────────────────────────────
    if (action === 'addferias') return addFerias_(body);

    // ── deleteFerias (e aliases) ──────────────────────────────────────────────
    if (['deleteferias', 'excluirferias', 'removeferias'].includes(action)) {
      return deleteFerias_(body);
    }

    return respond({ ok: false, error: 'Ação não reconhecida: ' + action });

  } catch (err) {
    return respond({ ok: false, error: String(err) });
  }
}

// ── addFuncionario ────────────────────────────────────────────────────────────

function addFuncionario_(payload) {
  const matricula = normalizeMatricula(payload.matricula);
  const nome      = String(payload.nome  || '').trim();
  const setor     = String(payload.setor || '').trim();

  if (!matricula || !nome) return respond({ ok: false, error: 'Matrícula e Nome são obrigatórios.' });

  const sh   = getSheet(SHEET_FUNC);
  const lock = LockService.getScriptLock();
  lock.tryLock(30000);
  try {
    const last = sh.getLastRow();
    if (last >= 2) {
      const values = sh.getRange(2, 1, last - 1, 1).getValues();
      if (values.some(r => normalizeMatricula(r[0]) === matricula)) {
        return respond({ ok: false, error: 'Matrícula já cadastrada.' });
      }
    }
    sh.getRange(last + 1, 1, 1, 4).setValues([[matricula, nome, setor, true]]);
    invalidateCache('cache_funcionarios');
    return respond({ ok: true });
  } catch (err) {
    return respond({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

// ── excluirFuncionario ────────────────────────────────────────────────────────

function excluirFuncionario_(payload) {
  const matricula = normalizeMatricula(payload.matricula);
  if (!matricula) return respond({ ok: false, error: 'Matrícula não informada.' });

  const sh   = getSheet(SHEET_FUNC);
  const lock = LockService.getScriptLock();
  lock.tryLock(30000);
  try {
    const last = sh.getLastRow();
    if (last < 2) return respond({ ok: false, error: 'Nenhum funcionário cadastrado.' });
    const values = sh.getRange(2, 1, last - 1, 1).getValues();
    let found = false;
    for (let i = values.length - 1; i >= 0; i--) {
      if (normalizeMatricula(values[i][0]) === matricula) {
        sh.deleteRow(i + 2);
        found = true;
      }
    }
    if (!found) return respond({ ok: false, error: 'Matrícula não encontrada.' });
    invalidateCache('cache_funcionarios');
    return respond({ ok: true, message: 'Funcionário removido.' });
  } catch (err) {
    return respond({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

// ── addFerias ─────────────────────────────────────────────────────────────────
// A partir de agora, Início e Fim são obrigatórios para TODAS as situações
// (Férias e Dispensa Médica). Registros sem período não são mais aceitos.

function addFerias_(payload) {
  const matricula    = normalizeMatricula(payload.matricula);
  const funcionario  = String(payload.funcionario  || '').trim();
  const situacao     = String(payload.situacao     || 'Férias').trim();
  const inicioFerias = String(payload.inicioFerias || '').trim();
  const fimFerias    = String(payload.fimFerias    || '').trim();

  if (!matricula || !funcionario) {
    return respond({ ok: false, error: 'Matrícula e Funcionario são obrigatórios.' });
  }
  // Datas obrigatórias para todas as situações
  if (!inicioFerias || !fimFerias) {
    return respond({ ok: false, error: 'Início e Fim do período são obrigatórios para todas as situações.' });
  }

  let sh;
  try {
    sh = getSheet(SHEET_FER);
  } catch (e) {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    sh = ss.insertSheet(SHEET_FER);
    sh.getRange(1, 1, 1, 5).setValues([['Matricula', 'Funcionario', 'Situacao', 'InicioFerias', 'FimFerias']]);
    sh.setFrozenRows(1);
  }

  const lock = LockService.getScriptLock();
  lock.tryLock(30000);
  try {
    sh.appendRow([matricula, funcionario, situacao, inicioFerias, fimFerias]);
    invalidateCache('cache_ferias');
    return respond({ ok: true, message: 'Registro inserido.' });
  } catch (err) {
    return respond({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

// ── deleteFerias ──────────────────────────────────────────────────────────────

function deleteFerias_(payload) {
  const matricula = normalizeMatricula(payload.matricula);
  const situacao  = payload.situacao ? String(payload.situacao).trim().toLowerCase() : null;
  if (!matricula) return respond({ ok: false, error: 'Matrícula não informada.' });

  const sh   = getSheet(SHEET_FER);
  const lock = LockService.getScriptLock();
  lock.tryLock(30000);
  try {
    const last = sh.getLastRow();
    if (last < 2) return respond({ ok: false, error: 'Aba Ferias está vazia.' });
    const data = sh.getRange(2, 1, last - 1, 3).getValues();
    let removed = 0;
    for (let i = data.length - 1; i >= 0; i--) {
      const rowMat = normalizeMatricula(data[i][0]);
      const rowSit = String(data[i][2]).trim().toLowerCase();
      if (rowMat !== matricula) continue;
      if (situacao && !rowSit.includes(situacao) && !situacao.includes(rowSit)) continue;
      sh.deleteRow(i + 2);
      removed++;
    }
    if (removed === 0) return respond({ ok: false, error: 'Registro não encontrado na aba Ferias.' });
    invalidateCache('cache_ferias');
    return respond({ ok: true, message: `${removed} registro(s) removido(s).` });
  } catch (err) {
    return respond({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

// ── AQUECIMENTO AUTOMÁTICO DE CACHE ─────────────────────────────────────────
// Sem isso, a cada 10 minutos (CACHE_TTL) o cache expira e o PRÓXIMO usuário
// a abrir o painel é quem "paga o pato": sofre a consulta lenta de reler a
// planilha inteira (e corre o risco de dar timeout numa planilha grande).
//
// Esta função relê as abas mais pesadas (Registros, Treinamentos,
// Funcionarios) e repõe o cache ANTES dele expirar de fato — assim nenhum
// usuário nunca chega a ver a versão "fria".
//
// Configuração necessária (só uma vez): no editor do Apps Script, vá em
// "Acionadores" (ícone de relógio na barra lateral esquerda) → "+ Adicionar
// acionador" → escolha a função "warmCache" → tipo de evento "Baseado em
// tempo" → "Temporizador por minutos" → a cada 5 minutos → Salvar.
function warmCache() {
  try { getCachedData('cache_registros_lite', () => getDataAsObjectsLite(SHEET_REG, ['AssinaturaPNG'])); } catch (e) {}
  try { getCachedData('cache_treinamentos', SHEET_TREI); } catch (e) {}
  try { getCachedData('cache_funcionarios', SHEET_FUNC); } catch (e) {}
}
