/**
 * DSS GIG — API (Apps Script) para Treinamentos Semanais de Segurança
 *
 * Abas esperadas na planilha:
 *   Funcionarios  (Matricula, Nome, Setor, Ativo)
 *   Treinamentos  (SemanaISO, Titulo, URL, Ativo)
 *   Registros     (Timestamp, Matricula, Nome, Setor, SemanaISO, TituloVideo, URLVideo, AssinaturaPNG, DeviceInfo)
 *   Ferias        (Matricula, Funcionario, Situacao, InicioFerias, FimFerias)
 *
 * Endpoints GET  (?action=...):
 *   funcionario    — busca 1 funcionário ativo por matrícula
 *   funcionarios   — retorna todos os funcionários (para o dashboard)
 *   treinamentos   — retorna TODOS os treinamentos ativos (mais recentes primeiro)
 *   registros      — retorna registros com filtros opcionais (matricula, nome, semana)
 *   ferias         — retorna todos os registros da aba Ferias
 *   getFeriasList  — alias de ferias (compatibilidade com o frontend)
 *
 * Endpoints POST (?action=...):
 *   registrar          — salva presença de um colaborador
 *   addFuncionario     — inclui novo funcionário (verifica duplicidade)
 *   excluirFuncionario — remove funcionário pelo número de matrícula
 *   deleteFuncionario  — alias de excluirFuncionario
 *   excluirColaborador — alias de excluirFuncionario
 *   addFerias          — inclui registro na aba Ferias
 *   deleteFerias       — remove registro da aba Ferias pela matrícula
 *   excluirFerias      — alias de deleteFerias
 *   removeFerias       — alias de deleteFerias
 */

// ── PLANILHA ALVO ─────────────────────────────────────────────────────────────
const SPREADSHEET_ID = '1zmEC0-JC-F2zi9oaP3Ea5lUXCNexl3MFf4KnSPMZlOs';
const SHEET_FUNC = 'Funcionarios';
const SHEET_TREI = 'Treinamentos';
const SHEET_REG  = 'Registros';
const SHEET_FER  = 'Ferias';

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

function getCachedData(cacheKey, sheetName) {
  const cache  = CacheService.getScriptCache();
  const cached = cache.get(cacheKey);
  if (cached) {
    try { return JSON.parse(cached); } catch (e) {}
  }
  const data = getDataAsObjects(sheetName);
  try {
    const serialized = JSON.stringify(data);
    if (serialized.length < 100000) cache.put(cacheKey, serialized, CACHE_TTL);
  } catch (e) {}
  return data;
}

function invalidateCache(cacheKey) {
  try { CacheService.getScriptCache().remove(cacheKey); } catch (e) {}
}

// ── NORMALIZAÇÕES ─────────────────────────────────────────────────────────────

function normalizeMatricula(v) {
  return String(v || '').trim().replace(/\D/g, '');
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

    // ── registros ─────────────────────────────────────────────────────────────
    if (action === 'registros') {
      const qMat  = normalizeMatricula(e.parameter.matricula || '');
      const qNome = String(e.parameter.nome   || '').trim().toLowerCase();
      const qSem  = normalizeSemanaISO(e.parameter.semana || '');

      const regs = getCachedData('cache_registros', SHEET_REG).filter(r => {
        const m = normalizeMatricula(r['Matricula']);
        const n = String(r['Nome']     || '').toLowerCase();
        const s = normalizeSemanaISO(r['SemanaISO']);
        return (!qMat  || m.includes(qMat))
            && (!qNome || n.includes(qNome))
            && (!qSem  || s === qSem);
      });
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

    return respond({
      ok:  true,
      msg: 'DSS GIG API — use ?action=[funcionario|funcionarios|treinamentos|registros|ferias]'
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
      let { matricula, semanaISO, tituloVideo, urlVideo, assinaturaPNG, deviceInfo } = body || {};
      matricula = normalizeMatricula(matricula);
      semanaISO = normalizeSemanaISO(semanaISO);

      if (!matricula || !semanaISO || !tituloVideo || !urlVideo || !assinaturaPNG) {
        return respond({ ok: false, error: 'Campos obrigatórios: matricula, semanaISO, tituloVideo, urlVideo, assinaturaPNG' });
      }
      const func = findFuncionarioByMatricula(matricula);
      if (!func) return respond({ ok: false, error: 'Funcionário não encontrado ou inativo.' });

      appendRow(SHEET_REG, {
        'Timestamp'    : new Date(),
        'Matricula'    : matricula,
        'Nome'         : func['Nome'],
        'Setor'        : func['Setor'],
        'SemanaISO'    : semanaISO,
        'TituloVideo'  : tituloVideo,
        'URLVideo'     : urlVideo,
        'AssinaturaPNG': assinaturaPNG,
        'DeviceInfo'   : deviceInfo || ''
      });
      invalidateCache('cache_registros');
      return respond({ ok: true, message: 'Registro salvo.' });
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

function addFerias_(payload) {
  const matricula   = normalizeMatricula(payload.matricula);
  const funcionario = String(payload.funcionario  || '').trim();
  const situacao    = String(payload.situacao     || 'Férias').trim();
  const inicioFerias = String(payload.inicioFerias || '').trim();
  const fimFerias    = String(payload.fimFerias    || '').trim();

  if (!matricula || !funcionario) return respond({ ok: false, error: 'Matrícula e Funcionario são obrigatórios.' });

  const isFerias = situacao.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').includes('ferias');
  if (isFerias && (!inicioFerias || !fimFerias)) return respond({ ok: false, error: 'Férias exigem InicioFerias e FimFerias.' });

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
    return respond({ ok: true, message: 'Registro de férias/ausência inserido.' });
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
