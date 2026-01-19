/**
 * DSS GIG — API (Apps Script) para Treinamentos Semanais de Segurança
 * Abas esperadas na planilha:
 *   Funcionarios(Matricula, Nome, Setor, Ativo)
 *   Treinamentos(SemanaISO, Titulo, URL, Ativo)
 *   Registros(Timestamp, Matricula, Nome, Setor, SemanaISO, TituloVideo, URLVideo, AssinaturaPNG, DeviceInfo)
 */

const SHEET_FUNC = 'Funcionarios';
const SHEET_TREI = 'Treinamentos';
const SHEET_REG  = 'Registros';

function respond(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function getSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error('Aba não encontrada: ' + name);
  return sh;
}

function getDataAsObjects(sheetName) {
  const sh = getSheet(sheetName);
  const rng = sh.getDataRange().getValues();
  if (rng.length < 2) return [];
  const headers = rng[0];
  return rng.slice(1).map(row => {
    const obj = {}; headers.forEach((h, i) => obj[String(h).trim()] = row[i]); return obj;
  });
}

function appendRow(sheetName, obj) {
  const sh = getSheet(sheetName);
  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const row = headers.map(h => obj.hasOwnProperty(h) ? obj[h] : '');
  sh.appendRow(row);
}

function computeRecentWeeks(treinamentos) {
  const active = treinamentos.filter(t => String(t['Ativo']).toLowerCase() === 'true' || t['Ativo'] === true);
  const sorted = active.sort((a,b) => String(b['SemanaISO']).localeCompare(String(a['SemanaISO'])));
  return sorted.slice(0,3);
}

function findFuncionarioByMatricula(matricula) {
  const list = getDataAsObjects(SHEET_FUNC);
  return list.find(f => String(f['Matricula']).trim() === String(matricula).trim() && (String(f['Ativo']).toLowerCase() === 'true' || f['Ativo'] === true));
}

function doGet(e) {
  try {
    const action = (e.parameter.action || '').toLowerCase();

    if (action === 'funcionario') {
      const matricula = e.parameter.matricula;
      if (!matricula) return respond({ ok:false, error:'Informe ?matricula=' });
      const func = findFuncionarioByMatricula(matricula);
      if (!func) return respond({ ok:true, found:false });
      return respond({ ok:true, found:true, data:func });
    }

    if (action === 'treinamentos') {
      const treinamentos = getDataAsObjects(SHEET_TREI);
      const recent = computeRecentWeeks(treinamentos);
      return respond({ ok:true, data:recent });
    }

    if (action === 'registros') {
      const qMat = (e.parameter.matricula || '').trim().toLowerCase();
      const qNome = (e.parameter.nome || '').trim().toLowerCase();
      const qSem  = (e.parameter.semana || '').trim().toLowerCase();
      const regs = getDataAsObjects(SHEET_REG).filter(r => {
        const m = String(r['Matricula']||'').toLowerCase();
        const n = String(r['Nome']||'').toLowerCase();
        const s = String(r['SemanaISO']||'').toLowerCase();
        return (!qMat || m.includes(qMat)) && (!qNome || n.includes(qNome)) && (!qSem || s === qSem);
      });
      return respond({ ok:true, data:regs });
    }

    return respond({ ok:true, msg:'DSS GIG API — use ?action=[funcionario|treinamentos|registros]' });
  } catch (err) {
    return respond({ ok:false, error:String(err) });
  }
}

function doPost(e) {
  try {
    const action = (e.parameter.action || '').toLowerCase();
    const body = (e.postData && e.postData.contents) ? JSON.parse(e.postData.contents) : {};

    if (action === 'registrar') {
      const { matricula, semanaISO, tituloVideo, urlVideo, assinaturaPNG, deviceInfo } = body || {};
      if (!matricula || !semanaISO || !tituloVideo || !urlVideo || !assinaturaPNG) {
        return respond({ ok:false, error:'Campos obrigatórios: matricula, semanaISO, tituloVideo, urlVideo, assinaturaPNG' });
      }

      const func = findFuncionarioByMatricula(matricula);
      if (!func) return respond({ ok:false, error:'Funcionário não encontrado ou inativo.' });

      const rowObj = {
        'Timestamp': new Date(),
        'Matricula': matricula,
        'Nome': func['Nome'],
        'Setor': func['Setor'],
        'SemanaISO': semanaISO,
        'TituloVideo': tituloVideo,
        'URLVideo': urlVideo,
        'AssinaturaPNG': assinaturaPNG,
        'DeviceInfo': deviceInfo || ''
      };
      appendRow(SHEET_REG, rowObj);
      return respond({ ok:true, message:'Registro salvo.' });
    }

    return respond({ ok:false, error:'Defina ?action=registrar' });
  } catch (err) {
    return respond({ ok:false, error:String(err) });
  }
}
