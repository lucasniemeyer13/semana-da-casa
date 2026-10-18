/**
 * Semana da Casa — a planilha é a fonte da verdade da página
 * https://lucasniemeyer13.github.io/semana-da-casa/
 *
 * A página grava aqui a cada alteração e lê daqui ao abrir e a cada 45 s. A leitura monta o estado
 * a partir das abas Tarefas, Eventos e Histórico, então o que for editado direto na planilha também
 * aparece na página (marcar "Lucas fez", trocar o responsável, renomear, mudar dia ou incluir linha).
 *
 * Como instalar (uma vez):
 *   1. Na planilha: Extensões > Apps Script, cole este arquivo e salve.
 *   2. Rode a função `configurar` (cria as abas e libera a planilha para leitura por qualquer pessoa com o link).
 *   3. Configurações do projeto > Propriedades do script: crie SENHA com a senha da casa.
 *   4. Implantar > Nova implantação > App da Web: Executar como "Eu", Quem pode acessar "Qualquer pessoa".
 *      Depois de mudar o código: Implantar > Gerenciar implantações > editar > Nova versão.
 *   5. Na página, cada navegador digita a senha uma vez.
 *
 * Toda leitura e gravação chega por POST com a senha no corpo; sem ela, nada é lido nem escrito.
 */

const ABAS = {
  tarefas: { nome: 'Tarefas', cab: ['id', 'Dia', 'Período', 'Ordem', 'Tarefa', 'Responsável', 'Lucas fez', 'Ana Júlia fez', 'Feita'] },
  eventos: { nome: 'Eventos', cab: ['id', 'Evento', 'De', 'Até', 'Quem'] },
  historico: { nome: 'Histórico', cab: ['Semana (início)', 'Tarefas', 'Feitas', '% geral', 'Lucas: atribuídas', 'Lucas: fez', 'Ana Júlia: atribuídas', 'Ana Júlia: fez'] },
  alteracoes: { nome: 'Alterações', cab: ['Salvo em', 'Semana (início)', 'Tarefas', 'Feitas', 'Lucas fez', 'Ana Júlia fez', 'Origem'] },
};
const DIAS = ['Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta', 'Sábado', 'Domingo'];
const PER = { m: 'Manhã', t: 'Tarde', n: 'Noite' };
const PESSOA = { L: 'Lucas', A: 'Ana Júlia', LA: 'Os dois' };
const MAX_LOG = 5000;      // linhas guardadas na aba Alterações

function configurar() {
  const ss = SpreadsheetApp.getActive();
  Object.keys(ABAS).forEach(k => aba_(ss, k));
  const nomes = Object.values(ABAS).map(a => a.nome);
  ss.getSheets().forEach(sh => {
    if (!nomes.includes(sh.getName()) && sh.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(sh);
  });
  DriveApp.getFileById(ss.getId()).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
}

function doGet() {
  return json_({ ok: false, erro: 'Use a página Semana da Casa.' });
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, erro: 'Envio inválido.' }); }
  if (!senhaOk_(body)) return json_({ ok: false, erro: 'senha' });
  try {
    if (body.acao === 'ler') {
      const estado = lerEstado_();
      return json_({ ok: true, estado: estado, versao: versao_(estado) });
    }
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      gravar_(body.estado, body.origem);
      const estado = lerEstado_();
      return json_({ ok: true, salvoEm: estado.salvoEm, versao: versao_(estado) });
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    return json_({ ok: false, erro: String(err) });
  }
}

function senhaOk_(body) {
  const senha = PropertiesService.getScriptProperties().getProperty('SENHA');
  return !!senha && !!body && typeof body.senha === 'string' && body.senha === senha;
}

// ---------- gravar: estado da página -> abas ----------
function gravar_(st, origem) {
  if (!st || typeof st.tasks !== 'object') throw new Error('Estado ausente no envio.');
  const ss = SpreadsheetApp.getActive();
  const agora = new Date();

  const tasks = Object.keys(st.tasks).map(id => Object.assign({ id: id }, st.tasks[id]))
    .sort((a, b) => a.dia - b.dia || 'mtn'.indexOf(a.periodo) - 'mtn'.indexOf(b.periodo) || (a.ordem || 0) - (b.ordem || 0));
  const linhas = tasks.map(t => {
    const f = t.feito || {};
    return [t.id, DIAS[t.dia] || t.dia, PER[t.periodo] || t.periodo, t.ordem || 0, t.titulo || '', PESSOA[resp_(t)] || 'Sem responsável', !!f.L, !!f.A, !!(f.L || f.A)];
  });
  const shT = aba_(ss, 'tarefas');
  escrever_(shT, ABAS.tarefas.cab.length, linhas);
  if (linhas.length) {
    shT.getRange(2, 7, linhas.length, 3).insertCheckboxes();
    shT.getRange(2, 9, linhas.length, 1).setFormulaR1C1('=OR(R[0]C[-2],R[0]C[-1])');
  }

  const bands = Object.keys(st.bands || {}).map(id => Object.assign({ id: id }, st.bands[id])).sort((a, b) => a.de - b.de);
  escrever_(aba_(ss, 'eventos'), ABAS.eventos.cab.length,
    bands.map(b => [b.id, b.texto || '', DIAS[b.de] || b.de, DIAS[b.ate] || b.ate, PESSOA[b.quem] || 'Os dois']));

  const hist = Object.keys(st.hist || {}).map(k => st.hist[k]).sort((a, b) => String(b.inicio).localeCompare(String(a.inicio)));
  const shH = aba_(ss, 'historico');
  escrever_(shH, ABAS.historico.cab.length, hist.map(h => [
    String(h.inicio), h.total || 0, h.feitas || 0, h.total ? h.feitas / h.total : 0,
    (h.L && h.L.atrib) || 0, (h.L && h.L.fez) || 0, (h.A && h.A.atrib) || 0, (h.A && h.A.fez) || 0,
  ]));
  if (hist.length) { shH.getRange(2, 1, hist.length, 1).setNumberFormat('@'); shH.getRange(2, 4, hist.length, 1).setNumberFormat('0%'); }

  PropertiesService.getScriptProperties().setProperty('META', JSON.stringify({ week: st.week || null, salvoEm: agora.toISOString() }));

  const feitas = tasks.filter(t => t.feito && (t.feito.L || t.feito.A)).length;
  const nL = tasks.filter(t => t.feito && t.feito.L).length;
  const nA = tasks.filter(t => t.feito && t.feito.A).length;
  const log = aba_(ss, 'alteracoes');
  log.appendRow([agora, (st.week && st.week.inicio) || '', tasks.length, feitas, nL, nA, String(origem || '').slice(0, 80)]);
  if (log.getLastRow() > MAX_LOG + 1) log.deleteRows(2, log.getLastRow() - 1 - MAX_LOG);
  SpreadsheetApp.flush();
}

// ---------- ler: abas -> estado da página ----------
function lerEstado_() {
  const ss = SpreadsheetApp.getActive();
  const meta = JSON.parse(PropertiesService.getScriptProperties().getProperty('META') || '{}');
  const estado = { tasks: {}, bands: {}, hist: {}, week: meta.week || null, salvoEm: meta.salvoEm || null };

  linhas_(ss, 'tarefas').forEach((r, i) => {
    const titulo = String(r[4] || '').trim();
    const dia = diaIdx_(r[1]);
    if (!titulo || dia < 0) return;
    const per = perId_(r[2]);
    const id = String(r[0] || '').trim() || ('linha-' + (i + 2));
    estado.tasks[id] = {
      titulo: titulo, dia: dia, periodo: per, ordem: Number(r[3]) || (i + 1) * 10,
      resp: respDeNome_(r[5]), feito: { L: r[6] === true, A: r[7] === true },
    };
  });

  linhas_(ss, 'eventos').forEach((r, i) => {
    const texto = String(r[1] || '').trim();
    const de = diaIdx_(r[2]), ate = diaIdx_(r[3]);
    if (!texto || de < 0) return;
    const q = respDeNome_(r[4]);
    estado.bands[String(r[0] || '').trim() || ('evento-' + (i + 2))] = {
      texto: texto, de: de, ate: ate < 0 ? de : ate, quem: q.length === 1 ? q[0] : 'LA',
    };
  });

  linhas_(ss, 'historico').forEach(r => {
    const inicio = r[0] instanceof Date ? Utilities.formatDate(r[0], ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd') : String(r[0] || '').trim();
    if (!inicio) return;
    estado.hist[inicio] = {
      inicio: inicio, total: Number(r[1]) || 0, feitas: Number(r[2]) || 0,
      L: { atrib: Number(r[4]) || 0, fez: Number(r[5]) || 0 }, A: { atrib: Number(r[6]) || 0, fez: Number(r[7]) || 0 },
    };
  });
  return estado;
}

// Versão = resumo do conteúdo; muda quando qualquer um grava pela página ou edita a planilha.
function versao_(estado) {
  const txt = JSON.stringify([estado.tasks, estado.bands, estado.hist, estado.week]);
  return Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, txt, Utilities.Charset.UTF_8));
}

// ---------- utilidades ----------
const sem_ = s => String(s || '').trim().toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '');
function diaIdx_(v) {
  if (typeof v === 'number') return v >= 0 && v < 7 ? v : -1;
  const s = sem_(v);
  return DIAS.findIndex(d => sem_(d) === s || sem_(d).slice(0, 3) === s.slice(0, 3) && s.length >= 3);
}
function perId_(v) {
  const s = sem_(v);
  if (s.startsWith('t')) return 't';
  if (s.startsWith('n')) return 'n';
  return 'm';
}
function respDeNome_(v) {
  const s = sem_(v);
  if (!s || s.startsWith('sem')) return [];
  if (s.includes('dois') || s.includes('ambos') || s === 'la') return ['L', 'A'];
  if (s.startsWith('l')) return ['L'];
  if (s.startsWith('a')) return ['A'];
  return [];
}
function resp_(t) {
  let r = Array.isArray(t.resp) ? t.resp : null;
  if (!r) { const q = Array.isArray(t.quem) ? t.quem : []; r = q.length === 1 ? q : []; }
  r = r.filter(x => x === 'L' || x === 'A');
  return r.length === 2 ? 'LA' : r.length === 1 ? r[0] : '';
}

function aba_(ss, k) {
  const a = ABAS[k];
  let sh = ss.getSheetByName(a.nome);
  if (!sh) {
    sh = ss.insertSheet(a.nome);
    sh.getRange(1, 1, 1, a.cab.length).setValues([a.cab]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function linhas_(ss, k) {
  const sh = ss.getSheetByName(ABAS[k].nome);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, ABAS[k].cab.length).getValues();
}

function escrever_(sh, ncol, linhas) {
  const ultima = sh.getLastRow();
  if (ultima > 1) sh.getRange(2, 1, ultima - 1, Math.max(ncol, sh.getLastColumn())).clear();
  if (linhas.length) sh.getRange(2, 1, linhas.length, ncol).setValues(linhas);
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
