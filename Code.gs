/* =========================================================================
   DIÁRIO DE BORDO — BACKEND (Google Apps Script)
   Cole este código em Extensões > Apps Script da planilha "Diário de Bordo".
   Depois rode a função setupSheet() UMA VEZ para criar as abas e popular
   a biblioteca padrão de tarefas.
   ========================================================================= */

const SHEET_BIBLIOTECA = 'Biblioteca';
const SHEET_LOG = 'Log';
const SHEET_METRICAS = 'Metricas';
const SHEET_CONFIG = 'Config';

/* ---------- ENDPOINTS ---------- */

function doGet(e) {
  return respond({
    library: getLibrary(),
    log: getLog(),
    metrics: getMetrics(),
    config: getConfig(),
  });
}

function doPost(e) {
  const body = JSON.parse(e.postData.contents);
  let result;
  switch (body.action) {
    case 'toggleTask': result = toggleTask(body.date, body.taskId); break;
    case 'addTask': result = addTask(body.task); break;
    case 'removeTask': result = removeTask(body.id); break;
    case 'updateMetric': result = updateMetric(body.week, body.field, body.value); break;
    case 'setConfig': result = setConfig(body.key, body.value); break;
    default: result = { error: 'ação desconhecida: ' + body.action };
  }
  return respond(result);
}

function respond(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- HELPERS DE PLANILHA ---------- */

function sh(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

function readSheet(name) {
  const s = sh(name);
  const data = s.getDataRange().getValues();
  const headers = data.shift();
  return data
    .filter(r => r.join('') !== '')
    .map(r => {
      const obj = {};
      headers.forEach((h, i) => (obj[h] = r[i]));
      return obj;
    });
}

function findRow(name, matchFn) {
  const s = sh(name);
  const data = s.getDataRange().getValues();
  const headers = data[0];
  for (let i = 1; i < data.length; i++) {
    const obj = {};
    headers.forEach((h, c) => (obj[h] = data[i][c]));
    if (matchFn(obj)) return { rowIndex: i + 1, headers, obj };
  }
  return null;
}

/* ---------- BIBLIOTECA ---------- */

function getLibrary() {
  return readSheet(SHEET_BIBLIOTECA).filter(t => t.ativo !== false);
}

function addTask(task) {
  const s = sh(SHEET_BIBLIOTECA);
  const id = 'c' + new Date().getTime();
  s.appendRow([id, task.day, task.pillar, task.portfolio, task.title, task.desc, true]);
  return { ok: true, id };
}

function removeTask(id) {
  const found = findRow(SHEET_BIBLIOTECA, o => o.id === id);
  if (!found) return { error: 'tarefa não encontrada' };
  const s = sh(SHEET_BIBLIOTECA);
  const ativoCol = found.headers.indexOf('ativo') + 1;
  s.getRange(found.rowIndex, ativoCol).setValue(false);
  return { ok: true };
}

/* ---------- LOG DIÁRIO ---------- */

function getLog() {
  return readSheet(SHEET_LOG);
}

function toggleTask(date, taskId) {
  const s = sh(SHEET_LOG);
  const found = findRow(SHEET_LOG, o => o.date === date && o.taskId === taskId);
  if (found) {
    const doneCol = found.headers.indexOf('done') + 1;
    const newVal = !found.obj.done;
    s.getRange(found.rowIndex, doneCol).setValue(newVal);
    return { ok: true, done: newVal };
  }
  s.appendRow([date, taskId, true, new Date()]);
  return { ok: true, done: true };
}

/* ---------- MÉTRICAS SEMANAIS ---------- */

function getMetrics() {
  return readSheet(SHEET_METRICAS);
}

function updateMetric(week, field, value) {
  const s = sh(SHEET_METRICAS);
  const found = findRow(SHEET_METRICAS, o => o.week === week);
  if (found) {
    const col = found.headers.indexOf(field) + 1;
    s.getRange(found.rowIndex, col).setValue(value);
    return { ok: true };
  }
  const headers = s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0];
  const row = headers.map(h => (h === 'week' ? week : h === field ? value : ''));
  s.appendRow(row);
  return { ok: true };
}

/* ---------- CONFIG (chave/valor) ---------- */

function getConfig() {
  const rows = readSheet(SHEET_CONFIG);
  const obj = {};
  rows.forEach(r => (obj[r.key] = r.value));
  return obj;
}

function setConfig(key, value) {
  const found = findRow(SHEET_CONFIG, o => o.key === key);
  const s = sh(SHEET_CONFIG);
  if (found) {
    s.getRange(found.rowIndex, 2).setValue(value);
  } else {
    s.appendRow([key, value]);
  }
  return { ok: true };
}

/* =========================================================================
   SETUP — rode esta função UMA VEZ (menu Executar > setupSheet) para criar
   as abas do zero e semear a biblioteca padrão de tarefas.
   ========================================================================= */

function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const bib = ss.getSheetByName(SHEET_BIBLIOTECA) || ss.insertSheet(SHEET_BIBLIOTECA);
  bib.clear();
  bib.appendRow(['id', 'day', 'pillar', 'portfolio', 'title', 'desc', 'ativo']);
  DEFAULT_LIBRARY.forEach(t => bib.appendRow([t.id, t.day, t.pillar, t.portfolio, t.title, t.desc, true]));
  bib.setFrozenRows(1);

  const log = ss.getSheetByName(SHEET_LOG) || ss.insertSheet(SHEET_LOG);
  log.clear();
  log.appendRow(['date', 'taskId', 'done', 'timestamp']);
  log.setFrozenRows(1);

  const met = ss.getSheetByName(SHEET_METRICAS) || ss.insertSheet(SHEET_METRICAS);
  met.clear();
  met.appendRow(['week', 'visitas', 'propostas', 'fechamentos', 'faturamento', 'obs']);
  met.setFrozenRows(1);

  const cfg = ss.getSheetByName(SHEET_CONFIG) || ss.insertSheet(SHEET_CONFIG);
  cfg.clear();
  cfg.appendRow(['key', 'value']);
  cfg.appendRow(['startDate', Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd')]);
  cfg.setFrozenRows(1);

  // remove a aba padrão "Página1"/"Sheet1" se ainda existir e estiver vazia
  const def = ss.getSheetByName('Sheet1') || ss.getSheetByName('Página1');
  if (def && ss.getSheets().length > 1) ss.deleteSheet(def);

  SpreadsheetApp.getUi().alert('Planilha configurada! Agora vá em Implantar > Nova implantação para publicar a API.');
}

const DEFAULT_LIBRARY = [
  { id: 't1', day: 1, pillar: 'gestao', portfolio: 'ambos', title: 'Revisar funil da semana', desc: 'Abra o CRM, revise as oportunidades em aberto e defina a meta de visitas e propostas da semana.' },
  { id: 't2', day: 1, pillar: 'prospeccao', portfolio: 'bondmann', title: 'Prospecção Bondmann', desc: 'Liste 10 novos prospects (CNAE metal-mecânica) da base BigQuery e envie o primeiro contato.' },
  { id: 't3', day: 1, pillar: 'prospeccao', portfolio: 'bransales', title: 'Prospecção Bransales', desc: 'Liste 10 novos prospects (frotistas/transportadoras) e envie o primeiro contato.' },
  { id: 't4', day: 1, pillar: 'marca', portfolio: 'pessoal', title: 'Planejar conteúdo da semana', desc: 'Defina os temas dos posts/stories da semana para @rdpne_us e o site.' },
  { id: 't5', day: 2, pillar: 'tecnico', portfolio: 'bondmann', title: 'Ficha técnica Bondmann', desc: 'Estude 1 produto do catálogo: aplicação, diferencial técnico e argumento de valor.' },
  { id: 't6', day: 2, pillar: 'negociacao', portfolio: 'bondmann', title: 'Treino de objeção (SPIN)', desc: 'Escolha 1 objeção comum de Bondmann e escreva a resposta usando SPIN Selling.' },
  { id: 't7', day: 2, pillar: 'prospeccao', portfolio: 'bondmann', title: 'Follow-up Bondmann', desc: 'Retome contato com leads Bondmann parados há mais de 3 dias.' },
  { id: 't8', day: 2, pillar: 'gestao', portfolio: 'ambos', title: 'Atualizar CRM do dia', desc: 'Registre no CRM todas as visitas e ligações feitas hoje.' },
  { id: 't9', day: 3, pillar: 'marca', portfolio: 'pessoal', title: 'Publicar conteúdo', desc: 'Publique 1 post ou stories com bastidor de trabalho ou dica técnica.' },
  { id: 't10', day: 3, pillar: 'prospeccao', portfolio: 'ambos', title: 'Follow-up de propostas', desc: 'Envie follow-up para propostas (Bondmann e Bransales) enviadas há 5+ dias sem retorno.' },
  { id: 't11', day: 3, pillar: 'marca', portfolio: 'pessoal', title: 'Interação estratégica', desc: 'Comente ou interaja com 5 perfis de clientes ou parceiros do setor.' },
  { id: 't12', day: 3, pillar: 'gestao', portfolio: 'ambos', title: 'Registrar visitas com GPS', desc: 'Confira se todas as visitas do dia foram registradas no app de prospecção.' },
  { id: 't13', day: 4, pillar: 'tecnico', portfolio: 'bransales', title: 'Linha de pneus Bransales', desc: 'Estude 1 linha de pneus: aplicação, durabilidade e argumento de custo por km rodado.' },
  { id: 't14', day: 4, pillar: 'negociacao', portfolio: 'bransales', title: 'Simulação Challenger Sale', desc: 'Simule um fechamento real usando a técnica Challenger Sale para um cliente do funil.' },
  { id: 't15', day: 4, pillar: 'prospeccao', portfolio: 'bransales', title: 'Follow-up Bransales', desc: 'Retome contato com leads Bransales parados há mais de 3 dias.' },
  { id: 't16', day: 4, pillar: 'gestao', portfolio: 'ambos', title: 'Atualizar CRM do dia', desc: 'Registre no CRM todas as visitas e ligações feitas hoje.' },
  { id: 't17', day: 5, pillar: 'negociacao', portfolio: 'ambos', title: 'Revisão de propostas em aberto', desc: 'Revise todas as propostas em aberto e defina a próxima ação de cada uma.' },
  { id: 't18', day: 5, pillar: 'negociacao', portfolio: 'ambos', title: 'Pós-mortem de negociação', desc: 'Escolha a negociação mais difícil da semana e escreva o que faria diferente.' },
  { id: 't19', day: 5, pillar: 'gestao', portfolio: 'ambos', title: 'Taxa de conversão da semana', desc: 'Calcule visitas → propostas → fechamentos e registre na aba Métricas.' },
  { id: 't20', day: 5, pillar: 'marca', portfolio: 'pessoal', title: 'Pedido de indicação', desc: 'Peça 1 indicação ou depoimento a um cliente satisfeito.' },
  { id: 't21', day: 6, pillar: 'gestao', portfolio: 'ambos', title: 'Revisão semanal completa', desc: 'Revise KPIs, funil e defina as metas da próxima semana.' },
  { id: 't22', day: 6, pillar: 'marca', portfolio: 'pessoal', title: 'Conteúdo de autoridade', desc: 'Planeje ou grave um vídeo curto, artigo ou case para publicar na semana seguinte.' },
  { id: 't23', day: 6, pillar: 'tecnico', portfolio: 'ambos', title: 'Leitura técnica/setorial', desc: 'Leia 1 material sobre o setor metal-mecânico ou de transporte/logística.' },
  { id: 't24', day: 0, pillar: 'marca', portfolio: 'pessoal', title: 'Planejamento leve', desc: 'Revise a semana e ajuste a agenda da próxima. Opcional, sem pressão.' },
];
