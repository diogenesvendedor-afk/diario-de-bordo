/* =========================================================================
   DIÁRIO DE BORDO — BACKEND (Google Apps Script) — v2
   Cole este código em Extensões > Apps Script da planilha "Diário de Bordo".
   Depois rode a função setupSheet() UMA VEZ (ou de novo, é seguro repetir
   nas abas Biblioteca/Config/Conhecimento — NÃO apaga Visitas/Clientes/Log
   se já existirem dados).
   ========================================================================= */

const SHEET_BIBLIOTECA = 'Biblioteca';
const SHEET_LOG = 'Log';
const SHEET_METRICAS = 'Metricas';
const SHEET_CONFIG = 'Config';
const SHEET_VISITAS = 'Visitas';
const SHEET_CLIENTES = 'Clientes';
const SHEET_CONHECIMENTO = 'Conhecimento';

/* ---------- ENDPOINTS ---------- */

function doGet(e) {
  return respond({
    library: getLibrary(),
    log: getLog(),
    metrics: getMetrics(),
    config: getConfig(),
    visits: getVisits(),
    clients: getClients(),
    knowledge: getKnowledge(),
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

    case 'addVisit': result = addVisit(body.visit); break;
    case 'updateVisit': result = updateVisit(body.id, body.field, body.value); break;
    case 'removeVisit': result = deleteRowById(SHEET_VISITAS, body.id); break;
    case 'importVisits': result = importVisits(body.rows); break;

    case 'addClient': result = addClient(body.client); break;
    case 'updateClient': result = updateClient(body.id, body.field, body.value); break;
    case 'removeClient': result = deleteRowById(SHEET_CLIENTES, body.id); break;

    case 'suggestStrategy': result = suggestStrategy(body.clientId); break;

    default: result = { error: 'ação desconhecida: ' + body.action };
  }
  return respond(result);
}

function respond(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- HELPERS GENÉRICOS DE PLANILHA ---------- */

function sh(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

function readSheet(name) {
  const s = sh(name);
  if (!s) return [];
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

function deleteRowById(sheetName, id) {
  const found = findRow(sheetName, o => o.id === id);
  if (!found) return { error: 'registro não encontrado' };
  sh(sheetName).deleteRow(found.rowIndex);
  return { ok: true };
}

function genId(prefix) {
  return prefix + new Date().getTime() + Math.floor(Math.random() * 1000);
}

/* ---------- BIBLIOTECA (tarefas recorrentes por dia da semana) ---------- */

function getLibrary() {
  return readSheet(SHEET_BIBLIOTECA).filter(t => t.ativo !== false);
}

function addTask(task) {
  const id = genId('c');
  sh(SHEET_BIBLIOTECA).appendRow([id, task.day, task.pillar, task.portfolio, task.title, task.desc, true]);
  return { ok: true, id };
}

function removeTask(id) {
  const found = findRow(SHEET_BIBLIOTECA, o => o.id === id);
  if (!found) return { error: 'tarefa não encontrada' };
  const ativoCol = found.headers.indexOf('ativo') + 1;
  sh(SHEET_BIBLIOTECA).getRange(found.rowIndex, ativoCol).setValue(false);
  return { ok: true };
}

/* ---------- LOG DIÁRIO ---------- */

function getLog() {
  return readSheet(SHEET_LOG);
}

function toggleTask(date, taskId) {
  const found = findRow(SHEET_LOG, o => o.date === date && o.taskId === taskId);
  if (found) {
    const doneCol = found.headers.indexOf('done') + 1;
    const newVal = !found.obj.done;
    sh(SHEET_LOG).getRange(found.rowIndex, doneCol).setValue(newVal);
    return { ok: true, done: newVal };
  }
  sh(SHEET_LOG).appendRow([date, taskId, true, new Date()]);
  return { ok: true, done: true };
}

/* ---------- MÉTRICAS SEMANAIS ---------- */

function getMetrics() {
  return readSheet(SHEET_METRICAS);
}

function updateMetric(week, field, value) {
  const found = findRow(SHEET_METRICAS, o => o.week === week);
  const s = sh(SHEET_METRICAS);
  if (found) {
    const col = found.headers.indexOf(field) + 1;
    s.getRange(found.rowIndex, col).setValue(value);
    return { ok: true };
  }
  const headers = s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0];
  s.appendRow(headers.map(h => (h === 'week' ? week : h === field ? value : '')));
  return { ok: true };
}

/* ---------- CONFIG (chave/valor) ---------- */

function getConfig() {
  const obj = {};
  readSheet(SHEET_CONFIG).forEach(r => (obj[r.key] = r.value));
  return obj;
}

function setConfig(key, value) {
  const found = findRow(SHEET_CONFIG, o => o.key === key);
  const s = sh(SHEET_CONFIG);
  if (found) s.getRange(found.rowIndex, 2).setValue(value);
  else s.appendRow([key, value]);
  return { ok: true };
}

/* ---------- AGENDA DE VISITAS ---------- */

function getVisits() {
  return readSheet(SHEET_VISITAS);
}

function addVisit(v) {
  const id = genId('v');
  sh(SHEET_VISITAS).appendRow([id, v.date, v.cliente, v.empresa, v.portfolio || 'ambos', v.endereco || '', v.objetivo || '', v.status || 'Pendente', v.notas || '']);
  return { ok: true, id };
}

function updateVisit(id, field, value) {
  const found = findRow(SHEET_VISITAS, o => o.id === id);
  if (!found) return { error: 'visita não encontrada' };
  const col = found.headers.indexOf(field) + 1;
  if (col <= 0) return { error: 'campo inválido' };
  sh(SHEET_VISITAS).getRange(found.rowIndex, col).setValue(value);
  return { ok: true };
}

// Importação em massa — usado pela tela "Importar visitas" do app,
// onde o Diogenes cola linhas copiadas da planilha de visitas dele.
function importVisits(rows) {
  const s = sh(SHEET_VISITAS);
  let count = 0;
  rows.forEach(r => {
    if (!r.cliente && !r.empresa) return;
    const id = genId('v');
    s.appendRow([id, r.date || '', r.cliente || '', r.empresa || '', r.portfolio || 'ambos', r.endereco || '', r.objetivo || '', 'Pendente', r.notas || '']);
    count++;
  });
  return { ok: true, imported: count };
}

/* ---------- CLIENTES ---------- */

function getClients() {
  return readSheet(SHEET_CLIENTES);
}

function addClient(c) {
  const id = genId('cl');
  sh(SHEET_CLIENTES).appendRow([id, c.empresa, c.cnpj || '', c.cnae || '', c.segmento || '', c.portfolio || 'ambos', c.contato || '', c.telefone || '', c.cidade || '', c.potencial || 'Médio', c.notas || '']);
  return { ok: true, id };
}

function updateClient(id, field, value) {
  const found = findRow(SHEET_CLIENTES, o => o.id === id);
  if (!found) return { error: 'cliente não encontrado' };
  const col = found.headers.indexOf(field) + 1;
  if (col <= 0) return { error: 'campo inválido' };
  sh(SHEET_CLIENTES).getRange(found.rowIndex, col).setValue(value);
  return { ok: true };
}

/* ---------- CONHECIMENTO (tópicos validados / guias) ---------- */

function getKnowledge() {
  return readSheet(SHEET_CONHECIMENTO);
}

/* ---------- SUGESTÃO DE ESTRATÉGIA ----------
   Motor simples baseado em regras: cruza segmento/CNAE/portfólio do cliente
   com as tags dos guias da aba Conhecimento. Sem chamadas externas — roda
   inteiramente dentro da planilha, então não depende de nenhuma chave de API. */

function suggestStrategy(clientId) {
  const client = readSheet(SHEET_CLIENTES).find(c => c.id === clientId);
  if (!client) return { error: 'cliente não encontrado' };

  const kb = readSheet(SHEET_CONHECIMENTO);
  const segText = ((client.segmento || '') + ' ' + (client.cnae || '')).toLowerCase();
  const segWords = segText.split(/[^a-zà-ú0-9]+/).filter(w => w.length > 2);

  const scored = kb.map(k => {
    let score = 0;
    if (k.portfolio === client.portfolio || k.portfolio === 'ambos') score += 2;
    const tags = (k.tags || '').toLowerCase().split(',').map(t => t.trim()).filter(Boolean);
    segWords.forEach(w => { if (tags.some(t => t.includes(w) || w.includes(t))) score += 1; });
    return Object.assign({}, k, { score });
  });

  let suggestions = scored.filter(k => k.score > 0).sort((a, b) => b.score - a.score).slice(0, 5);
  if (!suggestions.length) {
    suggestions = scored.filter(k => k.portfolio === client.portfolio || k.portfolio === 'ambos').slice(0, 4);
  }

  return { client, suggestions };
}

/* =========================================================================
   SETUP — rode esta função para criar as abas que ainda não existem e
   (re)semear Biblioteca e Conhecimento. Visitas, Clientes e Log NÃO são
   apagados se já existirem — é seguro rodar de novo a qualquer momento.
   ========================================================================= */

function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const bib = ss.getSheetByName(SHEET_BIBLIOTECA) || ss.insertSheet(SHEET_BIBLIOTECA);
  bib.clear();
  bib.appendRow(['id', 'day', 'pillar', 'portfolio', 'title', 'desc', 'ativo']);
  DEFAULT_LIBRARY.forEach(t => bib.appendRow([t.id, t.day, t.pillar, t.portfolio, t.title, t.desc, true]));
  bib.setFrozenRows(1);

  const log = ss.getSheetByName(SHEET_LOG) || ss.insertSheet(SHEET_LOG);
  if (log.getLastRow() === 0) { log.appendRow(['date', 'taskId', 'done', 'timestamp']); log.setFrozenRows(1); }

  const met = ss.getSheetByName(SHEET_METRICAS) || ss.insertSheet(SHEET_METRICAS);
  if (met.getLastRow() === 0) { met.appendRow(['week', 'visitas', 'propostas', 'fechamentos', 'faturamento', 'obs']); met.setFrozenRows(1); }

  const cfg = ss.getSheetByName(SHEET_CONFIG) || ss.insertSheet(SHEET_CONFIG);
  if (cfg.getLastRow() === 0) {
    cfg.appendRow(['key', 'value']);
    cfg.appendRow(['startDate', Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd')]);
    cfg.setFrozenRows(1);
  }

  const vis = ss.getSheetByName(SHEET_VISITAS) || ss.insertSheet(SHEET_VISITAS);
  if (vis.getLastRow() === 0) {
    vis.appendRow(['id', 'date', 'cliente', 'empresa', 'portfolio', 'endereco', 'objetivo', 'status', 'notas']);
    vis.setFrozenRows(1);
  }

  const cli = ss.getSheetByName(SHEET_CLIENTES) || ss.insertSheet(SHEET_CLIENTES);
  if (cli.getLastRow() === 0) {
    cli.appendRow(['id', 'empresa', 'cnpj', 'cnae', 'segmento', 'portfolio', 'contato', 'telefone', 'cidade', 'potencial', 'notas']);
    cli.setFrozenRows(1);
  }

  const con = ss.getSheetByName(SHEET_CONHECIMENTO) || ss.insertSheet(SHEET_CONHECIMENTO);
  con.clear();
  con.appendRow(['id', 'pillar', 'portfolio', 'title', 'content', 'tags']);
  DEFAULT_KNOWLEDGE.forEach(k => con.appendRow([k.id, k.pillar, k.portfolio, k.title, k.content, k.tags]));
  con.setFrozenRows(1);

  const def = ss.getSheetByName('Sheet1') || ss.getSheetByName('Página1');
  if (def && ss.getSheets().length > 1) ss.deleteSheet(def);

  Logger.log('Planilha configurada! Agora vá em Implantar > Gerenciar implantações > editar > Nova versão.');
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

const DEFAULT_KNOWLEDGE = [
  { id: 'k1', pillar: 'negociacao', portfolio: 'bondmann', title: 'SPIN Selling aplicado a produtos químicos industriais',
    content: 'Situação: entenda o processo produtivo atual (qual produto usam hoje, frequência de troca). Problema: pergunte sobre falhas, retrabalho, paradas de máquina, não conformidade. Implicação: quantifique o custo dessas falhas (horas paradas, refugo, garantia). Necessidade de solução: leve o cliente a verbalizar o ganho de resolver isso — só depois apresente o produto Bondmann como resposta a essa necessidade que ele mesmo descreveu.',
    tags: 'metal-mecanica,usinagem,serralheria,industria,quimica' },
  { id: 'k2', pillar: 'negociacao', portfolio: 'bransales', title: 'Challenger Sale para frotistas e transportadoras',
    content: 'Ensine algo que o gestor de frota não sabia (ex: como o custo por km rodado muda com a durabilidade do pneu, não só o preço de compra). Adapte o discurso ao que mais importa para aquele cliente (uptime, segurança, custo). Assuma o controle da conversa levando a discussão para o TCO da frota, não para desconto no pedido.',
    tags: 'frota,transportadora,logistica,transporte' },
  { id: 'k3', pillar: 'marca', portfolio: 'pessoal', title: 'Checklist de marca pessoal para ser procurado',
    content: '1) Perfil do LinkedIn com foto profissional, headline clara com os dois portfólios. 2) Postar 1x/semana conteúdo técnico de verdade, não genérico. 3) Responder rápido (menos de 2h em horário comercial). 4) Pedir depoimento a cada fechamento importante. 5) Manter o site/portfólio sempre atualizado com cases reais.',
    tags: 'geral' },
  { id: 'k4', pillar: 'tecnico', portfolio: 'bondmann', title: 'Como argumentar valor técnico Bondmann',
    content: 'Foque em custo total, não em preço do litro/kg: menos paradas de manutenção, menos refugo, conformidade com normas do setor. Traga um exemplo numérico sempre que possível — clientes técnicos respondem melhor a números do que a adjetivos.',
    tags: 'metal-mecanica,quimica,industria' },
  { id: 'k5', pillar: 'tecnico', portfolio: 'bransales', title: 'Como argumentar valor técnico Bransales',
    content: 'Traduza tudo para custo por km rodado e não para preço do pneu. Fale de uptime da frota, menos trocas por ano, menos veículos parados na oficina — isso importa mais para o gestor do que o valor de compra isolado.',
    tags: 'frota,transportadora,logistica' },
  { id: 'k6', pillar: 'prospeccao', portfolio: 'ambos', title: 'Roteiro de primeiro contato (cold approach)',
    content: 'Abra se identificando e citando algo específico da empresa (não genérico). Faça 1 pergunta de diagnóstico real (ex: "hoje vocês trabalham com qual fornecedor de X? Como está a relação custo x desempenho?"). Não tente vender no primeiro contato — o objetivo é agendar uma conversa mais longa.',
    tags: 'geral' },
  { id: 'k7', pillar: 'negociacao', portfolio: 'ambos', title: 'Como responder objeção de preço',
    content: 'Não entre em disputa de desconto direto. Reancore a conversa no custo total (falhas evitadas, tempo economizado, durabilidade). Pergunte "comparado com o quê?" para entender contra o que o cliente está comparando, e só então reposicione o valor.',
    tags: 'geral' },
  { id: 'k8', pillar: 'gestao', portfolio: 'ambos', title: 'Como priorizar o funil da semana',
    content: 'Classifique cada oportunidade em dois eixos: potencial de valor (alto/médio/baixo) e proximidade do fechamento (quente/morno/frio). Priorize sempre alto potencial + quente primeiro, depois alto potencial + morno. Não gaste a semana toda em oportunidades frias de baixo potencial.',
    tags: 'geral' },
  { id: 'k9', pillar: 'marca', portfolio: 'pessoal', title: 'Pauta de conteúdo semanal para LinkedIn/Instagram',
    content: 'Alterne 4 tipos de post ao longo do mês: bastidor de visita/trabalho, dica técnica curta (Bondmann ou Bransales), case real de cliente (com autorização), opinião/posicionamento sobre o setor. Isso evita ficar repetitivo e constrói autoridade em várias frentes.',
    tags: 'geral' },
  { id: 'k10', pillar: 'prospeccao', portfolio: 'bondmann', title: 'Como identificar bons prospects Bondmann via CNAE',
    content: 'Priorize CNAEs de metalurgia, usinagem, serralheria e metal-mecânica em geral. Sinais de bom potencial: porte médio/grande, atividade fabril contínua (não só comércio), histórico de mais de 2 anos de operação.',
    tags: 'metal-mecanica,cnae,usinagem,serralheria' },
  { id: 'k11', pillar: 'prospeccao', portfolio: 'bransales', title: 'Como identificar bons prospects Bransales via CNAE',
    content: 'Priorize CNAEs de transporte rodoviário de carga, logística e locação de veículos. Sinais de bom potencial: frota própria (não só terceirizada), operação regional/nacional, crescimento recente de contratos.',
    tags: 'transporte,logistica,cnae,frota' },
  { id: 'k12', pillar: 'negociacao', portfolio: 'pessoal', title: 'Como pedir indicação sem parecer forçado',
    content: 'O melhor momento é logo após uma entrega bem-sucedida ou um elogio espontâneo do cliente. Peça de forma específica: não "você conhece alguém?", mas "tem algum outro gestor do setor que você acha que eu deveria conhecer?".',
    tags: 'geral' },
];
