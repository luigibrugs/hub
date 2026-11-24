
/**
 * HUB ALPHA
 * Integração Google Sheets + Kanban
 */

const ABA = 'Tarefas';

// Cole o ID da sua planilha Google Sheets.
const PLANILHA_ID = '1VGvYnKPllv3rHFhq0MzYFCdj0hJglW45pQ5mHUEceUo';

// Abre o HUB
function doGet() {
  return HtmlService
    .createHtmlOutputFromFile('index')
    .setTitle('ALPHA | Portal Corporativo');
}

// Acessa a aba Tarefas
function obterAba() {
  const planilha = SpreadsheetApp.openById(PLANILHA_ID);
  const aba = planilha.getSheetByName(ABA);

  if (!aba) {
    throw new Error('A aba Tarefas não foi encontrada.');
  }

  return aba;
}

// Busca tarefas e preenche campos automáticos


// Atualiza o status no Google Sheets


// Gera IDs e dados pendentes
function gerarIdsPendentes() {
  return listarTarefas();
}

// Testa a conexão com a planilha
function testarConexao() {
  const aba = obterAba();

  Logger.log('Conexão realizada com sucesso!');
  Logger.log('Aba: ' + aba.getName());
  Logger.log('Total de tarefas: ' + listarTarefas().length);

  return true;
}


function atualizarResponsavel(id, novoResponsavel) {
  const responsaveis = [
    'Luiz Carlos',
    'Luigi Trintoni',
    'Gabriela Feitosa',
    'Guilherme Crisostomo',
    'Ana Clara',
    'Luiza Santos',
    'Letícia de Souza',
    'Leonardo dos Santos',
    'Andrea Conti',
    'Tessie Simonato',
    'Edineis Colombo'
  ];

  if (!responsaveis.includes(novoResponsavel)) {
    throw new Error('Responsável inválido.');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const aba = obterAba();
    const ultima = aba.getLastRow();

    if (ultima < 2) {
      throw new Error('Tarefa não encontrada.');
    }

    const ids = aba.getRange(2, 1, ultima - 1, 1)
      .getValues();

    const indice = ids.findIndex(
      linha => String(linha[0]) === String(id)
    );

    if (indice === -1) {
      throw new Error('Tarefa não encontrada.');
    }

    // Coluna D = RESPONSÁVEL
    aba.getRange(indice + 2, 4)
      .setValue(novoResponsavel);

    SpreadsheetApp.flush();

    return { sucesso: true };
  } finally {
    lock.releaseLock();
  }
}



// Edita campos de uma tarefa existente, preservando seu ID.
function salvarTarefa(dados) {
  if (!dados || typeof dados !== 'object') throw new Error('Dados da tarefa inválidos.');
  const id = String(dados.id || '').trim();
  const titulo = String(dados.titulo || '').trim();
  const descricao = String(dados.descricao || '');
  const responsavel = String(dados.responsavel || '').trim();
  const solicitante = String(dados.solicitante || '').trim();
  const prioridade = String(dados.prioridade || 'Média').trim();
  const status = String(dados.status || '').trim();
  const dataTexto = String(dados.data || '').trim();

  if (!id) throw new Error('ID da tarefa obrigatório.');
  if (!titulo) throw new Error('O título não pode estar vazio.');
  if (titulo.length > 500) throw new Error('Título muito longo.');
  if (descricao.length > 10000) throw new Error('Descrição muito longa.');
  if (!['Backlog', 'Em progresso', 'Revisão', 'Concluído'].includes(status)) throw new Error('Status inválido.');

  const partes = /^(\d{4})-(\d{2})-(\d{2})$/.exec(dataTexto);
  if (!partes) throw new Error('Informe uma data válida.');
  const ano = Number(partes[1]);
  const mes = Number(partes[2]);
  const dia = Number(partes[3]);
  const data = new Date(ano, mes - 1, dia, 12, 0, 0);
  if (data.getFullYear() !== ano || data.getMonth() !== mes - 1 || data.getDate() !== dia) {
    throw new Error('Data inválida.');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const local=localizarCard_(id,dados.quadroId||QUADRO_PADRAO);
    local.aba.getRange(local.linha, 2, 1, 7).setValues([[
      titulo, descricao, responsavel, solicitante, prioridade, status, data
    ]]);
    SpreadsheetApp.flush();
    return { sucesso: true, id: id };
  } finally {
    lock.releaseLock();
  }
}


const ESPACOS_ABA = 'Espacos';
const QUADROS_ABA = 'Quadros';
const ESPACO_FIXO = 'nucleo-aprendizagem';
const ESPACO_PESSOAL = 'meus-quadros';
const QUADRO_PADRAO = 'quadro-administrativo';
function getOrCreateAba_(nome, cabecalhos) {
  const planilha = SpreadsheetApp.openById(PLANILHA_ID);
  let aba = planilha.getSheetByName(nome);
  if (!aba) aba = planilha.insertSheet(nome);
  if (aba.getLastRow() === 0) aba.appendRow(cabecalhos);
  return aba;
}
function inicializarQuadros_() {
  const espacos = getOrCreateAba_(ESPACOS_ABA, ['ID','NOME','FIXO']);
  const quadros = getOrCreateAba_(QUADROS_ABA, ['ID','ESPACO_ID','NOME','COR']);
  if (espacos.getLastRow() === 1) {
    espacos.getRange(2,1,2,3).setValues([
      [ESPACO_PESSOAL,'O que vamos criar hoje?',false],
      [ESPACO_FIXO,'Governança Corporativa - Núcleo de Aprendizes',true]
    ]);
  }
  if (quadros.getLastRow() === 1) {
    quadros.getRange(2,1,3,4).setValues([
      [QUADRO_PADRAO,ESPACO_FIXO,'Administrativo','#00b6ce'],
      ['quadro-transformacao',ESPACO_FIXO,'Transformação Digital','#8155df'],
      ['quadro-atendimento',ESPACO_FIXO,'Atendimento VIP - TI','#02c4d4']
    ]);
  }
  const tarefas = obterAba();
  if (tarefas.getRange(1,9).getValue() !== 'QUADRO') {
    tarefas.getRange(1,9).setValue('QUADRO');
  }
  return {espacos,quadros};
}

function criarEspaco(nome) {
  nome = String(nome || '').trim();
  if (!nome || nome.length > 70) throw new Error('Informe um nome de espaço com até 70 caracteres.');
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const abas = inicializarQuadros_();
    const novo = {id:Utilities.getUuid(),nome,fixo:false};
    abas.espacos.appendRow([novo.id,novo.nome,false]);
    return novo;
  } finally { lock.releaseLock(); }
}




/* ===== Multiplas planilhas por quadro ===== */
function quadrosConfigurados_() {
  const planilha=SpreadsheetApp.openById(PLANILHA_ID);
  const aba=planilha.getSheetByName(QUADROS_ABA);
  if(!aba)return [];
  if(aba.getRange(1,5).getValue()!=='PLANILHA_ID')aba.getRange(1,5).setValue('PLANILHA_ID');
  const n=aba.getLastRow();
  return n<2?[]:aba.getRange(2,1,n-1,5).getValues().map((r,i)=>({
    id:String(r[0]),espacoId:String(r[1]),nome:String(r[2]),cor:String(r[3]||'#00DBFF'),
    planilhaId:String(r[4]||''),linha:i+2
  }));
}
function idSheets_(texto) {
  const s=String(texto||'').trim();
  if(!s)return '';
  const m=s.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
  const id=m?m[1]:s;
  if(!/^[a-zA-Z0-9_-]{20,}$/.test(id))throw new Error('Cole um link ou ID válido da planilha.');
  return id;
}
function validarSheet_(id) {
  const aba=SpreadsheetApp.openById(id).getSheetByName('Tarefas');
  if(!aba)throw new Error('A planilha vinculada deve conter a aba Tarefas.');
  const h=aba.getRange(1,1,1,8).getValues()[0].map(v=>String(v).trim().toUpperCase());
  if(['ID','TÍTULO','DESCRIÇÃO','RESPONSÁVEL','SOLICITANTE','PRIORIDADE','STATUS','DATA'].some((x,i)=>x!==h[i])){
    throw new Error('A aba Tarefas precisa ter os cabeçalhos ID, TÍTULO, DESCRIÇÃO, RESPONSÁVEL, SOLICITANTE, PRIORIDADE, STATUS e DATA.');
  }
  return aba;
}
function abaQuadro_(quadroId) {
  const q=quadrosConfigurados_().find(x=>x.id===String(quadroId));
  if(!q)throw new Error('Quadro não encontrado.');
  return q.planilhaId?validarSheet_(q.planilhaId):obterAba();
}
function localizarCard_(id,quadroId) {
  const aba=abaQuadro_(quadroId);
  const n=aba.getLastRow();
  if(n<2)throw new Error('Tarefa não encontrada.');
  const principal=aba.getParent().getId()===PLANILHA_ID;
  const dados=aba.getRange(2,1,n-1,principal?9:1).getValues();
  const hits=dados.map((r,i)=>{
    if(String(r[0])!==String(id))return null;
    if(principal&&String(r[8]||QUADRO_PADRAO)!==String(quadroId))return null;
    return i+2;
  }).filter(Boolean);
  if(hits.length!==1)throw new Error(hits.length?'ID duplicado.':'Tarefa não encontrada.');
  return {aba,linha:hits[0]};
}
function configurarPlanilhaQuadro(dados) {
  const quadroId=String(dados&&dados.quadroId||'');
  const id=idSheets_(dados&&dados.planilha||'');
  const lock=LockService.getScriptLock();lock.waitLock(10000);
  try{
    inicializarQuadros_();
    const aba=SpreadsheetApp.openById(PLANILHA_ID).getSheetByName(QUADROS_ABA);
    const todos=quadrosConfigurados_();
    const quadro=todos.find(q=>q.id===quadroId);
    if(!quadro)throw new Error('Quadro não encontrado.');
    if(id===PLANILHA_ID)throw new Error('Para usar a planilha principal, deixe o vínculo vazio.');
    if(id&&todos.some(q=>q.id!==quadroId&&q.planilhaId===id))throw new Error('Essa planilha já está vinculada a outro quadro.');
    if(id)validarSheet_(id);
    aba.getRange(quadro.linha,5).setValue(id);
    return {sucesso:true,conectado:!!id};
  }finally{lock.releaseLock();}
}
function listarTarefas() {
  const lock=LockService.getScriptLock();lock.waitLock(10000);
  try{
    inicializarQuadros_();
    const configurados=quadrosConfigurados_();
    const fontes=[{aba:obterAba(),principal:true}];
    configurados.filter(q=>q.planilhaId).forEach(q=>fontes.push({aba:validarSheet_(q.planilhaId),principal:false,quadroId:q.id}));
    const tarefas=[];
    fontes.forEach(f=>{
      const n=f.aba.getLastRow();
      if(n<2)return;
      const valores=f.aba.getRange(2,1,n-1,f.principal?9:8).getValues();
      valores.forEach((r,i)=>{
        let [id,titulo,descricao,responsavel,solicitante,prioridade,status,data,quadroId]=r;
        if(!String(titulo||'').trim())return;
        const destino=f.principal?String(quadroId||QUADRO_PADRAO):f.quadroId;
        if(f.principal&&configurados.some(q=>q.id===destino&&q.planilhaId))return;
        if(!id){id=Utilities.getUuid();f.aba.getRange(i+2,1).setValue(id);}
        if(!status){status='Backlog';f.aba.getRange(i+2,7).setValue(status);}
        if(!data){data=new Date();f.aba.getRange(i+2,8).setValue(data);}
        tarefas.push({
          id:String(id),quadroId:destino,titulo:String(titulo),
          descricao:String(descricao||''),responsavel:String(responsavel||''),
          solicitante:String(solicitante||''),prioridade:String(prioridade||'Média'),
          status:String(status),data:data instanceof Date?Utilities.formatDate(data,Session.getScriptTimeZone(),'dd/MM/yyyy'):String(data)
        });
      });
    });
    return tarefas;
  }finally{lock.releaseLock();}
}
function atualizarStatus(id,novoStatus,quadroId) {
  if(!['Backlog','Em progresso','Revisão','Concluído'].includes(novoStatus))throw new Error('Status inválido.');
  const lock=LockService.getScriptLock();lock.waitLock(10000);
  try {
    const x=localizarCard_(id,quadroId||QUADRO_PADRAO);
    x.aba.getRange(x.linha,7).setValue(novoStatus);
    SpreadsheetApp.flush();
    return {sucesso:true};
  }finally{lock.releaseLock();}
}
function criarQuadro(dados) {
  const nome=String(dados&&dados.nome||'').trim();
  const espacoId=String(dados&&dados.espacoId||'').trim();
  const cor=String(dados&&dados.cor||'#00b6ce');
  const planilhaId=idSheets_(dados&&dados.planilha||'');
  if(!nome||nome.length>70)throw new Error('Informe nome com até 70 caracteres.');
  if(!['#00b6ce','#8155df','#02c4d4','#cd20a0','#13a889','#ef8e4a'].includes(cor))throw new Error('Cor inválida.');
  const lock=LockService.getScriptLock();lock.waitLock(10000);
  try {
    const abas=inicializarQuadros_();
    const espacos=abas.espacos.getLastRow()>1?abas.espacos.getRange(2,1,abas.espacos.getLastRow()-1,1).getValues().flat().map(String):[];
    if(!espacos.includes(espacoId))throw new Error('Espaço inexistente.');
    if(planilhaId===PLANILHA_ID)throw new Error('Deixe o vínculo vazio para usar a planilha principal.');
    if(planilhaId&&quadrosConfigurados_().some(q=>q.planilhaId===planilhaId))throw new Error('Planilha já vinculada.');
    if(planilhaId)validarSheet_(planilhaId);
    abas.quadros.getRange(1,5).setValue('PLANILHA_ID');
    const novo={id:Utilities.getUuid(),espacoId,nome,cor,conectado:!!planilhaId};
    abas.quadros.appendRow([novo.id,novo.espacoId,novo.nome,novo.cor,planilhaId]);
    return novo;
  }finally{lock.releaseLock();}
}
function criarTarefa(dados) {
  const titulo=String(dados&&dados.titulo||'').trim();
  const quadroId=String(dados&&dados.quadroId||'');
  if(!titulo||titulo.length>500)throw new Error('Informe um título válido.');
  const lock=LockService.getScriptLock();lock.waitLock(10000);
  try {
    const aba=abaQuadro_(quadroId);
    const isMain=aba.getParent().getId()===PLANILHA_ID;
    const id=Utilities.getUuid();
    aba.appendRow([id,titulo,'','','','Média','Backlog',new Date(),isMain?quadroId:'']);
    return {sucesso:true,id};
  }finally{lock.releaseLock();}
}
function estruturaPainel() {
  const lock=LockService.getScriptLock();lock.waitLock(10000);
  try {
    const abas=inicializarQuadros_();
    abas.quadros.getRange(1,5).setValue('PLANILHA_ID');
    const espacos=abas.espacos.getLastRow()>1?abas.espacos.getRange(2,1,abas.espacos.getLastRow()-1,3).getValues().map(r=>({
      id:String(r[0]),nome:String(r[1]),fixo:r[2]===true||String(r[2]).toLowerCase()==='true'
    })):[];
    const quadros=quadrosConfigurados_().map(q=>({id:q.id,espacoId:q.espacoId,nome:q.nome,cor:q.cor,conectado:!!q.planilhaId}));
    return {espacos,quadros};
  }finally{lock.releaseLock();}
}
