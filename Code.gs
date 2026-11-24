
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
function listarTarefas() {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const aba = obterAba();
    const ultima = aba.getLastRow();

    if (ultima < 2) return [];

    const dados = aba
      .getRange(2, 1, ultima - 1, 8)
      .getValues();

    const tarefas = [];

    dados.forEach((linha, i) => {
      const numeroLinha = i + 2;

      let [
        id,
        titulo,
        descricao,
        responsavel,
        solicitante,
        prioridade,
        status,
        data,
        quadroId
      ] = linha;

      if (!String(titulo || '').trim()) return;

      if (!id) {
        id = Utilities.getUuid();
        aba.getRange(numeroLinha, 1).setValue(id);
      }

      if (!status) {
        status = 'Backlog';
        aba.getRange(numeroLinha, 7).setValue(status);
      }

      // Preenche a data automaticamente
      if (!data) {
        data = new Date();
        aba.getRange(numeroLinha, 8).setValue(data);
      }

      tarefas.push({
        id: String(id),
        quadroId: String(quadroId || QUADRO_PADRAO),
        titulo: String(titulo),
        descricao: String(descricao || ''),
        responsavel: String(responsavel || ''),
        solicitante: String(solicitante || ''),
        prioridade: String(prioridade || 'Média'),
        status: String(status),
        data: data instanceof Date
          ? Utilities.formatDate(
              data,
              Session.getScriptTimeZone(),
              'dd/MM/yyyy'
            )
          : String(data)
      });
    });

    return tarefas;

  } finally {
    lock.releaseLock();
  }
}

// Atualiza o status no Google Sheets
function atualizarStatus(id, novoStatus) {
  const permitidos = [
    'Backlog',
    'Em progresso',
    'Revisão',
    'Concluído'
  ];

  if (!permitidos.includes(novoStatus)) {
    throw new Error('Status inválido: ' + novoStatus);
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const aba = obterAba();
    const ultima = aba.getLastRow();

    if (ultima < 2) {
      throw new Error('Tarefa não encontrada.');
    }

    const ids = aba
      .getRange(2, 1, ultima - 1, 1)
      .getValues();

    const indice = ids.findIndex(
      linha => String(linha[0]) === String(id)
    );

    if (indice === -1) {
      throw new Error('Tarefa não encontrada.');
    }

    // Coluna G = STATUS
    aba.getRange(indice + 2, 7)
      .setValue(novoStatus);

    SpreadsheetApp.flush();

    return {
      sucesso: true,
      id: id,
      status: novoStatus
    };

  } finally {
    lock.releaseLock();
  }
}

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
    const aba = obterAba();
    const ultima = aba.getLastRow();
    if (ultima < 2) throw new Error('Tarefa não encontrada.');
    const ids = aba.getRange(2, 1, ultima - 1, 1).getValues();
    const encontrados = ids
      .map((linha, indice) => String(linha[0]) === id ? indice + 2 : null)
      .filter(numero => numero !== null);
    if (encontrados.length === 0) throw new Error('Tarefa não encontrada.');
    if (encontrados.length > 1) throw new Error('ID duplicado na planilha.');

    aba.getRange(encontrados[0], 2, 1, 7).setValues([[
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
function estruturaPainel() {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const abas = inicializarQuadros_();
    const espacos = abas.espacos.getLastRow() > 1
      ? abas.espacos.getRange(2,1,abas.espacos.getLastRow()-1,3).getValues().map(r => ({
          id:String(r[0]), nome:String(r[1]), fixo:r[2] === true || String(r[2]).toLowerCase() === 'true'
        })) : [];
    const quadros = abas.quadros.getLastRow() > 1
      ? abas.quadros.getRange(2,1,abas.quadros.getLastRow()-1,4).getValues().map(r => ({
          id:String(r[0]), espacoId:String(r[1]), nome:String(r[2]), cor:String(r[3]||'#00DBFF')
        })) : [];
    return {espacos,quadros};
  } finally { lock.releaseLock(); }
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
function criarQuadro(dados) {
  const nome = String(dados && dados.nome || '').trim();
  const espacoId = String(dados && dados.espacoId || '').trim();
  const paleta = ['#00b6ce','#8155df','#02c4d4','#cd20a0','#13a889','#ef8e4a'];
  const cor = String(dados && dados.cor || paleta[0]);
  if (!nome || nome.length > 70) throw new Error('Informe um nome de quadro com até 70 caracteres.');
  if (!paleta.includes(cor)) throw new Error('Cor inválida.');
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const abas = inicializarQuadros_();
    const ids = abas.espacos.getLastRow() > 1
      ? abas.espacos.getRange(2,1,abas.espacos.getLastRow()-1,1).getValues().flat().map(String) : [];
    if (!ids.includes(espacoId)) throw new Error('Espaço inexistente.');
    const novo = {id:Utilities.getUuid(),espacoId,nome,cor};
    abas.quadros.appendRow([novo.id,novo.espacoId,novo.nome,novo.cor]);
    return novo;
  } finally { lock.releaseLock(); }
}
function criarTarefa(dados) {
  dados = dados || {};
  const titulo = String(dados.titulo||'').trim();
  const quadroId = String(dados.quadroId||'').trim();
  if (!titulo || titulo.length>500) throw new Error('Informe um título de até 500 caracteres.');
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const abas = inicializarQuadros_();
    const ids = abas.quadros.getLastRow()>1
      ? abas.quadros.getRange(2,1,abas.quadros.getLastRow()-1,1).getValues().flat().map(String) : [];
    if (!ids.includes(quadroId)) throw new Error('Quadro inexistente.');
    const tarefaId=Utilities.getUuid();
    obterAba().appendRow([tarefaId,titulo,'','','','Média','Backlog',new Date(),quadroId]);
    return {sucesso:true,id:tarefaId};
  } finally {lock.releaseLock();}
}
