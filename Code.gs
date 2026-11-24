
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
        data
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

