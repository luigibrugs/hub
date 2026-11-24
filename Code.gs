/**
 * HUB ALPHA — Um Google Sheets, uma aba de tarefas por quadro.
 * A planilha original é a única fonte de dados.
 * Quadros e Espacos são abas de configuração.
 * Tarefas é mantida como arquivo legado (backup, não apagado).
 */
const PLANILHA_ID = '1VGvYnKPllv3rHFhq0MzYFCdj0hJglW45pQ5mHUEceUo';
const ABA = 'Tarefas';
const ESPACOS_ABA = 'Espacos';
const QUADROS_ABA = 'Quadros';
const ESPACO_FIXO = 'nucleo-aprendizagem';
const ESPACO_PESSOAL = 'meus-quadros';
const QUADRO_PADRAO = 'quadro-administrativo';
const CAMPOS_TAREFA = ['ID','TÍTULO','DESCRIÇÃO','RESPONSÁVEL','SOLICITANTE','PRIORIDADE','STATUS','DATA'];
const STATUS_VALIDOS = ['Backlog','Em progresso','Revisão','Concluído'];
const RESPONSAVEIS_FIXOS = [
  'Luiz Carlos','Luigi Trintoni','Gabriela Feitosa','Guilherme Crisostomo',
  'Ana Clara','Luiza Santos','Letícia de Souza','Leonardo dos Santos',
  'Andrea Conti','Tessie Simonato','Edineis Colombo'
];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index').setTitle('ALPHA | Portal Corporativo');
}
function planilha_() { return SpreadsheetApp.openById(PLANILHA_ID); }
function trava_() { const l=LockService.getScriptLock(); l.waitLock(15000); return l; }
function obterOuCriar_(nome,cabecalhos) {
  const ss=planilha_();
  let aba=ss.getSheetByName(nome);
  if (!aba) aba=ss.insertSheet(nome);
  if (aba.getLastRow()===0) aba.getRange(1,1,1,cabecalhos.length).setValues([cabecalhos]);
  return aba;
}
function validarCabecalho_(aba) {
  const atual=aba.getRange(1,1,1,8).getValues()[0].map(x=>String(x).trim().toUpperCase());
  if (CAMPOS_TAREFA.some((x,i)=>x!==atual[i])) {
    throw new Error('A aba "'+aba.getName()+'" precisa ter os cabeçalhos A:H: '+CAMPOS_TAREFA.join(', '));
  }
}
function nomeAbaSugerido_(nome,id) {
  let n=String(nome||'').trim();
  if (id===QUADRO_PADRAO) n='Administração';
  if (id==='quadro-atendimento') n='Suporte VIP';
  n=n.replace(/[\\\/\?\*\[\]:]/g,'').trim().slice(0,85);
  if (!n || ['Tarefas','Quadros','Espacos'].includes(n)) n='Quadro '+String(id).slice(0,8);
  return n;
}
function iniciarEstrutura_() {
  const espacos=obterOuCriar_(ESPACOS_ABA,['ID','NOME','FIXO']);
  const quadros=obterOuCriar_(QUADROS_ABA,['ID','ESPACO_ID','NOME','COR','PLANILHA_ID_LEGADO','ABA_TAREFAS']);
  if (espacos.getLastRow()===1) espacos.getRange(2,1,2,3).setValues([
    [ESPACO_PESSOAL,'O que vamos criar hoje?',false],
    [ESPACO_FIXO,'Governança Corporativa - Núcleo de Aprendizes',true]
  ]);
  if (quadros.getLastRow()===1) quadros.getRange(2,1,3,6).setValues([
    [QUADRO_PADRAO,ESPACO_FIXO,'Administrativo','#00b6ce','','Administração'],
    ['quadro-transformacao',ESPACO_FIXO,'Transformação Digital','#8155df','','Transformação Digital'],
    ['quadro-atendimento',ESPACO_FIXO,'Atendimento VIP - TI','#02c4d4','','Suporte VIP']
  ]);
  quadros.getRange(1,5,1,2).setValues([['PLANILHA_ID_LEGADO','ABA_TAREFAS']]);
  const qtd=quadros.getLastRow()-1;
  if (qtd>0) {
    const linhas=quadros.getRange(2,1,qtd,6).getValues();
    const usados=new Set();
    const reservados=new Set([ABA,ESPACOS_ABA,QUADROS_ABA]);
    let alterou=false;
    linhas.forEach((linha,i)=>{
      let nome=String(linha[5]||'').trim();
      if (!nome || usados.has(nome.toLowerCase()) || reservados.has(nome)) {
        const base=nomeAbaSugerido_(linha[2],linha[0]);
        nome=base;
        let t=2;
        while (usados.has(nome.toLowerCase()) || reservados.has(nome)) nome=base.slice(0,85)+' '+t++;
        linhas[i][5]=nome;
        alterou=true;
      }
      usados.add(nome.toLowerCase());
      const aba=obterOuCriar_(nome,CAMPOS_TAREFA);
      validarCabecalho_(aba);
    });
    if(alterou)quadros.getRange(2,1,qtd,6).setValues(linhas);
  }
  return {espacos,quadros};
}
function obterQuadros_() {
  const aba=planilha_().getSheetByName(QUADROS_ABA);
  if(!aba || aba.getLastRow()<2)return [];
  return aba.getRange(2,1,aba.getLastRow()-1,6).getValues().map((r,i)=>({
    id:String(r[0]),espacoId:String(r[1]),nome:String(r[2]),
    cor:String(r[3]||'#00b6ce'),legado:String(r[4]||''),
    aba:String(r[5]||''),linha:i+2
  }));
}
function abaQuadro_(id) {
  const q=obterQuadros_().find(q=>q.id===String(id));
  if(!q)throw new Error('Quadro não encontrado.');
  const aba=planilha_().getSheetByName(q.aba);
  if(!aba)throw new Error('A aba do quadro "'+q.nome+'" não existe.');
  validarCabecalho_(aba);
  return aba;
}
function copiarSemDuplicar_(origem,destino,quadroId,comColunaQuadro) {
  if(!origem || origem.getLastRow()<2)return 0;
  validarCabecalho_(origem);
  const dados=origem.getRange(2,1,origem.getLastRow()-1,comColunaQuadro?9:8).getValues();
  const ids=new Set();
  if(destino.getLastRow()>1) destino.getRange(2,1,destino.getLastRow()-1,1).getValues().forEach(r=>ids.add(String(r[0])));
  const novas=[];
  dados.forEach((linha,i)=>{
    if(!String(linha[1]||'').trim())return;
    if(comColunaQuadro && String(linha[8]||QUADRO_PADRAO)!==quadroId)return;
    let id=String(linha[0]||'');
    if(!id) {id=Utilities.getUuid();origem.getRange(i+2,1).setValue(id);}
    if(ids.has(id))return;
    ids.add(id);
    const t=linha.slice(0,8);
    t[0]=id;
    if(!t[6])t[6]='Backlog';
    if(!t[7])t[7]=new Date();
    novas.push(t);
  });
  if(novas.length)destino.getRange(destino.getLastRow()+1,1,novas.length,8).setValues(novas);
  return novas.length;
}
function migrarTarefasLegadas_() {
  const key='HUB_ABAS_MIGRADAS_V1_'+PLANILHA_ID;
  const props=PropertiesService.getScriptProperties();
  if(props.getProperty(key)==='sim')return {copiadas:0,concluida:true};
  const ss=planilha_(),quadros=obterQuadros_();
  const legado=ss.getSheetByName(ABA);
  let copiadas=0;
  const erros=[];
  quadros.forEach(q=>{
    const destino=ss.getSheetByName(q.aba);
    if(legado)copiadas+=copiarSemDuplicar_(legado,destino,q.id,true);
    if(q.legado) {
      try {
        const planilhaExterna=SpreadsheetApp.openById(q.legado);
        const origem=planilhaExterna.getSheetByName(ABA);
        if(!origem)throw new Error('Aba Tarefas não encontrada');
        copiadas+=copiarSemDuplicar_(origem,destino,q.id,false);
      }catch(e){erros.push(q.nome+': '+e.message);}
    }
  });
  if(erros.length)throw new Error('Algumas planilhas antigas não puderam ser copiadas: '+erros.join('; '));
  props.setProperty(key,'sim');
  return {copiadas,concluida:true};
}
function preparar_() {
  iniciarEstrutura_();
  return migrarTarefasLegadas_();
}
function migrarTarefasLegadas() {
  const lock=trava_();
  try{return preparar_();}finally{lock.releaseLock();}
}
function obterAba() { return planilha_().getSheetByName(ABA); }
function estruturaPainel() {
  const lock=trava_();
  try {
    const migracao=preparar_();
    const aba=planilha_().getSheetByName(ESPACOS_ABA);
    const espacos=aba.getLastRow()<2?[]:aba.getRange(2,1,aba.getLastRow()-1,3).getValues().map(r=>({
      id:String(r[0]),nome:String(r[1]),fixo:r[2]===true||String(r[2])==='true'
    }));
    const ss=planilha_();
    return {espacos,quadros:obterQuadros_().map(q=>({
      id:q.id,espacoId:q.espacoId,nome:q.nome,cor:q.cor,
      aba:q.aba,abaUrl:ss.getUrl()+'#gid='+abaQuadro_(q.id).getSheetId()
    })),migracao};
  }finally{lock.releaseLock();}
}
function listarTarefas() {
  const lock=trava_();
  try {
    preparar_();
    const tarefas=[];
    obterQuadros_().forEach(q=>{
      const aba=abaQuadro_(q.id);
      const n=aba.getLastRow();
      if(n<2)return;
      aba.getRange(2,1,n-1,8).getValues().forEach((r,i)=>{
        let [id,titulo,descricao,responsavel,solicitante,prioridade,status,data]=r;
        if(!String(titulo||'').trim())return;
        if(!id){id=Utilities.getUuid();aba.getRange(i+2,1).setValue(id);}
        if(!status){status='Backlog';aba.getRange(i+2,7).setValue(status);}
        if(!data){data=new Date();aba.getRange(i+2,8).setValue(data);}
        tarefas.push({
          id:String(id),quadroId:q.id,titulo:String(titulo),
          descricao:String(descricao||''),responsavel:String(responsavel||''),
          solicitante:String(solicitante||''),prioridade:String(prioridade||'Média'),
          status:String(status),
          data:data instanceof Date?Utilities.formatDate(data,Session.getScriptTimeZone(),'dd/MM/yyyy'):String(data)
        });
      });
    });
    return tarefas;
  }finally{lock.releaseLock();}
}
function localizarCard_(id,quadroId) {
  const aba=abaQuadro_(quadroId||QUADRO_PADRAO);
  if(aba.getLastRow()<2)throw new Error('Tarefa não encontrada');
  const rows=aba.getRange(2,1,aba.getLastRow()-1,1).getValues();
  const hits=rows.map((r,i)=>String(r[0])===String(id)?i+2:null).filter(Boolean);
  if(hits.length!==1)throw new Error(hits.length?'ID duplicado na aba.':'Tarefa não encontrada');
  return {aba,linha:hits[0]};
}
function atualizarStatus(id,status,quadroId) {
  if(!STATUS_VALIDOS.includes(status))throw new Error('Status inválido');
  const lock=trava_();
  try {
    preparar_();
    const x=localizarCard_(id,quadroId);
    x.aba.getRange(x.linha,7).setValue(status);
    return {sucesso:true};
  }finally{lock.releaseLock();}
}
function atualizarResponsavel(id,nome,quadroId) {
  if(!RESPONSAVEIS_FIXOS.includes(nome))throw new Error('Responsável inválido');
  const lock=trava_();
  try {
    preparar_();
    const x=localizarCard_(id,quadroId);
    x.aba.getRange(x.linha,4).setValue(nome);
    return {sucesso:true};
  }finally{lock.releaseLock();}
}
function salvarTarefa(dados) {
  if(!dados||typeof dados!=='object')throw new Error('Dados inválidos');
  const id=String(dados.id||'').trim();
  const titulo=String(dados.titulo||'').trim();
  const descricao=String(dados.descricao||'');
  const responsavel=String(dados.responsavel||'').trim();
  const solicitante=String(dados.solicitante||'').trim();
  const prioridade=String(dados.prioridade||'Média');
  const status=String(dados.status||'');
  const t=String(dados.data||'');
  const p=/^(\d{4})-(\d{2})-(\d{2})$/.exec(t);
  if(!id||!titulo||titulo.length>500||descricao.length>10000)throw new Error('Revise o ID, título ou descrição.');
  if(!STATUS_VALIDOS.includes(status))throw new Error('Status inválido.');
  if(!p)throw new Error('Data inválida.');
  const dt=new Date(Number(p[1]),Number(p[2])-1,Number(p[3]),12);
  if(dt.getFullYear()!==Number(p[1])||dt.getMonth()!==Number(p[2])-1||dt.getDate()!==Number(p[3]))throw new Error('Data inválida.');
  const lock=trava_();
  try {
    preparar_();
    const x=localizarCard_(id,dados.quadroId);
    x.aba.getRange(x.linha,2,1,7).setValues([[
      titulo,descricao,responsavel,solicitante,prioridade,status,dt
    ]]);
    return {sucesso:true};
  }finally{lock.releaseLock();}
}
function criarEspaco(nome) {
  nome=String(nome||'').trim();
  if(!nome||nome.length>70)throw new Error('Nome de espaço inválido');
  const lock=trava_();
  try {
    preparar_();
    const id=Utilities.getUuid();
    planilha_().getSheetByName(ESPACOS_ABA).appendRow([id,nome,false]);
    return {id,nome,fixo:false};
  }finally{lock.releaseLock();}
}
function criarQuadro(dados) {
  const nome=String(dados&&dados.nome||'').trim();
  const espacoId=String(dados&&dados.espacoId||'').trim();
  const cores=['#00b6ce','#8155df','#02c4d4','#cd20a0','#13a889','#ef8e4a'];
  const cor=String(dados&&dados.cor||cores[0]);
  if(!nome||nome.length>70||!cores.includes(cor))throw new Error('Nome ou cor inválidos');
  const lock=trava_();
  try {
    preparar_();
    const e=planilha_().getSheetByName(ESPACOS_ABA);
    const ids=e.getLastRow()>1?e.getRange(2,1,e.getLastRow()-1,1).getValues().flat().map(String):[];
    if(!ids.includes(espacoId))throw new Error('Espaço não encontrado.');
    const id=Utilities.getUuid();
    const usados=new Set(planilha_().getSheets().map(x=>x.getName().toLowerCase()));
    const base=nomeAbaSugerido_(nome,id);
    let abaNome=base,n=2;
    while(usados.has(abaNome.toLowerCase()))abaNome=base.slice(0,85)+' '+n++;
    const aba=obterOuCriar_(abaNome,CAMPOS_TAREFA);
    planilha_().getSheetByName(QUADROS_ABA).appendRow([id,espacoId,nome,cor,'',abaNome]);
    return {id,espacoId,nome,cor,aba:abaNome,abaUrl:planilha_().getUrl()+'#gid='+aba.getSheetId()};
  }finally{lock.releaseLock();}
}
function criarTarefa(dados) {
  const titulo=String(dados&&dados.titulo||'').trim();
  const quadroId=String(dados&&dados.quadroId||'').trim();
  if(!titulo||titulo.length>500)throw new Error('Título inválido');
  const lock=trava_();
  try {
    preparar_();
    const aba=abaQuadro_(quadroId);
    const id=Utilities.getUuid();
    aba.appendRow([id,titulo,'','','','Média','Backlog',new Date()]);
    return {sucesso:true,id};
  }finally{lock.releaseLock();}
}
function resolverAbaExistente_(valor) {
  const raw=String(valor||'').trim();
  const ss=planilha_();
  if(!raw)throw new Error('Informe o nome ou link da aba.');
  if(raw.includes('/spreadsheets/d/')) {
    const match=raw.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
    if(!match||match[1]!==PLANILHA_ID)throw new Error('O link precisa ser da planilha principal do HUB.');
    const gid=raw.match(/[?#&]gid=(\d+)/);
    if(!gid)throw new Error('Copie o link da aba, incluindo o trecho gid=...');
    const sheet=ss.getSheets().find(x=>String(x.getSheetId())===gid[1]);
    if(!sheet)throw new Error('Aba não encontrada nesse link.');
    return sheet;
  }
  const sheet=ss.getSheetByName(raw);
  if(!sheet)throw new Error('Aba "'+raw+'" não encontrada na planilha.');
  return sheet;
}
function configurarAbaQuadro(dados) {
  const quadroId=String(dados&&dados.quadroId||'');
  const lock=trava_();
  try {
    preparar_();
    const q=obterQuadros_().find(x=>x.id===quadroId);
    if(!q)throw new Error('Quadro inexistente');
    const destino=resolverAbaExistente_(dados&&dados.aba);
    if([ABA,ESPACOS_ABA,QUADROS_ABA].includes(destino.getName()))throw new Error('Aba reservada para configurações ou backup.');
    if(obterQuadros_().some(x=>x.id!==quadroId&&x.aba===destino.getName()))throw new Error('Essa aba já pertence a outro quadro.');
    validarCabecalho_(destino);
    const antiga=abaQuadro_(quadroId);
    const copiadas=antiga.getSheetId()===destino.getSheetId()?0:copiarSemDuplicar_(antiga,destino,quadroId,false);
    planilha_().getSheetByName(QUADROS_ABA).getRange(q.linha,6).setValue(destino.getName());
    return {sucesso:true,aba:destino.getName(),abaUrl:planilha_().getUrl()+'#gid='+destino.getSheetId(),copiadas};
  }finally{lock.releaseLock();}
}
function gerarIdsPendentes() {return listarTarefas();}
function testarConexao() {
  const resultado=estruturaPainel();
  Logger.log('Planilha conectada. Quadros: '+resultado.quadros.length);
  Logger.log('Tarefas: '+listarTarefas().length);
  return true;
}
