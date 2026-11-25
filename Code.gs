/**
 * HUB ALPHA — Um Google Sheets, uma aba de tarefas por quadro.
 * A planilha original é a única fonte de dados.
 * Quadros e Espacos são abas de configuração.
 * Tarefas é mantida como arquivo legado (backup, não apagado).
 */
const PLANILHA_ID = '16Tu5wve7bFeS0adr0PuUhs0Mf7Jgn6AQEm2Shz0XJQQ';
const ABA = 'Tarefas';
const ESPACOS_ABA = 'Espacos';
const QUADROS_ABA = 'Quadros';
const LISTAS_ABA = 'Listas';
const ESPACO_FIXO = 'nucleo-aprendizagem';
const ESPACO_PESSOAL = 'meus-quadros';
const QUADRO_PADRAO = 'quadro-administrativo';
const CAMPOS_TAREFA = ['ID','TÍTULO','DESCRIÇÃO','RESPONSÁVEL','SOLICITANTE','PRIORIDADE','STATUS','DATA'];
const STATUS_VALIDOS = ['Backlog','Em progresso','Revisão','Concluído'];
const LISTAS_PADRAO = [['backlog','Backlog','#00DBFF'],['progresso','Em progresso','#A44DFF'],['espera','Revisão','#FF8900'],['resolvido','Concluído','#00FFD4']];
const RESPONSAVEIS_FIXOS = [
  'Luiz Carlos','Luigi Trintoni','Gabriela Feitosa','Guilherme Crisostomo',
  'Ana Clara','Luiza Santos','Letícia de Souza','Leonardo dos Santos',
  'Andrea Conti','Tessie Simonato','Edineis Colombo'
];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index').setTitle('ALPHA | Portal Corporativo');
}
function planilha_() { return SpreadsheetApp.openById(PLANILHA_ID); }
function trava_() { const l=LockService.getScriptLock(); l.waitLock(30000); return l; }
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
  const listas=obterOuCriar_(LISTAS_ABA,['QUADRO_ID','CHAVE','NOME','COR','ORDEM']);
  // Após a primeira inicialização, excluir todas as linhas NÃO recria dados padrão.
  const bootstrap=PropertiesService.getScriptProperties();
  const initialized='HUB_ESTRUTURA_INICIALIZADA_'+PLANILHA_ID;
  const firstRun=bootstrap.getProperty(initialized)!=='sim';
  if (firstRun && espacos.getLastRow()===1) espacos.getRange(2,1,2,3).setValues([
    [ESPACO_PESSOAL,'O que vamos criar hoje?',false],
    [ESPACO_FIXO,'Governança Corporativa - Núcleo de Aprendizes',true]
  ]);
  if (firstRun && quadros.getLastRow()===1) quadros.getRange(2,1,3,6).setValues([
    [QUADRO_PADRAO,ESPACO_FIXO,'Administrativo','#00b6ce','','Administração'],
    ['quadro-transformacao',ESPACO_FIXO,'Transformação Digital','#8155df','','Transformação Digital'],
    ['quadro-atendimento',ESPACO_FIXO,'Atendimento VIP - TI','#02c4d4','','Suporte VIP']
  ]);
  quadros.getRange(1,5,1,2).setValues([['PLANILHA_ID_LEGADO','ABA_TAREFAS']]);
  // Primeira adaptação da planilha: oferece quatro listas para cada quadro existente.
  // Depois disso, qualquer remoção em Listas será respeitada.
  const listasKey='HUB_LISTAS_INICIALIZADAS_'+PLANILHA_ID;
  if(bootstrap.getProperty(listasKey)!=='sim'){
    if(listas.getLastRow()===1&&quadros.getLastRow()>1){
      const ids=quadros.getRange(2,1,quadros.getLastRow()-1,1).getValues().flat().map(String);
      const rows=[];ids.forEach(id=>LISTAS_PADRAO.forEach((c,i)=>rows.push([id,c[0],c[1],c[2],i+1])));
      if(rows.length)listas.getRange(2,1,rows.length,5).setValues(rows);
    }
    bootstrap.setProperty(listasKey,'sim');
  }
  const qtd=quadros.getLastRow()-1;
  if (qtd>0) {
    const linhas=quadros.getRange(2,1,qtd,6).getValues();
    const usados=new Set();
    const reservados=new Set([ABA,ESPACOS_ABA,QUADROS_ABA,LISTAS_ABA]);
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
  bootstrap.setProperty(initialized,'sim');
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
  if(props.getProperty(key)==='sim')return {copiadas:0,concluida:true,avisos:[]};

  const ss=planilha_(),quadros=obterQuadros_();
  const legado=ss.getSheetByName(ABA);
  let copiadas=0;
  const avisos=[];
  quadros.forEach(q=>{
    const destino=ss.getSheetByName(q.aba);
    // A migração nunca apaga ou substitui tarefas: apenas copia IDs ausentes.
    if(legado) {
      try {copiadas+=copiarSemDuplicar_(legado,destino,q.id,true);}
      catch(e){avisos.push('Aba Tarefas → '+q.nome+': '+e.message);}
    }
    if(q.legado) {
      try {
        const planilhaExterna=SpreadsheetApp.openById(q.legado);
        const origem=planilhaExterna.getSheetByName(ABA);
        if(!origem)throw new Error('Aba Tarefas não encontrada');
        copiadas+=copiarSemDuplicar_(origem,destino,q.id,false);
      } catch(e) {
        avisos.push('Importação anterior de '+q.nome+' pendente: '+e.message);
      }
    }
  });
  // Não interromper o Kanban porque uma planilha externa antiga está inacessível.
  // Mantém o marcador pendente para possibilitar uma nova tentativa.
  if(!avisos.length)props.setProperty(key,'sim');
  if(avisos.length)Logger.log('HUB ALPHA — avisos de migração: '+avisos.join(' | '));
  return {copiadas,concluida:!avisos.length,avisos};
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
    const l=planilha_().getSheetByName(LISTAS_ABA);
    const listas=l&&l.getLastRow()>1?l.getRange(2,1,l.getLastRow()-1,5).getValues().filter(r=>r[0]&&r[1]).map(r=>({quadroId:String(r[0]),key:String(r[1]),name:String(r[2]),color:String(r[3]||'#00DBFF'),ordem:Number(r[4])||0})):[];
    return {espacos,listas,quadros:obterQuadros_().map(q=>({
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
  if(!status||status.length>80)throw new Error('Status inválido.');
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
    const listas=planilha_().getSheetByName(LISTAS_ABA);
    LISTAS_PADRAO.forEach((c,i)=>listas.appendRow([id,c[0],c[1],c[2],i+1]));
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
    if([ABA,ESPACOS_ABA,QUADROS_ABA,LISTAS_ABA].includes(destino.getName()))throw new Error('Aba reservada para configurações ou backup.');
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

/**
 * Execute pelo editor do Apps Script para encontrar o erro real.
 * Não exclui dados.
 */
function diagnosticarIntegracao() {
  const resultado={planilha:'',abas:[],quadros:[],migracao:null,erros:[]};
  try{
    const ss=planilha_();
    resultado.planilha=ss.getName();
    resultado.abas=ss.getSheets().map(a=>({nome:a.getName(),linhas:a.getLastRow()}));
    try{
      const inicio=iniciarEstrutura_();
      resultado.quadros=obterQuadros_().map(q=>({
        nome:q.nome,aba:q.aba,
        existe:!!ss.getSheetByName(q.aba),
        origemLegadaConfigurada:!!q.legado
      }));
    }catch(e){resultado.erros.push('Estrutura: '+e.message);}
    try{resultado.migracao=migrarTarefasLegadas_();}
    catch(e){resultado.erros.push('Migração: '+e.message);}
    try{
      resultado.totalTarefas=listarTarefas().length;
    }catch(e){resultado.erros.push('Leitura: '+e.message);}
  }catch(e){resultado.erros.push('Acesso: '+e.message);}
  Logger.log(JSON.stringify(resultado,null,2));
  return resultado;
}


/** Exclui um card real da aba deste quadro. Não toca nas outras abas. */
function excluirTarefa(dados) {
  const id=String(dados&&dados.id||'').trim();
  const quadroId=String(dados&&dados.quadroId||'').trim();
  if(!id||!quadroId)throw new Error('Informe o card e o quadro.');
  const lock=trava_();
  try {
    preparar_();
    const x=localizarCard_(id,quadroId);
    x.aba.deleteRow(x.linha);
    return {sucesso:true,id};
  }finally{lock.releaseLock();}
}
/** Exclui quadro, tarefas e a aba exclusiva, após validação de propriedade. */
function excluirQuadro(quadroId) {
  quadroId=String(quadroId||'').trim();
  const lock=trava_();
  try {
    preparar_();
    const q=obterQuadros_().find(x=>x.id===quadroId);
    if(!q)throw new Error('Quadro não encontrado.');
    const ss=planilha_();
    const meta=ss.getSheetByName(QUADROS_ABA);
    const tab=ss.getSheetByName(q.aba);
    const usedElsewhere=obterQuadros_().some(x=>x.id!==q.id&&x.aba===q.aba);
    if(usedElsewhere)throw new Error('A aba também está vinculada a outro quadro.');
    if([ABA,ESPACOS_ABA,QUADROS_ABA,LISTAS_ABA].includes(q.aba))throw new Error('Aba de configuração reservada.');
    // Nunca exclui aba sem a confirmação de que pertence ao quadro.
    if(!tab)throw new Error('Aba vinculada não encontrada. Corrija a planilha antes de excluir.');
    meta.deleteRow(q.linha);
    const listas=ss.getSheetByName(LISTAS_ABA);
    if(listas&&listas.getLastRow()>1){const v=listas.getRange(2,1,listas.getLastRow()-1,1).getValues();for(let i=v.length-1;i>=0;i--)if(String(v[i][0])===quadroId)listas.deleteRow(i+2);}
    ss.deleteSheet(tab);
    return {sucesso:true,id:quadroId};
  }finally{lock.releaseLock();}
}
/** Exclui espaço e todos os quadros e abas associados (cascata). */
function excluirEspaco(espacoId) {
  espacoId=String(espacoId||'').trim();
  const lock=trava_();
  try {
    preparar_();
    const ss=planilha_(),esp=ss.getSheetByName(ESPACOS_ABA);
    const rows=esp.getLastRow()>1?esp.getRange(2,1,esp.getLastRow()-1,1).getValues():[];
    const pos=rows.findIndex(x=>String(x[0])===espacoId);
    if(pos<0)throw new Error('Espaço não encontrado.');
    const attached=obterQuadros_().filter(q=>q.espacoId===espacoId);
    const all=obterQuadros_();
    for(const q of attached) {
      if(!ss.getSheetByName(q.aba))throw new Error('Aba "'+q.aba+'" não encontrada.');
      if(all.some(other=>other.id!==q.id&&other.aba===q.aba))throw new Error('Aba compartilhada: '+q.aba);
      if([ABA,ESPACOS_ABA,QUADROS_ABA,LISTAS_ABA].includes(q.aba))throw new Error('Aba reservada: '+q.aba);
    }
    const metadata=ss.getSheetByName(QUADROS_ABA);
    attached.sort((a,b)=>b.linha-a.linha).forEach(q=>metadata.deleteRow(q.linha));
    const listas=ss.getSheetByName(LISTAS_ABA);
    if(listas&&listas.getLastRow()>1){const v=listas.getRange(2,1,listas.getLastRow()-1,1).getValues();for(let i=v.length-1;i>=0;i--)if(attached.some(q=>q.id===String(v[i][0])))listas.deleteRow(i+2);}
    attached.forEach(q=>ss.deleteSheet(ss.getSheetByName(q.aba)));
    esp.deleteRow(pos+2);
    return {sucesso:true,id:espacoId,quadrosExcluidos:attached.length};
  }finally{lock.releaseLock();}
}
/** Renomeia quadro mantendo o nome da aba e a associação da planilha. */
function renomearQuadro(dados) {
  const id=String(dados&&dados.id||''),nome=String(dados&&dados.nome||'').trim();
  if(!nome||nome.length>70)throw new Error('Nome inválido.');
  const lock=trava_();
  try{
    preparar_();
    const q=obterQuadros_().find(x=>x.id===id);
    if(!q)throw new Error('Quadro não encontrado.');
    planilha_().getSheetByName(QUADROS_ABA).getRange(q.linha,3).setValue(nome);
    return {sucesso:true,id,nome};
  }finally{lock.releaseLock();}
}
function renomearEspaco(dados) {
  const id=String(dados&&dados.id||''),nome=String(dados&&dados.nome||'').trim();
  if(!nome||nome.length>70)throw new Error('Nome inválido.');
  const lock=trava_();
  try {
    preparar_();
    const sheet=planilha_().getSheetByName(ESPACOS_ABA);
    const all=sheet.getLastRow()>1?sheet.getRange(2,1,sheet.getLastRow()-1,1).getValues():[];
    const idx=all.findIndex(x=>String(x[0])===id);
    if(idx<0)throw new Error('Espaço não encontrado.');
    sheet.getRange(idx+2,2).setValue(nome);
    return {sucesso:true,id,nome};
  }finally{lock.releaseLock();}
}


/** Lista/coluna Kanban editável pela aba Listas do Google Sheets. */
function salvarLista(dados) {
  const quadroId=String(dados&&dados.quadroId||'').trim(),key=String(dados&&dados.key||'').trim();
  const name=String(dados&&dados.name||'').trim();
  const color=String(dados&&dados.color||'#00DBFF');
  const ordem=Number(dados&&dados.ordem)||1;
  if(!quadroId||!key||!name||key.length>90||name.length>80)throw new Error('Dados inválidos da lista.');
  if(!/^#[0-9a-fA-F]{6}$/.test(color))throw new Error('Cor inválida.');
  const lock=trava_();
  try{
    preparar_();
    if(!obterQuadros_().some(q=>q.id===quadroId))throw new Error('Quadro não encontrado.');
    const sheet=planilha_().getSheetByName(LISTAS_ABA);
    const last=sheet.getLastRow();
    const rows=last>1?sheet.getRange(2,1,last-1,5).getValues():[];
    const index=rows.findIndex(r=>String(r[0])===quadroId&&String(r[1])===key);
    if(rows.some((r,i)=>i!==index&&String(r[0])===quadroId&&String(r[2]).toLowerCase()===name.toLowerCase()))throw new Error('Nome de lista já usado.');
    if(index>=0){
      const oldName=String(rows[index][2]);
      sheet.getRange(index+2,1,1,5).setValues([[quadroId,key,name,color,ordem]]);
      if(oldName!==name&&!['backlog','progresso','espera','resolvido'].includes(key)){
        const tasks=abaQuadro_(quadroId),n=tasks.getLastRow();
        if(n>1){
          const st=tasks.getRange(2,7,n-1,1),v=st.getValues();
          let changed=false;
          v.forEach(r=>{if(String(r[0])===oldName){r[0]=name;changed=true;}});
          if(changed)st.setValues(v);
        }
      }
    }else sheet.appendRow([quadroId,key,name,color,ordem]);
    return {sucesso:true,quadroId,key,name,color,ordem};
  }finally{lock.releaseLock();}
}
function excluirLista(dados) {
  const quadroId=String(dados&&dados.quadroId||'').trim(),key=String(dados&&dados.key||'').trim();
  const lock=trava_();
  try{
    preparar_();
    const sheet=planilha_().getSheetByName(LISTAS_ABA),last=sheet.getLastRow();
    const rows=last>1?sheet.getRange(2,1,last-1,5).getValues():[];
    const i=rows.findIndex(r=>String(r[0])===quadroId&&String(r[1])===key);
    if(i<0)throw new Error('Lista não encontrada.');
    const name=String(rows[i][2]);
    const status={backlog:'Backlog',progresso:'Em progresso',espera:'Revisão',resolvido:'Concluído'}[key]||name;
    const tab=abaQuadro_(quadroId),lastTask=tab.getLastRow();
    let deleted=0;
    if(lastTask>1){
      const values=tab.getRange(2,7,lastTask-1,1).getValues();
      for(let row=values.length-1;row>=0;row--){
        if(String(values[row][0])===status){tab.deleteRow(row+2);deleted++;}
      }
    }
    sheet.deleteRow(i+2);
    return {sucesso:true,deleted};
  }finally{lock.releaseLock();}
}
