/* =========================================================
   Studio GS — Agenda + KPIs Operacionais + Financeiros + Vendas
   Versão estável (sem ROI)
   ========================================================= */

/* --------- NOMES DE ABAS --------- */
const NOME_ABA_CONFIG        = 'Config';
const NOME_ABA_RAW           = 'Agenda_Raw';
const NOME_ABA_KPIS_OP       = 'KPIs_Mensais';
const NOME_ABA_KPIS_FIN      = 'KPIs_Financeiros';
const NOME_ABA_KPIS_VEND_C   = 'KPIs_Vendas_Campanha';
const NOME_ABA_KPIS_VEND_M   = 'KPIs_Vendas_Mensal';
const NOME_ABA_LEADS_DEF     = 'Leads';

/* ========= UTILS ========= */
function getSheet_(name){ const ss=SpreadsheetApp.getActive(); let sh=ss.getSheetByName(name); if(!sh) sh=ss.insertSheet(name); return sh; }
function clearBelowHeader_(sheet){ const lr=sheet.getLastRow(); if(lr>1) sheet.getRange(2,1,lr-1,sheet.getLastColumn()).clearContent(); }
function round2_(n){ return Math.round((Number(n)||0)*100)/100; }
function round4_(n){ return Math.round((Number(n)||0)*10000)/10000; } // para frações -> %
function stripDiacritics_(s){ return String(s||'').normalize('NFD').replace(/[\u0300-\u036f]/g,''); }
function parseCurrencyBR_(v){
  if(typeof v==='number') return v;
  const s=String(v||'').replace(/\s/g,'').replace(/R\$/i,'').replace(/\./g,'').replace(',', '.');
  const n=Number(s); return Number.isFinite(n)?n:0;
}
function parseDateFlexible_(val){
  if (val instanceof Date && !isNaN(val.getTime())) return val;
  const s = String(val||'').trim(); if (!s) return null;
  let d = new Date(s); if (!isNaN(d.getTime())) return d;
  const m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  if (m){ const dd=+m[1], mm=+m[2]-1, yy=+m[3]; d=new Date(yy,mm,dd); if(!isNaN(d.getTime())) return d; }
  return null;
}
function ymKey_(d){ return `${d.getFullYear()}-${('0'+(d.getMonth()+1)).slice(-2)}`; }
function toYMD_(d){ const y=d.getFullYear(), m=('0'+(d.getMonth()+1)).slice(-2), day=('0'+d.getDate()).slice(-2); return `${y}-${m}-${day}`; }
function fmtHM_(d){ return ('0'+d.getHours()).slice(-2)+':'+('0'+d.getMinutes()).slice(-2); }
function normalizeName_(s){ return stripDiacritics_(String(s||'').toLowerCase().trim()); }

/* Helpers extras (datas/chaves/marketing) */
function _parseDateMaybeYM_(val){
  const d1 = parseDateFlexible_(val);
  if (d1) return d1;
  const s = String(val||'').trim();
  const m1 = s.match(/^(\d{4})-(\d{2})$/);      // "YYYY-MM"
  if (m1) return new Date(+m1[1], +m1[2]-1, 1);
  const m2 = s.match(/^(\d{1,2})\/(\d{4})$/);   // "MM/YYYY"
  if (m2) return new Date(+m2[2], +m2[1]-1, 1);
  return null;
}
function _normKey_(s){
  return stripDiacritics_(String(s||'').toLowerCase())
    .replace(/[()–—\-_%*:/]/g,' ')
    .replace(/\s+/g,' ')
    .trim();
}
function _isMarketing_(fornecedor, categoria){
  const s = stripDiacritics_(`${fornecedor||''} ${categoria||''}`).toLowerCase();
  return /(meta\/facebook|facebook|meta|anuncio online|ads|google|marketing|\bmkt\b)/i.test(s);
}

/* ========= OPEN EXTERNAL SHEET ========= */
function openExternalSpreadsheet_(sheetId){
  const id = String(sheetId||'').trim();
  if (!id) throw new Error('ID do arquivo não informado na Config.');
  return SpreadsheetApp.openById(id);
}
function getExternalSheet_(sheetId, tabName, gid){
  const ss = openExternalSpreadsheet_(sheetId);
  if (tabName){ const sh = ss.getSheetByName(tabName); if (sh) return sh; }
  if (gid !== undefined && gid !== null && String(gid).trim()!==''){
    const gidNum = Number(String(gid).trim());
    for (const sh of ss.getSheets()) if (sh.getSheetId()===gidNum) return sh;
  }
  return ss.getSheets()[0];
}

/* ========= CONFIG ========= */
function readConfig_(){
  const sh=getSheet_(NOME_ABA_CONFIG), vals=sh.getDataRange().getValues(), cfg={};
  for (let i=1;i<vals.length;i++){ const k=String(vals[i][0]||'').trim(); const v=vals[i][1]; if(k) cfg[k]=v; }
  if(!cfg['CALENDAR_ID']) cfg['CALENDAR_ID']='primary';
  if(!cfg['MESES_BACK']) cfg['MESES_BACK']=6;

  cfg['RECEITAS_SHEET_ID']=String(cfg['RECEITAS_SHEET_ID']||'').trim();
  cfg['RECEITAS_ANO']=String(cfg['RECEITAS_ANO']||new Date().getFullYear()).trim();

  cfg['DESPESAS_SHEET_ID']=String(cfg['DESPESAS_SHEET_ID']||'').trim();
  cfg['DESPESAS_TAB']=String(cfg['DESPESAS_TAB']||'Despesas').trim(); // preferir a aba "Despesas"
  cfg['DESPESAS_GID']=String(cfg['DESPESAS_GID']||'').trim();

  cfg['LEADS_TAB']=String(cfg['LEADS_TAB']||NOME_ABA_LEADS_DEF).trim();

  cfg['NOTION_TOKEN']=String(cfg['NOTION_TOKEN']||'').trim();
  cfg['NOTION_DATABASE_ID']=String(cfg['NOTION_DATABASE_ID']||'').trim();

  // Cores (IDs do Calendar como strings)
  cfg['COLOR_INICIAL']=String(cfg['COLOR_INICIAL']||'11').trim();   // amarelo = "11"
  cfg['COLOR_INDIVIDUAL']=String(cfg['COLOR_INDIVIDUAL']||'').trim(); // vazio "" (laranja)
  cfg['COLOR_DUO']=String(cfg['COLOR_DUO']||'1').trim();             // lavanda = "1"

  // (opcional) palavra-chave para detectar presença no título/descrição
  cfg['KEYWORD_PRESENCA']=String(cfg['KEYWORD_PRESENCA']||'').trim();

  return cfg;
}

/* ========= HEADERS ========= */
function ensureAgendaRawHeaders_(){
  const sh=getSheet_(NOME_ABA_RAW);
  const headers=['Data','Hora_Início','Hora_Fim','Título','Descrição','ColorId','Tipo','Aluno','Presença','Foi_Inicial','Foi_Para_Receita','Slot_Key','Serviço','Tipo_Cliente'];
  const first=sh.getRange(1,1,1,headers.length).getValues()[0];
  const ok = first.length===headers.length && first.every((v,i)=>v===headers[i]);
  if(!ok){ sh.clear(); sh.getRange(1,1,1,headers.length).setValues([headers]); }
}
function ensureKPIsOpHeaders_(){
  const sh=getSheet_(NOME_ABA_KPIS_OP);
  const headers=[
    'Mês','Aulas_Previstas','Aulas_Dadas','Presença_%',
    'Ocupação_Teórica_%','Ocupação_Preferência_%','Ocupação_Pagantes_%',
    'Iniciais_Agendadas','Iniciais_Dadas','Iniciais_Convertidas',
    'Slots_Distintos',
    'Ag_Pilates','Dadas_Pilates','Ag_Liberacao','Dadas_Liberacao','Ag_Fisio','Dadas_Fisio',
    'Lib_Ag_Novos','Lib_Ag_Alunos','Fisio_Ag_Novos','Fisio_Ag_Alunos'
  ];
  const first=sh.getRange(1,1,1,headers.length).getValues()[0];
  const ok = first.length===headers.length && first.every((v,i)=>v===headers[i]);
  if(!ok){ sh.clear(); sh.getRange(1,1,1,headers.length).setValues([headers]); }
}

function ensureKPIsFinHeaders_(){
  const sh=getSheet_(NOME_ABA_KPIS_FIN);
  const headers=[
    'Mês','Receita_Total','Despesa_Total','Fluxo_Caixa','Lucro_Líquido',
    'Alunos_Ativos','Receita_por_Aluno','Custo_por_Aluno',
    'Novas_Matrículas','Marketing','CAC','Ticket_Médio',
    'Churn_%','Retenção_%','LTV','%Marketing/Receita'
  ];
  const first=sh.getRange(1,1,1,headers.length).getValues()[0];
  const ok = first.length===headers.length && first.every((v,i)=>v===headers[i]);
  if(!ok){ sh.clear(); sh.getRange(1,1,1,headers.length).setValues([headers]); }
}
function ensureKPIsVendCampHeaders_(){
  const sh=getSheet_(NOME_ABA_KPIS_VEND_C);
  const headers=[
    'Campanha','Leads','Agendadas','Dadas',
    'Convertidos','Perdidos_sem_aula','Perdidos_com_aula',
    'Tx_Lead_Agendada_%','Show_Inicial_%','Conv_Inicial_%','Conv_Lead_%',
    'Lost_sem_aula_%','Lost_com_aula_%',
    'Ticket_Medio','Receita','Leads_por_Origem',
    'Leads_por_Venda','Leads_por_Real'
  ];
  const first=sh.getRange(1,1,1,headers.length).getValues()[0];
  const ok = first.length===headers.length && first.every((v,i)=>v===headers[i]);
  if(!ok){ sh.clear(); sh.getRange(1,1,1,headers.length).setValues([headers]); }
}
function ensureKPIsVendMensalHeaders_(){
  const sh=getSheet_(NOME_ABA_KPIS_VEND_M);
  const headers=[
    'Mês','Leads','Agendadas','Dadas',
    'Convertidos','Perdidos_sem_aula','Perdidos_com_aula',
    'Tx_Lead_Agendada_%','Show_Inicial_%','Conv_Inicial_%','Conv_Lead_%',
    'Ticket_Medio','Receita',
    'Leads_por_Venda','Leads_por_Real'
  ];
  const first=sh.getRange(1,1,1,headers.length).getValues()[0];
  const ok = first.length===headers.length && first.every((v,i)=>v===headers[i]);
  if(!ok){ sh.clear(); sh.getRange(1,1,1,headers.length).setValues([headers]); }
}
function ensureAllHeaders_(){ ensureAgendaRawHeaders_(); ensureKPIsOpHeaders_(); ensureKPIsFinHeaders_(); ensureKPIsVendCampHeaders_(); ensureKPIsVendMensalHeaders_(); }

/* ========= RECEITAS ========= */
const PT_MONTHS={'janeiro':1,'fevereiro':2,'março':3,'marco':3,'abril':4,'maio':5,'junho':6,'julho':7,'agosto':8,'setembro':9,'outubro':10,'novembro':11,'dezembro':12};
function monthNumberFromSheetName_(name){ return PT_MONTHS[String(name||'').toLowerCase().trim()]||null; }
function readReceitasRows_(cfg){
  if(!cfg['RECEITAS_SHEET_ID']) return [];
  const ss=openExternalSpreadsheet_(cfg['RECEITAS_SHEET_ID']), ano=Number(cfg['RECEITAS_ANO']), rows=[];
  ss.getSheets().forEach(sh=>{
    const mNum=monthNumberFromSheetName_(sh.getName()); if(!mNum) return;
    const range=sh.getRange(1,1,200,20).getValues();
    let rHdr=-1, cNome=-1, cValor=-1;
    for(let r=1;r<range.length;r++){
       for(let c=0;c<range[r].length;c++){
         const cell=String(range[r][c]||'').toLowerCase().trim();
         if(cell==='nome aluno/paciente'){ rHdr=r; cNome=c; }
         if(cell.indexOf('valor total')===0 || cell==='valor total recebido' || cell==='valor'){
           rHdr=(rHdr<0?r:rHdr); cValor=c;
         }
       }
       if(rHdr>=0&&cNome>=0&&cValor>=0) break;
    }
    if(rHdr<0||cNome<0||cValor<0) return;
    for(let r=rHdr+1;r<range.length;r++){
      const nomeRaw=String(range[r][cNome]||'').trim();
      if(!nomeRaw) continue;
      const nomeNorm=normalizeName_(nomeRaw);
      if (/^(total|soma|subtotal|fechamento)$/i.test(nomeNorm)) continue;
      const vTot=range[r][cValor];
      const valor=parseCurrencyBR_(vTot); if(valor<=0) continue;
      rows.push({aluno:nomeRaw, alunoNorm:nomeNorm, data:new Date(ano,mNum-1,1), valor});
    }
  });
  return rows;
}

/* ========= DESPESAS (arquivo externo) ========= */
// Usa preferencialmente a aba "Despesas" (tabela transacional). Fallback: "Dashboard".
function readDespesasRows_(cfg){
  const sheetId = String(cfg['DESPESAS_SHEET_ID']||'').trim();
  if (!sheetId) return [];

  const ss = openExternalSpreadsheet_(sheetId);

  // ordem de preferência de abas: Config -> "Despesas" -> "Dashboard" -> primeira
  const prefTabs = [ String(cfg['DESPESAS_TAB']||'Despesas').trim(), 'Despesas', 'Dashboard' ];
  let sh = null;
  for (const t of prefTabs){
    if (!t) continue;
    const cand = ss.getSheetByName(t);
    if (cand){ sh=cand; break; }
  }
  if (!sh) sh = ss.getSheets()[0];

  const vals = sh.getDataRange().getValues();
  if (!vals || vals.length < 2) return [];

  // 1) tenta tabela transacional: Data|Valor (+ Categoria/Fornecedor)
  const hdr = vals[0].map(h => _normKey_(h));
  const idxLike = (res)=>{ for (let i=0;i<hdr.length;i++){ if (res.some(r=>r.test(hdr[i]))) return i; } return -1; };

  const iData = idxLike([/^data\b/, /\bcompetenc/, /\bmes\b/]);
  const iVal  = idxLike([/^valor\b/, /\bvalor total\b/, /\bvalor r\$/]);
  const iCat  = idxLike([/\bcategoria\b/]);
  const iForn = idxLike([/\bfornecedor\b/, /\bfornecedora\b/]);

  if (iData>=0 && iVal>=0){
    const out=[];
    for (let r=1;r<vals.length;r++){
      const d = _parseDateMaybeYM_(vals[r][iData]);
      const v = parseCurrencyBR_(vals[r][iVal]);
      if (!d || !v) continue;
      const cat  = (iCat>=0  ? vals[r][iCat]  : '');
      const forn = (iForn>=0 ? vals[r][iForn] : '');
      const isMkt = _isMarketing_(forn, cat);
      out.push({ data:d, valor:v, categoria: isMkt ? 'marketing' : 'total' });
    }
    return out;
  }

  // 2) Fallback: Dashboard chave-valor (Indicador | Valor)
  let ymStr = '', desp=0, mkt=0;
  for (let r=0; r<vals.length; r++){
    const key = _normKey_(vals[r][0]);
    const val = vals[r][1];
    if (key.startsWith('mes aaaa mm') || key === 'mes'){
      ymStr = String(val||'').trim();
      continue;
    }
    if (key.startsWith('despesa')) {
      desp += parseCurrencyBR_(val);
      continue;
    }
    if (key.startsWith('marketing')) {
      mkt += parseCurrencyBR_(val);
      continue;
    }
  }
  const dYM = _parseDateMaybeYM_(ymStr) || new Date();
  const out=[];
  if (desp>0) out.push({ data:dYM, valor:desp, categoria:'total' });
  if (mkt>0)  out.push({ data:dYM, valor:mkt,  categoria:'marketing' });
  return out;
}

/* ========= VENDAS / LEADS ========= */
// (NÃO MEXER)
function runInitLeadsHeaders(){
  const sh=getSheet_('Leads');
  const headers=['Nome','Status Funil','Campanha','Data do primeiro contato','Valor Plano','Origem'];
  const first=sh.getRange(1,1,1,headers.length).getValues()[0];
  const ok = first.length===headers.length && first.every((v,i)=>v===headers[i]);
  if(!ok){ sh.clear(); sh.getRange(1,1,1,headers.length).setValues([headers]); }
}
function readLeads_(cfg){
  const sh=getSheet_(cfg['LEADS_TAB']||NOME_ABA_LEADS_DEF);
  const vals=sh.getDataRange().getValues();
  if(!vals || vals.length<2) return [];

  const hdr = vals[0].map(h => String(h||'').trim().toLowerCase());
  const iNome  = hdr.indexOf('nome');
  const iFunil = hdr.indexOf('status funil'); // COLUNA B
  const iCamp  = hdr.indexOf('campanha');
  const iData  = hdr.indexOf('data do primeiro contato');
  const iValor = hdr.indexOf('valor plano');
  const iOrig  = hdr.indexOf('origem');

  if ([iNome,iFunil,iCamp,iData,iValor,iOrig].some(i=>i<0)) {
    throw new Error('Leads: Esperado cabeçalhos: Nome | Status Funil | Campanha | Data do primeiro contato | Valor Plano | Origem');
  }

  const leads=[];
  for(let r=1;r<vals.length;r++){
    const row=vals[r];
    const nome  = String(row[iNome] || '').trim();
    const funil = String(row[iFunil]|| '').trim();
    const camp  = String(row[iCamp] || '').trim();
    const data  = parseDateFlexible_(row[iData]);
    const valor = parseCurrencyBR_(row[iValor]||0);
    const origem= String(row[iOrig] || '').trim();

    if (!nome && !camp && !funil) continue;
    leads.push({ nome, statusFunil: funil, campanha: camp, dataEntrada: data, valorPlano: valor, origem });
  }
  return leads;
}
function _normFunil_(s){ return stripDiacritics_(String(s||'').toLowerCase()).replace(/\s+/g,' ').trim(); }
function classifyLead_(l){
  const sf = _normFunil_(l.statusFunil);
  let ag=false, dada=false, conv=false, perdSem=false, perdCom=false;

  if (sf.includes('perdido sem')) { perdSem = true; }
  else if (sf.includes('novo')) { /* só lead */ }
  else if (sf.includes('aula dada') || sf.includes('consulta dada') || sf.includes('inicial dada') || sf.includes('consulta realizada')) { ag = true; dada = true; }
  else if (sf.includes('aula agendada') || sf.includes('consulta agendada') || sf.includes('inicial agendada') || sf.includes('agendada')) { ag = true; }
  else if (sf.includes('convertido')) { ag = true; dada = true; conv = true; }
  else if (sf.includes('perdido')) { ag = true; dada = true; perdCom = true; }

  return {ag,dada,conv,perdSem,perdCom};
}
function summarizeByCampaign_(leads){
  const map=new Map(); const norm=s=>String(s||'').trim()||'Sem campanha';
  leads.forEach(l=>{
    const camp=norm(l.campanha);
    if(!map.has(camp)) map.set(camp,{leads:0,ag:0,dadas:0,conv:0,perdSem:0,perdCom:0, receita:0, tickets:[], origens:new Map()});
    const m=map.get(camp);
    m.leads += 1;

    const f=classifyLead_(l);
    if(f.ag)     m.ag+=1;
    if(f.dada)   m.dadas+=1;
    if(f.conv){  m.conv+=1; if(l.valorPlano){ m.receita+=(Number(l.valorPlano)||0); m.tickets.push(Number(l.valorPlano)||0); } }
    if(f.perdSem) m.perdSem+=1;
    if(f.perdCom) m.perdCom+=1;

    const o=String(l.origem||'').trim()||'—';
    m.origens.set(o, (m.origens.get(o)||0)+1);
  });

  const rows=[]; let T={leads:0,ag:0,dadas:0,conv:0,perdSem:0,perdCom:0, receita:0, tickets:[]};
  for(const [camp,m] of map.entries()){
    const txLeadAg = m.leads? (m.ag/m.leads):0;
    const show     = m.ag?    (m.dadas/m.ag):0;
    const convIni  = m.dadas? (m.conv/m.dadas):0;
    const convLead = m.leads? (m.conv/m.leads):0;
    const lostSem  = m.leads? (m.perdSem/m.leads):0;
    const lostCom  = m.ag?    (m.perdCom/m.ag):0;
    const ticket   = m.tickets.length? (m.tickets.reduce((a,b)=>a+b,0)/m.tickets.length):0;
    const origemStr = Array.from(m.origens.entries()).map(([k,v])=>`${k}:${v}`).join('; ');
    const leadsPorVenda = m.conv>0 ? round2_(m.leads/m.conv) : '';
    const leadsPorReal  = m.receita>0 ? round2_(m.leads/m.receita) : '';

    rows.push([
      camp, m.leads, m.ag, m.dadas, m.conv, m.perdSem, m.perdCom,
      round4_(txLeadAg), round4_(show), round4_(convIni), round4_(convLead),
      round4_(lostSem), round4_(lostCom),
      round2_(ticket), round2_(m.receita), origemStr,
      leadsPorVenda, leadsPorReal
    ]);

    T.leads+=m.leads; T.ag+=m.ag; T.dadas+=m.dadas; T.conv+=m.conv; T.perdSem+=m.perdSem; T.perdCom+=m.perdCom; T.receita+=m.receita; T.tickets.push(...m.tickets);
  }

  if(rows.length){
    rows.sort((a,b)=>String(a[0]).localeCompare(String(b[0])));
    const txLeadAg = T.leads? (T.ag/T.leads):0;
    const show     = T.ag?    (T.dadas/T.ag):0;
    const convIni  = T.dadas? (T.conv/T.dadas):0;
    const convLead = T.leads? (T.conv/T.leads):0;
    const lostSem  = T.leads? (T.perdSem/T.leads):0;
    const lostCom  = T.ag?    (T.perdCom/T.ag):0;
    const ticket   = T.tickets.length? (T.tickets.reduce((a,b)=>a+b,0)/T.tickets.length):0;
    const leadsPorVenda = T.conv>0 ? round2_(T.leads/T.conv) : '';
    const leadsPorReal  = T.receita>0 ? round2_(T.leads/T.receita) : '';
    rows.push(['TOTAL', T.leads, T.ag, T.dadas, T.conv, T.perdSem, T.perdCom,
      round4_(txLeadAg), round4_(show), round4_(convIni), round4_(convLead),
      round4_(lostSem), round4_(lostCom),
      round2_(ticket), round2_(T.receita), '', leadsPorVenda, leadsPorReal
    ]);
  }
  return rows;
}
function summarizeByMonth_(leads){
  const map=new Map();
  leads.forEach(l=>{
    const d=l.dataEntrada, ym=(d instanceof Date && !isNaN(d))? ymKey_(d) : 'SemData';
    if(!map.has(ym)) map.set(ym,{leads:0,ag:0,dadas:0,conv:0,perdSem:0,perdCom:0, receita:0, tickets:[]});
    const m=map.get(ym);
    m.leads+=1;
    const f=classifyLead_(l);
    if(f.ag) m.ag+=1;
    if(f.dada) m.dadas+=1;
    if(f.conv){ m.conv+=1; if(l.valorPlano){ m.receita+=(Number(l.valorPlano)||0); m.tickets.push(Number(l.valorPlano)||0); } }
    if(f.perdSem) m.perdSem+=1;
    if(f.perdCom) m.perdCom+=1;
  });
  const rows=[];
  Array.from(map.keys()).sort().forEach(ym=>{
    const m=map.get(ym);
    const txLeadAg=m.leads?(m.ag/m.leads):0;
    const show=m.ag?(m.dadas/m.ag):0;
    const convIni=m.dadas?(m.conv/m.dadas):0;
    const convLead=m.leads?(m.conv/m.leads):0;
    const ticket=m.tickets.length?(m.tickets.reduce((a,b)=>a+b,0)/m.tickets.length):0;
    const leadsPorVenda = m.conv>0 ? round2_(m.leads/m.conv) : '';
    const leadsPorReal  = m.receita>0 ? round2_(m.leads/m.receita) : '';
    rows.push([ym, m.leads, m.ag, m.dadas, m.conv, m.perdSem, m.perdCom,
      round4_(txLeadAg), round4_(show), round4_(convIni), round4_(convLead),
      round2_(ticket), round2_(m.receita),
      leadsPorVenda, leadsPorReal
    ]);
  });
  return rows;
}
function updateKPIsVendas_(cfg){
  const leads=readLeads_(cfg);

  const rowsCamp = summarizeByCampaign_(leads);
  const shC=getSheet_(NOME_ABA_KPIS_VEND_C);
  ensureKPIsVendCampHeaders_();
  clearBelowHeader_(shC);
  if(rowsCamp.length){
    shC.getRange(2,1,rowsCamp.length,rowsCamp[0].length).setValues(rowsCamp);
    const n=rowsCamp.length;
    shC.getRange(2,8,n,6).setNumberFormat('0.0%');           // frações -> %
    shC.getRange(2,14,n,2).setNumberFormat('"R$" #,##0.00'); // Ticket, Receita
  }

  const rowsMes = summarizeByMonth_(leads);
  const shM=getSheet_(NOME_ABA_KPIS_VEND_M);
  ensureKPIsVendMensalHeaders_();
  clearBelowHeader_(shM);
  if(rowsMes.length){
    shM.getRange(2,1,rowsMes.length,rowsMes[0].length).setValues(rowsMes);
    const n=rowsMes.length;
    shM.getRange(2,8,n,4).setNumberFormat('0.0%');           // frações -> %
    shM.getRange(2,12,n,2).setNumberFormat('"R$" #,##0.00'); // Ticket, Receita
  }
}
function runSyncVendas(){
  const cfg=readConfig_();
  ensureKPIsVendCampHeaders_();
  ensureKPIsVendMensalHeaders_();
  updateKPIsVendas_(cfg);
}

/* ========= AGENDA (Google Calendar) ========= */
function detectServicoFromText_(title, desc){
  const t = stripDiacritics_((title||'')+' '+(desc||'')).toLowerCase();
  if (t.includes('liberacao') || t.includes('liberação')) return 'Liberação';
  if (t.includes('fisio') || t.includes('fisioterapia')) return 'Fisioterapia';
  return 'Pilates'; // padrão
}
function extractAlunoFromTitle_(title){
  const s = String(title||'');
  const sep = s.split(/[-—–|•]/); // pega antes de hífen/traço/comum
  const first = sep[0].trim();
  return first || s.trim();
}
// aceita emojis, 'v'/'V' e palavra-chave vinda da agenda
function presenceIsGiven_(cellValue, titleOrDesc, cfg){
  const s = String(cellValue || '').trim().toLowerCase();
  const td = String(titleOrDesc || '').toLowerCase();
  const kw = String(cfg && cfg['KEYWORD_PRESENCA'] || '').trim().toLowerCase();

  // marcações diretas
  if (/[✅✔✓]/.test(s)) return true;
  if (s === 'v' || s === '1' || s === 'p' || s === 'ok' || s === 'x' || s === 'presente') return true;

  // se a célula estiver vazia, tenta identificar no título/descrição
  if (/[✅✔✓]/.test(td)) return true;
  if (kw && td.indexOf(kw) !== -1) return true;

  // também aceita 'v'/'V' isolado na descrição/título
  if (/\bv\b/i.test(td)) return true;

  return false;
}

function runSyncAgenda(){
  const cfg=readConfig_();
  ensureAgendaRawHeaders_();
  const sh = getSheet_(NOME_ABA_RAW);
  clearBelowHeader_(sh);

  const cal = CalendarApp.getCalendarById(String(cfg['CALENDAR_ID']||'primary'));
  if (!cal) throw new Error('Calendar não encontrado (ver CALENDAR_ID na Config).');

  const mesesBack=+cfg['MESES_BACK']||6, now=new Date();
  const start = new Date(now.getFullYear(), now.getMonth()-mesesBack, 1, 0,0,0);
  const end   = new Date(now.getFullYear(), now.getMonth()+1, 1, 0,0,0);

  const colorMap = {
    inicial: String(cfg['COLOR_INICIAL']||'11').trim(),  // "11" amarelo
    indiv:   String(cfg['COLOR_INDIVIDUAL']||'').trim(), // "" (vazio) laranja
    duo:     String(cfg['COLOR_DUO']||'1').trim()        // "1" lavanda
  };
  function classifyColorId_(cRaw){
    const c = String(cRaw||'').trim();
    const isInit  = (c === colorMap.inicial);
    const isIndiv = (c === colorMap.indiv);   // só vazio
    const isDuo   = (c === colorMap.duo);
    const isValida= (isInit || isIndiv || isDuo);
    return {isInit,isIndiv,isDuo,isValida,colorId:c};
  }

  const eventos = cal.getEvents(start, end);
  const rows=[];
  eventos.forEach(ev=>{
    if (ev.isAllDayEvent && ev.isAllDayEvent()) return;
    const st=ev.getStartTime(), en=ev.getEndTime();
    if (!st || !en) return;

    const titulo=ev.getTitle()||'', desc=ev.getDescription()||'';
    const colorRaw = (typeof ev.getColor === 'function') ? String(ev.getColor()||'').trim() : '';
    const {isInit,isIndiv,isDuo,isValida,colorId} = classifyColorId_(colorRaw);
    if (!isValida) return;

    const servico = detectServicoFromText_(titulo, desc);
    const aluno = extractAlunoFromTitle_(titulo);

    const foiInicial = !!isInit;
    let tipoCliente = '';
    if (isDuo) tipoCliente='Duo';
    else if (isIndiv) tipoCliente='Individual';

    const slotKey = `${toYMD_(st)}|${fmtHM_(st)}|${fmtHM_(en)}|${titulo}`;

    // detecta presença olhando descrição/título + keyword (preenche ✅ ou vazio)
    const presenca = presenceIsGiven_('', `${titulo} ${desc}`, cfg) ? '✅' : '';

    rows.push([
      new Date(st.getFullYear(), st.getMonth(), st.getDate()),
      fmtHM_(st),
      fmtHM_(en),
      titulo,
      desc,
      colorId,
      servico,
      aluno,
      presenca,
      foiInicial,
      '',
      slotKey,
      servico,
      tipoCliente
    ]);
  });

  if (rows.length){
    rows.sort((a,b)=> (a[0]-b[0]) || String(a[1]).localeCompare(String(b[1])) || String(a[3]).localeCompare(String(b[3])));
    sh.getRange(2,1,rows.length,rows[0].length).setValues(rows);
  }
}

/* ========= OCUPAÇÃO ========= */
// Slots fixos do estúdio (1h): seg–sex 08:30,09:30,10:30,11:30,14:00,15:00,16:00
function _hmToMin_(hm){ const m=String(hm||'').match(/^(\d{1,2}):(\d{2})$/); if(!m) return null; return (+m[1])*60+(+m[2]); }
function _monthSlots_(year, mon){ // mon: 0-11
  const MORNING = ['08:30','09:30','10:30','11:30'];
  const AFTERNOON = ['14:00','15:00','16:00'];
  const out=[];
  const daysInMonth = new Date(year, mon+1, 0).getDate();
  for(let d=1; d<=daysInMonth; d++){
    const dt = new Date(year, mon, d);
    const dow = dt.getDay(); // 0=dom,1=seg,...6=sáb
    if(dow===0 || dow===6) continue; // seg–sex
    const ymd = `${dt.getFullYear()}-${('0'+(dt.getMonth()+1)).slice(-2)}-${('0'+dt.getDate()).slice(-2)}`;
    MORNING.concat(AFTERNOON).forEach(hm=>{
      out.push({ slotKey: `${ymd}|${hm}`, ymd, hm });
    });
  }
  return out;
}
// Consolida eventos por slot (por mês) a partir da Agenda_Raw
function _collectEventsPerSlot_(vals){
  const hdr = vals[0].map(h=>String(h||'').toLowerCase());
  const ixDate = hdr.indexOf('data');
  const ixStart= hdr.indexOf('hora_início');   // pode vir Date ou String
  const ixAluno= hdr.indexOf('aluno');
  const ixIni  = hdr.indexOf('foi_inicial');
  const ixTipo = hdr.indexOf('tipo_cliente');  // 'Individual' | 'Duo'

  const map=new Map(); // ym -> slotKey -> {indiv, duo, inicial, alunos:Set}

  for(let r=1; r<vals.length; r++){
    const row=vals[r];

    // Data
    const d = parseDateFlexible_(row[ixDate]);
    if(!(d instanceof Date) || isNaN(d)) continue;
    const ymd = toYMD_(d);
    const ym  = ymKey_(d);

    // Hora início: normaliza para "HH:MM" mesmo se vier Date ou "08:30:00"
    const rawStart = row[ixStart];
    let hm = '';
    if (rawStart instanceof Date && !isNaN(rawStart)) {
      hm = fmtHM_(rawStart);
    } else {
      const s = String(rawStart||'').trim();
      const m = s.match(/(\d{1,2}):(\d{2})/); // aceita "8:30", "08:30", "08:30:00"
      if (!m) continue;
      hm = ('0'+m[1]).slice(-2)+':'+m[2];
    }

    const slotKey = `${ymd}|${hm}`;
    if(!map.has(ym)) map.set(ym, new Map());
    const mYM = map.get(ym);
    if(!mYM.has(slotKey)) mYM.set(slotKey,{indiv:0,duo:0,inicial:0, alunos:new Set()});
    const node = mYM.get(slotKey);

    const aluno = String(row[ixAluno]||'').trim();
    const isInicial = !!row[ixIni];
    const tipo = String(row[ixTipo]||'').trim(); // 'Individual' | 'Duo' | ''

    if (isInicial){ node.inicial++; node.alunos.add(aluno); }
    else if (tipo==='Individual'){ node.indiv++; node.alunos.add(aluno); }
    else if (tipo==='Duo'){ node.duo++; node.alunos.add(aluno); }
  }
  return map; // Map(ym -> Map(slotKey -> {...}))
}

function _computeOccupancyForMonth_(ym, slotMap){
  const y = +ym.slice(0,4), m = +ym.slice(5,7)-1;
  const slots = _monthSlots_(y,m);
  const totalSlots = slots.length;
  const capSeats = totalSlots * 2; // teórica: 2 lugares por slot
  let seatsUsed = 0;
  let slotsPrefOcc = 0;
  let slotsPagOcc  = 0;

  slots.forEach(s=>{
    const node = slotMap.get(s.slotKey);
    if(!node) return;

    const eventos = node.indiv + node.inicial + node.duo;
    // Teórica: até 2 por slot
    seatsUsed += Math.min(2, eventos);

    // Preferência: conta inicial OU individual; DUO precisa 2 no mesmo slot
    const prefFull = (node.inicial > 0) || (node.indiv > 0) || (node.duo >= 2);
    if (prefFull) slotsPrefOcc += 1;

    // Pagantes: EXCLUI iniciais → conta individual; DUO precisa 2 no mesmo slot
    const pagFull = (node.indiv > 0) || (node.duo >= 2);
    if (pagFull) slotsPagOcc += 1;
  });

  const occTeorica = capSeats>0 ? (seatsUsed/capSeats) : 0;
  const occPref    = totalSlots>0 ? (slotsPrefOcc/totalSlots) : 0;
  const occPag     = totalSlots>0 ? (slotsPagOcc/totalSlots) : 0;

  return {occTeorica, occPref, occPag};
}

/* ========= KPIs OPERACIONAIS ========= */
function runSyncKpisOperacionais(){
  ensureKPIsOpHeaders_();
  const shRaw=getSheet_(NOME_ABA_RAW);
  const vals=shRaw.getDataRange().getValues();
  if (!vals || vals.length<2) { clearBelowHeader_(getSheet_(NOME_ABA_KPIS_OP)); return; }

  const hdr = vals[0].map(h=>String(h||'').toLowerCase());
  const iData = hdr.indexOf('data');
  const iPres = hdr.indexOf('presença');
  const iServ = hdr.indexOf('serviço');
  const iAluno= hdr.indexOf('aluno');
  const iColor= hdr.indexOf('colorid');
  const iIni  = hdr.indexOf('foi_inicial');
  const iTipo = hdr.indexOf('tipo_cliente');

  if ([iData,iPres,iServ,iAluno,iColor,iIni,iTipo].some(i=>i<0)) throw new Error('Agenda_Raw com cabeçalhos inesperados.');

  const cfg = readConfig_();
  const colorValidSet = new Set([ String(cfg['COLOR_INICIAL']||'11').trim(), String(cfg['COLOR_INDIVIDUAL']||'').trim(), String(cfg['COLOR_DUO']||'1').trim() ]);

  const byMonth = new Map();
  const iniciais = []; // {alunoNorm, data, ym}
  const aulasPorAlunoFuturas = []; // {alunoNorm, data, ym}

  for (let r=1;r<vals.length;r++){
    const row=vals[r];
    const d=parseDateFlexible_(row[iData]); if(!d) continue;
    const ym=ymKey_(d);

    const colorId = String(row[iColor]||'').trim();
    if (!colorValidSet.has(colorId)) continue;

    const serv=String(row[iServ]||'').trim(); if(!serv) continue;
    const presGiven = presenceIsGiven_(row[iPres]);
    const aluno = String(row[iAluno]||'').trim();
    const alunoNorm = normalizeName_(aluno);
    const foiInicial = !!row[iIni];
    const tipoCliente = String(row[iTipo]||'').trim(); // 'Individual' | 'Duo'

    if (!byMonth.has(ym)) byMonth.set(ym, {
      prev:0, dadas:0, slots:new Set(),
      agPil:0, ddPil:0, agLib:0, ddLib:0, agFis:0, ddFis:0,
      inicAg:0, inicDd:0
    });
    const m=byMonth.get(ym);

    if (colorId === String(cfg['COLOR_INICIAL']||'11').trim()){
      m.inicAg += 1;
      if (presGiven) m.inicDd += 1;
      iniciais.push({alunoNorm, data:d, ym});
      continue; // iniciais não contam como aula de serviço
    }

    m.prev += 1;
    m.slots.add(`${toYMD_(d)}|${aluno}|${serv}`);

    if (serv==='Pilates'){ m.agPil++; if(presGiven) m.ddPil++; }
    else if(serv==='Liberação'){ m.agLib++; if(presGiven) m.ddLib++; }
    else if(serv==='Fisioterapia'){ m.agFis++; if(presGiven) m.ddFis++; }

    if (presGiven) m.dadas += 1;

    if (tipoCliente==='Individual' || tipoCliente==='Duo'){
      aulasPorAlunoFuturas.push({alunoNorm, data:d, ym});
    }
  }

  // Conversões: primeira aula futura após a inicial
  const convPorMes = new Map();
  iniciais.forEach(ini=>{
    const futuros = aulasPorAlunoFuturas
      .filter(x=> x.alunoNorm===ini.alunoNorm && x.data>ini.data)
      .sort((a,b)=> a.data - b.data);
    if (futuros.length){
      const convMes = futuros[0].ym;
      convPorMes.set(convMes, (convPorMes.get(convMes)||0)+1);
    }
  });

  // Ocupação por mês a partir da Agenda_Raw
  const perMonthSlots = _collectEventsPerSlot_(vals); // Map(ym -> Map(slotKey -> data))

  const rows=[];
  Array.from(byMonth.keys()).sort().forEach(ym=>{
    const m=byMonth.get(ym);
    const prev = m.prev;
    const dadas = m.dadas;
    const presPct = prev>0 ? dadas/prev : 0;
    const slotsDistintos = m.slots.size;
    const inicConv = convPorMes.get(ym)||0;

    const slotMap = perMonthSlots.get(ym) || new Map();
    const {occTeorica, occPref, occPag} = _computeOccupancyForMonth_(ym, slotMap);
    rows.push([
      ym,
      prev, dadas, round4_(presPct),
      round4_(occTeorica),           // Ocupação_Teórica_%
      round4_(occPref),              // Ocupação_Preferência_%
      round4_(occPag),               // Ocupação_Pagantes_% (NOVA)
      m.inicAg, m.inicDd, inicConv,
      slotsDistintos,
      m.agPil, m.ddPil, m.agLib, m.ddLib, m.agFis, m.ddFis,
      '', '', '', ''
    ]);
  });

  const sh=getSheet_(NOME_ABA_KPIS_OP);
  clearBelowHeader_(sh);
  if (rows.length){
    sh.getRange(2,1,rows.length,rows[0].length).setValues(rows);
    const n=rows.length;
   sh.getRange(2,4,n,1).setNumberFormat('0.0%'); // Presença_%
   sh.getRange(2,5,n,3).setNumberFormat('0.0%'); // Ocupações: Teórica, Preferência, Pagantes
  }
}

/* ========= FINANCEIRO (sem ROI) ========= */
function ensureKPIsFin_(cfg,start,end){
  const sh = getSheet_(NOME_ABA_KPIS_FIN);
  clearBelowHeader_(sh);

  // RECEITAS
  const recRows = readReceitasRows_(cfg);
  const receitaMes = {};
  const alunosMes = {};
  const firstMonthAluno = {}; // por alunoNorm -> menor YM
  recRows.forEach(r=>{
    const ym = ymKey_(r.data);
    receitaMes[ym] = (receitaMes[ym]||0) + (Number(r.valor)||0);

    const norm = r.alunoNorm;
    if(!alunosMes[ym]) alunosMes[ym] = new Set();
    alunosMes[ym].add(norm);

    if(!firstMonthAluno[norm]) firstMonthAluno[norm] = ym;
    else { if (ym < firstMonthAluno[norm]) firstMonthAluno[norm] = ym; }
  });

  // DESPESAS
  const desRows = readDespesasRows_(cfg);
  const despesaMes = {}, marketingMes = {};
  desRows.forEach(d=>{
    const ym = ymKey_(d.data), v = Number(d.valor)||0;
    despesaMes[ym]  = (despesaMes[ym] || 0) + v;
    if (d.categoria === 'marketing') marketingMes[ym] = (marketingMes[ym]||0) + v;
  });

  // Linha do tempo
  const meses=[]; const dd=new Date(start);
  while(dd<end){ meses.push(ymKey_(dd)); dd.setMonth(dd.getMonth()+1); }

  const out=[]; let prevActives=new Set();
  meses.forEach(ym=>{
    const receita = Number(receitaMes[ym]||0);
    const despesa = Number(despesaMes[ym]||0);
    const fluxo   = receita - despesa;
    const lucro   = fluxo;

    const activesSet = alunosMes[ym] || new Set(); // já normalizado
    const ativos = activesSet.size;

    const receitaPorAluno = ativos>0 ? receita/ativos : 0;
    const custoPorAluno   = ativos>0 ? despesa/ativos : 0;

    // novas matrículas: alunos cujo firstMonthAluno == ym
    let novas = 0; activesSet.forEach(aNorm=>{ if(firstMonthAluno[aNorm]===ym) novas++; });

    const marketing = Number(marketingMes[ym]||0);
    const cac = (novas>0 && marketing>0) ? (marketing/novas) : '';

    const ticketMedio = receitaPorAluno;

    // Churn/Retenção como fração
    let churnFrac = null, retencaoFrac = null;
    if (prevActives.size>0){
      let saidos=0; prevActives.forEach(a=>{ if(!activesSet.has(a)) saidos++; });
      churnFrac    = saidos/prevActives.size;
      retencaoFrac = 1 - churnFrac;
    }
    let ltv = '';
    if (typeof churnFrac==='number' && churnFrac>0) ltv = ticketMedio/churnFrac;

    const pctMktReceita = receita>0 ? (marketing/receita) : ''; // fração

    out.push([
      ym,
      round2_(receita), round2_(despesa), round2_(fluxo), round2_(lucro),
      ativos, round2_(receitaPorAluno), round2_(custoPorAluno), novas,
      round2_(marketing), (cac===''?'':round2_(cac)), round2_(ticketMedio),
      (typeof churnFrac==='number'? round2_(churnFrac):''),          // FRAÇÃO
      (typeof retencaoFrac==='number'? round2_(retencaoFrac):''),    // FRAÇÃO
      (ltv===''?'':round2_(ltv)),
      (pctMktReceita===''?'':round2_(pctMktReceita))                 // FRAÇÃO
    ]);

    prevActives = activesSet;
  });

  if (out.length){
    sh.getRange(2,1,out.length,out[0].length).setValues(out);

    const lastR = out.length;
    // R$
    sh.getRange(2,2,lastR,1).setNumberFormat('"R$" #,##0.00');
    sh.getRange(2,3,lastR,1).setNumberFormat('"R$" #,##0.00');
    sh.getRange(2,4,lastR,1).setNumberFormat('"R$" #,##0.00');
    sh.getRange(2,5,lastR,1).setNumberFormat('"R$" #,##0.00');
    sh.getRange(2,10,lastR,1).setNumberFormat('"R$" #,##0.00');
    sh.getRange(2,11,lastR,1).setNumberFormat('"R$" #,##0.00');
    sh.getRange(2,12,lastR,1).setNumberFormat('"R$" #,##0.00');
    // % (frações -> %)
    sh.getRange(2,13,lastR,1).setNumberFormat('0.0%'); // Churn
    sh.getRange(2,14,lastR,1).setNumberFormat('0.0%'); // Retenção
    sh.getRange(2,16,lastR,1).setNumberFormat('0.0%'); // %Marketing/Receita
  }
}
function runSyncFinance(){
  const cfg=readConfig_();
  const mesesBack=+cfg['MESES_BACK']||6, now=new Date(),
        start=new Date(now.getFullYear(),now.getMonth()-mesesBack,1),
        end=new Date(now.getFullYear(),now.getMonth()+1,1);
  ensureKPIsFinHeaders_();
  ensureKPIsFin_(cfg,start,end);
}

/* ========= NOTION (Status Funil BRUTO → coluna B) ========= */
function _cleanNotionId_(s) {
  const str = String(s || '').trim();
  const m36 = str.match(/[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/i);
  if (m36) return m36[0].replace(/-/g, '');
  const m32 = str.match(/[0-9a-f]{32}/i);
  if (m32) return m32[0];
  return '';
}
// Converte qualquer ID do Notion para o formato COM hífens (o novo Notion exige assim)
function _normalizeNotionIdDashed_(s){
  var raw = String(s || '').trim();
  var m36 = raw.match(/[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/i);
  if (m36) return m36[0].toLowerCase();
  var m32 = raw.match(/[0-9a-f]{32}/i);
  if (m32){
    var x = m32[0].toLowerCase();
    return x.slice(0,8) + '-' + x.slice(8,12) + '-' + x.slice(12,16) + '-' + x.slice(16,20) + '-' + x.slice(20);
  }
  throw new Error('NOTION_DATABASE_ID inválido: ' + s);
}

// Lê o database e descobre o data_source_id conectado à sua integração
function _getDataSourceId_(token, databaseId) {
  var dbIdDashed = _normalizeNotionIdDashed_(databaseId);
  var url = 'https://api.notion.com/v1/databases/' + dbIdDashed;
  var res = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: {
      'Authorization': 'Bearer ' + token,
      'Notion-Version': '2025-09-03'
    },
    muteHttpExceptions: true
  });
  if (res.getResponseCode() >= 300) {
    throw new Error('Notion GET databases: ' + res.getResponseCode() + ' ' + res.getContentText());
  }
  var json = JSON.parse(res.getContentText());
  var list = (json.data_sources || []);
  if (!list.length) {
    throw new Error('Esse database não tem data source conectada para a sua integração.');
  }
  return list[0].id; // se tiver várias, aqui você poderia escolher por nome
}

// Usa o novo endpoint baseado em data_sources (exigido pelo Notion em 2025-09-03)
function _notionQueryAll_(token, databaseId) {
  var dataSourceId = _getDataSourceId_(token, databaseId);
  var url = 'https://api.notion.com/v1/data_sources/' + dataSourceId + '/query';
  var out = [], body = { page_size: 100 }, more = true;
  while (more) {
    var res = UrlFetchApp.fetch(url, {
      method: 'post',
      headers: {
        'Authorization': 'Bearer ' + token,
        'Notion-Version': '2025-09-03',
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(body),
      muteHttpExceptions: true
    });
    if (res.getResponseCode() >= 300) {
      throw new Error('Notion query: ' + res.getResponseCode() + ' ' + res.getContentText());
    }
    var json = JSON.parse(res.getContentText());
    (json.results || []).forEach(function(p){ out.push(p); });
    more = !!json.has_more;
    if (more) body.start_cursor = json.next_cursor;
  }
  return out;
}

function _mapNotionLead_(page) {
  const props = page.properties || {};
  const getTitle = (...names)=>{ for(const n of names){ const p=props[n]; if(p && p.type==='title' && p.title?.length) return p.title.map(t=>t.plain_text).join(' ').trim(); } return ''; };
  const getSelect = (...names)=>{ for(const n of names){ const p=props[n]; if(!p) continue; if(p.type==='select' && p.select) return p.select.name||''; if(p.type==='status' && p.status) return p.status.name||''; } return ''; };
  const getRich = (...names)=>{ for(const n of names){ const p=props[n]; if(p && p.type==='rich_text' && p.rich_text?.length) return p.rich_text.map(t=>t.plain_text).join(' ').trim(); } return ''; };
  const getNumber=(...names)=>{ for(const n of names){ const p=props[n]; if(p && p.type==='number' && typeof p.number==='number') return p.number; } return ''; };
  const getDate=(...names)=>{ for(const n of names){ const p=props[n]; if(p && p.type==='date' && p.date){ const iso=p.date.start||p.date.end; if(iso){ const d=new Date(iso); if(!isNaN(d)) return d; } } } return null; };

  const nome         = getTitle('Nome','Name','Lead','Cliente');
  const statusFunil  = getSelect('Status Funil') || getRich('Status Funil'); // BRUTO
  const campanha     = getSelect('Campanha') || getRich('Campanha');
  const origem       = getSelect('Origem')   || getRich('Origem');
  const valor        = getNumber('Valor Plano','Valor','Plano','Preço');
  const data         = getDate('Data do primeiro contato','Data','Primeiro Contato');

  const dataBr = (data ? (('0'+data.getDate()).slice(-2)+'/'+('0'+(data.getMonth()+1)).slice(-2)+'/'+data.getFullYear()) : '');

  return { nome, statusFunil, campanha, origem, valor, dataBr };
}
function runSyncLeadsFromNotion() {
  const cfg = readConfig_();
  if (!cfg.NOTION_TOKEN || !cfg.NOTION_DATABASE_ID) {
    SpreadsheetApp.getActive().toast('Notion não configurado.');
    return {created:0, updated:0};
  }
  const sh = getSheet_(cfg['LEADS_TAB'] || NOME_ABA_LEADS_DEF);
  const vals = sh.getDataRange().getValues();
  if (vals.length < 1) throw new Error('Aba Leads sem cabeçalho.');

  // A: Nome | B: Status Funil | C: Campanha | D: Data 1º contato | E: Valor Plano | F: Origem
  const hdr = vals[0].map(h => String(h||'').trim().toLowerCase());
  const idx = {
    nome:  hdr.indexOf('nome'),
    funil: hdr.indexOf('status funil'),
    camp:  hdr.indexOf('campanha'),
    data:  hdr.indexOf('data do primeiro contato'),
    valor: hdr.indexOf('valor plano'),
    orig:  hdr.indexOf('origem')
  };
  if ([idx.nome, idx.funil, idx.camp, idx.data, idx.valor, idx.orig].some(i => i < 0))
    throw new Error('Leads: cabeçalhos inválidos.');

  const pos = new Map();
  for (let r=1; r<vals.length; r++){
    const n = String(vals[r][idx.nome]||'').trim();
    if (n && !pos.has(n)) pos.set(n, r);
  }

  const pages = _notionQueryAll_(cfg.NOTION_TOKEN, cfg.NOTION_DATABASE_ID);
  let created=0, updated=0; const append=[];

  pages.forEach(pg=>{
    const rec = _mapNotionLead_(pg);
    if(!rec || !rec.nome) return;
    const funil = rec.statusFunil || '';

    if (pos.has(rec.nome)) {
      const r = pos.get(rec.nome);
      const rowNum = r + 1;
      sh.getRange(rowNum, idx.funil+1).setValue(funil);
      sh.getRange(rowNum, idx.camp +1).setValue(rec.campanha||'');
      sh.getRange(rowNum, idx.data +1).setValue(rec.dataBr||'');
      sh.getRange(rowNum, idx.valor+1).setValue(rec.valor||'');
      sh.getRange(rowNum, idx.orig +1).setValue(rec.origem||'');
      updated++;
    } else {
      const cols = vals[0].length; // 6 colunas
      const row = new Array(cols).fill('');
      row[idx.nome]  = rec.nome;
      row[idx.funil] = funil;
      row[idx.camp]  = rec.campanha || '';
      row[idx.data]  = rec.dataBr || '';
      row[idx.valor] = rec.valor || '';
      row[idx.orig]  = rec.origem || '';
      append.push(row);
      created++;
    }
  });

  if (append.length) sh.getRange(sh.getLastRow()+1,1,append.length,append[0].length).setValues(append);
  SpreadsheetApp.getActive().toast(`Notion: ${created} criados, ${updated} atualizados.`);
  return {created, updated};
}
// Pega o data_source_id do seu database do Notion
function _getDataSourceId_(token, databaseId) {
  var url = 'https://api.notion.com/v1/databases/' + databaseId;
  var res = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: {
      'Authorization': 'Bearer ' + token,
      'Notion-Version': '2025-09-03'
    },
    muteHttpExceptions: true
  });
  if (res.getResponseCode() >= 300) {
    throw new Error('Notion GET databases: ' + res.getResponseCode() + ' ' + res.getContentText());
  }
  var json = JSON.parse(res.getContentText());
  var list = (json.data_sources || []);
  if (!list.length) throw new Error('Nenhuma data source vinculada a esse database.');
  // se você tiver mais de uma data source, pode escolher pela posição ou pelo nome:
  // var ds = list.find(x => (x.name||'').trim() === 'Minhas Aulas') || list[0];
  return list[0].id; // pega a primeira
}

/* ========= MENU / RUNNERS ========= */
function sanityCheck(){
  const cfg=readConfig_(); const msgs=[];
  if(!cfg['LEADS_TAB']) msgs.push('LEADS_TAB vazio');
  if(msgs.length) throw new Error('Config com pendências: '+msgs.join(' | '));
}
function runAll(){
  const ui=SpreadsheetApp.getUi();
  const lock=LockService.getScriptLock();
  try{
    lock.waitLock(30000);
    ui.alert('Atualização iniciada.');
    sanityCheck();
    runSyncAgenda();
    runSyncKpisOperacionais();
    runSyncFinance();
    runSyncVendas();
    ui.alert('Pronto! KPIs atualizados.');
  } catch(e){
    ui.alert('Erro ao atualizar: ' + (e && e.message ? e.message : e));
    throw e;
  } finally { try{lock.releaseLock();}catch(_){ } }
}
function onOpen(){
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('KPIs GS')
    .addItem('Atualizar Tudo','runAll')
    .addSeparator()
    .addItem('Atualizar Agenda','runSyncAgenda')
    .addItem('Atualizar KPIs Operacionais','runSyncKpisOperacionais')
    .addItem('Atualizar Financeiro','runSyncFinance')
    .addItem('Atualizar Vendas','runSyncVendas')
    .addItem('Atualizar Leads (Notion)','runSyncLeadsFromNotion')
    .addSeparator()
    .addItem('Criar cabeçalhos da aba Leads','runInitLeadsHeaders')
    .addToUi();
}
