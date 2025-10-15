/********************
 * Studio GS — Agenda → Meses → Contratos → Receita (v2)
 ********************/

/* ===== CONFIG AGENDA & PLANILHA ===== */
const CALENDAR_ID    = 'primary';
const NAME_COL       = 'D';   // nome nas abas mensais
const MODALITY_COL   = 'E';   // i/d
const QTY_COL        = 'H';   // quantidade do mês
const FIRST_DATA_ROW = 11;    // 1ª linha de dados
const CHECK_PAT = /✅|✔|☑|✓/;
var INDIVIDUAL_COLORS = [""];        // laranja = individual
var DUO_COLOR         = "1";         // lavanda = dupla
var EXCLUDED_COLORS   = ["5"];       // amarelo = ignorar

/* ===== ABAS ===== */
const SH_CFG_FIN = 'Config_Financas';
const SH_CONTR   = 'Contratos';
const SH_RECON   = 'Reconhecimento';
const SH_ROLLFWD = 'Rollforward_Passivo';

/* ===== HEADERS CONTRATOS ===== */
const CONTR_HEADERS = [
  'ID_CONTRATO','DATA_INICIO','ALUNO','NOME_NORMALIZADO',
  'PLANO','FREQUENCIA_SEMANAL','MESES_DURACAO',
  'PRECO_CHEIO_AULA','DESCONTO_%','PRECO_UNIT','QTDE_AULAS_CONTR',
  'STATUS',
  'DATA_AJUSTE','AJUSTE_AULAS_PRE_BASE','AJUSTE_MODALIDADE','AJUSTE_APLICADO',
  'RENOVAR','DATA_INICIO_NOVO','PLANO_NOVO','FREQ_NOVA','MESES_DUR_NOVO',
  'ID_RENOVACAO_DE','ID_RENOVACAO_PARA',
  'MODALIDADE_CONTR',
  'DATA_PROG_INICIO',
  'AULA_ATUAL','AULAS_RESTANTES',
  'AULAS_FEITAS_ATE_CUTOVER','AULAS_PRE_CUTOVER','AULAS_POS_CUTOVER'
];

/* ===== MENU ===== */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Financeiro do Studio')
    // ROTINA
    .addItem('1) Atualizar presenças (Agenda → Abas)', 'runUpdateAllMonths')
    .addItem('2) Atualizar progresso (CUTOVER)', 'fin_atualizarProgressoContratos')
    .addSeparator()
    .addItem('Novo contrato (da Agenda — com plano)', 'fin_baixarNovosDaAgendaPrompt')
    .addItem('Criar renovação (linhas selecionadas)', 'fin_criarRenovacaoLinhasSelecionadas')
    .addItem('Alterar contrato no meio (modalidade/freq/plano)', 'fin_dividirContratoNoMeio')
    .addSeparator()
    .addItem('Resumo saldos (Mensal x Planos)', 'fin_gerarResumoSaldos')
    .addItem('Reconhecer (GERAL – todos os planos)', 'fin_reconhecerGERALTodosPlanos') // << aqui
    .addSeparator()
    // APOIO
    .addItem('Diagnóstico pós-cutover', 'fin_diagPosCutover')
    .addItem('Corrigir descontos (todos os contratos)', 'fin_corrigirDescontosContratosAntigos')
    .addItem('Corrigir validação (Desconto/Mod Nova)', 'fin_fixDescontoNovoValidation')
    .addToUi();
}


/* ===== HELPERS ===== */
function notify_(msg){
  try{ SpreadsheetApp.getActive().toast(msg); }catch(e){}
  try{ SpreadsheetApp.getUi().alert(msg);}catch(e){}
  Logger.log(msg);
}
function getSheet_(name){ const ss=SpreadsheetApp.getActive(); let sh=ss.getSheetByName(name); if(!sh) sh=ss.insertSheet(name); return sh; }
function normalizeName_(s){ if(!s) return ''; return String(s).toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/[^a-z\s]/g,'').replace(/\s+/g,' ').trim(); }
function pickBetterDisplay_(a,b){ const A=(a||'').trim(), B=(b||'').trim(); if(!A) return B; if(!B) return A; return B.length>A.length?B:A; }
function getMonthNamePt_(i){ return ['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'][i]; }
function pad2_(n){ return (n<10?'0':'')+n; }

/* === parse seguro de data/hora da planilha === */
function parseSheetDateTimeCell_(sheet, a1){
  const raw = sheet.getRange(a1).getValue();
  if (raw instanceof Date) return new Date(raw.getFullYear(), raw.getMonth(), raw.getDate(), raw.getHours(), raw.getMinutes(), 0, 0);
  const s = String(raw||'').trim(); if(!s) return null;
  let m;
  // yyyy-mm-dd HH:mm / yyyy/mm/dd HH:mm
  if((m = s.match(/^(\d{4})[-\/](\d{2})[-\/](\d{2})\s+(\d{1,2}):(\d{2})$/))) return new Date(+m[1],+m[2]-1,+m[3],+m[4],+m[5],0,0);
  // dd/mm/yyyy HH:mm
  if((m = s.match(/^(\d{2})[\/\-](\d{2})[\/\-](\d{4})\s+(\d{1,2}):(\d{2})$/))) return new Date(+m[3],+m[2]-1,+m[1],+m[4],+m[5],0,0);
  // yyyy-mm-dd
  if((m = s.match(/^(\d{4})[-\/](\d{2})[-\/](\d{2})$/))) return new Date(+m[1],+m[2]-1,+m[3],0,0,0,0);
  // dd/mm/yyyy
  if((m = s.match(/^(\d{2})[\/\-](\d{2})[\/\-](\d{4})$/))) return new Date(+m[3],+m[2]-1,+m[1],0,0,0,0);
  const d = new Date(s); return isNaN(d)?null:d;
}
/** Conta aulas pós-cutover por aluno (norm) a partir da agenda, considerando ✅ e cores válidas */
function countPosCutByNorm_(cfg){
  const cal = CalendarApp.getCalendarById(CALENDAR_ID);
  const start = cfg.cutoverDateTime;                         // data/hora fixa do cutover
  const end   = new Date(); end.setHours(23,59,59,999);
  const events = cal.getEvents(start, end);

  const posByNorm = {};
  events.forEach(ev=>{
    const color = ev.getColor() || "";
    if (cfg.excludedColors.includes(color)) return;
    const title = ev.getTitle()||'', desc = ev.getDescription()||'';
    if (!CHECK_PAT.test(title) && !CHECK_PAT.test(desc)) return; // só aula dada
    const raw  = extractSingleNameFromTitle_(title);
    const norm = normalizeName_(raw);
    if(!norm) return;
    posByNorm[norm] = (posByNorm[norm]||0) + 1;
  });
  return posByNorm;
}

/** Lê contratos ativos com campos essenciais */
function getContractsData_(){
  const sh = ensureContratosSheet_();
  const lr = sh.getLastRow(); if(lr<2) return [];
  const vals = sh.getRange(1,1,lr,sh.getLastColumn()).getValues();
  const head = vals[0]; const idx = Object.fromEntries(head.map((h,i)=>[h,i]));

  // qual coluna é o PRE-CUTOVER?
  const preKey = idx.hasOwnProperty('AULAS_FEITAS_ATE_CUTOVER') ? 'AULAS_FEITAS_ATE_CUTOVER'
               : idx.hasOwnProperty('AULAS_PRE_CUTOVER')       ? 'AULAS_PRE_CUTOVER'
               : null;

  const out = [];
  for(let r=1;r<vals.length;r++){
    const row = vals[r];
    if(String(row[idx['STATUS']]||'').toLowerCase()!=='ativo') continue;
    const id   = row[idx['ID_CONTRATO']];
    const aluno= row[idx['ALUNO']];
    const norm = String(row[idx['NOME_NORMALIZADO']]||'').trim() || normalizeName_(aluno);
    if(!id || !norm) continue;

    const qt   = Number(row[idx['QTDE_AULAS_CONTR']]||0);
    const pre  = preKey ? Number(row[idx[preKey]]||0) : 0;
    const di   = row[idx['DATA_INICIO']] ? new Date(row[idx['DATA_INICIO']]) : new Date(1900,0,1);
    const dsc  = parseDiscountValueString_(row[idx['DESCONTO_%']]);
    const mod  = (String(row[idx['MODALIDADE_CONTR']]||'i').toLowerCase()==='d')?'d':'i';

    out.push({
      id, aluno, norm, qt, pre, dtInicio: di, desconto: (isNaN(dsc)?0:dsc), modalidade: mod
    });
  }
  // agrupa por aluno e ordena por DATA_INICIO asc
  out.sort((a,b)=> (a.norm===b.norm ? a.dtInicio - b.dtInicio : a.norm.localeCompare(b.norm)));
  return out;
}

/** Distribui aulas POS-CUTOVER por contrato do mesmo aluno, respeitando saldo de cada contrato */
function allocatePosToContracts_(posByNorm, contracts){
  // contratos já estão ordenados por (norm, dtInicio)
  const byId = {};
  let i = 0;
  while(i < contracts.length){
    const start = i, currNorm = contracts[i].norm;
    while(i < contracts.length && contracts[i].norm === currNorm) i++;
    const slice = contracts.slice(start, i); // todos os contratos desse aluno
    let pos = Number(posByNorm[currNorm]||0);

    for(const c of slice){
      const saldo = Math.max(0, Number(c.qt||0) - Number(c.pre||0));
      const entrega = Math.min(pos, saldo);
      byId[c.id] = entrega;
      pos -= entrega;
    }
  }
  return { byId };
}
function extractNamesFromTitleMulti_(title) {
  let t = (title||'')
    .replace(/dupla[:\-]?/i,'')
    .replace(/atendimento[:\-]?/i,'')
    .replace(/exp[:\-]?/i,'')
    .replace(CHECK_PAT,'')
    .trim();

  t = t.replace(/\b\d{1,2}[:h]\d{0,2}\b/gi,'')
       .replace(/\b\d{1,2}\b/g,'')
       .trim();

  t = t.replace(/\se\s/gi, ','); // "A e B" → "A,B"
  const parts = t.split(/[,|+|&|\-|–|—|\/]+/).map(s => s.trim()).filter(Boolean);

  const seen = new Set(), out = [];
  for (const p of parts) {
    const norm = normalizeName_(p);
    if (norm && !seen.has(norm)) { seen.add(norm); out.push({display:p, norm}); }
  }
  if (!out.length) {
    const norm = normalizeName_(t);
    if (norm) out.push({display:t, norm});
  }
  return out.slice(0, 4);
}
function countPosCutByNorm_(cfg){
  const cal   = CalendarApp.getCalendarById(CALENDAR_ID);
  const start = cfg.cutoverDateTime;
  const end   = new Date(); end.setHours(23,59,59,999);

  const posByNorm = {};
  try{
    const events = cal.getEvents(start, end);
    events.forEach(ev=>{
      const color = ev.getColor() || "";
      if (cfg.excludedColors.includes(color)) return;

      const title = ev.getTitle()||'', desc = ev.getDescription()||'';
      if (!CHECK_PAT.test(title) && !CHECK_PAT.test(desc)) return; // só aula dada (✅)

      extractNamesFromTitleMulti_(title).forEach(n=>{
        posByNorm[n.norm] = (posByNorm[n.norm]||0) + 1; // dupla conta para os 2
      });
    });
  }catch(e){ /* se Calendar falhar, cai no fallback */ }

  // se Agenda vier vazia, usa as abas mensais (Out, Nov...) como fallback
  const empty = Object.keys(posByNorm).length===0 || Object.values(posByNorm).every(v=>v===0);
  if (empty){
    const fromSheets = countPosCutByNormFromSheets_(cfg);
    Object.keys(fromSheets).forEach(k=> posByNorm[k] = (posByNorm[k]||0) + fromSheets[k]);
  }
  return posByNorm;
}
// Desconto "padrão" quando a coluna DESCONTO_% / DESCONTO_NOVO vier vazia
function resolveDefaultDiscount_(plano, freqSemanal, cfg){
  const p = String(plano||'').toLowerCase();
  const f = Number(freqSemanal||1);

  // Mensalidade: 1x = 5% ; 2x = 10%
  if (p.includes('mensal')) {
    if (f >= 2) return 0.10;
    return 0.05;
  }

  // Planos fechados (fallbacks conhecidos)
  if (p.includes('12') && p.includes('aula')) return 0.15; // 12 aulas
  if (p.includes('6')  && p.includes('mes'))  return 0.25; // 6 meses
  if (p.includes('12') && p.includes('mes'))  return 0.35; // 12 meses

  // Se tiver tabela cfg.descontos[…], mantém compatibilidade
  if (cfg && cfg.descontos && typeof cfg.descontos[p] === 'number') return cfg.descontos[p];

  return 0;
}


/* ===== DESCONTO ===== */
function parseDiscountValueString_(s){
  if(s===null || s===undefined) return NaN;
  let t = String(s).trim().replace(',', '.');
  if(!t) return NaN;
  if(t.endsWith('%')) t = t.slice(0,-1).trim();
  const n = Number(t);
  if(isNaN(n)) return NaN;
  return n > 1 ? n/100 : n;
}
function parseDiscountFromCell_(val, numFmt){
  if(typeof val === 'number'){
    if(numFmt && /%/.test(numFmt)){
      if(val > 0 && val < 0.01) return val * 100; // 0.0025 -> 0.25
      return val;
    }
    return val > 1 ? val/100 : val;
  }
  return parseDiscountValueString_(val);
}
function coerceDiscountMagnitude_(d, plano, cfg){
  if(typeof d !== 'number' || isNaN(d)) return d;
  const eps = 0.001; const near = (x,y)=> Math.abs(x-y) <= eps;
  const targets = [0,0.05,0.15,0.25,0.35];
  if(d > 0 && d < 0.01){ const d100=d*100; for(const t of targets){ if(near(d100,t)) return t; } return d100; }
  const tab = cfg?.descontos?.[(plano||'').toLowerCase()];
  if(typeof tab==='number' && !isNaN(tab) && d < tab/10) return tab;
  return d;
}
function mapDiscountToPlan_(d){
  if(typeof d!=='number' || isNaN(d)) return null;
  const eps=0.005, close=(x,y)=>Math.abs(x-y)<=eps;
  if(close(d,0)) return {plano:'Avulsa',meses:0};
  if(close(d,0.05)) return {plano:'Mensalidade',meses:1};
  if(close(d,0.15)) return {plano:'12 aulas',meses:0};
  if(close(d,0.25)) return {plano:'6 meses',meses:6};
  if(close(d,0.35)) return {plano:'12 meses',meses:12};
  return null;
}

/* ===== CONFIG ===== */
function readConfigFin_(){
  const sh = SpreadsheetApp.getActive().getSheetByName(SH_CFG_FIN);
  if(!sh) throw new Error(`Crie a aba ${SH_CFG_FIN}.`);

  const vals = sh.getDataRange().getValues();
  const descontos = {};
  for(let i=1;i<vals.length;i++){
    const plano = String(vals[i][0]||'').trim(); if(!plano) continue;
    const dPct = (String(vals[i][1]||'').toString().replace('%',''));
    descontos[plano.toLowerCase()] = Number(dPct)/100 || 0;
  }

  const precoIndividual = Number(sh.getRange('D1').getValue());
  const precoDuo        = Number(sh.getRange('E1').getValue());
  const semanasMes      = Number(sh.getRange('F1').getValue()) || 4.00;
  const presenceStart   = parseSheetDateTimeCell_(sh,'G1');
  const planDefault     = String(sh.getRange('H1').getValue()||'Mensalidade').trim();
  const indivCfg        = String(sh.getRange('I1').getValue()||'').trim();
  const duoCfg          = String(sh.getRange('J1').getValue()||'1').trim();
  const exclCfg         = String(sh.getRange('K1').getValue()||'5').trim();
  const cutoverDT       = parseSheetDateTimeCell_(sh,'L1') || new Date(2025,9,1,20,0,0,0); // 01/10/2025 20:00
  const zeroHMes        = String(sh.getRange('M1').getValue()||'NAO').toUpperCase() === 'SIM';

  const individualColors = indivCfg ? indivCfg.split(',').map(s=>s.trim()) : [""];
  const duoColor         = duoCfg || "1";
  const excludedColors   = exclCfg ? exclCfg.split(',').map(s=>s.trim()) : ["5"];

  return { descontos, precoIndividual, precoDuo, semanasMes,
           presenceStartDate: presenceStart, planDefault,
           individualColors, duoColor, excludedColors,
           cutoverDateTime: cutoverDT, zeroHMes };
}

/* ===== ESTRUTURA INICIAL ===== */
function createStudioStructure(){
  const ss = SpreadsheetApp.getActive();
  ['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro']
  .forEach(m=>{
    let sh = ss.getSheetByName(m); if(!sh) sh = ss.insertSheet(m);
    sh.getRange('A1:H9').clearContent();
    sh.getRange('D10').setValue('NOME');
    sh.getRange('E10').setValue('MODALIDADE (i/d)');
    sh.getRange('H10').setValue('QTD_MÊS');
  });

  let shCfg = ss.getSheetByName(SH_CFG_FIN); if(!shCfg) shCfg = ss.insertSheet(SH_CFG_FIN);
  if(shCfg.getLastRow()===0){
    shCfg.getRange('A1:B1').setValues([['PLANO','DESCONTO_%']]);
    shCfg.getRange('A2:B6').setValues([
      ['Avulsa','0%'],['Mensalidade','5%'],['12 aulas','15%'],['6 meses','25%'],['12 meses','35%']
    ]);
  }
  if(!ss.getSheetByName(SH_CONTR)) ss.insertSheet(SH_CONTR).getRange(1,1,1,CONTR_HEADERS.length).setValues([CONTR_HEADERS]);
  if(!ss.getSheetByName(SH_RECON)) ss.insertSheet(SH_RECON).getRange(1,1,1,11).setValues([[
    'ID_CONTRATO','ALUNO','PLANO','MODALIDADE','AAAA-MM',
    'AULAS_ENTREGUES_NO_MÊS','PRECO_UNIT_APLICADO','RECEITA_RECONHECIDA',
    'AULAS_ACUM_ENTREGUES','AULAS_SALDO','VALOR_SALDO'
  ]]);
  if(!ss.getSheetByName(SH_ROLLFWD)) ss.insertSheet(SH_ROLLFWD).getRange(1,1,1,5).setValues([[
    'AAAA-MM','SALDO_INICIAL_PASSIVO','(+ ) NOVOS CONTRATOS (aprox.)','(-) RECEITA RECONHECIDA','SALDO_FINAL_PASSIVO'
  ]]);
  notify_('Estrutura ok ✅');
}

/* ===== AGENDA → ABAS MENSAIS (i/d) ===== */
function extractSingleNameFromTitle_(title){
  let t = (title||'').replace(/dupla[:\-]?/i,'').replace(/atendimento[:\-]?/i,'').replace(/exp[:\-]?/i,'').replace(CHECK_PAT,'').trim();
  t = t.split(/[,|+|&|\-|–|—]/)[0];
  t = t.replace(/\b\d{1,2}[:h]\d{0,2}\b/gi,'').replace(/\b\d{1,2}\b/g,'').trim();
  return t;
}
function runUpdateAllMonths(){
  const ss  = SpreadsheetApp.getActive();
  const cal = CalendarApp.getCalendarById(CALENDAR_ID);
  const cfg = readConfigFin_(); // usa cores, presenceStartDate e zeroHMes

  const year  = new Date().getFullYear();
  const start = new Date(year, 0, 1, 0, 0, 0);
  const end   = new Date(year + 1, 0, 1, 0, 0, 0);
  const events = cal.getEvents(start, end);

  // monthCounters[mes][norm] = { display, i, d }
  const monthCounters = {
    'Janeiro': {}, 'Fevereiro': {}, 'Março': {}, 'Abril': {}, 'Maio': {}, 'Junho': {},
    'Julho': {}, 'Agosto': {}, 'Setembro': {}, 'Outubro': {}, 'Novembro': {}, 'Dezembro': {}
  };

  let processed = 0;

  events.forEach(ev => {
    // ignora eventos antes do marco de presença (se configurado)
    if (cfg.presenceStartDate && ev.getStartTime() < cfg.presenceStartDate) return;

    // filtra por cor (modalidade) e excluídos
    const color = ev.getColor() || "";
    if (cfg.excludedColors.includes(color)) return;
    const modality = cfg.individualColors.includes(color)
      ? 'i'
      : (color === cfg.duoColor ? 'd' : null);
    if (!modality) return;

    // só conta presença (✅ no título ou na descrição)
    const title = ev.getTitle() || '';
    const desc  = ev.getDescription() || '';
    if (!CHECK_PAT.test(title) && !CHECK_PAT.test(desc)) return;

    const monthName = getMonthNamePt_(ev.getStartTime().getMonth());

    // 👇 NOVO: suporta múltiplos nomes no título (dupla conta para os 2)
    const names = extractNamesFromTitleMulti_(title); // [{display, norm}, ...]

    names.forEach(n => {
      if (!n.norm) return;
      if (!monthCounters[monthName][n.norm]) {
        monthCounters[monthName][n.norm] = { display: n.display, i: 0, d: 0 };
      } else {
        monthCounters[monthName][n.norm].display =
          pickBetterDisplay_(monthCounters[monthName][n.norm].display, n.display);
      }
      monthCounters[monthName][n.norm][modality] += 1;
      processed++;
    });
  });

  // escreve nas abas (zera H só se cfg.zeroHMes = true)
  Object.keys(monthCounters).forEach(monthName => {
    const sheet = ss.getSheetByName(monthName);
    if (!sheet) return;
    writeCountsToSheet_(sheet, monthCounters[monthName], cfg.zeroHMes);
  });

  notify_(`Eventos com presença processados: ${processed}`);
}

function writeCountsToSheet_(sheet, counterObj, zeroColumn){
  // se zeroColumn = true, zera a coluna H; caso contrário, preserva valores não tocados
  if (zeroColumn){
    const lastRow = sheet.getLastRow();
    if (lastRow >= FIRST_DATA_ROW) {
      const qtyRange = sheet.getRange(`${QTY_COL}${FIRST_DATA_ROW}:${QTY_COL}${lastRow}`);
      const qtyVals  = qtyRange.getValues();
      for (let i = 0; i < qtyVals.length; i++) qtyVals[i][0] = '';
      qtyRange.setValues(qtyVals);
    }
  }

  // índice existente (nome+mod)
  const lastRow = sheet.getLastRow();
  const index = new Map();
  if (lastRow >= FIRST_DATA_ROW) {
    const nameVals = sheet.getRange(`${NAME_COL}${FIRST_DATA_ROW}:${NAME_COL}${lastRow}`).getValues().flat();
    const modVals  = sheet.getRange(`${MODALITY_COL}${FIRST_DATA_ROW}:${MODALITY_COL}${lastRow}`).getValues().flat();
    for (let i = 0; i < nameVals.length; i++) {
      const nm  = normalizeName_(String(nameVals[i] || ''));
      const mod = String(modVals[i] || '').toLowerCase().trim();
      if (!nm || !mod) continue;
      index.set(`${nm}|${mod}`, FIRST_DATA_ROW + i);
    }
  }

  function upsertRow(displayName, norm, modality, qty){
    const key = `${norm}|${modality}`;
    let row = index.get(key);
    if (row) {
      sheet.getRange(`${QTY_COL}${row}`).setValue(qty);
      const nameCell = sheet.getRange(`${NAME_COL}${row}`);
      if (!String(nameCell.getValue()||'').trim()) nameCell.setValue(displayName || norm);
      return;
    }
    let insertRow = null;
    const currentLast = sheet.getLastRow();
    for (let r = FIRST_DATA_ROW; r <= currentLast; r++) {
      const v = sheet.getRange(`${NAME_COL}${r}`).getValue();
      const h = sheet.getRange(`${QTY_COL}${r}`).getValue();
      if (!v && !h) { insertRow = r; break; }
    }
    if (!insertRow) insertRow = currentLast + 1;
    sheet.getRange(`${NAME_COL}${insertRow}`).setValue(displayName || norm);
    sheet.getRange(`${MODALITY_COL}${insertRow}`).setValue(modality);
    sheet.getRange(`${QTY_COL}${insertRow}`).setValue(qty);
    index.set(key, insertRow);
  }

  Object.keys(counterObj).forEach(norm => {
    const { display, i = 0, d = 0 } = counterObj[norm] || {};
    if (i > 0) upsertRow(display, norm, 'i', i);
    if (d > 0) upsertRow(display, norm, 'd', d);
  });
}
function resetAllMonths(){
  const ss = SpreadsheetApp.getActive();
  ['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro']
  .forEach(m => {
    const sheet = ss.getSheetByName(m);
    if (!sheet) return;
    const lastRow = sheet.getLastRow();
    if (lastRow < FIRST_DATA_ROW) return;
    const rng = sheet.getRange(`${QTY_COL}${FIRST_DATA_ROW}:${QTY_COL}${lastRow}`);
    const vals = rng.getValues();
    for (let i = 0; i < vals.length; i++) vals[i][0] = '';
    rng.setValues(vals);
  });
  notify_('QTD_MÊS (H) zerado em todas as abas.');
}
function logColorsToday(){
  const cal = CalendarApp.getCalendarById(CALENDAR_ID);
  const start = new Date(); start.setHours(0,0,0,0);
  const end   = new Date(); end.setHours(23,59,59,999);
  cal.getEvents(start, end).forEach(ev => Logger.log(`Título: ${ev.getTitle()} | Cor: ${ev.getColor()}`));
}

/* ===== QUANTIDADE PELO PLANO ===== */
function aulasTotaisDoPlano_(plano,freq,mesesDur,semanasMes){
  plano = (plano||'').toLowerCase();
  const f = Number(freq||0), w = Number(semanasMes||4.00);
  if(plano==='avulsa'||plano==='avulso') return 1;
  if(plano==='12 aulas'||plano==='12aulas') return 12;
  if(plano==='mensalidade'){ const m = Number(mesesDur||1); return Math.max(1, Math.round(m*f*w)); }
  if(plano==='6 meses'||plano==='6meses')  return Math.max(1, Math.round(6*f*w));
  if(plano==='12 meses'||plano==='12meses')return Math.max(1, Math.round(12*f*w));
  if(Number(mesesDur)>0) return Math.max(1, Math.round(Number(mesesDur)*f*w));
  return Math.max(1, Math.round(f*w));
}

/* ===== CONTRATOS (autofill / recalcular) ===== */
function ensureContratosSheet_(){
  const sh = getSheet_(SH_CONTR);
  let lastCol = sh.getLastColumn();
  if(lastCol === 0){ sh.getRange(1,1,1,CONTR_HEADERS.length).setValues([CONTR_HEADERS]); return sh; }
  const head = sh.getRange(1,1,1,lastCol).getValues()[0];
  const have = new Set(head);
  const missing = CONTR_HEADERS.filter(h => !have.has(h));
  if(missing.length){ sh.getRange(1,lastCol+1,1,missing.length).setValues([missing]); }
  return sh;
}
function generateNextContractId_(){
  const sh = getSheet_(SH_CONTR);
  const lr = sh.getLastRow();
  const now=new Date(); const ym = now.getFullYear().toString()+pad2_(now.getMonth()+1);
  let maxSeq = 0;
  if(lr>=2){
    const vals = sh.getRange(2,1,lr-1,1).getValues().flat();
    vals.forEach(v=>{
      const m = String(v||'').match(/^C-(\d{6})-(\d{3})$/);
      if(m && m[1]===ym){ maxSeq = Math.max(maxSeq, Number(m[2])); }
    });
  }
  const seq = (maxSeq+1).toString().padStart(3,'0');
  return `C-${ym}-${seq}`;
}
function inferPresenceStatsForAllStudents_(){
  const cal = CalendarApp.getCalendarById(CALENDAR_ID);
  const cfg = readConfigFin_();
  const start = cfg.presenceStartDate || new Date(new Date().getFullYear(),0,1);
  const end   = new Date(); end.setHours(23,59,59,999);
  const events = cal.getEvents(start, end);
  const per = {};

  events.forEach(ev=>{
    if (cfg.presenceStartDate && ev.getStartTime() < cfg.presenceStartDate) return;
    const color = ev.getColor() || ""; if (cfg.excludedColors.includes(color)) return;
    const modality = cfg.individualColors.includes(color) ? 'i' : (color === cfg.duoColor ? 'd' : null);
    if(!modality) return;
    const title=(ev.getTitle()||''), desc=(ev.getDescription()||'');
    if(!CHECK_PAT.test(title) && !CHECK_PAT.test(desc)) return;
    const raw=extractSingleNameFromTitle_(title), norm=normalizeName_(raw); if(!norm) return;
    if(!per[norm]) per[norm]={display:raw, dates:[], firstMod:modality, firstDate:ev.getStartTime()};
    else { per[norm].display=pickBetterDisplay_(per[norm].display, raw);
           if(ev.getStartTime()<per[norm].firstDate){ per[norm].firstDate=ev.getStartTime(); per[norm].firstMod=modality; } }
    per[norm].dates.push(ev.getStartTime());
  });

  const out={};
  Object.keys(per).forEach(norm=>{
    const disp=per[norm].display, dates=per[norm].dates.sort((a,b)=>a-b);
    if(!dates.length) return; const first=dates[0];
    // frequência média na 1ª janela de 4 semanas
    const windowEnd = new Date(first.getTime() + 28*24*60*60*1000);
    const inWindow = dates.filter(d => d >= first && d < windowEnd);
    const weeks=[0,0,0,0]; inWindow.forEach(d=>{ const k=Math.floor((d-first)/(7*24*60*60*1000)); if(k>=0&&k<4) weeks[k]++; });
    const total = weeks.reduce((s,x)=>s+x,0); const f = Math.max(1, Math.min(3, Math.round(total/4.00)||1));
    out[norm]={display:disp, firstDate:first, freq:f, firstMod:per[norm].firstMod};
  });
  return out;
}
function fin_autopreencherContratos(){
  const shC = ensureContratosSheet_();
  const cfg = readConfigFin_();

  const lr = shC.getLastRow();
  const idx = Object.fromEntries(CONTR_HEADERS.map((h,i)=>[h,i]));
  const existing = new Map();
  if(lr>=2){
    const vals = shC.getRange(2,1,lr-1,CONTR_HEADERS.length).getValues();
    vals.forEach((r,i)=>{
      const norm = String(r[idx['NOME_NORMALIZADO']]||'').trim() || normalizeName_(r[idx['ALUNO']]||'');
      if(norm) existing.set(norm, 2+i);
    });
  }

  const stats = inferPresenceStatsForAllStudents_();
  const plan  = (cfg.planDefault||'Mensalidade').toLowerCase();
  const monthsDur = plan==='6 meses'||plan==='6meses' ? 6
                    : plan==='12 meses'||plan==='12meses' ? 12
                    : plan==='mensalidade' ? 1
                    : plan==='12 aulas'||plan==='12aulas' ? 0
                    : plan==='avulsa'||plan==='avulso' ? 0 : 1;

  const rowsToAppend = []; const rowsToPatch  = [];

  Object.keys(stats).forEach(norm=>{
    const { display, firstDate, freq, firstMod } = stats[norm];
    if(!firstDate) return;

    const aulasTot = aulasTotaisDoPlano_(plan, freq, monthsDur, cfg.semanasMes);

    if(existing.has(norm)){
      const rowNum = existing.get(norm);
      const rowVals = shC.getRange(rowNum,1,1,CONTR_HEADERS.length).getValues()[0];

      if(!rowVals[idx['ALUNO']])              rowVals[idx['ALUNO']] = display;
      if(!rowVals[idx['NOME_NORMALIZADO']])   rowVals[idx['NOME_NORMALIZADO']] = norm;
      if(!rowVals[idx['PLANO']])              rowVals[idx['PLANO']] = cfg.planDefault;
      if(!rowVals[idx['FREQUENCIA_SEMANAL']]) rowVals[idx['FREQUENCIA_SEMANAL']] = freq;
      if(!rowVals[idx['MESES_DURACAO']])      rowVals[idx['MESES_DURACAO']] = monthsDur;
      if(!rowVals[idx['PRECO_CHEIO_AULA']])   rowVals[idx['PRECO_CHEIO_AULA']] = cfg.precoIndividual;
      if(!rowVals[idx['QTDE_AULAS_CONTR']])   rowVals[idx['QTDE_AULAS_CONTR']] = aulasTot;
      if(!rowVals[idx['STATUS']])             rowVals[idx['STATUS']] = 'ativo';
      if(!rowVals[idx['MODALIDADE_CONTR']])   rowVals[idx['MODALIDADE_CONTR']] = firstMod;
      if(!rowVals[idx['DATA_INICIO']])        rowVals[idx['DATA_INICIO']] = new Date(firstDate.getFullYear(), firstDate.getMonth(), 1);
      rowsToPatch.push({row: rowNum, values: rowVals});
    } else {
      const id = generateNextContractId_();
      const newRow = new Array(CONTR_HEADERS.length).fill('');
      newRow[idx['ID_CONTRATO']]       = id;
      newRow[idx['DATA_INICIO']]       = new Date(firstDate.getFullYear(), firstDate.getMonth(), 1);
      newRow[idx['ALUNO']]             = display;
      newRow[idx['NOME_NORMALIZADO']]  = norm;
      newRow[idx['PLANO']]             = cfg.planDefault;
      newRow[idx['FREQUENCIA_SEMANAL']]= freq;
      newRow[idx['MESES_DURACAO']]     = monthsDur;
      newRow[idx['PRECO_CHEIO_AULA']]  = cfg.precoIndividual;
      newRow[idx['QTDE_AULAS_CONTR']]  = aulasTot;
      newRow[idx['STATUS']]            = 'ativo';
      newRow[idx['RENOVAR']]           = 'NÃO';
      newRow[idx['MODALIDADE_CONTR']]  = firstMod;
      rowsToAppend.push(newRow);
    }
  });

  rowsToPatch.forEach(p=> shC.getRange(p.row,1,1,CONTR_HEADERS.length).setValues([p.values]));
  if(rowsToAppend.length){
    shC.getRange(shC.getLastRow()+1,1,rowsToAppend.length,CONTR_HEADERS.length).setValues(rowsToAppend);
  }
  notify_(`Autopreenchimento: ${rowsToPatch.length} atualizados, ${rowsToAppend.length} novos.`);
}
function fin_recalcularContratos(){
  const sh = ensureContratosSheet_(); const cfg = readConfigFin_();
  const lr = sh.getLastRow(); if(lr<2){ notify_('Aba Contratos vazia.'); return; }

  const rng  = sh.getRange(1,1,lr,CONTR_HEADERS.length);
  const vals = rng.getValues(); const fms  = rng.getNumberFormats();
  const head = vals[0]; const idx  = Object.fromEntries(head.map((h,i)=>[h,i]));
  const planDefaultLc = (cfg.planDefault||'').toLowerCase();

  for(let r=1;r<vals.length;r++){
    const row = vals[r], fmt = fms[r];
    const aluno = row[idx['ALUNO']];
    row[idx['NOME_NORMALIZADO']] = normalizeName_(aluno);
    row[idx['PRECO_CHEIO_AULA']] = cfg.precoIndividual;

    let plano = String(row[idx['PLANO']]||'').trim();
    let freq  = Number(row[idx['FREQUENCIA_SEMANAL']]||0);
    let meses = Number(row[idx['MESES_DURACAO']]||0);
    let modC  = String(row[idx['MODALIDADE_CONTR']]||'').toLowerCase().trim(); // i/d

    let dsc = parseDiscountFromCell_(row[idx['DESCONTO_%']], fmt[idx['DESCONTO_%']]);
    dsc = coerceDiscountMagnitude_(dsc, plano, cfg);

    const mapped = (!isNaN(dsc)) ? mapDiscountToPlan_(dsc) : null;
    if(mapped && (plano==='' || plano.toLowerCase()===planDefaultLc)){
      plano = mapped.plano; row[idx['PLANO']] = plano;
      if(!meses){ meses = mapped.meses; row[idx['MESES_DURACAO']] = meses; }
    }

    if(!plano){ plano = cfg.planDefault; row[idx['PLANO']] = plano; }
    if(!freq){ freq = 1; row[idx['FREQUENCIA_SEMANAL']] = 1; }
    if(!modC){ modC = 'i'; row[idx['MODALIDADE_CONTR']] = 'i'; }

    if(isNaN(dsc)) dsc = cfg.descontos[(plano||'').toLowerCase()] ?? 0;

    const basePrice = (modC==='d') ? cfg.precoDuo : cfg.precoIndividual;
    row[idx['DESCONTO_%']] = dsc;
    row[idx['PRECO_UNIT']] = basePrice * (1 - dsc);

    // respeita se você já digitou a quantidade
    if(!row[idx['QTDE_AULAS_CONTR']] || Number(row[idx['QTDE_AULAS_CONTR']])<=0){
      row[idx['QTDE_AULAS_CONTR']] = aulasTotaisDoPlano_(plano, freq, meses, cfg.semanasMes);
    }
    if(!row[idx['STATUS']]) row[idx['STATUS']] = 'ativo';
  }

  rng.setValues(vals);
  const colDsc = idx['DESCONTO_%'] + 1;
  if(lr>1) sh.getRange(2, colDsc, lr-1, 1).setNumberFormat('0.00%');
  notify_('Contratos recalculados ✅');
}

/* ===== RECONHECIMENTO / AJUSTES / ROLLFWD (igual antes) ===== */
// ... (mantém suas funções fin_reconhecerReceitaMesAtual, fin_reconhecerReceitaMes,
// fin_aplicarAjustesPreBase e fin_gerarRollforward do seu último código)

/* ===== Correção rápida da coluna de descontos ===== */
function fin_corrigirDescontos(){
  const sh=ensureContratosSheet_(); const cfg=readConfigFin_(); const lr=sh.getLastRow(); if(lr<2){ notify_('Aba Contratos vazia.'); return; }
  const rng=sh.getRange(1,1,lr,CONTR_HEADERS.length), vals=rng.getValues(), fms=rng.getNumberFormats();
  const head=vals[0], idx=Object.fromEntries(head.map((h,i)=>[h,i]));
  for(let r=1;r<vals.length;r++){
    const row=vals[r], fmt=fms[r], plano=String(row[idx['PLANO']]||'').trim();
    let dsc=parseDiscountFromCell_(row[idx['DESCONTO_%']], fmt[idx['DESCONTO_%']]); dsc=coerceDiscountMagnitude_(dsc,plano,cfg);
    if(isNaN(dsc)&&plano){ const tab=cfg.descontos[(plano||'').toLowerCase()]; if(typeof tab==='number') dsc=tab; }
    row[idx['DESCONTO_%']]=dsc;
  }
  rng.setValues(vals); const colD=idx['DESCONTO_%']+1; if(lr>1) sh.getRange(2,colD,lr-1,1).setNumberFormat('0.00%');
  notify_('Descontos corrigidos na coluna ✅');
}

/* ===== PROGRESSO POR QUANTIDADE — **APÓS CUTOVER (por EVENTO)** ===== */
function fin_atualizarProgressoContratos(){
  const cfg = readConfigFin_();
  const shC = ensureContratosSheet_();
  const lr  = shC.getLastRow();
  if (lr < 2) { notify_('Sem contratos.'); return; }

  // cabeçalho/índices
  const vals = shC.getRange(1,1,lr,shC.getLastColumn()).getValues();
  const head = vals[0];
  const idx  = Object.fromEntries(head.map((h,i)=>[h,i]));

  const colId     = idx['ID_CONTRATO'];
  const colAluno  = idx['ALUNO'];
  const colNorm   = idx['NOME_NORMALIZADO'];
  const colQt     = idx['QTDE_AULAS_CONTR'];
  const colPre1   = idx['AULAS_FEITAS_ATE_CUTOVER'];
  const colPre2   = idx['AULAS_PRE_CUTOVER'];
  const colPos    = idx['AULAS_POS_CUTOVER'];
  const colAtual  = idx['AULA_ATUAL'];
  const colRest   = idx['AULAS_RESTANTES'];
  const colStatus = idx['STATUS'];
  const colIni    = idx['DATA_INICIO'];

  if ([colPos,colAtual,colRest].some(v => v === undefined)) {
    notify_('Faltam colunas: AULAS_POS_CUTOVER / AULA_ATUAL / AULAS_RESTANTES.');
    return;
  }

  // 1) pós-cutover por aluno (Agenda com fallback nas abas)
  const posByNorm = countPosCutByNorm_(cfg);

  // 2) prepara grupos por aluno
  /** @type {Record<string, Array<{r:number, qt:number, pre:number, saldo:number, dt:Date}>>} */
  const groups = {};
  for (let r = 1; r < vals.length; r++) {
    const row = vals[r];
    if (!row[colId]) continue;
    if (String(row[colStatus] || '').toLowerCase() !== 'ativo') continue;

    let norm = String(row[colNorm] || '').trim();
    if (!norm) norm = normalizeName_(row[colAluno]);
    if (!norm) continue;

    const qt   = Number(row[colQt] || 0);
    const pre  = Number(colPre1 != null ? (row[colPre1] || 0) : (colPre2 != null ? (row[colPre2] || 0) : 0));
    const dt   = row[colIni] instanceof Date ? row[colIni] : (row[colIni] ? new Date(row[colIni]) : new Date(1900,0,1));
    const saldo= Math.max(0, qt - pre);

    if (!groups[norm]) groups[norm] = [];
    groups[norm].push({ r, qt, pre, saldo, dt });
  }

  // 3) zera colunas de progresso antes de distribuir (evita lixo antigo)
  for (let r = 1; r < vals.length; r++) {
    vals[r][colPos]   = 0;
    const qt = Number(vals[r][colQt] || 0);
    const pre= Number(colPre1 != null ? (vals[r][colPre1] || 0) : (colPre2 != null ? (vals[r][colPre2] || 0) : 0));
    const atual = Math.min(qt, pre);
    vals[r][colAtual] = atual;
    vals[r][colRest]  = Math.max(0, qt - atual);
  }

  // 4) distribui aulas pós-cutover por aluno, do contrato mais antigo para o mais novo
  let wrote = 0; const leftovers = [];
  Object.keys(groups).forEach(norm => {
    let pos = Number(posByNorm[norm] || 0);
    if (pos <= 0) return;

    const list = groups[norm]
      .filter(c => c.saldo > 0)
      .sort((a,b) => a.dt - b.dt);

    for (const c of list) {
      if (pos <= 0) break;
      const take = Math.min(c.saldo, pos);
      vals[c.r][colPos]   = take;

      const atual = Math.min(c.qt, c.pre + take);
      vals[c.r][colAtual] = atual;
      vals[c.r][colRest]  = Math.max(0, c.qt - atual);

      pos -= take;
      wrote++;
    }
    if (pos > 0) leftovers.push(`${norm}: ${pos}`);
  });

  // 5) grava
  shC.getRange(1,1,lr,shC.getLastColumn()).setValues(vals);
  notify_(`Progresso atualizado ✅ | contratos com POS>0: ${wrote}` + (leftovers.length ? ` | sobras não alocadas: ${leftovers.join(', ')}` : ''));
}



/* ===== AÇÕES DE RENOVAÇÃO ===== */
function fin_zerarProgressoSelecao(){
  const sh = ensureContratosSheet_();
  const sel = sh.getActiveRange(); 
  if(!sel){ notify_('Selecione as linhas dos contratos.'); return; }

  const head = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const idx  = Object.fromEntries(head.map((h,i)=>[h,i]));

  const vals = sh.getRange(sel.getRow(),1,sel.getNumRows(),sh.getLastColumn()).getValues();

  let n=0;
  vals.forEach(row=>{
    if(!row[idx['ID_CONTRATO']]) return;
    const qtd = Number(row[idx['QTDE_AULAS_CONTR']]||0);

    if (idx.hasOwnProperty('AULAS_FEITAS_ATE_CUTOVER')) row[idx['AULAS_FEITAS_ATE_CUTOVER']] = 0;
    if (idx.hasOwnProperty('AULAS_PRE_CUTOVER'))        row[idx['AULAS_PRE_CUTOVER']]        = 0;
    if (idx.hasOwnProperty('AULAS_POS_CUTOVER'))        row[idx['AULAS_POS_CUTOVER']]        = 0;

    row[idx['AULA_ATUAL']]      = 0;
    row[idx['AULAS_RESTANTES']] = qtd;
    n++;
  });

  sh.getRange(sel.getRow(),1,sel.getNumRows(),sh.getLastColumn()).setValues(vals);
  notify_(`Zerado progresso de ${n} contrato(s).`);
}

// marca RENOVAR=SIM p/ mensalidade que zerou saldo
function fin_sugerirRenovacoesMensalidade(){
  const sh = ensureContratosSheet_(); const lr = sh.getLastRow(); if(lr<2){ notify_('Sem contratos.'); return; }
  const vals = sh.getRange(1,1,lr,sh.getLastColumn()).getValues();
  const head = vals[0]; const idx = Object.fromEntries(head.map((h,i)=>[h,i]));
  let marcados=0;
  for(let r=1;r<vals.length;r++){
    const plano = String(vals[r][idx['PLANO']]||'').toLowerCase();
    const rest  = Number(vals[r][idx['AULAS_RESTANTES']]||0);
    if(plano==='mensalidade' && rest===0) { vals[r][idx['RENOVAR']]='SIM'; marcados++; }
  }
  sh.getRange(1,1,lr,sh.getLastColumn()).setValues(vals);
  notify_(`Sugeridas ${marcados} renovações de Mensalidade (RENOVAR=SIM).`);
}
function fin_reconhecerGERALTodosPlanos(){
  const cfg = readConfigFin_();
  const shR = getSheet_(SH_RECON);
  const shC = ensureContratosSheet_();
  const lr  = shC.getLastRow(); if(lr<2){ notify_('Sem contratos.'); return; }

  // cabeçalho padrão (reutilizamos a mesma estrutura)
  const header=['ID_CONTRATO','ALUNO','PLANO','MODALIDADE','AAAA-MM','AULAS_ENTREGUES_NO_MÊS','PRECO_UNIT_APLICADO','RECEITA_RECONHECIDA','AULAS_ACUM_ENTREGUES','AULAS_SALDO','VALOR_SALDO'];
  shR.clear(); shR.getRange(1,1,1,header.length).setValues([header]);

  const vals = shC.getRange(1,1,lr,shC.getLastColumn()).getValues();
  const head = vals[0]; const idx = Object.fromEntries(head.map((h,i)=>[h,i]));

  const preKey = idx.hasOwnProperty('AULAS_FEITAS_ATE_CUTOVER') ? 'AULAS_FEITAS_ATE_CUTOVER'
               : idx.hasOwnProperty('AULAS_PRE_CUTOVER')       ? 'AULAS_PRE_CUTOVER'
               : null;

  // pós-cutover por aluno (agenda) + alocação por contrato
  const posByNorm = countPosCutByNorm_(cfg);
  const contracts = getContractsData_();
  const alloc     = allocatePosToContracts_(posByNorm, contracts);

  const out = [];
  for(let r=1;r<vals.length;r++){
    const row = vals[r];
    if(String(row[idx['STATUS']]||'').toLowerCase()!=='ativo') continue;

    const id    = row[idx['ID_CONTRATO']];
    const aluno = row[idx['ALUNO']];
    const plano = row[idx['PLANO']];
    const norm  = String(row[idx['NOME_NORMALIZADO']]||'').trim() || normalizeName_(aluno);
    const qt    = Number(row[idx['QTDE_AULAS_CONTR']]||0);
    const pre   = preKey ? Number(row[idx[preKey]]||0) : 0;
    const pos   = Number(alloc.byId[id]||0);

    const totalEntregues = Math.min(qt, pre + pos);
    const saldoAulas     = Math.max(0, qt - totalEntregues);

    const dsc  = parseDiscountValueString_(row[idx['DESCONTO_%']]);
    const mod  = (String(row[idx['MODALIDADE_CONTR']]||'i').toLowerCase()==='d')?'d':'i';
    const pRef = (mod==='d' ? cfg.precoDuo : cfg.precoIndividual) * (1 - (isNaN(dsc)?0:dsc));

    out.push([
      id, aluno, plano, mod, 'TOTAL',
      totalEntregues,                      // usamos a mesma coluna "AULAS_ENTREGUES_NO_MÊS"
      pRef,
      totalEntregues * pRef,               // RECEITA_RECONHECIDA (total até aqui)
      totalEntregues,                      // AULAS_ACUM_ENTREGUES
      saldoAulas,                          // AULAS_SALDO
      saldoAulas * pRef                    // VALOR_SALDO
    ]);
  }

  if(out.length) shR.getRange(shR.getLastRow()+1,1,out.length,header.length).setValues(out);
  notify_(`Reconhecimento GERAL gerado ✅ (${out.length} contratos).`);
}
function fin_diagPosCutover(){
  const cfg = readConfigFin_();
  const pos = countPosCutByNorm_(cfg);
  const lines = Object.entries(pos).sort((a,b)=>b[1]-a[1]).slice(0,30)
                .map(([n,c])=> `${n}: ${c}`).join('\n');
  notify_(`Cutover: ${cfg.cutoverDateTime}\nTop alunos pós-cutover:\n${lines || '(nenhum encontrado)'}`);
}
function fin_criarRenovacaoLinhasSelecionadas(){
  const ui  = SpreadsheetApp.getUi();
  const cfg = readConfigFin_();
  const sh  = ensureContratosSheet_();
  const sel = sh.getActiveRange();
  if (!sel) { ui.alert('Selecione ao menos uma linha.'); return; }

  const lr   = sh.getLastRow();
  const head = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const idx  = Object.fromEntries(head.map((h,i)=>[h,i]));
  const gi = (n) => { const i = idx[n]; if (i === undefined) throw new Error('Coluna faltando: '+n); return i; };

  // colunas base
  const cID     = gi('ID_CONTRATO');
  const cAluno  = gi('ALUNO');
  const cNorm   = gi('NOME_NORMALIZADO');
  const cPlano  = gi('PLANO');
  const cFreq   = gi('FREQUENCIA_SEMANAL');
  const cMeses  = gi('MESES_DURACAO');
  const cMod    = gi('MODALIDADE_CONTR');
  const cIni    = gi('DATA_INICIO');
  const cDesc   = gi('DESCONTO_%');
  const cPreco  = gi('PRECO_CHEIO_AULA');
  const cUnit   = gi('PRECO_UNIT');
  const cQt     = gi('QTDE_AULAS_CONTR');
  const cStatus = gi('STATUS');

  // campos “NOVO”
  const cDataNovo  = idx['DATA_INICIO_NOVO'];
  const cPlanoNovo = idx['PLANO_NOVO'];
  const cFreqNova  = idx['FREQ_NOVA'];
  const cMesesNovo = (idx['MESES_DUR_N'] ?? idx['MESES_DUR_NOVO'] ?? idx['MESES_DURACAO_N']);
  const cModNova   = idx['MOD_NOVA'];
  const cDescNovo  = idx['DESCONTO_NOVO'];

  // vínculos
  const cRenDe   = idx['ID_RENOVACAO_DE'];
  const cRenPara = idx['ID_RENOVACAO_PARA'];

  // progresso
  const cPre1  = idx['AULAS_FEITAS_ATE_CUTOVER'];
  const cPre2  = idx['AULAS_PRE_CUTOVER'];
  const cPos   = idx['AULAS_POS_CUTOVER'];
  const cAtual = idx['AULA_ATUAL'];
  const cRest  = idx['AULAS_RESTANTES'];
  const cProg  = idx['DATA_PROG_INICIO'];

  const mapPlano = {
    'mensal' : 'mensalidade', 'mensalidade': 'mensalidade',
    '6m'     : '6 meses',     '6 mes'     : '6 meses', '6 meses' : '6 meses',
    '12a'    : '12 aulas',    '12 aulas'  : '12 aulas',
    '12m'    : '12 meses',    '12 mes'    : '12 meses','12 meses': '12 meses',
    'avulsa' : 'avulsa'
  };

  // pega as linhas selecionadas (1+)
  const start = sel.getRow();
  const count = sel.getNumRows();
  let created = 0;

  for (let k = 0; k < count; k++){
    const r = start + k;
    if (r <= 1) continue; // pula cabeçalho

    const row = sh.getRange(r,1,1,head.length).getValues()[0];
    if (String(row[cStatus]||'').toLowerCase() !== 'ativo') continue;

    const idOld   = row[cID];
    const aluno   = row[cAluno];
    const norm    = (String(row[cNorm]||'').trim() || normalizeName_(aluno));
    const planoOld= String(row[cPlano]||'mensalidade');
    const freqOld = Number(row[cFreq] || 1);            // <<< número (sem parênteses!)
    const mesesOld= Number(row[cMeses] || (/\bmes/i.test(planoOld) ? 1 : 0));
    const modOld  = (String(row[cMod]||'i').toLowerCase()==='d') ? 'd' : 'i';

    // --- lê os campos NOVOS (com defaults) ---
    let planoNew = row[cPlanoNovo] ? String(row[cPlanoNovo]).trim().toLowerCase() : 'mensalidade';
    planoNew = mapPlano[planoNew] || planoNew;

    let freqNew = (cFreqNova!=null) ? Number(row[cFreqNova]||0) : 0;
    if (!freqNew || isNaN(freqNew)) freqNew = freqOld;

    let mesesNew;
    if (planoNew === '12 aulas') mesesNew = 0;
    else if (planoNew === '6 meses') mesesNew = 6;
    else if (planoNew === '12 meses') mesesNew = 12;
    else if (cMesesNovo!=null && Number(row[cMesesNovo])) mesesNew = Number(row[cMesesNovo]);
    else mesesNew = 1; // mensalidade

    let modNew = (cModNova!=null && row[cModNova]) ? String(row[cModNova]).toLowerCase() : modOld;
    if (modNew !== 'd') modNew = 'i';

    const dataNovo = (cDataNovo!=null && row[cDataNovo])
      ? (row[cDataNovo] instanceof Date ? row[cDataNovo] : new Date(row[cDataNovo]))
      : new Date();

    // desconto novo: usa célula se válida; senão regra padrão
    const descNovoRaw = (cDescNovo!=null) ? row[cDescNovo] : '';
    let descNew = (typeof parseDiscountValueString_==='function')
      ? parseDiscountValueString_(descNovoRaw) : NaN;
    if (isNaN(descNew)) descNew = resolveDefaultDiscount_(planoNew, freqNew, cfg) || 0;

    // preço base por modalidade → unitário
    const base = (modNew==='d' ? cfg.precoDuo : cfg.precoIndividual);
    const unit = base * (1 - descNew);

    // quantidade de aulas do NOVO
    let qtNew = 0;
    if (planoNew === '12 aulas') qtNew = 12;
    else if (planoNew === 'avulsa') qtNew = 0; // avulsa não carrega saldo
    else qtNew = aulasTotaisDoPlano_(planoNew, freqNew, mesesNew, cfg.semanasMes || 4);

    // --- fecha o antigo ---
    row[cStatus] = 'encerrado';
    if (cRenPara!=null) row[cRenPara] = ''; // setaremos já já
    sh.getRange(r,1,1,head.length).setValues([row]);

    // --- cria o novo ---
    const newRow = new Array(head.length).fill('');
    const newId  = generateNextContractId_();

    if (cID    !=null) newRow[cID]    = newId;
    if (cAluno !=null) newRow[cAluno] = aluno;
    if (cNorm  !=null) newRow[cNorm]  = norm;
    if (cPlano !=null) newRow[cPlano] = planoNew;
    if (cFreq  !=null) newRow[cFreq]  = freqNew;
    if (cMeses !=null) newRow[cMeses] = mesesNew;
    if (cMod   !=null) newRow[cMod]   = modNew;
    if (cIni   !=null) newRow[cIni]   = dataNovo;
    if (cDesc  !=null) newRow[cDesc]  = descNew;
    if (cPreco !=null) newRow[cPreco] = base;
    if (cUnit  !=null) newRow[cUnit]  = unit;
    if (cQt    !=null) newRow[cQt]    = qtNew;
    if (cStatus!=null) newRow[cStatus]= 'ativo';

    if (cPre1  !=null) newRow[cPre1]  = 0;
    if (cPre2  !=null) newRow[cPre2]  = 0;
    if (cPos   !=null) newRow[cPos]   = 0;
    if (cAtual !=null) newRow[cAtual] = 0;
    if (cRest  !=null) newRow[cRest]  = qtNew;
    if (cProg  !=null) newRow[cProg]  = dataNovo;

    if (cRenDe !=null) newRow[cRenDe] = idOld;
    if (cRenPara!=null){
      // grava o id de "para" na linha antiga também
      const old = sh.getRange(r,1,1,head.length).getValues()[0];
      old[cRenPara] = newId;
      sh.getRange(r,1,1,head.length).setValues([old]);
    }

    // append
    const wr = sh.getLastRow()+1;
    sh.getRange(wr,1,1,head.length).setValues([newRow]);

    // formatinhos
    if (cDesc !=null) sh.getRange(wr, cDesc+1, 1, 1).setNumberFormat('0.00%');
    if (cUnit !=null) sh.getRange(wr, cUnit+1, 1, 1).setNumberFormat('R$ #,##0.00');
    if (cPreco!=null) sh.getRange(wr, cPreco+1,1, 1).setNumberFormat('R$ #,##0.00');

    created++;
  }

  ui.alert(`Renovação criada para ${created} contrato(s).`);
}


function fin_gerarRollforwardSplit(){
  const cfg = readConfigFin_();
  const year = new Date().getFullYear();

  const shC = ensureContratosSheet_();
  const shR = getSheet_(SH_RECON);
  const shF = getSheet_(SH_ROLLFWD);

  // limpa e escreve cabeçalho novo
  shF.clear();
  const header = [
    'AAAA-MM',
    'SALDO_INICIAL_MENSAL','SALDO_INICIAL_PLANOS',
    '(+) NOVOS_MENSAL','(+) NOVOS_PLANOS',
    '(-) RECEITA_MENSAL','(-) RECEITA_PLANOS',
    'SALDO_FINAL_MENSAL','SALDO_FINAL_PLANOS','SALDO_FINAL_TOTAL'
  ];
  shF.getRange(1,1,1,header.length).setValues([header]);

  // helpers
  const ym = (y,m)=> y + '-' + String(m+1).padStart(2,'0');
  const isMensal = (p)=> String(p||'').toLowerCase().includes('mensal');

  // ---- NOVOS CONTRATOS (por mês) a partir da aba Contratos
  const lrC = shC.getLastRow();
  const valsC = lrC>1 ? shC.getRange(1,1,lrC,shC.getLastColumn()).getValues() : [];
  const headC = valsC[0]||[];
  const idxC  = Object.fromEntries(headC.map((h,i)=>[h,i]));

  const addMensal = Array(12).fill(0);
  const addPlanos = Array(12).fill(0);

  for(let r=1;r<valsC.length;r++){
    const row = valsC[r];
    const dt  = row[idxC['DATA_INICIO']];
    if(!dt) continue;
    const d   = (dt instanceof Date) ? dt : new Date(dt);
    if (isNaN(d) || d.getFullYear() !== year) continue;

    const m   = d.getMonth();
    const plano = row[idxC['PLANO']];
    const qt   = Number(row[idxC['QTDE_AULAS_CONTR']]||0);
    let unit   = Number(row[idxC['PRECO_UNIT']]||0);

    // fallback: se PRECO_UNIT estiver vazio, calcula por modalidade + desconto
    if(!unit){
      const mod = (String(row[idxC['MODALIDADE_CONTR']]||'i').toLowerCase()==='d')?'d':'i';
      const base = (mod==='d'? cfg.precoDuo : cfg.precoIndividual);
      const dsc  = parseDiscountValueString_(row[idxC['DESCONTO_%']]) || 0;
      unit = base * (1 - dsc);
    }

    const valor = unit * qt;
    if (isMensal(plano)) addMensal[m] += valor;
    else                 addPlanos[m] += valor;
  }

  // ---- RECEITA RECONHECIDA (por mês) a partir da aba Reconhecimento
  const lrR = shR.getLastRow();
  const valsR = lrR>1 ? shR.getRange(1,1,lrR,shR.getLastColumn()).getValues() : [];
  const headR = valsR[0]||[];
  const idxR  = Object.fromEntries(headR.map((h,i)=>[h,i]));

  const recMensal = Array(12).fill(0);
  const recPlanos = Array(12).fill(0);

  for(let r=1;r<valsR.length;r++){
    const row = valsR[r];
    const key = String(row[idxR['AAAA-MM']]||'').trim();
    if (!/^\d{4}-\d{2}$/.test(key)) continue; // ignora "TOTAL"
    const y = Number(key.slice(0,4));
    const m = Number(key.slice(5,7)) - 1;
    if (y !== year) continue;

    const plano = row[idxR['PLANO']];
    const receita = Number(row[idxR['RECEITA_RECONHECIDA']]||0);

    if (isMensal(plano)) recMensal[m] += receita;
    else                 recPlanos[m] += receita;
  }

  // ---- monta o rollforward mês a mês
  let siMensal = 0, siPlanos = 0; // saldos iniciais começam em 0 (ano corrente)
  const out = [];

  for(let m=0;m<12;m++){
    const novosM = addMensal[m], novosP = addPlanos[m];
    const recM   = recMensal[m], recP   = recPlanos[m];

    const sfMensal = Math.max(0, siMensal + novosM - recM);
    const sfPlanos = Math.max(0, siPlanos + novosP - recP);

    out.push([
      ym(year,m),
      siMensal, siPlanos,
      novosM, novosP,
      recM, recP,
      sfMensal, sfPlanos, sfMensal + sfPlanos
    ]);

    // próximo mês começa com o saldo final do atual
    siMensal = sfMensal;
    siPlanos = sfPlanos;
  }

  if(out.length) shF.getRange(2,1,out.length,header.length).setValues(out);

  // formatação
  const fmtMoeda = 'R$ #,##0.00';
  shF.getRange(2,2,out.length,8).setNumberFormat(fmtMoeda);

  notify_('Rollforward (Mensalidade x Planos) gerado ✅');
}
function fin_gerarResumoSaldos(){
  const ss  = SpreadsheetApp.getActive();
  const cfg = readConfigFin_();
  const shC = ensureContratosSheet_();
  const lr  = shC.getLastRow();
  if (lr < 2) { notify_('Sem contratos.'); return; }

  const vals = shC.getRange(1,1,lr,shC.getLastColumn()).getValues();
  const head = vals[0];
  const idx  = Object.fromEntries(head.map((h,i)=>[h,i]));
  const gi = col => { const i = idx[col]; if (i===undefined) throw new Error(`Coluna não encontrada: ${col}`); return i; };

  const cID     = gi('ID_CONTRATO');
  const cAluno  = gi('ALUNO');
  const cStatus = gi('STATUS');
  const cPlano  = gi('PLANO');
  const cFreq   = gi('FREQUENCIA_SEMANAL');
  const cRest   = gi('AULAS_RESTANTES');
  const cQt     = gi('QTDE_AULAS_CONTR');
  const cUnit   = idx['PRECO_UNIT'];
  const cMod    = gi('MODALIDADE_CONTR');
  const cDesc   = gi('DESCONTO_%');

  const isMensal = (p)=> String(p||'').toLowerCase().includes('mensal');
  const isAvulsa = (p)=> String(p||'').toLowerCase().includes('avulsa');

  let aulasMensal = 0, valorMensal = 0;
  let aulasPlano  = 0, valorPlano  = 0;

  let renovarMensal = 0, renovarPlano = 0;
  let valorRenovarMensal = 0, valorRenovarPlano = 0;
  const listaMensal = [], listaPlano = [];

  for (let r=1; r<vals.length; r++){
    const row    = vals[r];
    const status = String(row[cStatus]||'').toLowerCase();
    if (status !== 'ativo') continue;

    const plano = row[cPlano];
    if (isAvulsa(plano)) continue;

    const aulas = Number(row[cRest]||0);
    const qt    = Number(row[cQt] || 0);
    const freq  = Number(row[cFreq] || 1);

    // PRECO_UNIT: usa o da coluna; se estiver vazio, calcula pela modalidade + desconto (freq p/ mensalidade)
    let unit = (cUnit!==undefined) ? Number(row[cUnit]||0) : 0;
    if (!unit || isNaN(unit)){
      const mod = (String(row[cMod]||'i').toLowerCase()==='d') ? 'd' : 'i';
      const base= (mod==='d' ? cfg.precoDuo : cfg.precoIndividual);
      let dsc   = parseDiscountValueString_(row[cDesc]);
      if (isNaN(dsc) || dsc===0) dsc = resolveDefaultDiscount_(plano, freq, cfg);
      unit = base * (1 - (dsc||0));
    }

    const valSaldo = aulas * unit;
    const valCiclo = qt * unit;

    if (isMensal(plano)) {
      aulasMensal += aulas; valorMensal += valSaldo;
      if (aulas === 0) { renovarMensal++; valorRenovarMensal += valCiclo; listaMensal.push(`${row[cAluno]} (${row[cID]})`); }
    } else {
      aulasPlano  += aulas; valorPlano  += valSaldo;
      if (aulas === 0) { renovarPlano++;  valorRenovarPlano  += valCiclo; listaPlano.push(`${row[cAluno]} (${row[cID]})`); }
    }
  }

  const shS = (function getOrCreate(name){
    const s = ss.getSheetByName(name);
    return s || ss.insertSheet(name);
  })('Resumo_Saldos');

  shS.clear();
  shS.getRange(1,1,1,3).setValues([['Atualizado em', new Date(), '']]);
  shS.getRange(2,1,1,3).setValues([['Tipo','AULAS / QTDE','VALOR (R$)']]);

  const linhas = [
    ['Mensalidades ativas',                    aulasMensal,                       valorMensal],
    ['Planos ativos',                          aulasPlano,                        valorPlano ],
    ['Mensalidades a renovar (qtd contratos)', renovarMensal,                     valorRenovarMensal],
    ['Planos a renovar (qtd contratos)',       renovarPlano,                      valorRenovarPlano ],
    ['TOTAL',                                  aulasMensal + aulasPlano,          valorMensal + valorPlano]
  ];
  shS.getRange(3,1,linhas.length,3).setValues(linhas);
  shS.getRange(3,3,linhas.length,1).setNumberFormat('R$ #,##0.00');
  if (renovarMensal > 0) shS.getRange(5,1).setNote('Mensalidades a renovar:\n' + listaMensal.join('\n'));
  if (renovarPlano  > 0) shS.getRange(6,1).setNote('Planos a renovar:\n'       + listaPlano.join('\n'));
  notify_('Resumo_Saldos gerado ✅ (desconto de mensalidade por frequência aplicado).');
}

// Cria, se faltar, as colunas usadas na renovação (inclui DESCONTO_NOVO)
function fin_addMissingRenovationColumns(){
  const sh = ensureContratosSheet_();
  const lr = sh.getLastRow();
  const lc = sh.getLastColumn();
  if (lr < 1) return;

  const head = sh.getRange(1,1,1,lc).getValues()[0];
  const have = name => head.indexOf(name) !== -1;

  // Colunas que queremos garantir (as já existentes serão ignoradas)
  const needed = [
    'DATA_INICIO_NOVO','PLANO_NOVO','FREQ_NOVA','MESES_DUR_NOVO',
    'MOD_NOVA',          // i ou d
    'DESCONTO_NOVO',     // % do novo contrato
    'ID_RENOVACAO_DE','ID_RENOVACAO_PARA'
  ];

  let added = 0;
  needed.forEach(name=>{
    if (have(name)) return;
    const col = sh.getLastColumn() + 1;              // adiciona no final
    sh.insertColumnAfter(sh.getLastColumn());
    sh.getRange(1, col).setValue(name);

    // formatos/validações úteis
    if (name === 'DESCONTO_NOVO') {
      sh.getRange(2, col, Math.max(1, lr-1), 1).setNumberFormat('0.00%'); // 5% -> 5,00%
      sh.getRange(1, col).setNote('Ex.: 5% | 0,05 | 0.05');
    }
    if (name === 'MOD_NOVA') {
      const rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['i','d'], true).setAllowInvalid(false).build();
      sh.getRange(2, col, Math.max(1, lr-1), 1).setDataValidation(rule);
      sh.getRange(1, col).setNote('i = individual | d = dupla');
    }
    added++;
  });

  notify_(added ? `Colunas criadas: ${added}` : 'Todas as colunas já existiam ✅');
}
function fin_corrigirDescontosContratosAntigos(){
  const cfg = readConfigFin_();
  const sh  = ensureContratosSheet_();
  const lr  = sh.getLastRow();
  if (lr < 2) { notify_('Sem contratos.'); return; }

  const rng  = sh.getRange(1,1,lr,sh.getLastColumn());
  const vals = rng.getValues();
  const head = vals[0];
  const idx  = Object.fromEntries(head.map((h,i)=>[h,i]));
  const gi   = name => { const i = idx[name]; if (i===undefined) throw new Error('Coluna faltando: '+name); return i; };

  const cPlano = gi('PLANO');
  const cFreq  = gi('FREQUENCIA_SEMANAL');
  const cMod   = gi('MODALIDADE_CONTR');      // 'i' | 'd'
  const cDesc  = gi('DESCONTO_%');
  const cUnit  = idx['PRECO_UNIT'];           // pode não existir em algumas versões

  let changed = 0;

  for (let r=1; r<vals.length; r++){
    const row   = vals[r];
    const plano = row[cPlano];
    const freq  = Number(row[cFreq]||1);
    const mod   = (String(row[cMod]||'i').toLowerCase()==='d') ? 'd' : 'i';

    // desconto ideal (mensalidade usa a regra 1x=5% | ≥2x=10%)
    const want = resolveDefaultDiscount_(plano, freq, cfg);

    // desconto atual na célula
    let curr = parseDiscountValueString_(row[cDesc]);
    if (isNaN(curr)) curr = 0;

    // preço unitário "certo" para esse contrato
    const base = (mod==='d' ? cfg.precoDuo : cfg.precoIndividual);
    const unitShould = base * (1 - want);

    // decide se precisa corrigir
    const needsDesc = Math.abs(curr - want) > 0.0005;
    const needsUnit = (cUnit !== undefined) && (isNaN(Number(row[cUnit])) || Math.abs(Number(row[cUnit]) - unitShould) > 0.001);

    if (needsDesc || needsUnit) {
      vals[r][cDesc] = want;
      if (cUnit !== undefined) vals[r][cUnit] = unitShould;
      changed++;
    }
  }

  rng.setValues(vals);

  // formatação
  const colDesc = gi('DESCONTO_%') + 1;
  sh.getRange(2, colDesc, lr-1, 1).setNumberFormat('0.00%');
  if (cUnit !== undefined) {
    sh.getRange(2, cUnit+1, lr-1, 1).setNumberFormat('R$ #,##0.00');
  }

  notify_(`Descontos/valores corrigidos em ${changed} linha(s).`);
}
function fin_baixarNovosDaAgendaPrompt(){
  const ui  = SpreadsheetApp.getUi();
  const cfg = readConfigFin_();

  // ========= 1) Escolha do plano e (opcional) frequência =========
  const p1 = ui.prompt(
    'Novo contrato (baixar da Agenda)',
    'Escolha o PLANO:  mensal  |  6m  |  12a  |  12m  |  avulsa',
    ui.ButtonSet.OK_CANCEL
  );
  if (p1.getSelectedButton() !== ui.Button.OK) return;

  const mapPlano = {
    'mensal': 'mensalidade', 'mensalidade':'mensalidade',
    '6m':'6 meses', '6 mes':'6 meses', '6 meses':'6 meses',
    '12a':'12 aulas', '12 aulas':'12 aulas',
    '12m':'12 meses', '12 mes':'12 meses', '12 meses':'12 meses',
    'avulsa':'avulsa'
  };
  const planoIn = String(p1.getResponseText()||'').trim().toLowerCase();
  const planoEscolhido = mapPlano[planoIn];
  if (!planoEscolhido) { ui.alert('Valor inválido. Use: mensal | 6m | 12a | 12m | avulsa'); return; }

  let freqOverride = null;
  if (planoEscolhido !== '12 aulas' && planoEscolhido !== 'avulsa') {
    const p2 = ui.prompt(
      'Frequência por semana',
      'Digite 1, 2, 3… (deixe vazio para estimar pela agenda).',
      ui.ButtonSet.OK_CANCEL
    );
    if (p2.getSelectedButton() === ui.Button.CANCEL) return;
    const t = String(p2.getResponseText()||'').trim();
    if (t) {
      const f = Number(t);
      if (!isNaN(f) && f>0 && f<=7) freqOverride = f;
      else ui.alert('Frequência inválida. Vou estimar pela agenda.');
    }
  }

  // ========= 2) Preparos de planilha/índices =========
  const shC = ensureContratosSheet_();
  const lr  = shC.getLastRow();
  const lc  = shC.getLastColumn();
  const vals= lr ? shC.getRange(1,1,lr,lc).getValues() : [[]];
  const head= vals[0] || [];
  const idx = Object.fromEntries(head.map((h,i)=>[h,i]));
  const gi  = h => (idx[h]!==undefined ? idx[h] : null);

  const cID=gi('ID_CONTRATO'), cAluno=gi('ALUNO'), cNorm=gi('NOME_NORMALIZADO'),
        cPlano=gi('PLANO'), cFreq=gi('FREQUENCIA_SEMANAL'), cMeses=gi('MESES_DURACAO'),
        cMod=gi('MODALIDADE_CONTR'), cIni=gi('DATA_INICIO'),
        cDesc=gi('DESCONTO_%'), cPreco=gi('PRECO_CHEIO_AULA'), cUnit=gi('PRECO_UNIT'),
        cQt=gi('QTDE_AULAS_CONTR'), cStatus=gi('STATUS'),
        cPre1=gi('AULAS_FEITAS_ATE_CUTOVER'), cPre2=gi('AULAS_PRE_CUTOVER'),
        cAtual=gi('AULA_ATUAL'), cRest=gi('AULAS_RESTANTES'),
        cPos=gi('AULAS_POS_CUTOVER'), cProg=gi('DATA_PROG_INICIO');

  // quem já tem contrato ATIVO
  const ativos = new Set();
  for (let r=1; r<vals.length; r++){
    if (String(vals[r][cStatus]||'').toLowerCase()==='ativo'){
      const nm = String(vals[r][cNorm]||'').trim() || normalizeName_(vals[r][cAluno]);
      if (nm) ativos.add(nm);
    }
  }

  // ========= 3) Lê agenda (últimos 60 dias) =========
  const cal   = CalendarApp.getCalendarById(CALENDAR_ID);
  const now   = new Date();
  const start = new Date(); start.setDate(start.getDate()-60); start.setHours(0,0,0,0);
  const events = cal.getEvents(start, now);

  /** norm -> {display, first:Date, cntI:number, cntD:number, dates:Date[]} */
  const seen = {};
  events.forEach(ev=>{
    const color = ev.getColor() || "";
    if (EXCLUDED_COLORS.includes(color)) return;

    const modality = INDIVIDUAL_COLORS.includes(color) ? 'i'
                    : (color === DUO_COLOR ? 'd' : null);
    if (!modality) return;

    const title = ev.getTitle()||'', desc = ev.getDescription()||'';
    if (!CHECK_PAT.test(title) && !CHECK_PAT.test(desc)) return; // só presença dada

    // suporta DUPLA com vários nomes no título (se tiver o helper multi)
    const names = (typeof extractNamesFromTitleMulti_ === 'function')
      ? extractNamesFromTitleMulti_(title)
      : [{display: extractSingleNameFromTitle_(title),
          norm: normalizeName_(extractSingleNameFromTitle_(title))}];

    names.forEach(n=>{
      if (!n.norm) return;
      if (!seen[n.norm]) seen[n.norm] = { display: n.display, first: ev.getStartTime(), cntI:0, cntD:0, dates:[] };
      const o = seen[n.norm];
      if (ev.getStartTime() < o.first) o.first = ev.getStartTime();
      if (modality==='i') o.cntI++; else o.cntD++;
      o.dates.push(ev.getStartTime());
    });
  });

  // ========= 4) Monta novos contratos =========
  const cutover = cfg.cutoverDateTime || new Date(now.getFullYear(), now.getMonth(), 1, 20, 0, 0);
  const newRows = [];
  const writeStart = lr ? lr+1 : 2;

  Object.keys(seen).forEach(norm=>{
    if (ativos.has(norm)) return; // já tem contrato ativo

    const o = seen[norm];
    const mod = (o.cntD > o.cntI) ? 'd' : 'i';

    // frequência: override ou estimada pelos últimos 28 dias
    let freq = freqOverride;
    if (!freq){
      const d28 = new Date(now); d28.setDate(d28.getDate()-28);
      const presUlt28 = o.dates.filter(d=>d>=d28).length;
      freq = Math.max(1, Math.min(5, Math.round(presUlt28/4)));
    }

    // meses conforme plano
    let meses = 1;
    if (planoEscolhido === '6 meses') meses = 6;
    if (planoEscolhido === '12 meses') meses = 12;
    if (planoEscolhido === 'avulsa')   meses = 0;

    // desconto padrão por plano/frequência (avulsa = 0)
    const dsc = (planoEscolhido === 'avulsa')
      ? 0
      : (typeof resolveDefaultDiscount_==='function'
          ? resolveDefaultDiscount_(planoEscolhido, freq, cfg)
          : (planoEscolhido==='12 aulas'?0.15:planoEscolhido==='6 meses'?0.25:planoEscolhido==='12 meses'?0.35:(freq>=2?0.10:0.05)));

    // preço unitário
    const base = (mod==='d' ? cfg.precoDuo : cfg.precoIndividual);
    const unit = base * (1 - (dsc||0));

    // quantidade contratada
    let qtde = 0;
    if (planoEscolhido === '12 aulas') qtde = 12;
    else if (planoEscolhido === 'avulsa') qtde = 0; // avulsa não carrega saldo
    else qtde = aulasTotaisDoPlano_(planoEscolhido, freq, meses, cfg.semanasMes||4);

    // aulas antes do cutover
    const preCut = o.dates.filter(d=>d < cutover).length;

    // monta linha
    const row = new Array(lc).fill('');
    if (cID     != null) row[cID]     = generateNextContractId_();
    if (cAluno  != null) row[cAluno]  = o.display;
    if (cNorm   != null) row[cNorm]   = norm;
    if (cPlano  != null) row[cPlano]  = planoEscolhido;
    if (cFreq   != null) row[cFreq]   = freq;
    if (cMeses  != null) row[cMeses]  = meses;
    if (cMod    != null) row[cMod]    = mod;
    if (cIni    != null) row[cIni]    = o.first;
    if (cDesc   != null) row[cDesc]   = dsc;
    if (cPreco  != null) row[cPreco]  = base;
    if (cUnit   != null) row[cUnit]   = unit;
    if (cQt     != null) row[cQt]     = qtde;
    if (cStatus != null) row[cStatus] = 'ativo';
    if (cPre1   != null) row[cPre1]   = preCut;
    if (cPre2   != null) row[cPre2]   = preCut;
    if (cAtual  != null) row[cAtual]  = Math.min(qtde, preCut);
    if (cRest   != null) row[cRest]   = Math.max(0, qtde - (row[cAtual]||0));
    if (cPos    != null) row[cPos]    = 0;
    if (cProg   != null) row[cProg]   = o.first;

    newRows.push(row);
  });

  if (!newRows.length) { ui.alert('Nenhum novo aluno com presença encontrada na Agenda.'); return; }

  shC.getRange(writeStart, 1, newRows.length, lc).setValues(newRows);

  if (cDesc != null) shC.getRange(writeStart, cDesc+1, newRows.length, 1).setNumberFormat('0.00%');
  if (cUnit != null) shC.getRange(writeStart, cUnit+1, newRows.length, 1).setNumberFormat('R$ #,##0.00');
  if (cPreco!= null) shC.getRange(writeStart, cPreco+1,newRows.length, 1).setNumberFormat('R$ #,##0.00');

  ui.alert(`Criados ${newRows.length} contrato(s) como "${planoEscolhido}".`);
}

function fin_testCores(){
  SpreadsheetApp.getUi().alert(
    'INDIVIDUAL_COLORS = ' + JSON.stringify(INDIVIDUAL_COLORS) + '\n' +
    'DUO_COLOR         = ' + DUO_COLOR + '\n' +
    'EXCLUDED_COLORS   = ' + JSON.stringify(EXCLUDED_COLORS)
  );
}
function fin_dividirContratoNoMeio(){
  const ui  = SpreadsheetApp.getUi();
  const ss  = SpreadsheetApp.getActive();
  const sh  = ensureContratosSheet_();
  const sel = sh.getActiveRange();
  if (!sel || sel.getNumRows() !== 1) { ui.alert('Selecione UMA linha do contrato para alterar.'); return; }

  const head = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const idx  = Object.fromEntries(head.map((h,i)=>[h,i]));
  const gi = n => { const i = idx[n]; if (i===undefined) throw new Error('Coluna faltando: '+n); return i; };

  const r0 = sel.getRow();
  if (r0 <= 1) { ui.alert('Selecione uma linha de dados (abaixo do cabeçalho).'); return; }

  const row = sh.getRange(r0,1,1,sh.getLastColumn()).getValues()[0];

  const cAluno = gi('ALUNO'), cNorm = gi('NOME_NORMALIZADO'), cStatus=gi('STATUS'),
        cPlano = gi('PLANO'), cFreq = gi('FREQUENCIA_SEMANAL'), cMeses=gi('MESES_DURACAO'),
        cMod   = gi('MODALIDADE_CONTR'), cIni = gi('DATA_INICIO'),
        cQt    = gi('QTDE_AULAS_CONTR'), cUnit = idx['PRECO_UNIT'],
        cDesc  = gi('DESCONTO_%'),
        cAtual = gi('AULA_ATUAL'), cRest = gi('AULAS_RESTANTES'),
        cPre1  = idx['AULAS_FEITAS_ATE_CUTOVER'], cPre2 = idx['AULAS_PRE_CUTOVER'],
        cPos   = idx['AULAS_POS_CUTOVER'],
        cID    = gi('ID_CONTRATO'), cRenDe = idx['ID_RENOVACAO_DE'], cRenPara = idx['ID_RENOVACAO_PARA'];

  if (String(row[cStatus]||'').toLowerCase()!=='ativo'){
    ui.alert('Contrato não está ATIVO.'); return;
  }

  // ===== prompts
  const pData = ui.prompt('Data efetiva da mudança','Informe a data (dd/mm/aaaa).', ui.ButtonSet.OK_CANCEL);
  if (pData.getSelectedButton()!==ui.Button.OK) return;
  const sData = String(pData.getResponseText()||'').trim();
  const dt    = parseBrDate_(sData);
  if (!dt) { ui.alert('Data inválida. Use dd/mm/aaaa.'); return; }

  const pMod  = ui.prompt('Nova modalidade','Digite i (individual) ou d (dupla).', ui.ButtonSet.OK_CANCEL);
  if (pMod.getSelectedButton()!==ui.Button.OK) return;
  let modNew  = String(pMod.getResponseText()||'i').trim().toLowerCase();
  if (modNew!=='d') modNew='i';

  const pFreq = ui.prompt('Nova frequência (opcional)','Digite 1/2/3… (deixe vazio para manter a atual).', ui.ButtonSet.OK_CANCEL);
  if (pFreq.getSelectedButton()===ui.Button.CANCEL) return;
  let freqNew = Number(String(pFreq.getResponseText()||'').trim());
  if (!freqNew || isNaN(freqNew)) freqNew = Number(row[cFreq]||1);

  const pPlano= ui.prompt('Novo plano (opcional)','Digite: mensal | 6m | 12a | 12m  (vazio = manter).', ui.ButtonSet.OK_CANCEL);
  if (pPlano.getSelectedButton()===ui.Button.CANCEL) return;
  const map = {'mensal':'mensalidade','6m':'6 meses','12a':'12 aulas','12m':'12 meses'};
  let planoNew = String(pPlano.getResponseText()||'').trim().toLowerCase();
  planoNew = planoNew ? (map[planoNew]||planoNew) : String(row[cPlano]||'mensalidade');

  const pDesc = ui.prompt('Desconto novo (opcional)','Ex.: 5% ou 0,05 (vazio = regra padrão).', ui.ButtonSet.OK_CANCEL);
  if (pDesc.getSelectedButton()===ui.Button.CANCEL) return;
  const descIn = parseDiscountValueString_(pDesc.getResponseText());

  // ===== dados atuais
  const aluno = row[cAluno];
  const norm  = String(row[cNorm]||'').trim() || normalizeName_(aluno);
  const dtIni = (row[cIni] instanceof Date) ? row[cIni] : new Date(row[cIni]);
  const qt    = Number(row[cQt]||0);

  // aulas dadas até a data (conta pela Agenda com ✅)
  const dadas = countPresencasAluno_(norm, dtIni, dt);
  const restante = Math.max(0, qt - dadas);
  if (restante === 0){
    ui.alert('Não há aulas restantes para dividir.'); return;
  }

  // ===== calcula preço/desconto do NOVO
  const cfg = readConfigFin_();
  let descNew = !isNaN(descIn) ? descIn : resolveDefaultDiscount_(planoNew, freqNew, cfg);
  const base  = (modNew==='d' ? cfg.precoDuo : cfg.precoIndividual);
  const unit  = base * (1 - (descNew||0));

  // meses aprox. para o novo (apenas informativo)
  let mesesNew = Number(row[cMeses]||1);
  if (/meses$/i.test(planoNew)){
    const porMes = Math.max(1, freqNew) * (cfg.semanasMes||4);
    mesesNew = Math.ceil(restante / porMes);
  } else if (/12 aulas/i.test(planoNew)) {
    mesesNew = 0; // não se aplica
  }

  // ===== fecha o antigo
  const idOld = row[cID];
  row[cStatus] = 'encerrado';
  if (cAtual!=null) row[cAtual] = dadas;
  if (cRest !=null) row[cRest]  = 0;

  // zera campos de progresso no novo
  const newId = generateNextContractId_();
  const newRow = new Array(head.length).fill('');
  if (cID   !=null) newRow[cID]   = newId;
  if (cAluno!=null) newRow[cAluno]= aluno;
  if (cNorm !=null) newRow[cNorm] = norm;
  if (cPlano!=null) newRow[cPlano]= planoNew;
  if (cFreq !=null) newRow[cFreq] = freqNew;
  if (cMeses!=null) newRow[cMeses]= mesesNew;
  if (cMod  !=null) newRow[cMod]  = modNew;
  if (cIni  !=null) newRow[cIni]  = dt;
  if (cDesc !=null) newRow[cDesc] = descNew;
  if (cUnit !=null) newRow[cUnit] = unit;
  if (cQt   !=null) newRow[cQt]   = restante;
  if (cStatus!=null)newRow[cStatus] = 'ativo';
  if (cPre1 !=null) newRow[cPre1] = 0;
  if (cPre2 !=null) newRow[cPre2] = 0;
  if (cPos  !=null) newRow[cPos]  = 0;
  if (cAtual!=null) newRow[cAtual]= 0;
  if (cRest !=null) newRow[cRest] = restante;
  if (cRenDe!=null) newRow[cRenDe]= idOld;
  if (cRenPara!=null) row[cRenPara] = newId;

  // grava
  sh.getRange(r0,1,1,sh.getLastColumn()).setValues([row]);
  sh.getRange(sh.getLastRow()+1,1,1,sh.getLastColumn()).setValues([newRow]);

  // formats
  if (cDesc!=null) sh.getRange(sh.getLastRow(), cDesc+1, 1, 1).setNumberFormat('0.00%');
  if (cUnit!=null) sh.getRange(sh.getLastRow(), cUnit+1, 1, 1).setNumberFormat('R$ #,##0.00');

  ui.alert('Contrato dividido ✅\nAntigo encerrado; novo criado com as aulas restantes.');
}

// === helpers ===

// conta presenças por aluno (nome normalizado) entre datas
function countPresencasAluno_(normName, startDate, endDate){
  const cal = CalendarApp.getCalendarById(CALENDAR_ID);
  const events = cal.getEvents(startDate, endDate);
  let count = 0;
  events.forEach(ev=>{
    const color = ev.getColor() || "";
    if (EXCLUDED_COLORS.includes(color)) return;

    const t = ev.getTitle()||'', d = ev.getDescription()||'';
    if (!CHECK_PAT.test(t) && !CHECK_PAT.test(d)) return;

    const names = (typeof extractNamesFromTitleMulti_==='function')
      ? extractNamesFromTitleMulti_(t)
      : [{display: extractSingleNameFromTitle_(t), norm: normalizeName_(extractSingleNameFromTitle_(t))}];

    if (names.some(n => n.norm === normName)) count++;
  });
  return count;
}

// dd/mm/aaaa -> Date (meia-noite)
function parseBrDate_(s){
  const m = String(s||'').trim().match(/^(\d{2})[\/\-](\d{2})[\/\-](\d{4})$/);
  if (!m) return null;
  const d = new Date(Number(m[3]), Number(m[2])-1, Number(m[1]), 0,0,0,0);
  return isNaN(d) ? null : d;
}
