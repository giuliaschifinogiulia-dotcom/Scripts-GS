/***** =========================
 * CONFIGURAÇÕES GERAIS
 * =========================*****/
const CALENDAR_ID       = 'primary';     // Agenda do estúdio
const RAW_SHEET         = 'Página1';     // Data | Aluno | Status
const REPORT_SHEET      = 'Relatório';   // Resumo mensal por aluno
const ALERTS_SHEET      = 'Alertas';     // Marcos 2/4/6 meses
const MONTH_PANEL_SHEET = 'Painel_Mensal'; // Painel (B1 = AAAA-MM)

// DATA DE INÍCIO (ignora eventos anteriores ao importar RAW)
const START_YEAR  = 2025;
const START_MONTH = 7;  // 0=Jan, 7=Agosto
const START_DAY   = 18;

// Cores aceitas no PAINEL (Calendar → Painel). Para incluir experimentais (banana=5), adicione "5".
const ACCEPTED_COLOR_IDS = ["", "1"]; // "", "1" e opcionalmente "5"

/***** =========================
 * MENU / ENTRADA PRINCIPAL
 * =========================*****/
function atualizar() {
  importarPresencas();   // Calendar → RAW (até AGORA), cores "" e "1"
  montarRelatorio();     // Consolida por aluno x mês (a partir do RAW)
  checarMarcos();        // Gera alertas 2/4/6 meses
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Frequência')
    .addItem('Atualizar agora', 'atualizar')
    .addItem('Painel mensal (B1 = AAAA-MM)', 'montarPainelMensal')
    .addItem('Consulta por período', 'consultarPeriodoAluno')
    .addItem('Criar gatilho diário (20h)', 'criarGatilhoDiario')
    .addItem('Remover gatilhos', 'removerGatilhos')
    .addToUi();
}

/***** =========================
 * PASSO 1 — IMPORTAÇÃO (só cor "" e "1")
 * =========================*****/
function importarPresencas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(RAW_SHEET) || ss.insertSheet(RAW_SHEET);

  // Limpa e recria cabeçalho
  sh.clearContents();
  sh.appendRow(['Data', 'Aluno', 'Status']);

  const inicio = new Date(START_YEAR, START_MONTH, START_DAY, 0, 0, 0);
  const agora  = new Date();

  const cal = CalendarApp.getCalendarById(CALENDAR_ID);
  const eventos = cal.getEvents(inicio, agora);

  const linhas = [];
  eventos.forEach(ev => {
    const dataInicio = ev.getStartTime();
    if (dataInicio < inicio || dataInicio > agora) return;

    // Aceita somente laranja (padrão "") e lavanda ("1") no RAW
    const colorRaw = ev.getColor(); // null quando padrão
    const colorId  = (colorRaw == null) ? "" : String(colorRaw);
    if (colorId !== "" && colorId !== "1") return;

    const titulo = ev.getTitle() || '';
    const aluno  = normalizeAlunoFromTitle_(titulo);
    if (!aluno) return;

    const status = isPresenteFromTitle_(titulo) ? 'Presente' : 'Falta';
    linhas.push([dataInicio, aluno, status]);
  });

  linhas.sort((a,b) => a[0] - b[0]);
  if (linhas.length) {
    sh.getRange(2,1,linhas.length,3).setValues(linhas);
    sh.getRange(2,1,linhas.length,1).setNumberFormat('dd/MM/yyyy HH:mm');
  }
}

/***** =========================
 * PASSO 2 — RELATÓRIO (por mês, a partir do RAW)
 * =========================*****/
function montarRelatorio() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const raw = ss.getSheetByName(RAW_SHEET);
  let rep   = ss.getSheetByName(REPORT_SHEET);
  if (!rep) rep = ss.insertSheet(REPORT_SHEET);

  rep.clearContents();
  const filtroAtual = rep.getFilter ? rep.getFilter() : null;
  if (filtroAtual) filtroAtual.remove();
  rep.appendRow(['Aluno', 'Mês', 'Presenças', 'Faltas', 'Aulas Esperadas', '% Real', '% Contratual']);

  if (!raw || raw.getLastRow() < 2) return;
  const dados = raw.getDataRange().getValues(); // inclui cabeçalho

  const mapa = {}; // { aluno: { 'YYYY-MM': {p, f} } }
  for (let i = 1; i < dados.length; i++) {
    const [data, aluno, status] = dados[i];
    if (!aluno || !status || !data) continue;
    const d  = new Date(data);
    const ym = `${d.getFullYear()}-${('0' + (d.getMonth() + 1)).slice(-2)}`;
    if (!mapa[aluno]) mapa[aluno] = {};
    if (!mapa[aluno][ym]) mapa[aluno][ym] = { p: 0, f: 0 };
    if (status === 'Presente') mapa[aluno][ym].p++;
    if (status === 'Falta')    mapa[aluno][ym].f++;
  }

  const planos = getPlanos(); // { 'Nome': aulas/semana }
  const agora  = new Date();

  const linhas = [];
  Object.keys(mapa).sort().forEach(aluno => {
    Object.keys(mapa[aluno]).sort().forEach(ym => {
      const { p, f } = mapa[aluno][ym];
      const [y, m]   = ym.split('-').map(Number);
      const mesStr   = Utilities.formatDate(new Date(y, m - 1, 1), Session.getScriptTimeZone(), 'MMMM/yyyy');
      const total    = p + f;
      const percReal = total ? p / total : 0;  // 0 se não houver base
      const aulasSemana       = planos[aluno] || 0;
      const semanasEsperadas  = weeksForMonthUpToNow(y, m, agora);
      const esperado          = aulasSemana * semanasEsperadas;
      const percContratual    = esperado > 0 ? p / esperado : 0;
      linhas.push([aluno, mesStr, p, f, esperado, percReal, percContratual]);
    });
  });

  if (linhas.length) {
    rep.getRange(2,1,linhas.length,7).setValues(linhas);
    rep.getRange(2,6,linhas.length,2).setNumberFormat('0.0%');
  }
  rep.getRange(1,1,Math.max(2,linhas.length+1),7).createFilter();
  rep.autoResizeColumns(1,7);
}

/***** =========================
 * PASSO 3 — ALERTAS 2/4/6 Meses
 * =========================*****/
function checarMarcos() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const raw = ss.getSheetByName(RAW_SHEET);
  if (!raw || raw.getLastRow() < 2) return;

  const alertSheet   = ensureAlertsSheet();
  const existingKeys = loadExistingAlertKeys(alertSheet);

  const dados       = raw.getRange(2, 1, raw.getLastRow() - 1, 3).getValues(); // Data | Aluno | Status
  const inicioGeral = new Date(START_YEAR, START_MONTH, START_DAY);
  const cal         = CalendarApp.getCalendarById(CALENDAR_ID);
  const hoje        = new Date();
  const marcosMeses = [2, 4, 6];

  // 1ª aula por aluno (>= START_DATE)
  const firstDateByAluno = {};
  dados.forEach(([data, aluno]) => {
    if (!aluno || !data) return;
    const d = new Date(data);
    if (d < inicioGeral) return;
    if (!firstDateByAluno[aluno] || d < firstDateByAluno[aluno]) {
      firstDateByAluno[aluno] = d;
    }
  });

  const rowsToAppend = [];

  Object.keys(firstDateByAluno).forEach(aluno => {
    const start = firstDateByAluno[aluno];

    marcosMeses.forEach(m => {
      const end = addMonths(start, m);

      const eventosPeriodo = cal.getEvents(start, end).filter(ev => {
        const colorRaw = ev.getColor();
        const colorId  = (colorRaw == null) ? "" : String(colorRaw);
        if (colorId !== "" && colorId !== "1") return false;
        const titulo = normalizeAlunoFromTitle_(ev.getTitle() || '');
        return titulo === aluno;
      });
      if (eventosPeriodo.length === 0) return;

      let presencasAteAgora = 0;
      let faltasAteAgora    = 0;
      let futurosNoPeriodo  = 0;

      eventosPeriodo.forEach(ev => {
        const d = ev.getStartTime();
        if (d > hoje) {
          futurosNoPeriodo++;
        } else {
          const titulo = ev.getTitle() || '';
          if (isPresenteFromTitle_(titulo)) presencasAteAgora++;
          else                               faltasAteAgora++;
        }
      });

      if (faltasAteAgora === 0) {
        if (futurosNoPeriodo === 2) {
          const key = makeAlertKey(aluno, m, 'Faltam 2 aulas', start, end);
          if (!existingKeys.has(key)) {
            rowsToAppend.push([
              new Date(), aluno, m, 'Faltam 2 aulas',
              `${fmt(start)} a ${fmt(end)}`, fmt(start), fmt(end),
              presencasAteAgora, faltasAteAgora, futurosNoPeriodo, key
            ]);
            existingKeys.add(key);
          }
        }
        if (hoje >= end && futurosNoPeriodo === 0) {
          const key = makeAlertKey(aluno, m, 'Completou 100%', start, end);
          if (!existingKeys.has(key)) {
            rowsToAppend.push([
              new Date(), aluno, m, 'Completou 100%',
              `${fmt(start)} a ${fmt(end)}`, fmt(start), fmt(end),
              presencasAteAgora, faltasAteAgora, futurosNoPeriodo, key
            ]);
            existingKeys.add(key);
          }
        }
      }
    });
  });

  if (rowsToAppend.length) {
    alertSheet
      .getRange(alertSheet.getLastRow() + 1, 1, rowsToAppend.length, rowsToAppend[0].length)
      .setValues(rowsToAppend);
    const last = alertSheet.getLastRow();
    alertSheet.getRange(last - rowsToAppend.length + 1, 1, rowsToAppend.length, 1)
      .setNumberFormat('dd/MM/yyyy HH:mm');
  }
}

/***** =========================
 * PLANOS (Aluno | Aulas/Semana)
 * =========================*****/
function getPlanos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Planos');
  if (!sheet || sheet.getLastRow() < 2) return {};
  const dados = sheet.getRange(2,1,sheet.getLastRow()-1,2).getValues();
  const planos = {};
  dados.forEach(([aluno, aulasSemana]) => {
    if (!aluno) return;
    let nome = toTitleCase(String(aluno).trim().replace(/\s+/g,' '));
    const n = Number(aulasSemana);
    if (nome && !isNaN(n)) planos[nome] = n;
  });
  return planos;
}

/***** =========================
 * PAINEL MENSAL (B1 = AAAA-MM) — lê direto do Calendar
 * =========================*****/
function montarPainelMensal() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(MONTH_PANEL_SHEET);
  if (!sh) sh = ss.insertSheet(MONTH_PANEL_SHEET);

  if (sh.getLastRow() < 1) sh.appendRow(['']);
  sh.getRange('A1').setValue('Mês (AAAA-MM)');
  sh.getRange('B1').setNote('Ex.: 2025-08. Mês passado = mês completo; mês atual = até agora.');

  // Limpa área e filtro
  const filtro = sh.getFilter ? sh.getFilter() : null;
  if (filtro) filtro.remove();
  sh.getRange(2,1,Math.max(0, sh.getMaxRows()-1), 10).clearContent().clearFormat();

  const mesStr = getYearMonthFromB1_(sh);
  if (!mesStr) {
    sh.getRange('A2').setValue('Informe um mês válido em B1, ex.: 2025-08');
    return;
  }

  const [year, month] = mesStr.split('-').map(Number); // 1..12
  const now   = new Date();
  const range = monthRangeConsideringNow_(year, month, now);

  // Cabeçalho
  sh.getRange(2,1,1,5).setValues([['Aluno','Aulas Agendadas','Presenças','Faltas','% Presença']]);

  // === LÊ DIRETO DO CALENDAR (ignora START_DATE) ===
  const cal = CalendarApp.getCalendarById(CALENDAR_ID);
  const events = cal.getEvents(range.start, range.end);

  const mapa = {}; // key -> { display, total, pres, freqDisplay: {nome:cont} }
  events.forEach(ev => {
    // Filtra por cor aceita
    if (!isAcceptedColorId_(ev.getColor())) return;

    const titulo = ev.getTitle() || '';
    const aluno  = normalizeAlunoFromTitle_(titulo);
    if (!aluno) return;

    const display = toTitleCase(aluno);
    const key = canonicalKey_(display);

    if (!mapa[key]) mapa[key] = { display, total: 0, pres: 0, freqDisplay: {} };
    mapa[key].total++;
    if (isPresenteFromTitle_(titulo)) mapa[key].pres++;

    // Escolhe a grafia mais frequente
    mapa[key].freqDisplay[display] = (mapa[key].freqDisplay[display] || 0) + 1;
    if (mapa[key].freqDisplay[display] > (mapa[key].freqDisplay[mapa[key].display] || 0)) {
      mapa[key].display = display;
    }
  });

  const linhas = Object.values(mapa)
    .sort((a,b) => a.display.localeCompare(b.display, 'pt-BR'))
    .map(({display, total, pres}) => {
      const faltas = Math.max(0, total - pres);
      const perc   = total > 0 ? pres / total : 0;
      return [display, total, pres, faltas, perc];
    });

  if (!linhas.length) {
    sh.getRange(3,1,1,5).setValues([['(sem eventos no período)', 0, 0, 0, 0]]);
    sh.getRange(3,5).setNumberFormat('0.0%');
    return;
  }

  // Dados
  sh.getRange(3,1,linhas.length,5).setValues(linhas);
  sh.getRange(3,5,linhas.length,1).setNumberFormat('0.0%');

  // TOTAL (até agora)
  const totalAgendadas = linhas.reduce((s, r) => s + Number(r[1]||0), 0);
  const totalPresencas = linhas.reduce((s, r) => s + Number(r[2]||0), 0);
  const totalFaltas    = Math.max(0, totalAgendadas - totalPresencas);
  const percGeral      = totalAgendadas > 0 ? totalPresencas / totalAgendadas : 0;

  const totalRowIndex = 3 + linhas.length;
  sh.getRange(totalRowIndex, 1, 1, 5).setValues([
    ['TOTAL (até agora)', totalAgendadas, totalPresencas, totalFaltas, percGeral]
  ]);
  sh.getRange(totalRowIndex, 1, 1, 5).setFontWeight('bold');
  sh.getRange(totalRowIndex, 5).setNumberFormat('0.0%');

  // Filtro só em cabeçalho + dados (TOTAL fora)
  sh.getRange(2,1,linhas.length+1,5).createFilter();
  sh.autoResizeColumns(1,5);
}

/***** =========================
 * GATILHOS
 * =========================*****/
function criarGatilhoDiario() {
  const existentes = ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'atualizar');
  if (existentes.length === 0) {
    ScriptApp.newTrigger('atualizar').timeBased().everyDays(1).atHour(20).create();
  }
  SpreadsheetApp.getActive().toast('Gatilho diário ativo às 20h.', 'Frequência', 5);
}

function removerGatilhos() {
  ScriptApp.getProjectTriggers().forEach(tr => ScriptApp.deleteTrigger(tr));
  SpreadsheetApp.getActive().toast('Todos os gatilhos removidos.', 'Frequência', 5);
}

/***** =========================
 * HELPERS
 * =========================*****/
function weeksForMonthUpToNow(year, month, now) { // month = 1..12
  const monthIdx = month - 1;
  const isPast   = (year < now.getFullYear()) || (year === now.getFullYear() && monthIdx < now.getMonth());
  const isFuture = (year > now.getFullYear()) || (year === now.getFullYear() && monthIdx > now.getMonth());
  if (isPast)   return getWeeksInMonth(year, month);
  if (isFuture) return 0;
  const day = now.getDate();
  return Math.max(1, Math.ceil(day / 7));
}

function getWeeksInMonth(year, month) { // month = 1..12
  const firstDay = new Date(year, month-1, 1);
  const lastDay  = new Date(year, month, 0);
  const days = Math.round((lastDay - firstDay)/(1000*60*60*24)) + 1;
  return Math.ceil(days / 7);
}

function toTitleCase(str) {
  return String(str || '')
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .replace(/\b\w/g, c => c.toUpperCase())
    .trim();
}

function addMonths(date, months) {
  const d = new Date(date);
  d.setMonth(d.getMonth() + months);
  return d;
}

function fmt(d) {
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM/yyyy');
}

function isAcceptedColorId_(colorRaw) {
  const id = (colorRaw == null) ? "" : String(colorRaw);
  return ACCEPTED_COLOR_IDS.indexOf(id) !== -1;
}

// Detecta presença via: ✅, ✔, ✓, (v)
function isPresenteFromTitle_(tituloOriginal) {
  const t = String(tituloOriginal || '');
  return /✅|✔|✓|\(v\)/i.test(t);
}

// Normaliza nome — corta qualificadores (Duo/Reposição/etc.)
function normalizeAlunoFromTitle_(tituloOriginal) {
  let t = String(tituloOriginal || '')
    .replace(/✅|✔|✓/g, ' ')
    .replace(/\(v\)/ig, ' ')
    .replace(/^[–—-]+|[–—-]+$/g, ' ')
    .replace(/^(aula\s+)(m[uú]sica\s+)?/i, ' ');
  // corta no primeiro separador comum
  t = t.split(/[\(\[\-\/#|]/)[0];
  t = t.replace(/\s+/g, ' ').trim();
  return toTitleCase(t);
}

// Chave canônica p/ deduplicar nomes
function canonicalKey_(nome) {
  return String(nome || '')
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/[^a-z0-9 ]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

/***** =========================
 * ALERTS SHEET (criação/duplicados)
 * =========================*****/
function ensureAlertsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(ALERTS_SHEET);
  if (!sh) {
    sh = ss.insertSheet(ALERTS_SHEET);
    sh.appendRow(['Gerado em','Aluno','Marco (meses)','Tipo','Período','Início','Fim','Presenças até agora','Faltas até agora','Futuras no período','Chave']);
  } else if (sh.getLastRow() < 1) {
    sh.appendRow(['Gerado em','Aluno','Marco (meses)','Tipo','Período','Início','Fim','Presenças até agora','Faltas até agora','Futuras no período','Chave']);
  }
  return sh;
}

function makeAlertKey(aluno, m, tipo, start, end) {
  return [aluno, m, tipo, fmt(start), fmt(end)].join('|');
}

function loadExistingAlertKeys(sh) {
  const keys = new Set();
  const lastRow = sh.getLastRow();
  if (lastRow >= 2) {
    const data = sh.getRange(2, 11, lastRow - 1, 1).getValues(); // coluna 'Chave'
    data.forEach(([k]) => { if (k) keys.add(String(k)); });
  }
  return keys;
}

/***** =========================
 * onEdit — atualiza painel quando B1 muda
 * =========================*****/
function onEdit(e) {
  try {
    const rng = e.range;
    const sh  = rng.getSheet();
    if (sh.getName() === MONTH_PANEL_SHEET && rng.getA1Notation() === 'B1') {
      montarPainelMensal();
    }
  } catch (_) {}
}

/***** =========================
 * Utilitários do Painel
 * =========================*****/
function getYearMonthFromB1_(sh) {
  const raw = sh.getRange('B1').getValue();
  if (!raw) return null;
  if (raw instanceof Date) {
    const y = raw.getFullYear();
    const m = ('0' + (raw.getMonth() + 1)).slice(-2);
    return `${y}-${m}`;
  }
  const s = String(raw).trim();
  const mmYYYY = s.match(/^(\d{1,2})[\/\-](\d{4})$/);
  if (mmYYYY) {
    const m = ('0' + Number(mmYYYY[1])).slice(-2);
    const y = Number(mmYYYY[2]);
    if (y >= 2000 && y <= 2100 && Number(m) >= 1 && Number(m) <= 12) {
      return `${y}-${m}`;
    }
  }
  const YYYYmm = s.match(/^(\d{4})-(\d{2})$/);
  if (YYYYmm) return s;
  return null;
}

function monthRangeConsideringNow_(year, month, now) { // month = 1..12
  const start = new Date(year, month - 1, 1, 0, 0, 0);
  let end = new Date(year, month, 0, 23, 59, 59);
  const isCurrent = (year === now.getFullYear() && (month - 1) === now.getMonth());
  if (isCurrent) end = new Date(now); // até agora
  return { start, end };
}

/***** =========================
 * CONSULTA POR PERÍODO
 * =========================*****/
function consultarPeriodoAluno() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ensureConsultaPeriodoSheet_(ss); // garante layout SEM limpar inputs

  const vIni = sh.getRange('B1').getValue();
  const vFim = sh.getRange('B2').getValue();
  const alunoInput = String(sh.getRange('B3').getValue() || '').trim();

  if (!vIni || !vFim || !alunoInput) {
    sh.getRange('A8').setValue('Preencha B1 (início), B2 (fim) e B3 (aluno) e rode de novo.');
    return;
  }

  // Parser robusto + normalizações + swap se invertido
  let inicio = normalizeAsDateStart_(vIni);
  let fim    = normalizeAsDateEnd_(vFim);
  if (inicio > fim) { const tmp = inicio; inicio = fim; fim = tmp; }
  const alunoKey = canonicalKey_(toTitleCase(alunoInput));

  // Mostra o intervalo interpretado (debug amigável)
  sh.getRange('A9').setValue('Início interpretado:');
  sh.getRange('B9').setValue(inicio).setNumberFormat('dd/MM/yyyy HH:mm');
  sh.getRange('A10').setValue('Fim interpretado:');
  sh.getRange('B10').setValue(fim).setNumberFormat('dd/MM/yyyy HH:mm');

  // Lê RAW
  const raw = ss.getSheetByName(RAW_SHEET);
  if (!raw || raw.getLastRow() < 2) {
    sh.getRange('A8').setValue('Sem dados no RAW. Rode "Frequência → Atualizar agora".');
    return;
  }
  const vals = raw.getRange(2,1,raw.getLastRow()-1,3).getValues(); // Data | Aluno | Status

  let agendadas = 0, presencas = 0, faltas = 0;
  for (const [data, aluno, status] of vals) {
    if (!data || !aluno) continue;
    const d = (data instanceof Date && !isNaN(data)) ? new Date(data) : new Date(data);
    if (d < inicio || d > fim) continue;

    const key = canonicalKey_(toTitleCase(String(aluno)));
    if (key !== alunoKey) continue;

    agendadas++;
    if (/^presente$/i.test(String(status))) presencas++;
    else if (/^falta$/i.test(String(status))) faltas++;
  }

  const perc = agendadas > 0 ? presencas / agendadas : 0;

  sh.getRange('A6:D6').clearContent();
  sh.getRange('A6:D6').setValues([[agendadas, presencas, faltas, perc]]);
  sh.getRange('D6').setNumberFormat('0.0%');
  sh.getRange('A8').setValue('Atualizado em: ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'));
}

function ensureConsultaPeriodoSheet_(ss) {
  let sh = ss.getSheetByName('Consulta_Periodo');
  if (!sh) sh = ss.insertSheet('Consulta_Periodo');

  // NÃO limpar tudo — apenas garante rótulos/formatos se estiverem vazios
  if (!sh.getRange('A1').getValue()) sh.getRange('A1').setValue('Data inicial:');
  sh.getRange('B1').setNumberFormat('dd/mm/yyyy');
  if (!sh.getRange('A2').getValue()) sh.getRange('A2').setValue('Data final:');
  sh.getRange('B2').setNumberFormat('dd/mm/yyyy');
  if (!sh.getRange('A3').getValue()) sh.getRange('A3').setValue('Aluno:');

  if (!sh.getRange('A5').getValue()) {
    sh.getRange('A5:D5').setValues([['Agendadas','Presenças','Faltas','% Presença']]).setFontWeight('bold');
    sh.getRange('D6').setNumberFormat('0.0%');
    sh.getRange('A4').setValue('Como usar: preencha B1, B2 e B3 e rode "Frequência → Consulta por período".');
    sh.setColumnWidths(1, 4, 150);
  }

  // Sugerir validação de lista para B3 (deduplicada por chave canônica)
  try {
    const hasValidation = sh.getRange('B3').getDataValidation();
    if (!hasValidation) {
      const raw = ss.getSheetByName(RAW_SHEET);
      if (raw && raw.getLastRow() >= 2) {
        const nomes = raw.getRange(2,2,raw.getLastRow()-1,1).getValues()
          .map(r => String(r[0]||'').trim())
          .filter(Boolean)
          .map(n => toTitleCase(n));

        const freqByKey = {}; // key -> { displayVariant -> count }
        for (const nome of nomes) {
          const key = canonicalKey_(nome);
          if (!key) continue;
          if (!freqByKey[key]) freqByKey[key] = {};
          freqByKey[key][nome] = (freqByKey[key][nome] || 0) + 1;
        }

        const unicos = Object.keys(freqByKey).map(key => {
          const variants = Object.entries(freqByKey[key]); // [display, count]
          variants.sort((a, b) => {
            if (b[1] !== a[1]) return b[1] - a[1]; // mais frequente primeiro
            return a[0].localeCompare(b[0], 'pt-BR'); // empate: alfabética
          });
          return variants[0][0]; // melhor grafia
        }).sort((a, b) => a.localeCompare(b, 'pt-BR'));

        if (unicos.length) {
          const rule = SpreadsheetApp.newDataValidation()
            .requireValueInList(unicos, true)
            .setAllowInvalid(true)
            .build();
          sh.getRange('B3').setDataValidation(rule);
        }
      }
    }
  } catch (_) {}

  return sh;
}

/***** =========================
 * PARSE/Normalização de datas
 * =========================*****/
// Início do dia: 00:00:00
function normalizeAsDateStart_(v) {
  const d = parseDateSmart_(v);
  d.setHours(0,0,0,0);
  return d;
}
// Fim do dia: 23:59:59.999
function normalizeAsDateEnd_(v) {
  const d = parseDateSmart_(v);
  d.setHours(23,59,59,999);
  return d;
}

// Parser abrangente: Date | número | dd/mm/yyyy | dd-mm-yyyy | dd.mm.yyyy | yyyy-mm-dd (+hora opcional)
function parseDateSmart_(v) {
  if (v instanceof Date && !isNaN(v)) {
    return new Date(v.getTime());
  }

  // Número serial estilo Excel/Sheets (base 1899-12-30)
  if (typeof v === 'number' && isFinite(v)) {
    const ms = Math.round((v - 25569) * 86400000); // 25569 = dias até 1970-01-01
    return new Date(ms);
  }

  const s = String(v || '').trim().replace(/\s+/g, ' ');

  // yyyy-mm-dd[ [T]HH:MM[:SS]] ou yyyy/mm/dd...
  let m = s.match(/^(\d{4})[\/\-\.](\d{1,2})[\/\-\.](\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
  if (m) {
    const y = +m[1], mo = +m[2], d = +m[3], hh = +(m[4]||0), mm = +(m[5]||0), ss = +(m[6]||0);
    return new Date(y, mo-1, d, hh, mm, ss, 0);
  }

  // dd/mm/yyyy | dd-mm-yyyy | dd.mm.yyyy [HH:MM[:SS]]
  m = s.match(/^(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{2,4})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
  if (m) {
    let d = +m[1], mo = +m[2], y = +m[3];
    if (y < 100) y += (y >= 70 ? 1900 : 2000); // 2 dígitos → heurística
    const hh = +(m[4]||0), mm = +(m[5]||0), ss = +(m[6]||0);
    return new Date(y, mo-1, d, hh, mm, ss, 0);
  }

  // Fallback controlado
  const guess = new Date(s);
  if (guess instanceof Date && !isNaN(guess)) return guess;

  throw new Error('Data inválida: ' + s);
}


