/********************
 * CONFIGURAÇÕES
 ********************/
const CALENDAR_ID   = 'primary'; // troque se não for o calendário principal
const NAME_COL      = 'D';       // coluna NOME ALUNO/PACIENTE
const MODALITY_COL  = 'E';       // coluna MODALIDADE (vamos preencher 'i' ou 'd')
const QTY_COL       = 'H';       // coluna QUANTIDADE AULA MÊS
const FIRST_DATA_ROW = 11;       // <<< ajuste se seus alunos começarem em outra linha

// Cores no seu calendário: laranja="", lavanda="1", amarelo="5" (ignorar)
const INDIVIDUAL_COLORS = [""];   // laranja claro → individual
const DUO_COLOR         = "1";    // lavanda → dupla
const EXCLUDED_COLORS   = ["5"];  // amarelo (experimental/cortesia) → ignorar

// Checkmarks aceitos (título ou descrição)
const CHECK_PAT = /✅|✔|☑|✓/;

/********************
 * MENU
 ********************/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Fechamento do Mês')
    .addItem('Atualizar todas as abas', 'runUpdateAllMonths')
    .addItem('Atualizar mês atual', 'runUpdateCurrentMonth') // ← novo botão
    .addItem('Zerar todas as abas', 'resetAllMonths')
    .addItem('Diagnóstico (cores de hoje)', 'logColorsToday')
    .addToUi();
}

/********************
 * PRINCIPAL – atualiza TODAS as abas (Jan..Dez)
 ********************/
function runUpdateAllMonths() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const cal = CalendarApp.getCalendarById(CALENDAR_ID);

  const year  = new Date().getFullYear();
  const start = new Date(year, 0, 1, 0, 0, 0);
  const end   = new Date(year + 1, 0, 1, 0, 0, 0);
  const events = cal.getEvents(start, end);

  // monthCounters[mes][normName] = { count, display, modality: 'i'|'d' }
  const monthCounters = {};
  let processed = 0;

  events.forEach(ev => {
    const color = ev.getColor() || "";
    if (EXCLUDED_COLORS.includes(color)) return;     // ignora amarelo

    const modality = INDIVIDUAL_COLORS.includes(color)
      ? 'i'
      : (color === DUO_COLOR ? 'd' : null);
    if (!modality) return;                           // só laranja/lavanda

    const title = (ev.getTitle() || '');
    const desc  = (ev.getDescription() || '');
    if (!CHECK_PAT.test(title) && !CHECK_PAT.test(desc)) return; // só presença

    const monthName = getMonthNamePt_(ev.getStartTime().getMonth());
    const sheet = ss.getSheetByName(monthName);
    if (!sheet) return;

    const rawName = extractSingleNameFromTitle_(title);
    const norm    = normalizeName_(rawName);
    if (!norm) return;

    if (!monthCounters[monthName]) monthCounters[monthName] = {};
    if (!monthCounters[monthName][norm]) {
      monthCounters[monthName][norm] = { count: 0, display: rawName, modality };
    } else {
      monthCounters[monthName][norm].count += 0; // só pra não perder referência
      // se ainda não tinha modalidade (não deve ocorrer), define
      if (!monthCounters[monthName][norm].modality) {
        monthCounters[monthName][norm].modality = modality;
      }
      // mantemos o display "mais longo"
      monthCounters[monthName][norm].display =
        pickBetterDisplay_(monthCounters[monthName][norm].display, rawName);
    }
    monthCounters[monthName][norm].count += 1;
    processed++;
  });

  Object.keys(monthCounters).forEach(monthName => {
    writeCountsToSheet_(ss.getSheetByName(monthName), monthCounters[monthName]);
  });

  SpreadsheetApp.getUi().alert(`Eventos processados: ${processed}`);
}

/********************
 * TESTE RÁPIDO – atualiza apenas Agosto
 ********************/
function runUpdateOnlyMonth_Agosto() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const cal = CalendarApp.getCalendarById(CALENDAR_ID);

  const year  = new Date().getFullYear();
  const start = new Date(year, 7, 1, 0, 0, 0); // Agosto
  const end   = new Date(year, 8, 1, 0, 0, 0); // Setembro
  const events = cal.getEvents(start, end);

  const counter = {}; // norm -> {count, display, modality}
  let processed = 0;

  events.forEach(ev => {
    const color = ev.getColor() || "";
    if (EXCLUDED_COLORS.includes(color)) return;

    const modality = INDIVIDUAL_COLORS.includes(color)
      ? 'i'
      : (color === DUO_COLOR ? 'd' : null);
    if (!modality) return;

    const title = (ev.getTitle() || '');
    const desc  = (ev.getDescription() || '');
    if (!CHECK_PAT.test(title) && !CHECK_PAT.test(desc)) return;

    const rawName = extractSingleNameFromTitle_(title);
    const norm    = normalizeName_(rawName);
    if (!norm) return;

    if (!counter[norm]) counter[norm] = { count: 0, display: rawName, modality };
    counter[norm].count    += 1;
    counter[norm].display   = pickBetterDisplay_(counter[norm].display, rawName);
    if (!counter[norm].modality) counter[norm].modality = modality;

    processed++;
  });

  writeCountsToSheet_(ss.getSheetByName('Agosto'), counter);
  SpreadsheetApp.getUi().alert(`Agosto processado: ${processed} eventos`);
}

/********************
 * ESCREVE NA ABA:
 * - Zera H nas linhas existentes;
 * - Atualiza quem já existe;
 * - Se NÃO existir, insere NA PRIMEIRA LINHA VAZIA DO BLOCO
 *   e grava Nome (D), Modalidade (E) e Quantidade (H).
 ********************/
function writeCountsToSheet_(sheet, counterObj) {
  if (!sheet) return false;

  const lastRow = sheet.getLastRow();

  // 1) Zera coluna H no bloco de dados, se houver
  if (lastRow >= FIRST_DATA_ROW) {
    const qtyRange = sheet.getRange(`${QTY_COL}${FIRST_DATA_ROW}:${QTY_COL}${lastRow}`);
    const qtyVals  = qtyRange.getValues();
    for (let i = 0; i < qtyVals.length; i++) qtyVals[i][0] = '';
    qtyRange.setValues(qtyVals);
  }

  // 2) Índice de nomes existentes -> linha
  let nameIndex = new Map();
  if (lastRow >= FIRST_DATA_ROW) {
    const nameVals = sheet.getRange(`${NAME_COL}${FIRST_DATA_ROW}:${NAME_COL}${lastRow}`).getValues().flat();
    for (let i = 0; i < nameVals.length; i++) {
      const nm = normalizeName_(String(nameVals[i] || ''));
      if (!nm) continue;
      nameIndex.set(nm, FIRST_DATA_ROW + i);
    }
  }

  // 3) Atualiza ou inclui
  Object.keys(counterObj).forEach(norm => {
    const { count, display, modality } = counterObj[norm] || {};
    if (!count) return;

    const existingRow = nameIndex.get(norm);
    if (existingRow) {
      // Atualiza quantidade; se modalidade estiver vazia, preenche
      sheet.getRange(`${QTY_COL}${existingRow}`).setValue(count);
      const modCell = sheet.getRange(`${MODALITY_COL}${existingRow}`);
      if (!String(modCell.getValue() || '').trim()) modCell.setValue(modality || '');
    } else {
      // Encontra primeira linha vazia no bloco (coluna D sem valor)
      let insertRow = null;
      const currentLast = sheet.getLastRow();
      for (let r = FIRST_DATA_ROW; r <= currentLast; r++) {
        const v = sheet.getRange(`${NAME_COL}${r}`).getValue();
        if (!v) { insertRow = r; break; }
      }
      if (!insertRow) insertRow = currentLast + 1;

      // Nome (D), Modalidade (E), Quantidade (H)
      sheet.getRange(`${NAME_COL}${insertRow}`).setValue(display || norm);
      sheet.getRange(`${MODALITY_COL}${insertRow}`).setValue(modality || '');
      sheet.getRange(`${QTY_COL}${insertRow}`).setValue(count);

      // Atualiza índice caso apareça de novo
      nameIndex.set(norm, insertRow);
    }
  });

  return true;
}

/********************
 * ZERAR TODAS AS ABAS (col. H)
 ********************/
function resetAllMonths() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const months = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho',
                  'Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'];
  months.forEach(m => {
    const sheet = ss.getSheetByName(m);
    if (!sheet) return;
    const lastRow = sheet.getLastRow();
    if (lastRow < FIRST_DATA_ROW) return;
    const rng = sheet.getRange(`${QTY_COL}${FIRST_DATA_ROW}:${QTY_COL}${lastRow}`);
    const vals = rng.getValues();
    for (let i = 0; i < vals.length; i++) vals[i][0] = '';
    rng.setValues(vals);
  });
  SpreadsheetApp.getUi().alert('Todas as abas zeradas 🧽');
}

/********************
 * DIAGNÓSTICO: lista cores de hoje
 ********************/
function logColorsToday() {
  const cal = CalendarApp.getCalendarById(CALENDAR_ID);
  const start = new Date(); start.setHours(0,0,0,0);
  const end   = new Date(); end.setHours(23,59,59,999);
  const events = cal.getEvents(start, end);
  events.forEach(ev => Logger.log(`Título: ${ev.getTitle()} | Cor retornada: ${ev.getColor()}`));
}

/********************
 * HELPERS
 ********************/
function extractSingleNameFromTitle_(title) {
  let t = title.replace(/dupla[:\-]?/i,'')
               .replace(/atendimento[:\-]?/i,'')
               .replace(/exp[:\-]?/i,'')
               .replace(CHECK_PAT,'')
               .trim();
  // usa o primeiro "token" caso tenha separadores
  t = t.split(/[,|+|&|\-|–|—]/)[0];
  // remove horários e números comuns no título
  t = t.replace(/\b\d{1,2}[:h]\d{0,2}\b/gi,'').replace(/\b\d{1,2}\b/g,'').trim();
  return t;
}

function normalizeName_(s) {
  if (!s) return '';
  return s.toLowerCase()
          .normalize('NFD').replace(/[\u0300-\u036f]/g,'')
          .replace(/[^a-z\s]/g,'')
          .replace(/\s+/g,' ')
          .trim();
}

function pickBetterDisplay_(a, b) {
  const A = (a || '').trim(), B = (b || '').trim();
  if (!A) return B;
  if (!B) return A;
  return B.length > A.length ? B : A;
}

function getMonthNamePt_(idx) {
  const meses = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho',
                 'Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'];
  return meses[idx];
}
function runUpdateAllMonths() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const cal = CalendarApp.getCalendarById(CALENDAR_ID);

  const year  = new Date().getFullYear();
  const start = new Date(year, 0, 1, 0, 0, 0);
  const end   = new Date(year + 1, 0, 1, 0, 0, 0);
  const events = cal.getEvents(start, end);

  // monthCounters[mes][normName] = { count, display, modality: 'i'|'d' }
  const monthCounters = {
    'Janeiro': {}, 'Fevereiro': {}, 'Março': {}, 'Abril': {}, 'Maio': {}, 'Junho': {},
    'Julho': {}, 'Agosto': {}, 'Setembro': {}, 'Outubro': {}, 'Novembro': {}, 'Dezembro': {}
  };

  let processed = 0;

  events.forEach(ev => {
    const color = ev.getColor() || "";
    if (EXCLUDED_COLORS.includes(color)) return; // ignora amarelo

    const modality = INDIVIDUAL_COLORS.includes(color)
      ? 'i'
      : (color === DUO_COLOR ? 'd' : null);
    if (!modality) return; // só laranja/lavanda

    const title = (ev.getTitle() || '');
    const desc  = (ev.getDescription() || '');
    if (!CHECK_PAT.test(title) && !CHECK_PAT.test(desc)) return; // só presença

    const monthName = getMonthNamePt_(ev.getStartTime().getMonth());
    const sheet = ss.getSheetByName(monthName);
    if (!sheet) return; // se não existir a aba, pula

    const rawName = extractSingleNameFromTitle_(title);
    const norm    = normalizeName_(rawName);
    if (!norm) return;

    if (!monthCounters[monthName][norm]) {
      monthCounters[monthName][norm] = { count: 0, display: rawName, modality };
    } else {
      monthCounters[monthName][norm].display =
        pickBetterDisplay_(monthCounters[monthName][norm].display, rawName);
      if (!monthCounters[monthName][norm].modality) {
        monthCounters[monthName][norm].modality = modality;
      }
    }
    monthCounters[monthName][norm].count += 1;
    processed++;
  });

  // 🔁 Agora escrevemos em TODAS as abas (mesmo que não haja presenças → zera H)
  const months = Object.keys(monthCounters);
  months.forEach(monthName => {
    const sheet = ss.getSheetByName(monthName);
    if (!sheet) return;
    writeCountsToSheet_(sheet, monthCounters[monthName]); // passa {} quando vazio → zera H
  });

  SpreadsheetApp.getUi().alert(`Eventos processados: ${processed}\nMeses atualizados: ${months.length}`);
}
function runUpdateCurrentMonth() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const cal = CalendarApp.getCalendarById(CALENDAR_ID);

  const now   = new Date();
  const year  = now.getFullYear();
  const month = now.getMonth(); // 0 = Jan
  const start = new Date(year, month, 1, 0, 0, 0);
  const end   = new Date(year, month + 1, 1, 0, 0, 0);

  const events = cal.getEvents(start, end);
  const counter = {}; // norm -> {count, display, modality}
  let processed = 0;

  events.forEach(ev => {
    const color = ev.getColor() || "";
    if (EXCLUDED_COLORS.includes(color)) return;

    const modality = INDIVIDUAL_COLORS.includes(color)
      ? 'i'
      : (color === DUO_COLOR ? 'd' : null);
    if (!modality) return;

    const title = (ev.getTitle() || '');
    const desc  = (ev.getDescription() || '');
    if (!CHECK_PAT.test(title) && !CHECK_PAT.test(desc)) return;

    const rawName = extractSingleNameFromTitle_(title);
    const norm    = normalizeName_(rawName);
    if (!norm) return;

    if (!counter[norm]) counter[norm] = { count: 0, display: rawName, modality };
    counter[norm].count  += 1;
    counter[norm].display = pickBetterDisplay_(counter[norm].display, rawName);
    if (!counter[norm].modality) counter[norm].modality = modality;

    processed++;
  });

  const monthName = getMonthNamePt_(month);
  const sheet = ss.getSheetByName(monthName);
  if (!sheet) {
    SpreadsheetApp.getUi().alert(`Aba "${monthName}" não encontrada. Crie a aba para este mês.`);
    return;
  }

  // writeCountsToSheet_ já zera a coluna H do bloco e atualiza/inclui
  writeCountsToSheet_(sheet, counter);

  SpreadsheetApp.getUi().alert(`${monthName} processado: ${processed} eventos`);
}

