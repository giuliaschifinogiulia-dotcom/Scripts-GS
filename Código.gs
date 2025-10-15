/***********************
 * MENU
 ***********************/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🗓️ Agenda Livre")
    .addItem("Atualizar horários livres", "atualizarHorariosLivresPeriodo")
    .addItem("Diagnóstico (eventos no período)", "Diagnostico_ListarEventos")
    .addToUi();
}
// compat com nome antigo
function atualizarHorariosLivres(){ return atualizarHorariosLivresPeriodo(); }

/***********************
 * PRINCIPAL
 ***********************/
function atualizarHorariosLivresPeriodo() {
  const ss  = SpreadsheetApp.getActive();
  const cfg = ss.getSheetByName("Config");
  const out = ss.getSheetByName("Horarios_Livres");
  if (!cfg || !out) throw new Error("Crie as abas 'Config' e 'Horarios_Livres'.");

  const CALENDAR_ID  = getCfg_(cfg,"CALENDAR_ID") || "primary";
  const tz           = Session.getScriptTimeZone();

  // --- PERÍODO DINÂMICO: hoje → hoje + 1 mês ---
  const { dtIni, dtFim } = periodoDinamico_(tz);
  if (!(dtIni instanceof Date) || !(dtFim instanceof Date) || dtIni >= dtFim)
    throw new Error("Período inválido.");

  // 1=Dom … 7=Sáb (ex.: 2,3,4,5,6 = seg–sex)
  const DIAS_ATIVOS_STR = String(getCfgDisplay_(cfg,"DIAS_ATIVOS") || "2,3,4,5,6").trim();
  const diasDom1 = DIAS_ATIVOS_STR.split(/[,;]\s*|,/).map(s=>parseInt(s,10)).filter(n=>!isNaN(n));

  // Horários (lidos como DISPLAY para evitar o bug do “1899”)
  const INICIO_MANHA = getCfgDisplay_(cfg,"INICIO_MANHA") || "08:30";
  const FIM_MANHA    = getCfgDisplay_(cfg,"FIM_MANHA")    || "12:30";
  const INICIO_TARDE = getCfgDisplay_(cfg,"INICIO_TARDE") || "14:00";
  const FIM_TARDE    = getCfgDisplay_(cfg,"FIM_TARDE")    || "17:00";
  const SLOT_MIN     = parseInt(getCfgDisplay_(cfg,"DURACAO_SLOT_MIN") || "60", 10);

  // Cores: individual ("" + o que você colocar) e dupla ("1")
  const INDIVID_IDS = splitIdsAllowBlankAsEmpty_(getCfgDisplay_(cfg,"INDIVIDUAL_COLOR_IDS"));
  const DUO_IDS     = splitIds_(getCfgDisplay_(cfg,"DUO_COLOR_IDS"));

  // Busca eventos (Calendar API avançada)
  const eventos = listEvents_(CALENDAR_ID, dtIni, dtFim, tz);

  // Limpa saída
  const lr = out.getLastRow();
  if (lr > 1) out.getRange(2,1,lr-1,7).clearContent();

  const rows = [];
  const d = new Date(dtIni); d.setHours(0,0,0,0);

  while (d <= dtFim) {
    const dowDom1 = getDom1Weekday_(d); // 1=Dom..7=Sáb
    if (diasDom1.includes(dowDom1)) {
      const blocos = [[INICIO_MANHA,FIM_MANHA],[INICIO_TARDE,FIM_TARDE]];
      for (const [iniStr, fimStr] of blocos) {
        if (!iniStr || !fimStr) continue;

        const janelaIni = setTime_(d, iniStr);
        const janelaFim = setTime_(d, fimStr);

        for (let tIni = new Date(janelaIni); tIni < janelaFim; tIni = new Date(tIni.getTime() + SLOT_MIN*60000)) {
          const tFim = new Date(tIni.getTime() + SLOT_MIN*60000);
          if (tFim > janelaFim) break;

          const eventosNoSlot = eventos.filter(ev => !ev.allDay && overlapEv_(ev, tIni, tFim));

          // individual bloqueia tudo
          const temInd = eventosNoSlot.some(ev => hasId_(ev.colorId, INDIVID_IDS));
          if (temInd) continue;

          // cada dupla ocupa 1 de 2 pistas
          const duplas = eventosNoSlot.filter(ev => hasId_(ev.colorId, DUO_IDS)).length;
          const livres = Math.max(0, 2 - duplas);
          if (livres <= 0) continue;

          // Dia da semana em PT-BR
          const diasPt = ["Domingo","Segunda","Terça","Quarta","Quinta","Sexta","Sábado"];
          const diaSemana = diasPt[tIni.getDay()];

          rows.push([
            Utilities.formatDate(tIni, tz, "dd/MM/yyyy"),
            diaSemana,
            Utilities.formatDate(tIni, tz, "HH:mm"),
            Utilities.formatDate(tFim, tz, "HH:mm"),
            livres,
            livres >= 1 ? "SIM" : "NÃO",
            livres === 2 ? "SIM" : "NÃO"
          ]);
        }
      }
    }
    d.setDate(d.getDate() + 1);
    d.setHours(0,0,0,0);
  }

  if (rows.length) out.getRange(2,1,rows.length,7).setValues(rows);
  SpreadsheetApp.getUi().alert(`✅ ${rows.length} slots disponíveis listados.`);
}

// Helper: hoje 00:00 → hoje + 1 mês 23:59
function periodoDinamico_(tz) {
  const agora = new Date();
  const hoje = new Date(agora.getFullYear(), agora.getMonth(), agora.getDate(), 0, 0, 0, 0);
  const fim = new Date(hoje);
  fim.setMonth(fim.getMonth() + 1);
  fim.setHours(23,59,59,999);
  return { dtIni: hoje, dtFim: fim };
}


/***********************
 * DIAGNÓSTICO
 ***********************/
function Diagnostico_ListarEventos(){
  const ss  = SpreadsheetApp.getActive();
  const cfg = ss.getSheetByName("Config");
  const tz  = Session.getScriptTimeZone();
  const CALENDAR_ID = getCfg_(cfg,"CALENDAR_ID") || "primary";
  const ini = parseFlexibleDate_(getCfgDisplay_(cfg,"DATA_INICIO"), false);
  const fim = parseFlexibleDate_(getCfgDisplay_(cfg,"DATA_FIM"),   true);
  const evs = listEvents_(CALENDAR_ID, ini, fim, tz);

  const sh = ss.getSheetByName("Diagnostico") || ss.insertSheet("Diagnostico");
  sh.clear();
  sh.getRange(1,1,1,6).setValues([["title","start","end","colorId","allDay","calendarId"]]);
  const rows = evs.map(e => [
    e.summary || "",
    e.startStr, e.endStr,
    String(e.colorId ?? ""),
    e.allDay ? "Y" : "",
    CALENDAR_ID
  ]);
  if (rows.length) sh.getRange(2,1,rows.length,6).setValues(rows);
  SpreadsheetApp.getUi().alert(`Diagnóstico: ${rows.length} eventos.`);
}

/***********************
 * HELPERS
 ***********************/
// Lê o VALOR bruto (pode ser Date/number/etc)
function getCfg_(sheet, key){
  const vals = sheet.getRange(1,1,sheet.getLastRow(),2).getValues();
  for (let i=0;i<vals.length;i++) if ((vals[i][0]+"").trim()===key) return vals[i][1];
  return "";
}
// Lê o DISPLAY (texto visível na célula)
function getCfgDisplay_(sheet, key) {
  const rng = sheet.getRange(1,1,sheet.getLastRow(),2);
  const values = rng.getValues();
  const displays = rng.getDisplayValues();
  for (let i = 0; i < values.length; i++) {
    if ((values[i][0] + "").trim() === key) return (displays[i][1] + "").trim();
  }
  return "";
}

// Divide IDs genéricos (retorna [] se vazio)
function splitIds_(raw){
  if (raw === null || raw === undefined) return [];
  if (Array.isArray(raw)) raw = raw.join(",");
  const s = String(raw).trim();
  if (s === "") return [];
  return s.split(/[;,]/).map(t=>t.trim()).map(t=> t==='""' ? "" : t);
}
// Específico para INDIVIDUAL_COLOR_IDS: célula vazia => [""] (id vazio)
function splitIdsAllowBlankAsEmpty_(raw){
  if (raw === null || raw === undefined) return [""];
  if (Array.isArray(raw)) raw = raw.join(",");
  const s = String(raw).trim();
  if (s === "") return [""];
  return s.split(/[;,]/).map(t=>t.trim()).map(t=> t==='""' ? "" : t);
}

// Converte dd/MM[/yyyy] [HH:mm]
function parseFlexibleDate_(str, endOfDay){
  if (!str) return null;
  str = String(str).trim();
  // número serial?
  if (!isNaN(str) && !str.includes("/")) {
    const base = new Date(Math.round((Number(str)-25569)*86400*1000));
    base.setHours(endOfDay?23:0, endOfDay?59:0, 0, 0);
    return base;
  }
  const parts = str.split(" ");
  const [d,m,yRaw] = (parts[0] || "").split("/");
  const y = yRaw ? parseInt(yRaw,10) : new Date().getFullYear();
  let hh=0, mm=0;
  if (parts[1]) { const [H,M] = parts[1].split(":"); hh=parseInt(H,10)||0; mm=parseInt(M,10)||0; }
  else { hh=endOfDay?23:0; mm=endOfDay?59:0; }
  return new Date(y, (parseInt(m,10)||1)-1, parseInt(d,10)||1, hh, mm, 0, 0);
}

// Aceita "08:30" OU Date "1899-12-30 08:30" OU serial de hora
function setTime_(dateBase, timeVal) {
  const d = new Date(dateBase);
  if (timeVal instanceof Date) {
    d.setHours(timeVal.getHours(), timeVal.getMinutes(), 0, 0);
    return d;
  }
  const s = String(timeVal || "").trim();
  const m = s.match(/^(\d{1,2}):(\d{2})(?::\d{2})?$/);
  if (m) {
    const hh = parseInt(m[1], 10) || 0;
    const mm = parseInt(m[2], 10) || 0;
    d.setHours(hh, mm, 0, 0);
    return d;
  }
  if (!isNaN(s) && !s.includes("/")) {
    const serial = Number(s);
    const totalMinutes = Math.round(serial * 24 * 60);
    const hh = Math.floor(totalMinutes / 60);
    const mm = totalMinutes % 60;
    d.setHours(hh % 24, mm, 0, 0);
    return d;
  }
  d.setHours(0, 0, 0, 0);
  return d;
}

// 1=Dom .. 7=Sáb (compatível com teu uso: 2..6 seg–sex)
function getDom1Weekday_(date){ const js=date.getDay(); return (js===0)?1:(js+1); }
function hasId_(id, list){ return list.includes(String(id ?? "")); }
function overlapEv_(ev, s, e){ return ev.start < e && ev.end > s; }

// Calendar Advanced API (pega colorId correto)
function listEvents_(calendarId, start, end, tz){
  const timeMin = new Date(start.getTime() - start.getTimezoneOffset()*60000).toISOString();
  const timeMax = new Date(end.getTime()   - end.getTimezoneOffset()*60000).toISOString();
  const items = [];
  let pageToken;
  do {
    const resp = Calendar.Events.list(calendarId, {
      timeMin, timeMax, singleEvents: true, orderBy: "startTime",
      pageToken, maxResults: 2500
    });
    (resp.items || []).forEach(it => {
      const allDay = !!it.start?.date && !it.start?.dateTime;
      const s = allDay ? new Date(it.start.date+"T00:00:00") : new Date(it.start.dateTime);
      const e = allDay ? new Date(it.end.date  +"T00:00:00") : new Date(it.end.dateTime);
      items.push({
        summary: it.summary || "",
        colorId: (it.colorId != null) ? String(it.colorId) : "",
        allDay: allDay,
        start: s,
        end: e,
        startStr: Utilities.formatDate(s, tz, "dd/MM/yyyy HH:mm"),
        endStr:   Utilities.formatDate(e, tz, "dd/MM/yyyy HH:mm"),
      });
    });
    pageToken = resp.nextPageToken;
  } while (pageToken);
  return items;
}
function periodoDinamico_(tz) {
  const agora = new Date();
  const hoje = new Date(agora.getFullYear(), agora.getMonth(), agora.getDate(), 0, 0, 0, 0);
  const fim = new Date(hoje);
  // +1 mês (mantém fim do dia)
  fim.setMonth(fim.getMonth() + 1);
  fim.setHours(23,59,59,999);
  return { dtIni: hoje, dtFim: fim };
}

