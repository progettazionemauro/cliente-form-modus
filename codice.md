/***************
 * CONFIG
 ***************/
const SHEET_NAME = "SWOT_test";   // <-- QUI: niente SWOT_v3
const API_KEY = "Djungo-Test-2026-01-17-chiave-lunga-XYZ"; // deve combaciare con index.html

// Intestazioni attese (solo per chiarezza; non serve che siano identiche nel foglio)
const HEADERS = [
  "ID",
  "Anno",
  "Contesto",
  "Attività",
  "Parti interessate",
  "Obiettivo (Ante)",
  "SWOT (Ante)",
  "Fattore Interno - Esterno (Ante)",
  "R / O (Ante)",
  "Anno di Rilevazione (POST)",
  "Esito POST",
  "SWOT (POST)",
  "Fattore Interno - Esterno (POST)",
  "R / O (POST)",
  "Commento POST"
];

/***************
 * Entry point
 ***************/
function doGet(e) {
  const p = (e && e.parameter) ? e.parameter : {};
  const cb = p.cb;

  if (!cb) return ContentService.createTextOutput("Missing cb (JSONP)");

  try {
    const mode = String(p.mode || "").trim();
    const debug = String(p.debug || "").trim() === "1";

    // LOG: sempre, così lo vedi in "Esecuzioni"
    console.log("doGet mode=", mode, "debug=", debug);
    console.log("PARAMS=", JSON.stringify(p));

    if (mode === "test") {
      return jsonp_(cb, { ok: true, msg: "OK test", sheet: SHEET_NAME });
    }

    if (mode === "view") {
      const limit = clampInt_(p.limit, 1, 200, 20);
      const out = handleView_(limit);
      return jsonp_(cb, out);
    }

    if (mode === "insert") {
      const apiKey = String(p.apiKey || "");
      if (apiKey !== API_KEY) {
        const out = { ok: false, error: "Invalid apiKey" };
        if (debug) out.debug = { receivedApiKey: apiKey };
        return jsonp_(cb, out);
      }

      const out = handleInsert_(p, debug);
      return jsonp_(cb, out);
    }

    return jsonp_(cb, { ok: false, error: "Unknown mode" });

  } catch (err) {
    console.error("ERROR:", err);
    return jsonp_(cb, { ok: false, error: String(err && err.message ? err.message : err) });
  }
}

/***************
 * Insert
 ***************/


function handleInsert_(p, debug) {
  const sh = getSheet_();

  // Leggo i campi ANTE
  const anno = norm_(p.anno);
  const contesto = norm_(p.contesto);
  const attivita = norm_(p.attivita);
  const partiInteressate = norm_(p.partiInteressate);
  const obiettivoAnte = norm_(p.obiettivoAnte);
  const swotAnte = norm_(p.swotAnte);

  // DEBUG snapshot campi
  const snap = {
    anno, contesto, attivita,
    partiInteressate, obiettivoAnte, swotAnte,
    keys: Object.keys(p).sort()
  };

  // Validazione required ANTE
  if (!anno || !contesto || !attivita || !partiInteressate || !obiettivoAnte || !swotAnte) {
    const missing = [];
    if (!anno) missing.push("anno");
    if (!contesto) missing.push("contesto");
    if (!attivita) missing.push("attivita");
    if (!partiInteressate) missing.push("partiInteressate");
    if (!obiettivoAnte) missing.push("obiettivoAnte");
    if (!swotAnte) missing.push("swotAnte");

    const out = { ok: false, error: "Missing required fields (ANTE)", missing, snap };
    if (debug) out.debug = snap;

    console.log("MISSING FIELDS:", missing, "SNAP=", JSON.stringify(snap));
    return out;
  }

  // ... resto invariato (deriveIE_, deriveRO_, appendRow, ecc.)
  const fattoreIEAnte = deriveIE_(swotAnte);
  const roAnte = deriveRO_(swotAnte);
  if (!fattoreIEAnte) {
    const out = { ok: false, error: "SWOT (ANTE) must contain (E) or (I)" };
    if (debug) out.debug = snap;
    return out;
  }

  const annoPost = norm_(p.annoPost);
  const esitoPost = norm_(p.esitoPost);
  const swotPost = norm_(p.swotPost);
  const commentoPost = norm_(p.commentoPost);

  const fattoreIEPost = swotPost ? deriveIE_(swotPost) : "";
  const roPost = swotPost ? deriveRO_(swotPost) : "";

  if (swotPost && !fattoreIEPost) {
    const out = { ok: false, error: "SWOT (POST) must contain (E) or (I)" };
    if (debug) out.debug = snap;
    return out;
  }

  const nextId = computeNextId_(sh);

  const row = [
    nextId,
    anno,
    contesto,
    attivita,
    partiInteressate,
    obiettivoAnte,
    swotAnte,
    fattoreIEAnte,
    roAnte,
    annoPost,
    esitoPost,
    swotPost,
    fattoreIEPost,
    roPost,
    commentoPost
  ];

  sh.appendRow(row);

  const okOut = { ok: true, id: nextId };
  if (debug) okOut.debug = snap;
  return okOut;
}



/***************
 * View
 ***************/
function handleView_(limit) {
  const sh = getSheet_();
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();

  if (lastRow < 1 || lastCol < 1) {
    return { ok: true, headers: HEADERS, rows: [] };
  }

  // Provo a leggere intestazioni reali in riga 1 (se presenti); altrimenti uso HEADERS
  let headers = [];
  let dataStartRow = 1;

  const firstRow = sh.getRange(1, 1, 1, Math.min(lastCol, HEADERS.length)).getValues()[0];
  const looksLikeHeader = String(firstRow[0] || "").toUpperCase() === "ID";

  if (looksLikeHeader) {
    headers = firstRow.map(v => String(v ?? ""));
    dataStartRow = 2;
  } else {
    headers = HEADERS.slice();
    dataStartRow = 1;
  }

  const availableRows = lastRow - dataStartRow + 1;
  const take = Math.max(0, Math.min(limit, availableRows));

  if (take === 0) return { ok: true, headers, rows: [] };

  // Prendo le ultime "take" righe, e solo fino alle 15 colonne attese
  const numCols = Math.min(lastCol, HEADERS.length);
  const start = lastRow - take + 1;

  const rows = sh.getRange(start, 1, take, numCols).getValues();
  return { ok: true, headers: headers.slice(0, numCols), rows };
}

/***************
 * Helpers: derivazioni
 ***************/
function deriveIE_(swotText) {
  const s = String(swotText || "");
  if (s.indexOf("(E)") !== -1) return "Fattore Esterno";
  if (s.indexOf("(I)") !== -1) return "Fattore Interno";
  return "";
}

// Replico la tua formula:
// IF(G4="";"";IF(REGEXMATCH(G4;"OPPORTUNITA|FORZA");"OPPORTUNITA'";"RISCHIO"))
function deriveRO_(swotText) {
  const s = String(swotText || "");
  if (!s) return "";
  const re = /OPPORTUNITA|FORZA/i;
  return re.test(s) ? "OPPORTUNITA'" : "RISCHIO";
}

/***************
 * Helpers: sheet + utils
 ***************/
function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) throw new Error("Sheet not found: " + SHEET_NAME);
  return sh;
}

function computeNextId_(sh) {
  // Se riga1 è header, il primo ID sarà 1
  const lastRow = sh.getLastRow();
  if (lastRow < 1) return 1;

  // se c'è header in A1, allora last data row può essere 1 (solo header)
  const a1 = String(sh.getRange(1, 1).getValue() || "").toUpperCase();
  const hasHeader = (a1 === "ID");
  if (hasHeader && lastRow === 1) return 1;

  // Leggo l’ultimo ID in colonna A e incremento se numerico
  const lastId = sh.getRange(lastRow, 1).getValue();
  const n = Number(lastId);
  if (!isNaN(n) && isFinite(n) && n >= 0) return n + 1;

  // fallback
  return lastRow + (hasHeader ? -1 : 0);
}

function clampInt_(v, min, max, defVal) {
  const n = parseInt(String(v || ""), 10);
  if (isNaN(n)) return defVal;
  return Math.max(min, Math.min(max, n));
}

function norm_(v) {
  return String(v ?? "").trim();
}

function jsonp_(cb, obj) {
  const out = cb + "(" + JSON.stringify(obj) + ");";
  return ContentService
    .createTextOutput(out)
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function text_(s) {
  return ContentService.createTextOutput(String(s));
}

/***************
 * Funzione di formattazione
 ***************/

function formatSWOTSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) throw new Error("Sheet not found: " + SHEET_NAME);

  // Colonne attese (max 15)
  const lc = Math.min(sh.getLastColumn(), HEADERS.length);
  if (lc < 1) return;

  // Header detection / creation
  const a1 = String(sh.getRange(1, 1).getValue() || "").toUpperCase();
  const hasHeader = (a1 === "ID");
  if (!hasHeader) {
    sh.getRange(1, 1, 1, lc).setValues([HEADERS.slice(0, lc)]);
  }

  const lr = sh.getLastRow();
  if (lr < 1) return;

  // === 1) Header style ===
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, lc)
    .setFontWeight("bold")
    .setBackground("#f3d36b")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setWrap(true);

  // === 2) Reset area data (solo fino a lr) e banding righe alternate (visibile) ===
  if (lr >= 2) {
    const data = sh.getRange(2, 1, lr - 1, lc);

    // Reset sfondi e bordi nell’area dati (così i colori "si vedono")
    data.setBackground(null);
    data.setBorder(false, false, false, false, false, false);

    // Banding: rimuovo eventuali banding precedenti e applico il mio
    sh.getBandings().forEach(b => b.remove());
    const banding = data.applyRowBanding();
    banding.setFirstRowColor("#e8f5e9");   // verde chiarissimo
    banding.setSecondRowColor("#c8e6c9");  // verde chiaro
    banding.setHeaderRowColor("#f3d36b");
  }

  // === 3) Wrap (a capo automatico) mirato per colonne testo (Attività, Obiettivo, Commento) ===
  // Indici 1-based: 4=Attività, 6=Obiettivo, 15=Commento POST
  const wrapCols = [4, 6, 15];
  wrapCols.forEach(col => {
    if (col <= lc && lr >= 2) {
      sh.getRange(2, col, lr - 1, 1).setWrap(true);
    }
  });

  // Allineamento generale
  sh.getRange(2, 1, Math.max(0, lr - 1), lc).setVerticalAlignment("top");

  // === 4) Conditional formatting su R/O ANTE e POST ===
  // 9 = R/O ANTE ; 14 = R/O POST
  const colROAnte = 9;
  const colROPost = 14;

  // Togli SOLO regole che insistono su col 9 o col 14, e ricreale pulite
  const keep = sh.getConditionalFormatRules().filter(rule => {
    return !rule.getRanges().some(r => {
      const c = r.getColumn();
      return (c === colROAnte || c === colROPost);
    });
  });

  if (lr >= 2) {
    const rAnte = sh.getRange(2, colROAnte, lr - 1, 1);
    const rPost = sh.getRange(2, colROPost, lr - 1, 1);

    const riskBg = "#f8d7da";  // alert
    const oppBg  = "#d4edda";  // tranquillo

    keep.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo("RISCHIO")
        .setBackground(riskBg)
        .setFontWeight("bold")
        .setRanges([rAnte, rPost])
        .build()
    );

    keep.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo("OPPORTUNITA'")
        .setBackground(oppBg)
        .setFontWeight("bold")
        .setRanges([rAnte, rPost])
        .build()
    );
  }

  sh.setConditionalFormatRules(keep);

  // === 5) Bordi leggeri per leggibilità ===
  if (lr >= 1) {
    sh.getRange(1, 1, lr, lc).setBorder(true, true, true, true, true, true, "#e5e5e5", SpreadsheetApp.BorderStyle.SOLID);
  }

  // === 6) Larghezze colonne (leggibili) ===
  if (4 <= lc) sh.setColumnWidth(4, 420);  // Attività
  if (6 <= lc) sh.setColumnWidth(6, 420);  // Obiettivo
  if (15 <= lc) sh.setColumnWidth(15, 380);// Commento POST

  // Colonne “brevi” auto-resize
  [2,3,5,7,8,9,10,11,12,13,14].forEach(c => {
    if (c <= lc) sh.autoResizeColumn(c);
  });
}
