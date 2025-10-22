

// #region Utilities: LS-JSON, Escape, Debounce, Dates, Array ---------------- */

function lsGetJson(key, def = null) {
  try {
    const s = localStorage.getItem(key);
    return s ? JSON.parse(s) : def;
  } catch {
    return def;
  }
}

function lsSetJson(key, obj) {
  try {
    localStorage.setItem(key, JSON.stringify(obj));
  } catch { /* ignore */ }
}

function escapeText(s) {
  // Für Sicherheit bei textContent (bei Verwendung von innerHTML vermeiden)
  return String(s ?? '');
}

function debounce(fn, ms = 100) {
  let t;
  return (...args) => {
    clearTimeout(t);
    t = setTimeout(() => fn(...args), ms);
  };
}

function sameDay(a, b) {
  return a.getFullYear() === b.getFullYear()
    && a.getMonth() === b.getMonth()
    && a.getDate() === b.getDate();
}

function addDays(date, days) {
  const d = new Date(date);
  d.setDate(d.getDate() + days);
  return d;
}

// Feiertage (bundesweit + ein paar bewegliche) */
function easterSunday(y) {
  const a = y % 19,
    b = Math.floor(y / 100),
    c = y % 100,
    d = Math.floor(b / 4),
    e = b % 4,
    f = Math.floor((b + 8) / 25),
    g = Math.floor((b - f + 1) / 3),
    h = (19 * a + b - d - g + 15) % 30,
    i = Math.floor(c / 4),
    k = c % 4,
    l = (32 + 2 * e + 2 * i - h - k) % 7,
    m = Math.floor((a + 11 * h + 22 * l) / 451),
    M = Math.floor((h + l - 7 * m + 114) / 31) - 1,
    D = (h + l - 7 * m + 114) % 31 + 1;
  return new Date(y, M, D);
}

function feiertagsNameFür(d) {
  const y = d.getFullYear();
  const feste = [
    { m: 0, d: 1, n: 'Neujahr' },
    { m: 4, d: 1, n: 'Tag der Arbeit' },
    { m: 9, d: 3, n: 'Tag der Deutschen Einheit' },
    { m: 9, d: 31, n: 'Reformationstag' },
    { m: 11, d: 25, n: '1. Weihnachtstag' },
    { m: 11, d: 26, n: '2. Weihnachtstag' }
  ];
  for (const x of feste) if (d.getMonth() === x.m && d.getDate() === x.d) return x.n;
  const e = easterSunday(y),
    k = addDays(e, -2),
    o = addDays(e, 1),
    c = addDays(e, 39),
    p = addDays(e, 50);
  if (sameDay(d, k)) return 'Karfreitag';
  if (sameDay(d, o)) return 'Ostermontag';
  if (sameDay(d, c)) return 'Christi Himmelfahrt';
  if (sameDay(d, p)) return 'Pfingstmontag';
  return '';
}


// #endregion ---------------------------------------------------------------- */


// #region Globale Konfiguration & State ------------------------------------- */

// UI-Grenzen
const MIN_ZEILEN_PRO_TAG = 1;
const MAX_ZEILEN_PRO_TAG = 30;

// Eingabefelder-Limits
const INPUT_MAX = {
  fahrzeug: 5,
  dienstposten: 20,
  ort: 20,
  von: 5,
  bis: 5,
  bemerkung: 30,
  eingetragen: 30,
};

// Spalten (nach Datum)
const COLS = ['fahrzeug', 'dienstposten', 'ort', 'von', 'bis', 'bemerkung', 'eingetragen'];

// Jahresbereich
const YEAR_START = 2020;
const YEAR_END = 2030;


// Basis-Fahrzeuge kommen NUR aus Excel (und optionale Namen aus Excel / Changes).
let FAHRZEUGE = [];
let FAHRZEUGNAMEN = {};
let ZEILEN_PRO_TAG = MIN_ZEILEN_PRO_TAG; // wird nach Excel-Import angepasst

// Save-Bar
let unsavedChanges = false;

// „Sperren“ – rein UI-seitig (kein echter Schutz)
const PASSWORT = '1234';
let bearbeitungGesperrt = false;

// Laufzeit-Merker für tatsächlich gerenderte Zeilen je Tag (für Speichern)
let LAST_RPD_PER_DAY = [];   // Index 1..tageImMonat

// #endregion ---------------------------------------------------------------- */


// #region Storage-Keys & Meta ------------------------------------------------ */

function speicherKey(jahr, monat) { return `tabelle_${jahr}_${monat}`; }
function speicherMetaKey(jahr, monat) { return `tabelle_meta_${jahr}_${monat}`; }

function speichereMeta(jahr, monat, meta) { lsSetJson(speicherMetaKey(jahr, monat), meta); }
function ladeMeta(jahr, monat) { return lsGetJson(speicherMetaKey(jahr, monat), null); }

function ladeDaten(jahr, monat) {
  const s = localStorage.getItem(speicherKey(jahr, monat));
  return s ? JSON.parse(s) : null;
}

// #endregion ---------------------------------------------------------------- */


// #region UI: Jahresauswahl, Titel, Export-Buttonlabel, Save-Bar ------------- */

function baueJahresAuswahl(start = YEAR_START, end = YEAR_END) {
  const sel = document.getElementById('jahr');
  if (!sel) return;
  sel.innerHTML = '';
  for (let y = start; y <= end; y++) {
    const opt = document.createElement('option');
    opt.value = y;
    opt.textContent = y;
    sel.appendChild(opt);
  }
}

function titelSetzen(jahr, monatIndex) {
  const span = document.getElementById('monatName');
  if (!span) return;
  if (monatIndex === -1) { span.textContent = `Alle Monate ${jahr}`; return; }
  const mName = new Date(jahr, monatIndex).toLocaleString('de-DE', { month: 'long', year: 'numeric' });
  span.textContent = mName.charAt(0).toUpperCase() + mName.slice(1);
}

function updateExportButtonLabel() {
  const btn = document.getElementById('exportXlsxBtn');
  if (!btn) return;
  const monatSel = parseInt(document.getElementById('monat').value, 10);
  const jahr = parseInt(document.getElementById('jahr').value, 10);
  const monate = ["Januar", "Februar", "März", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November", "Dezember"];
  btn.textContent = (monatSel === -1)
    ? `Excel exportieren (Alle Monate ${jahr})`
    : `Excel exportieren (${monate[monatSel]} ${jahr})`;
}

function setUnsaved(on) {
  unsavedChanges = !!on;
  const bar = document.getElementById('saveBar');
  if (!bar) return;
  bar.style.display = on ? 'block' : 'none';
  document.body.classList.toggle('savebar-visible', on);
}

function clearChangeMarkersAndRebase() {
  document.querySelectorAll('#monatsTabelle tbody input').forEach(inp => {
    inp.dataset.orig = inp.value;
    inp.classList.remove('changed');
  });
}

function discardUnsavedChanges() {
  if (document.activeElement && typeof document.activeElement.blur === 'function') {
    document.activeElement.blur();
  }
  document.querySelectorAll('#monatsTabelle tbody input').forEach(inp => {
    if ('orig' in inp.dataset) inp.value = inp.dataset.orig;
    inp.classList.remove('changed');
  });
  setUnsaved(false);
  document.body.classList.add('no-hover');
  setTimeout(() => document.body.classList.remove('no-hover'), 150);
}

function updateChangedFlag(el) {
  const norm = s => String(s ?? '').replace(/\s+/g, ' ').trim();
  el.classList.toggle('changed', norm(el.value) !== norm(el.dataset.orig || ''));
}

// #endregion ---------------------------------------------------------------- */


// #region Layout: Right Rail (Legende/Buttons) ------------------------------- */

function reflowRightRail() {
  const RAIL_RIGHT = 20;
  const GAP = 16;
  const MIN_TOP = 60;

  const legendTopBar = document.querySelector('.legend-card');
  const vehLegend = document.getElementById('vehicleLegend');
  const btns = document.querySelector('.export-buttons');
  const actions = document.getElementById('vehActions');

  const fix = (el, top) => {
    if (!el) return 0;
    el.style.position = 'fixed';
    el.style.right = `${RAIL_RIGHT}px`;
    el.style.top = `${Math.max(MIN_TOP, Math.round(top))}px`;
    return el.offsetHeight || 0;
  };

  const hLegendTopBar = legendTopBar?.offsetHeight || 0;
  const hVehLegend = vehLegend?.offsetHeight || 0;
  const hBtn = btns?.offsetHeight || 0;
  const hActions = actions?.offsetHeight || 0;

  const gapA = (hActions && (hLegendTopBar || hVehLegend || hBtn)) ? GAP : 0;
  const gap1 = (hLegendTopBar && hVehLegend) ? GAP : 0;
  const gap2 = ((hLegendTopBar || hVehLegend) && hBtn) ? GAP : 0;

  const total = hActions + gapA + hLegendTopBar + gap1 + hVehLegend + gap2 + hBtn;

  const viewportH = window.innerHeight || document.documentElement.clientHeight || 800;
  const startTop = Math.max(MIN_TOP, (viewportH - total) / 2);

  let top = startTop;
  top += fix(actions, top) + (hActions ? GAP : 0);
  top += fix(legendTopBar, top) + (hLegendTopBar ? GAP : 0);
  top += fix(vehLegend, top) + (hVehLegend ? GAP : 0);
  fix(btns, top);
}
window.addEventListener('resize', debounce(reflowRightRail, 100));
reflowRightRail();

// #endregion ---------------------------------------------------------------- */


// #region Dom-Helfer: Inputs & Tabellen-Aufnahme ---------------------------- */
// Minimal: Eingabefeld-Erzeuger (ohne type="time")
function makeInput(colKey, value = "") {
  const inp = document.createElement('input');
  inp.type = 'text';                 
  inp.maxLength = INPUT_MAX[colKey] ?? 10;
  inp.value = value;
  return inp;
}

// #endregion ---------------------------------------------------------------- */


// #region Jahr/Monat Auswahl & Aktualisieren -------------------------------- */

function baueGrundAuswahl() {
  baueJahresAuswahl();
  const heute = new Date();
  const jahrSel = document.getElementById('jahr');
  const monatSel = document.getElementById('monat');
  const yr = heute.getFullYear();
  const m = heute.getMonth();
  const yrClamped = (yr >= YEAR_START && yr <= YEAR_END) ? yr : YEAR_START;

  if (jahrSel) jahrSel.value = String(yrClamped);
  if (monatSel) {
    monatSel.value = String(m);
    if (monatSel.value !== String(m)) monatSel.selectedIndex = m;
  }
}

function titelFürDatei() {
  const monate = ["Januar", "Februar", "März", "April", "Mai", "Juni", "Juli",
    "August", "September", "Oktober", "November", "Dezember"];
  const m = parseInt(document.getElementById('monat').value, 10);
  const j = parseInt(document.getElementById('jahr').value, 10);
  return (m === -1) ? `Monatstabelle-AlleMonate-${j}` : `Monatstabelle-${monate[m]}-${j}`;
}

function aktualisieren() {
  const monat = parseInt(document.getElementById('monat').value, 10);
  let jahr = parseInt(document.getElementById('jahr').value, 10);

  if (!Number.isFinite(jahr)) {
    jahr = (new Date()).getFullYear();
    const sel = document.getElementById('jahr');
    if (sel && sel.value !== String(jahr)) sel.value = String(jahr);
  }
  if (jahr < YEAR_START || jahr > YEAR_END) {
    alert(`Bitte ein Jahr zwischen ${YEAR_START} und ${YEAR_END} auswählen.`);
    return;
  }

  const singleTable = document.getElementById('monatsTabelle');
  const multi = document.getElementById('multiMonate');

  if (monat === -1) {
    titelSetzen(jahr, -1);
    anzeigenAlleMonate(jahr);
  } else {
    if (multi) { multi.style.display = 'none'; multi.innerHTML = ''; }
    singleTable.style.display = 'table';
    titelSetzen(jahr, monat);
    tabelleErstellen(jahr, monat);
  }
  afterRenderApply?.();
}

// #endregion ---------------------------------------------------------------- */


// #region Fahrzeug-Legende ------------------------------ */

function renderFahrzeugLegende() {
  const box = document.getElementById('vehicleLegend');
  if (!box) return;

  const ids = [...FAHRZEUGE].sort((a, b) => a.localeCompare(b, 'de', { numeric: true }));

  const rows = ids.map(id => {
    const name = FAHRZEUGNAMEN[id] || '';
    return `
      <div class="veh-legend-row">
        <div class="veh-legend-id">${escapeText(id)}</div>
        <div class="veh-legend-body">
          <div class="veh-legend-head">
            <div class="veh-legend-name">${escapeText(name)}</div>
          </div>
          <div class="veh-meta muted"></div>
        </div>
      </div>
    `;
  }).join('');

  box.innerHTML = `<h3>Fahrzeuge</h3><div class="veh-group">${rows}</div>`;
  reflowRightRail?.();
}

function setVehicleName(id, name) {
  if (!id) return;
  const clean = String(name || '').trim();
  if (clean) {
    FAHRZEUGNAMEN[id] = clean;
  }
}

// #endregion ---------------------------------------------------------------- */


// #region Tabelle: Erstellen, Monats-/Jahresansicht

function tabelleErstellen(jahr, monatIndex, targetTbody = null) {
  const tbody = targetTbody || document.querySelector('#monatsTabelle tbody');
  tbody.innerHTML = '';

  const readOnlyMulti = !!targetTbody;
  const tageImMonat = new Date(jahr, monatIndex + 1, 0).getDate();

  const rowsPerDayToday = Math.max(MIN_ZEILEN_PRO_TAG, FAHRZEUGE.length || 1);

  const loadSavedRowsMapForDay = (j, m, day, dim) => {
    const saved = ladeDaten(j, m) || [];
    const meta = ladeMeta(j, m);
    const tage = dim || new Date(j, m + 1, 0).getDate();

    let savedRPD = parseInt(meta?.rowsPerDay, 10);
    if (!Number.isFinite(savedRPD) || savedRPD < 1) {
      savedRPD = Math.max(1, Math.round(saved.length / Math.max(1, tage)));
    }
    const start = (day - 1) * savedRPD;
    const block = saved.slice(start, start + savedRPD);

    const map = new Map();
    for (const row of block) {
      if (!row || !row.length) continue;
      const id = String(row[0] ?? '').trim();
      if (!id) continue;
      map.set(id, [row[1] ?? '', row[2] ?? '', row[3] ?? '', row[4] ?? '', row[5] ?? '', row[6] ?? '']);
    }
    return map;
  };

  LAST_RPD_PER_DAY = [];
  let zeilenIndexGlobal = 0;

  for (let tag = 1; tag <= tageImMonat; tag++) {
    const datum = new Date(jahr, monatIndex, tag);
    const datumText = datum.toLocaleDateString('de-DE', { weekday: 'long', day: 'numeric', month: 'long' });
    const feiertagName = feiertagsNameFür(datum);

    const rpdToday = rowsPerDayToday;
    LAST_RPD_PER_DAY[tag] = rpdToday;

    const savedMap = loadSavedRowsMapForDay(jahr, monatIndex, tag, tageImMonat);

    for (let zeile = 0; zeile < rpdToday; zeile++) {
      const tr = document.createElement('tr');

      if (feiertagName) tr.classList.add('feiertag');
      if (datum.getDay() === 0 || datum.getDay() === 6) tr.classList.add('weekend');
      if (!tr.classList.contains('weekend') && !tr.classList.contains('feiertag')) {
        if (zeile % 2 === 1) tr.classList.add('row-alt');
      }

      // Datumsspalte
      if (zeile === 0) {
        tr.classList.add('first-of-day');

        const tdDate = document.createElement('td');
        tdDate.classList.add('col-datum');

        const wrap = document.createElement('div');
        const dt = document.createElement('div');
        dt.className = 'datum-text';
        const datumTextFormatted = datumText.charAt(0).toUpperCase() + datumText.slice(1);
        dt.textContent = datumTextFormatted;
        wrap.appendChild(dt);
        if (feiertagName) {
          const fh = document.createElement('div');
          fh.className = 'feiertag-name';
          fh.textContent = feiertagName;
          wrap.appendChild(fh);
        }
        tdDate.appendChild(wrap);
        tdDate.rowSpan = rpdToday;
        tr.appendChild(tdDate);
      }

      // restliche Spalten
      for (let i = 0; i < 7; i++) {
        const colKey = COLS[i];                  // 'fahrzeug' | 'dienstposten' | 'ort' | 'von' | 'bis' | 'bemerkung' | 'eingetragen'
        const td = document.createElement('td');
        td.classList.add(`col-${colKey}`);       // <-- feste Spaltenklasse
        td.dataset.colKey = colKey;

        if (colKey === 'fahrzeug') {
          td.classList.add('text-center');       // ID mittig, wenn gewünscht
          const id = FAHRZEUGE[zeile] || '';
          const label = document.createElement('span');
          label.textContent = id;
          label.style.fontWeight = 'bold';
          if (id && FAHRZEUGNAMEN && FAHRZEUGNAMEN[id]) label.title = FAHRZEUGNAMEN[id];
          td.appendChild(label);

          if (!readOnlyMulti) {
            const hidden = document.createElement('input');
            hidden.type = 'hidden';
            hidden.value = id;
            td.appendChild(hidden);
          }
        } else {
          const idForRow = FAHRZEUGE[zeile] || '';
          const savedFields = (savedMap.get(idForRow) || ['', '', '', '', '', '']);
          const fieldIndex = i - 1;
          const vorbefuellt = (fieldIndex >= 0 ? savedFields[fieldIndex] : '') || '';

          if (readOnlyMulti) {
            const div = document.createElement('div');
            div.className = 'ro';
            div.textContent = vorbefuellt;
            td.appendChild(div);
          } else {
            // Fallback, falls makeInput entfernt wurde
            let inputEl;
            if (typeof makeInput === 'function') {
              inputEl = makeInput(colKey, vorbefuellt);
            } else {
              inputEl = document.createElement('input');
              // wenn du keine nativen Time-Picker willst, lass überall 'text'
              inputEl.type = (colKey === 'von' || colKey === 'bis') ? 'text' : 'text';
              inputEl.maxLength = INPUT_MAX[colKey] ?? 10;
              inputEl.value = vorbefuellt;
            }
            inputEl.classList.add('cell-input', `cell-${colKey}`);
            inputEl.dataset.orig = vorbefuellt;
            inputEl.dataset.row = zeilenIndexGlobal;
            inputEl.dataset.col = i;

            inputEl.addEventListener('input', () => {
              if (bearbeitungGesperrt) return;
              setUnsaved(true);
              updateChangedFlag(inputEl);
            });

            td.appendChild(inputEl);
          }
        }

        tr.appendChild(td);
      }

      tbody.appendChild(tr);
      zeilenIndexGlobal++;
    }
  }
}




function anzeigenAlleMonate(jahr) {
  const table = document.getElementById('monatsTabelle');
  const multi = document.getElementById('multiMonate');
  if (!table || !multi) return;

  table.style.display = 'none';
  multi.style.display = 'block';
  multi.innerHTML = '';

  const monate = ["Januar", "Februar", "März", "April", "Mai", "Juni", "Juli",
    "August", "September", "Oktober", "November", "Dezember"];

  if (!Number.isFinite(jahr)) {
    const uiVal = parseInt(document.getElementById('jahr')?.value, 10);
    jahr = Number.isFinite(uiVal) ? uiVal : (new Date()).getFullYear();
  }

  for (let m = 0; m < 12; m++) {
    const wrapper = document.createElement('section');

    const h2 = document.createElement('h2');
    h2.className = 'section-month-title';
    h2.textContent = `${monate[m]} ${jahr}`;
    wrapper.appendChild(h2);

    const table = document.createElement('table');
    table.className = 'miniMonat';
    table.innerHTML = `
      <thead>
        <tr>
          <th>Tag</th>
          <th>Fahrzeug</th>
          <th>Dienstposten</th>
          <th class="ort-cell">Ort</th>
          <th>Von (Zeit)</th>
          <th>Bis (Zeit)</th>
          <th>Bemerkung</th>
          <th>eingetragen</th>
        </tr>
      </thead>
      <tbody></tbody>
    `;

    const tbody = table.querySelector('tbody');
    tabelleErstellen(jahr, m, tbody);

    wrapper.appendChild(table);
    multi.appendChild(wrapper);
  }
}

// #endregion 



//#region Import Excel

async function importAllMonthYearSheets(file) {
  const arrayBuffer = await file.arrayBuffer();
  const wb = XLSX.read(arrayBuffer, { type: 'array', cellDates: true, cellNF: true, cellText: false });

  // Fahrzeugliste (optional) ziehen
  applyVehicleListFromWorkbook(wb);
  ZEILEN_PRO_TAG = Math.min(Math.max(FAHRZEUGE.length, MIN_ZEILEN_PRO_TAG), MAX_ZEILEN_PRO_TAG);

  const imported = [];
  const skippedNoHeader = [];
  const skippedNoMY = [];

  for (const sheetName of wb.SheetNames) {
    const ws = wb.Sheets[sheetName];
    if (!ws) continue;

    const my = parseMonthYearFromSheetName(sheetName);
    if (!my) { skippedNoMY.push(sheetName); continue; }

    const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false }) || [];
    if (!hasExpectedHeader(aoa)) { skippedNoHeader.push(sheetName); continue; }

    parseSheetToStorage(ws, my.year, my.monthIndex);
    imported.push({ name: sheetName, ...my });
  }

  if (imported.length) {
    const first = imported[0];
    document.getElementById('jahr').value = first.year;
    document.getElementById('monat').value = first.monthIndex;
  }

  // UI refresh
  aktualisieren();
  updateExportButtonLabel();
  renderFahrzeugLegende();
  reflowRightRail?.();

  // Zusätzlich: aktuelle Auswahl explizit neu rendern
  const curYear = parseInt(document.getElementById('jahr').value, 10) || (new Date()).getFullYear();
  const curMonth = parseInt(document.getElementById('monat').value, 10) || 0;
  tabelleErstellen(curYear, curMonth);
  renderFahrzeugLegende();
  reflowRightRail?.();
}

function applyVehicleListFromWorkbook(wb) {
  for (const name of wb.SheetNames || []) {
    const ws = wb.Sheets[name];
    if (!ws) continue;
    const listOwn = extractVehicleList_OWN(XLSX.utils.sheet_to_json(ws, { header: 1, raw: false }) || []);
    if (listOwn.length) {
      FAHRZEUGE = listOwn.map(x => x.id);
      FAHRZEUGNAMEN = {};
      for (const { id, name } of listOwn) FAHRZEUGNAMEN[id] = name;
      ZEILEN_PRO_TAG = Math.min(Math.max(FAHRZEUGE.length, MIN_ZEILEN_PRO_TAG), MAX_ZEILEN_PRO_TAG);
      return true;
    }
  }

  for (const name of wb.SheetNames || []) {
    const ws = wb.Sheets[name];
    if (!ws) continue;
    const list = extractVehicleList_FALLBACK(XLSX.utils.sheet_to_json(ws, { header: 1, raw: false }) || []);
    if (list.length) {
      FAHRZEUGE = list.map(x => x.id);
      FAHRZEUGNAMEN = {};
      for (const { id, name } of list) FAHRZEUGNAMEN[id] = name;
      ZEILEN_PRO_TAG = Math.min(Math.max(FAHRZEUGE.length, MIN_ZEILEN_PRO_TAG), MAX_ZEILEN_PRO_TAG);
      return true;
    }
  }

  const ids = collectVehiclesFromWorkbook(wb);
  if (ids.length) {
    FAHRZEUGE = ids;
    FAHRZEUGNAMEN = FAHRZEUGNAMEN || {};
    ZEILEN_PRO_TAG = Math.min(Math.max(FAHRZEUGE.length, MIN_ZEILEN_PRO_TAG), MAX_ZEILEN_PRO_TAG);
    return true;
  }

  return false;
}

function extractVehicleList_OWN(aoa) {
  const rows = aoa.length;
  const cols = Math.max(...aoa.map(r => r.length)) || 0;

  const norm = v => String(v ?? '').trim().toLowerCase();

  let minCol = 0;
  outer:
  for (let r = 0; r < Math.min(rows, 20); r++) {
    for (let c = 0; c < cols; c++) {
      const v = norm(aoa[r]?.[c]);
      if (v === 'wochenende' || v === 'feiertag') {
        minCol = Math.max(minCol, c);
        break outer;
      }
    }
  }

  for (let r = 0; r < rows; r++) {
    for (let c = Math.max(0, minCol); c < cols - 1; c++) {
      const a = norm(aoa[r]?.[c]);
      const b = norm(aoa[r]?.[c + 1]);
      if ((a === 'id') && (b === 'fahrzeugname' || b === 'fahrzeug')) {
        const out = [];
        for (let rr = r + 1; rr < rows; rr++) {
          const rawId = String(aoa[rr]?.[c] ?? '').trim();
          const name = String(aoa[rr]?.[c + 1] ?? '').trim();
          if (!rawId && !name) break;

          if (/^\d{2,4}$/.test(rawId) && name) {
            out.push({ id: rawId.replace(/^0+/, ''), name });
          } else if (!rawId && name) {
            break;
          }
        }
        return out;
      }
    }
  }
  return [];
}

function extractVehicleList_FALLBACK(aoa) {
  const rows = aoa.length;
  const cols = Math.max(...aoa.map(r => r.length)) || 0;

  const isId = (v) => /^\s*\d{2,4}\s*$/.test(String(v ?? ''));
  const isText = (v) => String(v ?? '').trim().length > 0;

  let best = [];
  for (let c = 0; c < cols - 1; c++) {
    let block = [];
    let started = false;

    for (let r = 0; r < Math.min(rows, 1000); r++) {
      const left = aoa[r]?.[c];
      const right = aoa[r]?.[c + 1];

      if (isId(left) && isText(right)) {
        block.push({ id: String(left).trim().replace(/^0+/, ''), name: String(right).trim() });
        started = true;
        if (block.length >= 150) break;
      } else {
        if (started) break;
      }
    }

    if (block.length >= 2 && block.length > best.length) {
      best = block;
    }
  }
  return best;
}


function collectVehiclesFromWorkbook(wb) {
  const ids = new Set();
  const looksLikeId = v => /^\s*\d{2,4}\s*$/.test(String(v ?? ''));
  const norm = s => String(s ?? '').trim().toLowerCase();

  for (const sheetName of wb.SheetNames || []) {
    const ws = wb.Sheets[sheetName];
    if (!ws) continue;
    const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false }) || [];
    if (!aoa.length) continue;

    const headerRow = aoa.find(r => Array.isArray(r) && r.some(c => norm(c) === 'fahrzeug'));
    if (!headerRow) continue;
    const hdrIdx = aoa.indexOf(headerRow);
    const colFzg = headerRow.findIndex(c => norm(c) === 'fahrzeug');
    if (colFzg < 0) continue;

    for (let r = hdrIdx + 1; r < aoa.length; r++) {
      const v = aoa[r]?.[colFzg];
      if (looksLikeId(v)) ids.add(String(v).trim().replace(/^0+/, ''));
    }
  }
  return Array.from(ids).sort((a, b) => a.localeCompare(b, 'de', { numeric: true }));
}

function parseSheetToStorage(ws, targetYear, monthIndex) {
  const norm = (s) => String(s ?? '').trim().toLowerCase();
  const safeCell = (row, idx) =>
    (Number.isInteger(idx) && idx >= 0 && idx < row.length && row[idx] != null)
      ? String(row[idx]).trim() : '';

  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false });
  if (!aoa || !aoa.length) throw new Error('Keine Daten im Blatt.');

  let headerRowIdx = -1;
  for (let i = 0; i < Math.min(10, aoa.length); i++) {
    const row = (aoa[i] || []).map(x => norm(x));
    if (
      row.includes('fahrzeug') &&
      row.includes('dienstposten') &&
      (row.includes('bemerkung') || row.includes('bemerkungen'))
    ) { headerRowIdx = i; break; }
  }
  if (headerRowIdx === -1) headerRowIdx = 0;

  const own = tryParseOwnExport(aoa, headerRowIdx, targetYear, monthIndex);
  if (own) return own;

  const header = (aoa[headerRowIdx] || []).map(x => norm(x));
  const dataRows = aoa.slice(headerRowIdx + 1);

  const idx = {
    fahrzeug: header.findIndex(h => h === 'fahrzeug'),
    dienstposten: header.findIndex(h => h === 'dienstposten'),
    ort: header.findIndex(h => h === 'ort'),
    von: header.findIndex(h => typeof h === 'string' && h.includes('von')),
    bis: header.findIndex(h => typeof h === 'string' && h.includes('bis')),
    bemerkung: header.findIndex(h => typeof h === 'string' && h.startsWith('bemerk')),
    eingetragen: header.findIndex(h => typeof h === 'string' && (h.includes('eingetragen') || h.includes('geändert') || h.includes('geaendert'))),
    tag: header.findIndex(h => h === 'tag' || h === 'wochentag'),
    datum: header.findIndex(h => h === 'datum' || (typeof h === 'string' && h.includes('tag'))),
  };
  const use = {
    fahrzeug: idx.fahrzeug >= 0 ? idx.fahrzeug : 1,
    dienstposten: idx.dienstposten >= 0 ? idx.dienstposten : 2,
    ort: idx.ort >= 0 ? idx.ort : 3,
    von: idx.von >= 0 ? idx.von : 4,
    bis: idx.bis >= 0 ? idx.bis : 5,
    bemerkung: idx.bemerkung >= 0 ? idx.bemerkung : 6,
    eingetragen: idx.eingetragen >= 0 ? idx.eingetragen : 7,
    tag: idx.tag,
    datum: idx.datum,
  };
  const maxCol = Math.max(0, header.length - 1);
  if (use.von === use.ort) use.von = Math.min(use.ort + 1, maxCol);
  if (use.bis === use.ort || use.bis === use.von) {
    use.bis = Math.min(((Number.isInteger(use.von) ? use.von : use.ort) + 1), maxCol);
  }

  const tageImMonat = new Date(targetYear, monthIndex + 1, 0).getDate();

  function extractDayNumber(txt) {
    if (!txt) return null;
    const s = String(txt).trim();
    const m = s.match(/\b(0?[1-9]|[12]\d|3[01])(?=\D|$)/);
    return m ? parseInt(m[1], 10) : null;
  }

  const dayRows = Array.from({ length: tageImMonat + 1 }, () => []);
  let currentDay = null;

  const emptyRow = () => ['', '', '', '', '', '', ''];

  for (const r of dataRows) {
    const wt = safeCell(r, use.tag);
    const dt = safeCell(r, use.datum);
    const dayNr = extractDayNumber(wt || dt);
    if (dayNr && dayNr >= 1 && dayNr <= tageImMonat) currentDay = dayNr;
    if (!currentDay) continue;

    const vals = [
      safeCell(r, use.fahrzeug),
      safeCell(r, use.dienstposten),
      safeCell(r, use.ort),
      safeCell(r, use.von),
      safeCell(r, use.bis),
      safeCell(r, use.bemerkung),
      safeCell(r, use.eingetragen),
    ];

    if (vals.every(v => v === '')) continue;

    if (vals[0] !== '') {
      const asNum = Number(vals[0].replace(',', '.'));
      if (Number.isFinite(asNum)) vals[0] = String(asNum);
    }
    dayRows[currentDay].push(vals);
  }

  const anyData = dayRows.some((arr, i) => i && arr.length);
  if (!anyData) {
    const rowsPerDayEmpty = Math.min(
      Math.max((FAHRZEUGE.length || ZEILEN_PRO_TAG || 1), MIN_ZEILEN_PRO_TAG),
      MAX_ZEILEN_PRO_TAG
    );
    localStorage.setItem(speicherKey(targetYear, monthIndex), JSON.stringify([]));
    speichereMeta(targetYear, monthIndex, { rowsPerDay: rowsPerDayEmpty });
    return { rowsPerDay: rowsPerDayEmpty };
  }

  let rowsPerDay = Math.max(...dayRows.map((b, i) => i ? b.length : 0));
  rowsPerDay = Math.min(Math.max(rowsPerDay, MIN_ZEILEN_PRO_TAG), MAX_ZEILEN_PRO_TAG);

  const normalized2D = [];
  for (let d = 1; d <= tageImMonat; d++) {
    const block = (dayRows[d] || []).slice(0, rowsPerDay);
    while (block.length < rowsPerDay) block.push(emptyRow());
    normalized2D.push(...block);
  }

  localStorage.setItem(speicherKey(targetYear, monthIndex), JSON.stringify(normalized2D));
  speichereMeta(targetYear, monthIndex, { rowsPerDay });
  return { rowsPerDay };
}

function tryParseOwnExport(aoa, headerRowIdx, targetYear, monthIndex) {
  const hdr = (aoa[headerRowIdx] || []).map(s => String(s ?? '').trim().toLowerCase());
  const want = ["tag", "fahrzeug", "dienstposten", "ort", "von", "bis", "bemerkung", "eingetragen"];
  const looksOk = want.every((h, i) => (hdr[i] || '').startsWith(h));
  if (!looksOk) return null;

  const MIN = 1, MAX = 30;
  let rowsPerDay = Math.min(Math.max(FAHRZEUGE.length || ZEILEN_PRO_TAG || 1, MIN), MAX);

  const bodyRows = aoa.slice(headerRowIdx + 1);

  // Wir lesen Nutzspalten B..H (Index 1..7)
  const rows = bodyRows.map(r => ([
    String(r[1] ?? '').trim(),
    String(r[2] ?? '').trim(),
    String(r[3] ?? '').trim(),
    String(r[4] ?? '').trim(),
    String(r[5] ?? '').trim(),
    String(r[6] ?? '').trim(),
    String(r[7] ?? '').trim(),
  ]));

  const daysInMonth = new Date(targetYear, monthIndex + 1, 0).getDate();
  const needed = rowsPerDay * daysInMonth;

  const out = rows.slice(0, needed);
  while (out.length < needed) out.push(['', '', '', '', '', '', '']);

  localStorage.setItem(speicherKey(targetYear, monthIndex), JSON.stringify(out));
  speichereMeta(targetYear, monthIndex, { rowsPerDay });
  return { rowsPerDay };
}

function hasExpectedHeader(aoa) {
  const norm = s => String(s ?? "").trim().toLowerCase();
  for (let i = 0; i < Math.min(10, aoa.length); i++) {
    const row = (aoa[i] || []).map(norm);
    const hasFahrzeug = row.some(c => c.includes("fahrzeug"));
    const hasDienstposten = row.some(c => c.includes("dienstposten"));
    const hasBemerkung = row.some(c => c.includes("bemerkung"));
    if (hasFahrzeug && hasDienstposten && hasBemerkung) return true;
  }
  return false;
}


function parseMonthYearFromSheetName(name) {
  if (!name) return null;
  let s = String(name).toLowerCase();
  try { s = s.normalize('NFD').replace(/\p{Diacritic}/gu, ''); } catch { }
  const months = [
    ['jan', 'januar'],
    ['feb', 'februar'],
    ['maer', 'maerz', 'marz', 'mar', 'mär'],
    ['apr', 'april'],
    ['mai', 'may'],
    ['jun', 'juni'],
    ['jul', 'juli'],
    ['aug', 'august'],
    ['sep', 'sept', 'september'],
    ['okt', 'oct', 'oktober', 'october'],
    ['nov', 'november'],
    ['dez', 'dec', 'dezember', 'december'],
  ];

  const ym = s.match(/\b(19|20)\d{2}\b/);
  if (!ym) return null;
  const year = parseInt(ym[0], 10);

  let monthIndex = null;
  for (let i = 0; i < months.length; i++) {
    if (months[i].some(tok => s.includes(tok))) { monthIndex = i; break; }
  }
  if (monthIndex == null) {
    const mNum = s.match(/\b(0?[1-9]|1[0-2])\b/);
    if (mNum) monthIndex = parseInt(mNum[1], 10) - 1;
  }
  if (monthIndex == null) return null;

  return { year, monthIndex };
}




//#endregion



// #region Excel-Export

 //Hilfsfunktionen nur für den Excel-Export
function tabelleAlsArray() {
  const tbody = document.querySelector('#monatsTabelle tbody');
  const rows = Array.from(tbody.querySelectorAll('tr'));
  const header = ['Datum', 'Fahrzeug', 'Dienstposten', 'Ort', 'Von', 'Bis', 'Bemerkung', 'Eingetragen'];

  const out = [header];
  let lastDatumText = '';

  rows.forEach(tr => {
    const hasDateCell = tr.cells.length === 8;

    if (hasDateCell) {
      const raw = tr.cells[0].innerText.trim();
      lastDatumText = raw.split('\n')[0].trim();
    }

    const fahrzeugTd = hasDateCell ? tr.cells[1] : tr.cells[0];
    let fahrzeug = fahrzeugTd.querySelector('span')?.textContent?.trim() ?? '';
    if (!fahrzeug) {
      fahrzeug = fahrzeugTd.querySelector('input[type="hidden"]')?.value?.trim() ?? '';
    }

    const startCellIndex = hasDateCell ? 2 : 1;
    const rest = [];
    for (let c = startCellIndex; c < tr.cells.length; c++) {
      const cell = tr.cells[c];
      const v = cell.querySelector('input[type="text"],input[type="time"]')?.value
        ?? cell.querySelector('.ro')?.textContent
        ?? '';
      rest.push(v);
    }
    if (!rest.length) rest.push('', '', '', '', '', '');
    out.push([lastDatumText, fahrzeug, ...rest]);
  });

  return out;
}

function tabelleAlsArrayFromTbody(tb) {
  let out = [['Datum', 'Fahrzeug', 'Dienstposten', 'Ort', 'Von', 'Bis', 'Bemerkung', 'Eingetragen']];
  let lastDate = '';
  [...tb.querySelectorAll('tr')].forEach(tr => {
    const hasDate = tr.cells.length === 8;
    if (hasDate) {
      const raw = tr.cells[0].innerText.trim();
      lastDate = raw.split('\n')[0].trim();
    }
    const fTd = hasDate ? tr.cells[1] : tr.cells[0];
    const fahrz = fTd.querySelector('span')?.textContent?.trim() || fTd.querySelector('input[type="hidden"]')?.value?.trim() || '';
    const startIdx = hasDate ? 2 : 1;
    const rest = [];
    for (let c = startIdx; c < tr.cells.length; c++) {
      const v = tr.cells[c].querySelector('input[type="text"],input[type="time"]')?.value
        ?? tr.cells[c].querySelector('.ro')?.textContent
        ?? '';
      rest.push(v);
    }
    out.push([lastDate, fahrz, ...rest]);
  });
  return out;
}

function rowsPerDayArrayFromDOM(tb) {
  return [...tb.querySelectorAll('tr.first-of-day td[rowspan]')].map(td =>
    parseInt(td.getAttribute('rowspan') || '1', 10)
  );
}

function renderMonthOffscreen(y, m) {
  const wrapper = document.createElement('div');
  wrapper.style.cssText = 'position:fixed;left:-99999px;top:-99999px;visibility:hidden';
  const table = document.createElement('table');
  table.innerHTML = '<thead><tr><th>Tag</th><th>Fahrzeug</th><th>Dienstposten</th><th>Ort</th><th>Von</th><th>Bis</th><th>Bemerkung</th><th>eingetragen</th></tr></thead><tbody></tbody>';
  wrapper.appendChild(table);
  document.body.appendChild(wrapper);
  const tb = table.querySelector('tbody');
  tabelleErstellen(y, m, tb);
  return { tbody: tb, cleanup: () => wrapper.remove() };
}

 //Einzelmonat exportieren
async function exportMitExcelJS() {
  if (typeof ExcelJS === 'undefined') {
    alert('ExcelJS ist nicht geladen.');
    return;
  }

  const wb = new ExcelJS.Workbook();

  const jahr = parseInt(document.getElementById('jahr').value, 10);
  const monat = parseInt(document.getElementById('monat').value, 10);
  const monate = ["Januar", "Februar", "März", "April", "Mai", "Juni", "Juli",
    "August", "September", "Oktober", "November", "Dezember"];

  const blueHeader = "2F75B5";
  const zebraARGB = "F2F2F2";
  const weekendARGB = "FFF6DFD1";
  const holidayARGB = "FFDFEEDD";

  const ws = wb.addWorksheet(`${monate[monat]} ${jahr}`, {
    properties: { defaultRowHeight: 18 },
    pageSetup: { paperSize: 9, orientation: "landscape", fitToPage: true }
  });

  // Titel
  const titel = `Tabelle für ${monate[monat]} ${jahr}`;
  ws.spliceRows(1, 0, [titel]);
  ws.mergeCells(1, 1, 1, 9);
  const tcell = ws.getCell('A1');
  tcell.font = { bold: true, size: 16 };
  tcell.alignment = { horizontal: 'center', vertical: 'middle' };
  ws.getRow(1).height = 24;

  // Kopf
  ws.addRow(["Tag", "Fahrzeug", "Dienstposten", "Ort", "Von", "Bis", "Bemerkung", "eingetragen"]);
  ws.getRow(2).eachCell(c => {
    c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: blueHeader } };
    c.font = { bold: true, color: { argb: "FFFFFFFF" } };
    c.alignment = { horizontal: "center", vertical: "middle" };
    c.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };
  });

  // Spaltenbreiten + Legende-Spalten
  ws.columns = [
    { width: 25 }, // A Datum
    { width: 12 }, // B Fahrzeug
    { width: 18 }, // C Dienstposten
    { width: 22 }, // D Ort
    { width: 12 }, // E Von
    { width: 12 }, // F Bis
    { width: 40 }, // G Bemerkung
    { width: 16 }, // H eingetragen
    { width: 2 },  // I Spacer
    { width: 14 }, // J Legende: ID
    { width: 28 }, // K Legende: Fahrzeugname
  ];

  const aoa = tabelleAlsArray(); // inkl. Datumstext
  const raw = aoa.slice(1);

  let rowsPerDay = FAHRZEUGE.length || 1;
  const firstDateCell = document.querySelector('#monatsTabelle tbody tr td[rowspan]');
  if (firstDateCell) {
    const rs = parseInt(firstDateCell.getAttribute('rowspan'), 10);
    if (!Number.isNaN(rs) && rs > 0) rowsPerDay = rs;
  }

  const dataOnly = raw.map(r => {
    const asNum = Number(String(r[1]).trim().replace(',', '.'));
    const fahrzeug = Number.isFinite(asNum) ? asNum : r[1];
    return [r[0], fahrzeug, r[2], r[3], r[4], r[5], r[6], r[7]];
  });
  ws.addRows(dataOnly);

  // Rahmen + Zebra
  const startDataRow = 3;
  const endDataRow = ws.lastRow.number;
  for (let r = startDataRow; r <= endDataRow; r++) {
    const row = ws.getRow(r);
    row.eachCell({ includeEmpty: true }, cell => {
      cell.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };
    });
    if ((r - startDataRow) % 2 === 1) {
      row.eachCell({ includeEmpty: true }, cell => {
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: zebraARGB } };
      });
    }
  }

  // Tag-Blöcke mergen + WE/Feiertags-Färbung
  const tageImMonat = new Date(jahr, monat + 1, 0).getDate();
  for (let blockStart = 0, day = 1; day <= tageImMonat; day++, blockStart += rowsPerDay) {
    const excelStart = startDataRow + blockStart;
    const excelEnd = Math.min(excelStart + rowsPerDay - 1, endDataRow);
    if (excelStart > endDataRow) break;

    ws.mergeCells(excelStart, 1, excelEnd, 1);

    const master = ws.getRow(excelStart).getCell(1);
    if (!master.value) {
      const dObj = new Date(jahr, monat, day);
      const dateText = dObj.toLocaleDateString('de-DE', { weekday: 'long', day: 'numeric', month: 'long' });
      master.value = { richText: [{ text: dateText }] };
    }
    master.alignment = { vertical: "middle", horizontal: "left", wrapText: true };

    const dObj = new Date(jahr, monat, day);
    const isWeekend = dObj.getDay() === 0 || dObj.getDay() === 6;
    const holiday = feiertagsNameFür(dObj);
    if (isWeekend || holiday) {
      const color = holiday ? holidayARGB : weekendARGB;
      for (let rr = excelStart; rr <= excelEnd; rr++) {
        for (let c = 1; c <= 8; c++) {
          ws.getRow(rr).getCell(c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: color } };
        }
      }
    }
    for (let rr = excelStart; rr <= excelEnd; rr++) {
      ws.getRow(rr).getCell(1).border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };
    }
  }

  // Rechte Legende: nur Fahrzeugliste (ID + Name) + Farberklärung WE/Feiertag
  (function addRightLegend() {
    const idCol = 10, nameCol = 11, headRow = 2;

    ws.getCell(headRow, idCol).value = "Wochenende";
    ws.getCell(headRow + 1, idCol).value = "Feiertag";
    ws.getCell(headRow, nameCol).fill = { type: "pattern", pattern: "solid", fgColor: { argb: weekendARGB } };
    ws.getCell(headRow + 1, nameCol).fill = { type: "pattern", pattern: "solid", fgColor: { argb: holidayARGB } };
    ws.getCell(headRow, nameCol).border = ws.getCell(headRow + 1, nameCol).border =
      { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };

    const startRow = headRow + 3;
    ws.getCell(startRow, idCol).value = "ID";
    ws.getCell(startRow, nameCol).value = "Fahrzeugname";
    [ws.getCell(startRow, idCol), ws.getCell(startRow, nameCol)].forEach(c => {
      c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: blueHeader } };
      c.font = { bold: true, color: { argb: "FFFFFFFF" } };
      c.alignment = { horizontal: "center", vertical: "middle" };
      c.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };
    });

    let r = startRow + 1;
    const ordered = [...FAHRZEUGE].sort((a, b) => a.localeCompare(b, 'de', { numeric: true }));
    for (const id of ordered) {
      ws.getCell(r, idCol).value = id;
      ws.getCell(r, idCol).border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };
      const cell = ws.getCell(r, nameCol);
      cell.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };
      cell.value = FAHRZEUGNAMEN[id] || '';
      r++;
    }
  })();

  const buf = await wb.xlsx.writeBuffer();
  const blob = new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = `Monatstabelle-${monate[monat]}-${jahr}.xlsx`;
  a.click();
  URL.revokeObjectURL(a.href);
}

// Ganzes Jahr (12 Blätter) exportieren – DOM-Mirror

async function exportYearWithExcelJS_DOMMirror() {
  if (typeof ExcelJS === 'undefined') {
    alert('ExcelJS ist nicht geladen.');
    return;
  }

  const wb = new ExcelJS.Workbook();
  const jahr = parseInt(document.getElementById('jahr').value, 10) || (new Date()).getFullYear();
  const monate = ["Januar", "Februar", "März", "April", "Mai", "Juni", "Juli",
    "August", "September", "Oktober", "November", "Dezember"];

  const blueHeader = "2F75B5", zebraARGB = "F2F2F2", weekendARGB = "FFF6DFD1", holidayARGB = "FFDFEEDD";

  const styleHeaderRow = (ws, rowIdx) =>
    ws.getRow(rowIdx).eachCell(c => {
      c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: blueHeader } };
      c.font = { bold: true, color: { argb: "FFFFFFFF" } };
      c.alignment = { horizontal: "center", vertical: "middle" };
      c.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };
    });

  for (let monat = 0; monat < 12; monat++) {
    const { tbody, cleanup } = renderMonthOffscreen(jahr, monat);
    const aoa = tabelleAlsArrayFromTbody(tbody);
    const data = aoa.slice(1);
    const rpdAr = rowsPerDayArrayFromDOM(tbody);
    const tageImMonat = new Date(jahr, monat + 1, 0).getDate();

    const offsets = rpdAr.reduce((acc, n, i) => {
      acc.push((acc[i - 1] ?? 0) + (rpdAr[i - 1] ?? 0));
      return acc;
    }, []);

    const ws = wb.addWorksheet(`${monate[monat]} ${jahr}`, {
      properties: { defaultRowHeight: 18 },
      pageSetup: { paperSize: 9, orientation: "landscape", fitToPage: true }
    });

    const titel = `Tabelle für ${monate[monat]} ${jahr}`;
    ws.spliceRows(1, 0, [titel]); ws.mergeCells(1, 1, 1, 9);
    const tcell = ws.getCell('A1');
    tcell.font = { bold: true, size: 16 };
    tcell.alignment = { horizontal: 'center', vertical: 'middle' };
    ws.getRow(1).height = 24;

    ws.addRow(["Tag", "Fahrzeug", "Dienstposten", "Ort", "Von", "Bis", "Bemerkung", "eingetragen"]);
    styleHeaderRow(ws, 2);

    ws.columns = [
      { width: 25 }, { width: 12 }, { width: 18 }, { width: 22 },
      { width: 12 }, { width: 12 }, { width: 40 }, { width: 16 },
      { width: 2 }, { width: 14 }, { width: 28 }
    ];

    ws.addRows(data);

    const startDataRow = 3;
    const endDataRow = ws.lastRow.number;
    for (let r = startDataRow; r <= endDataRow; r++) {
      ws.getRow(r).eachCell({ includeEmpty: true }, cell => {
        cell.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };
      });
      if ((r - startDataRow) % 2 === 1) {
        ws.getRow(r).eachCell({ includeEmpty: true }, cell => {
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: zebraARGB } };
        });
      }
    }

    for (let day = 1; day <= tageImMonat; day++) {
      const rpd = rpdAr[day - 1] ?? 1;
      const excelStart = startDataRow + (offsets[day - 1] || 0);
      const excelEnd = Math.min(excelStart + rpd - 1, endDataRow);
      if (excelStart > endDataRow) break;

      ws.mergeCells(excelStart, 1, excelEnd, 1);

      const master = ws.getRow(excelStart).getCell(1);
      if (!master.value) {
        const dObj = new Date(jahr, monat, day);
        const dateText = dObj.toLocaleDateString('de-DE', { weekday: 'long', day: 'numeric', month: 'long' });
        master.value = { richText: [{ text: dateText }] };
      }
      master.alignment = { vertical: "middle", horizontal: "left", wrapText: true };

      const dObj = new Date(jahr, monat, day);
      const isWeekend = dObj.getDay() === 0 || dObj.getDay() === 6;
      const holiday = feiertagsNameFür(dObj);
      if (isWeekend || holiday) {
        const color = holiday ? holidayARGB : weekendARGB;
        for (let rr = excelStart; rr <= excelEnd; rr++) {
          for (let c = 1; c <= 8; c++) {
            ws.getRow(rr).getCell(c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: color } };
          }
        }
      }
      for (let rr = excelStart; rr <= excelEnd; rr++) {
        ws.getRow(rr).getCell(1).border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };
      }
    }

    // rechte Legende je Monat
    (function addLegends() {
      const idCol = 10, nameCol = 11, headRow = 2;

      ws.getCell(headRow, idCol).value = "Wochenende";
      ws.getCell(headRow + 1, idCol).value = "Feiertag";
      ws.getCell(headRow, nameCol).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: weekendARGB } };
      ws.getCell(headRow + 1, nameCol).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: holidayARGB } };
      ws.getCell(headRow, nameCol).border = ws.getCell(headRow + 1, nameCol).border =
        { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };

      const startRow = headRow + 3;
      ws.getCell(startRow, idCol).value = "ID";
      ws.getCell(startRow, nameCol).value = "Fahrzeugname";
      [ws.getCell(startRow, idCol), ws.getCell(startRow, nameCol)].forEach(c => {
        c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: blueHeader } };
        c.font = { bold: true, color: { argb: "FFFFFFFF" } };
        c.alignment = { horizontal: "center", vertical: "middle" };
        c.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };
      });

      let r = startRow + 1;
      const ordered = [...FAHRZEUGE].sort((a, b) => a.localeCompare(b, 'de', { numeric: true }));
      for (const id of ordered) {
        ws.getCell(r, idCol).value = id;
        ws.getCell(r, idCol).border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };
        const cell = ws.getCell(r, nameCol);
        cell.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };
        cell.value = FAHRZEUGNAMEN[id] || '';
        r++;
      }
    })();

    cleanup();
  }

  const buf = await wb.xlsx.writeBuffer();
  const blob = new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = `Monatstabelle-AlleMonate-${jahr}.xlsx`;
  a.click();
  URL.revokeObjectURL(a.href);
}


// #endregion



// #region Speichern (Tabelle → localStorage)

function speichereDaten(jahr, monat) {
  const tbody = document.querySelector('#monatsTabelle tbody');
  const rows = Array.from(tbody.querySelectorAll('tr'));

  const tageImMonat = new Date(jahr, monat + 1, 0).getDate();
  const meta = ladeMeta(jahr, monat);

  let STORAGE_RPD = parseInt(meta?.rowsPerDay, 10);
  if (!Number.isFinite(STORAGE_RPD) || STORAGE_RPD < MIN_ZEILEN_PRO_TAG || STORAGE_RPD > MAX_ZEILEN_PRO_TAG) {
    const existing = ladeDaten(jahr, monat) || [];
    if (existing.length) {
      STORAGE_RPD = Math.max(
        MIN_ZEILEN_PRO_TAG,
        Math.min(MAX_ZEILEN_PRO_TAG, Math.round(existing.length / Math.max(1, tageImMonat)))
      );
    } else {
      STORAGE_RPD = Math.min(Math.max((FAHRZEUGE.length || ZEILEN_PRO_TAG || 1), MIN_ZEILEN_PRO_TAG), MAX_ZEILEN_PRO_TAG);
    }
  }

  const daten = [];
  let ptr = 0;

  for (let day = 1; day <= tageImMonat; day++) {
    const rpdToday = LAST_RPD_PER_DAY[day] || 1;
    const maxRows = Math.min(rpdToday, STORAGE_RPD);

    for (let r = 0; r < maxRows; r++) {
      const tr = rows[ptr++];
      const zeile = [];
      const tds = Array.from(tr.querySelectorAll('td'));
      const fahrzeugTd = tds.length === 8 ? tds[1] : tds[0];
      const id = fahrzeugTd.querySelector('span')?.textContent?.trim()
        || fahrzeugTd.querySelector('input[type="hidden"]')?.value?.trim() || '';
      zeile.push(id);

      tr.querySelectorAll('input[type="text"],input[type="time"]').forEach(inp => zeile.push(inp.value ?? ''));
      while (zeile.length < 7) zeile.push('');
      daten.push(zeile);
    }

    for (let r = maxRows; r < STORAGE_RPD; r++) {
      daten.push(['', '', '', '', '', '', '']);
    }

    for (let r = STORAGE_RPD; r < rpdToday; r++) ptr++;
  }

  localStorage.setItem(speicherKey(jahr, monat), JSON.stringify(daten));
}

// #endregion



// #region Lock/Unlock 

function setDisabledState(gesperrt) {
  document.querySelectorAll('#monatsTabelle input').forEach(inp => inp.disabled = gesperrt);
  const badge = document.getElementById('lockBadge');
  const lockBtn = document.getElementById('lockBtn');
  const unlockBtn = document.getElementById('unlockBtn');
  const statusAnzeige = document.getElementById('statusAnzeige');

  if (!badge || !lockBtn || !unlockBtn || !statusAnzeige) return;

  if (gesperrt) {
    document.body.classList.add('locked-mode');
    badge.textContent = 'Gesperrt - nur Lesen';
    badge.classList.remove('unlocked');
    badge.classList.add('locked');
    lockBtn.disabled = true;
    unlockBtn.disabled = false;

    statusAnzeige.textContent = 'Gesperrt - nur Lesen (kein Serverschutz)';
    statusAnzeige.style.backgroundColor = 'red';
  } else {
    document.body.classList.remove('locked-mode');
    badge.textContent = 'Bearbeitungsmodus aktiv';
    badge.classList.remove('locked');
    badge.classList.add('unlocked');
    lockBtn.disabled = false;
    unlockBtn.disabled = true;

    statusAnzeige.textContent = 'Aktiv - Bearbeitung möglich';
    statusAnzeige.style.backgroundColor = 'green';
  }
}

function entsperren() {
  const modal = document.getElementById('passwortModal');
  const input = document.getElementById('passwortEingabe');
  const okBtn = document.getElementById('passwortOk');
  const cancelBtn = document.getElementById('passwortCancel');
  const fehlerText = document.getElementById('passwortFehler');

  if (!modal || !input || !okBtn || !cancelBtn || !fehlerText) return;

  modal.style.display = 'flex';
  input.value = '';
  fehlerText.style.display = 'none';
  input.focus();

  function closeModal() {
    modal.style.display = 'none';
    okBtn.removeEventListener('click', confirm);
    cancelBtn.removeEventListener('click', cancel);
    input.removeEventListener('keydown', keyHandler);
    window.removeEventListener('keydown', escHandler);
  }

  function confirm() {
    if (input.value === PASSWORT) {
      bearbeitungGesperrt = false;
      setDisabledState(false);
      closeModal();
    } else {
      fehlerText.style.display = 'block';
      input.value = '';
      input.focus();
    }
  }

  function cancel() { closeModal(); }
  function keyHandler(e) { if (e.key === 'Enter') confirm(); }
  function escHandler(e) { if (e.key === 'Escape') closeModal(); }

  okBtn.addEventListener('click', confirm);
  cancelBtn.addEventListener('click', cancel);
  input.addEventListener('keydown', keyHandler);
  window.addEventListener('keydown', escHandler);
}

function sperren() {
  discardUnsavedChanges();
  bearbeitungGesperrt = true;
  setDisabledState(true);
}

// #endregion



// #region Misc UI Helpers

function afterRenderApply() {
  setDisabledState?.(bearbeitungGesperrt);
  renderFahrzeugLegende?.();
  reflowRightRail?.();
  updateExportButtonLabel?.();

  requestAnimationFrame(() => {
    window.dispatchEvent(new Event('resize'));
  });
}

document.addEventListener('click', (e) => {
  const btn = e.target.closest('[data-cancel-edits]');
  if (!btn) return;
  discardUnsavedChanges();
});

document.addEventListener('DOMContentLoaded', () => {
  const topBtn = document.getElementById('backToTop');
  if (!topBtn) return;

  topBtn.addEventListener('click', () => {
    const noMotion = window.matchMedia('(prefers-reduced-motion: reduce)').matches;
    window.scrollTo({ top: 0, behavior: noMotion ? 'auto' : 'smooth' });
  });

  const THRESHOLD = 200;
  function toggleTopBtn() {
    topBtn.style.display = (window.scrollY > THRESHOLD) ? 'block' : 'none';
  }
  window.addEventListener('scroll', toggleTopBtn, { passive: true });
  toggleTopBtn();
});

// #endregion



// #region Fahrzeug-hinzufügen oder entfernen

// -- Helfer -------------------------------------------------------------------
const _fz_formatDE = (d) =>
  `${String(d.getDate()).padStart(2,'0')}.${String(d.getMonth()+1).padStart(2,'0')}.${d.getFullYear()}`;

// ISO + deutsch (25.02.2025) erlauben
const _fz_parseDateFlex = (s) => {
  if (!s) return null;
  const str = String(s).trim();
  // ISO: 2025-03-09
  let m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(str);
  if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
  // Deutsch: 09.03.2025 (führende Nullen optional)
  m = /^(0?[1-9]|[12]\d|3[01])\.(0?[1-9]|1[0-2])\.(\d{4})$/.exec(str);
  if (m) return new Date(+m[3], +m[2] - 1, +m[1]);
  return null;
};

// Freundlicher Parser: nimmt Text, Array aus Strings oder bereits Objekte
function _fz_normalizeChanges(raw) {
  if (!raw) return [];
  // Bereits Objekt-Form?
  if (Array.isArray(raw) && raw.length && typeof raw[0] === 'object') return raw;

  // Text → Zeilen oder Array aus Strings
  let lines = [];
  if (typeof raw === 'string') {
    lines = raw.split(/\r?\n/);
  } else if (Array.isArray(raw)) {
    lines = raw.map(String);
  } else {
    return [];
  }

  const out = [];
  for (const line of lines) {
    const s = String(line).trim();
    if (!s || s.startsWith('#')) continue;

    // Muster: <Datum> <Aktion> <ID> [Name...]
    // Bsp.: 09.03.2025 hinzufügen 224 NEF 224
    //       2025-05-15 remove 221
    const m = s.match(
      /^(\d{4}-\d{2}-\d{2}|(?:0?[1-9]|[12]\d|3[01])\.(?:0?[1-9]|1[0-2])\.\d{4})\s+(hinzufügen|add|entfernen|remove)\s+(\d{2,4})(?:\s+(.*))?$/i
    );
    if (!m) continue;

    const [, dateStr, actRaw, idRaw, nameRaw] = m;
    const date = _fz_parseDateFlex(dateStr);
    const action = /hinzuf|add/i.test(actRaw) ? 'add' : 'remove';
    const id = String(idRaw).replace(/^0+/, '');
    const name = (nameRaw || '').trim();

    if (date && id) {
      const obj = { date: dateStr, action, id };
      if (name && action === 'add') obj.name = name;
      out.push(obj);
    }
  }
  return out;
}

// -- Konfig lesen (aus externer Datei via window.VEHICLE_CHANGES) ------------
const VEHICLE_CHANGES = _fz_normalizeChanges(window.VEHICLE_CHANGES);

// -- Sortiert & validiert (alt → neu) ---------------------------------------
const _VEHICLE_CHANGES_SORTED = (VEHICLE_CHANGES || [])
  .map(x => ({ ...x, _date: _fz_parseDateFlex(x.date) }))
  .filter(x => x._date instanceof Date && !Number.isNaN(x._date.valueOf()))
  .sort((a, b) => a._date - b._date);

// -- Aktive Fahrzeuge für ein Datum -----------------------------------------
function vehiclesForDate(d) {
  let ids = Array.isArray(FAHRZEUGE) ? [...FAHRZEUGE] : [];
  const names = { ...(FAHRZEUGNAMEN || {}) };

  for (const ch of _VEHICLE_CHANGES_SORTED) {
    if (ch._date > d) break;
    const vid = String(ch.id);
    if (ch.action === 'add') {
      if (!ids.includes(vid)) ids.push(vid);
      if (ch.name) names[vid] = ch.name;
    } else if (ch.action === 'remove') {
      ids = ids.filter(x => x !== vid);
      delete names[vid]; // Konsistenz: evtl. gesetzten Namen ebenfalls entfernen
    }
  }

  const rpd = Math.min(Math.max((ids.length || 1), MIN_ZEILEN_PRO_TAG), MAX_ZEILEN_PRO_TAG);
  return { ids: ids.slice(0, rpd), names };
}

// -- Hilfen für Legende ------------------------------------------------------
function _legendBaseDateFromSelection() {
  const jahr = parseInt(document.getElementById('jahr')?.value, 10) || (new Date()).getFullYear();
  const monat = parseInt(document.getElementById('monat')?.value, 10);
  return (monat === -1)
    ? { year: jahr, month: 0, isYear: true }
    : { year: jahr, month: monat, isYear: false };
}

function _buildChangeNotesForLegend() {
  const sel = _legendBaseDateFromSelection();
  const notes = [];
  for (const ch of _VEHICLE_CHANGES_SORTED) {
    const y = ch._date.getFullYear();
    const m = ch._date.getMonth();
    const inScope = sel.isYear ? (y === sel.year) : (y === sel.year && m === sel.month);
    if (!inScope) continue;

    const label = (ch.action === 'add') ? 'hinzugefügt' : 'entfernt';
    const nm = (ch.name || (FAHRZEUGNAMEN?.[ch.id]) || '').trim();
    const who = nm ? `${ch.id} – ${nm}` : `${ch.id}`;
    notes.push(`${who} ${label} am ${_fz_formatDE(ch._date)}`);
  }
  return notes;
}

// -- OVERRIDE: Legende -------------------------------------------------------
// ersetzt die vorhandene Funktion vollständig
function renderFahrzeugLegende() {
  const box = document.getElementById('vehicleLegend');
  if (!box) return;

  const { year, month, isYear } = _legendBaseDateFromSelection();
  const { ids, names } = vehiclesForDate(new Date(year, month, 1));
  const ordered = [...ids].sort((a, b) => a.localeCompare(b, 'de', { numeric: true }));

  const listRows = ordered.map(id => `
    <div class="veh-legend-row">
      <div class="veh-legend-id">${escapeText(id)}</div>
      <div class="veh-legend-body">
        <div class="veh-legend-head">
          <div class="veh-legend-name">${escapeText(names[id] || '')}</div>
        </div>
        <div class="veh-meta muted"></div>
      </div>
    </div>
  `).join('');

  // Änderungs-Hinweise (nur in der Legende)
  const notes = _buildChangeNotesForLegend();
  const notesHtml = notes.length
    ? `
      <div class="veh-legend-changes" style="margin-top:12px;border-top:1px solid #ddd;padding-top:8px;">
        <div class="muted" style="font-weight:600;margin-bottom:4px;">
          Änderungen ${isYear ? `im Jahr ${year}` : `in diesem Monat`}
        </div>
        <ul style="margin:0;padding-left:18px;">
          ${notes.map(n => `<li style="font-size:12px;color:#666;">${escapeText(n)}</li>`).join('')}
        </ul>
      </div>
    `
    : '';

  box.innerHTML = `<h3>Fahrzeuge</h3><div class="veh-group">${listRows}</div>${notesHtml}`;
  reflowRightRail?.();
}

// -- OVERRIDE: Tabelle erstellen --------------------------------------------
// ersetzt die vorhandene Funktion vollständig
function tabelleErstellen(jahr, monatIndex, targetTbody = null) {
  const tbody = targetTbody || document.querySelector('#monatsTabelle tbody');
  if (!tbody) return;
  tbody.innerHTML = '';

  const readOnlyMulti = !!targetTbody;
  const tageImMonat = new Date(jahr, monatIndex + 1, 0).getDate();

  // gespeicherte Monatsdaten blockweise (pro Tag) nach Fahrzeug-ID abbilden
  const loadSavedRowsMapForDay = (j, m, day, dim) => {
    const saved = ladeDaten(j, m) || [];
    const meta = ladeMeta(j, m);
    const tage = dim || new Date(j, m + 1, 0).getDate();

    let savedRPD = parseInt(meta?.rowsPerDay, 10);
    if (!Number.isFinite(savedRPD) || savedRPD < 1) {
      savedRPD = Math.max(1, Math.round(saved.length / Math.max(1, tage)));
    }
    const start = (day - 1) * savedRPD;
    const block = saved.slice(start, start + savedRPD);

    const map = new Map();
    for (const row of block) {
      if (!row || !row.length) continue;
      const id = String(row[0] ?? '').trim();
      if (!id) continue;
      map.set(id, [row[1] ?? '', row[2] ?? '', row[3] ?? '', row[4] ?? '', row[5] ?? '', row[6] ?? '']);
    }
    return map;
  };

  LAST_RPD_PER_DAY = [];
  let zeilenIndexGlobal = 0;

  for (let tag = 1; tag <= tageImMonat; tag++) {
    const datum = new Date(jahr, monatIndex, tag);
    const datumText = datum.toLocaleDateString('de-DE', { weekday: 'long', day: 'numeric', month: 'long' });
    const feiertagName = feiertagsNameFür(datum);

    // Tages-Fahrzeugliste laut Zeitachse
    const { ids: dayVehicles, names: dayNames } = vehiclesForDate(datum);
    const rpdToday = Math.max(MIN_ZEILEN_PRO_TAG, dayVehicles.length || 1);
    LAST_RPD_PER_DAY[tag] = rpdToday;

    const savedMap = loadSavedRowsMapForDay(jahr, monatIndex, tag, tageImMonat);

    for (let zeile = 0; zeile < rpdToday; zeile++) {
      const tr = document.createElement('tr');

      if (feiertagName) tr.classList.add('feiertag');
      if (datum.getDay() === 0 || datum.getDay() === 6) tr.classList.add('weekend');
      if (!tr.classList.contains('weekend') && !tr.classList.contains('feiertag') && (zeile % 2 === 1)) {
        tr.classList.add('row-alt');
      }

      // Datumsspalte (gemergt)
      if (zeile === 0) {
        tr.classList.add('first-of-day');

        const tdDate = document.createElement('td');
        tdDate.classList.add('col-datum');

        const wrap = document.createElement('div');
        const dt = document.createElement('div');
        dt.className = 'datum-text';
        const datumTextFormatted = datumText.charAt(0).toUpperCase() + datumText.slice(1);
        dt.textContent = datumTextFormatted;
        wrap.appendChild(dt);
        if (feiertagName) {
          const fh = document.createElement('div');
          fh.className = 'feiertag-name';
          fh.textContent = feiertagName;
          wrap.appendChild(fh);
        }
        tdDate.appendChild(wrap);
        tdDate.rowSpan = rpdToday;
        tr.appendChild(tdDate);
      }

      // restliche Spalten
      for (let i = 0; i < 7; i++) {
        const colKey = COLS[i]; // 'fahrzeug' | 'dienstposten' | 'ort' | 'von' | 'bis' | 'bemerkung' | 'eingetragen'
        const td = document.createElement('td');
        td.classList.add(`col-${colKey}`);
        td.dataset.colKey = colKey;

        if (colKey === 'fahrzeug') {
          td.classList.add('text-center');

          const id = dayVehicles[zeile] || '';
          const label = document.createElement('span');
          label.textContent = id;
          label.style.fontWeight = 'bold';
          if (id) {
            const t = (dayNames && dayNames[id]) || (FAHRZEUGNAMEN && FAHRZEUGNAMEN[id]) || '';
            if (t) label.title = t;
          }
          td.appendChild(label);

          if (!readOnlyMulti) {
            const hidden = document.createElement('input');
            hidden.type = 'hidden';
            hidden.value = id;
            td.appendChild(hidden);
          }
        } else {
          const idForRow = dayVehicles[zeile] || '';
          const savedFields = (savedMap.get(idForRow) || ['', '', '', '', '', '']);
          const fieldIndex = i - 1;
          const vorbefuellt = (fieldIndex >= 0 ? savedFields[fieldIndex] : '') || '';

          if (readOnlyMulti) {
            const div = document.createElement('div');
            div.className = 'ro';
            div.textContent = vorbefuellt;
            td.appendChild(div);
          } else {
            let inputEl;
            if (typeof makeInput === 'function') {
              inputEl = makeInput(colKey, vorbefuellt);
            } else {
              inputEl = document.createElement('input');
              inputEl.type = 'text';
              inputEl.maxLength = INPUT_MAX[colKey] ?? 10;
              inputEl.value = vorbefuellt;
            }
            inputEl.classList.add('cell-input', `cell-${colKey}`);
            inputEl.dataset.orig = vorbefuellt;
            inputEl.dataset.row = zeilenIndexGlobal;
            inputEl.dataset.col = i;
            inputEl.addEventListener('input', () => {
              if (bearbeitungGesperrt) return;
              setUnsaved(true);
              updateChangedFlag(inputEl);
            });
            td.appendChild(inputEl);
          }
        }

        tr.appendChild(td);
      }

      tbody.appendChild(tr);
      zeilenIndexGlobal++;
    }
  }
}
// -----------------------------------------------------------------------------
// #endregion





// #region Supabase Realtime (CLIENT) – Einbindung & sanfte Overrides
//
// 1) 🔐 Supabase-Projekt-Schlüssel eintragen
const SUPABASE_URL = "https://<DEIN-PROJEKT>.supabase.co";
const SUPABASE_ANON_KEY = "<DEIN-ANON-KEY>";
const SB_TABLE = "monthly_html";     // Tabellenname wie verwendet

// ─────────────────────────────────────────────────────────────────────────────
// 2) 📌 Einmalig in Supabase (SQL) ausführen – NICHT hier im JS, nur zur Info:
//
//   -- Realtime braucht vollständige Row-Images
//   alter table public.monthly_html replica identity full;
//
//   -- Publication anlegen (falls noch nicht vorhanden)
//  do $$
//   begin
//     if not exists (select 1 from pg_publication where pubname = 'supabase_realtime') then
//       create publication supabase_realtime;
//     end if;
//   end $$;
//
//   -- Tabelle zu Realtime-Publication hinzufügen
//   alter publication supabase_realtime add table public.monthly_html;
// ─────────────────────────────────────────────────────────────────────────────
//
// 3) ⤵️ Client-Setup (lazy import)
let supabase = null;
async function ensureSupabase() {
  if (supabase) return supabase;
  const { createClient } = await import("https://esm.sh/@supabase/supabase-js@2");
  supabase = createClient(SUPABASE_URL, SUPABASE_ANON_KEY);
  return supabase;
}

// 4) Helpers (IDs der aktuellen Auswahl)
function _selYM() {
  const y = parseInt(document.getElementById('jahr')?.value ?? new Date().getFullYear(), 10);
  const m = parseInt(document.getElementById('monat')?.value ?? new Date().getMonth(), 10);
  return { year: y, month: m };
}

// 5) Datensatz laden → tbody ersetzen (OVERRIDE-Logik)
async function sbLoadTbody(year, month) {
  try {
    await ensureSupabase();
    if (month === -1) return; // Jahresansicht: kein einzelnes <tbody>
    const { data, error } = await supabase
      .from(SB_TABLE)
      .select("html")
      .eq("year", year)
      .eq("month", month)
      .maybeSingle();

    if (error) { console.warn("[Supabase] SELECT Fehler:", error); return; }
    if (!data?.html) return;

    const tb = document.querySelector("#monatsTabelle tbody");
    if (!tb) return;

    tb.innerHTML = data.html;

    // Nachladen: Eingabe-Events & UI-Status wiederherstellen
    if (typeof bindCellInputEvents === "function") bindCellInputEvents();
    if (typeof clearChangeMarkersAndRebase === "function") clearChangeMarkersAndRebase();
    if (typeof setUnsaved === "function") setUnsaved(false);
  } catch (e) {
    console.warn("[Supabase] Load Fehler:", e);
  }
}

// 6) tbody speichern (Upsert) (OVERRIDE-Logik)
async function sbSaveTbody(year, month) {
  try {
    await ensureSupabase();
    if (month === -1) return;
    const tb = document.querySelector("#monatsTabelle tbody");
    if (!tb) return;

    const html = tb.innerHTML;

    // Existiert bereits?
    const { data: existing, error: selErr } = await supabase
      .from(SB_TABLE)
      .select("id")
      .eq("year", year)
      .eq("month", month)
      .maybeSingle();
    if (selErr) throw selErr;

    if (existing?.id) {
      const { error: updErr } = await supabase
        .from(SB_TABLE)
        .update({ html, updated_at: new Date().toISOString() })
        .eq("id", existing.id);
      if (updErr) throw updErr;
    } else {
      const { error: insErr } = await supabase
        .from(SB_TABLE)
        .insert({ year, month, html });
      if (insErr) throw insErr;
    }
  } catch (e) {
    console.error("[Supabase] Save Fehler:", e);
    throw e;
  }
}

// 7) Realtime abonnieren (INSERT/UPDATE) (OVERRIDE-Logik)
let _sbRealtimeSubscribed = false;
async function sbSubscribeRealtime() {
  try {
    await ensureSupabase();
    if (_sbRealtimeSubscribed) return;
    supabase
      .channel(`public:${SB_TABLE}`)
      .on(
        "postgres_changes",
        { event: "*", schema: "public", table: SB_TABLE },
        (payload) => {
          // Nur reloaden, falls die Nachricht den aktuell angezeigten Monat betrifft
          const { year, month } = _selYM();
          const row = payload.new ?? payload.old ?? {};
          if (row.year === year && row.month === month) {
            const tb = document.querySelector("#monatsTabelle tbody");
            if (!tb) return;
            if (payload.new?.html) {
              tb.innerHTML = payload.new.html;
              if (typeof bindCellInputEvents === "function") bindCellInputEvents();
              if (typeof clearChangeMarkersAndRebase === "function") clearChangeMarkersAndRebase();
              if (typeof setUnsaved === "function") setUnsaved(false);
            }
          }
        }
      )
      .subscribe();
    _sbRealtimeSubscribed = true;
  } catch (e) {
    console.warn("[Supabase] Realtime subscribe Fehler:", e);
  }
}

// 8) VERÄNDERT: „aktualisieren“ sanft patchen, damit nach jedem Render
//    (Monat wechselt) automatisch aus Supabase geladen und Realtime aktiv ist,
//    ohne deinen Originalcode hart zu überschreiben.
(function patchAktualisieren() {
  const original = window.aktualisieren;
  window.aktualisieren = function patchedAktualisieren(...args) {
    // 8a) Original ausführen (dein bestehendes Rendern)
    try { original?.apply(this, args); } catch (e) { console.warn("aktualisieren(original) Fehler:", e); }

    // 8b) Danach: für Einzelmonat aus Supabase nachziehen
    try {
      const jahr = parseInt(document.getElementById('jahr')?.value ?? "0", 10);
      const monat = parseInt(document.getElementById('monat')?.value ?? "-1", 10);
      if (Number.isFinite(jahr) && Number.isFinite(monat) && monat !== -1) {
        sbLoadTbody(jahr, monat);
      }
    } catch {}

    // 8c) Realtime beim ersten Mal aktivieren
    sbSubscribeRealtime();
  };
})();

// 9) VERÄNDERT: Speichern-Button zusätzlich mit Supabase-Speichern verbinden.
//    (Falls du bereits einen Listener hast, macht doppelt nichts kaputt –
//     wir hängen uns nur hinzu und rufen sbSaveTbody auf.)
(function wireSaveButtonToSupabase() {
  const saveBtn = document.getElementById('saveBtn');
  if (!saveBtn) return;
  saveBtn.addEventListener('click', async () => {
    try {
      const { year, month } = _selYM();
      await sbSaveTbody(year, month);
    } catch (e) {
      // Fehler bereits geloggt
    }
  });
})();

// 10) Beim ersten Laden direkt Realtime scharf schalten
sbSubscribeRealtime();
// #endregion





// #region Init/Wiring

(function init() {
  (async function init() {
    // Auswahl vorbereiten
    baueGrundAuswahl();

    const monatSel = document.getElementById('monat');
    const jahrSel = document.getElementById('jahr');
    const saveBtn = document.getElementById('saveBtn');

    const onPickChange = () => { aktualisieren(); afterRenderApply(); };
    monatSel?.addEventListener('change', onPickChange);
    jahrSel?.addEventListener('change', onPickChange);
    jahrSel?.addEventListener('input', onPickChange);

    // erster Render
    aktualisieren();
    afterRenderApply();


    // Lock/Unlock
    document.getElementById('unlockBtn')?.addEventListener('click', entsperren);
    document.getElementById('lockBtn')?.addEventListener('click', sperren);

    // Export-Buttons
    document.getElementById('exportXlsxBtn')?.addEventListener('click', async () => {
      const curMonth = parseInt(monatSel.value, 10);
      if (curMonth === -1) await exportYearWithExcelJS_DOMMirror();
      else await exportMitExcelJS();
    });

    // Import: ein versteckter File-Input
    let fileInput = document.getElementById('importExcelInput');
    if (!fileInput) {
      fileInput = document.createElement('input');
      fileInput.type = 'file';
      fileInput.accept = '.xlsx,.xls,.xlsm';
      fileInput.id = 'importExcelInput';
      fileInput.style.display = 'none';
      document.body.appendChild(fileInput);
    }
    document.getElementById('importXlsxBtn')?.addEventListener('click', () => fileInput.click());
    fileInput.addEventListener('change', async (e) => {
      const file = e.target.files?.[0];
      if (!file) return;
      try {
        await importAllMonthYearSheets(file);
      } catch (err) {
        console.error(err);
        alert('Import fehlgeschlagen: ' + (err?.message || err));
      } finally {
        e.target.value = '';
      }
    });

    // Speichern (inkl. Platz für externen Backup-Block)
    saveBtn?.addEventListener('click', async (e) => {
      e.preventDefault();
      e.stopPropagation();

      const jahr = parseInt(document.getElementById('jahr')?.value, 10);
      const monat = parseInt(document.getElementById('monat')?.value, 10);

      try {
        speichereDaten(jahr, monat);
      } catch (err) {
        console.warn('Speichern fehlgeschlagen:', err);
        return;
      }

      setUnsaved(false);
      clearChangeMarkersAndRebase();
      document.activeElement?.blur?.();
      document.body.classList.add('no-hover');
      setTimeout(() => document.body.classList.remove('no-hover'), 200);
    });

  })();

  // Optional: Lightbox/Escape
  const lightbox = document.getElementById('imgLightbox');
  if (lightbox) {
    lightbox.addEventListener('click', () => (lightbox.style.display = 'none'));
    window.addEventListener('keydown', (e) => { if (e.key === 'Escape') lightbox.style.display = 'none'; });
  }

  window.addEventListener('resize', reflowRightRail);
})();

// #endregion







