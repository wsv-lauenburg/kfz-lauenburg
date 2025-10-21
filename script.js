// --- RAM-Store (keine Persistenz) ------------------------------------------
const STORE = {
  // key: `${year}-${monthIndex}` -> { rowsPerDay: number, rows: Array<[id, dienstposten, ort, von, bis, bemerkung, eingetragen]> }
  byMonth: new Map()
};

function monthKey(y, m) { return `${y}-${m}`; }

function setMonthData(y, m, rowsPerDay, rows) {
  STORE.byMonth.set(monthKey(y, m), { rowsPerDay, rows });
}

function getMonthData(y, m) {
  return STORE.byMonth.get(monthKey(y, m)) || {
    rowsPerDay: Math.max(1, FAHRZEUGE.length || 1),
    rows: []
  };
}

// === Supabase: Setup =======================================================
const SUPABASE_URL = "https://gsigwwrepcafkwdvadhk.supabase.co";
const SUPABASE_ANON_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImdzaWd3d3JlcGNhZmt3ZHZhZGhrIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NjA2NjY4MzIsImV4cCI6MjA3NjI0MjgzMn0.cNNzK_OhUlMTCtUdBMxXaazQCZbf0oa0JlS7abA1Lwk";

let supabase; // wird gleich gesetzt
(async () => {
  const { createClient } = await import("https://esm.sh/@supabase/supabase-js@2");
  supabase = createClient(SUPABASE_URL, SUPABASE_ANON_KEY);
})();




// #region Handy Input verhindern ------------------ */


// Nur auf echten Phones (kleine Device-Breite + Touch + kein Hover) -> read-only
function applyPhoneReadOnly() {
  const isPhone =
    matchMedia("(max-device-width: 820px)").matches &&   // echte Gerätebreite, kein Desktop-Resize
    matchMedia("(pointer: coarse)").matches &&           // Touch-Eingabe
    matchMedia("(hover: none)").matches;                 // kein Hover

  document.body.classList.toggle("mobile-readonly", isPhone);

  // Inputs hard read-only machen
  document.querySelectorAll("#monatsTabelle input").forEach(inp => {
    inp.disabled = isPhone;
  });
}

// Initial + bei Orientierung/Resize (Gerätewechsel)
document.addEventListener("DOMContentLoaded", applyPhoneReadOnly);
addEventListener("orientationchange", applyPhoneReadOnly);
addEventListener("resize", applyPhoneReadOnly);


// #endregion ---------------------------------------------------------------- */


// #region Heute-Quelle (DEV-Override von "heute") --------------------------- */
const IS_DEV = ['localhost', '127.0.0.1'].includes(location.hostname);

let __t = null;
const toISO = d => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
const parseISO = s => {
  const m = s?.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  return m ? new Date(+m[1], +m[2] - 1, +m[3]) : null;
};
const getTodayDate = () => (IS_DEV && __t ? parseISO(__t) : new Date());
const getTodayISO = () => toISO(getTodayDate());

if (IS_DEV) {
  window.testToday = s => {
    const d = parseISO(s);
    if (d) {
      __t = toISO(d);
      console.info('Test-Heute:', __t);
      aktualisieren?.();
      renderFahrzeugLegende?.();
    }
  };
  window.clearTestToday = () => {
    __t = null;
    console.info('Echtes Heute aktiv');
    aktualisieren?.();
    renderFahrzeugLegende?.();
  };
  const q = new URLSearchParams(location.search).get('today');
  if (q) __t = toISO(parseISO(q));
}
// #endregion ---------------------------------------------------------------- */


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

function excelDateUTC(y, m, d) {
  return new Date(Date.UTC(y, m, d)); // 00:00 UTC (kein Shift)
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

function parseGermanDateToISO(d) {
  if (!d) return null;
  if (d instanceof Date) return toISO(d);
  const s = String(d).trim();
  const m = s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{2}|\d{4})$/);
  if (m) {
    let [_, dd, mm, yy] = m;
    if (yy.length === 2) yy = `20${yy}`;
    return `${yy}-${mm.padStart(2, '0')}-${dd.padStart(2, '0')}`;
  }
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  return null;
}

function deDate(iso) {
  const [y, m, d] = iso.split('-').map(Number);
  return `${String(d).padStart(2, '0')}.${String(m).padStart(2, '0')}.${y}`;
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

// Persistenz-Keys
const FAHRZEUGE_KEY = 'fahrzeugListe';
const FAHRZEUGNAMEN_KEY = 'fahrzeugNamen';

// Basis-Fahrzeuge (ohne Aktiv/Events)
let FAHRZEUGE = ["700", "220", "221", "222", "223", "231"];
let FAHRZEUGNAMEN = {};
let ZEILEN_PRO_TAG = Math.min(Math.max(FAHRZEUGE.length, MIN_ZEILEN_PRO_TAG), MAX_ZEILEN_PRO_TAG);

// Save-Bar
let unsavedChanges = false;

// „Sperren“ – rein UI-seitig (kein echter Schutz)
const PASSWORT = '1234';
let bearbeitungGesperrt = false;

// Laufzeit-Merker für tatsächlich gerenderte Zeilen je Tag (für Speichern)
let LAST_RPD_PER_DAY = [];   // Index 1..tageImMonat

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
  inp.type = 'text';                 // immer Text – kein time
  inp.maxLength = INPUT_MAX[colKey] ?? 10;
  inp.value = value;
  return inp;
}



function snapshotTable() {
  const rows = [];
  document.querySelectorAll('#monatsTabelle tbody tr').forEach((tr) => {
    const cols = [];
    tr.querySelectorAll('input').forEach((input) => cols.push(input.value));
    rows.push(cols);
  });
  return rows;
}

// #endregion ---------------------------------------------------------------- */


// #region Jahr/Monat Auswahl & Aktualisieren -------------------------------- */

function baueGrundAuswahl() {
  baueJahresAuswahl();
  const heute = getTodayDate();
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
    jahr = (getTodayDate()).getFullYear();
    const sel = document.getElementById('jahr');
    if (sel && sel.value !== String(jahr)) sel.value = String(jahr);
  }
  if (jahr < YEAR_START || jahr > YEAR_END) {
    alert(`Bitte ein Jahr zwischen ${YEAR_START} und ${YEAR_END} auswählen.`);
    return;
  }

  try { persistCurrentMonthToStore(); persistStoreToLS(); } catch { }

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
    sbLoadTbody(jahr, monat);
  }
  afterRenderApply?.();
}

// #endregion ---------------------------------------------------------------- */


// #region Fahrzeug-Legende (ohne aktiv/inaktiv) ------------------------------ */

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
    saveVehicleState();
  }
}

// #endregion ---------------------------------------------------------------- */


// #region Tabelle: Erstellen, Monats-/Jahresansicht

function tabelleErstellen(jahr, monatIndex, targetTbody = null) {
  const tbody = targetTbody || document.querySelector('#monatsTabelle tbody');
  tbody.innerHTML = '';

  const readOnlyMulti = !!targetTbody;
  const tageImMonat = new Date(jahr, monatIndex + 1, 0).getDate();

  const { rowsPerDay: rowsPerDayFromStore, rows } = getMonthData(jahr, monatIndex);
  const rowsPerDayToday = Math.max(MIN_ZEILEN_PRO_TAG, rowsPerDayFromStore || FAHRZEUGE.length || 1);

  // tag -> Map(fahrzeugId -> [dienstposten, ort, von, bis, bemerkung, eingetragen])
  const savedPerDay = new Map();
  for (let day = 1; day <= tageImMonat; day++) {
    const start = (day - 1) * rowsPerDayToday;
    const slice = rows.slice(start, start + rowsPerDayToday);
    const map = new Map();
    (slice || []).forEach(r => {
      const id = String(r?.[0] ?? '').trim();
      if (!id) return;
      map.set(id, [r[1] ?? '', r[2] ?? '', r[3] ?? '', r[4] ?? '', r[5] ?? '', r[6] ?? '']);
    });
    savedPerDay.set(day, map);
  }

  LAST_RPD_PER_DAY = [];
  let zeilenIndexGlobal = 0;

  for (let tag = 1; tag <= tageImMonat; tag++) {
    const datum = new Date(jahr, monatIndex, tag);
    const datumText = datum.toLocaleDateString('de-DE', { weekday: 'long', day: 'numeric', month: 'long' });
    const feiertagName = feiertagsNameFür(datum);

    LAST_RPD_PER_DAY[tag] = rowsPerDayToday;
    const savedMap = savedPerDay.get(tag) || new Map();

    for (let zeile = 0; zeile < rowsPerDayToday; zeile++) {
      const tr = document.createElement('tr');

      if (feiertagName) tr.classList.add('feiertag');
      if (datum.getDay() === 0 || datum.getDay() === 6) tr.classList.add('weekend');
      if (!tr.classList.contains('weekend') && !tr.classList.contains('feiertag')) {
        if (zeile % 2 === 1) tr.classList.add('row-alt');
      }

      if (zeile === 0) {
        tr.classList.add('first-of-day');
        const tdDate = document.createElement('td');
        tdDate.classList.add('col-datum');
        const wrap = document.createElement('div');
        const dt = document.createElement('div');
        dt.className = 'datum-text';
        dt.textContent = datumText.charAt(0).toUpperCase() + datumText.slice(1);
        wrap.appendChild(dt);
        if (feiertagName) {
          const fh = document.createElement('div');
          fh.className = 'feiertag-name';
          fh.textContent = feiertagName;
          wrap.appendChild(fh);
        }
        tdDate.appendChild(wrap);
        tdDate.rowSpan = rowsPerDayToday;
        tr.appendChild(tdDate);
      }

      for (let i = 0; i < 7; i++) {
        const colKey = COLS[i];
        const td = document.createElement('td');
        td.classList.add(`col-${colKey}`);
        td.dataset.colKey = colKey;

        if (colKey === 'fahrzeug') {
          td.classList.add('text-center');
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
            const inputEl = makeInput(colKey, vorbefuellt);
            inputEl.classList.add('cell-input', `cell-${colKey}`);
            inputEl.dataset.orig = vorbefuellt;
            inputEl.dataset.row = zeilenIndexGlobal;
            inputEl.dataset.col = i;
            inputEl.addEventListener('input', () => {
              if (bearbeitungGesperrt) return;
              setUnsaved(true);
              updateChangedFlag(inputEl);
              persistDebounced();
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
    jahr = Number.isFinite(uiVal) ? uiVal : (getTodayDate()).getFullYear();
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


// #region Excel: Sheet-/Workbook-Utilities 

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

function monthIndexFromName(name) {
  if (!name) return null;
  let s = String(name).toLowerCase();
  try { s = s.normalize('NFD').replace(/\p{Diacritic}/gu, ''); } catch { }
  const candidates = [
    ['jan', 'januar'],
    ['feb', 'februar'],
    ['maer', 'maerz', 'marz', 'mar', 'mär', 'maerz'],
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
  for (let i = 0; i < candidates.length; i++) {
    if (candidates[i].some(tok => s.includes(tok))) return i;
  }
  const mNum = s.match(/\b(0?[1-9]|1[0-2])(?=(?:\D|$))/);
  if (mNum) return parseInt(mNum[1], 10) - 1;
  return null;
}

function detectYear(sheetName, aoa) {
  if (typeof sheetName === 'string') {
    const m = sheetName.match(/\b(20\d{2}|19\d{2})\b/);
    if (m) return parseInt(m[1], 10);
  }
  const MAX_R = Math.min(25, aoa.length);
  const freq = new Map();
  for (let r = 0; r < MAX_R; r++) {
    const row = aoa[r] || [];
    const MAX_C = Math.min(12, row.length);
    for (let c = 0; c < MAX_C; c++) {
      const text = String(row[c] ?? '');
      const hits = text.match(/\b(20\d{2}|19\d{2})\b/g);
      if (hits) for (const y of hits) {
        const yi = parseInt(y, 10);
        if (yi >= 1900 && yi <= 2100) freq.set(yi, (freq.get(yi) || 0) + 1);
      }
    }
  }
  if (freq.size) return [...freq.entries()].sort((a, b) => b[1] - a[1])[0][0];
  return null;
}

// #endregion


// #region Excel: Import – eigenes Exportformat + generisch

function tryParseOwnExport(aoa, headerRowIdx, targetYear, monthIndex) {
  const hdr = (aoa[headerRowIdx] || []).map(s => String(s ?? '').trim().toLowerCase());
  const want = ["tag", "fahrzeug", "dienstposten", "ort", "von", "bis", "bemerkung", "eingetragen"];
  const looksOk = want.every((h, i) => (hdr[i] || '').startsWith(h));
  if (!looksOk) return null;

  const MIN = 1, MAX = 30;
  let rowsPerDay = Math.min(Math.max(FAHRZEUGE.length || ZEILEN_PRO_TAG || 1, MIN), MAX);

  const bodyRows = aoa.slice(headerRowIdx + 1);

  // Nutzspalten B..H (Index 1..7)
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

  // WICHTIG: NICHT speichern, nur zurückgeben
  return { rowsPerDay, rowsArray: out };
}


function parseSheetToStorage(ws, targetYear, monthIndex) {
  const norm = (s) => String(s ?? '').trim().toLowerCase();
  const safeCell = (row, idx) =>
    (Number.isInteger(idx) && idx >= 0 && idx < row.length && row[idx] != null)
      ? String(row[idx]).trim() : '';

  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false });
  if (!aoa || !aoa.length) throw new Error('Keine Daten im Blatt.');

  // Headerzeile suchen
  let headerRowIdx = -1;
  for (let i = 0; i < Math.min(10, aoa.length); i++) {
    const row = (aoa[i] || []).map(x => norm(x));
    if (row.includes('fahrzeug') && row.includes('dienstposten') &&
      (row.includes('bemerkung') || row.includes('bemerkungen'))) { headerRowIdx = i; break; }
  }
  if (headerRowIdx === -1) headerRowIdx = 0;

  // eigener Export?
  const own = tryParseOwnExport(aoa, headerRowIdx, targetYear, monthIndex);
  if (own) {
    setMonthData(targetYear, monthIndex, own.rowsPerDay, own.rowsArray);
    return own;
  }

  // generischer Parser
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
    setMonthData(targetYear, monthIndex, rowsPerDayEmpty, []);
    return { rowsPerDay: rowsPerDayEmpty, rowsArray: [] };
  }

  let rowsPerDay = Math.max(...dayRows.map((b, i) => i ? b.length : 0));
  rowsPerDay = Math.min(Math.max(rowsPerDay, MIN_ZEILEN_PRO_TAG), MAX_ZEILEN_PRO_TAG);

  const normalized2D = [];
  for (let d = 1; d <= tageImMonat; d++) {
    const block = (dayRows[d] || []).slice(0, rowsPerDay);
    while (block.length < rowsPerDay) block.push(emptyRow());
    normalized2D.push(...block);
  }

  setMonthData(targetYear, monthIndex, rowsPerDay, normalized2D);
  return { rowsPerDay, rowsArray: normalized2D };
}

// #endregion 


// #region Tabelle → Array & Helfer für Export

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

// #endregion


// #region Excel: Shared Builder (identische Optik für Export & Backup)

function monthNamesDE() {
  return ["Januar", "Februar", "März", "April", "Mai", "Juni",
    "Juli", "August", "September", "Oktober", "November", "Dezember"];
}

function styleHeaderRowExcel(ws, rowIdx, blueHeader) {
  ws.getRow(rowIdx).eachCell(c => {
    c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: blueHeader } };
    c.font = { bold: true, color: { argb: "FFFFFFFF" } };
    c.alignment = { horizontal: "center", vertical: "middle" };
    c.border = {
      top: { style: "thin" }, left: { style: "thin" },
      bottom: { style: "thin" }, right: { style: "thin" }
    };
  });
}

/** Baut EIN Monats-Blatt exakt wie im normalen Export */
async function buildWorkbookForMonthFromDOM(jahr, monat) {
  if (typeof ExcelJS === 'undefined') throw new Error('ExcelJS ist nicht geladen.');

  const wb = new ExcelJS.Workbook();
  const monate = monthNamesDE();

  const blueHeader = "2F75B5";
  const zebraARGB = "F2F2F2";
  const weekendARGB = "FFF6DFD1";
  const holidayARGB = "FFDFEEDD";

  const ws = wb.addWorksheet(`${monate[monat]} ${jahr}`, {
    properties: { defaultRowHeight: 18 },
    pageSetup: { paperSize: 9, orientation: "landscape", fitToPage: true }
  });

  // Titel wie im Export
  const titel = `Tabelle für ${monate[monat]} ${jahr}`;
  ws.spliceRows(1, 0, [titel]);
  ws.mergeCells(1, 1, 1, 9);
  const tcell = ws.getCell('A1');
  tcell.font = { bold: true, size: 16 };
  tcell.alignment = { horizontal: 'center', vertical: 'middle' };
  ws.getRow(1).height = 24;

  // Kopf
  ws.addRow(["Tag", "Fahrzeug", "Dienstposten", "Ort", "Von", "Bis", "Bemerkung", "eingetragen"]);
  styleHeaderRowExcel(ws, 2, blueHeader);

  // Spaltenbreiten + Legende-Spalten (GENAU wie im Export)
  ws.columns = [
    { width: 25 }, // A Datum
    { width: 12 }, // B Fahrzeug
    { width: 18 }, // C Dienstposten
    { width: 22 }, // D Ort
    { width: 12 }, // E Von
    { width: 12 }, // F Bis
    { width: 40 }, // G Bemerkung
    { width: 16 }, // H eingetragen
    { width: 2 }, // I Spacer
    { width: 14 }, // J Legende: ID
    { width: 28 }, // K Legende: Fahrzeugname
  ];

  // Daten AUS DEM DOM (wie exportMitExcelJS)
  const { tbody, cleanup } = renderMonthOffscreen(jahr, monat);
  const aoa = tabelleAlsArrayFromTbody(tbody);
  const rows = aoa.slice(1);
  cleanup();

  // rowsPerDay bestimmen wie im Export
  let rowsPerDay = FAHRZEUGE.length || 1;
  const firstDateCell = document.querySelector('#monatsTabelle tbody tr td[rowspan]');
  if (firstDateCell) {
    const rs = parseInt(firstDateCell.getAttribute('rowspan'), 10);
    if (!Number.isNaN(rs) && rs > 0) rowsPerDay = rs;
  }

  // evtl. numerische Fahrzeug-IDs als Zahl schreiben – wie Export
  const dataOnly = rows.map(r => {
    const asNum = Number(String(r[1]).trim().replace(',', '.'));
    const fahrzeug = Number.isFinite(asNum) ? asNum : r[1];
    return [r[0], fahrzeug, r[2], r[3], r[4], r[5], r[6], r[7]];
  });
  ws.addRows(dataOnly);

  // Rahmen + Zebra
  const startDataRow = 3;
  const endDataRow = ws.lastRow.number;
  for (let r = startDataRow; r <= endDataRow; r++) {
    ws.getRow(r).eachCell({ includeEmpty: true }, cell => {
      cell.border = {
        top: { style: "thin" }, left: { style: "thin" },
        bottom: { style: "thin" }, right: { style: "thin" }
      };
    });
    if ((r - startDataRow) % 2 === 1) {
      ws.getRow(r).eachCell({ includeEmpty: true }, cell => {
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: zebraARGB } };
      });
    }
  }

  // Tag-Blöcke mergen + WE/Feiertag färben
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
  }

  // Rechte Legende (WE/Feiertag + Fahrzeugliste)
  (function addRightLegend() {
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

  return wb;
}

/** Baut 12 Monats-Blätter (Jahr) exakt wie im normalen Export */
async function buildWorkbookForYearFromDOM(jahr) {
  if (typeof ExcelJS === 'undefined') throw new Error('ExcelJS ist nicht geladen.');
  const wb = new ExcelJS.Workbook();

  for (let m = 0; m < 12; m++) {
    const sub = await buildWorkbookForMonthFromDOM(jahr, m);
    const src = sub.worksheets[0];

    // Zielblatt mit denselben Basiseinstellungen
    const dst = wb.addWorksheet(src.name, {
      properties: src.properties,
      pageSetup: src.pageSetup
    });

    // Spaltenbreiten kopieren
    if (src.columns && src.columns.length) {
      dst.columns = src.columns.map(col => ({ width: col.width }));
    }

    // Zeilen + Zellen kopieren (Werte + Styles)
    src.eachRow({ includeEmpty: true }, (row, r) => {
      const newRow = dst.getRow(r);
      if (row.height) newRow.height = row.height;

      row.eachCell({ includeEmpty: true }, (cell, c) => {
        const n = newRow.getCell(c);
        n.value = cell.value;
        if (cell.style) n.style = { ...cell.style };
        if (cell.font) n.font = { ...cell.font };
        if (cell.alignment) n.alignment = { ...cell.alignment };
        if (cell.fill) n.fill = { ...cell.fill };
        if (cell.border) n.border = { ...cell.border };
        if (cell.numFmt) n.numFmt = cell.numFmt;
      });

      newRow.commit?.();
    });

    // Merge-Bereiche kopieren
    if (src._merges && src._merges.size) {
      for (const range of src._merges) {
        dst.mergeCells(range);
      }
    }
  }

  return wb;
}


// #endregion



// #region Excel-Export: Einzelmonat

async function exportMitExcelJS() {
  const jahr = parseInt(document.getElementById('jahr').value, 10);
  const monat = parseInt(document.getElementById('monat').value, 10);
  const monate = monthNamesDE();
  const wb = await buildWorkbookForMonthFromDOM(jahr, monat);
  const buf = await wb.xlsx.writeBuffer();
  const blob = new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = `Monatstabelle-${monate[monat]}-${jahr}.xlsx`;
  a.click();
  URL.revokeObjectURL(a.href);
}

// #endregion


// #region Excel-Export: Jahresdatei (DOM-Mirror pro Monat)

async function exportYearWithExcelJS_DOMMirror() {
  const jahr = parseInt(document.getElementById('jahr').value, 10) || (getTodayDate()).getFullYear();
  const wb = await buildWorkbookForYearFromDOM(jahr);
  const buf = await wb.xlsx.writeBuffer();
  const blob = new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = `Monatstabelle-AlleMonate-${jahr}.xlsx`;
  a.click();
  URL.revokeObjectURL(a.href);
}
const exportYearWithExcelJS = exportYearWithExcelJS_DOMMirror;

// #endregion


// #region Excel: Fahrzeugliste aus Workbook extrahieren

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


function saveVehicleState() {
  try {
    lsSetJson(FAHRZEUGE_KEY, FAHRZEUGE);
    lsSetJson(FAHRZEUGNAMEN_KEY, FAHRZEUGNAMEN);
  } catch {  }
}

(function restoreVehicleState() {
  try {
    const f = lsGetJson(FAHRZEUGE_KEY, null);
    if (Array.isArray(f) && f.length) FAHRZEUGE = f;

    const n = lsGetJson(FAHRZEUGNAMEN_KEY, null);
    if (n && typeof n === 'object') FAHRZEUGNAMEN = n;

    ZEILEN_PRO_TAG = Math.min(Math.max(FAHRZEUGE.length, MIN_ZEILEN_PRO_TAG), MAX_ZEILEN_PRO_TAG);
  } catch {  }
})();



function applyVehicleListFromWorkbook(wb) {
  for (const name of wb.SheetNames || []) {
    const ws = wb.Sheets[name];
    if (!ws) continue;
    const listOwn = extractVehicleList_OWN(XLSX.utils.sheet_to_json(ws, { header: 1, raw: false }) || []);
    if (listOwn.length) {
      FAHRZEUGE = listOwn.map(x => x.id);
      FAHRZEUGNAMEN = {};
      for (const { id, name } of listOwn) FAHRZEUGNAMEN[id] = name;
      saveVehicleState();
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
      saveVehicleState();
      ZEILEN_PRO_TAG = Math.min(Math.max(FAHRZEUGE.length, MIN_ZEILEN_PRO_TAG), MAX_ZEILEN_PRO_TAG);
      return true;
    }
  }

  const ids = collectVehiclesFromWorkbook(wb);
  if (ids.length) {
    FAHRZEUGE = ids;
    FAHRZEUGNAMEN = FAHRZEUGNAMEN || {};
    saveVehicleState();
    ZEILEN_PRO_TAG = Math.min(Math.max(FAHRZEUGE.length, MIN_ZEILEN_PRO_TAG), MAX_ZEILEN_PRO_TAG);
    return true;
  }

  return false;
}

// #endregion


// #region Excel: Import-Controller

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
  const curYear = parseInt(document.getElementById('jahr').value, 10) || (getTodayDate()).getFullYear();
  const curMonth = parseInt(document.getElementById('monat').value, 10) || 0;
  tabelleErstellen(curYear, curMonth);
  renderFahrzeugLegende();
  reflowRightRail?.();
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

// #endregion



// #region Lock/Unlock (UI, kein echter Schutz)

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



// #region Synchronisieren mit alln Benutzern




async function ensureLibs({ needXLSX=false, needExcelJS=false } = {}) {
  const deadline = Date.now() + 10000; // 10s
  while (true) {
    const ok = (!needXLSX || !!window.XLSX) && (!needExcelJS || !!window.ExcelJS);
    if (ok) return;
    if (Date.now() > deadline) throw new Error('Bibliotheken nicht geladen (XLSX/ExcelJS).');
    await new Promise(r => setTimeout(r, 50));
  }
}






// #region Auto-Import: neuestes Backup beim Laden --------------------------------

/** yyyy und (optional) mm aus Dateiname wie
 *  "Monatstabelle-AlleMonate-2026-....xlsx"  oder
 *  "Monatstabelle-2025-10-....xlsx"
 */
function _parseYearMonthFromBackupName(name) {
  const mYearAll = name.match(/AlleMonate-(20\d{2})/i);
  if (mYearAll) return { year: parseInt(mYearAll[1], 10), month: -1 };
  const mYearMon = name.match(/Monatstabelle-(20\d{2})-(0[1-9]|1[0-2])/i);
  if (mYearMon) return { year: parseInt(mYearMon[1], 10), month: parseInt(mYearMon[2], 10) - 1 };
  return null;
}

/** Neueste .xlsx im gespeicherten Ordner finden (still) und importieren.
 *  Gibt true zurück, wenn etwas importiert wurde.
 */
async function autoImportLatestBackup() {
  try {
    const dirHandle = await idbGet('sharedBackupDir');
    if (!dirHandle) { console.info('[autoImport] kein Ordner gespeichert'); return false; }

    let perm = 'granted';
    if (dirHandle.queryPermission) perm = await dirHandle.queryPermission({ mode: 'read' });
    console.info('[autoImport] queryPermission:', perm);

    // Ohne User-Geste wird requestPermission oft ignoriert → nicht hier versuchen.
    if (perm !== 'granted') {
      // UI-Hinweis anzeigen (Button klickbar machen)
      document.getElementById('backupFolderHint')?.classList.add('need-permission');
      console.info('[autoImport] keine Leserechte – zeige Hinweis/Button an');
      return false;
    }

    const files = [];
    for await (const [name, handle] of dirHandle.entries()) {
      if (handle.kind !== 'file' || !/\.xlsx?$/i.test(name)) continue;
      const f = await handle.getFile();
      files.push({ name, file: f, ts: f.lastModified });
    }
    if (!files.length) { console.info('[autoImport] kein .xlsx gefunden'); return false; }

    files.sort((a,b) => b.ts - a.ts);
    const latest = files[0];
    console.info('[autoImport] importiere:', latest.name);

    await ensureLibs({ needXLSX: true });
    await importAllMonthYearSheets(latest.file);
    return true;
  } catch (err) {
    console.warn('[autoImport] Fehler:', err);
    return false;
  }
}


/** Beim Start zuerst versuchen zu importieren, optional nur wenn STORE leer ist */
async function autoImportOnLoad({ onlyIfStoreEmpty = true } = {}) {
  const empty = STORE.byMonth.size === 0;
  if (!onlyIfStoreEmpty || empty) {
    const ok = await autoImportLatestBackup();
    return ok;
  }
  return false;
}

// #endregion -----------------------------------------------------------------



function idbSet(key, value) {
  return new Promise((res, rej) => {
    const req = indexedDB.open('shared-backup', 1);
    req.onupgradeneeded = () => req.result.createObjectStore('kv');
    req.onerror = () => rej(req.error);
    req.onsuccess = () => {
      const tx = req.result.transaction('kv', 'readwrite');
      tx.objectStore('kv').put(value, key);
      tx.oncomplete = () => res();
      tx.onerror = () => rej(tx.error);
    };
  });
}
function idbGet(key) {
  return new Promise((res, rej) => {
    const req = indexedDB.open('shared-backup', 1);
    req.onupgradeneeded = () => req.result.createObjectStore('kv');
    req.onerror = () => rej(req.error);
    req.onsuccess = () => {
      const tx = req.result.transaction('kv', 'readonly');
      const q = tx.objectStore('kv').get(key);
      q.onsuccess = () => res(q.result ?? null);
      q.onerror = () => rej(q.error);
    };
  });
}




async function autoRestoreFromSharedFolder() {
  try {
    const dirHandle = await idbGet('sharedBackupDir'); // zuvor via showDirectoryPicker gespeichert
    if (!dirHandle) return;

    // Berechtigung still prüfen/anfragen (funktioniert ohne User-Geste; zeigt evtl. Browser-Dialog)
    const perm = await dirHandle.requestPermission?.({ mode: 'read' }) ?? 'granted';
    if (perm !== 'granted') return; // kein stiller Zugriff möglich -> ggf. Button zeigen

    // Neueste XLSX finden
    const files = [];
    for await (const [name, handle] of dirHandle.entries()) {
      if (handle.kind !== 'file' || !/\.xlsx?$/i.test(name)) continue;
      const f = await handle.getFile();
      files.push({ name, file: f, ts: f.lastModified });
    }
    if (!files.length) return;

    files.sort((a, b) => b.ts - a.ts);
    await importAllMonthYearSheets(files[0].file);

    aktualisieren?.();
    afterRenderApply?.();
    console.info('Auto-Import aus gemeinsamem Ordner:', files[0].name);
  } catch (err) {
    console.warn('Auto-Import (shared) fehlgeschlagen:', err);
  }
}
//#endregion



// #region Backup (Excel-Dateien in gemeinsamen Ordner erzeugen & wiederherstellen)

// Wie viele Backups behalten?
const BACKUP_KEEP = 10;

/** Hinweistext rechts neben dem Button aktualisieren */
async function renderBackupFolderHint(dirHandle) {
  const span = document.getElementById('backupFolderHint');
  if (!span) return;
  try {
    span.textContent = dirHandle ? `Backup-Ordner: ${dirHandle.name}` : '';
  } catch {
    span.textContent = '';
  }
}

/** Ordner wählen & merken (in IndexedDB via idbSet) */
async function pickSharedBackupFolder() {
  try {
    const dirHandle = await window.showDirectoryPicker({ mode: 'readwrite' });
    const perm = await dirHandle.requestPermission?.({ mode: 'readwrite' }) ?? 'granted';
    if (perm !== 'granted') {
      alert('Keine Schreib-Berechtigung für den gewählten Ordner.');
      return;
    }
    await idbSet('sharedBackupDir', dirHandle);
    console.info('Backup-Ordner gespeichert:', dirHandle.name);
  } catch (err) {
    if (err?.name !== 'AbortError') {
      console.error('Ordnerwahl fehlgeschlagen:', err);
      alert('Ordnerwahl fehlgeschlagen: ' + (err?.message || err));
    }
  }
}

/** Manuelles Wiederherstellen: neueste .xlsx aus Ordner importieren */
async function manualRestoreFromSharedFolder() {
  try {
    const dirHandle = await idbGet('sharedBackupDir');
    if (!dirHandle) {
      alert('Kein Backup-Ordner gespeichert. Bitte zuerst „Backup-Ordner wählen“.');
      return;
    }
    const perm = await dirHandle.requestPermission?.({ mode: 'read' }) ?? 'granted';
    if (perm !== 'granted') {
      alert('Keine Leserechte für den Backup-Ordner.');
      return;
    }

    const files = [];
    for await (const [name, handle] of dirHandle.entries()) {
      if (handle.kind !== 'file' || !/\.xlsx?$/i.test(name)) continue;
      const f = await handle.getFile();
      files.push({ name, file: f, ts: f.lastModified });
    }
    if (!files.length) {
      alert('Im Backup-Ordner wurden keine Excel-Dateien gefunden.');
      return;
    }

    files.sort((a, b) => b.ts - a.ts);
    await importAllMonthYearSheets(files[0].file);
    aktualisieren?.();
    afterRenderApply?.();
    alert(`Backup wiederhergestellt: ${files[0].name}`);
  } catch (e) {
    console.warn(e);
    alert('Wiederherstellen fehlgeschlagen: ' + (e?.message || e));
  }
}


async function requestBackupReadAndImport() {
  try {
    await ensureLibs({ needXLSX: true });
    const dirHandle = await idbGet('sharedBackupDir');
    if (!dirHandle) { alert('Bitte zuerst „Backup-Ordner wählen“.'); return; }

    const perm = await dirHandle.requestPermission?.({ mode: 'read' }) ?? 'granted';
    if (perm !== 'granted') { alert('Zugriff nicht erlaubt.'); return; }

    // UI-Hinweis entfernen
    document.getElementById('backupFolderHint')?.classList.remove('need-permission');

    const ok = await autoImportLatestBackup();
    if (!ok) alert('Kein passendes Backup gefunden.');
    else { aktualisieren(); afterRenderApply(); }
  } catch (e) {
    console.error('Permission/Import fehlgeschlagen:', e);
    alert('Fehler beim Import: ' + (e?.message || e));
  }
}



/** Beim Start einmal den Hinweistext setzen, falls bereits ein Ordner gespeichert ist */
async function initBackupUI() {
  const handle = await idbGet('sharedBackupDir');
  await renderBackupFolderHint(handle || null);
}

/** Hilfsfunktion: Blob in Ordner schreiben (überschreiben) */
async function writeBlobToDir(dirHandle, fileName, blob) {
  const fileHandle = await dirHandle.getFileHandle(fileName, { create: true });
  const writable = await fileHandle.createWritable();
  await writable.write(blob);
  await writable.close();
  return fileHandle;
}

/** Alte Backups nach Prefix/Endung drehen (letzten N behalten) */
async function rotateBackups(dirHandle, { prefix = '', suffix = /\.xlsx?$/i, keep = BACKUP_KEEP }) {
  const files = [];
  for await (const [name, handle] of dirHandle.entries()) {
    if (handle.kind !== 'file') continue;
    if (!(suffix instanceof RegExp ? suffix.test(name) : String(name).endsWith(suffix))) continue;
    if (prefix && !name.startsWith(prefix)) continue;
    const f = await handle.getFile();
    files.push({ name, handle, ts: f.lastModified });
  }
  files.sort((a, b) => b.ts - a.ts);
  for (let i = keep; i < files.length; i++) {
    try { await dirHandle.removeEntry(files[i].name); } catch { }
  }
}

/** Timestamp für Dateinamen */
function tsStamp(d = new Date()) {
  const pad = n => String(n).padStart(2, '0');
  return `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}-${pad(d.getHours())}${pad(d.getMinutes())}${pad(d.getSeconds())}`;
}

/** Jahres-Excel als Blob bauen (DOM-Mirror, kein Download) */
async function buildExcelBlobForMonth(year, monthIndex) {
  const wb = await buildWorkbookForMonthFromDOM(year, monthIndex);
  const buf = await wb.xlsx.writeBuffer();
  return new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
}

async function buildExcelBlobForYear(year) {
  const wb = await buildWorkbookForYearFromDOM(year);
  const buf = await wb.xlsx.writeBuffer();
  return new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
}



// #endregion



// #region Backup schreiben (Jahr & Monat)
// --- Persistenz für Monatsdaten ---------------------------------------------

function persistCurrentMonthToStore() {
  const jahr = parseInt(document.getElementById('jahr').value, 10);
  const monat = parseInt(document.getElementById('monat').value, 10);
  if (!Number.isFinite(jahr) || !Number.isFinite(monat) || monat === -1) return;

  // rowsPerDay aus dem DOM (Rowspan der ersten Datumzelle)
  let rowsPerDay = FAHRZEUGE.length || 1;
  const firstRS = document.querySelector('#monatsTabelle tbody tr td[rowspan]');
  if (firstRS) {
    const rs = parseInt(firstRS.getAttribute('rowspan') || '0', 10);
    if (rs > 0) rowsPerDay = rs;
  }

  const rows = snapshotTable(); // gibt [fahrzeug, dienstposten, ort, von, bis, bemerkung, eingetragen]
  setMonthData(jahr, monat, rowsPerDay, rows);
}

function persistStoreToLS() {
  const obj = {};
  STORE.byMonth.forEach((v, k) => { obj[k] = v; });
  lsSetJson('monthStoreV1', obj);
}

function restoreStoreFromLS() {
  const obj = lsGetJson('monthStoreV1', null);
  if (!obj) return;
  STORE.byMonth = new Map(Object.entries(obj));
}

// für „Tippen speichert automatisch“
const persistDebounced = debounce(() => {
  persistCurrentMonthToStore();
  persistStoreToLS();
}, 400);



async function backupYearToSharedFolder(year) {
  try {
    const dirHandle = await idbGet('sharedBackupDir');
    if (!dirHandle) throw new Error('Kein Backup-Ordner gesetzt. Bitte zuerst „Backup-Ordner wählen“.');

    const perm = await dirHandle.requestPermission?.({ mode: 'readwrite' }) ?? 'granted';
    if (perm !== 'granted') throw new Error('Keine Schreib-Berechtigung für den Backup-Ordner.');

    const blob = await buildExcelBlobForYear(year);
    const fileName = `Monatstabelle-AlleMonate-${year}-${tsStamp()}.xlsx`;
    await writeBlobToDir(dirHandle, fileName, blob);
    await rotateBackups(dirHandle, { prefix: `Monatstabelle-AlleMonate-${year}-`, suffix: /\.xlsx?$/i, keep: BACKUP_KEEP });
  } catch (err) {
    console.error('Backup (Jahr) fehlgeschlagen:', err);
    throw err;
  }
}

async function backupMonthToSharedFolder(year, monthIndex) {
  try {
    const dirHandle = await idbGet('sharedBackupDir');
    if (!dirHandle) throw new Error('Kein Backup-Ordner gesetzt. Bitte zuerst „Backup-Ordner wählen“.');

    const perm = await dirHandle.requestPermission?.({ mode: 'readwrite' }) ?? 'granted';
    if (perm !== 'granted') throw new Error('Keine Schreib-Berechtigung für den Backup-Ordner.');

    const blob = await buildExcelBlobForMonth(year, monthIndex);
    const mm = String(monthIndex + 1).padStart(2, '0');
    const fileName = `Monatstabelle-${year}-${mm}-${tsStamp()}.xlsx`;
    await writeBlobToDir(dirHandle, fileName, blob);
    await rotateBackups(dirHandle, { prefix: `Monatstabelle-${year}-${mm}-`, suffix: /\.xlsx?$/i, keep: BACKUP_KEEP });
  } catch (err) {
    console.error('Backup (Monat) fehlgeschlagen:', err);
    throw err;
  }
}

// #endregion

// #region Fix: Restore-Button auf manuellen Restore umverdrahten

document.getElementById('restoreBackupBtn')
  ?.addEventListener('click', requestBackupReadAndImport);

// #endregion


// === Supabase: HTML des <tbody> pro (year, month) laden/speichern =========

// Hilfsfunktion: aktuelle Auswahl aus den Selects
function sbGetYearMonth() {
  const y = parseInt(document.getElementById('jahr')?.value ?? new Date().getFullYear(), 10);
  const m = parseInt(document.getElementById('monat')?.value ?? new Date().getMonth(), 10);
  return { year: y, month: m };
}

// Laden: überschreibt das aktuelle <tbody> mit dem gespeicherten HTML
async function sbLoadTbody(year, month) {
  try {
    if (!supabase || month === -1) return; // -1 = "Alle Monate" -> überspringen
    const { data, error } = await supabase
      .from("monthly_html")
      .select("html")
      .eq("year", year)
      .eq("month", month)
      .maybeSingle();

    if (error) { console.warn("Supabase SELECT Fehler:", error); return; }
    if (data?.html) {
      const tb = document.querySelector("#monatsTabelle tbody");
      if (!tb) return;
      tb.innerHTML = data.html;

      // Falls dein Code danach noch Events re-binden muss:
      if (window.afterTbodyRestore) window.afterTbodyRestore();

      // Markierungen/Save-Bar zurücksetzen, damit „alles sauber“ ist
      clearChangeMarkersAndRebase?.();
      setUnsaved?.(false);
    }
  } catch (e) {
    console.warn("sbLoadTbody Fehler:", e);
  }
}

// Speichern: aktuelles <tbody> hochladen (Upsert je year+month)
async function sbSaveTbody(year, month) {
  if (!supabase || month === -1) return;
  const tb = document.querySelector("#monatsTabelle tbody");
  if (!tb) return;
  const html = tb.innerHTML;

  const { error } = await supabase
    .from("monthly_html")
    .upsert({ year, month, html })
    .select();
  if (error) throw error;
}

// (Optional) Realtime-Updates (wenn Seite in mehreren Fenstern offen ist)
function sbSubscribeRealtime() {
  if (!supabase) return;
  supabase
    .channel("public:monthly_html")
    .on("postgres_changes", { event: "INSERT", schema: "public", table: "monthly_html" }, payload => {
      const { year, month } = sbGetYearMonth();
      if (payload.new?.year === year && payload.new?.month === month) {
        const tb = document.querySelector("#monatsTabelle tbody");
        if (tb) {
          tb.innerHTML = payload.new.html;
          window.afterTbodyRestore?.();
          clearChangeMarkersAndRebase?.();
          setUnsaved?.(false);
        }
      }
    })
    .on("postgres_changes", { event: "UPDATE", schema: "public", table: "monthly_html" }, payload => {
      const { year, month } = sbGetYearMonth();
      if (payload.new?.year === year && payload.new?.month === month) {
        const tb = document.querySelector("#monatsTabelle tbody");
        if (tb) {
          tb.innerHTML = payload.new.html;
          window.afterTbodyRestore?.();
          clearChangeMarkersAndRebase?.();
          setUnsaved?.(false);
        }
      }
    })
    .subscribe();
}





// #region Init/Wiring (vollständig)

(function init() {
  async function refreshBackupHint() {
    const hint = document.getElementById('backupFolderHint');
    if (!hint) return;
    try {
      const dirHandle = await idbGet('sharedBackupDir');
      if (!dirHandle) {
        hint.textContent = '(kein Backup-Ordner gewählt)';
        return;
      }
      // Permission-Status prüfen
      let perm = 'prompt';
      if (dirHandle.queryPermission) {
        perm = await dirHandle.queryPermission({ mode: 'read' });
      }
      if (perm !== 'granted' && dirHandle.requestPermission) {
        // nicht anstoßen – nur Hinweis anzeigen
        hint.textContent = `Ordner gesetzt (${dirHandle.name}) – Zugriff noch nicht bestätigt`;
      } else {
        hint.textContent = `Backup-Ordner: ${dirHandle.name}`;
      }
    } catch {
      // still
    }
  }

  (async function start() {
    // 1) Grund-UI aufsetzen (Selects etc.)
    baueGrundAuswahl();

    // 2) (optional) persistente Origin-Quota anfragen
    if (navigator.storage && navigator.storage.persist) {
      try { await navigator.storage.persist(); } catch { }
    }

    // 3) Referenzen & Auswahl-Listener
    const monatSel = document.getElementById('monat');
    const jahrSel = document.getElementById('jahr');

    const onPickChange = () => {
      // vor Rendern aktuelle Eingaben in STORE übernehmen
      persistCurrentMonthToStore?.();
      persistStoreToLS?.();

      aktualisieren();
      afterRenderApply();
    };
    monatSel?.addEventListener('change', onPickChange);
    jahrSel?.addEventListener('change', onPickChange);
    jahrSel?.addEventListener('input', onPickChange);

    // 4) Lokale Sicherung (LS) laden – falls vorhanden
    restoreStoreFromLS?.();

    // 5) Versuche vor dem ersten Render die neueste Backup-Excel still zu importieren
    //    (funktioniert nur, wenn der Ordner schon gewählt & Berechtigung ok ist)
    try {
      await ensureLibs({ needXLSX: true }); 
      const did = await autoImportOnLoad?.({ onlyIfStoreEmpty: false });
      if (did) {
        persistStoreToLS?.();
      }
    } catch (e) {
      console.warn('Auto-Import beim Start übersprungen:', e);
    }

    // 6) Erstes Rendern (nach LS/Backup-Import)
    aktualisieren();
    afterRenderApply();

    // 7) Backup-Hinweis initial anzeigen
    await refreshBackupHint?.();

    // 8) Buttons verdrahten

    // 8a) Import (versteckter <input type="file"> bei Bedarf erstellen)
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
        // nach Import direkt rendern & LS aktualisieren
        aktualisieren();
        afterRenderApply();
        // Realtime-Listener starten
        sbSubscribeRealtime();
        persistStoreToLS?.();
      } catch (err) {
        console.error(err);
        alert('Import fehlgeschlagen: ' + (err?.message || err));
      } finally {
        e.target.value = '';
      }
    });

    // 8b) Export (Monat oder Jahr)
    document.getElementById('exportXlsxBtn')?.addEventListener('click', async () => {
      const curMonth = parseInt(monatSel?.value ?? '0', 10);
      // aktuelle Eingaben sichern
      persistCurrentMonthToStore?.();
      persistStoreToLS?.();

      if (curMonth === -1) {
        await exportYearWithExcelJS_DOMMirror();
      } else {
        await exportMitExcelJS();
      }
    });

    // 8c) Backup-Ordner wählen
    document.getElementById('chooseBackupFolderBtn')?.addEventListener('click', async () => {
      await pickSharedBackupFolder();
      await refreshBackupHint?.();
    });

    // 8e) Speichern-Button unten: Export + optionales Backup in Ordner
    const saveBtn = document.getElementById('saveBtn');
    saveBtn?.addEventListener('click', async (e) => {
      e.preventDefault();
      e.stopPropagation();

      const y = parseInt(document.getElementById('jahr')?.value ?? '0', 10) || (getTodayDate()).getFullYear();
      const m = parseInt(document.getElementById('monat')?.value ?? '0', 10) || 0;

      // aktuelle Eingaben vor dem Export sichern
      persistCurrentMonthToStore?.();
      persistStoreToLS?.();

      // --- Supabase: aktuelles <tbody> hochladen --------------------------------
try {
  await sbSaveTbody(y, m);
} catch (err) {
  console.error('Supabase Save Fehler:', err);
  alert('Speichern in der Cloud fehlgeschlagen: ' + (err?.message || err));
}


      try {
        if (m === -1) {
          await backupYearToSharedFolder(y);
        } else {
          await backupMonthToSharedFolder(y, m);
        }
        setUnsaved(false);
        clearChangeMarkersAndRebase();
        document.activeElement?.blur?.();
      } catch (err) {
        console.error('Speichern/Export fehlgeschlagen:', err);
        alert('Speichern/Export fehlgeschlagen: ' + (err?.message || err));
      }
    });

    // 9) Lock/Unlock
    document.getElementById('unlockBtn')?.addEventListener('click', entsperren);
    document.getElementById('lockBtn')?.addEventListener('click', sperren);
  })();

  // 7) sonstiges UI
  const lightbox = document.getElementById('imgLightbox');
  if (lightbox) {
    lightbox.addEventListener('click', () => (lightbox.style.display = 'none'));
    window.addEventListener('keydown', (e) => { if (e.key === 'Escape') lightbox.style.display = 'none'; });
  }

  // Back-to-top Button (du hast #backToTop in deiner Button-Leiste)
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

  // Reflow Right Rail bei Resize
  window.addEventListener('resize', reflowRightRail);
})();
















// #endregion





