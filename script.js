// Polling and sheet configuration
const FETCH_INTERVAL_MS = 60_000;
const FETCH_TIMEOUT_MS = 25_000;
const SCRIPT_STALE_MS = 15 * 60_000;
const AGE_HINT_DAYS = 3;
const SHEET_ID = '10mfm9SVVDiWcxnfK2QuUCj3msaVFBQIQx34NnPlUEo4';
const SHEET_URL = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?tqx=out:json`;
const IMPORT_LOG_URL = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?tqx=out:json&sheet=Statistik&range=A2:C`;
const IMPORT_RUN_URL = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?tqx=out:json&sheet=Statistik&range=F1`;
const TAGES_STAT_URL = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?tqx=out:json&sheet=Statistik&range=H2:J`;

// Board configuration
const columns = ['Eingang', 'Hadi', 'Ramazan', 'Robar', 'Osama', 'Geprüft'];
const workerColumnAliases = new Map([
  ['hadi', 'Hadi'],
  ['hadi issa', 'Hadi'],
  ['ramazan', 'Ramazan'],
  ['ramazan dag', 'Ramazan'],
  ['robar', 'Robar'],
  ['robar kassem', 'Robar'],
  ['osama', 'Osama'],
  ['osama sleiman', 'Osama'],
  ['osama souleiman', 'Osama']
]);
const externalWorkerAliases = new Map([
  ['h', { cls: 'hj', label: 'H' }],
  ['hj', { cls: 'hj', label: 'H' }],
  ['hussein jaber', { cls: 'hj', label: 'H' }],
  ['b', { cls: 'hussein', label: 'B' }],
  ['hussein selman', { cls: 'hussein', label: 'B' }],
  ['hu', { cls: 'hu', label: 'HU' }],
  ['hussein souleiman', { cls: 'hu', label: 'HU' }]
]);

let lastFetchTime = null;
let nextFetchTime = null;
let inFlightController = null;
let lastFetchError = '';
let openCountEl = null;
let statsInFlightController = null;
let importDateByAkte = new Map();
let lastImportRunTime = null;
let lastBoardData = [];
let previousCardPositions = new Map();
let knownAkten = new Set();
let isFirstBoardRender = true;

function todayDateKey(date = new Date()) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

function parseGvizRows(text) {
  const json = parseGviz(text);
  return Array.isArray(json?.table?.rows) ? json.table.rows : [];
}

function normalizeSyncLabel(raw) {
  const label = String(raw ?? '').trim();
  if (['OK', 'Vollständig', 'Komplett'].includes(label)) return 'OK';
  if (['Offen', 'Fehlt im Sheet', 'Fehlt', 'Missing', 'Gap', 'NO'].includes(label)) return 'NO';
  return label || '–';
}

function isSyncOkLabel(label) {
  return label === 'OK';
}

function setUploadStats({ syncLabel, rbCount }) {
  const syncEl = document.getElementById('uploadSyncValue');
  const syncCell = document.getElementById('syncStatCell');
  const displayLabel = normalizeSyncLabel(syncLabel);

  if (syncEl) syncEl.textContent = displayLabel;

  if (syncCell) {
    syncCell.classList.remove('sync-ok', 'sync-open');
    if (isSyncOkLabel(displayLabel)) syncCell.classList.add('sync-ok');
    else if (displayLabel === 'NO') syncCell.classList.add('sync-open');
  }

  setRbCount(rbCount);
}

function setRbCount(count) {
  const widget = document.getElementById('rbCountWidget');
  const valueEl = document.getElementById('rbCountValue');
  if (!widget || !valueEl) return;

  const n = Number.isFinite(Number(count)) ? Number(count) : null;
  valueEl.textContent = n === null ? '–' : String(n);

  widget.classList.remove('state-green', 'state-yellow', 'state-orange', 'state-burn');
  if (n === null) widget.classList.add('state-green');
  else if (n >= 15) widget.classList.add('state-burn');
  else if (n >= 10) widget.classList.add('state-orange');
  else if (n >= 5) widget.classList.add('state-yellow');
  else widget.classList.add('state-green');

  widget.title = n === null
    ? 'Reparaturbestätigungen im Drive-Ordner'
    : `Offene Reparaturbestätigungen: ${n}`;
}

function parseTagesStatRow(row) {
  const cell = (idx) => String(row.c?.[idx]?.v ?? '').trim();

  // Neues Format: Datum | Drive-Abgleich | RB_Offene
  if (['OK', 'Offen', 'Vollständig', 'Fehlt im Sheet', 'Komplett', 'Fehlt', 'Missing', 'Gap', 'NO'].includes(cell(1))) {
    const rb = Number.parseInt(cell(2), 10);
    return { syncLabel: cell(1), rbCount: Number.isFinite(rb) ? rb : null };
  }

  // Altes 6-Spalten-Format: Datum | Neu | Offen | Drive | Sync | RB
  if (['OK', 'Offen', 'Vollständig', 'Fehlt im Sheet', 'Komplett', 'Fehlt', 'Missing', 'Gap', 'NO'].includes(cell(4))) {
    const rb = Number.parseInt(cell(5), 10);
    return { syncLabel: cell(4), rbCount: Number.isFinite(rb) ? rb : null };
  }

  // Altes 5-Spalten-Format: Datum | Neu | Offen | Sync | RB
  if (['OK', 'Offen', 'Vollständig', 'Fehlt im Sheet', 'Komplett', 'Fehlt', 'Missing', 'Gap', 'NO'].includes(cell(3))) {
    const rb = Number.parseInt(cell(4), 10);
    return { syncLabel: cell(3), rbCount: Number.isFinite(rb) ? rb : null };
  }

  return { syncLabel: '–', rbCount: null };
}

function computeImportStats(tagesStatRows) {
  const today = todayDateKey();

  let syncLabel = '–';
  let rbCount = null;
  for (let i = tagesStatRows.length - 1; i >= 0; i -= 1) {
    const date = String(tagesStatRows[i].c?.[0]?.v ?? '').trim();
    if (!/^\d{4}-\d{2}-\d{2}$/.test(date) || date !== today) continue;
    const parsed = parseTagesStatRow(tagesStatRows[i]);
    syncLabel = parsed.syncLabel;
    rbCount = parsed.rbCount;
    break;
  }

  return { syncLabel, rbCount };
}

async function fetchImportStats() {
  if (statsInFlightController) return;

  statsInFlightController = new AbortController();
  const timeoutId = window.setTimeout(() => {
    if (statsInFlightController) statsInFlightController.abort();
  }, FETCH_TIMEOUT_MS);

  try {
    const [logRes, statRes, runRes] = await Promise.all([
      fetch(IMPORT_LOG_URL, { signal: statsInFlightController.signal, cache: 'no-store' }),
      fetch(TAGES_STAT_URL, { signal: statsInFlightController.signal, cache: 'no-store' }),
      fetch(IMPORT_RUN_URL, { signal: statsInFlightController.signal, cache: 'no-store' })
    ]);

    if (!logRes.ok && !statRes.ok && !runRes.ok) return;

    const logRows = logRes.ok ? parseGvizRows(await logRes.text()) : [];
    const tagesStatRows = statRes.ok ? parseGvizRows(await statRes.text()) : [];
    const { dateByAkte } = parseImportLogData(logRows);
    importDateByAkte = dateByAkte;

    if (runRes.ok) {
      const runJson = parseGviz(await runRes.text());
      const runValue = runJson?.table?.rows?.[0]?.c?.[0]?.v;
      lastImportRunTime = parseImportRunTime(runValue);
    }
    setUploadStats(computeImportStats(tagesStatRows));
    if (lastBoardData.length) renderBoard(lastBoardData);
    updateTimerDisplay();
  } catch (err) {
    if (err.name !== 'AbortError') {
      console.error('fetchImportStats() Fehler:', err);
    }
  } finally {
    window.clearTimeout(timeoutId);
    statsInFlightController = null;
  }
}

function extractAktenNummer(value) {
  const match = String(value || '').match(/\d+/);
  return match ? match[0] : '';
}

function parseImportRunTime(text) {
  const value = String(text || '').trim();
  if (!value) return null;

  const dt = new Date(value.includes('T') ? value : value.replace(' ', 'T'));
  return Number.isNaN(dt.getTime()) ? null : dt;
}

function parseImportLogData(logRows) {
  const dateByAkte = new Map();

  logRows.forEach((row) => {
    const date = String(row.c?.[0]?.v ?? '').trim();
    const nummer = String(row.c?.[2]?.v ?? '').replace(/[^0-9]/g, '');
    if (!/^\d{4}-\d{2}-\d{2}$/.test(date) || !nummer) return;

    const existing = dateByAkte.get(nummer);
    if (!existing || date < existing) dateByAkte.set(nummer, date);
  });

  return { dateByAkte };
}

function getAgeDays(importDateStr) {
  if (!importDateStr) return null;
  const imported = new Date(`${importDateStr}T12:00:00`);
  const today = new Date();
  today.setHours(12, 0, 0, 0);
  return Math.floor((today - imported) / 86_400_000);
}

function formatScriptStaleMessage(date) {
  const diffMs = Date.now() - date.getTime();
  const minutes = Math.floor(diffMs / 60_000);
  if (minutes < 60) return `⚠ Skript seit ${minutes} Min nicht gelaufen`;
  const hours = Math.floor(minutes / 60);
  if (hours < 24) return `⚠ Skript seit ${hours} Std nicht gelaufen`;
  const days = Math.floor(hours / 24);
  return `⚠ Skript seit ${days} Tag${days === 1 ? '' : 'en'} nicht gelaufen`;
}

function isScriptStale(date) {
  if (!date) return false;
  return Date.now() - date.getTime() > SCRIPT_STALE_MS;
}

function isUnknownStatus(status) {
  const normalized = String(status || '').trim().toLowerCase();
  return normalized === 'unbekannt' || normalized === 'fehler beim auslesen';
}

function makeEmptyMap() {
  return columns.reduce((map, col) => {
    map[col] = [];
    return map;
  }, {});
}

function escapeHtml(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function setTickerText(text) {
  const ticker = document.querySelector('.ticker');
  if (!ticker) return;

  ticker.textContent = '';
  const span = document.createElement('span');
  span.className = 'ticker-span';
  span.textContent = text;
  ticker.appendChild(span);
  startTickerAnimation();
}

function startTickerAnimation() {
  const spans = document.querySelectorAll('.ticker .ticker-span');
  spans.forEach((span) => {
    span.style.animation = 'ticker-scroll 25s linear infinite';
  });
}

function ensureOpenCountWidget() {
  const existing = document.getElementById('openCountWidget');
  if (existing) {
    openCountEl = existing;
    return existing;
  }
  return null;
}

function setOpenCount(count) {
  const widget = ensureOpenCountWidget();
  if (!widget) return;

  const valueEl = widget.querySelector('.value');
  if (valueEl) valueEl.textContent = String(count ?? 0);

  widget.classList.remove('state-green', 'state-yellow', 'state-orange', 'state-burn');
  const n = Number(count) || 0;
  if (n >= 30) widget.classList.add('state-burn');
  else if (n >= 20) widget.classList.add('state-orange');
  else if (n >= 10) widget.classList.add('state-yellow');
  else widget.classList.add('state-green');

  widget.title = `Offene Akten: ${n}`;
}

function extractAktenzeichen(text) {
  if (!text) return '';
  const match = text.match(/^(RB|KVA)?\s*\d+/i);
  return match ? match[0].toUpperCase().replace(/\s+/, ' ') : String(text).trim();
}

function parseGviz(text) {
  const start = text.indexOf('{');
  const end = text.lastIndexOf('}');
  if (start === -1 || end === -1) throw new Error('GViz JSON nicht gefunden');
  return JSON.parse(text.slice(start, end + 1));
}

function isGeprueftStatus(status) {
  return /^geprüft\s*(o|1|2|hj|hk)$/i.test(status);
}

function normalizeWorkerName(rawValue) {
  return String(rawValue || '')
    .replace(/\s+/g, ' ')
    .trim()
    .replace(/^herr\s+/i, '')
    .replace(/^frau\s+/i, '')
    .trim()
    .toLowerCase();
}

function resolveWorkerColumn(rawWorker) {
  const normalized = normalizeWorkerName(rawWorker);
  if (!normalized) return null;
  if (workerColumnAliases.has(normalized)) return workerColumnAliases.get(normalized);

  const firstName = normalized.split(' ')[0];
  if (workerColumnAliases.has(firstName)) return workerColumnAliases.get(firstName);
  return null;
}

function resolveExternalBadge(rawWorker) {
  const normalized = normalizeWorkerName(rawWorker);
  if (!normalized) return null;
  if (externalWorkerAliases.has(normalized)) return externalWorkerAliases.get(normalized);

  // Robust Hussein mapping by surname to avoid exact-string dependency
  if (normalized.includes('hussein')) {
    if (/\bselman\b/.test(normalized)) return externalWorkerAliases.get('b');
    if (/\b(souleiman|suleiman|sleiman)\b/.test(normalized)) return externalWorkerAliases.get('hu');
    if (/\bjaber\b/.test(normalized)) return externalWorkerAliases.get('h');
  }

  const firstName = normalized.split(' ')[0];
  if (externalWorkerAliases.has(firstName)) return externalWorkerAliases.get(firstName);
  return null;
}

function scheduleNextFetch() {
  nextFetchTime = new Date(Date.now() + FETCH_INTERVAL_MS);
}

async function fetchData() {
  if (inFlightController) return;

  inFlightController = new AbortController();
  let didTimeout = false;
  const timeoutId = window.setTimeout(() => {
    didTimeout = true;
    if (inFlightController) inFlightController.abort();
  }, FETCH_TIMEOUT_MS);

  try {
    const res = await fetch(SHEET_URL, {
      signal: inFlightController.signal,
      cache: 'no-store'
    });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);

    const raw = await res.text();
    const json = parseGviz(raw);
    const tableRows = Array.isArray(json?.table?.rows) ? json.table.rows : [];

    const rows = tableRows.map((row) => {
      const eingangRaw = row.c?.[0]?.v ?? '';
      const eingang = extractAktenzeichen(eingangRaw);
      const bearbeiter = String(row.c?.[1]?.v ?? '').trim();
      const status = String(row.c?.[2]?.v ?? '').trim().toLowerCase();
      return { Eingang: eingang, Bearbeiter: bearbeiter, Status: status };
    }).filter((row) => row.Eingang && row.Eingang.toLowerCase() !== 'eingang');

    const cleanedRows = rows.filter((row) => !/^versendet\b/i.test(row.Status));
    setOpenCount(cleanedRows.length);

    const numbers = cleanedRows
      .map((row) => (row.Eingang.match(/\d+/) ?? [null])[0])
      .filter((num) => num !== null)
      .map((num) => Number.parseInt(num, 10))
      .filter((num) => Number.isFinite(num));

    const nextNumber = numbers.length ? Math.max(...numbers) + 1 : '–';
    setTickerText(`💥 Aktuelle Nummer: ${nextNumber} 🚗`);

    renderBoard(cleanedRows);
    lastBoardData = cleanedRows;
    lastFetchTime = new Date();
    lastFetchError = '';
    await fetchImportStats();
  } catch (err) {
    if (err.name === 'AbortError' && didTimeout) {
      lastFetchError = `Timeout nach ${Math.round(FETCH_TIMEOUT_MS / 1000)}s`;
      console.error('fetchData() Timeout:', err);
    } else if (err.name !== 'AbortError') {
      lastFetchError = err.message || 'Unbekannter Fehler';
      console.error('fetchData() Fehler:', err);
    }
  } finally {
    window.clearTimeout(timeoutId);
    inFlightController = null;
    scheduleNextFetch();
    updateTimerDisplay();
  }
}

function updateTimerDisplay() {
  const el = document.getElementById('updateInfo');
  if (!el) return;

  const now = Date.now();
  const lastUpdateText = lastFetchTime
    ? new Intl.DateTimeFormat('de-DE', {
      hour: '2-digit',
      minute: '2-digit',
      second: '2-digit'
    }).format(lastFetchTime)
    : 'unbekannt';

  const diffMs = nextFetchTime ? Math.max(0, nextFetchTime.getTime() - now) : 0;
  const minutesLeft = Math.floor(diffMs / 60_000);
  const secondsLeft = Math.floor((diffMs % 60_000) / 1000);

  el.className = 'stat-cell stat-cell-time';
  if (diffMs > FETCH_INTERVAL_MS * 0.5) el.classList.add('green');
  else if (diffMs > FETCH_INTERVAL_MS * 0.2) el.classList.add('orange');
  else el.classList.add('red');

  let html = `${escapeHtml(lastUpdateText)} · ↻ ${minutesLeft}:${String(secondsLeft).padStart(2, '0')}`;
  if (isScriptStale(lastImportRunTime)) {
    html += ` · ${escapeHtml(formatScriptStaleMessage(lastImportRunTime))}`;
    el.classList.remove('green', 'orange');
    el.classList.add('red');
  }
  if (lastFetchError) {
    html += ` · ⚠ ${escapeHtml(lastFetchError)}`;
    el.classList.remove('green', 'orange');
    el.classList.add('red');
  }
  el.textContent = html;
}

const columnClassMap = {
  Eingang: 'column-eingang',
  Hadi: 'column-hadi',
  Ramazan: 'column-ramazan',
  Robar: 'column-robar',
  Osama: 'column-osama',
  'Geprüft': 'column-gepruft'
};

function applyCardStatus(card, status) {
  card.classList.remove('status-geprueft', 'status-unvollstaendig', 'status-vollstaendig');
  if (isGeprueftStatus(status)) card.classList.add('status-geprueft');
  else if (status.includes('unvollständig')) card.classList.add('status-unvollstaendig');
  else if (status.includes('vollständig')) card.classList.add('status-vollstaendig');
}

function buildBoardMap(data) {
  const map = makeEmptyMap();

  data.forEach((row) => {
    const status = row.Status.toLowerCase().trim();
    const akte = { nummer: row.Eingang, status, bearbeiter: row.Bearbeiter };

    if (isGeprueftStatus(status)) {
      map.Geprüft.push(akte);
      return;
    }

    if (status === 'vollständig') {
      map.Osama.push(akte);
      return;
    }

    const workerColumn = resolveWorkerColumn(row.Bearbeiter);
    if (workerColumn) {
      map[workerColumn].push(akte);
      return;
    }

    map.Eingang.push(akte);
  });

  for (const col of columns) {
    map[col].sort((a, b) => {
      const aNum = Number.parseInt((a.nummer.match(/\d+/) ?? ['0'])[0], 10);
      const bNum = Number.parseInt((b.nummer.match(/\d+/) ?? ['0'])[0], 10);
      return aNum - bNum;
    });
  }

  return map;
}

function applyCardHighlight(card, { nummer, col }) {
  const isNew = !knownAkten.has(nummer);
  const prevCol = previousCardPositions.get(nummer);
  const moved = !isFirstBoardRender && prevCol && prevCol !== col;

  if (isNew && col === 'Eingang') {
    card.classList.add('card-new');
  } else if (moved) {
    card.classList.add('card-moved');
  }

  knownAkten.add(nummer);
}

function renderBoard(data) {
  const board = document.getElementById('board');
  if (!board) return;

  const map = buildBoardMap(data);

  board.textContent = '';
  const fragment = document.createDocumentFragment();

  columns.forEach((col) => {
    const colDiv = document.createElement('div');
    colDiv.className = `column ${columnClassMap[col] || ''}`;

    const count = map[col].length;

    const header = document.createElement('div');
    header.className = 'column-header';

    const title = document.createElement('h2');
    title.textContent = col;

    const countBadge = document.createElement('span');
    countBadge.className = 'column-count';
    countBadge.textContent = String(count);

    header.appendChild(title);
    header.appendChild(countBadge);
    colDiv.appendChild(header);

    const cardsWrap = document.createElement('div');
    cardsWrap.className = 'column-cards';

    map[col].forEach((item) => {
      const { nummer, status, bearbeiter } = item;
      if (nummer.toLowerCase() === col.toLowerCase()) return;

      const card = document.createElement('div');
      card.className = 'card';

      applyCardHighlight(card, { nummer, col });

      const nummerEl = document.createElement('div');
      nummerEl.className = 'card-number';
      nummerEl.textContent = nummer;
      card.appendChild(nummerEl);

      const aktenNummer = extractAktenNummer(nummer);
      const importDate = aktenNummer ? importDateByAkte.get(aktenNummer) : null;
      const ageDays = getAgeDays(importDate);

      if (ageDays !== null && ageDays >= AGE_HINT_DAYS) {
        card.classList.add('card-aged');
        card.title = ageDays === 1 ? 'Seit 1 Tag im System' : `Seit ${ageDays} Tagen im System`;
      }

      if (isUnknownStatus(status)) {
        card.classList.add('card-unknown');
        const unknownBadge = document.createElement('div');
        unknownBadge.className = 'unknown-badge';
        unknownBadge.textContent = '?';
        unknownBadge.title = 'Status unbekannt – UX-Sync prüfen';
        card.appendChild(unknownBadge);
      }

      applyCardStatus(card, status);

      const externalBadge = resolveExternalBadge(bearbeiter);
      if (externalBadge) {
        card.classList.add('extern', externalBadge.cls);
        const badge = document.createElement('div');
        badge.className = 'extern-badge';
        badge.textContent = externalBadge.label;
        card.appendChild(badge);
      }

      cardsWrap.appendChild(card);
    });

    colDiv.appendChild(cardsWrap);
    fragment.appendChild(colDiv);
  });

  board.appendChild(fragment);

  previousCardPositions = new Map();
  columns.forEach((col) => {
    map[col].forEach((item) => previousCardPositions.set(item.nummer, col));
  });
  isFirstBoardRender = false;
}

const THEME_STORAGE_KEY = 'svs-dashboard-theme';

function readStoredTheme() {
  try {
    const saved = localStorage.getItem(THEME_STORAGE_KEY);
    if (saved === 'dark' || saved === 'light') return saved;
  } catch (_) {
    /* localStorage blockiert (Privatmodus / iframe) */
  }
  try {
    const saved = sessionStorage.getItem(THEME_STORAGE_KEY);
    if (saved === 'dark' || saved === 'light') return saved;
  } catch (_) {
    /* sessionStorage blockiert */
  }
  return null;
}

function storeTheme(theme) {
  try {
    localStorage.setItem(THEME_STORAGE_KEY, theme);
    return;
  } catch (_) {
    /* Fallback wenn localStorage nicht verfügbar */
  }
  try {
    sessionStorage.setItem(THEME_STORAGE_KEY, theme);
  } catch (_) {
    /* Speichern optional */
  }
}

function applyTheme(theme) {
  const isDark = theme === 'dark';
  document.documentElement.setAttribute('data-theme', isDark ? 'dark' : 'light');
  document.documentElement.style.colorScheme = isDark ? 'dark' : 'light';

  const meta = document.getElementById('themeColorMeta');
  if (meta) meta.content = isDark ? '#0b0f14' : '#eef2f7';

  const btn = document.getElementById('themeToggle');
  if (!btn) return;

  const icon = btn.querySelector('.theme-toggle-icon');
  if (icon) icon.textContent = isDark ? '☀️' : '🌙';
  btn.title = isDark ? 'Helles Design' : 'Dunkles Design';
  btn.setAttribute('aria-label', isDark ? 'Helles Design aktivieren' : 'Dunkles Design aktivieren');
}

function initTheme() {
  const stored = readStoredTheme();
  const current = document.documentElement.getAttribute('data-theme');
  applyTheme(stored || (current === 'dark' ? 'dark' : 'light'));
}

function toggleTheme() {
  const isDark = document.documentElement.getAttribute('data-theme') === 'dark';
  const next = isDark ? 'light' : 'dark';
  storeTheme(next);
  applyTheme(next);
}

function bindThemeToggle() {
  const themeBtn = document.getElementById('themeToggle');
  if (!themeBtn || themeBtn.dataset.bound === '1') return;
  themeBtn.dataset.bound = '1';

  themeBtn.addEventListener('click', (event) => {
    event.preventDefault();
    toggleTheme();
  });
}

function isKioskMode() {
  return document.documentElement.getAttribute('data-kiosk') === '1';
}

function initKioskMode() {
  if (!isKioskMode()) return;

  document.body.classList.add('kiosk');
  document.title = 'SVS Dashboard';

  const themeBtn = document.getElementById('themeToggle');
  if (themeBtn) themeBtn.hidden = true;

  document.addEventListener('contextmenu', (e) => e.preventDefault());

  let cursorTimer;
  const hideCursor = () => document.body.classList.add('kiosk-idle');
  const showCursor = () => {
    document.body.classList.remove('kiosk-idle');
    window.clearTimeout(cursorTimer);
    cursorTimer = window.setTimeout(hideCursor, 4000);
  };
  document.addEventListener('mousemove', showCursor, { passive: true });
  cursorTimer = window.setTimeout(hideCursor, 4000);

  tryEnterFullscreen();
  requestWakeLock();
}

async function tryEnterFullscreen() {
  try {
    if (!document.fullscreenElement && document.documentElement.requestFullscreen) {
      await document.documentElement.requestFullscreen();
    }
  } catch (_) {
    /* Browser blockiert Vollbild ohne Nutzeraktion — auf dem Pi mit Chromium --kiosk ok */
  }
}

async function requestWakeLock() {
  try {
    if ('wakeLock' in navigator) {
      await navigator.wakeLock.request('screen');
    }
  } catch (_) {
    /* Optional — nicht überall verfügbar */
  }
}

window.addEventListener('DOMContentLoaded', () => {
  initKioskMode();
  initTheme();
  bindThemeToggle();

  ensureOpenCountWidget();
  startTickerAnimation();
  scheduleNextFetch();
  fetchData();
  setInterval(fetchData, FETCH_INTERVAL_MS);
  setInterval(updateTimerDisplay, 1000);
});
