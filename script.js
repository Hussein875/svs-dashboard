// Polling and sheet configuration
const FETCH_INTERVAL_MS = 60_000;
const FETCH_TIMEOUT_MS = 25_000;
const SHEET_ID = '10mfm9SVVDiWcxnfK2QuUCj3msaVFBQIQx34NnPlUEo4';
const SHEET_URL = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?tqx=out:json`;

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
  if (openCountEl) return openCountEl;

  const allWidgets = document.querySelectorAll('#openCountWidget');
  if (allWidgets.length > 1) {
    const keep = Array.from(allWidgets).find((el) => el.closest('.bottom-bar')) || allWidgets[0];
    allWidgets.forEach((el) => {
      if (el !== keep) el.remove();
    });
  }

  const existingWidgets = document.querySelectorAll('#openCountWidget');
  const existing = Array.from(existingWidgets).find((el) => el.closest('.bottom-bar')) || existingWidgets[0];
  if (existing) {
    openCountEl = existing;
    return openCountEl;
  }

  const widget = document.createElement('div');
  widget.id = 'openCountWidget';
  widget.className = 'state-green';
  widget.setAttribute('aria-label', 'Offene Akten');

  const chip = document.createElement('div');
  chip.className = 'chip';
  const label = document.createElement('div');
  label.className = 'label';
  label.textContent = 'Offene Akten';
  const value = document.createElement('div');
  value.className = 'value';
  value.textContent = '0';

  widget.appendChild(chip);
  widget.appendChild(label);
  widget.appendChild(value);

  const bottomBar = document.querySelector('.bottom-bar');
  const updateInfo = document.getElementById('updateInfo');
  if (bottomBar && updateInfo) bottomBar.insertBefore(widget, updateInfo);
  else if (bottomBar) bottomBar.appendChild(widget);
  else document.body.appendChild(widget);

  openCountEl = widget;
  return widget;
}

function setOpenCount(count) {
  const widget = ensureOpenCountWidget();
  const valueEl = widget.querySelector('.value');
  if (valueEl) valueEl.textContent = String(count ?? 0);

  widget.classList.remove('state-green', 'state-yellow', 'state-orange', 'state-red', 'state-burn');
  const n = Number(count) || 0;
  if (n >= 30) widget.classList.add('state-burn');
  else if (n >= 20) widget.classList.add('state-orange');
  else if (n >= 10) widget.classList.add('state-yellow');
  else widget.classList.add('state-green');

  widget.title = `Offene / nicht versendete Akten: ${n}`;
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
    lastFetchTime = new Date();
    lastFetchError = '';
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

  el.className = 'update-info';
  if (diffMs > FETCH_INTERVAL_MS * 0.5) el.classList.add('green');
  else if (diffMs > FETCH_INTERVAL_MS * 0.2) el.classList.add('orange');
  else el.classList.add('red');

  let html = `⏱️ <b>Letztes Update:</b> ${escapeHtml(lastUpdateText)} &nbsp;|&nbsp; 🔄 <b>Nächstes in:</b> ${minutesLeft}m ${secondsLeft}s`;
  if (lastFetchError) {
    html += ` &nbsp;|&nbsp; ⚠️ <b>Fehler:</b> ${escapeHtml(lastFetchError)}`;
    el.classList.remove('green');
    el.classList.add('red');
  }
  el.innerHTML = html;
}

function renderBoard(data) {
  const board = document.getElementById('board');
  if (!board) return;

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

  board.textContent = '';
  const fragment = document.createDocumentFragment();

  columns.forEach((col) => {
    const colDiv = document.createElement('div');
    colDiv.className = 'column';

    const title = document.createElement('h2');
    const count = map[col].length;
    title.textContent = count > 0 ? `${col} (${count})` : col;
    colDiv.appendChild(title);

    map[col].forEach((item) => {
      const { nummer, status, bearbeiter } = item;
      if (nummer.toLowerCase() === col.toLowerCase()) return;

      const card = document.createElement('div');
      card.className = 'card';
      card.textContent = nummer;

      if (isGeprueftStatus(status)) {
        card.style.backgroundColor = '#13e339ff';
        card.style.color = 'white';
      } else if (status.includes('unvollständig')) {
        card.style.backgroundColor = '#e31313ff';
        card.style.color = 'white';
      } else if (status.includes('vollständig')) {
        card.style.backgroundColor = '#ff7300ff';
        card.style.color = 'white';
      }

      const externalBadge = resolveExternalBadge(bearbeiter);
      if (externalBadge) {
        card.classList.add('extern', externalBadge.cls);
        const badge = document.createElement('div');
        badge.className = 'extern-badge';
        badge.textContent = externalBadge.label;
        card.appendChild(badge);
      }

      colDiv.appendChild(card);
    });

    fragment.appendChild(colDiv);
  });

  board.appendChild(fragment);
}

window.addEventListener('DOMContentLoaded', () => {
  ensureOpenCountWidget();
  startTickerAnimation();
  scheduleNextFetch();
  fetchData();
  setInterval(fetchData, FETCH_INTERVAL_MS);
  setInterval(updateTimerDisplay, 1000);
});
