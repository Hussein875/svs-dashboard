// --- 1) Ticker-Animation: nach JEDEM Update neu starten ---
function setTickerText(text) {
  const ticker = document.querySelector('.ticker');
  if (!ticker) return;
  ticker.textContent = '';                       // XSS-sicher
  const span = document.createElement('span');
  span.className = 'ticker-span';
  span.textContent = text;
  ticker.appendChild(span);
  startTickerAnimation();                        // <- neu: Animation reaktivieren
}

function startTickerAnimation() {
  const spans = document.querySelectorAll('.ticker .ticker-span');
  spans.forEach(span => {
    span.style.animation = 'ticker-scroll 25s linear infinite';
  });
}

// --- 2) Fetch: Überlappungen vermeiden + robustere JSON-Extraktion ---
let lastFetchTime = null;
let inFlightController = null;

const sheetID = '10mfm9SVVDiWcxnfK2QuUCj3msaVFBQIQx34NnPlUEo4';
const url = `https://docs.google.com/spreadsheets/d/${sheetID}/gviz/tq?tqx=out:json`;

// Spalten-Definition und Board-Map automatisch synchron halten
const columns = ["Eingang", "Hadi", "Ramazan", "Osama", "Geprüft"];
// Key für die einfache Tages-/Wochenstatistik (nur Nummern-Differenzen)
const STATS_ENDPOINT = 'https://script.google.com/macros/s/AKfycbza16YNgG6pXVWEUJmKEKpNgVGjyx1A9y-e2M0d0L4Bjsf7_jZpuD3j7wPwrFy1AGtLfw/exec';

window.addEventListener("DOMContentLoaded", () => {
  fetchData();
  startTickerAnimation();
  setInterval(updateTimerDisplay, 1000);   // Live-Countdown
  setInterval(fetchData, 60000);           // alle 60s neu
});

function extractAktenzeichen(text) {
  if (!text) return '';
  const match = text.match(/^(RB|KVA)?\s*\d+/i);
  return match ? match[0].toUpperCase().replace(/\s+/, ' ') : String(text).trim();
}

// robuster Parser für gviz-JSONP
function parseGviz(text) {
  const start = text.indexOf('{');
  const end = text.lastIndexOf('}');
  if (start === -1 || end === -1) throw new Error('GViz JSON nicht gefunden');
  return JSON.parse(text.slice(start, end + 1));
}

async function fetchData() {
  try {
    // laufenden Request abbrechen (z. B. bei langsamer Verbindung)
    if (inFlightController) inFlightController.abort();
    inFlightController = new AbortController();

    const res = await fetch(url, { signal: inFlightController.signal, cache: 'no-store' });
    const raw = await res.text();
    const json = parseGviz(raw);

    const rows = json.table.rows.map(row => {
      const eingangRaw = row.c[0]?.v ?? '';
      const eingang = extractAktenzeichen(eingangRaw);
      const bearbeiter = (row.c[1]?.v ?? '').toString().trim();
      const status = (row.c[2]?.v ?? '').toString().trim().toLowerCase();

      return { Eingang: eingang, Bearbeiter: bearbeiter, Status: status };
    }).filter(r => r.Eingang && r.Eingang.toLowerCase() !== 'eingang');

    // 3) "versendet" einheitlich (case-insensitive) filtern, inkl. Varianten
    const cleaned = rows.filter(r => !/^versendet\b/i.test(r.Status));

    // 4) nächste Nummer berechnen (nur Ziffern)
    const nummern = cleaned
      .map(r => (r.Eingang.match(/\d+/) ?? [null])[0])
      .filter(n => n !== null)
      .map(n => parseInt(n, 10))
      .filter(n => Number.isFinite(n));

    const nextNummer = nummern.length ? Math.max(...nummern) + 1 : 0;

    // 3a) Tages- und Wochenstatistik aus der aktuellen Nummer ableiten
    const stats = await computeStatsFromNext(nextNummer);
    updateStatsBar(stats);

    setTickerText(`🏖️ Aktuelle Nummer: ${nextNummer || '–'} 🌴`);

    renderBoard(cleaned);
    lastFetchTime = new Date();   // Zeitpunkt des echten Abrufs
    updateTimerDisplay();
  } catch (err) {
    if (err.name === 'AbortError') return; // ignorieren
    console.error('fetchData() Fehler:', err);
  } finally {
    inFlightController = null;
  }
}

// --- 5) Timer robuster machen (Guard + TZ konsequent) ---
function updateTimerDisplay() {
  const el = document.getElementById("updateInfo");
  if (!el) return;

  const now = new Date();
  const minutes = now.getMinutes();
  const nextFive = Math.ceil((minutes + 1) / 5) * 5;
  const nextUpdate = new Date(now);
  nextUpdate.setMinutes(nextFive === 60 ? 0 : nextFive, 0, 0);
  if (nextFive === 60) nextUpdate.setHours(nextUpdate.getHours() + 1);

  const diffMs = nextUpdate - now;
  const minutesLeft = Math.max(0, Math.floor(diffMs / 60000));
  const secondsLeft = Math.max(0, Math.floor((diffMs % 60000) / 1000));

  const timeFmt = new Intl.DateTimeFormat('de-DE', {
    hour: '2-digit', minute: '2-digit', second: '2-digit'
  });

  el.className = 'update-info';
  if (minutesLeft >= 3) el.classList.add('green');
  else if (minutesLeft >= 1) el.classList.add('orange');
  else el.classList.add('red');

  el.innerHTML = `⏱️ <b>Letztes Update:</b> ${lastFetchTime ? timeFmt.format(lastFetchTime) : "unbekannt"} ` +
                 `&nbsp;|&nbsp; 🔄 <b>Nächstes in:</b> ${minutesLeft}m ${secondsLeft}s`;
}

// --- 6) Board-Render: DocumentFragment, sichere textContent, konsistente Regeln ---
function makeEmptyMap() {
  return columns.reduce((map, col) => {
    map[col] = [];
    return map;
  }, {});
}
function renderBoard(data) {
  const board = document.getElementById('board');
  if (!board) return;

  const map = makeEmptyMap();

  data.forEach(row => {
    const status = row.Status.toLowerCase().trim();
    const akte = { nummer: row.Eingang, status, bearbeiter: row.Bearbeiter };

    // geprüft (o/1/2/HJ/HK) – einheitlich
      if (/^geprüft\s*(O|1|2|HJ|HK)$/i.test(status)) {
      map.Geprüft.push(akte);
    } else if (status === 'vollständig') {
      map.Osama.push(akte);
    } else if (columns.includes(row.Bearbeiter)) {
      map[row.Bearbeiter].push(akte);
    } else {
      map.Eingang.push(akte);
    }
  });

  // optional: innerhalb der Spalten nach Nummer sortieren (absteigend)
  for (const col of columns) {
    map[col].sort((a, b) => {
      const na = parseInt((a.nummer.match(/\d+/) ?? ['0'])[0], 10);
      const nb = parseInt((b.nummer.match(/\d+/) ?? ['0'])[0], 10);
      return nb - na;
    });
  }

  // DOM effizient aufbauen
  board.textContent = '';
  const frag = document.createDocumentFragment();

  columns.forEach(col => {
    const colDiv = document.createElement('div');
    colDiv.className = 'column';

    const h2 = document.createElement('h2');
    const anzahl = map[col].length;
    h2.textContent = anzahl > 0 ? `${col} (${anzahl})` : col;
    colDiv.appendChild(h2);

    map[col].forEach(item => {
      const { nummer: aktenzeichen, status } = item;
      if (aktenzeichen.toLowerCase() === col.toLowerCase()) return;

      const card = document.createElement('div');
      card.className = 'card';

      // Farben
      if (/^geprüft\s*(O|1|2|HJ|HK)$/i.test(status)) {
        card.style.backgroundColor = '#13e339ff'; // grün
        card.style.color = 'white';
      } else if (status.includes('unvollständig')) {
        card.style.backgroundColor = '#e31313ff'; // rot
        card.style.color = 'white';
      } else if (status.includes('vollständig')) {
        card.style.backgroundColor = '#ff7300ff'; // orange
        card.style.color = 'white';
      }

      card.textContent = aktenzeichen;  // sicher
      colDiv.appendChild(card);
    });

    frag.appendChild(colDiv);
  });

  board.appendChild(frag);
}
// --- 7) Tages- und Wochenzähler ---
function getMonday(date) {
  const d = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  const day = d.getDay(); // So = 0, Mo = 1, ...
  const diff = (day + 6) % 7; // Mo = 0, Di = 1, ...
  d.setDate(d.getDate() - diff);
  d.setHours(0, 0, 0, 0);
  return d;
}

function toDateKey(d) {
  const year = d.getFullYear();
  const month = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

function normalizeDateKey(value) {
  if (!value) return null;

  // Versuchen, das, was vom Backend kommt (Date-Objekt-String wie
  // "Fri Nov 07 2025 03:00:00 GMT+0400 (Golf-Zeit)" oder "2025-11-07")
  // in unser YYYY-MM-DD-Format zu bringen.
  const d = new Date(value);
  if (!isNaN(d.getTime())) {
    return toDateKey(d);
  }

  return String(value).trim();
}

async function loadSimpleStats() {
  try {
    const res = await fetch(`${STATS_ENDPOINT}?t=${Date.now()}`, { cache: 'no-store' });
    if (!res.ok) {
      throw new Error(`Stats-GET HTTP ${res.status}`);
    }
    const data = await res.json();
    return {
      day: data.day || null,
      week: data.week || null,
    };
  } catch (e) {
    console.warn('Konnte Stats nicht vom Backend laden, starte neu:', e);
    return { day: null, week: null };
  }
}

async function saveSimpleStats(stats) {
  try {
    await fetch(STATS_ENDPOINT, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(stats),
    });
  } catch (e) {
    console.warn('Konnte Stats nicht zum Backend speichern:', e);
  }
}

async function computeStatsFromNext(currentNext) {
  const today = new Date();
  const todayKey = toDateKey(today);
  const mondayKey = toDateKey(getMonday(today));

  const stats = await loadSimpleStats() || {};

  // vorhandene Keys aus dem Backend normalisieren (können Date-Strings sein)
  if (stats.day && stats.day.dateKey) {
    stats.day.dateKey = normalizeDateKey(stats.day.dateKey);
  }
  if (stats.week && (stats.week.mondayKey || stats.week.dateKey)) {
    const rawWeekKey = stats.week.mondayKey || stats.week.dateKey;
    stats.week.mondayKey = normalizeDateKey(rawWeekKey);
  }

  // Falls du initial nur base eingetragen hast (dateKey/mondayKey leer),
  // übernehmen wir diese Basis für HEUTE / DIESE WOCHE und setzen NUR das Datum.
  if (stats.day && !stats.day.dateKey) {
    stats.day.dateKey = todayKey;
  }
  if (stats.week && !stats.week.mondayKey) {
    stats.week.mondayKey = mondayKey;
  }

  // Sicherheits-Defaults, falls noch gar nichts vorhanden ist
  if (!stats.day) {
       stats.day = { dateKey: todayKey, base: currentNext };
  }
  if (!stats.week) {
    stats.week = { mondayKey, base: currentNext };
  }

  // Roll-Over: wenn ein neuer Tag/Woche beginnt, Basis auf aktuelle Nummer setzen
  if (stats.day.dateKey !== todayKey) {
    stats.day = { dateKey: todayKey, base: currentNext };
  }
  if (stats.week.mondayKey !== mondayKey) {
    stats.week = { mondayKey, base: currentNext };
  }

  await saveSimpleStats(stats);

  const todayBase = Number(stats.day.base) || 0;
  const weekBase = Number(stats.week.base) || 0;

  const todayCount = Math.max(0, currentNext - todayBase);
  const weekCount = Math.max(0, currentNext - weekBase);

  return { todayCount, weekCount };
}

function updateStatsBar(stats) {
  const board = document.getElementById('board');
  if (!board) return;

  let bar = document.getElementById('statsBar');
  if (!bar) {
    bar = document.createElement('div');
    bar.id = 'statsBar';
    bar.className = 'stats-bar';

    if (board.parentNode) {
      board.parentNode.insertBefore(bar, board);
    } else {
      document.body.insertBefore(bar, document.body.firstChild);
    }
  }

  bar.innerHTML =
    `📊 &nbsp;<b>Heute:&nbsp;</b> ${stats.todayCount} Gutachten` +
    ` &nbsp;|&nbsp; ` +
    `📅 &nbsp;<b>Diese Woche:&nbsp;</b> ${stats.weekCount} Gutachten`;
}