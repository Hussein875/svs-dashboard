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
function makeEmptyMap() {
  return columns.reduce((m, c) => (m[c] = [], m), {});
}

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

    const nextNummer = nummern.length ? Math.max(...nummern) + 1 : '–';

    setTickerText(`🕸️ Aktuelle Nummer: ${nextNummer} 🎃`);

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