// Ticker-Animation starten, wenn DOM geladen ist
window.addEventListener("DOMContentLoaded", () => {
  fetchData();
  startTickerAnimation();
});

function startTickerAnimation() {
  const ticker = document.querySelector('.ticker');
  if (ticker) {
    const spans = ticker.querySelectorAll('.ticker-span');
    spans.forEach(span => {
      span.style.animation = 'ticker-scroll 25s linear infinite';
    });
  }
}
const sheetID = '10mfm9SVVDiWcxnfK2QuUCj3msaVFBQIQx34NnPlUEo4';
const url = `https://docs.google.com/spreadsheets/d/${sheetID}/gviz/tq?tqx=out:json`;

const columns = ["Eingang", "Ahmet", "Hadi", "Osama", "Geprüft"];

let lastFetchTime = null;

// Aktenzeichen extrahieren (z. B. RB 1012 → "RB 1012", 1078 → "1078")
function extractAktenzeichen(text) {
  if (!text) return '';
  const match = text.match(/^(RB|KVA)?\s*\d+/i);
  return match ? match[0].toUpperCase().replace(/\s+/, ' ') : text;
}
// Daten aus dem Google Sheet abrufen und verarbeiten
function fetchData() {
  fetch(url)
    .then(res => res.text())
    .then(text => {
      const json = JSON.parse(text.substr(47).slice(0, -2));

      const rows = json.table.rows.map(row => {
        const eingangRaw = row.c[0]?.v !== null && row.c[0]?.v !== undefined ? String(row.c[0].v).trim() : '';
        const eingang = extractAktenzeichen(eingangRaw);
        const bearbeiter = row.c[1]?.v !== null && row.c[1]?.v !== undefined ? String(row.c[1].v).trim() : '';
        const status = row.c[2]?.v !== null && row.c[2]?.v !== undefined ? String(row.c[2].v).trim().toLowerCase() : '';

        return { Eingang: eingang, Bearbeiter: bearbeiter, Status: status };
      }).filter(row =>
        row.Eingang !== '' &&
        row.Eingang.toLowerCase() !== 'eingang'
      );

      // Berechne höchste Aktennummer für "nächste Nummer"
      // 1. Nach dem Parsen der Daten:
      const nummern = rows
        .map(r => r.Eingang.match(/\d+/))              // Ziffern aus Aktenzeichen extrahieren
        .filter(match => match && !isNaN(match[0]))     // nur gültige Nummern
        .map(match => parseInt(match[0], 10));          // als Zahl umwandeln

      let nextNummer = '–';
      if (nummern.length > 0) {
        const maxNummer = Math.max(...nummern);
        nextNummer = maxNummer + 1;
      }

      const tickerText = `Aktuelle Nummer: ${nextNummer} `;
      const tickerElement = document.querySelector('.ticker');

      if (tickerElement) {
        tickerElement.innerHTML = `<span class="ticker-span">${tickerText}</span>`;
      }

      renderBoard(rows);
      lastFetchTime = new Date(); // Merke Uhrzeit des echten Abrufs
      updateTimerDisplay(); // ✅ Das ist die einzige wichtige Zeile hier für den Timer

    });
}

//Timer visualiserung
function updateTimerDisplay() {
  const now = new Date();
  const minutes = now.getMinutes();

  // Nächstes 5-Minuten-Intervall berechnen
  const nextFive = Math.ceil((minutes + 1) / 5) * 5;
  const nextUpdate = new Date(now);
  nextUpdate.setMinutes(nextFive);
  nextUpdate.setSeconds(0);

  // Falls nextFive 60 ist, auf nächste Stunde setzen
  if (nextUpdate.getMinutes() === 60) {
    nextUpdate.setHours(nextUpdate.getHours() + 1);
    nextUpdate.setMinutes(0);
  }

  const diffMs = nextUpdate - now;
  const minutesLeft = Math.floor(diffMs / 60000);
  const secondsLeft = Math.floor((diffMs % 60000) / 1000);

  const updateInfo = document.getElementById("updateInfo");
  updateInfo.innerHTML =
    `⏱️ <b>Letztes Update:</b> ${lastFetchTime ? lastFetchTime.toLocaleTimeString() : "unbekannt"} &nbsp;|&nbsp; 🔄 <b>Nächstes in:</b> ${minutesLeft}m ${secondsLeft}s`;

  // Farbe zurücksetzen
  updateInfo.className = 'update-info';

  if (minutesLeft >= 3) {
    updateInfo.classList.add('green');
  } else if (minutesLeft >= 1) {
    updateInfo.classList.add('orange');
  } else {
    updateInfo.classList.add('red');
  }
}

// Board visualisieren
function renderBoard(data) {
  const board = document.getElementById('board');
  board.innerHTML = '';

  const map = {
    Eingang: [],
    Ahmet: [],
    Hadi: [],
    Osama: [],
    Geprüft: []
  };

  data.forEach(row => {
    const status = row.Status.toLowerCase();
    const akte = {
      nummer: row.Eingang,
      status: status,
      bearbeiter: row.Bearbeiter
    };

    // trifft auf geprüft o, geprüft 1 oder geprüft 2 zu (Groß-/Kleinschreibung egal)
    console.log(akte.status)
    if (/geprüft [o12hjhk]/i.test(status)) {
      map.Geprüft.push(akte);
    } else if (status.trim().toLowerCase() === 'vollständig') {
      map.Osama.push(akte);
    } else if (columns.includes(row.Bearbeiter)) {
      map[row.Bearbeiter].push(akte);
    } else {
      map.Eingang.push(akte);
    }
  });

  columns.forEach(col => {
    const colDiv = document.createElement('div');
    colDiv.className = 'column';

    const h2 = document.createElement('h2');
    const anzahl = map[col].length;
    h2.innerText = `${col} (${anzahl})`;
    colDiv.appendChild(h2);

    map[col].forEach(item => {
      const aktenzeichen = typeof item === 'string' ? item : item.nummer;
      const status = typeof item === 'object' ? item.status : '';
      const bearbeiter = item.bearbeiter;

      if (aktenzeichen.toLowerCase() === col.toLowerCase()) return;

      const card = document.createElement('div');
      card.className = 'card';

      if (bearbeiter === 'HJ') {
        card.classList.add('hj');
        card.title = 'Bearbeitung durch Hannover (HJ)';
      }

      // ✅ Farbe erst nach Erzeugung der Karte setzen
      if (/geprüft [o12]/i.test(status)) {
        card.style.backgroundColor = '#13e339ff'; // grün
        card.style.color = 'white';
      } else if (status.includes('unvollständig')) {
        card.style.backgroundColor = '#e31313ff'; // rot
        card.style.color = 'white';
      } else if (status.includes('vollständig')) {
        card.style.backgroundColor = '#ff7300ff'; // orange
        card.style.color = 'white';
      }

      card.innerText = aktenzeichen;
      colDiv.appendChild(card);
    });

    board.appendChild(colDiv);
  });
}


setInterval(updateTimerDisplay, 1000); // Live-Countdown

// Automatisch alle 60 Sekunden neu laden
setInterval(fetchData, 60000);