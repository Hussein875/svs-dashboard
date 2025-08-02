const sheetID = '10mfm9SVVDiWcxnfK2QuUCj3msaVFBQIQx34NnPlUEo4';
const url = `https://docs.google.com/spreadsheets/d/${sheetID}/gviz/tq?tqx=out:json`;

const columns = ["Eingang", "Ahmet", "Hadi", "Osama", "Erledigt"];

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

      renderBoard(rows);
      updateTimerDisplay(); // ✅ Das ist die einzige wichtige Zeile hier für den Timer

    });
}

//Timer visualiserung
function updateTimerDisplay() {
  const now = new Date();
  const minutes = now.getMinutes();
  const nextQuarter = Math.ceil((minutes + 1) / 15) * 15;
  const nextUpdate = new Date(now);
  nextUpdate.setMinutes(nextQuarter);
  nextUpdate.setSeconds(0);

  if (nextUpdate.getMinutes() === 60) {
    nextUpdate.setHours(nextUpdate.getHours() + 1);
    nextUpdate.setMinutes(0);
  }

  const diffMs = nextUpdate - now;
  const minutesLeft = Math.floor(diffMs / 60000);
  const secondsLeft = Math.floor((diffMs % 60000) / 1000);

  const updateInfo = document.getElementById("updateInfo");
  updateInfo.innerHTML =
    `⏱️ <b>Letztes Update:</b> ${now.toLocaleTimeString()} &nbsp;|&nbsp; 🔄 <b>Nächstes in:</b> ${minutesLeft}m ${secondsLeft}s`;

  // Farbe zurücksetzen
  updateInfo.className = 'update-info';

  if (minutesLeft >= 5) {
    updateInfo.classList.add('green');
  } else if (minutesLeft >= 1) {
    updateInfo.classList.add('orange');
  } else {
    updateInfo.classList.add('red');
    updateInfo.classList.add('blink');
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
    Erledigt: []
  };

  data.forEach(row => {
    const status = row.Status.toLowerCase();
    const akte = {
      nummer: row.Eingang,
      status: status
    };

    if (status.includes('geprüft o')) {
      map.Erledigt.push(akte);
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
    h2.innerText = col;
    colDiv.appendChild(h2);

    map[col].forEach(item => {
      const aktenzeichen = typeof item === 'string' ? item : item.nummer;
      const status = typeof item === 'object' ? item.status : '';

      if (aktenzeichen.toLowerCase() === col.toLowerCase()) return;

      const card = document.createElement('div');
      card.className = 'card';

      if (status.includes('vollständig')) {
        card.style.backgroundColor = '#007bff';
        card.style.color = 'white';
      }

      card.innerText = aktenzeichen;
      colDiv.appendChild(card);
    });

    board.appendChild(colDiv);
  });
}

// Beim Laden starten
window.addEventListener("DOMContentLoaded", fetchData);

setInterval(updateTimerDisplay, 1000); // Live-Countdown

// Automatisch alle 60 Sekunden neu laden
setInterval(fetchData, 60000);