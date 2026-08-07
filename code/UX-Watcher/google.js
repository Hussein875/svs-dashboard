const path = require('path');
const { google } = require('googleapis');

const KEY_FILE = process.env.GOOGLE_APPLICATION_CREDENTIALS
  || path.join(__dirname, 'ux-dashboard-465511-29cd7fce4011.json');
const SHEET_ID = process.env.SHEET_ID || '10mfm9SVVDiWcxnfK2QuUCj3msaVFBQIQx34NnPlUEo4';
const TAB_NAME = process.env.SHEET_TAB_NAME || 'Dashboard';
const ASSIGNEE_COLUMN = process.env.SHEET_ASSIGNEE_COLUMN || 'B';
const STATUS_COLUMN = process.env.SHEET_STATUS_COLUMN || 'C';

const auth = new google.auth.GoogleAuth({
  keyFile: KEY_FILE,
  scopes: ['https://www.googleapis.com/auth/spreadsheets']
});

function normalizeAktenzeichen(rawValue) {
  return String(rawValue || '').replace(/[^0-9]/g, '');
}

function isVersendetStatus(rawStatus) {
  return /^versendet\b/i.test(String(rawStatus || '').trim());
}

async function getSheetsClient() {
  const client = await auth.getClient();
  return google.sheets({ version: 'v4', auth: client });
}

async function readAktenzeichen() {
  const sheets = await getSheetsClient();
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range: `${TAB_NAME}!A2:A`
  });

  return Array.isArray(res.data.values)
    ? res.data.values.map((row) => row[0]).filter(Boolean)
    : [];
}

function isNonEmptyBearbeiter(rawValue) {
  return String(rawValue ?? '').trim().length > 0;
}

async function updateStatus(aktenzeichen, status, bearbeiter = null) {
  const normalizedAkte = normalizeAktenzeichen(aktenzeichen);
  if (!normalizedAkte) return;

  const statusValue = String(status || '').trim();
  const sheets = await getSheetsClient();
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range: `${TAB_NAME}!A2:C`
  });

  const rows = res.data.values || [];
  const rowIndex = rows.findIndex((row) => normalizeAktenzeichen(row[0]) === normalizedAkte);

  if (rowIndex < 0) {
    console.log(`⚠️ Aktenzeichen ${aktenzeichen} nicht in Tabelle gefunden.`);
    return;
  }

  if (isVersendetStatus(statusValue)) {
    const filteredRows = rows.filter((_, idx) => idx !== rowIndex);

    await sheets.spreadsheets.values.clear({
      spreadsheetId: SHEET_ID,
      range: `${TAB_NAME}!A2:C`
    });

    if (filteredRows.length > 0) {
      await sheets.spreadsheets.values.update({
        spreadsheetId: SHEET_ID,
        range: `${TAB_NAME}!A2`,
        valueInputOption: 'RAW',
        requestBody: { values: filteredRows }
      });
    }

    console.log(`🧹 Akte ${aktenzeichen} entfernt (Versendet).`);
    return;
  }

  const targetRow = rowIndex + 2;
  const data = [{ range: `${TAB_NAME}!${STATUS_COLUMN}${targetRow}`, values: [[statusValue]] }];
  const bearbeiterName = String(bearbeiter ?? '').trim();

  // Watcher: leeren UX-Bearbeiter nicht ins Sheet schreiben — Spalte B bleibt erhalten.
  if (isNonEmptyBearbeiter(bearbeiterName)) {
    data.push({ range: `${TAB_NAME}!${ASSIGNEE_COLUMN}${targetRow}`, values: [[bearbeiterName]] });
  }

  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId: SHEET_ID,
    requestBody: {
      valueInputOption: 'RAW',
      data
    }
  });

  if (isNonEmptyBearbeiter(bearbeiterName)) {
    console.log(`✅ Status für ${aktenzeichen} → ${statusValue}, Bearbeiter → ${bearbeiterName}`);
  } else {
    console.log(`✅ Status für ${aktenzeichen} → ${statusValue} (Bearbeiter unverändert)`);
  }
}

const SHORT_NAME_TO_SHEET = {
  Hadi: 'Hadi Issa',
  Ramazan: 'Ramazan Dag',
  Robar: 'Robar Kassem',
  Osama: 'Osama Sleiman',
  Jad: 'Jad',
};

function sheetAkteLabel(aktenzeichen) {
  const text = String(aktenzeichen || '').trim();
  const slashMatch = text.match(/\b(\d{3,5})\s*\/\s*(\d{2})\b/);
  if (slashMatch) return slashMatch[1];
  const digits = normalizeAktenzeichen(text);
  return digits || text;
}

async function updateAssignee(aktenzeichen, bearbeiterShortName) {
  const normalizedAkte = normalizeAktenzeichen(aktenzeichen);
  if (!normalizedAkte) return false;

  const shortName = String(bearbeiterShortName ?? '').trim();
  const assigneeValue = shortName === '' ? '' : (SHORT_NAME_TO_SHEET[shortName] || shortName);

  const sheets = await getSheetsClient();
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range: `${TAB_NAME}!A2:C`,
  });

  const rows = res.data.values || [];
  const rowIndex = rows.findIndex((row) => normalizeAktenzeichen(row[0]) === normalizedAkte);

  if (rowIndex < 0) {
    // Noch nicht importiert: Zeile anlegen (Drive-Kürzel-Assign vor Cron-Import)
    await sheets.spreadsheets.values.append({
      spreadsheetId: SHEET_ID,
      range: `${TAB_NAME}!A:B`,
      valueInputOption: 'RAW',
      requestBody: { values: [[sheetAkteLabel(aktenzeichen), assigneeValue]] },
    });
    console.log(`✅ Akte ${aktenzeichen} neu angelegt, Bearbeiter → ${assigneeValue || '(leer)'}`);
    return true;
  }

  const targetRow = rowIndex + 2;
  await sheets.spreadsheets.values.update({
    spreadsheetId: SHEET_ID,
    range: `${TAB_NAME}!${ASSIGNEE_COLUMN}${targetRow}`,
    valueInputOption: 'RAW',
    requestBody: { values: [[assigneeValue]] },
  });

  console.log(`✅ Bearbeiter für ${aktenzeichen} → ${assigneeValue || '(leer)'}`);
  return true;
}

module.exports = { readAktenzeichen, updateStatus, updateAssignee };
