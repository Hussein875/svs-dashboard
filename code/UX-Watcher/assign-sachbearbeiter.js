const fs = require('fs/promises');
const path = require('path');
const { chromium } = require('playwright');

const WORKER_BY_KEY = {
  hadi: 'Hadi Issa',
  ramazan: 'Ramazan Dag',
  robar: 'Robar Kassem',
  osama: 'Osama Sleiman',
  jad: 'Jad',
  hb: 'Hussein Selman',
  mz: 'Mohamed Zahreddine',
  hj: 'Hussein Jaber',
};

const WORKER_UX_CANDIDATES = {
  hadi: ['Hadi Issa', 'Herr Hadi Issa'],
  ramazan: ['Ramazan Dag', 'Herr Ramazan Dag'],
  robar: ['Robar Kassem', 'Herr Robar Kassem', 'Herr Robar Kassam'],
  osama: ['Osama Sleiman', 'Herr Osama Sleiman'],
  jad: ['Jad'],
  hb: ['Hussein Selman', 'Herr Hussein Selman'],
  mz: ['Mohamed Zahreddine', 'Herr Mohamed Zahreddine'],
  hj: ['Hussein Jaber', 'Herr Hussein Jaber'],
};

/** Drive-Ordner-Kürzel in Klammern → workerKey */
const FOLDER_SHORTCODE_TO_WORKER = {
  RO: 'robar',
  HB: 'hb',
  MZ: 'mz',
  HJ: 'hj',
};

function parsePositiveInt(rawValue, fallback) {
  const value = Number.parseInt(rawValue ?? '', 10);
  return Number.isFinite(value) && value > 0 ? value : fallback;
}

function requiredEnv(name) {
  const value = (process.env[name] || '').trim();
  if (!value) {
    throw new Error(`Fehlende Umgebungsvariable: ${name}`);
  }
  return value;
}

function truthyEnv(name) {
  return ['1', 'true', 'yes', 'on'].includes((process.env[name] || '').trim().toLowerCase());
}

function normalizeAkteRef(rawValue) {
  const text = String(rawValue || '').trim();
  const slashMatch = text.match(/\b(\d{3,5})\s*\/\s*(\d{2})\b/);
  if (slashMatch) return `${slashMatch[1]}/${slashMatch[2]}`;
  const digits = text.replace(/[^0-9]/g, '');
  if (digits.length >= 5) {
    return `${digits.slice(0, -2)}/${digits.slice(-2)}`;
  }
  return text;
}

function resolveAssigneeFromOptions(options = {}) {
  if (options.clear) return '';

  const assigneeName = String(options.assigneeName || '').trim();
  if (assigneeName) return assigneeName;

  const workerKey = String(options.workerKey || '').trim().toLowerCase();
  if (workerKey && WORKER_BY_KEY[workerKey]) return WORKER_BY_KEY[workerKey];

  throw new Error('assignSachbearbeiter: workerKey, assigneeName oder clear erforderlich');
}

function resolveUxCandidates(options = {}) {
  if (options.clear) return [];

  const workerKey = String(options.workerKey || '').trim().toLowerCase();
  if (workerKey && WORKER_UX_CANDIDATES[workerKey]) {
    return WORKER_UX_CANDIDATES[workerKey];
  }

  const assigneeName = resolveAssigneeFromOptions(options);
  return assigneeName ? [assigneeName] : [];
}

function resolveAssigneeNameFromEnv() {
  if (truthyEnv('UX_ASSIGN_CLEAR')) return '';

  const direct = (process.env.UX_ASSIGN_NAME || '').trim();
  if (direct) return direct;

  const key = (process.env.UX_ASSIGN_WORKER || '').trim().toLowerCase();
  if (key && WORKER_BY_KEY[key]) return WORKER_BY_KEY[key];

  throw new Error(
    'Setze UX_ASSIGN_NAME, UX_ASSIGN_WORKER (hadi|ramazan|robar|osama|jad|hb|mz|hj) oder UX_ASSIGN_CLEAR=1',
  );
}

function extractFolderShortcode(folderName) {
  const match = String(folderName || '').trim().match(/\(([A-Za-z]{1,4})\)\s*$/);
  return match ? match[1].toUpperCase() : '';
}

function extractAkteFromFolderName(folderName) {
  const text = String(folderName || '').trim();
  const slashMatch = text.match(/\b(\d{3,5})\s*\/\s*(\d{2})\b/);
  if (slashMatch) return `${slashMatch[1]}/${slashMatch[2]}`;
  return '';
}

function resolveWorkerFromFolderShortcode(folderName) {
  const shortcode = extractFolderShortcode(folderName);
  if (!shortcode) {
    return { shortcode: '', workerKey: null, skipped: true, reason: 'Kein Kürzel in Klammern' };
  }
  const workerKey = FOLDER_SHORTCODE_TO_WORKER[shortcode] || null;
  if (!workerKey) {
    return {
      shortcode,
      workerKey: null,
      skipped: true,
      reason: `Kürzel ${shortcode} wird nicht automatisch zugewiesen`,
    };
  }
  return {
    shortcode,
    workerKey,
    skipped: false,
    assigneeName: WORKER_BY_KEY[workerKey],
  };
}

function buildConfig(overrides = {}) {
  return {
    baseUrl: (process.env.UX_BASE_URL || 'https://ux.winvalue.de').replace(/\/+$/, ''),
    customerNr: requiredEnv('UX_CUSTOMER_NR'),
    password: requiredEnv('UX_PASSWORD'),
    akte: normalizeAkteRef(overrides.akte || process.env.UX_ASSIGN_AKTE || '1320/26'),
    dossierUrl: String(overrides.dossierUrl || process.env.UX_ASSIGN_DOSSIER_URL || '').trim(),
    assigneeName: overrides.assigneeName ?? resolveAssigneeNameFromEnv(),
    timeoutMs: parsePositiveInt(process.env.UX_TIMEOUT_MS, 45000),
    headless: overrides.headless ?? !truthyEnv('UX_ASSIGN_HEADED'),
    debugScreenshotPath: (process.env.UX_ASSIGN_DEBUG_SCREENSHOT || '').trim(),
    lockFile: process.env.UX_ASSIGN_LOCK_FILE || path.join(__dirname, '.assign.lock'),
    lockMaxAgeMs: parsePositiveInt(process.env.UX_ASSIGN_LOCK_MAX_AGE_MS, 20 * 60 * 1000),
  };
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function isStaleLock(lockFile, lockMaxAgeMs) {
  try {
    const stat = await fs.stat(lockFile);
    return (Date.now() - stat.mtimeMs) > lockMaxAgeMs;
  } catch {
    return false;
  }
}

async function acquireLock(lockFile, lockMaxAgeMs) {
  try {
    const handle = await fs.open(lockFile, 'wx');
    await handle.writeFile(`${process.pid}\n${new Date().toISOString()}\n`);
    return handle;
  } catch (err) {
    if (err.code !== 'EEXIST') throw err;

    if (await isStaleLock(lockFile, lockMaxAgeMs)) {
      await fs.unlink(lockFile).catch(() => {});
      return acquireLock(lockFile, lockMaxAgeMs);
    }
    return null;
  }
}

async function releaseLock(lockHandle, lockFile) {
  if (!lockHandle) return;
  await lockHandle.close().catch(() => {});
  await fs.unlink(lockFile).catch(() => {});
}

async function closeUpdateModal(page) {
  const closeBtn = page.locator('div.modal-content .modal-header button.close').first();
  try {
    await closeBtn.waitFor({ state: 'visible', timeout: 5000 });
  } catch {
    return;
  }

  try {
    await closeBtn.click({ timeout: 3000 });
  } catch {
    await page.keyboard.press('Escape').catch(() => {});
  }

  await page.locator('div.modal.show, div.modal.in').first()
    .waitFor({ state: 'hidden', timeout: 3000 })
    .catch(() => {});
}

async function login(page, config) {
  await page.goto(`${config.baseUrl}/`, { waitUntil: 'domcontentloaded' });
  await page.fill('input[name="customerNr"]', config.customerNr);
  await page.fill('input[name="password"]', config.password);
  await Promise.all([
    page.waitForLoadState('networkidle', { timeout: config.timeoutMs }).catch(() => {}),
    page.click('button[type="submit"]'),
  ]);
  await closeUpdateModal(page);
  console.log('✅ Login ausgeführt');
}

async function openAkteByUrl(page, dossierUrl, config) {
  const url = String(dossierUrl || '').trim();
  if (!url) return false;

  await page.goto(url, { waitUntil: 'domcontentloaded' });
  await page.waitForLoadState('networkidle', { timeout: config.timeoutMs }).catch(() => {});
  await closeUpdateModal(page);
  await sleep(800);
  console.log(`✅ Akte per Dossier-URL geöffnet: ${url}`);
  return true;
}

async function openAkteFromDossiers(page, aktenzeichen, config) {
  if (config.dossierUrl) {
    const opened = await openAkteByUrl(page, config.dossierUrl, config);
    if (opened) return;
  }

  await page.goto(`${config.baseUrl}/ux/home/dossiers`, { waitUntil: 'domcontentloaded' });
  await page.waitForLoadState('networkidle', { timeout: config.timeoutMs }).catch(() => {});
  await closeUpdateModal(page);
  await sleep(800);

  const row = page.locator('table.table-hover tbody tr').filter({ hasText: aktenzeichen }).first();
  if (await row.count()) {
    await row.click();
    await page.waitForLoadState('networkidle', { timeout: config.timeoutMs }).catch(() => {});
    console.log(`✅ Akte ${aktenzeichen} aus Übersicht geöffnet`);
    return;
  }

  const card = page.locator('.databox.databox-widget').filter({ hasText: aktenzeichen }).first();
  if (await card.count()) {
    await card.click();
    await page.waitForURL(/\/ux\/home\/dossiers\/edit\//, { timeout: config.timeoutMs });
    await page.waitForLoadState('networkidle', { timeout: config.timeoutMs }).catch(() => {});
    await sleep(1200);
    console.log(`✅ Akte ${aktenzeichen} aus Kartenansicht geöffnet`);
    return;
  }

  const link = page.getByRole('link', { name: new RegExp(aktenzeichen.replace('/', '\\/')) }).first();
  if (await link.count()) {
    await link.click();
    await page.waitForLoadState('networkidle', { timeout: config.timeoutMs }).catch(() => {});
    console.log(`✅ Akte ${aktenzeichen} per Link geöffnet`);
    return;
  }

  throw new Error(`Akte ${aktenzeichen} in der Übersicht nicht gefunden.`);
}

async function dismissMultipleAccessWarning(page) {
  const banner = page.locator('.alert').filter({ hasText: /Mehrfachzugriff/i }).first();
  if (!(await banner.count())) return;

  const closeBtn = banner.locator('button.close, .close').first();
  if (await closeBtn.count()) {
    await closeBtn.click().catch(() => {});
    await sleep(300);
  }
}

async function openAuftragTab(page, config) {
  await dismissMultipleAccessWarning(page);

  const current = page.url();
  const match = current.match(/^(https?:\/\/[^/]+\/ux\/home\/dossiers\/edit\/[^/]+)/);
  if (match) {
    const orderUrl = `${match[1]}/order/general`;
    await page.goto(orderUrl, { waitUntil: 'domcontentloaded' });
    await page.waitForLoadState('networkidle', { timeout: config.timeoutMs }).catch(() => {});
    await page.getByText(/Weitere Sachbearbeiter/i).first()
      .waitFor({ state: 'visible', timeout: config.timeoutMs })
      .catch(() => {});
    console.log('✅ Auftrag (order/general) geöffnet');
    return;
  }

  const auftrag = page.getByText('Auftrag', { exact: true }).first();
  if (await auftrag.count()) {
    await auftrag.click({ force: true });
    await page.waitForLoadState('networkidle', { timeout: config.timeoutMs }).catch(() => {});
    await sleep(1000);
    console.log('✅ Bereich Auftrag geöffnet');
    return;
  }

  throw new Error('Bereich „Auftrag“ nicht erreichbar.');
}

async function findWeitereSachbearbeiterField(page) {
  const byLabel = page.locator('label').filter({ hasText: /Weitere Sachbearbeiter/i }).first();
  if (await byLabel.count()) {
    const group = byLabel.locator('xpath=ancestor::*[contains(@class,"form-group")][1]');
    const input = group.locator('input, textarea, [contenteditable="true"]').first();
    if (await input.count()) return input;

    const siblingInput = byLabel.locator('xpath=following::input[1]');
    if (await siblingInput.count()) return siblingInput;
  }

  const placeholder = page.locator('input[placeholder*="Sachbearbeiter" i]').first();
  if (await placeholder.count()) return placeholder;

  const textNear = page.getByText(/Weitere Sachbearbeiter/i).first();
  if (await textNear.count()) {
    const container = textNear.locator('xpath=ancestor::div[1]');
    const input = container.locator('input').first();
    if (await input.count()) return input;
  }

  throw new Error('Feld „Weitere Sachbearbeiter“ nicht gefunden.');
}

async function clearTypeaheadField(page, field, config) {
  await field.click({ timeout: config.timeoutMs });
  await sleep(200);

  const clearBtn = page.locator(
    '.select2-selection__clear, .glyphicon-remove, button[title*="löschen" i], button[aria-label*="clear" i]',
  ).first();
  if (await clearBtn.count()) {
    await clearBtn.click().catch(() => {});
    await sleep(200);
  }

  await field.fill('');
  await page.keyboard.press('ControlOrMeta+A').catch(() => {});
  await page.keyboard.press('Backspace').catch(() => {});
  await sleep(200);
}

async function pickTypeaheadOption(page, candidates, config) {
  const names = [...new Set(candidates.filter(Boolean))];
  const searchTerms = [];

  for (const name of names) {
    searchTerms.push(name);
    const withoutTitle = name.replace(/^Herr\s+/i, '').trim();
    if (withoutTitle) searchTerms.push(withoutTitle);
    const lastName = withoutTitle.split(/\s+/).pop();
    if (lastName && lastName.length > 2) searchTerms.push(lastName);
  }

  for (const term of [...new Set(searchTerms)]) {
    const field = await findWeitereSachbearbeiterField(page);
    await field.click();
    await field.fill('');
    await page.keyboard.type(term, { delay: 35 });
    await sleep(700);

    for (const name of names) {
      const optionSelectors = [
        page.locator('.select2-results__option').filter({ hasText: name }),
        page.locator('[role="option"]').filter({ hasText: name }),
        page.locator('.dropdown-menu li, .autocomplete-suggestions div').filter({ hasText: name }),
        page.getByText(name, { exact: true }),
      ];

      for (const locator of optionSelectors) {
        const option = locator.first();
        if (!(await option.count())) continue;
        await option.click();
        console.log(`✅ Auswahl gesetzt: ${name}`);
        return name;
      }
    }

    const fuzzy = page.locator('.select2-results__option, [role="option"], .dropdown-menu li')
      .filter({ hasText: new RegExp(names.map((n) => n.replace(/^Herr\s+/i, '').split(' ').pop()).join('|'), 'i') })
      .first();
    if (await fuzzy.count()) {
      const picked = ((await fuzzy.innerText()) || '').trim();
      await fuzzy.click();
      console.log(`✅ Auswahl gesetzt (Treffer): ${picked}`);
      return picked;
    }
  }

  await page.keyboard.press('ArrowDown').catch(() => {});
  await page.keyboard.press('Enter').catch(() => {});
  console.log(`⚠️ Keine Dropdown-Option geklickt — Enter auf „${names[0]}“ gesendet`);
  return names[0] || '';
}

async function setWeitereSachbearbeiter(page, candidates, config) {
  const label = page.getByText(/Weitere Sachbearbeiter/i).first();
  await label.scrollIntoViewIfNeeded().catch(() => {});
  await sleep(300);

  const field = await findWeitereSachbearbeiterField(page);
  await field.scrollIntoViewIfNeeded().catch(() => {});
  await clearTypeaheadField(page, field, config);

  if (!candidates.length) {
    console.log('ℹ️ Weitere Sachbearbeiter wird geleert (Eingang).');
    return '';
  }

  return pickTypeaheadOption(page, candidates, config);
}

async function saveAuftrag(page, config) {
  await page.evaluate(() => window.scrollTo(0, document.body.scrollHeight));
  await sleep(400);

  const candidates = [
    page.getByRole('button', { name: /Speichern/i }),
    page.locator('button').filter({ hasText: /^Speichern$/i }),
    page.locator('.btn, a, div[role="button"]').filter({ hasText: /^Speichern$/i }),
  ];

  for (const locator of candidates) {
    const saveBtn = locator.first();
    if (!(await saveBtn.count())) continue;
    await saveBtn.scrollIntoViewIfNeeded().catch(() => {});
    await saveBtn.click({ timeout: config.timeoutMs });
    await page.waitForLoadState('networkidle', { timeout: config.timeoutMs }).catch(() => {});
    await sleep(800);
    console.log('✅ Auftrag gespeichert');
    return;
  }

  throw new Error('Speichern-Button nicht gefunden.');
}

async function runAssignOnce(config, uxCandidates = []) {
  const candidates = uxCandidates.length
    ? uxCandidates
    : (config.assigneeName ? [config.assigneeName] : []);

  const browser = await chromium.launch({ headless: config.headless });
  const page = await browser.newPage();
  page.setDefaultTimeout(config.timeoutMs);
  page.setDefaultNavigationTimeout(config.timeoutMs);

  try {
    await login(page, config);
    await openAkteFromDossiers(page, config.akte, config);
    await openAuftragTab(page, config);
    await setWeitereSachbearbeiter(page, candidates, config);
    await saveAuftrag(page, config);

    if (config.debugScreenshotPath) {
      await page.screenshot({ path: config.debugScreenshotPath, fullPage: true });
      console.log(`📸 Screenshot: ${config.debugScreenshotPath}`);
    }
  } catch (err) {
    const shotPath = config.debugScreenshotPath || path.join(__dirname, 'assign-error.png');
    await page.screenshot({ path: shotPath, fullPage: true }).catch(() => {});
    console.error(`❌ Fehler: ${err.message}`);
    console.error(`📸 Debug-Screenshot: ${shotPath}`);
    throw err;
  } finally {
    await browser.close().catch(() => {});
  }
}

async function assignSachbearbeiter(options = {}) {
  const assigneeName = resolveAssigneeFromOptions(options);
  const uxCandidates = resolveUxCandidates(options);
  const config = buildConfig({
    akte: options.akte,
    dossierUrl: options.dossierUrl,
    assigneeName,
    headless: options.headless,
  });

  const lockHandle = await acquireLock(config.lockFile, config.lockMaxAgeMs);
  if (!lockHandle) {
    const err = new Error('Assign läuft bereits. Dieser Lauf wird übersprungen.');
    err.code = 'ASSIGN_BUSY';
    throw err;
  }

  try {
    console.log(`ℹ️ Akte: ${config.akte}`);
    console.log(`ℹ️ Weitere Sachbearbeiter: ${uxCandidates.length ? uxCandidates.join(' | ') : '(leer)'}`);
    await runAssignOnce(config, uxCandidates);
    return { ok: true, akte: config.akte, assigneeName };
  } finally {
    await releaseLock(lockHandle, config.lockFile);
  }
}

async function runCli() {
  const config = buildConfig();
  console.log(`ℹ️ Akte: ${config.akte}`);
  console.log(`ℹ️ Weitere Sachbearbeiter: ${config.assigneeName || '(leer)'}`);

  const lockHandle = await acquireLock(config.lockFile, config.lockMaxAgeMs);
  if (!lockHandle) {
    console.log('⏭️ Assign läuft bereits. Dieser Lauf wird übersprungen.');
    return;
  }

  try {
    await runAssignOnce(config);
  } catch {
    process.exitCode = 1;
  } finally {
    await releaseLock(lockHandle, config.lockFile);
  }
}

module.exports = {
  WORKER_BY_KEY,
  WORKER_UX_CANDIDATES,
  FOLDER_SHORTCODE_TO_WORKER,
  assignSachbearbeiter,
  normalizeAkteRef,
  extractFolderShortcode,
  extractAkteFromFolderName,
  resolveWorkerFromFolderShortcode,
};

if (require.main === module) {
  runCli();
}
