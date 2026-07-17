const express = require('express');
const cors = require('cors');
const jwt = require('jsonwebtoken');
const { updateAssignee } = require('./google');
const {
  assignSachbearbeiter,
  extractAkteFromFolderName,
  normalizeAkteRef,
  resolveWorkerFromFolderShortcode,
  WORKER_BY_KEY,
} = require('./assign-sachbearbeiter');

const PORT = Number.parseInt(process.env.ASSIGN_PORT || '3080', 10);
const ADMIN_PIN = (process.env.ADMIN_PIN || '').trim();
const JWT_SECRET = (process.env.JWT_SECRET || ADMIN_PIN || 'dev-secret').trim();
const FOLDER_ASSIGN_SECRET = (process.env.FOLDER_ASSIGN_SECRET || '').trim();
const CORS_ORIGIN = (process.env.ASSIGN_CORS_ORIGIN || '*').trim();
const JWT_TTL = '24h';
const UX_ASSIGN_RETRY_MS = Number.parseInt(process.env.UX_ASSIGN_RETRY_MS || '5000', 10);
const UX_ASSIGN_MAX_RETRIES = Number.parseInt(process.env.UX_ASSIGN_MAX_RETRIES || '12', 10);

const uxAssignQueue = [];
let uxQueueRunning = false;

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function runUxAssignWithRetry(assignOptions, meta) {
  for (let attempt = 1; attempt <= UX_ASSIGN_MAX_RETRIES; attempt += 1) {
    try {
      await assignSachbearbeiter(assignOptions);
      console.log(`✅ UX-Assign abgeschlossen: ${meta.akte} → ${meta.column}`);
      return;
    } catch (err) {
      if (err.code === 'ASSIGN_BUSY' && attempt < UX_ASSIGN_MAX_RETRIES) {
        console.warn(`⚠️ UX-Assign wartet (busy), Versuch ${attempt}/${UX_ASSIGN_MAX_RETRIES}: ${meta.akte}`);
        await sleep(UX_ASSIGN_RETRY_MS);
        continue;
      }
      console.error(`❌ UX-Assign fehlgeschlagen (${meta.akte}):`, err.message);
      return;
    }
  }
}

async function processUxAssignQueue() {
  if (uxQueueRunning) return;
  uxQueueRunning = true;

  while (uxAssignQueue.length > 0) {
    const job = uxAssignQueue.shift();
    await runUxAssignWithRetry(job.assignOptions, job.meta);
  }

  uxQueueRunning = false;
}

function enqueueUxAssign(assignOptions, meta) {
  uxAssignQueue.push({ assignOptions, meta });
  processUxAssignQueue().catch((err) => {
    console.error('UX-Assign-Queue Fehler:', err.message);
  });
}

const ASSIGN_COLUMNS = new Set(['Hadi', 'Ramazan', 'Robar', 'Jad']);

function columnToShortName(column) {
  if (column === 'Eingang') return '';
  return column;
}

function folderAssignAuth(req, res, next) {
  if (!FOLDER_ASSIGN_SECRET) {
    res.status(500).json({ ok: false, error: 'FOLDER_ASSIGN_SECRET ist nicht konfiguriert' });
    return;
  }
  const provided = String(req.headers['x-folder-assign-secret'] || '').trim();
  if (!provided || provided !== FOLDER_ASSIGN_SECRET) {
    res.status(401).json({ ok: false, error: 'unauthorized' });
    return;
  }
  next();
}

async function assignByWorkerKey(akteRaw, workerKey, options = {}) {
  const akte = normalizeAkteRef(akteRaw);
  const key = String(workerKey || '').trim().toLowerCase();
  const dossierUrl = String(options.dossierUrl || '').trim();
  if (!akte) {
    return { ok: false, status: 400, error: 'akte erforderlich' };
  }
  if (!WORKER_BY_KEY[key]) {
    return { ok: false, status: 400, error: `Unbekannter workerKey: ${key}` };
  }

  const sheetName = WORKER_BY_KEY[key];
  let sheetUpdated = false;
  try {
    sheetUpdated = await updateAssignee(akte, sheetName);
  } catch (err) {
    console.error('Sheet-Update fehlgeschlagen:', err.message);
    return { ok: false, status: 500, error: 'Sheet-Update fehlgeschlagen', sheetUpdated: false };
  }

  if (!sheetUpdated) {
    return {
      ok: false,
      status: 404,
      error: `Akte ${akte} nicht im Sheet gefunden`,
      sheetUpdated: false,
    };
  }

  enqueueUxAssign(
    { akte, workerKey: key, dossierUrl: dossierUrl || undefined },
    { akte, column: key },
  );
  return {
    ok: true,
    status: 200,
    sheetUpdated: true,
    uxPending: true,
    akte,
    workerKey: key,
    assigneeName: sheetName,
  };
}

function authMiddleware(req, res, next) {
  const header = req.headers.authorization || '';
  const token = header.startsWith('Bearer ') ? header.slice(7).trim() : '';
  if (!token) {
    res.status(401).json({ ok: false, error: 'Authorization erforderlich' });
    return;
  }

  try {
    req.auth = jwt.verify(token, JWT_SECRET);
    next();
  } catch {
    res.status(401).json({ ok: false, error: 'Ungültiges oder abgelaufenes Token' });
  }
}

function createApp() {
  const app = express();
  app.use(express.json({ limit: '16kb' }));
  app.use(cors({
    origin: CORS_ORIGIN === '*' ? true : (origin, callback) => {
      const allowed = CORS_ORIGIN.split(',').map((v) => v.trim()).filter(Boolean);
      if (!origin || allowed.includes(origin) || allowed.includes('*')) {
        callback(null, true);
        return;
      }
      if (/^https?:\/\/(localhost|127\.0\.0\.1)(:\d+)?$/.test(origin)) {
        callback(null, true);
        return;
      }
      callback(new Error(`CORS blocked: ${origin}`));
    },
    methods: ['GET', 'POST', 'OPTIONS'],
    allowedHeaders: ['Content-Type', 'Authorization'],
  }));

  app.get('/api/health', (_req, res) => {
    res.json({ ok: true, service: 'assign-service' });
  });

  app.post('/api/auth', (req, res) => {
    const pin = String(req.body?.pin ?? '').trim();
    if (!/^\d{4}$/.test(ADMIN_PIN)) {
      res.status(500).json({ ok: false, error: 'ADMIN_PIN ist nicht konfiguriert' });
      return;
    }
    if (!/^\d{4}$/.test(pin)) {
      res.status(400).json({ ok: false, error: 'PIN muss 4 Ziffern haben' });
      return;
    }
    if (pin !== ADMIN_PIN) {
      res.status(401).json({ ok: false, error: 'Falscher PIN' });
      return;
    }

    const token = jwt.sign({ role: 'admin' }, JWT_SECRET, { expiresIn: JWT_TTL });
    res.json({ ok: true, token, expiresIn: JWT_TTL });
  });

  app.post('/api/assign', authMiddleware, async (req, res) => {
    const akte = String(req.body?.akte ?? '').trim();
    const column = String(req.body?.column ?? '').trim();

    if (!akte) {
      res.status(400).json({ ok: false, error: 'akte erforderlich' });
      return;
    }
    if (!ASSIGN_COLUMNS.has(column)) {
      res.status(400).json({
        ok: false,
        error: 'column muss Hadi, Ramazan, Robar oder Jad sein',
      });
      return;
    }

    const bearbeiterShortName = columnToShortName(column);
    let sheetUpdated = false;

    try {
      sheetUpdated = await updateAssignee(akte, bearbeiterShortName);
    } catch (err) {
      console.error('Sheet-Update fehlgeschlagen:', err.message);
      res.status(500).json({ ok: false, error: 'Sheet-Update fehlgeschlagen', sheetUpdated: false });
      return;
    }

    if (!sheetUpdated) {
      res.status(404).json({ ok: false, error: `Akte ${akte} nicht im Sheet gefunden`, sheetUpdated: false });
      return;
    }

    const assignOptions = { akte, workerKey: column.toLowerCase() };

    res.json({ ok: true, sheetUpdated: true, uxAssigned: false, uxPending: true });

    enqueueUxAssign(assignOptions, { akte, column });
  });

  // Drive-Ordner → Kürzel (RO/HB/MZ/HJ) → Bearbeiter setzen (Sheet + UX)
  app.post('/api/assign-from-folder', folderAssignAuth, async (req, res) => {
    const folderName = String(req.body?.folderName ?? '').trim();
    const akteRaw = String(req.body?.akte ?? '').trim()
      || extractAkteFromFolderName(folderName);

    if (!folderName && !akteRaw) {
      res.status(400).json({ ok: false, error: 'folderName oder akte erforderlich' });
      return;
    }

    const resolved = resolveWorkerFromFolderShortcode(folderName);
    if (resolved.skipped) {
      res.json({
        ok: true,
        skipped: true,
        reason: resolved.reason,
        shortcode: resolved.shortcode || null,
        folderName,
        akte: akteRaw || null,
      });
      return;
    }

    if (!akteRaw) {
      res.status(400).json({
        ok: false,
        error: 'Aktennummer konnte aus Ordnername nicht gelesen werden',
        folderName,
        shortcode: resolved.shortcode,
      });
      return;
    }

    const dossierUrl = String(req.body?.dossierUrl ?? '').trim();
    const result = await assignByWorkerKey(akteRaw, resolved.workerKey, { dossierUrl });
    if (!result.ok) {
      res.status(result.status || 500).json(result);
      return;
    }

    res.json({
      ...result,
      skipped: false,
      shortcode: resolved.shortcode,
      folderName,
      dossierUrl: dossierUrl || null,
    });
  });

  return app;
}

function startServer() {
  if (!/^\d{4}$/.test(ADMIN_PIN)) {
    console.error('❌ ADMIN_PIN muss eine 4-stellige Zahl sein.');
    process.exit(1);
  }

  const app = createApp();
  app.listen(PORT, () => {
    console.log(`✅ assign-service lauscht auf Port ${PORT}`);
  });
}

module.exports = { createApp, startServer };

if (require.main === module) {
  startServer();
}
