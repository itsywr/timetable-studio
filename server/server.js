import express from 'express';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const PROD = process.env.NODE_ENV === 'production';

// Pick a writable location that survives restarts:
//   1. DATA_DIR env (Fly sets this to /data)
//   2. /data if it exists and is writable (Fly volume)
//   3. project-root/.data (Glitch persistent folder, also fine locally)
//   4. server/ (dev fallback)
const DATA_DIR = (() => {
  const candidates = [
    process.env.DATA_DIR,
    '/data',
    path.join(__dirname, '..', '.data'),
    __dirname,
  ].filter(Boolean);
  for (const dir of candidates) {
    try {
      fs.mkdirSync(dir, { recursive: true });
      fs.accessSync(dir, fs.constants.W_OK);
      return dir;
    } catch {}
  }
  return __dirname;
})();
const DATA_FILE = path.join(DATA_DIR, 'state.json');

const PORT = Number(process.env.PORT) || 3001;

let state = null;
let savedAt = null;
try {
  if (fs.existsSync(DATA_FILE)) {
    state = JSON.parse(fs.readFileSync(DATA_FILE, 'utf-8'));
    const stat = fs.statSync(DATA_FILE);
    savedAt = stat.mtime.toISOString();
    console.log(`[state] loaded snapshot from ${DATA_FILE}`);
  } else {
    console.log('[state] no snapshot on disk yet');
  }
} catch (e) {
  console.error('[state] failed to load snapshot:', e.message);
}

const app = express();
app.use(express.json({ limit: '10mb' }));

app.get('/api/state', (_req, res) => {
  res.json({ state, savedAt });
});

// In-memory only — survives page reload, lost on server restart unless saved.
app.patch('/api/state', (req, res) => {
  state = req.body;
  res.json({ ok: true });
});

// Persist current state to disk.
app.post('/api/state/save', (req, res) => {
  if (req.body) state = req.body;
  try {
    fs.writeFileSync(DATA_FILE, JSON.stringify(state, null, 2));
    savedAt = new Date().toISOString();
    res.json({ ok: true, savedAt });
  } catch (e) {
    res.status(500).json({ ok: false, error: e.message });
  }
});

// If a built React app exists in dist/, serve it alongside the API.
// (Production on Fly/Glitch — dev runs Vite separately on 5173.)
const distDir = path.join(__dirname, '..', 'dist');
if (fs.existsSync(path.join(distDir, 'index.html'))) {
  app.use(express.static(distDir));
  app.get('*', (req, res, next) => {
    if (req.path.startsWith('/api/')) return next();
    res.sendFile(path.join(distDir, 'index.html'));
  });
  console.log(`[web] serving static from ${distDir}`);
}

app.listen(PORT, '0.0.0.0', () => console.log(`[api] listening on http://0.0.0.0:${PORT} (prod=${PROD}, data=${DATA_DIR})`));
