import express from 'express';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(__dirname, '..');

const DATA_DIR = process.env.DATA_DIR || path.join(ROOT, 'data');
const DATA_FILE = path.join(DATA_DIR, 'locations.json');

if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });
if (!fs.existsSync(DATA_FILE)) fs.writeFileSync(DATA_FILE, '[]');

const app = express();
app.use(express.json({ limit: '50mb' }));

// GET current saved locations (all uploaded reports + their parsed data)
app.get('/api/locations', (req, res) => {
  try {
    const raw = fs.readFileSync(DATA_FILE, 'utf-8');
    res.json(JSON.parse(raw));
  } catch (err) {
    console.error('Failed to read data file:', err);
    res.status(500).json({ error: 'Failed to read saved data' });
  }
});

// Replace saved locations wholesale (the frontend sends the full array on every change)
app.post('/api/locations', (req, res) => {
  try {
    fs.writeFileSync(DATA_FILE, JSON.stringify(req.body, null, 2));
    res.json({ ok: true });
  } catch (err) {
    console.error('Failed to write data file:', err);
    res.status(500).json({ error: 'Failed to save data' });
  }
});

app.get('/api/health', (req, res) => res.json({ status: 'ok' }));

// Serve the built frontend
const DIST_DIR = path.join(ROOT, 'dist');
app.use(express.static(DIST_DIR));
app.get('*', (req, res) => {
  res.sendFile(path.join(DIST_DIR, 'index.html'));
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Shelf Signal listening on port ${PORT}`);
  console.log(`Data file: ${DATA_FILE}`);
});
