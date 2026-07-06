import express from 'express';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { guessCategoryFromProduct } from '../src/lib/categorize.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(__dirname, '..');

const DATA_DIR = process.env.DATA_DIR || path.join(ROOT, 'data');
const DATA_FILE = path.join(DATA_DIR, 'locations.json');
const CATEGORY_CACHE_FILE = path.join(DATA_DIR, 'category-cache.json');

if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });
if (!fs.existsSync(DATA_FILE)) fs.writeFileSync(DATA_FILE, '[]');

// Persistent cache of product name -> category, so repeated uploads/restarts
// don't keep re-querying Open Food Facts for names we've already resolved.
let categoryCache = {};
if (fs.existsSync(CATEGORY_CACHE_FILE)) {
  try { categoryCache = JSON.parse(fs.readFileSync(CATEGORY_CACHE_FILE, 'utf-8')); } catch { categoryCache = {}; }
} else {
  fs.writeFileSync(CATEGORY_CACHE_FILE, '{}');
}

function saveCategoryCache() {
  fs.writeFileSync(CATEGORY_CACHE_FILE, JSON.stringify(categoryCache, null, 2));
}

function titleCase(s) {
  return s.replace(/\b\w/g, (c) => c.toUpperCase());
}

// Looks up a product name against Open Food Facts (free, no API key) and
// buckets the result into our own category taxonomy where possible.
async function lookupProductCategory(name) {
  const url = `https://world.openfoodfacts.org/cgi/search.pl?json=1&page_size=1&search_terms=${encodeURIComponent(name)}`;
  const res = await fetch(url, { headers: { 'User-Agent': 'ShelfSignal/1.0 (inventory categorizer)' } });
  if (!res.ok) return null;
  const data = await res.json();
  const product = data.products && data.products[0];
  if (!product) return null;

  const descriptiveText = [product.categories, product.product_name, product.generic_name].filter(Boolean).join(' ');
  const bucketed = guessCategoryFromProduct(descriptiveText);
  if (bucketed) return bucketed;

  const tags = (product.categories_tags || []).map((t) => t.replace(/^en:/, '').replace(/-/g, ' ')).filter(Boolean);
  if (!tags.length) return null;
  return titleCase(tags[tags.length - 1]);
}

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

// Look up categories for product names via Open Food Facts, backed by an
// on-disk cache. Sequential (not parallel) to stay polite to the public API.
app.post('/api/categorize', async (req, res) => {
  const names = Array.isArray(req.body.names) ? req.body.names : [];
  const results = {};
  let cacheDirty = false;
  for (const name of names) {
    const key = String(name || '').trim().toLowerCase();
    if (!key) continue;
    if (Object.prototype.hasOwnProperty.call(categoryCache, key)) {
      results[name] = categoryCache[key];
      continue;
    }
    try {
      const category = await lookupProductCategory(name);
      categoryCache[key] = category;
      results[name] = category;
      cacheDirty = true;
    } catch (err) {
      console.error('Category lookup failed for', name, err.message);
      results[name] = null;
    }
  }
  if (cacheDirty) saveCategoryCache();
  res.json(results);
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
