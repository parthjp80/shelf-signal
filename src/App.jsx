import React, { useState, useEffect, useMemo, useRef } from 'react';
import * as XLSX from 'xlsx';
import Papa from 'papaparse';
import {
  Upload, Trash2, TrendingUp, AlertTriangle, CheckCircle2, MapPin,
  Sparkles, PackageX, ChevronDown, ChevronRight, ClipboardCopy, RefreshCw, Layers
} from 'lucide-react';

/* ---------------------------------- helpers ---------------------------------- */

const FIELD_KEYWORDS = {
  sku: ['sku', 'item number', 'item id', 'upc', 'plu', 'item no', 'item #', 'product barcode', 'barcode'],
  product: ['product name', 'product', 'item description', 'description', 'item name', 'name'],
  category: ['category', 'dept', 'department', 'type', 'segment'],
  brand: ['brand', 'manufacturer', 'maker'],
  units: ['units sold', 'qty sold', 'quantity sold', 'units', 'qty', 'quantity', 'sold', 'sales qty'],
  revenue: ['net sales', 'total sales', 'revenue', 'sales $', 'sales', 'amount', 'total revenue', 'gross sales', 'sales amount'],
  cost: ['unit cost', 'cost', 'cogs', 'cost per unit', 'unit price'],
};

const FIELD_LABELS = {
  sku: 'SKU / Item #',
  product: 'Product name',
  category: 'Category (optional)',
  brand: 'Brand (use if no category)',
  units: 'Units sold',
  revenue: 'Revenue',
  cost: 'Unit cost (optional)',
};

const PALETTE = ['#2E6F7E', '#4C7A5E', '#D6912B', '#C0512B', '#5B6F9E', '#8B5E83', '#7C8452', '#B79A6B'];

const FLAG_META = {
  top:    { label: 'Top performer', color: '#3F7A5C', bg: '#E9F2EC' },
  steady: { label: 'Steady',        color: '#2E6F7E', bg: '#EAF1F2' },
  watch:  { label: 'Watch',         color: '#B4791F', bg: '#FBF0DD' },
  cut:    { label: 'Cut candidate', color: '#B23E1F', bg: '#FBEAE4' },
};

function normalizeHeader(h) {
  return String(h || '').toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
}

function guessMapping(headers) {
  const used = new Set();
  const mapping = {};
  ['sku', 'product', 'category', 'brand', 'units', 'revenue', 'cost'].forEach((field) => {
    const keywords = FIELD_KEYWORDS[field];
    let found = '';
    for (const h of headers) {
      if (used.has(h)) continue;
      const norm = normalizeHeader(h);
      if (keywords.some((k) => norm === k || norm.includes(k))) { found = h; break; }
    }
    if (found) used.add(found);
    mapping[field] = found;
  });
  return mapping;
}

function parseNumber(v) {
  if (v === null || v === undefined || v === '') return 0;
  if (typeof v === 'number') return v;
  const s = String(v).replace(/[$,]/g, '').replace(/\((.*)\)/, '-$1').trim();
  const n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}

function hashString(str) {
  let hash = 0;
  for (let i = 0; i < str.length; i++) hash = (hash * 31 + str.charCodeAt(i)) | 0;
  return Math.abs(hash);
}

function categoryColor(category) {
  return PALETTE[hashString(category || 'Uncategorized') % PALETTE.length];
}

// Fallback classifier: infer a category from the product name itself when no
// category/brand column is available (or the cell is blank for that row).
// Order matters — more specific/branded keywords are listed before generic
// catch-alls (e.g. "snickers" before the bare "bar" in Bars) so a product
// like "Snickers Bar" resolves to Candy rather than Bars.
const PRODUCT_CATEGORY_KEYWORDS = {
  'Drinks': ['soda', 'cola', 'pepsi', 'coke', 'sprite', 'juice', 'drink', 'beverage', 'tea', 'lemonade',
    'gatorade', 'powerade', 'water', 'sparkling', 'seltzer', 'energy drink', 'red bull', 'redbull',
    'monster', 'iced coffee', 'coffee', 'smoothie', 'kombucha'],
  'Chips': ['chip', 'crisp', 'tortilla', 'doritos', 'lays', "lay's", 'pringles', 'fritos', 'cheetos',
    'sun chips', 'popcorn', 'pretzel'],
  'Candy': ['candy', 'chocolate', 'gummy', 'gummies', 'snickers', 'skittles', 'm&m', "reese's", 'reeses',
    'twix', 'kitkat', 'kit kat', 'sour patch', 'starburst', 'mint', 'gum', 'licorice', 'jelly bean'],
  'Cookies': ['cookie', 'oreo', 'wafer'],
  'Crackers': ['cracker', 'ritz', 'goldfish', "cheez-it", 'cheez it'],
  'Nuts & Trail Mix': ['almond', 'peanut', 'cashew', 'trail mix', 'pistachio', 'nut'],
  'Jerky & Meat Snacks': ['jerky', 'meat stick', 'slim jim'],
  'Ice Cream & Frozen': ['ice cream', 'popsicle', 'frozen'],
  'Bakery': ['muffin', 'donut', 'doughnut', 'pastry', 'danish', 'bun'],
  'Bars': ['granola bar', 'protein bar', 'energy bar', 'snack bar', 'cereal bar', 'bar'],
};

// Matches keyword at a word boundary (start of a token), tolerating simple
// plurals like "donut" -> "donuts". Deliberately NOT a plain substring test:
// naive .includes() matched "cola" inside "chocolate" and "nut" inside "donuts".
function guessCategoryFromProduct(productName) {
  const norm = normalizeHeader(productName);
  if (!norm) return '';
  for (const [category, keywords] of Object.entries(PRODUCT_CATEGORY_KEYWORDS)) {
    for (const k of keywords) {
      const kn = normalizeHeader(k);
      if (kn && new RegExp(`(^|\\s)${kn}`).test(norm)) return category;
    }
  }
  return '';
}

const money = (v) => new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD', maximumFractionDigits: 0 }).format(v || 0);
const pct = (v) => `${((v || 0) * 100).toFixed(1)}%`;

// Reads every tab in a workbook (or the single table in a CSV) and returns
// them as separate "sheets" so the caller can pick which one holds real data.
// Handles reports with title/metadata rows before the actual header row.
function readWorkbookSheets(file) {
  return new Promise((resolve, reject) => {
    const name = file.name.toLowerCase();
    const reader = new FileReader();
    if (name.endsWith('.csv')) {
      reader.onload = (e) => {
        Papa.parse(e.target.result, {
          header: true,
          skipEmptyLines: true,
          complete: (res) => resolve([{ name: 'Sheet1', headers: res.meta.fields || [], rows: res.data }]),
          error: reject,
        });
      };
      reader.onerror = reject;
      reader.readAsText(file);
    } else {
      reader.onload = (e) => {
        try {
          const wb = XLSX.read(e.target.result, { type: 'binary' });
          const sheets = wb.SheetNames.map((sheetName) => {
            const sheet = wb.Sheets[sheetName];
            // Get all raw values to detect where headers actually are
            const range = XLSX.utils.decode_range(sheet['!ref'] || 'A1');
            const allRows = [];
            for (let r = range.s.r; r <= range.e.r; r++) {
              const row = [];
              for (let c = range.s.c; c <= range.e.c; c++) {
                const cellAddr = XLSX.utils.encode_cell({ r, c });
                const cell = sheet[cellAddr];
                row.push(cell ? cell.v : '');
              }
              allRows.push(row);
            }

            // Find the header row: first row where most cells have content and no row above looks like headers
            let headerRowIdx = 0;
            for (let i = 0; i < Math.min(5, allRows.length); i++) {
              const row = allRows[i];
              const nonEmptyCells = row.filter(c => c && String(c).trim().length > 0).length;
              // A header row typically has many non-empty cells and text (not numbers)
              if (nonEmptyCells > 3 && row.some(c => c && String(c).match(/[a-zA-Z]/))) {
                headerRowIdx = i;
                break;
              }
            }

            // Extract headers and data starting from the detected header row
            const headers = allRows[headerRowIdx].map(h => String(h || '').trim());
            const dataRows = allRows.slice(headerRowIdx + 1).map(row => {
              const obj = {};
              headers.forEach((h, i) => { obj[h] = row[i] || ''; });
              return obj;
            });

            return { name: sheetName, headers, rows: dataRows };
          });
          resolve(sheets);
        } catch (err) { reject(err); }
      };
      reader.onerror = reject;
      reader.readAsBinaryString(file);
    }
  });
}

function getActiveSheet(location) {
  return location.sheets.find((s) => s.name === location.activeSheet) || location.sheets[0];
}

function aggregateLocation(location) {
  const sheet = getActiveSheet(location);
  const rows = sheet.rows;
  const mapping = location.mappingsBySheet[sheet.name] || {};
  const map = new Map();
  let totalRevenue = 0;
  let totalUnits = 0;
  let hasCost = !!mapping.cost;
  let totalCost = 0;

  rows.forEach((row) => {
    const productRaw = mapping.product ? row[mapping.product] : (mapping.sku ? row[mapping.sku] : '');
    const product = String(productRaw || '').trim() || 'Unlisted item';
    const sku = mapping.sku ? String(row[mapping.sku] || '').trim() : '';
    // Use category if mapped, otherwise brand, otherwise guess from the product
    // name (e.g. "drink", "chips", "bars"), otherwise "Uncategorized"
    let category = '';
    if (mapping.category) {
      category = String(row[mapping.category] || '').trim();
    } else if (mapping.brand) {
      category = String(row[mapping.brand] || '').trim();
    }
    if (!category) category = guessCategoryFromProduct(product);
    if (!category) category = 'Uncategorized';
    const units = mapping.units ? parseNumber(row[mapping.units]) : 0;
    const revenue = mapping.revenue ? parseNumber(row[mapping.revenue]) : 0;
    const unitCost = mapping.cost ? parseNumber(row[mapping.cost]) : 0;
    const key = (sku || product).toLowerCase();

    if (!map.has(key)) map.set(key, { key, product, sku, category, units: 0, revenue: 0, cost: 0 });
    const entry = map.get(key);
    entry.units += units;
    entry.revenue += revenue;
    entry.cost += unitCost * units;
    if (entry.category === 'Uncategorized' && category !== 'Uncategorized') entry.category = category;

    totalRevenue += revenue;
    totalUnits += units;
    totalCost += unitCost * units;
  });

  const items = Array.from(map.values());
  const shares = items.map((i) => (totalRevenue > 0 ? i.revenue / totalRevenue : 0));
  const sortedShares = [...shares].sort((a, b) => a - b);

  function percentileRank(v) {
    if (!sortedShares.length) return 0;
    let idx = sortedShares.findIndex((s) => s >= v);
    if (idx === -1) idx = sortedShares.length - 1;
    return idx / (sortedShares.length - 1 || 1);
  }

  items.forEach((item, i) => {
    item.share = shares[i];
    item.margin = hasCost ? item.revenue - item.cost : null;
    const p = percentileRank(item.share);
    if (item.units === 0 || item.revenue === 0) item.flag = 'cut';
    else if (p >= 0.85) item.flag = 'top';
    else if (p <= 0.2 && item.share < 0.02) item.flag = 'watch';
    else item.flag = 'steady';
  });
  items.sort((a, b) => b.revenue - a.revenue);

  const catMap = {};
  items.forEach((i) => { catMap[i.category] = (catMap[i.category] || 0) + i.revenue; });
  const categoryMix = Object.entries(catMap)
    .map(([category, revenue]) => ({ category, revenue, share: totalRevenue > 0 ? revenue / totalRevenue : 0 }))
    .sort((a, b) => b.revenue - a.revenue);

  return {
    items, totalRevenue, totalUnits, skuCount: items.length, categoryMix, hasCost,
    totalMargin: hasCost ? totalRevenue - totalCost : null,
    rowCount: rows.length,
  };
}

function computeRollup(locations, analysisByLocation) {
  const list = locations
    .map((loc) => ({ id: loc.id, name: loc.name, ...analysisByLocation[loc.id] }))
    .sort((a, b) => b.totalRevenue - a.totalRevenue);

  const n = list.length;
  list.forEach((loc, idx) => {
    if (n >= 3) {
      if (idx < Math.ceil(n / 3)) loc.tier = 'High';
      else if (idx < Math.ceil((2 * n) / 3)) loc.tier = 'Mid';
      else loc.tier = 'Value';
    } else if (n === 2) {
      loc.tier = idx === 0 ? 'High' : 'Value';
    } else {
      loc.tier = 'High';
    }
  });

  const productMap = new Map();
  locations.forEach((loc) => {
    const analysis = analysisByLocation[loc.id];
    analysis.items.forEach((item) => {
      const pk = item.product.toLowerCase();
      if (!productMap.has(pk)) productMap.set(pk, { product: item.product, locations: {} });
      productMap.get(pk).locations[loc.id] = { revenue: item.revenue, share: item.share, flag: item.flag, name: loc.name };
    });
  });

  const allIds = locations.map((l) => l.id);
  const expansion = [];
  const universalCut = [];

  productMap.forEach((entry) => {
    const presentIn = Object.keys(entry.locations);
    const missingFrom = allIds.filter((id) => !presentIn.includes(id));
    const topIn = presentIn.filter((id) => entry.locations[id].flag === 'top');
    if (topIn.length && missingFrom.length && locations.length > 1) {
      expansion.push({
        product: entry.product,
        topInNames: topIn.map((id) => entry.locations[id].name),
        missingNames: missingFrom.map((id) => locations.find((l) => l.id === id)?.name).filter(Boolean),
      });
    }
    if (presentIn.length >= 2 && presentIn.every((id) => entry.locations[id].flag === 'cut')) {
      universalCut.push({ product: entry.product, inNames: presentIn.map((id) => entry.locations[id].name) });
    }
  });

  return { list, expansion: expansion.slice(0, 12), universalCut: universalCut.slice(0, 12) };
}

function buildSummaryText(location, analysis) {
  const lines = [];
  lines.push(`${location.name} — Inventory Summary`);
  lines.push(`Revenue: ${money(analysis.totalRevenue)} | Units sold: ${analysis.totalUnits} | SKUs: ${analysis.skuCount}`);
  if (analysis.hasCost) lines.push(`Estimated margin: ${money(analysis.totalMargin)}`);
  lines.push('');
  const top = analysis.items.filter((i) => i.flag === 'top').slice(0, 5);
  const cut = analysis.items.filter((i) => i.flag === 'cut');
  const watch = analysis.items.filter((i) => i.flag === 'watch');
  if (top.length) {
    lines.push('Top performers (protect facings / consider expanding):');
    top.forEach((i) => lines.push(`  - ${i.product} — ${money(i.revenue)} (${pct(i.share)} of revenue)`));
    lines.push('');
  }
  if (watch.length) {
    lines.push(`Watch list — low velocity, monitor for a cycle (${watch.length} items):`);
    watch.slice(0, 8).forEach((i) => lines.push(`  - ${i.product} — ${i.units} units, ${money(i.revenue)}`));
    lines.push('');
  }
  if (cut.length) {
    lines.push(`Cut candidates — zero or negligible sales (${cut.length} items):`);
    cut.slice(0, 10).forEach((i) => lines.push(`  - ${i.product}`));
    lines.push('');
  }
  const topCat = analysis.categoryMix[0];
  if (topCat) {
    lines.push(`Category mix is led by ${topCat.category} at ${pct(topCat.share)} of revenue.`);
    if (topCat.share > 0.4) lines.push(`Consider diversifying — this category may be overrepresented on the shelf.`);
  }
  return lines.join('\n');
}

/* ---------------------------------- component ---------------------------------- */

export default function App() {
  const [locations, setLocations] = useState([]);
  const [activeId, setActiveId] = useState(null);
  const [tab, setTab] = useState('detail');
  const [hydrated, setHydrated] = useState(false);
  const [mappingOpenFor, setMappingOpenFor] = useState(null);
  const [copiedId, setCopiedId] = useState(null);
  const [toast, setToast] = useState('');
  const fileInputRef = useRef(null);

  // Load saved locations from the backend on startup
  useEffect(() => {
    (async () => {
      try {
        const res = await fetch('/api/locations');
        if (res.ok) {
          const data = await res.json();
          setLocations(data || []);
          if (data && data.length) setActiveId(data[0].id);
        }
      } catch (e) {
        console.error('Failed to load saved locations', e);
      }
      setHydrated(true);
    })();
  }, []);

  // Persist to the backend whenever locations change
  useEffect(() => {
    if (!hydrated) return;
    fetch('/api/locations', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(locations),
    }).catch((e) => console.error('Failed to save locations', e));
  }, [locations, hydrated]);

  const analysisByLocation = useMemo(() => {
    const out = {};
    locations.forEach((loc) => { out[loc.id] = aggregateLocation(loc); });
    return out;
  }, [locations]);

  const rollup = useMemo(() => computeRollup(locations, analysisByLocation), [locations, analysisByLocation]);

  async function handleFiles(fileList) {
    const files = Array.from(fileList || []);
    for (const file of files) {
      try {
        const sheets = await readWorkbookSheets(file);
        if (!sheets.length) {
          setToast(`${file.name} has no readable sheets.`);
          setTimeout(() => setToast(''), 4000);
          continue;
        }
        // Auto-pick the sheet most likely to hold SKU-level data: the one
        // with the most rows (a "Summary" tab is usually much shorter).
        const defaultSheet = sheets.reduce((best, s) => (s.rows.length > best.rows.length ? s : best), sheets[0]);
        const mappingsBySheet = {};
        sheets.forEach((s) => { mappingsBySheet[s.name] = guessMapping(s.headers); });

        const id = `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
        const name = file.name.replace(/\.(csv|xlsx|xls)$/i, '');
        const newLoc = {
          id, name, fileName: file.name,
          sheets, activeSheet: defaultSheet.name, mappingsBySheet,
          uploadedAt: new Date().toISOString(),
        };
        setLocations((prev) => [...prev, newLoc]);
        setActiveId(id);
        setTab('detail');
        if (sheets.length > 1) {
          setToast(`${file.name} has ${sheets.length} tabs — using "${defaultSheet.name}". Switch tabs above the table if needed.`);
          setTimeout(() => setToast(''), 6000);
        }
      } catch (e) {
        setToast(`Couldn't read ${file.name} — check it's a valid CSV or Excel export.`);
        setTimeout(() => setToast(''), 4000);
      }
    }
  }

  function updateMapping(id, field, header) {
    setLocations((prev) => prev.map((l) => {
      if (l.id !== id) return l;
      const sheetName = l.activeSheet;
      return { ...l, mappingsBySheet: { ...l.mappingsBySheet, [sheetName]: { ...l.mappingsBySheet[sheetName], [field]: header } } };
    }));
  }

  function switchSheet(id, sheetName) {
    setLocations((prev) => prev.map((l) => (l.id === id ? { ...l, activeSheet: sheetName } : l)));
  }

  function renameLocation(id, name) {
    setLocations((prev) => prev.map((l) => (l.id === id ? { ...l, name } : l)));
  }

  function removeLocation(id) {
    setLocations((prev) => prev.filter((l) => l.id !== id));
    if (activeId === id) setActiveId(null);
  }

  function resetAll() {
    setLocations([]);
    setActiveId(null);
  }

  async function copySummary(location, analysis) {
    const text = buildSummaryText(location, analysis);
    try {
      await navigator.clipboard.writeText(text);
      setCopiedId(location.id);
      setTimeout(() => setCopiedId(null), 1800);
    } catch (e) { /* clipboard unavailable */ }
  }

  const activeLocation = locations.find((l) => l.id === activeId) || null;
  const activeAnalysis = activeLocation ? analysisByLocation[activeLocation.id] : null;
  const activeTierEntry = rollup.list.find((l) => l.id === activeId);

  return (
    <div className="ss-root">
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;600;700&family=IBM+Plex+Mono:wght@400;500;600&family=Inter:wght@400;500;600&display=swap');

        .ss-root {
          --bg: #F2F6F5;
          --panel: #FFFFFF;
          --ink: #142229;
          --ink-soft: #52666C;
          --line: #DCE6E3;
          --teal: #2E6F7E;
          --teal-deep: #1C4A54;
          font-family: 'Inter', sans-serif;
          color: var(--ink);
          background: var(--bg);
          flex: 1;
          min-height: 100vh;
          display: flex;
          flex-direction: column;
        }
        .ss-root * { box-sizing: border-box; }
        .ss-mono { font-family: 'IBM Plex Mono', monospace; }
        .ss-display { font-family: 'Space Grotesk', sans-serif; }

        .ss-topbar {
          display: flex; align-items: center; justify-content: space-between;
          padding: 18px 24px; background: var(--teal-deep); color: #EAF3F1;
        }
        .ss-wordmark { display: flex; align-items: center; gap: 10px; }
        .ss-wordmark-text { font-family: 'Space Grotesk', sans-serif; font-weight: 700; font-size: 19px; letter-spacing: 0.02em; }
        .ss-tagline { font-size: 12px; color: #A9CBC9; margin-top: 2px; }
        .ss-tabs { display: flex; gap: 4px; background: rgba(255,255,255,0.08); padding: 4px; border-radius: 8px; }
        .ss-tab {
          padding: 7px 14px; border-radius: 6px; font-size: 13px; font-weight: 600; cursor: pointer;
          color: #C9E2E0; border: none; background: transparent; transition: all 0.15s;
        }
        .ss-tab.active { background: #EAF3F1; color: var(--teal-deep); }

        .ss-body { display: flex; flex: 1; min-height: 0; }

        .ss-sidebar { width: 260px; background: #EAF1EF; border-right: 1px solid var(--line); display: flex; flex-direction: column; }
        .ss-sidebar-title { font-size: 11px; font-weight: 600; letter-spacing: 0.08em; text-transform: uppercase; color: var(--ink-soft); padding: 16px 16px 8px; }
        .ss-stop-list { flex: 1; overflow-y: auto; padding: 0 10px; }
        .ss-stop {
          display: flex; flex-direction: column; gap: 3px; padding: 10px 10px; border-radius: 8px; cursor: pointer;
          margin-bottom: 4px; border: 1px solid transparent;
        }
        .ss-stop:hover { background: #DFEAE7; }
        .ss-stop.active { background: var(--panel); border-color: var(--line); box-shadow: 0 1px 2px rgba(20,34,41,0.06); }
        .ss-stop-top { display: flex; align-items: center; justify-content: space-between; gap: 6px; }
        .ss-stop-name { font-size: 13.5px; font-weight: 600; color: var(--ink); overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
        .ss-stop-rev { font-size: 12px; color: var(--ink-soft); }
        .ss-stop-trash { opacity: 0; color: #A85A3D; }
        .ss-stop:hover .ss-stop-trash { opacity: 0.8; }
        .ss-tier-dot { width: 7px; height: 7px; border-radius: 50%; flex-shrink: 0; }

        .ss-sidebar-footer { padding: 12px; border-top: 1px solid var(--line); display: flex; flex-direction: column; gap: 8px; }
        .ss-add-btn {
          display: flex; align-items: center; justify-content: center; gap: 7px;
          padding: 10px; border-radius: 8px; border: 1.5px dashed #9CB8B3; background: transparent;
          color: var(--teal-deep); font-weight: 600; font-size: 13px; cursor: pointer;
        }
        .ss-add-btn:hover { background: #DFEAE7; }
        .ss-reset-btn {
          display: flex; align-items: center; justify-content: center; gap: 6px;
          font-size: 11.5px; color: var(--ink-soft); background: none; border: none; cursor: pointer; padding: 4px;
        }
        .ss-reset-btn:hover { color: #B23E1F; }

        .ss-main { flex: 1; overflow-y: auto; padding: 24px 28px; }

        .ss-empty { display: flex; flex-direction: column; align-items: center; justify-content: center; height: 100%; text-align: center; color: var(--ink-soft); gap: 10px; padding: 40px; }
        .ss-empty-icon { color: var(--teal); margin-bottom: 4px; }
        .ss-empty h3 { font-family: 'Space Grotesk', sans-serif; color: var(--ink); font-size: 18px; margin: 0; }
        .ss-empty p { max-width: 440px; font-size: 13.5px; line-height: 1.6; margin: 0; }
        .ss-empty-cols { display: flex; flex-wrap: wrap; gap: 6px; justify-content: center; margin-top: 6px; }
        .ss-chip { font-family: 'IBM Plex Mono', monospace; font-size: 11px; background: #EAF1EF; border: 1px solid var(--line); padding: 3px 8px; border-radius: 5px; color: var(--teal-deep); }
        .ss-cta { margin-top: 8px; display: flex; align-items: center; gap: 8px; background: var(--teal); color: white; border: none; padding: 10px 18px; border-radius: 8px; font-weight: 600; font-size: 13.5px; cursor: pointer; }
        .ss-cta:hover { background: var(--teal-deep); }

        .ss-header-row { display: flex; align-items: flex-start; justify-content: space-between; margin-bottom: 12px; gap: 16px; }
        .ss-loc-name-input {
          font-family: 'Space Grotesk', sans-serif; font-size: 22px; font-weight: 700; border: none; background: none;
          color: var(--ink); padding: 0; width: 100%; outline: none; border-bottom: 2px solid transparent;
        }
        .ss-loc-name-input:focus { border-bottom-color: var(--teal); }
        .ss-loc-sub { font-size: 12px; color: var(--ink-soft); margin-top: 2px; }
        .ss-header-actions { display: flex; align-items: center; gap: 8px; flex-shrink: 0; }
        .ss-icon-btn { display: flex; align-items: center; gap: 6px; font-size: 12.5px; font-weight: 600; color: var(--teal-deep); background: #EAF1EF; border: 1px solid var(--line); padding: 7px 12px; border-radius: 7px; cursor: pointer; }
        .ss-icon-btn:hover { background: #DFEAE7; }

        .ss-tier-badge { font-size: 11px; font-weight: 700; padding: 3px 9px; border-radius: 20px; text-transform: uppercase; letter-spacing: 0.04em; }

        .ss-sheet-row { display: flex; align-items: center; gap: 8px; margin-bottom: 18px; }
        .ss-sheet-row label { font-size: 11.5px; font-weight: 600; color: var(--ink-soft); display: flex; align-items: center; gap: 5px; }
        .ss-sheet-row select { padding: 6px 8px; border-radius: 6px; border: 1px solid var(--line); font-size: 12.5px; background: white; color: var(--ink); font-weight: 600; }

        .ss-kpi-row { display: grid; grid-template-columns: repeat(4, 1fr); gap: 12px; margin-bottom: 22px; }
        .ss-kpi { background: var(--panel); border: 1px solid var(--line); border-radius: 10px; padding: 14px 16px; }
        .ss-kpi-label { font-size: 11px; color: var(--ink-soft); text-transform: uppercase; letter-spacing: 0.06em; font-weight: 600; }
        .ss-kpi-value { font-family: 'IBM Plex Mono', monospace; font-size: 21px; font-weight: 600; margin-top: 4px; color: var(--ink); }

        .ss-section-title { font-family: 'Space Grotesk', sans-serif; font-weight: 600; font-size: 15px; margin: 26px 0 10px; display: flex; align-items: center; gap: 8px; }

        .ss-shelf { display: flex; height: 46px; border-radius: 8px; overflow: hidden; border: 1px solid var(--line); }
        .ss-shelf-seg { display: flex; align-items: center; justify-content: center; color: white; font-size: 10.5px; font-weight: 600; overflow: hidden; white-space: nowrap; text-shadow: 0 1px 1px rgba(0,0,0,0.25); }
        .ss-legend { display: flex; flex-wrap: wrap; gap: 10px; margin-top: 8px; }
        .ss-legend-item { display: flex; align-items: center; gap: 5px; font-size: 11.5px; color: var(--ink-soft); }
        .ss-legend-dot { width: 8px; height: 8px; border-radius: 2px; }

        .ss-flag-legend { display: flex; gap: 14px; margin: 4px 0 14px; flex-wrap: wrap; }
        .ss-flag-badge { display: inline-flex; align-items: center; gap: 5px; font-size: 11.5px; font-weight: 600; padding: 3px 9px; border-radius: 20px; }

        table.ss-table { width: 100%; border-collapse: collapse; background: var(--panel); border: 1px solid var(--line); border-radius: 10px; overflow: hidden; }
        table.ss-table th { text-align: left; font-size: 11px; text-transform: uppercase; letter-spacing: 0.05em; color: var(--ink-soft); background: #EAF1EF; padding: 9px 12px; font-weight: 600; }
        table.ss-table td { padding: 9px 12px; font-size: 13px; border-top: 1px solid var(--line); }
        table.ss-table tr:hover td { background: #F7FAF9; }
        .ss-num { font-family: 'IBM Plex Mono', monospace; text-align: right; }

        .ss-mapping-panel { background: #FBFAF6; border: 1px solid var(--line); border-radius: 10px; padding: 14px 16px; margin-bottom: 18px; }
        .ss-mapping-grid { display: grid; grid-template-columns: repeat(3, 1fr); gap: 10px; margin-top: 10px; }
        .ss-map-field label { font-size: 11px; font-weight: 600; color: var(--ink-soft); display: block; margin-bottom: 4px; }
        .ss-map-field select { width: 100%; padding: 7px 8px; border-radius: 6px; border: 1px solid var(--line); font-size: 12.5px; background: white; color: var(--ink); }

        .ss-summary-box { background: #142229; color: #DCEBE8; border-radius: 10px; padding: 16px 18px; font-family: 'IBM Plex Mono', monospace; font-size: 12px; line-height: 1.7; white-space: pre-wrap; margin-top: 10px; }

        .ss-rollup-bar-row { display: flex; align-items: center; gap: 10px; margin-bottom: 8px; }
        .ss-rollup-bar-label { width: 150px; font-size: 12.5px; font-weight: 600; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
        .ss-rollup-bar-track { flex: 1; background: #EAF1EF; border-radius: 5px; height: 20px; overflow: hidden; }
        .ss-rollup-bar-fill { height: 100%; border-radius: 5px 0 0 5px; }
        .ss-rollup-bar-val { width: 90px; text-align: right; font-family: 'IBM Plex Mono', monospace; font-size: 12px; }

        .ss-tier-blocks { display: grid; grid-template-columns: repeat(3, 1fr); gap: 12px; margin-bottom: 22px; }
        .ss-tier-block { border-radius: 10px; padding: 14px 16px; border: 1px solid var(--line); background: var(--panel); }
        .ss-tier-block h4 { margin: 0 0 6px; font-family: 'Space Grotesk', sans-serif; font-size: 14px; }
        .ss-tier-block p { margin: 0 0 8px; font-size: 12px; color: var(--ink-soft); line-height: 1.5; }
        .ss-tier-block .ss-tier-locs { font-size: 12px; font-weight: 600; }

        .ss-insight-list { list-style: none; padding: 0; margin: 0; display: flex; flex-direction: column; gap: 8px; }
        .ss-insight-item { background: var(--panel); border: 1px solid var(--line); border-radius: 8px; padding: 10px 14px; font-size: 12.5px; }
        .ss-insight-item b { color: var(--ink); }
        .ss-insight-item .ss-sub { color: var(--ink-soft); font-size: 11.5px; margin-top: 3px; }

        .ss-toast { position: fixed; bottom: 20px; left: 50%; transform: translateX(-50%); background: #142229; color: white; padding: 10px 18px; border-radius: 8px; font-size: 12.5px; z-index: 50; max-width: 480px; text-align: center; }
        .ss-hint { font-size: 11.5px; color: var(--ink-soft); margin-top: -4px; margin-bottom: 14px; }
        .ss-select-msg { display: flex; align-items: center; justify-content: center; height: 100%; color: var(--ink-soft); font-size: 14px; }
      `}</style>

      <div className="ss-topbar">
        <div className="ss-wordmark">
          <MapPin size={20} />
          <div>
            <div className="ss-wordmark-text">SHELF SIGNAL</div>
            <div className="ss-tagline">Cooler inventory intelligence across your locations</div>
          </div>
        </div>
        <div className="ss-tabs">
          <button className={`ss-tab ${tab === 'detail' ? 'active' : ''}`} onClick={() => setTab('detail')}>Location detail</button>
          <button className={`ss-tab ${tab === 'rollup' ? 'active' : ''}`} onClick={() => setTab('rollup')}>Cross-location rollup</button>
        </div>
      </div>

      <div className="ss-body">
        <div className="ss-sidebar">
          <div className="ss-sidebar-title">Locations ({locations.length})</div>
          <div className="ss-stop-list">
            {locations.map((loc) => {
              const tierEntry = rollup.list.find((l) => l.id === loc.id);
              const tierColor = tierEntry?.tier === 'High' ? '#3F7A5C' : tierEntry?.tier === 'Mid' ? '#B4791F' : '#8B5E83';
              const a = analysisByLocation[loc.id];
              return (
                <div key={loc.id} className={`ss-stop ${activeId === loc.id ? 'active' : ''}`} onClick={() => { setActiveId(loc.id); setTab('detail'); }}>
                  <div className="ss-stop-top">
                    <div style={{ display: 'flex', alignItems: 'center', gap: 6, minWidth: 0 }}>
                      {tierEntry && <span className="ss-tier-dot" style={{ background: tierColor }} />}
                      <span className="ss-stop-name">{loc.name}</span>
                    </div>
                    <Trash2 size={13} className="ss-stop-trash" onClick={(e) => { e.stopPropagation(); removeLocation(loc.id); }} />
                  </div>
                  <div className="ss-stop-rev">{money(a?.totalRevenue)} · {a?.skuCount} SKUs</div>
                </div>
              );
            })}
          </div>
          <div className="ss-sidebar-footer">
            <button className="ss-add-btn" onClick={() => fileInputRef.current?.click()}>
              <Upload size={15} /> Add report
            </button>
            {locations.length > 0 && (
              <button className="ss-reset-btn" onClick={() => { if (confirm('Clear all uploaded locations? This cannot be undone.')) resetAll(); }}>
                <RefreshCw size={12} /> Clear all data
              </button>
            )}
          </div>
        </div>

        <div className="ss-main">
          <input
            ref={fileInputRef}
            type="file"
            accept=".csv,.xlsx,.xls"
            multiple
            style={{ display: 'none' }}
            onChange={(e) => { handleFiles(e.target.files); e.target.value = ''; }}
          />

          {locations.length === 0 && (
            <div className="ss-empty">
              <Sparkles size={30} className="ss-empty-icon" />
              <h3>Upload your first location report</h3>
              <p>
                Drop in a sales / inventory-sold export (CSV or Excel) from any cooler location.
                If your file has multiple tabs, Shelf Signal reads all of them and defaults to
                the one with the most rows — usually the detailed data, not a summary tab.
              </p>
              <div className="ss-empty-cols">
                {['SKU', 'Product Name', 'Category', 'Units Sold', 'Revenue', 'Unit Cost (optional)'].map((c) => (
                  <span key={c} className="ss-chip">{c}</span>
                ))}
              </div>
              <button className="ss-cta" onClick={() => fileInputRef.current?.click()}>
                <Upload size={15} /> Upload a report
              </button>
            </div>
          )}

          {locations.length > 0 && tab === 'detail' && !activeLocation && (
            <div className="ss-select-msg">Select a location on the left to view its breakdown.</div>
          )}

          {tab === 'detail' && activeLocation && activeAnalysis && (
            <LocationDetail
              location={activeLocation}
              analysis={activeAnalysis}
              tierEntry={activeTierEntry}
              mappingOpen={mappingOpenFor === activeLocation.id}
              onToggleMapping={() => setMappingOpenFor(mappingOpenFor === activeLocation.id ? null : activeLocation.id)}
              onUpdateMapping={(field, header) => updateMapping(activeLocation.id, field, header)}
              onSwitchSheet={(sheetName) => switchSheet(activeLocation.id, sheetName)}
              onRename={(name) => renameLocation(activeLocation.id, name)}
              onCopy={() => copySummary(activeLocation, activeAnalysis)}
              copied={copiedId === activeLocation.id}
            />
          )}

          {tab === 'rollup' && (
            <RollupView locations={locations} rollup={rollup} />
          )}
        </div>
      </div>

      {toast && <div className="ss-toast">{toast}</div>}
    </div>
  );
}

/* ---------------------------------- subviews ---------------------------------- */

function LocationDetail({ location, analysis, tierEntry, mappingOpen, onToggleMapping, onUpdateMapping, onSwitchSheet, onRename, onCopy, copied }) {
  const tierColor = tierEntry?.tier === 'High' ? { c: '#3F7A5C', bg: '#E9F2EC' } : tierEntry?.tier === 'Mid' ? { c: '#B4791F', bg: '#FBF0DD' } : { c: '#8B5E83', bg: '#F2ECF1' };
  const sheet = getActiveSheet(location);
  const mapping = location.mappingsBySheet[sheet.name] || {};
  const unmapped = ['sku', 'product', 'units', 'revenue'].filter((f) => !mapping[f]);
  const hasMultipleSheets = location.sheets.length > 1;

  return (
    <div>
      <div className="ss-header-row">
        <div style={{ flex: 1, minWidth: 0 }}>
          <input className="ss-loc-name-input" value={location.name} onChange={(e) => onRename(e.target.value)} />
          <div className="ss-loc-sub">{location.fileName} · {sheet.rows.length} rows{hasMultipleSheets ? ` on "${sheet.name}"` : ''}</div>
        </div>
        <div className="ss-header-actions">
          {tierEntry && <span className="ss-tier-badge" style={{ color: tierColor.c, background: tierColor.bg }}>{tierEntry.tier} tier</span>}
          <button className="ss-icon-btn" onClick={onCopy}><ClipboardCopy size={13} /> {copied ? 'Copied!' : 'Copy summary'}</button>
          <button className="ss-icon-btn" onClick={onToggleMapping}>{mappingOpen ? <ChevronDown size={13} /> : <ChevronRight size={13} />} Column mapping</button>
        </div>
      </div>

      {hasMultipleSheets && (
        <div className="ss-sheet-row">
          <label><Layers size={13} /> Tab:</label>
          <select value={location.activeSheet} onChange={(e) => onSwitchSheet(e.target.value)}>
            {location.sheets.map((s) => (
              <option key={s.name} value={s.name}>{s.name} ({s.rows.length} rows)</option>
            ))}
          </select>
        </div>
      )}

      {unmapped.length > 0 && !mappingOpen && (
        <div className="ss-hint" style={{ color: '#B23E1F' }}>
          ⚠ {unmapped.map((f) => FIELD_LABELS[f]).join(', ')} not mapped — numbers below may be incomplete. Open Column mapping to fix.
        </div>
      )}

      {mappingOpen && (
        <div className="ss-mapping-panel">
          <div style={{ fontSize: 12, color: 'var(--ink-soft)' }}>Map each field to the matching column from your upload.</div>
          <div className="ss-mapping-grid">
            {Object.entries(FIELD_LABELS).map(([field, label]) => (
              <div className="ss-map-field" key={field}>
                <label>{label}</label>
                <select value={mapping[field] || ''} onChange={(e) => onUpdateMapping(field, e.target.value)}>
                  <option value="">— none —</option>
                  {sheet.headers.map((h) => <option key={h} value={h}>{h}</option>)}
                </select>
              </div>
            ))}
          </div>
        </div>
      )}

      <div className="ss-kpi-row">
        <div className="ss-kpi"><div className="ss-kpi-label">Revenue</div><div className="ss-kpi-value">{money(analysis.totalRevenue)}</div></div>
        <div className="ss-kpi"><div className="ss-kpi-label">Units sold</div><div className="ss-kpi-value">{analysis.totalUnits.toLocaleString()}</div></div>
        <div className="ss-kpi"><div className="ss-kpi-label">SKU count</div><div className="ss-kpi-value">{analysis.skuCount}</div></div>
        <div className="ss-kpi">
          <div className="ss-kpi-label">{analysis.hasCost ? 'Est. margin' : 'Avg rev / SKU'}</div>
          <div className="ss-kpi-value">{analysis.hasCost ? money(analysis.totalMargin) : money(analysis.totalRevenue / (analysis.skuCount || 1))}</div>
        </div>
      </div>

      <div className="ss-section-title"><Sparkles size={15} /> Shelf, by revenue share</div>
      <div className="ss-shelf">
        {analysis.items.slice(0, 40).map((item) => (
          <div
            key={item.key}
            className="ss-shelf-seg"
            title={`${item.product} — ${money(item.revenue)} (${pct(item.share)})`}
            style={{ flex: `${Math.max(item.share, 0.003)} 0 auto`, background: categoryColor(item.category) }}
          >
            {item.share > 0.05 ? item.product.slice(0, 16) : ''}
          </div>
        ))}
      </div>
      <div className="ss-legend">
        {analysis.categoryMix.map((c) => (
          <div className="ss-legend-item" key={c.category}>
            <span className="ss-legend-dot" style={{ background: categoryColor(c.category) }} />
            {c.category} · {pct(c.share)}
          </div>
        ))}
      </div>

      <div className="ss-section-title">Recommendations</div>
      <div className="ss-flag-legend">
        {Object.entries(FLAG_META).map(([k, v]) => (
          <span key={k} className="ss-flag-badge" style={{ color: v.color, background: v.bg }}>
            {k === 'top' ? <TrendingUp size={12} /> : k === 'cut' ? <PackageX size={12} /> : k === 'watch' ? <AlertTriangle size={12} /> : <CheckCircle2 size={12} />}
            {v.label}
          </span>
        ))}
      </div>
      <table className="ss-table">
        <thead>
          <tr>
            <th>Product</th><th>Category</th><th className="ss-num">Units</th><th className="ss-num">Revenue</th><th className="ss-num">Share</th><th>Flag</th>
          </tr>
        </thead>
        <tbody>
          {analysis.items.slice(0, 60).map((item) => (
            <tr key={item.key}>
              <td>{item.product}</td>
              <td>{item.category}</td>
              <td className="ss-num ss-mono">{item.units.toLocaleString()}</td>
              <td className="ss-num ss-mono">{money(item.revenue)}</td>
              <td className="ss-num ss-mono">{pct(item.share)}</td>
              <td><span className="ss-flag-badge" style={{ color: FLAG_META[item.flag].color, background: FLAG_META[item.flag].bg }}>{FLAG_META[item.flag].label}</span></td>
            </tr>
          ))}
        </tbody>
      </table>

      <div className="ss-section-title">Written summary</div>
      <div className="ss-summary-box">{buildSummaryText(location, analysis)}</div>
    </div>
  );
}

function RollupView({ locations, rollup }) {
  if (locations.length === 0) {
    return <div className="ss-select-msg">Upload at least one location to see the rollup.</div>;
  }
  const maxRevenue = Math.max(...rollup.list.map((l) => l.totalRevenue), 1);
  const tierGroups = { High: [], Mid: [], Value: [] };
  rollup.list.forEach((l) => tierGroups[l.tier]?.push(l));

  const tierCopy = {
    High: 'Highest revenue locations. Prioritize premium and net-new SKUs here first — these coolers can absorb higher price points and support experimentation before rolling out elsewhere.',
    Mid: 'Solid, steady performers. Keep the core mix stable; test one or two of the High-tier winners here before a full rollout.',
    Value: 'Lower revenue locations. Focus on tightening the SKU count, cutting duplicate facings, and leaning into value multi-packs rather than adding premium items.',
  };

  return (
    <div>
      {locations.length < 2 && (
        <div className="ss-hint" style={{ marginBottom: 16 }}>Add at least one more location for a full comparison — showing what's available so far.</div>
      )}

      <div className="ss-section-title">Revenue by location</div>
      {rollup.list.map((l) => (
        <div className="ss-rollup-bar-row" key={l.id}>
          <div className="ss-rollup-bar-label">{l.name}</div>
          <div className="ss-rollup-bar-track">
            <div className="ss-rollup-bar-fill" style={{ width: `${(l.totalRevenue / maxRevenue) * 100}%`, background: l.tier === 'High' ? '#3F7A5C' : l.tier === 'Mid' ? '#D6912B' : '#8B5E83' }} />
          </div>
          <div className="ss-rollup-bar-val">{money(l.totalRevenue)}</div>
        </div>
      ))}

      <div className="ss-section-title">Tier strategy</div>
      <div className="ss-tier-blocks">
        {['High', 'Mid', 'Value'].map((tier) => (
          tierGroups[tier].length > 0 && (
            <div className="ss-tier-block" key={tier}>
              <h4>{tier} tier</h4>
              <p>{tierCopy[tier]}</p>
              <div className="ss-tier-locs">{tierGroups[tier].map((l) => l.name).join(', ')}</div>
            </div>
          )
        ))}
      </div>

      <div className="ss-section-title">Expansion opportunities</div>
      {rollup.expansion.length === 0 ? (
        <div className="ss-hint">No clear cross-location gaps yet — add more locations for richer comparisons.</div>
      ) : (
        <ul className="ss-insight-list">
          {rollup.expansion.map((e, i) => (
            <li className="ss-insight-item" key={i}>
              <b>{e.product}</b> is a top performer at {e.topInNames.join(', ')} but isn't stocked at {e.missingNames.join(', ')}.
              <div className="ss-sub">Consider trialing it at the missing location(s).</div>
            </li>
          ))}
        </ul>
      )}

      <div className="ss-section-title">Universal cut candidates</div>
      {rollup.universalCut.length === 0 ? (
        <div className="ss-hint">No SKUs are underperforming across every location that carries them.</div>
      ) : (
        <ul className="ss-insight-list">
          {rollup.universalCut.map((c, i) => (
            <li className="ss-insight-item" key={i}>
              <b>{c.product}</b> shows zero/negligible sales everywhere it's stocked ({c.inNames.join(', ')}).
              <div className="ss-sub">Strong candidate to drop network-wide and free up facings for a better fit.</div>
            </li>
          ))}
        </ul>
      )}
    </div>
  );
}
