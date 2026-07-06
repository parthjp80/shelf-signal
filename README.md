# Shelf Signal

Cooler inventory intelligence across multiple locations. Upload a sales/inventory-sold
report (CSV or Excel) per location and get automated SKU-level recommendations
(top performers, cut candidates, watch list) plus a cross-location rollup that tiers
your locations by revenue and suggests where premium vs. value SKUs belong.

**Key features:**
- Auto-detects CSV/Excel columns (SKU, product, category, units sold, revenue, cost)
- Reads **all tabs** in multi-sheet Excel files; auto-picks the data-rich sheet
- Flags each SKU as a top performer, steady, watch-list, or cut candidate based on velocity & revenue contribution
- Categorizes products with no category column via an internet lookup (Open Food Facts) with a local keyword fallback
- Surfaces a **Suggestions** panel per location — concrete, data-backed calls to action (cut candidates, thin-margin items, category concentration risk, etc.), separate from the raw per-SKU flag table
- **Download PDF**: exports the shelf visualization, category mix, recommendations table, and suggestions as a shareable report (per-location or the cross-location rollup), titled with your business name and an editable reporting period
- Builds cross-location rollups with tier strategy (High/Mid/Value) and expansion/consolidation suggestions
- Persists all uploaded reports to a JSON file on a mounted Docker volume — survives restarts and is accessible from any device on your network

## Quick start (any Docker host)

```bash
unzip shelf-signal-git.zip
cd shelf-signal
docker compose up -d --build
```

Then visit `http://<your-server-ip>:3000` in a browser. Upload your first report and Shelf Signal will auto-detect the columns and pick the tab with the most rows (usually your detailed data, not a summary tab). If it picks the wrong tab, use the "Tab" dropdown above the table to switch.

## Running on TrueNAS SCALE

**Option A — SSH / shell, using docker compose directly (recommended)**
1. Copy this folder onto your TrueNAS box (e.g. via SMB share or `scp`), for example to
   `/mnt/YOUR_POOL/apps/shelf-signal`.
2. Edit `docker-compose.yml` and point the volume at a real dataset path instead of `./data`, e.g.:
   ```yaml
   volumes:
     - /mnt/YOUR_POOL/apps/shelf-signal/data:/app/data
   ```
3. From that directory, run:
   ```bash
   docker compose up -d --build
   ```
4. Visit `http://<truenas-ip>:3000` from any device on your network.
5. Upload your first sales/inventory report. If your Excel file has multiple tabs (e.g. Summary + Details), Shelf Signal reads all of them and automatically uses the tab with the most rows. You can switch tabs using the "Tab" dropdown if needed.

**Option B — TrueNAS SCALE "Custom App" (Apps UI)**
1. Build and push the image somewhere TrueNAS can pull it (a local registry, Docker Hub,
   or GHCR), or build it once locally with `docker build -t shelf-signal .` and load it
   onto the TrueNAS Docker host.
2. In the TrueNAS Apps UI, create a Custom App:
   - Image: `shelf-signal:latest`
   - Port: container port `3000` → whatever host port you want (e.g. `3000`)
   - Storage: mount a host path/dataset to `/app/data` (this is where reports are saved)
3. Deploy, then visit `http://<truenas-ip>:<port>`.

## Local development (without Docker)

```bash
npm install
npm run dev        # starts Vite dev server on :5173, proxies /api to :3000
# in a second terminal:
npm run build && npm start   # or just: node server/server.js (after a build)
```

## How it works

- **Frontend** (`src/App.jsx`): React + Vite. Parses CSV/XLSX client-side (Papaparse / SheetJS):
  - Reads **every tab** in multi-sheet Excel files
  - Auto-detects columns (SKU, product, category, units sold, revenue, unit cost)
  - Auto-picks the sheet with the most rows (typically your detailed data, not a summary)
  - Shows a "Tab" dropdown if the file has multiple sheets — switch manually if needed
  - Lets you fix column mappings if auto-detect misses something
- **Backend** (`server/server.js`): a small Express server that serves the built frontend
  and exposes these endpoints:
  - `GET /api/locations` — returns all saved location reports
  - `POST /api/locations` — overwrites the saved data (the frontend sends the full set
    whenever something changes)
  - `POST /api/categorize` — looks up categories for product names with no mapped
    category column, via [Open Food Facts](https://world.openfoodfacts.org) (free, no
    API key). Results are cached to `/app/data/category-cache.json` so the same product
    name is never looked up twice. Requires the container to have outbound internet
    access; if it doesn't, this silently falls back to the local keyword guesser.
  - Data is stored as plain JSON at `/app/data/locations.json` inside the container — back
    up that file (or the whole mounted volume) if you want a copy outside the container.

## File format & column mapping

Shelf Signal expects your CSV or Excel file to have columns like:
- **SKU / Item #** — unique product identifier
- **Product Name** — human-readable product name
- **Category** — category/department (optional but helpful for segmentation)
- **Units Sold** — quantity sold in the period
- **Revenue** — total sales dollars (or net sales)
- **Unit Cost** (optional) — cost per unit, for margin estimation

Column names don't have to match exactly — Shelf Signal searches for keywords
(e.g. "Item Desc" matches "Product Name", "Qty" matches "Units Sold"). If auto-detect
doesn't work, open the "Column mapping" panel and map each field manually.

**Multi-sheet Excel files:** If your export has a Summary tab and a detailed tab,
upload the whole file — Shelf Signal reads all sheets and defaults to the one with
the most rows. Use the "Tab" dropdown to switch if it picks the wrong one.

## Notes

- The business name shown on PDF reports is set via the `COMPANY_NAME` constant near the
  top of `src/App.jsx` — edit it there if it ever needs to change.
- No authentication is built in — if you expose this beyond your home network (e.g. via
  a reverse proxy), put it behind something like Authelia, Tailscale, or basic auth.
- Uploaded reports are stored as JSON in `/app/data/locations.json` — back up this file
  regularly or mount a backup-aware volume if you plan to rely on this for production use.
- The app is optimized for cooler/vending inventory (SKU-level velocity & revenue flagging),
  but works for any retail inventory file with similar structure.
- The `data/` folder in this repo is just a placeholder for local `docker compose` runs;
  on TrueNAS, use a real dataset path as shown above.
