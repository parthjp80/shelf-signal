# Shelf Signal

Cooler inventory intelligence across multiple locations. Upload a sales/inventory-sold
report (CSV or Excel) per location and get automated SKU-level recommendations
(top performers, cut candidates, watch list) plus a cross-location rollup that tiers
your locations by revenue and suggests where premium vs. value SKUs belong.

Data persists to a JSON file on a mounted volume, so it survives container restarts
and rebuilds, and is reachable from any device on your home network.

## Quick start (any Docker host)

```bash
docker compose up -d --build
```

Then visit `http://<your-server-ip>:3000`.

## Running on TrueNAS SCALE

**Option A — SSH / shell, using docker compose directly**
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
4. Visit `http://<truenas-ip>:3000`.

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

- **Frontend** (`src/`): React + Vite. Parses CSV/XLSX client-side (Papaparse / SheetJS),
  auto-detects columns (SKU, product, category, units sold, revenue, unit cost), and lets
  you fix the mapping if it guesses wrong.
- **Backend** (`server/server.js`): a small Express server that serves the built frontend
  and exposes two endpoints:
  - `GET /api/locations` — returns all saved location reports
  - `POST /api/locations` — overwrites the saved data (the frontend sends the full set
    whenever something changes)
  Data is stored as plain JSON at `/app/data/locations.json` inside the container — back
  up that file (or the whole mounted volume) if you want a copy outside the container.

## Notes

- No authentication is built in — if you expose this beyond your home network (e.g. via
  a reverse proxy), put it behind something like Authelia, Tailscale, or basic auth.
- The `data/` folder in this repo is just a placeholder for local `docker compose` runs;
  on TrueNAS, use a real dataset path as shown above.
