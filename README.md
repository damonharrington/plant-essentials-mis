<<<<<<< HEAD
# Plant Essentials — Investor MIS Portal

A production-ready Management Information System dashboard for Plant Essentials,
a D2C food/beverage startup. Built with React + Vite, served by Nginx, fully
containerised with Docker.

---

## Features

| Module | What it does |
|---|---|
| **Dashboard** | Revenue KPIs, EBITDA, gross margin, SKU projections (n+1/n+2/n+3), moving averages, geo heatmap |
| **Upload** | Drag-and-drop Excel/CSV ingestion per channel, smart column mapping, auto-invoice generation |
| **Sales Ops** | Invoice ledger, manual invoice creation, weekly revenue reports |
| **MIS Input** | Manual entry for OPEX, COGS, marketing spend, channel revenue |
| **P&L** | Full income statement: Revenue → Gross Profit → CM1 → CM2 → EBITDA |

---

## Tech Stack

- **Frontend**: React 18 + Vite 5
- **Charts**: Recharts
- **Excel I/O**: SheetJS (xlsx)
- **Server**: Nginx 1.25 (Alpine)
- **Container**: Docker multi-stage build (Node 20 builder → Nginx runner)
- **Storage**: Browser `localStorage` (persists across sessions, no backend needed)

---

## Prerequisites

- [Docker](https://docs.docker.com/get-docker/) ≥ 24
- [Docker Compose](https://docs.docker.com/compose/install/) ≥ 2 (bundled with Docker Desktop)

---

## Quick Start (Docker — recommended)

```bash
# 1. Clone / unzip the project
cd plant-essentials-mis

# 2. Build and start (first run ~2 min, subsequent runs use cache)
docker compose up --build -d

# 3. Open in browser
open http://localhost:3000
```

To stop:
```bash
docker compose down
```

To rebuild after code changes:
```bash
docker compose up --build -d
```

---

## Local Development (without Docker)

```bash
# Install dependencies
npm install

# Start dev server with hot reload
npm run dev
# → http://localhost:3000

# Build production bundle
npm run build

# Preview production build locally
npm run preview
```

---

## Project Structure

```
plant-essentials-mis/
├── Dockerfile              # Multi-stage: Node (build) → Nginx (serve)
├── docker-compose.yml      # One-command deployment
├── nginx.conf              # SPA routing + caching + security headers
├── package.json
├── vite.config.js
├── index.html
└── src/
    ├── main.jsx            # React entry point
    └── App.jsx             # Full MIS portal (single-file component)
```

---

## Port Configuration

The app runs on **port 3000** by default. To change it, edit `docker-compose.yml`:

```yaml
ports:
  - "8080:80"   # now accessible at http://localhost:8080
```

---

## Data Persistence

All data is stored in the browser's `localStorage`:

| Key | Contents |
|---|---|
| `oatey-mis:FY_2025-26` | Monthly MIS inputs (OPEX, COGS, marketing) |
| `oatey-sales:FY_2025-26` | Invoice ledger |
| `oatey-maps` | Saved column mappings per channel |

Use **Export / Import → Download JSON** to back up or transfer data between browsers.

---

## Deployment Notes

- The Docker image is **stateless** — data lives in the client browser.
- To deploy to a server: push the image to a registry (Docker Hub, ECR, GCR),
  pull on the server, and run `docker compose up -d`.
- For HTTPS, place Nginx or Traefik as a reverse proxy in front with a TLS cert.

---

## SKU Projection Model

Forecasts use a weighted blend:

```
Combined growth = 50% × SKU velocity + 50% × channel momentum
Floor: 3% / month
Revenue = Σ (projected SKU units × last known price)
```

Moving averages (7-day STMA, 30-day LTMA) are computed from daily invoice data
and displayed on the dashboard when 2+ days of invoices are present.

---

*Plant Essentials Pvt Ltd — Confidential*
=======
# plant-essentials-mis
Production ready Investor MIS Portal for a D2C startup - React 18, Vite, Recharts, Docker. Features P&amp;L, Sales Ops, SKU forecasting &amp; Excel ingestion.
>>>>>>> 4edaa9a862f95b52008e694f640b73c9b80b906c
