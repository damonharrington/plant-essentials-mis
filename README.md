# 🌿 Plant Essentials — Investor MIS Portal

A production-ready Management Information System built for a D2C food & beverage startup. Designed to give founders and investors a real-time view of revenue, margins, and operational health — all from the browser, no backend required.

Built by **Damon Harrington** as part of an AI internship application.

---

## What It Does

This portal replaces messy spreadsheets with a structured, interactive dashboard that tracks the full financial picture of a multi-channel D2C business — from raw sales invoices all the way through to EBITDA.

### Modules

| Module | Description |
|---|---|
| **Dashboard** | Live KPIs — Revenue, Gross Margin, EBITDA, CM1/CM2. SKU-level projections (n+1, n+2, n+3 months) with weighted growth model. 7-day and 30-day moving averages computed from invoice data. |
| **Upload** | Drag-and-drop Excel/CSV ingestion for any sales channel. Smart column auto-mapping with memory — upload once, it remembers your format. Auto-generates invoices from raw data. |
| **Sales Ops** | Full invoice ledger across 8 channels (HORECA, Quick Commerce, E-Commerce, Physical, B2B, Vending, Community, Website). Manual invoice creation. Weekly revenue reports by week. |
| **MIS Input** | Manual entry for OPEX line items, COGS, marketing spend, and channel revenue. Saves per month, per financial year. |
| **P&L Statement** | Full income statement: Revenue → Gross Profit → CM1 → CM2 → EBITDA, with % of revenue for every line item. |

---

## Tech Stack

| Layer | Technology |
|---|---|
| Frontend | React 18 + Vite 5 |
| Charts | Recharts |
| Excel I/O | SheetJS (xlsx) |
| Server | Nginx 1.25 (Alpine) |
| Container | Docker multi-stage build |
| Storage | Browser localStorage (no backend needed) |

---

## Quick Start

### Option 1 — Run locally (fastest)

```bash
git clone https://github.com/damonharrington/plant-essentials-mis.git
cd plant-essentials-mis
npm install
npm run dev
```
Open [http://localhost:3000](http://localhost:3000)

### Option 2 — Docker

```bash
docker compose up --build -d
```
Open [http://localhost:3000](http://localhost:3000)

---

## Project Structure

```
plant-essentials-mis/
├── Dockerfile              # Multi-stage: Node 20 (build) → Nginx (serve)
├── docker-compose.yml      # One-command deployment
├── nginx.conf              # SPA routing + gzip + security headers
├── package.json
├── vite.config.js
├── index.html
└── src/
    ├── main.jsx            # React entry point
    └── App.jsx             # Full MIS portal (~1,100 lines)
```

---

## Key Design Decisions

- **No backend** — All data lives in `localStorage`, making this fully portable and instantly deployable anywhere.
- **Multi-channel architecture** — Built around 8 distinct sales channels, each with its own upload mapping memory and revenue tracking.
- **SKU forecasting model** — Revenue projections use a weighted blend of SKU velocity (50%) and channel momentum (50%), with a 3% monthly floor.
- **Single-file component** — The entire portal lives in `App.jsx` for easy review and portability, while still being cleanly structured with named sections.

---

## Data Persistence

| localStorage Key | Contents |
|---|---|
| `oatey-mis:FY_2025-26` | Monthly MIS inputs (OPEX, COGS, marketing) |
| `oatey-sales:FY_2025-26` | Invoice ledger |
| `oatey-maps` | Saved column mappings per channel |

Use the built-in **Export / Import** feature to back up or transfer data between browsers.

---

*Plant Essentials Pvt Ltd — Confidential*
