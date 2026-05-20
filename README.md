# 📦 Bin Tag Generator

A browser-based tool that reads a warehouse Put Away Excel file and automatically generates **PDF bin tags** — one per row — matching your Bin Tag template format.

---

## Features

- Upload `.xlsx` / `.xls` Put Away sheets directly in the browser
- Reads data from **Put Away** sheet and shipment date from **SRV** sheet
- Generates one A4 landscape PDF page per bin tag row
- Matches your template: SRV#, Invoice#, Container#, SKU, Batch/Lot No, Item Description, Date of Expiry, QTY Received, Shipment Received On
- No server required — runs entirely in the browser

---

## Bin Tag Layout (mirrors Excel template)

```
┌──────────────────────────────────────────────────┐
│                   SRV# 26100                     │
├─────────┬──────────┬────────┬───────────┬────────┤
│Invoice# │4882028881│ 26100  │Container# │ 51770  │
├─────────┼──────────┼────────┼───────────┼────────┤
│  SKU    │Batch/Lot │  Item  │  Expiry   │  QTY   │
│         │    No    │ Descr. │           │Received│
├─────────┼──────────┼────────┼───────────┼────────┤
│ ASOU386 │   DV9F   │PARODO… │ 27/6/2028 │  420   │
├──────────────────────────────────────────────────┤
│                                                  │
├───────────────────────────┬──────────────────────┤
│   Shipment Received On    │      24/3/2026        │
└───────────────────────────┴──────────────────────┘
```

---

## Project Structure

```
bin-tag-generator/
├── index.html   ← Main app page
├── style.css    ← Styling
├── app.js       ← All logic: Excel parsing + PDF generation
└── README.md    ← This file
```

---

## Deploying to GitHub Pages

1. Create a new GitHub repository (e.g. `bin-tag-generator`)
2. Push these three files to the `main` branch
3. Go to **Settings → Pages**
4. Set **Source** to `main` branch, root folder (`/`)
5. Click **Save** — your site will be live at:
   ```
   https://<your-username>.github.io/bin-tag-generator/
   ```

---

## Excel File Requirements

Your `.xlsx` must contain these sheets:

### SRV sheet
| Cell | Value |
|------|-------|
| B2   | Shipment Received Date |
| B3   | Invoice Number |
| H3   | SRV Number |
| M3   | Container / Seal Number |

### Put Away sheet
Must have a header row containing exactly these columns:
- `Sl #`
- `SKU`
- `Batch/     Lot No` (or any column containing "Batch")
- `MFG Date` (any column containing "MFG")
- `Date of Expiry` (any column containing "Expiry")
- `Description of Goods` (any column containing "Description")
- `Counted Qnty` (any column containing "Qnty", "Qty", or "QTY")
- `Location` (optional)

---

## Libraries Used (loaded from CDN — no install needed)

| Library | Purpose |
|---------|---------|
| [SheetJS (xlsx)](https://sheetjs.com/) | Parse Excel files in the browser |
| [jsPDF](https://github.com/parallax/jsPDF) | Generate PDF files in the browser |

---

## Local Testing

Just open `index.html` in any modern browser. No build step, no server needed.

```bash
# Optional: serve locally with Python
python3 -m http.server 8080
# then open http://localhost:8080
```
