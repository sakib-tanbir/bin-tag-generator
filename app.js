/* ─────────────────────────────────────────────────────────────────
   Bin Tag Generator — app.js
   Reads Put Away sheet + SRV sheet from an uploaded .xlsx file,
   then generates one PDF page per row (bin tag format).
───────────────────────────────────────────────────────────────── */

const { jsPDF } = window.jspdf;

// ── DOM refs ──────────────────────────────────────────────────────
const dropZone       = document.getElementById('dropZone');
const fileInput      = document.getElementById('fileInput');
const uploadSection  = document.getElementById('uploadSection');
const previewSection = document.getElementById('previewSection');
const progressSection= document.getElementById('progressSection');
const fileInfoEl     = document.getElementById('fileInfo');
const dataSummaryEl  = document.getElementById('dataSummary');
const btnReset       = document.getElementById('btnReset');
const btnGenerate    = document.getElementById('btnGenerate');
const progressText   = document.getElementById('progressText');

let parsedData = null; // holds { tags[], meta{} }

// ── Drag & drop ───────────────────────────────────────────────────
dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.classList.add('drag-over'); });
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));
dropZone.addEventListener('drop', e => {
  e.preventDefault();
  dropZone.classList.remove('drag-over');
  handleFile(e.dataTransfer.files[0]);
});
fileInput.addEventListener('change', e => handleFile(e.target.files[0]));
dropZone.addEventListener('click', e => {
  if (e.target !== fileInput) fileInput.click();
});

btnReset.addEventListener('click', resetUI);
btnGenerate.addEventListener('click', generatePDF);

// ── File handling ─────────────────────────────────────────────────
function handleFile(file) {
  if (!file) return;
  if (!file.name.match(/\.(xlsx|xls)$/i)) {
    alert('Please upload an Excel file (.xlsx or .xls).');
    return;
  }

  const reader = new FileReader();
  reader.onload = e => {
    try {
      parsedData = parseWorkbook(e.target.result);
      showPreview(file.name, parsedData);
    } catch (err) {
      alert('Error reading file: ' + err.message);
      console.error(err);
    }
  };
  reader.readAsArrayBuffer(file);
}

// ── Excel parsing ─────────────────────────────────────────────────
function parseWorkbook(arrayBuffer) {
  const wb = XLSX.read(arrayBuffer, { type: 'array', cellDates: true });

  // ── SRV sheet → shipment received date, SRV#, Invoice#, Container# ──
  const srvSheet = wb.Sheets['SRV'];
  if (!srvSheet) throw new Error('Sheet "SRV" not found in this workbook.');

  const srvData = XLSX.utils.sheet_to_json(srvSheet, { header: 1, raw: false, dateNF: 'DD/MM/YYYY' });

  // Row 1 (index 1): Date is at col B (index 1)
  const shipmentDate = getCellValue(srvSheet, 'B3') || '';
  // Row 2 (index 2): INV# col B, SRV# col H, Container col M (approx)
  const invoiceNo    = getCellValue(srvSheet, 'B4') || '';
  const srvNo        = getCellValue(srvSheet, 'H4') || '';
  const containerNo  = getCellValue(srvSheet, 'M4') || '';

  // ── Put Away sheet ──
  const paSheet = wb.Sheets['Put Away'];
  if (!paSheet) throw new Error('Sheet "Put Away" not found in this workbook.');

  const paData = XLSX.utils.sheet_to_json(paSheet, { header: 1, raw: false, dateNF: 'DD/M/YY' });

  // Find the header row (contains "SKU")
  let headerRowIdx = -1;
  for (let i = 0; i < paData.length; i++) {
    if (paData[i].some(c => String(c).trim() === 'SKU')) {
      headerRowIdx = i;
      break;
    }
  }
  if (headerRowIdx < 0) throw new Error('Could not find header row with "SKU" in Put Away sheet.');

  const headers = paData[headerRowIdx].map(h => String(h || '').trim());
  const colIdx = {
    sl:       headers.indexOf('Sl #'),
    sku:      headers.indexOf('SKU'),
    batch:    headers.findIndex(h => h.includes('Batch')),
    mfg:      headers.findIndex(h => h.includes('MFG')),
    expiry:   headers.findIndex(h => h.includes('Expiry')),
    desc:     headers.findIndex(h => h.includes('Description')),
    qty:      headers.findIndex(h => h.includes('Qnty') || h.includes('Qty') || h.includes('QTY')),
    location: headers.findIndex(h => h.includes('Location')),
  };

  const tags = [];
  for (let i = headerRowIdx + 1; i < paData.length; i++) {
    const row = paData[i];
    const sl  = String(row[colIdx.sl] || '').trim();
    const sku = String(row[colIdx.sku] || '').trim();
    // Stop at totals/footer rows
    if (!sku || sku === '' || /^put away/i.test(sku)) break;
    if (!sl || isNaN(Number(sl))) continue; // skip non-data rows

    tags.push({
      slNo:        sl,
      sku:         sku,
      batch:       String(row[colIdx.batch]  || '').trim(),
      mfgDate:     formatDate(row[colIdx.mfg]),
      expiry:      formatDate(row[colIdx.expiry]),
      description: String(row[colIdx.desc]   || '').trim(),
      qty:         String(row[colIdx.qty]     || '').trim(),
      location:    String(colIdx.location >= 0 ? (row[colIdx.location] || '') : '').trim(),
    });
  }

  if (tags.length === 0) throw new Error('No data rows found in Put Away sheet.');

  return {
    tags,
    meta: {
      shipmentDate: formatDateStr(shipmentDate),
      invoiceNo,
      srvNo,
      containerNo,
    }
  };
}
//new date code
const MONTH_NAMES = ['January','February','March','April','May','June',
                     'July','August','September','October','November','December'];

// Parse any date value coming from SheetJS into a JS Date, or null.
function parseToDate(val) {
  if (!val && val !== 0) return null;
  // Already a Date object (cell.t === 'd')
  //if (val instanceof Date) return isNaN(val) ? null : val;
  //new date parser 
  if (val instanceof Date) {
  if (isNaN(val)) return null;
  return new Date(Date.UTC(val.getUTCFullYear(), val.getUTCMonth(), val.getUTCDate(), 12, 0, 0));
}
   //new date parser
  const s = String(val).trim();
  if (!s) return null;
  // Numeric Excel serial
  if (!isNaN(s) && s !== '') {
    const serial = Number(s);
    return new Date(Math.round((serial - 25569) * 86400 * 1000) + 43200000); // noon UTC
  }
  // DD/MM/YYYY or D/M/YY style
  const parts = s.split('/');
  if (parts.length === 3) {
    let [d, m, y] = parts.map(Number);
    if (y < 100) y += 2000;
    return new Date(Date.UTC(y, m - 1, d));
  }
  // Fallback: native parse
  const d = new Date(s);
  return isNaN(d) ? null : d;
}

function getCellValue(sheet, addr) {
  const cell = sheet[addr];
  if (!cell) return '';
  if (cell.t === 'd' || cell.t === 'n') {
    if (cell.w) return cell.w.trim();
  }
  return String(cell.w || cell.v || '').trim();
}

function formatDate(val) {
  if (!val) return '';
  const s = String(val).trim();
  // Already DD/M/YY or DD/MM/YYYY — normalise to DD/MM/YYYY
  const parts = s.split('/');
  if (parts.length === 3) {
    let [d, m, y] = parts;
    if (y.length === 2) y = '20' + y;
    return `${d.padStart(2,'0')}/${m.padStart(2,'0')}/${y}`;
  }
  return s;
}

// Formats shipment date as DD/Month/YYYY  e.g. 01/June/2026
function formatDateStr(val) {
  if (!val) return '';
  const parts = String(val).trim().split('/');
  if (parts.length === 3) {
    let [d, m, y] = parts.map(Number);
    if (y < 100) y += 2000;
    return `${String(d).padStart(2,'0')}/${MONTH_NAMES[m - 1]}/${y}`;
  }
  return String(val).trim();
}
//new date code

// ── UI: Preview ───────────────────────────────────────────────────
function showPreview(fileName, data) {
  const { tags, meta } = data;

  fileInfoEl.innerHTML = `
    <strong>📄 ${escHtml(fileName)}</strong>
    SRV#: <b>${escHtml(meta.srvNo)}</b> &nbsp;|&nbsp;
    Invoice#: <b>${escHtml(meta.invoiceNo)}</b> &nbsp;|&nbsp;
    Container#: <b>${escHtml(meta.containerNo)}</b> &nbsp;|&nbsp;
    Shipment Date: <b>${escHtml(meta.shipmentDate)}</b>
  `;

  // Group by SKU for summary
  const skuMap = {};
  tags.forEach(t => {
    if (!skuMap[t.sku]) skuMap[t.sku] = { desc: t.description, count: 0 };
    skuMap[t.sku].count++;
  });

  let rows = '';
  Object.entries(skuMap).forEach(([sku, info]) => {
    rows += `<tr>
      <td>${escHtml(sku)}</td>
      <td>${escHtml(info.desc)}</td>
      <td style="text-align:center"><span class="tag-count">${info.count}</span></td>
    </tr>`;
  });

  dataSummaryEl.innerHTML = `
    <table>
      <thead><tr><th>SKU</th><th>Description</th><th>Tags</th></tr></thead>
      <tbody>${rows}</tbody>
    </table>
    <div class="summary-footer">Total bin tags to generate: <strong>${tags.length}</strong></div>
  `;

  uploadSection.style.display = 'none';
  previewSection.style.display = 'block';
}

function resetUI() {
  parsedData = null;
  fileInput.value = '';
  previewSection.style.display = 'none';
  progressSection.style.display = 'none';
  uploadSection.style.display = 'block';
}

// ── PDF Generation ────────────────────────────────────────────────
async function generatePDF() {
  if (!parsedData) return;
  const { tags, meta } = parsedData;

  previewSection.style.display = 'none';
  progressSection.style.display = 'block';
  progressText.textContent = `Generating ${tags.length} bin tag(s)…`;

  // Small delay so UI updates render
  await sleep(80);

  try {
    // A4 landscape for wide tags: 297 × 210 mm
    const pdf = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });
    const W = 297, H = 210;
    const margin = 12;
    const usable = W - margin * 2;

    for (let i = 0; i < tags.length; i++) {
      if (i > 0) pdf.addPage();
      progressText.textContent = `Generating tag ${i+1} of ${tags.length}…`;
      await sleep(0);
      drawBinTag(pdf, tags[i], meta, margin, margin, usable, H - margin * 2);
    }

    pdf.save(`BinTags_SRV${meta.srvNo || 'export'}.pdf`);
    progressText.textContent = `✅ Done! ${tags.length} bin tag(s) saved.`;
    await sleep(1500);
    previewSection.style.display = 'block';
    progressSection.style.display = 'none';
  } catch (err) {
    alert('PDF generation error: ' + err.message);
    console.error(err);
    previewSection.style.display = 'block';
    progressSection.style.display = 'none';
  }
}

/* ── Draw one bin tag onto the PDF canvas ──
   Mirrors the Excel Bin Tag template:
   Row 1 : SRV# (large, centred across page)
   Row 2 : Invoice# | SRV value | Container# | value
   Row 3 : Column headers (SKU | Batch/Lot No | Item Description | Date of Expiry | QTY Received)
   Row 4 : Data values (large fonts)
   Row 5 : gap
   Row 6 : Shipment Received On  |  date value
*/
function drawBinTag(pdf, tag, meta, x, y, w, h) {
  const FONT = 'helvetica';

  // ── Outer border ──────────────────────────────────────────────
  pdf.setDrawColor(0);
  pdf.setLineWidth(0.8);
  pdf.rect(x, y, w, h);

  // ── Column proportions (match Excel template) ─────────────────
  // 5 columns: SKU | Batch/Lot | Description | Expiry | QTY
  const colPcts = [0.14, 0.12, 0.40, 0.17, 0.17];
  const colW    = colPcts.map(p => p * w);
  const colX    = [];
  let cx = x;
  colW.forEach(cw => { colX.push(cx); cx += cw; });

  // ── Row heights (proportional to page height) ─────────────────
  const rowH = [
    h * 0.11,  // Row 0: SRV# title
    h * 0.13,  // Row 1: Invoice / SRV / Container
    h * 0.11,  // Row 2: column headers
    h * 0.38,  // Row 3: data values (tallest)
    h * 0.06,  // Row 4: spacer
    h * 0.21,  // Row 5: Shipment Received On
  ];
  const rowY = [];
  let ry = y;
  rowH.forEach(rh => { rowY.push(ry); ry += rh; });

  // helper: draw a cell with border, text, given style
  function cell(rx, ry, rw, rh, text, opts = {}) {
    const {
      fontSize  = 10,
      bold      = false,
      valign    = 'middle',
      halign    = 'center',
      wrap      = false,
      color     = [0, 0, 0],
      bgColor   = null,
      border    = true,
      paddingX  = 2,
    } = opts;

    if (bgColor) {
      pdf.setFillColor(...bgColor);
      pdf.rect(rx, ry, rw, rh, 'F');
    }
    if (border) {
      pdf.setDrawColor(0);
      pdf.setLineWidth(0.4);
      pdf.rect(rx, ry, rw, rh);
    }

    if (!text && text !== 0) return;

    pdf.setFont(FONT, bold ? 'bold' : 'normal');
    pdf.setFontSize(fontSize);
    pdf.setTextColor(...color);

    const textStr = String(text);
    const tx = halign === 'center' ? rx + rw / 2 : rx + paddingX;
    let ty;
    if (valign === 'middle') ty = ry + rh / 2 + fontSize * 0.18;
    else if (valign === 'top') ty = ry + fontSize * 0.45 + 1;
    else ty = ry + rh - 2;

    if (wrap) {
      const lines = pdf.splitTextToSize(textStr, rw - paddingX * 2);
      const lineH = fontSize * 0.45;
      const totalH = lines.length * lineH;
      let startY = ry + (rh - totalH) / 2 + lineH * 0.8;
      lines.forEach(line => {
        pdf.text(line, tx, startY, { align: halign });
        startY += lineH;
      });
    } else {
      pdf.text(textStr, tx, ty, { align: halign });
    }
  }

  // ── ROW 0: SRV# spanning full width ───────────────────────────
  cell(x, rowY[0], w, rowH[0],
    `SRV# ${meta.srvNo}`,
    { fontSize: 20, bold: true, halign: 'center' }
  );

  // ── ROW 1: Invoice# | value | SRV value | Container# | value ─
  const r1 = rowY[1], r1h = rowH[1];
  const seg1 = w / 5;
  cell(x + seg1*0, r1, seg1, r1h, 'Invoice#',          { fontSize: 11, bold: true });
  cell(x + seg1*1, r1, seg1, r1h, meta.invoiceNo,       { fontSize: 13, bold: true });
  cell(x + seg1*2, r1, seg1, r1h, meta.srvNo,           { fontSize: 18, bold: true });
  cell(x + seg1*3, r1, seg1, r1h, 'Container#',         { fontSize: 11, bold: true });
  cell(x + seg1*4, r1, seg1, r1h, meta.containerNo,     { fontSize: 13, bold: true });

  // ── ROW 2: Column headers ─────────────────────────────────────
  const headers = ['SKU', 'Batch/Lot No', 'Item Description', 'Date of Expiry', 'QTY Received'];
  const hFontSizes = [13, 12, 11, 10, 10];
  headers.forEach((h, ci) => {
    cell(colX[ci], rowY[2], colW[ci], rowH[2], h,
      { fontSize: hFontSizes[ci], bold: true, bgColor: [230, 240, 255] }
    );
  });

  // ── ROW 3: Data values ──────────────────────────────────────── 
  // ── ROW 3: Data values ────────────────────────────────────────
  // Columns 0,1,3,4 = vertical text | Column 2 (Description) = horizontal
  const dataValues = [tag.sku, tag.batch, tag.description, tag.expiry, tag.qty];

  // helper: draw a cell with vertical (90°) rotated text, auto-fit font
  function verticalCell(ci, maxFontSize, bold) {
    pdf.setDrawColor(0);
    pdf.setLineWidth(0.4);
    pdf.rect(colX[ci], rowY[3], colW[ci], rowH[3]);

    const val = String(dataValues[ci] || '');
    if (!val) return;

    pdf.setFont(FONT, bold ? 'bold' : 'normal');
    pdf.setTextColor(0, 0, 0);

    // Shrink font until text width fits within cell height, and font height fits cell width
    let fs = maxFontSize;
    while (fs > 6) {
      pdf.setFontSize(fs);
      const tw = pdf.getTextWidth(val);
      const fh = fs * 0.35; // approximate font body height in mm
      if (tw <= rowH[3] - 4 && fh <= colW[ci] - 2) break;
      fs -= 0.5;
    }

    const tw = pdf.getTextWidth(val);
    const fh = fs * 0.35;
    pdf.text(
      val,
      colX[ci] + colW[ci] / 2 + fh / 2,
      rowY[3] + rowH[3] / 2 + tw / 2,
      { align: 'left', angle: 90 }
    );
  }

  verticalCell(0, 36, true); // SKU
  verticalCell(1, 32, true); // Batch/Lot No
  // Description stays horizontal (col index 2)
  cell(colX[2], rowY[3], colW[2], rowH[3], tag.description,
    { fontSize: 18, bold: true, wrap: true }
  );
  verticalCell(3, 24, true); // Date of Expiry
  verticalCell(4, 38, true); // QTY Received

  // ── ROW 4: spacer (empty row with border) ────────────────────
  pdf.setDrawColor(0);
  pdf.setLineWidth(0.4);
  pdf.rect(x, rowY[4], w, rowH[4]);

  // ── ROW 5: Shipment Received On ───────────────────────────────
  const r5w1 = w * 0.55, r5w2 = w - r5w1;
  cell(x,        rowY[5], r5w1, rowH[5], 'Shipment Received On',
    { fontSize: 16, bold: true }
  );
  cell(x + r5w1, rowY[5], r5w2, rowH[5], meta.shipmentDate,
    { fontSize: 18, bold: true }
  );
}

// ── Helpers ───────────────────────────────────────────────────────
function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }
function escHtml(s) {
  return String(s || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}
