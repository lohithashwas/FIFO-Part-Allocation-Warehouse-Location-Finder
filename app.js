/* ============================================================
   WareFlow – FIFO Part Allocation Engine
   app.js
   ============================================================ */

'use strict';

// ─── Required columns per file ────────────────────────────────
const REQUIRED = {
  request:   ['Part Code', 'Part Name', 'Quantity', 'Requested Date', 'Destination Location'],
  inventory: ['Part Code', 'Quantity', 'Order Date', 'Location', 'Type Location'],
  container: ['Part No', 'Quantity', 'Order Date']
};

const FILE_LABELS = {
  request:   'Part Request List',
  inventory: 'Consolidated Inventory',
  container: 'Container List Case Details'
};

// ─── State ────────────────────────────────────────────────────
const state = {
  files: { request: null, inventory: null, container: null },
  data:  { requests: [], inventory: [], containers: [] },
  results: { fulfilled: [], shortages: [], summary: {} }
};

// ─── DATE HELPERS ─────────────────────────────────────────────
function parseExcelDate(val) {
  if (val == null || val === '') return new Date(0);
  if (val instanceof Date) return isNaN(val) ? new Date(0) : val;
  if (typeof val === 'number') {
    // Excel serial date (days since 1899-12-30)
    return new Date((val - 25569) * 86400 * 1000);
  }
  if (typeof val === 'string') {
    // "10-Jul-26" or "10-Jul-2026" or "2026-07-10"
    const cleaned = val.trim();
    // Try native parse first
    const d = new Date(cleaned);
    if (!isNaN(d)) return d;
    // "dd-Mon-yy" pattern
    const m = cleaned.match(/^(\d{1,2})[\/\-]([A-Za-z]+)[\/\-](\d{2,4})$/);
    if (m) {
      const year = m[3].length === 2 ? '20' + m[3] : m[3];
      return new Date(`${m[1]} ${m[2]} ${year}`);
    }
  }
  return new Date(0);
}

function formatDate(val) {
  if (!val) return '—';
  const d = (val instanceof Date) ? val : parseExcelDate(val);
  if (!d || isNaN(d)) return String(val);
  return d.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' });
}

// ─── UI HELPERS ───────────────────────────────────────────────
function showToast(msg, type = 'info', duration = 4000) {
  const c = document.getElementById('toast-container');
  const t = document.createElement('div');
  t.className = `toast ${type}`;
  const icon = type === 'success' ? '✅' : type === 'error' ? '❌' : 'ℹ️';
  t.innerHTML = `<span>${icon}</span><span>${msg}</span>`;
  c.appendChild(t);
  setTimeout(() => t.remove(), duration);
}

function setLoadingSub(txt) {
  document.getElementById('loading-sub').textContent = txt;
}

function updateProcessBtn() {
  const all = state.files.request && state.files.inventory && state.files.container;
  const btn = document.getElementById('process-btn');
  const status = document.getElementById('upload-status');
  btn.disabled = !all;
  const uploaded = [state.files.request, state.files.inventory, state.files.container].filter(Boolean).length;
  status.textContent = all ? 'All files ready — click Process to run FIFO allocation'
                           : `${uploaded}/3 files uploaded`;
}

// ─── DRAG & DROP SETUP ────────────────────────────────────────
['request', 'inventory', 'container'].forEach(type => {
  const zone = document.getElementById(`drop-${type}`);
  const input = document.getElementById(`file-${type}`);

  zone.addEventListener('dragover', e => { e.preventDefault(); zone.classList.add('drag-over'); });
  zone.addEventListener('dragleave', () => zone.classList.remove('drag-over'));
  zone.addEventListener('drop', e => {
    e.preventDefault();
    zone.classList.remove('drag-over');
    const file = e.dataTransfer.files[0];
    if (file) handleFile(type, file);
  });

  input.addEventListener('change', () => {
    if (input.files[0]) handleFile(type, input.files[0]);
  });
});

// ─── FILE HANDLING ────────────────────────────────────────────
function handleFile(type, file) {
  clearError(type);
  if (!file.name.match(/\.xlsx$/i)) {
    showError(type, 'Only .xlsx files are accepted.');
    return;
  }
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const wb = XLSX.read(e.target.result, { type: 'array', cellDates: true });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });

      if (rows.length === 0) {
        showError(type, 'The file appears to be empty or has no data rows.');
        return;
      }

      // Column validation
      const cols = Object.keys(rows[0]);
      const missing = REQUIRED[type].filter(r => !cols.includes(r));
      if (missing.length) {
        showError(type, `Missing required columns: ${missing.join(', ')}`);
        return;
      }

      state.files[type] = file;
      // Store parsed rows
      if (type === 'request')   state.data.requests   = rows;
      if (type === 'inventory') state.data.inventory  = rows;
      if (type === 'container') state.data.containers = rows;

      // Update UI
      document.getElementById(`fp-${type}-name`).textContent = file.name;
      document.getElementById(`fp-${type}-rows`).textContent = `${rows.length.toLocaleString()} rows`;
      document.getElementById(`fp-${type}`).classList.add('visible');
      document.getElementById(`card-${type}`).classList.add('has-file');
      document.getElementById(`card-${type}`).classList.remove('has-error');

      updateProcessBtn();
      showToast(`${FILE_LABELS[type]} loaded (${rows.length.toLocaleString()} rows)`, 'success');
    } catch (err) {
      showError(type, `Could not read file: ${err.message}`);
    }
  };
  reader.readAsArrayBuffer(file);
}

function removeFile(type) {
  state.files[type] = null;
  if (type === 'request')   state.data.requests   = [];
  if (type === 'inventory') state.data.inventory  = [];
  if (type === 'container') state.data.containers = [];

  document.getElementById(`fp-${type}`).classList.remove('visible');
  document.getElementById(`card-${type}`).classList.remove('has-file', 'has-error');
  document.getElementById(`file-${type}`).value = '';
  clearError(type);
  updateProcessBtn();
}

function showError(type, msg) {
  const el = document.getElementById(`err-${type}`);
  el.textContent = msg;
  el.classList.add('visible');
  document.getElementById(`card-${type}`).classList.add('has-error');
  document.getElementById(`card-${type}`).classList.remove('has-file');
}

function clearError(type) {
  const el = document.getElementById(`err-${type}`);
  el.textContent = '';
  el.classList.remove('visible');
}

// ─── FIFO ENGINE ──────────────────────────────────────────────
function buildSupplyPool(inventoryRows, containerRows) {
  // pool: Map<partCode, Array<batch>>
  const pool = new Map();

  const addBatch = (partCode, batch) => {
    if (!pool.has(partCode)) pool.set(partCode, []);
    pool.get(partCode).push(batch);
  };

  // Inventory
  for (const row of inventoryRows) {
    const code = String(row['Part Code'] || '').trim();
    if (!code) continue;
    // Exclude blocked rows
    const blocked = row['Blocked'];
    if (blocked === true || blocked === 1 || String(blocked).toUpperCase() === 'TRUE') continue;
    const qty = parseFloat(row['Quantity']) || 0;
    if (qty <= 0) continue;

    addBatch(code, {
      source:    'Inventory',
      orderDate: parseExcelDate(row['Order Date']),
      remaining: qty,
      location:  String(row['Location'] || '').trim(),
      pickLoc:   String(row['Type Location'] || row['Location'] || '').trim(),
      status:    String(row['Status'] || '').trim(),
      caseNo:    String(row['Case No'] || '').trim(),
      refNo:     String(row['Order No'] || '').trim(),
      plant:     String(row['Plant'] || '').trim(),
      partName:  String(row['Part Name'] || '').trim(),
      type:      String(row['Type'] || row['Pack Type'] || row['Packing Type'] || row['Packing'] || '').trim()
    });
  }

  // ── Containers
  // Add containers to FIFO pool UNLESS they are 'In Transit' or 'At Port'
  let lastContainerNo = '';

  for (const row of containerRows) {
    const code = String(row['Part No'] || '').trim();
    if (!code) continue;
    const qty = parseFloat(row['Quantity']) || 0;
    if (qty <= 0) continue;

    const caseNo = String(row['Case No'] || '').trim();
    const status = String(row['Status'] || '').trim();
    const destuff = String(row['Destuff Status'] || '').trim();
    const contLoc  = String(row['Container Location'] || '').trim();
    const loc = String(row['Location'] || '').trim();

    // Forward-fill container number
    const rawContainerNo = String(
      row['C/T No'] || row['C/T No.'] || row['Container No.'] ||
      row['Container No'] || row['Container Number'] || ''
    ).trim();
    if (rawContainerNo) lastContainerNo = rawContainerNo;
    const containerNo = lastContainerNo;

    // Is this container still in transit or at the port? (Check ONLY the Status column)
    const statusUpper = status.toUpperCase();
    const isTransitOrPort = statusUpper.includes('TRANSIT') || statusUpper.includes('PORT');

    // If it's in transit/port, skip adding it to the available FIFO pool
    // (It will be picked up by the shortage cross-check instead)
    if (isTransitOrPort) continue;

    const pickLoc = [caseNo ? `Case: ${caseNo}` : '', status ? `Status: ${status}` : '', destuff ? `Destuff: ${destuff}` : '']
                    .filter(Boolean).join(' | ') || contLoc || 'Container';

    addBatch(code, {
      source:      'Container',
      orderDate:   parseExcelDate(row['Order Date']),
      remaining:   qty,
      location:    loc,
      pickLoc,
      status,
      caseNo,
      containerNo,
      refNo:       String(row['Invoice No'] || '').trim(),
      portETA:     String(row['Port ETA'] || '').trim(),
      portATA:     String(row['Port ATA'] || '').trim(),
      destuffStatus: destuff,
      partName:    String(row['Part Name'] || '').trim(),
      type:        String(row['Type'] || row['Pack Type'] || row['Packing Type'] || row['Packing'] || '').trim()
    });
  }

  // Sort each part's batches by Order Date ASC (FIFO)
  for (const [, batches] of pool) {
    batches.sort((a, b) => a.orderDate - b.orderDate);
  }
  return pool;
}

// Build Combined Shortages report: for each shortage, find matching parts in container data
function buildCombinedShortages(shortages, containerRows) {
  if (!containerRows || containerRows.length === 0) return { combined: shortages, partsWithTransitCount: 0 };

  // Forward-fill container number in container rows
  let lastCT = '';
  const enriched = containerRows.map(row => {
    const raw = String(
      row['C/T No'] || row['C/T No.'] || row['Container No.'] ||
      row['Container No'] || row['Container Number'] || ''
    ).trim();
    if (raw) lastCT = raw;
    return { ...row, _ctNo: lastCT };
  });

  // Build map: Part No → container rows (ONLY the ones that are in transit/port)
  const ctMap = new Map();
  for (const row of enriched) {
    const code = String(row['Part No'] || '').trim();
    if (!code) continue;

    const status = String(row['Status'] || '').trim();
    
    // Only include in this report if it actually is in transit / at port (Check ONLY the Status column)
    const statusUpper = status.toUpperCase();
    const isTransitOrPort = statusUpper.includes('TRANSIT') || statusUpper.includes('PORT');
      
    if (!isTransitOrPort) continue; // It was already used in the main FIFO pool

    if (!ctMap.has(code)) ctMap.set(code, []);
    ctMap.get(code).push(row);
  }

  const combined = [];
  // Track which shortage parts have in-transit (for stat)
  const partsWithTransit = new Set();

  for (const shortage of shortages) {
    const partCode = shortage['Part Code'];
    const containers = ctMap.get(partCode) || [];
    
    // Calculate net status
    const shortQty = shortage['Shortage Quantity'];
    const totalTransit = containers.reduce((s, r) => s + (parseFloat(r['Quantity']) || 0), 0);
    const netShort = Math.max(0, shortQty - totalTransit);
    
    let netStatus = '';
    if (totalTransit === 0) {
      netStatus = `Absolute Shortage: ${fmt(shortQty)}`;
    } else if (netShort === 0) {
      netStatus = `Fully Covered by Transit`;
    } else {
      netStatus = `Short: ${fmt(netShort)} | In Transit: ${fmt(totalTransit)}`;
    }
    
    if (containers.length === 0) {
      combined.push({
        ...shortage,
        'Net Status': netStatus,
        'Container No.': '',
        'In Transit Qty': '',
        'Container Status': '',
        'Port ETA': ''
      });
      continue;
    }

    partsWithTransit.add(partCode);

    for (const row of containers) {
      combined.push({
        ...shortage,
        'Net Status':        netStatus,
        'Container No.':     row._ctNo,
        'In Transit Qty':    parseFloat(row['Quantity']) || 0,
        'Container Status':  String(row['Status'] || '').trim(),
        'Port ETA':          String(row['Port ETA'] || '').trim()
      });
    }
  }

  return { combined, partsWithTransitCount: partsWithTransit.size };
}

function runFIFO(requests, pool) {
  const fulfilled = [];
  const shortages = [];

  for (const req of requests) {
    const partCode = String(req['Part Code'] || '').trim();
    const requestedQty = parseFloat(req['Quantity']) || 0;
    const partName = String(req['Part Name'] || '').trim();
    const reqDate  = req['Requested Date'];
    const destLoc  = String(req['Destination Location'] || '').trim();

    const batches = pool.get(partCode) || [];

    // Total available at this moment (before this allocation)
    const totalAvailable = batches.reduce((s, b) => s + b.remaining, 0);

    let needed = requestedQty;
    let totalAllocated = 0;
    let runningTotal = 0;

    // Track which rows belong to this request (for multi-batch marking)
    const startIdx = fulfilled.length;

    for (const batch of batches) {
      if (needed <= 0) break;
      if (batch.remaining <= 0) continue;

      const allocate = Math.min(batch.remaining, needed);
      batch.remaining -= allocate;
      needed -= allocate;
      totalAllocated += allocate;
      runningTotal += allocate;

      fulfilled.push({
        'Container No.':                 batch.source === 'Container' ? (batch.containerNo || batch.caseNo) : '',
        'Type':                          batch.type || batch.source,
        'Case No':                          batch.caseNo,
        'Destination Location':             destLoc,
        'Pick Location':                    batch.pickLoc,
        'Batch Order Date':                 formatDate(batch.orderDate),
        'Order No / Invoice No':            batch.refNo,
        'Part Code':                        partCode,
        'Part Name':                        partName || batch.partName,
        'Quantity Allocated From This Batch': allocate,
        'Plant':                            batch.plant || '',
        'Source Location':                  String(req['Source Location'] || '').trim(),
        // ── remaining reference fields ──────────────────────
        'Requested Quantity':               requestedQty,
        'Requested Date':                   formatDate(reqDate),
        'Running Total Fulfilled':          runningTotal,
        'Allocation Source':                batch.source,
        'Status':                           'Pending',
        'Batch Location':                   batch.location,
        'Plant Sub Line':                   String(req['Plant Sub Line'] || '').trim(),
        'Shift Name':                       String(req['Shift Name'] || '').trim(),
        // ── internal markers (filtered out of Excel columns) ─
        _multiSource: false,   // filled in below
        _batchNum:    0,
        _batchTotal:  0
      });
    }

    // Mark multi-batch groups
    const endIdx = fulfilled.length;
    const batchCount = endIdx - startIdx;
    if (batchCount > 1) {
      for (let i = startIdx; i < endIdx; i++) {
        fulfilled[i]._multiSource = true;
        fulfilled[i]._batchNum   = i - startIdx + 1;
        fulfilled[i]._batchTotal = batchCount;
      }
    }

    // Fix final status on all rows added for this request
    const finalStatus = (needed <= 0) ? 'Fulfilled' : 'Partial Shortage';
    const rowsAddedCount = runningTotal > 0 ? fulfilled.filter(r =>
      r['Part Code'] === partCode && r['Running Total Fulfilled'] <= runningTotal).length : 0;
    // Mark status on every row emitted for this specific request pass
    let cumCheck = 0;
    for (let i = fulfilled.length - 1; i >= 0 && cumCheck < totalAllocated; i--) {
      cumCheck += fulfilled[i]['Quantity Allocated From This Batch'];
      fulfilled[i]['Status'] = finalStatus;
    }

    // Shortage row if needed
    const shortage = requestedQty - totalAllocated;
    if (shortage > 0) {
      shortages.push({
        'Part Code':              partCode,
        'Part Name':              partName,
        'Requested Quantity':     requestedQty,
        'Requested Date':         formatDate(reqDate),
        'Destination Location':   destLoc,
        'Total Quantity Available': totalAvailable,
        'Total Quantity Allocated': totalAllocated,
        'Shortage Quantity':      shortage,
        'Status':                 'Shortage'
      });
    }
  }

  return { fulfilled, shortages };
}

// ─── PROCESS FILES ────────────────────────────────────────────
async function processFiles() {
  if (!state.files.request || !state.files.inventory || !state.files.container) {
    showToast('Please upload all 3 files first.', 'error');
    return;
  }

  // Show loading
  document.getElementById('loading-overlay').hidden = false;
  document.getElementById('results-section').hidden = true;
  document.getElementById('process-btn').disabled = true;

  // Small delay to let the UI render
  await new Promise(r => setTimeout(r, 50));

  try {
    setLoadingSub('Building FIFO supply pool from Inventory & Containers…');
    await new Promise(r => setTimeout(r, 20));
    const pool = buildSupplyPool(state.data.inventory, state.data.containers);

    setLoadingSub(`Processing ${state.data.requests.length.toLocaleString()} part requests (FIFO)…`);
    await new Promise(r => setTimeout(r, 20));
    const { fulfilled, shortages } = runFIFO(state.data.requests, pool);

    setLoadingSub('Cross-checking shortages against in-transit containers…');
    await new Promise(r => setTimeout(r, 20));
    const { combined, partsWithTransitCount } = buildCombinedShortages(shortages, state.data.containers);

    state.results.fulfilled  = fulfilled;
    state.results.shortages  = combined; // Using the new combined list
    
    // Build summary
    const totalReqs    = state.data.requests.length;
    const fulfilledFull  = totalReqs - shortages.length;
    const unitsReq   = state.data.requests.reduce((s, r) => s + (parseFloat(r['Quantity']) || 0), 0);
    const unitsAlloc = fulfilled.reduce((s, r) => s + r['Quantity Allocated From This Batch'], 0);
    const unitsShort = shortages.reduce((s, r) => s + r['Shortage Quantity'], 0);

    state.results.summary = { totalReqs, fulfilledFull, shortageCount: shortages.length,
                               unitsReq, unitsAlloc, unitsShort, partsWithTransitCount };

    setLoadingSub('Rendering results…');
    await new Promise(r => setTimeout(r, 20));

    renderResults();

    document.getElementById('loading-overlay').hidden = true;
    document.getElementById('results-section').hidden = false;
    document.getElementById('results-section').scrollIntoView({ behavior: 'smooth', block: 'start' });
    showToast(`Done! ${fulfilledFull} fulfilled, ${shortages.length} shortages, ${partsWithTransitCount} found in transit.`, 'success', 6000);

  } catch (err) {
    document.getElementById('loading-overlay').hidden = true;
    document.getElementById('process-btn').disabled = false;
    showToast(`Processing error: ${err.message}`, 'error', 7000);
    console.error(err);
  }
}

// ─── RENDER RESULTS ───────────────────────────────────────────
const MAX_PREVIEW = 500;

function renderResults() {
  const s = state.results.summary;
  document.getElementById('stat-total').textContent        = s.totalReqs.toLocaleString();
  document.getElementById('stat-fulfilled').textContent    = s.fulfilledFull.toLocaleString();
  document.getElementById('stat-shortage').textContent     = s.shortageCount.toLocaleString();
  document.getElementById('stat-intransit').textContent    = s.partsWithTransitCount.toLocaleString();
  document.getElementById('stat-units-short').textContent  = s.unitsShort.toLocaleString();
  document.getElementById('stat-units-req').textContent    = s.unitsReq.toLocaleString();
  document.getElementById('stat-units-alloc').textContent  = s.unitsAlloc.toLocaleString();

  document.getElementById('tc-fulfilled').textContent  = state.results.fulfilled.length.toLocaleString();
  document.getElementById('tc-shortages').textContent  = state.results.shortages.length.toLocaleString();

  renderTable('fulfilled', state.results.fulfilled);
  renderTable('shortages', state.results.shortages);
}

function renderTable(type, rows) {
  const tbody = document.getElementById(`tbody-${type}`);
  const footer = document.getElementById(`footer-${type}`);
  tbody.innerHTML = '';

  const preview = rows.slice(0, MAX_PREVIEW);

  if (preview.length === 0) {
    const tr = document.createElement('tr');
    tr.innerHTML = `<td colspan="20" style="text-align:center;padding:32px;color:var(--grey-text)">No ${type} records.</td>`;
    tbody.appendChild(tr);
    footer.textContent = '';
    return;
  }

  if (type === 'fulfilled') {
    preview.forEach((row, idx) => {
      const tr = document.createElement('tr');
      const status = row['Status'];
      const isMulti = row._multiSource;
      const isLast  = isMulti && row._batchNum === row._batchTotal;

      const statusBadge = status === 'Fulfilled'
        ? `<span class="badge badge-green">✓ Fulfilled</span>`
        : `<span class="badge badge-orange">⚠ Partial</span>`;
      const srcBadge = row['Allocation Source'] === 'Inventory'
        ? `<span class="source-badge source-inv">Inventory</span>`
        : `<span class="source-badge source-cont">Container</span>`;

      tr.innerHTML = `
        <td><strong>${esc(row['Part Code'])}</strong></td>
        <td>${esc(row['Part Name'])}</td>
        <td>${fmt(row['Requested Quantity'])}</td>
        <td>${esc(row['Requested Date'])}</td>
        <td>${esc(row['Destination Location'])}</td>
        <td>${srcBadge}</td>
        <td>${esc(row['Pick Location'])}</td>
        <td>${esc(row['Batch Order Date'])}</td>
        <td style="font-weight:600;color:var(--green)">${fmt(row['Quantity Allocated From This Batch'])}</td>
        <td>${fmt(row['Running Total Fulfilled'])}</td>
        <td>${statusBadge}</td>
        <td>${esc(row['Case No'])}</td>`;
      tbody.appendChild(tr);

      // Insert summary separator row after the last batch of a multi-batch group
      if (isLast) {
        const reqQty    = row['Requested Quantity'];
        const allocated = row['Running Total Fulfilled'];
        const shortage  = reqQty - allocated;
        const ok        = shortage <= 0;
        const sepClass  = ok ? 'group-sep-ok' : 'group-sep-short';
        const icon      = ok ? '✓' : '⚠';
        const statusTxt = ok
          ? `<span class="gs-ok">${icon} Fully Fulfilled</span>`
          : `<span class="gs-short">${icon} Short by ${fmt(shortage)}</span>`;

        const sep = document.createElement('tr');
        sep.className = `group-separator ${sepClass}`;
        sep.innerHTML = `<td colspan="12">
          <div class="group-summary">
            <span class="gs-part">📦 ${esc(row['Part Code'])} &nbsp;·&nbsp; ${row._batchTotal} batches</span>
            <span class="gs-item">Requested: <strong>${fmt(reqQty)}</strong></span>
            <span class="gs-item">Allocated: <strong>${fmt(allocated)}</strong></span>
            ${statusTxt}
          </div>
        </td>`;
        tbody.appendChild(sep);
      }
    });
  } else if (type === 'shortages') {
    preview.forEach(row => {
      const tr = document.createElement('tr');
      const transitQty = row['In Transit Qty'] || '';
      
      tr.innerHTML = `
        <td><strong>${esc(row['Part Code'])}</strong></td>
        <td>${esc(row['Part Name'])}</td>
        <td>${fmt(row['Requested Quantity'])}</td>
        <td>${fmt(row['Total Quantity Allocated'])}</td>
        <td style="font-weight:700;color:var(--red)">${fmt(row['Shortage Quantity'])}</td>
        <td style="font-size:0.75rem;font-weight:600;color:var(--navy)">${esc(row['Net Status'])}</td>
        <td>${esc(row['Container No.'])}</td>
        <td style="font-weight:700;color:#2563EB">${fmt(transitQty)}</td>
        <td>${esc(row['Container Status'])}</td>
        <td>${esc(row['Port ETA'])}</td>
        <td>${esc(row['Destination Location'])}</td>`;
      tbody.appendChild(tr);
    });
  }

  footer.textContent = rows.length > MAX_PREVIEW
    ? `Showing first ${MAX_PREVIEW.toLocaleString()} of ${rows.length.toLocaleString()} rows. Download the Excel file for the full data.`
    : `Showing all ${rows.length.toLocaleString()} rows.`;
}

function esc(str) {
  if (str == null) return '—';
  return String(str).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;') || '—';
}
function fmt(num) {
  if (num == null || num === '') return '—';
  const n = parseFloat(num);
  return isNaN(n) ? String(num) : n.toLocaleString();
}

// ─── TAB SWITCHER ─────────────────────────────────────────────
function switchTab(tab) {
  ['fulfilled', 'shortages'].forEach(t => {
    document.getElementById(`tab-${t}`).classList.toggle('active', t === tab);
    document.getElementById(`panel-${t}`).classList.toggle('active', t === tab);
  });
}

// ─── EXCEL EXPORT ─────────────────────────────────────────────

// Priority columns — shown FIRST and highlighted light blue
const PRIORITY_COLS = [
  'Container No.',
  'Type',
  'Case No',
  'Destination Location',
  'Pick Location',
  'Batch Order Date',
  'Order No / Invoice No',
  'Part Code',
  'Part Name',
  'Quantity Allocated From This Batch',
  'Plant',
  'Source Location'
];

// Styles
const BORDER_THICK = { style: 'medium', color: { rgb: 'A0AAB5' } };
const BORDER_STD   = {
  top: BORDER_THICK, bottom: BORDER_THICK,
  left: BORDER_THICK, right: BORDER_THICK
};

const ALIGN_CENTER = { horizontal: 'center', vertical: 'center', wrapText: true };

const S_HDR_PRIORITY = {
  fill: { patternType: 'solid', fgColor: { rgb: '0A1F44' } },
  font: { bold: true, color: { rgb: 'FFFFFF' }, sz: 11 },
  alignment: ALIGN_CENTER,
  border: BORDER_STD
};
const S_HDR_NORMAL = {
  fill: { patternType: 'solid', fgColor: { rgb: '1A3566' } },
  font: { bold: true, color: { rgb: 'FFFFFF' }, sz: 10 },
  alignment: ALIGN_CENTER,
  border: BORDER_STD
};
const S_CELL_PRIORITY_ODD  = { fill: { patternType: 'solid', fgColor: { rgb: 'D6EEFF' } }, alignment: ALIGN_CENTER, border: BORDER_STD };
const S_CELL_PRIORITY_EVEN = { fill: { patternType: 'solid', fgColor: { rgb: 'BFE0FF' } }, alignment: ALIGN_CENTER, border: BORDER_STD };
const S_CELL_ODD           = { fill: { patternType: 'solid', fgColor: { rgb: 'FFFFFF' } }, alignment: ALIGN_CENTER, border: BORDER_STD };
const S_CELL_EVEN          = { fill: { patternType: 'solid', fgColor: { rgb: 'F4F6FA' } }, alignment: ALIGN_CENTER, border: BORDER_STD };

function buildStyledSheet(rows, priorityCols) {
  if (!rows || rows.length === 0) {
    return XLSX.utils.json_to_sheet([{ Note: 'No data.' }]);
  }

  // Filter out internal _ marker keys
  const cleanRows = rows.map(r => {
    const o = {};
    for (const k of Object.keys(r)) { if (!k.startsWith('_')) o[k] = r[k]; }
    return o;
  });

  // Build ordered column list: priority first, then remaining
  const seen = new Set(priorityCols);
  const allCols = [...priorityCols];
  for (const key of Object.keys(cleanRows[0])) {
    if (!seen.has(key)) { allCols.push(key); seen.add(key); }
  }
  const prioritySet = new Set(priorityCols);

  // Build AOA — insert a blank separator row after each multi-batch group's last row
  const aoa = [allCols];
  const isBlankRow = [false]; // track which AOA rows are separators (header = false)
  for (let i = 0; i < cleanRows.length; i++) {
    aoa.push(allCols.map(k => (cleanRows[i][k] != null ? cleanRows[i][k] : '')));
    isBlankRow.push(false);
    // If this is the last batch of a multi-batch group, add a summary separator row
    const orig = rows[i];
    if (orig && orig._multiSource && orig._batchNum === orig._batchTotal) {
      const reqQty    = cleanRows[i]['Requested Quantity'];
      const allocated = cleanRows[i]['Running Total Fulfilled'];
      const shortage  = reqQty - allocated;
      const ok        = shortage <= 0;
      const summaryText = ok
        ? `→ ${cleanRows[i]['Part Code']} | Batches: ${orig._batchTotal} | Requested: ${reqQty} | Allocated: ${allocated} | ✓ Fulfilled`
        : `→ ${cleanRows[i]['Part Code']} | Batches: ${orig._batchTotal} | Requested: ${reqQty} | Allocated: ${allocated} | ⚠ SHORT by ${shortage}`;
      const summaryRow = new Array(allCols.length).fill('');
      summaryRow[0] = summaryText;
      aoa.push(summaryRow);
      isBlankRow.push({ isSummary: true, ok });
    }
  }

  const ws = XLSX.utils.aoa_to_sheet(aoa);
  const range = XLSX.utils.decode_range(ws['!ref']);

  // Apply styles
  for (let R = range.s.r; R <= range.e.r; R++) {
    const even = R % 2 === 0;
    for (let C = range.s.c; C <= range.e.c; C++) {
      const addr = XLSX.utils.encode_cell({ r: R, c: C });
      if (!ws[addr]) ws[addr] = { v: '', t: 's' };
      const isPri = prioritySet.has(allCols[C]);

      if (R === 0) {
        ws[addr].s = isPri ? S_HDR_PRIORITY : S_HDR_NORMAL;
      } else if (isBlankRow[R] && isBlankRow[R].isSummary) {
        // Summary separator row
        const ok = isBlankRow[R].ok;
        if (C === 0) {
          ws[addr].s = {
            fill: { patternType: 'solid', fgColor: { rgb: ok ? 'F0FDF4' : 'FFF7ED' } },
            font: { italic: true, bold: false, color: { rgb: ok ? '166534' : '9A3412' }, sz: 9 },
            border: {
              top:    { style: 'medium', color: { rgb: ok ? '86EFAC' : 'FDBA74' } },
              bottom: { style: 'medium', color: { rgb: ok ? '86EFAC' : 'FDBA74' } },
              left:   { style: 'medium', color: { rgb: ok ? '86EFAC' : 'FDBA74' } },
              right:  { style: 'medium', color: { rgb: ok ? '86EFAC' : 'FDBA74' } }
            },
            alignment: ALIGN_CENTER
          };
        } else {
          ws[addr].s = {
            fill: { patternType: 'solid', fgColor: { rgb: ok ? 'F0FDF4' : 'FFF7ED' } },
            border: {
              top:    { style: 'medium', color: { rgb: ok ? '86EFAC' : 'FDBA74' } },
              bottom: { style: 'medium', color: { rgb: ok ? '86EFAC' : 'FDBA74' } },
              left:   { style: 'medium', color: { rgb: ok ? '86EFAC' : 'FDBA74' } },
              right:  { style: 'medium', color: { rgb: ok ? '86EFAC' : 'FDBA74' } }
            },
            alignment: ALIGN_CENTER
          };
        }
      } else if (isBlankRow[R]) {
        ws[addr].s = {
          fill: { patternType: 'solid', fgColor: { rgb: 'FFFFFF' } },
          border: { top: { style: 'medium', color: { rgb: 'A0AAB5' } } }
        };
      } else {
        ws[addr].s = isPri
          ? (even ? S_CELL_PRIORITY_EVEN : S_CELL_PRIORITY_ODD)
          : (even ? S_CELL_EVEN : S_CELL_ODD);
      }
    }
  }

  // Column widths & freeze header
  ws['!cols'] = allCols.map(k => ({ wch: Math.max(14, Math.min(k.length + 6, 42)) }));
  ws['!freeze'] = { xSplit: 0, ySplit: 1, topLeftCell: 'A2', activePane: 'bottomLeft', state: 'frozen' };

  return ws;
}

function downloadExcel() {
  try {
    const now = new Date();
    const stamp = `${now.getFullYear()}${String(now.getMonth()+1).padStart(2,'0')}${String(now.getDate()).padStart(2,'0')}_${String(now.getHours()).padStart(2,'0')}${String(now.getMinutes()).padStart(2,'0')}`;

    // Helper to generate and download a single workbook
    const createWorkbook = (fRows, sRows, summaryData, suffix) => {
      const wb = XLSX.utils.book_new();

      // Sheet 1: Fulfilled
      const ws1 = buildStyledSheet(fRows.length ? fRows : [{ Note: 'No fulfilled allocations.' }], PRIORITY_COLS);
      XLSX.utils.book_append_sheet(wb, ws1, 'Fulfilled');

      // Sheet 2: Shortages
      const shortagePri = [
        'Part Code', 'Part Name', 'Requested Quantity', 'Shortage Quantity', 'Net Status',
        'Container No.', 'In Transit Qty', 'Container Status', 'Port ETA', 'Destination Location'
      ];
      const ws2 = buildStyledSheet(sRows.length ? sRows : [{ Note: 'No shortages.' }], shortagePri);
      XLSX.utils.book_append_sheet(wb, ws2, 'Shortages');

      // Sheet 3: Summary
      const ws4 = buildStyledSheet(summaryData, []);
      XLSX.utils.book_append_sheet(wb, ws4, 'Summary');

      // Sanitize the suffix to ensure a valid filename
      const safeSuffix = String(suffix).replace(/[^a-z0-9_-]/gi, '_').toUpperCase();
      XLSX.writeFile(wb, `FIFO_Allocation_${safeSuffix}_${stamp}.xlsx`);
    };

    // 1. Identify all unique Destination Locations
    const locations = new Set();
    state.results.fulfilled.forEach(r => {
      const loc = String(r['Destination Location'] || '').trim();
      if (loc) locations.add(loc);
    });
    state.results.shortages.forEach(r => {
      const loc = String(r['Destination Location'] || '').trim();
      if (loc) locations.add(loc);
    });

    // 2. Pre-calculate summaries for each location
    const locSummaries = new Map();
    
    locations.forEach(loc => {
      const fRows = state.results.fulfilled.filter(r => String(r['Destination Location'] || '').trim() === loc);
      const sRows = state.results.shortages.filter(r => String(r['Destination Location'] || '').trim() === loc);
      
      let locUnitsAlloc = 0;
      let locUnitsShort = 0;
      let locUnitsReq = 0;
      let locUnitsTransit = 0;
      
      const seenReqs = new Set();
      fRows.forEach(r => {
        locUnitsAlloc += (parseFloat(r['Quantity Allocated From This Batch']) || 0);
        const reqKey = r['Part Code'] + '|' + r['Requested Date'] + '|' + r['Requested Quantity'];
        if (!seenReqs.has(reqKey)) {
          seenReqs.add(reqKey);
          locUnitsReq += (parseFloat(r['Requested Quantity']) || 0);
        }
      });
      
      const uniqueShortages = new Set();
      sRows.forEach(r => {
        locUnitsTransit += (parseFloat(r['In Transit Qty']) || 0);
        
        const shortKey = r['Part Code'] + '|' + r['Requested Date'] + '|' + r['Shortage Quantity'];
        if (!uniqueShortages.has(shortKey)) {
          uniqueShortages.add(shortKey);
          locUnitsShort += (parseFloat(r['Shortage Quantity']) || 0);
        }
        
        const reqKey = r['Part Code'] + '|' + r['Requested Date'] + '|' + r['Requested Quantity'];
        if (!seenReqs.has(reqKey)) {
          seenReqs.add(reqKey);
          locUnitsReq += (parseFloat(r['Requested Quantity']) || 0);
        }
      });
      
      locSummaries.set(loc, {
        totalReqs: seenReqs.size,
        fulfilledFull: seenReqs.size - uniqueShortages.size,
        shortageCount: uniqueShortages.size,
        unitsReq: locUnitsReq,
        unitsAlloc: locUnitsAlloc,
        unitsShort: locUnitsShort,
        unitsTransit: locUnitsTransit
      });
    });

    // 3. Build multi-row Summary for the FULL report
    let overallUnitsTransit = 0;
    state.results.shortages.forEach(r => {
      overallUnitsTransit += (parseFloat(r['In Transit Qty']) || 0);
    });

    const fullSummaryRows = [
      {
        'Destination Location': 'ALL (OVERALL)',
        'Requests Processed': state.results.summary.totalReqs,
        'Fully Fulfilled': state.results.summary.fulfilledFull,
        'Shortage Count': state.results.summary.shortageCount,
        'Units Requested': state.results.summary.unitsReq,
        'Units Allocated': state.results.summary.unitsAlloc,
        'Units Short': state.results.summary.unitsShort,
        'Units In Transit': overallUnitsTransit,
        'Generated At': new Date().toLocaleString()
      }
    ];

    locations.forEach(loc => {
      const s = locSummaries.get(loc);
      fullSummaryRows.push({
        'Destination Location': loc,
        'Requests Processed': s.totalReqs,
        'Fully Fulfilled': s.fulfilledFull,
        'Shortage Count': s.shortageCount,
        'Units Requested': s.unitsReq,
        'Units Allocated': s.unitsAlloc,
        'Units Short': s.unitsShort,
        'Units In Transit': s.unitsTransit,
        'Generated At': ''
      });
    });

    // 4. Download the FULL Master Report
    createWorkbook(state.results.fulfilled, state.results.shortages, fullSummaryRows, 'FULL');

    // 5. Generate and download a separate file for each unique location
    let delay = 300;
    locations.forEach(loc => {
      setTimeout(() => {
        const fRows = state.results.fulfilled.filter(r => String(r['Destination Location'] || '').trim() === loc);
        const sRows = state.results.shortages.filter(r => String(r['Destination Location'] || '').trim() === loc);
        
        const s = locSummaries.get(loc);
        const locSummaryRow = [{
          'Destination Location': loc,
          'Requests Processed': s.totalReqs,
          'Fully Fulfilled': s.fulfilledFull,
          'Shortage Count': s.shortageCount,
          'Units Requested': s.unitsReq,
          'Units Allocated': s.unitsAlloc,
          'Units Short': s.unitsShort,
          'Units In Transit': s.unitsTransit,
          'Generated At': new Date().toLocaleString()
        }];

        createWorkbook(fRows, sRows, locSummaryRow, loc);
      }, delay);
      delay += 400; // Stagger downloads slightly to prevent browser blocking
    });

    showToast(`Downloading FULL report + ${locations.size} location reports...`, 'success', 5000);
  } catch (err) {
    showToast(`Export error: ${err.message}`, 'error');
    console.error(err);
  }
}

// ─── RESET ────────────────────────────────────────────────────
function resetAll() {
  ['request', 'inventory', 'container'].forEach(removeFile);
  document.getElementById('results-section').hidden = true;
  document.getElementById('process-btn').disabled = true;
  state.results = { fulfilled: [], shortages: [], summary: {} };
  document.getElementById('upload-section').scrollIntoView({ behavior: 'smooth' });
  showToast('Ready for a new batch.', 'info');
}

// ─── INIT ─────────────────────────────────────────────────────
updateProcessBtn();
