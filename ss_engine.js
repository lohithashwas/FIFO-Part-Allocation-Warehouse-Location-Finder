/* ============================================================
   WareFlow – Section C: Shortage Statement Engine
   ============================================================ */

// ─── SS State ──────────────────────────────────────────────────
const ssState = {
  files:   { plan: null, inventory: null, container: null },
  data:    { planRequests: [], inventory: [], containers: [] },
  results: { fulfilled: [], shortages: [], summary: {} }
};

// Helper: Normalize Part Codes for hyphen-insensitive & space-insensitive matching
function ssNormCode(code) {
  if (!code) return '';
  return String(code).replace(/[^A-Z0-9]/gi, '').toUpperCase();
}

// ─── Drag & Drop Event Handlers ────────────────────────────────
['plan', 'inventory', 'container'].forEach(type => {
  const zone  = document.getElementById(`ss-drop-${type}`);
  const input = document.getElementById(`ss-file-${type}`);

  if (!zone || !input) return;

  zone.addEventListener('dragover', e => { e.preventDefault(); zone.classList.add('drag-over'); });
  zone.addEventListener('dragleave', () => zone.classList.remove('drag-over'));
  zone.addEventListener('drop', e => {
    e.preventDefault(); zone.classList.remove('drag-over');
    const file = e.dataTransfer.files[0];
    if (file) ssHandleFile(type, file);
  });
  input.addEventListener('change', () => { if (input.files[0]) ssHandleFile(type, input.files[0]); });
});

// ─── File Input Processing ────────────────────────────────────
function ssHandleFile(type, file) {
  ssClearError(type);
  if (!file.name.match(/\.xlsx$/i)) {
    ssShowError(type, 'Only .xlsx files are accepted.');
    return;
  }
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const wb   = XLSX.read(e.target.result, { type: 'array', cellDates: true });
      const ws   = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });

      if (rows.length === 0) { ssShowError(type, 'File is empty or has no data rows.'); return; }

      if (type === 'plan') {
        const firstRow = rows[0];
        const keys = Object.keys(firstRow);

        // Find Part Code column (e.g. "Material No.", "Part Code", "Material No", "Part No")
        const partKey = keys.find(k => k.match(/material\s*no/i) || k.match(/part\s*code/i) || k.match(/part\s*no/i) || k.match(/material/i)) || keys[0];
        
        // Find Shortage column (e.g. "P401\nShortage", "Shortage", "Shortage Qty", "Quantity")
        const shortKey = keys.find(k => k.match(/shortage/i)) || keys.find(k => k.match(/qty|quantity/i));

        if (!partKey) {
          ssShowError(type, "Missing part column: 'Material No.' or 'Part Code'");
          return;
        }
        if (!shortKey) {
          ssShowError(type, "Missing shortage column: 'Shortage' or 'P401\\nShortage'");
          return;
        }

        // Parse plan file into standard request format
        const planRequests = [];
        rows.forEach((row, idx) => {
          const rawCode = String(row[partKey] || '').trim();
          if (!rawCode) return;

          // Find Part Name / Description
          const descKey = keys.find(k => k.match(/desc/i) || k.match(/name/i) || k.match(/description/i));
          const partName = descKey ? String(row[descKey] || '').trim() : '';

          // Find Plant
          const plantKey = keys.find(k => k.match(/plant/i));
          const plant = plantKey ? String(row[plantKey] || '').trim() : 'HVF1';

          // Shortage Qty — take absolute value (handles negative -32 as positive shortage quantity 32)
          let rawQty = row[shortKey];
          let qty = Math.abs(parseFloat(rawQty) || 0);

          if (qty <= 0) return; // Only process actual shortages

          // Find Requested Date (Base date column)
          const dateKey = keys.find(k => k.match(/date/i));
          let reqDate = dateKey ? formatDate(row[dateKey]) : formatDate(new Date());

          planRequests.push({
            'Part Code':            rawCode,
            'Part Name':            partName,
            'Quantity':             qty,
            'Requested Date':       reqDate,
            'Destination Location': '', // empty string so FIFO allocation searches ALL warehouse locations
            'Source Location':      plant || 'HVF1',
            'Requested Quantity':   qty,
            'Shortage Quantity':    qty,
            _reqIdx:                idx
          });
        });

        if (planRequests.length === 0) {
          ssShowError(type, 'No positive shortage rows found in plan file.');
          return;
        }

        ssState.files.plan = file;
        ssState.data.planRequests = planRequests;

      } else if (type === 'inventory') {
        const colsLower = Object.keys(rows[0]).map(k => k.trim().toLowerCase());
        const hasPartCol = colsLower.includes('part code') || colsLower.includes('part no') || colsLower.includes('material no.') || colsLower.includes('material no');
        const hasQty     = colsLower.includes('quantity') || colsLower.includes('qty');
        if (!hasPartCol) { ssShowError(type, `Missing column: 'Part Code' or 'Part No'`); return; }
        if (!hasQty)     { ssShowError(type, `Missing column: 'Quantity'`); return; }

        ssState.files.inventory = file;
        ssState.data.inventory = rows;

      } else if (type === 'container') {
        const colsLower = Object.keys(rows[0]).map(k => k.trim().toLowerCase());
        const hasPartCol = colsLower.includes('part no') || colsLower.includes('part code') || colsLower.includes('material no.') || colsLower.includes('material no');
        const hasQty     = colsLower.includes('quantity') || colsLower.includes('qty');
        if (!hasPartCol) { ssShowError(type, `Missing column: 'Part No' or 'Part Code'`); return; }
        if (!hasQty)     { ssShowError(type, `Missing column: 'Quantity'`); return; }

        ssState.files.container = file;
        ssState.data.containers = rows;
      }

      document.getElementById(`ss-fp-${type}-name`).textContent = file.name;
      document.getElementById(`ss-fp-${type}-rows`).textContent = type === 'plan' 
        ? `${ssState.data.planRequests.length.toLocaleString()} shortage items` 
        : `${rows.length.toLocaleString()} rows`;
      
      document.getElementById(`ss-fp-${type}`).classList.add('visible');
      document.getElementById(`ss-card-${type}`).classList.add('has-file');
      document.getElementById(`ss-card-${type}`).classList.remove('has-error');

      ssUpdateProcessBtn();

    } catch (err) {
      ssShowError(type, 'Could not read file: ' + err.message);
    }
  };
  reader.readAsArrayBuffer(file);
}

function ssRemoveFile(type) {
  ssState.files[type] = null;
  if (type === 'plan') ssState.data.planRequests = [];
  if (type === 'inventory') ssState.data.inventory = [];
  if (type === 'container') ssState.data.containers = [];

  const input = document.getElementById(`ss-file-${type}`);
  if (input) input.value = '';

  document.getElementById(`ss-fp-${type}`).classList.remove('visible');
  document.getElementById(`ss-card-${type}`).classList.remove('has-file');
  ssClearError(type);
  ssUpdateProcessBtn();
}

function ssShowError(type, msg) {
  const el = document.getElementById(`ss-err-${type}`);
  if (el) { el.textContent = msg; el.classList.add('visible'); }
  document.getElementById(`ss-card-${type}`).classList.add('has-error');
}

function ssClearError(type) {
  const el = document.getElementById(`ss-err-${type}`);
  if (el) { el.textContent = ''; el.classList.remove('visible'); }
  document.getElementById(`ss-card-${type}`).classList.remove('has-error');
}

function ssUpdateProcessBtn() {
  const allSet = ssState.files.plan && ssState.files.inventory && ssState.files.container;
  const btn    = document.getElementById('ss-process-btn');
  const status = document.getElementById('ss-upload-status');

  if (btn) btn.disabled = !allSet;
  if (status) {
    if (allSet) {
      status.textContent = 'All 3 files loaded. Ready to calculate shortage allocation!';
      status.style.color = 'var(--green)';
    } else {
      const missing = [];
      if (!ssState.files.plan)      missing.push('Plan File');
      if (!ssState.files.inventory) missing.push('Inventory');
      if (!ssState.files.container) missing.push('Container List');
      status.textContent = `Upload missing: ${missing.join(', ')}`;
      status.style.color = 'var(--grey-text)';
    }
  }
}

// ─── PROCESS SHORTAGE STATEMENT ───────────────────────────────
function processSS() {
  if (!ssState.files.plan || !ssState.files.inventory || !ssState.files.container) return;

  const overlay = document.getElementById('ss-loading-overlay');
  if (overlay) overlay.hidden = false;

  setTimeout(() => {
    try {
      // 1. Build supply pool from Inventory and Containers
      const { pool, blockedMap } = buildSupplyPool(ssState.data.inventory, ssState.data.containers);

      // 2. Run FIFO Allocation on plan requests
      const fifoResults = runFIFO(ssState.data.planRequests, pool);

      // 3. Enrich shortage rows with transit & blocked stock info
      const { combined: combinedShortages } = buildCombinedShortages(fifoResults.shortages, ssState.data.containers, blockedMap);

      // 4. Compute Summary Metrics
      const totalReqs       = ssState.data.planRequests.length;
      const totalUnitsReq   = ssState.data.planRequests.reduce((s, r) => s + (parseFloat(r['Quantity']) || 0), 0);
      const totalUnitsAlloc = (fifoResults.fulfilled || []).reduce((s, r) => s + (parseFloat(r['Quantity Allocated From This Batch']) || 0), 0);
      const totalUnitsShort = (fifoResults.shortages || []).reduce((s, r) => s + (parseFloat(r['Shortage Quantity']) || 0), 0);
      const fulfilledFull   = ssState.data.planRequests.filter(r => !(fifoResults.shortages || []).some(s => s._reqIdx === r._reqIdx)).length;
      const shortageCount   = totalReqs - fulfilledFull;

      const summary = {
        totalReqs,
        fulfilledFull,
        shortageCount,
        totalUnitsReq,
        totalUnitsAlloc,
        totalUnitsShort
      };

      ssState.results = {
        fulfilled:  fifoResults.fulfilled || [],
        shortages:  combinedShortages || [],
        summary:    summary,
        blockedMap: blockedMap
      };

      if (overlay) overlay.hidden = true;

      // 5. Display Results
      ssRenderResults();
      document.getElementById('ss-results-section').hidden = false;
      document.getElementById('ss-results-section').scrollIntoView({ behavior: 'smooth' });

      showToast('Shortage Statement allocation complete!', 'success');

    } catch (err) {
      if (overlay) overlay.hidden = true;
      showToast('Error processing Shortage Statement: ' + err.message, 'error');
      console.error(err);
    }
  }, 100);
}

// ─── RENDER RESULTS ───────────────────────────────────────────
function ssRenderResults() {
  const { fulfilled = [], shortages = [], summary = {} } = ssState.results;

  // Stat cards
  document.getElementById('ss-stat-total').textContent       = (summary.totalReqs || 0).toLocaleString();
  document.getElementById('ss-stat-full').textContent        = (summary.fulfilledFull || 0).toLocaleString();
  document.getElementById('ss-stat-short').textContent       = (summary.shortageCount || 0).toLocaleString();
  document.getElementById('ss-stat-units-req').textContent   = (summary.totalUnitsReq || 0).toLocaleString();
  document.getElementById('ss-stat-units-alloc').textContent = (summary.totalUnitsAlloc || 0).toLocaleString();
  document.getElementById('ss-stat-units-short').textContent = (summary.totalUnitsShort || 0).toLocaleString();

  // Tab counts
  const fCount = Array.isArray(fulfilled) ? fulfilled.length : 0;
  const sCount = Array.isArray(shortages) ? shortages.length : 0;

  document.getElementById('ss-tc-fulfilled').textContent = fCount.toLocaleString();
  document.getElementById('ss-tc-shortages').textContent = sCount.toLocaleString();

  // Render tables
  ssRenderTable('fulfilled', fulfilled);
  ssRenderTable('shortages', shortages);
}

function ssRenderTable(type, rows) {
  const tbody  = document.getElementById(`ss-tbody-${type}`);
  const footer = document.getElementById(`ss-footer-${type}`);
  if (!tbody) return;
  tbody.innerHTML = '';

  const list = Array.isArray(rows) ? rows : [];
  const preview = list.slice(0, 500);

  if (preview.length === 0) {
    const tr = document.createElement('tr');
    tr.innerHTML = `<td colspan="20" style="text-align:center;padding:32px;color:var(--grey-text)">No ${type} records.</td>`;
    tbody.appendChild(tr);
    if (footer) footer.textContent = '';
    return;
  }

  if (type === 'fulfilled') {
    preview.forEach(row => {
      const tr = document.createElement('tr');
      const status = row['Status'];
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
    });
  } else if (type === 'shortages') {
    preview.forEach(row => {
      const tr = document.createElement('tr');
      const transitQty = row['In Transit Qty'] || '';
      const blockedQty = row['Blocked Qty in Storage'];
      const blockedLocs = row['Blocked Storage Locations'];
      
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
        <td>${esc(row['Destination Location'])}</td>
        <td style="font-weight:700;color:#D97706">${blockedQty ? fmt(blockedQty) : '—'}</td>
        <td style="font-size:0.75rem;color:#92400E">${blockedLocs ? esc(blockedLocs) : '—'}</td>`;
      tbody.appendChild(tr);
    });
  }

  if (footer) {
    footer.textContent = list.length > 500
      ? `Showing first 500 of ${list.length.toLocaleString()} rows. Download the Excel file for the full data.`
      : `Showing all ${list.length.toLocaleString()} rows.`;
  }
}

// ─── TAB SWITCHER ─────────────────────────────────────────────
function ssSwitchTab(tab) {
  ['fulfilled', 'shortages'].forEach(t => {
    const tabBtn = document.getElementById(`ss-tab-${t}`);
    const panel  = document.getElementById(`ss-panel-${t}`);
    if (tabBtn) tabBtn.classList.toggle('active', t === tab);
    if (panel)  panel.classList.toggle('active', t === tab);
  });
}

// ─── REPORT BUILDER HELPERS ───────────────────────────────────

function buildFIFOWaterfallReport(shortages, fulfilled, containerRows, blockedMap) {
  if (!shortages || shortages.length === 0) return [];

  const allocMap = new Map();
  (fulfilled || []).forEach(f => {
    const code = ssNormCode(f['Part Code']);
    if (!code) return;
    if (!allocMap.has(code)) allocMap.set(code, { total: 0, locs: new Set() });
    const entry = allocMap.get(code);
    entry.total += (parseFloat(f['Quantity Allocated From This Batch']) || 0);
    const loc = String(f['Pick Location'] || f['Destination Location'] || '').trim();
    if (loc) entry.locs.add(loc);
  });

  const ctMap = new Map();
  (containerRows || []).forEach(r => {
    const rawCode = String(r['Part No'] || r['Part Code'] || r['Material No.'] || r['Material No'] || '').trim();
    if (!rawCode) return;
    const code = ssNormCode(rawCode);
    const status = String(r['Status'] || '').trim().toUpperCase();
    if (!status.includes('TRANSIT') && !status.includes('PORT')) return;
    if (!ctMap.has(code)) ctMap.set(code, []);
    const _normCtKey = Object.keys(r).find(k => k.replace(/[^A-Z0-9]/gi, '').toUpperCase() === 'CTNO' || k.replace(/[^A-Z0-9]/gi, '').toUpperCase() === 'CONTAINERNO');
    const ctNo = _normCtKey ? String(r[_normCtKey] || '').trim() : String(r['C/T No'] || r['C/T No.'] || r['Container No'] || '').trim();
    ctMap.get(code).push({
      qty: parseFloat(r['Quantity']) || 0,
      ctNo: ctNo,
      eta: formatDate(r['ETA'] || r['Port ETA']),
      status: String(r['Status'] || '').trim()
    });
  });

  return shortages.map(s => {
    const partCode = String(s['Part Code'] || '').trim();
    const normCode = ssNormCode(partCode);
    const reqQty   = parseFloat(s['Shortage Quantity'] || s['Quantity'] || 0);

    const allocInfo = allocMap.get(normCode) || { total: 0, locs: new Set() };
    const whAlloc   = allocInfo.total;
    const whLocs    = Array.from(allocInfo.locs).join(', ');
    const remAfterWh = Math.max(0, reqQty - whAlloc);

    const transitBatches = ctMap.get(normCode) || [];
    const transitQty = transitBatches.reduce((acc, b) => acc + b.qty, 0);
    const ctNos = [...new Set(transitBatches.map(b => b.ctNo).filter(Boolean))].join(', ');
    const etas  = [...new Set(transitBatches.map(b => b.eta).filter(Boolean))].join(', ');
    const stats = [...new Set(transitBatches.map(b => b.status).filter(Boolean))].join(', ');
    const remAfterTransit = Math.max(0, remAfterWh - transitQty);

    const blockedBatches = (blockedMap && (blockedMap.get(normCode) || blockedMap.get(partCode))) || [];
    const blockedQty = blockedBatches.reduce((acc, b) => acc + b.qty, 0);
    const blockedLocs = [...new Set(blockedBatches.map(b => b.location).filter(Boolean))].join(', ');

    const absoluteShortage = Math.max(0, remAfterTransit);
    let overallStatus = 'Absolute Shortage';
    if (remAfterWh === 0) overallStatus = 'Covered by Warehouse';
    else if (remAfterTransit === 0) overallStatus = 'Covered by Transit/Port';
    else if (transitQty > 0) overallStatus = 'Partial Cover by Transit';

    return {
      'Part Code': partCode,
      'Part Name': s['Part Name'] || '',
      'Destination Location': s['Destination Location'] || '',
      'Requested Date': s['Requested Date'] || '',
      'Required Qty (Shortage)': reqQty,
      '① Allocated from Warehouse': whAlloc,
      '① Warehouse Locations': whLocs,
      '① Remaining After Warehouse': remAfterWh,
      '② In Transit / Port Qty': transitQty,
      '② Container Nos': ctNos,
      '② Port ETA': etas,
      '② Transit Status': stats,
      '② Remaining After Transit': remAfterTransit,
      '③ Blocked Qty in Storage': blockedQty > 0 ? blockedQty : 0,
      '③ Blocked Locations': blockedLocs,
      '❌ Absolute Shortage': absoluteShortage,
      'Status': overallStatus
    };
  });
}

function buildShortagePartRequestRows(shortages, inventory, containerRows) {
  if (!shortages || shortages.length === 0) return [];
  return shortages.map(s => ({
    'Part Code': s['Part Code'] || '',
    'Part Name': s['Part Name'] || '',
    'Quantity': s['Shortage Quantity'] || s['Quantity'] || 0,
    'Requested Date': s['Requested Date'] || '',
    'Destination Location': s['Destination Location'] || '',
    'Shortage Quantity': s['Shortage Quantity'] || s['Quantity'] || 0,
    'Requested Quantity': s['Requested Quantity'] || s['Quantity'] || 0,
    'Total Qty Available': s['Total Quantity Available'] || 0,
    'Total Qty Allocated': s['Total Quantity Allocated'] || 0,
    'Net Status': s['Net Status'] || '',
    'In Transit Qty': s['In Transit Qty'] || '',
    'Port ETA': s['Port ETA'] || '',
    'Blocked Qty in Storage': s['Blocked Qty in Storage'] || '',
    'Blocked Storage Location': s['Blocked Storage Locations'] || ''
  }));
}

function buildPortTransitReport(shortages, containerRows, blockedMap) {
  if (!shortages || shortages.length === 0) return [];
  const results = [];

  shortages.forEach(s => {
    const partCode = String(s['Part Code'] || '').trim();
    const normCode = ssNormCode(partCode);
    const reqQty   = parseFloat(s['Shortage Quantity'] || s['Quantity'] || 0);

    const ctMap = new Map();
    (containerRows || []).forEach(r => {
      const rawCode = String(r['Part No'] || r['Part Code'] || r['Material No.'] || r['Material No'] || '').trim();
      if (!rawCode) return;
      const code = ssNormCode(rawCode);
      const status = String(r['Status'] || '').trim().toUpperCase();
      if (!status.includes('TRANSIT') && !status.includes('PORT')) return;
      if (!ctMap.has(code)) ctMap.set(code, []);
      const _normCtKey = Object.keys(r).find(k => k.replace(/[^A-Z0-9]/gi, '').toUpperCase() === 'CTNO' || k.replace(/[^A-Z0-9]/gi, '').toUpperCase() === 'CONTAINERNO');
      const ctNo = _normCtKey ? String(r[_normCtKey] || '').trim() : String(r['C/T No'] || r['C/T No.'] || r['Container No'] || '').trim();
      ctMap.get(code).push({
        qty: parseFloat(r['Quantity']) || 0,
        ctNo: ctNo,
        eta: formatDate(r['ETA'] || r['Port ETA']),
        status: String(r['Status'] || '').trim()
      });
    });

    const transitBatches = ctMap.get(normCode) || [];
    const transitQty = transitBatches.reduce((acc, b) => acc + b.qty, 0);
    const ctNos = [...new Set(transitBatches.map(b => b.ctNo).filter(Boolean))].join(', ');
    const etas  = [...new Set(transitBatches.map(b => b.eta).filter(Boolean))].join(', ');
    const stats = [...new Set(transitBatches.map(b => b.status).filter(Boolean))].join(', ');

    const blockedBatches = (blockedMap && (blockedMap.get(normCode) || blockedMap.get(partCode))) || [];
    const blockedQty = blockedBatches.reduce((acc, b) => acc + b.qty, 0);
    const blockedLocs = [...new Set(blockedBatches.map(b => b.location).filter(Boolean))].join(', ');

    if (transitQty > 0 || blockedQty > 0) {
      results.push({
        'Part Code': partCode,
        'Part Name': s['Part Name'] || '',
        'Quantity': reqQty,
        'Requested Date': s['Requested Date'] || '',
        'Destination Location': s['Destination Location'] || '',
        'Container No.': ctNos || '—',
        'In Transit Qty': transitQty > 0 ? transitQty : '—',
        'Container Status': stats || '—',
        'Port ETA': etas || '—',
        'Blocked Qty': blockedQty > 0 ? blockedQty : '—',
        'Blocked Storage Locations': blockedLocs || '—',
        'Availability Status': transitQty >= reqQty ? 'Fully Covered by Transit' : transitQty > 0 ? 'Partially Covered by Transit' : 'Blocked Stock Only'
      });
    }
  });

  return results;
}

// ─── EXCEL EXPORT ─────────────────────────────────────────────
function ssDownloadExcel() {
  try {
    const now = new Date();
    const stamp = `${now.getFullYear()}${String(now.getMonth()+1).padStart(2,'0')}${String(now.getDate()).padStart(2,'0')}_${String(now.getHours()).padStart(2,'0')}${String(now.getMinutes()).padStart(2,'0')}`;

    const fulfillPri = [
      'Part Code', 'Part Name', 'Quantity', 'Requested Date',
      'Destination Location',
      'Quantity Allocated From This Batch', 'Status',
      'Container No.', 'Pick Location', 'Batch Order Date',
      'Order No / Invoice No', 'Case No', 'Type', 'Plant'
    ];

    const shortagePri = [
      'Part Code', 'Part Name', 'Quantity', 'Requested Date',
      'Destination Location',
      'Shortage Quantity', 'Requested Quantity', 'Total Quantity Available', 'Total Quantity Allocated',
      'Status', 'Net Status',
      'Container No.', 'In Transit Qty', 'Container Status', 'Port ETA',
      'Blocked Qty in Storage', 'Blocked Storage Locations'
    ];

    const btnContainer = document.getElementById('ss-download-buttons');
    if (btnContainer) btnContainer.innerHTML = '';
    const mgr = document.getElementById('ss-download-manager');
    if (mgr) mgr.style.display = 'block';

    const addManualDownloadBtn = (wb, filename) => {
      XLSX.writeFile(wb, filename);
      if (btnContainer) {
        const btn = document.createElement('button');
        btn.className = 'btn btn-outline';
        btn.style.cssText = 'font-size:0.8rem;padding:6px 12px;border-color:var(--primary);color:var(--primary);';
        btn.innerHTML = `📄 ${filename}`;
        btn.onclick = () => XLSX.writeFile(wb, filename);
        btnContainer.appendChild(btn);
      }
    };

    const createWorkbook = (fRows, sRows, summaryData, partMasterData, suffix) => {
      const wb = XLSX.utils.book_new();

      const ws1 = buildStyledSheet(fRows.length ? fRows : [{ Note: 'No fulfilled allocations.' }], fulfillPri);
      XLSX.utils.book_append_sheet(wb, ws1, 'Fulfilled');

      const ws2 = buildStyledSheet(sRows.length ? sRows : [{ Note: 'No shortages.' }], shortagePri);
      XLSX.utils.book_append_sheet(wb, ws2, 'Shortages');

      const ws4 = buildStyledSheet(summaryData, []);
      XLSX.utils.book_append_sheet(wb, ws4, 'Summary');

      if (partMasterData && partMasterData.length > 0) {
        const pmPri = ['Part Code', 'Part Name', 'Total Shortage Qty', 'Net Status', 'Total In Transit Qty', 'Containers', 'Transit Statuses'];
        const wsPM = buildStyledSheet(partMasterData, pmPri);
        XLSX.utils.book_append_sheet(wb, wsPM, 'Shortage Parts Master');
      }

      const safeSuffix = String(suffix).replace(/[^a-z0-9_-]/gi, '_').toUpperCase();
      const filename = `SHORTAGE_STATEMENT_${safeSuffix}_${stamp}.xlsx`;
      addManualDownloadBtn(wb, filename);
    };

    // Part Master List for Shortages
    const partMasterMap = new Map();
    (ssState.results.shortages || []).forEach(r => {
      const code = String(r['Part Code'] || '').trim();
      if (!code) return;
      if (!partMasterMap.has(code)) {
        partMasterMap.set(code, {
          'Part Code': code,
          'Part Name': r['Part Name'],
          _shortKeys: new Set(),
          _transitKeys: new Set(),
          'Total Shortage Qty': 0,
          'Total In Transit Qty': 0,
          'Containers': new Set(),
          'Statuses': new Set()
        });
      }
      
      const p = partMasterMap.get(code);
      const shortKey = code + '|' + r['Requested Date'] + '|' + r['Shortage Quantity'] + '|' + r['Destination Location'];
      if (!p._shortKeys.has(shortKey)) {
        p._shortKeys.add(shortKey);
        p['Total Shortage Qty'] += (parseFloat(r['Shortage Quantity']) || 0);
      }
      
      const cont = String(r['Container No.'] || '').trim();
      const transitQty = parseFloat(r['In Transit Qty']) || 0;
      if (cont && transitQty > 0) {
        const transKey = cont + '|' + transitQty;
        if (!p._transitKeys.has(transKey)) {
          p._transitKeys.add(transKey);
          p['Total In Transit Qty'] += transitQty;
          p['Containers'].add(cont);
          const stat = String(r['Container Status'] || '').trim();
          if (stat) p['Statuses'].add(stat);
        }
      }
    });

    const partMasterList = Array.from(partMasterMap.values()).map(p => {
      const netShort = Math.max(0, p['Total Shortage Qty'] - p['Total In Transit Qty']);
      let netStatus = '';
      if (p['Total In Transit Qty'] === 0) netStatus = `Absolute Shortage`;
      else if (netShort === 0) netStatus = `Fully Covered`;
      else netStatus = `Partial Shortage (${netShort} missing)`;

      return {
        'Part Code': p['Part Code'],
        'Part Name': p['Part Name'],
        'Total Shortage Qty': p['Total Shortage Qty'],
        'Net Status': netStatus,
        'Total In Transit Qty': p['Total In Transit Qty'],
        'Containers': Array.from(p['Containers']).join(', '),
        'Transit Statuses': Array.from(p['Statuses']).join(', ')
      };
    });

    // FILE 1: SHORTAGE WATERFALL REPORT
    const waterfallRows = buildFIFOWaterfallReport(
      ssState.results.shortages, ssState.results.fulfilled,
      ssState.data.containers, ssState.results.blockedMap
    );
    const waterfallPri = [
      'Part Code', 'Part Name', 'Destination Location', 'Requested Date',
      'Required Qty (Shortage)',
      '① Allocated from Warehouse', '① Warehouse Locations', '① Remaining After Warehouse',
      '② In Transit / Port Qty', '② Container Nos', '② Port ETA', '② Transit Status', '② Remaining After Transit',
      '③ Blocked Qty in Storage', '③ Blocked Locations',
      '❌ Absolute Shortage', 'Status'
    ];
    const wbWater = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wbWater,
      buildStyledSheet(waterfallRows.length ? waterfallRows : [{Note:'No shortages.'}], waterfallPri),
      'Shortage Waterfall');
    
    const shortReqRows = buildShortagePartRequestRows(
      ssState.results.shortages, ssState.data.inventory, ssState.data.containers
    );
    const partReqPri6 = [
      'Part Code', 'Part Name', 'Quantity', 'Requested Date',
      'Destination Location',
      'Quantity Allocated From This Batch', 'Status',
      'Container No.', 'Pick Location', 'Batch Order Date',
      'Order No / Invoice No', 'Case No', 'Type', 'Plant',
      'Shortage Quantity', 'Requested Quantity',
      'Total Qty Available', 'Total Qty Allocated',
      'Net Status', 'In Transit Qty', 'Port ETA',
      'Blocked Qty in Storage', 'Blocked Storage Location'
    ];
    XLSX.utils.book_append_sheet(wbWater,
      buildStyledSheet(shortReqRows.length ? shortReqRows : [{Note:'No shortages.'}], partReqPri6),
      'Shortage Part Request');
    addManualDownloadBtn(wbWater, `SHORTAGE_STATEMENT_Waterfall_${stamp}.xlsx`);

    // FILE 2: Port / In-Transit / Blocked
    const portPriCols = ['Part Code','Part Name','Quantity','Requested Date',
      'Destination Location',
      'Container No.','In Transit Qty','Container Status','Port ETA',
      'Blocked Qty','Blocked Storage Locations','Availability Status'];
    const portReqPriCols = ['Part Code','Part Name','Quantity','Requested Date',
      'Destination Location'];
    const portRows = buildPortTransitReport(ssState.results.shortages, ssState.data.containers, ssState.results.blockedMap);
    const portReqRows = portRows.map(r => ({
      'Part Code': r['Part Code'], 'Part Name': r['Part Name'], 'Quantity': r['Quantity'],
      'Requested Date': r['Requested Date'], 'Destination Location': r['Destination Location']
    }));
    const wbPort = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wbPort,
      buildStyledSheet(portReqRows.length ? portReqRows : [{Note:'No parts at Port/Transit/Blocked.'}], portReqPriCols),
      'Port & Transit (Req Format)');
    XLSX.utils.book_append_sheet(wbPort,
      buildStyledSheet(portRows.length ? portRows : [{Note:'No parts at Port/Transit/Blocked.'}], portPriCols),
      'Port Transit Blocked Detail');
    addManualDownloadBtn(wbPort, `SHORTAGE_STATEMENT_Port_Transit_Blocked_${stamp}.xlsx`);

    // FULL Report Workbook
    const fullSummaryRows = [
      {
        'Report Name': 'SHORTAGE STATEMENT FULL REPORT',
        'Requests Processed': ssState.results.summary.totalReqs,
        'Fully Fulfilled': ssState.results.summary.fulfilledFull,
        'Shortage Count': ssState.results.summary.shortageCount,
        'Generated At': new Date().toLocaleString()
      }
    ];

    createWorkbook(ssState.results.fulfilled, ssState.results.shortages, fullSummaryRows, partMasterList, 'FULL');

    showToast(`Downloads started! Click manual buttons below if browser blocked any files.`, 'success', 6000);
  } catch (err) {
    showToast(`Export error: ${err.message}`, 'error');
    console.error(err);
  }
}

// ─── RESET ────────────────────────────────────────────────────
window.ssDownloadSSExcel = ssDownloadExcel;

function ssReset() {
  ['plan', 'inventory', 'container'].forEach(ssRemoveFile);
  document.getElementById('ss-results-section').hidden = true;
  ssState.results = { fulfilled: [], shortages: [], summary: {} };
  document.getElementById('ss-upload-section').scrollIntoView({ behavior: 'smooth' });
  showToast('Ready for a new Shortage Statement batch.', 'info');
}
