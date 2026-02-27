// ============================================================
// BioPharma Indirect Spend Tracker - Application Logic
// ============================================================

(function() {
  'use strict';

  // ---- Constants ----
  const CATEGORIES = [
    'Clinical, Lab and scientific services',
    'Production Equipment',
    'External Warehouse and distribution',
    'Professional Services',
    'Miscellaneous Indirect Costs',
    'Office and Print'
  ];

  const CATEGORY_COLORS = {
    'Clinical, Lab and scientific services': '#3b82f6',
    'Production Equipment': '#f59e0b',
    'External Warehouse and distribution': '#8b5cf6',
    'Professional Services': '#06b6d4',
    'Miscellaneous Indirect Costs': '#10b981',
    'Office and Print': '#ec4899'
  };

  // Currency state
  const EUR_USD_RATE = 0.851; // 1 EUR = 0.851 USD (2026)
  let displayCurrency = 'EUR'; // 'EUR' or 'USD'

  const COLUMNS = [
    { key: 'date', label: 'Date', type: 'text' },
    { key: 'cost_category', label: 'Category', type: 'select', options: CATEGORIES },
    { key: 'sub_category', label: 'Sub-Category', type: 'text' },
    { key: 'sku', label: 'SKU', type: 'text' },
    { key: 'item_description', label: 'Description', type: 'text' },
    { key: 'supplier', label: 'Supplier', type: 'text' },
    { key: 'ordered_by', label: 'Ordered By', type: 'text' },
    { key: 'department', label: 'Department', type: 'text' },
    { key: 'cost_center', label: 'Cost Center', type: 'text' },
    { key: 'po_number', label: 'PO Number', type: 'text' },
    { key: 'quantity', label: 'Qty', type: 'number' },
    { key: 'unit_price_usd', label: 'Unit Price', type: 'number' },
    { key: 'total_amount_usd', label: 'Total (EUR)', type: 'number' },
    { key: 'budget_type', label: 'Budget Type', type: 'select', options: ['Actual','Baseline','Target'] },
    { key: 'price_impact_usd', label: 'Price Impact', type: 'number' },
    { key: 'volume_impact_usd', label: 'Volume Impact', type: 'number' },
    { key: 'insourcing_savings_usd', label: 'Insourcing', type: 'number' },
    { key: 'notes', label: 'Notes', type: 'text' }
  ];

  const STORAGE_KEY = 'biopharma_indirect_spend_data';
  const TARGETS_KEY = 'biopharma_category_targets';
  const MONTH_NAMES = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];

  // ---- State ----
  let allData = [];
  let filteredData = [];
  let charts = {};
  let currentPage = 1;
  let rowsPerPage = 50;
  let sortColumn = null;
  let sortDirection = 'asc';
  let columnFilters = {};
  let showColumnFilters = false;
  let selectedCategory = null;
  let editingRowIndex = -1;
  let confirmCallback = null;
  let activePage = 'overview';
  let dirtyPages = new Set();
  let aiAdvisorKey = '';
  let categoryTargets = {};
  const AI_KEY_STORAGE = 'biopharma_claude_api_key';

  // ---- Utility Functions ----
  function fmt(n, decimals = 0) {
    if (n == null || isNaN(n)) return '--';
    return Number(n).toLocaleString('en-US', { minimumFractionDigits: decimals, maximumFractionDigits: decimals });
  }

  function fmtK(n) {
    if (n == null || isNaN(n)) return '--';
    const val = displayCurrency === 'USD' ? Number(n) * EUR_USD_RATE : Number(n);
    return (val / 1000).toLocaleString('en-US', { minimumFractionDigits: 1, maximumFractionDigits: 1 });
  }

  function fmtEUR(n) {
    if (n == null || isNaN(n)) return '--';
    const val = displayCurrency === 'USD' ? Number(n) * EUR_USD_RATE : Number(n);
    const sym = displayCurrency === 'USD' ? '$' : '€';
    return sym + val.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }

  function curLabel(base) {
    return displayCurrency === 'USD' ? base.replace('EUR', 'USD') : base;
  }

  function curSym() { return displayCurrency === 'USD' ? '$' : '€'; }
  function curK() { return displayCurrency === 'USD' ? 'k USD' : 'k EUR'; }
  function convK(v) { return displayCurrency === 'USD' ? v * EUR_USD_RATE : v; }

  function pct(n, total) {
    if (!total) return '0.0';
    return ((n / total) * 100).toFixed(1);
  }

  function parseNum(v) {
    if (v == null || v === '') return 0;
    let s = String(v).replace(/[€\s]/g, '').trim();
    // EU format detection: comma exists AND is after the last dot → dot=thousands, comma=decimal
    const lastComma = s.lastIndexOf(',');
    const lastDot = s.lastIndexOf('.');
    if (lastComma > -1 && lastComma > lastDot) {
      s = s.replace(/\./g, '').replace(',', '.');
    } else {
      s = s.replace(/,/g, '');
    }
    const n = parseFloat(s);
    return isNaN(n) ? 0 : n;
  }

  function getMonthKey(dateStr) {
    return dateStr ? dateStr.substring(0, 7) : '';
  }

  function getYear(dateStr) {
    return dateStr ? parseInt(dateStr.substring(0, 4)) : null;
  }

  function getMonth(dateStr) {
    return dateStr ? parseInt(dateStr.substring(5, 7)) : null;
  }

  function monthLabel(mk) {
    const parts = mk.split('-');
    if (parts.length < 2) return mk;
    return MONTH_NAMES[parseInt(parts[1]) - 1] + ' ' + parts[0].slice(2);
  }

  function debounce(fn, ms) {
    let t;
    return function(...args) { clearTimeout(t); t = setTimeout(() => fn.apply(this, args), ms); };
  }

  function toast(message, type = 'info') {
    const container = document.getElementById('toast-container');
    const el = document.createElement('div');
    el.className = 'toast ' + type;
    const msg = document.createElement('div');
    msg.className = 'toast-message';
    msg.textContent = message;
    el.appendChild(msg);
    container.appendChild(el);
    setTimeout(() => { el.style.opacity = '0'; setTimeout(() => el.remove(), 300); }, 4000);
  }

  function showConfirm(title, message, callback) {
    document.getElementById('confirm-title').textContent = title;
    document.getElementById('confirm-message').textContent = message;
    confirmCallback = callback;
    document.getElementById('modal-confirm').classList.add('active');
  }

  function hideConfirm() {
    document.getElementById('modal-confirm').classList.remove('active');
    confirmCallback = null;
  }

  // ---- Data Management ----
  function loadFromStorage() {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (raw) {
        allData = JSON.parse(raw);
        reindexData();
        return true;
      }
    } catch (e) { console.error('Load error:', e); }
    return false;
  }

  function saveToStorage() {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(allData));
    } catch (e) { console.error('Save error:', e); toast('Failed to save data: ' + e.message, 'error'); }
  }

  function loadTargets() {
    try { categoryTargets = JSON.parse(localStorage.getItem(TARGETS_KEY) || '{}'); } catch(e) { categoryTargets = {}; }
  }

  function saveTargets() {
    try { localStorage.setItem(TARGETS_KEY, JSON.stringify(categoryTargets)); } catch(e) {}
  }

  function getBudgetTarget(cat, year) {
    return (categoryTargets[year] && categoryTargets[year][cat] != null)
      ? categoryTargets[year][cat] : null;
  }

  function applyGlobalFilters() {
    const yearVal = document.getElementById('filter-year').value;
    const monthVal = document.getElementById('filter-month').value;
    const catVal = document.getElementById('filter-category').value;

    filteredData = allData.filter(row => {
      if (yearVal !== 'all' && getYear(row.date) !== parseInt(yearVal)) return false;
      if (monthVal !== 'all' && getMonth(row.date) !== parseInt(monthVal)) return false;
      if (catVal !== 'all' && row.cost_category !== catVal) return false;
      return true;
    });
  }

  function updateFilterOptions() {
    const yearSelect = document.getElementById('filter-year');
    const catSelect = document.getElementById('filter-category');
    const years = [...new Set(allData.map(r => getYear(r.date)).filter(Boolean))].sort();
    const cats = [...new Set(allData.map(r => r.cost_category).filter(Boolean))].sort();

    const yearVal = yearSelect.value;
    yearSelect.innerHTML = '<option value="all">All Years</option>' +
      years.map(y => '<option value="' + y + '">' + y + '</option>').join('');
    yearSelect.value = yearVal;

    const catVal = catSelect.value;
    catSelect.innerHTML = '<option value="all">All Categories</option>' +
      cats.map(c => '<option value="' + c + '">' + c + '</option>').join('');
    catSelect.value = catVal;
  }

  function processCSV(csvText, append = false) {
    Papa.parse(csvText, {
      header: true,
      skipEmptyLines: true,
      dynamicTyping: false,
      complete: function(results) {
        if (results.errors.length > 0) {
          toast('CSV parsing warnings: ' + results.errors.length + ' issues found', 'warning');
        }

        // ---- Izvoz spend format detection ----
        const headers = results.meta.fields || [];
        const isIzvoz = headers.some(h => h.trim().toLowerCase().includes('indirect category mapping'));
        if (isIzvoz) {
          processIzvozData(results.data, append);
          return;
        }

        const newData = results.data.map(row => {
          const clean = {};
          COLUMNS.forEach(col => {
            let val = row[col.key];
            if (val === undefined || val === null) val = '';
            if (col.type === 'number') {
              clean[col.key] = parseNum(val);
            } else {
              clean[col.key] = String(val).trim();
            }
          });
          if (clean.total_amount_usd === 0 && clean.quantity && clean.unit_price_usd) {
            clean.total_amount_usd = clean.quantity * clean.unit_price_usd;
          }
          if (!clean.budget_type) clean.budget_type = 'Actual';
          return clean;
        });

        if (append) {
          const existingKeys = new Set(allData.map(r => {
            return r.po_number ? r.po_number : (r.date + '|' + r.sku + '|' + r.supplier + '|' + r.total_amount_usd);
          }));
          const unique = newData.filter(r => {
            const k = r.po_number ? r.po_number : (r.date + '|' + r.sku + '|' + r.supplier + '|' + r.total_amount_usd);
            return !existingKeys.has(k);
          });
          const skipped = newData.length - unique.length;
          allData = allData.concat(unique);
          reindexData();
          saveToStorage();
          updateFilterOptions();
          applyGlobalFilters();
          refreshAll();
          toast('Added ' + unique.length + ' new records' + (skipped ? ' — ' + skipped + ' duplicates skipped' : ''), 'success');
        } else {
          allData = newData;
          reindexData();
          saveToStorage();
          updateFilterOptions();
          applyGlobalFilters();
          refreshAll();
          toast('Loaded ' + newData.length + ' records successfully', 'success');
        }
        updateFooter();
      },
      error: function(err) {
        toast('CSV parse error: ' + err.message, 'error');
      }
    });
  }

  function processIzvozData(rows, append) {
    // Find column headers (flexible matching)
    const findCol = (match) => Object.keys(rows[0] || {}).find(h => h.toLowerCase().includes(match));
    const catCol = findCol('indirect category') || findCol('category');
    const vendorCol = findCol('vendor');
    const spendCol = findCol('ytd spend') || findCol('spend');
    const targetCol = findCol('target');

    // Aggregate targets per category
    const targetAgg = {};
    const newData = [];

    rows.forEach(row => {
      const category = (row[catCol] || '').trim();
      if (!category) return;

      const vendor = (row[vendorCol] || '').trim();
      const spendRaw = parseNum(row[spendCol]);
      const spendK = Math.abs(spendRaw); // values are in k EUR, negative = spend
      const spendEUR = spendK * 1000; // convert from k EUR to EUR

      // Aggregate target per category
      if (targetCol && row[targetCol]) {
        const tgtRaw = parseNum(row[targetCol]);
        if (tgtRaw !== 0) {
          const tgtEUR = Math.abs(tgtRaw) * 1000;
          if (!targetAgg[category]) targetAgg[category] = 0;
          targetAgg[category] += tgtEUR;
        }
      }

      newData.push({
        date: '2025-12',
        cost_category: category,
        sub_category: '',
        sku: '',
        item_description: vendor || category,
        supplier: vendor,
        ordered_by: '',
        department: '',
        cost_center: '',
        po_number: '',
        quantity: 1,
        unit_price_usd: spendEUR,
        total_amount_usd: spendEUR,
        budget_type: 'Actual',
        price_impact_usd: 0,
        volume_impact_usd: 0,
        insourcing_savings_usd: 0,
        notes: ''
      });
    });

    // Save targets for 2026 (only if we have targets)
    if (Object.keys(targetAgg).length > 0) {
      if (!categoryTargets['2026']) categoryTargets['2026'] = {};
      Object.entries(targetAgg).forEach(([cat, val]) => {
        categoryTargets['2026'][cat] = val;
      });
      saveTargets();
    }

    if (append) {
      allData = allData.concat(newData);
    } else {
      allData = newData;
    }
    reindexData();
    saveToStorage();
    updateFilterOptions();
    applyGlobalFilters();
    refreshAll();
    updateFooter();
    toast('Loaded ' + newData.length + ' records from Izvoz spend data' +
      (Object.keys(targetAgg).length > 0 ? ' — Budget targets for 2026 imported' : ''), 'success');
  }

  function exportCSV(data, filename) {
    const headers = COLUMNS.map(c => c.key);
    let csv = headers.join(',') + '\n';
    data.forEach(row => {
      csv += headers.map(h => {
        let val = row[h] != null ? String(row[h]) : '';
        if (val.includes(',') || val.includes('"') || val.includes('\n')) {
          val = '"' + val.replace(/"/g, '""') + '"';
        }
        return val;
      }).join(',') + '\n';
    });
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    triggerDownload(blob, filename);
  }

  function triggerDownload(blob, filename) {
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    setTimeout(() => URL.revokeObjectURL(url), 100);
  }

  function downloadTemplate() {
    const headers = COLUMNS.map(c => c.key).join(',');
    const example = '2026-01,"Clinical, Lab and scientific services",Analytical testing,LAB-0042,HPLC Column C18 250mm,Biorelliance,Jan Novak,QC Laboratory,CC-4200,PO-2026-0142,10,450.00,4500.00,Actual,-120.00,-50.00,0,Sample entry';
    const csv = headers + '\n' + example + '\n';
    triggerDownload(new Blob([csv], { type: 'text/csv;charset=utf-8;' }), 'indirect_spend_template.csv');
  }

  function reindexData() {
    allData.forEach((row, i) => { row._idx = i; });
  }

  function updateFooter() {
    document.getElementById('record-count').textContent = allData.length;
    document.getElementById('last-update').textContent = allData.length > 0 ? new Date().toLocaleDateString() : '--';
  }

  // ---- Render a single page by ID ----
  function renderPage(pageId) {
    switch (pageId) {
      case 'overview':    renderKPIs(); renderCategorySummary(); renderTopLists(); renderCharts(); break;
      case 'datatable':   renderDataTable(); break;
      case 'categories':  renderCategoryPage(); break;
      case 'savings':     renderSavingsPage(); break;
      case 'suppliers':   renderSupplierPage(); break;
      case 'requesters':  renderRequesterPage(); break;
      case 'datamanage':  updateDataSummary(); renderTargetSettings(); break;
      case 'ai-advisor':  renderAIAdvisor(); break;
    }
    dirtyPages.delete(pageId);
  }

  // ---- Refresh All Views ----
  function refreshAll() {
    renderPage(activePage);
    ['overview','datatable','categories','savings','suppliers','requesters','datamanage','ai-advisor']
      .filter(p => p !== activePage)
      .forEach(p => dirtyPages.add(p));
  }

  // ---- KPI Cards ----
  function renderKPIs() {
    const grid = document.getElementById('kpi-grid');
    const data = filteredData.filter(r => r.budget_type === 'Actual' || !r.budget_type || r.budget_type === '');
    const totalSpend = data.reduce((s, r) => s + r.total_amount_usd, 0);
    const uniqueSKUs = new Set(data.map(r => r.sku).filter(Boolean)).size;
    const uniqueSuppliers = new Set(data.map(r => r.supplier).filter(Boolean)).size;
    const uniqueRequesters = new Set(data.map(r => r.ordered_by).filter(Boolean)).size;
    const totalPriceImpact = data.reduce((s, r) => s + (r.price_impact_usd || 0), 0);
    const totalVolumeImpact = data.reduce((s, r) => s + (r.volume_impact_usd || 0), 0);
    const totalInsourcing = data.reduce((s, r) => s + (r.insourcing_savings_usd || 0), 0);
    const totalSavings = totalPriceImpact + totalVolumeImpact + totalInsourcing;

    const avgOrderValue = data.length > 0 ? totalSpend / data.length : 0;

    grid.innerHTML = `
      <div class="kpi-card">
        <div class="kpi-icon blue"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="24" height="24"><path d="M12 2v20M17 5H9.5a3.5 3.5 0 000 7h5a3.5 3.5 0 010 7H6"/></svg></div>
        <div class="kpi-label">Total Spend</div>
        <div class="kpi-value">${fmtEUR(totalSpend)}</div>
        <div class="kpi-change neutral">${fmtK(totalSpend)}${curK()}</div>
      </div>
      <div class="kpi-card">
        <div class="kpi-icon green"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="24" height="24"><polyline points="23 6 13.5 15.5 8.5 10.5 1 18"/></svg></div>
        <div class="kpi-label">Total Savings</div>
        <div class="kpi-value">${fmtEUR(Math.abs(totalSavings))}</div>
        <div class="kpi-change ${totalSavings <= 0 ? 'positive' : 'negative'}">${totalSavings <= 0 ? 'Savings' : 'Increase'}</div>
      </div>
      <div class="kpi-card">
        <div class="kpi-icon cyan"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="24" height="24"><path d="M20 7h-3a2 2 0 01-2-2V2"/><path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z"/></svg></div>
        <div class="kpi-label">Unique SKUs</div>
        <div class="kpi-value">${fmt(uniqueSKUs)}</div>
        <div class="kpi-change neutral">${fmt(data.length)} line items</div>
      </div>
      <div class="kpi-card">
        <div class="kpi-icon orange"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="24" height="24"><path d="M20 21v-2a4 4 0 00-4-4H8a4 4 0 00-4 4v2"/><circle cx="12" cy="7" r="4"/></svg></div>
        <div class="kpi-label">Suppliers</div>
        <div class="kpi-value">${fmt(uniqueSuppliers)}</div>
        <div class="kpi-change neutral">Avg order: ${fmtEUR(avgOrderValue)}</div>
      </div>
      <div class="kpi-card">
        <div class="kpi-icon red"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="24" height="24"><path d="M17 21v-2a4 4 0 00-4-4H5a4 4 0 00-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 00-3-3.87M16 3.13a4 4 0 010 7.75"/></svg></div>
        <div class="kpi-label">Requesters</div>
        <div class="kpi-value">${fmt(uniqueRequesters)}</div>
        <div class="kpi-change neutral">People ordering</div>
      </div>
    `;
  }

  // ---- Category Summary Table ----
  function renderCategorySummary() {
    const body = document.getElementById('category-summary-body');
    const foot = document.getElementById('category-summary-foot');
    const thead = document.querySelector('#category-summary-table thead tr');
    const data = filteredData;

    // Determine if we have a specific year filter active for budget columns
    const yearVal = document.getElementById('filter-year').value;
    const showBudget = yearVal !== 'all';

    // Update header
    if (thead) {
      let headerHtml = '<th>Cost Category</th><th>ACT Spend</th><th>Baseline</th><th>Price Impact</th><th>Volume Impact</th><th>Insourcing</th><th>Target Prelim.</th>';
      if (showBudget) {
        headerHtml += `<th>Budget TGT${yearVal}</th><th>vs Budget</th>`;
      }
      thead.innerHTML = headerHtml;
    }

    const catData = {};
    CATEGORIES.forEach(c => { catData[c] = { actual: 0, baseline: 0, price: 0, volume: 0, insourcing: 0 }; });

    data.forEach(row => {
      const cat = row.cost_category;
      if (!catData[cat]) catData[cat] = { actual: 0, baseline: 0, price: 0, volume: 0, insourcing: 0 };
      const bt = (row.budget_type || 'Actual').toLowerCase();
      if (bt === 'actual' || bt === '') catData[cat].actual += row.total_amount_usd;
      else if (bt === 'baseline') catData[cat].baseline += row.total_amount_usd;
      catData[cat].price += row.price_impact_usd || 0;
      catData[cat].volume += row.volume_impact_usd || 0;
      catData[cat].insourcing += row.insourcing_savings_usd || 0;
    });

    let totals = { actual: 0, baseline: 0, price: 0, volume: 0, insourcing: 0, budget: 0, hasBudget: false };
    let html = '';
    CATEGORIES.forEach(cat => {
      const d = catData[cat];
      const target = d.actual + d.price + d.volume + d.insourcing;
      totals.actual += d.actual;
      totals.baseline += d.baseline;
      totals.price += d.price;
      totals.volume += d.volume;
      totals.insourcing += d.insourcing;

      html += '<tr data-category="' + cat + '" style="cursor:pointer">';
      html += '<td>' + cat + '</td>';
      html += '<td>' + fmtK(d.actual) + '</td>';
      html += '<td>' + (d.baseline ? fmtK(d.baseline) : fmtK(d.actual)) + '</td>';
      html += '<td class="' + (d.price < 0 ? 'positive' : d.price > 0 ? 'negative' : '') + '">' + (d.price ? fmtK(d.price) : '') + '</td>';
      html += '<td class="' + (d.volume < 0 ? 'positive' : d.volume > 0 ? 'negative' : '') + '">' + (d.volume ? fmtK(d.volume) : '') + '</td>';
      html += '<td class="' + (d.insourcing < 0 ? 'positive' : '') + '">' + (d.insourcing ? fmtK(d.insourcing) : '') + '</td>';
      html += '<td class="positive">' + fmtK(target) + '</td>';
      if (showBudget) {
        const budget = getBudgetTarget(cat, yearVal);
        if (budget != null) {
          totals.budget += budget;
          totals.hasBudget = true;
          const variance = d.actual - budget;
          html += '<td>' + fmtK(budget) + '</td>';
          html += '<td class="' + (variance <= 0 ? 'variance-neg' : 'variance-pos') + '">' + (variance <= 0 ? '' : '+') + fmtK(variance) + '</td>';
        } else {
          html += '<td>—</td><td>—</td>';
        }
      }
      html += '</tr>';
    });
    body.innerHTML = html;

    const totalTarget = totals.actual + totals.price + totals.volume + totals.insourcing;
    let footHtml = '<tr><td>TOTAL</td><td>' + fmtK(totals.actual) + '</td><td>' +
      (totals.baseline ? fmtK(totals.baseline) : fmtK(totals.actual)) + '</td><td class="' +
      (totals.price < 0 ? 'positive' : '') + '">' + fmtK(totals.price) + '</td><td class="' +
      (totals.volume < 0 ? 'positive' : '') + '">' + fmtK(totals.volume) + '</td><td class="' +
      (totals.insourcing < 0 ? 'positive' : '') + '">' + fmtK(totals.insourcing) + '</td><td class="positive">' + fmtK(totalTarget) + '</td>';
    if (showBudget) {
      if (totals.hasBudget) {
        const totalVariance = totals.actual - totals.budget;
        footHtml += '<td>' + fmtK(totals.budget) + '</td>';
        footHtml += '<td class="' + (totalVariance <= 0 ? 'variance-neg' : 'variance-pos') + '">' + (totalVariance <= 0 ? '' : '+') + fmtK(totalVariance) + '</td>';
      } else {
        footHtml += '<td>—</td><td>—</td>';
      }
    }
    footHtml += '</tr>';
    foot.innerHTML = footHtml;

    // Row click -> navigate to category analysis (event delegation)
    body.onclick = (e) => {
      const tr = e.target.closest('tr[data-category]');
      if (tr) {
        selectedCategory = tr.dataset.category;
        dirtyPages.add('categories');
        navigateTo('categories');
      }
    };
  }

  // ---- Top Lists ----
  function renderTopLists() {
    const data = filteredData.filter(r => !r.budget_type || r.budget_type === 'Actual' || r.budget_type === '');

    // Top SKUs
    const skuMap = {};
    data.forEach(r => {
      const key = r.sku || 'Unknown';
      if (!skuMap[key]) skuMap[key] = { name: r.item_description || key, amount: 0 };
      skuMap[key].amount += r.total_amount_usd;
    });
    const topSKUs = Object.entries(skuMap).sort((a, b) => b[1].amount - a[1].amount).slice(0, 10);
    const maxSKU = topSKUs.length > 0 ? topSKUs[0][1].amount : 1;
    document.getElementById('top-skus-list').innerHTML = topSKUs.map(([sku, d]) =>
      '<li class="top-list-item"><span class="name" title="' + sku + '">' + (d.name.length > 30 ? d.name.slice(0, 28) + '..' : d.name) +
      '</span><div class="bar-container"><div class="bar-fill" style="width:' + (d.amount / maxSKU * 100) + '%"></div></div><span class="amount">' + fmtEUR(d.amount) + '</span></li>'
    ).join('') || '<li class="top-list-item"><span class="name">No data</span></li>';

    // Top Suppliers
    const supMap = {};
    data.forEach(r => {
      const key = r.supplier || 'Unknown';
      if (!supMap[key]) supMap[key] = 0;
      supMap[key] += r.total_amount_usd;
    });
    const topSup = Object.entries(supMap).sort((a, b) => b[1] - a[1]).slice(0, 10);
    const maxSup = topSup.length > 0 ? topSup[0][1] : 1;
    document.getElementById('top-suppliers-list').innerHTML = topSup.map(([name, amount]) =>
      '<li class="top-list-item"><span class="name">' + (name.length > 25 ? name.slice(0, 23) + '..' : name) +
      '</span><div class="bar-container"><div class="bar-fill" style="width:' + (amount / maxSup * 100) + '%;background:#8b5cf6"></div></div><span class="amount">' + fmtEUR(amount) + '</span></li>'
    ).join('') || '<li class="top-list-item"><span class="name">No data</span></li>';
  }

  // ---- Charts ----
  function destroyChart(name) {
    if (charts[name]) { charts[name].destroy(); charts[name] = null; }
  }

  function renderCharts() {
    renderMonthlyTrendChart();
    renderCategoryPieChart();
  }

  function renderMonthlyTrendChart() {
    destroyChart('monthlyTrend');
    const data = filteredData.filter(r => !r.budget_type || r.budget_type === 'Actual' || r.budget_type === '');
    const monthMap = {};
    data.forEach(r => {
      const mk = getMonthKey(r.date);
      if (!mk) return;
      if (!monthMap[mk]) monthMap[mk] = 0;
      monthMap[mk] += r.total_amount_usd;
    });
    const months = Object.keys(monthMap).sort();
    const values = months.map(m => monthMap[m] / 1000);

    const ctx = document.getElementById('chart-monthly-trend');
    if (!ctx) return;
    charts.monthlyTrend = new Chart(ctx, {
      type: 'bar',
      data: {
        labels: months.map(monthLabel),
        datasets: [{
          label: 'Spend (' + curK() + ')',
          data: values,
          backgroundColor: '#3b82f6',
          borderRadius: 4,
          barThickness: months.length > 12 ? undefined : 40
        }]
      },
      options: {
        responsive: true, maintainAspectRatio: false,
        plugins: { legend: { display: false }, tooltip: { callbacks: { label: ctx => fmtK(ctx.raw * 1000) + curK() } } },
        scales: {
          y: { beginAtZero: true, grid: { color: 'rgba(255,255,255,0.05)' }, ticks: { color: '#94a3b8', callback: v => v + 'k' } },
          x: { grid: { display: false }, ticks: { color: '#94a3b8' } }
        }
      }
    });
  }

  function renderCategoryPieChart() {
    destroyChart('categoryPie');
    const data = filteredData.filter(r => !r.budget_type || r.budget_type === 'Actual' || r.budget_type === '');
    const catMap = {};
    data.forEach(r => {
      const cat = r.cost_category || 'Miscellaneous Indirect Costs';
      if (!catMap[cat]) catMap[cat] = 0;
      catMap[cat] += r.total_amount_usd;
    });
    const entries = Object.entries(catMap).sort((a, b) => b[1] - a[1]);
    const labels = entries.map(e => e[0]);
    const values = entries.map(e => e[1]);
    const colors = labels.map(l => CATEGORY_COLORS[l] || '#94a3b8');

    const ctx = document.getElementById('chart-category-pie');
    if (!ctx) return;
    charts.categoryPie = new Chart(ctx, {
      type: 'doughnut',
      data: { labels, datasets: [{ data: values, backgroundColor: colors, borderWidth: 0, hoverOffset: 8 }] },
      options: {
        responsive: true, maintainAspectRatio: false, cutout: '55%',
        plugins: {
          legend: { position: 'right', labels: { color: '#e2e8f0', font: { size: 11 }, padding: 12, usePointStyle: true } },
          tooltip: { callbacks: { label: ctx => ctx.label + ': ' + fmtEUR(ctx.raw) + ' (' + pct(ctx.raw, values.reduce((a,b) => a+b, 0)) + '%)' } }
        }
      }
    });
  }

  // ---- Data Table ----
  function renderDataTable() {
    renderTableHeader();
    renderTableFilters();
    renderTableBody();
  }

  function renderTableHeader() {
    const header = document.getElementById('main-table-header');
    header.innerHTML = '<th style="width:36px"></th>' + COLUMNS.map((col, i) =>
      '<th data-col="' + i + '" class="' + (sortColumn === i ? (sortDirection === 'asc' ? 'sorted-asc' : 'sorted-desc') : '') + '">' +
      col.label + '<span class="sort-indicator"></span></th>'
    ).join('');

    header.querySelectorAll('th[data-col]').forEach(th => {
      th.addEventListener('click', () => {
        const colIdx = parseInt(th.dataset.col);
        if (sortColumn === colIdx) {
          sortDirection = sortDirection === 'asc' ? 'desc' : 'asc';
        } else {
          sortColumn = colIdx;
          sortDirection = 'asc';
        }
        currentPage = 1;
        renderDataTable();
      });
    });
  }

  function renderTableFilters() {
    const row = document.getElementById('main-table-filters');
    row.style.display = showColumnFilters ? '' : 'none';
    if (!showColumnFilters) return;

    row.innerHTML = '<th></th>' + COLUMNS.map((col, i) => {
      if (col.type === 'select') {
        const opts = (col.options || []).map(o => '<option value="' + o + '">' + o + '</option>').join('');
        return '<th><select data-filter-col="' + i + '"><option value="">All</option>' + opts + '</select></th>';
      }
      return '<th><input type="text" data-filter-col="' + i + '" placeholder="Filter..." value="' + (columnFilters[i] || '') + '"></th>';
    }).join('');

    row.querySelectorAll('[data-filter-col]').forEach(el => {
      const colIdx = parseInt(el.dataset.filterCol);
      el.value = columnFilters[colIdx] || '';
      el.addEventListener('input', () => {
        columnFilters[colIdx] = el.value;
        currentPage = 1;
        renderTableBody();
      });
      el.addEventListener('change', () => {
        columnFilters[colIdx] = el.value;
        currentPage = 1;
        renderTableBody();
      });
    });
  }

  function getTableData() {
    let data = [...filteredData];
    const searchTerm = (document.getElementById('table-search-input').value || '').toLowerCase();

    // Apply column filters
    Object.entries(columnFilters).forEach(([colIdx, filterVal]) => {
      if (!filterVal) return;
      const col = COLUMNS[parseInt(colIdx)];
      const fv = filterVal.toLowerCase();
      data = data.filter(row => {
        const val = String(row[col.key] || '').toLowerCase();
        return val.includes(fv);
      });
    });

    // Apply search
    if (searchTerm) {
      data = data.filter(row =>
        COLUMNS.some(col => String(row[col.key] || '').toLowerCase().includes(searchTerm))
      );
    }

    // Sort
    if (sortColumn !== null) {
      const col = COLUMNS[sortColumn];
      data.sort((a, b) => {
        let va = a[col.key], vb = b[col.key];
        if (col.type === 'number') {
          va = parseNum(va); vb = parseNum(vb);
          return sortDirection === 'asc' ? va - vb : vb - va;
        }
        va = String(va || '').toLowerCase(); vb = String(vb || '').toLowerCase();
        return sortDirection === 'asc' ? va.localeCompare(vb) : vb.localeCompare(va);
      });
    }

    return data;
  }

  function renderTableBody() {
    const body = document.getElementById('main-table-body');
    const data = getTableData();
    const totalRows = data.length;
    const totalPages = Math.max(1, Math.ceil(totalRows / rowsPerPage));
    if (currentPage > totalPages) currentPage = totalPages;
    const startIdx = (currentPage - 1) * rowsPerPage;
    const pageData = data.slice(startIdx, startIdx + rowsPerPage);

    body.innerHTML = pageData.map((row) => {
      const globalIdx = row._idx;
      return '<tr data-idx="' + globalIdx + '"><td><button class="btn-icon btn-edit-row" data-idx="' + globalIdx + '" title="Edit"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="14" height="14"><path d="M11 4H4a2 2 0 00-2 2v14a2 2 0 002 2h14a2 2 0 002-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 013 3L12 15l-4 1 1-4 9.5-9.5z"/></svg></button></td>' + COLUMNS.map(col => {
        const val = row[col.key];
        if (col.type === 'number') {
          if (col.key === 'total_amount_usd' || col.key === 'unit_price_usd') return '<td class="currency">' + fmtEUR(val) + '</td>';
          if (col.key.includes('impact') || col.key.includes('savings')) {
            const cls = val < 0 ? 'positive' : val > 0 ? 'negative' : '';
            return '<td class="num ' + cls + '">' + (val ? fmtEUR(val) : '') + '</td>';
          }
          return '<td class="num">' + fmt(val) + '</td>';
        }
        return '<td>' + (val || '') + '</td>';
      }).join('') + '</tr>';
    }).join('');

    if (pageData.length === 0) {
      body.innerHTML = '<tr><td colspan="' + (COLUMNS.length + 2) + '" style="text-align:center;padding:40px;color:var(--text-muted)">No data. Go to Data Management to import a CSV file or load sample data.</td></tr>';
    }

    // Edit buttons
    body.querySelectorAll('.btn-edit-row').forEach(btn => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        openEditModal(parseInt(btn.dataset.idx));
      });
    });

    // Pagination
    document.getElementById('pagination-info').textContent = 'Showing ' + (totalRows > 0 ? startIdx + 1 : 0) + '-' + Math.min(startIdx + rowsPerPage, totalRows) + ' of ' + totalRows + ' entries';
    renderPagination(totalPages);
  }

  function renderPagination(totalPages) {
    const controls = document.getElementById('pagination-controls');
    let html = '<button class="pagination-btn" data-page="prev" ' + (currentPage <= 1 ? 'disabled' : '') + '>&laquo; Prev</button>';

    let startPage = Math.max(1, currentPage - 2);
    let endPage = Math.min(totalPages, startPage + 4);
    if (endPage - startPage < 4) startPage = Math.max(1, endPage - 4);

    for (let i = startPage; i <= endPage; i++) {
      html += '<button class="pagination-btn' + (i === currentPage ? ' active' : '') + '" data-page="' + i + '">' + i + '</button>';
    }
    html += '<button class="pagination-btn" data-page="next" ' + (currentPage >= totalPages ? 'disabled' : '') + '>Next &raquo;</button>';
    controls.innerHTML = html;

    controls.querySelectorAll('.pagination-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        const page = btn.dataset.page;
        if (page === 'prev') currentPage = Math.max(1, currentPage - 1);
        else if (page === 'next') currentPage++;
        else currentPage = parseInt(page);
        renderTableBody();
      });
    });
  }

  // ---- Edit Modal ----
  function openEditModal(rowIdx) {
    editingRowIndex = rowIdx;
    const row = allData[rowIdx];
    if (!row) return;

    const body = document.getElementById('edit-row-body');
    body.innerHTML = '<div class="add-row-form" style="display:grid">' +
      COLUMNS.map(col => {
        const val = row[col.key] != null ? row[col.key] : '';
        if (col.type === 'select') {
          const opts = (col.options || []).map(o => '<option value="' + o + '"' + (o === val ? ' selected' : '') + '>' + o + '</option>').join('');
          return '<div class="form-group"><label>' + col.label + '</label><select class="form-control" data-key="' + col.key + '">' + opts + '</select></div>';
        }
        const inputType = col.type === 'number' ? 'number' : (col.key === 'date' ? 'month' : 'text');
        return '<div class="form-group"><label>' + col.label + '</label><input type="' + inputType + '" class="form-control" data-key="' + col.key + '" value="' + val + '" ' + (col.type === 'number' ? 'step="0.01"' : '') + '></div>';
      }).join('') + '</div>';

    const modal = document.getElementById('modal-edit-row');
    modal.classList.add('active');

    // Auto-calculate total when qty or unit price changes
    const qtyEl = body.querySelector('[data-key="quantity"]');
    const priceEl = body.querySelector('[data-key="unit_price_usd"]');
    const totalEl = body.querySelector('[data-key="total_amount_usd"]');
    function recalcTotal() {
      const q = parseNum(qtyEl && qtyEl.value);
      const p = parseNum(priceEl && priceEl.value);
      if (q && p && totalEl) totalEl.value = (q * p).toFixed(2);
    }
    if (qtyEl) qtyEl.addEventListener('input', recalcTotal);
    if (priceEl) priceEl.addEventListener('input', recalcTotal);
  }

  function saveEditRow() {
    if (editingRowIndex < 0) return;
    const body = document.getElementById('edit-row-body');
    const row = allData[editingRowIndex];

    body.querySelectorAll('[data-key]').forEach(el => {
      const key = el.dataset.key;
      const col = COLUMNS.find(c => c.key === key);
      if (col && col.type === 'number') {
        row[key] = parseNum(el.value);
      } else {
        row[key] = el.value;
      }
    });

    if (row.total_amount_usd === 0 && row.quantity && row.unit_price_usd) {
      row.total_amount_usd = row.quantity * row.unit_price_usd;
    }

    const savedIdx = editingRowIndex;
    saveToStorage();
    applyGlobalFilters();
    refreshAll();
    closeEditModal();
    toast('Entry updated', 'success');
    // Flash the updated row green
    requestAnimationFrame(() => {
      const tr = document.querySelector('#main-table-body tr[data-idx="' + savedIdx + '"]');
      if (tr) { tr.classList.add('row-saved'); setTimeout(() => tr.classList.remove('row-saved'), 1200); }
    });
  }

  function deleteEditRow() {
    if (editingRowIndex < 0) return;
    showConfirm('Delete Entry', 'Are you sure you want to delete this entry? This cannot be undone.', () => {
      allData.splice(editingRowIndex, 1);
      reindexData();
      saveToStorage();
      applyGlobalFilters();
      refreshAll();
      updateFooter();
      closeEditModal();
      toast('Entry deleted', 'success');
    });
  }

  function closeEditModal() {
    document.getElementById('modal-edit-row').classList.remove('active');
    editingRowIndex = -1;
  }

  // ---- Add Row ----
  function addNewRow() {
    const fields = {
      date: document.getElementById('add-date').value,
      cost_category: document.getElementById('add-category').value,
      sub_category: document.getElementById('add-subcategory').value,
      sku: document.getElementById('add-sku').value,
      item_description: document.getElementById('add-description').value,
      supplier: document.getElementById('add-supplier').value,
      ordered_by: document.getElementById('add-orderedby').value,
      department: document.getElementById('add-department').value,
      cost_center: document.getElementById('add-costcenter').value,
      po_number: document.getElementById('add-po').value,
      quantity: parseNum(document.getElementById('add-quantity').value),
      unit_price_usd: parseNum(document.getElementById('add-unitprice').value),
      total_amount_usd: parseNum(document.getElementById('add-totalamount').value),
      budget_type: document.getElementById('add-budgettype').value,
      price_impact_usd: 0,
      volume_impact_usd: 0,
      insourcing_savings_usd: 0,
      notes: document.getElementById('add-notes').value
    };

    if (!fields.date || !fields.cost_category || !fields.sku) {
      toast('Date, Category, and SKU are required', 'error');
      return;
    }

    if (fields.total_amount_usd === 0 && fields.quantity && fields.unit_price_usd) {
      fields.total_amount_usd = fields.quantity * fields.unit_price_usd;
    }

    allData.push(fields);
    reindexData();
    saveToStorage();
    updateFilterOptions();
    applyGlobalFilters();
    refreshAll();
    updateFooter();

    // Clear form
    document.getElementById('add-row-form').style.display = 'none';
    ['add-date','add-subcategory','add-sku','add-description','add-supplier','add-orderedby','add-department','add-costcenter','add-po','add-quantity','add-unitprice','add-totalamount','add-notes'].forEach(id => {
      const el = document.getElementById(id);
      if (el) el.value = '';
    });
    document.getElementById('add-category').value = '';
    document.getElementById('add-budgettype').value = 'Actual';

    toast('Entry added successfully', 'success');
  }

  // ---- Category Analysis Page ----
  function renderCategoryPage() {
    renderCategoryChips();
    renderCategoryComparisonChart();
    renderCategoryTrendsChart();
    renderCategoryDetail();
  }

  function renderCategoryChips() {
    const container = document.getElementById('category-chips');
    const data = filteredData;
    const catSpend = {};
    data.forEach(r => {
      if (!catSpend[r.cost_category]) catSpend[r.cost_category] = 0;
      catSpend[r.cost_category] += r.total_amount_usd;
    });

    container.innerHTML = '<div class="filter-chip' + (!selectedCategory ? ' active' : '') + '" data-cat="all">All Categories</div>' +
      CATEGORIES.filter(c => catSpend[c]).map(cat =>
        '<div class="filter-chip' + (selectedCategory === cat ? ' active' : '') + '" data-cat="' + cat + '">' +
        '<span style="display:inline-block;width:8px;height:8px;border-radius:50%;background:' + CATEGORY_COLORS[cat] + '"></span>' +
        cat.split(',')[0] + '</div>'
      ).join('');

    container.querySelectorAll('.filter-chip').forEach(chip => {
      chip.addEventListener('click', () => {
        selectedCategory = chip.dataset.cat === 'all' ? null : chip.dataset.cat;
        renderCategoryPage();
      });
    });
  }

  function renderCategoryComparisonChart() {
    destroyChart('catComparison');
    const data = filteredData;
    const catData = {};
    data.forEach(r => {
      const cat = r.cost_category;
      if (!catData[cat]) catData[cat] = { actual: 0, baseline: 0, target: 0, savings: 0 };
      const bt = (r.budget_type || 'Actual').toLowerCase();
      if (bt === 'actual' || bt === '') {
        catData[cat].actual += r.total_amount_usd;
        catData[cat].savings += (r.price_impact_usd || 0) + (r.volume_impact_usd || 0) + (r.insourcing_savings_usd || 0);
      } else if (bt === 'baseline') {
        catData[cat].baseline += r.total_amount_usd;
      } else if (bt === 'target') {
        catData[cat].target += r.total_amount_usd;
      }
    });

    const cats = Object.keys(catData).sort((a, b) => catData[b].actual - catData[a].actual);
    const yearVal = document.getElementById('filter-year').value;

    const datasets = [
      { label: 'Actual', data: cats.map(c => catData[c].actual / 1000), backgroundColor: '#3b82f6', borderRadius: 3 },
      { label: 'Baseline', data: cats.map(c => (catData[c].baseline || catData[c].actual) / 1000), backgroundColor: '#64748b', borderRadius: 3 },
      { label: 'Target', data: cats.map(c => {
        const d = catData[c];
        return d.target ? d.target / 1000 : (d.actual + d.savings) / 1000;
      }), backgroundColor: '#10b981', borderRadius: 3 }
    ];

    // Add Budget dataset if a specific year is selected and targets exist
    if (yearVal !== 'all') {
      const budgetData = cats.map(c => {
        const b = getBudgetTarget(c, yearVal);
        return b != null ? b / 1000 : null;
      });
      if (budgetData.some(v => v != null)) {
        datasets.push({ label: 'Budget', data: budgetData, backgroundColor: '#f59e0b', borderRadius: 3 });
      }
    }

    const ctx = document.getElementById('chart-category-comparison');
    if (!ctx) return;
    charts.catComparison = new Chart(ctx, {
      type: 'bar',
      data: {
        labels: cats.map(c => c.length > 20 ? c.slice(0, 18) + '..' : c),
        datasets
      },
      options: {
        responsive: true, maintainAspectRatio: false,
        plugins: { legend: { labels: { color: '#e2e8f0', usePointStyle: true } }, tooltip: { callbacks: { label: ctx => ctx.dataset.label + ': ' + (ctx.raw != null ? convK(ctx.raw).toFixed(1) : '—') + curK() } } },
        scales: {
          y: { beginAtZero: true, grid: { color: 'rgba(255,255,255,0.05)' }, ticks: { color: '#94a3b8', callback: v => v + 'k' } },
          x: { grid: { display: false }, ticks: { color: '#94a3b8', maxRotation: 45 } }
        }
      }
    });
  }

  function renderCategoryTrendsChart() {
    destroyChart('catTrends');
    const data = filteredData.filter(r => !r.budget_type || r.budget_type === 'Actual' || r.budget_type === '');
    const monthCat = {};
    const allMonths = new Set();
    const activeCats = selectedCategory ? [selectedCategory] : [...new Set(data.map(r => r.cost_category))];

    data.forEach(r => {
      const mk = getMonthKey(r.date);
      if (!mk) return;
      allMonths.add(mk);
      if (!monthCat[r.cost_category]) monthCat[r.cost_category] = {};
      if (!monthCat[r.cost_category][mk]) monthCat[r.cost_category][mk] = 0;
      monthCat[r.cost_category][mk] += r.total_amount_usd;
    });

    const months = [...allMonths].sort();
    const datasets = activeCats.map(cat => ({
      label: cat.length > 25 ? cat.slice(0, 23) + '..' : cat,
      data: months.map(m => ((monthCat[cat] || {})[m] || 0) / 1000),
      borderColor: CATEGORY_COLORS[cat] || '#94a3b8',
      backgroundColor: (CATEGORY_COLORS[cat] || '#94a3b8') + '20',
      fill: activeCats.length === 1,
      tension: 0.3,
      pointRadius: 4,
      borderWidth: 2
    }));

    const ctx = document.getElementById('chart-category-trends');
    if (!ctx) return;
    charts.catTrends = new Chart(ctx, {
      type: 'line',
      data: { labels: months.map(monthLabel), datasets },
      options: {
        responsive: true, maintainAspectRatio: false,
        plugins: { legend: { labels: { color: '#e2e8f0', usePointStyle: true, font: { size: 10 } } }, tooltip: { callbacks: { label: ctx => ctx.dataset.label + ': ' + ctx.raw.toFixed(1) + 'k' } } },
        scales: {
          y: { beginAtZero: true, grid: { color: 'rgba(255,255,255,0.05)' }, ticks: { color: '#94a3b8', callback: v => v + 'k' } },
          x: { grid: { display: false }, ticks: { color: '#94a3b8' } }
        }
      }
    });
  }

  function renderCategoryDetail() {
    const title = document.getElementById('cat-detail-title');
    const catFilter = selectedCategory || null;
    const data = filteredData.filter(r => {
      if (catFilter && r.cost_category !== catFilter) return false;
      if (!r.budget_type || r.budget_type === 'Actual' || r.budget_type === '') return true;
      return false;
    });

    title.textContent = catFilter ? catFilter + ' - Detail' : 'All Categories - Detail';

    // By SKU
    const skuMap = {};
    data.forEach(r => {
      const k = r.sku || 'Unknown';
      if (!skuMap[k]) skuMap[k] = { desc: r.item_description, spend: 0, qty: 0, count: 0 };
      skuMap[k].spend += r.total_amount_usd;
      skuMap[k].qty += r.quantity;
      skuMap[k].count++;
    });
    const skuRows = Object.entries(skuMap).sort((a, b) => b[1].spend - a[1].spend);
    renderSortableTable('cat-sku-table', skuRows.map(([sku, d]) => [sku, d.desc, d.spend, d.qty, d.qty > 0 ? d.spend / d.qty : 0, d.count]),
      [{ fmt: 'text' }, { fmt: 'text' }, { fmt: 'usd' }, { fmt: 'num' }, { fmt: 'usd' }, { fmt: 'num' }]);

    // By Supplier
    const totalSpend = data.reduce((s, r) => s + r.total_amount_usd, 0);
    const supMap = {};
    data.forEach(r => {
      const k = r.supplier || 'Unknown';
      if (!supMap[k]) supMap[k] = { spend: 0, skus: new Set(), count: 0 };
      supMap[k].spend += r.total_amount_usd;
      supMap[k].skus.add(r.sku);
      supMap[k].count++;
    });
    const supRows = Object.entries(supMap).sort((a, b) => b[1].spend - a[1].spend);
    renderSortableTable('cat-supplier-table', supRows.map(([sup, d]) => [sup, d.spend, d.skus.size, d.count, pct(d.spend, totalSpend) + '%']),
      [{ fmt: 'text' }, { fmt: 'usd' }, { fmt: 'num' }, { fmt: 'num' }, { fmt: 'text' }]);

    // By Requester
    const reqMap = {};
    data.forEach(r => {
      const k = r.ordered_by || 'Unknown';
      if (!reqMap[k]) reqMap[k] = { dept: r.department, spend: 0, count: 0 };
      reqMap[k].spend += r.total_amount_usd;
      reqMap[k].count++;
    });
    const reqRows = Object.entries(reqMap).sort((a, b) => b[1].spend - a[1].spend);
    renderSortableTable('cat-requester-table', reqRows.map(([req, d]) => [req, d.dept, d.spend, d.count, pct(d.spend, totalSpend) + '%']),
      [{ fmt: 'text' }, { fmt: 'text' }, { fmt: 'usd' }, { fmt: 'num' }, { fmt: 'text' }]);

    // Monthly detail chart
    renderCatMonthlyDetail(data);
  }

  function renderCatMonthlyDetail(data) {
    destroyChart('catMonthlyDetail');
    const monthMap = {};
    data.forEach(r => {
      const mk = getMonthKey(r.date);
      if (!mk) return;
      if (!monthMap[mk]) monthMap[mk] = 0;
      monthMap[mk] += r.total_amount_usd;
    });
    const months = Object.keys(monthMap).sort();
    const ctx = document.getElementById('chart-cat-monthly-detail');
    if (!ctx) return;
    charts.catMonthlyDetail = new Chart(ctx, {
      type: 'bar',
      data: {
        labels: months.map(monthLabel),
        datasets: [{ label: 'Spend (EUR)', data: months.map(m => monthMap[m]), backgroundColor: selectedCategory ? CATEGORY_COLORS[selectedCategory] : '#3b82f6', borderRadius: 4 }]
      },
      options: {
        responsive: true, maintainAspectRatio: false,
        plugins: { legend: { display: false }, tooltip: { callbacks: { label: ctx => fmtEUR(ctx.raw) } } },
        scales: {
          y: { beginAtZero: true, grid: { color: 'rgba(255,255,255,0.05)' }, ticks: { color: '#94a3b8', callback: v => curSym() + (convK(v)/1000).toFixed(0) + 'k' } },
          x: { grid: { display: false }, ticks: { color: '#94a3b8' } }
        }
      }
    });
  }

  // ---- Sortable Mini Tables ----
  function renderSortableTable(tableId, rows, colFormats) {
    const table = document.getElementById(tableId);
    const tbody = table.querySelector('tbody');
    const headers = table.querySelectorAll('thead th');

    function render(data) {
      tbody.innerHTML = data.map(row =>
        '<tr>' + row.map((val, i) => {
          const f = colFormats[i];
          if (f.fmt === 'usd') return '<td class="currency">' + fmtEUR(val) + '</td>';
          if (f.fmt === 'num') return '<td class="num">' + fmt(val) + '</td>';
          return '<td>' + (val || '') + '</td>';
        }).join('') + '</tr>'
      ).join('') || '<tr><td colspan="' + colFormats.length + '" style="text-align:center;padding:20px;color:var(--text-muted)">No data</td></tr>';
    }

    render(rows);

    // Make headers sortable
    headers.forEach((th, colIdx) => {
      th.style.cursor = 'pointer';
      th.onclick = () => {
        const currentDir = th.classList.contains('sorted-asc') ? 'asc' : (th.classList.contains('sorted-desc') ? 'desc' : null);
        headers.forEach(h => { h.classList.remove('sorted-asc', 'sorted-desc'); });

        const newDir = currentDir === 'asc' ? 'desc' : 'asc';
        th.classList.add('sorted-' + newDir);

        const sorted = [...rows].sort((a, b) => {
          let va = a[colIdx], vb = b[colIdx];
          if (typeof va === 'number' && typeof vb === 'number') return newDir === 'asc' ? va - vb : vb - va;
          va = String(va || ''); vb = String(vb || '');
          const na = parseFloat(va), nb = parseFloat(vb);
          if (!isNaN(na) && !isNaN(nb)) return newDir === 'asc' ? na - nb : nb - na;
          return newDir === 'asc' ? va.localeCompare(vb) : vb.localeCompare(va);
        });
        render(sorted);
      };
    });
  }

  // ---- Savings Page ----
  function renderSavingsPage() {
    renderSavingsKPIs();
    renderWaterfallChart();
    renderSavingsByCatChart();
    renderSavingsTimeline();
  }

  function renderSavingsKPIs() {
    const data = filteredData;
    const totalPrice = data.reduce((s, r) => s + (r.price_impact_usd || 0), 0);
    const totalVolume = data.reduce((s, r) => s + (r.volume_impact_usd || 0), 0);
    const totalInsourcing = data.reduce((s, r) => s + (r.insourcing_savings_usd || 0), 0);
    const totalSavings = totalPrice + totalVolume + totalInsourcing;
    const totalSpend = data.filter(r => !r.budget_type || r.budget_type === 'Actual').reduce((s, r) => s + r.total_amount_usd, 0);
    const savingsPct = totalSpend > 0 ? (Math.abs(totalSavings) / totalSpend * 100).toFixed(1) : '0.0';

    // vs Budget KPI — only when a specific year is selected and targets exist
    const yearVal = document.getElementById('filter-year').value;
    let vsBudgetHtml = '';
    if (yearVal !== 'all') {
      let totalBudget = 0, hasBudget = false;
      CATEGORIES.forEach(cat => {
        const b = getBudgetTarget(cat, yearVal);
        if (b != null) { totalBudget += b; hasBudget = true; }
      });
      if (hasBudget) {
        const vsBudget = totalBudget - totalSpend; // positive = under budget (good)
        const pctOfBudget = totalBudget > 0 ? Math.abs(vsBudget / totalBudget * 100).toFixed(1) : '0.0';
        vsBudgetHtml = `<div class="kpi-card"><div class="kpi-icon ${vsBudget >= 0 ? 'green' : 'red'}"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="24" height="24"><path d="M12 2v20M17 5H9.5a3.5 3.5 0 000 7h5a3.5 3.5 0 010 7H6"/><polyline points="20 4 4 4"/></svg></div>
          <div class="kpi-label">vs Budget ${yearVal}</div><div class="kpi-value">${fmtEUR(Math.abs(vsBudget))}</div>
          <div class="kpi-change ${vsBudget >= 0 ? 'positive' : 'negative'}">${vsBudget >= 0 ? pctOfBudget + '% under budget' : pctOfBudget + '% over budget'}</div></div>`;
      }
    }

    document.getElementById('savings-kpis').innerHTML = `
      <div class="kpi-card"><div class="kpi-icon red"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="24" height="24"><path d="M12 2v20M17 5H9.5a3.5 3.5 0 000 7h5a3.5 3.5 0 010 7H6"/></svg></div>
        <div class="kpi-label">Price Impact</div><div class="kpi-value">${fmtEUR(totalPrice)}</div>
        <div class="kpi-change ${totalPrice <= 0 ? 'positive' : 'negative'}">${totalPrice <= 0 ? 'Savings' : 'Increase'}</div></div>
      <div class="kpi-card"><div class="kpi-icon orange"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="24" height="24"><polyline points="23 6 13.5 15.5 8.5 10.5 1 18"/></svg></div>
        <div class="kpi-label">Volume Impact</div><div class="kpi-value">${fmtEUR(totalVolume)}</div>
        <div class="kpi-change ${totalVolume <= 0 ? 'positive' : 'negative'}">${totalVolume <= 0 ? 'Reduction' : 'Increase'}</div></div>
      <div class="kpi-card"><div class="kpi-icon blue"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="24" height="24"><path d="M3 9l9-7 9 7v11a2 2 0 01-2 2H5a2 2 0 01-2-2z"/></svg></div>
        <div class="kpi-label">Insourcing Savings</div><div class="kpi-value">${fmtEUR(totalInsourcing)}</div>
        <div class="kpi-change ${totalInsourcing <= 0 ? 'positive' : 'negative'}">${totalInsourcing <= 0 ? 'Savings' : 'Cost'}</div></div>
      <div class="kpi-card"><div class="kpi-icon green"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="24" height="24"><circle cx="12" cy="12" r="10"/><path d="M8 14s1.5 2 4 2 4-2 4-2"/><line x1="9" y1="9" x2="9.01" y2="9"/><line x1="15" y1="9" x2="15.01" y2="9"/></svg></div>
        <div class="kpi-label">Total Savings</div><div class="kpi-value">${fmtEUR(totalSavings)}</div>
        <div class="kpi-change ${totalSavings <= 0 ? 'positive' : 'negative'}">${savingsPct}% of spend</div></div>
      ${vsBudgetHtml}
    `;
  }

  function renderWaterfallChart() {
    destroyChart('waterfall');
    const data = filteredData;
    const actuals = data.filter(r => !r.budget_type || r.budget_type === 'Actual' || r.budget_type === '');
    const baseline = actuals.reduce((s, r) => s + r.total_amount_usd, 0);
    const priceImpact = data.reduce((s, r) => s + (r.price_impact_usd || 0), 0);
    const volumeImpact = data.reduce((s, r) => s + (r.volume_impact_usd || 0), 0);
    const insourcing = data.reduce((s, r) => s + (r.insourcing_savings_usd || 0), 0);
    const target = baseline + priceImpact + volumeImpact + insourcing;

    const labels = ['Baseline Spend', 'Price Impact', 'Volume Impact', 'Insourcing', 'Target'];
    let running = baseline;
    const bgColors = [];
    const barData = [];
    const hiddenData = [];

    // Baseline
    barData.push(baseline / 1000);
    hiddenData.push(0);
    bgColors.push('#1e293b');

    // Price Impact
    const pv = priceImpact / 1000;
    if (pv < 0) {
      hiddenData.push((running + priceImpact) / 1000);
      barData.push(Math.abs(pv));
      bgColors.push('#10b981');
    } else {
      hiddenData.push(running / 1000);
      barData.push(pv);
      bgColors.push('#ef4444');
    }
    running += priceImpact;

    // Volume Impact
    const vv = volumeImpact / 1000;
    if (vv < 0) {
      hiddenData.push((running + volumeImpact) / 1000);
      barData.push(Math.abs(vv));
      bgColors.push('#10b981');
    } else {
      hiddenData.push(running / 1000);
      barData.push(vv);
      bgColors.push('#ef4444');
    }
    running += volumeImpact;

    // Insourcing
    const iv = insourcing / 1000;
    if (iv < 0) {
      hiddenData.push((running + insourcing) / 1000);
      barData.push(Math.abs(iv));
      bgColors.push('#3b82f6');
    } else {
      hiddenData.push(running / 1000);
      barData.push(iv);
      bgColors.push('#3b82f6');
    }
    running += insourcing;

    // Target
    barData.push(target / 1000);
    hiddenData.push(0);
    bgColors.push('#1e293b');

    const ctx = document.getElementById('chart-waterfall');
    if (!ctx) return;
    charts.waterfall = new Chart(ctx, {
      type: 'bar',
      data: {
        labels,
        datasets: [
          { label: 'Hidden', data: hiddenData, backgroundColor: 'transparent', borderWidth: 0, barThickness: 60 },
          { label: 'Value', data: barData, backgroundColor: bgColors, borderRadius: 3, barThickness: 60 }
        ]
      },
      options: {
        responsive: true, maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          tooltip: {
            filter: (item) => item.datasetIndex === 1,
            callbacks: {
              label: ctx => {
                const idx = ctx.dataIndex;
                if (idx === 0) return 'Baseline: ' + convK(ctx.raw).toFixed(1) + curK();
                if (idx === labels.length - 1) return 'Target: ' + convK(ctx.raw).toFixed(1) + curK();
                return labels[idx] + ': ' + (ctx.raw > 0 ? '-' : '+') + convK(ctx.raw).toFixed(1) + curK();
              }
            }
          }
        },
        scales: {
          x: { stacked: true, grid: { display: false }, ticks: { color: '#94a3b8' } },
          y: { stacked: true, grid: { color: 'rgba(255,255,255,0.05)' }, ticks: { color: '#94a3b8', callback: v => v + 'k' } }
        }
      }
    });
  }

  function renderSavingsByCatChart() {
    destroyChart('savingsByCat');
    const data = filteredData;
    const catSavings = {};
    data.forEach(r => {
      const cat = r.cost_category;
      if (!catSavings[cat]) catSavings[cat] = { price: 0, volume: 0, insourcing: 0 };
      catSavings[cat].price += r.price_impact_usd || 0;
      catSavings[cat].volume += r.volume_impact_usd || 0;
      catSavings[cat].insourcing += r.insourcing_savings_usd || 0;
    });

    const cats = Object.keys(catSavings).filter(c => {
      const d = catSavings[c];
      return d.price !== 0 || d.volume !== 0 || d.insourcing !== 0;
    }).sort((a, b) => {
      const ta = catSavings[a].price + catSavings[a].volume + catSavings[a].insourcing;
      const tb = catSavings[b].price + catSavings[b].volume + catSavings[b].insourcing;
      return ta - tb;
    });

    const ctx = document.getElementById('chart-savings-by-cat');
    if (!ctx) return;
    charts.savingsByCat = new Chart(ctx, {
      type: 'bar',
      data: {
        labels: cats.map(c => c.length > 20 ? c.slice(0, 18) + '..' : c),
        datasets: [
          { label: 'Price Impact', data: cats.map(c => catSavings[c].price / 1000), backgroundColor: '#ef4444', borderRadius: 2 },
          { label: 'Volume Impact', data: cats.map(c => catSavings[c].volume / 1000), backgroundColor: '#f59e0b', borderRadius: 2 },
          { label: 'Insourcing', data: cats.map(c => catSavings[c].insourcing / 1000), backgroundColor: '#3b82f6', borderRadius: 2 }
        ]
      },
      options: {
        indexAxis: 'y',
        responsive: true, maintainAspectRatio: false,
        plugins: { legend: { labels: { color: '#e2e8f0', usePointStyle: true } }, tooltip: { callbacks: { label: ctx => ctx.dataset.label + ': ' + convK(ctx.raw).toFixed(2) + curK() } } },
        scales: {
          x: { stacked: true, grid: { color: 'rgba(255,255,255,0.05)' }, ticks: { color: '#94a3b8', callback: v => v + 'k' } },
          y: { stacked: true, grid: { display: false }, ticks: { color: '#94a3b8' } }
        }
      }
    });
  }

  function renderSavingsTimeline() {
    destroyChart('savingsTimeline');
    const data = filteredData;
    const monthSavings = {};
    data.forEach(r => {
      const mk = getMonthKey(r.date);
      if (!mk) return;
      if (!monthSavings[mk]) monthSavings[mk] = 0;
      monthSavings[mk] += (r.price_impact_usd || 0) + (r.volume_impact_usd || 0) + (r.insourcing_savings_usd || 0);
    });
    const months = Object.keys(monthSavings).sort();
    let cumulative = 0;
    const cumData = months.map(m => { cumulative += monthSavings[m]; return Math.abs(cumulative) / 1000; });

    const ctx = document.getElementById('chart-savings-timeline');
    if (!ctx) return;
    charts.savingsTimeline = new Chart(ctx, {
      type: 'line',
      data: {
        labels: months.map(monthLabel),
        datasets: [{
          label: 'Cumulative Savings (' + curK() + ')',
          data: cumData,
          borderColor: '#10b981',
          backgroundColor: 'rgba(16,185,129,0.1)',
          fill: true,
          tension: 0.3,
          pointRadius: 5,
          pointBackgroundColor: '#10b981',
          borderWidth: 3
        }]
      },
      options: {
        responsive: true, maintainAspectRatio: false,
        plugins: { legend: { labels: { color: '#e2e8f0' } }, tooltip: { callbacks: { label: ctx => convK(ctx.raw).toFixed(1) + curK() + ' saved' } } },
        scales: {
          y: { beginAtZero: true, grid: { color: 'rgba(255,255,255,0.05)' }, ticks: { color: '#94a3b8', callback: v => v + 'k' } },
          x: { grid: { display: false }, ticks: { color: '#94a3b8' } }
        }
      }
    });
  }

  // ---- Supplier Analysis Page ----
  let cachedSuppliers = null;
  let cachedSupplierTotalSpend = 0;

  function renderSupplierPage() {
    const data = filteredData.filter(r => !r.budget_type || r.budget_type === 'Actual' || r.budget_type === '');
    const totalSpend = data.reduce((s, r) => s + r.total_amount_usd, 0);

    const supMap = {};
    data.forEach(r => {
      const s = r.supplier || 'Unknown';
      if (!supMap[s]) supMap[s] = { spend: 0, count: 0, skus: new Set(), cats: new Set() };
      supMap[s].spend += r.total_amount_usd;
      supMap[s].count++;
      supMap[s].skus.add(r.sku);
      supMap[s].cats.add(r.cost_category);
    });

    cachedSuppliers = Object.entries(supMap).sort((a, b) => b[1].spend - a[1].spend);
    cachedSupplierTotalSpend = totalSpend;

    renderSupplierBarChart(cachedSuppliers.slice(0, 15));
    renderParetoChart(cachedSuppliers, totalSpend);
    renderSupplierDetailTable(cachedSuppliers, totalSpend);
  }

  function renderSupplierBarChart(top15) {
    destroyChart('supplierBar');
    const ctx = document.getElementById('chart-supplier-bar');
    if (!ctx) return;
    charts.supplierBar = new Chart(ctx, {
      type: 'bar',
      data: {
        labels: top15.map(([n]) => n.length > 18 ? n.slice(0, 16) + '..' : n),
        datasets: [{ label: 'Spend (EUR)', data: top15.map(([, d]) => d.spend), backgroundColor: '#8b5cf6', borderRadius: 4 }]
      },
      options: {
        indexAxis: 'y',
        responsive: true, maintainAspectRatio: false,
        plugins: { legend: { display: false }, tooltip: { callbacks: { label: ctx => fmtEUR(ctx.raw) } } },
        scales: {
          x: { beginAtZero: true, grid: { color: 'rgba(255,255,255,0.05)' }, ticks: { color: '#94a3b8', callback: v => curSym() + (convK(v)/1000).toFixed(0) + 'k' } },
          y: { grid: { display: false }, ticks: { color: '#94a3b8' } }
        }
      }
    });
  }

  function renderParetoChart(suppliers, totalSpend) {
    destroyChart('supplierPareto');
    let cumulative = 0;
    const cumPct = suppliers.map(([, d]) => { cumulative += d.spend; return (cumulative / totalSpend * 100); });

    const ctx = document.getElementById('chart-supplier-pareto');
    if (!ctx) return;
    charts.supplierPareto = new Chart(ctx, {
      type: 'line',
      data: {
        labels: suppliers.map(([n], i) => i + 1),
        datasets: [{
          label: 'Cumulative % of Spend',
          data: cumPct,
          borderColor: '#f59e0b',
          backgroundColor: 'rgba(245,158,11,0.1)',
          fill: true,
          tension: 0.3,
          pointRadius: suppliers.length > 30 ? 0 : 3,
          borderWidth: 2
        }]
      },
      options: {
        responsive: true, maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          tooltip: { callbacks: { label: ctx => ctx.raw.toFixed(1) + '% of total spend', title: items => 'Supplier #' + items[0].label + ': ' + suppliers[parseInt(items[0].label) - 1][0] } }
        },
        scales: {
          y: { min: 0, max: 100, grid: { color: 'rgba(255,255,255,0.05)' }, ticks: { color: '#94a3b8', callback: v => v + '%' } },
          x: { grid: { display: false }, ticks: { color: '#94a3b8' }, title: { display: true, text: 'Supplier Rank', color: '#94a3b8' } }
        }
      }
    });
  }

  function renderSupplierDetailTable(suppliers, totalSpend) {
    const searchInput = document.getElementById('supplier-search-input');
    const searchTerm = (searchInput ? searchInput.value : '').toLowerCase();

    let filtered = suppliers;
    if (searchTerm) {
      filtered = suppliers.filter(([name]) => name.toLowerCase().includes(searchTerm));
    }

    const rows = filtered.map(([name, d]) => [
      name, d.spend, d.count, d.skus.size, d.spend / d.count,
      [...d.cats].join(', '), pct(d.spend, totalSpend) + '%'
    ]);

    renderSortableTable('supplier-detail-table', rows,
      [{ fmt: 'text' }, { fmt: 'usd' }, { fmt: 'num' }, { fmt: 'num' }, { fmt: 'usd' }, { fmt: 'text' }, { fmt: 'text' }]);
  }

  // ---- Requester Analysis Page ----
  function renderRequesterPage() {
    const data = filteredData.filter(r => !r.budget_type || r.budget_type === 'Actual' || r.budget_type === '');
    const totalSpend = data.reduce((s, r) => s + r.total_amount_usd, 0);

    const reqMap = {};
    data.forEach(r => {
      const k = r.ordered_by || 'Unknown';
      if (!reqMap[k]) reqMap[k] = { dept: r.department, spend: 0, count: 0, catSpend: {} };
      reqMap[k].spend += r.total_amount_usd;
      reqMap[k].count++;
      if (!reqMap[k].catSpend[r.cost_category]) reqMap[k].catSpend[r.cost_category] = 0;
      reqMap[k].catSpend[r.cost_category] += r.total_amount_usd;
    });

    const requesters = Object.entries(reqMap).sort((a, b) => b[1].spend - a[1].spend);

    // Requester bar chart
    renderRequesterBarChart(requesters.slice(0, 15));
    // Department pie
    renderDepartmentPieChart(data);
    // Detail table
    const rows = requesters.map(([name, d]) => {
      const topCat = Object.entries(d.catSpend).sort((a, b) => b[1] - a[1])[0];
      return [name, d.dept || '', d.spend, d.count, d.spend / d.count, topCat ? topCat[0].split(',')[0] : '', pct(d.spend, totalSpend) + '%'];
    });
    renderSortableTable('requester-detail-table', rows,
      [{ fmt: 'text' }, { fmt: 'text' }, { fmt: 'usd' }, { fmt: 'num' }, { fmt: 'usd' }, { fmt: 'text' }, { fmt: 'text' }]);
  }

  function renderRequesterBarChart(top15) {
    destroyChart('requesterBar');
    const ctx = document.getElementById('chart-requester-bar');
    if (!ctx) return;
    charts.requesterBar = new Chart(ctx, {
      type: 'bar',
      data: {
        labels: top15.map(([n]) => n),
        datasets: [{ label: 'Spend (EUR)', data: top15.map(([, d]) => d.spend), backgroundColor: '#06b6d4', borderRadius: 4 }]
      },
      options: {
        indexAxis: 'y',
        responsive: true, maintainAspectRatio: false,
        plugins: { legend: { display: false }, tooltip: { callbacks: { label: ctx => fmtEUR(ctx.raw) } } },
        scales: {
          x: { beginAtZero: true, grid: { color: 'rgba(255,255,255,0.05)' }, ticks: { color: '#94a3b8', callback: v => curSym() + (convK(v)/1000).toFixed(0) + 'k' } },
          y: { grid: { display: false }, ticks: { color: '#94a3b8' } }
        }
      }
    });
  }

  function renderDepartmentPieChart(data) {
    destroyChart('departmentPie');
    const deptMap = {};
    data.forEach(r => {
      const d = r.department || 'Unassigned';
      if (!deptMap[d]) deptMap[d] = 0;
      deptMap[d] += r.total_amount_usd;
    });
    const entries = Object.entries(deptMap).sort((a, b) => b[1] - a[1]);
    const colors = ['#3b82f6','#f59e0b','#8b5cf6','#06b6d4','#10b981','#ec4899','#64748b','#ef4444','#a78bfa','#22d3ee'];

    const ctx = document.getElementById('chart-department-pie');
    if (!ctx) return;
    charts.departmentPie = new Chart(ctx, {
      type: 'doughnut',
      data: {
        labels: entries.map(e => e[0]),
        datasets: [{ data: entries.map(e => e[1]), backgroundColor: colors, borderWidth: 0, hoverOffset: 8 }]
      },
      options: {
        responsive: true, maintainAspectRatio: false, cutout: '55%',
        plugins: {
          legend: { position: 'right', labels: { color: '#e2e8f0', font: { size: 11 }, padding: 12, usePointStyle: true } },
          tooltip: { callbacks: { label: ctx => ctx.label + ': ' + fmtEUR(ctx.raw) } }
        }
      }
    });
  }

  // ---- Budget Targets ----
  function renderTargetSettings() {
    const body = document.getElementById('target-settings-body');
    if (!body) return;

    const years = Object.keys(categoryTargets).sort();

    if (years.length === 0) {
      body.innerHTML = '<p style="padding:20px 16px;color:var(--text-muted);font-size:13px;">No budget years defined. Click <strong style="color:var(--text)">+ Add Year</strong> to get started.</p>';
      body.onchange = null;
      body.onclick = null;
      return;
    }

    let html = '<div class="data-table-wrapper"><table class="target-table"><thead><tr>';
    html += '<th style="min-width:220px">Category</th>';
    years.forEach(yr => {
      html += `<th class="target-year-header">${yr} <button class="target-year-remove" data-year="${yr}" title="Remove year ${yr}">×</button></th>`;
    });
    html += '</tr></thead><tbody>';

    // Pre-compute actual spend per category per year from allData
    const actualByYearCat = {};
    allData.forEach(r => {
      if (r.budget_type && r.budget_type !== 'Actual' && r.budget_type !== '') return;
      const yr = r.date ? r.date.substring(0, 4) : null;
      if (!yr) return;
      if (!actualByYearCat[yr]) actualByYearCat[yr] = {};
      actualByYearCat[yr][r.cost_category] = (actualByYearCat[yr][r.cost_category] || 0) + r.total_amount_usd;
    });

    CATEGORIES.forEach(cat => {
      html += `<tr><td>${cat}</td>`;
      years.forEach(yr => {
        const stored = (categoryTargets[yr] && categoryTargets[yr][cat] != null) ? categoryTargets[yr][cat] : null;
        const displayVal = stored != null ? Math.round(stored / 1000) : '';
        const actual = (actualByYearCat[yr] && actualByYearCat[yr][cat]) || 0;
        let barHtml = '';
        if (stored != null && stored > 0) {
          const pctUsed = Math.min(100, actual / stored * 100);
          const barColor = pctUsed >= 100 ? '#ef4444' : pctUsed >= 85 ? '#f59e0b' : '#10b981';
          barHtml = `<div class="target-util-bar"><div class="target-util-fill" style="width:${pctUsed.toFixed(1)}%;background:${barColor}"></div></div>`;
        }
        html += `<td style="text-align:center"><input class="target-input" type="number" step="1" min="0" placeholder="—" data-cat="${cat}" data-year="${yr}" value="${displayVal}">${barHtml}</td>`;
      });
      html += '</tr>';
    });

    // Total row
    html += '<tr class="target-total-row"><td>TOTAL (' + curK() + ')</td>';
    years.forEach(yr => {
      let total = 0, hasAny = false;
      CATEGORIES.forEach(cat => {
        const v = categoryTargets[yr] && categoryTargets[yr][cat];
        if (v != null && v > 0) { total += v; hasAny = true; }
      });
      html += `<td id="target-total-${yr}" style="text-align:center;padding:6px 12px;">${hasAny ? fmtK(total) : '—'}</td>`;
    });
    html += '</tr></tbody></table></div>';

    body.innerHTML = html;

    body.onchange = function(e) {
      const input = e.target;
      if (!input.classList.contains('target-input')) return;
      const cat = input.dataset.cat;
      const year = input.dataset.year;
      const val = parseFloat(input.value);
      if (!categoryTargets[year]) categoryTargets[year] = {};
      if (input.value === '' || isNaN(val) || val < 0) {
        delete categoryTargets[year][cat];
      } else {
        categoryTargets[year][cat] = val * 1000;
      }
      saveTargets();
      // Update year total
      let total = 0, hasAny = false;
      CATEGORIES.forEach(c => {
        const v = categoryTargets[year] && categoryTargets[year][c];
        if (v != null && v > 0) { total += v; hasAny = true; }
      });
      const totalEl = document.getElementById('target-total-' + year);
      if (totalEl) totalEl.textContent = hasAny ? fmtK(total) : '—';
      dirtyPages.add('overview');
      dirtyPages.add('categories');
      dirtyPages.add('savings');
    };

    body.onclick = function(e) {
      const btn = e.target.closest('.target-year-remove');
      if (!btn) return;
      const year = btn.dataset.year;
      if (!confirm('Remove all budget targets for ' + year + '?')) return;
      delete categoryTargets[year];
      saveTargets();
      renderTargetSettings();
      dirtyPages.add('overview');
      dirtyPages.add('categories');
      dirtyPages.add('savings');
    };
  }

  function addTargetYear() {
    const yr = prompt('Enter year to add budget targets (e.g. 2027):');
    if (!yr) return;
    const trimmed = yr.trim();
    const num = parseInt(trimmed);
    if (isNaN(num) || trimmed.length !== 4 || num < 2000 || num > 2099) {
      toast('Please enter a valid 4-digit year between 2000 and 2099.', 'error');
      return;
    }
    const key = String(num);
    if (!categoryTargets[key]) categoryTargets[key] = {};
    saveTargets();
    renderTargetSettings();
  }

  // ---- Data Summary ----
  function updateDataSummary() {
    const el = document.getElementById('data-summary');
    if (!el) return;
    if (allData.length === 0) {
      el.innerHTML = 'No data loaded. Upload a CSV file or load sample data to get started.';
      return;
    }
    const dates = allData.map(r => r.date).filter(Boolean).sort();
    const cats = [...new Set(allData.map(r => r.cost_category).filter(Boolean))];
    const suppliers = [...new Set(allData.map(r => r.supplier).filter(Boolean))];
    const requesters = [...new Set(allData.map(r => r.ordered_by).filter(Boolean))];
    const totalSpend = allData.filter(r => !r.budget_type || r.budget_type === 'Actual').reduce((s, r) => s + r.total_amount_usd, 0);

    el.innerHTML =
      'Total records: <strong>' + allData.length + '</strong><br>' +
      'Date range: <strong>' + (dates[0] || '--') + '</strong> to <strong>' + (dates[dates.length - 1] || '--') + '</strong><br>' +
      'Categories: <strong>' + cats.length + '</strong><br>' +
      'Suppliers: <strong>' + suppliers.length + '</strong><br>' +
      'Requesters: <strong>' + requesters.length + '</strong><br>' +
      'Total actual spend: <strong>' + fmtEUR(totalSpend) + '</strong>';
  }

  // ---- Sample Data ----
  function generateSampleData() {
    const categories = CATEGORIES;
    const subCategories = {
      'Clinical, Lab and scientific services': ['Analytical testing', 'Bioassay services', 'Stability testing', 'Method validation', 'Microbiology testing'],
      'Production Equipment': ['Bioreactor parts', 'Filtration systems', 'Chromatography columns', 'Sensors & probes', 'Tubing & connectors'],
      'External Warehouse and distribution': ['Cold storage', 'Ambient storage', 'Distribution services', 'Packaging materials', 'Temperature monitoring'],
      'Professional Services': ['Consulting', 'Regulatory affairs', 'Quality auditing', 'Training services', 'IT services'],
      'Miscellaneous Indirect Costs': ['Travel', 'Subscriptions', 'Memberships', 'General supplies'],
      'Office and Print': ['Office supplies', 'Printing services', 'IT equipment', 'Furniture']
    };
    const suppliers = {
      'Clinical, Lab and scientific services': ['Biorelliance', 'Eurofins', 'SGS', 'Charles River Labs', 'WuXi AppTec', 'Lek Pharmaceuticals'],
      'Production Equipment': ['Sartorius', 'Pall Corporation', 'GE Healthcare', 'Merck Millipore', 'Cytiva', 'Thermo Fisher'],
      'External Warehouse and distribution': ['DHL Life Sciences', 'World Courier', 'Marken', 'FedEx Custom Critical', 'Kuehne+Nagel'],
      'Professional Services': ['Deloitte', 'KPMG', 'PwC', 'Accenture', 'McKinsey'],
      'Miscellaneous Indirect Costs': ['Various', 'Amazon Business', 'Local vendors', 'NIL d.o.o.'],
      'Office and Print': ['Mladinska Knjiga', 'Office Depot', 'Dell Technologies', 'HP Inc']
    };
    const requesters = [
      { name: 'Jan Novak', dept: 'QC Laboratory', cc: 'CC-4200' },
      { name: 'Maria Schmidt', dept: 'Production', cc: 'CC-3100' },
      { name: 'Peter Horvat', dept: 'Quality Assurance', cc: 'CC-4100' },
      { name: 'Ana Krajnc', dept: 'R&D', cc: 'CC-5000' },
      { name: 'Thomas Mueller', dept: 'Engineering', cc: 'CC-3200' },
      { name: 'Elena Popovic', dept: 'Supply Chain', cc: 'CC-2100' },
      { name: 'Marco Rossi', dept: 'Facilities', cc: 'CC-6000' },
      { name: 'Sophie Weber', dept: 'Procurement', cc: 'CC-2000' },
      { name: 'Luka Matic', dept: 'IT', cc: 'CC-7000' },
      { name: 'Katja Zupan', dept: 'Regulatory', cc: 'CC-4300' }
    ];

    const data = [];
    let poCounter = 1;
    const months = ['2025-01','2025-02','2025-03','2025-04','2025-05','2025-06','2025-07','2025-08','2025-09','2025-10','2025-11','2025-12',
                     '2026-01','2026-02'];

    // Weight categories by typical spend
    const catWeights = { 'Clinical, Lab and scientific services': 35, 'Production Equipment': 25, 'External Warehouse and distribution': 15,
      'Professional Services': 5, 'Miscellaneous Indirect Costs': 3, 'Office and Print': 2 };

    months.forEach(month => {
      categories.forEach(cat => {
        const weight = catWeights[cat] || 3;
        const numItems = Math.max(1, Math.floor(weight / 3 + Math.random() * weight / 2));
        const subs = subCategories[cat] || ['General'];
        const sups = suppliers[cat] || ['Various'];

        for (let i = 0; i < numItems; i++) {
          const sub = subs[Math.floor(Math.random() * subs.length)];
          const sup = sups[Math.floor(Math.random() * sups.length)];
          const req = requesters[Math.floor(Math.random() * requesters.length)];
          const qty = Math.floor(1 + Math.random() * 50);
          const basePrice = (weight * 100 + Math.random() * weight * 500);
          const unitPrice = Math.round(basePrice * 100) / 100;
          const totalAmount = Math.round(qty * unitPrice * 100) / 100;

          const skuPrefix = cat.substring(0, 3).toUpperCase().replace(/[^A-Z]/g, '');
          const skuNum = String(Math.floor(1000 + Math.random() * 9000));
          const sku = skuPrefix + '-' + skuNum;

          const priceImpact = Math.random() < 0.3 ? Math.round(-totalAmount * (0.02 + Math.random() * 0.05) * 100) / 100 : 0;
          const volumeImpact = Math.random() < 0.2 ? Math.round(-totalAmount * (0.01 + Math.random() * 0.03) * 100) / 100 : 0;
          const insourcingSavings = (cat === 'Clinical, Lab and scientific services' && Math.random() < 0.15) ?
            Math.round(-totalAmount * (0.05 + Math.random() * 0.1) * 100) / 100 : 0;

          data.push({
            date: month,
            cost_category: cat,
            sub_category: sub,
            sku: sku,
            item_description: sub + ' - ' + sup + ' service/product',
            supplier: sup,
            ordered_by: req.name,
            department: req.dept,
            cost_center: req.cc,
            po_number: 'PO-' + month.replace('-', '') + '-' + String(poCounter++).padStart(4, '0'),
            quantity: qty,
            unit_price_usd: unitPrice,
            total_amount_usd: totalAmount,
            budget_type: 'Actual',
            price_impact_usd: priceImpact,
            volume_impact_usd: volumeImpact,
            insourcing_savings_usd: insourcingSavings,
            notes: insourcingSavings < 0 ? 'Internalized from external provider' : ''
          });
        }
      });
    });

    return data;
  }

  // ---- Navigation ----
  function navigateTo(pageId) {
    document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));
    document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
    const navItem = document.querySelector('.nav-item[data-page="' + pageId + '"]');
    const page = document.getElementById('page-' + pageId);
    if (navItem) navItem.classList.add('active');
    if (page) page.classList.add('active');

    const titles = {
      overview: 'Overview Dashboard', datatable: 'Spend Data Table', categories: 'Category Analysis',
      savings: 'Savings Tracker', suppliers: 'Supplier Analysis', requesters: 'Requester Analysis',
      datamanage: 'Data Management', 'ai-advisor': 'AI Savings Advisor'
    };
    document.getElementById('page-title').textContent = titles[pageId] || pageId;

    activePage = pageId;
    if (dirtyPages.has(pageId)) {
      renderPage(pageId);
    }
  }

  // ---- Tab System ----
  function setupTabs() {
    document.querySelectorAll('.tabs').forEach(tabBar => {
      tabBar.querySelectorAll('.tab').forEach(tab => {
        tab.addEventListener('click', () => {
          tabBar.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
          tab.classList.add('active');
          const tabId = tab.dataset.tab;
          const parent = tabBar.parentElement;
          parent.querySelectorAll('.tab-content').forEach(tc => tc.classList.remove('active'));
          const content = parent.querySelector('#tab-' + tabId);
          if (content) content.classList.add('active');
        });
      });
    });
  }

  // ---- Event Listeners ----
  function setupEventListeners() {
    // Navigation
    document.querySelectorAll('.nav-item').forEach(item => {
      item.addEventListener('click', () => navigateTo(item.dataset.page));
    });

    // Global filters
    ['filter-year', 'filter-month', 'filter-category'].forEach(id => {
      document.getElementById(id).addEventListener('change', () => {
        applyGlobalFilters();
        refreshAll();
      });
    });

    // Currency toggle
    const currencyToggle = document.getElementById('currency-toggle');
    if (currencyToggle) {
      currencyToggle.addEventListener('click', (e) => {
        const label = e.target.closest('.toggle-label');
        if (!label) return;
        const cur = label.dataset.cur;
        if (cur === displayCurrency) return;
        displayCurrency = cur;
        currencyToggle.querySelectorAll('.toggle-label').forEach(l => l.classList.toggle('active', l.dataset.cur === cur));
        // Force re-render all pages
        dirtyPages = new Set(['overview','datatable','categories','savings','suppliers','requesters','datamanage','ai-advisor']);
        renderPage(activePage);
      });
    }

    // Table search
    document.getElementById('table-search-input').addEventListener('input', debounce(() => { currentPage = 1; renderTableBody(); }, 200));

    // Rows per page
    document.getElementById('rows-per-page').addEventListener('change', (e) => { rowsPerPage = parseInt(e.target.value); currentPage = 1; renderTableBody(); });

    // Toggle column filters
    document.getElementById('btn-toggle-filters').addEventListener('click', () => { showColumnFilters = !showColumnFilters; renderTableFilters(); });

    // Add row
    document.getElementById('btn-add-row').addEventListener('click', () => {
      const form = document.getElementById('add-row-form');
      form.style.display = form.style.display === 'none' ? 'grid' : 'none';
    });
    document.getElementById('btn-save-row').addEventListener('click', addNewRow);
    document.getElementById('btn-cancel-row').addEventListener('click', () => { document.getElementById('add-row-form').style.display = 'none'; });

    // Auto-calc total
    ['add-quantity', 'add-unitprice'].forEach(id => {
      document.getElementById(id).addEventListener('input', () => {
        const q = parseNum(document.getElementById('add-quantity').value);
        const p = parseNum(document.getElementById('add-unitprice').value);
        if (q && p) document.getElementById('add-totalamount').value = (q * p).toFixed(2);
      });
    });

    // Export from data table
    document.getElementById('btn-export-csv').addEventListener('click', () => {
      exportCSV(getTableData(), 'indirect_spend_export.csv');
      toast('Data exported successfully', 'success');
    });

    // File upload
    const uploadArea = document.getElementById('upload-area');
    const fileInput = document.getElementById('file-input');

    uploadArea.addEventListener('click', () => fileInput.click());
    uploadArea.addEventListener('dragover', (e) => { e.preventDefault(); uploadArea.classList.add('dragover'); });
    uploadArea.addEventListener('dragleave', () => uploadArea.classList.remove('dragover'));
    function handleFileText(text) {
      const firstLine = text.split('\n')[0];
      const headers = firstLine.split(/[,;\t]/).map(h => h.trim().replace(/"/g, ''));
      const sapKeys = Object.keys(SAP_FIELD_MAP);
      if (headers.filter(h => sapKeys.includes(h)).length >= 2) {
        startSAPWizard(text);
      } else {
        processCSV(text, false);
      }
    }

    uploadArea.addEventListener('drop', (e) => {
      e.preventDefault();
      uploadArea.classList.remove('dragover');
      const file = e.dataTransfer.files[0];
      if (file && file.name.endsWith('.csv')) {
        const reader = new FileReader();
        reader.onload = (ev) => handleFileText(ev.target.result);
        reader.readAsText(file);
      } else {
        toast('Please drop a CSV file', 'error');
      }
    });

    fileInput.addEventListener('change', (e) => {
      const file = e.target.files[0];
      if (file) {
        const reader = new FileReader();
        reader.onload = (ev) => handleFileText(ev.target.result);
        reader.readAsText(file);
        fileInput.value = '';
      }
    });

    // Mobile sidebar toggle
    const hamburger = document.getElementById('btn-hamburger');
    const sidebar = document.querySelector('.sidebar');
    const sidebarOverlay = document.getElementById('sidebar-overlay');
    function closeSidebar() { sidebar.classList.remove('open'); sidebarOverlay.classList.remove('active'); }
    hamburger.addEventListener('click', () => { sidebar.classList.toggle('open'); sidebarOverlay.classList.toggle('active'); });
    sidebarOverlay.addEventListener('click', closeSidebar);
    document.querySelectorAll('.nav-item').forEach(item => item.addEventListener('click', closeSidebar, { capture: true }));

    // Budget targets
    document.getElementById('btn-add-target-year').addEventListener('click', addTargetYear);

    // Data management buttons
    document.getElementById('btn-download-template').addEventListener('click', downloadTemplate);

    document.getElementById('btn-load-sample').addEventListener('click', () => {
      allData = generateSampleData();
      reindexData();
      saveToStorage();
      updateFilterOptions();
      applyGlobalFilters();
      refreshAll();
      updateFooter();
      toast('Sample data loaded: ' + allData.length + ' records', 'success');
    });

    document.getElementById('btn-export-all').addEventListener('click', () => {
      if (allData.length === 0) { toast('No data to export', 'warning'); return; }
      exportCSV(allData, 'indirect_spend_all.csv');
      toast('All data exported', 'success');
    });

    document.getElementById('btn-export-filtered').addEventListener('click', () => {
      if (filteredData.length === 0) { toast('No filtered data to export', 'warning'); return; }
      exportCSV(filteredData, 'indirect_spend_filtered.csv');
      toast('Filtered data exported', 'success');
    });

    document.getElementById('btn-clear-data').addEventListener('click', () => {
      showConfirm('Clear All Data', 'This will permanently delete all loaded data. Are you sure?', () => {
        allData = [];
        filteredData = [];
        localStorage.removeItem(STORAGE_KEY);
        updateFilterOptions();
        refreshAll();
        updateFooter();
        toast('All data cleared', 'success');
      });
    });

    // Confirm modal
    document.getElementById('confirm-ok').addEventListener('click', () => { if (confirmCallback) confirmCallback(); hideConfirm(); });
    document.getElementById('confirm-cancel').addEventListener('click', hideConfirm);
    document.getElementById('confirm-close').addEventListener('click', hideConfirm);

    // Edit row modal
    document.getElementById('edit-row-save').addEventListener('click', saveEditRow);
    document.getElementById('edit-row-delete').addEventListener('click', deleteEditRow);
    document.getElementById('edit-row-cancel').addEventListener('click', closeEditModal);
    document.getElementById('edit-row-close').addEventListener('click', closeEditModal);

    // Keyboard shortcuts: Esc closes edit modal, Enter saves it
    document.addEventListener('keydown', (e) => {
      const modal = document.getElementById('modal-edit-row');
      if (!modal.classList.contains('active')) return;
      if (e.key === 'Escape') { e.preventDefault(); closeEditModal(); }
      if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) { e.preventDefault(); saveEditRow(); }
    });

    // Close modals on overlay click
    document.querySelectorAll('.modal-overlay').forEach(overlay => {
      overlay.addEventListener('click', (e) => {
        if (e.target === overlay) {
          overlay.classList.remove('active');
          editingRowIndex = -1;
        }
      });
    });

    // Print report
    document.getElementById('btn-print-report').addEventListener('click', generateReport);

    // AI Advisor
    document.getElementById('ai-api-key').addEventListener('change', (e) => {
      aiAdvisorKey = e.target.value.trim();
      localStorage.setItem(AI_KEY_STORAGE, aiAdvisorKey);
    });
    document.getElementById('btn-ai-rescan').addEventListener('click', () => {
      dirtyPages.add('ai-advisor');
      renderPage('ai-advisor');
    });
    document.getElementById('btn-ai-generate').addEventListener('click', () => {
      const key = document.getElementById('ai-api-key').value.trim() || aiAdvisorKey;
      if (!key) { toast('Enter your Anthropic API key first', 'error'); return; }
      const model = document.getElementById('ai-model').value;
      const findings = analyzeSpendOpportunities(filteredData);
      callClaudeAdvisor(findings, key, model);
    });
    document.getElementById('btn-ai-copy').addEventListener('click', () => {
      const text = document.getElementById('ai-response-content').textContent;
      navigator.clipboard.writeText(text).then(() => toast('Copied to clipboard', 'success'));
    });

    // Supplier search — only re-filters the table, no chart rebuild
    const supSearch = document.getElementById('supplier-search-input');
    if (supSearch) {
      supSearch.addEventListener('input', () => {
        if (cachedSuppliers) {
          renderSupplierDetailTable(cachedSuppliers, cachedSupplierTotalSpend);
        } else {
          renderSupplierPage();
        }
      });
    }
  }


  // ============================================================
  // PRINT REPORT
  // ============================================================

  function generateReport() {
    if (allData.length === 0) { toast('No data loaded — import data first', 'warning'); return; }

    const actual = filteredData.filter(r => !r.budget_type || r.budget_type === 'Actual' || r.budget_type === '');
    const totalSpend    = actual.reduce((s, r) => s + r.total_amount_usd, 0);
    const totalPrice    = filteredData.reduce((s, r) => s + (r.price_impact_usd || 0), 0);
    const totalVolume   = filteredData.reduce((s, r) => s + (r.volume_impact_usd || 0), 0);
    const totalInsource = filteredData.reduce((s, r) => s + (r.insourcing_savings_usd || 0), 0);
    const totalSavings  = totalPrice + totalVolume + totalInsource;
    const uniqueSuppliers  = new Set(actual.map(r => r.supplier).filter(Boolean)).size;
    const uniqueRequesters = new Set(actual.map(r => r.ordered_by).filter(Boolean)).size;
    const dates = filteredData.map(r => r.date).filter(Boolean).sort();
    const dateRange = dates.length ? dates[0] + ' — ' + dates[dates.length - 1] : 'All dates';

    // Category breakdown
    const catMap = {};
    CATEGORIES.forEach(c => { catMap[c] = { actual: 0, baseline: 0, price: 0, volume: 0, insourcing: 0 }; });
    filteredData.forEach(r => {
      const cat = r.cost_category;
      if (!catMap[cat]) catMap[cat] = { actual: 0, baseline: 0, price: 0, volume: 0, insourcing: 0 };
      const bt = (r.budget_type || 'Actual').toLowerCase();
      if (bt === 'actual' || bt === '') catMap[cat].actual += r.total_amount_usd;
      else if (bt === 'baseline') catMap[cat].baseline += r.total_amount_usd;
      catMap[cat].price     += r.price_impact_usd || 0;
      catMap[cat].volume    += r.volume_impact_usd || 0;
      catMap[cat].insourcing += r.insourcing_savings_usd || 0;
    });
    const catRows = CATEGORIES.filter(c => catMap[c] && (catMap[c].actual || catMap[c].baseline))
      .sort((a, b) => catMap[b].actual - catMap[a].actual);

    // Top 10 suppliers
    const supMap = {};
    actual.forEach(r => {
      const s = r.supplier || 'Unknown';
      if (!supMap[s]) supMap[s] = { spend: 0, orders: 0 };
      supMap[s].spend += r.total_amount_usd;
      supMap[s].orders++;
    });
    const top10Sup = Object.entries(supMap).sort((a, b) => b[1].spend - a[1].spend).slice(0, 10);

    // Monthly spend for sparkline-like text
    const monthMap = {};
    actual.forEach(r => {
      const mk = getMonthKey(r.date);
      if (!mk) return;
      if (!monthMap[mk]) monthMap[mk] = 0;
      monthMap[mk] += r.total_amount_usd;
    });
    const months = Object.keys(monthMap).sort();

    // Filter description
    const yearEl  = document.getElementById('filter-year');
    const monthEl = document.getElementById('filter-month');
    const catEl   = document.getElementById('filter-category');
    const filterDesc = [
      yearEl.value  !== 'all' ? 'Year: ' + yearEl.value : '',
      monthEl.value !== 'all' ? 'Month: ' + monthEl.options[monthEl.selectedIndex].text : '',
      catEl.value   !== 'all' ? 'Category: ' + catEl.value : ''
    ].filter(Boolean).join(' | ') || 'All data (no filters applied)';

    const r = (n) => curSym() + Number(displayCurrency === 'USD' ? n * EUR_USD_RATE : n).toLocaleString('en-US', { minimumFractionDigits: 0, maximumFractionDigits: 0 });
    const rk = (n) => curSym() + ((displayCurrency === 'USD' ? n * EUR_USD_RATE : n) / 1000).toLocaleString('en-US', { minimumFractionDigits: 1, maximumFractionDigits: 1 }) + 'k';
    const sign = (n) => n < 0 ? '<span style="color:#16a34a">' + r(n) + '</span>' : n > 0 ? '<span style="color:#dc2626">' + r(n) + '</span>' : '—';

    const html = `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>Indirect Spend Report — ${new Date().toLocaleDateString('en-US', { year:'numeric', month:'long', day:'numeric' })}</title>
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: 'Segoe UI', Arial, sans-serif; font-size: 13px; color: #1e293b; background: #fff; line-height: 1.5; }
  .report-wrap { max-width: 900px; margin: 0 auto; padding: 32px 24px; }

  /* Header */
  .report-header { display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 28px; padding-bottom: 16px; border-bottom: 3px solid #1e3a5f; }
  .report-logo h1 { font-size: 22px; font-weight: 700; color: #1e3a5f; }
  .report-logo .sub { font-size: 12px; color: #64748b; margin-top: 2px; }
  .report-meta { text-align: right; font-size: 12px; color: #64748b; }
  .report-meta strong { display: block; font-size: 13px; color: #1e293b; }

  /* Section */
  .section { margin-bottom: 28px; }
  .section-title { font-size: 14px; font-weight: 700; color: #1e3a5f; text-transform: uppercase; letter-spacing: 0.06em; border-bottom: 1px solid #e2e8f0; padding-bottom: 6px; margin-bottom: 14px; }

  /* KPI grid */
  .kpi-row { display: grid; grid-template-columns: repeat(4, 1fr); gap: 12px; margin-bottom: 4px; }
  .kpi-box { border: 1px solid #e2e8f0; border-radius: 8px; padding: 14px 16px; }
  .kpi-box .label { font-size: 11px; color: #64748b; text-transform: uppercase; font-weight: 600; letter-spacing: 0.05em; }
  .kpi-box .value { font-size: 22px; font-weight: 700; color: #0f172a; margin: 4px 0 2px; }
  .kpi-box .sub   { font-size: 11px; color: #94a3b8; }
  .kpi-box.green  { border-left: 4px solid #16a34a; }
  .kpi-box.blue   { border-left: 4px solid #2563eb; }
  .kpi-box.orange { border-left: 4px solid #d97706; }
  .kpi-box.slate  { border-left: 4px solid #475569; }

  /* Tables */
  table { width: 100%; border-collapse: collapse; font-size: 12px; }
  th { background: #f8fafc; font-weight: 600; color: #475569; text-transform: uppercase; font-size: 10px; letter-spacing: 0.05em; padding: 8px 10px; text-align: left; border-bottom: 2px solid #e2e8f0; }
  th.num, td.num { text-align: right; }
  td { padding: 7px 10px; border-bottom: 1px solid #f1f5f9; color: #334155; }
  tr:last-child td { border-bottom: none; }
  tr:hover td { background: #f8fafc; }
  tfoot td { font-weight: 700; background: #f1f5f9; border-top: 2px solid #e2e8f0; }

  /* Savings grid */
  .savings-grid { display: grid; grid-template-columns: repeat(3, 1fr); gap: 12px; }
  .savings-box { border: 1px solid #e2e8f0; border-radius: 8px; padding: 14px; }
  .savings-box .label { font-size: 11px; color: #64748b; font-weight: 600; text-transform: uppercase; }
  .savings-box .value { font-size: 19px; font-weight: 700; margin-top: 4px; }
  .savings-box .detail { font-size: 11px; color: #94a3b8; margin-top: 3px; }
  .savings { color: #16a34a; } .increase { color: #dc2626; } .neutral { color: #475569; }

  /* Monthly bar row */
  .month-bars { display: flex; align-items: flex-end; gap: 4px; height: 60px; margin-top: 12px; }
  .month-bar-wrap { flex: 1; display: flex; flex-direction: column; align-items: center; gap: 3px; height: 100%; justify-content: flex-end; }
  .month-bar { width: 100%; background: #bfdbfe; border-radius: 2px 2px 0 0; min-height: 2px; }
  .month-label { font-size: 8px; color: #94a3b8; white-space: nowrap; }

  /* Footer */
  .report-footer { margin-top: 32px; padding-top: 12px; border-top: 1px solid #e2e8f0; font-size: 11px; color: #94a3b8; display: flex; justify-content: space-between; }

  /* Print button */
  .print-btn { display: flex; gap: 8px; margin-bottom: 24px; }
  .print-btn button { padding: 8px 20px; border: none; border-radius: 6px; cursor: pointer; font-size: 13px; font-weight: 600; }
  .print-btn .primary { background: #1e3a5f; color: #fff; }
  .print-btn .secondary { background: #f1f5f9; color: #475569; }
  @media print {
    .print-btn { display: none; }
    body { font-size: 11px; }
    .report-wrap { padding: 16px; }
    .kpi-box .value { font-size: 18px; }
    table { page-break-inside: avoid; }
    .section { page-break-inside: avoid; }
  }
</style>
</head>
<body>
<div class="report-wrap">

  <div class="print-btn">
    <button class="primary" onclick="window.print()">&#x1F5A8; Print / Save as PDF</button>
    <button class="secondary" onclick="window.close()">Close</button>
  </div>

  <div class="report-header">
    <div class="report-logo">
      <h1>Indirect Spend Report</h1>
      <div class="sub">BioPharma Indirect Spend Tracker</div>
    </div>
    <div class="report-meta">
      <strong>${new Date().toLocaleDateString('en-US', { year:'numeric', month:'long', day:'numeric' })}</strong>
      Period: ${dateRange}<br>
      Filter: ${filterDesc}<br>
      Records: ${filteredData.length.toLocaleString()}
    </div>
  </div>

  <!-- Executive Summary -->
  <div class="section">
    <div class="section-title">Executive Summary</div>
    <div class="kpi-row">
      <div class="kpi-box blue">
        <div class="label">Total Spend</div>
        <div class="value">${r(totalSpend)}</div>
        <div class="sub">${actual.length} line items (Actual)</div>
      </div>
      <div class="kpi-box ${totalSavings <= 0 ? 'green' : 'orange'}">
        <div class="label">Total Savings</div>
        <div class="value ${totalSavings <= 0 ? 'savings' : 'increase'}">${r(Math.abs(totalSavings))}</div>
        <div class="sub">${totalSpend > 0 ? ((Math.abs(totalSavings) / totalSpend) * 100).toFixed(1) : '0'}% of spend</div>
      </div>
      <div class="kpi-box slate">
        <div class="label">Suppliers</div>
        <div class="value">${uniqueSuppliers}</div>
        <div class="sub">Active vendors</div>
      </div>
      <div class="kpi-box slate">
        <div class="label">Requesters</div>
        <div class="value">${uniqueRequesters}</div>
        <div class="sub">Ordering users</div>
      </div>
    </div>
  </div>

  <!-- Monthly Spend Trend -->
  ${months.length > 0 ? (() => {
    const maxV = Math.max(...months.map(m => monthMap[m]));
    return '<div class="section"><div class="section-title">Monthly Spend Trend</div>' +
      '<div class="month-bars">' +
      months.map(m => {
        const h = Math.max(4, Math.round((monthMap[m] / maxV) * 56));
        return '<div class="month-bar-wrap"><div class="month-bar" style="height:' + h + 'px" title="' + monthLabel(m) + ': ' + r(monthMap[m]) + '"></div><div class="month-label">' + monthLabel(m).replace(' ', '\'') + '</div></div>';
      }).join('') +
      '</div></div>';
  })() : ''}

  <!-- Category Breakdown -->
  <div class="section">
    <div class="section-title">Spend by Category</div>
    <table>
      <thead><tr>
        <th>Category</th>
        <th class="num">Actual Spend</th>
        <th class="num">Share %</th>
        <th class="num">Baseline</th>
        <th class="num">Price Impact</th>
        <th class="num">Volume Impact</th>
        <th class="num">Insourcing</th>
        <th class="num">Target</th>
      </tr></thead>
      <tbody>
        ${catRows.map(cat => {
          const d = catMap[cat];
          const target = d.actual + d.price + d.volume + d.insourcing;
          return '<tr><td>' + cat + '</td>' +
            '<td class="num">' + r(d.actual) + '</td>' +
            '<td class="num">' + (totalSpend > 0 ? ((d.actual / totalSpend) * 100).toFixed(1) : '0') + '%</td>' +
            '<td class="num">' + (d.baseline ? r(d.baseline) : r(d.actual)) + '</td>' +
            '<td class="num">' + sign(d.price) + '</td>' +
            '<td class="num">' + sign(d.volume) + '</td>' +
            '<td class="num">' + sign(d.insourcing) + '</td>' +
            '<td class="num"><strong>' + r(target) + '</strong></td>' +
            '</tr>';
        }).join('')}
      </tbody>
      <tfoot><tr>
        <td>TOTAL</td>
        <td class="num">${r(totalSpend)}</td>
        <td class="num">100%</td>
        <td class="num">${r(catRows.reduce((s, c) => s + (catMap[c].baseline || catMap[c].actual), 0))}</td>
        <td class="num">${sign(totalPrice)}</td>
        <td class="num">${sign(totalVolume)}</td>
        <td class="num">${sign(totalInsource)}</td>
        <td class="num"><strong>${r(totalSpend + totalSavings)}</strong></td>
      </tr></tfoot>
    </table>
  </div>

  <!-- Savings Summary -->
  <div class="section">
    <div class="section-title">Savings Summary</div>
    <div class="savings-grid">
      <div class="savings-box">
        <div class="label">Price Impact</div>
        <div class="value ${totalPrice <= 0 ? 'savings' : 'increase'}">${r(totalPrice)}</div>
        <div class="detail">${totalPrice <= 0 ? 'Procurement savings vs baseline' : 'Cost increase vs baseline'}</div>
      </div>
      <div class="savings-box">
        <div class="label">Volume Impact</div>
        <div class="value ${totalVolume <= 0 ? 'savings' : 'increase'}">${r(totalVolume)}</div>
        <div class="detail">${totalVolume <= 0 ? 'Demand reduction savings' : 'Volume increase impact'}</div>
      </div>
      <div class="savings-box">
        <div class="label">Insourcing Savings</div>
        <div class="value ${totalInsource <= 0 ? 'savings' : 'increase'}">${r(totalInsource)}</div>
        <div class="detail">${totalInsource <= 0 ? 'From internalization initiatives' : 'External cost increase'}</div>
      </div>
    </div>
  </div>

  <!-- Top Suppliers -->
  <div class="section">
    <div class="section-title">Top 10 Suppliers by Spend</div>
    <table>
      <thead><tr>
        <th>#</th><th>Supplier</th>
        <th class="num">Total Spend</th>
        <th class="num">Orders</th>
        <th class="num">Avg Order</th>
        <th class="num">Share %</th>
      </tr></thead>
      <tbody>
        ${top10Sup.map(([name, d], i) =>
          '<tr><td style="color:#94a3b8;font-weight:600">' + (i + 1) + '</td>' +
          '<td><strong>' + name + '</strong></td>' +
          '<td class="num">' + r(d.spend) + '</td>' +
          '<td class="num">' + d.orders + '</td>' +
          '<td class="num">' + r(d.spend / d.orders) + '</td>' +
          '<td class="num">' + (totalSpend > 0 ? ((d.spend / totalSpend) * 100).toFixed(1) : '0') + '%</td></tr>'
        ).join('')}
      </tbody>
      <tfoot><tr>
        <td colspan="2">Top 10 Total</td>
        <td class="num">${r(top10Sup.reduce((s, [, d]) => s + d.spend, 0))}</td>
        <td class="num">${top10Sup.reduce((s, [, d]) => s + d.orders, 0)}</td>
        <td class="num">—</td>
        <td class="num">${totalSpend > 0 ? ((top10Sup.reduce((s, [, d]) => s + d.spend, 0) / totalSpend) * 100).toFixed(1) : '0'}%</td>
      </tr></tfoot>
    </table>
  </div>

  <div class="report-footer">
    <span>BioPharma Indirect Spend Tracker</span>
    <span>Generated ${new Date().toLocaleString('en-US')} — Filtered: ${filterDesc}</span>
  </div>
</div>
</body>
</html>`;

    const win = window.open('', '_blank', 'width=980,height=800');
    if (!win) { toast('Pop-up blocked — allow pop-ups for this page', 'error'); return; }
    win.document.write(html);
    win.document.close();
    setTimeout(() => win.print(), 600);
  }

  // ============================================================
  // AI ADVISOR
  // ============================================================

  function analyzeSpendOpportunities(data) {
    const actual = data.filter(r => !r.budget_type || r.budget_type === 'Actual' || r.budget_type === '');
    const findings = [];

    // --- 1. Supplier Consolidation ---
    CATEGORIES.forEach(cat => {
      const catRows = actual.filter(r => r.cost_category === cat);
      if (!catRows.length) return;
      const catSpend = catRows.reduce((s, r) => s + r.total_amount_usd, 0);
      const supMap = {};
      catRows.forEach(r => {
        const s = r.supplier || 'Unknown';
        if (!supMap[s]) supMap[s] = 0;
        supMap[s] += r.total_amount_usd;
      });
      const suppliers = Object.entries(supMap).sort((a, b) => b[1] - a[1]);
      if (suppliers.length < 3) return;
      const top3Share = suppliers.slice(0, 3).reduce((s, [, v]) => s + v, 0) / catSpend;
      if (top3Share < 0.65 && suppliers.length >= 5) {
        const estimatedSavings = Math.round(catSpend * 0.07);
        findings.push({
          type: 'supplier_consolidation',
          priority: catSpend > 50000 ? 'high' : 'medium',
          category: cat,
          title: 'Supplier Consolidation',
          detail: suppliers.length + ' suppliers in ' + cat.split(',')[0] + ' — top 3 cover only ' + Math.round(top3Share * 100) + '% of spend (' + curSym() + Math.round(convK(catSpend) / 1000) + 'k). Fragmented buying reduces negotiating leverage.',
          affected: suppliers.slice(0, 4).map(([n]) => n),
          estimatedSavings,
          action: 'Issue an RFQ to consolidate to 2–3 preferred suppliers with volume commitments.',
          icon: 'purple'
        });
      }
    });

    // --- 2. Price Variance (same SKU, big price range) ---
    const skuPrices = {};
    actual.forEach(r => {
      const k = r.sku;
      if (!k || !r.unit_price_usd) return;
      if (!skuPrices[k]) skuPrices[k] = { prices: [], qty: 0, spend: 0, desc: r.item_description, cat: r.cost_category };
      skuPrices[k].prices.push(r.unit_price_usd);
      skuPrices[k].qty += r.quantity;
      skuPrices[k].spend += r.total_amount_usd;
    });
    const varianceHits = [];
    Object.entries(skuPrices).forEach(([sku, d]) => {
      if (d.prices.length < 2) return;
      const minP = Math.min(...d.prices), maxP = Math.max(...d.prices);
      const variance = (maxP - minP) / minP;
      if (variance > 0.15 && d.spend > 2000) {
        const avgP = d.prices.reduce((a, b) => a + b, 0) / d.prices.length;
        const savingsEstimate = Math.round((avgP - minP) * d.qty * 0.5);
        if (savingsEstimate > 200) varianceHits.push({ sku, desc: d.desc, variance, savingsEstimate, cat: d.cat });
      }
    });
    varianceHits.sort((a, b) => b.savingsEstimate - a.savingsEstimate);
    if (varianceHits.length > 0) {
      const topHits = varianceHits.slice(0, 5);
      const totalSavings = topHits.reduce((s, h) => s + h.savingsEstimate, 0);
      findings.push({
        type: 'price_variance',
        priority: totalSavings > 10000 ? 'high' : 'medium',
        category: topHits[0].cat,
        title: 'Price Variance Detected',
        detail: topHits.length + ' SKU(s) purchased at significantly different prices across POs. Top offender: ' + (topHits[0].desc || topHits[0].sku).substring(0, 50) + ' — ' + Math.round(topHits[0].variance * 100) + '% price spread.',
        affected: topHits.map(h => h.sku),
        estimatedSavings: totalSavings,
        action: 'Standardize pricing via blanket POs or catalogue agreements. Enforce approved price list.',
        icon: 'orange'
      });
    }

    // --- 3. Tail Spend ---
    const allSupMap = {};
    actual.forEach(r => {
      const s = r.supplier || 'Unknown';
      if (!allSupMap[s]) allSupMap[s] = { spend: 0, count: 0 };
      allSupMap[s].spend += r.total_amount_usd;
      allSupMap[s].count++;
    });
    const tailSuppliers = Object.entries(allSupMap).filter(([, d]) => d.spend < 5000 && d.count <= 5);
    if (tailSuppliers.length >= 3) {
      const tailSpend = tailSuppliers.reduce((s, [, d]) => s + d.spend, 0);
      findings.push({
        type: 'tail_spend',
        priority: tailSuppliers.length > 10 ? 'medium' : 'low',
        category: 'All Categories',
        title: 'Tail Spend Cleanup',
        detail: tailSuppliers.length + ' suppliers each account for less than ' + curSym() + '5k in total spend (combined ' + curSym() + Math.round(convK(tailSpend) / 1000) + 'k). Tail spend increases admin cost and reduces leverage.',
        affected: tailSuppliers.slice(0, 5).map(([n]) => n),
        estimatedSavings: Math.round(tailSpend * 0.05),
        action: 'Consolidate tail suppliers into preferred vendors or a marketplace (e.g. Amazon Business). Target <20 active suppliers per category.',
        icon: 'blue'
      });
    }

    // --- 4. Volume Bundling (same SKU, multiple small POs in one month) ---
    const skuMonthOrders = {};
    actual.forEach(r => {
      const k = (r.sku || '') + '|' + (r.date || '').substring(0, 7);
      if (!skuMonthOrders[k]) skuMonthOrders[k] = { sku: r.sku, desc: r.item_description, count: 0, spend: 0 };
      skuMonthOrders[k].count++;
      skuMonthOrders[k].spend += r.total_amount_usd;
    });
    const bundleHits = Object.values(skuMonthOrders).filter(d => d.count >= 3 && d.spend > 1000);
    bundleHits.sort((a, b) => b.spend - a.spend);
    if (bundleHits.length > 0) {
      const totalBundleSpend = bundleHits.reduce((s, h) => s + h.spend, 0);
      findings.push({
        type: 'volume_bundling',
        priority: totalBundleSpend > 30000 ? 'medium' : 'low',
        category: 'Multiple',
        title: 'Volume Bundling Opportunity',
        detail: bundleHits.length + ' SKU(s) are ordered 3+ times per month in separate POs. Top case: ' + (bundleHits[0].desc || bundleHits[0].sku || '').substring(0, 45) + ' (' + bundleHits[0].count + ' orders/month).',
        affected: bundleHits.slice(0, 4).map(h => h.sku).filter(Boolean),
        estimatedSavings: Math.round(totalBundleSpend * 0.03),
        action: 'Consolidate repeat orders into monthly blanket POs. Reduces processing cost and enables volume discounts.',
        icon: 'cyan'
      });
    }

    // --- 5. Untapped savings categories ---
    CATEGORIES.forEach(cat => {
      const catRows = data.filter(r => r.cost_category === cat);
      if (!catRows.length) return;
      const catActualSpend = catRows.filter(r => !r.budget_type || r.budget_type === 'Actual' || r.budget_type === '')
        .reduce((s, r) => s + r.total_amount_usd, 0);
      if (catActualSpend < 50000) return;
      const hasSavings = catRows.some(r =>
        (r.price_impact_usd && r.price_impact_usd !== 0) ||
        (r.volume_impact_usd && r.volume_impact_usd !== 0) ||
        (r.insourcing_savings_usd && r.insourcing_savings_usd !== 0)
      );
      if (!hasSavings) {
        findings.push({
          type: 'untapped_savings',
          priority: catActualSpend > 100000 ? 'high' : 'medium',
          category: cat,
          title: 'No Savings Initiatives',
          detail: cat.split(',')[0] + ' has ' + curSym() + Math.round(convK(catActualSpend) / 1000) + 'k in spend but zero recorded savings initiatives. Industry benchmark is 3–7% savings annually.',
          affected: [],
          estimatedSavings: Math.round(catActualSpend * 0.05),
          action: 'Launch a sourcing initiative: market benchmarking, RFQ, or demand management review.',
          icon: 'red'
        });
      }
    });

    // --- 6. Single-source risk ---
    CATEGORIES.forEach(cat => {
      const catRows = actual.filter(r => r.cost_category === cat);
      if (!catRows.length) return;
      const catSpend = catRows.reduce((s, r) => s + r.total_amount_usd, 0);
      if (catSpend < 20000) return;
      const uniqueSups = new Set(catRows.map(r => r.supplier).filter(Boolean));
      if (uniqueSups.size === 1) {
        const [supName] = uniqueSups;
        findings.push({
          type: 'single_source',
          priority: catSpend > 80000 ? 'high' : 'medium',
          category: cat,
          title: 'Single-Source Risk',
          detail: cat.split(',')[0] + ' is 100% sourced from ' + supName + ' (' + curSym() + Math.round(convK(catSpend) / 1000) + 'k). No competitive leverage or supply continuity fallback.',
          affected: [supName],
          estimatedSavings: Math.round(catSpend * 0.08),
          action: 'Qualify a second supplier and run a competitive RFQ. Even a 20% split creates leverage for pricing negotiations.',
          icon: 'green'
        });
      }
    });

    return findings.sort((a, b) => b.estimatedSavings - a.estimatedSavings);
  }

  function renderAIAdvisor() {
    const data = filteredData;
    if (data.length === 0) {
      document.getElementById('ai-kpi-bar').innerHTML = '';
      document.getElementById('ai-opportunities-list').innerHTML =
        '<div class="ai-empty-state">' +
        '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" width="48" height="48"><path d="M13 2L3 14h9l-1 8 10-12h-9l1-8z"/></svg>' +
        '<h3>No data loaded</h3><p>Import spend data to generate savings insights.</p></div>';
      return;
    }

    const findings = analyzeSpendOpportunities(data);
    const totalSavings = findings.reduce((s, f) => s + f.estimatedSavings, 0);
    const highCount = findings.filter(f => f.priority === 'high').length;
    const actualSpend = data.filter(r => !r.budget_type || r.budget_type === 'Actual' || r.budget_type === '')
      .reduce((s, r) => s + r.total_amount_usd, 0);
    const savingsPct = actualSpend > 0 ? ((totalSavings / actualSpend) * 100).toFixed(1) : '0';

    // KPI bar
    document.getElementById('ai-kpi-bar').innerHTML =
      '<div class="kpi-card"><div class="kpi-icon red"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="24" height="24"><path d="M13 2L3 14h9l-1 8 10-12h-9l1-8z"/></svg></div>' +
      '<div class="kpi-label">Opportunities Found</div><div class="kpi-value">' + findings.length + '</div>' +
      '<div class="kpi-change ' + (highCount > 0 ? 'negative' : 'neutral') + '">' + highCount + ' high priority</div></div>' +

      '<div class="kpi-card"><div class="kpi-icon green"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="24" height="24"><path d="M12 2v20M17 5H9.5a3.5 3.5 0 000 7h5a3.5 3.5 0 010 7H6"/></svg></div>' +
      '<div class="kpi-label">Potential Savings</div><div class="kpi-value">' + fmtEUR(totalSavings) + '</div>' +
      '<div class="kpi-change positive">' + savingsPct + '% of spend</div></div>' +

      '<div class="kpi-card"><div class="kpi-icon blue"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="24" height="24"><circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/></svg></div>' +
      '<div class="kpi-label">Quick Wins (0–3 months)</div>' +
      '<div class="kpi-value">' + findings.filter(f => f.type === 'price_variance' || f.type === 'volume_bundling' || f.type === 'tail_spend').length + '</div>' +
      '<div class="kpi-change neutral">Tactical opportunities</div></div>' +

      '<div class="kpi-card"><div class="kpi-icon orange"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="24" height="24"><path d="M20 21v-2a4 4 0 00-4-4H8a4 4 0 00-4 4v2"/><circle cx="12" cy="7" r="4"/></svg></div>' +
      '<div class="kpi-label">Suppliers Analyzed</div>' +
      '<div class="kpi-value">' + new Set(data.map(r => r.supplier).filter(Boolean)).size + '</div>' +
      '<div class="kpi-change neutral">' + new Set(data.map(r => r.cost_category).filter(Boolean)).size + ' categories</div></div>';

    // Opportunity cards
    const iconSVGs = {
      purple: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="18" height="18"><path d="M17 21v-2a4 4 0 00-4-4H5a4 4 0 00-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 00-3-3.87M16 3.13a4 4 0 010 7.75"/></svg>',
      orange: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="18" height="18"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>',
      blue:   '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="18" height="18"><path d="M21 16V8a2 2 0 00-1-1.73l-7-4a2 2 0 00-2 0l-7 4A2 2 0 003 8v8a2 2 0 001 1.73l7 4a2 2 0 002 0l7-4A2 2 0 0021 16z"/></svg>',
      cyan:   '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="18" height="18"><polyline points="16 3 21 3 21 8"/><line x1="4" y1="20" x2="21" y2="3"/><polyline points="21 16 21 21 16 21"/><line x1="15" y1="15" x2="21" y2="21"/></svg>',
      red:    '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="18" height="18"><path d="M13 2L3 14h9l-1 8 10-12h-9l1-8z"/></svg>',
      green:  '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="18" height="18"><path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/></svg>'
    };

    if (findings.length === 0) {
      document.getElementById('ai-opportunities-list').innerHTML =
        '<div class="ai-empty-state"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" width="48" height="48"><path d="M22 11.08V12a10 10 0 11-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/></svg>' +
        '<h3>No issues detected</h3><p>Your spend data looks well-managed. Keep monitoring as new data comes in.</p></div>';
    } else {
      document.getElementById('ai-opportunities-list').innerHTML = findings.map(f =>
        '<div class="ai-opp-card priority-' + f.priority + '">' +
        '<div class="ai-opp-icon ' + f.icon + '">' + (iconSVGs[f.icon] || '') + '</div>' +
        '<div class="ai-opp-body">' +
        '<div class="ai-opp-header"><span class="ai-opp-title">' + f.title + '</span>' +
        '<span class="priority-badge ' + f.priority + '">' + f.priority + '</span></div>' +
        '<div class="ai-opp-detail">' + f.detail + '</div>' +
        '<div class="ai-opp-action">' + f.action + '</div>' +
        (f.affected.length ? '<div class="ai-opp-affected">Affected: ' + f.affected.slice(0, 4).join(', ') + (f.affected.length > 4 ? '…' : '') + '</div>' : '') +
        '</div>' +
        '<div class="ai-opp-savings"><div class="amount">' + (f.estimatedSavings >= 1000 ? curSym() + Math.round(convK(f.estimatedSavings) / 1000) + 'k' : curSym() + Math.round(convK(f.estimatedSavings))) + '</div><div class="label">Est. savings</div></div>' +
        '</div>'
      ).join('');
    }

    // Restore saved API key into input
    const keyInput = document.getElementById('ai-api-key');
    if (aiAdvisorKey && !keyInput.value) keyInput.value = aiAdvisorKey;
  }

  async function callClaudeAdvisor(findings, apiKey, model) {
    const responseArea = document.getElementById('ai-response-area');
    const content = document.getElementById('ai-response-content');
    responseArea.style.display = '';
    content.innerHTML = '<span class="ai-loading">Generating action plan</span>';

    const actual = filteredData.filter(r => !r.budget_type || r.budget_type === 'Actual' || r.budget_type === '');
    const totalSpend = actual.reduce((s, r) => s + r.total_amount_usd, 0);
    const dates = filteredData.map(r => r.date).filter(Boolean).sort();
    const catBreakdown = {};
    actual.forEach(r => {
      const c = r.cost_category || 'Miscellaneous Indirect Costs';
      if (!catBreakdown[c]) catBreakdown[c] = 0;
      catBreakdown[c] += r.total_amount_usd;
    });
    const topCats = Object.entries(catBreakdown).sort((a, b) => b[1] - a[1]).slice(0, 5);
    const totalSavings = findings.reduce((s, f) => s + f.estimatedSavings, 0);

    const prompt = `You are a senior procurement savings expert analyzing indirect spend data for a pharmaceutical company.

SPEND SUMMARY:
- Total actual spend: ${fmtEUR(totalSpend)}
- Date range: ${dates[0] || '?'} to ${dates[dates.length - 1] || '?'}
- Total records: ${filteredData.length}
- Suppliers: ${new Set(actual.map(r => r.supplier).filter(Boolean)).size}
- Categories: ${new Set(actual.map(r => r.cost_category).filter(Boolean)).size}

CATEGORY BREAKDOWN (top 5):
${topCats.map(([c, v]) => '- ' + c.split(',')[0] + ': ' + fmtEUR(v) + ' (' + pct(v, totalSpend) + '%)').join('\n')}

AUTOMATED ANALYSIS — ${findings.length} FINDINGS (est. total savings: ${fmtEUR(totalSavings)}):
${findings.slice(0, 6).map((f, i) => (i + 1) + '. [' + f.priority.toUpperCase() + '] ' + f.title + ' — ' + f.category.split(',')[0] + '\n   ' + f.detail + '\n   Est. savings: ' + fmtEUR(f.estimatedSavings)).join('\n\n')}

Please provide a concise, actionable procurement action plan:
1. **Quick Wins (0–3 months)**: Highest-ROI actions requiring minimal lead time
2. **Strategic Initiatives (3–12 months)**: Sourcing projects, supplier negotiations, framework agreements
3. **Negotiation Tactics**: Specific leverage points for the top 2 opportunities
4. **KPIs to Track**: 3–4 metrics to measure progress

Keep the response under 450 words. Be specific and pharma-industry aware.`;

    try {
      const resp = await fetch('https://api.anthropic.com/v1/messages', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'x-api-key': apiKey,
          'anthropic-version': '2023-06-01',
          'anthropic-dangerous-direct-browser-access': 'true'
        },
        body: JSON.stringify({
          model,
          max_tokens: 900,
          messages: [{ role: 'user', content: prompt }]
        })
      });

      if (!resp.ok) {
        const err = await resp.json().catch(() => ({}));
        throw new Error(err.error?.message || 'API error ' + resp.status);
      }

      const result = await resp.json();
      const text = result.content?.[0]?.text || '';

      // Simple markdown → HTML (bold, headers, bullets)
      const formatted = text
        .replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>')
        .replace(/^### (.+)$/gm, '<strong style="display:block;margin-top:10px;color:var(--text)">$1</strong>')
        .replace(/^## (.+)$/gm, '<strong style="display:block;margin-top:12px;font-size:14px;color:var(--text)">$1</strong>')
        .replace(/^# (.+)$/gm, '<strong style="display:block;margin-top:12px;font-size:15px;color:var(--text)">$1</strong>')
        .replace(/^(\d+)\. /gm, '<br><strong>$1.</strong> ')
        .replace(/^[-•] /gm, '<br>• ');

      content.innerHTML = formatted;
    } catch (e) {
      content.innerHTML = '<span style="color:var(--danger)">Error: ' + e.message + '</span>';
      toast('AI request failed: ' + e.message, 'error');
    }
  }

  // ============================================================
  // SAP S4 HANA IMPORT WIZARD
  // ============================================================

  const SAP_FIELD_MAP = {
    // Purchase Order fields
    'EBELN': { target: 'po_number', label: 'PO Number' },
    'Purchasing Document': { target: 'po_number', label: 'PO Number' },
    'Purchase Order': { target: 'po_number', label: 'PO Number' },
    'PO Number': { target: 'po_number', label: 'PO Number' },
    'Einkaufsbeleg': { target: 'po_number', label: 'PO Number' },
    'Einkaufsbel.': { target: 'po_number', label: 'PO Number' },
    
    // Material / SKU
    'MATNR': { target: 'sku', label: 'Material Number / SKU' },
    'Material': { target: 'sku', label: 'Material Number / SKU' },
    'Material Number': { target: 'sku', label: 'Material Number / SKU' },
    'Materialnr.': { target: 'sku', label: 'Material Number / SKU' },
    'Materialnummer': { target: 'sku', label: 'Material Number / SKU' },
    
    // Description
    'TXZ01': { target: 'item_description', label: 'Description' },
    'Short Text': { target: 'item_description', label: 'Description' },
    'Description': { target: 'item_description', label: 'Description' },
    'Material Description': { target: 'item_description', label: 'Description' },
    'Item Text': { target: 'item_description', label: 'Description' },
    'Kurztext': { target: 'item_description', label: 'Description' },
    'Bezeichnung': { target: 'item_description', label: 'Description' },
    'Material description': { target: 'item_description', label: 'Description' },
    
    // Supplier / Vendor
    'LIFNR': { target: 'supplier', label: 'Vendor Number' },
    'NAME1': { target: 'supplier', label: 'Vendor Name' },
    'Vendor': { target: 'supplier', label: 'Vendor' },
    'Vendor Name': { target: 'supplier', label: 'Vendor Name' },
    'Supplier': { target: 'supplier', label: 'Supplier' },
    'Supplier Name': { target: 'supplier', label: 'Supplier Name' },
    'Lieferant': { target: 'supplier', label: 'Vendor' },
    'Kreditor': { target: 'supplier', label: 'Vendor' },
    'Name 1': { target: 'supplier', label: 'Vendor Name' },
    'Vendor name': { target: 'supplier', label: 'Vendor Name' },
    
    // Quantity
    'MENGE': { target: 'quantity', label: 'Quantity' },
    'PO Quantity': { target: 'quantity', label: 'Quantity' },
    'Order Quantity': { target: 'quantity', label: 'Quantity' },
    'Quantity': { target: 'quantity', label: 'Quantity' },
    'Bestellmenge': { target: 'quantity', label: 'Quantity' },
    'Qty': { target: 'quantity', label: 'Quantity' },
    'PO quantity': { target: 'quantity', label: 'Quantity' },
    
    // Price
    'NETPR': { target: 'unit_price_usd', label: 'Net Price' },
    'Net Price': { target: 'unit_price_usd', label: 'Net Price' },
    'Price': { target: 'unit_price_usd', label: 'Net Price' },
    'Unit Price': { target: 'unit_price_usd', label: 'Net Price' },
    'Nettopreis': { target: 'unit_price_usd', label: 'Net Price' },
    'Net price': { target: 'unit_price_usd', label: 'Net Price' },
    
    // Net Value / Total
    'NETWR': { target: 'total_amount_usd', label: 'Net Value' },
    'Net Value': { target: 'total_amount_usd', label: 'Net Value' },
    'Net Order Value': { target: 'total_amount_usd', label: 'Net Value' },
    'Net order value': { target: 'total_amount_usd', label: 'Net Value' },
    'Amount': { target: 'total_amount_usd', label: 'Net Value' },
    'Total Amount': { target: 'total_amount_usd', label: 'Net Value' },
    'Nettowert': { target: 'total_amount_usd', label: 'Net Value' },
    'Nettobest.wert': { target: 'total_amount_usd', label: 'Net Value' },
    'Value': { target: 'total_amount_usd', label: 'Net Value' },
    'Net order val.': { target: 'total_amount_usd', label: 'Net Value' },
    
    // Cost Center
    'KOSTL': { target: 'cost_center', label: 'Cost Center' },
    'Cost Center': { target: 'cost_center', label: 'Cost Center' },
    'CostCenter': { target: 'cost_center', label: 'Cost Center' },
    'Cost center': { target: 'cost_center', label: 'Cost Center' },
    'Kostenstelle': { target: 'cost_center', label: 'Cost Center' },
    
    // Date
    'BEDAT': { target: 'date', label: 'PO Date' },
    'PO Date': { target: 'date', label: 'PO Date' },
    'Document Date': { target: 'date', label: 'Document Date' },
    'Posting Date': { target: 'date', label: 'Posting Date' },
    'Created On': { target: 'date', label: 'Created On' },
    'Belegdatum': { target: 'date', label: 'Document Date' },
    'Doc. Date': { target: 'date', label: 'Document Date' },
    'Order Date': { target: 'date', label: 'Order Date' },
    'Delivery Date': { target: 'date', label: 'Delivery Date' },
    
    // Created By / Ordered By
    'ERNAM': { target: 'ordered_by', label: 'Created By' },
    'Created By': { target: 'ordered_by', label: 'Created By' },
    'Created by': { target: 'ordered_by', label: 'Created By' },
    'Requisitioner': { target: 'ordered_by', label: 'Requisitioner' },
    'Angelegt von': { target: 'ordered_by', label: 'Created By' },
    'Anforderer': { target: 'ordered_by', label: 'Requisitioner' },
    'Requisitioner name': { target: 'ordered_by', label: 'Requisitioner' },
    
    // Material Group -> Category
    'MATKL': { target: 'cost_category', label: 'Material Group' },
    'Material Group': { target: 'cost_category', label: 'Material Group' },
    'Material Grp': { target: 'cost_category', label: 'Material Group' },
    'Mat. Group': { target: 'cost_category', label: 'Material Group' },
    'Warengruppe': { target: 'cost_category', label: 'Material Group' },
    'Commodity': { target: 'cost_category', label: 'Material Group' },
    'Purchasing Group': { target: 'department', label: 'Purchasing Group' },
    'Purch. Group': { target: 'department', label: 'Purchasing Group' },
    'Einkaufsgruppe': { target: 'department', label: 'Purchasing Group' },
    
    // Plant
    'WERKS': { target: 'department', label: 'Plant' },
    'Plant': { target: 'department', label: 'Plant' },
    'Werk': { target: 'department', label: 'Plant' },
    
    // Currency
    'WAERS': { target: '_currency', label: 'Currency' },
    'Currency': { target: '_currency', label: 'Currency' },
    'Währung': { target: '_currency', label: 'Currency' },
    'Doc. Currency': { target: '_currency', label: 'Currency' },
    'Document Currency': { target: '_currency', label: 'Currency' },
    
    // PO Item
    'EBELP': { target: '_po_item', label: 'PO Item' },
    'Item': { target: '_po_item', label: 'PO Item' },
    'PO Item': { target: '_po_item', label: 'PO Item' },
    
    // Company Code
    'BUKRS': { target: '_company_code', label: 'Company Code' },
    'Company Code': { target: '_company_code', label: 'Company Code' },
    'Buchungskreis': { target: '_company_code', label: 'Company Code' },
    
    // GL Account
    'SAKTO': { target: 'sub_category', label: 'GL Account' },
    'G/L Account': { target: 'sub_category', label: 'GL Account' },
    'GL Account': { target: 'sub_category', label: 'GL Account' },
    'Sachkonto': { target: 'sub_category', label: 'GL Account' }
  };

  const SAP_STORAGE_KEY = 'biopharma_sap_mappings';

  let sapWizardState = {
    step: 0,
    rawData: [],
    headers: [],
    columnMapping: {},
    categoryMapping: {},
    numberFormat: 'auto',
    dateFormat: 'auto',
    detectedSAP: false,
    issues: [],
    previewData: []
  };

  function isSAPExport(headers) {
    const sapKeys = Object.keys(SAP_FIELD_MAP);
    let matchCount = 0;
    headers.forEach(h => {
      const trimmed = h.trim();
      if (sapKeys.includes(trimmed)) matchCount++;
    });
    return matchCount >= 2;
  }

  function autoDetectNumberFormat(data, headers) {
    let euCount = 0, usCount = 0;
    const numericHeaders = headers.filter(h => {
      const mapped = SAP_FIELD_MAP[h.trim()];
      return mapped && ['quantity','unit_price_usd','total_amount_usd'].includes(mapped.target);
    });
    
    data.slice(0, 50).forEach(row => {
      numericHeaders.forEach(h => {
        const val = String(row[h] || '');
        if (val.match(/\d+\.\d{3},\d/)) euCount++;
        else if (val.match(/\d+,\d{3}\.\d/)) usCount++;
        else if (val.match(/,\d{2}$/) && !val.match(/\.\d{2}$/)) euCount++;
        else if (val.match(/\.\d{2}$/) && !val.match(/,\d{2}$/)) usCount++;
      });
    });
    return euCount > usCount ? 'EU' : 'US';
  }

  function parseEUNumber(val) {
    if (val == null || val === '') return 0;
    let s = String(val).trim();
    s = s.replace(/\s/g, '');
    // Remove thousands separator (dot) and convert decimal comma to dot
    s = s.replace(/\./g, '').replace(',', '.');
    // Handle negative with trailing minus
    if (s.endsWith('-')) s = '-' + s.slice(0, -1);
    const n = parseFloat(s);
    return isNaN(n) ? 0 : n;
  }

  function parseUSNumber(val) {
    if (val == null || val === '') return 0;
    let s = String(val).trim();
    s = s.replace(/\s/g, '');
    s = s.replace(/,/g, '');
    if (s.endsWith('-')) s = '-' + s.slice(0, -1);
    const n = parseFloat(s);
    return isNaN(n) ? 0 : n;
  }

  function parseSAPDate(val) {
    if (!val) return '';
    let s = String(val).trim();
    // SAP formats: DD.MM.YYYY, YYYY-MM-DD, YYYYMMDD, MM/DD/YYYY, DD/MM/YYYY
    let match;
    if ((match = s.match(/^(\d{4})-(\d{2})-(\d{2})/))) return match[1] + '-' + match[2];
    if ((match = s.match(/^(\d{2})\.(\d{2})\.(\d{4})/))) return match[3] + '-' + match[2];
    if ((match = s.match(/^(\d{8})$/))) return s.substring(0, 4) + '-' + s.substring(4, 6);
    if ((match = s.match(/^(\d{2})\/(\d{2})\/(\d{4})/))) return match[3] + '-' + match[1];
    if ((match = s.match(/^(\d{4})\/(\d{2})/))) return match[1] + '-' + match[2];
    return s.substring(0, 7);
  }

  function loadSavedCategoryMappings() {
    try {
      const raw = localStorage.getItem(SAP_STORAGE_KEY);
      if (raw) return JSON.parse(raw);
    } catch(e) {}
    return {};
  }

  function saveCategoryMappings(mappings) {
    try { localStorage.setItem(SAP_STORAGE_KEY, JSON.stringify(mappings)); } catch(e) {}
  }

  function openSAPWizard() {
    document.getElementById('sap-file-input').click();
  }

  function startSAPWizard(csvText) {
    Papa.parse(csvText, {
      header: true,
      skipEmptyLines: true,
      dynamicTyping: false,
      complete: function(results) {
        if (results.data.length === 0) { toast('No data found in file', 'error'); return; }
        
        sapWizardState = {
          step: 0,
          rawData: results.data,
          headers: results.meta.fields || [],
          columnMapping: {},
          categoryMapping: loadSavedCategoryMappings(),
          numberFormat: 'auto',
          dateFormat: 'auto',
          detectedSAP: false,
          issues: [],
          previewData: []
        };
        
        // Auto-detect SAP columns
        sapWizardState.detectedSAP = isSAPExport(sapWizardState.headers);
        
        // Auto-map columns
        sapWizardState.headers.forEach(h => {
          const trimmed = h.trim();
          const mapping = SAP_FIELD_MAP[trimmed];
          if (mapping) {
            // Don't overwrite if already mapped (e.g., two vendor fields)
            const alreadyMapped = Object.values(sapWizardState.columnMapping).some(v => v === mapping.target);
            if (!alreadyMapped || mapping.target === '_currency' || mapping.target === '_po_item' || mapping.target === '_company_code') {
              sapWizardState.columnMapping[h] = mapping.target;
            }
          }
        });
        
        // Auto-detect number format
        const detectedFormat = autoDetectNumberFormat(results.data, sapWizardState.headers);
        sapWizardState.numberFormat = detectedFormat;
        
        document.getElementById('modal-sap-wizard').classList.add('active');
        renderSAPStep();
      },
      error: function(err) { toast('Parse error: ' + err.message, 'error'); }
    });
  }

  function renderSAPStep() {
    const steps = ['Upload & Detect', 'Map Columns', 'Map Categories', 'Settings & Preview', 'Import'];
    const stepsDiv = document.getElementById('sap-wizard-steps');
    stepsDiv.innerHTML = steps.map((s, i) =>
      '<div class="sap-step-dot ' + (i === sapWizardState.step ? 'active' : (i < sapWizardState.step ? 'completed' : '')) + '" title="' + s + '"></div>'
    ).join('');
    
    document.getElementById('sap-wizard-title').textContent = 'SAP Import - ' + steps[sapWizardState.step];
    
    const backBtn = document.getElementById('sap-wizard-back');
    const nextBtn = document.getElementById('sap-wizard-next');
    backBtn.style.display = sapWizardState.step > 0 ? '' : 'none';
    
    if (sapWizardState.step === steps.length - 1) {
      nextBtn.textContent = 'Import Data';
      nextBtn.className = 'btn btn-success';
    } else {
      nextBtn.textContent = 'Next';
      nextBtn.className = 'btn btn-primary';
    }
    
    const body = document.getElementById('sap-wizard-body');
    
    switch (sapWizardState.step) {
      case 0: renderSAPStep0(body); break;
      case 1: renderSAPStep1(body); break;
      case 2: renderSAPStep2(body); break;
      case 3: renderSAPStep3(body); break;
      case 4: executeSAPImport(); break;
    }
  }

  // Step 0: Detection Summary
  function renderSAPStep0(body) {
    const h = sapWizardState.headers;
    const d = sapWizardState.rawData;
    const mapped = Object.keys(sapWizardState.columnMapping).length;
    const unmapped = h.length - mapped;
    
    let html = '';
    if (sapWizardState.detectedSAP) {
      html += '<div class="sap-detected-banner"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M22 11.08V12a10 10 0 11-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/></svg>SAP S4 HANA export detected! Auto-mapped ' + mapped + ' of ' + h.length + ' columns.</div>';
    } else {
      html += '<div class="sap-warning-banner"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M10.29 3.86L1.82 18a2 2 0 001.71 3h16.94a2 2 0 001.71-3L13.71 3.86a2 2 0 00-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/></svg>File does not appear to be a standard SAP export. You can still map columns manually in the next step.</div>';
    }
    
    html += '<div class="preview-stats">';
    html += '<div class="preview-stat"><div class="label">Rows</div><div class="value">' + d.length + '</div></div>';
    html += '<div class="preview-stat"><div class="label">Columns</div><div class="value">' + h.length + '</div></div>';
    html += '<div class="preview-stat"><div class="label">Auto-Mapped</div><div class="value" style="color:var(--success)">' + mapped + '</div></div>';
    html += '<div class="preview-stat"><div class="label">Need Review</div><div class="value" style="color:' + (unmapped > 0 ? 'var(--warning)' : 'var(--success)') + '">' + unmapped + '</div></div>';
    html += '<div class="preview-stat"><div class="label">Number Format</div><div class="value" style="font-size:16px">' + sapWizardState.numberFormat + '</div></div>';
    html += '</div>';
    
    html += '<div class="card-title" style="margin-bottom:8px">Detected Columns</div>';
    html += '<div style="max-height:200px;overflow-y:auto;border:1px solid var(--border);border-radius:var(--radius)">';
    html += '<table class="mapping-table"><thead><tr><th>Source Column</th><th>Sample Value</th><th>Auto-Mapped To</th></tr></thead><tbody>';
    h.forEach(col => {
      const sample = d[0] ? (d[0][col] || '') : '';
      const target = sapWizardState.columnMapping[col];
      const targetCol = target ? COLUMNS.find(c => c.key === target) : null;
      const cls = targetCol ? 'mapped' : (target && target.startsWith('_') ? 'mapped' : 'unmapped');
      const label = targetCol ? targetCol.label : (target ? target.replace('_', '') : 'Not mapped');
      html += '<tr><td class="sap-col-name">' + col + '</td><td class="sap-sample" title="' + String(sample).replace(/"/g, '&quot;') + '">' + String(sample).substring(0, 40) + '</td>';
      html += '<td class="' + cls + '">' + label + '</td></tr>';
    });
    html += '</tbody></table></div>';
    
    body.innerHTML = html;
  }

  // Step 1: Column Mapping
  function renderSAPStep1(body) {
    const h = sapWizardState.headers;
    const d = sapWizardState.rawData;
    
    const targetOptions = COLUMNS.map(c => '<option value="' + c.key + '">' + c.label + '</option>').join('');
    const extraOptions = '<option value="_skip">-- Skip / Ignore --</option><option value="_currency">Currency (for reference)</option><option value="_po_item">PO Item (for reference)</option><option value="_company_code">Company Code (for ref)</option>';
    
    let html = '<div class="sap-step-title">Map SAP Columns to App Fields</div>';
    html += '<div class="sap-step-desc">Review and adjust the column mapping. Each SAP column should map to one app field.</div>';
    
    html += '<div style="max-height:350px;overflow-y:auto;">';
    html += '<table class="mapping-table"><thead><tr><th>SAP Column</th><th>Sample Values</th><th>Maps To</th></tr></thead><tbody>';
    
    h.forEach((col, idx) => {
      const samples = d.slice(0, 3).map(r => String(r[col] || '')).filter(Boolean);
      const sampleStr = samples.join(' | ');
      const currentMapping = sapWizardState.columnMapping[col] || '_skip';
      const isAuto = SAP_FIELD_MAP[col.trim()] !== undefined;
      
      html += '<tr>';
      html += '<td><span class="sap-col-name">' + col + '</span>';
      if (isAuto) html += '<span class="mapping-auto-badge auto">auto</span>';
      html += '</td>';
      html += '<td class="sap-sample" title="' + sampleStr.replace(/"/g, '&quot;') + '">' + sampleStr.substring(0, 50) + '</td>';
      html += '<td><select data-sap-col="' + idx + '">';
      html += '<option value="_skip"' + (currentMapping === '_skip' ? ' selected' : '') + '>-- Skip --</option>';
      
      COLUMNS.forEach(c => {
        html += '<option value="' + c.key + '"' + (currentMapping === c.key ? ' selected' : '') + '>' + c.label + '</option>';
      });
      html += '<option value="_currency"' + (currentMapping === '_currency' ? ' selected' : '') + '>Currency (ref)</option>';
      html += '<option value="_po_item"' + (currentMapping === '_po_item' ? ' selected' : '') + '>PO Item (ref)</option>';
      html += '<option value="_company_code"' + (currentMapping === '_company_code' ? ' selected' : '') + '>Company Code (ref)</option>';
      html += '</select></td>';
      html += '</tr>';
    });
    
    html += '</tbody></table></div>';
    body.innerHTML = html;
    
    // Bind change events
    body.querySelectorAll('select[data-sap-col]').forEach(sel => {
      sel.addEventListener('change', () => {
        const colIdx = parseInt(sel.dataset.sapCol);
        const colName = sapWizardState.headers[colIdx];
        sapWizardState.columnMapping[colName] = sel.value;
      });
    });
  }

  // Step 2: Category Mapping (MATKL -> our categories)
  function renderSAPStep2(body) {
    const d = sapWizardState.rawData;
    const catCol = Object.entries(sapWizardState.columnMapping).find(([k, v]) => v === 'cost_category');
    
    if (!catCol) {
      body.innerHTML = '<div class="sap-step-title">Category Mapping</div><div class="sap-step-desc">No column is mapped to Cost Category. You can skip this step - all entries will be assigned to "Miscellaneous Indirect Costs".</div><p style="color:var(--text-muted)">Go back and map a column (like Material Group / MATKL) to Category, or proceed and all data will be categorized as "Miscellaneous Indirect Costs".</p>';
      return;
    }
    
    // Collect unique values and counts
    const valueCounts = {};
    d.forEach(r => {
      const val = String(r[catCol[0]] || '').trim();
      if (val) {
        if (!valueCounts[val]) valueCounts[val] = 0;
        valueCounts[val]++;
      }
    });
    
    const sortedValues = Object.entries(valueCounts).sort((a, b) => b[1] - a[1]);
    const savedMappings = sapWizardState.categoryMapping;
    
    let html = '<div class="sap-step-title">Map SAP Values to Spend Categories</div>';
    html += '<div class="sap-step-desc">Map each unique SAP material group / value (' + sortedValues.length + ' found) to one of the 6 spend categories. Previously saved mappings are loaded automatically.</div>';
    
    if (sortedValues.length > 30) {
      html += '<div class="sap-warning-banner"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="16" height="16"><path d="M10.29 3.86L1.82 18a2 2 0 001.71 3h16.94a2 2 0 001.71-3L13.71 3.86a2 2 0 00-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/></svg>' + sortedValues.length + ' unique values found. Review the most frequent ones - rare values can default to "Other".</div>';
    }
    
    html += '<div style="max-height:350px;overflow-y:auto">';
    html += '<table class="cat-mapping-table"><thead><tr><th>SAP Value</th><th>Count</th><th>Map To Category</th></tr></thead><tbody>';
    
    sortedValues.forEach(([val, count]) => {
      const saved = savedMappings[val] || '';
      // Try to auto-guess based on keywords
      let autoGuess = saved;
      if (!autoGuess) {
        const lower = val.toLowerCase();
        if (lower.includes('lab') || lower.includes('clinical') || lower.includes('test') || lower.includes('analyt') || lower.includes('scientific')) autoGuess = 'Clinical, Lab and scientific services';
        else if (lower.includes('equip') || lower.includes('machine') || lower.includes('reactor') || lower.includes('prod') || lower.includes('manufactur')) autoGuess = 'Production Equipment';
        else if (lower.includes('warehouse') || lower.includes('logistics') || lower.includes('distrib') || lower.includes('transport') || lower.includes('storage')) autoGuess = 'External Warehouse and distribution';
        else if (lower.includes('consult') || lower.includes('professional') || lower.includes('advisory') || lower.includes('legal') || lower.includes('audit')) autoGuess = 'Professional Services';
        else if (lower.includes('office') || lower.includes('print') || lower.includes('stationery') || lower.includes('paper') || lower.includes('toner')) autoGuess = 'Office and Print';
        else if (lower.includes('misc') || lower.includes('other') || lower.includes('facility') || lower.includes('utilit') || lower.includes('general')) autoGuess = 'Miscellaneous Indirect Costs';
      }
      
      html += '<tr><td class="sap-value">' + val + '</td>';
      html += '<td class="count">' + count + '</td>';
      html += '<td><select data-cat-val="' + val.replace(/"/g, '&quot;') + '">';
      html += '<option value=""' + (!autoGuess ? ' selected' : '') + '>-- Select Category --</option>';
      CATEGORIES.forEach(cat => {
        html += '<option value="' + cat + '"' + (autoGuess === cat ? ' selected' : '') + '>' + cat + '</option>';
      });
      html += '</select></td></tr>';
      
      if (autoGuess) sapWizardState.categoryMapping[val] = autoGuess;
    });
    
    html += '</tbody></table></div>';
    body.innerHTML = html;
    
    body.querySelectorAll('select[data-cat-val]').forEach(sel => {
      sel.addEventListener('change', () => {
        sapWizardState.categoryMapping[sel.dataset.catVal] = sel.value;
      });
    });
  }

  // Step 3: Settings & Preview
  function renderSAPStep3(body) {
    // Process full data once — use it for both issues and stats
    const fullPreview = processSAPData(sapWizardState.rawData);
    sapWizardState.previewData = fullPreview.data.slice(0, 100);
    sapWizardState.issues = fullPreview.issues;

    const totalAmount = fullPreview.data.reduce((s, r) => s + r.total_amount_usd, 0);
    const uniqueSuppliers = new Set(fullPreview.data.map(r => r.supplier).filter(Boolean)).size;
    const uniqueSKUs = new Set(fullPreview.data.map(r => r.sku).filter(Boolean)).size;
    const dateRange = fullPreview.data.map(r => r.date).filter(Boolean).sort();
    
    let html = '<div class="sap-step-title">Review & Settings</div>';
    html += '<div class="sap-step-desc">Verify the import settings and preview the transformed data before importing.</div>';
    
    // Number format setting
    html += '<div class="number-format-toggle">';
    html += '<label>Number Format:</label>';
    html += '<select id="sap-number-format">';
    html += '<option value="US"' + (sapWizardState.numberFormat === 'US' ? ' selected' : '') + '>US (1,234.56 - comma=thousands, dot=decimal)</option>';
    html += '<option value="EU"' + (sapWizardState.numberFormat === 'EU' ? ' selected' : '') + '>European (1.234,56 - dot=thousands, comma=decimal)</option>';
    html += '</select></div>';
    
    // Stats
    html += '<div class="preview-stats">';
    html += '<div class="preview-stat"><div class="label">Total Rows</div><div class="value">' + fullPreview.data.length + '</div></div>';
    html += '<div class="preview-stat"><div class="label">Total Spend</div><div class="value" style="font-size:16px">' + fmtEUR(totalAmount) + '</div></div>';
    html += '<div class="preview-stat"><div class="label">Suppliers</div><div class="value">' + uniqueSuppliers + '</div></div>';
    html += '<div class="preview-stat"><div class="label">SKUs</div><div class="value">' + uniqueSKUs + '</div></div>';
    html += '<div class="preview-stat"><div class="label">Date Range</div><div class="value" style="font-size:14px">' + (dateRange[0] || '?') + ' - ' + (dateRange[dateRange.length-1] || '?') + '</div></div>';
    html += '</div>';
    
    // Issues
    if (fullPreview.issues.length > 0) {
      html += '<div class="card-title" style="margin:12px 0 8px">Issues Found (' + fullPreview.issues.length + ')</div>';
      html += '<ul class="sap-issues-list">';
      fullPreview.issues.slice(0, 15).forEach(issue => {
        const icon = issue.type === 'error' ? '&#x26D4;' : (issue.type === 'warn' ? '&#x26A0;' : '&#x2139;');
        html += '<li><span class="issue-icon ' + issue.type + '">' + icon + '</span>' + issue.message + '</li>';
      });
      if (fullPreview.issues.length > 15) html += '<li style="color:var(--text-muted)">... and ' + (fullPreview.issues.length - 15) + ' more</li>';
      html += '</ul>';
    }
    
    // Preview table
    html += '<div class="card-title" style="margin:16px 0 8px">Data Preview (first 10 rows after transformation)</div>';
    html += '<div class="preview-table-wrap"><table class="data-table"><thead><tr>';
    ['Date','Category','SKU','Description','Supplier','Ordered By','Qty','Total (EUR)'].forEach(h => { html += '<th>' + h + '</th>'; });
    html += '</tr></thead><tbody>';
    fullPreview.data.slice(0, 10).forEach(row => {
      html += '<tr>';
      html += '<td>' + (row.date || '') + '</td>';
      html += '<td>' + (row.cost_category || '') + '</td>';
      html += '<td>' + (row.sku || '') + '</td>';
      html += '<td>' + (row.item_description || '').substring(0, 35) + '</td>';
      html += '<td>' + (row.supplier || '') + '</td>';
      html += '<td>' + (row.ordered_by || '') + '</td>';
      html += '<td class="num">' + fmt(row.quantity) + '</td>';
      html += '<td class="currency">' + fmtEUR(row.total_amount_usd) + '</td>';
      html += '</tr>';
    });
    html += '</tbody></table></div>';
    
    body.innerHTML = html;
    
    document.getElementById('sap-number-format').addEventListener('change', (e) => {
      sapWizardState.numberFormat = e.target.value;
      renderSAPStep3(body);
    });
  }

  function processSAPData(rawRows) {
    const mapping = sapWizardState.columnMapping;
    const catMapping = sapWizardState.categoryMapping;
    const numFmt = sapWizardState.numberFormat;
    const parseNumber = numFmt === 'EU' ? parseEUNumber : parseUSNumber;
    const issues = [];
    const reverseMap = {};
    
    Object.entries(mapping).forEach(([sapCol, appField]) => {
      if (appField && !appField.startsWith('_')) {
        reverseMap[appField] = sapCol;
      }
    });
    
    const data = rawRows.map((raw, idx) => {
      const row = {};
      COLUMNS.forEach(col => {
        const sapCol = reverseMap[col.key];
        if (!sapCol) {
          row[col.key] = col.type === 'number' ? 0 : '';
          return;
        }
        let val = raw[sapCol];
        if (val === undefined || val === null) val = '';
        
        if (col.key === 'date') {
          row[col.key] = parseSAPDate(val);
        } else if (col.key === 'cost_category') {
          const catVal = String(val).trim();
          row[col.key] = catMapping[catVal] || 'Miscellaneous Indirect Costs';
          if (!catMapping[catVal] && catVal && idx < 5) {
            issues.push({ type: 'warn', message: 'Row ' + (idx+1) + ': Unknown category "' + catVal + '" mapped to Miscellaneous Indirect Costs' });
          }
        } else if (col.type === 'number') {
          row[col.key] = parseNumber(val);
        } else {
          row[col.key] = String(val).trim();
        }
      });
      
      // Auto-calculate total if missing
      if (row.total_amount_usd === 0 && row.quantity > 0 && row.unit_price_usd > 0) {
        row.total_amount_usd = row.quantity * row.unit_price_usd;
      }
      if (!row.budget_type) row.budget_type = 'Actual';
      
      // Validation
      if (!row.date && idx < 5) issues.push({ type: 'warn', message: 'Row ' + (idx+1) + ': Missing date' });
      if (row.total_amount_usd === 0 && idx < 5) issues.push({ type: 'info', message: 'Row ' + (idx+1) + ': Zero amount' });
      
      return row;
    });
    
    // Filter out completely empty rows
    const filtered = data.filter(r => r.total_amount_usd !== 0 || r.sku || r.supplier || r.item_description);
    if (data.length - filtered.length > 0) {
      issues.push({ type: 'info', message: (data.length - filtered.length) + ' empty rows removed' });
    }
    
    return { data: filtered, issues };
  }

  function executeSAPImport() {
    const result = processSAPData(sapWizardState.rawData);
    
    if (result.data.length === 0) {
      toast('No data to import after processing', 'error');
      return;
    }
    
    // Save category mappings for next time
    saveCategoryMappings(sapWizardState.categoryMapping);
    
    // Add to existing data
    allData = allData.concat(result.data);
    reindexData();
    saveToStorage();
    updateFilterOptions();
    applyGlobalFilters();
    refreshAll();
    updateFooter();
    
    // Close wizard
    document.getElementById('modal-sap-wizard').classList.remove('active');
    
    toast('SAP import complete: ' + result.data.length + ' records added', 'success');
    if (result.issues.length > 0) {
      toast(result.issues.filter(i => i.type === 'warn').length + ' warnings during import', 'warning');
    }
  }

  function setupSAPWizardEvents() {
    // SAP import button
    document.getElementById('btn-sap-import').addEventListener('click', openSAPWizard);
    
    // SAP file input
    document.getElementById('sap-file-input').addEventListener('change', (e) => {
      const file = e.target.files[0];
      if (file) {
        const reader = new FileReader();
        reader.onload = (ev) => startSAPWizard(ev.target.result);
        reader.readAsText(file);
        e.target.value = '';
      }
    });
    
    // Wizard navigation
    document.getElementById('sap-wizard-next').addEventListener('click', () => {
      if (sapWizardState.step < 4) {
        sapWizardState.step++;
        renderSAPStep();
      }
    });
    
    document.getElementById('sap-wizard-back').addEventListener('click', () => {
      if (sapWizardState.step > 0) {
        sapWizardState.step--;
        renderSAPStep();
      }
    });
    
    document.getElementById('sap-wizard-close').addEventListener('click', () => {
      document.getElementById('modal-sap-wizard').classList.remove('active');
    });
    
    // Also auto-detect SAP format on regular file upload
    const origFileInput = document.getElementById('file-input');
    const origUpload = document.getElementById('upload-area');
    
    // Override the upload handler to auto-detect SAP
    origUpload.removeEventListener('click', origUpload._clickHandler);
    origUpload._clickHandler = () => origFileInput.click();
    origUpload.addEventListener('click', origUpload._clickHandler);
  }

  // ---- Initialize ----
  function init() {
    aiAdvisorKey = localStorage.getItem(AI_KEY_STORAGE) || '';
    loadTargets();
    setupEventListeners();
    setupTabs();
    setupSAPWizardEvents();

    const hasData = loadFromStorage();
    if (hasData) {
      updateFilterOptions();
      applyGlobalFilters();
      refreshAll();
      updateFooter();
      toast('Data restored from previous session (' + allData.length + ' records)', 'info');
    } else {
      applyGlobalFilters();
      refreshAll();
    }
  }

  // Start
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }

})();
