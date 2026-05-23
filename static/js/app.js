/**
 * CustomsMID Automation — Frontend Logic
 * Handles drag & drop, file selection, upload, progress, and results rendering.
 */

'use strict';

// Scroll to top on refresh — don't remember scroll position
if (history.scrollRestoration) {
  history.scrollRestoration = 'manual';
}
window.scrollTo(0, 0);

// ── DOM References ────────────────────────────────────────────────────────────
const dropZone      = document.getElementById('drop-zone');
const fileInput     = document.getElementById('file-input');
const filePreview   = document.getElementById('file-preview');
const fileNameEl    = document.getElementById('file-name');
const fileSizeEl    = document.getElementById('file-size');
const fileRemoveBtn = document.getElementById('file-remove');
const processBtn    = document.getElementById('process-btn');
const btnLabel      = document.getElementById('btn-label');
const progressWrap  = document.getElementById('progress-wrap');
const progressBar   = document.getElementById('progress-bar');
const progressText  = document.getElementById('progress-text');
const statusMsg     = document.getElementById('status-message');
const resultsCard   = document.getElementById('results-card');
const resultsMessage= document.getElementById('results-message');
const resultsStats  = document.getElementById('results-stats');
const downloadBtn   = document.getElementById('download-btn');
const loadingOverlay= document.getElementById('loading-overlay');
const loadingSub    = document.getElementById('loading-sub');
const dropTitle     = document.getElementById('drop-title');
const dropIcon      = document.getElementById('drop-icon');
const countrySelect = document.getElementById('country-select');

// No-MID Card DOM
const noMidCard      = document.getElementById('no-mid-card');
const noMidBadge     = document.getElementById('no-mid-badge');
const noMidTableBody = document.getElementById('no-mid-table-body');

// Fixed Card DOM
const fixedCard = document.getElementById('fixed-card');
const fixedBadge = document.getElementById('fixed-badge');
const fixedTableBody = document.getElementById('fixed-table-body');

// Preview DOM
const previewCard   = document.getElementById('preview-card');
const previewCountry= document.getElementById('preview-country-label');
const previewTotal  = document.getElementById('preview-total');
const previewValid  = document.getElementById('preview-valid');
const previewInvalid= document.getElementById('preview-invalid');
const previewEnvelope = document.getElementById('preview-envelope');
const previewTableBody = document.getElementById('preview-table-body');
const previewSpaces = document.getElementById('preview-spaces');
const previewModified = document.getElementById('preview-modified');

// ── State ─────────────────────────────────────────────────────────────────────
let selectedFile = null;
let isProcessing = false;
let currentSessionId = sessionStorage.getItem('currentSessionId') || null;
let currentPreviewType = 'valid';
let isProcessed = false;

// API base configuration: supports deploying frontend separately (e.g., Vercel)
// Set a meta tag in the HTML: <meta name="api-base" content="https://api.example.com" />
const _apiMeta = document.querySelector('meta[name="api-base"]');
const API_BASE = (window.API_BASE && typeof window.API_BASE === 'string' && window.API_BASE.trim())
  ? window.API_BASE.trim()
  : (_apiMeta ? (_apiMeta.content || '/') : '/');

function apiJoin(path) {
  const base = (API_BASE === '/' || !API_BASE) ? '' : (API_BASE.endsWith('/') ? API_BASE.slice(0, -1) : API_BASE);
  return path.startsWith('/') ? base + path : base + '/' + path;
}

// ── Utilities ─────────────────────────────────────────────────────────────────
function formatBytes(bytes) {
  if (bytes < 1024) return bytes + ' B';
  if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
  return (bytes / (1024 * 1024)).toFixed(2) + ' MB';
}

function setStatus(message, type = 'info') {
  statusMsg.textContent = message;
  statusMsg.className = `status-message ${type}`;
  statusMsg.hidden = false;
}

function clearStatus() {
  statusMsg.hidden = true;
  statusMsg.textContent = '';
}

function setProgress(pct, text) {
  progressWrap.hidden = false;
  progressBar.style.width = pct + '%';
  progressBar.setAttribute('aria-valuenow', pct);
  progressText.textContent = text;
}

function hideProgress() {
  progressWrap.hidden = true;
  progressBar.style.width = '0%';
}

function showLoading(sub = 'Cleaning MID codes & HS codes…') {
  loadingSub.textContent = sub;
  loadingOverlay.hidden = false;
  loadingOverlay.removeAttribute('aria-hidden');
}

function hideLoading() {
  loadingOverlay.hidden = true;
  loadingOverlay.setAttribute('aria-hidden', 'true');
}

const loadingMessages = [
  'Removing invalid rows…',
  'Parsing Manifested Descriptions…',
  'Cleaning MID codes…',
  'Recovering missing HS codes…',
  'Formatting output…',
  'Exporting to Excel…',
];

// ── File Handling ─────────────────────────────────────────────────────────────
function isValidFile(file) {
  const ext = file.name.split('.').pop().toLowerCase();
  return ['csv', 'xlsx', 'xls'].includes(ext);
}

function setFile(file) {
  if (!file) return;

  if (!isValidFile(file)) {
    setStatus('❌ Invalid file type. Please upload a CSV, XLSX, or XLS file.', 'error');
    return;
  }
  if (file.size > 16 * 1024 * 1024) {
    setStatus('❌ File too large. Maximum allowed size is 16 MB.', 'error');
    return;
  }

  selectedFile = file;
  clearStatus();

  // Update preview
  fileNameEl.textContent = file.name;
  fileSizeEl.textContent = formatBytes(file.size);
  filePreview.hidden = false;

  // Update drop zone appearance
  dropZone.classList.add('has-file');
  dropTitle.textContent = 'File ready to process';

  // Enable process button
  processBtn.disabled = false;

  // Hide any old results
  resultsCard.hidden = true;
  hideProgress();
}

function clearFile(keepResults = false) {
  selectedFile = null;
  fileInput.value = '';
  filePreview.hidden = true;
  dropZone.classList.remove('has-file');
  dropTitle.textContent = 'Drop your file here';
  processBtn.disabled = true;
  previewCard.hidden = true;
  if (!keepResults) {
    clearStatus();
    hideProgress();
    resultsCard.hidden = true;
    if (noMidCard) noMidCard.hidden = true;
    currentSessionId = null;
    sessionStorage.removeItem('currentSessionId');
  }
}

// ── Drag & Drop ───────────────────────────────────────────────────────────────
dropZone.addEventListener('dragenter', (e) => {
  e.preventDefault();
  dropZone.classList.add('drag-over');
});
dropZone.addEventListener('dragover', (e) => {
  e.preventDefault();
  dropZone.classList.add('drag-over');
});
dropZone.addEventListener('dragleave', (e) => {
  if (!dropZone.contains(e.relatedTarget)) {
    dropZone.classList.remove('drag-over');
  }
});
dropZone.addEventListener('drop', (e) => {
  e.preventDefault();
  dropZone.classList.remove('drag-over');
  const file = e.dataTransfer.files[0];
  if (file) setFile(file);
});

// Click to browse
dropZone.addEventListener('click', () => {
  if (!isProcessing) fileInput.click();
});
dropZone.addEventListener('keydown', (e) => {
  if (e.key === 'Enter' || e.key === ' ') {
    e.preventDefault();
    if (!isProcessing) fileInput.click();
  }
});

fileInput.addEventListener('change', () => {
  if (fileInput.files[0]) setFile(fileInput.files[0]);
});

fileRemoveBtn.addEventListener('click', (e) => {
  e.stopPropagation();
  clearFile();
});

// ── Results Rendering ─────────────────────────────────────────────────────────
function renderResults(data) {
  const { stats, message, download_token } = data;

  resultsMessage.textContent =
    `${stats.removed_rows} rows removed · ${stats.processed_rows} rows processed · ` +
    `${stats.mid_formatted} MID codes cleaned`;

  resultsStats.innerHTML = `
    <div class="stat-box removed">
      <div class="stat-box-value">${stats.removed_rows}</div>
      <div class="stat-box-label">Rows Removed</div>
    </div>
    <div class="stat-box processed">
      <div class="stat-box-value">${stats.processed_rows}</div>
      <div class="stat-box-label">Rows Processed</div>
    </div>
    <div class="stat-box formatted">
      <div class="stat-box-value">${stats.mid_formatted}</div>
      <div class="stat-box-label">MID Codes Cleaned</div>
    </div>
  `;

  downloadBtn.href = apiJoin(`/download/${encodeURIComponent(download_token)}`);
  downloadBtn.setAttribute('download', '');
  resultsCard.hidden = false;

  // Smooth scroll to results
  setTimeout(() => {
    resultsCard.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
  }, 100);
}

// ── No-MID Card Rendering ────────────────────────────────────────────────────
function renderNoMidCard(noMidRows) {
  if (!noMidCard || !noMidTableBody || !noMidBadge) return;

  if (!noMidRows || noMidRows.length === 0) {
    noMidCard.hidden = true;
    return;
  }

  noMidBadge.textContent = noMidRows.length;
  noMidTableBody.innerHTML = '';

  noMidRows.forEach(row => {
    const tr = document.createElement('tr');
    const reason = row['Reason'] || '';
    // Color-code the reason badge
    const reasonColor = reason.includes('No HS') && reason.includes('No valid MID')
      ? '#dc2626'   // both missing → red
      : reason.includes('No HS')
        ? '#d97706' // only HS missing → amber
        : '#7c3aed'; // only MID missing → purple
    tr.innerHTML = `
      <td class="col-tracking">${row['Tracking Number'] || ''}</td>
      <td class="col-desc" title="${row['Manifested Description'] || ''}">${row['Manifested Description'] || ''}</td>
      <td><span style="background:${reasonColor};color:white;padding:2px 8px;border-radius:12px;font-size:0.78rem;white-space:nowrap;">${reason}</span></td>
    `;
    noMidTableBody.appendChild(tr);
  });

  noMidCard.hidden = false;
  setTimeout(() => {
    noMidCard.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
  }, 200);
}

// Render Fixed Shipments (original vs cleaned)
function renderFixedCard(fixedRows) {
  if (!fixedCard || !fixedTableBody || !fixedBadge) return;

  if (!fixedRows || fixedRows.length === 0) {
    fixedCard.hidden = true;
    return;
  }

  fixedBadge.textContent = fixedRows.length;
  fixedTableBody.innerHTML = '';

  fixedRows.forEach(row => {
    const tr = document.createElement('tr');
    const orig = row['original_manifested'] || '';
    const cleaned = row['cleaned_manifested'] || '';
    tr.innerHTML = `
      <td class="col-tracking">${row['Tracking Number'] || ''}</td>
      <td class="col-desc" title="${orig}">${orig}</td>
      <td class="col-desc" title="${cleaned}">${cleaned}</td>
    `;
    fixedTableBody.appendChild(tr);
  });

  fixedCard.hidden = false;
}

// Render Spaces Found Card
function renderSpacesCard(spacesRows) {
  const card = document.getElementById('spaces-found-card');
  const badge = document.getElementById('spaces-found-badge');
  const tbody = document.getElementById('spaces-found-table-body');
  if (!card || !tbody || !badge) return;

  if (!spacesRows || spacesRows.length === 0) {
    card.hidden = true;
    return;
  }

  badge.textContent = spacesRows.length;
  tbody.innerHTML = '';

  spacesRows.forEach(row => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td class="col-tracking">${row['Tracking Number'] || ''}</td>
      <td class="col-desc" title="${row['Manifested Description'] || ''}">${row['Manifested Description'] || ''}</td>
      <td>${row['Spaces Detail'] || ''}</td>
    `;
    tbody.appendChild(tr);
  });

  card.hidden = false;
}

// ── Processing ────────────────────────────────────────────────────────────────
async function analyzeFile() {
  if (!selectedFile || isProcessing) return;

  // Cache check!
  if (!previewCard) {
    alert("⚠️ New features have been added! Please hard refresh your browser (Ctrl+Shift+R or Ctrl+F5) to use the app.");
    return;
  }

  isProcessing = true;

  // UI state
  processBtn.disabled = true;
  btnLabel.textContent = 'Analyzing…';
  clearStatus();
  resultsCard.hidden = true;
  previewCard.hidden = true;
  showLoading('Analyzing file shipments…');

  const formData = new FormData();
  formData.append('file', selectedFile);
  const countryVal = countrySelect ? countrySelect.value : 'US';
  formData.append('country', countryVal);

  try {
    const response = await fetch(apiJoin('/analyze'), {
      method: 'POST',
      body: formData,
    });
    const result = await response.json();

    if (result.success) {
      currentSessionId = result.preview.session_id;
      sessionStorage.setItem('currentSessionId', currentSessionId);
      
      // Show Preview
      if (previewCountry) previewCountry.textContent = result.preview.country === 'ALL' ? 'All Countries' : result.preview.country;
      if (previewTotal) previewTotal.textContent = result.preview.valid_rows;
      if (previewValid) previewValid.textContent = result.preview.perfect_count;
      if (previewInvalid) previewInvalid.textContent = result.preview.invalid_rows;
      if (previewEnvelope) previewEnvelope.textContent = result.preview.envelope_count;
      if (previewSpaces) previewSpaces.textContent = result.preview.spaces_count || 0;
      if (previewModified) previewModified.textContent = result.preview.modified_count || 0;
      const allTabs = document.querySelectorAll('.preview-stat-item');
      allTabs.forEach(el => el.classList.remove('active'));
      const validTab = document.querySelector('.preview-stat-item[data-type="valid"]');
      if (validTab) {
        validTab.classList.add('active');
      }
      currentPreviewType = 'valid';
      updateContextButton();
      
      // Fetch valid rows data (not sample_data which has all rows)
      // sample_data already filters valid rows from backend, so use it
      renderPreviewTable(result.preview.sample_data);

      // Render fixed shipments card (original vs cleaned)
      try {
        const fixedSample = result.preview.fixed_sample || [];
        renderFixedCard(fixedSample);
      } catch (err) {
        console.warn('No fixed sample data in preview');
        if (fixedCard) fixedCard.hidden = true;
      }

      // Render spaces found stat + card
      if (previewSpaces) previewSpaces.textContent = result.preview.spaces_count || 0;
      try {
        const spacesSample = result.preview.spaces_sample || [];
        renderSpacesCard(spacesSample);
      } catch (err) {
        console.warn('No spaces sample in preview');
      }
      
      previewCard.hidden = false;
      previewCard.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
      // Inform user about any rows removed due to blank CE Commodity Description
      let previewMsg = `Analyzed ${result.preview.total_rows} shipments.`;
      if (result.preview.blank_removed && result.preview.blank_removed > 0) {
        previewMsg += ` ${result.preview.blank_removed} rows removed (blank CE Commodity Description).`;
      }
      previewMsg += ' Please review and confirm.';
      setStatus(previewMsg, 'info');
    } else {
      setStatus('❌ ' + result.error, 'error');
    }
  } catch (err) {
    console.error('Analyze error:', err);
    setStatus('❌ Network error during analysis.', 'error');
  } finally {
    hideLoading();
    isProcessing = false;
    processBtn.disabled = false;
    btnLabel.textContent = 'Analyze File';
  }
}

processBtn.addEventListener('click', analyzeFile);

// ── Contextual Download Button ────────────────────────────────────────────────
const contextDownloadBtn = document.getElementById('context-download-btn');
const contextBtnLabel = document.getElementById('context-btn-label');

async function contextDownload() {
  if (!currentSessionId || isProcessing) return;

  const type = currentPreviewType;
  const isProcessType = type === 'spaces' || type === 'modified';

  // Step 1: Process preview first (for spaces/modified)
  if (isProcessType && !isProcessed) {
    isProcessing = true;
    showLoading(`Processing ${type === 'spaces' ? 'spaces' : 'MID words'}...`);
    try {
      const res = await fetch(apiJoin(`/api/preview/${currentSessionId}?type=${type}&process=true`));
      const result = await res.json();
      if (result.success) {
        // Update preview table with Original vs Cleaned columns
        currentPreviewType = type;
        renderProcessedPreview(result.data);
        isProcessed = true;
        updateContextButton();
        contextDownloadBtn.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
        setStatus(`✅ ${type === 'spaces' ? 'Spaces' : 'MID Words'} processed! Preview shows Original vs Cleaned.`, 'success');
      } else {
        setStatus('❌ Processing failed: ' + result.error, 'error');
      }
    } catch (err) {
      console.error('Process preview error:', err);
      setStatus('❌ Network error during processing.', 'error');
    } finally {
      hideLoading();
      isProcessing = false;
    }
    return;
  }

  // Step 2: Download (already processed OR non-process type)
  try {
    const response = await fetch(apiJoin(`/api/download_by_type/${currentSessionId}?type=${type}`));
    if (!response.ok) {
      const errText = await response.text();
      throw new Error(errText || 'Download failed');
    }

    const blob = await response.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    const disposition = response.headers.get('Content-Disposition');
    a.download = disposition
      ? disposition.split('filename=')[1]?.replace(/"/g, '') || `${type}_shipments.xlsx`
      : `${type}_shipments.xlsx`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  } catch (err) {
    console.error('Download error:', err);
    setStatus('❌ Download failed: ' + err.message, 'error');
  }
}

if (contextDownloadBtn) {
  contextDownloadBtn.addEventListener('click', contextDownload);
}

function updateContextButton() {
  if (!contextDownloadBtn || !contextBtnLabel) return;

  const isProcessType = currentPreviewType === 'spaces' || currentPreviewType === 'modified';

  if (isProcessType && !isProcessed) {
    contextBtnLabel.textContent = currentPreviewType === 'spaces' ? 'Process Spaces' : 'Process MID Words';
    contextDownloadBtn.className = 'btn btn-primary btn-process-all';
  } else {
    const labelMap = {
      total: 'Download Total Shipments',
      valid: 'Download Valid MID',
      invalid: 'Download Missing MID',
      envelope: 'Download Envelopes',
      spaces: 'Download Processed Spaces',
      modified: 'Download Processed MID Words',
    };
    contextBtnLabel.textContent = labelMap[currentPreviewType] || 'Download';
    contextDownloadBtn.className = 'btn btn-success btn-process-all';
  }
}

// ── Interactive Preview Tabs ──────────────────────────────────────────────────
function renderPreviewTable(data) {
  if (!previewTableBody) return;
  previewTableBody.innerHTML = '';

  // Update column headers depending on preview type
  try {
    const headers = document.querySelectorAll('#preview-table thead tr th');
    if (headers.length >= 3) {
      if (currentPreviewType === 'spaces') {
        headers[0].textContent = 'Tracking Number';
        headers[1].textContent = 'Manifested Description';
        headers[2].textContent = 'Spaces Detail';
      } else if (currentPreviewType === 'modified') {
        headers[0].textContent = 'Tracking Number';
        headers[1].textContent = 'Original Manifested Description';
        headers[2].textContent = 'Cleaned Manifested Description';
      } else {
        headers[0].textContent = 'Tracking Number';
        headers[1].textContent = 'Manifested Description';
        headers[2].textContent = 'CE Item HSCode';
      }
    }
  } catch (err) {
    // ignore
  }

  if (data && data.length > 0) {
    data.forEach(row => {
      const firstVal = row['Tracking Number'] || '';
      let secondVal, thirdVal;
      if (currentPreviewType === 'spaces') {
        secondVal = row['Manifested Description'] || '';
        thirdVal = row['Spaces Detail'] || '';
      } else if (currentPreviewType === 'modified') {
        secondVal = row['Original Manifested Description'] || row['_orig_manifested'] || '';
        thirdVal = row['Cleaned Manifested Description'] || row['Manifested Description'] || '';
      } else {
        secondVal = row['Manifested Description'] || '';
        thirdVal = row['CE Item HSCode'] || '';
      }
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td class="col-tracking">${firstVal}</td>
        <td class="col-desc" title="${secondVal}">${secondVal}</td>
        <td>${thirdVal}</td>
      `;
      previewTableBody.appendChild(tr);
    });
  } else {
    previewTableBody.innerHTML = '<tr><td colspan="3" style="text-align:center">No data available</td></tr>';
  }
}

function renderProcessedPreview(data) {
  if (!previewTableBody) return;
  previewTableBody.innerHTML = '';

  try {
    const headers = document.querySelectorAll('#preview-table thead tr th');
    if (headers.length >= 3) {
      headers[0].textContent = 'Tracking Number';
      headers[1].textContent = 'Original Manifested Description';
      headers[2].textContent = 'Cleaned Manifested Description';
    }
  } catch (err) {}

  if (data && data.length > 0) {
    data.forEach(row => {
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td class="col-tracking">${row['Tracking Number'] || ''}</td>
        <td class="col-desc" title="${row['Original Manifested Description'] || ''}">${row['Original Manifested Description'] || ''}</td>
        <td class="col-desc" title="${row['Cleaned Manifested Description'] || ''}">${row['Cleaned Manifested Description'] || ''}</td>
      `;
      previewTableBody.appendChild(tr);
    });
  } else {
    previewTableBody.innerHTML = '<tr><td colspan="3" style="text-align:center">No data available</td></tr>';
  }
}

async function fetchPreviewData(type) {
  if (!currentSessionId) return;
  try {
    isProcessed = false;
    const res = await fetch(apiJoin(`/api/preview/${currentSessionId}?type=${type}`));
    const result = await res.json();
    if (result.success) {
      renderPreviewTable(result.data);
    }
  } catch(e) {
    console.error(e);
  }
}

document.querySelectorAll('.preview-stat-item').forEach(el => {
  el.addEventListener('click', () => {
    document.querySelectorAll('.preview-stat-item').forEach(item => item.classList.remove('active'));
    el.classList.add('active');
    currentPreviewType = el.getAttribute('data-type');
    isProcessed = false;
    updateContextButton();
    fetchPreviewData(currentPreviewType);
    // Hide separate cards when switching preview tabs to avoid confusion
    if (noMidCard) noMidCard.hidden = true;
    if (fixedCard) fixedCard.hidden = true;
    const spacesCard = document.getElementById('spaces-found-card');
    if (spacesCard) spacesCard.hidden = true;
  });
});

// ── Download click handler (ensure it works as normal link) ──────────────────
downloadBtn.addEventListener('click', (e) => {
  if (!downloadBtn.href || downloadBtn.href === '#') {
    e.preventDefault();
    setStatus('❌ No file ready to download. Please process a file first.', 'error');
  }
});

// ── Smooth scroll for anchor links ───────────────────────────────────────────
document.querySelectorAll('a[href^="#"]').forEach(anchor => {
  anchor.addEventListener('click', (e) => {
    const target = document.querySelector(anchor.getAttribute('href'));
    if (target) {
      e.preventDefault();
      target.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }
  });
});

// ── Intersection Observer for step card animations ────────────────────────────
const observer = new IntersectionObserver((entries) => {
  entries.forEach(entry => {
    if (entry.isIntersecting) {
      entry.target.style.animationPlayState = 'running';
    }
  });
}, { threshold: 0.1 });

document.querySelectorAll('.step-card').forEach(card => {
  card.style.animationPlayState = 'paused';
  observer.observe(card);
});

// ── Paste anywhere to upload ──────────────────────────────────────────────────
document.addEventListener('paste', (e) => {
  const items = e.clipboardData?.files;
  if (items && items.length > 0) setFile(items[0]);
});
