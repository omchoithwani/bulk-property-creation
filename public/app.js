/* ── State ─────────────────────────────────────────────────────────── */
let parsedRows = [];    // parsed CSV rows
let isCreating = false; // lock to prevent double-submit

// Manage panel state
let allProperties    = [];    // raw property objects from HubSpot
let usageContext     = null;  // { workflowSet: Set, formSet: Set }
let analysisStarted  = false;
let pendingDelete    = [];    // property names queued for deletion

/* ── Token field toggle ────────────────────────────────────────────── */
function toggleToken() {
  const input = document.getElementById('token');
  input.type = input.type === 'password' ? 'text' : 'password';
}

/* ── CSV template download ─────────────────────────────────────────── */
function downloadTemplate() {
  const rows = [
    ['Name', 'Type', 'Description', 'Options'],
    ['Lead Source', 'Drop-down Select', 'How the contact discovered us', 'Website;Referral;Social Media;Email Campaign;Event'],
    ['Preferred Contact Method', 'Radio Select', 'Contact preferred communication channel', 'Phone;Email;Text Message'],
    ['Product Interests', 'Multiple Checkboxes', 'Products the contact is interested in', 'Product A;Product B;Product C;Product D'],
    ['Decision Stage', 'Drop-down Select', 'Where the contact is in the buying journey', 'Awareness;Consideration;Decision'],
  ];

  const csv = rows.map((r) => r.map(escapeCSV).join(',')).join('\r\n');
  downloadBlob(csv, 'hubspot-properties-template.csv', 'text/csv');
}

function escapeCSV(value) {
  const str = String(value);
  if (str.includes(',') || str.includes('"') || str.includes('\n')) {
    return '"' + str.replace(/"/g, '""') + '"';
  }
  return str;
}

function downloadBlob(content, filename, mimeType) {
  const blob = new Blob([content], { type: mimeType });
  const url  = URL.createObjectURL(blob);
  const a    = Object.assign(document.createElement('a'), { href: url, download: filename });
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

/* ── Drag & drop ───────────────────────────────────────────────────── */
function handleDragOver(e) {
  e.preventDefault();
  document.getElementById('dropzone').classList.add('dragging');
}

function handleDragLeave(e) {
  e.preventDefault();
  document.getElementById('dropzone').classList.remove('dragging');
}

function handleDrop(e) {
  e.preventDefault();
  document.getElementById('dropzone').classList.remove('dragging');
  const file = e.dataTransfer.files[0];
  if (file) processFile(file);
}

function handleFileSelect(e) {
  const file = e.target.files[0];
  if (file) processFile(file);
}

/* ── File processing ───────────────────────────────────────────────── */
async function processFile(file) {
  if (!file.name.endsWith('.csv')) {
    showCSVError(['Please upload a .csv file.']);
    return;
  }

  clearPreview();
  showCSVError(null);

  const formData = new FormData();
  formData.append('file', file);

  // Show a light loading indicator on the dropzone
  const dz = document.getElementById('dropzone');
  dz.querySelector('.dz-primary').textContent = 'Parsing…';

  try {
    const res  = await fetch('/api/parse-csv', { method: 'POST', body: formData });
    const data = await res.json();

    if (!data.success) {
      showCSVError(data.errors || ['Failed to parse CSV.']);
      resetDropzone();
      return;
    }

    parsedRows = data.data;
    renderPreview(parsedRows);
    updateCreateBtn();
  } catch (err) {
    showCSVError([`Network error: ${err.message}`]);
    resetDropzone();
  }
}

function resetDropzone() {
  const dz = document.getElementById('dropzone').querySelector('.dz-primary');
  dz.textContent = 'Drag & drop your CSV here';
  document.getElementById('fileInput').value = '';
}

function clearFile() {
  parsedRows = [];
  clearPreview();
  showCSVError(null);
  resetDropzone();
  updateCreateBtn();

  // Reset results area
  document.getElementById('resultsSection').style.display = 'none';
  document.getElementById('progressSection').style.display = 'none';
}

/* ── Error display ─────────────────────────────────────────────────── */
function showCSVError(errors) {
  const box = document.getElementById('csvErrors');
  if (!errors || errors.length === 0) {
    box.style.display = 'none';
    box.innerHTML = '';
    return;
  }
  box.style.display = 'block';
  if (errors.length === 1) {
    box.textContent = errors[0];
  } else {
    box.innerHTML = `<strong>${errors.length} errors found:</strong><ul>${errors.map((e) => `<li>${e}</li>`).join('')}</ul>`;
  }
}

/* ── Preview table ─────────────────────────────────────────────────── */
function renderPreview(rows) {
  const preview = document.getElementById('csvPreview');
  const count   = document.getElementById('previewCount');
  const tbody   = document.getElementById('previewBody');

  count.textContent = `${rows.length} propert${rows.length === 1 ? 'y' : 'ies'} ready to create`;
  tbody.innerHTML   = '';

  rows.forEach((row, i) => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td class="muted">${i + 1}</td>
      <td><strong>${esc(row.Name)}</strong></td>
      <td>${typeBadge(row.Type)}</td>
      <td class="${row.Description ? '' : 'muted'}">${row.Description ? esc(row.Description) : '—'}</td>
      <td class="${row.Options ? 'options-cell' : 'muted'}" title="${esc(row.Options || '')}">${row.Options ? esc(row.Options) : '—'}</td>
    `;
    tbody.appendChild(tr);
  });

  preview.style.display = 'block';
}

function clearPreview() {
  document.getElementById('csvPreview').style.display = 'none';
  document.getElementById('previewBody').innerHTML = '';
}

/* ── Create button state ───────────────────────────────────────────── */
function updateCreateBtn() {
  const btn = document.getElementById('createBtn');
  btn.disabled = parsedRows.length === 0 || isCreating;
}

/* ── Property creation ─────────────────────────────────────────────── */
async function createProperties() {
  const token      = document.getElementById('token').value.trim();
  const objectType = document.getElementById('objectType').value;

  if (!token) {
    alert('Please enter your HubSpot Private App Token before creating properties.');
    document.getElementById('token').focus();
    return;
  }

  if (parsedRows.length === 0) return;

  // Lock UI
  isCreating = true;
  updateCreateBtn();
  document.getElementById('createBtn').textContent = 'Creating…';

  // Prepare results UI
  const progressSection = document.getElementById('progressSection');
  const resultsSection  = document.getElementById('resultsSection');
  const resultsBody     = document.getElementById('resultsBody');
  const resultsSummary  = document.getElementById('resultsSummary');

  progressSection.style.display = 'block';
  resultsSection.style.display  = 'block';
  resultsBody.innerHTML         = '';
  resultsSummary.innerHTML      = '';

  // Pre-populate results table with pending rows
  parsedRows.forEach((row, i) => {
    const tr = document.createElement('tr');
    tr.id = `result-row-${i}`;
    tr.innerHTML = `
      <td class="muted">${i + 1}</td>
      <td><strong>${esc(row.Name)}</strong></td>
      <td>${typeBadge(row.Type)}</td>
      <td class="muted" id="result-iname-${i}">—</td>
      <td id="result-status-${i}"><span class="badge badge-pending">Pending</span></td>
      <td id="result-detail-${i}"></td>
    `;
    resultsBody.appendChild(tr);
  });

  let successCount = 0;
  let failCount    = 0;
  const total      = parsedRows.length;

  for (let i = 0; i < total; i++) {
    const row = parsedRows[i];

    // Mark as processing
    document.getElementById(`result-status-${i}`).innerHTML =
      '<span class="badge badge-progress">Creating…</span>';

    updateProgress(i, total, `Creating property ${i + 1} of ${total}: "${row.Name}"`);

    try {
      const res  = await fetch('/api/create-property', {
        method:  'POST',
        headers: { 'Content-Type': 'application/json' },
        body:    JSON.stringify({ token, objectType, property: row }),
      });
      const data = await res.json();

      if (data.success) {
        successCount++;
        document.getElementById(`result-status-${i}`).innerHTML =
          '<span class="badge badge-success">✓ Created</span>';
        document.getElementById(`result-iname-${i}`).innerHTML =
          `<span class="success-detail">${esc(data.internalName || '')}</span>`;
        document.getElementById(`result-detail-${i}`).innerHTML = '';
      } else {
        failCount++;
        document.getElementById(`result-status-${i}`).innerHTML =
          '<span class="badge badge-error">✗ Failed</span>';
        document.getElementById(`result-detail-${i}`).innerHTML =
          `<span class="error-detail" title="${esc(data.error || '')}">${esc(data.error || 'Unknown error')}</span>`;
      }
    } catch (err) {
      failCount++;
      document.getElementById(`result-status-${i}`).innerHTML =
        '<span class="badge badge-error">✗ Failed</span>';
      document.getElementById(`result-detail-${i}`).innerHTML =
        `<span class="error-detail" title="${esc(err.message)}">Network error: ${esc(err.message)}</span>`;
    }
  }

  // Final state
  updateProgress(total, total, `Done — ${successCount} created, ${failCount} failed`);

  resultsSummary.innerHTML = `
    <div class="summary-stat total">
      <span class="stat-number">${total}</span>
      <span class="stat-label">Total</span>
    </div>
    <div class="summary-stat success">
      <span class="stat-number">${successCount}</span>
      <span class="stat-label">Created</span>
    </div>
    <div class="summary-stat failed">
      <span class="stat-number">${failCount}</span>
      <span class="stat-label">Failed</span>
    </div>
  `;

  // Unlock UI
  isCreating = false;
  const btn = document.getElementById('createBtn');
  btn.disabled    = false;
  btn.textContent = 'Create Properties in HubSpot';
  btn.innerHTML = `
    <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none"
      stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">
      <line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/>
    </svg>
    Create Properties in HubSpot
  `;
}

/* ── Progress helpers ──────────────────────────────────────────────── */
function updateProgress(done, total, label) {
  const pct = total === 0 ? 100 : Math.round((done / total) * 100);
  document.getElementById('progressFill').style.width = pct + '%';
  document.getElementById('progressText').textContent  = label || `${done} / ${total}`;
}

/* ── Badge helpers ─────────────────────────────────────────────────── */
function typeBadge(type) {
  const map = {
    'Drop-down Select':    ['badge-dropdown', '▾ Dropdown'],
    'Radio Select':        ['badge-radio',    '◉ Radio'],
    'Multiple Checkboxes': ['badge-checkbox', '☑ Checkboxes'],
  };
  const [cls, label] = map[type] || ['badge-pending', type];
  return `<span class="badge ${cls}">${label}</span>`;
}

/* ── XSS-safe escaping ─────────────────────────────────────────────── */
function esc(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

// ════════════════════════════════════════════════════════════════════════════
// TAB NAVIGATION
// ════════════════════════════════════════════════════════════════════════════

function switchTab(tab) {
  document.getElementById('panel-create').style.display = tab === 'create' ? '' : 'none';
  document.getElementById('panel-manage').style.display = tab === 'manage' ? '' : 'none';
  document.getElementById('tab-create').classList.toggle('active', tab === 'create');
  document.getElementById('tab-manage').classList.toggle('active', tab === 'manage');
}

/** Called when object type changes — reset the manage panel state. */
function onObjectTypeChange() {
  allProperties   = [];
  usageContext    = null;
  analysisStarted = false;
  document.getElementById('propsTableCard').style.display  = 'none';
  document.getElementById('propsEmpty').style.display      = 'none';
  document.getElementById('filterBar').style.display       = 'none';
  document.getElementById('bulkBar').style.display         = 'none';
  document.getElementById('analyzeBtn').disabled           = true;
  document.getElementById('exportBtn').disabled            = true;
  document.getElementById('analyzeProgress').style.display = 'none';
  document.getElementById('analyzeWarnings').style.display = 'none';
  document.getElementById('mgmtSubtitle').textContent      = 'Load properties to get started';
}

// ════════════════════════════════════════════════════════════════════════════
// MANAGE — LOAD PROPERTIES
// ════════════════════════════════════════════════════════════════════════════

async function loadProperties() {
  const token      = document.getElementById('token').value.trim();
  const objectType = document.getElementById('objectType').value;

  if (!token) {
    alert('Please enter your HubSpot Private App Token first.');
    document.getElementById('token').focus();
    return;
  }

  // Reset state
  allProperties   = [];
  usageContext    = null;
  analysisStarted = false;

  const btn = document.getElementById('loadPropsBtn');
  btn.disabled     = true;
  btn.textContent  = 'Loading…';

  document.getElementById('propsTableCard').style.display  = 'none';
  document.getElementById('propsEmpty').style.display      = 'none';
  document.getElementById('filterBar').style.display       = 'none';
  document.getElementById('bulkBar').style.display         = 'none';
  document.getElementById('analyzeBtn').disabled           = true;
  document.getElementById('exportBtn').disabled            = true;
  document.getElementById('analyzeProgress').style.display = 'none';
  document.getElementById('analyzeWarnings').style.display = 'none';

  try {
    const res  = await fetch(`/api/list-properties?token=${encodeURIComponent(token)}&objectType=${encodeURIComponent(objectType)}`);
    const data = await res.json();

    if (!data.success) {
      alert(`Failed to load properties: ${data.error}`);
      return;
    }

    // Annotate each property with analysis placeholders
    allProperties = data.properties.map((p) => ({
      ...p,
      _recordCount:  null,  // null = not checked yet
      _inWorkflow:   null,
      _inForm:       null,
      _inList:       null,
      _inPipeline:   null,
      _inReport:     null,
      _inEmail:      null,
    }));

    renderPropertiesTable(visibleProperties());
    document.getElementById('analyzeBtn').disabled    = false;
    document.getElementById('exportBtn').disabled     = false;
    document.getElementById('filterBar').style.display = 'flex';

    const custom = allProperties.filter((p) => !p.hubspotDefined).length;
    document.getElementById('mgmtSubtitle').textContent =
      `${data.count} properties loaded (${custom} custom)`;
    document.getElementById('filterCount').textContent =
      `${data.count} shown`;
  } catch (err) {
    alert(`Network error: ${err.message}`);
  } finally {
    btn.disabled    = false;
    btn.innerHTML   = `
      <svg xmlns="http://www.w3.org/2000/svg" width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polyline points="1 4 1 10 7 10"/><path d="M3.51 15a9 9 0 1 0 .49-4.55"/></svg>
      Load Properties`;
  }
}

// ════════════════════════════════════════════════════════════════════════════
// MANAGE — RENDER TABLE
// ════════════════════════════════════════════════════════════════════════════

function renderPropertiesTable(props) {
  const card  = document.getElementById('propsTableCard');
  const empty = document.getElementById('propsEmpty');
  const tbody = document.getElementById('propsBody');

  tbody.innerHTML = '';

  if (props.length === 0) {
    card.style.display  = 'none';
    empty.style.display = 'block';
    return;
  }

  card.style.display  = 'block';
  empty.style.display = 'none';

  for (const prop of props) {
    tbody.appendChild(buildPropertyRow(prop));
  }

  updateSelectAllState();
  updateBulkBar();
}

function buildPropertyRow(prop) {
  const isSystem  = prop.hubspotDefined;
  const canDelete = !isSystem;

  const tr = document.createElement('tr');
  tr.id            = `prop-row-${prop.name}`;
  tr.dataset.name  = prop.name;
  tr.dataset.system = isSystem ? '1' : '0';

  tr.innerHTML = `
    <td class="col-cb">
      <input
        type="checkbox"
        class="prop-cb"
        data-name="${esc(prop.name)}"
        ${isSystem ? 'disabled title="System properties cannot be deleted"' : ''}
        onchange="onRowCheckChange()"
      />
    </td>
    <td>
      <span class="prop-label">${esc(prop.label || prop.name)}</span>
      <span class="prop-internal">${esc(prop.name)}</span>
    </td>
    <td>${propTypeBadge(prop.fieldType)}</td>
    <td class="muted">${esc(prop.groupName || '—')}</td>
    <td class="col-source">
      <span class="badge ${isSystem ? 'badge-system' : 'badge-custom'}">${isSystem ? 'System' : 'Custom'}</span>
    </td>
    <td class="col-usage" id="usage-records-${esc(prop.name)}">${usageCellHtml(prop._recordCount, 'records')}</td>
    <td class="col-usage" id="usage-workflow-${esc(prop.name)}">${usageCellHtml(prop._inWorkflow, 'workflow', prop.name)}</td>
    <td class="col-usage" id="usage-form-${esc(prop.name)}">${usageCellHtml(prop._inForm, 'form', prop.name)}</td>
    <td class="col-usage" id="usage-list-${esc(prop.name)}">${usageCellHtml(prop._inList, 'list', prop.name)}</td>
    <td class="col-usage" id="usage-pipeline-${esc(prop.name)}">${usageCellHtml(prop._inPipeline, 'pipeline', prop.name)}</td>
    <td class="col-usage" id="usage-report-${esc(prop.name)}">${usageCellHtml(prop._inReport, 'report', prop.name)}</td>
    <td class="col-usage" id="usage-email-${esc(prop.name)}">${usageCellHtml(prop._inEmail, 'email', prop.name)}</td>
    <td class="col-date">${formatDate(prop.updatedAt)}</td>
    <td class="col-actions">
      <button
        class="btn-icon"
        title="${canDelete ? 'Delete property' : 'System properties cannot be deleted'}"
        ${canDelete ? `onclick="deleteSingleProperty('${esc(prop.name)}')"` : 'disabled'}
      >
        <svg xmlns="http://www.w3.org/2000/svg" width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="3 6 5 6 21 6"/><path d="M19 6l-1 14a2 2 0 0 1-2 2H8a2 2 0 0 1-2-2L5 6"/><path d="M9 6V4a1 1 0 0 1 1-1h4a1 1 0 0 1 1 1v2"/></svg>
      </button>
    </td>
  `;
  return tr;
}

/** Render the content of a usage cell. */
function usageCellHtml(value, kind, propName) {
  if (value === null) return '<span class="usage-none">—</span>';
  if (kind === 'records') {
    if (value === 'loading') return '<span class="usage-spinner">⏳</span>';
    if (value === 'error')   return '<span class="usage-error">error</span>';
    const n = Number(value);
    return `<span class="usage-count ${n > 0 ? 'has-values' : 'no-values'}">${n > 0 ? n.toLocaleString() : '0'}</span>`;
  }
  // bool (workflow / form / list / pipeline / report) — clickable when in use
  if (value === true) {
    const typeMap = { workflow: 'workflows', form: 'forms', list: 'lists', pipeline: 'pipelines', report: 'reports', email: 'emails' };
    if (propName && typeMap[kind]) {
      return `<button class="usage-check-btn" title="Click to see where" onclick="showUsageDetails('${esc(propName)}', '${typeMap[kind]}')">✓</button>`;
    }
    return '<span class="usage-check" title="Used">✓</span>';
  }
  if (value === false) return '<span class="usage-cross" title="Not found">✕</span>';
  return '<span class="usage-none">—</span>';
}

/** Re-render a single usage cell without touching the rest of the row. */
function updateUsageCell(propName, kind, value) {
  // Update in allProperties array
  const prop = allProperties.find((p) => p.name === propName);
  if (prop) {
    if (kind === 'records')  prop._recordCount = value;
    if (kind === 'workflow') prop._inWorkflow  = value;
    if (kind === 'form')     prop._inForm      = value;
    if (kind === 'list')     prop._inList      = value;
    if (kind === 'pipeline') prop._inPipeline  = value;
    if (kind === 'report')   prop._inReport    = value;
    if (kind === 'email')    prop._inEmail     = value;
  }
  const cell = document.getElementById(`usage-${kind}-${propName}`);
  if (cell) cell.innerHTML = usageCellHtml(value, kind, propName);
}

/** Type badge from HubSpot fieldType */
function propTypeBadge(fieldType) {
  const map = {
    select:          ['badge-dropdown', '▾ Dropdown'],
    radio:           ['badge-radio',    '◉ Radio'],
    checkbox:        ['badge-checkbox', '☑ Checkboxes'],
    booleancheckbox: ['badge-checkbox', '☑ Checkbox'],
    text:            ['badge-pending',  'T  Text'],
    textarea:        ['badge-pending',  'T  Textarea'],
    number:          ['badge-pending',  '# Number'],
    date:            ['badge-pending',  '📅 Date'],
    file:            ['badge-pending',  '📎 File'],
  };
  const [cls, label] = map[fieldType] || ['badge-pending', fieldType || '—'];
  return `<span class="badge ${cls}">${label}</span>`;
}

// ════════════════════════════════════════════════════════════════════════════
// MANAGE — FILTER / SEARCH
// ════════════════════════════════════════════════════════════════════════════

function visibleProperties() {
  const search      = (document.getElementById('searchProps')?.value || '').toLowerCase();
  const filterSrc   = document.getElementById('filterSource')?.value || 'all';
  const filterUsage = document.getElementById('filterUsage')?.value  || 'all';

  return allProperties.filter((p) => {
    // Text search
    if (search) {
      const haystack = `${p.label} ${p.name} ${p.groupName}`.toLowerCase();
      if (!haystack.includes(search)) return false;
    }
    // Source filter
    if (filterSrc === 'custom' && p.hubspotDefined) return false;
    if (filterSrc === 'system' && !p.hubspotDefined) return false;
    // Usage filter (only meaningful after analysis)
    if (filterUsage === 'unused') {
      if (!analysisStarted) return !p.hubspotDefined; // show custom before analysis
      const noRecords  = p._recordCount === null || Number(p._recordCount) === 0;
      const noWorkflow = p._inWorkflow  === false || p._inWorkflow  === null;
      const noForm     = p._inForm      === false || p._inForm      === null;
      const noList     = p._inList      === false || p._inList      === null;
      const noPipeline = p._inPipeline  === false || p._inPipeline  === null;
      const noReport   = p._inReport    === false || p._inReport    === null;
      const noEmail    = p._inEmail     === false || p._inEmail     === null;
      return !p.hubspotDefined && noRecords && noWorkflow && noForm && noList && noPipeline && noReport && noEmail;
    }
    if (filterUsage === 'used') {
      const hasRecords  = Number(p._recordCount) > 0;
      const inWorkflow  = p._inWorkflow === true;
      const inForm      = p._inForm     === true;
      const inList      = p._inList     === true;
      const inPipeline  = p._inPipeline === true;
      const inReport    = p._inReport   === true;
      const inEmail     = p._inEmail    === true;
      return hasRecords || inWorkflow || inForm || inList || inPipeline || inReport || inEmail;
    }
    return true;
  });
}

function filterProperties() {
  const vis = visibleProperties();
  renderPropertiesTable(vis);
  document.getElementById('filterCount').textContent = `${vis.length} of ${allProperties.length} shown`;
}

// ════════════════════════════════════════════════════════════════════════════
// MANAGE — ANALYZE USAGE
// ════════════════════════════════════════════════════════════════════════════

async function analyzeUsage() {
  const token      = document.getElementById('token').value.trim();
  const objectType = document.getElementById('objectType').value;

  if (!token) { alert('Please enter your HubSpot Private App Token first.'); return; }
  if (!allProperties.length) { alert('Load properties first.'); return; }

  analysisStarted = true;

  const analyzeBtn  = document.getElementById('analyzeBtn');
  analyzeBtn.disabled = true;

  const progressEl  = document.getElementById('analyzeProgress');
  const fillEl      = document.getElementById('analyzeProgressFill');
  const textEl      = document.getElementById('analyzeProgressText');
  const warningsEl  = document.getElementById('analyzeWarnings');

  progressEl.style.display  = 'block';
  warningsEl.style.display  = 'none';
  fillEl.style.width = '0%';
  textEl.textContent = 'Fetching workflows and forms…';

  // ── Step 1: fetch workflow + form context ──
  let ctx;
  try {
    const res  = await fetch('/api/fetch-usage-context', {
      method:  'POST',
      headers: { 'Content-Type': 'application/json' },
      body:    JSON.stringify({ token }),
    });
    ctx = await res.json();
  } catch (err) {
    textEl.textContent = `Network error: ${err.message}`;
    analyzeBtn.disabled = false;
    return;
  }

  if (ctx.warnings && ctx.warnings.length) {
    warningsEl.style.display = 'block';
    warningsEl.innerHTML =
      `<strong>Some checks were skipped (missing API scopes):</strong><ul>` +
      ctx.warnings.map((w) => `<li>${esc(w)}</li>`).join('') +
      `</ul>`;
  }

  const wfSet       = new Set(ctx.workflowProperties  || []);
  const formSet     = new Set(ctx.formProperties       || []);
  const listSet     = new Set(ctx.listProperties       || []);
  const pipelineSet = new Set(ctx.pipelineProperties   || []);
  const reportSet   = new Set(ctx.reportProperties     || []);
  const emailSet    = new Set(ctx.emailProperties      || []);
  usageContext  = { wfSet, formSet, listSet, pipelineSet, reportSet, emailSet, usageDetails: ctx.propertyUsageDetails || {} };

  textEl.textContent =
    `Found ${ctx.workflowCount} workflows, ${ctx.formCount} forms, ` +
    `${ctx.listCount} lists, ${ctx.pipelineCount} pipelines, ${ctx.reportCount} reports, ` +
    `${ctx.emailCount} emails. Checking records…`;
  fillEl.style.width = '10%';

  // ── Step 2: update all non-record columns (in-memory, instant) ──
  for (const prop of allProperties) {
    updateUsageCell(prop.name, 'workflow',  wfSet.has(prop.name));
    updateUsageCell(prop.name, 'form',      formSet.has(prop.name));
    updateUsageCell(prop.name, 'list',      listSet.has(prop.name));
    updateUsageCell(prop.name, 'pipeline',  pipelineSet.has(prop.name));
    updateUsageCell(prop.name, 'report',    reportSet.has(prop.name));
    updateUsageCell(prop.name, 'email',     emailSet.has(prop.name));
  }

  // ── Step 3: check record counts for custom properties ──
  const customProps = allProperties.filter((p) => !p.hubspotDefined);
  const total       = customProps.length;

  if (total === 0) {
    fillEl.style.width = '100%';
    textEl.textContent = 'Analysis complete — no custom properties to check.';
    analyzeBtn.disabled = false;
    filterProperties();
    return;
  }

  textEl.textContent = `Checking record values for ${total} custom properties…`;

  for (let i = 0; i < total; i++) {
    const prop = customProps[i];
    const pct  = Math.round(10 + ((i / total) * 88));
    fillEl.style.width = pct + '%';
    textEl.textContent = `Checking records: ${i + 1} / ${total} — "${prop.label || prop.name}"`;

    // Mark as loading
    updateUsageCell(prop.name, 'records', 'loading');

    try {
      const res  = await fetch('/api/check-property-records', {
        method:  'POST',
        headers: { 'Content-Type': 'application/json' },
        body:    JSON.stringify({ token, objectType, propertyName: prop.name }),
      });
      const data = await res.json();
      updateUsageCell(prop.name, 'records', data.success ? data.total : 'error');
    } catch {
      updateUsageCell(prop.name, 'records', 'error');
    }

    // Small delay to respect HubSpot rate limits (~10 req/s safe)
    if (i < total - 1) await sleep(120);
  }

  fillEl.style.width = '100%';

  const unusedCount = allProperties.filter((p) =>
    !p.hubspotDefined &&
    Number(p._recordCount) === 0 &&
    p._inWorkflow === false &&
    p._inForm     === false &&
    p._inList     === false &&
    p._inPipeline === false &&
    p._inReport   === false &&
    p._inEmail    === false
  ).length;

  textEl.textContent =
    `Analysis complete — ${unusedCount} unused custom propert${unusedCount === 1 ? 'y' : 'ies'} found.`;

  analyzeBtn.disabled = false;
  filterProperties(); // re-render with updated data
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

// ════════════════════════════════════════════════════════════════════════════
// MANAGE — SELECTION
// ════════════════════════════════════════════════════════════════════════════

function toggleSelectAll() {
  const master   = document.getElementById('selectAllCb');
  const checkboxes = document.querySelectorAll('.prop-cb:not([disabled])');
  checkboxes.forEach((cb) => { cb.checked = master.checked; });
  updateBulkBar();
}

function onRowCheckChange() {
  updateSelectAllState();
  updateBulkBar();
}

function updateSelectAllState() {
  const all      = document.querySelectorAll('.prop-cb:not([disabled])');
  const checked  = document.querySelectorAll('.prop-cb:not([disabled]):checked');
  const master   = document.getElementById('selectAllCb');
  master.checked       = all.length > 0 && checked.length === all.length;
  master.indeterminate = checked.length > 0 && checked.length < all.length;
}

function updateBulkBar() {
  const checked = document.querySelectorAll('.prop-cb:checked');
  const bar     = document.getElementById('bulkBar');
  const count   = document.getElementById('bulkCount');
  if (checked.length === 0) {
    bar.style.display = 'none';
  } else {
    bar.style.display = 'flex';
    count.textContent = `${checked.length} propert${checked.length === 1 ? 'y' : 'ies'} selected`;
  }
}

function clearSelection() {
  document.querySelectorAll('.prop-cb:checked').forEach((cb) => { cb.checked = false; });
  document.getElementById('selectAllCb').checked       = false;
  document.getElementById('selectAllCb').indeterminate = false;
  updateBulkBar();
}

/** Select all visible custom properties that appear to be unused. */
function selectUnused() {
  const vis = visibleProperties();
  let count = 0;
  for (const prop of vis) {
    if (prop.hubspotDefined) continue;
    const noRecords  = prop._recordCount === null || Number(prop._recordCount) === 0;
    const noWorkflow = prop._inWorkflow  !== true;
    const noForm     = prop._inForm      !== true;
    const noList     = prop._inList      !== true;
    const noPipeline = prop._inPipeline  !== true;
    const noReport   = prop._inReport    !== true;
    const noEmail    = prop._inEmail     !== true;
    const isUnused   = noRecords && noWorkflow && noForm && noList && noPipeline && noReport && noEmail;
    const cb = document.querySelector(`.prop-cb[data-name="${CSS.escape(prop.name)}"]`);
    if (cb && isUnused) { cb.checked = true; count++; }
  }
  updateSelectAllState();
  updateBulkBar();
  if (count === 0) {
    alert(analysisStarted
      ? 'No unused properties found among the visible rows.'
      : 'Run "Analyze Usage" first to identify unused properties.');
  }
}

// ════════════════════════════════════════════════════════════════════════════
// MANAGE — DELETE
// ════════════════════════════════════════════════════════════════════════════

function deleteSelected() {
  const checked = [...document.querySelectorAll('.prop-cb:checked')];
  if (!checked.length) return;
  pendingDelete = checked.map((cb) => cb.dataset.name);
  showDeleteModal(pendingDelete);
}

function deleteSingleProperty(propName) {
  pendingDelete = [propName];
  showDeleteModal(pendingDelete);
}

function showDeleteModal(names) {
  const modal    = document.getElementById('deleteModal');
  const bodyEl   = document.getElementById('deleteModalBody');
  const n        = names.length;
  const examples = names.slice(0, 5).map((n) => `<strong>${esc(n)}</strong>`).join(', ');
  bodyEl.innerHTML =
    `You are about to permanently delete ${n} propert${n === 1 ? 'y' : 'ies'}` +
    (n <= 5 ? `: ${examples}` : ` including ${examples} and ${n - 5} more`) +
    `.<br/><br/>This action <strong>cannot be undone</strong>. All data stored in ${n === 1 ? 'this property' : 'these properties'} will be removed from every record.`;
  modal.style.display = 'flex';
}

function closeDeleteModal(e) {
  if (e && e.target !== document.getElementById('deleteModal')) return;
  document.getElementById('deleteModal').style.display = 'none';
}

async function confirmDelete() {
  document.getElementById('deleteModal').style.display = 'none';

  const token      = document.getElementById('token').value.trim();
  const objectType = document.getElementById('objectType').value;
  const names      = pendingDelete;
  pendingDelete    = [];

  const analyzeProgress = document.getElementById('analyzeProgress');
  const fillEl          = document.getElementById('analyzeProgressFill');
  const textEl          = document.getElementById('analyzeProgressText');

  analyzeProgress.style.display = 'block';
  fillEl.style.width = '0%';

  let deleted = 0, failed = 0;

  for (let i = 0; i < names.length; i++) {
    const propName = names[i];
    fillEl.style.width = `${Math.round((i / names.length) * 100)}%`;
    textEl.textContent = `Deleting ${i + 1} / ${names.length}: "${propName}"…`;

    try {
      const res  = await fetch('/api/delete-property', {
        method:  'DELETE',
        headers: { 'Content-Type': 'application/json' },
        body:    JSON.stringify({ token, objectType, propertyName: propName }),
      });
      const data = await res.json();
      if (data.success) {
        deleted++;
        // Visually mark the row as deleted then fade it out
        const row = document.getElementById(`prop-row-${propName}`);
        if (row) {
          row.classList.add('row-deleted');
          setTimeout(() => row.remove(), 1200);
        }
        // Remove from allProperties array
        allProperties = allProperties.filter((p) => p.name !== propName);
      } else {
        failed++;
        console.warn(`Failed to delete ${propName}: ${data.error}`);
      }
    } catch (err) {
      failed++;
      console.warn(`Network error deleting ${propName}: ${err.message}`);
    }
  }

  fillEl.style.width = '100%';
  textEl.textContent = `Deleted ${deleted} propert${deleted === 1 ? 'y' : 'ies'}` +
    (failed > 0 ? `, ${failed} failed (check console)` : '.');

  // Update subtitle and filter count
  const custom = allProperties.filter((p) => !p.hubspotDefined).length;
  document.getElementById('mgmtSubtitle').textContent =
    `${allProperties.length} properties (${custom} custom)`;

  clearSelection();
  filterProperties();
}

// ════════════════════════════════════════════════════════════════════════════
// USAGE DETAIL MODAL
// ════════════════════════════════════════════════════════════════════════════

function showUsageDetails(propName, type) {
  const details = usageContext?.usageDetails?.[propName]?.[type] || [];
  const prop = allProperties.find((p) => p.name === propName);
  const label = prop?.label || propName;
  const typeLabels = { workflows: 'Workflows', forms: 'Forms', lists: 'Lists', pipelines: 'Pipelines', reports: 'Reports', emails: 'Marketing Emails' };
  const typeLabel = typeLabels[type] || type;

  document.getElementById('usageDetailTitle').textContent =
    `"${label}" — Used in ${typeLabel}`;

  const list = document.getElementById('usageDetailList');
  if (details.length === 0) {
    list.innerHTML = '<li class="usage-detail-none">No names available</li>';
  } else {
    list.innerHTML = details
      .map((name) => `<li>${esc(name)}</li>`)
      .join('');
  }

  document.getElementById('usageDetailModal').style.display = 'flex';
}

function closeUsageDetailModal(e) {
  if (e && e.target !== document.getElementById('usageDetailModal')) return;
  document.getElementById('usageDetailModal').style.display = 'none';
}

// ════════════════════════════════════════════════════════════════════════════
// HELPERS
// ════════════════════════════════════════════════════════════════════════════

function formatDate(dateStr) {
  if (!dateStr) return '—';
  try {
    return new Date(dateStr).toLocaleDateString(undefined, { year: 'numeric', month: 'short', day: 'numeric' });
  } catch {
    return '—';
  }
}

// ════════════════════════════════════════════════════════════════════════════
// EXPORT CSV
// ════════════════════════════════════════════════════════════════════════════

function exportCSV() {
  const vis = visibleProperties();
  const objectType = document.getElementById('objectType').value;

  const headers = [
    'Label', 'Internal Name', 'Type', 'Group', 'Source',
    'Records', 'In Workflow', 'In Form', 'In List', 'In Pipeline', 'In Report', 'In Email',
    'Last Modified',
  ];

  const boolCell = (v) => v === true ? 'Yes' : v === false ? 'No' : '';

  const rows = vis.map((p) => [
    p.label || p.name,
    p.name,
    p.fieldType || '',
    p.groupName || '',
    p.hubspotDefined ? 'System' : 'Custom',
    p._recordCount !== null && p._recordCount !== 'loading' && p._recordCount !== 'error'
      ? p._recordCount
      : '',
    boolCell(p._inWorkflow),
    boolCell(p._inForm),
    boolCell(p._inList),
    boolCell(p._inPipeline),
    boolCell(p._inReport),
    boolCell(p._inEmail),
    p.updatedAt ? new Date(p.updatedAt).toISOString().slice(0, 10) : '',
  ]);

  const csv = [headers, ...rows].map((r) => r.map(escapeCSV).join(',')).join('\r\n');
  downloadBlob(csv, `${objectType}-properties.csv`, 'text/csv');
}
