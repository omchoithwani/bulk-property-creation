/* ── State ─────────────────────────────────────────────────────────── */
let parsedRows = [];
let isCreating = false;
let allProperties    = [];
let usageContext     = null;
let analysisStarted  = false;
let pendingDelete    = [];

/* ── Auth init ─────────────────────────────────────────────────────── */
async function initAuth() {
  try {
    // Check for OAuth error redirect
    const params = new URLSearchParams(window.location.search);
    const authError = params.get('auth_error');
    if (authError) {
      window.history.replaceState({}, '', '/');
      showAuthConnect(authError);
      return;
    }

    const res  = await fetch('/oauth/status');
    const data = await res.json();

    if (data.connected) {
      showAuthConnected(data.portalId, data.hubDomain);
      await loadObjectTypes();
    } else {
      showAuthConnect(null);
    }
  } catch {
    showAuthConnect('Could not reach server.');
  }
}

function showAuthConnect(errorMsg) {
  document.getElementById('auth-loading').style.display   = 'none';
  document.getElementById('auth-connected').style.display = 'none';
  document.getElementById('auth-connect').style.display   = '';
  document.getElementById('tab-nav').style.display        = 'none';
  document.getElementById('panel-create').style.display   = 'none';
  document.getElementById('panel-manage').style.display   = 'none';

  document.getElementById('connection-subtitle').textContent =
    'Connect your HubSpot account to get started';

  const errEl = document.getElementById('auth-error');
  if (errorMsg) {
    errEl.textContent = errorMsg;
    errEl.style.display = '';
  } else {
    errEl.style.display = 'none';
  }

  document.getElementById('header-auth-status').innerHTML = '';
}

function showAuthConnected(portalId, hubDomain) {
  document.getElementById('auth-loading').style.display   = 'none';
  document.getElementById('auth-connect').style.display   = 'none';
  document.getElementById('auth-connected').style.display = '';
  document.getElementById('tab-nav').style.display        = '';
  document.getElementById('panel-create').style.display   = '';
  document.getElementById('panel-manage').style.display   = 'none';

  document.getElementById('portal-id').textContent  = portalId  || '—';
  document.getElementById('hub-domain').textContent = hubDomain || '—';
  document.getElementById('connection-subtitle').textContent = 'Connected to HubSpot';

  document.getElementById('header-auth-status').innerHTML =
    `<span class="conn-dot"></span><span>Portal ${esc(String(portalId || '—'))}</span>`;
}

async function disconnectHubSpot() {
  await fetch('/oauth/disconnect', { method: 'POST' });
  showAuthConnect(null);
}

/* ── Handle 401 from any API call ──────────────────────────────────── */
function handleUnauth() {
  showAuthConnect('Your session expired. Please reconnect.');
}

/* ── CSV template download ─────────────────────────────────────────── */
function downloadTemplate() {
  const rows = [
    ['Name', 'Type', 'Description', 'Options'],
    ['Lead Source',          'Drop-down Select',     'How the contact discovered us',           'Website;Referral;Social Media;Email Campaign;Event'],
    ['Preferred Contact',    'Radio Select',          'Preferred communication channel',         'Phone;Email;Text Message'],
    ['Product Interests',    'Multiple Checkboxes',   'Products the contact is interested in',   'Product A;Product B;Product C'],
    ['Job Title',            'Single Line Text',      'Contact job title',                       ''],
    ['Notes',                'Multi-line Text',       'Additional notes about the contact',      ''],
    ['Office Phone',         'Phone Number',          'Primary office phone number',             ''],
    ['Website',              'URL',                   'Company or personal website URL',         ''],
    ['Bio',                  'Rich Text',             'Formatted biography or description',      ''],
    ['Annual Revenue',       'Number',                'Annual revenue amount in USD',            ''],
    ['Contract Date',        'Date Picker',           'Date the contract was signed',            ''],
    ['Meeting Scheduled',    'Date and Time Picker',  'Date and time of the next meeting',       ''],
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
function handleDragOver(e)  { e.preventDefault(); document.getElementById('dropzone').classList.add('dragging'); }
function handleDragLeave(e) { e.preventDefault(); document.getElementById('dropzone').classList.remove('dragging'); }
function handleDrop(e) {
  e.preventDefault();
  document.getElementById('dropzone').classList.remove('dragging');
  const file = e.dataTransfer.files[0];
  if (file) processFile(file);
}
function handleFileSelect(e) { const file = e.target.files[0]; if (file) processFile(file); }

/* ── File processing ───────────────────────────────────────────────── */
async function processFile(file) {
  if (!file.name.endsWith('.csv')) { showCSVError(['Please upload a .csv file.']); return; }
  clearPreview();
  showCSVError(null);
  const formData = new FormData();
  formData.append('file', file);
  document.getElementById('dropzone').querySelector('.dz-primary').textContent = 'Parsing…';
  try {
    const res  = await fetch('/api/parse-csv', { method: 'POST', body: formData });
    const data = await res.json();
    if (!data.success) { showCSVError(data.errors || ['Failed to parse CSV.']); resetDropzone(); return; }
    parsedRows = data.data;
    renderPreview(parsedRows);
    updateCreateBtn();
  } catch (err) {
    showCSVError([`Network error: ${err.message}`]);
    resetDropzone();
  }
}

function resetDropzone() {
  document.getElementById('dropzone').querySelector('.dz-primary').textContent = 'Drag & drop your CSV here';
  document.getElementById('fileInput').value = '';
}

function clearFile() {
  parsedRows = [];
  clearPreview();
  showCSVError(null);
  resetDropzone();
  updateCreateBtn();
  document.getElementById('resultsSection').style.display  = 'none';
  document.getElementById('progressSection').style.display = 'none';
}

/* ── Error display ─────────────────────────────────────────────────── */
function showCSVError(errors) {
  const box = document.getElementById('csvErrors');
  if (!errors || errors.length === 0) { box.style.display = 'none'; box.innerHTML = ''; return; }
  box.style.display = 'block';
  box.innerHTML = errors.length === 1
    ? errors[0]
    : `<strong>${errors.length} errors found:</strong><ul>${errors.map((e) => `<li>${e}</li>`).join('')}</ul>`;
}

/* ── Preview table ─────────────────────────────────────────────────── */
function renderPreview(rows) {
  document.getElementById('previewCount').textContent = `${rows.length} propert${rows.length === 1 ? 'y' : 'ies'} ready to create`;
  const tbody = document.getElementById('previewBody');
  tbody.innerHTML = '';
  rows.forEach((row, i) => {
    const tr = document.createElement('tr');
    tr.innerHTML = `<td class="muted">${i + 1}</td><td><strong>${esc(row.Name)}</strong></td><td>${typeBadge(row.Type)}</td><td class="${row.Description ? '' : 'muted'}">${row.Description ? esc(row.Description) : '—'}</td><td class="${row.Options ? 'options-cell' : 'muted'}" title="${esc(row.Options || '')}">${row.Options ? esc(row.Options) : '—'}</td>`;
    tbody.appendChild(tr);
  });
  document.getElementById('csvPreview').style.display = 'block';
}

function clearPreview() {
  document.getElementById('csvPreview').style.display = 'none';
  document.getElementById('previewBody').innerHTML = '';
}

/* ── Create button state ───────────────────────────────────────────── */
function updateCreateBtn() {
  document.getElementById('createBtn').disabled = parsedRows.length === 0 || isCreating;
}

/* ── Property creation ─────────────────────────────────────────────── */
async function createProperties() {
  await loadObjectTypes();
  const objectType = getObjectType();
  if (parsedRows.length === 0) return;

  isCreating = true;
  updateCreateBtn();
  document.getElementById('createBtn').textContent = 'Creating…';

  const progressSection = document.getElementById('progressSection');
  const resultsSection  = document.getElementById('resultsSection');
  const resultsBody     = document.getElementById('resultsBody');
  const resultsSummary  = document.getElementById('resultsSummary');

  progressSection.style.display = 'block';
  resultsSection.style.display  = 'block';
  resultsBody.innerHTML         = '';
  resultsSummary.innerHTML      = '';

  parsedRows.forEach((row, i) => {
    const tr = document.createElement('tr');
    tr.id = `result-row-${i}`;
    tr.innerHTML = `<td class="muted">${i + 1}</td><td><strong>${esc(row.Name)}</strong></td><td>${typeBadge(row.Type)}</td><td class="muted" id="result-iname-${i}">—</td><td id="result-status-${i}"><span class="badge badge-pending">Pending</span></td><td id="result-detail-${i}"></td>`;
    resultsBody.appendChild(tr);
  });

  let successCount = 0, failCount = 0;
  const total = parsedRows.length;

  for (let i = 0; i < total; i++) {
    const row = parsedRows[i];
    document.getElementById(`result-status-${i}`).innerHTML = '<span class="badge badge-progress">Creating…</span>';
    updateProgress(i, total, `Creating property ${i + 1} of ${total}: "${row.Name}"`);
    try {
      const res  = await fetch('/api/create-property', {
        method:  'POST',
        headers: { 'Content-Type': 'application/json' },
        body:    JSON.stringify({ objectType, property: row, defaultGroup: objectTypeDefaultGroups[objectType] || null }),
      });
      if (res.status === 401) { handleUnauth(); return; }
      const data = await res.json();
      if (data.success) {
        successCount++;
        document.getElementById(`result-status-${i}`).innerHTML = '<span class="badge badge-success">✓ Created</span>';
        document.getElementById(`result-iname-${i}`).innerHTML  = `<span class="success-detail">${esc(data.internalName || '')}</span>`;
        document.getElementById(`result-detail-${i}`).innerHTML = '';
      } else {
        failCount++;
        document.getElementById(`result-status-${i}`).innerHTML  = '<span class="badge badge-error">✗ Failed</span>';
        document.getElementById(`result-detail-${i}`).innerHTML  = `<span class="error-detail" title="${esc(data.error || '')}">${esc(data.error || 'Unknown error')}</span>`;
      }
    } catch (err) {
      failCount++;
      document.getElementById(`result-status-${i}`).innerHTML  = '<span class="badge badge-error">✗ Failed</span>';
      document.getElementById(`result-detail-${i}`).innerHTML  = `<span class="error-detail">Network error: ${esc(err.message)}</span>`;
    }
  }

  updateProgress(total, total, `Done — ${successCount} created, ${failCount} failed`);
  resultsSummary.innerHTML = `
    <div class="summary-stat total"><span class="stat-number">${total}</span><span class="stat-label">Total</span></div>
    <div class="summary-stat success"><span class="stat-number">${successCount}</span><span class="stat-label">Created</span></div>
    <div class="summary-stat failed"><span class="stat-number">${failCount}</span><span class="stat-label">Failed</span></div>`;

  isCreating = false;
  const btn = document.getElementById('createBtn');
  btn.disabled = false;
  btn.innerHTML = `<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg> Create Properties in HubSpot`;
}

/* ── Progress helpers ──────────────────────────────────────────────── */
function updateProgress(done, total, label) {
  const pct = total === 0 ? 100 : Math.round((done / total) * 100);
  document.getElementById('progressFill').style.width   = pct + '%';
  document.getElementById('progressText').textContent   = label || `${done} / ${total}`;
}

/* ── Badge helpers ─────────────────────────────────────────────────── */
function typeBadge(type) {
  const map = {
    'Drop-down Select':     ['badge-dropdown', '▾ Dropdown'],
    'Radio Select':         ['badge-radio',    '◉ Radio'],
    'Multiple Checkboxes':  ['badge-checkbox', '☑ Checkboxes'],
    'Single Line Text':     ['badge-text',     'T  Text'],
    'Multi-line Text':      ['badge-text',     '¶  Textarea'],
    'Phone Number':         ['badge-text',     '☏ Phone'],
    'URL':                  ['badge-text',     '⌁ URL'],
    'Rich Text':            ['badge-text',     '⟨⟩ Rich Text'],
    'Number':               ['badge-number',   '# Number'],
    'Date Picker':          ['badge-date',     '📅 Date'],
    'Date and Time Picker': ['badge-date',     '⊙ DateTime'],
  };
  const [cls, label] = map[type] || ['badge-pending', type];
  return `<span class="badge ${cls}">${label}</span>`;
}

/* ── XSS-safe escaping ─────────────────────────────────────────────── */
function esc(str) {
  return String(str).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

/* ── Tab navigation ────────────────────────────────────────────────── */
function switchTab(tab) {
  document.getElementById('panel-create').style.display = tab === 'create' ? '' : 'none';
  document.getElementById('panel-manage').style.display = tab === 'manage' ? '' : 'none';
  document.getElementById('tab-create').classList.toggle('active', tab === 'create');
  document.getElementById('tab-manage').classList.toggle('active', tab === 'manage');
}

function onObjectTypeChange() {
  allProperties = []; usageContext = null; analysisStarted = false;
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

/* ── Object types ──────────────────────────────────────────────────── */
const objectTypeDefaultGroups = {};

function getObjectType() {
  const manual = (document.getElementById('objectTypeManual')?.value || '').trim();
  return manual || document.getElementById('objectType').value;
}

async function loadObjectTypes() {
  const select  = document.getElementById('objectType');
  const current = select.value;
  try {
    const res  = await fetch('/api/list-object-types');
    if (res.status === 401) { handleUnauth(); return; }
    const data = await res.json();
    if (!data.success || !data.objectTypes?.length) return;
    data.objectTypes.forEach(ot => { if (ot.defaultGroup) objectTypeDefaultGroups[ot.value] = ot.defaultGroup; });
    select.innerHTML = data.objectTypes
      .map(ot => `<option value="${esc(ot.value)}"${ot.value === current ? ' selected' : ''}>${esc(ot.label)}</option>`)
      .join('');
    if ([...select.options].some(o => o.value === current)) select.value = current;
  } catch { /* silent — HTML defaults remain */ }
}

/* ── Load properties ───────────────────────────────────────────────── */
async function loadProperties() {
  await loadObjectTypes();
  const objectType = getObjectType();

  allProperties = []; usageContext = null; analysisStarted = false;

  const btn = document.getElementById('loadPropsBtn');
  btn.disabled    = true;
  btn.textContent = 'Loading…';

  document.getElementById('propsTableCard').style.display  = 'none';
  document.getElementById('propsEmpty').style.display      = 'none';
  document.getElementById('filterBar').style.display       = 'none';
  document.getElementById('bulkBar').style.display         = 'none';
  document.getElementById('analyzeBtn').disabled           = true;
  document.getElementById('exportBtn').disabled            = true;
  document.getElementById('analyzeProgress').style.display = 'none';
  document.getElementById('analyzeWarnings').style.display = 'none';

  try {
    const res  = await fetch(`/api/list-properties?objectType=${encodeURIComponent(objectType)}`);
    if (res.status === 401) { handleUnauth(); return; }
    const data = await res.json();
    if (!data.success) { alert(`Failed to load properties: ${data.error}`); return; }

    allProperties = data.properties.map((p) => ({
      ...p, _recordCount: null, _inWorkflow: null, _inForm: null,
      _inList: null, _inPipeline: null, _inReport: null, _inEmail: null,
    }));

    renderPropertiesTable(visibleProperties());
    document.getElementById('analyzeBtn').disabled    = false;
    document.getElementById('exportBtn').disabled     = false;
    document.getElementById('filterBar').style.display = 'flex';
    const custom = allProperties.filter((p) => !p.hubspotDefined).length;
    document.getElementById('mgmtSubtitle').textContent = `${data.count} properties loaded (${custom} custom)`;
    document.getElementById('filterCount').textContent  = `${data.count} shown`;
  } catch (err) {
    alert(`Network error: ${err.message}`);
  } finally {
    btn.disabled  = false;
    btn.innerHTML = `<svg xmlns="http://www.w3.org/2000/svg" width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polyline points="1 4 1 10 7 10"/><path d="M3.51 15a9 9 0 1 0 .49-4.55"/></svg> Load Properties`;
  }
}

/* ── Render table ──────────────────────────────────────────────────── */
function renderPropertiesTable(props) {
  const card  = document.getElementById('propsTableCard');
  const empty = document.getElementById('propsEmpty');
  const tbody = document.getElementById('propsBody');
  tbody.innerHTML = '';
  if (props.length === 0) { card.style.display = 'none'; empty.style.display = 'block'; return; }
  card.style.display = 'block'; empty.style.display = 'none';
  for (const prop of props) tbody.appendChild(buildPropertyRow(prop));
  updateSelectAllState();
  updateBulkBar();
}

function buildPropertyRow(prop) {
  const isSystem  = prop.hubspotDefined;
  const canDelete = !isSystem;
  const tr = document.createElement('tr');
  tr.id = `prop-row-${prop.name}`;
  tr.dataset.name   = prop.name;
  tr.dataset.system = isSystem ? '1' : '0';
  tr.innerHTML = `
    <td class="col-cb"><input type="checkbox" class="prop-cb" data-name="${esc(prop.name)}" ${isSystem ? 'disabled title="System properties cannot be deleted"' : ''} onchange="onRowCheckChange()"/></td>
    <td><span class="prop-label">${esc(prop.label || prop.name)}</span><span class="prop-internal">${esc(prop.name)}</span></td>
    <td>${propTypeBadge(prop.fieldType, prop.type)}</td>
    <td class="muted">${esc(prop.groupName || '—')}</td>
    <td class="col-source"><span class="badge ${isSystem ? 'badge-system' : 'badge-custom'}">${isSystem ? 'System' : 'Custom'}</span></td>
    <td class="col-usage" id="usage-records-${esc(prop.name)}">${usageCellHtml(prop._recordCount, 'records')}</td>
    <td class="col-usage" id="usage-workflow-${esc(prop.name)}">${usageCellHtml(prop._inWorkflow, 'workflow', prop.name)}</td>
    <td class="col-usage" id="usage-form-${esc(prop.name)}">${usageCellHtml(prop._inForm, 'form', prop.name)}</td>
    <td class="col-usage" id="usage-list-${esc(prop.name)}">${usageCellHtml(prop._inList, 'list', prop.name)}</td>
    <td class="col-usage" id="usage-pipeline-${esc(prop.name)}">${usageCellHtml(prop._inPipeline, 'pipeline', prop.name)}</td>
    <td class="col-usage" id="usage-report-${esc(prop.name)}">${usageCellHtml(prop._inReport, 'report', prop.name)}</td>
    <td class="col-usage" id="usage-email-${esc(prop.name)}">${usageCellHtml(prop._inEmail, 'email', prop.name)}</td>
    <td class="col-date">${formatDate(prop.updatedAt)}</td>
    <td class="col-actions"><button class="btn-icon" title="${canDelete ? 'Delete property' : 'System properties cannot be deleted'}" ${canDelete ? `onclick="deleteSingleProperty('${esc(prop.name)}')"` : 'disabled'}><svg xmlns="http://www.w3.org/2000/svg" width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="3 6 5 6 21 6"/><path d="M19 6l-1 14a2 2 0 0 1-2 2H8a2 2 0 0 1-2-2L5 6"/><path d="M9 6V4a1 1 0 0 1 1-1h4a1 1 0 0 1 1 1v2"/></svg></button></td>`;
  return tr;
}

function usageCellHtml(value, kind, propName) {
  if (value === null) return '<span class="usage-none">—</span>';
  if (kind === 'records') {
    if (value === 'loading') return '<span class="usage-spinner">⏳</span>';
    if (value === 'error')   return '<span class="usage-error">error</span>';
    const n = Number(value);
    return `<span class="usage-count ${n > 0 ? 'has-values' : 'no-values'}">${n > 0 ? n.toLocaleString() : '0'}</span>`;
  }
  if (value === true) {
    const typeMap = { workflow: 'workflows', form: 'forms', list: 'lists', pipeline: 'pipelines', report: 'reports', email: 'emails' };
    if (propName && typeMap[kind]) return `<button class="usage-check-btn" title="Click to see where" onclick="showUsageDetails('${esc(propName)}', '${typeMap[kind]}')">✓</button>`;
    return '<span class="usage-check" title="Used">✓</span>';
  }
  if (value === false) return '<span class="usage-cross" title="Not found">✕</span>';
  return '<span class="usage-none">—</span>';
}

function updateUsageCell(propName, kind, value) {
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

function propTypeBadge(fieldType, type) {
  const map = {
    select: ['badge-dropdown','▾ Dropdown'], radio: ['badge-radio','◉ Radio'],
    checkbox: ['badge-checkbox','☑ Checkboxes'], booleancheckbox: ['badge-checkbox','☑ Boolean'],
    text: ['badge-text','T  Text'], textarea: ['badge-text','¶  Textarea'],
    phonenumber: ['badge-text','☏ Phone'], html: ['badge-text','⟨⟩ Rich Text'],
    number: ['badge-number','# Number'],
    date: type === 'datetime' ? ['badge-date','⊙ DateTime'] : ['badge-date','📅 Date'],
    file: ['badge-pending','📎 File'],
  };
  const [cls, label] = map[fieldType] || ['badge-pending', fieldType || '—'];
  return `<span class="badge ${cls}">${label}</span>`;
}

/* ── Filter / search ───────────────────────────────────────────────── */
function visibleProperties() {
  const search      = (document.getElementById('searchProps')?.value || '').toLowerCase();
  const filterSrc   = document.getElementById('filterSource')?.value || 'all';
  const filterUsage = document.getElementById('filterUsage')?.value  || 'all';

  return allProperties.filter((p) => {
    if (search && !`${p.label} ${p.name} ${p.groupName}`.toLowerCase().includes(search)) return false;
    if (filterSrc === 'custom' &&  p.hubspotDefined) return false;
    if (filterSrc === 'system' && !p.hubspotDefined) return false;
    if (filterUsage === 'unused') {
      if (!analysisStarted) return !p.hubspotDefined;
      return !p.hubspotDefined &&
        (p._recordCount === null || Number(p._recordCount) === 0) &&
        p._inWorkflow !== true && p._inForm !== true && p._inList !== true &&
        p._inPipeline !== true && p._inReport !== true && p._inEmail !== true;
    }
    if (filterUsage === 'used') {
      return Number(p._recordCount) > 0 || p._inWorkflow === true || p._inForm === true ||
        p._inList === true || p._inPipeline === true || p._inReport === true || p._inEmail === true;
    }
    return true;
  });
}

function filterProperties() {
  const vis = visibleProperties();
  renderPropertiesTable(vis);
  document.getElementById('filterCount').textContent = `${vis.length} of ${allProperties.length} shown`;
}

/* ── Analyze usage ─────────────────────────────────────────────────── */
async function analyzeUsage() {
  if (!allProperties.length) { alert('Load properties first.'); return; }
  analysisStarted = true;

  const analyzeBtn = document.getElementById('analyzeBtn');
  analyzeBtn.disabled = true;
  const progressEl = document.getElementById('analyzeProgress');
  const fillEl     = document.getElementById('analyzeProgressFill');
  const textEl     = document.getElementById('analyzeProgressText');
  const warningsEl = document.getElementById('analyzeWarnings');

  progressEl.style.display = 'block';
  warningsEl.style.display = 'none';
  fillEl.style.width = '0%';
  textEl.textContent = 'Fetching workflows and forms…';

  let ctx;
  try {
    const res = await fetch('/api/fetch-usage-context', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: '{}' });
    if (res.status === 401) { handleUnauth(); return; }
    ctx = await res.json();
  } catch (err) { textEl.textContent = `Network error: ${err.message}`; analyzeBtn.disabled = false; return; }

  if (ctx.warnings?.length) {
    warningsEl.style.display = 'block';
    warningsEl.innerHTML = `<strong>Some checks were skipped (missing API scopes):</strong><ul>${ctx.warnings.map((w) => `<li>${esc(w)}</li>`).join('')}</ul>`;
  }

  const wfSet       = new Set(ctx.workflowProperties  || []);
  const formSet     = new Set(ctx.formProperties       || []);
  const listSet     = new Set(ctx.listProperties       || []);
  const pipelineSet = new Set(ctx.pipelineProperties   || []);
  const reportSet   = new Set(ctx.reportProperties     || []);
  const emailSet    = new Set(ctx.emailProperties      || []);
  usageContext = { wfSet, formSet, listSet, pipelineSet, reportSet, emailSet, usageDetails: ctx.propertyUsageDetails || {} };

  textEl.textContent = `Found ${ctx.workflowCount} workflows, ${ctx.formCount} forms, ${ctx.listCount} lists, ${ctx.pipelineCount} pipelines, ${ctx.reportCount} reports, ${ctx.emailCount} emails. Checking records…`;
  fillEl.style.width = '10%';

  for (const prop of allProperties) {
    updateUsageCell(prop.name, 'workflow',  wfSet.has(prop.name));
    updateUsageCell(prop.name, 'form',      formSet.has(prop.name));
    updateUsageCell(prop.name, 'list',      listSet.has(prop.name));
    updateUsageCell(prop.name, 'pipeline',  pipelineSet.has(prop.name));
    updateUsageCell(prop.name, 'report',    reportSet.has(prop.name));
    updateUsageCell(prop.name, 'email',     emailSet.has(prop.name));
  }

  const customProps = allProperties.filter((p) => !p.hubspotDefined);
  const total = customProps.length;

  if (total === 0) {
    fillEl.style.width = '100%';
    textEl.textContent = 'Analysis complete — no custom properties to check.';
    analyzeBtn.disabled = false;
    filterProperties();
    return;
  }

  for (let i = 0; i < total; i++) {
    const prop = customProps[i];
    fillEl.style.width = Math.round(10 + (i / total) * 88) + '%';
    textEl.textContent = `Checking records: ${i + 1} / ${total} — "${prop.label || prop.name}"`;
    updateUsageCell(prop.name, 'records', 'loading');
    try {
      const res  = await fetch('/api/check-property-records', {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ objectType: getObjectType(), propertyName: prop.name }),
      });
      const data = await res.json();
      updateUsageCell(prop.name, 'records', data.success ? data.total : 'error');
    } catch { updateUsageCell(prop.name, 'records', 'error'); }
    if (i < total - 1) await sleep(120);
  }

  fillEl.style.width = '100%';
  const unusedCount = allProperties.filter((p) =>
    !p.hubspotDefined && Number(p._recordCount) === 0 &&
    p._inWorkflow === false && p._inForm === false && p._inList === false &&
    p._inPipeline === false && p._inReport === false && p._inEmail === false
  ).length;
  textEl.textContent = `Analysis complete — ${unusedCount} unused custom propert${unusedCount === 1 ? 'y' : 'ies'} found.`;
  analyzeBtn.disabled = false;
  filterProperties();
}

function sleep(ms) { return new Promise((resolve) => setTimeout(resolve, ms)); }

/* ── Selection ─────────────────────────────────────────────────────── */
function toggleSelectAll() {
  const master = document.getElementById('selectAllCb');
  document.querySelectorAll('.prop-cb:not([disabled])').forEach((cb) => { cb.checked = master.checked; });
  updateBulkBar();
}
function onRowCheckChange() { updateSelectAllState(); updateBulkBar(); }
function updateSelectAllState() {
  const all     = document.querySelectorAll('.prop-cb:not([disabled])');
  const checked = document.querySelectorAll('.prop-cb:not([disabled]):checked');
  const master  = document.getElementById('selectAllCb');
  master.checked       = all.length > 0 && checked.length === all.length;
  master.indeterminate = checked.length > 0 && checked.length < all.length;
}
function updateBulkBar() {
  const checked = document.querySelectorAll('.prop-cb:checked');
  const bar     = document.getElementById('bulkBar');
  if (checked.length === 0) { bar.style.display = 'none'; return; }
  bar.style.display = 'flex';
  document.getElementById('bulkCount').textContent = `${checked.length} propert${checked.length === 1 ? 'y' : 'ies'} selected`;
}
function clearSelection() {
  document.querySelectorAll('.prop-cb:checked').forEach((cb) => { cb.checked = false; });
  document.getElementById('selectAllCb').checked       = false;
  document.getElementById('selectAllCb').indeterminate = false;
  updateBulkBar();
}
function selectUnused() {
  const vis = visibleProperties();
  let count = 0;
  for (const prop of vis) {
    if (prop.hubspotDefined) continue;
    const isUnused = (prop._recordCount === null || Number(prop._recordCount) === 0) &&
      prop._inWorkflow !== true && prop._inForm !== true && prop._inList !== true &&
      prop._inPipeline !== true && prop._inReport !== true && prop._inEmail !== true;
    const cb = document.querySelector(`.prop-cb[data-name="${CSS.escape(prop.name)}"]`);
    if (cb && isUnused) { cb.checked = true; count++; }
  }
  updateSelectAllState(); updateBulkBar();
  if (count === 0) alert(analysisStarted ? 'No unused properties found.' : 'Run "Analyze Usage" first.');
}

/* ── Delete ────────────────────────────────────────────────────────── */
function deleteSelected() {
  const checked = [...document.querySelectorAll('.prop-cb:checked')];
  if (!checked.length) return;
  pendingDelete = checked.map((cb) => cb.dataset.name);
  showDeleteModal(pendingDelete);
}
function deleteSingleProperty(propName) { pendingDelete = [propName]; showDeleteModal(pendingDelete); }
function showDeleteModal(names) {
  const n        = names.length;
  const examples = names.slice(0, 5).map((n) => `<strong>${esc(n)}</strong>`).join(', ');
  document.getElementById('deleteModalBody').innerHTML =
    `You are about to permanently delete ${n} propert${n === 1 ? 'y' : 'ies'}` +
    (n <= 5 ? `: ${examples}` : ` including ${examples} and ${n - 5} more`) +
    `.<br/><br/>This action <strong>cannot be undone</strong>.`;
  document.getElementById('deleteModal').style.display = 'flex';
}
function closeDeleteModal(e) {
  if (e && e.target !== document.getElementById('deleteModal')) return;
  document.getElementById('deleteModal').style.display = 'none';
}
async function confirmDelete() {
  document.getElementById('deleteModal').style.display = 'none';
  const objectType = getObjectType();
  const names      = pendingDelete;
  pendingDelete    = [];

  const fillEl = document.getElementById('analyzeProgressFill');
  const textEl = document.getElementById('analyzeProgressText');
  document.getElementById('analyzeProgress').style.display = 'block';
  fillEl.style.width = '0%';

  let deleted = 0, failed = 0;
  for (let i = 0; i < names.length; i++) {
    const propName = names[i];
    fillEl.style.width = `${Math.round((i / names.length) * 100)}%`;
    textEl.textContent = `Deleting ${i + 1} / ${names.length}: "${propName}"…`;
    try {
      const res  = await fetch('/api/delete-property', {
        method: 'DELETE', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ objectType, propertyName: propName }),
      });
      if (res.status === 401) { handleUnauth(); return; }
      const data = await res.json();
      if (data.success) {
        deleted++;
        const row = document.getElementById(`prop-row-${propName}`);
        if (row) { row.classList.add('row-deleted'); setTimeout(() => row.remove(), 1200); }
        allProperties = allProperties.filter((p) => p.name !== propName);
      } else { failed++; }
    } catch { failed++; }
  }

  fillEl.style.width = '100%';
  textEl.textContent = `Deleted ${deleted} propert${deleted === 1 ? 'y' : 'ies'}` + (failed > 0 ? `, ${failed} failed` : '.');
  const custom = allProperties.filter((p) => !p.hubspotDefined).length;
  document.getElementById('mgmtSubtitle').textContent = `${allProperties.length} properties (${custom} custom)`;
  clearSelection();
  filterProperties();
}

/* ── Usage detail modal ────────────────────────────────────────────── */
function showUsageDetails(propName, type) {
  const details  = usageContext?.usageDetails?.[propName]?.[type] || [];
  const prop     = allProperties.find((p) => p.name === propName);
  const label    = prop?.label || propName;
  const typeLabels = { workflows: 'Workflows', forms: 'Forms', lists: 'Lists', pipelines: 'Pipelines', reports: 'Reports', emails: 'Marketing Emails' };
  document.getElementById('usageDetailTitle').textContent = `"${label}" — Used in ${typeLabels[type] || type}`;
  document.getElementById('usageDetailList').innerHTML = details.length
    ? details.map((name) => `<li>${esc(name)}</li>`).join('')
    : '<li class="usage-detail-none">No names available</li>';
  document.getElementById('usageDetailModal').style.display = 'flex';
}
function closeUsageDetailModal(e) {
  if (e && e.target !== document.getElementById('usageDetailModal')) return;
  document.getElementById('usageDetailModal').style.display = 'none';
}

/* ── Helpers ───────────────────────────────────────────────────────── */
function formatDate(dateStr) {
  if (!dateStr) return '—';
  try { return new Date(dateStr).toLocaleDateString(undefined, { year: 'numeric', month: 'short', day: 'numeric' }); }
  catch { return '—'; }
}

/* ── Export CSV ────────────────────────────────────────────────────── */
function exportCSV() {
  const vis        = visibleProperties();
  const objectType = getObjectType();
  const headers    = ['Label','Internal Name','Type','Group','Source','Records','In Workflow','In Form','In List','In Pipeline','In Report','In Email','Last Modified'];
  const boolCell   = (v) => v === true ? 'Yes' : v === false ? 'No' : '';
  const rows = vis.map((p) => [
    p.label || p.name, p.name, p.fieldType || '', p.groupName || '',
    p.hubspotDefined ? 'System' : 'Custom',
    (p._recordCount !== null && p._recordCount !== 'loading' && p._recordCount !== 'error') ? p._recordCount : '',
    boolCell(p._inWorkflow), boolCell(p._inForm), boolCell(p._inList),
    boolCell(p._inPipeline), boolCell(p._inReport), boolCell(p._inEmail),
    p.updatedAt ? new Date(p.updatedAt).toISOString().slice(0, 10) : '',
  ]);
  const csv = [headers, ...rows].map((r) => r.map(escapeCSV).join(',')).join('\r\n');
  downloadBlob(csv, `${objectType}-properties.csv`, 'text/csv');
}

/* ── Boot ──────────────────────────────────────────────────────────── */
initAuth();
