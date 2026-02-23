/* ── State ─────────────────────────────────────────────────────────── */
let parsedRows = [];    // parsed CSV rows
let isCreating = false; // lock to prevent double-submit

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
