const express = require('express');
const multer = require('multer');
const { parse } = require('csv-parse/sync');
const axios = require('axios');

const app = express();
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 5 * 1024 * 1024 }, // 5 MB max
  fileFilter: (_req, file, cb) => {
    if (!file.originalname.endsWith('.csv')) {
      return cb(new Error('Only .csv files are allowed'));
    }
    cb(null, true);
  },
});

app.use(express.json());
app.use(express.static('public'));

// ── Constants ──────────────────────────────────────────────────────────────

const PROPERTY_TYPES = {
  'Drop-down Select':    { type: 'enumeration', fieldType: 'select' },
  'Radio Select':        { type: 'enumeration', fieldType: 'radio' },
  'Multiple Checkboxes': { type: 'enumeration', fieldType: 'checkbox' },
};

const VALID_TYPES = Object.keys(PROPERTY_TYPES);

const DEFAULT_GROUPS = {
  contacts:  'contactinformation',
  companies: 'companyinformation',
  deals:     'dealinformation',
  tickets:   'ticketinformation',
  products:  'productinformation',
};

// ── Helpers ────────────────────────────────────────────────────────────────

/**
 * Converts a display label to a valid HubSpot internal property name.
 * Rules: lowercase, only a-z / 0-9 / underscore, must start with a letter.
 */
function toInternalName(label) {
  let name = label
    .toLowerCase()
    .trim()
    .replace(/\s+/g, '_')
    .replace(/[^a-z0-9_]/g, '');

  // HubSpot names must start with a letter
  if (/^[0-9]/.test(name)) name = 'p_' + name;

  return name.slice(0, 250); // HubSpot max length
}

/**
 * Builds the HubSpot options array from a semicolon-separated string.
 */
function parseOptions(optionsStr) {
  if (!optionsStr || !optionsStr.trim()) return [];
  return optionsStr
    .split(';')
    .map((o) => o.trim())
    .filter(Boolean)
    .map((label, i) => ({
      label,
      value: toInternalName(label) || `option_${i}`,
      displayOrder: i,
      hidden: false,
    }));
}

// ── Routes ─────────────────────────────────────────────────────────────────

/**
 * POST /api/parse-csv
 * Accepts a multipart CSV upload, validates columns and types, returns parsed rows.
 */
app.post('/api/parse-csv', upload.single('file'), (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ success: false, errors: ['No file uploaded.'] });
    }

    const records = parse(req.file.buffer.toString('utf-8'), {
      columns: true,
      skip_empty_lines: true,
      trim: true,
    });

    if (records.length === 0) {
      return res.status(400).json({ success: false, errors: ['The CSV file is empty.'] });
    }

    // Validate header columns
    const headers = Object.keys(records[0]).map((h) => h.trim());
    if (!headers.includes('Name') || !headers.includes('Type')) {
      return res.status(400).json({
        success: false,
        errors: ['CSV must have at least the columns: Name, Type'],
      });
    }

    // Validate each row
    const errors = [];
    records.forEach((row, i) => {
      const rowNum = i + 2; // +2 because row 1 is the header
      if (!row.Name || !row.Name.trim()) {
        errors.push(`Row ${rowNum}: Name is required`);
      }
      if (!row.Type || !row.Type.trim()) {
        errors.push(`Row ${rowNum}: Type is required`);
      } else if (!PROPERTY_TYPES[row.Type.trim()]) {
        errors.push(
          `Row ${rowNum}: Invalid type "${row.Type}". Must be one of: ${VALID_TYPES.join(', ')}`
        );
      }
    });

    if (errors.length) {
      return res.status(400).json({ success: false, errors });
    }

    // Normalise rows before returning
    const data = records.map((row) => ({
      Name:        row.Name.trim(),
      Type:        row.Type.trim(),
      Description: (row.Description || '').trim(),
      Options:     (row.Options || '').trim(),
    }));

    res.json({ success: true, data, count: data.length });
  } catch (err) {
    res.status(400).json({ success: false, errors: [err.message] });
  }
});

/**
 * POST /api/create-property
 * Creates a single HubSpot property.
 * Body: { token, objectType, property: { Name, Type, Description, Options } }
 */
app.post('/api/create-property', async (req, res) => {
  const { token, objectType, property } = req.body;

  if (!token || !objectType || !property) {
    return res.status(400).json({ success: false, error: 'Missing required fields.' });
  }

  const typeInfo = PROPERTY_TYPES[property.Type];
  if (!typeInfo) {
    return res.status(400).json({ success: false, error: `Unknown property type: ${property.Type}` });
  }

  const groupName = DEFAULT_GROUPS[objectType] || `${objectType}information`;
  const internalName = toInternalName(property.Name);

  const body = {
    name:        internalName,
    label:       property.Name,
    type:        typeInfo.type,
    fieldType:   typeInfo.fieldType,
    groupName,
    description: property.Description || '',
    options:     parseOptions(property.Options),
  };

  try {
    const response = await axios.post(
      `https://api.hubapi.com/crm/v3/properties/${objectType}`,
      body,
      {
        headers: {
          Authorization: `Bearer ${token}`,
          'Content-Type': 'application/json',
        },
      }
    );

    res.json({
      success:      true,
      internalName: response.data.name,
      label:        response.data.label,
    });
  } catch (err) {
    const hsError = err.response?.data;
    const message =
      hsError?.message ||
      (hsError?.errors?.[0]?.message) ||
      err.message ||
      'Unknown error';

    res.status(err.response?.status || 500).json({ success: false, error: message });
  }
});

// ── Manage-properties helpers ───────────────────────────────────────────────

/**
 * Follow HubSpot cursor pagination, collecting every item.
 * @param {string} token  Bearer token
 * @param {string} url    API URL (no query string)
 * @param {string} key    Key in the response body that holds the array
 * @param {object} extra  Extra query params
 */
async function paginateHubSpot(token, url, key, extra = {}) {
  const items = [];
  let after = null;
  do {
    const res = await axios.get(url, {
      headers: { Authorization: `Bearer ${token}` },
      params: { limit: 100, ...extra, ...(after ? { after } : {}) },
    });
    const batch = res.data[key];
    if (Array.isArray(batch)) items.push(...batch);
    after = res.data.paging?.next?.after ?? null;
  } while (after);
  return items;
}

/**
 * Scan a single workflow object for HubSpot property name references.
 * Uses JSON stringification + regex — fast, handles any workflow shape.
 */
function extractWorkflowProps(workflow) {
  const str = JSON.stringify(workflow);
  const names = new Set();
  // Matches: "propertyName":"foo"  "property":"foo"  "filterProperty":"foo"
  const re = /"(?:propertyName|property|filterProperty)"\s*:\s*"([^"]+)"/g;
  let m;
  while ((m = re.exec(str)) !== null) names.add(m[1]);
  return names;
}

/**
 * Scan a HubSpot form object (v3 or legacy) for referenced property names.
 */
function extractFormProps(form) {
  const names = new Set();

  // Recursively scan a list of fields, including any dependent/conditional fields.
  function scanFields(fields) {
    for (const field of (fields || [])) {
      if (field.name) names.add(field.name);
      // v3 forms can nest conditional fields under each field's dependentFields array
      for (const dep of (field.dependentFields || [])) {
        scanFields(dep.fields);
      }
    }
  }

  // v3 forms
  for (const group of (form.fieldGroups || [])) {
    scanFields(group.fields);
  }
  // legacy v2 forms
  scanFields(form.formFields);

  return names;
}

// ── Manage-properties routes ────────────────────────────────────────────────

/**
 * GET /api/list-properties?token=&objectType=
 * Returns all non-archived properties for the given object type.
 */
app.get('/api/list-properties', async (req, res) => {
  const { token, objectType } = req.query;
  if (!token || !objectType) {
    return res.status(400).json({ success: false, error: 'token and objectType are required.' });
  }
  try {
    // Properties API returns everything in one page (no cursor pagination)
    const response = await axios.get(
      `https://api.hubapi.com/crm/v3/properties/${objectType}`,
      {
        headers: { Authorization: `Bearer ${token}` },
        params: { archived: false },
      }
    );
    const properties = (response.data.results || [])
      .sort((a, b) => (a.label || a.name).localeCompare(b.label || b.name));
    res.json({ success: true, properties, count: properties.length });
  } catch (err) {
    const msg = err.response?.data?.message || err.message;
    res.status(err.response?.status || 500).json({ success: false, error: msg });
  }
});

/**
 * POST /api/fetch-usage-context
 * Body: { token, objectType }
 * Fetches all workflows and all forms, then returns:
 *   - workflowProperties: string[]   (property names referenced in any workflow)
 *   - formProperties:     string[]   (property names used as form fields)
 * Errors on either source are non-fatal — returned as warning messages.
 */
app.post('/api/fetch-usage-context', async (req, res) => {
  const { token } = req.body;
  if (!token) return res.status(400).json({ success: false, error: 'token is required.' });

  const warnings = [];

  // ── Workflows (v3 automation API) ──
  let workflowProps = new Set();
  let workflowCount = 0;
  try {
    const workflows = await paginateHubSpot(
      token,
      'https://api.hubapi.com/automation/v3/workflows',
      'workflows'
    );
    workflowCount = workflows.length;
    for (const wf of workflows) {
      for (const name of extractWorkflowProps(wf)) workflowProps.add(name);
    }
  } catch (err) {
    warnings.push(`Workflows: ${err.response?.data?.message || err.message}`);
  }

  // ── Forms (marketing v3 API) ──
  let formProps = new Set();
  let formCount = 0;
  try {
    const forms = await paginateHubSpot(
      token,
      'https://api.hubapi.com/marketing/v3/forms',
      'results'
    );
    formCount = forms.length;
    for (const f of forms) {
      for (const name of extractFormProps(f)) formProps.add(name);
    }
  } catch (err) {
    warnings.push(`Forms: ${err.response?.data?.message || err.message}`);
  }

  res.json({
    success: true,
    workflowProperties: Array.from(workflowProps),
    formProperties: Array.from(formProps),
    workflowCount,
    formCount,
    warnings,
  });
});

/**
 * POST /api/check-property-records
 * Body: { token, objectType, propertyName }
 * Uses the CRM search API to count how many records have this property set.
 */
app.post('/api/check-property-records', async (req, res) => {
  const { token, objectType, propertyName } = req.body;
  if (!token || !objectType || !propertyName) {
    return res.status(400).json({ success: false, error: 'token, objectType, propertyName are required.' });
  }
  try {
    const response = await axios.post(
      `https://api.hubapi.com/crm/v3/objects/${objectType}/search`,
      {
        filterGroups: [{ filters: [{ propertyName, operator: 'HAS_PROPERTY' }] }],
        limit: 1,
        properties: [],
      },
      { headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' } }
    );
    res.json({ success: true, total: response.data.total ?? 0 });
  } catch (err) {
    const msg = err.response?.data?.message || err.message;
    res.status(err.response?.status || 500).json({ success: false, error: msg });
  }
});

/**
 * DELETE /api/delete-property
 * Body: { token, objectType, propertyName }
 * Permanently deletes a custom HubSpot property. System properties will fail.
 */
app.delete('/api/delete-property', async (req, res) => {
  const { token, objectType, propertyName } = req.body;
  if (!token || !objectType || !propertyName) {
    return res.status(400).json({ success: false, error: 'token, objectType, propertyName are required.' });
  }
  try {
    await axios.delete(
      `https://api.hubapi.com/crm/v3/properties/${objectType}/${propertyName}`,
      { headers: { Authorization: `Bearer ${token}` } }
    );
    res.json({ success: true });
  } catch (err) {
    const msg = err.response?.data?.message || err.message;
    res.status(err.response?.status || 500).json({ success: false, error: msg });
  }
});

// ── Error handler ──────────────────────────────────────────────────────────

// eslint-disable-next-line no-unused-vars
app.use((err, _req, res, _next) => {
  res.status(500).json({ success: false, errors: [err.message] });
});

// ── Start ──────────────────────────────────────────────────────────────────

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`HubSpot Bulk Property Creator running at http://localhost:${PORT}`);
});
