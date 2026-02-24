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
  // Enumeration types — require Options
  'Drop-down Select':      { type: 'enumeration', fieldType: 'select',       enumeration: true },
  'Radio Select':          { type: 'enumeration', fieldType: 'radio',        enumeration: true },
  'Multiple Checkboxes':   { type: 'enumeration', fieldType: 'checkbox',     enumeration: true },
  // Text / string types
  'Single Line Text':      { type: 'string',      fieldType: 'text' },
  'Multi-line Text':       { type: 'string',      fieldType: 'textarea' },
  'Phone Number':          { type: 'string',      fieldType: 'phonenumber' },
  'URL':                   { type: 'string',      fieldType: 'text' },
  'Rich Text':             { type: 'string',      fieldType: 'html' },
  // Numeric
  'Number':                { type: 'number',      fieldType: 'number' },
  // Date / time
  'Date Picker':           { type: 'date',        fieldType: 'date' },
  'Date and Time Picker':  { type: 'datetime',    fieldType: 'date' },
};

const VALID_TYPES = Object.keys(PROPERTY_TYPES);

const DEFAULT_GROUPS = {
  contacts:  'contactinformation',
  companies: 'companyinformation',
  deals:     'dealinformation',
  tickets:   'ticketinformation',
  products:  'productinformation',
};

// Standard CRM objects shown in the UI by default (no API call needed)
const STANDARD_OBJECTS = [
  { value: 'contacts',  label: 'Contacts' },
  { value: 'companies', label: 'Companies' },
  { value: 'deals',     label: 'Deals' },
  { value: 'tickets',   label: 'Tickets' },
  { value: 'products',  label: 'Products' },
];

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
 * Returns the property group name to use when creating a property.
 * For standard objects the group is known; for custom objects we fetch
 * the first available group from the HubSpot groups API.
 */
async function resolveGroupName(token, objectType) {
  if (DEFAULT_GROUPS[objectType]) return DEFAULT_GROUPS[objectType];
  try {
    const res = await axios.get(
      `https://api.hubapi.com/crm/v3/properties/groups/${objectType}`,
      { headers: { Authorization: `Bearer ${token}` } }
    );
    const groups = res.data.results || [];
    // Prefer a non-HubSpot-defined group (i.e. the object's own default group)
    const preferred = groups.find(g => !g.hubspotDefined) || groups[0];
    if (preferred) return preferred.name;
  } catch { /* fall through to the guessed name */ }
  return `${objectType}information`;
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
  const { token, objectType, property, defaultGroup } = req.body;

  if (!token || !objectType || !property) {
    return res.status(400).json({ success: false, error: 'Missing required fields.' });
  }

  const typeInfo = PROPERTY_TYPES[property.Type];
  if (!typeInfo) {
    return res.status(400).json({ success: false, error: `Unknown property type: ${property.Type}` });
  }

  // Use the group name provided by the caller (resolved at object-type load time)
  // and only fall back to the dynamic lookup if it was not supplied.
  const groupName = defaultGroup || await resolveGroupName(token, objectType);
  const internalName = toInternalName(property.Name);

  const body = {
    name:        internalName,
    label:       property.Name,
    type:        typeInfo.type,
    fieldType:   typeInfo.fieldType,
    groupName,
    description: property.Description || '',
    // Only enumeration types use options — sending options for other types causes API errors
    ...(typeInfo.enumeration ? { options: parseOptions(property.Options) } : {}),
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
 * Shared regex scan: extracts property internal names from a JSON-stringified object.
 *
 * Catches ANY JSON key whose name ends in "Property" or "PropertyName"
 * (case-insensitive on the suffix), e.g.:
 *   property, propertyName, filterProperty, targetProperty, sourceProperty,
 *   fromPropertyName, toPropertyName, sourcePropertyName, targetPropertyName,
 *   associatedProperty, actionPropertyName, …
 *
 * Value is constrained to look like a HubSpot internal name:
 *   starts with a-z, followed by a-z / 0-9 / _ only.
 * This avoids false-positives on things like "propertyType":"ENUMERATION".
 */
function extractPropNamesFromJson(str) {
  const names = new Set();
  // Key: anything + "Property" optionally + "Name"  (both casings)
  // Value: must look like a HubSpot internal property name (lowercase snake_case)
  const re = /"[a-zA-Z]*[Pp]roperty(?:[Nn]ame)?"\s*:\s*"([a-z][a-z0-9_]*)"/g;
  let m;
  while ((m = re.exec(str)) !== null) names.add(m[1]);
  return names;
}

/**
 * Scan any HubSpot object for property references using two strategies:
 *  1. JSON key scan   — catches "propertyName":"foo", "targetProperty":"foo", etc.
 *  2. Token scan      — catches {{contact.foo}} personalization tokens in text fields
 * Used by both workflow and email scanners.
 */
function extractPropsFromJsonWithTokens(obj) {
  const str = JSON.stringify(obj);
  const names = extractPropNamesFromJson(str);
  const tokenRe = /\{\{[a-z_]+\.([a-z0-9_]+)\}\}/g;
  let m;
  while ((m = tokenRe.exec(str)) !== null) names.add(m[1]);
  return names;
}

/**
 * Scan a single workflow object for HubSpot property name references.
 */
function extractWorkflowProps(workflow) {
  return extractPropsFromJsonWithTokens(workflow);
}

/**
 * Scan a HubSpot marketing email object for property references.
 * Catches {{contact.my_property}} tokens in subject, htmlBody, plainTextBody, etc.
 */
function extractEmailProps(email) {
  return extractPropsFromJsonWithTokens(email);
}

/**
 * Scan a HubSpot list object for referenced property names.
 * List filter branches use "property":"propName" in their filter conditions.
 */
function extractListProps(list) {
  return extractPropNamesFromJson(JSON.stringify(list));
}

/**
 * Scan a HubSpot pipeline object for referenced property names.
 * Catches any property references in stage metadata / required properties.
 */
function extractPipelineProps(pipeline) {
  return extractPropNamesFromJson(JSON.stringify(pipeline));
}

/**
 * Scan a HubSpot report object for referenced property names.
 */
function extractReportProps(report) {
  return extractPropNamesFromJson(JSON.stringify(report));
}

/**
 * Scan a HubSpot form object (v3 or legacy) for referenced property names.
 */
function extractFormProps(form) {
  const names = new Set();

  // Recursively scan a list of field objects.
  function scanFields(fields) {
    for (const field of (fields || [])) {
      if (field.name) names.add(field.name);

      // v3 forms: conditional fields live at
      //   field.dependentFields[].dependentFieldFilters[].dependentFormField
      // (NOT dep.fields — that path doesn't exist in the v3 response)
      for (const dep of (field.dependentFields || [])) {
        for (const filter of (dep.dependentFieldFilters || [])) {
          if (filter.dependentFormField) {
            scanFields([filter.dependentFormField]); // recurse for further nesting
          }
        }
        // legacy fallback: dep.fields[] (old embedded-forms structure)
        scanFields(dep.fields);
      }
    }
  }

  // v3 forms: fieldGroups[].fields[]
  for (const group of (form.fieldGroups || [])) {
    scanFields(group.fields);
  }

  // legacy v2 forms: formFields can be a flat array OR an array-of-row-arrays
  //   [[field1, field2], [field3]]  ← rows
  const legacyFields = form.formFields;
  if (Array.isArray(legacyFields)) {
    if (legacyFields.length > 0 && Array.isArray(legacyFields[0])) {
      for (const row of legacyFields) scanFields(row);
    } else {
      scanFields(legacyFields);
    }
  }

  return names;
}

// ── Manage-properties routes ────────────────────────────────────────────────

/**
 * GET /api/list-object-types?token=
 * Returns standard CRM objects plus any custom objects in the portal
 * (fetched from GET /crm/v3/schemas — requires crm.schemas.custom.read scope).
 * Each custom object entry includes a `defaultGroup` so callers can pass it
 * directly to create-property without an extra round-trip.
 * If the schemas call fails (missing scope, network, etc.) the standard objects
 * are still returned; a non-fatal warning is included in the response.
 */
app.get('/api/list-object-types', async (req, res) => {
  const { token } = req.query;
  if (!token) return res.status(400).json({ success: false, error: 'token is required.' });

  let customObjects = [];
  let warning = null;

  try {
    const schemasRes = await axios.get('https://api.hubapi.com/crm/v3/schemas', {
      headers: { Authorization: `Bearer ${token}` },
      params:  { archived: false },
    });

    // Fetch the property groups for each custom object in parallel so we know
    // the correct defaultGroup to use when creating properties later.
    customObjects = await Promise.all(
      (schemasRes.data.results || []).map(async schema => {
        let defaultGroup = null;
        try {
          const groupsRes = await axios.get(
            `https://api.hubapi.com/crm/v3/properties/groups/${schema.objectTypeId}`,
            { headers: { Authorization: `Bearer ${token}` } }
          );
          const groups = groupsRes.data.results || [];
          // Prefer a group that belongs to the object itself (hubspotDefined=false)
          const preferred = groups.find(g => !g.hubspotDefined) || groups[0];
          defaultGroup = preferred?.name || null;
        } catch { /* defaultGroup stays null; create-property will handle it */ }
        return {
          value:        schema.objectTypeId,
          label:        schema.labels?.plural || schema.name,
          defaultGroup,
        };
      })
    );
  } catch (err) {
    warning = `Custom objects could not be loaded: ${err.response?.data?.message || err.message}`;
  }

  res.json({
    success:     true,
    objectTypes: [...STANDARD_OBJECTS, ...customObjects],
    warning,
  });
});

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
 * Fetches workflows, forms, lists, pipelines, reports, and marketing emails, then
 * returns property names referenced in each source.
 * Errors on any source are non-fatal — returned as warning messages.
 */
app.post('/api/fetch-usage-context', async (req, res) => {
  const { token } = req.body;
  if (!token) return res.status(400).json({ success: false, error: 'token is required.' });

  const warnings = [];

  // propName -> { workflows: string[], forms: string[], lists: string[], pipelines: string[], reports: string[] }
  const usageDetails = {};
  function addUsage(propName, type, sourceName) {
    if (!usageDetails[propName]) usageDetails[propName] = {};
    if (!usageDetails[propName][type]) usageDetails[propName][type] = [];
    if (!usageDetails[propName][type].includes(sourceName)) {
      usageDetails[propName][type].push(sourceName);
    }
  }

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
      const wfName = wf.name || `Workflow ${wf.id || ''}`.trim();
      for (const name of extractWorkflowProps(wf)) {
        workflowProps.add(name);
        addUsage(name, 'workflows', wfName);
      }
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
      const formName = f.name || f.id || 'Unnamed Form';
      for (const name of extractFormProps(f)) {
        formProps.add(name);
        addUsage(name, 'forms', formName);
      }
    }
  } catch (err) {
    warnings.push(`Forms: ${err.response?.data?.message || err.message}`);
  }

  // ── Lists (CRM v3 lists API) ──
  let listProps = new Set();
  let listCount = 0;
  try {
    const lists = await paginateHubSpot(
      token,
      'https://api.hubapi.com/crm/v3/lists',
      'lists',
      { includeFilters: true }
    );
    listCount = lists.length;
    for (const list of lists) {
      const listName = list.name || list.listId || 'Unnamed List';
      for (const name of extractListProps(list)) {
        listProps.add(name);
        addUsage(name, 'lists', listName);
      }
    }
  } catch (err) {
    warnings.push(`Lists: ${err.response?.data?.message || err.message}`);
  }

  // ── Pipelines (CRM v3 — deals + tickets only) ──
  let pipelineProps = new Set();
  let pipelineCount = 0;
  try {
    for (const pipelineObj of ['deals', 'tickets']) {
      const res = await axios.get(
        `https://api.hubapi.com/crm/v3/pipelines/${pipelineObj}`,
        { headers: { Authorization: `Bearer ${token}` } }
      );
      const pipelines = res.data.results || [];
      pipelineCount += pipelines.length;
      for (const pipeline of pipelines) {
        const pipelineName = pipeline.label || pipeline.id || 'Unnamed Pipeline';
        for (const name of extractPipelineProps(pipeline)) {
          pipelineProps.add(name);
          addUsage(name, 'pipelines', pipelineName);
        }
      }
    }
  } catch (err) {
    warnings.push(`Pipelines: ${err.response?.data?.message || err.message}`);
  }

  // ── Marketing emails (marketing v3 API) ──
  // Catches {{contact.property}} personalization tokens in subject lines and body HTML.
  let emailProps = new Set();
  let emailCount = 0;
  try {
    const emails = await paginateHubSpot(
      token,
      'https://api.hubapi.com/marketing/v3/emails',
      'results'
    );
    emailCount = emails.length;
    for (const email of emails) {
      const emailName = email.name || email.id || 'Unnamed Email';
      for (const name of extractEmailProps(email)) {
        emailProps.add(name);
        addUsage(name, 'emails', emailName);
      }
    }
  } catch (err) {
    warnings.push(`Marketing emails: ${err.response?.data?.message || err.message}`);
  }

  // ── Reports (reporting v1 API) ──
  let reportProps = new Set();
  let reportCount = 0;
  try {
    const reportRes = await axios.get(
      'https://api.hubapi.com/reporting/v1/reports',
      {
        headers: { Authorization: `Bearer ${token}` },
        params: { limit: 300 },
      }
    );
    const reports = reportRes.data.objects || reportRes.data.results || [];
    reportCount = reports.length;
    for (const report of reports) {
      const reportName = report.name || report.id || 'Unnamed Report';
      for (const name of extractReportProps(report)) {
        reportProps.add(name);
        addUsage(name, 'reports', reportName);
      }
    }
  } catch (err) {
    warnings.push(`Reports: ${err.response?.data?.message || err.message}`);
  }

  res.json({
    success: true,
    workflowProperties:  Array.from(workflowProps),
    formProperties:      Array.from(formProps),
    listProperties:      Array.from(listProps),
    pipelineProperties:  Array.from(pipelineProps),
    reportProperties:    Array.from(reportProps),
    emailProperties:     Array.from(emailProps),
    propertyUsageDetails: usageDetails,
    workflowCount,
    formCount,
    listCount,
    pipelineCount,
    reportCount,
    emailCount,
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
