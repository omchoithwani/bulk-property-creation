require('dotenv').config();

const express = require('express');
const multer = require('multer');
const { parse } = require('csv-parse/sync');
const axios = require('axios');
const session = require('express-session');

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

// ── OAuth config ────────────────────────────────────────────────────────────

const CLIENT_ID     = process.env.HUBSPOT_CLIENT_ID;
const CLIENT_SECRET = process.env.HUBSPOT_CLIENT_SECRET;
const REDIRECT_URI  = process.env.HUBSPOT_REDIRECT_URI || 'http://localhost:3000/oauth/callback';

const SCOPES = [
  'automation',
  'business-intelligence',
  'content',
  'crm.lists.read',
  'crm.objects.companies.read',
  'crm.objects.contacts.read',
  'crm.objects.custom.read',
  'crm.objects.deals.read',
  'crm.objects.leads.read',
  'crm.objects.line_items.read',
  'crm.objects.listings.read',
  'crm.objects.orders.write',
  'crm.objects.owners.read',
  'crm.objects.products.read',
  'crm.objects.projects.read',
  'crm.objects.services.read',
  'crm.schemas.appointments.read',
  'crm.schemas.appointments.write',
  'crm.schemas.carts.read',
  'crm.schemas.carts.write',
  'crm.schemas.commercepayments.read',
  'crm.schemas.commercepayments.write',
  'crm.schemas.companies.read',
  'crm.schemas.companies.write',
  'crm.schemas.contacts.read',
  'crm.schemas.contacts.write',
  'crm.schemas.courses.read',
  'crm.schemas.courses.write',
  'crm.schemas.custom.read',
  'crm.schemas.deals.read',
  'crm.schemas.deals.write',
  'crm.schemas.forecasts.read',
  'crm.schemas.invoices.read',
  'crm.schemas.invoices.write',
  'crm.schemas.line_items.read',
  'crm.schemas.listings.read',
  'crm.schemas.listings.write',
  'crm.schemas.orders.read',
  'crm.schemas.orders.write',
  'crm.schemas.projects.read',
  'crm.schemas.projects.write',
  'crm.schemas.quotes.read',
  'crm.schemas.quotes.write',
  'crm.schemas.services.read',
  'crm.schemas.services.write',
  'crm.schemas.subscriptions.read',
  'crm.schemas.subscriptions.write',
  'forms',
  'oauth',
  'tickets',
].join(' ');

// ── Middleware ──────────────────────────────────────────────────────────────

app.use(express.json());
app.use(session({
  secret: process.env.SESSION_SECRET || 'hubspot-property-tool-secret',
  resave: false,
  saveUninitialized: false,
  cookie: {
    // In production (HTTPS) cookies should be secure; locally they work without it
    secure: process.env.NODE_ENV === 'production',
    httpOnly: true,
    maxAge: 8 * 60 * 60 * 1000, // 8 hours — matches HubSpot token lifetime
  },
}));
app.use(express.static('public'));

// ── Session token helper ────────────────────────────────────────────────────

/**
 * Returns a valid access token from the session.
 * Automatically refreshes the token if it is within 5 minutes of expiry.
 * Throws a 401 error if the session has no token at all.
 */
async function getValidToken(req) {
  if (!req.session?.accessToken) {
    const err = new Error('Not connected. Please connect your HubSpot account.');
    err.statusCode = 401;
    throw err;
  }

  // Refresh if the token expires within the next 5 minutes
  if (
    req.session.expiresAt &&
    Date.now() > req.session.expiresAt - 5 * 60 * 1000
  ) {
    if (!req.session.refreshToken || !CLIENT_SECRET) {
      const err = new Error('Session expired. Please reconnect your HubSpot account.');
      err.statusCode = 401;
      throw err;
    }

    const tokenRes = await axios.post(
      'https://api.hubapi.com/oauth/v1/token',
      new URLSearchParams({
        grant_type:    'refresh_token',
        client_id:     CLIENT_ID,
        client_secret: CLIENT_SECRET,
        refresh_token: req.session.refreshToken,
      }).toString(),
      { headers: { 'Content-Type': 'application/x-www-form-urlencoded' } }
    );

    req.session.accessToken  = tokenRes.data.access_token;
    req.session.refreshToken = tokenRes.data.refresh_token;
    req.session.expiresAt    = Date.now() + tokenRes.data.expires_in * 1000;
  }

  return req.session.accessToken;
}

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

const STANDARD_OBJECTS = [
  { value: 'contacts',              label: 'Contacts' },
  { value: 'companies',             label: 'Companies' },
  { value: 'deals',                 label: 'Deals' },
  { value: 'tickets',               label: 'Tickets' },
  { value: 'products',              label: 'Products' },
  { value: 'line_items',            label: 'Line Items' },
  { value: 'quotes',                label: 'Quotes' },
  { value: 'calls',                 label: 'Calls' },
  { value: 'emails',                label: 'Emails' },
  { value: 'meetings',              label: 'Meetings' },
  { value: 'notes',                 label: 'Notes' },
  { value: 'tasks',                 label: 'Tasks' },
  { value: 'communications',        label: 'Communications' },
  { value: 'feedback_submissions',  label: 'Feedback Submissions' },
  { value: 'leads',                 label: 'Leads' },
];

// ── Helpers ────────────────────────────────────────────────────────────────

function toInternalName(label) {
  let name = label
    .toLowerCase()
    .trim()
    .replace(/\s+/g, '_')
    .replace(/[^a-z0-9_]/g, '');
  if (/^[0-9]/.test(name)) name = 'p_' + name;
  return name.slice(0, 250);
}

/**
 * Resolves the correct property group name for an object type.
 * Strategy:
 *  1. Known default group (instant)
 *  2. Groups API (first non-HubSpot-defined group)
 *  3. Most-used group from existing properties
 *  4. Throw
 */
async function resolveGroupName(token, objectType) {
  if (DEFAULT_GROUPS[objectType]) return DEFAULT_GROUPS[objectType];

  try {
    const res = await axios.get(
      `https://api.hubapi.com/crm/v3/properties/groups/${objectType}`,
      { headers: { Authorization: `Bearer ${token}` } }
    );
    const groups = res.data.results || [];
    const preferred = groups.find(g => !g.hubspotDefined) || groups[0];
    if (preferred?.name) return preferred.name;
  } catch { /* fall through */ }

  try {
    const res = await axios.get(
      `https://api.hubapi.com/crm/v3/properties/${objectType}`,
      { headers: { Authorization: `Bearer ${token}` } }
    );
    const counts = {};
    for (const prop of (res.data.results || [])) {
      if (prop.groupName) counts[prop.groupName] = (counts[prop.groupName] || 0) + 1;
    }
    const top = Object.entries(counts).sort((a, b) => b[1] - a[1])[0];
    if (top?.[0]) return top[0];
  } catch { /* fall through */ }

  throw new Error(
    `Cannot determine a valid property group for object type "${objectType}". ` +
    `Ensure the app has crm.schemas.custom.read scope, or that the object already has at least one property.`
  );
}

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

// ── OAuth routes ────────────────────────────────────────────────────────────

/**
 * GET /oauth/authorize
 * Redirects the browser to HubSpot's OAuth consent screen.
 */
app.get('/oauth/authorize', (req, res) => {
  if (!CLIENT_ID) {
    return res.status(500).send(
      'HUBSPOT_CLIENT_ID is not set. Add it to your environment variables.'
    );
  }
  const url =
    `https://app.hubspot.com/oauth/authorize` +
    `?client_id=${encodeURIComponent(CLIENT_ID)}` +
    `&redirect_uri=${encodeURIComponent(REDIRECT_URI)}` +
    `&scope=${encodeURIComponent(SCOPES)}`;
  res.redirect(url);
});

/**
 * GET /oauth/callback
 * HubSpot redirects here after the user grants (or denies) access.
 * Exchanges the auth code for access + refresh tokens and stores them in the session.
 */
app.get('/oauth/callback', async (req, res) => {
  const { code, error, error_description } = req.query;

  if (error || !code) {
    const msg = error_description || error || 'Authorization was denied.';
    return res.redirect(`/?auth_error=${encodeURIComponent(msg)}`);
  }

  try {
    const tokenRes = await axios.post(
      'https://api.hubapi.com/oauth/v1/token',
      new URLSearchParams({
        grant_type:    'authorization_code',
        client_id:     CLIENT_ID,
        client_secret: CLIENT_SECRET,
        redirect_uri:  REDIRECT_URI,
        code,
      }).toString(),
      { headers: { 'Content-Type': 'application/x-www-form-urlencoded' } }
    );

    req.session.accessToken  = tokenRes.data.access_token;
    req.session.refreshToken = tokenRes.data.refresh_token;
    req.session.expiresAt    = Date.now() + tokenRes.data.expires_in * 1000;

    // Fetch portal info (best-effort — not fatal if it fails)
    try {
      const infoRes = await axios.get(
        `https://api.hubapi.com/oauth/v1/access-tokens/${tokenRes.data.access_token}`
      );
      req.session.portalId  = infoRes.data.hub_id;
      req.session.hubDomain = infoRes.data.hub_domain;
    } catch { /* portal info is non-critical */ }

    res.redirect('/');
  } catch (err) {
    const msg = err.response?.data?.message || err.message;
    res.redirect(`/?auth_error=${encodeURIComponent(msg)}`);
  }
});

/**
 * GET /oauth/status
 * Returns whether the current session has a valid connection and portal info.
 */
app.get('/oauth/status', (req, res) => {
  if (req.session?.accessToken) {
    res.json({
      connected: true,
      portalId:  req.session.portalId  || null,
      hubDomain: req.session.hubDomain || null,
    });
  } else {
    res.json({ connected: false });
  }
});

/**
 * POST /oauth/disconnect
 * Destroys the session, effectively logging the user out.
 */
app.post('/oauth/disconnect', (req, res) => {
  req.session.destroy((err) => {
    res.json({ success: !err });
  });
});

// ── API routes ─────────────────────────────────────────────────────────────

/**
 * POST /api/parse-csv
 * No auth required — only parses and validates the uploaded file.
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

    const headers = Object.keys(records[0]).map((h) => h.trim());
    if (!headers.includes('Name') || !headers.includes('Type')) {
      return res.status(400).json({
        success: false,
        errors: ['CSV must have at least the columns: Name, Type'],
      });
    }

    const errors = [];
    records.forEach((row, i) => {
      const rowNum = i + 2;
      if (!row.Name || !row.Name.trim()) errors.push(`Row ${rowNum}: Name is required`);
      if (!row.Type || !row.Type.trim()) {
        errors.push(`Row ${rowNum}: Type is required`);
      } else if (!PROPERTY_TYPES[row.Type.trim()]) {
        errors.push(`Row ${rowNum}: Invalid type "${row.Type}". Must be one of: ${VALID_TYPES.join(', ')}`);
      }
    });

    if (errors.length) return res.status(400).json({ success: false, errors });

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
 * Body: { objectType, property: { Name, Type, Description, Options }, defaultGroup? }
 */
app.post('/api/create-property', async (req, res) => {
  const { objectType, property, defaultGroup } = req.body;

  if (!objectType || !property) {
    return res.status(400).json({ success: false, error: 'Missing required fields.' });
  }

  const typeInfo = PROPERTY_TYPES[property.Type];
  if (!typeInfo) {
    return res.status(400).json({ success: false, error: `Unknown property type: ${property.Type}` });
  }

  try {
    const token = await getValidToken(req);
    const groupName = defaultGroup || await resolveGroupName(token, objectType);
    const internalName = toInternalName(property.Name);

    const body = {
      name:        internalName,
      label:       property.Name,
      type:        typeInfo.type,
      fieldType:   typeInfo.fieldType,
      groupName,
      description: property.Description || '',
      ...(typeInfo.enumeration ? { options: parseOptions(property.Options) } : {}),
    };

    const response = await axios.post(
      `https://api.hubapi.com/crm/v3/properties/${objectType}`,
      body,
      { headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' } }
    );

    res.json({ success: true, internalName: response.data.name, label: response.data.label });
  } catch (err) {
    const status = err.statusCode || err.response?.status || 500;
    const message = err.response?.data?.message || err.response?.data?.errors?.[0]?.message || err.message || 'Unknown error';
    res.status(status).json({ success: false, error: message });
  }
});

// ── Manage-properties helpers ───────────────────────────────────────────────

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

function extractPropNamesFromJson(str) {
  const names = new Set();
  const re = /"[a-zA-Z]*[Pp]roperty(?:[Nn]ame)?"\s*:\s*"([a-z][a-z0-9_]*)"/g;
  let m;
  while ((m = re.exec(str)) !== null) names.add(m[1]);
  return names;
}

function extractPropsFromJsonWithTokens(obj) {
  const str = JSON.stringify(obj);
  const names = extractPropNamesFromJson(str);
  const tokenRe = /\{\{[a-z_]+\.([a-z0-9_]+)\}\}/g;
  let m;
  while ((m = tokenRe.exec(str)) !== null) names.add(m[1]);
  return names;
}

function extractWorkflowProps(workflow)   { return extractPropsFromJsonWithTokens(workflow); }
function extractEmailProps(email)         { return extractPropsFromJsonWithTokens(email); }
function extractListProps(list)           { return extractPropNamesFromJson(JSON.stringify(list)); }
function extractPipelineProps(pipeline)   { return extractPropNamesFromJson(JSON.stringify(pipeline)); }
function extractReportProps(report)       { return extractPropNamesFromJson(JSON.stringify(report)); }

function extractFormProps(form) {
  const names = new Set();

  function scanFields(fields) {
    for (const field of (fields || [])) {
      if (field.name) names.add(field.name);
      for (const dep of (field.dependentFields || [])) {
        for (const filter of (dep.dependentFieldFilters || [])) {
          if (filter.dependentFormField) scanFields([filter.dependentFormField]);
        }
        scanFields(dep.fields);
      }
    }
  }

  for (const group of (form.fieldGroups || [])) scanFields(group.fields);

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

// ── Manage routes ───────────────────────────────────────────────────────────

/**
 * GET /api/list-object-types
 */
app.get('/api/list-object-types', async (req, res) => {
  let customObjects = [];
  let warning = null;

  try {
    const token = await getValidToken(req);

    const allSchemas = [];
    let after = undefined;
    do {
      const schemasRes = await axios.get('https://api.hubapi.com/crm/v3/schemas', {
        headers: { Authorization: `Bearer ${token}` },
        params:  { archived: false, limit: 100, ...(after ? { after } : {}) },
      });
      allSchemas.push(...(schemasRes.data.results || []));
      after = schemasRes.data.paging?.next?.after;
    } while (after);

    customObjects = await Promise.all(
      allSchemas.map(async schema => {
        let defaultGroup = null;
        try {
          const groupsRes = await axios.get(
            `https://api.hubapi.com/crm/v3/properties/groups/${schema.objectTypeId}`,
            { headers: { Authorization: `Bearer ${token}` } }
          );
          const groups = groupsRes.data.results || [];
          const preferred = groups.find(g => !g.hubspotDefined) || groups[0];
          defaultGroup = preferred?.name || null;
        } catch { /* defaultGroup stays null */ }
        const value = schema.objectTypeId || schema.fullyQualifiedName || schema.name;
        const label = schema.labels?.plural || schema.labels?.singular || schema.name || value;
        if (!value) return null;
        return { value, label, defaultGroup };
      })
    );
    customObjects = customObjects.filter(Boolean);
  } catch (err) {
    if (err.statusCode === 401) {
      return res.status(401).json({ success: false, error: err.message, unauthenticated: true });
    }
    warning = `Custom objects could not be loaded: ${err.response?.data?.message || err.message}`;
  }

  res.json({ success: true, objectTypes: [...STANDARD_OBJECTS, ...customObjects], warning });
});

/**
 * GET /api/list-properties?objectType=
 */
app.get('/api/list-properties', async (req, res) => {
  const { objectType } = req.query;
  if (!objectType) {
    return res.status(400).json({ success: false, error: 'objectType is required.' });
  }
  try {
    const token = await getValidToken(req);
    const response = await axios.get(
      `https://api.hubapi.com/crm/v3/properties/${objectType}`,
      { headers: { Authorization: `Bearer ${token}` }, params: { archived: false } }
    );
    const properties = (response.data.results || [])
      .sort((a, b) => (a.label || a.name).localeCompare(b.label || b.name));
    res.json({ success: true, properties, count: properties.length });
  } catch (err) {
    const status = err.statusCode || err.response?.status || 500;
    const msg = err.response?.data?.message || err.message;
    res.status(status).json({ success: false, error: msg, unauthenticated: status === 401 });
  }
});

/**
 * POST /api/fetch-usage-context
 */
app.post('/api/fetch-usage-context', async (req, res) => {
  const warnings = [];
  const usageDetails = {};

  function addUsage(propName, type, sourceName) {
    if (!usageDetails[propName]) usageDetails[propName] = {};
    if (!usageDetails[propName][type]) usageDetails[propName][type] = [];
    if (!usageDetails[propName][type].includes(sourceName)) {
      usageDetails[propName][type].push(sourceName);
    }
  }

  let token;
  try {
    token = await getValidToken(req);
  } catch (err) {
    return res.status(err.statusCode || 401).json({ success: false, error: err.message, unauthenticated: true });
  }

  let workflowProps = new Set(), workflowCount = 0;
  try {
    const workflows = await paginateHubSpot(token, 'https://api.hubapi.com/automation/v3/workflows', 'workflows');
    workflowCount = workflows.length;
    for (const wf of workflows) {
      const wfName = wf.name || `Workflow ${wf.id || ''}`.trim();
      for (const name of extractWorkflowProps(wf)) { workflowProps.add(name); addUsage(name, 'workflows', wfName); }
    }
  } catch (err) { warnings.push(`Workflows: ${err.response?.data?.message || err.message}`); }

  let formProps = new Set(), formCount = 0;
  try {
    const forms = await paginateHubSpot(token, 'https://api.hubapi.com/marketing/v3/forms', 'results');
    formCount = forms.length;
    for (const f of forms) {
      const formName = f.name || f.id || 'Unnamed Form';
      for (const name of extractFormProps(f)) { formProps.add(name); addUsage(name, 'forms', formName); }
    }
  } catch (err) { warnings.push(`Forms: ${err.response?.data?.message || err.message}`); }

  let listProps = new Set(), listCount = 0;
  try {
    const lists = await paginateHubSpot(token, 'https://api.hubapi.com/crm/v3/lists', 'lists', { includeFilters: true });
    listCount = lists.length;
    for (const list of lists) {
      const listName = list.name || list.listId || 'Unnamed List';
      for (const name of extractListProps(list)) { listProps.add(name); addUsage(name, 'lists', listName); }
    }
  } catch (err) { warnings.push(`Lists: ${err.response?.data?.message || err.message}`); }

  let pipelineProps = new Set(), pipelineCount = 0;
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
        for (const name of extractPipelineProps(pipeline)) { pipelineProps.add(name); addUsage(name, 'pipelines', pipelineName); }
      }
    }
  } catch (err) { warnings.push(`Pipelines: ${err.response?.data?.message || err.message}`); }

  let emailProps = new Set(), emailCount = 0;
  try {
    const emails = await paginateHubSpot(token, 'https://api.hubapi.com/marketing/v3/emails', 'results');
    emailCount = emails.length;
    for (const email of emails) {
      const emailName = email.name || email.id || 'Unnamed Email';
      for (const name of extractEmailProps(email)) { emailProps.add(name); addUsage(name, 'emails', emailName); }
    }
  } catch (err) { warnings.push(`Marketing emails: ${err.response?.data?.message || err.message}`); }

  let reportProps = new Set(), reportCount = 0;
  try {
    const reportRes = await axios.get(
      'https://api.hubapi.com/reporting/v1/reports',
      { headers: { Authorization: `Bearer ${token}` }, params: { limit: 300 } }
    );
    const reports = reportRes.data.objects || reportRes.data.results || [];
    reportCount = reports.length;
    for (const report of reports) {
      const reportName = report.name || report.id || 'Unnamed Report';
      for (const name of extractReportProps(report)) { reportProps.add(name); addUsage(name, 'reports', reportName); }
    }
  } catch (err) { warnings.push(`Reports: ${err.response?.data?.message || err.message}`); }

  res.json({
    success: true,
    workflowProperties:   Array.from(workflowProps),
    formProperties:       Array.from(formProps),
    listProperties:       Array.from(listProps),
    pipelineProperties:   Array.from(pipelineProps),
    reportProperties:     Array.from(reportProps),
    emailProperties:      Array.from(emailProps),
    propertyUsageDetails: usageDetails,
    workflowCount, formCount, listCount, pipelineCount, reportCount, emailCount,
    warnings,
  });
});

/**
 * POST /api/check-property-records
 * Body: { objectType, propertyName }
 */
app.post('/api/check-property-records', async (req, res) => {
  const { objectType, propertyName } = req.body;
  if (!objectType || !propertyName) {
    return res.status(400).json({ success: false, error: 'objectType and propertyName are required.' });
  }
  try {
    const token = await getValidToken(req);
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
    const status = err.statusCode || err.response?.status || 500;
    res.status(status).json({ success: false, error: err.response?.data?.message || err.message });
  }
});

/**
 * DELETE /api/delete-property
 * Body: { objectType, propertyName }
 */
app.delete('/api/delete-property', async (req, res) => {
  const { objectType, propertyName } = req.body;
  if (!objectType || !propertyName) {
    return res.status(400).json({ success: false, error: 'objectType and propertyName are required.' });
  }
  try {
    const token = await getValidToken(req);
    await axios.delete(
      `https://api.hubapi.com/crm/v3/properties/${objectType}/${propertyName}`,
      { headers: { Authorization: `Bearer ${token}` } }
    );
    res.json({ success: true });
  } catch (err) {
    const status = err.statusCode || err.response?.status || 500;
    res.status(status).json({ success: false, error: err.response?.data?.message || err.message });
  }
});

// ── Error handler ──────────────────────────────────────────────────────────

// eslint-disable-next-line no-unused-vars
app.use((err, _req, res, _next) => {
  const status = err.statusCode || 500;
  res.status(status).json({ success: false, error: err.message });
});

// ── Start ──────────────────────────────────────────────────────────────────

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`HubSpot Property Manager running at http://localhost:${PORT}`);
});
