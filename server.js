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
