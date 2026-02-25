require('dotenv').config();
const express = require('express');
const session = require('express-session');
const { google } = require('googleapis');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 3000;

// Middleware
app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(session({
  secret: process.env.SESSION_SECRET || 'orderly-crm-secret',
  resave: false,
  saveUninitialized: false,
  cookie: { maxAge: 24 * 60 * 60 * 1000 }
}));

// Static files
app.use(express.static(path.join(__dirname, 'public')));

// ==================== Google Sheets Auth ====================
let sheetsClient = null;
let driveClient = null;

async function getClients() {
  if (sheetsClient) return { sheets: sheetsClient, drive: driveClient };

  let auth;
  if (process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL && process.env.GOOGLE_PRIVATE_KEY) {
    auth = new google.auth.JWT(
      process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
      null,
      process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, '\n'),
      ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    );
  } else {
    auth = new google.auth.GoogleAuth({
      keyFile: process.env.GOOGLE_APPLICATION_CREDENTIALS || './credentials.json',
      scopes: ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'],
    });
  }

  sheetsClient = google.sheets({ version: 'v4', auth });
  driveClient = google.drive({ version: 'v3', auth });
  return { sheets: sheetsClient, drive: driveClient };
}

async function getSheets() {
  const { sheets } = await getClients();
  return sheets;
}

const SHEET_ID = process.env.GOOGLE_SHEET_ID;

// Tab name mapping - UPDATE THESE if your tab names differ
const TAB_NAMES = {
  pipeline: 'Pipeline',
  seed: 'Seed Re-Approach',
  kol: 'KOL Round',
  outreach: 'Outreach Plan',
  token: 'Token Holdings'
};

// ==================== Auth Middleware ====================
function requireAuth(req, res, next) {
  if (req.session && req.session.authenticated) return next();
  res.status(401).json({ error: 'Not authenticated' });
}

// ==================== Auth Routes ====================
app.post('/api/login', (req, res) => {
  const { password } = req.body;
  if (password === (process.env.CRM_PASSWORD || 'orderly2026')) {
    req.session.authenticated = true;
    res.json({ success: true });
  } else {
    res.status(401).json({ error: 'Incorrect password' });
  }
});

app.post('/api/logout', (req, res) => {
  req.session.destroy();
  res.json({ success: true });
});

app.get('/api/auth-check', (req, res) => {
  res.json({ authenticated: !!(req.session && req.session.authenticated) });
});

// ==================== ADD POC COLUMN TO ALL TABS ====================
app.get('/api/add-poc-column', async (req, res) => {
  console.log('\n=== Adding POC column to all tabs ===');
  const results = [];
  const tabsToUpdate = ['Pipeline', 'Seed Re-Approach', 'KOL Round', 'Outreach Plan'];

  try {
    const sheets = await getSheets();

    for (const tabName of tabsToUpdate) {
      try {
        // Get current headers
        const headerRes = await sheets.spreadsheets.values.get({
          spreadsheetId: SHEET_ID,
          range: `'${tabName}'!1:1`,
        });
        const headers = headerRes.data.values ? headerRes.data.values[0] : [];

        // Check if POC already exists
        if (headers.some(h => h.toLowerCase() === 'poc')) {
          results.push({ tab: tabName, status: 'SKIPPED', detail: 'POC column already exists' });
          console.log(`${tabName}: POC already exists, skipping`);
          continue;
        }

        // Find where to insert POC (after Org or after Name)
        let insertIdx = headers.findIndex(h => h.toLowerCase() === 'org');
        if (insertIdx === -1) insertIdx = headers.findIndex(h => h.toLowerCase() === 'name');
        insertIdx = insertIdx === -1 ? headers.length : insertIdx + 1;

        // Get all data
        const dataRes = await sheets.spreadsheets.values.get({
          spreadsheetId: SHEET_ID,
          range: `'${tabName}'`,
        });
        const allRows = dataRes.data.values || [];

        // Insert POC into each row
        const updatedRows = allRows.map((row, i) => {
          const newRow = [...row];
          // Pad row if shorter than insertIdx
          while (newRow.length < insertIdx) newRow.push('');
          newRow.splice(insertIdx, 0, i === 0 ? 'POC' : '');
          return newRow;
        });

        // Write back
        await sheets.spreadsheets.values.update({
          spreadsheetId: SHEET_ID,
          range: `'${tabName}'!A1`,
          valueInputOption: 'USER_ENTERED',
          resource: { values: updatedRows },
        });

        results.push({ tab: tabName, status: 'OK', detail: `POC column added at position ${insertIdx + 1}` });
        console.log(`${tabName}: POC column added`);
      } catch (e) {
        results.push({ tab: tabName, status: 'FAILED', detail: e.message });
        console.error(`${tabName}: Failed - ${e.message}`);
      }
    }

    console.log('=== Done ===\n');
    res.json({ success: true, results });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ==================== DIAGNOSTIC ROUTE ====================
app.get('/api/diagnose', async (req, res) => {
  console.log('\n=== RUNNING DIAGNOSTICS ===');
  console.log('Sheet ID:', SHEET_ID);
  const results = { sheetId: SHEET_ID, steps: [] };

  try {
    const { sheets, drive } = await getClients();
    results.steps.push({ step: 'Auth', status: 'OK', detail: 'Google API client created' });

    // Step 1: Check file via Drive API
    try {
      const fileInfo = await drive.files.get({
        fileId: SHEET_ID,
        fields: 'id,name,mimeType,owners,shared'
      });
      const info = fileInfo.data;
      console.log('File info:', JSON.stringify(info, null, 2));
      results.steps.push({
        step: 'Drive API - File Info',
        status: 'OK',
        detail: {
          name: info.name,
          mimeType: info.mimeType,
          isGoogleSheet: info.mimeType === 'application/vnd.google-apps.spreadsheet',
          isXlsx: info.mimeType === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        }
      });
    } catch (e) {
      console.error('Drive API error:', e.message);
      results.steps.push({ step: 'Drive API - File Info', status: 'FAILED', detail: e.message });
    }

    // Step 2: Try spreadsheets.get (metadata only)
    try {
      const meta = await sheets.spreadsheets.get({
        spreadsheetId: SHEET_ID,
        fields: 'spreadsheetId,properties.title,sheets.properties'
      });
      const sheetNames = meta.data.sheets.map(s => s.properties.title);
      console.log('Sheet tabs:', sheetNames);
      results.steps.push({
        step: 'Sheets API - Metadata',
        status: 'OK',
        detail: { title: meta.data.properties.title, tabs: sheetNames }
      });
    } catch (e) {
      console.error('Sheets metadata error:', e.message);
      results.steps.push({ step: 'Sheets API - Metadata', status: 'FAILED', detail: e.message });
    }

    // Step 3: Try reading first tab
    try {
      const firstTab = Object.values(TAB_NAMES)[0];
      const valRes = await sheets.spreadsheets.values.get({
        spreadsheetId: SHEET_ID,
        range: `'${firstTab}'!A1:C3`,
      });
      console.log('Values read OK, rows:', valRes.data.values?.length || 0);
      results.steps.push({
        step: `Sheets API - Read "${firstTab}"`,
        status: 'OK',
        detail: { rows: valRes.data.values?.length || 0, sample: valRes.data.values }
      });
    } catch (e) {
      console.error('Values read error:', e.message);
      results.steps.push({ step: 'Sheets API - Read Values', status: 'FAILED', detail: e.message });
    }

  } catch (e) {
    results.steps.push({ step: 'Auth', status: 'FAILED', detail: e.message });
  }

  console.log('=== DIAGNOSTICS COMPLETE ===\n');
  res.json(results);
});

// ==================== Read Routes ====================
app.get('/api/sheets/:tab', requireAuth, async (req, res) => {
  const { tab } = req.params;
  const tabName = TAB_NAMES[tab];

  if (!tabName) {
    return res.status(400).json({ error: `Unknown tab: ${tab}` });
  }

  try {
    const sheets = await getSheets();
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: `'${tabName}'`,
    });

    const rows = response.data.values || [];
    if (rows.length < 2) {
      return res.json({ headers: rows[0] || [], data: [] });
    }

    const headers = rows[0];
    const data = rows.slice(1).map(row => {
      const obj = {};
      headers.forEach((header, i) => {
        obj[header] = row[i] || '';
      });
      return obj;
    });

    res.json({ headers, data });
  } catch (error) {
    console.error(`Error reading ${tab}:`, error.message);
    res.status(500).json({ error: `Failed to read ${tabName}: ${error.message}` });
  }
});

// ==================== Write Routes ====================

// Add a new row
app.post('/api/sheets/:tab', requireAuth, async (req, res) => {
  const { tab } = req.params;
  const tabName = TAB_NAMES[tab];
  const { rowData } = req.body;

  if (!tabName) return res.status(400).json({ error: `Unknown tab: ${tab}` });

  try {
    const sheets = await getSheets();

    const headerResponse = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: `'${tabName}'!1:1`,
    });

    const headers = headerResponse.data.values ? headerResponse.data.values[0] : [];
    if (!headers.length) return res.status(400).json({ error: 'Sheet has no headers' });

    const newRow = headers.map(header => rowData[header] || '');

    // Auto-set Last Updated
    const lastUpdatedIdx = headers.findIndex(h =>
      h.toLowerCase().includes('last updated') || h.toLowerCase().includes('date')
    );
    if (lastUpdatedIdx !== -1 && !newRow[lastUpdatedIdx]) {
      newRow[lastUpdatedIdx] = new Date().toISOString().split('T')[0];
    }

    await sheets.spreadsheets.values.append({
      spreadsheetId: SHEET_ID,
      range: `'${tabName}'`,
      valueInputOption: 'USER_ENTERED',
      insertDataOption: 'INSERT_ROWS',
      resource: { values: [newRow] },
    });

    res.json({ success: true, message: `Added to ${tabName}` });
  } catch (error) {
    console.error(`Error writing to ${tab}:`, error.message);
    res.status(500).json({ error: `Failed to write to ${tabName}: ${error.message}` });
  }
});

// Update an existing row
app.put('/api/sheets/:tab/:rowIndex', requireAuth, async (req, res) => {
  const { tab, rowIndex } = req.params;
  const tabName = TAB_NAMES[tab];
  const { rowData } = req.body;
  const rowNum = parseInt(rowIndex) + 2;

  if (!tabName) return res.status(400).json({ error: `Unknown tab: ${tab}` });

  try {
    const sheets = await getSheets();

    const headerResponse = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: `'${tabName}'!1:1`,
    });

    const headers = headerResponse.data.values ? headerResponse.data.values[0] : [];
    const updatedRow = headers.map(header => {
      if (rowData.hasOwnProperty(header)) return rowData[header];
      return '';
    });

    // Auto-set Last Updated
    const lastUpdatedIdx = headers.findIndex(h =>
      h.toLowerCase().includes('last updated')
    );
    if (lastUpdatedIdx !== -1) {
      updatedRow[lastUpdatedIdx] = new Date().toISOString().split('T')[0];
    }

    await sheets.spreadsheets.values.update({
      spreadsheetId: SHEET_ID,
      range: `'${tabName}'!A${rowNum}`,
      valueInputOption: 'USER_ENTERED',
      resource: { values: [updatedRow] },
    });

    res.json({ success: true, message: `Updated row ${rowIndex} in ${tabName}` });
  } catch (error) {
    console.error(`Error updating ${tab}:`, error.message);
    res.status(500).json({ error: `Failed to update ${tabName}: ${error.message}` });
  }
});

// ==================== Summary Stats ====================
app.get('/api/summary', requireAuth, async (req, res) => {
  try {
    const sheets = await getSheets();
    const results = {};

    for (const [key, tabName] of Object.entries(TAB_NAMES)) {
      try {
        const response = await sheets.spreadsheets.values.get({
          spreadsheetId: SHEET_ID,
          range: `'${tabName}'`,
        });
        const rows = response.data.values || [];
        results[key] = { count: Math.max(0, rows.length - 1), headers: rows[0] || [] };
      } catch (e) {
        results[key] = { count: 0, headers: [] };
      }
    }

    res.json(results);
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// ==================== Catch-all: serve frontend ====================
app.get('*', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// ==================== Start ====================
app.listen(PORT, () => {
  console.log(`\n  Orderly CRM running at http://localhost:${PORT}`);
  console.log(`  Sheet ID: ${SHEET_ID}`);
  console.log(`  Run diagnostics: http://localhost:3000/api/diagnose\n`);
});
