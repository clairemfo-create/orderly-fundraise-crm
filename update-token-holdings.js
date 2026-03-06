require('dotenv').config();
const { google } = require('googleapis');
const XLSX = require('xlsx');
const path = require('path');

const SHEET_ID = process.env.GOOGLE_SHEET_ID;
const TAB_NAME = 'Token Holdings';

async function main() {
  // Parse the Excel file
  const excelPath = process.argv[2];
  if (!excelPath) {
    console.error('Usage: node update-token-holdings.js <path-to-excel>');
    process.exit(1);
  }

  console.log(`Reading Excel: ${excelPath}`);
  const wb = XLSX.readFile(excelPath);
  const ws = wb.Sheets[wb.SheetNames[0]];
  const data = XLSX.utils.sheet_to_json(ws);
  console.log(`Found ${data.length} rows`);

  // Map Excel columns to CRM headers
  const crmHeaders = [
    'Name', 'Source', 'Categories', 'Wallet Address',
    'Allocation', 'Remaining Vesting Bal', 'Vested - Claimed',
    'In Wallet', 'Staking', 'Sold', 'Total', '% Vested', '% Sold'
  ];

  const excelToCrm = {
    'Name/Code': 'Name',
    'Source': 'Source',
    'Categories': 'Categories',
    'Wallet Address': 'Wallet Address',
    'Allocation': 'Allocation',
    'Remaining Vesting Bal': 'Remaining Vesting Bal',
    'Vested - Claimed': 'Vested - Claimed',
    'In Wallet': 'In Wallet',
    'Staking': 'Staking',
    'Sold': 'Sold',
    'Total': 'Total',
    '% Vested': '% Vested',
    '% Sold (calc)': '% Sold'
  };

  const rows = data.map(row => {
    return crmHeaders.map(h => {
      const excelKey = Object.keys(excelToCrm).find(k => excelToCrm[k] === h);
      if (!excelKey) return '';
      let val = row[excelKey];
      if (val === undefined || val === null) return '';
      // Format numbers nicely
      if (typeof val === 'number') {
        if (h === '% Vested' || h === '% Sold') {
          return (val * 100).toFixed(1) + '%';
        }
        if (Number.isInteger(val)) return val.toString();
        return Math.round(val).toString();
      }
      return val.toString();
    });
  });

  // Auth with Google
  let auth;
  if (process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL && process.env.GOOGLE_PRIVATE_KEY) {
    auth = new google.auth.JWT(
      process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
      null,
      process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, '\n'),
      ['https://www.googleapis.com/auth/spreadsheets']
    );
  } else {
    auth = new google.auth.GoogleAuth({
      keyFile: process.env.GOOGLE_APPLICATION_CREDENTIALS || './credentials.json',
      scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    });
  }

  const sheets = google.sheets({ version: 'v4', auth });

  // Clear existing data (keep header)
  console.log('Clearing existing token data...');
  await sheets.spreadsheets.values.clear({
    spreadsheetId: SHEET_ID,
    range: `'${TAB_NAME}'!A2:Z1000`,
  });

  // Write new header + data
  console.log('Writing new data...');
  const allRows = [crmHeaders, ...rows];
  await sheets.spreadsheets.values.update({
    spreadsheetId: SHEET_ID,
    range: `'${TAB_NAME}'!A1`,
    valueInputOption: 'USER_ENTERED',
    resource: { values: allRows },
  });

  console.log(`\nDone! Updated ${rows.length} rows in Token Holdings tab.`);
  console.log('Refresh your CRM to see the changes.');
}

main().catch(e => {
  console.error('Error:', e.message);
  process.exit(1);
});
