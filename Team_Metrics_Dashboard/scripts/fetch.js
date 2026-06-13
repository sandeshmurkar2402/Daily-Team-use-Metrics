const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');

const ROOT = path.join(__dirname, '..');

async function fetchSheetData() {
  const config = JSON.parse(fs.readFileSync(path.join(ROOT, 'config.json'), 'utf8'));

  let authOptions;

  if (process.env.GOOGLE_CREDENTIALS) {
    // GitHub Actions: credentials stored as base64-encoded JSON secret
    const credentials = JSON.parse(
      Buffer.from(process.env.GOOGLE_CREDENTIALS, 'base64').toString('utf8')
    );
    authOptions = { credentials };
  } else {
    // Local development: point to the credentials JSON file
    const keyFile = path.isAbsolute(config.credentialsPath)
      ? config.credentialsPath
      : path.join(ROOT, config.credentialsPath);
    authOptions = { keyFile };
  }

  const auth = new google.auth.GoogleAuth({
    ...authOptions,
    scopes: ['https://www.googleapis.com/auth/spreadsheets.readonly'],
  });

  const sheets = google.sheets({ version: 'v4', auth });

  console.log(`Fetching ${config.range} from sheet ${config.spreadsheetId}…`);

  const response = await sheets.spreadsheets.values.get({
    spreadsheetId: config.spreadsheetId,
    range: config.range,
  });

  const rawRows = response.data.values;
  if (!rawRows || rawRows.length < 2) {
    throw new Error('Sheet is empty or has only a header row.');
  }

  const headers = rawRows[0];
  const rows = rawRows.slice(1).map(row => {
    const obj = {};
    headers.forEach((header, i) => {
      obj[header] = row[i] ?? '';
    });
    return obj;
  });

  const output = {
    lastUpdated: new Date().toISOString(),
    headers,
    rows,
  };

  const dataDir = path.join(ROOT, 'data');
  if (!fs.existsSync(dataDir)) fs.mkdirSync(dataDir, { recursive: true });

  fs.writeFileSync(
    path.join(dataDir, 'metrics.json'),
    JSON.stringify(output, null, 2)
  );

  console.log(`✓ Saved ${rows.length} rows to data/metrics.json`);
}

fetchSheetData().catch(err => {
  console.error('✗ Fetch failed:', err.message);
  process.exit(1);
});
