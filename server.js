const express = require('express');
const { google } = require('googleapis');
const path = require('path');
const fs = require('fs');

const app = express();
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// Load service account credentials
const CREDENTIALS_PATH = process.env.CREDENTIALS_PATH || './credentials.json';
let auth;
try {
  const credentials = JSON.parse(fs.readFileSync(CREDENTIALS_PATH, 'utf8'));
  auth = new google.auth.GoogleAuth({
    credentials,
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });
} catch (e) {
  console.error('credentials.json 파일을 찾을 수 없습니다:', e.message);
  console.error('서비스 계정 키 파일을 credentials.json으로 저장하세요.');
  process.exit(1);
}

const sheets = google.sheets({ version: 'v4', auth });

function getSpreadsheetId(urlOrId) {
  const m = (urlOrId || '').match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  return m ? m[1] : urlOrId;
}

// GET /api/meta?url=...
app.get('/api/meta', async (req, res) => {
  const id = getSpreadsheetId(req.query.url);
  if (!id) return res.status(400).json({ error: '올바른 URL이 아닙니다.' });
  try {
    const resp = await sheets.spreadsheets.get({
      spreadsheetId: id,
      fields: 'properties.title,sheets.properties',
    });
    res.json(resp.data);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// GET /api/values?url=...&sheet=...
app.get('/api/values', async (req, res) => {
  const id = getSpreadsheetId(req.query.url);
  const sheet = req.query.sheet;
  if (!id) return res.status(400).json({ error: '올바른 URL이 아닙니다.' });
  try {
    const resp = await sheets.spreadsheets.values.get({
      spreadsheetId: id,
      range: sheet,
    });
    res.json(resp.data);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// POST /api/values  body: { url, sheet, data: [{range, values}] }
app.post('/api/values', async (req, res) => {
  const { url, data } = req.body;
  const id = getSpreadsheetId(url);
  if (!id || !data) return res.status(400).json({ error: '잘못된 요청입니다.' });
  try {
    const resp = await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: id,
      requestBody: {
        valueInputOption: 'USER_ENTERED',
        data,
      },
    });
    res.json({ updatedCells: resp.data.totalUpdatedCells });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// POST /api/append  body: { url, sheet, values: [[...]] }
app.post('/api/append', async (req, res) => {
  const { url, sheet, values } = req.body;
  const id = getSpreadsheetId(url);
  if (!id || !values) return res.status(400).json({ error: '잘못된 요청입니다.' });
  try {
    const resp = await sheets.spreadsheets.values.append({
      spreadsheetId: id,
      range: sheet,
      valueInputOption: 'USER_ENTERED',
      requestBody: { values },
    });
    res.json(resp.data);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`\n✅ Google Sheets Editor 서버 실행 중: http://localhost:${PORT}`);
  console.log('   브라우저에서 위 주소를 여세요.\n');
});
