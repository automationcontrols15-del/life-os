'use strict';
const express = require('express');
const { google } = require('googleapis');
const path = require('path');

const app = express();
app.use(express.json({ limit: '10mb' }));
app.use(express.static(path.join(__dirname, 'public')));

// ── CONFIG ────────────────────────────────────────────────────────────────────
const SPREADSHEET_ID = process.env.SPREADSHEET_ID;
const SHEET_NAME     = process.env.SHEET_NAME || 'LifeOS';

if (!SPREADSHEET_ID) {
  console.error('❌  SPREADSHEET_ID env var is missing. Set it in Railway → Variables.');
}
if (!process.env.GOOGLE_SERVICE_ACCOUNT_JSON) {
  console.error('❌  GOOGLE_SERVICE_ACCOUNT_JSON env var is missing. Set it in Railway → Variables.');
}

// ── GOOGLE AUTH ───────────────────────────────────────────────────────────────
function getAuth() {
  const credentials = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
  return new google.auth.GoogleAuth({
    credentials,
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });
}

async function getSheetsClient() {
  const auth = await getAuth().getClient();
  return google.sheets({ version: 'v4', auth });
}

// ── GET ALL DATA ──────────────────────────────────────────────────────────────
// Reads every row from the LifeOS sheet; each row is [key, JSON_value].
// Returns a plain object  { key: parsedValue, … }
async function getAllData() {
  const sheets = await getSheetsClient();
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `${SHEET_NAME}!A:B`,
  });
  const rows = res.data.values || [];
  const data = {};
  for (const [key, val] of rows) {
    if (!key) continue;
    try { data[key] = JSON.parse(val); }
    catch (_) { data[key] = val ?? null; }
  }
  return data;
}

// ── BATCH UPSERT ──────────────────────────────────────────────────────────────
// updates = { key: value, … }
// Reads current column A once, then batch-updates existing rows and appends new ones.
async function batchUpsert(updates) {
  const sheets = await getSheetsClient();

  // 1. Read existing keys → build rowIndex map (1-based)
  const readRes = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `${SHEET_NAME}!A:A`,
  });
  const existingRows = readRes.data.values || [];
  const rowMap = {};
  existingRows.forEach((r, i) => { if (r[0]) rowMap[r[0]] = i + 1; });

  const toUpdate = []; // { range, values } for batchUpdate
  const toAppend = []; // [key, json] rows for append

  for (const [key, value] of Object.entries(updates)) {
    const jsonVal = JSON.stringify(value);
    if (rowMap[key]) {
      toUpdate.push({
        range: `${SHEET_NAME}!A${rowMap[key]}:B${rowMap[key]}`,
        values: [[key, jsonVal]],
      });
    } else {
      toAppend.push([key, jsonVal]);
    }
  }

  // 2. Batch update existing rows (one API call)
  if (toUpdate.length) {
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: { valueInputOption: 'RAW', data: toUpdate },
    });
  }

  // 3. Append new rows (one API call)
  if (toAppend.length) {
    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: `${SHEET_NAME}!A:B`,
      valueInputOption: 'RAW',
      requestBody: { values: toAppend },
    });
  }
}

// ── ROUTES ────────────────────────────────────────────────────────────────────

// GET /api/data  →  { key: value, … }
app.get('/api/data', async (req, res) => {
  try {
    const data = await getAllData();
    res.json(data);
  } catch (e) {
    console.error('GET /api/data error:', e.message);
    res.status(500).json({ error: e.message });
  }
});

// POST /api/data  body: { key: value, … }  →  batch upsert
app.post('/api/data', async (req, res) => {
  try {
    const updates = req.body;
    if (!updates || typeof updates !== 'object' || Array.isArray(updates)) {
      return res.status(400).json({ error: 'Body must be a plain object of { key: value }' });
    }
    await batchUpsert(updates);
    res.json({ ok: true });
  } catch (e) {
    console.error('POST /api/data error:', e.message);
    res.status(500).json({ error: e.message });
  }
});

// Health check
app.get('/api/ping', (req, res) => res.json({ ok: true, ts: Date.now() }));

// SPA fallback — serve index.html for all other routes
app.get('*', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// ── START ─────────────────────────────────────────────────────────────────────
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`✅  Life OS running on port ${PORT}`));
