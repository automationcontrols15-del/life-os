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

// ── GOOGLE FIT OAUTH ─────────────────────────────────────────────────────────
const GOOGLE_CLIENT_ID     = process.env.GOOGLE_CLIENT_ID;
const GOOGLE_CLIENT_SECRET = process.env.GOOGLE_CLIENT_SECRET;
const RAILWAY_URL          = process.env.RAILWAY_PUBLIC_DOMAIN
  ? `https://${process.env.RAILWAY_PUBLIC_DOMAIN}`
  : 'https://life-os-production-57af.up.railway.app';
const REDIRECT_URI         = `${RAILWAY_URL}/api/auth/google/callback`;

const FITNESS_SCOPES = [
  'https://www.googleapis.com/auth/fitness.activity.read',
  'https://www.googleapis.com/auth/fitness.heart_rate.read',
  'https://www.googleapis.com/auth/fitness.sleep.read',
  'https://www.googleapis.com/auth/fitness.body.read',
];

function getOAuth2Client() {
  return new google.auth.OAuth2(GOOGLE_CLIENT_ID, GOOGLE_CLIENT_SECRET, REDIRECT_URI);
}

// ── ROUTES ────────────────────────────────────────────────────────────────────

// GET /api/auth/google → redirect to Google OAuth consent
app.get('/api/auth/google', (req, res) => {
  if (!GOOGLE_CLIENT_ID || !GOOGLE_CLIENT_SECRET) {
    return res.status(500).json({ error: 'GOOGLE_CLIENT_ID or GOOGLE_CLIENT_SECRET not set' });
  }
  const oauth2Client = getOAuth2Client();
  const url = oauth2Client.generateAuthUrl({
    access_type: 'offline',
    scope: FITNESS_SCOPES,
    prompt: 'consent',
  });
  res.redirect(url);
});

// GET /api/auth/google/callback → exchange code → store tokens → redirect to SPA
app.get('/api/auth/google/callback', async (req, res) => {
  const { code, error } = req.query;
  if (error) return res.redirect('/?gfit=error');
  try {
    const oauth2Client = getOAuth2Client();
    const { tokens } = await oauth2Client.getToken(code);
    await batchUpsert({ googleFitTokens: tokens });
    res.redirect('/?gfit=connected');
  } catch (e) {
    console.error('OAuth callback error:', e.message);
    res.redirect('/?gfit=error');
  }
});

// GET /api/fit/status → is Google Fit connected?
app.get('/api/fit/status', async (req, res) => {
  try {
    const data = await getAllData();
    const tokens = data.googleFitTokens;
    res.json({ connected: !!(tokens && tokens.access_token) });
  } catch (e) {
    res.json({ connected: false });
  }
});

// GET /api/fit/data → today's fitness data
app.get('/api/fit/data', async (req, res) => {
  try {
    const data = await getAllData();
    const tokens = data.googleFitTokens;
    if (!tokens) return res.status(401).json({ error: 'Not connected to Google Fit' });

    const oauth2Client = getOAuth2Client();
    oauth2Client.setCredentials(tokens);

    // Persist refreshed tokens automatically
    oauth2Client.on('tokens', async (newTokens) => {
      const merged = { ...tokens, ...newTokens };
      await batchUpsert({ googleFitTokens: merged });
    });

    const fitness = google.fitness({ version: 'v1', auth: oauth2Client });
    const now = Date.now();
    const startOfDay = new Date();
    startOfDay.setHours(0, 0, 0, 0);

    const fitRes = await fitness.users.dataset.aggregate({
      userId: 'me',
      requestBody: {
        aggregateBy: [
          { dataTypeName: 'com.google.step_count.delta' },
          { dataTypeName: 'com.google.heart_rate.bpm' },
          { dataTypeName: 'com.google.calories.expended' },
          { dataTypeName: 'com.google.active_minutes' },
        ],
        bucketByTime: { durationMillis: 86400000 },
        startTimeMillis: startOfDay.getTime().toString(),
        endTimeMillis: now.toString(),
      },
    });

    let steps = 0, heartRate = 0, calories = 0, activeMinutes = 0;
    for (const bucket of (fitRes.data.bucket || [])) {
      for (const dataset of (bucket.dataset || [])) {
        for (const point of (dataset.point || [])) {
          const dtype = point.dataTypeName;
          const val = point.value?.[0];
          if (!val) continue;
          if (dtype === 'com.google.step_count.delta') steps += (val.intVal || 0);
          if (dtype === 'com.google.heart_rate.bpm') heartRate = (val.fpVal || 0);
          if (dtype === 'com.google.calories.expended') calories += (val.fpVal || 0);
          if (dtype === 'com.google.active_minutes') activeMinutes += (val.intVal || 0);
        }
      }
    }

    res.json({ steps, heartRate: Math.round(heartRate), calories: Math.round(calories), activeMinutes });
  } catch (e) {
    console.error('Fit data error:', e.message);
    if (e.message && e.message.includes('invalid_grant')) {
      await batchUpsert({ googleFitTokens: null });
    }
    res.status(500).json({ error: e.message });
  }
});

// POST /api/auth/google/disconnect → remove tokens
app.post('/api/auth/google/disconnect', async (req, res) => {
  try {
    await batchUpsert({ googleFitTokens: null });
    res.json({ ok: true });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

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
