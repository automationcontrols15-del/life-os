'use strict';
const express      = require('express');
const { google }   = require('googleapis');
const path         = require('path');
const bcrypt       = require('bcryptjs');
const jwt          = require('jsonwebtoken');
const cookieParser = require('cookie-parser');

const app = express();
app.use(express.json({ limit: '10mb' }));
app.use(cookieParser());
app.use(express.static(path.join(__dirname, 'public')));

// ── CONFIG ──────────────────────────────────────────────────────────────────
const SPREADSHEET_ID = process.env.SPREADSHEET_ID;
const SHEET_NAME     = process.env.SHEET_NAME || 'LifeOS';
const USERS_SHEET    = 'Users';
const JWT_SECRET     = process.env.JWT_SECRET;
const COOKIE_NAME    = 'lifeos_token';
const COOKIE_OPTS    = {
  httpOnly: true,
  secure: process.env.NODE_ENV === 'production',
  sameSite: 'lax',
  maxAge: 7 * 24 * 60 * 60 * 1000
};

if (!SPREADSHEET_ID)   console.error('⚠  SPREADSHEET_ID env var is missing.');
if (!JWT_SECRET)       console.error('⚠  JWT_SECRET env var is missing.');
if (!process.env.GOOGLE_SERVICE_ACCOUNT_JSON)
  console.error('⚠  GOOGLE_SERVICE_ACCOUNT_JSON env var is missing.');

// ── GOOGLE AUTH ──────────────────────────────────────────────────────────────
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

// ── USERS SHEET ──────────────────────────────────────────────────────────────
// Columns: user_id | email | password_hash | name | role | status | created_at

async function ensureUsersSheet() {
  const sheets = await getSheetsClient();
  const meta = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
  const exists = meta.data.sheets.some(s => s.properties.title === USERS_SHEET);
  if (!exists) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: {
        requests: [{ addSheet: { properties: { title: USERS_SHEET } } }]
      }
    });
    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: `${USERS_SHEET}!A1`,
      valueInputOption: 'RAW',
      requestBody: { values: [['user_id','email','password_hash','name','role','status','created_at']] }
    });
  }
}

async function getUsers() {
  const sheets = await getSheetsClient();
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `${USERS_SHEET}!A:G`,
  });
  const rows = res.data.values || [];
  if (rows.length < 2) return [];
  return rows.slice(1).map(r => ({
    user_id:       r[0] || '',
    email:         r[1] || '',
    password_hash: r[2] || '',
    name:          r[3] || '',
    role:          r[4] || 'user',
    status:        r[5] || 'active',
    created_at:    r[6] || ''
  })).filter(u => u.user_id);
}

async function getUserByEmail(email) {
  const users = await getUsers();
  return users.find(u => u.email.toLowerCase() === email.toLowerCase()) || null;
}

async function countUsers() {
  const users = await getUsers();
  return users.length;
}

async function createUser(userData) {
  const sheets = await getSheetsClient();
  await sheets.spreadsheets.values.append({
    spreadsheetId: SPREADSHEET_ID,
    range: `${USERS_SHEET}!A:G`,
    valueInputOption: 'RAW',
    requestBody: {
      values: [[
        userData.user_id,
        userData.email,
        userData.password_hash,
        userData.name,
        userData.role,
        userData.status,
        userData.created_at
      ]]
    }
  });
}

async function updateUserRow(user_id, patch) {
  const sheets = await getSheetsClient();
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `${USERS_SHEET}!A:G`,
  });
  const rows = res.data.values || [];
  let rowIdx = -1;
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === user_id) { rowIdx = i + 1; break; }
  }
  if (rowIdx === -1) throw new Error('User not found');
  const current = rows[rowIdx - 1];
  const updated = [
    current[0],
    patch.email         !== undefined ? patch.email         : current[1],
    patch.password_hash !== undefined ? patch.password_hash : current[2],
    patch.name          !== undefined ? patch.name          : current[3],
    patch.role          !== undefined ? patch.role          : current[4],
    patch.status        !== undefined ? patch.status        : current[5],
    current[6]
  ];
  await sheets.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID,
    range: `${USERS_SHEET}!A${rowIdx}:G${rowIdx}`,
    valueInputOption: 'RAW',
    requestBody: { values: [updated] }
  });
}

// ── JWT HELPERS ──────────────────────────────────────────────────────────────
function signToken(payload) {
  return jwt.sign(payload, JWT_SECRET, { expiresIn: '7d' });
}

// ── AUTH MIDDLEWARE ──────────────────────────────────────────────────────────
function requireAuth(req, res, next) {
  const token = req.cookies[COOKIE_NAME];
  if (!token) return res.status(401).json({ error: 'Not authenticated' });
  try {
    req.user = jwt.verify(token, JWT_SECRET);
    next();
  } catch (e) {
    res.clearCookie(COOKIE_NAME);
    res.status(401).json({ error: 'Session expired' });
  }
}

function requireAdmin(req, res, next) {
  if (req.user.role !== 'admin') return res.status(403).json({ error: 'Admin only' });
  next();
}

// ── DATA HELPERS (per-user isolation) ────────────────────────────────────────
// Keys in LifeOS sheet are prefixed: "{userId}:{key}" for regular users.
// Admin sees legacy unprefixed keys (backward compat) + their own prefixed keys.

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

async function getUserDataSafe(userId, isAdmin) {
  const all = await getAllData();
  const prefix = userId + ':';
  const result = {};

  // Collect prefixed keys for this user
  for (const [key, val] of Object.entries(all)) {
    if (key.startsWith(prefix)) {
      result[key.slice(prefix.length)] = val;
    }
  }

  // Admin: also load legacy unprefixed keys (backward compat) if not already present
  if (isAdmin) {
    for (const [key, val] of Object.entries(all)) {
      if (!key.includes(':') && result[key] === undefined) {
        result[key] = val;
      }
    }
  }

  return result;
}

async function batchUpsert(updates) {
  const sheets = await getSheetsClient();
  const readRes = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `${SHEET_NAME}!A:A`,
  });
  const existingRows = readRes.data.values || [];
  const rowMap = {};
  existingRows.forEach((r, i) => { if (r[0]) rowMap[r[0]] = i + 1; });

  const toUpdate = [];
  const toAppend = [];

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

  if (toUpdate.length) {
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: { valueInputOption: 'RAW', data: toUpdate },
    });
  }
  if (toAppend.length) {
    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: `${SHEET_NAME}!A:B`,
      valueInputOption: 'RAW',
      requestBody: { values: toAppend },
    });
  }
}

async function batchUpsertUser(userId, updates) {
  const prefixed = {};
  for (const [key, val] of Object.entries(updates)) {
    prefixed[`${userId}:${key}`] = val;
  }
  await batchUpsert(prefixed);
}

// ── GOOGLE OAUTH CLIENT (for Fit) ─────────────────────────────────────────────
function getOAuthClient() {
  return new google.auth.OAuth2(
    process.env.GOOGLE_CLIENT_ID,
    process.env.GOOGLE_CLIENT_SECRET,
    process.env.OAUTH_REDIRECT_URI ||
      `https://${process.env.RAILWAY_PUBLIC_DOMAIN}/api/auth/google/callback`
  );
}

// ── AUTH ROUTES ──────────────────────────────────────────────────────────────

// Check if any users exist (for initial setup flow)
app.get('/api/auth/setup-status', async (req, res) => {
  try {
    await ensureUsersSheet();
    const count = await countUsers();
    res.json({ needsSetup: count === 0 });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// Create first admin (only allowed when no users exist)
app.post('/api/auth/setup', async (req, res) => {
  try {
    await ensureUsersSheet();
    const count = await countUsers();
    if (count > 0) return res.status(403).json({ error: 'Setup already complete' });

    const { email, password, name } = req.body;
    if (!email || !password || !name)
      return res.status(400).json({ error: 'email, password and name required' });

    const hash = await bcrypt.hash(password, 12);
    const user_id = 'u_' + Date.now();
    await createUser({
      user_id, email, password_hash: hash, name,
      role: 'admin', status: 'active',
      created_at: new Date().toISOString()
    });

    const token = signToken({ user_id, email, name, role: 'admin' });
    res.cookie(COOKIE_NAME, token, COOKIE_OPTS);
    res.json({ ok: true, user: { user_id, email, name, role: 'admin' } });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// Login
app.post('/api/auth/login', async (req, res) => {
  try {
    const { email, password } = req.body;
    if (!email || !password)
      return res.status(400).json({ error: 'email and password required' });

    const user = await getUserByEmail(email);
    if (!user) return res.status(401).json({ error: 'Invalid credentials' });
    if (user.status !== 'active')
      return res.status(403).json({ error: 'Account is not active' });

    const match = await bcrypt.compare(password, user.password_hash);
    if (!match) return res.status(401).json({ error: 'Invalid credentials' });

    const token = signToken({
      user_id: user.user_id, email: user.email,
      name: user.name, role: user.role
    });
    res.cookie(COOKIE_NAME, token, COOKIE_OPTS);
    res.json({ ok: true, user: { user_id: user.user_id, email: user.email, name: user.name, role: user.role } });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// Logout
app.post('/api/auth/logout', (req, res) => {
  res.clearCookie(COOKIE_NAME);
  res.json({ ok: true });
});

// Who am I
app.get('/api/auth/me', requireAuth, (req, res) => {
  res.json({ user: req.user });
});

// ── USER MANAGEMENT (Admin only) ─────────────────────────────────────────────

// List all users
app.get('/api/users', requireAuth, requireAdmin, async (req, res) => {
  try {
    const users = await getUsers();
    res.json(users.map(u => ({
      user_id: u.user_id, email: u.email,
      name: u.name, role: u.role, status: u.status, created_at: u.created_at
    })));
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// Add user
app.post('/api/users', requireAuth, requireAdmin, async (req, res) => {
  try {
    const { email, password, name, role } = req.body;
    if (!email || !password || !name)
      return res.status(400).json({ error: 'email, password and name required' });

    const existing = await getUserByEmail(email);
    if (existing) return res.status(409).json({ error: 'Email already registered' });

    const hash = await bcrypt.hash(password, 12);
    const user_id = 'u_' + Date.now();
    await createUser({
      user_id, email, password_hash: hash, name,
      role: role === 'admin' ? 'admin' : 'user',
      status: 'active',
      created_at: new Date().toISOString()
    });
    res.json({ ok: true, user_id });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// Edit user (name, email, role, optionally password)
app.put('/api/users/:id', requireAuth, requireAdmin, async (req, res) => {
  try {
    const { name, email, role, password } = req.body;
    const patch = {};
    if (name)  patch.name  = name;
    if (email) patch.email = email;
    if (role)  patch.role  = role === 'admin' ? 'admin' : 'user';
    if (password) patch.password_hash = await bcrypt.hash(password, 12);
    await updateUserRow(req.params.id, patch);
    res.json({ ok: true });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// Activate / deactivate
app.patch('/api/users/:id/status', requireAuth, requireAdmin, async (req, res) => {
  try {
    const { status } = req.body;
    if (!['active','inactive'].includes(status))
      return res.status(400).json({ error: 'status must be active or inactive' });
    await updateUserRow(req.params.id, { status });
    res.json({ ok: true });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ── DATA ROUTES ───────────────────────────────────────────────────────────────

// GET /api/data  →  user-scoped data
app.get('/api/data', requireAuth, async (req, res) => {
  try {
    const data = await getUserDataSafe(req.user.user_id, req.user.role === 'admin');
    res.json(data);
  } catch (e) {
    console.error('GET /api/data error:', e.message);
    res.status(500).json({ error: e.message });
  }
});

// POST /api/data  →  save user-scoped data
app.post('/api/data', requireAuth, async (req, res) => {
  try {
    const updates = req.body;
    if (!updates || typeof updates !== 'object' || Array.isArray(updates))
      return res.status(400).json({ error: 'Body must be a plain object' });
    await batchUpsertUser(req.user.user_id, updates);
    res.json({ ok: true });
  } catch (e) {
    console.error('POST /api/data error:', e.message);
    res.status(500).json({ error: e.message });
  }
});

// ── GOOGLE FIT / OAUTH ROUTES ─────────────────────────────────────────────────

// Start Google OAuth (user must be logged in)
app.get('/api/auth/google', requireAuth, (req, res) => {
  const oauth2 = getOAuthClient();
  const url = oauth2.generateAuthUrl({
    access_type: 'offline',
    prompt: 'consent',
    scope: [
      'https://www.googleapis.com/auth/fitness.activity.read',
      'https://www.googleapis.com/auth/fitness.sleep.read',
      'https://www.googleapis.com/auth/fitness.heart_rate.read',
    ],
    state: req.user.user_id
  });
  res.redirect(url);
});

// OAuth callback
app.get('/api/auth/google/callback', async (req, res) => {
  try {
    const { code, state: userId } = req.query;
    if (!code || !userId) return res.status(400).send('Missing code or state');

    const oauth2 = getOAuthClient();
    const { tokens } = await oauth2.getToken(code);
    await batchUpsertUser(userId, { google_fit_tokens: tokens });
    res.redirect('/?fit=connected');
  } catch (e) {
    console.error('OAuth callback error:', e.message);
    res.redirect('/?fit=error');
  }
});

// Google Fit connection status
app.get('/api/fit/status', requireAuth, async (req, res) => {
  try {
    const data = await getUserDataSafe(req.user.user_id, false);
    res.json({ connected: !!data.google_fit_tokens });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// Fetch Google Fit data
app.get('/api/fit/data', requireAuth, async (req, res) => {
  try {
    const data = await getUserDataSafe(req.user.user_id, false);
    const tokens = data.google_fit_tokens;
    if (!tokens) return res.status(400).json({ error: 'Google Fit not connected' });

    const oauth2 = getOAuthClient();
    oauth2.setCredentials(tokens);
    const fitness = google.fitness({ version: 'v1', auth: oauth2 });

    const now = Date.now();
    const weekAgo = now - 7 * 24 * 60 * 60 * 1000;

    const fitRes = await fitness.users.dataset.aggregate({
      userId: 'me',
      requestBody: {
        aggregateBy: [
          { dataTypeName: 'com.google.step_count.delta' },
          { dataTypeName: 'com.google.sleep.segment' },
        ],
        bucketByTime: { durationMillis: 86400000 },
        startTimeMillis: weekAgo,
        endTimeMillis: now,
      }
    });

    // Refresh tokens if changed
    if (oauth2.credentials.access_token !== tokens.access_token) {
      await batchUpsertUser(req.user.user_id, { google_fit_tokens: oauth2.credentials });
    }

    res.json(fitRes.data);
  } catch (e) {
    console.error('Fit data error:', e.message);
    res.status(500).json({ error: e.message });
  }
});

// Disconnect Google Fit
app.post('/api/auth/google/disconnect', requireAuth, async (req, res) => {
  try {
    await batchUpsertUser(req.user.user_id, { google_fit_tokens: null });
    res.json({ ok: true });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ── HEALTH CHECK ──────────────────────────────────────────────────────────────
app.get('/api/ping', (req, res) => res.json({ ok: true, ts: Date.now() }));

// ── SPA FALLBACK ──────────────────────────────────────────────────────────────
app.get('*', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// ── START ─────────────────────────────────────────────────────────────────────
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`✓ Life OS running on port ${PORT}`));
