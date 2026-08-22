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

// ── CONFIG ────────────────────────────────────────────────────────────────────
const SPREADSHEET_ID = process.env.SPREADSHEET_ID;
const SHEET_NAME     = process.env.SHEET_NAME || 'LifeOS';
const USERS_SHEET    = 'Users';
const JWT_SECRET     = process.env.JWT_SECRET;
const COOKIE_NAME    = 'lifeos_token';
const COOKIE_OPTS    = {
  httpOnly: true,
  secure:   process.env.NODE_ENV === 'production',
  sameSite: 'lax',
  maxAge:   7 * 24 * 60 * 60 * 1000, // 7 days
};

if (!SPREADSHEET_ID)  console.error('❌  SPREADSHEET_ID env var missing.');
if (!JWT_SECRET)      console.error('❌  JWT_SECRET env var missing — generate one and add to Railway.');
if (!process.env.GOOGLE_SERVICE_ACCOUNT_JSON)
  console.error('❌  GOOGLE_SERVICE_ACCOUNT_JSON env var missing.');

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

// ── USERS SHEET ───────────────────────────────────────────────────────────────
// Columns: A=user_id  B=name  C=email  D=contact  E=password_hash  F=role  G=status  H=created_at

async function ensureUsersSheet() {
  const sheets = await getSheetsClient();
  // Check if the Users sheet exists
  const meta = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
  const exists = meta.data.sheets.some(s => s.properties.title === USERS_SHEET);
  if (!exists) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: { requests: [{ addSheet: { properties: { title: USERS_SHEET } } }] },
    });
    // Write header row
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `${USERS_SHEET}!A1:H1`,
      valueInputOption: 'RAW',
      requestBody: { values: [['user_id','name','email','contact','password_hash','role','status','created_at']] },
    });
  }
}

async function getUsers() {
  await ensureUsersSheet();
  const sheets = await getSheetsClient();
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `${USERS_SHEET}!A:H`,
  });
  const rows = res.data.values || [];
  if (rows.length < 2) return [];
  return rows.slice(1).map(r => ({
    user_id:       r[0] || '',
    name:          r[1] || '',
    email:         r[2] || '',
    contact:       r[3] || '',
    password_hash: r[4] || '',
    role:          r[5] || 'user',
    status:        r[6] || 'active',
    created_at:    r[7] || '',
  })).filter(u => u.user_id);
}

async function getUserByEmail(email) {
  const users = await getUsers();
  return users.find(u => u.email.toLowerCase() === email.toLowerCase()) || null;
}

async function getUserById(user_id) {
  const users = await getUsers();
  return users.find(u => u.user_id === user_id) || null;
}

async function countUsers() {
  const users = await getUsers();
  return users.length;
}

async function createUser(userData) {
  await ensureUsersSheet();
  const sheets = await getSheetsClient();
  const row = [
    userData.user_id,
    userData.name,
    userData.email,
    userData.contact || '',
    userData.password_hash,
    userData.role || 'user',
    userData.status || 'active',
    userData.created_at || new Date().toISOString(),
  ];
  await sheets.spreadsheets.values.append({
    spreadsheetId: SPREADSHEET_ID,
    range: `${USERS_SHEET}!A:H`,
    valueInputOption: 'RAW',
    requestBody: { values: [row] },
  });
}

async function updateUserRow(user_id, patch) {
  await ensureUsersSheet();
  const sheets = await getSheetsClient();
  // Find the row number
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `${USERS_SHEET}!A:H`,
  });
  const rows = res.data.values || [];
  const rowIdx = rows.findIndex((r, i) => i > 0 && r[0] === user_id);
  if (rowIdx === -1) throw new Error('User not found');

  const existing = rows[rowIdx];
  const updated = [
    existing[0],                              // user_id
    patch.name      ?? existing[1] ?? '',
    existing[2],                              // email (immutable)
    patch.contact   ?? existing[3] ?? '',
    patch.password_hash ?? existing[4] ?? '',
    patch.role      ?? existing[5] ?? 'user',
    patch.status    ?? existing[6] ?? 'active',
    existing[7],                              // created_at (immutable)
  ];
  const sheetRow = rowIdx + 1; // 1-indexed
  await sheets.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID,
    range: `${USERS_SHEET}!A${sheetRow}:H${sheetRow}`,
    valueInputOption: 'RAW',
    requestBody: { values: [updated] },
  });
}

// ── JWT AUTH ──────────────────────────────────────────────────────────────────
function signToken(payload) {
  return jwt.sign(payload, JWT_SECRET, { expiresIn: '7d' });
}

function requireAuth(req, res, next) {
  try {
    const token = req.cookies[COOKIE_NAME];
    if (!token) return res.status(401).json({ error: 'Not authenticated. Please log in.' });
    const decoded = jwt.verify(token, JWT_SECRET);
    req.user = decoded;
    next();
  } catch (e) {
    res.clearCookie(COOKIE_NAME);
    res.status(401).json({ error: 'Session expired. Please log in again.' });
  }
}

function requireAdmin(req, res, next) {
  if (req.user?.role !== 'admin') {
    return res.status(403).json({ error: 'Access denied. Admin privileges required.' });
  }
  next();
}

// ── DATA HELPERS (user-scoped) ─────────────────────────────────────────────────
// User data keys are stored as "{userId}:{key}" in the LifeOS sheet.
// Admin users also see legacy keys (no prefix) for backward compatibility.

async function getUserData(userId, isAdmin) {
  const sheets = await getSheetsClient();
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `${SHEET_NAME}!A:B`,
  });
  const rows = res.data.values || [];
  const data = {};
  const prefix = userId + ':';

  for (const [key, val] of rows) {
    if (!key) continue;
    if (key.startsWith(prefix)) {
      // This user's data — strip the prefix
      const shortKey = key.slice(prefix.length);
      // Skip internal system keys
      if (shortKey === 'googleFitTokens') { data[shortKey] = tryParse(val); continue; }
      data[shortKey] = tryParse(val);
    } else if (isAdmin && !key.includes(':')) {
      // Legacy data (no prefix) visible only to admin, for backward compat
      // Only if this key doesn't already have a prefixed version
      if (!(prefix + key in rows.map(r => r[0]))) {
        data[key] = tryParse(val);
      }
    }
  }
  return data;
}

// More efficient version: build a set of prefixed keys to check legacy overlap
async function getUserDataSafe(userId, isAdmin) {
  const sheets = await getSheetsClient();
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `${SHEET_NAME}!A:B`,
  });
  const rows = res.data.values || [];
  const prefix = userId + ':';
  const data = {};
  const prefixedKeysFound = new Set();

  // First pass: collect all prefixed keys for this user
  for (const [key] of rows) {
    if (key && key.startsWith(prefix)) {
      prefixedKeysFound.add(key.slice(prefix.length));
    }
  }

  // Second pass: build data object
  for (const [key, val] of rows) {
    if (!key) continue;
    if (key.startsWith(prefix)) {
      data[key.slice(prefix.length)] = tryParse(val);
    } else if (isAdmin && !key.includes(':')) {
      // Legacy key: only include if no prefixed version exists yet
      if (!prefixedKeysFound.has(key)) {
        data[key] = tryParse(val);
      }
    }
  }
  return data;
}

function tryParse(val) {
  try { return JSON.parse(val); }
  catch (_) { return val ?? null; }
}

// ── BATCH UPSERT (user-scoped) ────────────────────────────────────────────────
async function batchUpsert(updates) {
  // Raw upsert — keys used as-is. Used internally only.
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
      toUpdate.push({ range: `${SHEET_NAME}!A${rowMap[key]}:B${rowMap[key]}`, values: [[key, jsonVal]] });
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
  // Prefix all keys with userId: before saving
  const prefixed = {};
  for (const [key, value] of Object.entries(updates)) {
    prefixed[`${userId}:${key}`] = value;
  }
  await batchUpsert(prefixed);
}

// ── GOOGLE FIT OAUTH ──────────────────────────────────────────────────────────
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

// ── ROUTES ─────────────────────────────────────────────────────────────────────

// ── SETUP (first-run admin creation) ─────────────────────────────────────────
// Only works when ZERO users exist. Creates the first admin account.
app.post('/api/auth/setup', async (req, res) => {
  try {
    const count = await countUsers();
    if (count > 0) return res.status(403).json({ error: 'Setup already complete. Admin account exists.' });

    const { name, email, password, contact } = req.body;
    if (!name || !email || !password)
      return res.status(400).json({ error: 'Name, email, and password are required.' });
    if (password.length < 8)
      return res.status(400).json({ error: 'Password must be at least 8 characters.' });

    const password_hash = await bcrypt.hash(password, 12);
    const user_id = 'u' + Date.now().toString(36).toUpperCase();
    const newUser = { user_id, name, email: email.toLowerCase(), contact: contact || '', password_hash, role: 'admin', status: 'active', created_at: new Date().toISOString() };
    await createUser(newUser);

    const token = signToken({ user_id, name, email: newUser.email, role: 'admin' });
    res.cookie(COOKIE_NAME, token, COOKIE_OPTS);
    res.json({ ok: true, user: { user_id, name, email: newUser.email, role: 'admin' } });
  } catch (e) {
    console.error('Setup error:', e.message);
    res.status(500).json({ error: e.message });
  }
});

// Check if setup is needed
app.get('/api/auth/setup-status', async (req, res) => {
  try {
    const count = await countUsers();
    res.json({ needsSetup: count === 0 });
  } catch (e) {
    res.json({ needsSetup: false });
  }
});

// ── LOGIN ──────────────────────────────────────────────────────────────────────
app.post('/api/auth/login', async (req, res) => {
  try {
    const { email, password } = req.body;
    if (!email || !password) return res.status(400).json({ error: 'Email and password are required.' });

    const user = await getUserByEmail(email);
    if (!user) return res.status(401).json({ error: 'Your email is not registered. Please contact the administrator.' });
    if (user.status !== 'active') return res.status(401).json({ error: 'Your account has been deactivated. Please contact the administrator.' });

    const valid = await bcrypt.compare(password, user.password_hash);
    if (!valid) return res.status(401).json({ error: 'Incorrect password.' });

    const token = signToken({ user_id: user.user_id, name: user.name, email: user.email, role: user.role });
    res.cookie(COOKIE_NAME, token, COOKIE_OPTS);
    res.json({ ok: true, user: { user_id: user.user_id, name: user.name, email: user.email, role: user.role } });
  } catch (e) {
    console.error('Login error:', e.message);
    res.status(500).json({ error: 'Login failed. Please try again.' });
  }
});

// ── LOGOUT ──────────────────────────────────────────────────────────────────────
app.post('/api/auth/logout', (req, res) => {
  res.clearCookie(COOKIE_NAME);
  res.json({ ok: true });
});

// ── CURRENT USER ──────────────────────────────────────────────────────────────
app.get('/api/auth/me', requireAuth, (req, res) => {
  res.json({ user: req.user });
});

// ── USER MANAGEMENT (admin only) ──────────────────────────────────────────────

// GET /api/users — list all users (admin)
app.get('/api/users', requireAuth, requireAdmin, async (req, res) => {
  try {
    const users = await getUsers();
    // Never return password_hash to the frontend
    res.json(users.map(u => ({ user_id: u.user_id, name: u.name, email: u.email, contact: u.contact, role: u.role, status: u.status, created_at: u.created_at })));
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// POST /api/users — add new user (admin)
app.post('/api/users', requireAuth, requireAdmin, async (req, res) => {
  try {
    const { name, email, password, contact, role } = req.body;
    if (!name || !email || !password) return res.status(400).json({ error: 'Name, email, and password are required.' });
    if (password.length < 8) return res.status(400).json({ error: 'Password must be at least 8 characters.' });

    const existing = await getUserByEmail(email);
    if (existing) return res.status(409).json({ error: 'A user with this email already exists.' });

    const password_hash = await bcrypt.hash(password, 12);
    const user_id = 'u' + Date.now().toString(36).toUpperCase();
    const newUser = { user_id, name, email: email.toLowerCase(), contact: contact || '', password_hash, role: role === 'admin' ? 'admin' : 'user', status: 'active', created_at: new Date().toISOString() };
    await createUser(newUser);
    res.json({ ok: true, user: { user_id, name, email: newUser.email, role: newUser.role, status: 'active' } });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// PUT /api/users/:id — edit user (admin)
app.put('/api/users/:id', requireAuth, requireAdmin, async (req, res) => {
  try {
    const { id } = req.params;
    const { name, contact, role, status, password } = req.body;

    // Prevent admin from removing their own admin role
    if (id === req.user.user_id && role === 'user') {
      return res.status(400).json({ error: 'You cannot remove your own admin role.' });
    }

    const patch = {};
    if (name)    patch.name    = name;
    if (contact !== undefined) patch.contact = contact;
    if (role)    patch.role    = role === 'admin' ? 'admin' : 'user';
    if (status)  patch.status  = status === 'active' ? 'active' : 'inactive';
    if (password) {
      if (password.length < 8) return res.status(400).json({ error: 'Password must be at least 8 characters.' });
      patch.password_hash = await bcrypt.hash(password, 12);
    }

    await updateUserRow(id, patch);
    res.json({ ok: true });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// PATCH /api/users/:id/status — activate or deactivate (admin)
app.patch('/api/users/:id/status', requireAuth, requireAdmin, async (req, res) => {
  try {
    const { id } = req.params;
    const { status } = req.body;
    if (id === req.user.user_id && status === 'inactive') {
      return res.status(400).json({ error: 'You cannot deactivate your own account.' });
    }
    await updateUserRow(id, { status: status === 'active' ? 'active' : 'inactive' });
    res.json({ ok: true });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ── GOOGLE FIT ROUTES (require auth) ─────────────────────────────────────────

app.get('/api/auth/google', requireAuth, (req, res) => {
  if (!GOOGLE_CLIENT_ID || !GOOGLE_CLIENT_SECRET) {
    return res.status(500).json({ error: 'GOOGLE_CLIENT_ID or GOOGLE_CLIENT_SECRET not set' });
  }
  const oauth2Client = getOAuth2Client();
  const url = oauth2Client.generateAuthUrl({
    access_type: 'offline',
    scope: FITNESS_SCOPES,
    prompt: 'consent',
    state: req.user.user_id, // pass user_id through OAuth flow
  });
  res.redirect(url);
});

app.get('/api/auth/google/callback', async (req, res) => {
  const { code, error, state: userId } = req.query;
  if (error) return res.redirect('/?gfit=error');
  try {
    const oauth2Client = getOAuth2Client();
    const { tokens } = await oauth2Client.getToken(code);
    // Store tokens scoped to the user who initiated the flow
    if (userId) {
      await batchUpsertUser(userId, { googleFitTokens: tokens });
    }
    res.redirect('/?gfit=connected');
  } catch (e) {
    console.error('OAuth callback error:', e.message);
    res.redirect('/?gfit=error');
  }
});

app.get('/api/fit/status', requireAuth, async (req, res) => {
  try {
    const data = await getUserDataSafe(req.user.user_id, req.user.role === 'admin');
    const tokens = data.googleFitTokens;
    res.json({ connected: !!(tokens && tokens.access_token) });
  } catch (e) {
    res.json({ connected: false });
  }
});

app.get('/api/fit/data', requireAuth, async (req, res) => {
  try {
    const data = await getUserDataSafe(req.user.user_id, req.user.role === 'admin');
    const tokens = data.googleFitTokens;
    if (!tokens) return res.status(401).json({ error: 'Not connected to Google Fit' });

    const oauth2Client = getOAuth2Client();
    oauth2Client.setCredentials(tokens);
    oauth2Client.on('tokens', async (newTokens) => {
      await batchUpsertUser(req.user.user_id, { googleFitTokens: { ...tokens, ...newTokens } });
    });

    const fitness = google.fitness({ version: 'v1', auth: oauth2Client });
    const now = Date.now();
    const startOfDay = new Date(); startOfDay.setHours(0, 0, 0, 0);

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
      await batchUpsertUser(req.user.user_id, { googleFitTokens: null });
    }
    res.status(500).json({ error: e.message });
  }
});

app.post('/api/auth/google/disconnect', requireAuth, async (req, res) => {
  try {
    await batchUpsertUser(req.user.user_id, { googleFitTokens: null });
    res.json({ ok: true });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ── DATA ROUTES (require auth, user-scoped) ────────────────────────────────────

// GET /api/data → returns ONLY the authenticated user's data
app.get('/api/data', requireAuth, async (req, res) => {
  try {
    const data = await getUserDataSafe(req.user.user_id, req.user.role === 'admin');
    res.json(data);
  } catch (e) {
    console.error('GET /api/data error:', e.message);
    res.status(500).json({ error: e.message });
  }
});

// POST /api/data → saves data scoped to the authenticated user
app.post('/api/data', requireAuth, async (req, res) => {
  try {
    const updates = req.body;
    if (!updates || typeof updates !== 'object' || Array.isArray(updates)) {
      return res.status(400).json({ error: 'Body must be a plain object of { key: value }' });
    }
    await batchUpsertUser(req.user.user_id, updates);
    res.json({ ok: true });
  } catch (e) {
    console.error('POST /api/data error:', e.message);
    res.status(500).json({ error: e.message });
  }
});

// ── HEALTH CHECK ───────────────────────────────────────────────────────────────
app.get('/api/ping', (req, res) => res.json({ ok: true, ts: Date.now() }));

// ── SPA FALLBACK ───────────────────────────────────────────────────────────────
app.get('*', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// ── START ──────────────────────────────────────────────────────────────────────
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`✅  Life OS running on port ${PORT}`));
