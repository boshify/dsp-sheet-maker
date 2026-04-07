require('dotenv').config();
const express = require('express');
const bodyParser = require('body-parser');
const { google } = require('googleapis');
const { Pool } = require('pg');
const { buildBriefRequests } = require('./sheet-renderer');

const app = express();

const BODY_SIZE_LIMIT = process.env.BODY_SIZE_LIMIT || '10mb';
app.use(bodyParser.json({ limit: BODY_SIZE_LIMIT }));
app.use(bodyParser.text({ limit: BODY_SIZE_LIMIT, type: ['text/plain'] }));

const APP_BASE_URL = process.env.APP_BASE_URL;

// --- Postgres pool ---
const pool = new Pool({ connectionString: process.env.DATABASE_URL });

// ------------------------------------------------------------
// DB init
// ------------------------------------------------------------
async function initDb() {
  await pool.query(`
    CREATE TABLE IF NOT EXISTS user_tokens (
      id SERIAL PRIMARY KEY,
      user_key TEXT UNIQUE NOT NULL,
      refresh_token TEXT NOT NULL,
      created_at TIMESTAMPTZ DEFAULT NOW(),
      updated_at TIMESTAMPTZ DEFAULT NOW()
    );
  `);
  await pool.query(`
    CREATE TABLE IF NOT EXISTS error_logs (
      id SERIAL PRIMARY KEY,
      user_key TEXT,
      path TEXT NOT NULL,
      message TEXT NOT NULL,
      stack TEXT,
      meta JSONB,
      created_at TIMESTAMPTZ DEFAULT NOW()
    );
  `);
  console.log('Database tables ensured.');
}

async function logError(userKey, path, err, meta = null) {
  try {
    const message = err?.message || String(err);
    const stack = err?.stack || null;
    const result = await pool.query(
      `INSERT INTO error_logs (user_key, path, message, stack, meta)
       VALUES ($1, $2, $3, $4, $5) RETURNING id;`,
      [userKey || null, path, message, stack, meta ? JSON.stringify(meta) : null]
    );
    const errorId = result.rows[0].id;
    console.error(`Logged error #${errorId}:`, message);
    return errorId;
  } catch (logErr) {
    console.error('Failed to log error:', logErr);
    return null;
  }
}

// --- OAuth2 client ---
const oauth2Client = new google.auth.OAuth2(
  process.env.GOOGLE_CLIENT_ID,
  process.env.GOOGLE_CLIENT_SECRET,
  `${APP_BASE_URL}/oauth2/callback/google`
);

// ------------------------------------------------------------
// Health check
// ------------------------------------------------------------
app.get('/', (req, res) => {
  res.send('DSP Sheet Maker backend is running (Railway, DB, logging, modular).');
});

// ------------------------------------------------------------
// Connect Google (OAuth start)
// ------------------------------------------------------------
app.get('/connect/google', (req, res) => {
  const scopes = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive.file',
    'https://www.googleapis.com/auth/userinfo.email',
    'https://www.googleapis.com/auth/userinfo.profile'
  ];
  const url = oauth2Client.generateAuthUrl({
    access_type: 'offline',
    prompt: 'consent',
    scope: scopes
  });
  res.redirect(url);
});

async function storeUserToken(userKey, tokens) {
  if (!tokens.refresh_token) {
    throw new Error('No refresh_token provided from Google');
  }
  await pool.query(
    `INSERT INTO user_tokens (user_key, refresh_token, created_at, updated_at)
     VALUES ($1, $2, NOW(), NOW())
     ON CONFLICT (user_key)
     DO UPDATE SET refresh_token = EXCLUDED.refresh_token, updated_at = NOW();`,
    [userKey, tokens.refresh_token]
  );
}

// ------------------------------------------------------------
// OAuth callback
// ------------------------------------------------------------
app.get('/oauth2/callback/google', async (req, res) => {
  const code = req.query.code;
  if (!code) return res.status(400).send('Missing "code" in query params.');

  try {
    const { tokens } = await oauth2Client.getToken(code);
    oauth2Client.setCredentials(tokens);

    const oauth2 = google.oauth2({ auth: oauth2Client, version: 'v2' });
    const { data: userInfo } = await oauth2.userinfo.get();
    const email = userInfo.email;

    if (!email) {
      return res.status(500).send('Could not determine your Google account email.');
    }
    if (!tokens.refresh_token) {
      return res.send(`
        <h1>Already connected</h1>
        <p>Google did not return a new refresh token.</p>
        <p>Revoke the app in your Google Account Permissions and try again.</p>
      `);
    }

    await storeUserToken(email, tokens);
    console.log('Stored tokens for user:', email);

    res.send(`
      <h1>Connected!</h1>
      <p>Your Google account <b>${email}</b> is now linked.</p>
      <p>Use this as your <code>userKey</code> in n8n: <code>${email}</code></p>
      <p>You can close this tab.</p>
    `);
  } catch (err) {
    const errorId = await logError(null, '/oauth2/callback/google', err, { query: req.query });
    res.status(500).send(`
      <h1>OAuth error</h1>
      <p>Error ID: <code>${errorId || 'unknown'}</code></p>
    `);
  }
});

// ------------------------------------------------------------
// Helper: auth client for a given userKey
// ------------------------------------------------------------
async function getAuthClientForUser(userKey) {
  const { rows } = await pool.query(
    'SELECT * FROM user_tokens WHERE user_key = $1',
    [userKey]
  );
  if (!rows.length) throw new Error(`No stored tokens for userKey: ${userKey}`);
  const row = rows[0];
  const client = new google.auth.OAuth2(
    process.env.GOOGLE_CLIENT_ID,
    process.env.GOOGLE_CLIENT_SECRET,
    `${APP_BASE_URL}/oauth2/callback/google`
  );
  client.setCredentials({ refresh_token: row.refresh_token });
  return client;
}

// ------------------------------------------------------------
// Resolve the numeric sheetId (gid) for a given sheet.
// Accepts an explicit gid, or looks up by sheetName. If neither exists,
// creates a new sheet with the given name and returns its gid.
// ------------------------------------------------------------
async function resolveSheetId(sheets, spreadsheetId, sheetName, providedSheetId) {
  const resp = await sheets.spreadsheets.get({
    spreadsheetId,
    fields: 'sheets.properties(sheetId,title)'
  });
  const existing = resp.data.sheets || [];

  // 1. Match by provided gid
  if (providedSheetId != null && providedSheetId !== '') {
    const gid = Number(providedSheetId);
    const found = existing.find(s => s.properties.sheetId === gid);
    if (found) return found.properties.sheetId;
  }

  // 2. Match by name
  const byName = existing.find(s => s.properties.title === sheetName);
  if (byName) return byName.properties.sheetId;

  // 3. Create a new sheet
  const addResp = await sheets.spreadsheets.batchUpdate({
    spreadsheetId,
    requestBody: {
      requests: [{ addSheet: { properties: { title: sheetName } } }]
    }
  });
  const added = addResp.data.replies[0].addSheet.properties;
  return added.sheetId;
}

// ------------------------------------------------------------
// POST /render-sheet — the main endpoint
// Body: array of jobs OR single job. Shape:
// {
//   spreadsheetId, sheetName, sheetId?, userKey,
//   meta, tables, notes
// }
// ------------------------------------------------------------
app.post('/render-sheet', async (req, res) => {
  let body = req.body;
  if (!body) return res.status(400).json({ error: 'Missing request body' });

  // Accept either a single job or an array of jobs; also accept a top-level
  // { userKey, jobs: [...] } shape for convenience.
  let jobs;
  let userKey;

  if (Array.isArray(body)) {
    jobs = body;
    userKey = body[0]?.userKey;
  } else if (Array.isArray(body.jobs)) {
    jobs = body.jobs;
    userKey = body.userKey || jobs[0]?.userKey;
  } else {
    jobs = [body];
    userKey = body.userKey;
  }

  if (!userKey) {
    return res.status(400).json({
      error: 'userKey is required (the email used when you connected via /connect/google)'
    });
  }

  try {
    const auth = await getAuthClientForUser(userKey);
    const sheets = google.sheets({ version: 'v4', auth });

    const results = [];
    for (const job of jobs) {
      if (!job || !job.spreadsheetId || !job.sheetName) {
        throw new Error('Each job requires spreadsheetId and sheetName');
      }
      if (!job.meta || !job.tables) {
        throw new Error('Each job requires meta and tables');
      }

      const gid = await resolveSheetId(
        sheets, job.spreadsheetId, job.sheetName, job.sheetId
      );

      const { requests } = buildBriefRequests(gid, job);

      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: job.spreadsheetId,
        requestBody: { requests }
      });

      results.push({
        spreadsheetId: job.spreadsheetId,
        sheetName: job.sheetName,
        sheetId: gid,
        requestsApplied: requests.length
      });
    }

    return res.json({ status: 'success', results });
  } catch (err) {
    const errorId = await logError(userKey, '/render-sheet', err, {
      jobCount: jobs?.length || 0,
      firstSpreadsheetId: jobs?.[0]?.spreadsheetId || null,
      googleError: err?.response?.data || null
    });
    return res.status(500).json({
      error: 'Internal server error',
      errorId,
      detail: err?.message || 'An error occurred while rendering the sheet.'
    });
  }
});

// ------------------------------------------------------------
// Start the server AFTER DB is initialized
// ------------------------------------------------------------
const PORT = process.env.PORT || 3000;

(async () => {
  try {
    await initDb();
    console.log('Database initialized.');
    app.listen(PORT, () => {
      console.log(`DSP Sheet Maker listening on port ${PORT}`);
    });
  } catch (err) {
    console.error('Failed to initialize database:', err);
    process.exit(1);
  }
})();

module.exports = {};
