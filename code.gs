/**
 * AI Community Stamp Journey — Google Apps Script Web App
 * Version: Template Builder edition (v5)
 *
 * New in this version:
 *   - Flexible stamp-track template: admin sets total stamp count and
 *     any number of reward milestones (position, custom name, icon).
 *   - Config persisted in a new `Config` sheet (single JSON row).
 *   - Reward logic is now fully data-driven (no hardcoded 5/10 sbux/gadget).
 *   - Admin can bulk-create N placeholder sessions in one click.
 *   - Admin can register new admin accounts from the UI.
 *   - Backward compatible: if no Config row exists yet, falls back to the
 *     classic 2-milestone (5 stamps / 10 stamps) defaults.
 */

const SHEETS = {
  users:       'Users',
  sessions:    'Sessions',
  stamps:      'StampRecords',
  redemptions: 'Redemptions',
  admins:      'Admins',
  config:      'Config'
};

const FALLBACK_SECRET = 'ai-community-stamp-secret-2025';

const RANDOM_ICONS = [
  'ti-brain','ti-tools','ti-users','ti-layout-dashboard',
  'ti-shield-check','ti-rocket','ti-sparkles','ti-pencil',
  'ti-chart-bar','ti-confetti','ti-bulb','ti-star','ti-award',
  'ti-bolt','ti-flask','ti-cpu','ti-map','ti-microscope'
];

const SESSION_LOCATIONS = ['Building1', 'Building2', 'Building3'];

function isValidSessionLocation(location) {
  return SESSION_LOCATIONS.indexOf(String(location || '').trim()) !== -1;
}

// Default template used only if the Config sheet has no saved row yet.
const DEFAULT_TEMPLATE = {
  totalStamps: 10,
  milestones: [
    { id: 'm5',  position: 5,  name: 'Starbucks Card', icon: 'ti-coffee' },
    { id: 'm10', position: 10, name: 'Free Gadget',    icon: 'ti-device-laptop' }
  ]
};

// ─────────────────────────────────────────────────────────────
// HTTP ENTRY POINTS
// ─────────────────────────────────────────────────────────────

function doGet(e) {
  return HtmlService
    .createHtmlOutputFromFile('Index')
    .setTitle('AI Community Stamp Journey')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  try {
    const payload = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    return jsonResponse(processRequest(payload));
  } catch (err) {
    return jsonResponse({ ok: false, message: safeMessage(err) });
  }
}

// ─────────────────────────────────────────────────────────────
// DISPATCHER
// ─────────────────────────────────────────────────────────────

function processRequest(payload) {
  try {
    if (!payload || typeof payload !== 'object') return { ok: false, message: 'Invalid request payload.' };
    const action = String(payload.action || '').trim();
    if (!action) return { ok: false, message: 'No action specified.' };

    if (action === 'ping')                  return { ok: true, message: 'API is running.' };
    if (action === 'login')                 return login(payload);
    if (action === 'signup')                return signup(payload);
    if (action === 'getProfile')            return getProfile(payload);
    if (action === 'scanQR')                return scanQR(payload);
    if (action === 'generateQRToken')       return generateQRToken(payload);
    if (action === 'deactivateQR')          return deactivateQR(payload);
    if (action === 'getAdminParticipants')  return getAdminParticipants(payload);
    if (action === 'markRedeemed')          return markRedeemed(payload);
    if (action === 'updateSession')         return updateSession(payload);
    if (action === 'createSession')         return createSession(payload);
    if (action === 'deleteSession')         return deleteSession(payload);
    if (action === 'getTemplateConfig')     return getTemplateConfigAction(payload);
    if (action === 'saveTemplateConfig')    return saveTemplateConfig(payload);
    if (action === 'bulkCreateSessions')    return bulkCreateSessions(payload);
    if (action === 'registerAdmin')         return registerAdmin(payload);

    return { ok: false, message: 'Unknown action: ' + action };
  } catch (err) {
    console.error('processRequest error [' + (payload && payload.action) + ']:', err);
    return { ok: false, message: safeMessage(err) };
  }
}

// ─────────────────────────────────────────────────────────────
// UTILITIES
// ─────────────────────────────────────────────────────────────

function safeMessage(err) { return (err && (err.message || String(err))) || 'Unknown error.'; }

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data || { ok: false, message: 'Empty response.' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function ss() { return SpreadsheetApp.getActiveSpreadsheet(); }

function getSheet(name) {
  const sh = ss().getSheetByName(name);
  if (!sh) throw new Error('Sheet "' + name + '" not found. Run setupInitialSheets() first.');
  return sh;
}

function getSheetSafe(name) {
  return ss().getSheetByName(name) || null;
}

function getRows(sheetName) {
  const sh = getSheet(sheetName);
  const lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];
  const values = sh.getRange(1, 1, lastRow, lastCol).getValues();
  if (values.length < 2) return [];
  const headers = values[0].map(String);
  return values.slice(1)
    .filter(row => row.some(v => v !== '' && v !== null && v !== undefined))
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = row[i] !== undefined ? row[i] : ''; });
      return obj;
    });
}

function appendRow(sheetName, obj) {
  const sh = getSheet(sheetName);
  const lastCol = sh.getLastColumn();
  if (lastCol < 1) throw new Error('Sheet "' + sheetName + '" has no headers.');
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  sh.appendRow(headers.map(h => obj[h] !== undefined ? obj[h] : ''));
}

function updateRowWhere(sheetName, matchFn, updateFn) {
  const sh = getSheet(sheetName);
  const lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return false;
  const all = sh.getRange(1, 1, lastRow, lastCol).getValues();
  const headers = all[0].map(String);
  let updated = false;
  for (let i = 1; i < all.length; i++) {
    const obj = {};
    headers.forEach((h, j) => { obj[h] = all[i][j]; });
    if (matchFn(obj)) {
      updateFn(obj);
      sh.getRange(i + 1, 1, 1, headers.length).setValues([headers.map(h => obj[h] !== undefined ? obj[h] : '')]);
      updated = true;
    }
  }
  return updated;
}

function nowText() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'd MMM yyyy HH:mm:ss');
}
function nowMs() { return new Date().getTime(); }
function newId(prefix) { return (prefix || 'id') + '_' + Utilities.getUuid().slice(0, 8); }
function cleanEmail(email) { return String(email || '').trim().toLowerCase(); }
function randomIcon() { return RANDOM_ICONS[Math.floor(Math.random() * RANDOM_ICONS.length)]; }

// ─────────────────────────────────────────────────────────────
// HMAC TOKEN HELPERS
// ─────────────────────────────────────────────────────────────

function getHmacSecret() {
  return PropertiesService.getScriptProperties().getProperty('HMAC_SECRET') || FALLBACK_SECRET;
}
function setHmacSecret() {
  PropertiesService.getScriptProperties().setProperty('HMAC_SECRET', 'ai-stamp-' + Utilities.getUuid());
  Logger.log('HMAC_SECRET set.');
}
function createQRToken(sessionId, expiresAtMs) {
  const nonce = Utilities.getUuid().slice(0, 12);
  const payload = JSON.stringify({ sid: Number(sessionId), exp: expiresAtMs, nonce });
  const payloadB64 = Utilities.base64EncodeWebSafe(payload);
  const mac = Utilities.computeHmacSha256Signature(payloadB64, getHmacSecret());
  return payloadB64 + '.' + Utilities.base64EncodeWebSafe(mac);
}
function verifyQRToken(token) {
  try {
    if (!token || typeof token !== 'string') return { ok: false, error: 'Empty token.' };
    const parts = token.split('.');
    if (parts.length !== 2) return { ok: false, error: 'Malformed token.' };
    const [payloadB64, macB64] = parts;
    const expectedMac = Utilities.base64EncodeWebSafe(
      Utilities.computeHmacSha256Signature(payloadB64, getHmacSecret())
    );
    if (expectedMac !== macB64) return { ok: false, error: 'Invalid QR code.' };
    const payload = JSON.parse(Utilities.newBlob(Utilities.base64DecodeWebSafe(payloadB64)).getDataAsString());
    if (nowMs() > Number(payload.exp)) return { ok: false, error: 'This QR code has expired. Ask the facilitator for a new one.' };
    return { ok: true, sessionId: Number(payload.sid), expiresAtMs: Number(payload.exp) };
  } catch (err) {
    return { ok: false, error: 'Could not read QR code: ' + safeMessage(err) };
  }
}

// ─────────────────────────────────────────────────────────────
// ADMIN CHECK
// ─────────────────────────────────────────────────────────────

function isAdminEmail(email) {
  try {
    const target = cleanEmail(email);
    if (!target) return false;
    const user = getRows(SHEETS.users).find(u => cleanEmail(u.email) === target);
    return Boolean(user && String(user.role || '').trim().toLowerCase() === 'admin');
  } catch (err) { return false; }
}
function requireAdmin(email) {
  if (!isAdminEmail(email)) throw new Error('Admin access required for "' + cleanEmail(email) + '".');
}

// ─────────────────────────────────────────────────────────────
// USER / SESSION HELPERS
// ─────────────────────────────────────────────────────────────

function findUserByEmail(email) {
  return getRows(SHEETS.users).find(u => cleanEmail(u.email) === cleanEmail(email)) || null;
}
function findUserById(userId) {
  if (!userId) return null;
  return getRows(SHEETS.users).find(u => String(u.user_id) === String(userId)) || null;
}

function getActiveSessions() {
  return getRows(SHEETS.sessions).filter(s => String(s.is_active).toLowerCase() !== 'false');
}

function getSessionsForClient(includeAdminFields) {
  return getActiveSessions().map(s => {
    const obj = {
      id:       Number(s.session_id),
      name:     String(s.session_name || ''),
      type:     String(s.session_type || 'Sharing'),
      desc:     String(s.description  || ''),
      icon:     String(s.icon         || 'ti-brain'),
      location: String(s.location     || ''),
      time:     String(s.time         || '')
    };
    if (includeAdminFields) {
      obj.qrActive    = Boolean(s.qr_active && String(s.qr_active).toLowerCase() !== 'false' && s.qr_active !== '');
      obj.qrExpiresAt = s.qr_expires_at ? Number(s.qr_expires_at) : 0;
      obj.qrToken     = s.qr_token ? String(s.qr_token) : '';
    }
    return obj;
  });
}

function getStampRecordsForUser(userId) {
  return getRows(SHEETS.stamps).filter(r => String(r.user_id) === String(userId));
}
function getRedemptionsForUser(userId) {
  return getRows(SHEETS.redemptions).filter(r => String(r.user_id) === String(userId));
}

// ─────────────────────────────────────────────────────────────
// TEMPLATE CONFIG  (flexible stamp count + reward milestones)
// ─────────────────────────────────────────────────────────────

/**
 * Reads the saved template config from the Config sheet.
 * Falls back to DEFAULT_TEMPLATE if nothing has been saved yet.
 * Returns { totalStamps, milestones: [{id,position,name,icon}] }
 */
function getTemplateConfig() {
  const sh = getSheetSafe(SHEETS.config);
  if (!sh) return DEFAULT_TEMPLATE;
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return DEFAULT_TEMPLATE;
  const raw = sh.getRange(2, 1).getValue();
  if (!raw) return DEFAULT_TEMPLATE;
  try {
    const parsed = JSON.parse(String(raw));
    if (!parsed || !Array.isArray(parsed.milestones) || !parsed.totalStamps) return DEFAULT_TEMPLATE;
    // Normalise + sort by position, ensure each milestone has an id.
    parsed.milestones = parsed.milestones
      .map(m => ({
        id:       String(m.id || ('m' + m.position)),
        position: Number(m.position),
        name:     String(m.name || 'Reward'),
        icon:     String(m.icon || 'ti-gift')
      }))
      .filter(m => m.position > 0 && m.position <= parsed.totalStamps)
      .sort((a, b) => a.position - b.position);
    parsed.totalStamps = Number(parsed.totalStamps);
    return parsed;
  } catch (err) {
    console.error('getTemplateConfig parse error:', err);
    return DEFAULT_TEMPLATE;
  }
}

function getTemplateConfigAction(payload) {
  return { ok: true, config: getTemplateConfig() };
}

/**
 * Save a new template config.
 * payload: { adminEmail, totalStamps, milestones: [{position,name,icon}] }
 */
function saveTemplateConfig(payload) {
  requireAdmin(cleanEmail(payload.adminEmail || ''));

  const totalStamps = Math.max(1, Math.min(100, Number(payload.totalStamps) || 0));
  if (!totalStamps) return { ok: false, message: 'Total stamps must be a positive number.' };

  let milestones = Array.isArray(payload.milestones) ? payload.milestones : [];
  if (!milestones.length) return { ok: false, message: 'Add at least one reward milestone.' };

  // Validate + normalise each milestone.
  const seenPositions = {};
  milestones = milestones.map((m, idx) => {
    const position = Number(m.position);
    const name = String(m.name || '').trim();
    const icon = String(m.icon || 'ti-gift').trim() || 'ti-gift';
    if (!position || position < 1 || position > totalStamps) {
      throw new Error('Milestone ' + (idx + 1) + ': position must be between 1 and ' + totalStamps + '.');
    }
    if (!name) throw new Error('Milestone ' + (idx + 1) + ': name is required.');
    if (seenPositions[position]) throw new Error('Two milestones cannot share stamp position ' + position + '.');
    seenPositions[position] = true;
    return { id: 'm' + position, position, name, icon };
  }).sort((a, b) => a.position - b.position);

  const config = { totalStamps, milestones };

  let sh = getSheetSafe(SHEETS.config);
  if (!sh) {
    sh = ss().insertSheet(SHEETS.config);
    sh.getRange(1, 1, 1, 2).setValues([['config_json', 'updated_at']]);
    sh.setFrozenRows(1);
  }
  const lastRow = sh.getLastRow();
  const rowValues = [JSON.stringify(config), nowText()];
  if (lastRow < 2) {
    sh.getRange(2, 1, 1, 2).setValues([rowValues]);
  } else {
    sh.getRange(2, 1, 1, 2).setValues([rowValues]);
  }

  return { ok: true, config: config, message: 'Template saved.' };
}

// ─────────────────────────────────────────────────────────────
// USER PROFILE  (rewards now built dynamically from template config)
// ─────────────────────────────────────────────────────────────

function buildUserProfile(user) {
  if (!user) throw new Error('buildUserProfile: user is null.');
  const stampRecords = getStampRecordsForUser(user.user_id);
  const claimed      = stampRecords.map(r => Number(r.session_id)).filter(n => !isNaN(n) && n > 0);
  const redemptions  = getRedemptionsForUser(user.user_id);
  const scanTimestamps = {};
  stampRecords.forEach(r => { if (r.session_id) scanTimestamps[Number(r.session_id)] = String(r.claimed_at || ''); });

  const template = getTemplateConfig();
  const stampCount = claimed.length;

  const rewards = template.milestones.map(m => {
    const redemption = redemptions.find(r => String(r.reward_id || r.reward_type) === m.id) || null;
    return {
      id:         m.id,
      position:   m.position,
      name:       m.name,
      icon:       m.icon,
      eligible:   stampCount >= m.position,
      redeemed:   Boolean(redemption),
      redeemedAt: redemption ? String(redemption.redeemed_at || '') : null
    };
  });

  return {
    id:              String(user.user_id),
    name:            String(user.full_name || ''),
    facility:        String(user.facility  || ''),
    email:           cleanEmail(user.email),
    role:            String(user.role || 'participant').trim().toLowerCase(),
    stamps:          stampCount,
    totalStamps:     template.totalStamps,
    claimedSessions: claimed,
    scanTimestamps:  scanTimestamps,
    rewards:         rewards
  };
}

// ─────────────────────────────────────────────────────────────
// AUTH
// ─────────────────────────────────────────────────────────────

function login(payload) {
  const email = cleanEmail(payload.email), password = String(payload.password || '');
  if (!email || !password) return { ok: false, message: 'Email and password are required.' };
  const user = findUserByEmail(email);
  if (!user || String(user.password) !== password) return { ok: false, message: 'Incorrect email or password.' };
  const profile = buildUserProfile(user);
  return { ok: true, user: profile, sessions: getSessionsForClient(profile.role === 'admin') };
}

function signup(payload) {
  const name     = String(payload.name     || '').trim();
  const facility = String(payload.facility || '').trim();
  const email    = cleanEmail(payload.email);
  const password = String(payload.password || '');
  if (!name)                          return { ok: false, message: 'Full name is required.' };
  if (!facility)                      return { ok: false, message: 'Facility is required.' };
  if (!email || !email.includes('@')) return { ok: false, message: 'Valid email is required.' };
  if (password.length < 8)            return { ok: false, message: 'Password must be at least 8 characters.' };
  if (findUserByEmail(email))         return { ok: false, message: 'An account with this email already exists.' };
  const newUser = { user_id: newId('u'), full_name: name, facility, email, password, role: 'participant', created_at: nowText() };
  appendRow(SHEETS.users, newUser);
  return { ok: true, user: buildUserProfile(newUser), sessions: getSessionsForClient(false) };
}

function getProfile(payload) {
  const user = findUserById(payload.userId) || findUserByEmail(payload.email);
  if (!user) return { ok: false, message: 'User not found.' };
  const profile = buildUserProfile(user);
  return { ok: true, user: profile, sessions: getSessionsForClient(profile.role === 'admin') };
}

// ─────────────────────────────────────────────────────────────
// ADMIN REGISTRATION
// ─────────────────────────────────────────────────────────────

/**
 * Registers a new admin account.
 * payload: { adminEmail (requesting admin), name, email, password }
 */
function registerAdmin(payload) {
  requireAdmin(cleanEmail(payload.adminEmail || ''));

  const name     = String(payload.name || '').trim();
  const email    = cleanEmail(payload.email);
  const password = String(payload.password || '');

  if (!name)                          return { ok: false, message: 'Full name is required.' };
  if (!email || !email.includes('@')) return { ok: false, message: 'Valid email is required.' };
  if (password.length < 8)            return { ok: false, message: 'Password must be at least 8 characters.' };
  if (findUserByEmail(email))          return { ok: false, message: 'An account with this email already exists.' };

  appendRow(SHEETS.users, {
    user_id:    newId('a'),
    full_name:  name,
    facility:   'Admin',
    email:      email,
    password:   password,
    role:       'admin',
    created_at: nowText()
  });

  return { ok: true, message: 'Admin account created for ' + email + '.' };
}

// ─────────────────────────────────────────────────────────────
// QR STAMP CLAIMING
// ─────────────────────────────────────────────────────────────

function scanQR(payload) {
  const user = findUserById(payload.userId) || findUserByEmail(payload.email);
  if (!user) return { ok: false, message: 'User not found. Please sign in again.' };
  if (String(user.role || '').toLowerCase() === 'admin') return { ok: false, message: 'Admin accounts cannot collect stamps.' };

  const token = String(payload.qrToken || '').trim();
  if (!token) return { ok: false, message: 'No QR token received.' };

  const verify = verifyQRToken(token);
  if (!verify.ok) return { ok: false, message: verify.error };

  const sessionId = verify.sessionId;
  const session   = getActiveSessions().find(s => Number(s.session_id) === sessionId);
  if (!session) return { ok: false, message: 'Session not found or no longer active.' };
  if (!session.qr_active || String(session.qr_active).toLowerCase() === 'false' || session.qr_active === '') {
    return { ok: false, message: 'This QR code is no longer active. Please ask the facilitator.' };
  }

  const existing = getRows(SHEETS.stamps).find(r =>
    String(r.user_id) === String(user.user_id) && Number(r.session_id) === sessionId
  );
  if (existing) return { ok: false, message: 'You already have the stamp for ' + session.session_name + '.' };

  const scanTime = nowText();
  appendRow(SHEETS.stamps, { record_id: newId('stamp'), user_id: user.user_id, session_id: sessionId, claimed_at: scanTime });

  return {
    ok: true, message: 'Stamp collected for ' + session.session_name + '!',
    sessionName: String(session.session_name), scannedAt: scanTime,
    user: buildUserProfile(user), sessions: getSessionsForClient(false)
  };
}

// ─────────────────────────────────────────────────────────────
// QR MANAGEMENT  (admin)
// ─────────────────────────────────────────────────────────────

function generateQRToken(payload) {
  requireAdmin(cleanEmail(payload.adminEmail || ''));
  const sessionId       = Number(payload.sessionId);
  const durationMinutes = Math.max(1, Math.min(1440, Number(payload.durationMinutes) || 15));
  const expiresAtMs     = nowMs() + durationMinutes * 60 * 1000;
  if (!sessionId) return { ok: false, message: 'sessionId is required.' };
  const session = getActiveSessions().find(s => Number(s.session_id) === sessionId);
  if (!session) return { ok: false, message: 'Session not found.' };
  const token = createQRToken(sessionId, expiresAtMs);
  const updated = updateRowWhere(SHEETS.sessions,
    row => Number(row.session_id) === sessionId,
    row => { row.qr_active = true; row.qr_expires_at = expiresAtMs; row.qr_token = token; }
  );
  if (!updated) return { ok: false, message: 'Could not update session record.' };
  return { ok: true, token, sessionId, sessionName: String(session.session_name), expiresAtMs, durationMinutes };
}

function deactivateQR(payload) {
  requireAdmin(cleanEmail(payload.adminEmail || ''));
  const sessionId = Number(payload.sessionId);
  if (!sessionId) return { ok: false, message: 'sessionId is required.' };
  updateRowWhere(SHEETS.sessions,
    row => Number(row.session_id) === sessionId,
    row => { row.qr_active = false; row.qr_expires_at = ''; row.qr_token = ''; }
  );
  return { ok: true, message: 'QR deactivated.' };
}

// ─────────────────────────────────────────────────────────────
// SESSION MANAGEMENT  (admin)
// ─────────────────────────────────────────────────────────────

function updateSession(payload) {
  requireAdmin(cleanEmail(payload.adminEmail || ''));
  const sessionId = Number(payload.sessionId);
  if (!sessionId) return { ok: false, message: 'sessionId is required.' };
  const name     = String(payload.sessionName || '').trim();
  const type     = String(payload.sessionType || 'Sharing').trim();
  const desc     = String(payload.description || '').trim();
  const location = String(payload.location    || '').trim();
  const time     = String(payload.time        || '').trim();
  if (!name) return { ok: false, message: 'Session name is required.' };
  if (location && !isValidSessionLocation(location)) return { ok: false, message: 'Location must be Building1, Building2, or Building3.' };
  const updated = updateRowWhere(SHEETS.sessions,
    row => Number(row.session_id) === sessionId,
    row => { row.session_name = name; row.session_type = type; row.description = desc; row.location = location; row.time = time; }
  );
  if (!updated) return { ok: false, message: 'Session not found.' };
  return getAdminParticipants(payload);
}

function createSession(payload) {
  requireAdmin(cleanEmail(payload.adminEmail || ''));
  const name     = String(payload.sessionName || '').trim();
  const type     = String(payload.sessionType || 'Sharing').trim();
  const desc     = String(payload.description || '').trim();
  const location = String(payload.location    || '').trim();
  const time     = String(payload.time        || '').trim();
  if (!name) return { ok: false, message: 'Session name is required.' };
  if (location && !isValidSessionLocation(location)) return { ok: false, message: 'Location must be Building1, Building2, or Building3.' };
  const newSessionId = nextSessionId_();
  appendRow(SHEETS.sessions, {
    session_id: newSessionId, session_name: name, session_type: type,
    description: desc, icon: randomIcon(), location, time,
    is_active: true, qr_active: false, qr_expires_at: '', qr_token: ''
  });
  return getAdminParticipants(payload);
}

function nextSessionId_() {
  const existing = getRows(SHEETS.sessions);
  return existing.reduce((m, r) => Math.max(m, Number(r.session_id) || 0), 0) + 1;
}

/**
 * Bulk-creates N placeholder sessions: "Session {next}", "Session {next+1}", ...
 * Only names are filled in; type/time/location/description are left blank
 * for the admin to edit afterwards.
 * payload: { adminEmail, count }
 */
function bulkCreateSessions(payload) {
  requireAdmin(cleanEmail(payload.adminEmail || ''));
  const count = Math.max(1, Math.min(50, Number(payload.count) || 0));
  if (!count) return { ok: false, message: 'Enter how many sessions to create.' };

  let nextId = nextSessionId_();
  const created = [];
  for (let i = 0; i < count; i++) {
    const sessionId = nextId + i;
    const name = 'Session ' + sessionId;
    appendRow(SHEETS.sessions, {
      session_id: sessionId, session_name: name, session_type: 'Sharing',
      description: '', icon: randomIcon(), location: '', time: '',
      is_active: true, qr_active: false, qr_expires_at: '', qr_token: ''
    });
    created.push(sessionId);
  }

  const result = getAdminParticipants(payload);
  result.createdSessionIds = created;
  result.message = created.length + ' session(s) created. Edit each one to set time, location, and description.';
  return result;
}

function deleteSession(payload) {
  requireAdmin(cleanEmail(payload.adminEmail || ''));
  const sessionId = Number(payload.sessionId);
  if (!sessionId) return { ok: false, message: 'sessionId is required.' };
  const updated = updateRowWhere(SHEETS.sessions,
    row => Number(row.session_id) === sessionId,
    row => { row.is_active = false; }
  );
  if (!updated) return { ok: false, message: 'Session not found.' };
  return getAdminParticipants(payload);
}

// ─────────────────────────────────────────────────────────────
// ADMIN PARTICIPANTS
// ─────────────────────────────────────────────────────────────

function getAdminParticipants(payload) {
  const adminEmail = cleanEmail(payload.adminEmail || payload.email || '');
  if (!adminEmail) return { ok: false, message: 'adminEmail is required.' };
  requireAdmin(adminEmail);
  const participants = getRows(SHEETS.users)
    .filter(u => String(u.role || '').toLowerCase() === 'participant')
    .map(u => { try { return buildUserProfile(u); } catch(e) { return null; } })
    .filter(Boolean);
  return { ok: true, participants, sessions: getSessionsForClient(true), template: getTemplateConfig() };
}

/**
 * Marks a reward as redeemed for a participant.
 * payload: { adminEmail, userId, rewardId }   // rewardId e.g. "m5"
 */
function markRedeemed(payload) {
  requireAdmin(cleanEmail(payload.adminEmail || payload.email || ''));
  const user = findUserById(payload.userId);
  if (!user) return { ok: false, message: 'Participant not found.' };

  const rewardId = String(payload.rewardId || payload.rewardType || '').trim();
  if (!rewardId) return { ok: false, message: 'rewardId is required.' };

  const template = getTemplateConfig();
  const milestone = template.milestones.find(m => m.id === rewardId);
  if (!milestone) return { ok: false, message: 'Unknown reward milestone.' };

  const profile = buildUserProfile(user);
  if (profile.stamps < milestone.position) {
    return { ok: false, message: 'Participant does not yet have ' + milestone.position + ' stamps.' };
  }

  const already = getRows(SHEETS.redemptions).find(r =>
    String(r.user_id) === String(user.user_id) && String(r.reward_id || r.reward_type) === rewardId
  );

  if (!already) {
    appendRow(SHEETS.redemptions, {
      redemption_id:        newId('red'),
      user_id:              user.user_id,
      reward_id:            rewardId,
      reward_type:          rewardId, // kept for backward-compat with old sheets
      redeemed_at:          nowText(),
      redeemed_by_admin_id: cleanEmail(payload.adminEmail || payload.email || '')
    });
  }

  return getAdminParticipants(payload);
}

// ─────────────────────────────────────────────────────────────
// ONE-TIME SETUP
// ─────────────────────────────────────────────────────────────

function setupInitialSheets() {
  const book = ss();
  const schema = [
    { name: SHEETS.users,       headers: ['user_id','full_name','facility','email','password','role','created_at'] },
    { name: SHEETS.sessions,    headers: ['session_id','session_name','session_type','description','icon','location','time','is_active','qr_active','qr_expires_at','qr_token'] },
    { name: SHEETS.stamps,      headers: ['record_id','user_id','session_id','claimed_at'] },
    { name: SHEETS.redemptions, headers: ['redemption_id','user_id','reward_id','reward_type','redeemed_at','redeemed_by_admin_id'] },
    { name: SHEETS.admins,      headers: ['email'] },
    { name: SHEETS.config,      headers: ['config_json','updated_at'] }
  ];
  schema.forEach(cfg => {
    let sh = book.getSheetByName(cfg.name);
    if (!sh) sh = book.insertSheet(cfg.name);
    sh.clear();
    sh.getRange(1, 1, 1, cfg.headers.length).setValues([cfg.headers]);
    sh.setFrozenRows(1);
  });

  const ts = nowText();
  const users = [
    ['u1','Arisa Tanaka','Bangkok HQ','arisa@company.com','pass123','participant',ts],
    ['u2','Ben Morales','Chiang Mai Branch','ben@company.com','pass123','participant',ts],
    ['u3','Chanya Patel','Phuket Office','chanya@company.com','pass123','participant',ts],
    ['u4','Dana Kim','Remote / WFH','dana@company.com','pass123','participant',ts],
    ['u5','Emon Srisuk','Bangkok HQ','emon@company.com','pass123','participant',ts],
    ['u6','Fiona Walsh','Chiang Mai Branch','fiona@company.com','pass123','participant',ts],
    ['adm','Admin User','HQ','admin@company.com','admin123','admin',ts]
  ];
  getSheet(SHEETS.users).getRange(2,1,users.length,users[0].length).setValues(users);

  const sessions = [
    [1, 'Session 1', 'Sharing',           'Introduction to AI tools in the workplace', 'ti-brain',            'Building1','09:00–10:00', true, false, '', ''],
    [2, 'Session 2', 'Workshop',           'Hands-on prompt engineering basics',         'ti-tools',            'Building2','10:15–11:15', true, false, '', ''],
    [3, 'Session 3', 'Sharing',           'Real-world AI use cases from our teams',     'ti-users',            'Building3','11:30–12:30', true, false, '', ''],
    [4, 'Session 4', 'Workshop',           'Building AI-assisted workflows onsite',      'ti-layout-dashboard', 'Building1','13:30–14:30', true, false, '', ''],
    [5, 'Session 5', 'Sharing',           'Responsible AI: ethics and governance',      'ti-shield-check',     'Building2','14:45–15:45', true, false, '', ''],
    [6, 'Session 6', 'Workshop',           'Advanced automation deep-dive',              'ti-rocket',           'Building3','16:00–17:00', true, false, '', ''],
    [7, 'Session 7', 'Sharing & Workshop', 'AI productivity tips for daily work',        'ti-sparkles',         'Building1','09:00–10:30', true, false, '', ''],
    [8, 'Session 8', 'Workshop',           'Designing better prompts for teams',         'ti-pencil',           'Building2','10:45–11:45', true, false, '', ''],
    [9, 'Session 9', 'Sharing',           'Data storytelling with AI assistance',       'ti-chart-bar',        'Building3','13:00–14:00', true, false, '', ''],
    [10,'Session 10','Sharing & Workshop', 'AI Community wrap-up and next steps',        'ti-confetti',         'Building1','15:00–16:00', true, false, '', '']
  ];
  getSheet(SHEETS.sessions).getRange(2,1,sessions.length,sessions[0].length).setValues(sessions);

  const seedClaims = { u1:[1,2,3,4,5,6,7], u2:[1,2,3,4,5], u3:[1,2,3,4,5,6,7,8,9,10], u4:[1,2,3], u5:[1,2,3,4,5,6], u6:[1,2,3,4,5,6,7,8,9,10] };
  const stampRows = [];
  Object.keys(seedClaims).forEach(uid => {
    seedClaims[uid].forEach((sid, idx) => {
      const fakeDate = new Date(Date.now() - (seedClaims[uid].length - idx) * 3 * 24 * 60 * 60 * 1000);
      stampRows.push([newId('stamp'), uid, sid, Utilities.formatDate(fakeDate, Session.getScriptTimeZone(), 'd MMM yyyy HH:mm:ss')]);
    });
  });
  if (stampRows.length) getSheet(SHEETS.stamps).getRange(2,1,stampRows.length,stampRows[0].length).setValues(stampRows);

  // Redemptions seeded using the new reward_id scheme (m5 / m10) matching DEFAULT_TEMPLATE.
  const redemptions = [
    [newId('red'),'u2','m5','m5','15 Jun 2025 14:30:00','adm'],
    [newId('red'),'u3','m5','m5','10 Jun 2025 11:15:00','adm'],
    [newId('red'),'u3','m10','m10','15 Jun 2025 14:32:00','adm'],
    [newId('red'),'u6','m5','m5','12 Jun 2025 16:45:00','adm']
  ];
  getSheet(SHEETS.redemptions).getRange(2,1,redemptions.length,redemptions[0].length).setValues(redemptions);
  getSheet(SHEETS.admins).getRange(2,1).setValue('admin@company.com');

  // Seed the default template config row so Config sheet isn't empty.
  getSheet(SHEETS.config).getRange(2, 1, 1, 2).setValues([[JSON.stringify(DEFAULT_TEMPLATE), ts]]);

  Logger.log('setupInitialSheets() complete.');
}
