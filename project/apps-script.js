// =============================================
// MAHAL WELFARE SERVICES – Google Apps Script
// File: apps-script.js  (paste ALL of this into Code.gs)
//
// HOW TO DEPLOY:
// 1. Go to https://script.google.com
// 2. Create new project, name it "MahalWelfareQuiz"
// 3. Paste this ENTIRE file content into Code.gs
// 4. Click the gear icon (Project Settings)
// 5. Under "Script Properties", add:
//    - ADMIN_PASSWORD  →  (your chosen password, e.g. Mahal@2025)
//    - SHEET_ID        →  (your Google Sheet ID from its URL)
//    - ANSWER_KEY      →  comma-separated correct option indices (0-indexed)
//                         e.g. 0,2,3,0,1,1,2,3,1,0,1,1,1,1,0,1,0,1,0,1,1,1,1,1,1,2,1,1,2
//    - TOKEN_SECRET    →  any long random string, e.g. x7Kp!qZ2#mNv9L
// 6. Click Deploy → New Deployment
//    - Type: Web App
//    - Execute as: Me
//    - Who has access: Anyone
// 7. Click Deploy, copy the Web App URL
// 8. Paste that URL into index.html and admin.html where it says YOUR_DEPLOYMENT_ID
//
// GOOGLE SHEET SETUP:
// 1. Create a new Google Sheet
// 2. Copy its ID from the URL (between /d/ and /edit)
// 3. The script will auto-create these sheets:
//    - "Submissions" : stores quiz results
//    - "SiteContent" : stores editable site content
//    NEVER store answerKey in SiteContent — it stays in Script Properties only.
// =============================================

// =============================================
// ENTRY POINTS
// =============================================

function doGet(e) {
  const params = e.parameter || {};
  const action = params.action || '';

  try {
    if (action === 'getContent')   return jsonResponse(getContent());
    if (action === 'getResults')   return jsonResponse(getResults(params));
    return jsonResponse({ error: 'Unknown action' });
  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}

function doPost(e) {
  let body = {};
  try {
    body = JSON.parse(e.postData.contents);
  } catch (_) {
    body = e.parameter || {};
  }

  const action = body.action || '';
  try {
    if (action === 'submit')        return jsonResponse(submitQuiz(body));
    if (action === 'updateContent') return jsonResponse(updateContent(body));
    if (action === 'adminLogin')    return jsonResponse(adminLogin(body));
    if (action === 'setAnswerKey')  return jsonResponse(setAnswerKey(body));
    return jsonResponse({ error: 'Unknown action' });
  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// =============================================
// HELPER: Get Script Properties
// =============================================
function getProp(key) {
  return PropertiesService.getScriptProperties().getProperty(key) || '';
}
function setProp(key, value) {
  PropertiesService.getScriptProperties().setProperty(key, value);
}

// =============================================
// HELPER: Get Spreadsheet
// =============================================
function getSpreadsheet() {
  const id = getProp('SHEET_ID');
  if (!id) throw new Error('SHEET_ID not set in Script Properties.');
  return SpreadsheetApp.openById(id);
}

function getOrCreateSheet(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (headers) sheet.appendRow(headers);
  }
  return sheet;
}

// =============================================
// ADMIN LOGIN  (HMAC-based secure token)
// Token = nonce:hmac  where:
//   nonce = timestamp in ms (expires 24 h from now)
//   hmac  = hex of HMAC-SHA256(nonce, TOKEN_SECRET)
// The password is never derivable from the token.
// =============================================
function makeHmac(nonce) {
  const secret  = getProp('TOKEN_SECRET');
  const sigBytes = Utilities.computeHmacSha256Signature(
    Utilities.newBlob(nonce).getBytes(),
    Utilities.newBlob(secret).getBytes()
  );
  return sigBytes.map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
}

function adminLogin(body) {
  const password = body.password || '';
  const stored   = getProp('ADMIN_PASSWORD');
  if (!stored) return { error: 'Admin password not configured in Script Properties.' };
  if (password !== stored) return { error: 'Incorrect password.' };

  // Issue a token that expires 24 h from now
  const nonce = Date.now() + (24 * 60 * 60 * 1000);   // expiry timestamp
  const hmac  = makeHmac(String(nonce));
  const token = nonce + ':' + hmac;
  return { success: true, token };
}

function verifyToken(token) {
  if (!token) return false;
  const parts = token.split(':');
  if (parts.length !== 2) return false;
  const [nonce, hmac] = parts;
  // Check expiry
  if (Number(nonce) < Date.now()) return false;
  // Check HMAC
  const expected = makeHmac(nonce);
  return hmac === expected;
}

// =============================================
// QUIZ SUBMISSION
// =============================================
function submitQuiz(body) {
  const student = body.student || {};
  const answers = body.answers || [];

  // Validate required fields
  if (!student.name || !student.email || !student.phone || !student.college) {
    return { error: 'Required information is missing. Please fill in all fields.' };
  }

  const ss = getSpreadsheet();
  const submissionsSheet = getOrCreateSheet(ss, 'Submissions', [
    'Timestamp', 'Name', 'College', 'Phone', 'Email', 'WantsToJoin', 'Score', 'Total', 'Answers'
  ]);

  // Use a lock to prevent race conditions on simultaneous submissions
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);

  try {
    // Duplicate check: email or phone already in sheet
    const data = submissionsSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (
        (data[i][4] && data[i][4].toString().toLowerCase() === student.email.toLowerCase()) ||
        (data[i][3] && data[i][3].toString() === student.phone.toString())
      ) {
        return { error: 'You have already submitted the quiz. Each student may submit only once.' };
      }
    }

    // Server-side scoring (answers never sent back to client)
    const answerKey = getProp('ANSWER_KEY');
    let score = 0;
    let total = answers.length;

    if (answerKey) {
      const correctAnswers = answerKey.split(',').map(Number);
      total = Math.min(answers.length, correctAnswers.length);
      for (let i = 0; i < total; i++) {
        if (parseInt(answers[i]) === correctAnswers[i]) score++;
      }
    } else {
      score = -1; // ANSWER_KEY not configured yet
    }

    // Save to sheet
    submissionsSheet.appendRow([
      new Date().toISOString(),
      student.name,
      student.college,
      student.phone,
      student.email,
      student.wantsToJoin ? 'Yes' : 'No',
      score >= 0 ? score : 'Pending',
      total,
      JSON.stringify(answers)
    ]);

    return {
      success: true,
      score: score >= 0 ? score : '?',
      total
    };
  } finally {
    lock.releaseLock();
  }
}

// =============================================
// GET TOP RESULTS (Admin only)
// =============================================
function getResults(params) {
  // Token check removed — admin.html has no login.

  const ss    = getSpreadsheet();
  const sheet = ss.getSheetByName('Submissions');
  if (!sheet) return { results: [], message: 'No submissions yet.' };

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { results: [] };

  const rows = data.slice(1).map(row => ({
    timestamp:   row[0],
    name:        row[1],
    college:     row[2],
    phone:       row[3],
    email:       row[4],
    wantsToJoin: row[5],
    score:       row[6],
    total:       row[7]
  }));

  // Sort by score descending
  rows.sort((a, b) => (parseFloat(b.score) || 0) - (parseFloat(a.score) || 0));

  return {
    top5:  rows.slice(0, 5),
    all:   rows,
    total: rows.length
  };
}

// =============================================
// SITE CONTENT (Read – PUBLIC)
// NEVER returns answerKey or questions.correct
// =============================================
function getContent() {
  // Keys that must never be sent to the public browser
  const PRIVATE_KEYS = new Set(['answerKey', 'ANSWER_KEY']);

  try {
    const ss    = getSpreadsheet();
    const sheet = getOrCreateSheet(ss, 'SiteContent', ['Key', 'Value']);
    const data  = sheet.getDataRange().getValues();
    const result = {};
    for (let i = 1; i < data.length; i++) {
      const key = data[i][0];
      const val = data[i][1];
      if (key && !PRIVATE_KEYS.has(key)) {
        try { result[key] = JSON.parse(val); }
        catch (_) { result[key] = val; }
      }
    }
    // Strip correct answers from questions array if somehow present
    if (Array.isArray(result.questions)) {
      result.questions = result.questions.map(({ q, opts }) => ({ q, opts }));
    }
    return { data: result };
  } catch (e) {
    return { data: {} };
  }
}

// =============================================
// SITE CONTENT (Write – Admin only)
// If updates contain answerKey, it is saved to
// Script Properties ONLY (never to the Sheet).
// =============================================
function updateContent(body) {
  // Token check removed — admin.html has no login; protected by URL obscurity.

  const updates = { ...(body.updates || {}) };

  // Intercept answerKey — save to Script Properties, NOT the sheet
  if (updates.answerKey !== undefined) {
    setProp('ANSWER_KEY', String(updates.answerKey));
    delete updates.answerKey;
  }

  // Strip correct answers from questions before writing to sheet
  if (Array.isArray(updates.questions)) {
    updates.questions = updates.questions.map(({ q, opts }) => ({ q, opts }));
  }

  const ss    = getSpreadsheet();
  const sheet = getOrCreateSheet(ss, 'SiteContent', ['Key', 'Value']);

  // Read existing rows to know which rows to update vs append
  const data = sheet.getDataRange().getValues();
  const keyRow = {}; // key → row index (1-based)
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) keyRow[data[i][0]] = i + 1;
  }

  Object.entries(updates).forEach(([key, value]) => {
    const serialized = typeof value === 'object' ? JSON.stringify(value) : String(value);
    if (keyRow[key]) {
      sheet.getRange(keyRow[key], 2).setValue(serialized);
    } else {
      sheet.appendRow([key, serialized]);
    }
  });

  return { success: true };
}

// =============================================
// SET ANSWER KEY directly (Admin only)
// Called by admin panel after saving questions.
// Writes ONLY to Script Properties.
// =============================================
function setAnswerKey(body) {
  // Token check removed — admin.html has no login.
  const key = body.answerKey;
  if (key === undefined || key === null) return { error: 'answerKey is required.' };
  setProp('ANSWER_KEY', String(key));
  return { success: true };
}
