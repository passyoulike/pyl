// Quiz sets are discovered automatically from the spreadsheet's tabs — add a
// new tab and it appears in the app with no code changes required. Everything
// except these reserved sheets is treated as a quiz set.
const RESERVED_SHEETS = ['Members', 'Quiz Results', 'Contact Us', 'Email'];

// Non-quiz utility/admin tabs that have accumulated in the spreadsheet (a
// pivot table, a roster, duplicated backup copies, etc.) — excluded by name
// so they're never scanned at all, not even the cheap header check.
const IGNORED_SHEETS = [
  'MCN 2 - P1', 'Register', 'Pivot Table 1', 'Results', 'Radi', 'Copy of Radi', 'Copy of Med Surg - '
];

// Where registration/Contact Us notifications are sent — read from column A
// of the "Email" sheet (one address per row; a header row like "Email" is
// automatically skipped since it doesn't contain an "@"). Falls back to a
// hardcoded address if that sheet is missing or empty, so notifications
// don't silently stop if the sheet gets renamed or cleared by mistake.
function getNotifyEmails_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Email');
  if (sheet) {
    const values = sheet.getRange(1, 1, Math.max(sheet.getLastRow(), 1), 1).getValues();
    const emails = values
      .map(function (row) { return String(row[0] || '').trim(); })
      .filter(function (v) { return v.indexOf('@') !== -1; });
    if (emails.length > 0) return emails.join(',');
  }
  return 'raymerkado@gmail.com';
}

// A tab named "Prefix - Topic" (e.g. "MCN - Sensory") is auto-split into a
// category ("Prefix") and a short label ("Topic"). Only these two categories
// are ever shown — a sheet whose prefix isn't recognized here doesn't appear
// on the dashboard at all (see resolveCategory_), rather than spawning its
// own new category tile.
const CATEGORY_PREFIX_MAP = {
  'MCN': 'MCN - Maternal & Child',
  'Basic': 'Medical Surgical Nursing',
  'Med Surg': 'Medical Surgical Nursing'
};
// Display order for the two allowed categories.
const CATEGORY_ORDER = ['MCN - Maternal & Child', 'Medical Surgical Nursing'];

function splitSheetName_(name) {
  const idx = name.indexOf(' - ');
  if (idx === -1) return { prefix: name.trim(), topic: name.trim() };
  return { prefix: name.slice(0, idx).trim(), topic: name.slice(idx + 3).trim() };
}

// Returns the display category for a sheet-name prefix, or null if it isn't
// one of the recognized categories — callers should exclude that sheet
// entirely rather than showing it under a stray/unexpected category.
function resolveCategory_(prefix) {
  return CATEGORY_PREFIX_MAP[prefix] || null;
}

// Every spreadsheet tab except the reserved ones is a candidate quiz set.
// Tabs that don't actually contain gradable questions (wrong column layout,
// empty, a scratch/notes tab, etc.) are filtered out later by getTestList()
// once their question count comes back as zero.
function getQuestionSheetNames_() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheets()
    .map(function (s) { return s.getName(); })
    .filter(function (name) {
      return RESERVED_SHEETS.indexOf(name) === -1 && IGNORED_SHEETS.indexOf(name) === -1;
    });
}

// Student IDs allowed to view the class results summary.
const ADMIN_IDS = ['906223'];

function isAdminId_(id) {
  return ADMIN_IDS.indexOf(String(id || '').trim()) !== -1;
}

function doGet(e) {
  // ?action=<fnName>&args=<JSON array> switches this into a plain JSON API,
  // used by the standalone frontend hosted on GitHub Pages (which can't use
  // google.script.run since it's not served by HtmlService). No action param
  // means a normal visit, so the built-in HTML app still works as before.
  if (e && e.parameter && e.parameter.action) {
    return handleApiRequest_(e.parameter.action, e.parameter.args);
  }
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
        .setTitle('Passyoulike')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Functions callable through the ?action= JSON API. Deliberately a fixed
// allowlist rather than global-function lookup, so no other server function
// becomes reachable just by naming it in a URL.
const API_FUNCTIONS_ = {
  authenticate: authenticate,
  registerMember: registerMember,
  getTestListForStudent: getTestListForStudent,
  getCategorySummary: getCategorySummary,
  getTestsForCategory: getTestsForCategory,
  getQuestionSet: getQuestionSet,
  saveResult: saveResult,
  getResultsSummary: getResultsSummary,
  getMaintenanceStatus: getMaintenanceStatus,
  setMaintenanceMode: setMaintenanceMode,
  requestPasswordReset: requestPasswordReset,
  resetPasswordWithCode: resetPasswordWithCode,
  submitContactRequest: submitContactRequest
};

function handleApiRequest_(action, argsJson) {
  let payload;
  try {
    const fn = API_FUNCTIONS_[action];
    if (!fn) throw new Error('Unknown action: ' + action);
    const args = argsJson ? JSON.parse(argsJson) : [];
    payload = { ok: true, result: fn.apply(null, args) };
  } catch (err) {
    payload = { ok: false, error: (err && err.message) ? err.message : String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(payload)).setMimeType(ContentService.MimeType.JSON);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getSheet_(name) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sheet) throw new Error('Sheet not found: ' + name);
  return sheet;
}

// Members columns: A Email, B Name, C Password, D Referral Code, E Gcash
// confirmation, F Status. Existing rows predate the Status column, so a blank
// status is treated as approved — only self-registered rows (via
// registerMember) start out 'Pending' and need an admin to flip them.
function authenticate(studentId, password) {
  const data = getSheet_('Members').getDataRange().getValues();
  const idIn = String(studentId || '').trim().toLowerCase();
  const pwIn = String(password || '').trim();

  for (let i = 1; i < data.length; i++) {
    const id = String(data[i][0] || '').trim();
    const name = String(data[i][1] || '').trim();
    const pw = String(data[i][2] || '').trim();
    if (!id) continue;
    if (id.toLowerCase() === idIn && pw === pwIn) {
      const status = String(data[i][5] || '').trim().toLowerCase();
      if (status && status !== 'approved') {
        return { ok: false, pending: true };
      }
      const isAdmin = isAdminId_(id);
      if (isMaintenanceMode_() && !isAdmin) {
        return { ok: false, maintenance: true };
      }
      return { ok: true, id: id, name: name, isAdmin: isAdmin };
    }
  }
  return { ok: false };
}

// Maintenance mode blocks non-admin logins so an admin can safely edit the
// spreadsheet (e.g. reordering/renaming quiz tabs) without students hitting a
// half-updated state. Stored in Script Properties, not CacheService, since it
// must persist indefinitely until an admin turns it back off.
function isMaintenanceMode_() {
  return PropertiesService.getScriptProperties().getProperty('MAINTENANCE_MODE') === 'true';
}

function getMaintenanceStatus() {
  return { enabled: isMaintenanceMode_() };
}

function setMaintenanceMode(requesterId, enabled) {
  if (!isAdminId_(requesterId)) throw new Error('Not authorized');
  PropertiesService.getScriptProperties().setProperty('MAINTENANCE_MODE', enabled ? 'true' : 'false');
  return { ok: true, enabled: !!enabled };
}

const TRIAL_MONTHS = 5;

// Free 5-month trial: no payment to verify, so accounts are approved
// immediately rather than left Pending for manual review. Trial end date
// (column G) and a live "days left" countdown (column H, a formula that
// recalculates on its own whenever the sheet is opened) are recorded so you
// can see at a glance who's about to run out.
function registerMember(fields) {
  const name = String((fields && fields.name) || '').trim();
  const email = String((fields && fields.email) || '').trim();
  const password = String((fields && fields.password) || '').trim();

  if (!name || !email || !password) {
    throw new Error('Name, email, and password are required.');
  }

  const sheet = getSheet_('Members');
  const data = sheet.getDataRange().getValues();
  const emailLower = email.toLowerCase();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0] || '').trim().toLowerCase() === emailLower) {
      throw new Error('This email is already registered.');
    }
  }

  if (String(sheet.getRange(1, 6).getValue() || '').trim() === '') {
    sheet.getRange(1, 6).setValue('Status');
  }
  if (String(sheet.getRange(1, 7).getValue() || '').trim() === '') {
    sheet.getRange(1, 7).setValue('Trial Ends');
  }
  if (String(sheet.getRange(1, 8).getValue() || '').trim() === '') {
    sheet.getRange(1, 8).setValue('Trial Countdown');
  }

  const trialEnd = new Date();
  trialEnd.setMonth(trialEnd.getMonth() + TRIAL_MONTHS);

  // Columns D (Referral Code) and E (Gcash confirmation) are left blank —
  // kept so the sheet's column layout stays compatible with older rows.
  sheet.appendRow([email, name, password, '', '', 'Approved', trialEnd, '']);
  const newRow = sheet.getLastRow();
  sheet.getRange(newRow, 8).setFormula(
    '=IF(TODAY()>G' + newRow + ',"Expired",DATEDIF(TODAY(),G' + newRow + ',"D")&" days left")'
  );

  MailApp.sendEmail({
    to: getNotifyEmails_(),
    replyTo: email,
    subject: 'New Passyoulike registration (5-month trial): ' + name,
    body: 'A new account started its 5-month free trial and can log in immediately.\n\n' +
      'Name: ' + name + '\n' +
      'Email: ' + email + '\n' +
      'Trial ends: ' + trialEnd.toDateString()
  });

  return { ok: true };
}

// ---------- Forgot password ----------
// A 6-digit code is emailed to the address on file and cached (not stored in
// the sheet) for 10 minutes with a limited number of guesses, so a reset
// requires proving control of the email rather than just knowing it.
const RESET_CODE_TTL_SECONDS = 600;
const RESET_CODE_MAX_ATTEMPTS = 5;

function resetCacheKey_(email) {
  return 'pwreset_' + String(email || '').trim().toLowerCase();
}

function requestPasswordReset(email) {
  const emailIn = String(email || '').trim();
  if (!emailIn) throw new Error('Please enter your email address.');
  const emailLower = emailIn.toLowerCase();

  const data = getSheet_('Members').getDataRange().getValues();
  const found = data.some(function (row, i) {
    return i > 0 && String(row[0] || '').trim().toLowerCase() === emailLower;
  });

  // Always return ok, whether or not the email is registered, so this can't be
  // used to test which emails exist in the Members sheet.
  if (!found) return { ok: true };

  const code = String(Math.floor(100000 + Math.random() * 900000));
  CacheService.getScriptCache().put(
    resetCacheKey_(emailIn),
    JSON.stringify({ code: code, attempts: 0 }),
    RESET_CODE_TTL_SECONDS
  );

  MailApp.sendEmail(
    emailIn,
    'Passyoulike password reset code',
    'Your password reset code is: ' + code + '\n\n' +
    'This code expires in 10 minutes. If you did not request this, you can ignore this email.'
  );

  return { ok: true };
}

function resetPasswordWithCode(email, code, newPassword) {
  const emailIn = String(email || '').trim();
  const codeIn = String(code || '').trim();
  const pwIn = String(newPassword || '').trim();

  if (!emailIn || !codeIn || !pwIn) throw new Error('All fields are required.');
  if (pwIn.length < 6) throw new Error('New password must be at least 6 characters.');

  const cache = CacheService.getScriptCache();
  const key = resetCacheKey_(emailIn);
  const cached = cache.get(key);
  if (!cached) throw new Error('That code has expired. Please request a new one.');

  const entry = JSON.parse(cached);
  if (entry.attempts >= RESET_CODE_MAX_ATTEMPTS) {
    cache.remove(key);
    throw new Error('Too many incorrect attempts. Please request a new code.');
  }
  if (codeIn !== entry.code) {
    entry.attempts++;
    cache.put(key, JSON.stringify(entry), RESET_CODE_TTL_SECONDS);
    throw new Error('Incorrect code. Please try again.');
  }

  const sheet = getSheet_('Members');
  const data = sheet.getDataRange().getValues();
  const emailLower = emailIn.toLowerCase();
  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0] || '').trim().toLowerCase() === emailLower) {
      rowIndex = i + 1; // sheet rows are 1-indexed
      break;
    }
  }
  if (rowIndex === -1) throw new Error('Account not found.');

  sheet.getRange(rowIndex, 3).setValue(pwIn); // column C: Password
  cache.remove(key);
  return { ok: true };
}

// ---------- Contact Us ----------
function getOrCreateContactSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Contact Us');
  if (!sheet) {
    sheet = ss.insertSheet('Contact Us');
    sheet.appendRow(['Timestamp', 'Ticket #', 'Email', 'Subject', 'Message']);
  }
  return sheet;
}

function submitContactRequest(fields) {
  const email = String((fields && fields.email) || '').trim();
  const subject = String((fields && fields.subject) || '').trim();
  const message = String((fields && fields.message) || '').trim();

  if (!email || !subject || !message) {
    throw new Error('Email, subject, and message are all required.');
  }

  const sheet = getOrCreateContactSheet_();
  // Ticket number = next row number, so it's always unique and sequential
  // without needing a separate counter to keep in sync.
  const ticketNumber = 'TCK-' + String(sheet.getLastRow()).padStart(4, '0');

  sheet.appendRow([new Date(), ticketNumber, email, subject, message]);

  MailApp.sendEmail({
    to: getNotifyEmails_(),
    replyTo: email,
    subject: '[' + ticketNumber + '] ' + subject,
    body: 'New Contact Us ticket ' + ticketNumber + '\n\n' +
      'From: ' + email + '\n' +
      'Subject: ' + subject + '\n\n' +
      message + '\n\n' +
      'Reply directly to this email to respond to ' + email + '.'
  });

  return { ok: true, ticketNumber: ticketNumber };
}

// Locates each question sheet's columns by header name instead of assuming a
// fixed position, so a sheet can have extra columns (e.g. an "Item #" column
// before "Question", as some tabs do) without breaking parsing. Returns null
// if the header doesn't contain recognizable "Question"/"Answer" columns at
// all — that means the tab isn't actually a quiz sheet (a pivot table, a
// roster, a backup copy, etc.), and callers should skip it WITHOUT reading
// the sheet's full data, since a spreadsheet can accumulate a lot of large
// non-quiz tabs over time and reading all of them on every cache-cold
// dashboard load is what was making loading slow.
// Takes an already-fetched header row (rather than reading the sheet itself)
// so callers can get the layout from the same getDataRange() call they use
// for the sheet's actual rows — one spreadsheet call per sheet instead of
// two. Reading each quiz sheet twice (once for the header, once for the
// data) across 40+ tabs was the main reason loading was slow.
function deriveColumnLayout_(headerRow) {
  const header = (headerRow || []).map(function (h) {
    return String(h || '').trim().toLowerCase();
  });

  function find(aliases) {
    for (let i = 0; i < header.length; i++) {
      if (aliases.indexOf(header[i]) !== -1) return i;
    }
    return -1;
  }

  const layout = {
    question: find(['question']),
    A: find(['choice a', 'a']),
    B: find(['choice b', 'b']),
    C: find(['choice c', 'c']),
    D: find(['choice d', 'd']),
    answer: find(['correct answer', 'answer']),
    rationale: find(['rationale'])
  };

  // No recognizable Question/Answer columns — not actually a quiz sheet
  // (a pivot table, a roster, a backup copy, etc.).
  if (layout.question === -1 || layout.answer === -1) return null;
  return layout;
}

// Counts gradable questions in already-fetched sheet data the same way
// getQuestionSet() filters them, so the count shown to students always
// matches what they'll actually be asked. Split from countQuestions_ so
// getTestsForCategory() can reuse it against data fetched in one batch call
// instead of a separate read per sheet.
function countQuestionsFromData_(data) {
  if (!data || data.length === 0) return 0;
  const layout = deriveColumnLayout_(data[0]);
  if (!layout) return 0;
  let count = 0;
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[layout.question]) continue;
    const answer = String(row[layout.answer] || '').trim().toUpperCase();
    if (!answer || answer.indexOf(',') !== -1 || ['A', 'B', 'C', 'D'].indexOf(answer) === -1) continue;
    count++;
  }
  return count;
}

function countQuestions_(sheetName) {
  return countQuestionsFromData_(getSheet_(sheetName).getDataRange().getValues());
}

// Fetches every listed sheet's full data in a single Sheets API call instead
// of one SpreadsheetApp call per sheet. This is the difference between a
// category with ~20 sheets taking ~1-2s (one HTTP call) vs ~10-20s (20
// sequential calls, each with its own fixed round-trip latency) — the
// latter was slow enough to occasionally hit Google's own timeout for
// serving the web app request. Requires the Sheets Advanced Service
// (see appsscript.json).
function batchGetSheetValues_(sheetNames) {
  if (!sheetNames || sheetNames.length === 0) return {};
  const ssId = SpreadsheetApp.getActiveSpreadsheet().getId();
  const ranges = sheetNames.map(function (name) {
    return "'" + name.replace(/'/g, "''") + "'";
  });
  const response = Sheets.Spreadsheets.Values.batchGet(ssId, { ranges: ranges });
  const result = {};
  (response.valueRanges || []).forEach(function (vr, i) {
    result[sheetNames[i]] = vr.values || [];
  });
  return result;
}

// Building this list means reading quiz sheets in full just to count their
// questions. Cache it so concurrent students and repeat visits don't each
// pay that cost. getTestsForCategory() now fetches a whole category's sheets
// in one batched Sheets API call (see batchGetSheetValues_), so the cold
// path is fast enough (~1-2s) that a short TTL is fine — a newly added quiz
// sheet shows up within this long.
const TEST_LIST_CACHE_SECONDS = 300;

function getTestList() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('testList');
  if (cached) return JSON.parse(cached);

  // A single malformed/unreadable sheet shouldn't take down the whole
  // dashboard for every student — skip it and keep going.
  const tests = getQuestionSheetNames_().map(function (name) {
    try {
      const parts = splitSheetName_(name);
      const category = resolveCategory_(parts.prefix);
      if (!category) return null;
      return {
        key: name,
        label: name,
        shortLabel: parts.topic,
        category: category,
        totalQuestions: countQuestions_(name)
      };
    } catch (err) {
      return null;
    }
  }).filter(function (t) { return t && t.totalQuestions > 0; });

  cache.put('testList', JSON.stringify(tests), TEST_LIST_CACHE_SECONDS);
  return tests;
}

// Returns the test list annotated with each test's completion status for the given student,
// based on their best recorded attempt in the Quiz Results sheet.
function getTestListForStudent(studentId) {
  const tests = getTestList();
  const completedMap = getCompletedTestsForStudent_(studentId);
  return tests.map(function (t) {
    const c = completedMap[t.label];
    return {
      key: t.key,
      label: t.label,
      shortLabel: t.shortLabel,
      category: t.category,
      totalQuestions: t.totalQuestions,
      completed: !!c,
      bestPct: c ? c.bestPct : null,
      bestScore: c ? c.bestScore : null,
      bestTotal: c ? c.bestTotal : null
    };
  });
}

// ---------- Lazy per-category loading ----------
// getTestList()/getTestListForStudent() scan every quiz sheet across every
// category, which is expensive with 40+ tabs. These two split that into a
// cheap dashboard-level summary (sheet names/counts only, no per-sheet
// content reads) and a per-category scan that only touches the sheets in
// the category actually being opened — the dashboard no longer pays for
// categories the student never clicks into.

function getCategorySummary() {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'categorySummary';
  const cached = cache.get(cacheKey);
  if (cached) return JSON.parse(cached);

  const byCategory = {};
  getQuestionSheetNames_().forEach(function (name) {
    const parts = splitSheetName_(name);
    const cat = resolveCategory_(parts.prefix);
    if (!cat) return;
    byCategory[cat] = (byCategory[cat] || 0) + 1;
  });

  const summary = Object.keys(byCategory).map(function (cat) {
    return { category: cat, testCount: byCategory[cat] };
  });

  cache.put(cacheKey, JSON.stringify(summary), TEST_LIST_CACHE_SECONDS);
  return summary;
}

function getTestsForCategory(categoryName, studentId) {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'categoryTests_' + categoryName;
  let tests = cache.get(cacheKey);

  if (tests) {
    tests = JSON.parse(tests);
  } else {
    const sheetNames = getQuestionSheetNames_().filter(function (name) {
      const parts = splitSheetName_(name);
      return resolveCategory_(parts.prefix) === categoryName;
    });

    const dataBySheet = batchGetSheetValues_(sheetNames);

    tests = sheetNames.map(function (name) {
      try {
        const parts = splitSheetName_(name);
        return {
          key: name,
          label: name,
          shortLabel: parts.topic,
          category: categoryName,
          totalQuestions: countQuestionsFromData_(dataBySheet[name])
        };
      } catch (err) {
        return null;
      }
    }).filter(function (t) { return t && t.totalQuestions > 0; });

    cache.put(cacheKey, JSON.stringify(tests), TEST_LIST_CACHE_SECONDS);
  }

  const completedMap = getCompletedTestsForStudent_(studentId);
  return tests.map(function (t) {
    const c = completedMap[t.label];
    return {
      key: t.key,
      label: t.label,
      shortLabel: t.shortLabel,
      category: t.category,
      totalQuestions: t.totalQuestions,
      completed: !!c,
      bestPct: c ? c.bestPct : null,
      bestScore: c ? c.bestScore : null,
      bestTotal: c ? c.bestTotal : null
    };
  });
}

// This scans the entire (and ever-growing) Quiz Results sheet, so it's the
// slowest part of loading the dashboard/category view — cache each student's
// result briefly. saveResult() below clears that student's entry immediately
// so a just-finished test still shows up right away.
const COMPLETED_CACHE_SECONDS = 60;

function completedCacheKey_(studentId) {
  return 'completed_' + String(studentId || '').trim().toLowerCase();
}

function getCompletedTestsForStudent_(studentId) {
  const cache = CacheService.getScriptCache();
  const cacheKey = completedCacheKey_(studentId);
  const cached = cache.get(cacheKey);
  if (cached) return JSON.parse(cached);

  const result = {};
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Quiz Results');
  if (!sheet || sheet.getLastRow() < 2) return result;

  const data = sheet.getDataRange().getValues();
  const idIn = String(studentId || '').trim().toLowerCase();

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const sid = String(row[1] || '').trim().toLowerCase();
    if (sid !== idIn) continue;
    const test = String(row[3] || '').trim();
    if (!test) continue;
    const score = Number(row[4]);
    const total = Number(row[5]);
    const pct = Number(row[6]);
    if (!result[test] || pct > result[test].bestPct) {
      result[test] = { bestScore: score, bestTotal: total, bestPct: pct };
    }
  }

  cache.put(cacheKey, JSON.stringify(result), COMPLETED_CACHE_SECONDS);
  return result;
}

function getQuestionSet(setKey, studentId) {
  const sheetName = setKey;
  if (RESERVED_SHEETS.indexOf(sheetName) !== -1 || !SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName)) {
    throw new Error('Unknown question set: ' + setKey);
  }

  const data = getSheet_(sheetName).getDataRange().getValues();
  const layout = data.length ? deriveColumnLayout_(data[0]) : null;
  if (!layout) throw new Error('Unknown question set: ' + setKey);

  const questions = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const question = row[layout.question];
    if (!question) continue;

    const answer = String(row[layout.answer] || '').trim().toUpperCase();
    // Skip malformed rows: multi-answer "select all that apply" items (e.g. "A,C,E")
    // aren't gradable by this single-answer A-D quiz format.
    if (!answer || answer.indexOf(',') !== -1 || ['A', 'B', 'C', 'D'].indexOf(answer) === -1) continue;

    questions.push({
      question: String(question),
      choices: {
        A: String(row[layout.A] || ''),
        B: String(row[layout.B] || ''),
        C: String(row[layout.C] || ''),
        D: String(row[layout.D] || '')
      },
      answer: answer,
      rationale: String(row[layout.rationale] || '')
    });
  }
  return questions;
}

function saveResult(record) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Quiz Results');
  if (!sheet) {
    sheet = ss.insertSheet('Quiz Results');
    sheet.appendRow(['Timestamp', 'Student ID', 'Student Name', 'Test', 'Score', 'Total', 'Percentage']);
  }
  sheet.appendRow([
    new Date(),
    record.studentId,
    record.studentName,
    record.test,
    record.score,
    record.total,
    record.pct
  ]);
  CacheService.getScriptCache().remove(completedCacheKey_(record.studentId));
  return { ok: true };
}

function getResultsSummary(requesterId) {
  if (!isAdminId_(requesterId)) throw new Error('Not authorized');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Quiz Results');
  if (!sheet || sheet.getLastRow() < 2) return { tests: [], leaderboard: {} };

  const data = sheet.getDataRange().getValues();
  const testStats = {};          // test -> { attempts, sumPct, studentSet }
  const studentTestBest = {};    // "studentId||test" -> best record
  const studentTestAttempts = {};// "studentId||test" -> attempt count

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const timestamp = row[0];
    const studentId = String(row[1] || '').trim();
    const studentName = String(row[2] || '').trim();
    const test = String(row[3] || '').trim();
    const score = Number(row[4]);
    const total = Number(row[5]);
    const pct = Number(row[6]);
    if (!studentId || !test) continue;

    if (!testStats[test]) testStats[test] = { attempts: 0, sumPct: 0, studentSet: {} };
    testStats[test].attempts++;
    testStats[test].sumPct += pct;
    testStats[test].studentSet[studentId] = true;

    const key = studentId + '||' + test;
    studentTestAttempts[key] = (studentTestAttempts[key] || 0) + 1;

    const best = studentTestBest[key];
    if (!best || pct > best.bestPct) {
      studentTestBest[key] = {
        studentId: studentId,
        studentName: studentName,
        test: test,
        bestScore: score,
        bestTotal: total,
        bestPct: pct,
        lastDate: timestamp
      };
    }
  }

  const tests = Object.keys(testStats).map(function (test) {
    return {
      label: test,
      attempts: testStats[test].attempts,
      avgPct: Math.round(testStats[test].sumPct / testStats[test].attempts),
      uniqueStudents: Object.keys(testStats[test].studentSet).length
    };
  });

  const leaderboard = {};
  Object.keys(studentTestBest).forEach(function (key) {
    const rec = studentTestBest[key];
    rec.attempts = studentTestAttempts[key];
    rec.lastDate = rec.lastDate ? new Date(rec.lastDate).toISOString() : '';
    if (!leaderboard[rec.test]) leaderboard[rec.test] = [];
    leaderboard[rec.test].push(rec);
  });
  Object.keys(leaderboard).forEach(function (test) {
    leaderboard[test].sort(function (a, b) { return b.bestPct - a.bestPct; });
  });

  return { tests: tests, leaderboard: leaderboard };
}
