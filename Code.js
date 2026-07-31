// Quiz sets are discovered automatically from the spreadsheet's tabs — add a
// new tab and it appears in the app with no code changes required. Everything
// except these reserved sheets is treated as a quiz set.
const RESERVED_SHEETS = ['Members', 'Quiz Results'];

// A tab named "Prefix - Topic" (e.g. "MCN - Sensory") is auto-split into a
// category ("Prefix") and a short label ("Topic"). This map lets several
// prefixes share one display category; any prefix not listed here becomes
// its own category automatically, so new subject areas need no code change.
const CATEGORY_PREFIX_MAP = {
  'MCN': 'MCN - Maternal & Child',
  'Basic': 'Med Surg',
  'Med Surg': 'Med Surg'
};
// Known categories are shown first, in this order; any new/unrecognized
// category discovered from a sheet name is appended after these automatically.
const CATEGORY_ORDER = ['MCN - Maternal & Child', 'Med Surg'];

function splitSheetName_(name) {
  const idx = name.indexOf(' - ');
  if (idx === -1) return { prefix: name.trim(), topic: name.trim() };
  return { prefix: name.slice(0, idx).trim(), topic: name.slice(idx + 3).trim() };
}

// Every spreadsheet tab except the reserved ones is a candidate quiz set.
// Tabs that don't actually contain gradable questions (wrong column layout,
// empty, a scratch/notes tab, etc.) are filtered out later by getTestList()
// once their question count comes back as zero.
function getQuestionSheetNames_() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheets()
    .map(function (s) { return s.getName(); })
    .filter(function (name) { return RESERVED_SHEETS.indexOf(name) === -1; });
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
  getTestListForStudent: getTestListForStudent,
  getQuestionSet: getQuestionSet,
  saveResult: saveResult,
  getResultsSummary: getResultsSummary
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
      return { ok: true, id: id, name: name, isAdmin: isAdminId_(id) };
    }
  }
  return { ok: false };
}

// Counts gradable questions in a sheet the same way getQuestionSet() filters them,
// so the count shown to students always matches what they'll actually be asked.
function countQuestions_(sheetName) {
  const data = getSheet_(sheetName).getDataRange().getValues();
  let count = 0;
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    const answer = String(row[5] || '').trim().toUpperCase();
    if (!answer || answer.indexOf(',') !== -1 || ['A', 'B', 'C', 'D'].indexOf(answer) === -1) continue;
    count++;
  }
  return count;
}

// Building this list means reading every quiz sheet in full just to count its
// questions — slow, and identical for every student. Cache it for a few
// minutes so concurrent students and repeat dashboard visits don't each pay
// that cost. A newly added sheet may take up to this long to appear.
const TEST_LIST_CACHE_SECONDS = 300;

function getTestList() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('testList');
  if (cached) return JSON.parse(cached);

  const tests = getQuestionSheetNames_().map(function (name) {
    const parts = splitSheetName_(name);
    return {
      key: name,
      label: name,
      shortLabel: parts.topic,
      category: CATEGORY_PREFIX_MAP[parts.prefix] || parts.prefix,
      totalQuestions: countQuestions_(name)
    };
  }).filter(function (t) { return t.totalQuestions > 0; });

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

function getCompletedTestsForStudent_(studentId) {
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
  return result;
}

function getQuestionSet(setKey, studentId) {
  const sheetName = setKey;
  if (RESERVED_SHEETS.indexOf(sheetName) !== -1 || !SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName)) {
    throw new Error('Unknown question set: ' + setKey);
  }

  const data = getSheet_(sheetName).getDataRange().getValues();
  const questions = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const question = row[0];
    if (!question) continue;

    const answer = String(row[5] || '').trim().toUpperCase();
    // Skip malformed rows: multi-answer "select all that apply" items (e.g. "A,C,E")
    // aren't gradable by this single-answer A-D quiz format.
    if (!answer || answer.indexOf(',') !== -1 || ['A', 'B', 'C', 'D'].indexOf(answer) === -1) continue;

    questions.push({
      question: String(question),
      choices: {
        A: String(row[1] || ''),
        B: String(row[2] || ''),
        C: String(row[3] || ''),
        D: String(row[4] || '')
      },
      answer: answer,
      rationale: String(row[6] || '')
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
