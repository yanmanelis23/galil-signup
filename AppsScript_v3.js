// ═══════════════════════════════════════════════════
// גליל עליון — Google Apps Script v3
// חדש בגרסה זו: getSchedule + saveSchedule
// ═══════════════════════════════════════════════════

const SHEET_NAME = 'registrations';
const MAX_REG = 15;
const MIN_REG = 5;
const NOTIFY_EMAIL = 'yanmanelis23@gmail.com';

const DEFAULT_SCHEDULE = {
  'א': { label: 'יום א׳', slots: [{time:'14:00',coach:'יאן',key:'yan'},{time:'15:00',coach:'יאן',key:'yan'},{time:'16:00',coach:'יאן',key:'yan'},{time:'17:00',coach:'יאן',key:'yan'}]},
  'ב': { label: 'יום ב׳', slots: [{time:'15:00',coach:'תו',key:'tav'},{time:'16:00',coach:'תו',key:'tav'},{time:'17:00',coach:'יאן',key:'yan'}]},
  'ג': { label: 'יום ג׳', slots: [{time:'14:00',coach:'תו',key:'tav'},{time:'15:00',coach:'תו',key:'tav'}]},
  'ד': { label: 'יום ד׳', slots: [{time:'16:00',coach:'יאן',key:'yan'},{time:'17:00',coach:'תו',key:'tav'}]},
  'ה': { label: 'יום ה׳', slots: [{time:'14:00',coach:'יאן',key:'yan'},{time:'15:00',coach:'יאן',key:'yan'},{time:'16:00',coach:'תו',key:'tav'}]},
};

// ─── GET ───────────────────────────────────────────
function doGet(e) {
  const action = e.parameter.action;

  if (action === 'getSlotCounts') {
    return json_({ success: true, counts: getSlotCounts_() });
  }
  if (action === 'getAllRegistrations') {
    return json_({ success: true, registrations: getAllRegistrations_() });
  }
  if (action === 'getSchedule') {
    return json_({ success: true, schedule: getSchedule_() });
  }

  return json_({ success: false, error: 'Unknown action' });
}

// ─── POST ──────────────────────────────────────────
function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); }
  catch(err) { return json_({ success: false, error: 'Invalid JSON' }); }

  if (body.action === 'register')    return json_(registerPlayer_(body));
  if (body.action === 'removeReg')   return json_(removeReg_(body));
  if (body.action === 'resetSlot')   return json_(resetSlot_(body.slot_key));
  if (body.action === 'resetDay')    return json_(resetDay_(body.slot_keys));
  if (body.action === 'resetAll')    return json_(resetAll_());
  if (body.action === 'saveSchedule') {
    PropertiesService.getScriptProperties().setProperty('schedule', JSON.stringify(body.schedule));
    return json_({ success: true });
  }

  return json_({ success: false, error: 'Unknown action' });
}

// ─── HELPERS ───────────────────────────────────────
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet_() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
}

// ─── SLOT COUNTS ───────────────────────────────────
function getSlotCounts_() {
  const sheet = getSheet_();
  const data = sheet.getDataRange().getValues();
  const counts = {};
  for (let i = 1; i < data.length; i++) {
    const key = data[i][6]; // slot_key column
    if (key) counts[key] = (counts[key] || 0) + 1;
  }
  return counts;
}

// ─── ALL REGISTRATIONS ─────────────────────────────
function getAllRegistrations_() {
  const sheet = getSheet_();
  const data = sheet.getDataRange().getValues();
  const regs = {};
  for (let i = 1; i < data.length; i++) {
    const [timestamp, name, team, day, time, coach, key] = data[i];
    if (!key) continue;
    if (!regs[key]) regs[key] = [];
    regs[key].push({ name, team, day, time, coach, timestamp: String(timestamp) });
  }
  return regs;
}

// ─── SCHEDULE (stored in Script Properties) ────────
function getSchedule_() {
  const saved = PropertiesService.getScriptProperties().getProperty('schedule');
  if (saved) {
    try { return JSON.parse(saved); } catch(e) {}
  }
  return DEFAULT_SCHEDULE;
}

// ─── REGISTER ──────────────────────────────────────
function registerPlayer_(body) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = getSheet_();
    const counts = getSlotCounts_();
    const k = body.slot_key;
    const count = counts[k] || 0;
    if (count >= MAX_REG) return { success: false, error: 'SLOT_FULL', count };

    sheet.appendRow([new Date(), body.name, body.team, body.day, body.time, body.coach, k]);
    const newCount = count + 1;

    if (newCount === MIN_REG) {
      try {
        MailApp.sendEmail(
          NOTIFY_EMAIL,
          `✅ אימון מתקיים! ${body.day} ${body.time} — ${body.coach}`,
          `הגיע המינימום (${MIN_REG} נרשמים) עבור ${body.day} ${body.time} עם ${body.coach}.`
        );
      } catch(mailErr) {}
    }

    return { success: true, count: newCount };
  } catch(e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

// ─── REMOVE REGISTRATION ───────────────────────────
function removeReg_(body) {
  const sheet = getSheet_();
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][6] === body.slot_key &&
        data[i][1] === body.name &&
        String(data[i][0]) === body.timestamp) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, error: 'Not found' };
}

// ─── RESET SLOT ────────────────────────────────────
function resetSlot_(key) {
  const sheet = getSheet_();
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][6] === key) sheet.deleteRow(i + 1);
  }
  return { success: true };
}

// ─── RESET DAY ─────────────────────────────────────
function resetDay_(keys) {
  const sheet = getSheet_();
  const data = sheet.getDataRange().getValues();
  const keySet = new Set(keys);
  for (let i = data.length - 1; i >= 1; i--) {
    if (keySet.has(data[i][6])) sheet.deleteRow(i + 1);
  }
  return { success: true };
}

// ─── RESET ALL ─────────────────────────────────────
function resetAll_() {
  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);
  return { success: true };
}
