/**
 * COLLECTIONS — daily automatic backup of the Firestore data to Google Drive.
 * Runs inside the owner's own Google account (Google Apps Script, free).
 * Backups go to the Drive folder "COLLECTIONS Backups" (private to you).
 */
const PROJECT_ID  = "collections-388ec";
const API_KEY     = "AIzaSyD-53Kz_6_8WfxXe3b76GyCu-alUaDfjFg";
const FOLDER_NAME = "COLLECTIONS Backups";
const KEEP_DAYS   = 90;            // older backup files are moved to Drive trash
const BACKUP_HOUR = 2;             // runs daily around 2 AM

// ---- run this ONE time: creates the daily trigger and takes the first backup ----
function setup() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger("dailyBackup").timeBased().everyDays(1).atHour(BACKUP_HOUR).create();
  dailyBackup();
}

function dailyBackup() {
  try {
    const url = "https://firestore.googleapis.com/v1/projects/" + PROJECT_ID +
      "/databases/(default)/documents/collections_data/main?key=" + API_KEY;
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) throw new Error("Firestore HTTP " + res.getResponseCode() + ": " + res.getContentText().slice(0, 200));
    const raw = JSON.parse(res.getContentText());
    const data = fromFields(raw.fields || {});
    if (!Array.isArray(data.donors) || !Array.isArray(data.batches)) throw new Error("Unexpected data shape - backup NOT saved");

    const folder = getFolder_();
    const stamp = Utilities.formatDate(new Date(), "Asia/Kolkata", "yyyy-MM-dd_HHmm");
    folder.createFile("COLLECTIONS_backup_" + stamp + ".json", JSON.stringify(data, null, 1), MimeType.PLAIN_TEXT);
    cleanup_(folder);
    Logger.log("Backup OK: " + data.donors.length + " members, rev " + data.rev);
  } catch (e) {
    MailApp.sendEmail(Session.getEffectiveUser().getEmail(), "COLLECTIONS backup FAILED", String(e));
    throw e;
  }
}

function getFolder_() {
  const it = DriveApp.getFoldersByName(FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER_NAME);
}

function cleanup_(folder) {
  const limit = new Date().getTime() - KEEP_DAYS * 86400000;
  const files = folder.getFiles();
  while (files.hasNext()) { const f = files.next(); if (f.getDateCreated().getTime() < limit) f.setTrashed(true); }
}

// Firestore REST "typed values" -> plain JSON (same shape as the app's own backup button)
function fromValue(v) {
  if ("stringValue" in v) return v.stringValue;
  if ("integerValue" in v) return Number(v.integerValue);
  if ("doubleValue" in v) return Number(v.doubleValue);
  if ("booleanValue" in v) return v.booleanValue;
  if ("nullValue" in v) return null;
  if ("timestampValue" in v) return v.timestampValue;
  if ("arrayValue" in v) return (v.arrayValue.values || []).map(fromValue);
  if ("mapValue" in v) return fromFields(v.mapValue.fields || {});
  return null;
}
function fromFields(f) { const o = {}; Object.keys(f).forEach(k => { o[k] = fromValue(f[k]); }); return o; }
