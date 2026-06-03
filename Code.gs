/* ══════════════════════════════════════════════════════════════
   Ascendant 2026 — OC Application Backend (Google Apps Script)
   
   SETUP INSTRUCTIONS:
   1. Go to https://script.google.com  →  New Project
   2. Replace the default code with this entire file
   3. Click the gear icon (Project Settings):
      - Change "GCP project" → Leave as default
   4. Run the `setup()` function ONCE to create the Sheet + Drive folder
   5. Deploy:
      - Deploy → New deployment
      - Type: Web app
      - Execute as: Me
      - Who has access: Anyone
   6. Copy the deployed URL and paste it into register.html
      (replace YOUR_APPS_SCRIPT_WEB_APP_URL)
   ══════════════════════════════════════════════════════════════ */

/* ─── CONFIG ─── */
const SHEET_NAME = 'OC Applications';
const FOLDER_NAME = 'Ascendant 2026 – CV Uploads';

/* ─── ONE-TIME SETUP ─── */
function setup() {
  // Create or find the spreadsheet
  let ss;
  const files = DriveApp.getFilesByName(SHEET_NAME);
  if (files.hasNext()) {
    ss = SpreadsheetApp.open(files.next());
    Logger.log('Sheet already exists: ' + ss.getUrl());
  } else {
    ss = SpreadsheetApp.create(SHEET_NAME);
    Logger.log('Created sheet: ' + ss.getUrl());
  }

  // Add headers to the first sheet
  const sheet = ss.getSheets()[0];
  const headers = [
    'Timestamp',
    'Email',
    'Full Name',
    'Admission Number',
    'Contact Number',
    'Grade',
    'Section',
    'OC Role',
    'Suitability Statement',
    'CV/Portfolio Link'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  sheet.setFrozenRows(1);

  // Create or find the Drive folder for uploads
  const folders = DriveApp.getFoldersByName(FOLDER_NAME);
  if (folders.hasNext()) {
    Logger.log('Folder already exists: ' + folders.next().getUrl());
  } else {
    const folder = DriveApp.createFolder(FOLDER_NAME);
    Logger.log('Created folder: ' + folder.getUrl());
  }

  Logger.log('✅ Setup complete. Now deploy as a Web App.');
}

/* ─── HANDLE POST REQUESTS ─── */
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    // ── Validate required fields ──
    const required = ['fullName', 'admissionNo', 'contact', 'grade', 'section', 'role', 'statement'];
    for (const key of required) {
      if (!data[key] || data[key].toString().trim() === '') {
        return jsonResponse('error', 'Missing required field: ' + key);
      }
    }

    // ── Handle file upload (Base64 → Drive) ──
    let fileUrl = '';
    if (data.file && data.file.data) {
      const folder = DriveApp.getFoldersByName(FOLDER_NAME).next();
      const decoded = Utilities.base64Decode(data.file.data);
      const blob = Utilities.newBlob(decoded, data.file.type, data.file.name);
      const file = folder.createFile(blob);
      file.setDescription('Uploaded by: ' + data.fullName + ' (' + data.admissionNo + ')');
      fileUrl = file.getUrl();
    }

    // ── Append row to Sheet ──
    const ss = DriveApp.getFilesByName(SHEET_NAME).next();
    const sheet = SpreadsheetApp.open(ss).getSheets()[0];

    sheet.appendRow([
      new Date(),                          // Timestamp
      Session.getActiveUser().getEmail(),   // Email (script owner)
      data.fullName,
      data.admissionNo,
      data.contact,
      data.grade,
      data.section,
      data.role,
      data.statement,
      fileUrl || 'No file uploaded'
    ]);

    return jsonResponse('success', 'Application received.');

  } catch (err) {
    return jsonResponse('error', err.message);
  }
}

/* ─── CORS-friendly GET fallback ─── */
function doGet(e) {
  return jsonResponse('ok', 'Ascendant 2026 OC endpoint is live.');
}

/* ─── JSON helper ─── */
function jsonResponse(status, message) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: status, message: message }))
    .setMimeType(ContentService.MimeType.JSON);
}
