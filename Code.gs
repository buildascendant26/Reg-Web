/* ══════════════════════════════════════════════════════════════
   Ascendant 2026 — OC Application Backend (Google Apps Script)
   
   SETUP INSTRUCTIONS:
   1. Go to https://script.google.com  →  New Project
   2. Replace the default code with this entire file
   3. Save (Floppy disk icon)
   4. Deploy:
      - Deploy → New deployment
      - Type: Web app
      - Execute as: Me
      - Who has access: Anyone
   5. Copy the deployed URL and paste it into register.html
   ══════════════════════════════════════════════════════════════ */

/* ─── CONFIG ─── */
const SHEET_NAME = 'OC Applications';
const FOLDER_NAME = 'Ascendant 2026 – CV Uploads';

/* ─── ONE-TIME SETUP (OPTIONAL) ─── */
function setup() {
  // Find or Create Sheet
  let ss;
  const files = DriveApp.getFilesByName(SHEET_NAME);
  if (files.hasNext()) {
    ss = SpreadsheetApp.open(files.next());
    Logger.log('Sheet already exists: ' + ss.getUrl());
  } else {
    ss = SpreadsheetApp.create(SHEET_NAME);
    Logger.log('Created sheet: ' + ss.getUrl());
  }

  // Set up header columns
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

  // Find or Create Folder
  const folders = DriveApp.getFoldersByName(FOLDER_NAME);
  if (folders.hasNext()) {
    Logger.log('Folder already exists: ' + folders.next().getUrl());
  } else {
    const folder = DriveApp.createFolder(FOLDER_NAME);
    Logger.log('Created folder: ' + folder.getUrl());
  }

  Logger.log('✅ Setup complete.');
}

/* ─── HANDLE POST REQUESTS ─── */
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    // 1. Validate required fields
    const required = ['fullName', 'email', 'admissionNo', 'contact', 'grade', 'section', 'role', 'statement'];
    for (const key of required) {
      if (!data[key] || data[key].toString().trim() === '') {
        return jsonResponse('error', 'Missing required field: ' + key);
      }
    }

    // 2. Find or Create Uploads Folder
    let folder;
    const folders = DriveApp.getFoldersByName(FOLDER_NAME);
    if (folders.hasNext()) {
      folder = folders.next();
    } else {
      folder = DriveApp.createFolder(FOLDER_NAME);
    }

    // 3. Handle File Stream (Base64 to Drive File)
    let fileUrl = '';
    if (data.file && data.file.data) {
      const decoded = Utilities.base64Decode(data.file.data);
      const blob = Utilities.newBlob(decoded, data.file.type, data.file.name);
      const file = folder.createFile(blob);
      file.setDescription('Uploaded by: ' + data.fullName + ' (' + data.admissionNo + ')');
      fileUrl = file.getUrl();
    }

    // 4. Find or Create Spreadsheet
    let ss;
    const files = DriveApp.getFilesByName(SHEET_NAME);
    if (files.hasNext()) {
      ss = SpreadsheetApp.open(files.next());
    } else {
      ss = SpreadsheetApp.create(SHEET_NAME);
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
    }
    
    const sheet = ss.getSheets()[0];

    // 5. Append Submission Row
    sheet.appendRow([
      new Date(),                          // Timestamp
      data.email,                          // Email
      data.fullName,                       // Full Name
      data.admissionNo,                    // Admission Number
      data.contact,                        // Contact Number
      data.grade,                          // Grade
      data.section,                        // Section
      data.role,                           // OC Role
      data.statement,                      // Suitability Statement
      fileUrl || 'No file uploaded'        // CV File Link
    ]);

    return jsonResponse('success', 'Application submitted successfully.');

  } catch (err) {
    return jsonResponse('error', err.message);
  }
}

/* ─── GET FALLBACK ─── */
function doGet(e) {
  return jsonResponse('ok', 'Verification: Version 3 is successfully deployed and active.');
}

/* ─── UTILITY: JSON Output ─── */
function jsonResponse(status, message) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: status, message: message }))
    .setMimeType(ContentService.MimeType.JSON);
}
