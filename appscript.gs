function getLessonSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName('Lessons');
  if (!sheet) {
    sheet = spreadsheet.insertSheet('Lessons');
  }
  const headers = ['day', 'title', 'description', 'youtube', 'image', 'resource', 'codingTask', 'remarks'];
  const existing = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  const hasHeaders = existing.some((value) => String(value).trim());
  if (!hasHeaders) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  return sheet;
}

function getSubmissionSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName('Submissions');
  if (!sheet) {
    sheet = spreadsheet.insertSheet('Submissions');
  }
  const headers = ['day', 'lessonTitle', 'name', 'colabLink', 'submittedAt'];
  const existing = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  const hasHeaders = existing.some((value) => String(value).trim());
  if (!hasHeaders) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  return sheet;
}

function getNoticeSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName('Notice');
  if (!sheet) {
    sheet = spreadsheet.insertSheet('Notice');
  }
  const headers = ['message', 'updatedAt'];
  const existing = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  const hasHeaders = existing.some((value) => String(value).trim());
  if (!hasHeaders) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  return sheet;
}

function buildLessonRows(sheet) {
  const values = sheet.getDataRange().getDisplayValues();
  if (!values || values.length < 2) return [];
  const headers = values[0].map((header) => String(header).trim().toLowerCase());
  return values.slice(1).map((row, index) => {
    const item = {};
    headers.forEach((header, headerIndex) => {
      item[header] = row[headerIndex] || '';
    });
    item.rowIndex = index + 2;
    return item;
  });
}

function buildSubmissionRows(sheet) {
  const values = sheet.getDataRange().getDisplayValues();
  if (!values || values.length < 2) return [];
  const headers = values[0].map((header) => String(header).trim().toLowerCase());
  return values.slice(1).map((row, index) => {
    const item = {};
    headers.forEach((header, headerIndex) => {
      item[header] = row[headerIndex] || '';
    });
    item.rowIndex = index + 2;
    return item;
  });
}

function buildNoticeRows(sheet) {
  const values = sheet.getDataRange().getDisplayValues();
  if (!values || values.length < 2) return [];
  const headers = values[0].map((header) => String(header).trim().toLowerCase());
  return values.slice(1).map((row, index) => {
    const item = {};
    headers.forEach((header, headerIndex) => {
      item[header] = row[headerIndex] || '';
    });
    item.rowIndex = index + 2;
    return item;
  });
}

function parseRequest(e) {
  try {
    const rawBody = e && e.postData && e.postData.getDataAsString ? e.postData.getDataAsString() : '';
    if (rawBody) {
      const parsed = JSON.parse(rawBody);
      return parsed;
    }
  } catch (err) {}
  return {
    action: e && e.parameter ? e.parameter.action : '',
    payload: e && e.parameter ? e.parameter.payload : {},
  };
}

function doGet(e) {
  const action = String((e && e.parameter && e.parameter.action) || 'get-lessons').toLowerCase();
  if (action === 'get-submissions') {
    const sheet = getSubmissionSheet();
    const rows = buildSubmissionRows(sheet);
    return ContentService.createTextOutput(JSON.stringify({ rows }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  if (action === 'get-notice' || action === 'get-notices') {
    const sheet = getNoticeSheet();
    const rows = buildNoticeRows(sheet);
    const notice = rows.length ? rows[rows.length - 1] : null;
    return ContentService.createTextOutput(JSON.stringify({ notice }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  const sheet = getLessonSheet();
  const rows = buildLessonRows(sheet);
  return ContentService.createTextOutput(JSON.stringify({ rows }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const request = parseRequest(e);
  const action = String(request.action || '').toLowerCase();
  const payload = request.payload || {};

  if (action === 'add-lesson') {
    const sheet = getLessonSheet();
    sheet.appendRow([
      payload.day || '',
      payload.title || '',
      payload.description || '',
      payload.youtube || '',
      payload.image || '',
      payload.resource || '',
      payload.codingTask || '',
      payload.remarks || ''
    ]);
  } else if (action === 'update-lesson') {
    const sheet = getLessonSheet();
    const rowIndex = Number(payload.rowIndex || 0);
    const targetRow = rowIndex > 1 ? rowIndex : 2;
    sheet.getRange(targetRow, 1, 1, 8).setValues([[
      payload.day || '',
      payload.title || '',
      payload.description || '',
      payload.youtube || '',
      payload.image || '',
      payload.resource || '',
      payload.codingTask || '',
      payload.remarks || ''
    ]]);
  } else if (action === 'delete-lesson') {
    const sheet = getLessonSheet();
    const rowIndex = Number(payload.rowIndex || 0);
    const targetRow = rowIndex > 1 ? rowIndex : 2;
    sheet.deleteRow(targetRow);
  } else if (action === 'add-submission') {
    const sheet = getSubmissionSheet();
    sheet.appendRow([
      payload.day || '',
      payload.lessonTitle || '',
      payload.name || '',
      payload.colabLink || '',
      new Date().toISOString()
    ]);
  } else if (action === 'delete-submission') {
    const sheet = getSubmissionSheet();
    const rowIndex = Number(payload.rowIndex || 0);
    const targetRow = rowIndex > 1 ? rowIndex : 2;
    sheet.deleteRow(targetRow);
  } else if (action === 'save-notice') {
    const sheet = getNoticeSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.deleteRows(2, lastRow - 1);
    }
    const message = String(payload.message || '').trim();
    if (message) {
      sheet.appendRow([message, new Date().toISOString()]);
    }
  }

  return ContentService.createTextOutput(JSON.stringify({ success: true }))
    .setMimeType(ContentService.MimeType.JSON);
}
