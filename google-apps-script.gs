const RESPONSES_SHEET = 'Responses';
const REPORTS_FOLDER_NAME = 'PADFSG Questionnaire Reports';

function doPost(e) {
  try {
    const data = parseRequest(e);
    if (!data._submissionId || !data.respEmail) throw new Error('Invalid submission payload.');

    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    let result;
    try {
      const sheet = getResponseSheet();
      const duplicate = findSubmission(sheet, data._submissionId);
      if (duplicate) return jsonResponse({ success: true, duplicate: true, reference: duplicate });
      result = appendSubmission(sheet, data);
    } finally {
      lock.releaseLock();
    }

    const sheet = getResponseSheet();
    const report = createSubmissionReport(data, result.reference);

    sheet.getRange(result.rowNumber, result.reportColumn).setValue(report.documentUrl);
    sheet.getRange(result.rowNumber, result.pdfColumn).setValue(report.pdfUrl);

    return jsonResponse({
      success: true,
      reference: result.reference,
      documentUrl: report.documentUrl,
      pdfUrl: report.pdfUrl
    });
  } catch (error) {
    return jsonResponse({ success: false, message: error.message });
  }
}

function parseRequest(e) {
  if (!e || !e.postData || !e.postData.contents) throw new Error('Empty request.');
  return JSON.parse(e.postData.contents);
}

function doGet() {
  return jsonResponse({ success: true, service: 'PADFSG Questionnaire Receiver' });
}

function getResponseSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  return spreadsheet.getSheetByName(RESPONSES_SHEET) || spreadsheet.insertSheet(RESPONSES_SHEET);
}

function appendSubmission(sheet, data) {
  const fixedFields = {
    Reference: createReference(),
    'Submission ID': data._submissionId || '',
    'Received At': new Date(),
    'Respondent Name': data.respName || '',
    'Respondent Email': data.respEmail || '',
    'Respondent Role': data.respRole || '',
    'Google Document': '',
    'PDF Report': ''
  };

  const excluded = new Set(['respName', 'respEmail', 'respRole', '_submittedAt', '_submissionId', '_form', '_subject', 'email', 'name']);
  const dynamicFields = {};
  Object.keys(data).forEach(key => {
    if (!excluded.has(key)) dynamicFields[formatLabel(key)] = safeSheetValue(normaliseValue(data[key]));
  });

  const completeData = { ...fixedFields, ...dynamicFields };
  let headers = getHeaders(sheet);

  if (!headers.length) {
    headers = Object.keys(completeData);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    formatHeader(sheet, headers.length);
  } else {
    const missing = Object.keys(completeData).filter(header => !headers.includes(header));
    if (missing.length) {
      sheet.getRange(1, headers.length + 1, 1, missing.length).setValues([missing]);
      headers = [...headers, ...missing];
      formatHeader(sheet, headers.length);
    }
  }

  sheet.appendRow(headers.map(header => completeData[header] ?? ''));
  const rowNumber = sheet.getLastRow();
  sheet.getRange(rowNumber, 1, 1, headers.length).setVerticalAlignment('top').setWrap(true);

  return {
    reference: fixedFields.Reference,
    rowNumber,
    reportColumn: headers.indexOf('Google Document') + 1,
    pdfColumn: headers.indexOf('PDF Report') + 1
  };
}

function findSubmission(sheet, submissionId) {
  const headers = getHeaders(sheet);
  const idColumn = headers.indexOf('Submission ID') + 1;
  const referenceColumn = headers.indexOf('Reference') + 1;
  if (!idColumn || sheet.getLastRow() < 2) return '';
  const match = sheet.getRange(2, idColumn, sheet.getLastRow() - 1, 1)
    .createTextFinder(submissionId).matchEntireCell(true).findNext();
  return match && referenceColumn ? sheet.getRange(match.getRow(), referenceColumn).getDisplayValue() : '';
}

function getHeaders(sheet) {
  if (!sheet.getLastColumn()) return [];
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].filter(String);
}

function createSubmissionReport(data, reference) {
  const folder = getOrCreateFolder(REPORTS_FOLDER_NAME);
  const respondentName = data.respName || 'Unnamed Respondent';
  const document = DocumentApp.create(`${reference} - PADFSG Website Discovery - ${respondentName}`);
  const body = document.getBody();

  body.appendParagraph('PADFSG Website Discovery Questionnaire').setHeading(DocumentApp.ParagraphHeading.TITLE);
  body.appendParagraph(`Reference: ${reference}`);
  body.appendParagraph(`Submitted: ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMMM yyyy, HH:mm')}`);
  body.appendHorizontalRule();

  appendSection(body, 'Respondent details', {
    Name: data.respName,
    Email: data.respEmail,
    Role: data.respRole,
    Date: data.respDate
  });

  const answers = {};
  Object.keys(data).forEach(key => {
    if (!key.startsWith('_') && !['respName', 'respEmail', 'respRole', 'respDate'].includes(key)) {
      answers[formatLabel(key)] = normaliseValue(data[key]);
    }
  });
  appendSection(body, 'Questionnaire responses', answers);

  document.saveAndClose();
  const docFile = DriveApp.getFileById(document.getId());
  docFile.moveTo(folder);
  const pdfFile = folder.createFile(docFile.getAs(MimeType.PDF).setName(`${reference} - PADFSG Website Discovery Report.pdf`));

  return { documentUrl: document.getUrl(), pdfUrl: pdfFile.getUrl() };
}

function appendSection(body, title, values) {
  const entries = Object.entries(values).filter(([, value]) => value !== '' && value != null);
  if (!entries.length) return;
  body.appendParagraph(title).setHeading(DocumentApp.ParagraphHeading.HEADING1);
  entries.forEach(([label, value]) => {
    body.appendParagraph(label).setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph(String(value));
  });
}

function getOrCreateFolder(name) {
  const folders = DriveApp.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(name);
}

function formatHeader(sheet, count) {
  sheet.getRange(1, 1, 1, count)
    .setFontWeight('bold')
    .setBackground('#315C46')
    .setFontColor('#FFFFFF')
    .setWrap(true);
  sheet.setFrozenRows(1);
}

function formatLabel(key) {
  return key.replace(/([a-z])([A-Z])/g, '$1 $2').replace(/[_-]+/g, ' ').replace(/\b\w/g, c => c.toUpperCase()).trim();
}

function normaliseValue(value) {
  return Array.isArray(value) ? value.join(', ') : (value ?? '');
}

function safeSheetValue(value) {
  const text = String(value ?? '');
  return /^[=+\-@]/.test(text) ? "'" + text : text;
}

function createReference() {
  const date = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd');
  const random = Math.random().toString(36).substring(2, 7).toUpperCase();
  return `PADFSG-${date}-${random}`;
}

function jsonResponse(payload) {
  return ContentService.createTextOutput(JSON.stringify(payload)).setMimeType(ContentService.MimeType.JSON);
}
