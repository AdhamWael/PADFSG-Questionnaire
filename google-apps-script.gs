const RESPONSES_SHEET = 'Responses';
const REPORTS_FOLDER_NAME = 'PADFSG Questionnaire Reports';
const REPORT_RECIPIENT_EMAIL = ''; // Optional: enter the address that should always receive reports.
const SEND_COPY_TO_RESPONDENT = false;

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
      pdfUrl: report.pdfUrl,
      emailedTo: report.emailedTo
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

  const excluded = new Set(['respName', 'respEmail', 'respRole', '_submittedAt', '_submissionId', '_reportAnswers', '_form', '_subject', 'email', 'name']);
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
  body.setMarginTop(54).setMarginBottom(54).setMarginLeft(54).setMarginRight(54);

  appendReportTitle(body, 'PADFSG Website Discovery Report', 'Client responses, interpretation and recommended next steps');
  appendMetaTable(body, reference, data);

  const analysis = buildReportAnalysis(data);
  appendReportSection(body, 'Executive summary');
  appendCallout(body, `${analysis.scope}\n${analysis.summary}`);

  appendReportSection(body, 'What the answers suggest');
  appendInsight(body, 'Recommended experience', analysis.experience);
  appendInsight(body, 'Recommended website structure', analysis.structure);
  appendInsight(body, 'Delivery and pricing considerations', analysis.delivery);
  appendInsight(body, 'Content and approval readiness', analysis.readiness);

  appendReportSection(body, 'Recommended next steps');
  analysis.nextSteps.forEach(step => body.appendListItem(step).setGlyphType(DocumentApp.GlyphType.NUMBER));

  appendReportSection(body, 'Questions and answers');
  const readableAnswers = parseReadableAnswers(data._reportAnswers);
  appendReadableAnswers(body, readableAnswers.length ? readableAnswers : fallbackAnswers(data));

  body.appendHorizontalRule();
  const closing = body.appendParagraph(`Prepared automatically from submission ${reference}. Recommendations should be confirmed during the project scoping meeting.`);
  closing.editAsText().setFontSize(8).setForegroundColor('#68766F');

  document.saveAndClose();
  const docFile = DriveApp.getFileById(document.getId());
  docFile.moveTo(folder);
  const pdfName = `${reference} - PADFSG Website Discovery Report.pdf`;
  const pdfBlob = docFile.getAs(MimeType.PDF).setName(pdfName);
  const pdfFile = folder.createFile(pdfBlob);
  const emailedTo = emailSubmissionReport(data, reference, pdfBlob, document.getUrl(), pdfFile.getUrl());

  return { documentUrl: document.getUrl(), pdfUrl: pdfFile.getUrl(), emailedTo };
}

function appendReportTitle(body, title, subtitle) {
  const titleParagraph = body.appendParagraph(title).setHeading(DocumentApp.ParagraphHeading.TITLE);
  titleParagraph.editAsText().setForegroundColor('#173C31').setBold(true);
  const subtitleParagraph = body.appendParagraph(subtitle);
  subtitleParagraph.editAsText().setFontSize(12).setForegroundColor('#9B5A36');
  body.appendHorizontalRule();
}

function appendMetaTable(body, reference, data) {
  const rows = [
    ['Reference', reference],
    ['Respondent', data.respName || 'Not provided'],
    ['Role', data.respRole || 'Not provided'],
    ['Email', data.respEmail || 'Not provided'],
    ['Submitted', Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMMM yyyy, HH:mm')]
  ];
  const table = body.appendTable(rows);
  for (let row = 0; row < rows.length; row += 1) {
    table.getCell(row, 0).setBackgroundColor('#EAF1EC');
    table.getCell(row, 0).editAsText().setBold(true).setForegroundColor('#173C31');
    table.getCell(row, 1).editAsText().setForegroundColor('#26352E');
  }
}

function appendReportSection(body, title) {
  const paragraph = body.appendParagraph(title).setHeading(DocumentApp.ParagraphHeading.HEADING1);
  paragraph.editAsText().setForegroundColor('#173C31').setBold(true);
  return paragraph;
}

function appendCallout(body, text) {
  const table = body.appendTable([[text]]);
  const cell = table.getCell(0, 0).setBackgroundColor('#F4EEDC');
  cell.editAsText().setForegroundColor('#173C31').setFontSize(11);
}

function appendInsight(body, title, text) {
  const heading = body.appendParagraph(title).setHeading(DocumentApp.ParagraphHeading.HEADING2);
  heading.editAsText().setForegroundColor('#9B5A36').setBold(true);
  body.appendParagraph(text);
}

function parseReadableAnswers(raw) {
  if (!raw) return [];
  try {
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed.filter(item => item && item.question && item.answer) : [];
  } catch (error) {
    return [];
  }
}

function appendReadableAnswers(body, answers) {
  let currentSection = '';
  answers.forEach(item => {
    if (item.section !== currentSection) {
      currentSection = item.section;
      const section = body.appendParagraph(currentSection).setHeading(DocumentApp.ParagraphHeading.HEADING2);
      section.editAsText().setForegroundColor('#2F6B49').setBold(true);
    }
    const question = body.appendParagraph(item.question);
    question.editAsText().setBold(true).setForegroundColor('#26352E').setFontSize(10);
    const answer = body.appendParagraph(item.answer);
    answer.editAsText().setForegroundColor('#5A675F').setFontSize(10);
    answer.setSpacingAfter(8);
  });
}

function fallbackAnswers(data) {
  return Object.keys(data)
    .filter(key => !key.startsWith('_') && !['respName', 'respEmail', 'respRole', 'respDate', 'email', 'name'].includes(key))
    .map(key => ({ section: 'Website discovery', question: formatLabel(key), answer: normaliseValue(data[key]) }))
    .filter(item => item.answer !== '');
}

function values(data, key) {
  return String(data[key] || '').split(',').map(value => value.trim()).filter(Boolean);
}

function friendlyList(items, labels) {
  const translated = items.map(item => labels[item] || formatLabel(item));
  if (!translated.length) return 'No specific preference was selected.';
  if (translated.length === 1) return translated[0];
  return `${translated.slice(0, -1).join(', ')} and ${translated[translated.length - 1]}`;
}

function buildReportAnalysis(data) {
  const scope = data._proposalScope || 'Website scope to be confirmed';
  const impressionLabels = {
    'trusted-science': 'scientific trust', 'african-leadership': 'African leadership',
    'welcoming-network': 'an inclusive network', 'practical-help': 'practical usefulness',
    'active-movement': 'visible activity', 'established-body': 'institutional credibility',
    'future-focused': 'forward-looking ambition', 'patient-impact': 'patient-centred impact'
  };
  const metaphorLabels = {
    'leading-institute': 'a respected medical institute', 'continental-network': 'a connected continental meeting place',
    'learning-centre': 'an accessible learning centre', 'public-campaign': 'a movement for change'
  };
  const areaLabels = {
    about: 'About PADFSG', governance: 'Leadership and governance', regions: 'Regions and countries',
    projects: 'Projects and impact', research: 'Research and publications', resources: 'Clinical resources',
    education: 'Education and webinars', events: 'Events and conferences', membership: 'Membership',
    partners: 'Partners and supporters', news: 'News and media', awards: 'Awards', patients: 'Patient information'
  };
  const impressions = friendlyList(values(data, 'desiredImpression'), impressionLabels);
  const metaphor = metaphorLabels[data.experienceMetaphor] || 'a balanced professional organisation';
  const density = data.informationComfort === 'simple' ? 'short and highly focused pages' : data.informationComfort === 'rich' ? 'a rich experience with deeper information available' : 'clear summaries with optional detail';
  const areas = friendlyList(values(data, 'workAreas'), areaLabels);
  const features = values(data, 'platformFeatures');
  const complexFeatures = features.filter(item => ['member-accounts', 'payments', 'events', 'certificates', 'directory', 'interactive-map', 'translation'].includes(item));
  const pricingFactors = data._pricingFactors || 'No major complexity factors were identified.';

  const summary = `The answers point to ${metaphor}, built around ${impressions}. The recommended content approach is ${density}.`;
  const experience = `Lead with ${impressions}. Use ${density}, with African professionals, countries, results and patient impact shown according to the selected priorities.`;
  const structure = `The requested launch structure includes: ${areas}. Confirm the final navigation by testing the most important visitor actions before pages are written.`;
  const delivery = `${scope}. ${features.length} online activities were selected${complexFeatures.length ? `, including ${friendlyList(complexFeatures, {})}` : ''}. Main pricing considerations: ${pricingFactors}.`;
  const readiness = data.contentPreparation === 'full-support'
    ? 'PADFSG expects full support with research, writing and image selection; content production should be a separate priced workstream.'
    : data.contentPreparation === 'shared'
      ? 'PADFSG expects a shared content process; allow time for interviews, editing and approvals.'
      : 'PADFSG expects to provide finished content; confirm completeness and quality before the build schedule is fixed.';

  const nextSteps = [
    'Hold a short scope-confirmation meeting to agree the essential first release and any later phases.',
    'Turn the selected website areas into a proposed sitemap and confirm the top visitor journeys.',
    'Prepare a content checklist with an owner, approval person and deadline for every page.',
    'Create a homepage and key-page prototype that reflects the inferred experience and test it with PADFSG stakeholders.',
    'Confirm the services, languages and ongoing support included in the final quotation.',
    'Agree the build, content-entry, review, testing and launch schedule.'
  ];
  if (data.membershipReadiness === 'new') nextSteps.splice(2, 0, 'Define membership categories, fees, benefits and approval rules before building the member journey.');
  if (data.regionalReadiness === 'future') nextSteps.splice(2, 0, 'Agree a simple regional structure that can grow as more countries become active.');
  if (values(data, 'additionalLanguages').filter(item => item !== 'none').length) nextSteps.splice(-1, 0, 'Confirm who translates and approves each language, and include translation time in the schedule.');

  return { scope, summary, experience, structure, delivery, readiness, nextSteps };
}

function emailSubmissionReport(data, reference, pdfBlob, documentUrl, pdfUrl) {
  const ownerEmail = REPORT_RECIPIENT_EMAIL || Session.getEffectiveUser().getEmail();
  const recipients = [];
  if (ownerEmail) recipients.push(ownerEmail);
  if (SEND_COPY_TO_RESPONDENT && data.respEmail && !recipients.includes(data.respEmail)) recipients.push(data.respEmail);
  if (!recipients.length) return 'Not sent - add REPORT_RECIPIENT_EMAIL';

  try {
    MailApp.sendEmail({
      to: recipients.join(','),
      subject: `${reference} - PADFSG website discovery report`,
      name: 'PADFSG Website Discovery',
      body: `A new PADFSG website discovery response has been received from ${data.respName || 'a respondent'}. The attached PDF includes their answers, scope analysis and recommended next steps. Google Document: ${documentUrl} PDF: ${pdfUrl}`,
      htmlBody: `<p>A new PADFSG website discovery response has been received from <strong>${escapeHtml(data.respName || 'a respondent')}</strong>.</p><p>The PDF report is attached and includes their answers, scope analysis and recommended next steps.</p><p><a href="${documentUrl}">Open Google Document</a> &nbsp;|&nbsp; <a href="${pdfUrl}">Open PDF in Drive</a></p>`,
      attachments: [pdfBlob]
    });
    return recipients.join(', ');
  } catch (error) {
    console.error(`Report email failed: ${error.message}`);
    return `Not sent - ${error.message}`;
  }
}

function escapeHtml(value) {
  return String(value || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
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
