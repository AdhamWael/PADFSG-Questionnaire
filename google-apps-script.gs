const SUBMISSIONS_SHEET = 'Submissions';
const ANSWERS_SHEET = 'Answers';
const REPORT_LOG_SHEET = 'Report Log';
const DASHBOARD_SHEET = 'Dashboard';
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
      setupPADFSGWorkbook();
      const sheet = getSubmissionSheet();
      const duplicate = findSubmission(sheet, data._submissionId);
      if (duplicate) return jsonResponse({ success: true, duplicate: true, reference: duplicate });
      result = appendSubmission(sheet, data);
      appendAnswers(data, result.reference);
    } finally {
      lock.releaseLock();
    }

    const sheet = getSubmissionSheet();
    let report;
    try {
      report = createSubmissionReport(data, result.reference);
      updateSubmissionReport(sheet, result.rowNumber, {
        'Report Status': 'Complete',
        'Google Drive Folder': report.folderUrl,
        'Google Document': report.documentUrl,
        'PDF Report': report.pdfUrl,
        'Email Status': report.emailedTo,
        'Report Error': ''
      });
      appendReportLog(result.reference, 'Complete', report);
    } catch (reportError) {
      updateSubmissionReport(sheet, result.rowNumber, {
        'Report Status': 'Failed',
        'Email Status': 'Not sent',
        'Report Error': reportError.message
      });
      appendReportLog(result.reference, 'Failed', { error: reportError.message });
      throw new Error(`Response saved as ${result.reference}, but report creation failed: ${reportError.message}`);
    }

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
  return jsonResponse({ success: true, service: 'PADFSG Questionnaire Receiver', version: '2.0-managed-workbook' });
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('PADFSG Reports')
    .addItem('Set up / refresh workbook', 'setupPADFSGWorkbook')
    .addItem('Open reports folder', 'showReportsFolderLink')
    .addToUi();
}

function showReportsFolderLink() {
  const folder = getOrCreateFolder(REPORTS_FOLDER_NAME);
  const html = HtmlService.createHtmlOutput(`<p><a href="${folder.getUrl()}" target="_blank">Open ${REPORTS_FOLDER_NAME}</a></p>`).setWidth(420).setHeight(100);
  SpreadsheetApp.getUi().showModalDialog(html, 'PADFSG reports');
}

function getSubmissionSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  return spreadsheet.getSheetByName(SUBMISSIONS_SHEET) || spreadsheet.insertSheet(SUBMISSIONS_SHEET);
}

function appendSubmission(sheet, data) {
  const fixedFields = {
    Reference: createReference(),
    'Submission ID': data._submissionId || '',
    'Received At': new Date(),
    'Respondent Name': data.respName || '',
    'Respondent Email': data.respEmail || '',
    'Respondent Role': data.respRole || '',
    'Response Date': data.respDate || '',
    'Answer Count': data._answerCount || '',
    'Proposal Scope': data._proposalScope || '',
    'Scope Score': data._scopeScore || '',
    'Pricing Factors': data._pricingFactors || '',
    'Report Status': 'Pending',
    'Google Drive Folder': '',
    'Google Document': '',
    'PDF Report': '',
    'Email Status': 'Pending',
    'Report Error': ''
  };
  const readableAnswers = parseReadableAnswers(data._reportAnswers);
  if (readableAnswers.length) {
    if (data._answerCount && readableAnswers.length !== Number(data._answerCount)) throw new Error('Answer completeness check failed.');
  }

  const completeData = fixedFields;
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
  formatSubmissionRow(sheet, rowNumber, headers.length);

  return {
    reference: fixedFields.Reference,
    rowNumber,
    rowNumber
  };
}

function appendAnswers(data, reference) {
  const sheet = getOrCreateSheet(ANSWERS_SHEET);
  const headers = ['Reference', 'Received At', 'Respondent Name', 'Section', 'Question', 'Answer'];
  ensureSheetHeaders(sheet, headers);
  const answers = parseReadableAnswers(data._reportAnswers);
  if (!answers.length) throw new Error('The complete question-and-answer list was not received.');
  const receivedAt = new Date();
  const rows = answers.map(item => [
    reference,
    receivedAt,
    safeSheetValue(data.respName || ''),
    safeSheetValue(item.section || 'Website discovery'),
    safeSheetValue(item.question || ''),
    safeSheetValue(item.answer || 'No answer')
  ]);
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, headers.length).setValues(rows).setVerticalAlignment('top').setWrap(true);
}

function updateSubmissionReport(sheet, rowNumber, updates) {
  const headers = getHeaders(sheet);
  Object.keys(updates).forEach(header => {
    const column = headers.indexOf(header) + 1;
    if (column) sheet.getRange(rowNumber, column).setValue(safeSheetValue(updates[header]));
  });
}

function appendReportLog(reference, status, report) {
  const sheet = getOrCreateSheet(REPORT_LOG_SHEET);
  const headers = ['Timestamp', 'Reference', 'Status', 'Google Drive Folder', 'Google Document', 'PDF Report', 'Email Status', 'Error'];
  ensureSheetHeaders(sheet, headers);
  sheet.appendRow([
    new Date(), reference, status, report.folderUrl || '', report.documentUrl || '', report.pdfUrl || '',
    report.emailedTo || 'Not sent', report.error || ''
  ]);
  formatSubmissionRow(sheet, sheet.getLastRow(), headers.length);
}

function setupPADFSGWorkbook() {
  const submissions = getOrCreateSheet(SUBMISSIONS_SHEET);
  const submissionHeaders = [
    'Reference', 'Submission ID', 'Received At', 'Respondent Name', 'Respondent Email', 'Respondent Role',
    'Response Date', 'Answer Count', 'Proposal Scope', 'Scope Score', 'Pricing Factors', 'Report Status',
    'Google Drive Folder', 'Google Document', 'PDF Report', 'Email Status', 'Report Error'
  ];
  ensureSheetHeaders(submissions, submissionHeaders);
  ensureSheetHeaders(getOrCreateSheet(ANSWERS_SHEET), ['Reference', 'Received At', 'Respondent Name', 'Section', 'Question', 'Answer']);
  ensureSheetHeaders(getOrCreateSheet(REPORT_LOG_SHEET), ['Timestamp', 'Reference', 'Status', 'Google Drive Folder', 'Google Document', 'PDF Report', 'Email Status', 'Error']);
  setupDashboard();
  styleWorkbook();
}

function getOrCreateSheet(name) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  return spreadsheet.getSheetByName(name) || spreadsheet.insertSheet(name);
}

function ensureSheetHeaders(sheet, headers) {
  if (!sheet.getLastColumn()) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else {
    const current = getHeaders(sheet);
    const missing = headers.filter(header => !current.includes(header));
    if (missing.length) sheet.getRange(1, current.length + 1, 1, missing.length).setValues([missing]);
  }
  formatHeader(sheet, sheet.getLastColumn());
}

function setupDashboard() {
  const sheet = getOrCreateSheet(DASHBOARD_SHEET);
  if (sheet.getLastRow()) return;
  sheet.getRange('A1:D1').merge().setValue('PADFSG Questionnaire Dashboard');
  sheet.getRange('A3:A7').setValues([
    ['Total submissions'], ['Completed reports'], ['Failed reports'], ['Pending reports'], ['Last submission']
  ]);
  sheet.getRange('B3').setFormula(`=MAX(COUNTA('${SUBMISSIONS_SHEET}'!A:A)-1,0)`);
  sheet.getRange('B4').setFormula(`=COUNTIF('${SUBMISSIONS_SHEET}'!L:L,"Complete")`);
  sheet.getRange('B5').setFormula(`=COUNTIF('${SUBMISSIONS_SHEET}'!L:L,"Failed")`);
  sheet.getRange('B6').setFormula(`=COUNTIF('${SUBMISSIONS_SHEET}'!L:L,"Pending")`);
  sheet.getRange('B7').setFormula(`=IFERROR(MAX('${SUBMISSIONS_SHEET}'!C:C),"")`).setNumberFormat('dd mmm yyyy, hh:mm');
  sheet.getRange('A9:D9').merge().setValue('Where to find everything');
  sheet.getRange('A10:B13').setValues([
    ['Submissions', 'One clean management row per response, including report and email status.'],
    ['Answers', 'One row per question, including “No answer”, suitable for filtering and analysis.'],
    ['Report Log', 'A history of report creation, links, email delivery, and errors.'],
    ['Reports Drive folder', getOrCreateFolder(REPORTS_FOLDER_NAME).getUrl()]
  ]);
}

function styleWorkbook() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const dashboard = spreadsheet.getSheetByName(DASHBOARD_SHEET);
  if (dashboard) {
    dashboard.setTabColor('#9B5A36');
    dashboard.setColumnWidth(1, 180).setColumnWidth(2, 420);
    dashboard.getRange('A1:D1').setBackground('#173C31').setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(16).setHorizontalAlignment('center');
    dashboard.getRange('A3:A7').setBackground('#EAF1EC').setFontWeight('bold').setFontColor('#173C31');
    dashboard.getRange('B3:B7').setBackground('#F8FAF8').setFontWeight('bold');
    dashboard.getRange('A9:D9').setBackground('#315C46').setFontColor('#FFFFFF').setFontWeight('bold');
    dashboard.getRange('A10:B13').setWrap(true).setVerticalAlignment('top');
    dashboard.setFrozenRows(1);
  }
  const submissions = spreadsheet.getSheetByName(SUBMISSIONS_SHEET);
  if (submissions) {
    submissions.setTabColor('#315C46');
    submissions.setFrozenRows(1);
    if (!submissions.getFilter()) submissions.getRange(1, 1, Math.max(submissions.getLastRow(), 1), submissions.getLastColumn()).createFilter();
    [1, 2, 4, 5, 9, 11, 13, 14, 15, 16, 17].forEach(column => submissions.setColumnWidth(column, column >= 13 ? 190 : 150));
    submissions.getRange('C2:C').setNumberFormat('dd mmm yyyy, hh:mm');
    const statusRange = submissions.getRange('L2:L');
    submissions.setConditionalFormatRules([
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Complete').setBackground('#DCEFE2').setFontColor('#1F5B36').setRanges([statusRange]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Failed').setBackground('#FCE2E2').setFontColor('#9C1C1C').setRanges([statusRange]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Pending').setBackground('#FFF1CC').setFontColor('#76520A').setRanges([statusRange]).build()
    ]);
  }
  const answers = spreadsheet.getSheetByName(ANSWERS_SHEET);
  if (answers) {
    answers.setTabColor('#5E806C');
    answers.setFrozenRows(1);
    answers.setColumnWidth(1, 170).setColumnWidth(2, 145).setColumnWidth(3, 170).setColumnWidth(4, 220).setColumnWidth(5, 360).setColumnWidth(6, 320);
    answers.getRange('B2:B').setNumberFormat('dd mmm yyyy, hh:mm');
  }
  const log = spreadsheet.getSheetByName(REPORT_LOG_SHEET);
  if (log) {
    log.setTabColor('#C89242');
    log.setFrozenRows(1);
    log.setColumnWidth(1, 145).setColumnWidth(2, 170).setColumnWidth(3, 100);
    for (let column = 4; column <= 8; column += 1) log.setColumnWidth(column, 220);
    log.getRange('A2:A').setNumberFormat('dd mmm yyyy, hh:mm');
  }
}

function formatSubmissionRow(sheet, rowNumber, columnCount) {
  sheet.getRange(rowNumber, 1, 1, columnCount).setVerticalAlignment('top').setWrap(true);
  if (rowNumber % 2 === 1) sheet.getRange(rowNumber, 1, 1, columnCount).setBackground('#F7FAF8');
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
  appendInsight(body, 'Content readiness', analysis.readiness);

  const proposal = buildSiteProposal(data);
  appendReportSection(body, 'Automatic website proposal');
  appendCallout(body, proposal.recommendation);
  appendInsight(body, 'Primary audiences and journeys', proposal.audienceJourneys);
  appendInsight(body, 'Recommended design language', proposal.designLanguage);
  appendInsight(body, 'Recommended implementation approach', proposal.implementation);

  appendReportSection(body, 'Proposed site structure');
  appendSiteStructure(body, proposal.siteStructure);

  appendReportSection(body, 'Proposed features and services');
  appendDetailedList(body, proposal.features);

  appendReportSection(body, 'Content and asset requirements');
  appendDetailedList(body, proposal.contentRequirements);

  appendReportSection(body, 'Suggested delivery plan');
  appendDeliveryPlan(body, proposal.deliveryPlan);

  appendReportSection(body, 'Assumptions and items to confirm');
  appendDetailedList(body, proposal.openItems);

  appendReportSection(body, 'Recommended next steps');
  analysis.nextSteps.forEach(step => body.appendListItem(step).setGlyphType(DocumentApp.GlyphType.NUMBER));

  body.appendPageBreak();
  appendReportSection(body, 'Appendix: complete questions and answers');
  const readableAnswers = parseReadableAnswers(data._reportAnswers);
  appendResponseCoverage(body, readableAnswers);
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
  const recipients = getReportRecipients(data);
  shareReportFiles([docFile, pdfFile], recipients);
  const emailedTo = emailSubmissionReport(data, reference, pdfBlob, document.getUrl(), pdfFile.getUrl());

  return { folderUrl: folder.getUrl(), documentUrl: document.getUrl(), pdfUrl: pdfFile.getUrl(), emailedTo };
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
    ['Answers recorded', String(data._answerCount || parseReadableAnswers(data._reportAnswers).length || 'Not available')],
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

function appendSiteStructure(body, groups) {
  groups.forEach(group => {
    const heading = body.appendParagraph(group.section).setHeading(DocumentApp.ParagraphHeading.HEADING2);
    heading.editAsText().setForegroundColor('#2F6B49').setBold(true);
    const rows = [['Page or area', 'Purpose and recommended content']]
      .concat(group.pages.map(page => [page.name, page.purpose]));
    const table = body.appendTable(rows);
    table.getRow(0).getCell(0).setBackgroundColor('#315C46');
    table.getRow(0).getCell(1).setBackgroundColor('#315C46');
    table.getRow(0).getCell(0).editAsText().setForegroundColor('#FFFFFF').setBold(true);
    table.getRow(0).getCell(1).editAsText().setForegroundColor('#FFFFFF').setBold(true);
    for (let row = 1; row < rows.length; row += 1) {
      table.getCell(row, 0).setBackgroundColor(row % 2 ? '#EAF1EC' : '#F7FAF8');
      table.getCell(row, 0).editAsText().setBold(true).setForegroundColor('#173C31');
      table.getCell(row, 1).editAsText().setForegroundColor('#26352E');
    }
    body.appendParagraph('');
  });
}

function appendDetailedList(body, items) {
  items.forEach(item => {
    const paragraph = body.appendListItem(item.title).setGlyphType(DocumentApp.GlyphType.BULLET);
    paragraph.editAsText().setBold(true).setForegroundColor('#173C31');
    const detail = body.appendParagraph(item.detail);
    detail.setIndentStart(24).setSpacingAfter(6);
    detail.editAsText().setFontSize(9).setForegroundColor('#5A675F');
  });
}

function appendDeliveryPlan(body, stages) {
  const rows = [['Stage', 'Purpose', 'Main outputs']].concat(stages.map(stage => [stage.name, stage.purpose, stage.outputs]));
  const table = body.appendTable(rows);
  for (let column = 0; column < 3; column += 1) {
    table.getRow(0).getCell(column).setBackgroundColor('#315C46');
    table.getRow(0).getCell(column).editAsText().setForegroundColor('#FFFFFF').setBold(true);
  }
  for (let row = 1; row < rows.length; row += 1) {
    if (row % 2) for (let column = 0; column < 3; column += 1) table.getCell(row, column).setBackgroundColor('#F7FAF8');
    table.getCell(row, 0).editAsText().setBold(true).setForegroundColor('#173C31');
  }
}

function appendResponseCoverage(body, answers) {
  if (!answers.length) return;
  const unanswered = answers.filter(item => String(item.answer).trim().toLowerCase() === 'no answer');
  const answered = answers.length - unanswered.length;
  appendCallout(body, `${answers.length} questions recorded: ${answered} answered and ${unanswered.length} marked “No answer”. Every visible question is included below.`);
}

function buildSiteProposal(data) {
  const selectedAreas = values(data, 'workAreas');
  const selectedFeatures = values(data, 'platformFeatures');
  const selectedHomepage = values(data, 'homepageContent');
  const languages = [data.mainLanguage || 'english'].concat(values(data, 'additionalLanguages').filter(item => item !== 'none'));
  const scope = data._proposalScope || 'Website scope to be confirmed';

  const pageGroups = {
    about: { section: 'Organisation', pages: [
      { name: 'About PADFSG', purpose: 'Mission, African mandate, history, objectives and organisational profile.' },
      { name: 'Strategy and impact', purpose: 'Strategic priorities, progress indicators, achievements and evidence of impact.' }
    ]},
    governance: { section: 'Leadership and governance', pages: [
      { name: 'Leadership', purpose: 'President, board and committee profiles with roles, biographies and approved photographs.' },
      { name: 'Governance', purpose: 'Governance structure, constitution, policies, annual reports and official documents.' }
    ]},
    regions: { section: 'African network', pages: [
      { name: 'Regions and countries', purpose: 'Continental overview with regional groupings, representatives and participation.' },
      { name: 'Country or regional profiles', purpose: 'Reusable profile pages for contacts, projects, news, events and local achievements.' }
    ]},
    projects: { section: 'Programmes and impact', pages: [
      { name: 'Projects', purpose: 'Filterable overview of current and completed programmes.' },
      { name: 'Project detail template', purpose: 'Challenge, objectives, partners, locations, activities, results and related resources.' }
    ]},
    research: { section: 'Research', pages: [
      { name: 'Research and publications', purpose: 'Searchable studies, publications, calls for collaboration and research updates.' },
      { name: 'Publication detail template', purpose: 'Citation, authors, abstract, link or download, topic and related material.' }
    ]},
    resources: { section: 'Clinical knowledge', pages: [
      { name: 'Clinical resource library', purpose: 'Search and filter guidance by topic, audience, language, type and publication year.' },
      { name: 'Resource detail template', purpose: 'Plain summary, file or external link, authorship, date, language and related resources.' }
    ]},
    education: { section: 'Education', pages: [
      { name: 'Education and webinars', purpose: 'Upcoming learning opportunities, recordings, courses and professional development.' },
      { name: 'Learning detail template', purpose: 'Learning objectives, speakers, audience, date, registration, materials and certificate details.' }
    ]},
    events: { section: 'Events', pages: [
      { name: 'Events and conferences', purpose: 'Calendar and list views for upcoming and past events.' },
      { name: 'Event detail template', purpose: 'Programme, venue or joining link, speakers, registration, fees and related downloads.' }
    ]},
    membership: { section: 'Membership', pages: [
      { name: 'Membership overview', purpose: 'Categories, eligibility, benefits, fees and clear application journey.' },
      { name: 'Join or renew', purpose: 'Application, supporting information, payment or renewal steps according to the selected scope.' },
      { name: 'Member area', purpose: 'Account, status and private resources only if selected for the first release.' }
    ]},
    partners: { section: 'Partnerships', pages: [
      { name: 'Partners and supporters', purpose: 'Partner recognition grouped by relationship or programme.' },
      { name: 'Partner with PADFSG', purpose: 'Collaboration opportunities, value proposition and enquiry route.' }
    ]},
    news: { section: 'News and media', pages: [
      { name: 'News', purpose: 'Searchable organisation, regional, project and research updates.' },
      { name: 'News article template', purpose: 'Headline, date, author, imagery, article body and related content.' },
      { name: 'Media centre', purpose: 'Press contacts, organisational facts, approved assets and media releases.' }
    ]},
    awards: { section: 'Awards', pages: [
      { name: 'Awards and recognition', purpose: 'Award programmes, eligibility, nomination dates and previous recipients.' }
    ]},
    patients: { section: 'Patient information', pages: [
      { name: 'For patients and families', purpose: 'Clear prevention, warning signs, care guidance and routes to professional help.' },
      { name: 'Patient topic template', purpose: 'Plain-language explanation, actions, cautions, illustrations and downloadable guidance.' }
    ]}
  };

  const siteStructure = [{ section: 'Essential pages', pages: [
    { name: 'Home', purpose: `Communicate the mission, priority action and selected homepage evidence: ${friendlyList(selectedHomepage, {})}.` },
    { name: 'Contact', purpose: 'Contact routes, general enquiries, partnership enquiries and appropriate organisational details.' },
    { name: 'Search', purpose: 'A single route to find pages, resources, news, events and publications.' },
    { name: 'Privacy, cookies and website terms', purpose: 'Required privacy information, consent choices and website-use terms.' }
  ]}];
  selectedAreas.forEach(area => { if (pageGroups[area]) siteStructure.push(pageGroups[area]); });

  const featureDetails = {
    'site-search': ['Website search', 'Search across key content types with useful result labels and filters where appropriate.'],
    'member-accounts': ['Member accounts', 'Secure sign-in, profile and access rules; confirm identity, permissions and support process.'],
    forms: ['Online applications and requests', 'Structured forms with confirmations, notifications, consent and protected data handling.'],
    payments: ['Online payments', 'Payment journey for fees or donations, with receipts, status tracking and finance hand-off.'],
    newsletter: ['Newsletter subscription', 'Consent-based signup connected to the selected mailing platform.'],
    events: ['Event registration', 'Registration, confirmations, reminders and attendance information.'],
    certificates: ['Certificates', 'Generate or deliver attendance and training certificates against verified participation.'],
    directory: ['Searchable directory', 'Structured profiles and filters with clear decisions about public and private information.'],
    'interactive-map': ['Interactive African map', 'Explore participation, representatives, projects or activity by country.'],
    'resource-tracking': ['Resource-use reporting', 'Measure views and downloads to identify useful clinical content.'],
    analytics: ['Website analytics', 'Privacy-conscious reporting for audiences, content, referrals, events and conversions.'],
    translation: ['Multilingual publishing', `Support ${friendlyList(languages, {})} with a repeatable translation and review workflow.`]
  };
  const features = selectedFeatures.length
    ? selectedFeatures.map(feature => ({ title: (featureDetails[feature] || [formatLabel(feature)])[0], detail: (featureDetails[feature] || ['', 'Confirm detailed behaviour during discovery.'])[1] }))
    : [{ title: 'Core informational website', detail: 'No additional online service was selected; include reliable navigation, search-engine foundations, forms and analytics as agreed.' }];

  const contentRequirements = [
    { title: 'Organisation foundation', detail: 'Approved mission, history, objectives, legal name, contact details, leadership information and governance documents.' },
    { title: 'Homepage evidence', detail: `Prepare the selected homepage material: ${friendlyList(selectedHomepage, {})}. Confirm the primary action: ${data.primaryAction || 'No answer'}.` },
    { title: 'Structured content inventory', detail: `Create one inventory for the proposed sections: ${friendlyList(selectedAreas, {})}. Record owner, status, language, review date and source for each item.` },
    { title: 'Photography and media', detail: `Prioritise ${friendlyList(values(data, 'visualSubjects'), {})}. Confirm consent, ownership, captions, credits and suitable image quality.` },
    { title: 'Language preparation', detail: `Main and additional language plan: ${friendlyList(languages, {})}. Agree source language, translation responsibility and clinical review.` },
    { title: 'Clinical and patient review', detail: 'Apply named subject review, publication dates, version control and review dates to clinical or patient-facing guidance.' }
  ];

  const deliveryPlan = [
    { name: '1. Confirm', purpose: 'Turn questionnaire findings into an agreed scope.', outputs: 'Scope workshop, priorities, sitemap, feature list, assumptions and quotation.' },
    { name: '2. Organise', purpose: 'Define how information will be structured and governed.', outputs: 'Content model, page templates, navigation, content inventory and responsibilities.' },
    { name: '3. Prototype', purpose: 'Test the inferred experience before full production.', outputs: 'Homepage and key journey prototypes, stakeholder review and agreed design system.' },
    { name: '4. Build', purpose: 'Create the approved website and selected services.', outputs: 'Responsive templates, content types, integrations, forms, search and analytics.' },
    { name: '5. Populate and verify', purpose: 'Add content and check accuracy and usability.', outputs: 'Content entry, language review, clinical review, accessibility checks and device testing.' },
    { name: '6. Launch and improve', purpose: 'Release safely and establish ongoing management.', outputs: 'Launch checks, training, documentation, measurement dashboard and improvement backlog.' }
  ];

  const allAnswers = parseReadableAnswers(data._reportAnswers);
  const unanswered = allAnswers.filter(item => String(item.answer).trim().toLowerCase() === 'no answer');
  const openItems = [
    { title: 'Scope confirmation', detail: `${scope} with a scope score of ${data._scopeScore || 'not available'}. Validate what belongs in the first release and what should be phased.` },
    { title: 'Commercial assumptions', detail: `Pricing factors currently identified: ${data._pricingFactors || 'No major factors recorded'}. Confirm external service fees, content volume, integrations and ongoing support.` },
    { title: 'Content readiness', detail: buildReportAnalysis(data).readiness },
    { title: 'Unanswered decisions', detail: unanswered.length ? `${unanswered.length} visible questions were not answered. Review them in the appendix before final scope approval.` : 'All visible questions received an answer.' },
    { title: 'Privacy and ownership', detail: 'Confirm who owns the website, accounts, domain, analytics, mailing data, member data, source files and generated content.' }
  ];

  const impressions = friendlyList(values(data, 'desiredImpression'), {});
  const proposedPageCount = siteStructure.reduce((total, group) => total + group.pages.length, 0);
  const recommendation = `${scope}. Build a focused PADFSG platform around ${selectedAreas.length || 'the confirmed'} selected content areas and ${selectedFeatures.length} selected online services. This first-draft structure proposes approximately ${proposedPageCount} core pages or reusable page types. Validate and phase it in a scope workshop before quotation approval.`;
  const audienceJourneys = `Prioritise ${friendlyList(values(data, 'primaryAudiences'), {})}. Support the selected visitor priorities: ${friendlyList(values(data, 'visitorPriorities'), {})}. Each audience should reach its most important information or action from the homepage and primary navigation without needing organisational knowledge.`;
  const designLanguage = `Translate the selected impressions - ${impressions} - into a credible African health organisation system. Use the chosen experience reference (${formatLabel(data.experienceMetaphor || 'balanced professional organisation')}) and information density (${formatLabel(data.informationComfort || 'balanced')}) to guide hierarchy, imagery, spacing and editorial tone.`;
  const implementation = `Use reusable page templates and structured content so PADFSG can grow without redesigning each section. Responsive behaviour is a baseline requirement. Include secure hosting, backups, performance optimisation, accessibility, privacy, search foundations and measurement in the implementation plan.`;

  return { recommendation, audienceJourneys, designLanguage, implementation, siteStructure, features, contentRequirements, deliveryPlan, openItems };
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
    'Prepare a content checklist with an owner and deadline for every page.',
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
  const recipients = getReportRecipients(data);
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

function getReportRecipients(data) {
  const recipients = [];
  const configured = String(REPORT_RECIPIENT_EMAIL || '').split(',').map(email => email.trim()).filter(Boolean);
  configured.forEach(email => { if (!recipients.includes(email)) recipients.push(email); });
  if (!recipients.length) {
    try {
      const spreadsheetFile = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId());
      const ownerEmail = spreadsheetFile.getOwner().getEmail();
      if (ownerEmail) recipients.push(ownerEmail);
    } catch (error) {
      const effectiveEmail = Session.getEffectiveUser().getEmail();
      if (effectiveEmail) recipients.push(effectiveEmail);
    }
  }
  if (SEND_COPY_TO_RESPONDENT && data.respEmail && !recipients.includes(data.respEmail)) recipients.push(data.respEmail);
  return recipients;
}

function shareReportFiles(files, recipients) {
  recipients.forEach(email => {
    files.forEach(file => {
      try {
        file.addViewer(email);
      } catch (error) {
        console.warn(`Could not share ${file.getName()} with ${email}: ${error.message}`);
      }
    });
  });
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
