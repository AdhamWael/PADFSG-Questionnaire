const CONFIG = window.PADFSG_FORM_CONFIG || {};
const FORMSPREE_URL = CONFIG.formspreeUrl || '';
const GOOGLE_SHEET_URL = CONFIG.googleSheetsUrl || '';
const form = document.getElementById('discoveryForm');

let visibleSteps = [];
let currentVisibleIndex = 0;

const allSteps = Array.from(document.querySelectorAll('.wizard-step'));
const btnPrev = document.getElementById('btnPrev');
const btnNext = document.getElementById('btnNext');
const btnSubmit = document.getElementById('btnSubmit');
const progressFill = document.getElementById('progressFill');
const stepNavEl = document.getElementById('stepNav');
const successScreen = document.getElementById('successScreen');
const questionnaire = document.getElementById('questionnaire');
const submissionStatus = document.getElementById('submissionStatus');

const isConfigured = (url, placeholder) => Boolean(url && !url.includes(placeholder));

function enforceSelectionLimit(group) {
  const step = group.closest('.wizard-step');
  if (step && !visibleSteps.includes(step)) return;
  const limit = Number(group.dataset.maxSelections);
  const boxes = Array.from(group.querySelectorAll('input[type="checkbox"]'));
  const checked = boxes.filter(box => box.checked).length;
  boxes.forEach(box => { box.disabled = !box.checked && checked >= limit; });
  group.classList.toggle('selection-limit-reached', checked >= limit);
}

function selectedWorkAreas() {
  return new Set(Array.from(document.querySelectorAll('input[name="workAreas"]:checked')).map(input => input.value));
}

function shouldShow(step, selected) {
  const condition = step.dataset.condition;
  return !condition || condition.split(',').some(value => selected.has(value.trim()));
}

function scrollToQuestion() {
  const target = visibleSteps[currentVisibleIndex] || questionnaire;
  if (!target) return;
  const header = document.querySelector('.site-header');
  const progress = document.querySelector('.progress-track');
  const offset = (header?.offsetHeight || 0) + (progress?.offsetHeight || 0) + 20;
  window.scrollTo({ top: Math.max(target.getBoundingClientRect().top + window.scrollY - offset, 0), behavior: 'smooth' });
}

function rebuildFlow(keepCurrent = true) {
  const currentElement = visibleSteps[currentVisibleIndex];
  visibleSteps = allSteps.filter(step => shouldShow(step, selectedWorkAreas()));
  allSteps.forEach(step => {
    const excluded = !visibleSteps.includes(step);
    step.querySelectorAll('input, textarea, select').forEach(control => { control.disabled = excluded; });
  });
  document.querySelectorAll('[data-max-selections]').forEach(enforceSelectionLimit);
  currentVisibleIndex = keepCurrent && visibleSteps.includes(currentElement)
    ? visibleSteps.indexOf(currentElement)
    : Math.min(currentVisibleIndex, Math.max(visibleSteps.length - 1, 0));
  buildDots();
  showCurrentStep(false);
}

function buildDots() {
  stepNavEl.replaceChildren();
  visibleSteps.forEach((step, index) => {
    const dot = document.createElement('button');
    dot.type = 'button';
    dot.className = 'step-dot';
    dot.title = `Question ${index + 1}`;
    dot.setAttribute('aria-label', `Go to question ${index + 1}`);
    dot.addEventListener('click', () => { currentVisibleIndex = index; showCurrentStep(true); });
    stepNavEl.appendChild(dot);
  });
}

function showCurrentStep(scroll = true) {
  allSteps.forEach(step => { step.classList.remove('active'); step.setAttribute('aria-hidden', 'true'); });
  const current = visibleSteps[currentVisibleIndex];
  if (!current) return;
  current.classList.add('active');
  current.setAttribute('aria-hidden', 'false');
  visibleSteps.forEach((step, index) => {
    const eye = step.querySelector('.step-eyebrow');
    if (eye) {
      eye.dataset.label ||= eye.textContent.trim().replace(/^Question\s+\d+$/i, 'Discovery');
      eye.textContent = `${eye.dataset.label} · ${index + 1} of ${visibleSteps.length}`;
    }
  });
  document.querySelectorAll('.step-dot').forEach((dot, index) => {
    dot.classList.toggle('active', index === currentVisibleIndex);
    dot.classList.toggle('done', index < currentVisibleIndex);
    dot.setAttribute('aria-current', index === currentVisibleIndex ? 'step' : 'false');
  });
  progressFill.style.width = `${((currentVisibleIndex + 1) / visibleSteps.length) * 100}%`;
  btnPrev.style.visibility = currentVisibleIndex === 0 ? 'hidden' : 'visible';
  btnNext.classList.toggle('hidden', currentVisibleIndex === visibleSteps.length - 1);
  btnSubmit.classList.toggle('hidden', currentVisibleIndex !== visibleSteps.length - 1);
  if (scroll) scrollToQuestion();
}

btnNext.addEventListener('click', () => {
  if (currentVisibleIndex < visibleSteps.length - 1) { currentVisibleIndex += 1; showCurrentStep(true); }
});
btnPrev.addEventListener('click', () => {
  if (currentVisibleIndex > 0) { currentVisibleIndex -= 1; showCurrentStep(true); }
});

document.querySelectorAll('input[name="workAreas"]').forEach(input => input.addEventListener('change', () => rebuildFlow(true)));

document.querySelectorAll('[data-max-selections]').forEach(group => {
  group.querySelectorAll('input[type="checkbox"]').forEach(box => box.addEventListener('change', () => enforceSelectionLimit(group)));
  enforceSelectionLimit(group);
});

document.querySelectorAll('.null-choice input').forEach(input => {
  input.addEventListener('change', () => {
    if (!input.checked) return;
    document.querySelectorAll(`input[name="${CSS.escape(input.name)}"]`).forEach(peer => { if (peer !== input) peer.checked = false; });
  });
});

form.addEventListener('change', event => {
  const input = event.target;
  if (!(input instanceof HTMLInputElement) || !input.checked || input.closest('.null-choice')) return;
  document.querySelectorAll(`.null-choice input[name="${CSS.escape(input.name)}"]`).forEach(none => { none.checked = false; });
});

function collectFormData() {
  const data = {};
  for (const [key, value] of new FormData(form).entries()) {
    if (key === 'website') continue;
    data[key] = Object.hasOwn(data, key) ? `${data[key]}, ${value}` : value;
  }
  data._submissionId = crypto.randomUUID ? crypto.randomUUID() : `${Date.now()}-${Math.random().toString(36).slice(2)}`;
  data._submittedAt = new Date().toISOString();
  data._form = 'PADFSG Website Discovery Questionnaire';
  data.email = data.respEmail || '';
  data.name = data.respName || '';
  data._subject = `New PADFSG website questionnaire from ${data.respName || 'respondent'}`;
  data.respDate ||= new Date().toLocaleDateString('en-GB');
  Object.assign(data, buildScopeProfile(data));
  const readableReport = buildReadableAnswerReport();
  data._answerCount = readableReport.length;
  data._reportAnswers = JSON.stringify(readableReport);
  return data;
}

function readableOptionLabel(input) {
  const label = input.closest('label');
  if (!label) return input.value;
  const title = label.querySelector('strong');
  if (title) return title.textContent.trim();
  const textContainer = label.querySelector('span');
  return (textContainer?.textContent || label.textContent || input.value).replace(/\s+/g, ' ').trim();
}

function questionForControl(control) {
  const group = control.closest('.field-group');
  const groupLabel = group?.querySelector(':scope > .field-label');
  if (groupLabel) return groupLabel.textContent.replace(/\s+/g, ' ').trim();
  const directLabel = control.closest('label');
  const directText = directLabel?.querySelector(':scope > span');
  return (directText?.textContent || formatFieldName(control.name)).replace(/\s+/g, ' ').trim();
}

function formatFieldName(name) {
  return String(name || '').replace(/([a-z])([A-Z])/g, '$1 $2').replace(/[_-]+/g, ' ').replace(/\b\w/g, letter => letter.toUpperCase());
}

function buildReadableAnswerReport() {
  const excluded = new Set(['website', 'respName', 'respRole', 'respEmail', 'respDate']);
  const report = [];
  visibleSteps.forEach(step => {
    const section = step.querySelector('.step-title')?.textContent.trim() || 'Website discovery';
    const controls = Array.from(step.querySelectorAll('input[name], textarea[name], select[name]'));
    const names = [...new Set(controls.map(control => control.name).filter(name => !excluded.has(name)))];
    names.forEach(name => {
      const matching = controls.filter(control => control.name === name && !control.disabled);
      if (!matching.length) return;
      const type = matching[0].type;
      let answers = [];
      if (type === 'checkbox' || type === 'radio') {
        answers = matching.filter(control => control.checked).map(readableOptionLabel);
      } else {
        answers = matching.map(control => control.value.trim()).filter(Boolean);
      }
      if (answers.length) report.push({ section, question: questionForControl(matching[0]), answer: answers.join(', ') });
    });
  });
  return report;
}

function completeAnswerText(report) {
  let currentSection = '';
  const lines = [];
  report.forEach(item => {
    if (item.section !== currentSection) {
      currentSection = item.section;
      lines.push('', currentSection.toUpperCase());
    }
    lines.push(`${item.question}\nAnswer: ${item.answer}`);
  });
  return lines.join('\n\n').trim();
}

function selectedValues(data, key) {
  return String(data[key] || '').split(',').map(value => value.trim()).filter(Boolean);
}

function buildScopeProfile(data) {
  let score = selectedValues(data, 'workAreas').length;
  const factors = [];
  const featureWeights = {
    'site-search': 1, 'member-accounts': 5, forms: 2, payments: 4,
    newsletter: 1, events: 3, certificates: 3, directory: 4,
    'interactive-map': 4, 'resource-tracking': 2, analytics: 2, translation: 4
  };
  selectedValues(data, 'platformFeatures').forEach(feature => { score += featureWeights[feature] || 1; });

  const languages = selectedValues(data, 'additionalLanguages').filter(value => value !== 'none');
  score += languages.length * 2;
  if (languages.length) factors.push(`${languages.length} additional language(s)`);
  if (data.patientContentApproach === 'equal-audience') { score += 2; factors.push('equal professional and public audiences'); }
  if (data.launchApproach === 'complete') { score += 4; factors.push('complete first release'); }
  if (data.pageRange === 'medium') score += 3;
  if (data.pageRange === 'large') { score += 7; factors.push('more than 25 pages'); }
  if (data.contentPreparation === 'shared') { score += 3; factors.push('shared content preparation'); }
  if (data.contentPreparation === 'full-support') { score += 7; factors.push('full content support'); }
  if (data.approvalLevel === 'committee' || data.decisionModel === 'committee') { score += 2; factors.push('committee approvals'); }
  if (data.editorModel === 'regional-editors') { score += 3; factors.push('several publishing teams'); }
  if (data.editorModel === 'managed-service') { score += 3; factors.push('ongoing content support'); }

  const profile = score <= 12 ? 'Focused website' : score <= 25 ? 'Standard organisation website' : score <= 42 ? 'Advanced digital platform' : 'Full service platform';
  return {
    _proposalScope: profile,
    _scopeScore: score,
    _pricingFactors: factors.length ? factors.join('; ') : 'No major complexity factors selected'
  };
}

async function sendToFormspree(data) {
  if (!isConfigured(FORMSPREE_URL, 'YOUR_FORM_ID')) return { service: 'Formspree', configured: false };
  const report = JSON.parse(data._reportAnswers || '[]');
  const formspreeData = {
    name: data.respName || '',
    email: data.respEmail || '',
    respondent_role: data.respRole || '',
    response_date: data.respDate || '',
    proposal_scope: data._proposalScope || '',
    scope_score: data._scopeScore || '',
    answer_count: data._answerCount || report.length,
    pricing_factors: data._pricingFactors || '',
    complete_answers: completeAnswerText(report),
    _subject: data._subject
  };
  report.forEach((item, index) => {
    const number = String(index + 1).padStart(2, '0');
    const shortQuestion = item.question.replace(/\s+/g, ' ').trim().slice(0, 90);
    formspreeData[`Q${number} - ${shortQuestion}`] = item.answer;
  });
  const response = await fetch(FORMSPREE_URL, {
    method: 'POST', headers: { 'Content-Type': 'application/json', Accept: 'application/json' }, body: JSON.stringify(formspreeData)
  });
  const payload = await response.json().catch(() => ({}));
  if (!response.ok) throw new Error(payload?.errors?.map(error => error.message).join(' ') || 'Formspree rejected the submission.');
  return { service: 'Formspree', configured: true };
}

async function sendToGoogleSheets(data) {
  if (!isConfigured(GOOGLE_SHEET_URL, 'YOUR_SCRIPT_ID')) return { service: 'Google Sheets', configured: false };
  // text/plain avoids a CORS preflight that Google Apps Script web apps do not answer reliably.
  await fetch(GOOGLE_SHEET_URL, { method: 'POST', mode: 'no-cors', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify(data) });
  return { service: 'Google Sheets', configured: true };
}

function setStatus(message, type = '') {
  submissionStatus.textContent = message;
  submissionStatus.className = `submission-status${type ? ` is-${type}` : ''}`;
}

form.addEventListener('submit', async event => {
  event.preventDefault();
  if (form.website.value) return;
  if (!form.reportValidity()) { setStatus('Please complete the highlighted respondent details.', 'error'); return; }

  btnSubmit.textContent = 'Submitting…';
  btnSubmit.disabled = true;
  setStatus('Securely sending your response…');
  const data = collectFormData();
  const tasks = [sendToFormspree(data), sendToGoogleSheets(data)];
  const results = await Promise.allSettled(tasks);
  const successful = results.filter(result => result.status === 'fulfilled' && result.value.configured).map(result => result.value.service);
  const configuredCount = [isConfigured(FORMSPREE_URL, 'YOUR_FORM_ID'), isConfigured(GOOGLE_SHEET_URL, 'YOUR_SCRIPT_ID')].filter(Boolean).length;

  if (!configuredCount) {
    setStatus('Connect at least one submission service in config.js before publishing.', 'error');
  } else if (!successful.length) {
    const reason = results.find(result => result.status === 'rejected')?.reason?.message;
    setStatus(reason || 'The response could not be sent. Please try again.', 'error');
  } else {
    const partial = successful.length < configuredCount;
    const reference = `PADFSG-${data._submissionId.split('-')[0].toUpperCase()}`;
    document.getElementById('successRef').textContent = partial
      ? `Reference: ${reference}. Your response was saved; one backup service needs attention.`
      : `Reference: ${reference}`;
    successScreen.classList.remove('hidden');
    successScreen.querySelector('.success-inner').setAttribute('tabindex', '-1');
    successScreen.querySelector('.success-inner').focus();
    setStatus('Response submitted.', 'success');
    form.reset();
  }

  btnSubmit.textContent = 'Submit assessment →';
  btnSubmit.disabled = false;
});

rebuildFlow(false);
document.getElementById('heroStart')?.addEventListener('click', () => {
  scrollToQuestion();
  window.setTimeout(() => visibleSteps[0]?.querySelector('input, textarea, button')?.focus({ preventScroll: true }), 650);
});
