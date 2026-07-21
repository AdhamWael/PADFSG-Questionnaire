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
    if (eye) eye.textContent = `Question ${index + 1} of ${visibleSteps.length}`;
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
  data._form = 'PADFSG Website Planning Questionnaire';
  data.email = data.respEmail || '';
  data.name = data.respName || '';
  data._subject = `New PADFSG website questionnaire from ${data.respName || 'respondent'}`;
  data.respDate ||= new Date().toLocaleDateString('en-GB');
  return data;
}

async function sendToFormspree(data) {
  if (!isConfigured(FORMSPREE_URL, 'YOUR_FORM_ID')) return { service: 'Formspree', configured: false };
  const response = await fetch(FORMSPREE_URL, {
    method: 'POST', headers: { 'Content-Type': 'application/json', Accept: 'application/json' }, body: JSON.stringify(data)
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
