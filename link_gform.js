const surveyForm = document.getElementById('survey-form');
const surveyAction = document.getElementById('survey-action');
const submitButton = document.getElementById('submit');
const submitStatus = document.getElementById('submit-status');
const loadLatestButton = document.getElementById('load-latest-response');
const loadedVersionInfo = document.getElementById('loaded-version-info');

let activeSurveyConfig = {};
let surveyRequestTimeoutId;
let pendingSurveyAction = '';
let hasSubmittedVersion = false;

function isAppsScriptMessage(event) {
  try {
    const hostname = new URL(event.origin).hostname;
    return event.origin === 'https://script.google.com' ||
      hostname.endsWith('.googleusercontent.com');
  } catch (error) {
    return false;
  }
}

function formatSubmittedAt(value) {
  const date = new Date(value);
  if (!value || Number.isNaN(date.getTime())) {
    return '';
  }
  return date.toLocaleString('ko-KR');
}

function updateCharacterCounts() {
  surveyForm.querySelectorAll('.answer-box').forEach(function (answerBox) {
    const textarea = answerBox.querySelector('textarea');
    const counter = answerBox.querySelector('.character-count');
    const maxLength = textarea.maxLength > 0 ? textarea.maxLength : 1000;
    counter.textContent = `${textarea.value.length} / ${maxLength}자`;
  });
}

function updateCopyFlags(copyFlags) {
  for (let index = 0; index < 5; index++) {
    const flagInput = document.getElementById(`answer-copy-flag-${index + 1}`);
    flagInput.value = Array.isArray(copyFlags) && copyFlags[index]
      ? 'true'
      : 'false';
  }
}

function resetSurveyButtons() {
  submitButton.disabled = false;
  loadLatestButton.disabled = false;
  loadLatestButton.textContent =
    activeSurveyConfig.loadLatestButtonText || '최근 답변 불러오기';
  submitButton.textContent = hasSubmittedVersion
    ? (activeSurveyConfig.resubmitButtonText || '수정 내용 저장')
    : (activeSurveyConfig.submitButtonText || submitButton.dataset.defaultText || 'Submit');
}

function resetSurveyEditor() {
  clearTimeout(surveyRequestTimeoutId);
  pendingSurveyAction = '';
  hasSubmittedVersion = false;
  surveyForm.querySelectorAll('textarea').forEach(function (textarea) {
    textarea.value = '';
  });
  if (typeof window.syncAnswerIntegrityBaselines === 'function') {
    window.syncAnswerIntegrityBaselines();
  }
  updateCopyFlags([]);
  updateCharacterCounts();
  submitStatus.textContent = '';
  submitStatus.className = 'submit-status';
  loadedVersionInfo.textContent =
    activeSurveyConfig.newResponseMessage || '새 답변을 작성 중입니다.';
  resetSurveyButtons();
}

window.resetSurveyResponseEditor = resetSurveyEditor;

async function sendSurveyRequest(action) {
  if (pendingSurveyAction) {
    return;
  }

  if (action === 'submitSurvey' && typeof window.auditAnswerIntegrity === 'function') {
    window.auditAnswerIntegrity();
  }

  let googleScriptUrl;
  try {
    const config = await window.surveyConfigPromise;
    activeSurveyConfig = config.form || {};
    googleScriptUrl = activeSurveyConfig.googleScriptUrl || '';
  } catch (error) {
    submitStatus.textContent = 'questions.json을 불러오지 못했습니다.';
    submitStatus.className = 'submit-status error';
    return;
  }

  if (!googleScriptUrl.startsWith('https://script.google.com/macros/s/') ||
      !googleScriptUrl.endsWith('/exec')) {
    submitStatus.textContent = 'questions.json에 Apps Script 배포 URL을 입력해 주세요.';
    submitStatus.className = 'submit-status error';
    return;
  }

  pendingSurveyAction = action;
  surveyAction.value = action;
  surveyForm.action = googleScriptUrl;
  submitButton.disabled = true;
  loadLatestButton.disabled = true;

  if (action === 'loadLatestSurvey') {
    loadLatestButton.textContent = '불러오는 중...';
    loadedVersionInfo.textContent =
      activeSurveyConfig.loadingLatestMessage || '최근 답변을 불러오고 있습니다.';
  } else {
    submitButton.textContent = '제출 중...';
    submitStatus.textContent =
      activeSurveyConfig.submitProgressMessage || '답변을 전송하고 있습니다.';
    submitStatus.className = 'submit-status';
  }

  HTMLFormElement.prototype.submit.call(surveyForm);

  clearTimeout(surveyRequestTimeoutId);
  surveyRequestTimeoutId = setTimeout(function () {
    const timedOutAction = pendingSurveyAction;
    pendingSurveyAction = '';
    resetSurveyButtons();
    if (timedOutAction === 'loadLatestSurvey') {
      loadedVersionInfo.textContent = '최근 답변을 불러오지 못했습니다.';
    } else {
      submitStatus.textContent =
        '제출 결과를 확인하지 못했습니다. Apps Script 배포 상태를 확인해 주세요.';
      submitStatus.className = 'submit-status error';
    }
  }, 45000);
}

surveyForm.addEventListener('submit', function (event) {
  event.preventDefault();
  sendSurveyRequest('submitSurvey');
});

loadLatestButton.addEventListener('click', function () {
  sendSurveyRequest('loadLatestSurvey');
});

window.addEventListener('message', function (event) {
  if (!isAppsScriptMessage(event) ||
      !['survey-submit-result', 'survey-latest-result'].includes(event.data?.type)) {
    return;
  }

  clearTimeout(surveyRequestTimeoutId);
  pendingSurveyAction = '';
  const result = event.data;

  if (result.type === 'survey-submit-result') {
    if (!result.success) {
      resetSurveyButtons();
      submitStatus.textContent = result.message || '답변을 제출하지 못했습니다.';
      submitStatus.className = 'submit-status error';
      return;
    }

    hasSubmittedVersion = true;
    resetSurveyButtons();
    const submittedAt = formatSubmittedAt(result.submittedAt);
    loadedVersionInfo.textContent = submittedAt
      ? `최근 저장: ${submittedAt}`
      : '방금 저장한 제출본을 편집 중입니다.';
    submitStatus.textContent = activeSurveyConfig.submitSuccessMessage ||
      '답변을 저장했습니다. 이후 저장하면 기존 제출본이 수정됩니다.';
    submitStatus.className = 'submit-status success';
    return;
  }

  resetSurveyButtons();
  if (!result.success) {
    loadedVersionInfo.textContent = result.message || '최근 답변을 불러오지 못했습니다.';
    submitStatus.textContent = result.message || '';
    submitStatus.className = 'submit-status error';
    return;
  }

  if (!result.found) {
    hasSubmittedVersion = false;
    updateCopyFlags([]);
    resetSurveyButtons();
    loadedVersionInfo.textContent = activeSurveyConfig.noPreviousResponseMessage ||
      '이전에 제출한 답변이 없습니다.';
    return;
  }

  const answers = Array.isArray(result.answers) ? result.answers : [];
  updateCopyFlags(result.copyFlags);
  surveyForm.querySelectorAll('.answer-box textarea').forEach(function (textarea, index) {
    textarea.value = answers[index] || '';
  });
  if (typeof window.syncAnswerIntegrityBaselines === 'function') {
    window.syncAnswerIntegrityBaselines();
  }
  updateCharacterCounts();
  hasSubmittedVersion = true;
  resetSurveyButtons();

  const submittedAt = formatSubmittedAt(result.submittedAt);
  const loadedMessage = activeSurveyConfig.latestLoadedMessage ||
    '현재 제출본을 불러왔습니다.';
  loadedVersionInfo.textContent = submittedAt
    ? `${loadedMessage} (${submittedAt})`
    : loadedMessage;
  submitStatus.textContent = '답변을 수정한 뒤 기존 제출본에 저장할 수 있습니다.';
  submitStatus.className = 'submit-status success';
  window.scrollTo({ top: 0, behavior: 'smooth' });
});
