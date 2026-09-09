const SHEET_NAME = 'Responses';
const HEADER_ROW = 2;
const START_COLUMN = 2;
const HEADER_COLOR = '#1A73E8';
const HEADER_TEXT_COLOR = '#FFFFFF';
const HEADERS = [
  '학번',
  '이름',
  'Timestamp',
  'Answer 1',
  'Answer 2',
  'Answer 3',
  'Answer 4',
  'Answer 5'
];
const ANSWER_KEYS = [
  'writtenAnswer1',
  'writtenAnswer2',
  'writtenAnswer3',
  'writtenAnswer4',
  'writtenAnswer5'
];
const USERS_SHEET_NAME = 'Users';
const USER_HEADERS = ['학번', '이름', '비밀번호', '활성화', '초기화 여부'];
const CLIPBOARD_LOG_SHEET_NAME = 'ClipboardLogs';
const CLIPBOARD_LOG_HEADERS = ['Timestamp', '학번', '이름', 'Question', 'Log'];
const INITIAL_PASSWORD = '1234';
const SESSION_TTL_SECONDS = 21600;
const LOGIN_FAILURE_LIMIT = 5;
const LOGIN_LOCK_SECONDS = 600;
const MAX_CONFIGURED_ANSWER_LENGTH = 50000;
const COPY_MARKER = '<copy>';

/**
 * 스프레드시트에 연결된 Apps Script 편집기에서 최초 한 번 실행하세요.
 * Responses/Users/ClipboardLogs 시트와 제목 행을 만들고 스프레드시트 ID를 저장합니다.
 */
function setup() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  PropertiesService.getScriptProperties()
    .setProperty('SPREADSHEET_ID', spreadsheet.getId());

  let sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_NAME);
  }

  // 이전 9개 열 헤더(이메일 포함)가 남지 않도록 제목 행만 초기화합니다.
  sheet.getRange(HEADER_ROW, START_COLUMN, 1, 9).clearContent();
  sheet
    .getRange(HEADER_ROW, START_COLUMN, 1, HEADERS.length)
    .setValues([HEADERS]);
  sheet.setFrozenRows(HEADER_ROW);
  formatResponseSheet_(sheet);

  let usersSheet = spreadsheet.getSheetByName(USERS_SHEET_NAME);
  if (!usersSheet) {
    usersSheet = spreadsheet.insertSheet(USERS_SHEET_NAME);
  }
  setupUsersSheet_(usersSheet);

  let clipboardLogSheet = spreadsheet.getSheetByName(CLIPBOARD_LOG_SHEET_NAME);
  if (!clipboardLogSheet) {
    clipboardLogSheet = spreadsheet.insertSheet(CLIPBOARD_LOG_SHEET_NAME);
  }
  setupClipboardLogSheet_(clipboardLogSheet);
}

function doGet() {
  return jsonResponse_({ result: 'ready' });
}

function doPost(e) {
  const action = String(e.parameter.action || '');
  if (action === 'login') {
    return handleLoginRequest_(e);
  }
  if (action === 'changePassword') {
    return handlePasswordChangeRequest_(e);
  }
  if (action === 'loadLatestSurvey') {
    return handleLoadLatestSurvey_(e);
  }
  if (action === 'logClipboardAttempt') {
    return handleClipboardAttempt_(e);
  }
  if (action === 'submitSurvey') {
    return handleSurveySubmission_(e);
  }
  return jsonResponse_({ result: 'error', message: '지원하지 않는 요청입니다.' });
}

function handleSurveySubmission_(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const spreadsheetId = PropertiesService.getScriptProperties()
      .getProperty('SPREADSHEET_ID');

    if (!spreadsheetId) {
      throw new Error('Apps Script 편집기에서 setup 함수를 먼저 실행해 주세요.');
    }

    const authToken = String(e.parameter.authToken || '');
    const session = getLoginSession_(authToken);
    if (!session) {
      throw new Error('로그인이 만료되었거나 유효하지 않습니다.');
    }

    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName(SHEET_NAME);
    if (!sheet) {
      throw new Error('Responses 시트를 찾을 수 없습니다. setup 함수를 다시 실행해 주세요.');
    }

    const existingResponse = findLatestResponse_(sheet, session.studentId);
    const clipboardLogFlags = getClipboardLogFlags_(
      spreadsheet,
      session.studentId,
      session.respondentName
    );
    const clipboardCache = CacheService.getScriptCache();
    const answers = ANSWER_KEYS.map(function (key, index) {
      const answerMaxLength = Math.floor(
        Number(
          e.parameter['answerMaxLength' + (index + 1)] ||
          e.parameter.answerMaxLength
        )
      );
      if (!Number.isFinite(answerMaxLength) ||
          answerMaxLength < 1 ||
          answerMaxLength > MAX_CONFIGURED_ANSWER_LENGTH) {
        throw new Error(
          'Question ' + (index + 1) + '의 글자 수 제한 설정이 올바르지 않습니다.'
        );
      }

      const answer = String(e.parameter[key] || '');
      if (answer.length > answerMaxLength) {
        throw new Error(
          'Answer ' + (index + 1) + '은(는) ' +
          answerMaxLength + '자를 넘을 수 없습니다.'
        );
      }
      const copyDetected = clipboardLogFlags[index] ||
        String(e.parameter['answerCopy' + (index + 1)]) === 'true' ||
        clipboardCache.get(clipboardAuditKey_(authToken, index + 1)) === 'true';
      return safeCellValue_(copyDetected ? COPY_MARKER + answer : answer);
    });

    const submittedAt = new Date();
    const row = [
      safeCellValue_(session.studentId),
      safeCellValue_(session.respondentName),
      submittedAt
    ].concat(answers);
    const targetRow = existingResponse
      ? existingResponse.rowNumber
      : Math.max(sheet.getLastRow() + 1, HEADER_ROW + 1);

    const targetRange = sheet
      .getRange(targetRow, START_COLUMN, 1, row.length)
      .setValues([row]);
    formatSavedResponseRow_(sheet, targetRange, targetRow);
    SpreadsheetApp.flush();

    const savedStudentId = sheet
      .getRange(targetRow, START_COLUMN)
      .getDisplayValue()
      .trim();
    if (savedStudentId !== session.studentId) {
      throw new Error('시트에 제출 내용을 저장하지 못했습니다. 다시 시도해 주세요.');
    }

    return loginHtmlResponse_({
      type: 'survey-submit-result',
      success: true,
      row: targetRow,
      updated: Boolean(existingResponse),
      submittedAt: submittedAt.toISOString()
    });
  } catch (error) {
    return loginHtmlResponse_({
      type: 'survey-submit-result',
      success: false,
      message: String(error)
    });
  } finally {
    lock.releaseLock();
  }
}

function handleClipboardAttempt_(e) {
  const authToken = String(e.parameter.authToken || '');
  const session = getLoginSession_(authToken);
  const questionNumber = Math.floor(Number(e.parameter.questionNumber));
  const clipboardAction = String(e.parameter.clipboardAction || 'clipboard').slice(0, 40);

  if (!session) {
    return loginHtmlResponse_({
      type: 'survey-clipboard-log-result',
      success: false,
      message: '로그인이 만료되었거나 유효하지 않습니다.'
    });
  }
  if (questionNumber < 1 || questionNumber > ANSWER_KEYS.length) {
    return loginHtmlResponse_({
      type: 'survey-clipboard-log-result',
      success: false,
      message: '질문 번호가 올바르지 않습니다.'
    });
  }

  const actualTransfer = isActualAnswerTransfer_(clipboardAction);
  if (actualTransfer) {
    CacheService.getScriptCache().put(
      clipboardAuditKey_(authToken, questionNumber),
      'true',
      SESSION_TTL_SECONDS
    );
  }

  const spreadsheetId = PropertiesService.getScriptProperties()
    .getProperty('SPREADSHEET_ID');
  if (!spreadsheetId) {
    return loginHtmlResponse_({
      type: 'survey-clipboard-log-result',
      success: false,
      message: 'Apps Script 편집기에서 setup 함수를 먼저 실행해 주세요.'
    });
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    let logSheet = spreadsheet.getSheetByName(CLIPBOARD_LOG_SHEET_NAME);
    if (!logSheet) {
      logSheet = spreadsheet.insertSheet(CLIPBOARD_LOG_SHEET_NAME);
      setupClipboardLogSheet_(logSheet);
    }
    const targetRow = Math.max(logSheet.getLastRow() + 1, HEADER_ROW + 1);
    const logRange = logSheet
      .getRange(targetRow, START_COLUMN, 1, CLIPBOARD_LOG_HEADERS.length)
      .setValues([[
        new Date(),
        safeCellValue_(session.studentId),
        safeCellValue_(session.respondentName),
        'Question ' + questionNumber,
        COPY_MARKER + ' ' + (actualTransfer ? 'actual ' : 'attempt ') + clipboardAction
      ]]);
    logRange
      .setFontColor('#303030')
      .setWrap(true)
      .setVerticalAlignment('top');
    logSheet
      .getRange(targetRow, START_COLUMN)
      .setNumberFormat('yyyy-mm-dd hh:mm:ss');
    logSheet
      .getRange(targetRow, START_COLUMN + 1)
      .setNumberFormat('@');
  } catch (error) {
    console.error('클립보드 감사 로그 저장 실패:', error);
    return loginHtmlResponse_({
      type: 'survey-clipboard-log-result',
      success: false,
      message: '클립보드 감사 로그를 저장하지 못했습니다.'
    });
  } finally {
    if (lock.hasLock()) {
      lock.releaseLock();
    }
  }

  return loginHtmlResponse_({
    type: 'survey-clipboard-log-result',
    success: true
  });
}

function handleLoadLatestSurvey_(e) {
  try {
    const authToken = String(e.parameter.authToken || '');
    const session = getLoginSession_(authToken);
    if (!session) {
      throw new Error('로그인이 만료되었거나 유효하지 않습니다. 다시 로그인해 주세요.');
    }

    const spreadsheetId = PropertiesService.getScriptProperties()
      .getProperty('SPREADSHEET_ID');
    if (!spreadsheetId) {
      throw new Error('Apps Script 편집기에서 setup 함수를 먼저 실행해 주세요.');
    }

    const sheet = SpreadsheetApp
      .openById(spreadsheetId)
      .getSheetByName(SHEET_NAME);
    if (!sheet) {
      throw new Error('Responses 시트를 찾을 수 없습니다.');
    }

    const latestResponse = findLatestResponse_(sheet, session.studentId);
    if (!latestResponse) {
      return loginHtmlResponse_({
        type: 'survey-latest-result',
        success: true,
        found: false
      });
    }

    const latestRow = latestResponse.row;
    const submittedAt = latestRow[2] instanceof Date
      ? latestRow[2].toISOString()
      : String(latestRow[2] || '');
    const loadedAnswers = ANSWER_KEYS.map(function (_, index) {
      return readStoredAnswer_(latestRow[index + 3]);
    });
    const answers = loadedAnswers.map(function (entry) {
      return entry.answer;
    });
    const clipboardLogFlags = getClipboardLogFlags_(
      spreadsheet,
      session.studentId,
      session.respondentName
    );
    const copyFlags = loadedAnswers.map(function (_, index) {
      if (clipboardLogFlags[index]) {
        CacheService.getScriptCache().put(
          clipboardAuditKey_(authToken, index + 1),
          'true',
          SESSION_TTL_SECONDS
        );
      }
      return clipboardLogFlags[index];
    });

    return loginHtmlResponse_({
      type: 'survey-latest-result',
      success: true,
      found: true,
      submittedAt: submittedAt,
      answers: answers,
      copyFlags: copyFlags
    });
  } catch (error) {
    return loginHtmlResponse_({
      type: 'survey-latest-result',
      success: false,
      message: String(error)
    });
  }
}

function findLatestResponse_(sheet, studentId) {
  if (!sheet || sheet.getLastRow() <= HEADER_ROW) {
    return null;
  }

  const rows = sheet
    .getRange(
      HEADER_ROW + 1,
      START_COLUMN,
      sheet.getLastRow() - HEADER_ROW,
      HEADERS.length
    )
    .getValues();
  let latestResponse = null;
  let latestTime = -1;

  rows.forEach(function (row, index) {
    if (String(row[0]).trim() !== studentId) {
      return;
    }

    const timestamp = row[2] instanceof Date
      ? row[2].getTime()
      : new Date(row[2]).getTime();
    const comparableTime = Number.isFinite(timestamp) ? timestamp : index;
    if (!latestResponse || comparableTime >= latestTime) {
      latestResponse = {
        row: row,
        rowNumber: HEADER_ROW + 1 + index
      };
      latestTime = comparableTime;
    }
  });

  return latestResponse;
}

function handleLoginRequest_(e) {
  const studentId = String(e.parameter.studentId || '').trim();
  const respondentName = String(e.parameter.respondentName || '').trim();
  const password = String(e.parameter.password || '');

  if (!studentId || !respondentName || !password) {
    return loginHtmlResponse_({
      type: 'survey-login-result',
      success: false,
      message: '학번, 이름, 비밀번호를 모두 입력해 주세요.'
    });
  }

  const cache = CacheService.getScriptCache();
  const failureKey = 'login-fail:' + studentId;
  const failureCount = Number(cache.get(failureKey) || 0);
  if (failureCount >= LOGIN_FAILURE_LIMIT) {
    return loginHtmlResponse_({
      type: 'survey-login-result',
      success: false,
      message: '로그인 시도가 너무 많습니다. 10분 후 다시 시도해 주세요.'
    });
  }

  try {
    const spreadsheetId = PropertiesService.getScriptProperties()
      .getProperty('SPREADSHEET_ID');
    if (!spreadsheetId) {
      throw new Error('setup 함수를 먼저 실행해 주세요.');
    }

    const usersSheet = SpreadsheetApp
      .openById(spreadsheetId)
      .getSheetByName(USERS_SHEET_NAME);
    const user = findUser_(usersSheet, studentId, respondentName);

    if (!user || !user.active || user.password !== password) {
      cache.put(failureKey, String(failureCount + 1), LOGIN_LOCK_SECONDS);
      return loginHtmlResponse_({
        type: 'survey-login-result',
        success: false,
        message: '학번, 이름 또는 비밀번호가 올바르지 않습니다.'
      });
    }

    cache.remove(failureKey);
    if (user.passwordWasReset) {
      return loginHtmlResponse_({
        type: 'survey-login-result',
        success: true,
        studentId: user.studentId,
        respondentName: user.respondentName,
        passwordWasReset: true
      });
    }

    const token = createLoginSession_(
      cache,
      user.studentId,
      user.respondentName
    );

    return loginHtmlResponse_({
      type: 'survey-login-result',
      success: true,
      studentId: user.studentId,
      respondentName: user.respondentName,
      token: token,
      passwordWasReset: user.passwordWasReset
    });
  } catch (error) {
    console.error('로그인 처리 실패:', error);
    return loginHtmlResponse_({
      type: 'survey-login-result',
      success: false,
      message: '로그인 처리 중 오류가 발생했습니다.'
    });
  }
}

function handlePasswordChangeRequest_(e) {
  const responseType = 'survey-password-change-result';
  const studentId = String(e.parameter.studentId || '').trim();
  const respondentName = String(e.parameter.respondentName || '').trim();
  const currentPassword = String(e.parameter.password || '');
  const newPassword = String(e.parameter.newPassword || '');

  if (!studentId || !respondentName || !currentPassword || !newPassword) {
    return loginHtmlResponse_({
      type: responseType,
      success: false,
      message: '학번, 이름, 현재 비밀번호와 새 비밀번호를 모두 입력해 주세요.'
    });
  }
  if (newPassword.length < 4) {
    return loginHtmlResponse_({
      type: responseType,
      success: false,
      message: '새 비밀번호는 4자 이상으로 입력해 주세요.'
    });
  }
  if (newPassword === currentPassword) {
    return loginHtmlResponse_({
      type: responseType,
      success: false,
      message: '현재 비밀번호와 다른 새 비밀번호를 입력해 주세요.'
    });
  }
  if (newPassword === INITIAL_PASSWORD) {
    return loginHtmlResponse_({
      type: responseType,
      success: false,
      message: '초기 비밀번호 1234는 새 비밀번호로 사용할 수 없습니다.'
    });
  }

  const cache = CacheService.getScriptCache();
  const failureKey = 'login-fail:' + studentId;
  const failureCount = Number(cache.get(failureKey) || 0);
  if (failureCount >= LOGIN_FAILURE_LIMIT) {
    return loginHtmlResponse_({
      type: responseType,
      success: false,
      message: '확인 시도가 너무 많습니다. 10분 후 다시 시도해 주세요.'
    });
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    const spreadsheetId = PropertiesService.getScriptProperties()
      .getProperty('SPREADSHEET_ID');
    if (!spreadsheetId) {
      throw new Error('setup 함수를 먼저 실행해 주세요.');
    }

    const usersSheet = SpreadsheetApp
      .openById(spreadsheetId)
      .getSheetByName(USERS_SHEET_NAME);
    const user = findUser_(usersSheet, studentId, respondentName);

    if (!user || !user.active || user.password !== currentPassword) {
      cache.put(failureKey, String(failureCount + 1), LOGIN_LOCK_SECONDS);
      return loginHtmlResponse_({
        type: responseType,
        success: false,
        message: '학번, 이름 또는 현재 비밀번호가 올바르지 않습니다.'
      });
    }

    usersSheet
      .getRange(user.rowNumber, START_COLUMN + 2)
      .setNumberFormat('@')
      .setValue(newPassword);
    usersSheet
      .getRange(user.rowNumber, START_COLUMN + 4)
      .setValue(false);
    SpreadsheetApp.flush();
    cache.remove(failureKey);
    const token = createLoginSession_(
      cache,
      user.studentId,
      user.respondentName
    );

    return loginHtmlResponse_({
      type: responseType,
      success: true,
      message: '비밀번호가 변경되었습니다.',
      studentId: user.studentId,
      respondentName: user.respondentName,
      token: token
    });
  } catch (error) {
    console.error('비밀번호 변경 처리 실패:', error);
    return loginHtmlResponse_({
      type: responseType,
      success: false,
      message: '비밀번호 변경 중 오류가 발생했습니다.'
    });
  } finally {
    if (lock.hasLock()) {
      lock.releaseLock();
    }
  }
}

function findUser_(sheet, studentId, respondentName) {
  if (!sheet || sheet.getLastRow() <= HEADER_ROW) {
    return null;
  }

  const rows = sheet
    .getRange(
      HEADER_ROW + 1,
      START_COLUMN,
      sheet.getLastRow() - HEADER_ROW,
      USER_HEADERS.length
    )
    .getDisplayValues();

  for (let index = 0; index < rows.length; index++) {
    const row = rows[index];
    if (row[0].trim() === studentId && row[1].trim() === respondentName) {
      return {
        studentId: row[0].trim(),
        respondentName: row[1].trim(),
        password: row[2],
        active: row[3].trim().toUpperCase() !== 'FALSE',
        passwordWasReset:
          row[2] === INITIAL_PASSWORD &&
          row[4].trim().toUpperCase() === 'TRUE',
        rowNumber: HEADER_ROW + 1 + index
      };
    }
  }
  return null;
}

function createLoginSession_(cache, studentId, respondentName) {
  const token = Utilities.getUuid() + Utilities.getUuid();
  cache.put(
    'login-session:' + token,
    JSON.stringify({
      studentId: studentId,
      respondentName: respondentName
    }),
    SESSION_TTL_SECONDS
  );
  return token;
}

function getLoginSession_(token) {
  if (!token) {
    return null;
  }

  const cached = CacheService.getScriptCache().get('login-session:' + token);
  if (!cached) {
    return null;
  }

  try {
    return JSON.parse(cached);
  } catch (error) {
    return null;
  }
}

function loginHtmlResponse_(data) {
  const payload = JSON.stringify(data)
    .replace(/</g, '\\u003c')
    .replace(/>/g, '\\u003e')
    .replace(/&/g, '\\u0026')
    .replace(/\u2028/g, '\\u2028')
    .replace(/\u2029/g, '\\u2029');

  return HtmlService
    .createHtmlOutput(
      '<!doctype html><html><body><script>' +
      'window.top.postMessage(' + payload + ', "*");' +
      '</script></body></html>'
    )
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function setupUsersSheet_(sheet) {
  sheet
    .getRange(HEADER_ROW, START_COLUMN, 1, USER_HEADERS.length)
    .setValues([USER_HEADERS])
    .setBackground(HEADER_COLOR)
    .setFontColor(HEADER_TEXT_COLOR)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setFrozenRows(HEADER_ROW);
  sheet.setRowHeight(HEADER_ROW, 36);
  sheet.setColumnWidth(START_COLUMN, 120);
  sheet.setColumnWidth(START_COLUMN + 1, 120);
  sheet.setColumnWidth(START_COLUMN + 2, 140);
  sheet.setColumnWidth(START_COLUMN + 3, 90);
  sheet.setColumnWidth(START_COLUMN + 4, 110);
  sheet
    .getRange(HEADER_ROW + 1, START_COLUMN, sheet.getMaxRows() - HEADER_ROW, 1)
    .setNumberFormat('@');
  sheet
    .getRange(HEADER_ROW + 1, START_COLUMN + 2, sheet.getMaxRows() - HEADER_ROW, 1)
    .setNumberFormat('@');
  fillBlankPasswords_(sheet);
}

function setupClipboardLogSheet_(sheet) {
  sheet
    .getRange(HEADER_ROW, START_COLUMN, 1, CLIPBOARD_LOG_HEADERS.length)
    .setValues([CLIPBOARD_LOG_HEADERS])
    .setBackground(HEADER_COLOR)
    .setFontColor(HEADER_TEXT_COLOR)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setFrozenRows(HEADER_ROW);
  sheet.setRowHeight(HEADER_ROW, 36);
  sheet.setColumnWidth(START_COLUMN, 170);
  sheet.setColumnWidth(START_COLUMN + 1, 120);
  sheet.setColumnWidth(START_COLUMN + 2, 120);
  sheet.setColumnWidth(START_COLUMN + 3, 100);
  sheet.setColumnWidth(START_COLUMN + 4, 180);
  sheet
    .getRange(HEADER_ROW + 1, START_COLUMN, sheet.getMaxRows() - HEADER_ROW, 1)
    .setNumberFormat('yyyy-mm-dd hh:mm:ss');
  sheet
    .getRange(HEADER_ROW + 1, START_COLUMN + 1, sheet.getMaxRows() - HEADER_ROW, 1)
    .setNumberFormat('@');
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('설문 관리')
    .addItem('선택 사용자 비밀번호를 1234로 초기화', 'resetSelectedUserPasswords')
    .addItem('빈 비밀번호를 모두 1234로 설정', 'initializeBlankPasswords')
    .addSeparator()
    .addItem('응답을 학번순으로 정렬', 'sortResponses')
    .addToUi();
}

function resetSelectedUserPasswords() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getActiveSheet();
  const selectedRange = sheet.getActiveRange();
  const ui = SpreadsheetApp.getUi();

  if (sheet.getName() !== USERS_SHEET_NAME || !selectedRange) {
    ui.alert('Users 시트에서 초기화할 사용자 행을 선택해 주세요.');
    return;
  }

  const firstRow = Math.max(selectedRange.getRow(), HEADER_ROW + 1);
  const lastRow = selectedRange.getLastRow();
  if (lastRow < firstRow) {
    ui.alert('초기화할 사용자 행을 선택해 주세요.');
    return;
  }

  const studentIds = sheet
    .getRange(firstRow, START_COLUMN, lastRow - firstRow + 1, 1)
    .getDisplayValues();
  let resetCount = 0;

  studentIds.forEach(function (row, index) {
    if (!row[0].trim()) {
      return;
    }
    const targetRow = firstRow + index;
    sheet.getRange(targetRow, START_COLUMN + 2).setValue(INITIAL_PASSWORD);
    sheet.getRange(targetRow, START_COLUMN + 4).setValue(true);
    resetCount++;
  });

  ui.alert(resetCount + '명의 비밀번호를 1234로 초기화했습니다.');
}

function initializeBlankPasswords() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(USERS_SHEET_NAME);
  const ui = SpreadsheetApp.getUi();

  if (!sheet || sheet.getLastRow() <= HEADER_ROW) {
    ui.alert('Users 시트에 사용자를 먼저 입력해 주세요.');
    return;
  }

  const initializedCount = fillBlankPasswords_(sheet);
  ui.alert(initializedCount + '명의 빈 비밀번호를 1234로 설정했습니다.');
}

function fillBlankPasswords_(sheet) {
  if (!sheet || sheet.getLastRow() <= HEADER_ROW) {
    return 0;
  }

  const rowCount = sheet.getLastRow() - HEADER_ROW;
  const range = sheet.getRange(
    HEADER_ROW + 1,
    START_COLUMN,
    rowCount,
    USER_HEADERS.length
  );
  const values = range.getValues();
  let initializedCount = 0;

  values.forEach(function (row) {
    if (!String(row[0]).trim()) {
      return;
    }

    if (!String(row[2]).trim()) {
      row[2] = INITIAL_PASSWORD;
      initializedCount++;
    }

    // 초기 비밀번호인 계정은 다음 로그인 때 반드시 변경 화면으로 보냅니다.
    if (String(row[2]) === INITIAL_PASSWORD) {
      row[4] = true;
    }
    if (row[3] === '') {
      row[3] = true;
    }
  });

  range.setValues(values);
  return initializedCount;
}

function onEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();
  if (sheet.getName() !== USERS_SHEET_NAME || range.getRow() <= HEADER_ROW) {
    return;
  }

  // 학번을 새로 입력하면 빈 비밀번호/상태를 자동으로 초기화합니다.
  if (range.getColumn() <= START_COLUMN && range.getLastColumn() >= START_COLUMN) {
    fillBlankPasswords_(sheet);
    return;
  }

  // 비밀번호를 직접 변경하면 초기화 여부도 함께 갱신합니다.
  if (range.getColumn() === START_COLUMN + 2 && range.getNumColumns() === 1) {
    const passwords = range.getDisplayValues();
    const resetFlags = passwords.map(function (row) {
      return [row[0] === INITIAL_PASSWORD];
    });
    sheet
      .getRange(range.getRow(), START_COLUMN + 4, range.getNumRows(), 1)
      .setValues(resetFlags);
  }
}

/**
 * 기존 응답도 포함하여 표 서식을 다시 적용하고 학번 오름차순으로 정렬합니다.
 * Apps Script 편집기에서 이 함수를 직접 실행해도 됩니다.
 */
function sortResponses() {
  const spreadsheetId = PropertiesService.getScriptProperties()
    .getProperty('SPREADSHEET_ID');
  if (!spreadsheetId) {
    throw new Error('setup 함수를 먼저 실행해 주세요.');
  }

  const sheet = SpreadsheetApp
    .openById(spreadsheetId)
    .getSheetByName(SHEET_NAME);
  if (!sheet) {
    throw new Error('Responses 시트를 찾을 수 없습니다.');
  }

  formatResponseSheet_(sheet);
}

function formatResponseSheet_(sheet) {
  const lastDataRow = Math.max(sheet.getLastRow(), HEADER_ROW);
  const tableLastRow = Math.max(lastDataRow, HEADER_ROW + 1);
  const tableRowCount = tableLastRow - HEADER_ROW + 1;
  const dataRowCount = lastDataRow - HEADER_ROW;
  const tableRange = sheet.getRange(
    HEADER_ROW,
    START_COLUMN,
    tableRowCount,
    HEADERS.length
  );

  // 필터를 다시 만들기 전에 실제 데이터만 학번(B열) 기준으로 정렬합니다.
  const existingFilter = sheet.getFilter();
  if (existingFilter) {
    existingFilter.remove();
  }

  if (dataRowCount > 1) {
    sheet
      .getRange(HEADER_ROW + 1, START_COLUMN, dataRowCount, HEADERS.length)
      .sort({ column: START_COLUMN, ascending: true });
  }

  // Responses 전용 시트의 기존 교차 행 색상을 갱신합니다.
  sheet.getBandings().forEach(function (banding) {
    banding.remove();
  });
  tableRange.applyRowBanding(SpreadsheetApp.BandingTheme.BLUE, true, false);

  const headerRange = sheet.getRange(
    HEADER_ROW,
    START_COLUMN,
    1,
    HEADERS.length
  );
  headerRange
    .setBackground(HEADER_COLOR)
    .setFontColor(HEADER_TEXT_COLOR)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  tableRange
    .setBorder(
      true,
      true,
      true,
      true,
      true,
      true,
      '#B7C9E2',
      SpreadsheetApp.BorderStyle.SOLID
    )
    .setWrap(true)
    .setVerticalAlignment('top');
  headerRange.setVerticalAlignment('middle');
  if (tableRowCount > 1) {
    sheet
      .getRange(
        HEADER_ROW + 1,
        START_COLUMN,
        tableRowCount - 1,
        HEADERS.length
      )
      .setFontColor('#303030');
  }

  // 학번은 앞자리 0이 사라지지 않도록 텍스트 형식으로 유지합니다.
  sheet
    .getRange(HEADER_ROW + 1, START_COLUMN, sheet.getMaxRows() - HEADER_ROW, 1)
    .setNumberFormat('@');

  sheet.setRowHeight(HEADER_ROW, 36);
  sheet.setColumnWidth(START_COLUMN, 120);      // 학번
  sheet.setColumnWidth(START_COLUMN + 1, 100);  // 이름
  sheet.setColumnWidth(START_COLUMN + 2, 170);  // Timestamp
  for (let column = START_COLUMN + 3; column < START_COLUMN + HEADERS.length; column++) {
    sheet.setColumnWidth(column, 300);
  }

  tableRange.createFilter();
}

function formatSavedResponseRow_(sheet, range, rowNumber) {
  const background = rowNumber % 2 === 0 ? '#E8F0FE' : '#FFFFFF';
  range
    .setBackground(background)
    .setFontColor('#303030')
    .setBorder(
      true,
      true,
      true,
      true,
      true,
      true,
      '#B7C9E2',
      SpreadsheetApp.BorderStyle.SOLID
    )
    .setWrap(true)
    .setVerticalAlignment('top');
  sheet.getRange(rowNumber, START_COLUMN).setNumberFormat('@');
  sheet
    .getRange(rowNumber, START_COLUMN + 2)
    .setNumberFormat('yyyy-mm-dd hh:mm:ss');
}

// 입력값이 스프레드시트 수식으로 실행되지 않도록 보호합니다.
function safeCellValue_(value) {
  const text = String(value);
  return /^[=+\-@]/.test(text) ? "'" + text : text;
}

function restoreSafeCellValue_(value) {
  const text = String(value == null ? '' : value);
  return /^'[=+\-@]/.test(text) ? text.slice(1) : text;
}

function readStoredAnswer_(value) {
  const text = restoreSafeCellValue_(value);
  const copyDetected = text.indexOf(COPY_MARKER) === 0;
  return {
    answer: copyDetected ? text.slice(COPY_MARKER.length) : text,
    copyDetected: copyDetected
  };
}

function clipboardAuditKey_(authToken, questionNumber) {
  return 'clipboard-transfer:' + authToken + ':' + questionNumber;
}

function isActualAnswerTransfer_(clipboardAction) {
  return [
    'insertFromPaste',
    'insertFromDrop',
    'deleteByCut',
    'untrusted-input',
    'unobserved-value-change'
  ].indexOf(clipboardAction) !== -1;
}

function getClipboardLogFlags_(spreadsheet, studentId, respondentName) {
  const flags = ANSWER_KEYS.map(function () {
    return false;
  });
  const logSheet = spreadsheet.getSheetByName(CLIPBOARD_LOG_SHEET_NAME);
  if (!logSheet || logSheet.getLastRow() <= HEADER_ROW) {
    return flags;
  }

  const rows = logSheet
    .getRange(
      HEADER_ROW + 1,
      START_COLUMN,
      logSheet.getLastRow() - HEADER_ROW,
      CLIPBOARD_LOG_HEADERS.length
    )
    .getDisplayValues();
  rows.forEach(function (row) {
    if (row[1].trim() !== studentId || row[2].trim() !== respondentName) {
      return;
    }
    const questionMatch = /^Question (\d+)$/.exec(row[3].trim());
    const questionNumber = questionMatch ? Number(questionMatch[1]) : 0;
    const logText = row[4].trim();
    const loggedAction = logText.indexOf(COPY_MARKER + ' actual ') === 0
      ? logText.slice((COPY_MARKER + ' actual ').length)
      : '';
    if (questionNumber >= 1 &&
        questionNumber <= flags.length &&
        isActualAnswerTransfer_(loggedAction)) {
      flags[questionNumber - 1] = true;
    }
  });
  return flags;
}

function jsonResponse_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
