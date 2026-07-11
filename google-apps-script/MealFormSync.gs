const MEAL_SYNC_CONFIG = Object.freeze({
  TIME_ZONE: 'Asia/Seoul',
  SCHOOL_NAME: '율량중학교',
  OFFICE_CODE: 'M10',
  SCHOOL_CODE: '8011092',
  FORM_QUESTION_KEYWORD: '기대되는 급식 메뉴',
  TRIGGER_HANDLER: 'scheduledMealFormSync',
  TRIGGER_HOUR: 6,
  TRIGGER_NEAR_MINUTE: 5,
  NEXT_MEAL_SEARCH_DAYS: 14,
  NEIS_BASE_URL: 'https://open.neis.go.kr/hub/mealServiceDietInfo'
});

/**
 * 최초 1회 실행합니다.
 * - 현재 응답 스프레드시트와 연결된 Google Form을 자동 탐색
 * - 급식 질문을 확인
 * - 매일 오전 자동 실행 트리거 생성
 * - 오늘 급식으로 즉시 동기화
 */
function setupMealFormSync() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('설문 응답 스프레드시트에서 확장 프로그램 → Apps Script로 열어 실행해 주세요.');
  }

  const formUrl = findLinkedFormUrl_(ss);
  if (!formUrl) {
    throw new Error('이 스프레드시트에 연결된 Google Form을 찾지 못했습니다. 설문지의 응답 탭에서 이 시트를 응답 대상으로 연결해 주세요.');
  }

  const form = FormApp.openByUrl(formUrl);
  const question = findMealQuestion_(form);

  const props = PropertiesService.getScriptProperties();
  props.setProperties({
    SPREADSHEET_ID: ss.getId(),
    MEAL_FORM_ID: form.getId(),
    MEAL_FORM_QUESTION_ID: String(question.getId())
  }, false);

  // 같은 Form을 매일 재사용하므로 Google Forms의 '응답 1회 제한'은 해제합니다.
  // 이 설정이 켜져 있으면 학생은 다음 날 다시 응답할 수 없습니다.
  const oneResponseLimitWasEnabled = form.hasLimitOneResponsePerUser();
  if (oneResponseLimitWasEnabled) {
    form.setLimitOneResponsePerUser(false);
  }
  form.setShowLinkToRespondAgain(false);

  installMealFormDailyTrigger();
  const syncResult = syncTodayMealToForm();

  const result = {
    ok: true,
    school: MEAL_SYNC_CONFIG.SCHOOL_NAME,
    formTitle: form.getTitle(),
    questionTitle: question.getTitle(),
    oneResponseLimitWasDisabled: oneResponseLimitWasEnabled,
    trigger: `매일 ${MEAL_SYNC_CONFIG.TRIGGER_HOUR}시 전후`,
    sync: syncResult
  };

  ss.toast(
    syncResult.hasMeal
      ? `오늘 급식 ${syncResult.items.length}개를 설문 선택지로 반영했습니다.`
      : '오늘 급식이 없어 설문 응답을 닫았습니다.',
    '급식 설문 자동화 설정 완료',
    7
  );
  console.log(JSON.stringify(result, null, 2));
  return result;
}

/**
 * 매일 자동 트리거가 호출하는 진입점입니다.
 */
function scheduledMealFormSync() {
  return syncTodayMealToForm();
}

/**
 * 오늘 급식을 조회해 객관식/드롭다운 선택지를 교체합니다.
 */
function syncTodayMealToForm() {
  const ymd = Utilities.formatDate(new Date(), MEAL_SYNC_CONFIG.TIME_ZONE, 'yyyyMMdd');
  return syncMealToFormForDate_(ymd);
}

/**
 * 오늘 급식 데이터만 조회하고 Form은 수정하지 않습니다.
 */
function previewTodayMealChoices() {
  const ymd = Utilities.formatDate(new Date(), MEAL_SYNC_CONFIG.TIME_ZONE, 'yyyyMMdd');
  const meal = fetchMealForDate_(ymd);
  const result = {
    ok: true,
    date: ymd,
    dateLabel: formatKoreanDateLabel_(ymd),
    hasMeal: meal.items.length > 0,
    items: meal.items,
    mealType: meal.mealType
  };
  console.log(JSON.stringify(result, null, 2));
  return result;
}

/**
 * 주말이나 방학 중 테스트용입니다.
 * 오늘부터 지정 일수 안에서 가장 가까운 급식일을 찾아 실제 Form에 반영합니다.
 */
function syncNextAvailableMealToForm() {
  const base = startOfKoreanToday_();

  for (let offset = 0; offset <= MEAL_SYNC_CONFIG.NEXT_MEAL_SEARCH_DAYS; offset += 1) {
    const date = new Date(base.getTime());
    date.setDate(date.getDate() + offset);
    const ymd = Utilities.formatDate(date, MEAL_SYNC_CONFIG.TIME_ZONE, 'yyyyMMdd');
    const meal = fetchMealForDate_(ymd);

    if (meal.items.length > 0) {
      return applyMealToForm_(ymd, meal);
    }
  }

  throw new Error(`오늘부터 ${MEAL_SYNC_CONFIG.NEXT_MEAL_SEARCH_DAYS}일 안에 급식 데이터가 없습니다.`);
}

/**
 * 기존 동일 트리거를 제거하고 매일 오전 트리거를 하나만 생성합니다.
 */
function installMealFormDailyTrigger() {
  removeMealFormDailyTrigger();

  const trigger = ScriptApp.newTrigger(MEAL_SYNC_CONFIG.TRIGGER_HANDLER)
    .timeBased()
    .everyDays(1)
    .atHour(MEAL_SYNC_CONFIG.TRIGGER_HOUR)
    .nearMinute(MEAL_SYNC_CONFIG.TRIGGER_NEAR_MINUTE)
    .inTimezone(MEAL_SYNC_CONFIG.TIME_ZONE)
    .create();

  console.log(`급식 설문 동기화 트리거 생성: ${trigger.getUniqueId()}`);
  return trigger.getUniqueId();
}

function removeMealFormDailyTrigger() {
  const handler = MEAL_SYNC_CONFIG.TRIGGER_HANDLER;
  ScriptApp.getProjectTriggers()
    .filter((trigger) => trigger.getHandlerFunction() === handler)
    .forEach((trigger) => ScriptApp.deleteTrigger(trigger));
}

function syncMealToFormForDate_(ymd) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const meal = fetchMealForDate_(ymd);
    return applyMealToForm_(ymd, meal);
  } finally {
    lock.releaseLock();
  }
}

function applyMealToForm_(ymd, meal) {
  const form = getConfiguredMealForm_();
  const question = findMealQuestion_(form);
  const dateLabel = formatKoreanDateLabel_(ymd);
  const props = PropertiesService.getScriptProperties();

  if (!meal.items.length) {
    form.setCustomClosedFormMessage(`${dateLabel}은 급식이 없거나 급식 정보가 등록되지 않았습니다.`);
    form.setAcceptingResponses(false);

    props.setProperties({
      MEAL_LAST_SYNC_DATE: ymd,
      MEAL_LAST_SYNC_STATUS: 'NO_MEAL',
      MEAL_LAST_SYNC_ITEMS: '[]'
    }, false);

    const result = {
      ok: true,
      date: ymd,
      dateLabel,
      hasMeal: false,
      items: [],
      formAcceptingResponses: false
    };
    console.log(JSON.stringify(result, null, 2));
    return result;
  }

  question
    .setChoiceValues(meal.items)
    .setRequired(true)
    .setHelpText(`${dateLabel} ${meal.mealType || '중식'} 메뉴 중 가장 기대되는 메뉴를 하나 골라 주세요.`);

  form.setCustomClosedFormMessage('오늘의 급식 메뉴 투표가 마감되었습니다.');
  form.setAcceptingResponses(true);

  props.setProperties({
    MEAL_LAST_SYNC_DATE: ymd,
    MEAL_LAST_SYNC_STATUS: 'UPDATED',
    MEAL_LAST_SYNC_ITEMS: JSON.stringify(meal.items)
  }, false);

  const result = {
    ok: true,
    date: ymd,
    dateLabel,
    hasMeal: true,
    mealType: meal.mealType,
    items: meal.items,
    formAcceptingResponses: true
  };
  console.log(JSON.stringify(result, null, 2));
  return result;
}

function fetchMealForDate_(ymd) {
  validateYmd_(ymd);

  const params = {
    Type: 'json',
    pIndex: 1,
    pSize: 10,
    ATPT_OFCDC_SC_CODE: MEAL_SYNC_CONFIG.OFFICE_CODE,
    SD_SCHUL_CODE: MEAL_SYNC_CONFIG.SCHOOL_CODE,
    MLSV_YMD: ymd
  };

  const apiKey = PropertiesService.getScriptProperties().getProperty('NEIS_API_KEY');
  if (apiKey) params.KEY = apiKey;

  const query = Object.keys(params)
    .map((key) => `${encodeURIComponent(key)}=${encodeURIComponent(params[key])}`)
    .join('&');

  const response = UrlFetchApp.fetch(`${MEAL_SYNC_CONFIG.NEIS_BASE_URL}?${query}`, {
    method: 'get',
    muteHttpExceptions: true,
    followRedirects: true,
    headers: {
      Accept: 'application/json'
    }
  });

  const status = response.getResponseCode();
  const body = response.getContentText('UTF-8');

  if (status < 200 || status >= 300) {
    throw new Error(`나이스 급식 API HTTP ${status}: ${body.slice(0, 300)}`);
  }

  let data;
  try {
    data = JSON.parse(body);
  } catch (error) {
    throw new Error(`나이스 급식 API 응답을 JSON으로 해석하지 못했습니다: ${body.slice(0, 300)}`);
  }

  const rows = extractNeisMealRows_(data);
  const lunchRows = rows.filter((row) => {
    const code = String(row.MMEAL_SC_CODE || '');
    const name = String(row.MMEAL_SC_NM || '');
    return code === '2' || name.includes('중식');
  });
  const selectedRows = lunchRows.length ? lunchRows : rows;

  const items = [];
  selectedRows.forEach((row) => {
    splitNeisDishNames_(row.DDISH_NM || '').forEach((item) => items.push(item));
  });

  return {
    date: ymd,
    mealType: selectedRows[0] && selectedRows[0].MMEAL_SC_NM
      ? String(selectedRows[0].MMEAL_SC_NM)
      : '중식',
    items: uniqueStrings_(items)
  };
}

function extractNeisMealRows_(data) {
  const service = data && data.mealServiceDietInfo;
  if (Array.isArray(service)) {
    const rowBlock = service.find((block) => block && Array.isArray(block.row));
    return rowBlock ? rowBlock.row : [];
  }

  const result = data && data.RESULT;
  if (result && result.CODE === 'INFO-200') {
    return [];
  }
  if (result && result.MESSAGE) {
    throw new Error(`나이스 급식 API 오류: ${result.MESSAGE}`);
  }
  return [];
}

function splitNeisDishNames_(html) {
  return String(html || '')
    .split(/<br\s*\/?\s*>/i)
    .map((value) => decodeBasicHtml_(stripHtmlTags_(value)))
    .map(cleanDishName_)
    .filter(Boolean);
}

function cleanDishName_(value) {
  let name = String(value || '').replace(/\s+/g, ' ').trim();

  // 끝의 알레르기 번호: 메뉴명(1.2.5.6.) 또는 메뉴명 1.2.5.6.
  name = name.replace(/\s*\([0-9.\s]+\)\s*$/, '');
  name = name.replace(/\s+[0-9]+(?:\.[0-9]+)+\.?\s*$/, '');

  // 나이스 메뉴 앞쪽 장식 기호 정리
  name = name.replace(/^[-·•*★☆]+\s*/, '');
  return name.replace(/\s+/g, ' ').trim();
}

function stripHtmlTags_(value) {
  return String(value || '').replace(/<[^>]*>/g, '');
}

function decodeBasicHtml_(value) {
  return String(value || '')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&nbsp;/g, ' ');
}

function uniqueStrings_(values) {
  const seen = Object.create(null);
  return values.filter((value) => {
    const key = String(value).toLowerCase();
    if (!key || seen[key]) return false;
    seen[key] = true;
    return true;
  });
}

function getConfiguredMealForm_() {
  const props = PropertiesService.getScriptProperties();
  const savedFormId = props.getProperty('MEAL_FORM_ID');
  if (savedFormId) {
    return FormApp.openById(savedFormId);
  }

  const spreadsheetId = props.getProperty('SPREADSHEET_ID');
  if (!spreadsheetId) {
    throw new Error('먼저 setupMealFormSync 함수를 1회 실행해 주세요.');
  }

  const ss = SpreadsheetApp.openById(spreadsheetId);
  const formUrl = findLinkedFormUrl_(ss);
  if (!formUrl) {
    throw new Error('응답 스프레드시트에 연결된 Google Form을 찾지 못했습니다.');
  }

  const form = FormApp.openByUrl(formUrl);
  props.setProperty('MEAL_FORM_ID', form.getId());
  return form;
}

function findLinkedFormUrl_(ss) {
  const sheetUrl = ss.getSheets()
    .map((sheet) => {
      try {
        return sheet.getFormUrl();
      } catch (error) {
        return null;
      }
    })
    .find(Boolean);

  return sheetUrl || ss.getFormUrl() || null;
}

function findMealQuestion_(form) {
  const keyword = normalizeSearchText_(MEAL_SYNC_CONFIG.FORM_QUESTION_KEYWORD);
  const supported = form.getItems().filter((item) => {
    const type = item.getType();
    return type === FormApp.ItemType.MULTIPLE_CHOICE || type === FormApp.ItemType.LIST;
  });

  const matched = supported.find((item) =>
    normalizeSearchText_(item.getTitle()).includes(keyword)
  );

  if (!matched) {
    const titles = supported.map((item) => item.getTitle()).filter(Boolean);
    throw new Error(
      `설문에서 '${MEAL_SYNC_CONFIG.FORM_QUESTION_KEYWORD}' 문구가 들어간 객관식 또는 드롭다운 질문을 찾지 못했습니다. 현재 후보: ${titles.join(' / ') || '없음'}`
    );
  }

  if (matched.getType() === FormApp.ItemType.MULTIPLE_CHOICE) {
    return matched.asMultipleChoiceItem();
  }
  return matched.asListItem();
}

function normalizeSearchText_(value) {
  return String(value || '').toLowerCase().replace(/\s+/g, '').trim();
}

function validateYmd_(ymd) {
  if (!/^\d{8}$/.test(String(ymd))) {
    throw new Error(`날짜는 yyyyMMdd 형식이어야 합니다: ${ymd}`);
  }
}

function formatKoreanDateLabel_(ymd) {
  validateYmd_(ymd);
  const year = Number(ymd.slice(0, 4));
  const month = Number(ymd.slice(4, 6));
  const day = Number(ymd.slice(6, 8));
  const date = new Date(year, month - 1, day, 12, 0, 0);
  return Utilities.formatDate(date, MEAL_SYNC_CONFIG.TIME_ZONE, 'M월 d일 E요일');
}

function startOfKoreanToday_() {
  const ymd = Utilities.formatDate(new Date(), MEAL_SYNC_CONFIG.TIME_ZONE, 'yyyyMMdd');
  const year = Number(ymd.slice(0, 4));
  const month = Number(ymd.slice(4, 6));
  const day = Number(ymd.slice(6, 8));
  return new Date(year, month - 1, day, 12, 0, 0);
}
