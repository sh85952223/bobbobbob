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

function setupMealFormSync() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('설문 응답 스프레드시트에서 Apps Script를 열어 실행해 주세요.');

  const formUrl = ss.getFormUrl();
  if (!formUrl) throw new Error('이 스프레드시트에 연결된 Google Form을 찾지 못했습니다.');

  const form = FormApp.openByUrl(formUrl);
  const question = findMealQuestion_(form);
  const props = PropertiesService.getScriptProperties();

  props.setProperties({
    SPREADSHEET_ID: ss.getId(),
    MEAL_FORM_ID: form.getId(),
    MEAL_FORM_QUESTION_ID: String(question.getId())
  }, false);

  if (form.hasLimitOneResponsePerUser()) form.setLimitOneResponsePerUser(false);
  form.setShowLinkToRespondAgain(false);

  installMealFormDailyTrigger();
  const result = syncTodayMealToForm();
  ss.toast(`${result.dateLabel} 급식 메뉴 동기화 완료`, '급식 설문', 6);
  return result;
}

function scheduledMealFormSync() {
  return syncTodayMealToForm();
}

function syncTodayMealToForm() {
  const context = getCurrentMealVoteContext_();
  const meal = fetchMealForDate_(context.targetYmd);
  return applyMealToForm_(context, meal);
}

function previewTodayMealChoices() {
  const context = getCurrentMealVoteContext_();
  const meal = fetchMealForDate_(context.targetYmd);
  const result = {
    ok: true,
    today: context.todayYmd,
    date: context.targetYmd,
    dateLabel: context.dateLabel,
    mode: context.mode,
    items: meal.items,
    responseWindowStart: context.windowStartYmd,
    responseWindowEnd: context.windowEndYmd
  };
  console.log(JSON.stringify(result, null, 2));
  return result;
}

function syncNextAvailableMealToForm() {
  const todayYmd = getKoreanTodayYmd_();
  for (let offset = 0; offset <= MEAL_SYNC_CONFIG.NEXT_MEAL_SEARCH_DAYS; offset += 1) {
    const targetYmd = addDaysToYmd_(todayYmd, offset);
    const meal = fetchMealForDate_(targetYmd);
    if (meal.items.length) {
      return applyMealToForm_(buildMealVoteContext_(todayYmd, targetYmd, 'NEXT_AVAILABLE'), meal);
    }
  }
  throw new Error(`앞으로 ${MEAL_SYNC_CONFIG.NEXT_MEAL_SEARCH_DAYS}일 안에 급식 데이터가 없습니다.`);
}

function installMealFormDailyTrigger() {
  removeMealFormDailyTrigger();
  return ScriptApp.newTrigger(MEAL_SYNC_CONFIG.TRIGGER_HANDLER)
    .timeBased()
    .everyDays(1)
    .atHour(MEAL_SYNC_CONFIG.TRIGGER_HOUR)
    .nearMinute(MEAL_SYNC_CONFIG.TRIGGER_NEAR_MINUTE)
    .inTimezone(MEAL_SYNC_CONFIG.TIME_ZONE)
    .create()
    .getUniqueId();
}

function removeMealFormDailyTrigger() {
  ScriptApp.getProjectTriggers()
    .filter((trigger) => trigger.getHandlerFunction() === MEAL_SYNC_CONFIG.TRIGGER_HANDLER)
    .forEach((trigger) => ScriptApp.deleteTrigger(trigger));
}

function getCurrentMealVoteContext_() {
  const todayYmd = getKoreanTodayYmd_();
  const weekday = getWeekdayFromYmd_(todayYmd);

  if (weekday === 6) {
    return buildMealVoteContext_(todayYmd, addDaysToYmd_(todayYmd, 2), 'WEEKEND_MONDAY');
  }
  if (weekday === 0) {
    return buildMealVoteContext_(todayYmd, addDaysToYmd_(todayYmd, 1), 'WEEKEND_MONDAY');
  }
  return buildMealVoteContext_(todayYmd, todayYmd, 'TODAY');
}

function buildMealVoteContext_(todayYmd, targetYmd, mode) {
  const isMonday = getWeekdayFromYmd_(targetYmd) === 1;
  return {
    todayYmd,
    targetYmd,
    dateLabel: formatKoreanDateLabel_(targetYmd),
    mode,
    windowStartYmd: isMonday ? addDaysToYmd_(targetYmd, -2) : targetYmd,
    windowEndYmd: targetYmd
  };
}

function applyMealToForm_(context, meal) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const form = getConfiguredMealForm_();
    const question = findMealQuestion_(form);
    const props = PropertiesService.getScriptProperties();

    props.setProperties({
      MEAL_LAST_SYNC_DATE: context.todayYmd,
      MEAL_VOTE_TARGET_DATE: context.targetYmd,
      MEAL_VOTE_WINDOW_START: context.windowStartYmd,
      MEAL_VOTE_WINDOW_END: context.windowEndYmd,
      MEAL_VOTE_MODE: context.mode,
      MEAL_LAST_SYNC_ITEMS: JSON.stringify(meal.items)
    }, false);

    if (!meal.items.length) {
      form.setCustomClosedFormMessage(`${context.dateLabel}은 급식이 없거나 급식 정보가 등록되지 않았습니다.`);
      form.setAcceptingResponses(false);
      props.setProperty('MEAL_LAST_SYNC_STATUS', 'NO_MEAL');
      return {
        ok: true,
        date: context.targetYmd,
        dateLabel: context.dateLabel,
        mode: context.mode,
        hasMeal: false,
        items: [],
        formAcceptingResponses: false
      };
    }

    const prefix = context.mode === 'WEEKEND_MONDAY' ? '주말 미리 투표 · ' : '';
    question
      .setChoiceValues(meal.items)
      .setRequired(true)
      .setHelpText(`${prefix}${context.dateLabel} ${meal.mealType} 메뉴 중 가장 기대되는 메뉴를 하나 골라 주세요.`);

    form.setAcceptingResponses(true);
    props.setProperty('MEAL_LAST_SYNC_STATUS', 'UPDATED');

    return {
      ok: true,
      date: context.targetYmd,
      dateLabel: context.dateLabel,
      mode: context.mode,
      hasMeal: true,
      items: meal.items,
      responseWindowStart: context.windowStartYmd,
      responseWindowEnd: context.windowEndYmd,
      formAcceptingResponses: true
    };
  } finally {
    lock.releaseLock();
  }
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
    muteHttpExceptions: true,
    headers: { Accept: 'application/json' }
  });

  if (response.getResponseCode() < 200 || response.getResponseCode() >= 300) {
    throw new Error(`나이스 급식 API HTTP ${response.getResponseCode()}`);
  }

  const data = JSON.parse(response.getContentText('UTF-8'));
  const service = data.mealServiceDietInfo;
  const rowBlock = Array.isArray(service)
    ? service.find((block) => Array.isArray(block.row))
    : null;
  const rows = rowBlock ? rowBlock.row : [];
  const lunchRows = rows.filter((row) => String(row.MMEAL_SC_CODE || '') === '2' || String(row.MMEAL_SC_NM || '').includes('중식'));
  const selectedRows = lunchRows.length ? lunchRows : rows;
  const items = [];

  selectedRows.forEach((row) => {
    String(row.DDISH_NM || '')
      .split(/<br\s*\/?\s*>/i)
      .map((value) => cleanDishName_(decodeBasicHtml_(value.replace(/<[^>]*>/g, ''))))
      .filter(Boolean)
      .forEach((value) => items.push(value));
  });

  return {
    date: ymd,
    mealType: selectedRows[0] && selectedRows[0].MMEAL_SC_NM ? String(selectedRows[0].MMEAL_SC_NM) : '중식',
    items: [...new Set(items)]
  };
}

function cleanDishName_(value) {
  return String(value || '')
    .replace(/\s*\([0-9.\s]+\)\s*$/, '')
    .replace(/\s+[0-9]+(?:\.[0-9]+)+\.?\s*$/, '')
    .replace(/^[-·•*★☆]+\s*/, '')
    .replace(/\s+/g, ' ')
    .trim();
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

function getConfiguredMealForm_() {
  const props = PropertiesService.getScriptProperties();
  const formId = props.getProperty('MEAL_FORM_ID');
  if (formId) return FormApp.openById(formId);

  const spreadsheetId = props.getProperty('SPREADSHEET_ID');
  if (!spreadsheetId) throw new Error('먼저 setupMealFormSync 함수를 실행해 주세요.');
  const ss = SpreadsheetApp.openById(spreadsheetId);
  const formUrl = ss.getFormUrl();
  if (!formUrl) throw new Error('연결된 Google Form을 찾지 못했습니다.');
  const form = FormApp.openByUrl(formUrl);
  props.setProperty('MEAL_FORM_ID', form.getId());
  return form;
}

function findMealQuestion_(form) {
  const keyword = normalizeSearchText_(MEAL_SYNC_CONFIG.FORM_QUESTION_KEYWORD);
  const item = form.getItems().find((candidate) => {
    const type = candidate.getType();
    return (type === FormApp.ItemType.MULTIPLE_CHOICE || type === FormApp.ItemType.LIST)
      && normalizeSearchText_(candidate.getTitle()).includes(keyword);
  });

  if (!item) throw new Error(`'${MEAL_SYNC_CONFIG.FORM_QUESTION_KEYWORD}' 문구가 포함된 객관식 질문을 찾지 못했습니다.`);
  return item.getType() === FormApp.ItemType.MULTIPLE_CHOICE
    ? item.asMultipleChoiceItem()
    : item.asListItem();
}

function normalizeSearchText_(value) {
  return String(value || '').toLowerCase().replace(/\s+/g, '').trim();
}

function getKoreanTodayYmd_() {
  return Utilities.formatDate(new Date(), MEAL_SYNC_CONFIG.TIME_ZONE, 'yyyyMMdd');
}

function getWeekdayFromYmd_(ymd) {
  validateYmd_(ymd);
  return new Date(Date.UTC(
    Number(ymd.slice(0, 4)),
    Number(ymd.slice(4, 6)) - 1,
    Number(ymd.slice(6, 8))
  )).getUTCDay();
}

function addDaysToYmd_(ymd, days) {
  validateYmd_(ymd);
  const date = new Date(Date.UTC(
    Number(ymd.slice(0, 4)),
    Number(ymd.slice(4, 6)) - 1,
    Number(ymd.slice(6, 8)) + Number(days || 0)
  ));
  return `${date.getUTCFullYear()}${String(date.getUTCMonth() + 1).padStart(2, '0')}${String(date.getUTCDate()).padStart(2, '0')}`;
}

function validateYmd_(ymd) {
  if (!/^\d{8}$/.test(String(ymd))) throw new Error(`날짜 형식 오류: ${ymd}`);
}

function formatKoreanDateLabel_(ymd) {
  validateYmd_(ymd);
  const date = new Date(
    Number(ymd.slice(0, 4)),
    Number(ymd.slice(4, 6)) - 1,
    Number(ymd.slice(6, 8)),
    12
  );
  return Utilities.formatDate(date, MEAL_SYNC_CONFIG.TIME_ZONE, 'M월 d일 E요일');
}
