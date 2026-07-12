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
  MEAL_PROXY_URL: 'https://bobbobbob.vercel.app/api/meal',
  NEIS_BASE_URL: 'https://open.neis.go.kr/hub/mealServiceDietInfo',
  REQUEST_RETRY_COUNT: 3,
  REQUEST_RETRY_DELAY_MS: 1200
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
    source: meal.source,
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
      MEAL_LAST_SYNC_SOURCE: meal.source || 'unknown',
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
        source: meal.source,
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
      source: meal.source,
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

/**
 * 급식 조회 순서
 * 1) 기존 Vercel 급식 API: 나이스 API 키/캐시 설정을 재사용
 * 2) 실패 시 나이스 원본 API 직접 호출
 */
function fetchMealForDate_(ymd) {
  validateYmd_(ymd);
  const errors = [];

  try {
    return fetchMealFromProxy_(ymd);
  } catch (error) {
    errors.push(`Vercel API: ${error.message}`);
    console.warn(errors[errors.length - 1]);
  }

  try {
    return fetchMealFromNeis_(ymd);
  } catch (error) {
    errors.push(`나이스 직접 호출: ${error.message}`);
    console.warn(errors[errors.length - 1]);
  }

  throw new Error(`급식 정보를 불러오지 못했습니다. ${errors.join(' / ')}`);
}

function fetchMealFromProxy_(ymd) {
  const year = Number(ymd.slice(0, 4));
  const month = Number(ymd.slice(4, 6));
  const isoDate = `${ymd.slice(0, 4)}-${ymd.slice(4, 6)}-${ymd.slice(6, 8)}`;
  const url = `${MEAL_SYNC_CONFIG.MEAL_PROXY_URL}?year=${year}&month=${month}`;
  const data = fetchJsonWithRetry_(url, 'Vercel 급식 API');

  if (!data || data.ok !== true) {
    throw new Error(data && data.message ? data.message : '정상 응답이 아닙니다.');
  }

  const entry = data.meals && data.meals[isoDate];
  if (!entry) {
    return {
      date: ymd,
      mealType: '중식',
      source: 'vercel-proxy',
      items: []
    };
  }

  const items = Array.isArray(entry.items)
    ? entry.items
        .map((item) => typeof item === 'string' ? item : item && item.name)
        .map(cleanDishName_)
        .filter(Boolean)
    : [];

  return {
    date: ymd,
    mealType: entry.mealType ? String(entry.mealType) : '중식',
    source: 'vercel-proxy',
    items: [...new Set(items)]
  };
}

function fetchMealFromNeis_(ymd) {
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

  const data = fetchJsonWithRetry_(
    `${MEAL_SYNC_CONFIG.NEIS_BASE_URL}?${query}`,
    '나이스 급식 API'
  );

  if (data && data.RESULT && data.RESULT.CODE === 'INFO-200') {
    return { date: ymd, mealType: '중식', source: 'open-neis', items: [] };
  }
  if (data && data.RESULT && data.RESULT.MESSAGE) {
    throw new Error(data.RESULT.MESSAGE);
  }

  const service = data && data.mealServiceDietInfo;
  const rowBlock = Array.isArray(service)
    ? service.find((block) => Array.isArray(block.row))
    : null;
  const rows = rowBlock ? rowBlock.row : [];
  const lunchRows = rows.filter((row) =>
    String(row.MMEAL_SC_CODE || '') === '2' ||
    String(row.MMEAL_SC_NM || '').includes('중식')
  );
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
    mealType: selectedRows[0] && selectedRows[0].MMEAL_SC_NM
      ? String(selectedRows[0].MMEAL_SC_NM)
      : '중식',
    source: 'open-neis',
    items: [...new Set(items)]
  };
}

function fetchJsonWithRetry_(url, sourceName) {
  let lastError = null;

  for (let attempt = 1; attempt <= MEAL_SYNC_CONFIG.REQUEST_RETRY_COUNT; attempt += 1) {
    try {
      const response = UrlFetchApp.fetch(url, {
        method: 'get',
        muteHttpExceptions: true,
        followRedirects: true,
        headers: {
          Accept: 'application/json',
          'Cache-Control': 'no-cache'
        }
      });

      const status = response.getResponseCode();
      const body = response.getContentText('UTF-8');

      if (status >= 200 && status < 300) {
        try {
          return JSON.parse(body);
        } catch (error) {
          throw new Error(`JSON 해석 실패: ${body.slice(0, 180)}`);
        }
      }

      lastError = new Error(`HTTP ${status}${body ? ` · ${body.slice(0, 180)}` : ''}`);
      const retryable = status === 429 || status >= 500;
      if (!retryable || attempt === MEAL_SYNC_CONFIG.REQUEST_RETRY_COUNT) break;
    } catch (error) {
      lastError = error;
      if (attempt === MEAL_SYNC_CONFIG.REQUEST_RETRY_COUNT) break;
    }

    Utilities.sleep(MEAL_SYNC_CONFIG.REQUEST_RETRY_DELAY_MS * attempt);
  }

  throw new Error(`${sourceName} 호출 실패: ${lastError ? lastError.message : '알 수 없는 오류'}`);
}

/** 두 데이터 원본을 각각 시험하고 실행 로그에 결과를 남깁니다. */
function diagnoseMealApi() {
  const context = getCurrentMealVoteContext_();
  const result = {
    targetDate: context.targetYmd,
    proxy: null,
    neis: null
  };

  try {
    result.proxy = fetchMealFromProxy_(context.targetYmd);
  } catch (error) {
    result.proxy = { ok: false, error: error.message };
  }

  try {
    result.neis = fetchMealFromNeis_(context.targetYmd);
  } catch (error) {
    result.neis = { ok: false, error: error.message };
  }

  console.log(JSON.stringify(result, null, 2));
  return result;
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
