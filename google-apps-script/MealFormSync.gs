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
 * - 응답 스프레드시트에 연결된 Google Form 자동 탐색
 * - 급식 질문 확인
 * - 매일 오전 자동 실행 트리거 생성
 * - 평일에는 오늘 급식, 주말에는 돌아오는 월요일 급식 반영
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
      ? `${syncResult.dateLabel} 급식 ${syncResult.items.length}개를 설문 선택지로 반영했습니다.`
      : `${syncResult.dateLabel} 급식이 없어 설문 응답을 닫았습니다.`,
    '급식 설문 자동화 설정 완료',
    7
  );
  console.log(JSON.stringify(result, null, 2));
  return result;
}

function scheduledMealFormSync() {
  return syncTodayMealToForm();
}

function syncTodayMealToForm() {
  const context = getCurrentMealVoteContext_();
  return syncMealToFormForContext_(context);
}

function previewTodayMealChoices() {
  const context = getCurrentMealVoteContext_();
  const meal = fetchMealForDate_(context.targetYmd);
  const result = {
    ok: true,
    today: context.todayYmd,
    date: context.targetYmd,
    dateLabel: context.dateLabel,
    mode