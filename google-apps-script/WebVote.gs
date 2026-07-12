const EMBEDDED_VOTE_CONFIG = Object.freeze({
  TIME_ZONE: 'Asia/Seoul',
  DAILY_SWITCH_HOUR: 5,
  TRIGGER_HANDLER: 'scheduledMealFormSync'
});

/**
 * 최초 1회 실행합니다.
 * - 기존 급식 동기화 트리거를 제거
 * - 매일 오전 5시 전후 트리거 설치
 * - 현재 투표 대상 메뉴 즉시 동기화
 */
function setupEmbeddedMealVote() {
  const form = getConfiguredMealForm_();
  const question = findMealQuestion_(form);
  const triggerId = installFiveAmMealTrigger();
  const syncResult = syncTodayMealToForm();

  const result = {
    ok: true,
    triggerId,
    triggerTime: '매일 오전 5시 전후',
    formTitle: form.getTitle(),
    questionTitle: question.getTitle(),
    sync: syncResult
  };

  console.log(JSON.stringify(result, null, 2));
  return result;
}

/**
 * Apps Script 시간 트리거는 nearMinute(0) 기준 ±15분 범위에서 실행될 수 있습니다.
 * 페이지는 오전 5시가 지났는데 자동 실행이 늦은 경우 최초 방문 시 한 번 보정합니다.
 */
function installFiveAmMealTrigger() {
  ScriptApp.getProjectTriggers()
    .filter((trigger) => trigger.getHandlerFunction() === EMBEDDED_VOTE_CONFIG.TRIGGER_HANDLER)
    .forEach((trigger) => ScriptApp.deleteTrigger(trigger));

  return ScriptApp.newTrigger(EMBEDDED_VOTE_CONFIG.TRIGGER_HANDLER)
    .timeBased()
    .everyDays(1)
    .atHour(EMBEDDED_VOTE_CONFIG.DAILY_SWITCH_HOUR)
    .nearMinute(0)
    .inTimezone(EMBEDDED_VOTE_CONFIG.TIME_ZONE)
    .create()
    .getUniqueId();
}

/**
 * 차트에 표시할 투표 기간을 결정합니다.
 * 00:00~04:59에는 이전 투표 결과를 유지하고, 오전 5시부터 새 날짜로 전환합니다.
 */
function getDashboardContext_() {
  const scheduledContext = getCurrentMealVoteContext_();
  const hour = getKoreanCurrentHour_();
  const storedContext = getStoredMealVoteContext_();

  if (
    hour < EMBEDDED_VOTE_CONFIG.DAILY_SWITCH_HOUR &&
    storedContext &&
    storedContext.targetYmd !== scheduledContext.targetYmd
  ) {
    return storedContext;
  }

  return scheduledContext;
}

/**
 * HTML 드롭다운에 표시할 현재 투표 상태와 선택지를 반환합니다.
 */
function getVotePanelData() {
  const scheduledContext = getCurrentMealVoteContext_();
  const hour = getKoreanCurrentHour_();
  const props = PropertiesService.getScriptProperties();
  let storedTarget = props.getProperty('MEAL_VOTE_TARGET_DATE');

  // 새 날짜가 되었지만 아직 오전 5시 전이면 이전 메뉴로 잘못 투표하지 않도록 잠시 닫습니다.
  if (
    hour < EMBEDDED_VOTE_CONFIG.DAILY_SWITCH_HOUR &&
    storedTarget !== scheduledContext.targetYmd
  ) {
    return {
      ok: true,
      available: false,
      date: scheduledContext.targetYmd,
      dateLabel: scheduledContext.dateLabel,
      mode: scheduledContext.mode,
      options: [],
      message: '오전 5시에 오늘의 급식 투표가 열립니다.'
    };
  }

  // 오전 5시가 지났는데 트리거가 아직 실행되지 않은 경우 최초 방문이 자동 보정합니다.
  if (storedTarget !== scheduledContext.targetYmd) {
    syncTodayMealToForm();
    storedTarget = PropertiesService.getScriptProperties().getProperty('MEAL_VOTE_TARGET_DATE');
  }

  const form = getConfiguredMealForm_();
  const question = findMealQuestion_(form);
  const options = question.getChoices().map((choice) => choice.getValue()).filter(Boolean);
  const available = form.isAcceptingResponses() && options.length > 0 && storedTarget === scheduledContext.targetYmd;

  return {
    ok: true,
    available,
    date: scheduledContext.targetYmd,
    dateLabel: scheduledContext.dateLabel,
    mode: scheduledContext.mode,
    options: available ? options : [],
    message: available
      ? (scheduledContext.mode === 'WEEKEND_MONDAY'
          ? '돌아오는 월요일 급식 중 가장 기대되는 메뉴를 골라 주세요.'
          : '오늘 급식 중 가장 기대되는 메뉴를 골라 주세요.')
      : `${scheduledContext.dateLabel} 급식 투표를 이용할 수 없습니다.`
  };
}

/**
 * HTML 우측 상단 '동기화' 버튼에서 호출합니다.
 * 급식 API 재조회 → Google Form 선택지 변경 → 드롭다운/차트 데이터 반환을 한 번에 처리합니다.
 */
function manualSyncMealWidget() {
  const sync = syncTodayMealToForm();

  return {
    ok: true,
    syncedAt: Utilities.formatDate(new Date(), EMBEDDED_VOTE_CONFIG.TIME_ZONE, 'HH:mm:ss'),
    sync,
    vote: getVotePanelData(),
    dashboard: getDashboardData()
  };
}

/**
 * HTML 페이지에서 선택한 값을 실제 Google Form 응답으로 제출합니다.
 */
function submitMealVote(payload) {
  if (!payload || typeof payload !== 'object') {
    throw new Error('제출 데이터가 올바르지 않습니다.');
  }

  const menu = String(payload.menu || '').replace(/\s+/g, ' ').trim();
  const targetDate = String(payload.targetDate || '').trim();
  if (!menu) throw new Error('메뉴를 하나 선택해 주세요.');

  const panel = getVotePanelData();
  if (!panel.available) throw new Error(panel.message || '현재 투표를 받을 수 없습니다.');
  if (targetDate !== panel.date) {
    throw new Error('투표 날짜가 변경되었습니다. 화면을 갱신한 뒤 다시 선택해 주세요.');
  }
  if (!panel.options.includes(menu)) {
    throw new Error('현재 급식 메뉴에 없는 선택지입니다. 화면을 갱신해 주세요.');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const form = getConfiguredMealForm_();
    if (!form.isAcceptingResponses()) {
      throw new Error('현재 설문 응답이 마감되어 있습니다.');
    }

    const question = findMealQuestion_(form);
    const validOptions = question.getChoices().map((choice) => choice.getValue());
    if (!validOptions.includes(menu)) {
      throw new Error('선택지가 방금 변경되었습니다. 다시 선택해 주세요.');
    }

    const submitted = form
      .createResponse()
      .withItemResponse(question.createResponse(menu))
      .submit();

    return {
      ok: true,
      responseId: submitted.getId(),
      menu,
      date: panel.date,
      dateLabel: panel.dateLabel,
      submittedAt: Utilities.formatDate(new Date(), EMBEDDED_VOTE_CONFIG.TIME_ZONE, 'HH:mm:ss'),
      message: `${menu}에 투표했어요!`
    };
  } finally {
    lock.releaseLock();
  }
}

function getStoredMealVoteContext_() {
  const props = PropertiesService.getScriptProperties();
  const targetYmd = props.getProperty('MEAL_VOTE_TARGET_DATE');
  const windowStartYmd = props.getProperty('MEAL_VOTE_WINDOW_START');
  const windowEndYmd = props.getProperty('MEAL_VOTE_WINDOW_END');
  if (!targetYmd || !windowStartYmd || !windowEndYmd) return null;

  return {
    todayYmd: getKoreanTodayYmd_(),
    targetYmd,
    dateLabel: formatKoreanDateLabel_(targetYmd),
    mode: props.getProperty('MEAL_VOTE_MODE') || 'STORED',
    windowStartYmd,
    windowEndYmd
  };
}

function getKoreanCurrentHour_() {
  return Number(Utilities.formatDate(new Date(), EMBEDDED_VOTE_CONFIG.TIME_ZONE, 'H'));
}
