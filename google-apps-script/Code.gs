const CONFIG = Object.freeze({
  TIME_ZONE: 'Asia/Seoul',
  TOP_N: 3,
  MIN_RESPONSES_TO_SHOW: 1,
  MENU_HEADER_KEYWORD: '기대되는 급식 메뉴',
  TIMESTAMP_HEADER_KEYWORD: '타임스탬프',
  FALLBACK_TIMESTAMP_COLUMN: 1,
  FALLBACK_MENU_COLUMN: 2
});

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('급식 차트')
    .addItem('초기 설정', 'setup')
    .addItem('오늘 데이터 테스트', 'testDashboardData')
    .addToUi();
}

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('응답 스프레드시트에서 Apps Script를 열어 주세요.');
  const sheet = findResponseSheet_(ss);
  PropertiesService.getScriptProperties().setProperties({
    SPREADSHEET_ID: ss.getId(),
    RESPONSE_SHEET_ID: String(sheet.getSheetId())
  }, false);
  ss.toast(`초기 설정 완료: ${sheet.getName()}`, '급식 차트', 5);
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('급식 기대 메뉴')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getDashboardData() {
  const context = getCurrentMealVoteContext_();
  const ss = getConfiguredSpreadsheet_();
  const sheet = getConfiguredResponseSheet_(ss);
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  const updatedAt = Utilities.formatDate(new Date(), CONFIG.TIME_ZONE, 'HH:mm:ss');

  if (lastRow < 2 || lastColumn < 2) {
    return emptyDashboard_(context, updatedAt, sheet.getName());
  }

  const headers = sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0];
  const timestampColumn = findColumnNumber_(headers, CONFIG.TIMESTAMP_HEADER_KEYWORD, CONFIG.FALLBACK_TIMESTAMP_COLUMN, 'timestamp');
  const menuColumn = findColumnNumber_(headers, CONFIG.MENU_HEADER_KEYWORD, CONFIG.FALLBACK_MENU_COLUMN, 'menu');
  const rows = sheet.getRange(2, 1, lastRow - 1, Math.max(timestampColumn, menuColumn)).getValues();

  const counts = Object.create(null);
  let totalResponses = 0;

  rows.forEach((row) => {
    const submittedAt = normalizeDate_(row[timestampColumn - 1]);
    if (!submittedAt) return;
    const submittedYmd = Utilities.formatDate(submittedAt, CONFIG.TIME_ZONE, 'yyyyMMdd');
    if (submittedYmd < context.windowStartYmd || submittedYmd > context.windowEndYmd) return;

    const menu = String(row[menuColumn - 1] || '').replace(/\s+/g, ' ').trim();
    if (!menu) return;
    counts[menu] = (counts[menu] || 0) + 1;
    totalResponses += 1;
  });

  const sorted = Object.entries(counts)
    .map(([menu, count]) => ({ menu, count }))
    .sort((a, b) => b.count - a.count || a.menu.localeCompare(b.menu, 'ko-KR'));

  let previousCount = null;
  let previousRank = 0;
  const ranked = sorted.map((item, index) => {
    const rank = item.count === previousCount ? previousRank : index + 1;
    previousCount = item.count;
    previousRank = rank;
    return {
      rank,
      menu: item.menu,
      count: item.count,
      percentage: totalResponses ? Math.round(item.count / totalResponses * 1000) / 10 : 0
    };
  });

  const cutoff = ranked[Math.min(CONFIG.TOP_N, ranked.length) - 1];
  const items = cutoff ? ranked.filter((item) => item.rank <= cutoff.rank) : [];

  return {
    success: true,
    dateLabel: context.dateLabel,
    mode: context.mode,
    periodLabel: context.windowStartYmd === context.windowEndYmd
      ? context.dateLabel
      : `${formatShortDate_(context.windowStartYmd)}~${formatShortDate_(context.windowEndYmd)}`,
    updatedAt,
    sheetName: sheet.getName(),
    totalResponses,
    minimumResponses: CONFIG.MIN_RESPONSES_TO_SHOW,
    hasEnoughResponses: totalResponses >= CONFIG.MIN_RESPONSES_TO_SHOW,
    items
  };
}

function testDashboardData() {
  const result = getDashboardData();
  console.log(JSON.stringify(result, null, 2));
  return result;
}

function getConfiguredSpreadsheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) throw new Error('먼저 setup 함수를 실행해 주세요.');
  return SpreadsheetApp.openById(id);
}

function getConfiguredResponseSheet_(ss) {
  const savedId = Number(PropertiesService.getScriptProperties().getProperty('RESPONSE_SHEET_ID'));
  const saved = ss.getSheets().find((sheet) => sheet.getSheetId() === savedId);
  return saved || findResponseSheet_(ss);
}

function findResponseSheet_(ss) {
  return ss.getSheets().find((sheet) => {
    if (sheet.getLastColumn() < 2) return false;
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0].map(normalizeHeader_);
    return headers.some((h) => h.includes('타임스탬프') || h.includes('timestamp'))
      && headers.some((h) => h.includes('기대되는급식메뉴') || h.includes('급식메뉴'));
  }) || ss.getSheets()[0];
}

function findColumnNumber_(headers, keyword, fallback, type) {
  const normalized = headers.map(normalizeHeader_);
  let index = normalized.findIndex((header) => header.includes(normalizeHeader_(keyword)));
  if (index < 0 && type === 'timestamp') index = normalized.findIndex((header) => header.includes('타임스탬프') || header.includes('timestamp'));
  if (index < 0 && type === 'menu') index = normalized.findIndex((header) => header.includes('급식메뉴') || header.includes('메뉴'));
  return index >= 0 ? index + 1 : fallback;
}

function normalizeHeader_(value) {
  return String(value || '').toLowerCase().replace(/\s+/g, '').trim();
}

function normalizeDate_(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value;
  const parsed = new Date(value);
  return Number.isNaN(parsed.getTime()) ? null : parsed;
}

function formatShortDate_(ymd) {
  return `${Number(ymd.slice(4, 6))}/${Number(ymd.slice(6, 8))}`;
}

function emptyDashboard_(context, updatedAt, sheetName) {
  return {
    success: true,
    dateLabel: context.dateLabel,
    mode: context.mode,
    periodLabel: context.windowStartYmd === context.windowEndYmd
      ? context.dateLabel
      : `${formatShortDate_(context.windowStartYmd)}~${formatShortDate_(context.windowEndYmd)}`,
    updatedAt,
    sheetName,
    totalResponses: 0,
    minimumResponses: CONFIG.MIN_RESPONSES_TO_SHOW,
    hasEnoughResponses: false,
    items: []
  };
}
