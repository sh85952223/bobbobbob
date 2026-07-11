const SCHOOL_NAME = '율량중학교';
const REGION_CODE = 'M10'; // 충청북도교육청
const OPEN_NEIS_BASE = 'https://open.neis.go.kr/hub';
const REQUEST_TIMEOUT_MS = 7000;

let schoolInfoCache = null;

module.exports = async function handler(req, res) {
  if (req.method !== 'GET') {
    res.setHeader('Allow', 'GET');
    return res.status(405).json({ ok: false, message: 'GET 요청만 지원합니다.' });
  }

  const today = getKoreanDateParts();
  const year = toInteger(req.query.year, today.year);
  const month = toInteger(req.query.month, today.month);

  if (year < 2000 || year > 2100 || month < 1 || month > 12) {
    return res.status(400).json({ ok: false, message: '조회 연월을 확인해 주세요.' });
  }

  try {
    const result = await fetchFromOpenNeis(year, month);
    const source = 'open-neis';

    res.setHeader('Cache-Control', 'public, s-maxage=1800, stale-while-revalidate=86400');
    return res.status(200).json({
      ok: true,
      school: result.school,
      year,
      month,
      source,
      meals: result.meals,
      fetchedAt: new Date().toISOString()
    });
  } catch (error) {
    console.error('[meal-api]', error);
    return res.status(502).json({
      ok: false,
      message: '급식 정보를 불러오지 못했습니다. 잠시 후 다시 확인해 주세요.',
      detail: process.env.NODE_ENV === 'development' ? error.message : undefined
    });
  }
};

async function fetchFromOpenNeis(year, month) {
  const school = await getSchoolInfoFromOpenNeis();
  const lastDay = new Date(Date.UTC(year, month, 0)).getUTCDate();
  const from = `${year}${pad(month)}01`;
  const to = `${year}${pad(month)}${pad(lastDay)}`;

  const url = makeOpenNeisUrl('mealServiceDietInfo', {
    ATPT_OFCDC_SC_CODE: school.officeCode,
    SD_SCHUL_CODE: school.schoolCode,
    MLSV_FROM_YMD: from,
    MLSV_TO_YMD: to,
    pIndex: 1,
    pSize: 100
  });

  const data = await fetchJson(url);
  const rows = extractRows(data, 'mealServiceDietInfo');
  const meals = {};

  for (const row of rows) {
    const date = formatCompactDate(row.MLSV_YMD);
    if (!date) continue;

    const mealType = row.MMEAL_SC_NM || '중식';
    const items = splitMealItems(row.DDISH_NM || '');
    if (!items.length) continue;

    if (!meals[date]) {
      meals[date] = {
        date,
        mealType,
        items,
        calories: normalizeText(row.CAL_INFO || ''),
        nutrition: normalizeMultiline(row.NTR_INFO || ''),
        origin: normalizeMultiline(row.ORPLC_INFO || '')
      };
    } else {
      meals[date].items.push(...items);
    }
  }

  return {
    school: {
      name: school.name,
      address: school.address,
      officeCode: school.officeCode,
      schoolCode: school.schoolCode
    },
    meals
  };
}

async function getSchoolInfoFromOpenNeis() {
  if (schoolInfoCache) return schoolInfoCache;

  const url = makeOpenNeisUrl('schoolInfo', {
    ATPT_OFCDC_SC_CODE: REGION_CODE,
    SCHUL_NM: SCHOOL_NAME,
    pIndex: 1,
    pSize: 20
  });
  const data = await fetchJson(url);
  const rows = extractRows(data, 'schoolInfo');

  const exact = rows.find(
    (row) => row.SCHUL_NM === SCHOOL_NAME && row.ATPT_OFCDC_SC_CODE === REGION_CODE
  );
  const selected = exact || rows.find((row) => row.SCHUL_NM === SCHOOL_NAME) || rows[0];

  if (!selected) throw new Error(`${SCHOOL_NAME} 학교 정보를 찾지 못했습니다.`);

  schoolInfoCache = {
    name: selected.SCHUL_NM,
    address: selected.ORG_RDNMA || selected.ORG_RDNDA || '',
    officeCode: selected.ATPT_OFCDC_SC_CODE,
    schoolCode: selected.SD_SCHUL_CODE
  };
  return schoolInfoCache;
}

function makeOpenNeisUrl(service, params) {
  const url = new URL(`${OPEN_NEIS_BASE}/${service}`);
  const common = {
    Type: 'json',
    pIndex: 1,
    pSize: 100,
    ...params
  };

  if (process.env.NEIS_API_KEY) common.KEY = process.env.NEIS_API_KEY;

  Object.entries(common).forEach(([key, value]) => {
    if (value !== undefined && value !== null && value !== '') {
      url.searchParams.set(key, String(value));
    }
  });
  return url.toString();
}

async function fetchJson(url) {
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), REQUEST_TIMEOUT_MS);

  try {
    const response = await fetch(url, {
      signal: controller.signal,
      headers: { 'User-Agent': 'YullyangMealWidget/1.0' }
    });
    if (!response.ok) throw new Error(`나이스 API HTTP ${response.status}`);
    return await response.json();
  } finally {
    clearTimeout(timeout);
  }
}

function extractRows(data, serviceName) {
  if (Array.isArray(data?.[serviceName])) {
    const rowBlock = data[serviceName].find((block) => Array.isArray(block.row));
    return rowBlock?.row || [];
  }

  const code = data?.RESULT?.CODE;
  if (code === 'INFO-200') return [];
  if (data?.RESULT?.MESSAGE) throw new Error(data.RESULT.MESSAGE);
  throw new Error(`${serviceName} 응답 형식을 확인할 수 없습니다.`);
}

function splitMealItems(html) {
  return String(html)
    .split(/<br\s*\/?\s*>/i)
    .map((value) => decodeHtml(stripTags(value)).trim())
    .filter(Boolean)
    .map(parseMealItem)
    .filter((item) => item.name);
}

function parseMealItem(value) {
  const original = normalizeText(value);
  let name = original;
  let allergens = [];

  const parenthesized = name.match(/\(([0-9.\s]+)\)\s*$/);
  if (parenthesized) {
    allergens = parseAllergenNumbers(parenthesized[1]);
    name = name.slice(0, parenthesized.index).trim();
  } else {
    const attached = name.match(/([0-9]+(?:\.[0-9]+){1,}\.?)\s*$/);
    if (attached) {
      allergens = parseAllergenNumbers(attached[1]);
      name = name.slice(0, attached.index).trim();
    }
  }

  name = name.replace(/^[-·•]\s*/, '').replace(/\s+/g, ' ').trim();
  return { name, allergens, raw: original };
}

function parseAllergenNumbers(value) {
  return [...new Set(String(value).split('.').map((v) => v.trim()).filter(Boolean))];
}

function normalizeMultiline(value) {
  return decodeHtml(stripTags(String(value).replace(/<br\s*\/?\s*>/gi, '\n')))
    .replace(/\r/g, '')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

function normalizeText(value) {
  return decodeHtml(stripTags(String(value))).replace(/\s+/g, ' ').trim();
}

function stripTags(value) {
  return String(value).replace(/<[^>]*>/g, '');
}

function decodeHtml(value) {
  return String(value)
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&nbsp;/g, ' ');
}

function formatCompactDate(value) {
  const text = String(value || '').replace(/\D/g, '');
  if (text.length !== 8) return null;
  return `${text.slice(0, 4)}-${text.slice(4, 6)}-${text.slice(6, 8)}`;
}

function getKoreanDateParts() {
  const formatter = new Intl.DateTimeFormat('en-US', {
    timeZone: 'Asia/Seoul',
    year: 'numeric',
    month: 'numeric',
    day: 'numeric'
  });
  const parts = Object.fromEntries(
    formatter.formatToParts(new Date()).map((part) => [part.type, part.value])
  );
  return { year: Number(parts.year), month: Number(parts.month), day: Number(parts.day) };
}

function toInteger(value, fallback) {
  const parsed = Number.parseInt(String(value ?? ''), 10);
  return Number.isFinite(parsed) ? parsed : fallback;
}

function pad(value) {
  return String(value).padStart(2, '0');
}
