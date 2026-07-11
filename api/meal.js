const SCHOOL = {
  name: '율량중학교',
  officeCode: 'M10', // 충청북도교육청
  schoolCode: '8011092'
};

const OPEN_NEIS_BASE = 'https://open.neis.go.kr/hub';
const REQUEST_TIMEOUT_MS = 9000;

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
    const meals = await fetchMonthByDay(year, month);

    res.setHeader('Cache-Control', 'public, s-maxage=1800, stale-while-revalidate=86400');
    return res.status(200).json({
      ok: true,
      school: SCHOOL,
      year,
      month,
      source: 'open-neis',
      meals,
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

async function fetchMonthByDay(year, month) {
  const lastDay = new Date(Date.UTC(year, month, 0)).getUTCDate();
  const dates = [];

  for (let day = 1; day <= lastDay; day += 1) {
    const date = new Date(Date.UTC(year, month - 1, day));
    const weekday = date.getUTCDay();
    if (weekday !== 0 && weekday !== 6) dates.push(day);
  }

  const results = await runWithConcurrency(dates, 6, async (day) => {
    const ymd = `${year}${pad(month)}${pad(day)}`;
    const url = makeOpenNeisUrl('mealServiceDietInfo', {
      ATPT_OFCDC_SC_CODE: SCHOOL.officeCode,
      SD_SCHUL_CODE: SCHOOL.schoolCode,
      MLSV_YMD: ymd,
      pIndex: 1,
      pSize: 10
    });

    const data = await fetchJson(url);
    const rows = extractRows(data, 'mealServiceDietInfo');
    return rows.map(parseMealRow).filter(Boolean);
  });

  const meals = {};
  for (const group of results) {
    for (const meal of group) {
      if (!meals[meal.date]) {
        meals[meal.date] = meal;
      } else {
        meals[meal.date].items.push(...meal.items);
      }
    }
  }
  return meals;
}

function parseMealRow(row) {
  const date = formatCompactDate(row.MLSV_YMD);
  const items = splitMealItems(row.DDISH_NM || '');
  if (!date || !items.length) return null;

  return {
    date,
    mealType: row.MMEAL_SC_NM || '중식',
    items,
    calories: normalizeText(row.CAL_INFO || ''),
    nutrition: normalizeMultiline(row.NTR_INFO || ''),
    origin: normalizeMultiline(row.ORPLC_INFO || '')
  };
}

function makeOpenNeisUrl(service, params) {
  const url = new URL(`${OPEN_NEIS_BASE}/${service}`);
  const common = {
    Type: 'json',
    pIndex: 1,
    pSize: 10,
    ...params
  };

  if (process.env.NEIS_API_KEY) common.KEY = process.env.NEIS_API_KEY;

  for (const [key, value] of Object.entries(common)) {
    if (value !== undefined && value !== null && value !== '') {
      url.searchParams.set(key, String(value));
    }
  }
  return url.toString();
}

async function fetchJson(url) {
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), REQUEST_TIMEOUT_MS);

  try {
    const response = await fetch(url, {
      signal: controller.signal,
      headers: { 'User-Agent': 'YullyangMealWidget/1.1' }
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
  return [];
}

async function runWithConcurrency(items, limit, worker) {
  const output = new Array(items.length);
  let cursor = 0;

  async function runner() {
    while (cursor < items.length) {
      const index = cursor;
      cursor += 1;
      output[index] = await worker(items[index]);
    }
  }

  await Promise.all(Array.from({ length: Math.min(limit, items.length) }, runner));
  return output;
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
