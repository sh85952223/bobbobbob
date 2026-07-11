const state = {
  selectedDate: getKoreanToday(),
  today: getKoreanToday(),
  monthlyCache: new Map(),
  loading: false
};

const els = {
  previousDay: document.querySelector('#previousDay'),
  nextDay: document.querySelector('#nextDay'),
  dateDisplay: document.querySelector('#dateDisplay'),
  datePicker: document.querySelector('#datePicker'),
  dateHeadline: document.querySelector('#dateHeadline'),
  dateSubline: document.querySelector('#dateSubline'),
  weekStrip: document.querySelector('#weekStrip'),
  refreshButton: document.querySelector('#refreshButton'),
  loadingState: document.querySelector('#loadingState'),
  mealState: document.querySelector('#mealState'),
  emptyState: document.querySelector('#emptyState'),
  errorState: document.querySelector('#errorState'),
  mealType: document.querySelector('#mealType'),
  menuList: document.querySelector('#menuList'),
  calorieChip: document.querySelector('#calorieChip'),
  emptyMessage: document.querySelector('#emptyMessage'),
  errorMessage: document.querySelector('#errorMessage'),
  goTodayButton: document.querySelector('#goTodayButton'),
  retryButton: document.querySelector('#retryButton'),
  todayButton: document.querySelector('#todayButton')
};

initialize();

function initialize() {
  els.datePicker.value = state.selectedDate;
  bindEvents();
  loadSelectedMonth();
}

function bindEvents() {
  els.previousDay.addEventListener('click', () => changeDateBy(-1));
  els.nextDay.addEventListener('click', () => changeDateBy(1));
  els.dateDisplay.addEventListener('click', openDatePicker);
  els.datePicker.addEventListener('change', () => {
    if (els.datePicker.value) selectDate(els.datePicker.value);
  });
  els.refreshButton.addEventListener('click', () => loadSelectedMonth(true));
  els.goTodayButton.addEventListener('click', goToday);
  els.todayButton.addEventListener('click', goToday);
  els.retryButton.addEventListener('click', () => loadSelectedMonth(true));
}

async function loadSelectedMonth(force = false) {
  const { year, month } = parseDate(state.selectedDate);
  const cacheKey = `${year}-${pad(month)}`;

  if (!force && state.monthlyCache.has(cacheKey)) {
    render(state.monthlyCache.get(cacheKey));
    return;
  }

  setView('loading');
  state.loading = true;
  els.refreshButton.classList.add('is-spinning');

  try {
    const response = await fetch(`/api/meal?year=${year}&month=${month}`, {
      headers: force ? { 'Cache-Control': 'no-cache' } : {}
    });
    const data = await response.json();
    if (!response.ok || !data.ok) throw new Error(data.message || '급식 조회에 실패했습니다.');

    state.monthlyCache.set(cacheKey, data);
    render(data);
  } catch (error) {
    renderDateHeader();
    renderWeekStrip({ meals: {} });
    els.errorMessage.textContent = error.message || '네트워크 상태를 확인한 뒤 다시 시도해 주세요.';
    setView('error');
  } finally {
    state.loading = false;
    els.refreshButton.classList.remove('is-spinning');
  }
}

function render(data) {
  renderDateHeader();
  renderWeekStrip(data);

  const meal = data.meals?.[state.selectedDate];
  if (!meal || !Array.isArray(meal.items) || meal.items.length === 0) {
    const selected = fromDateKey(state.selectedDate);
    const weekday = selected.getUTCDay();
    els.emptyMessage.textContent = weekday === 0 || weekday === 6
      ? '주말에는 학교 급식이 운영되지 않아요. 다른 평일을 선택해 보세요.'
      : '방학이나 학교 일정에 따라 급식이 없을 수 있어요.';
    setView('empty');
    return;
  }

  els.mealType.textContent = meal.mealType || '중식';
  els.menuList.replaceChildren(...meal.items.map(createMenuItem));

  if (meal.calories) {
    els.calorieChip.textContent = meal.calories;
    els.calorieChip.classList.remove('hidden');
  } else {
    els.calorieChip.classList.add('hidden');
  }

  setView('meal');
}

function renderDateHeader() {
  const date = fromDateKey(state.selectedDate);
  const isToday = state.selectedDate === state.today;
  els.dateHeadline.textContent = new Intl.DateTimeFormat('ko-KR', {
    month: 'long', day: 'numeric', weekday: 'long', timeZone: 'UTC'
  }).format(date);
  els.dateSubline.textContent = isToday
    ? '오늘의 급식이에요'
    : new Intl.DateTimeFormat('ko-KR', { year: 'numeric', timeZone: 'UTC' }).format(date);
  els.datePicker.value = state.selectedDate;
}

function renderWeekStrip(data) {
  const selected = fromDateKey(state.selectedDate);
  const start = addDays(selected, -selected.getUTCDay());
  const weekdays = ['일', '월', '화', '수', '목', '금', '토'];
  const buttons = [];

  for (let offset = 0; offset < 7; offset += 1) {
    const date = addDays(start, offset);
    const dateKey = toDateKey(date);
    const button = document.createElement('button');
    button.type = 'button';
    button.className = 'day-button';
    button.setAttribute('role', 'listitem');
    button.setAttribute('aria-label', `${formatAccessibleDate(date)} 급식 보기`);
    button.innerHTML = `<span class="weekday">${weekdays[offset]}</span><span class="day-number">${date.getUTCDate()}</span>`;

    if (dateKey === state.selectedDate) {
      button.classList.add('is-selected');
      button.setAttribute('aria-current', 'date');
    }
    if (data.meals?.[dateKey]) button.classList.add('has-meal');
    if (offset === 0) button.classList.add('is-sunday');
    if (offset === 6) button.classList.add('is-saturday');

    button.addEventListener('click', () => selectDate(dateKey));
    buttons.push(button);
  }

  els.weekStrip.replaceChildren(...buttons);
}

function createMenuItem(item, index) {
  const li = document.createElement('li');
  li.className = 'menu-item';
  li.style.animationDelay = `${Math.min(index * 55, 330)}ms`;

  const icon = document.createElement('span');
  icon.className = 'menu-icon';
  icon.setAttribute('aria-hidden', 'true');
  icon.textContent = chooseIcon(item.name);

  const name = document.createElement('span');
  name.className = 'menu-name';
  name.textContent = item.name;

  const allergens = document.createElement('span');
  allergens.className = 'allergen-badges';
  (item.allergens || []).slice(0, 8).forEach((number) => {
    const badge = document.createElement('span');
    badge.className = 'allergen-badge';
    badge.textContent = number;
    badge.title = `알레르기 번호 ${number}`;
    allergens.appendChild(badge);
  });

  li.append(icon, name, allergens);
  return li;
}

function chooseIcon(name = '') {
  const value = name.toLowerCase();
  if (/밥|라이스|볶음밥|덮밥|비빔밥/.test(value)) return '🍚';
  if (/국|탕|찌개|스프/.test(value)) return '🥣';
  if (/면|국수|우동|파스타|라면|쫄면/.test(value)) return '🍜';
  if (/김치|깍두기|겉절이/.test(value)) return '🥬';
  if (/닭|치킨|오리/.test(value)) return '🍗';
  if (/돈|돼지|제육|함박|갈비|불고기|소고기|쇠고기/.test(value)) return '🥩';
  if (/생선|고등어|갈치|연어|새우|오징어|어묵/.test(value)) return '🐟';
  if (/샐러드|나물|무침|야채|채소/.test(value)) return '🥗';
  if (/과일|귤|사과|배|포도|수박|바나나/.test(value)) return '🍎';
  if (/우유|요거트|요구르트|주스|음료/.test(value)) return '🥛';
  if (/빵|케이크|쿠키|도넛|핫도그|피자/.test(value)) return '🥐';
  return '✨';
}

function setView(view) {
  els.loadingState.classList.toggle('hidden', view !== 'loading');
  els.mealState.classList.toggle('hidden', view !== 'meal');
  els.emptyState.classList.toggle('hidden', view !== 'empty');
  els.errorState.classList.toggle('hidden', view !== 'error');
}

function changeDateBy(amount) {
  selectDate(toDateKey(addDays(fromDateKey(state.selectedDate), amount)));
}

function selectDate(dateKey) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(dateKey)) return;
  const previousMonth = state.selectedDate.slice(0, 7);
  state.selectedDate = dateKey;
  const nextMonth = state.selectedDate.slice(0, 7);

  if (previousMonth === nextMonth) {
    const cached = state.monthlyCache.get(nextMonth);
    if (cached) render(cached);
    else loadSelectedMonth();
  } else {
    loadSelectedMonth();
  }
}

function goToday() { selectDate(state.today); }

function openDatePicker() {
  if (typeof els.datePicker.showPicker === 'function') {
    els.datePicker.showPicker();
  } else {
    els.datePicker.click();
  }
}

function getKoreanToday() {
  const parts = new Intl.DateTimeFormat('en-US', {
    timeZone: 'Asia/Seoul', year: 'numeric', month: '2-digit', day: '2-digit'
  }).formatToParts(new Date());
  const map = Object.fromEntries(parts.map((part) => [part.type, part.value]));
  return `${map.year}-${map.month}-${map.day}`;
}

function parseDate(dateKey) {
  const [year, month, day] = dateKey.split('-').map(Number);
  return { year, month, day };
}

function fromDateKey(dateKey) {
  const { year, month, day } = parseDate(dateKey);
  return new Date(Date.UTC(year, month - 1, day));
}

function toDateKey(date) {
  return `${date.getUTCFullYear()}-${pad(date.getUTCMonth() + 1)}-${pad(date.getUTCDate())}`;
}

function addDays(date, amount) {
  const next = new Date(date.getTime());
  next.setUTCDate(next.getUTCDate() + amount);
  return next;
}

function formatAccessibleDate(date) {
  return new Intl.DateTimeFormat('ko-KR', {
    year: 'numeric', month: 'long', day: 'numeric', weekday: 'long', timeZone: 'UTC'
  }).format(date);
}

function pad(value) { return String(value).padStart(2, '0'); }
