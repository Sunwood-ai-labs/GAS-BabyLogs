/** ===== ログ出力ユーティリティ ===== */
function logInfo(msg) {
  Logger.log(msg);
  try {
    console.log(msg);
  } catch (e) {
    // Apps Script のコンソールが利用できない場合でも無視
  }
}

function logWarn(msg) {
  Logger.log('[WARN] ' + msg);
  try {
    console.warn(msg);
  } catch (e) {}
}

function logError(msg) {
  Logger.log('[ERROR] ' + msg);
  try {
    console.error(msg);
  } catch (e) {}
}

/** ===== 日付ユーティリティ ===== */
function shiftDate_(base, days) {
  const d = new Date(base);
  d.setDate(d.getDate() + days);
  d.setHours(0, 0, 0, 0);
  return d;
}

function fmt(dt, tz) {
  return Utilities.formatDate(dt, tz, 'yyyy-MM-dd HH:mm');
}

function getScriptTimeZone_() {
  if (typeof SETTINGS !== 'undefined' && SETTINGS && SETTINGS.TIMEZONE) {
    return SETTINGS.TIMEZONE;
  }
  if (typeof Session !== 'undefined' && Session.getScriptTimeZone) {
    try {
      const tz = Session.getScriptTimeZone();
      if (tz) {
        return tz;
      }
    } catch (e) {}
  }
  return 'UTC';
}

function formatMonthDay(value, timeZone) {
  if (value === null || value === undefined || value === '') {
    return '';
  }

  let date;
  const isDateObject = value instanceof Date || Object.prototype.toString.call(value) === '[object Date]';

  if (isDateObject) {
    const time = typeof value.getTime === 'function' ? value.getTime() : Date.parse(value);
    if (!Number.isNaN(time)) {
      date = new Date(time);
    }
  } else if (typeof value === 'number') {
    date = new Date(value);
  } else if (typeof value === 'string') {
    const normalized = value.includes('T') || value.includes('/') ? value : value.replace(/-/g, '/');
    date = new Date(normalized);
    if (Number.isNaN(date.getTime())) {
      const parts = value.split('-');
      if (parts.length >= 3) {
        const year = Number(parts[0]);
        const month = Number(parts[1]);
        const day = Number(parts[2]);
        if (![year, month, day].some(n => Number.isNaN(n))) {
          date = new Date(year, month - 1, day);
        }
      }
    }
  }

  if (!(date instanceof Date) || Number.isNaN(date.getTime())) {
    return value;
  }

  const tz = timeZone || getScriptTimeZone_();
  if (typeof Utilities !== 'undefined' && Utilities.formatDate) {
    return Utilities.formatDate(date, tz, 'M/d');
  }

  if (typeof Intl !== 'undefined' && Intl.DateTimeFormat) {
    return new Intl.DateTimeFormat('en-US', {
      month: 'numeric',
      day: 'numeric',
      timeZone: tz,
    }).format(date);
  }

  const month = date.getUTCMonth() + 1;
  const day = date.getUTCDate();
  return month + '/' + day;
}

/** ===== テキストユーティリティ ===== */
function normalizeText_(s) {
  let value = (s || '').trim();
  try {
    value = value.replace(/[Ａ-Ｚａ-ｚ０-９]/g, ch => String.fromCharCode(ch.charCodeAt(0) - 0xFEE0));
  } catch (e) {}
  return value.toLowerCase();
}

/** ===== カレンダーユーティリティ ===== */
function normalizeCalendarId(raw) {
  if (!raw) return null;
  let s = String(raw).trim();

  const icsMatch = s.match(/\/calendar\/ical\/([^/]+)\/.*\/basic\.ics/i);
  if (icsMatch) s = icsMatch[1];

  const cidMatch = s.match(/[?&]cid=([^&]+)/i);
  if (cidMatch) s = cidMatch[1];

  try {
    s = decodeURIComponent(s);
  } catch (e) {}

  s = s
    .replace(/[\u200B-\u200D\uFEFF]/g, '')
    .replace(/^<|>$/g, '')
    .replace(/^['"]|['"]$/g, '')
    .trim();

  return s || null;
}

function resolveUsableCalendarIds(ids) {
  const unique = new Set();
  const usable = [];

  ids.forEach(raw => {
    const id = normalizeCalendarId(raw);
    if (!id || unique.has(id)) return;
    unique.add(id);

    const cal = CalendarApp.getCalendarById(id);
    if (!cal) {
      logWarn(`無効/未購読/権限不足の可能性: ${raw} → 正規化: ${id}`);
      return;
    }

    logInfo(`[OK] 使用: ${cal.getName()} (${id})`);
    usable.push(id);
  });

  if (usable.length === 0) {
    logError('使えるカレンダー ID がありません。ID・購読状態・権限（予定のすべての情報の表示）を確認してください。');
  }

  return usable;
}

function detectCategory(text, poopKeywords, peeKeywords) {
  if (!text) return '未分類';
  const s = normalizeText_(text);
  const hasPoop = poopKeywords.some(k => s.includes(normalizeText_(k)));
  const hasPee = peeKeywords.some(k => s.includes(normalizeText_(k)));
  if (hasPoop && hasPee) return '両方';
  if (hasPoop) return 'うんち';
  if (hasPee) return 'しっこ';
  return '未分類';
}

function parseMilkEventInfo(title) {
  if (!title || typeof title !== 'string') {
    return null;
  }

  const normalized = normalizeText_(title);
  const prefix = (MILK_SERIES_SETTINGS && MILK_SERIES_SETTINGS.TITLE_PREFIX) || '';
  const normalizedPrefix = prefix ? normalizeText_(prefix) : '';

  const containsMilkKeyword = normalized.includes('ミルク') || normalized.includes('milk');
  const containsPrefix = normalizedPrefix && normalized.includes(normalizedPrefix);
  const containsBottleEmoji = title.includes('🍼');

  if (!containsMilkKeyword && !containsPrefix && !containsBottleEmoji) {
    return null;
  }

  const matches = normalized.match(/(\d+)\s*ml/g);
  if (!matches || !matches.length) {
    return null;
  }

  const amount = matches.reduce((sum, part) => {
    const m = part.match(/(\d+)/);
    if (!m) return sum;
    const value = Number(m[1]);
    return Number.isFinite(value) ? sum + value : sum;
  }, 0);

  if (!Number.isFinite(amount)) {
    return null;
  }

  return {
    amount,
  };
}

/** ===== シートユーティリティ ===== */
function getOrCreateSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function autoResizeAllColumns_(sheet, columnCount) {
  for (let c = 1; c <= columnCount; c++) {
    sheet.autoResizeColumn(c);
  }
}

function setOrResetFilter_(sheet, headerRow, colCount) {
  const range = sheet.getRange(headerRow, 1, sheet.getMaxRows() - headerRow + 1, colCount);
  const filter = sheet.getFilter();
  if (filter) filter.remove();
  range.createFilter();
}
