/**
 * Baby Logs のメイン処理：対象カレンダーからイベントを抽出し、カテゴリ分けしてシートへ書き込みます。
 */
function extractBabyLogs() {
  const startedAt = new Date();
  logInfo(`=== extractBabyLogs start @ ${startedAt.toISOString()} ===`);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const TZ = SETTINGS.TIMEZONE;
  const now = new Date();
  const windowStart = shiftDate_(now, -SETTINGS.DAYS_BACK);
  const windowEnd = shiftDate_(now, SETTINGS.DAYS_AHEAD);
  logInfo(`Window: ${fmt(windowStart, TZ)} ～ ${fmt(windowEnd, TZ)}`);

  const calendarIds = resolveUsableCalendarIds(SETTINGS.CALENDAR_IDS);
  if (calendarIds.length === 0) {
    logError('処理を継続できません（利用可能なカレンダーが 0 件）');
    return;
  }

  const headers = ['Category', '日付', '開始', '終了', '終日', 'タイトル', 'カレンダー', 'イベントID', '更新日時', 'ミルク量(ml)'];
  const rows = [];
  let totalEvents = 0;
  let matchedEvents = 0;

  calendarIds.forEach(id => {
    const calendar = CalendarApp.getCalendarById(id);
    const events = calendar.getEvents(windowStart, windowEnd);
    totalEvents += events.length;
    logInfo(`Fetch: ${calendar.getName()} (${id}) -> ${events.length} events`);

    let matchedPerCalendar = 0;
    events.forEach(event => {
      const title = (event.getTitle() || '').trim();
      const milkInfo = parseMilkEventInfo(title);
      const isMilkLog = milkInfo && milkInfo.amount !== null;
      const category = isMilkLog ? CATEGORY_MILK : detectCategory(title, SETTINGS.KEYWORDS_POOP, SETTINGS.KEYWORDS_PEE);
      if (category === '未分類') return;

      matchedEvents++;
      matchedPerCalendar++;
      const isAllDay = event.isAllDayEvent();
      const start = event.getStartTime();
      const end = event.getEndTime();
      rows.push([
        category,
        Utilities.formatDate(start, TZ, 'yyyy-MM-dd'),
        isAllDay ? '' : Utilities.formatDate(start, TZ, 'HH:mm'),
        isAllDay ? '' : Utilities.formatDate(end, TZ, 'HH:mm'),
        isAllDay ? 'TRUE' : 'FALSE',
        title,
        calendar.getName() || id,
        event.getId(),
        Utilities.formatDate(new Date(event.getLastUpdated()), TZ, 'yyyy-MM-dd HH:mm:ss'),
        isMilkLog ? milkInfo.amount : '',
      ]);
    });

    logInfo(`[HIT] ${calendar.getName()} (${id}) => ${matchedPerCalendar} rows`);
  });

  logInfo(`Total events: ${totalEvents}, Matched: ${matchedEvents}`);
  rows.sort((a, b) => {
    const ak = `${a[1]} ${a[2] || '00:00'} ${a[0]}`;
    const bk = `${b[1]} ${b[2] || '00:00'} ${b[0]}`;
    return ak < bk ? -1 : ak > bk ? 1 : 0;
  });

  if (SETTINGS.DRY_RUN) {
    logInfo(`[DRY_RUN] rows prepared = ${rows.length} (no write)`);
    logInfo(`=== done (dry run, ${new Date() - startedAt} ms) ===`);
    return;
  }

  const sheet = getOrCreateSheet_(ss, SETTINGS.SHEET_NAME);
  sheet.clearContents();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (rows.length) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
  sheet.setFrozenRows(1);
  autoResizeAllColumns_(sheet, headers.length);
  setOrResetFilter_(sheet, 1, headers.length);

  logInfo(`Wrote ${rows.length} rows to "${SETTINGS.SHEET_NAME}"`);
  logInfo(`=== done (${new Date() - startedAt} ms) ===`);
}

/** 任意：毎朝 7 時に自動実行したい場合にトリガーを作成します。 */
function createDailyTrigger() {
  ScriptApp.newTrigger('extractBabyLogs').timeBased().atHour(7).everyDays(1).create();
}
