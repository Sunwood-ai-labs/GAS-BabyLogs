/**
 * ミルクタイムの定期予定を作成します。
 * 既存の同一タイトル・時刻の予定が存在する場合はスキップします。
 */
function setupMilkTime() {
  const TZ = SETTINGS.TIMEZONE || 'Asia/Tokyo';
  const targetCalId = (SETTINGS.CALENDAR_IDS && SETTINGS.CALENDAR_IDS[0]) || null;
  if (!targetCalId) {
    throw new Error('SETTINGS.CALENDAR_IDS に対象カレンダー ID が設定されていません。');
  }

  const calendar = CalendarApp.getCalendarById(targetCalId);
  if (!calendar) {
    throw new Error(`カレンダーが見つかりません: ${targetCalId}`);
  }

  const today = new Date();
  const base = new Date(today.getFullYear(), today.getMonth(), today.getDate(), MILK_SERIES_SETTINGS.START_HOUR, MILK_SERIES_SETTINGS.START_MINUTE, 0, 0);

  const labels = MILK_SERIES_SETTINGS.LABELS.slice(0, MILK_SERIES_SETTINGS.COUNT_PER_DAY);
  const titles = labels.map(label => `${MILK_SERIES_SETTINGS.TITLE_PREFIX}${label}`);

  const hasSimilarSeries = (title, startHour, startMinute) => {
    const rangeDays = MILK_SERIES_SETTINGS.SEARCH_RANGE_DAYS;
    const from = shiftDate_(base, -rangeDays);
    const to = new Date(shiftDate_(base, rangeDays).getTime() + 24 * 60 * 60 * 1000);
    const events = calendar.getEvents(from, to, { search: title });
    return events.some(ev => {
      const st = ev.getStartTime();
      return st.getHours() === startHour && st.getMinutes() === startMinute;
    });
  };

  for (let i = 0; i < labels.length; i++) {
    const start = new Date(base.getTime() + i * MILK_SERIES_SETTINGS.INTERVAL_HOURS * 60 * 60 * 1000);
    const end = new Date(start.getTime() + MILK_SERIES_SETTINGS.DURATION_MINUTES * 60 * 1000);
    const title = titles[i];

    if (hasSimilarSeries(title, start.getHours(), start.getMinutes())) {
      logInfo(`[SKIP] 既存あり: ${title} ${fmt(start, TZ)} - ${fmt(end, TZ)}`);
      continue;
    }

    const recurrence = CalendarApp.newRecurrence().addDailyRule();
    const series = calendar.createEventSeries(title, start, end, recurrence);

    try {
      if (series.removeAllReminders) series.removeAllReminders();
    } catch (e) {}
    series.addPopupReminder(30);
    series.addPopupReminder(10);
    series.setColor(CalendarApp.EventColor.YELLOW);

    logInfo(`[OK] 作成: ${title} ${fmt(start, TZ)} - ${fmt(end, TZ)} @ ${calendar.getName()}`);
  }

  SpreadsheetApp.getUi().alert('ミルクタイム（1時間枠）の定期予定を作成しました。');
}

/**
 * ミルクタイムの定期予定（TITLE_PREFIX を含む）を削除します。
 * createEventSeries で作成された予定を優先して削除し、単発イベントも対象とします。
 */
function deleteMilkTimeSeries() {
  const targetCalId = (SETTINGS.CALENDAR_IDS && SETTINGS.CALENDAR_IDS[0]) || null;
  if (!targetCalId) {
    throw new Error('SETTINGS.CALENDAR_IDS に対象カレンダー ID が設定されていません。');
  }

  const calendar = CalendarApp.getCalendarById(targetCalId);
  if (!calendar) {
    throw new Error(`カレンダーが見つかりません: ${targetCalId}`);
  }

  const base = new Date();
  const rangeDays = MILK_SERIES_SETTINGS.SEARCH_RANGE_DAYS;
  const from = shiftDate_(base, -rangeDays);
  const to = new Date(shiftDate_(base, rangeDays).getTime() + 24 * 60 * 60 * 1000);
  const prefix = MILK_SERIES_SETTINGS.TITLE_PREFIX;

  const events = calendar.getEvents(from, to, { search: prefix });
  if (!events.length) {
    SpreadsheetApp.getUi().alert('削除対象のミルクタイム予定は見つかりませんでした。');
    return;
  }

  const deletedSeriesIds = new Set();
  let deletedCount = 0;

  events.forEach(event => {
    const title = String(event.getTitle() || '');
    if (!title.startsWith(prefix)) return;

    if (event.isRecurringEvent()) {
      const series = event.getEventSeries();
      if (series && !deletedSeriesIds.has(series.getId())) {
        series.deleteEventSeries();
        deletedSeriesIds.add(series.getId());
        deletedCount++;
      }
    } else {
      event.deleteEvent();
      deletedCount++;
    }
  });

  SpreadsheetApp.getUi().alert(`ミルクタイムの予定を ${deletedCount} 件削除しました。`);
}
