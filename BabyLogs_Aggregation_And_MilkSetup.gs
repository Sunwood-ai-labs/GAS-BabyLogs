/** ========= ミルクタイムの定期予定セットアップ（1時間枠） =========
 *  仕様:
 *   - 開始: 毎日 01:30 / 1時間枠
 *   - 間隔: 3時間ごと（01:30, 04:30, 07:30, 10:30, 13:30, 16:30, 19:30, 22:30）
 *   - タイトル: 🍼ミルクタイム❶, ❷, ❸, …（全8本）
 *   - 通知: 30分前 & 10分前（ポップアップ）
 *   - 色: 黄色
 *   - 対象カレンダー: SETTINGS.CALENDAR_IDS[0]
 */
function setupMilkTime() {
  const TZ = SETTINGS.TIMEZONE || 'Asia/Tokyo';
  const targetCalId = (SETTINGS.CALENDAR_IDS && SETTINGS.CALENDAR_IDS[0]) || null;
  if (!targetCalId) throw new Error('SETTINGS.CALENDAR_IDS に対象カレンダーIDが設定されていません。');

  const cal = CalendarApp.getCalendarById(targetCalId);
  if (!cal) throw new Error(`カレンダーが見つかりません: ${targetCalId}`);

  // 基準日（今日の 01:30 から作成）
  const today = new Date();
  const base = new Date(today.getFullYear(), today.getMonth(), today.getDate(), 1, 30, 0, 0);

  const DURATION_MIN = 60;                 // ← ここが 60分（以前は 30）
  const INTERVAL_HOURS = 3;               // 3時間ごと
  const COUNT = 8;                        // 1日8本
  const labels = ['❶','❷','❸','❹','❺','❻','❼','❽'];
  const seriesTitles = labels.map(l => `🍼ミルクタイム${l}`);

  // 既存の重複を避ける簡易チェック
  const hasSimilarSeries = (title, startHour, startMinute) => {
    const from = new Date(base.getFullYear(), base.getMonth(), base.getDate() - 15, 0, 0, 0, 0);
    const to   = new Date(base.getFullYear(), base.getMonth(), base.getDate() + 15, 23, 59, 59, 999);
    const evs = cal.getEvents(from, to, { search: title });
    return evs.some(ev => {
      const st = ev.getStartTime();
      return st.getHours() === startHour && st.getMinutes() === startMinute;
    });
  };

  for (let i = 0; i < COUNT; i++) {
    const start = new Date(base.getTime() + i * INTERVAL_HOURS * 60 * 60 * 1000);
    const end   = new Date(start.getTime() + DURATION_MIN * 60 * 1000);
    const title = seriesTitles[i];

    if (hasSimilarSeries(title, start.getHours(), start.getMinutes())) {
      Logger.log(`[SKIP] 既存あり: ${title} ${fmt(start, TZ)} - ${fmt(end, TZ)}`);
      continue;
    }

    // 日次の繰り返し
    const recur = CalendarApp.newRecurrence().addDailyRule();
    const series = cal.createEventSeries(title, start, end, recur);

    // 通知（ポップアップ）
    try { series.removeAllReminders && series.removeAllReminders(); } catch (e) {}
    series.addPopupReminder(30);
    series.addPopupReminder(10);

    // 色: 黄色
    series.setColor(CalendarApp.EventColor.YELLOW);

    Logger.log(`[OK] 作成: ${title} ${fmt(start, TZ)} - ${fmt(end, TZ)} ＠ ${cal.getName()}`);
  }

  SpreadsheetApp.getUi().alert('ミルクタイム（1時間枠）の定期予定を作成しました。');
}
