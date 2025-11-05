/**
 * スプレッドシートを開いたときに便利メニューを追加します。
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('👶 Baby Logs')
    .addItem('抽出 → 集計 → グラフ（全部やる）', 'runAll')
    .addSeparator()
    .addItem('データ抽出のみ（カレンダー → baby_logs）', 'extractBabyLogs')
    .addItem('集計＆グラフのみ（baby_summary 更新）', 'aggregateAndChart')
    .addToUi();

  ui.createMenu('🍼 Milk Setup')
    .addItem('ミルクタイム定期予定を作成', 'setupMilkTime')
    .addItem('ミルクタイム定期予定を削除', 'deleteMilkTimeSeries')
    .addToUi();
}
