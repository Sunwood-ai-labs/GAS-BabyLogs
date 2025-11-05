/** Baby logs extractor ぜんぶ入り（カテゴリ分け・ID正規化・詳細ログ・フィルタ・自動整形）V1
 *  使い方：
 *   1) スプレッドシートを作成 → 拡張機能→Apps Script を開く
 *   2) このファイルを貼り付けて保存
 *   3) extractBabyLogs() を実行（初回は権限付与）
 *   4) シート "baby_logs" に結果が出力されます
 */

const SETTINGS = {
  // ★ 'primary' は入れない：実際に使う共有カレンダーIDだけを指定
  CALENDAR_IDS: [
    '352c174852fa30b97367fc0734341b2d1f0edf5c65998633f2d2d8fa4f021de8@group.calendar.google.com'
  ],

  // 取得期間（必要に応じて調整）
  DAYS_BACK: 60,
  DAYS_AHEAD: 7,

  // カテゴリ別キーワード（表記ゆれがあれば足してください）
  KEYWORDS_POOP: ['うんち','ウンチ','💩','便','排便'],
  KEYWORDS_PEE:  ['しっこ','おしっこ','オシッコ','尿','排尿'],

  // 出力先
  SHEET_NAME: 'baby_logs',

  // タイムゾーン
  TIMEZONE: 'Asia/Tokyo',

  // ログ確認だけしたい時は true（シートには書かない）
  DRY_RUN: false,
};

/** メイン：カテゴリ分けしてスプレッドシートに書き込み */
function extractBabyLogs() {
  const startedAt = new Date();
  logInfo(`=== extractBabyLogs start @ ${startedAt.toISOString()} ===`);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const TZ = SETTINGS.TIMEZONE;

    // 期間
    const now = new Date();
    const start = shiftDate_(now, -SETTINGS.DAYS_BACK);
    const end   = shiftDate_(now,  SETTINGS.DAYS_AHEAD);
    logInfo(`Window: ${fmt(start,TZ)} ～ ${fmt(end,TZ)}`);

    // カレンダーIDを正規化→有効なものだけに絞る
    const CAL_IDS = resolveUsableCalendarIds_(SETTINGS.CALENDAR_IDS);
    if (CAL_IDS.length === 0) {
      logError('使えるカレンダーIDが 0 件のため処理を終了します。');
      return;
    }

    const headers = ['Category','日付','開始','終了','終日','タイトル','カレンダー','イベントID','更新日時'];
    const rows = [];
    let total = 0, hit = 0;

    // 各カレンダーから取得
    CAL_IDS.forEach(id => {
      const cal = CalendarApp.getCalendarById(id);
      const events = cal.getEvents(start, end);
      total += events.length;
      logInfo(`Fetch: ${cal.getName()} (${id}) -> ${events.length} events`);

      let hitThisCal = 0;
      events.forEach(ev => {
        const title = (ev.getTitle() || '').trim();
        const cat = detectCategory_(title, SETTINGS.KEYWORDS_POOP, SETTINGS.KEYWORDS_PEE);
        if (cat === '未分類') return;

        hit++; hitThisCal++;
        const isAllDay = ev.isAllDayEvent();
        const st = ev.getStartTime(), et = ev.getEndTime();
        rows.push([
          cat,
          Utilities.formatDate(st, TZ, 'yyyy-MM-dd'),
          isAllDay ? '' : Utilities.formatDate(st, TZ, 'HH:mm'),
          isAllDay ? '' : Utilities.formatDate(et, TZ, 'HH:mm'),
          isAllDay ? 'TRUE' : 'FALSE',
          title,
          cal.getName() || id,
          ev.getId(),
          Utilities.formatDate(new Date(ev.getLastUpdated()), TZ, 'yyyy-MM-dd HH:mm:ss'),
        ]);
      });
      logInfo(`[HIT] ${cal.getName()} (${id}) => ${hitThisCal} rows`);
    });

    logInfo(`Total events: ${total}, Matched: ${hit}`);

    // 並べ替え：日付→開始→カテゴリ
    rows.sort((a,b)=>{
      const ak = `${a[1]} ${a[2]||'00:00'} ${a[0]}`, bk = `${b[1]} ${b[2]||'00:00'} ${b[0]}`;
      return ak < bk ? -1 : ak > bk ? 1 : 0;
    });

    if (SETTINGS.DRY_RUN) {
      logInfo(`[DRY_RUN] rows prepared = ${rows.length} (no write)`);
    } else {
      const sheet = getOrCreateSheet_(ss, SETTINGS.SHEET_NAME);
      sheet.clearContents();
      sheet.getRange(1,1,1,headers.length).setValues([headers]);
      if (rows.length) sheet.getRange(2,1,rows.length,headers.length).setValues(rows);
      sheet.setFrozenRows(1);
      autoResizeAllColumns_(sheet, headers.length);
      setOrResetFilter_(sheet, 1, headers.length);
      logInfo(`Wrote ${rows.length} rows to "${SETTINGS.SHEET_NAME}"`);
    }

    logInfo(`=== done (${new Date() - startedAt} ms) ===`);
  } catch (e) {
    logError(e && e.stack ? e.stack : e);
    throw e;
  }
}

/** 任意：毎朝7時に自動更新したい場合は一度だけ実行 */
function createDailyTrigger() {
  ScriptApp.newTrigger('extractBabyLogs').timeBased().atHour(7).everyDays(1).create();
}

/* ========= 補助 ========= */

// どんな貼り方（cid=URL/ics/生ID）でも内部IDへ正規化
function normalizeCalendarId(raw) {
  if (!raw) return null;
  let s = String(raw).trim();

  // ics 秘密アドレス → ID抽出
  const icsMatch = s.match(/\/calendar\/ical\/([^/]+)\/.*\/basic\.ics/i);
  if (icsMatch) s = icsMatch[1];

  // cid= 付きURL → 値抽出
  const cidMatch = s.match(/[?&]cid=([^&]+)/i);
  if (cidMatch) s = cidMatch[1];

  // URLデコード（%40 → @ など）
  try { s = decodeURIComponent(s); } catch (_) {}

  // 不可視スペース・引用符・山括弧を除去
  s = s.replace(/[\u200B-\u200D\uFEFF]/g, '').replace(/^<|>$/g, '').replace(/^['"]|['"]$/g, '').trim();

  return s || null;
}

// 有効なカレンダーだけを返す（ログ出力込み）
function resolveUsableCalendarIds_(ids) {
  const unique = new Set();
  const usable = [];
  ids.forEach(raw => {
    const id = normalizeCalendarId(raw);
    if (!id || unique.has(id)) return;
    unique.add(id);

    const cal = CalendarApp.getCalendarById(id);
    if (!cal) {
      logWarn(`無効/未購読/権限不足の可能性: ${raw}  → 正規化: ${id}`);
    } else {
      logInfo(`[OK] 使用: ${cal.getName()} (${id})`);
      usable.push(id);
    }
  });
  if (usable.length === 0) {
    logError('使えるIDがありません。ID/購読/権限（予定のすべての情報の表示）を確認してください。');
  }
  return usable;
}

// タイトルからカテゴリ判定
function detectCategory_(text, poopKeywords, peeKeywords) {
  if (!text) return '未分類';
  const s = normalize_(text);
  const hasPoop = poopKeywords.some(k => s.includes(normalize_(k)));
  const hasPee  = peeKeywords.some(k => s.includes(normalize_(k)));
  if (hasPoop && hasPee) return '両方';
  if (hasPoop) return 'うんち';
  if (hasPee)  return 'しっこ';
  return '未分類';
}

// 軽い正規化（全角英数→半角、lower）
function normalize_(s){
  s = (s||'').trim();
  try { s = s.replace(/[Ａ-Ｚａ-ｚ０-９]/g, ch => String.fromCharCode(ch.charCodeAt(0)-0xFEE0)); } catch(_){}
  return s.toLowerCase();
}

// シート関連
function getOrCreateSheet_(ss, name){ return ss.getSheetByName(name) || ss.insertSheet(name); }
function autoResizeAllColumns_(sheet, n){ for (let c=1;c<=n;c++) sheet.autoResizeColumn(c); }
function setOrResetFilter_(sheet, headerRow, colCount){ const range=sheet.getRange(headerRow,1,sheet.getMaxRows()-headerRow+1,colCount); const f=sheet.getFilter(); if (f) f.remove(); range.createFilter(); }

// 日付ユーティリティ
function shiftDate_(base, days){ const d=new Date(base); d.setDate(d.getDate()+days); d.setHours(0,0,0,0); return d; }
function fmt(dt,tz){ return Utilities.formatDate(dt, tz, 'yyyy-MM-dd HH:mm'); }

/* ===== ロガー（Logger と console の両方へ） ===== */
function logInfo(msg){ Logger.log(msg); try{console.log(msg);}catch(_){} }
function logWarn(msg){ Logger.log('[WARN] '+msg); try{console.warn(msg);}catch(_){} }
function logError(msg){ Logger.log('[ERROR] '+msg); try{console.error(msg);}catch(_){} }

/** ========= 集計＆グラフ =========
 *  前提：baby_logs シートの列は以下（1行目ヘッダ）
 *   A:Category / B:日付 / C:開始 / D:終了 / E:終日 / F:タイトル / G:カレンダー / H:イベントID / I:更新日時
 *  使い方：
 *   1) extractBabyLogs() を実行してデータ更新
 *   2) aggregateAndChart() を実行（またはメニューから）
 */

const SUMMARY_SHEET = 'baby_summary';  // 集計出力シート名

/** すべて：抽出 → 集計 → グラフ */
function runAll() {
  extractBabyLogs();       // 既存の抽出関数（あなたの環境にあるやつ）
  aggregateAndChart();     // 集計＋グラフ
}

/** 集計＋グラフ（これだけでもOK） */
function aggregateAndChart() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName(SETTINGS.SHEET_NAME || 'baby_logs');
  if (!dataSheet) {
    throw new Error(`データシート "${SETTINGS.SHEET_NAME || 'baby_logs'}" が見つかりません。先に extractBabyLogs() を実行してください。`);
  }
  const lastRow = dataSheet.getLastRow();
  const lastCol = dataSheet.getLastColumn();
  if (lastRow < 2) {
    throw new Error('baby_logs にデータ行がありません。');
  }

  // 全データ読み込み
  const values = dataSheet.getRange(2, 1, lastRow - 2 + 1, lastCol).getValues();
  const COL = { Category: 0, Date: 1 }; // 0-based index within values

  // 日付別にカウント
  /** mapByDate = {
   *   'yyyy-MM-dd': { poop: n, pee: n, both: n, total: n }
   * }
   */
  const mapByDate = {};
  values.forEach(row => {
    const category = String(row[COL.Category] || '').trim();
    const dateStr  = String(row[COL.Date] || '').trim();
    if (!dateStr) return;

    if (!mapByDate[dateStr]) mapByDate[dateStr] = { poop: 0, pee: 0, both: 0, total: 0 };

    if (category === 'うんち') mapByDate[dateStr].poop++;
    else if (category === 'しっこ') mapByDate[dateStr].pee++;
    else if (category === '両方') mapByDate[dateStr].both++;
    else return;

    mapByDate[dateStr].total++;
  });

  // 月別にカウント（yyyy-MM）
  const mapByMonth = {};
  Object.keys(mapByDate).forEach(d => {
    const ym = d.slice(0, 7); // 'yyyy-MM'
    if (!mapByMonth[ym]) mapByMonth[ym] = { poop: 0, pee: 0, both: 0, total: 0 };
    const v = mapByDate[d];
    mapByMonth[ym].poop += v.poop;
    mapByMonth[ym].pee  += v.pee;
    mapByMonth[ym].both += v.both;
    mapByMonth[ym].total += v.total;
  });

  // 出力シートを準備
  const sumSheet = getOrCreateSheet_(ss, SUMMARY_SHEET);
  sumSheet.clear();

  // 1. 日別テーブル
  const dayHeader = ['日付','うんち','しっこ','両方','合計'];
  const dayRows = Object.keys(mapByDate)
    .sort() // yyyy-MM-dd 文字列なのでこれで日付昇順
    .map(d => [d, mapByDate[d].poop, mapByDate[d].pee, mapByDate[d].both, mapByDate[d].total]);

  sumSheet.getRange(1, 1, 1, dayHeader.length).setValues([dayHeader]);
  if (dayRows.length) sumSheet.getRange(2, 1, dayRows.length, dayHeader.length).setValues(dayRows);

  // 2. 月別テーブル（隣に配置）
  const monthHeader = ['月','うんち','しっこ','両方','合計'];
  const monthRows = Object.keys(mapByMonth)
    .sort()
    .map(m => [m, mapByMonth[m].poop, mapByMonth[m].pee, mapByMonth[m].both, mapByMonth[m].total]);

  const monthStartCol = dayHeader.length + 2; // 日別の右に1列空けて配置
  sumSheet.getRange(1, monthStartCol, 1, monthHeader.length).setValues([monthHeader]);
  if (monthRows.length) sumSheet.getRange(2, monthStartCol, monthRows.length, monthHeader.length).setValues(monthRows);

  // 見た目
  sumSheet.setFrozenRows(1);
  autoResizeAllColumns_(sumSheet, monthStartCol + monthHeader.length - 1);

  // 既存グラフは削除して作り直し
  sumSheet.getCharts().forEach(c => sumSheet.removeChart(c));

  // ========== グラフ 1: 日別 積み上げ棒（直近30日） ==========
  const dayDataEndRow = 1 + Math.max(dayRows.length, 1);
  const dayRangeAll = sumSheet.getRange(1, 1, dayDataEndRow, dayHeader.length);

  // 直近30日の範囲（データが少ない場合は全件）
  const lastN = 30;
  const startRowForLastN = Math.max(2, dayDataEndRow - lastN + 1);
  const dayRangeLastN = sumSheet.getRange(startRowForLastN, 1, dayDataEndRow - startRowForLastN + 1, dayHeader.length);

  let chart1 = sumSheet.newChart()
    .asColumnChart()
    .addRange(sumSheet.getRange(1,1,1,1)) // タイトル列のヘッダ（軸ラベル用ダミー）
    .addRange(dayRangeLastN)               // 実データ
    .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
    .setStacked()
    .setPosition(2, monthStartCol + monthHeader.length + 1, 0, 0) // 月表のさらに右に配置
    .setOption('title', '日別件数（直近30日・積み上げ）')
    .setOption('legend', { position: 'top' })
    .setOption('hAxis', { slantedText: true })
    .build();
  sumSheet.insertChart(chart1);

  // ========== グラフ 2: 月別 合計（クラスター縦棒） ==========
  const monthDataEndRow = 1 + Math.max(monthRows.length, 1);
  const monthRange = sumSheet.getRange(1, monthStartCol, monthDataEndRow, monthHeader.length);

  let chart2 = sumSheet.newChart()
    .asColumnChart()
    .addRange(monthRange)
    .setPosition(20, monthStartCol + monthHeader.length + 1, 0, 0)
    .setOption('title', '月別件数')
    .setOption('legend', { position: 'top' })
    .build();
  sumSheet.insertChart(chart2);

  // ========== グラフ 3: カテゴリ内訳（円、全期間） ==========
  const totalPoop = dayRows.reduce((a,r)=>a+r[1],0);
  const totalPee  = dayRows.reduce((a,r)=>a+r[2],0);
  const totalBoth = dayRows.reduce((a,r)=>a+r[3],0);
  const pieStartRow = Math.max(20, 2 + dayRows.length) + 18;
  const pieTable = [
    ['カテゴリ','件数'],
    ['うんち', totalPoop],
    ['しっこ', totalPee],
    ['両方', totalBoth],
  ];
  const pieAnchor = sumSheet.getRange(pieStartRow, 1, pieTable.length, pieTable[0].length);
  pieAnchor.setValues(pieTable);

  let chart3 = sumSheet.newChart()
    .asPieChart()
    .addRange(pieAnchor)
    .setPosition(pieStartRow, 4, 0, 0)
    .setOption('title', 'カテゴリ内訳（期間合計）')
    .build();
  sumSheet.insertChart(chart3);

  // 仕上げ
  sumSheet.getRange(1,1,1,dayHeader.length).setFontWeight('bold');
  sumSheet.getRange(1,monthStartCol,1,monthHeader.length).setFontWeight('bold');

  Logger.log('aggregateAndChart: 集計とグラフの更新が完了しました。');
}


/** ===== メニュー追加 ===== */
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  // 抽出・集計・グラフ用メニュー
  ui.createMenu('👶 Baby Logs')
    .addItem('抽出 → 集計 → グラフ（全部やる）', 'runAll')
    .addSeparator()
    .addItem('データ抽出のみ（カレンダー → baby_logs）', 'extractBabyLogs')
    .addItem('集計＆グラフのみ（baby_summary 更新）', 'aggregateAndChart')
    .addToUi();

  // ミルクタイム定期予定セットアップ用メニュー
  ui.createMenu('🍼 Milk Setup')
    .addItem('ミルクタイム定期予定を作成', 'setupMilkTime')
    .addItem('ミルクタイム定期予定を削除', 'deleteMilkTimeSeries')
    .addToUi();
}
