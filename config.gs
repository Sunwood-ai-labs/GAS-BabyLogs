/**
 * グローバル設定：使用するカレンダーや出力シートなどを管理します。
 * 必要に応じてあなたの環境に合わせて編集してください。
 */
const SETTINGS = {
  /**
   * Baby Logs を取得する対象カレンダーの ID 一覧。
   * - 共有カレンダー ID（@group.calendar.google.com ...）を推奨。
   * - ics URL や cid= パラメータを貼り付けても内部で正規化されます。
   * - 'primary' は含めないでください。プライベートカレンダーを保護します。
   */
  CALENDAR_IDS: [
    '352c174852fa30b97367fc0734341b2d1f0edf5c65998633f2d2d8fa4f021de8@group.calendar.google.com',
  ],

  /**
   * 取得期間：今日からさかのぼる日数と、未来側の日数。
   * 例) DAYS_BACK = 60, DAYS_AHEAD = 7 → 過去 60 日〜未来 7 日を取得。
   */
  DAYS_BACK: 60,
  DAYS_AHEAD: 7,

  /**
   * カテゴリ判定キーワード（タイトルに含まれる文字列）。
   * 表記ゆれがあれば適宜追加してください。
   */
  KEYWORDS_POOP: ['うんち', 'ウンチ', '💩', '便', '排便'],
  KEYWORDS_PEE: ['しっこ', 'おしっこ', 'オシッコ', '尿', '排尿'],

  /** 出力シート名 */
  SHEET_NAME: 'baby_logs',

  /** スクリプト全体で利用するタイムゾーン */
  TIMEZONE: 'Asia/Tokyo',

  /**
   * true にすると抽出結果をログに出すだけでシートへ書き込まない。
   * テスト時に利用してください。
   */
  DRY_RUN: false,
};

/** 集計結果を出力するシート名 */
const SUMMARY_SHEET = 'baby_summary';

/** ミルク実績のカテゴリ名 */
const CATEGORY_MILK = 'ミルク';

/**
 * ミルクタイム定期予定関連の設定。
 * setupMilkTime / deleteMilkTimeSeries で共通利用します。
 */
const MILK_SERIES_SETTINGS = {
  /** 1 回あたりの枠（分） */
  DURATION_MINUTES: 60,
  /** 間隔（時間） */
  INTERVAL_HOURS: 3,
  /** 1 日あたりの作成本数 */
  COUNT_PER_DAY: 8,
  /** 初回の開始時刻 */
  START_HOUR: 1,
  START_MINUTE: 30,
  /** タイトルの番号ラベル */
  LABELS: ['❶', '❷', '❸', '❹', '❺', '❻', '❼', '❽'],
  /** タイトル接頭辞（検索・削除時にも利用） */
  TITLE_PREFIX: '🍼ミルクタイム',
  /** 重複チェック／削除で検索する期間（日） */
  SEARCH_RANGE_DAYS: 30,
};
