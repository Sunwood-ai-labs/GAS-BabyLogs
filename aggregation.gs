/**
 * 抽出 → 集計 → グラフまで一括実行します。
 * メニューやトリガーから呼び出せるよう公開関数にしています。
 */
function runAll() {
  extractBabyLogs();
  aggregateAndChart();
}

/**
 * baby_logs シートを集計し、SUMMARY_SHEET に日別・月別集計とグラフを描画します。
 */
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

  const header = dataSheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const values = dataSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const COL = { Category: 0, Date: 1 };
  const milkColIndex = header.indexOf('ミルク量(ml)');

  const mapByDate = {};
  const mapMilkByDate = {};
  values.forEach(row => {
    const category = String(row[COL.Category] || '').trim();
    const dateStr = String(row[COL.Date] || '').trim();
    if (!dateStr) return;

    if (!mapByDate[dateStr]) mapByDate[dateStr] = { poop: 0, pee: 0, both: 0, total: 0 };

    let countedInTotal = false;
    if (category === 'うんち') {
      mapByDate[dateStr].poop++;
      countedInTotal = true;
    } else if (category === 'しっこ') {
      mapByDate[dateStr].pee++;
      countedInTotal = true;
    } else if (category === '両方') {
      mapByDate[dateStr].both++;
      countedInTotal = true;
    } else if (category === CATEGORY_MILK && milkColIndex >= 0) {
      const rawAmount = row[milkColIndex];
      const amount = typeof rawAmount === 'number' ? rawAmount : Number(rawAmount) || 0;
      if (!mapMilkByDate[dateStr]) mapMilkByDate[dateStr] = { amount: 0, count: 0 };
      mapMilkByDate[dateStr].amount += amount;
      mapMilkByDate[dateStr].count += 1;
    } else return;

    if (countedInTotal) {
      mapByDate[dateStr].total++;
    }
  });

  const mapByMonth = {};
  const mapMilkByMonth = {};
  Object.keys(mapByDate).forEach(dateKey => {
    const ym = dateKey.slice(0, 7);
    if (!mapByMonth[ym]) mapByMonth[ym] = { poop: 0, pee: 0, both: 0, total: 0 };
    const dayValue = mapByDate[dateKey];
    mapByMonth[ym].poop += dayValue.poop;
    mapByMonth[ym].pee += dayValue.pee;
    mapByMonth[ym].both += dayValue.both;
    mapByMonth[ym].total += dayValue.total;
  });

  Object.keys(mapMilkByDate).forEach(dateKey => {
    const ym = dateKey.slice(0, 7);
    if (!mapMilkByMonth[ym]) mapMilkByMonth[ym] = { amount: 0, count: 0 };
    const milkValue = mapMilkByDate[dateKey];
    mapMilkByMonth[ym].amount += milkValue.amount;
    mapMilkByMonth[ym].count += milkValue.count;
  });

  const summarySheet = getOrCreateSheet_(ss, SUMMARY_SHEET);
  summarySheet.clear();

  const dayHeader = ['日付', 'うんち', 'しっこ', '両方', '合計'];
  const dayRows = Object.keys(mapByDate)
    .sort()
    .map(dateKey => [
      dateKey,
      mapByDate[dateKey].poop,
      mapByDate[dateKey].pee,
      mapByDate[dateKey].both,
      mapByDate[dateKey].total,
    ]);

  summarySheet.getRange(1, 1, 1, dayHeader.length).setValues([dayHeader]);
  if (dayRows.length) summarySheet.getRange(2, 1, dayRows.length, dayHeader.length).setValues(dayRows);

  const monthHeader = ['月', 'うんち', 'しっこ', '両方', '合計'];
  const monthRows = Object.keys(mapByMonth)
    .sort()
    .map(monthKey => [
      monthKey,
      mapByMonth[monthKey].poop,
      mapByMonth[monthKey].pee,
      mapByMonth[monthKey].both,
      mapByMonth[monthKey].total,
    ]);

  const monthStartCol = dayHeader.length + 2;
  summarySheet.getRange(1, monthStartCol, 1, monthHeader.length).setValues([monthHeader]);
  if (monthRows.length) {
    summarySheet.getRange(2, monthStartCol, monthRows.length, monthHeader.length).setValues(monthRows);
  }

  const milkStartCol = monthStartCol + monthHeader.length + 2;
  const milkDayHeader = ['日付', 'ミルク量(ml)', '授乳回数'];
  const milkDayRows = Object.keys(mapMilkByDate)
    .sort()
    .map(dateKey => [
      dateKey,
      Math.round(mapMilkByDate[dateKey].amount * 10) / 10,
      mapMilkByDate[dateKey].count,
    ]);

  summarySheet.getRange(1, milkStartCol, 1, milkDayHeader.length).setValues([milkDayHeader]);
  if (milkDayRows.length) {
    summarySheet.getRange(2, milkStartCol, milkDayRows.length, milkDayHeader.length).setValues(milkDayRows);
  }

  const milkMonthHeader = ['月', 'ミルク量(ml)', '授乳回数'];
  const milkMonthRows = Object.keys(mapMilkByMonth)
    .sort()
    .map(monthKey => [
      monthKey,
      Math.round(mapMilkByMonth[monthKey].amount * 10) / 10,
      mapMilkByMonth[monthKey].count,
    ]);

  const milkMonthStartRow = Math.max(3, milkDayRows.length + 3);
  summarySheet.getRange(milkMonthStartRow, milkStartCol, 1, milkMonthHeader.length).setValues([milkMonthHeader]);
  if (milkMonthRows.length) {
    summarySheet.getRange(milkMonthStartRow + 1, milkStartCol, milkMonthRows.length, milkMonthHeader.length).setValues(
      milkMonthRows
    );
  }

  summarySheet.setFrozenRows(1);

  const lastUsedCol = Math.max(
    monthStartCol + monthHeader.length - 1,
    milkStartCol + milkDayHeader.length - 1,
    milkStartCol + milkMonthHeader.length - 1
  );
  autoResizeAllColumns_(summarySheet, lastUsedCol);

  summarySheet.getCharts().forEach(chart => summarySheet.removeChart(chart));

  const dayDataEndRow = 1 + Math.max(dayRows.length, 1);
  const dayRangeLastN = (() => {
    const lastN = 30;
    const startRow = Math.max(2, dayDataEndRow - lastN + 1);
    return summarySheet.getRange(startRow, 1, dayDataEndRow - startRow + 1, dayHeader.length);
  })();

  const chartBaseCol = milkStartCol + Math.max(milkDayHeader.length, milkMonthHeader.length) + 4;

  const chart1 = summarySheet
    .newChart()
    .asColumnChart()
    .addRange(summarySheet.getRange(1, 1, 1, 1))
    .addRange(dayRangeLastN)
    .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
    .setStacked()
    .setPosition(2, chartBaseCol, 0, 0)
    .setOption('title', '日別件数（直近30日・積み上げ）')
    .setOption('legend', { position: 'top' })
    .setOption('hAxis', { slantedText: true })
    .build();
  summarySheet.insertChart(chart1);

  const monthDataEndRow = 1 + Math.max(monthRows.length, 1);
  const monthRange = summarySheet.getRange(1, monthStartCol, monthDataEndRow, monthHeader.length);
  const chart2 = summarySheet
    .newChart()
    .asColumnChart()
    .addRange(monthRange)
    .setPosition(20, chartBaseCol, 0, 0)
    .setOption('title', '月別件数')
    .setOption('legend', { position: 'top' })
    .build();
  summarySheet.insertChart(chart2);

  const totalPoop = dayRows.reduce((acc, row) => acc + row[1], 0);
  const totalPee = dayRows.reduce((acc, row) => acc + row[2], 0);
  const totalBoth = dayRows.reduce((acc, row) => acc + row[3], 0);
  const pieStartRow = Math.max(20, 2 + dayRows.length) + 18;
  const pieTable = [
    ['カテゴリ', '件数'],
    ['うんち', totalPoop],
    ['しっこ', totalPee],
    ['両方', totalBoth],
  ];

  const pieAnchor = summarySheet.getRange(pieStartRow, 1, pieTable.length, pieTable[0].length);
  pieAnchor.setValues(pieTable);

  const chart3 = summarySheet
    .newChart()
    .asPieChart()
    .addRange(pieAnchor)
    .setPosition(pieStartRow, 4, 0, 0)
    .setOption('title', 'カテゴリ内訳（期間合計）')
    .build();
  summarySheet.insertChart(chart3);

  if (milkDayRows.length) {
    const milkDayEndRow = 1 + Math.max(milkDayRows.length, 1);
    const milkDayRange = summarySheet.getRange(1, milkStartCol, milkDayEndRow, milkDayHeader.length);
    const chart4 = summarySheet
      .newChart()
      .asColumnChart()
      .addRange(milkDayRange)
      .setPosition(2, chartBaseCol + 6, 0, 0)
      .setOption('title', 'ミルク日別実績（ml）')
      .setOption('legend', { position: 'none' })
      .setOption('hAxis', { slantedText: true })
      .build();
    summarySheet.insertChart(chart4);
  }

  if (milkMonthRows.length) {
    const milkMonthEndRow = milkMonthStartRow + Math.max(milkMonthRows.length, 1);
    const milkMonthRange = summarySheet.getRange(
      milkMonthStartRow,
      milkStartCol,
      milkMonthEndRow - milkMonthStartRow + 1,
      milkMonthHeader.length
    );
    const chart5 = summarySheet
      .newChart()
      .asColumnChart()
      .addRange(milkMonthRange)
      .setPosition(20, chartBaseCol + 6, 0, 0)
      .setOption('title', 'ミルク月別実績（ml）')
      .setOption('legend', { position: 'none' })
      .build();
    summarySheet.insertChart(chart5);
  }

  summarySheet.getRange(1, 1, 1, dayHeader.length).setFontWeight('bold');
  summarySheet.getRange(1, monthStartCol, 1, monthHeader.length).setFontWeight('bold');
  summarySheet.getRange(1, milkStartCol, 1, milkDayHeader.length).setFontWeight('bold');
  summarySheet.getRange(milkMonthStartRow, milkStartCol, 1, milkMonthHeader.length).setFontWeight('bold');

  logInfo('aggregateAndChart: 集計とグラフの更新が完了しました。');
}
