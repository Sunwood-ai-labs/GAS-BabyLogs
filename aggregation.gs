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
  const startTimeColIndex = header.indexOf('開始');

  const hourMapByCategory = {};
  const registerHeatmapCount = (categoryKey, startTimeStr) => {
    if (categoryKey && startTimeStr) {
      const hour = parseInt(startTimeStr.slice(0, 2), 10);
      if (Number.isNaN(hour)) {
        return;
      }
      if (hourMapByCategory[categoryKey] === undefined) {
        hourMapByCategory[categoryKey] = Array(24).fill(0);
      }
      hourMapByCategory[categoryKey][hour] += 1;
    }
  };

  const mapByDate = {};
  const mapMilkByDate = {};
  values.forEach(row => {
    const category = String(row[COL.Category] || '').trim();
    const dateStr = String(row[COL.Date] || '').trim();
    if (!dateStr) return;

    if (!mapByDate[dateStr]) mapByDate[dateStr] = { poop: 0, pee: 0, both: 0, total: 0 };
    const startTimeStr =
      startTimeColIndex >= 0 ? String(row[startTimeColIndex] || '').trim() : '';

    let countedInTotal = false;
    if (category === 'うんち') {
      mapByDate[dateStr].poop++;
      countedInTotal = true;
      registerHeatmapCount(category, startTimeStr);
    } else if (category === 'しっこ') {
      mapByDate[dateStr].pee++;
      countedInTotal = true;
      registerHeatmapCount(category, startTimeStr);
    } else if (category === '両方') {
      mapByDate[dateStr].both++;
      countedInTotal = true;
      registerHeatmapCount(category, startTimeStr);
    } else if (category === CATEGORY_MILK && milkColIndex >= 0) {
      const rawAmount = row[milkColIndex];
      const amount = typeof rawAmount === 'number' ? rawAmount : Number(rawAmount) || 0;
      if (!mapMilkByDate[dateStr]) mapMilkByDate[dateStr] = { amount: 0, count: 0 };
      mapMilkByDate[dateStr].amount += amount;
      mapMilkByDate[dateStr].count += 1;
      registerHeatmapCount(category, startTimeStr);
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
  summarySheet.setConditionalFormatRules([]);

  const parseDateForSheet = dateStr => {
    if (!dateStr) return '';
    const parts = String(dateStr).split('-');
    if (parts.length < 3) return dateStr;
    const [yearStr, monthStr, dayStr] = parts;
    const year = Number(yearStr);
    const month = Number(monthStr);
    const day = Number(dayStr);
    if ([year, month, day].some(value => Number.isNaN(value))) {
      return dateStr;
    }
    return new Date(year, month - 1, day);
  };

  const dayHeader = ['日付', 'うんち', 'しっこ', '両方', '合計'];
  const dayRows = Object.keys(mapByDate)
    .sort()
    .map(dateKey => [
      parseDateForSheet(dateKey),
      mapByDate[dateKey].poop,
      mapByDate[dateKey].pee,
      mapByDate[dateKey].both,
      mapByDate[dateKey].total,
    ]);
  const dayRowsDisplay = dayRows.map(row => [formatMonthDay(row[0]), row[1], row[2], row[3], row[4]]);

  summarySheet.getRange(1, 1, 1, dayHeader.length).setValues([dayHeader]);
  if (dayRows.length) {
    const dayRange = summarySheet.getRange(2, 1, dayRows.length, dayHeader.length);
    dayRange.setValues(dayRows);
    dayRange.offset(0, 0, dayRows.length, 1).setNumberFormat('M/d');
  }

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
      parseDateForSheet(dateKey),
      Math.round(mapMilkByDate[dateKey].amount * 10) / 10,
      mapMilkByDate[dateKey].count,
    ]);
  const milkDayRowsDisplay = milkDayRows.map(row => [formatMonthDay(row[0]), row[1], row[2]]);

  summarySheet.getRange(1, milkStartCol, 1, milkDayHeader.length).setValues([milkDayHeader]);
  if (milkDayRows.length) {
    const milkDayRange = summarySheet.getRange(2, milkStartCol, milkDayRows.length, milkDayHeader.length);
    milkDayRange.setValues(milkDayRows);
    milkDayRange.offset(0, 0, milkDayRows.length, 1).setNumberFormat('M/d');
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

  let lastUsedCol = Math.max(
    monthStartCol + monthHeader.length - 1,
    milkStartCol + milkDayHeader.length - 1,
    milkStartCol + milkMonthHeader.length - 1
  );

  summarySheet.getCharts().forEach(chart => summarySheet.removeChart(chart));

  const dayDataEndRow = 1 + Math.max(dayRows.length, 1);
  const dayRangeLastN = (() => {
    const lastN = 30;
    const startRow = Math.max(2, dayDataEndRow - lastN + 1);
    return summarySheet.getRange(startRow, 1, dayDataEndRow - startRow + 1, dayHeader.length);
  })();

  const chartBaseCol = milkStartCol + Math.max(milkDayHeader.length, milkMonthHeader.length) + 4;

  const chartPositions = {
    event: { row: 2, col: chartBaseCol },
    milk: { row: 2, col: chartBaseCol + 6 },
  };
  const advanceChartPosition = (group, heightRows = 22) => {
    const target = chartPositions[group];
    if (!target) {
      throw new Error(`未定義のチャートグループです: ${group}`);
    }
    const { row, col } = target;
    target.row += heightRows;
    return { row, col };
  };

  const dayChartPosition = advanceChartPosition('event');
  const chart1Builder = summarySheet
    .newChart()
    .asColumnChart()
    .addRange(summarySheet.getRange(1, 1, 1, dayHeader.length))
    .addRange(dayRangeLastN)
    .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
    .setStacked()
    .setOption('title', '日別件数（直近30日・積み上げ）')
    .setOption('legend', { position: 'top' })
    .setOption('hAxis', { slantedText: true })
    .setOption('height', 320)
    .setOption('series', {
      0: { labelInLegend: 'うんち', color: '#8d6e63' },
      1: { labelInLegend: 'しっこ', color: '#fbc02d' },
      2: { labelInLegend: '両方', color: '#26a69a' },
      3: { labelInLegend: '合計', color: '#546e7a' },
    })
    .setPosition(dayChartPosition.row, dayChartPosition.col, 0, 0);
  summarySheet.insertChart(chart1Builder.build());

  const monthDataEndRow = 1 + Math.max(monthRows.length, 1);
  const monthRange = summarySheet.getRange(1, monthStartCol, monthDataEndRow, monthHeader.length);
  const monthChartPosition = advanceChartPosition('event');
  const chart2Builder = summarySheet
    .newChart()
    .asColumnChart()
    .addRange(monthRange)
    .setOption('title', '月別件数')
    .setOption('legend', { position: 'top' })
    .setOption('height', 280)
    .setOption('series', {
      0: { labelInLegend: 'うんち', color: '#8d6e63' },
      1: { labelInLegend: 'しっこ', color: '#fbc02d' },
      2: { labelInLegend: '両方', color: '#26a69a' },
      3: { labelInLegend: '合計', color: '#546e7a' },
    })
    .setPosition(monthChartPosition.row, monthChartPosition.col, 0, 0);
  summarySheet.insertChart(chart2Builder.build());

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
    const milkDayChartPosition = advanceChartPosition('milk');
    const chart4Builder = summarySheet
      .newChart()
      .asColumnChart()
      .addRange(milkDayRange)
      .setOption('title', 'ミルク日別実績（ml）')
      .setOption('legend', { position: 'none' })
      .setOption('hAxis', { slantedText: true })
      .setOption('height', 320)
      .setPosition(milkDayChartPosition.row, milkDayChartPosition.col, 0, 0);
    summarySheet.insertChart(chart4Builder.build());
  }

  if (milkMonthRows.length) {
    const milkMonthEndRow = milkMonthStartRow + Math.max(milkMonthRows.length, 1);
    const milkMonthRange = summarySheet.getRange(
      milkMonthStartRow,
      milkStartCol,
      milkMonthEndRow - milkMonthStartRow + 1,
      milkMonthHeader.length
    );
    const milkMonthChartPosition = advanceChartPosition('milk');
    const chart5Builder = summarySheet
      .newChart()
      .asColumnChart()
      .addRange(milkMonthRange)
      .setOption('title', 'ミルク月別実績（ml）')
      .setOption('legend', { position: 'none' })
      .setOption('height', 280)
      .setPosition(milkMonthChartPosition.row, milkMonthChartPosition.col, 0, 0);
    summarySheet.insertChart(chart5Builder.build());
  }

  const hourLabels = Array.from({ length: 24 }, (_, hour) => `${hour}:00`);
  const preferredHeatmapOrder = ['うんち', 'しっこ', '両方', CATEGORY_MILK];
  const categoriesForHeatmap = Object.keys(hourMapByCategory)
    .filter(key => hourMapByCategory[key].some(count => count > 0))
    .sort((a, b) => {
      const indexA = preferredHeatmapOrder.indexOf(a);
      const indexB = preferredHeatmapOrder.indexOf(b);
      if (indexA === -1 && indexB === -1) return a.localeCompare(b);
      if (indexA === -1) return 1;
      if (indexB === -1) return -1;
      return indexA - indexB;
    });

  const heatmapHeader = ['カテゴリ/時間帯', ...hourLabels];
  const heatmapRows = categoriesForHeatmap.map(categoryKey => [
    categoryKey,
    ...hourLabels.map((_, hourIndex) => hourMapByCategory[categoryKey][hourIndex] || 0),
  ]);

  if (heatmapRows.length) {
    const heatmapStartRow = pieStartRow;
    const heatmapStartCol = chartBaseCol + 12;
    summarySheet
      .getRange(heatmapStartRow, heatmapStartCol, 1, heatmapHeader.length)
      .setValues([heatmapHeader]);
    summarySheet
      .getRange(heatmapStartRow + 1, heatmapStartCol, heatmapRows.length, heatmapHeader.length)
      .setValues(heatmapRows);

    const heatmapLastCol = heatmapStartCol + heatmapHeader.length - 1;
    lastUsedCol = Math.max(lastUsedCol, heatmapLastCol);

    const heatmapChartRange = summarySheet.getRange(
      heatmapStartRow,
      heatmapStartCol,
      heatmapRows.length + 1,
      heatmapHeader.length
    );
    const heatmapChart = summarySheet
      .newChart()
      .setChartType(Charts.ChartType.HEATMAP)
      .addRange(heatmapChartRange)
      .setOption('title', 'カテゴリ別時間帯ヒートマップ')
      .setOption('colorAxis', { colors: ['#e8f5e9', '#1b5e20'] })
      .setPosition(heatmapStartRow + heatmapRows.length + 2, heatmapStartCol, 0, 0)
      .build();
    summarySheet.insertChart(heatmapChart);

    const maxHeatmapValue = heatmapRows.reduce((max, row) => {
      const rowMax = Math.max(...row.slice(1));
      return Math.max(max, rowMax);
    }, 0);
    if (maxHeatmapValue > 0) {
      const heatmapDataRange = summarySheet.getRange(
        heatmapStartRow + 1,
        heatmapStartCol + 1,
        heatmapRows.length,
        hourLabels.length
      );
      const gradientRule = SpreadsheetApp.newConditionalFormatRule()
        .setGradientMinpointWithValue('#f1f8e9', SpreadsheetApp.InterpolationType.NUMBER, '0')
        .setGradientMidpointWithValue('#aed581', SpreadsheetApp.InterpolationType.PERCENTILE, '50')
        .setGradientMaxpointWithValue('#1b5e20', SpreadsheetApp.InterpolationType.NUMBER, String(maxHeatmapValue))
        .setRanges([heatmapDataRange])
        .build();
      summarySheet.setConditionalFormatRules([gradientRule]);
    }
  }

  autoResizeAllColumns_(summarySheet, lastUsedCol);

  summarySheet.getRange(1, 1, 1, dayHeader.length).setFontWeight('bold');
  summarySheet.getRange(1, monthStartCol, 1, monthHeader.length).setFontWeight('bold');
  summarySheet.getRange(1, milkStartCol, 1, milkDayHeader.length).setFontWeight('bold');
  summarySheet.getRange(milkMonthStartRow, milkStartCol, 1, milkMonthHeader.length).setFontWeight('bold');

  logInfo('aggregateAndChart: 集計とグラフの更新が完了しました。');
}
