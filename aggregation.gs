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
  const timezone = SETTINGS.TIMEZONE || Session.getScriptTimeZone();
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
    let dateKey = '';
    const rawDate = row[COL.Date];
    if (rawDate instanceof Date && !Number.isNaN(rawDate.getTime())) {
      dateKey = Utilities.formatDate(rawDate, timezone, 'yyyy-MM-dd');
    } else {
      dateKey = String(rawDate || '').trim();
    }
    if (!dateKey) return;

    if (!mapByDate[dateKey]) mapByDate[dateKey] = { poop: 0, pee: 0, both: 0, total: 0 };
    const startTimeStr =
      startTimeColIndex >= 0 ? String(row[startTimeColIndex] || '').trim() : '';

    let countedInTotal = false;
    if (category === 'うんち') {
      mapByDate[dateKey].poop++;
      countedInTotal = true;
      registerHeatmapCount(category, startTimeStr);
    } else if (category === 'しっこ') {
      mapByDate[dateKey].pee++;
      countedInTotal = true;
      registerHeatmapCount(category, startTimeStr);
    } else if (category === '両方') {
      mapByDate[dateKey].both++;
      countedInTotal = true;
      registerHeatmapCount(category, startTimeStr);
    } else if (category === CATEGORY_MILK && milkColIndex >= 0) {
      const rawAmount = row[milkColIndex];
      const amount = typeof rawAmount === 'number' ? rawAmount : Number(rawAmount) || 0;
      if (!mapMilkByDate[dateKey]) mapMilkByDate[dateKey] = { amount: 0, count: 0 };
      mapMilkByDate[dateKey].amount += amount;
      mapMilkByDate[dateKey].count += 1;
      registerHeatmapCount(category, startTimeStr);
    } else return;

    if (countedInTotal) {
      mapByDate[dateKey].total++;
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

  const parseDateForSheet = value => {
    if (!value) return '';
    if (value instanceof Date && !Number.isNaN(value.getTime())) {
      return value;
    }
    const rawStr = String(value).trim();
    if (!rawStr) return '';
    const normalized = rawStr.replace(/[/.]/g, '-');
    const parts = normalized.split('-');
    if (parts.length < 3) return rawStr;
    const [yearStr, monthStr, dayStr] = parts;
    const year = Number(yearStr);
    const month = Number(monthStr);
    const day = Number(dayStr);
    if ([year, month, day].some(num => Number.isNaN(num))) {
      return rawStr;
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

  const summaryLastCol = Math.max(
    monthStartCol + monthHeader.length - 1,
    milkStartCol + milkDayHeader.length - 1,
    milkStartCol + milkMonthHeader.length - 1
  );
  autoResizeAllColumns_(summarySheet, summaryLastCol);

  const resetChartSheet = sheetName => {
    const sheet = getOrCreateSheet_(ss, sheetName);
    sheet.clear();
    sheet.getCharts().forEach(chart => sheet.removeChart(chart));
    sheet.setConditionalFormatRules([]);
    sheet.setFrozenRows(0);
    return sheet;
  };

  const dayChartSheet = resetChartSheet('chart_day_events');
  const recentDayCount = 30;
  dayChartSheet.getRange(1, 1, 1, dayHeader.length).setValues([dayHeader]);
  if (dayRows.length) {
    const dayChartRows = dayRows.slice(-recentDayCount);
    dayChartSheet.getRange(2, 1, dayChartRows.length, dayHeader.length).setValues(dayChartRows);
    dayChartSheet.getRange(2, 1, dayChartRows.length, 1).setNumberFormat('M/d');
    const dayChartRange = dayChartSheet.getRange(1, 1, dayChartRows.length + 1, dayHeader.length);
    const dayChartBuilder = dayChartSheet
      .newChart()
      .asColumnChart()
      .addRange(dayChartRange)
      .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
      .setStacked()
      .setOption('title', '日別件数（直近30日・積み上げ）')
      .setOption('legend', { position: 'top' })
      .setOption('hAxis', { slantedText: true, format: 'M/d' })
      .setOption('height', 320)
      .setOption('series', {
        0: { labelInLegend: 'うんち', color: '#8d6e63' },
        1: { labelInLegend: 'しっこ', color: '#fbc02d' },
        2: { labelInLegend: '両方', color: '#26a69a' },
        3: { labelInLegend: '合計', color: '#546e7a' },
      })
      .setPosition(1, dayHeader.length + 2, 0, 0);
    dayChartSheet.insertChart(dayChartBuilder.build());
    dayChartSheet.setFrozenRows(1);
    autoResizeAllColumns_(dayChartSheet, dayHeader.length);
  }

  const monthChartSheet = resetChartSheet('chart_month_events');
  monthChartSheet.getRange(1, 1, 1, monthHeader.length).setValues([monthHeader]);
  if (monthRows.length) {
    monthChartSheet.getRange(2, 1, monthRows.length, monthHeader.length).setValues(monthRows);
    const monthChartRange = monthChartSheet.getRange(1, 1, monthRows.length + 1, monthHeader.length);
    const monthChartBuilder = monthChartSheet
      .newChart()
      .asColumnChart()
      .addRange(monthChartRange)
      .setOption('title', '月別件数')
      .setOption('legend', { position: 'top' })
      .setOption('height', 280)
      .setOption('series', {
        0: { labelInLegend: 'うんち', color: '#8d6e63' },
        1: { labelInLegend: 'しっこ', color: '#fbc02d' },
        2: { labelInLegend: '両方', color: '#26a69a' },
        3: { labelInLegend: '合計', color: '#546e7a' },
      })
      .setPosition(1, monthHeader.length + 2, 0, 0);
    monthChartSheet.insertChart(monthChartBuilder.build());
    monthChartSheet.setFrozenRows(1);
    autoResizeAllColumns_(monthChartSheet, monthHeader.length);
  }

  const totalPoop = dayRows.reduce((acc, row) => acc + row[1], 0);
  const totalPee = dayRows.reduce((acc, row) => acc + row[2], 0);
  const totalBoth = dayRows.reduce((acc, row) => acc + row[3], 0);
  const pieTable = [
    ['カテゴリ', '件数'],
    ['うんち', totalPoop],
    ['しっこ', totalPee],
    ['両方', totalBoth],
  ];
  const pieChartSheet = resetChartSheet('chart_category_breakdown');
  pieChartSheet.getRange(1, 1, pieTable.length, pieTable[0].length).setValues(pieTable);
  const pieChart = pieChartSheet
    .newChart()
    .asPieChart()
    .addRange(pieChartSheet.getRange(1, 1, pieTable.length, pieTable[0].length))
    .setOption('title', 'カテゴリ内訳（期間合計）')
    .setPosition(1, pieTable[0].length + 2, 0, 0)
    .build();
  pieChartSheet.insertChart(pieChart);
  pieChartSheet.setFrozenRows(1);
  autoResizeAllColumns_(pieChartSheet, pieTable[0].length);

  if (milkDayRows.length) {
    const milkDayChartSheet = resetChartSheet('chart_milk_daily');
    milkDayChartSheet.getRange(1, 1, 1, milkDayHeader.length).setValues([milkDayHeader]);
    milkDayChartSheet
      .getRange(2, 1, milkDayRows.length, milkDayHeader.length)
      .setValues(milkDayRows);
    milkDayChartSheet.getRange(2, 1, milkDayRows.length, 1).setNumberFormat('M/d');
    const milkDayChartRange = milkDayChartSheet.getRange(
      1,
      1,
      milkDayRows.length + 1,
      milkDayHeader.length
    );
    const milkDayChartBuilder = milkDayChartSheet
      .newChart()
      .asColumnChart()
      .addRange(milkDayChartRange)
      .setOption('title', 'ミルク日別実績（ml）')
      .setOption('legend', { position: 'none' })
      .setOption('hAxis', { slantedText: true, format: 'M/d' })
      .setOption('height', 320)
      .setPosition(1, milkDayHeader.length + 2, 0, 0);
    milkDayChartSheet.insertChart(milkDayChartBuilder.build());
    milkDayChartSheet.setFrozenRows(1);
    autoResizeAllColumns_(milkDayChartSheet, milkDayHeader.length);
  }

  if (milkMonthRows.length) {
    const milkMonthChartSheet = resetChartSheet('chart_milk_monthly');
    milkMonthChartSheet.getRange(1, 1, 1, milkMonthHeader.length).setValues([milkMonthHeader]);
    milkMonthChartSheet
      .getRange(2, 1, milkMonthRows.length, milkMonthHeader.length)
      .setValues(milkMonthRows);
    const milkMonthChartRange = milkMonthChartSheet.getRange(
      1,
      1,
      milkMonthRows.length + 1,
      milkMonthHeader.length
    );
    const milkMonthChartBuilder = milkMonthChartSheet
      .newChart()
      .asColumnChart()
      .addRange(milkMonthChartRange)
      .setOption('title', 'ミルク月別実績（ml）')
      .setOption('legend', { position: 'none' })
      .setOption('height', 280)
      .setPosition(1, milkMonthHeader.length + 2, 0, 0);
    milkMonthChartSheet.insertChart(milkMonthChartBuilder.build());
    milkMonthChartSheet.setFrozenRows(1);
    autoResizeAllColumns_(milkMonthChartSheet, milkMonthHeader.length);
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
    const heatmapSheet = resetChartSheet('chart_category_heatmap');
    heatmapSheet
      .getRange(1, 1, 1, heatmapHeader.length)
      .setValues([heatmapHeader]);
    heatmapSheet
      .getRange(2, 1, heatmapRows.length, heatmapHeader.length)
      .setValues(heatmapRows);
    heatmapSheet.setFrozenRows(1);
    autoResizeAllColumns_(heatmapSheet, heatmapHeader.length);

    const heatmapChartRange = heatmapSheet.getRange(
      1,
      1,
      heatmapRows.length + 1,
      heatmapHeader.length
    );
    const heatmapChart = heatmapSheet
      .newChart()
      .setChartType(Charts.ChartType.HEATMAP)
      .addRange(heatmapChartRange)
      .setOption('title', 'カテゴリ別時間帯ヒートマップ')
      .setOption('colorAxis', { colors: ['#e8f5e9', '#1b5e20'] })
      .setPosition(1, heatmapHeader.length + 2, 0, 0)
      .build();
    heatmapSheet.insertChart(heatmapChart);

    const maxHeatmapValue = heatmapRows.reduce((max, row) => {
      const rowMax = Math.max(...row.slice(1));
      return Math.max(max, rowMax);
    }, 0);
    if (maxHeatmapValue > 0) {
      const heatmapDataRange = heatmapSheet.getRange(2, 2, heatmapRows.length, hourLabels.length);
      const gradientRule = SpreadsheetApp.newConditionalFormatRule()
        .setGradientMinpointWithValue('#f1f8e9', SpreadsheetApp.InterpolationType.NUMBER, '0')
        .setGradientMidpointWithValue('#aed581', SpreadsheetApp.InterpolationType.PERCENTILE, '50')
        .setGradientMaxpointWithValue('#1b5e20', SpreadsheetApp.InterpolationType.NUMBER, String(maxHeatmapValue))
        .setRanges([heatmapDataRange])
        .build();
      heatmapSheet.setConditionalFormatRules([gradientRule]);
    }
  }

  summarySheet.getRange(1, 1, 1, dayHeader.length).setFontWeight('bold');
  summarySheet.getRange(1, monthStartCol, 1, monthHeader.length).setFontWeight('bold');
  summarySheet.getRange(1, milkStartCol, 1, milkDayHeader.length).setFontWeight('bold');
  summarySheet.getRange(milkMonthStartRow, milkStartCol, 1, milkMonthHeader.length).setFontWeight('bold');

  logInfo('aggregateAndChart: 集計とグラフの更新が完了しました。');
}
