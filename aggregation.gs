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

  // 全体サマリーを最初に追加
  const summaryStartRow = 1;
  summarySheet.getRange(summaryStartRow, 1).setValue('📊 Baby Logs サマリー');
  summarySheet.getRange(summaryStartRow, 1).setFontSize(14);
  summarySheet.getRange(summaryStartRow, 1).setFontWeight('bold');
  summarySheet.getRange(summaryStartRow, 1).setBackground('#bbdefb');

  summarySheet.getRange(summaryStartRow + 1, 1).setValue('データ期間：');
  const sortedDateKeys = Object.keys(mapByDate).sort();
  summarySheet.getRange(summaryStartRow + 1, 2).setValue(
    sortedDateKeys.length > 0
      ? `${sortedDateKeys[0]} ～ ${sortedDateKeys[sortedDateKeys.length - 1]}`
      : 'データなし'
  );

  summarySheet.getRange(summaryStartRow + 2, 1).setValue('総記録日数：');
  summarySheet.getRange(summaryStartRow + 2, 2).setValue(sortedDateKeys.length);

  summarySheet.getRange(summaryStartRow + 3, 1).setValue('総イベント数：');
  const totalEvents = Object.values(mapByDate).reduce((sum, v) => sum + v.total, 0);
  summarySheet.getRange(summaryStartRow + 3, 2).setValue(totalEvents);

  if (Object.keys(mapMilkByDate).length > 0) {
    const totalMilk = Object.values(mapMilkByDate).reduce((sum, v) => sum + v.amount, 0);
    const totalMilkCount = Object.values(mapMilkByDate).reduce((sum, v) => sum + v.count, 0);
    summarySheet.getRange(summaryStartRow + 4, 1).setValue('総ミルク量：');
    summarySheet.getRange(summaryStartRow + 4, 2).setValue(`${Math.round(totalMilk * 10) / 10} ml`);
    summarySheet.getRange(summaryStartRow + 5, 1).setValue('総授乳回数：');
    summarySheet.getRange(summaryStartRow + 5, 2).setValue(`${totalMilkCount} 回`);
  }

  summarySheet.getRange(summaryStartRow, 1, 6, 1).setFontWeight('bold');
  summarySheet.getRange(summaryStartRow, 1, 6, 2).setBorder(true, true, true, true, true, true);

  const dataStartRow = summaryStartRow + 8;

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

  summarySheet.getRange(dataStartRow, 1, 1, dayHeader.length).setValues([dayHeader]);
  if (dayRows.length) {
    const dayRange = summarySheet.getRange(dataStartRow + 1, 1, dayRows.length, dayHeader.length);
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
  summarySheet.getRange(dataStartRow, monthStartCol, 1, monthHeader.length).setValues([monthHeader]);
  if (monthRows.length) {
    summarySheet.getRange(dataStartRow + 1, monthStartCol, monthRows.length, monthHeader.length).setValues(monthRows);
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

  summarySheet.getRange(dataStartRow, milkStartCol, 1, milkDayHeader.length).setValues([milkDayHeader]);
  if (milkDayRows.length) {
    const milkDayRange = summarySheet.getRange(dataStartRow + 1, milkStartCol, milkDayRows.length, milkDayHeader.length);
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

  const milkMonthStartRow = Math.max(dataStartRow + 2, dataStartRow + milkDayRows.length + 2);
  summarySheet.getRange(milkMonthStartRow, milkStartCol, 1, milkMonthHeader.length).setValues([milkMonthHeader]);
  if (milkMonthRows.length) {
    summarySheet.getRange(milkMonthStartRow + 1, milkStartCol, milkMonthRows.length, milkMonthHeader.length).setValues(
      milkMonthRows
    );
  }

  summarySheet.setFrozenRows(dataStartRow);

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

    // 分析情報を追加
    const calculateAverage = (rows, index) => {
      if (rows.length === 0) return 0;
      const total = rows.reduce((sum, row) => sum + row[index], 0);
      return Math.round((total / rows.length) * 10) / 10;
    };
    const avgPoop = calculateAverage(dayChartRows, 1);
    const avgPee = calculateAverage(dayChartRows, 2);
    const avgBoth = calculateAverage(dayChartRows, 3);
    const avgTotal = calculateAverage(dayChartRows, 4);

    const statsStartRow = dayChartRows.length + 3;
    dayChartSheet.getRange(statsStartRow, 1).setValue('平均（1日あたり）');
    dayChartSheet.getRange(statsStartRow, 2).setValue(avgPoop);
    dayChartSheet.getRange(statsStartRow, 3).setValue(avgPee);
    dayChartSheet.getRange(statsStartRow, 4).setValue(avgBoth);
    dayChartSheet.getRange(statsStartRow, 5).setValue(avgTotal);
    dayChartSheet.getRange(statsStartRow, 1, 1, dayHeader.length).setFontWeight('bold');
    dayChartSheet.getRange(statsStartRow, 1, 1, dayHeader.length).setBackground('#e8f5e9');

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
      .setOption('vAxis', { title: '件数' })
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

    // 分析情報を追加
    const avgMonthPoop = Math.round(monthRows.reduce((sum, row) => sum + row[1], 0) / monthRows.length * 10) / 10;
    const avgMonthPee = Math.round(monthRows.reduce((sum, row) => sum + row[2], 0) / monthRows.length * 10) / 10;
    const avgMonthBoth = Math.round(monthRows.reduce((sum, row) => sum + row[3], 0) / monthRows.length * 10) / 10;
    const avgMonthTotal = Math.round(monthRows.reduce((sum, row) => sum + row[4], 0) / monthRows.length * 10) / 10;

    const statsStartRow = monthRows.length + 3;
    monthChartSheet.getRange(statsStartRow, 1).setValue('平均（月あたり）');
    monthChartSheet.getRange(statsStartRow, 2).setValue(avgMonthPoop);
    monthChartSheet.getRange(statsStartRow, 3).setValue(avgMonthPee);
    monthChartSheet.getRange(statsStartRow, 4).setValue(avgMonthBoth);
    monthChartSheet.getRange(statsStartRow, 5).setValue(avgMonthTotal);
    monthChartSheet.getRange(statsStartRow, 1, 1, monthHeader.length).setFontWeight('bold');
    monthChartSheet.getRange(statsStartRow, 1, 1, monthHeader.length).setBackground('#e8f5e9');

    const monthChartRange = monthChartSheet.getRange(1, 1, monthRows.length + 1, monthHeader.length);
    const monthChartBuilder = monthChartSheet
      .newChart()
      .asColumnChart()
      .addRange(monthChartRange)
      .setOption('title', '月別件数')
      .setOption('legend', { position: 'top' })
      .setOption('vAxis', { title: '件数' })
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
  const totalAll = totalPoop + totalPee + totalBoth;

  const pieTable = [
    ['カテゴリ', '件数', '割合(%)'],
    ['うんち', totalPoop, totalAll > 0 ? Math.round(totalPoop / totalAll * 1000) / 10 : 0],
    ['しっこ', totalPee, totalAll > 0 ? Math.round(totalPee / totalAll * 1000) / 10 : 0],
    ['両方', totalBoth, totalAll > 0 ? Math.round(totalBoth / totalAll * 1000) / 10 : 0],
  ];
  const pieChartSheet = resetChartSheet('chart_category_breakdown');
  pieChartSheet.getRange(1, 1, pieTable.length, pieTable[0].length).setValues(pieTable);

  // サマリー情報を追加
  pieChartSheet.getRange(pieTable.length + 2, 1).setValue('期間合計');
  pieChartSheet.getRange(pieTable.length + 2, 2).setValue(totalAll);
  pieChartSheet.getRange(pieTable.length + 2, 1, 1, 2).setFontWeight('bold');
  pieChartSheet.getRange(pieTable.length + 2, 1, 1, 2).setBackground('#fff3e0');

  const pieChart = pieChartSheet
    .newChart()
    .asPieChart()
    .addRange(pieChartSheet.getRange(1, 1, pieTable.length, 2))
    .setOption('title', 'カテゴリ内訳（期間合計）')
    .setOption('pieSliceText', 'percentage')
    .setOption('slices', {
      0: { color: '#8d6e63' }, // うんち
      1: { color: '#fbc02d' }, // しっこ
      2: { color: '#26a69a' }, // 両方
    })
    .setPosition(1, pieTable[0].length + 2, 0, 0)
    .build();
  pieChartSheet.insertChart(pieChart);
  pieChartSheet.setFrozenRows(1);
  autoResizeAllColumns_(pieChartSheet, pieTable[0].length);

  if (milkDayRows.length) {
    // ミルク量のグラフ（ml）
    const milkAmountChartSheet = resetChartSheet('chart_milk_amount');
    const milkAmountHeader = ['日付', 'ミルク量(ml)'];
    const milkAmountRows = milkDayRows.map(row => [row[0], row[1]]);
    milkAmountChartSheet.getRange(1, 1, 1, milkAmountHeader.length).setValues([milkAmountHeader]);
    milkAmountChartSheet
      .getRange(2, 1, milkAmountRows.length, milkAmountHeader.length)
      .setValues(milkAmountRows);
    milkAmountChartSheet.getRange(2, 1, milkAmountRows.length, 1).setNumberFormat('M/d');

    // 平均値を計算
    const totalMilkAmount = milkAmountRows.reduce((sum, row) => sum + row[1], 0);
    const avgMilkAmount = Math.round(totalMilkAmount / milkAmountRows.length * 10) / 10;

    // 平均値を表示
    milkAmountChartSheet.getRange(milkAmountRows.length + 3, 1).setValue('平均ミルク量(ml)');
    milkAmountChartSheet.getRange(milkAmountRows.length + 3, 2).setValue(avgMilkAmount);
    milkAmountChartSheet.getRange(milkAmountRows.length + 3, 1).setFontWeight('bold');

    const milkAmountChartRange = milkAmountChartSheet.getRange(
      1,
      1,
      milkAmountRows.length + 1,
      milkAmountHeader.length
    );
    const milkAmountChartBuilder = milkAmountChartSheet
      .newChart()
      .asColumnChart()
      .addRange(milkAmountChartRange)
      .setOption('title', 'ミルク日別実績（ml）')
      .setOption('legend', { position: 'none' })
      .setOption('hAxis', { slantedText: true, format: 'M/d' })
      .setOption('vAxis', { title: 'ミルク量(ml)' })
      .setOption('height', 320)
      .setOption('colors', ['#42A5F5'])
      .setPosition(1, milkAmountHeader.length + 2, 0, 0);
    milkAmountChartSheet.insertChart(milkAmountChartBuilder.build());
    milkAmountChartSheet.setFrozenRows(1);
    autoResizeAllColumns_(milkAmountChartSheet, milkAmountHeader.length);

    // 授乳回数のグラフ
    const milkCountChartSheet = resetChartSheet('chart_milk_count');
    const milkCountHeader = ['日付', '授乳回数'];
    const milkCountRows = milkDayRows.map(row => [row[0], row[2]]);
    milkCountChartSheet.getRange(1, 1, 1, milkCountHeader.length).setValues([milkCountHeader]);
    milkCountChartSheet
      .getRange(2, 1, milkCountRows.length, milkCountHeader.length)
      .setValues(milkCountRows);
    milkCountChartSheet.getRange(2, 1, milkCountRows.length, 1).setNumberFormat('M/d');

    // 平均値を計算
    const totalMilkCount = milkCountRows.reduce((sum, row) => sum + row[1], 0);
    const avgMilkCount = Math.round(totalMilkCount / milkCountRows.length * 10) / 10;

    // 平均値を表示
    milkCountChartSheet.getRange(milkCountRows.length + 3, 1).setValue('平均授乳回数');
    milkCountChartSheet.getRange(milkCountRows.length + 3, 2).setValue(avgMilkCount);
    milkCountChartSheet.getRange(milkCountRows.length + 3, 1).setFontWeight('bold');

    const milkCountChartRange = milkCountChartSheet.getRange(
      1,
      1,
      milkCountRows.length + 1,
      milkCountHeader.length
    );
    const milkCountChartBuilder = milkCountChartSheet
      .newChart()
      .asColumnChart()
      .addRange(milkCountChartRange)
      .setOption('title', '授乳日別回数')
      .setOption('legend', { position: 'none' })
      .setOption('hAxis', { slantedText: true, format: 'M/d' })
      .setOption('vAxis', { title: '授乳回数' })
      .setOption('height', 320)
      .setOption('colors', ['#66BB6A'])
      .setPosition(1, milkCountHeader.length + 2, 0, 0);
    milkCountChartSheet.insertChart(milkCountChartBuilder.build());
    milkCountChartSheet.setFrozenRows(1);
    autoResizeAllColumns_(milkCountChartSheet, milkCountHeader.length);
  }

  if (milkMonthRows.length) {
    // 月別ミルク量のグラフ（ml）
    const milkMonthAmountChartSheet = resetChartSheet('chart_milk_amount_monthly');
    const milkMonthAmountHeader = ['月', 'ミルク量(ml)'];
    const milkMonthAmountRows = milkMonthRows.map(row => [row[0], row[1]]);
    milkMonthAmountChartSheet.getRange(1, 1, 1, milkMonthAmountHeader.length).setValues([milkMonthAmountHeader]);
    milkMonthAmountChartSheet
      .getRange(2, 1, milkMonthAmountRows.length, milkMonthAmountHeader.length)
      .setValues(milkMonthAmountRows);

    // 平均値を計算
    const totalMonthMilkAmount = milkMonthAmountRows.reduce((sum, row) => sum + row[1], 0);
    const avgMonthMilkAmount = Math.round(totalMonthMilkAmount / milkMonthAmountRows.length * 10) / 10;

    // 平均値を表示
    milkMonthAmountChartSheet.getRange(milkMonthAmountRows.length + 3, 1).setValue('平均月間ミルク量(ml)');
    milkMonthAmountChartSheet.getRange(milkMonthAmountRows.length + 3, 2).setValue(avgMonthMilkAmount);
    milkMonthAmountChartSheet.getRange(milkMonthAmountRows.length + 3, 1).setFontWeight('bold');

    const milkMonthAmountChartRange = milkMonthAmountChartSheet.getRange(
      1,
      1,
      milkMonthAmountRows.length + 1,
      milkMonthAmountHeader.length
    );
    const milkMonthAmountChartBuilder = milkMonthAmountChartSheet
      .newChart()
      .asColumnChart()
      .addRange(milkMonthAmountChartRange)
      .setOption('title', 'ミルク月別実績（ml）')
      .setOption('legend', { position: 'none' })
      .setOption('vAxis', { title: 'ミルク量(ml)' })
      .setOption('height', 280)
      .setOption('colors', ['#42A5F5'])
      .setPosition(1, milkMonthAmountHeader.length + 2, 0, 0);
    milkMonthAmountChartSheet.insertChart(milkMonthAmountChartBuilder.build());
    milkMonthAmountChartSheet.setFrozenRows(1);
    autoResizeAllColumns_(milkMonthAmountChartSheet, milkMonthAmountHeader.length);

    // 月別授乳回数のグラフ
    const milkMonthCountChartSheet = resetChartSheet('chart_milk_count_monthly');
    const milkMonthCountHeader = ['月', '授乳回数'];
    const milkMonthCountRows = milkMonthRows.map(row => [row[0], row[2]]);
    milkMonthCountChartSheet.getRange(1, 1, 1, milkMonthCountHeader.length).setValues([milkMonthCountHeader]);
    milkMonthCountChartSheet
      .getRange(2, 1, milkMonthCountRows.length, milkMonthCountHeader.length)
      .setValues(milkMonthCountRows);

    // 平均値を計算
    const totalMonthMilkCount = milkMonthCountRows.reduce((sum, row) => sum + row[1], 0);
    const avgMonthMilkCount = Math.round(totalMonthMilkCount / milkMonthCountRows.length * 10) / 10;

    // 平均値を表示
    milkMonthCountChartSheet.getRange(milkMonthCountRows.length + 3, 1).setValue('平均月間授乳回数');
    milkMonthCountChartSheet.getRange(milkMonthCountRows.length + 3, 2).setValue(avgMonthMilkCount);
    milkMonthCountChartSheet.getRange(milkMonthCountRows.length + 3, 1).setFontWeight('bold');

    const milkMonthCountChartRange = milkMonthCountChartSheet.getRange(
      1,
      1,
      milkMonthCountRows.length + 1,
      milkMonthCountHeader.length
    );
    const milkMonthCountChartBuilder = milkMonthCountChartSheet
      .newChart()
      .asColumnChart()
      .addRange(milkMonthCountChartRange)
      .setOption('title', '授乳月別回数')
      .setOption('legend', { position: 'none' })
      .setOption('vAxis', { title: '授乳回数' })
      .setOption('height', 280)
      .setOption('colors', ['#66BB6A'])
      .setPosition(1, milkMonthCountHeader.length + 2, 0, 0);
    milkMonthCountChartSheet.insertChart(milkMonthCountChartBuilder.build());
    milkMonthCountChartSheet.setFrozenRows(1);
    autoResizeAllColumns_(milkMonthCountChartSheet, milkMonthCountHeader.length);
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

  summarySheet.getRange(dataStartRow, 1, 1, dayHeader.length).setFontWeight('bold');
  summarySheet.getRange(dataStartRow, 1, 1, dayHeader.length).setBackground('#c8e6c9');
  summarySheet.getRange(dataStartRow, monthStartCol, 1, monthHeader.length).setFontWeight('bold');
  summarySheet.getRange(dataStartRow, monthStartCol, 1, monthHeader.length).setBackground('#c8e6c9');
  summarySheet.getRange(dataStartRow, milkStartCol, 1, milkDayHeader.length).setFontWeight('bold');
  summarySheet.getRange(dataStartRow, milkStartCol, 1, milkDayHeader.length).setBackground('#e1f5fe');
  summarySheet.getRange(milkMonthStartRow, milkStartCol, 1, milkMonthHeader.length).setFontWeight('bold');
  summarySheet.getRange(milkMonthStartRow, milkStartCol, 1, milkMonthHeader.length).setBackground('#e1f5fe');

  logInfo('aggregateAndChart: 集計とグラフの更新が完了しました。');
}
