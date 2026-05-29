function createMonthlySummaryChart() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  sheet.getCharts().forEach(c => sheet.removeChart(c));

  const finder = sheet.createTextFinder('月別内訳').findNext();
  if (!finder) {
    SpreadsheetApp.getUi().alert('「月別内訳」が見つかりません。');
    return;
  }
  const startRow = finder.getRow() + 1;
  const monthLabelRange = sheet.getRange(startRow, 1, 12);
  const amountRange = sheet.getRange(startRow, 2, 12);

  // 計算C1:J20の領域サイズ（ピクセル）
  const topRow = 1;
  const leftCol = 3; // C列
  const rightCol = 10; // J列
  const bottomRow = 20;

  // 列幅の合計を取得
  let totalWidth = 0;
  for (let col = leftCol; col <= rightCol; col++) {
    totalWidth += sheet.getColumnWidth(col);
  }
  // 行高の合計を取得
  let totalHeight = 0;
  for (let row = topRow; row <= bottomRow; row++) {
    totalHeight += sheet.getRowHeight(row);
  }

  // 余白を考慮して少し小さめに設定
  const chartWidth = Math.floor(totalWidth * 0.9);
  const chartHeight = Math.floor(totalHeight * 0.85);

  const chart = sheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(amountRange)
    .setPosition(topRow, leftCol, 0, 0) // 左上をC1セルに
    .setOption('title', '📆 月別受取金額（2026年）')
    .setOption('width', chartWidth)
    .setOption('height', chartHeight)
    .setOption('legend', { position: 'none' })
    .setOption('hAxis', {
      title: '月',
      ticks: monthLabelRange.getValues().flat(),
      slantedText: false,
      textStyle: { fontSize: 10 }
    })
    .setOption('vAxis', {
      title: '受取金額 (円)',
      format: '¥#,##0',
      minValue: 0
    })
    .setOption('colors', ['#1a7640'])
    .build();

  sheet.insertChart(chart);
  SpreadsheetApp.getUi().alert('✅ グラフを C1:J20 の範囲内に配置しました。');
}