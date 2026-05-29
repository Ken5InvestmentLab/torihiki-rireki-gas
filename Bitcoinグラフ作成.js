const SUMMARY_SHEET_NAME = 'Bitcoinサマリ';
const HISTORY_SHEET_NAME = 'Bitcoin履歴_システム用';

function recordDailyBitcoinData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const summarySheet = ss.getSheetByName(SUMMARY_SHEET_NAME);

  if (!summarySheet) return;

  const totalInvestment = summarySheet.getRange('C2').getValue();
  const currentValuation = summarySheet.getRange('D2').getValue();
  
  // 時刻を00:00:00に固定する
  const now = new Date();
  const timestamp = new Date(now.getFullYear(), now.getMonth(), now.getDate()); 

  let historySheet = ss.getSheetByName(HISTORY_SHEET_NAME);
  if (!historySheet) {
    historySheet = ss.insertSheet(HISTORY_SHEET_NAME);
    historySheet.appendRow(['日時', '合計投資額 (円)', '現在の評価額 (円)']);
    historySheet.hideSheet();
  }

  historySheet.appendRow([timestamp, totalInvestment, currentValuation]);
  updateChart(summarySheet, historySheet);
}

function updateChart(summarySheet, historySheet) {
  const chartTitle = '投資額と評価額の推移';
  
  const charts = summarySheet.getCharts();
  charts.forEach(chart => {
    if (chart.getOptions().get('title') === chartTitle) {
      summarySheet.removeChart(chart);
    }
  });

  const lastRow = historySheet.getLastRow();
  const dataRange = historySheet.getRange(1, 1, lastRow, 3);

  // Y軸の最大値に20%の余裕を持たせる計算
  let maxValue = 0;
  if (lastRow > 1) {
    const dataValues = historySheet.getRange(2, 2, lastRow - 1, 2).getValues();
    dataValues.forEach(row => {
      const val1 = Number(row[0]) || 0;
      const val2 = Number(row[1]) || 0;
      maxValue = Math.max(maxValue, val1, val2);
    });
  }
  const yAxisMax = maxValue > 0 ? maxValue * 1.2 : 5000;

  const chartBuilder = summarySheet.newChart()
    .asLineChart()
    .addRange(dataRange)
    .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
    .setNumHeaders(1)
    .setPosition(7, 5, 0, 0) 
    .setOption('width', getChartWidth(summarySheet, 5, 8))
    .setOption('height', getChartHeight(summarySheet, 7, 24))
    .setOption('title', chartTitle)
    .setOption('useFirstColumnAsDomain', true)
    .setOption('hAxis', {
      title: '日時',
      gridlines: { color: 'none' }, 
      minorGridlines: { color: 'none' }, 
      textStyle: { fontSize: 10 } 
    })
    .setOption('vAxis', {
      viewWindow: { min: 0, max: yAxisMax }
    })
    .setOption('legend', { 
      position: 'bottom', 
      textStyle: { fontSize: 12 } 
    })
    // --- 今回の追加箇所：データポイントのサイズを2に設定 ---
    .setOption('pointSize', 2)
    // ----------------------------------------------------
    .setOption('series', {
      0: { color: '#4285F4', labelInLegend: '合計投資額 (円)' },
      1: { color: '#EA4335', labelInLegend: '現在の評価額 (円)' }
    });

  summarySheet.insertChart(chartBuilder.build());
}

function getChartWidth(sheet, startCol, endCol) {
  let width = 0;
  for (let i = startCol; i <= endCol; i++) {
    width += sheet.getColumnWidth(i);
  }
  return width;
}

function getChartHeight(sheet, startRow, endRow) {
  let height = 0;
  for (let i = startRow; i <= endRow; i++) {
    height += sheet.getRowHeight(i);
  }
  return height;
}

// =================================================================
// 以下、トリガー設定（23:58ピンポイント制）
// =================================================================

function setupDailyTriggerManager() {
  const triggers = ScriptApp.getProjectTriggers();
  
  // --- 追加：セットアップの重複を防ぐ ---
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'createTodaySpecificTimeTrigger') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  // ------------------------------------

  // 毎日0時台に「今日の23:58の予約を作る」ためのトリガーを設定
  ScriptApp.newTrigger('createTodaySpecificTimeTrigger')
    .timeBased()
    .everyDays(1)
    .atHour(0) 
    .create();

  // 初回（今日分）の予約を今すぐ作成
  createTodaySpecificTimeTrigger();
  
  Browser.msgBox("自動実行のセットアップが完了しました。毎日23:58にデータが記録されます。");
}

function createTodaySpecificTimeTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  
  // 過去にセットされた「23:58の記録トリガー」だけを安全に削除
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'recordDailyBitcoinData') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  const now = new Date();
  // 今日の23:58:00を指定
  const scheduledTime = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 23, 58, 0);

  // もし既に今日の23:58を過ぎていたら、明日の23:58に設定
  if (scheduledTime.getTime() < now.getTime()) {
    scheduledTime.setDate(scheduledTime.getDate() + 1);
  }

  // 特定の日時に1回だけ実行されるトリガーをセット
  ScriptApp.newTrigger('recordDailyBitcoinData')
    .timeBased()
    .at(scheduledTime)
    .create();
}