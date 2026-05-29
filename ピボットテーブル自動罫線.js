/**
 * スプレッドシートを開いたときにカスタムメニューを追加します。
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('ピボットテーブル操作')
    .addItem('データのあるセルに罫線を設定 (A5:Z)', 'applyBordersToFixedRange')
    .addToUi();
}

/**
 * 指定したシートの固定範囲 (A5:Z) 内で、データのある範囲にのみ罫線を自動設定
 * UI コンテキスト外（トリガー実行）でも動作するよう、alert を排除し Logger に変更
 */
function applyBordersToFixedRange(e) {
  const SHEET_NAME = '資産サマリ';
  const TARGET_RANGE_NOTATION = 'A5:Z';

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    Logger.log(`シート "${SHEET_NAME}" が見つかりませんでした`);
    return;
  }

  const targetRange = sheet.getRange(TARGET_RANGE_NOTATION);
  const values = targetRange.getDisplayValues();
  const numRows = values.length;
  const numCols = values[0].length;
  const startRow = targetRange.getRow();
  const startCol = targetRange.getColumn();

  // 既存の罫線をクリア
  targetRange.setBorder(false, false, false, false, false, false);

  // データ存在チェック
  let lastDataRow = -1;
  let lastDataCol = -1;

  for (let r = numRows - 1; r >= 0; r--) {
    if (values[r].some(cell => cell !== '')) {
      lastDataRow = r;
      break;
    }
  }

  if (lastDataRow !== -1) {
    for (let c = numCols - 1; c >= 0; c--) {
      if (values.some(row => row[c] !== '')) {
        lastDataCol = c;
        break;
      }
    }
  }

  if (lastDataRow !== -1 && lastDataCol !== -1) {
    const dataRange = sheet.getRange(
      startRow,
      startCol,
      lastDataRow + 1,
      lastDataCol + 1
    );

    // 罫線スタイル設定
    const color = '#808080';
    const style = SpreadsheetApp.BorderStyle.SOLID;

    dataRange.setBorder(
      true, true, true, true, true, true,
      color, style
    );
  } else {
    Logger.log('指定範囲内にデータが見つかりませんでした');
  }
}
