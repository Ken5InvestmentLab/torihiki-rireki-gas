function pasteValuesOnSaturday() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("需給分析");

  if (!sheet) {
    Logger.log("シートが見つかりません。");
    return;
  }

  const range = sheet.getDataRange();
  const values = range.getValues();
  const lastRow = range.getLastRow();

  const maxWaitTimeMs = 60000; // 最大1分（60000ミリ秒）待つ
  const checkIntervalMs = 10000; // 10秒ごとにチェック
  const maxRetries = Math.floor(maxWaitTimeMs / checkIntervalMs);

  Logger.log('土曜日の貼り付け処理を開始します。');

  for (let i = 1; i < lastRow; i++) {
    const row = i + 1;
    const bValue = values[i][1]; // B列
    const lValue = values[i][11]; // L列
    const mCell = sheet.getRange(row, 13); // M列
    const lCell = sheet.getRange(row, 12); // L列

    if (bValue !== "" && lValue === "") {
      Logger.log(`行 ${row}: M列の値の取得を開始`);

      let retries = 0;
      let mValue = mCell.getValue();

      while (!(typeof mValue === 'number' && !isNaN(mValue)) && retries < maxRetries) {
        Logger.log(`行 ${row}: M列の値がまだ取得できません (${mValue})。リフレッシュします。(${retries + 1}回目)`);

        SpreadsheetApp.flush(); // セル再計算の強制
        Utilities.sleep(checkIntervalMs); // 10秒待ってから再取得

        // 数式再設定で再評価を促す
        mCell.setFormula(mCell.getFormula());

        SpreadsheetApp.flush();
        mValue = mCell.getValue();
        retries++;
      }

      if (typeof mValue === 'number' && !isNaN(mValue)) {
        mCell.copyTo(lCell, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
        Logger.log(`行 ${row}: M列の値 (${mValue}) をL列に貼り付けました。`);
      } else {
        Logger.log(`行 ${row}: M列の値が取得できませんでした。最終値: ${mValue}`);
      }
    }
  }

  Logger.log('土曜日の貼り付け処理が終了しました。');
}
