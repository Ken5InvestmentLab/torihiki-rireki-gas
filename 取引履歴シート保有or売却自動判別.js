function handleTransactionHistoryEdit(e) {
  // 基本チェック（高速化のため最低限に）
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  if (sheet.getName() !== "取引履歴") return;

  const startRow = e.range.getRow();
  const numRows = e.range.getNumRows();
  const lastRow = sheet.getLastRow();

  // ヘッダー行(1行目)のみ、またはデータ範囲外の編集なら終了
  if (startRow === 1 && numRows === 1) return;
  if (startRow > lastRow) return;

  // 範囲調整（実際にデータがある行までに制限する処理を入れるとより安全ですが、今回は速度重視でそのまま）
  // A列(1) 〜 H列(8) のデータを一括取得
  const dataRange = sheet.getRange(startRow, 1, numRows, 8);
  
  // getValuesではなくgetDisplayValuesを使う（見た目通りの空判定をするため）
  const values = dataRange.getDisplayValues();
  const newStatuses = [];
  let isChanged = false;

  for (let i = 0; i < numRows; i++) {
    // ヘッダー行(1行目)は無視して、元の値を保持（あるいはスキップ）
    if (startRow + i === 1) {
      newStatuses.push([values[i][1]]);
      continue;
    }

    const rowVal = values[i];
    const valA = rowVal[0]; // A列
    const currentB = rowVal[1]; // B列
    const valH = rowVal[7]; // H列

    let newB = "";

    // 判定ロジック
    if (valH !== "") {
      newB = "売却済";
    } else if (valA !== "") {
      newB = "保有中";
    } else {
      newB = "";
    }

    // 変更判定
    if (newB !== currentB) {
      isChanged = true;
    }
    newStatuses.push([newB]);
  }

  // 変更がある場合のみ書き込み
  if (isChanged) {
    // 書き込み範囲をB列に限定して一括更新
    sheet.getRange(startRow, 2, newStatuses.length, 1).setValues(newStatuses);
  }
}