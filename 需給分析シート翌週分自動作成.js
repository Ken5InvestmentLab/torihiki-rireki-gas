function copyAndShiftToFriday() {
  const SHEET_NAME = '需給分析';
  const TARGET_COLUMNS = [3,6,9,10,11,13,14,15,16,17,18]; // C,F,I,J,K,M,N,O,P,Q,R列（1ベース）
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error(`シート "${SHEET_NAME}" が見つかりません`);

  // データ範囲を取得（ヘッダーが1行目にある場合）
  const startRow = 2;
  const lastRow = sheet.getLastRow();
  if (lastRow < startRow) return;
  const numCols = sheet.getLastColumn();
  const numRows = lastRow - startRow + 1;

  // 元データの値と数式を取得
  const values = sheet.getRange(startRow, 1, numRows, numCols).getValues();
  const formulas = sheet.getRange(startRow, 1, numRows, numCols).getFormulas();

  // 金曜日の日付計算
  const today = new Date();
  const friday = getNextBusinessFriday(today);
  const fridayText = Utilities.formatDate(friday, ss.getSpreadsheetTimeZone(), 'yyyy/MM/dd');
  // 転記用配列
  const appendValues = [];
  const appendFormulas = [];

  values.forEach((row, i) => {
    if (row[0] !== '' && (row[3] === '' || row[3] === null)) {
      const sourceRow = startRow + i; // コピー元の行番号
      const targetRow = lastRow + 1 + appendValues.length; // コピー先の行番号

      const valRow = [];
      const formRow = [];
      
      for (let j = 0; j < numCols; j++) {
        if (j === 0) { // A列
          valRow[j] = fridayText;
          formRow[j] = null;
        } else if (j === 11) { // L列
          valRow[j] = '';
          formRow[j] = null;
        } else if (formulas[i][j]) { // 数式がある列
          let formula = formulas[i][j];
          // 対象列の場合のみ数式調整
          if (TARGET_COLUMNS.includes(j + 1)) { // 列番号を1ベースに変換
            formula = adjustFormula(formula, sourceRow, targetRow);
          }
          valRow[j] = values[i][j]; // 値はそのまま
          formRow[j] = formula;
        } else { // その他の列
          valRow[j] = values[i][j];
          formRow[j] = null;
        }
      }
      appendValues.push(valRow);
      appendFormulas.push(formRow);
    }
  });

  if (appendValues.length === 0) {
    Logger.log('対象行なし');
    return;
  }

  // シート末尾に貼り付け
  const appendRange = sheet.getRange(lastRow + 1, 1, appendValues.length, numCols);
  appendRange.setValues(appendValues);

  // 数式を個別に設定
  appendFormulas.forEach((formRow, rowIdx) => {
    formRow.forEach((formula, colIdx) => {
      if (formula) {
        sheet.getRange(lastRow + 1 + rowIdx, 1 + colIdx).setFormula(formula);
      }
    });
  });

  Logger.log(`${appendValues.length} 行を追加（A列:${fridayText}、L列クリア、数式調整済み）`);
}

// 数式調整関数
function adjustFormula(formula, sourceRow, targetRow) {
  const rowDiff = targetRow - sourceRow;
  return formula.replace(/(\$?[A-Z]+)(\$?)(\d+)/g, (match, colPart, rowDollar, rowNum) => {
    // 行が絶対参照でない場合のみ調整
    if (rowDollar !== '$') {
      const newRow = parseInt(rowNum) + rowDiff;
      return `${colPart}${rowDollar}${newRow}`;
    }
    return match;
  });
}

function isHoliday(date) {
  const calendarId = 'ja.japanese#holiday@group.v.calendar.google.com';
  const tz = 'Asia/Tokyo';

  const timeMin = Utilities.formatDate(date, tz, "yyyy-MM-dd'T'00:00:00XXX");
  const timeMax = Utilities.formatDate(date, tz, "yyyy-MM-dd'T'23:59:59XXX");

  let events;
  try {
    events = Calendar.Events.list(calendarId, {
      timeMin: timeMin,
      timeMax: timeMax,
      singleEvents: true
    });
  } catch (e) {
    Logger.log('Calendar API error: ' + e);
    return true;
  }

  if (events.items && events.items.length > 0) return true;

  // 土日
  const day = date.getDay();
  if (day === 0 || day === 6) return true;

  // 年末年始（銀行休業）
  const m = date.getMonth() + 1;
  const d = date.getDate();
  if ((m === 12 && d >= 31) || (m === 1 && d <= 3)) return true;

  return false;
}

function getNextBusinessFriday(baseDate) {
  const date = new Date(baseDate);

  // 翌週金曜
  const delta = ((5 - date.getDay()) + 7) % 7 || 7;
  date.setDate(date.getDate() + delta);

  // 祝日なら前営業日へ
  let guard = 0;
    while (isHoliday(date)) {
      date.setDate(date.getDate() - 1);
      guard++;
      if (guard > 7) throw new Error('祝日判定ループ異常');
  }

  return date;
}