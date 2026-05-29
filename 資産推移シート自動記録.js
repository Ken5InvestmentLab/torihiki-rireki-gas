/**
 * 資産サマリの D2 から株価（または資産額）を取得し、
 * 資産推移シートの末尾に日付と値を追記、降順ソートするスクリプト。
 * * - recordAssetValue()         : 手動実行用（週末もスキップせず実行）
 * - recordAssetValueTrigger()  : 時間駆動トリガー用（土日はスキップ）
 * - recordAssetValueInternal() : 共通ロジック
 * - fetchPriceWithRetry()      : カスタム関数の計算完了を待つポーリング取得
 */

/**
 * 手動起動（週末も実行）
 */
function recordAssetValue() {
  recordAssetValueInternal(false);
}

/**
 * トリガー起動（土日はスキップ）
 * ※Apps Script 管理画面で「時間主導型トリガー」に
 * この関数を登録してください。
 */
function recordAssetValueTrigger() {
  recordAssetValueInternal(true);
}

/**
 * 実処理本体
 * @param {boolean} skipWeekendCheck - true なら土日をスキップ
 */
function recordAssetValueInternal(skipWeekendCheck) {
  const today = new Date();
  const tz    = Session.getScriptTimeZone();
  const dow   = today.getDay(); // 0=日曜,6=土曜

  // 土日スキップ判定
  if (skipWeekendCheck && (dow === 0 || dow === 6)) {
    console.log('週末のため記録をスキップ:', Utilities.formatDate(today, tz, 'yyyy/MM/dd'));
    return;
  }

  // シート取得
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const src   = ss.getSheetByName('資産サマリ');
  const tgt   = ss.getSheetByName('資産推移');
  if (!src || !tgt) {
    console.error('シートが見つかりません');
    if (!skipWeekendCheck) SpreadsheetApp.getUi().alert('エラー: シートが見つかりません');
    return;
  }

  // D2 の値をポーリングで取得（数値になるまで待機） ★E2からD2に変更
  let assetValue;
  try {
    assetValue = fetchPriceWithRetry(src, 'D2', 10, 1000);
  } catch (e) {
    console.error(e.message);
    if (!skipWeekendCheck) SpreadsheetApp.getUi().alert('価格取得失敗: ' + e.message);
    return;
  }

  // MM/dd フォーマットの日付
  const formattedDate = Utilities.formatDate(today, tz, 'MM/dd');

  // 資産推移シートの最終データ行を探す (1行目=ヘッダー、2行目以降がデータ)
  const colA = tgt.getRange('A2:A').getValues();
  let lastDataRow = 1;  // ヘッダー行
  for (let i = 0; i < colA.length; i++) {
    if (colA[i][0] !== '') {
      lastDataRow = 2 + i;
    } else {
      break;
    }
  }
  const nextRow = lastDataRow + 1;

  // 追記：A列=日付、B列=資産額
  tgt.getRange(nextRow, 1).setValue(formattedDate);
  tgt.getRange(nextRow, 2).setValue(assetValue);

  // 2行目以降 A〜D 列を A 列基準で降順ソート
  const newLastRow = tgt.getLastRow();
  const numRows = newLastRow - 1;   // データ件数
  const numCols = 4;                // A〜D
  if (numRows > 1) {
    const sortRange = tgt.getRange(2, 1, numRows, numCols);
    sortRange.sort({ column: 1, ascending: false });
  }

  console.log(`記録完了: ${formattedDate} = ${assetValue} (行 ${nextRow})`);
}

/**
 * カスタム関数(GETPRICE 等)の計算完了を待って
 * 指定セルの表示値を数値化して返すポーリングロジック
 *
 * @param {Sheet}   sheet      - データ取得元シート
 * @param {string}  a1Range    - 取得セル (例: 'D2')
 * @param {number}  maxRetries - 最大リトライ回数
 * @param {number}  delayMs    - リトライ間隔 (ミリ秒)
 * @return {number}            - 取得した数値
 * @throws {Error}             - 取得失敗時
 */
function fetchPriceWithRetry(sheet, a1Range, maxRetries, delayMs) {
  let raw, val;
  for (let i = 0; i < maxRetries; i++) {
    SpreadsheetApp.flush();
    raw = sheet.getRange(a1Range).getDisplayValue();
    // ★通貨記号やカンマを除去→数字のみ残す
    const cleaned = raw.replace(/[^\d\.\-]/g, '');
    val = parseFloat(cleaned);
    if (!isNaN(val)) {
      return val;
    }
    Utilities.sleep(delayMs);
  }
  throw new Error(`セル ${a1Range} の取得に失敗しました (raw="${raw}")`);
}