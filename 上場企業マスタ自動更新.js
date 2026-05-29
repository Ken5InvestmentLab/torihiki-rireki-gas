/**
 * 指定のExcelファイルからB列とC列を取得し、
 * スプレッドシートの指定シートに追記するGASサンプル
 * 同じ証券コード（B列）がすでに存在する行はスキップ
 */

/**
 * 汎用インポート関数
 * @param {string} excelUrl  ExcelファイルのURL
 * @param {string} targetSheetName 対象シート名
 */
function importB2C(excelUrl, targetSheetName) {
  // デバッグ用ログ: 関数がどの引数で呼び出されたか確認
  console.log(`importB2C invoked with excelUrl='${excelUrl}', targetSheetName='${targetSheetName}'`);

  // 引数の存在チェック
  if (!excelUrl || !targetSheetName) {
    // エラー発生時には、どの引数が不足しているかを示すメッセージを投げる
    throw new Error(`必須パラメータが不足しています: excelUrl='${excelUrl}', targetSheetName='${targetSheetName}'`);
  }

  let converted; // スコープをtryブロック外に移動
  let xlsFile;   // スコープをtryブロック外に移動

  try {
    // 1. ExcelファイルをBlobとして取得
    console.log(`Fetching Excel file from: ${excelUrl}`);
    const response = UrlFetchApp.fetch(excelUrl, { muteHttpExceptions: true }); // エラー時にもレスポンスを取得
    const responseCode = response.getResponseCode();
    if (responseCode !== 200) {
        throw new Error(`Excelファイルの取得に失敗しました。ステータスコード: ${responseCode}, URL: ${excelUrl}`);
    }
    const blob = response.getBlob();
    console.log('Excel file fetched successfully.');

    // 2. Driveに一時XLSファイルとして保存
    console.log('Creating temporary XLS file in Google Drive...');
    xlsFile = DriveApp.createFile(blob.setName(`temp_list_${Date.now()}.xls`)); // ユニークなファイル名
    console.log(`Temporary XLS file created with ID: ${xlsFile.getId()}`);

    // 3. Googleスプレッドシートに変換（Drive API Advancedサービス）
    // 注意: Drive API Advanced Serviceを有効にする必要があります
    // スクリプトエディタの「サービス」+ボタンから「Drive API」を追加
    console.log('Converting XLS to Google Sheet using Drive API...');
    const resource = {
      title: `temp_list_converted_${Date.now()}`, // ユニークなファイル名
      mimeType: MimeType.GOOGLE_SHEETS
    };
    converted = Drive.Files.copy(resource, xlsFile.getId());
    console.log(`Converted Google Sheet created with ID: ${converted.id}`);
    const ssTemp = SpreadsheetApp.openById(converted.id);

    // 4. 変換後のシートからB列・C列を取得（ヘッダ下から最終行まで）
    console.log('Reading data from converted Google Sheet...');
    const sheetTemp = ssTemp.getSheets()[0]; // 最初のシートを対象とする
    const last = sheetTemp.getLastRow();
    if (last < 2) {
      console.log('インポート元にデータがありません (ヘッダー行のみ)');
      return; // データがない場合は処理終了
    }
    // B列(2列目)から2列分(B列とC列)を、2行目から (last - 1) 行分取得
    const data = sheetTemp.getRange(2, 2, last - 1, 2).getValues();
    console.log(`Read ${data.length} rows from source sheet.`);

    // 5. 追記先のスプレッドシートと既存マスタデータを取得
    console.log(`Accessing target sheet: ${targetSheetName}`);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(targetSheetName);
    if (!sheet) {
      throw new Error(`シートが見つかりません: ${targetSheetName}`);
    }
    const lastRow = sheet.getLastRow();
    // 既存データの取得（A列を基準とする）
    const existing = lastRow > 1
      ? sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat().map(String) // 既存コードを文字列配列に
      : [];
    console.log(`Found ${existing.length} existing codes in target sheet.`);

    // 6. 重複チェック＆追加するデータを作成
    console.log('Checking for duplicates and preparing data to add...');
    const toAdd = [];
    const addedCodes = new Set(); // 追加済みコードを管理し、入力データ内の重複も防ぐ
    data.forEach(function(r) {
      const code = (r[0] || '').toString().trim(); // B列の値（証券コード）
      const name = (r[1] || '').toString().trim(); // C列の値（企業名）

      // コードが空、またはヘッダー行のような文字列の場合はスキップ
      if (!code || code === 'コード' || code === '証券コード') return;

      // 既存データに含まれておらず、かつ今回の追加リストにもまだ含まれていない場合
      if (existing.indexOf(code) === -1 && !addedCodes.has(code)) {
        toAdd.push([code, name]); // A列にコード、B列に名前を追加
        addedCodes.add(code); // 追加リストに追加したことを記録
      }
    });

    // 7. マスタシートに追記
    if (toAdd.length > 0) {
      console.log(`Adding ${toAdd.length} new rows to sheet '${targetSheetName}'...`);
      // 最終行+1行目から、A列(1列目)に、追加データ(toAdd)の行数分、2列(A, B列)書き込む
      sheet.getRange(lastRow + 1, 1, toAdd.length, 2).setValues(toAdd);
      console.log(`${toAdd.length} 件を追加しました`);
    } else {
      console.log('新規データはありませんでした');
    }

  } catch (e) {
    // エラーハンドリング
    console.error(`エラーが発生しました: ${e.message}`);
    console.error(`スタックトレース: ${e.stack}`);
    // 必要に応じてエラーを再スローするか、通知などを追加
    // throw e;
  } finally {
    // 8. 一時ファイルをゴミ箱へ移動 (tryブロック内で定義されたか確認)
    if (xlsFile) {
      try {
        console.log(`Moving temporary XLS file (${xlsFile.getId()}) to trash...`);
        DriveApp.getFileById(xlsFile.getId()).setTrashed(true);
        console.log('Temporary XLS file moved to trash.');
      } catch (e) {
        console.error(`一時XLSファイル (${xlsFile.getId()}) の削除中にエラー: ${e.message}`);
      }
    }
    // 変換後の一時スプレッドシートを削除 (tryブロック内で定義されたか確認)
    if (converted && converted.id) {
      try {
        console.log(`Deleting temporary Google Sheet (${converted.id})...`);
        Drive.Files.remove(converted.id); // Drive APIを使って完全に削除
        console.log('Temporary Google Sheet deleted.');
      } catch (e) {
        console.error(`一時Google Sheet (${converted.id}) の削除中にエラー: ${e.message}`);
        // 代替としてゴミ箱へ移動を試みる
        try {
            DriveApp.getFileById(converted.id).setTrashed(true);
            console.log(`Temporary Google Sheet (${converted.id}) moved to trash as fallback.`);
        } catch (e2) {
            console.error(`一時Google Sheet (${converted.id}) のゴミ箱への移動も失敗: ${e2.message}`);
        }
      }
    }
  }
}

/**
 * 東証上場銘柄一覧Excelを取り込み、「上場企業マスタ」シートに反映する関数
 * この関数をGoogle Apps Scriptエディタから実行してください。
 */
function importListedCompanies() {
  // JPXから提供されている上場銘柄一覧データ(Excel形式)のURL
  const excelUrl = 'https://www.jpx.co.jp/markets/statistics-equities/misc/tvdivq0000001vg2-att/data_j.xls';
  // データを追記するシート名
  const targetSheet = '上場企業マスタ';

  console.log(`importListedCompanies starts. Target sheet: '${targetSheet}', Source URL: '${excelUrl}'`);

  // 汎用インポート関数を呼び出し
  importB2C(excelUrl, targetSheet);

  console.log('importListedCompanies finished.');
}
