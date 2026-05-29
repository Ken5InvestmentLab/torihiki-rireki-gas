/**
 * @fileoverview
 * スプレッドシートの需給分析データを管理するためのGoogle Apps Script。
 * JPX（日本取引所グループ）およびJSDA（日本証券業協会）からデータを取得・処理します。
 * Webhook URLはスクリプトプロパティに安全に保存されます。
 * * * 主な機能：
 * 1. onOpen: スプレッドシートを開いたときにカスタムメニューを追加します。
 * 2. setDiscordWebhookUrl: Webhook URLをスクリプトプロパティに設定します。
 * 3. findAndRecordMissingJpxPdfs: 「需給分析」シートでデータが不足している日付のJPX PDFを検索し、URLを記録します。
 * 4. sendToDiscord: 新しいPDFが見つかった場合にDiscordへ通知します。
 * 5. importJpxDataFromSheet: 「JPX元データ」シートに貼り付けられたPDF内容を解析し、「需給分析」シートに転記します。
 * 6. updateJSDA: JSDAから信用取引残高データを取得し、シートを更新します。
 */

// ===============================================================
// スプレッドシートを開いた時の処理
// ===============================================================
/**
 * スプレッドシートを開いた時にカスタムメニューを追加する関数です。
 * Webhook URLを設定するメニューもここに追加します。
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('データ更新')
    .addItem('1. JPXの未取得PDFを検索・記録', 'findAndRecordMissingJpxPdfs')
    .addItem('2. 「JPX元データ」から取込', 'importJpxDataFromSheet')
    .addSeparator() // メニューの区切り線
    .addItem('JPX毎日16:32の自動実行を設定', 'createDailyJpxTrigger')
    .addItem('JSDA毎日15:32の自動実行を設定', 'createDailyJsdaTrigger') // ★追加箇所
    .addSeparator() // メニューの区切り線
    .addItem('Webhook URLを設定/変更', 'setDiscordWebhookUrl')
    .addToUi();
}

// ===============================================================
// トリガー設定用関数
// ===============================================================
function createDailyJpxTrigger() {
  const functionNameToTrigger = 'findAndRecordMissingJpxPdfs';
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === functionNameToTrigger) {
      ScriptApp.deleteTrigger(trigger);
    }
  }
  ScriptApp.newTrigger(functionNameToTrigger)
    .timeBased()
    .everyDays(1)
    .atHour(16)
    .nearMinute(32)
    .create();
  SpreadsheetApp.getUi().alert(
    '自動実行トリガーの設定完了',
    `「${functionNameToTrigger}」が毎日16:32頃に自動で実行されるように設定しました。`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function createDailyJsdaTrigger() {
  const functionNameToTrigger = 'updateJSDA';
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === functionNameToTrigger) {
      ScriptApp.deleteTrigger(trigger);
    }
  }
  ScriptApp.newTrigger(functionNameToTrigger)
    .timeBased()
    .everyDays(1)
    .atHour(15)
    .nearMinute(32)
    .create();
  SpreadsheetApp.getUi().alert(
    '自動実行トリガーの設定完了',
    `「${functionNameToTrigger}」が毎日15:32頃に自動で実行されるように設定しました。`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// ===============================================================
// Discord Webhook設定
// ===============================================================
function setDiscordWebhookUrl() {
  const ui = SpreadsheetApp.getUi();
  const currentUrl = PropertiesService.getScriptProperties().getProperty('DISCORD_WEBHOOK_URL') || '';
  const result = ui.prompt(
    'Discord Webhook URL設定',
    'Discordから取得したWebhook URLを以下に貼り付けてください。\n（現在設定されているURL: ' + currentUrl + '）',
    ui.ButtonSet.OK_CANCEL
  );
  const button = result.getSelectedButton();
  const webhookUrl = result.getResponseText();
  if (button == ui.Button.OK) {
    if (webhookUrl && webhookUrl.trim().startsWith('https://discord.com/api/webhooks/')) {
      PropertiesService.getScriptProperties().setProperty('DISCORD_WEBHOOK_URL', webhookUrl.trim());
      ui.alert('成功', 'Webhook URLを保存しました。', ui.ButtonSet.OK);
    } else if (webhookUrl.trim() === '') {
      PropertiesService.getScriptProperties().deleteProperty('DISCORD_WEBHOOK_URL');
      ui.alert('クリアしました', 'Webhook URLの設定を削除しました。', ui.ButtonSet.OK);
    } else {
      ui.alert('エラー', '無効なURLのようです。URLを確認してもう一度お試しください。', ui.ButtonSet.OK);
    }
  }
}

// ===============================================================
// シート編集時に自動実行
// ===============================================================
function handleSheetEdit(e) {
  const sheetName = e.source.getSheetName();
  if (sheetName === 'JPX元データ') {
    SpreadsheetApp.getActiveSpreadsheet().toast('「JPX元データ」の編集を検知。データ転記を実行します...', '自動実行中', 5);
    importJpxDataFromSheet();
  }
}

// ===============================================================
// JPX用：半自動（ハイブリッド方式）
// ===============================================================

/**
 * 【JPX用スクリプト 1】
 * 「需給分析」シートをチェックし、データが未入力の日付に対応するJPXのPDFをすべて探し、
 * 「JPX_PDF_URL置場」シートに追記します。新しいPDFが見つかった場合はDiscordに通知します。
 */
function findAndRecordMissingJpxPdfs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const juyouSheet = ss.getSheetByName('需給分析');
  const urlSheet = ss.getSheetByName('JPX_PDF_URL置場');

  if (!juyouSheet || !urlSheet) {
    SpreadsheetApp.getUi().alert('エラー：「需給分析」または「JPX_PDF_URL置場」シートが見つかりません。');
    return;
  }

  const juyouData = juyouSheet.getDataRange().getValues();
  const targetDates = new Set();
  for (let i = 1; i < juyouData.length; i++) {
    const row = juyouData[i];
    const date = row[0];
    const jpxSell = row[3];
    const jpxBuy = row[4];
    if (date instanceof Date && !jpxSell && !jpxBuy) {
      targetDates.add(Utilities.formatDate(date, 'JST', 'yyyyMMdd'));
    }
  }

  if (targetDates.size === 0) {
    SpreadsheetApp.getActiveSpreadsheet().toast('JPXデータを入力すべき日付が見つかりませんでした。', '完了', 5);
    return;
  }

  const existingUrlsData = urlSheet.getDataRange().getValues();
  const existingDates = new Set();
  for (const row of existingUrlsData) {
    const url = row[1];
    if (typeof url === 'string' && url) {
      const match = url.match(/syumatsu(\d{8})00\.pdf/);
      if (match) existingDates.add(match[1]);
    }
  }
  
  const datesToCheck = [...targetDates].filter(date => !existingDates.has(date));

  if (datesToCheck.length === 0) {
    SpreadsheetApp.getActiveSpreadsheet().toast('チェック対象の新しい日付はありませんでした。', '完了', 5);
    return;
  }

  const newlyFoundPdfs = [];
  for (const ymd of datesToCheck) {
    const url = `https://www.jpx.co.jp/markets/statistics-equities/margin/tvdivq0000001rnl-att/syumatsu${ymd}00.pdf`;
    try {
      const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      if (response.getResponseCode() === 200) {
        const dateObj = new Date(ymd.slice(0, 4), ymd.slice(4, 6) - 1, ymd.slice(6, 8));
        const formattedDate = Utilities.formatDate(dateObj, 'JST', 'yyyy/MM/dd');
        newlyFoundPdfs.push([formattedDate, url]);
        Logger.log(`PDF発見: ${url}`);
      }
    } catch (e) {
      Logger.log(`URL ${url} のチェック中にエラーが発生しました: ${e.message}`);
    }
  }

  if (newlyFoundPdfs.length > 0) {
    urlSheet.getRange(urlSheet.getLastRow() + 1, 1, newlyFoundPdfs.length, 2).setValues(newlyFoundPdfs);
    const message = `新しいJPX信用残PDFが見つかりました🎉\n\n` +
                    newlyFoundPdfs.map(pdf => `- ${pdf[0]}: <${pdf[1]}>`).join('\n');
    const notified = sendToDiscord(message);
    SpreadsheetApp.getActiveSpreadsheet().toast(notified ? `${newlyFoundPdfs.length}件の新しいPDFを記録し、Discordに通知しました。` : `${newlyFoundPdfs.length}件のPDFを記録しましたが、Discord通知に失敗しました（実行ログを確認してください）。`, notified ? '成功' : '注意', 10);
  } else {
    SpreadsheetApp.getActiveSpreadsheet().toast('対象期間内に見つかった新しいPDFはありませんでした。', '完了', 5);
  }
}


/**
 * 【JPX用スクリプト 2】
 * 「JPX元データ」シートに貼り付けられたテキストを解析し、「需給分析」シートに転記します。
 */
function importJpxDataFromSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName('JPX元データ');
  const targetSheet = ss.getSheetByName('需給分析');
  const urlSheet = ss.getSheetByName('JPX_PDF_URL置場');

  if (!sourceSheet || !targetSheet || !urlSheet) {
    Logger.log('エラー：必要なシートが見つかりません。');
    SpreadsheetApp.getUi().alert('エラー：「JPX元データ」「需給分析」「JPX_PDF_URL置場」のいずれかのシートが見つかりません。');
    return;
  }

  const sourceData = sourceSheet.getDataRange().getValues();
  const urlData = urlSheet.getDataRange().getValues();
  const targetData = targetSheet.getDataRange().getValues();

  const filledDates = new Set();
  for (let i = 1; i < targetData.length; i++) {
    if (targetData[i][0] instanceof Date && targetData[i][3] && targetData[i][4]) {
      filledDates.add(Utilities.formatDate(targetData[i][0], 'JST', 'yyyyMMdd'));
    }
  }

  const pendingUrls = [];
  for (const row of urlData) {
    const url = row[1];
    if (typeof url !== 'string' || !url) continue;
    const dateMatch = url.match(/syumatsu(\d{8})00\.pdf/);
    if (dateMatch && !filledDates.has(dateMatch[1])) {
      pendingUrls.push({ date: dateMatch[1], url: url });
    }
  }

  if (pendingUrls.length === 0) {
    Logger.log('処理対象のPDFがURL置場に見つかりません。');
    SpreadsheetApp.getActiveSpreadsheet().toast('処理対象の新しいデータはありません。');
    return;
  }

  pendingUrls.sort((a, b) => b.date.localeCompare(a.date));
  const pdfDateStr = pendingUrls[0].date;
  Logger.log('処理対象の日付: ' + pdfDateStr);
  SpreadsheetApp.getActiveSpreadsheet().toast(`${pdfDateStr} のデータを処理します...`, '情報', 5);

  const marginMap = {};
  for (const row of sourceData) {
    const line = row.join(' ');
    if (typeof line !== 'string' || line.trim() === '') continue;

    const codeMatch = line.match(/\b(\d{4}[A-Z]?)0\b/);
    if (codeMatch) {
      Logger.log('銘柄コード行を発見: ' + line);
      
      const codeIndex = line.indexOf(codeMatch[0]);
      let dataString = line.substring(codeIndex + codeMatch[0].length);

      dataString = dataString.replace(/\s*JP\w+\s+/, ' ');

      const values = dataString.match(/([▲]?[\d,]+)/g);
      
      if (values && values.length >= 3) {
        Logger.log('抽出した数値リスト: ' + values.join(', '));
        
        const sellStr = values[0];
        const buyStr = values[2];

        const sell = parseInt(sellStr.replace(/,/g, ''), 10);
        const buy = parseInt(buyStr.replace(/,/g, ''), 10);

        if (!isNaN(sell) && !isNaN(buy) && sell >= 0 && buy >= 0) {
          const code4digit = codeMatch[1];
          marginMap[code4digit] = { sell, buy };
          Logger.log(`マップに追加: コード=${code4digit}, 売り=${sell}, 買い=${buy}`);
        } else {
          Logger.log(`パース後の数値が不正です: sell=${sell}, buy=${buy}`);
        }
      } else {
        Logger.log('値の抽出に失敗、または数値の数が不足: ' + dataString);
      }
    }
  }
  
  Logger.log('解析完了。マップの件数: ' + Object.keys(marginMap).length);

  let updatedCount = 0;
  for (let i = 1; i < targetData.length; i++) {
    const targetRow = targetData[i];
    const rowDate = targetRow[0];
    if (!(rowDate instanceof Date)) continue;

    const rowDateStr = Utilities.formatDate(rowDate, 'JST', 'yyyyMMdd');
    if (rowDateStr === pdfDateStr && !targetRow[3] && !targetRow[4]) {
      const code = String(targetRow[1]);
      if (marginMap[code]) {
        targetSheet.getRange(i + 1, 4).setValue(marginMap[code].sell);
        targetSheet.getRange(i + 1, 5).setValue(marginMap[code].buy);
        updatedCount++;
      }
    }
  }
  
  if (updatedCount > 0) {
    SpreadsheetApp.getActiveSpreadsheet().toast(`${pdfDateStr} のデータについて、${updatedCount}行を更新しました。`, '成功', 7);
    sourceSheet.clearContents();
    SpreadsheetApp.getActiveSpreadsheet().toast('「JPX元データ」シートをクリアしました。', '情報', 5);
  } else {
    Logger.log('更新対象のデータが見つかりませんでした。貼り付けたデータの日付や内容を確認してください。');
    SpreadsheetApp.getActiveSpreadsheet().toast('更新対象のデータが見つかりませんでした。', '情報', 5);
  }
}

// ===============================================================
// Discord通知用関数
// ===============================================================
function sendToDiscord(message) {
  const webhookUrl = PropertiesService.getScriptProperties().getProperty('DISCORD_WEBHOOK_URL');
  // セキュリティ: Webhook URL（トークンを含む）はログに出力しない
  if (!webhookUrl) {
    const msg = 'Webhook URLが設定されていません。\n\n' +
                'メニューの「データ更新」>「Webhook URLを設定/変更」からURLを登録してください。';
    Logger.log(msg);
    try { SpreadsheetApp.getUi().alert(msg); } catch (uiErr) { /* トリガー実行時はUIなし */ }
    return false;
  }

  const options = {
    'method': 'post',
    'contentType': 'application/json',
    // Discord/Cloudflareのbot判定による429回避のため User-Agent を明示（Discord APIは本来UA必須）
    'headers': { 'User-Agent': 'Ken5InvestmentLab-GAS/1.0 (+https://script.google.com)' },
    'payload': JSON.stringify({ content: message }),
    'muteHttpExceptions': true
  };

  const MAX_ATTEMPTS = 5;
  for (let attempt = 1; attempt <= MAX_ATTEMPTS; attempt++) {
    let response;
    try {
      response = UrlFetchApp.fetch(webhookUrl, options);
    } catch (e) {
      // ネットワーク例外（タイムアウト等）→ 指数バックオフで再試行
      Logger.log(`Discord通知で例外（試行${attempt}/${MAX_ATTEMPTS}）: ${e.message}`);
      if (attempt < MAX_ATTEMPTS) { Utilities.sleep(attempt * 2000); continue; }
      return false;
    }

    const code = response.getResponseCode();

    // 204(通常成功) / 200(?wait=true時) は成功
    if (code === 204 || code === 200) {
      Logger.log(`Discord通知成功: HTTP ${code}（試行${attempt}回目）`);
      return true;
    }

    // 429: レート制限 → Retry-After に従って待機して再試行
    if (code === 429) {
      const bodyText = response.getContentText();
      let waitSec = attempt * 2; // フォールバック（指数的）
      try {
        const body = JSON.parse(bodyText);
        // Discordのwebhook 429は retry_after が「秒」。念のため極端に大きい値はミリ秒とみなす
        if (body && typeof body.retry_after === 'number') {
          waitSec = body.retry_after > 600 ? body.retry_after / 1000 : body.retry_after;
        }
        if (body && body.global === true) Logger.log('Discordグローバルレート制限を検知。');
      } catch (parseErr) {
        // JSONでない（Cloudflareのブロックページ等）→ ヘッダの Retry-After を見る
        const headers = response.getAllHeaders();
        const ra = headers['Retry-After'] || headers['retry-after'];
        if (ra) { const h = parseFloat(ra); if (!isNaN(h)) waitSec = h; }
      }
      waitSec = Math.min(Math.max(waitSec, 1), 60); // 1〜60秒に丸める
      Logger.log(`Discord 429（レート制限）。${waitSec}秒待機して再試行 (試行${attempt}/${MAX_ATTEMPTS})。応答: ${bodyText.slice(0, 200)}`);
      if (attempt < MAX_ATTEMPTS) { Utilities.sleep(Math.ceil(waitSec * 1000) + 300); continue; }
      Logger.log('Discord通知: 429が続いたため断念しました。');
      return false;
    }

    // 5xx はサーバ側の一過性エラー → 再試行
    if (code >= 500) {
      Logger.log(`Discordサーバエラー HTTP ${code}（試行${attempt}/${MAX_ATTEMPTS}）。応答: ${response.getContentText().slice(0, 200)}`);
      if (attempt < MAX_ATTEMPTS) { Utilities.sleep(attempt * 2000); continue; }
      return false;
    }

    // その他の4xx（400/401/404等）は再試行しても無駄なので即終了
    Logger.log(`Discord通知エラー HTTP ${code}。応答: ${response.getContentText().slice(0, 300)}`);
    return false;
  }
  return false;
}

// ===============================================================
// JSDA用：完全自動
// ===============================================================
function updateJSDA() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('需給分析');
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  const dateGroupMap = {};

  for (let i = 1; i < data.length; i++) {
    if ((data[i][6] || data[i][7]) || !data[i][0] || !(data[i][0] instanceof Date)) continue;

    // let に変更して日付を操作できるようにする
    let cellDate = new Date(data[i][0]);
    const dayOfWeek = cellDate.getDay();
    
    // いったん週末(金曜)に寄せる
    cellDate.setDate(cellDate.getDate() + (5 - dayOfWeek)); 

    // ★追加：金曜が祝日や休業日の場合、直近の過去の営業日まで遡る
    while (!isBusinessDay(cellDate)) {
      cellDate.setDate(cellDate.getDate() - 1);
    }

    const ymd = Utilities.formatDate(cellDate, 'JST', 'yyyyMMdd');

    if (!dateGroupMap[ymd]) dateGroupMap[ymd] = { rows: [] };
    dateGroupMap[ymd].rows.push(i + 1);
  }

  const ymdList = Object.keys(dateGroupMap).sort();
  if (ymdList.length === 0) {
    Logger.log('更新対象なし');
    return;
  }

  // 1回の実行で最古の1週だけ処理
  const ymd = ymdList[0];

  const urls = [
    `https://www.jsda.or.jp/shiryoshitsu/toukei/kabu-taiw/files/${ymd}z.xlsx`,
    `https://www.jsda.or.jp/shiryoshitsu/toukei/kabu-taiw/files/${ymd}z.xls`
  ];

  let blob = null;
  let successUrl = null;

  for (const url of urls) {
    try {
      blob = fetchWithRetries(url, 2, 1).setName(`${ymd}${url.endsWith('.xlsx') ? '.xlsx' : '.xls'}`);
      successUrl = url;
      Logger.log(`ダウンロード成功: ${url}`);
      break;
    } catch (e) {
      Logger.log(`ファイル取得失敗: ${url} (${e.message})`);
    }
  }

  if (!blob) {
    Logger.log(`ファイル ${ymd} の取得に最終的に失敗`);
    return;
  }

  const resource = {
    name: blob.getName(),
    mimeType: 'application/vnd.google-apps.spreadsheet'
  };

  const file = Drive.Files.create(resource, blob, { supportsAllDrives: true });

  try {
    const excelRows = SpreadsheetApp.openById(file.id).getSheets()[0].getDataRange().getValues();
    const loanDataMap = {};

    for (let j = 1; j < excelRows.length; j++) {
      const r = excelRows[j];
      const excelCode5digit = r[1] ? r[1].toString() : '';
      if (!excelCode5digit) continue;

      const type = r[2];
      const amount = parseInt(r[3], 10) || 0;

      if (!loanDataMap[excelCode5digit]) {
        loanDataMap[excelCode5digit] = { secured: 0, unsecured: 0 };
      }

      if (type === '有担保') {
        loanDataMap[excelCode5digit].secured = amount;
      } else if (type === '無担保') {
        loanDataMap[excelCode5digit].unsecured = amount;
      }
    }

    const rowsToUpdate = dateGroupMap[ymd].rows;
    let updatedCount = 0;

    for (const rowIndex of rowsToUpdate) {
      const sheetCode4digit = sheet.getRange(rowIndex, 2).getValue();
      if (!sheetCode4digit) continue;

      const lookupCode5digit = sheetCode4digit.toString() + '0';
      const loanData = loanDataMap[lookupCode5digit];

      if (loanData) {
        sheet.getRange(rowIndex, 7).setValue(loanData.secured);
        sheet.getRange(rowIndex, 8).setValue(loanData.unsecured);
        updatedCount++;
      }
    }

    Logger.log(`更新完了: ${ymd}, ${updatedCount}件, URL=${successUrl}`);
  } finally {
    Drive.Files.remove(file.id);
  }
}

function fetchWithRetries(url, maxRetries = 2, delaySeconds = 1) {
  for (let i = 0; i < maxRetries; i++) {
    const response = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      followRedirects: true
    });

    const code = response.getResponseCode();

    if (code === 200) {
      return response.getBlob();
    }

    // 404 は再試行してもほぼ無意味
    if (code === 404) {
      throw new Error(`404 Not Found`);
    }

    if (i === maxRetries - 1) {
      throw new Error(`HTTP ${code}`);
    }

    Logger.log(`ダウンロード失敗（試行 ${i + 1}/${maxRetries}回, HTTP ${code}）。${delaySeconds}秒後に再試行します...`);
    Utilities.sleep(delaySeconds * 1000);
  }
}

// ===============================================================
// 営業日判定用関数
// ===============================================================
function isBusinessDay(date) {
  const day = date.getDay();
  // 1. 土日判定（0:日曜, 6:土曜）
  if (day === 0 || day === 6) return false;

  const m = date.getMonth() + 1; // getMonth()は0始まりのため+1
  const d = date.getDate();
  
  // 2. 大晦日・三が日（銀行休業日）判定
  if (m === 12 && d === 31) return false;
  if (m === 1 && d >= 1 && d <= 3) return false;

  // 3. 祝日判定（Google提供の日本の祝日カレンダーを利用）
  const calendarId = 'ja.japanese#holiday@group.v.calendar.google.com';
  const calendar = CalendarApp.getCalendarById(calendarId);
  
  if (calendar) {
    const events = calendar.getEventsForDay(date);
    if (events.length > 0) {
      // カレンダーに登録されているが、金融機関が休みになる「祝日」ではない行事リスト
      const nonHolidays = [
        '節分', '七夕', 'バレンタイン', 'ひな祭り', 'ひなまつり', 
        'ホワイトデー', 'ハロウィン', 'クリスマス', 'イースター', 
        '母の日', '父の日'
      ];
      
      for (const event of events) {
        const title = event.getTitle();
        let isTrueHoliday = true;
        
        for (const nh of nonHolidays) {
          if (title.includes(nh)) {
            isTrueHoliday = false; // 行事リストに一致したら「祝日ではない」とする
            break;
          }
        }
        
        // 振替休日、国民の休日、正規の祝日の場合は休業日（false）を返す
        if (isTrueHoliday) return false;
      }
    }
  }
  
  // 上記のいずれの休業条件にも引っかからなければ営業日
  return true;
}
