/**
 * スクリプトプロパティから情報を取得し、セキュアにDiscordへ送信
 */
function sendToDiscord_Nuro() {
  const props = PropertiesService.getScriptProperties();
  
  // 各種情報の取得
  const webhookUrl = props.getProperty('DISCORD_WEBHOOK_URL_NURO');
  const nuroId = props.getProperty('NURO_ID');
  const nuroPw = props.getProperty('NURO_PW');

  // バリデーション（設定漏れチェック）
  if (!webhookUrl || !nuroId || !nuroPw) {
    console.error("エラー: スクリプトプロパティが正しく設定されていません。");
    return;
  }

  // メッセージの組み立て
  const message = [
    "NURO 光の請求書をダウンロードしてGoogleドライブに保存して下さい",
    "https://www.nuro.jp/app/mypage/login",
    "",
    "ユーザーID：" + nuroId,
    "パスワード：" + nuroPw
  ].join("\n");

  const payload = {
    "content": message
  };

  const options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload)
  };

  try {
    UrlFetchApp.fetch(webhookUrl, options);
    console.log("メッセージを送信しました。");
  } catch (e) {
    console.error("送信エラー: " + e.toString());
  }
}