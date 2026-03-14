function sendSlackMessage(text) {

  const url = CONFIG.SLACK.WEBHOOK_URL;

  const payload = {
    text: text
  };

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload)
  };

  const response = UrlFetchApp.fetch(url, options);

  Logger.log(response.getContentText());

}

function testSlack() {
  sendSlackMessage("出席管理システム テスト通知");
}