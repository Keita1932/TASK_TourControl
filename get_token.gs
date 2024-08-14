function getApiToken() {
  const url = 'https://api.m2msystems.cloud/login';
  const mail = PropertiesService.getScriptProperties().getProperty("mail_address");
  const pass = PropertiesService.getScriptProperties().getProperty("pass");

  const payload = {
    email: mail,
    password: pass
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  
  if (response.getResponseCode() == 200) {
    const json = JSON.parse(response.getContentText());
    const token = json.accessToken;
    // Logger.log("取得したトークン: " + token);
    return token;
  } else {
    // Logger.log("エラーが発生しました。ステータスコード: " + response.getResponseCode());
    return null;
  }
}