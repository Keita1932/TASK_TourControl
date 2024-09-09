function findTroubleById() {
  let token = getApiToken();
  if (!token) {
    Logger.log("トークンを取得できませんでした。");
    throw new Error("トークン取得失敗");
  }

  const ss = SpreadsheetApp.openById("1ExSiRfy4df9yJafRvrMRdKFPw8vmUUzvJlpdSQHtdrQ");
  const troubleFormSheet = ss.getSheetByName("troubleform");
  const cleaningTroubleInfoSheet = ss.getSheetByName("cleaningTroubleInfo");

  // I列のtroubleIdとY列のステータスを全て取得
  let troubleIdRange = troubleFormSheet.getRange(3, 5, troubleFormSheet.getLastRow() - 2, 1).getValues().flat();  // E列 (troubleId)
  let statusRange = troubleFormSheet.getRange(3, 25, troubleFormSheet.getLastRow() - 2, 1).getValues().flat();    // Y列 (ステータス)

  // "済み" ではない行のtroubleIdをフィルタリング
  let troubleIds = troubleIdRange.filter((troubleId, index) => troubleId !== "" && statusRange[index] !== "済み");

  const api_base_url = 'https://api-cleaning.m2msystems.cloud/v3/troubles/';
  const options = {
    method: 'get',
    contentType: 'application/json',
    headers: {
      'Authorization': 'Bearer ' + token
    },
    muteHttpExceptions: true
  };

  let results = [];

  troubleIds.forEach(troubleId => {
    let api_url = api_base_url + troubleId;

    let maxAttempts = 3;  // 最大リトライ回数
    let attempts = 0;
    let success = false;

    while (attempts < maxAttempts && !success) {
      try {
        let response = UrlFetchApp.fetch(api_url, options);
        let responseCode = response.getResponseCode();

        if (responseCode === 200) {
          let jsonResponse = JSON.parse(response.getContentText());
          results.push(jsonResponse);  // 成功した結果を配列に追加
          success = true;
        } else {
          Logger.log('APIリクエスト失敗: ' + responseCode + ' - ' + response.getContentText());
          throw new Error('APIエラー: ' + responseCode);
        }
      } catch (error) {
        attempts++;
        if (attempts < maxAttempts) {
          let delay = Math.pow(2, attempts) * 1000;  // Exponential backoff: 2^attempts seconds
          Logger.log('リトライ中... (' + attempts + '/' + maxAttempts + ') - 次のリトライまで ' + delay / 1000 + ' 秒');
          Utilities.sleep(delay);
        } else {
          Logger.log('最大リトライ回数に達しました。エラー: ' + error.message);
          throw error;
        }
      }
    }
  });

  // "cleaningTroubleInfo" シートに結果を出力
  if (results.length > 0) {
    let lastRow = cleaningTroubleInfoSheet.getLastRow();  // 現在の最終行を取得
    let startRow = lastRow + 1;  // データを出力する開始行

    // データの出力
    results.forEach((result, index) => {
      // 出力するデータを適切に構成
      let outputData = [
        result.id || '', 
        result.status || '',
        result.description || '',
        result.reportedAt || '',
        // 他に必要なフィールドがあればここに追加
      ];
      
      // データをシートの指定行に出力
      cleaningTroubleInfoSheet.getRange(startRow + index, 1, 1, outputData.length).setValues([outputData]);
    });
  }

  Logger.log('全てのAPIリクエスト結果をシートに出力しました。');
}

function clearTrouble() {
  const ss = SpreadsheetApp.openById("1ExSiRfy4df9yJafRvrMRdKFPw8vmUUzvJlpdSQHtdrQ");
  const cleaningTroubleInfoSheet = ss.getSheetByName("cleaningTroubleInfo");
  
  // シートのすべての内容をクリア
  cleaningTroubleInfoSheet.clearContents();
}

