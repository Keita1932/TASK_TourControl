function makingTour() {
  let lock = LockService.getScriptLock();
  try {
    // 10秒以内にロックを取得しようと試みる
    if (!lock.tryLock(10000)) {
      Logger.log('他のプロセスが実行中のため、makingTour 関数を終了します。');
      return; // 他の関数が実行中の場合は終了
    }

    findTroubleById()

    const sheetId = "1ExSiRfy4df9yJafRvrMRdKFPw8vmUUzvJlpdSQHtdrQ";
    const ss = SpreadsheetApp.openById(sheetId);
    const operationSheet = ss.getSheetByName("ツアー作成用");

    let token = getApiToken();
    if (!token) {
      Logger.log("トークンを取得できませんでした。");
      throw new Error("トークン取得失敗");
    }

    Logger.log(token);

    let lastRow = getLastNonEmptyRowByColumn(operationSheet, 3); // C列を基準に最終行を取得

    if (lastRow < 2) {
      Logger.log("データがありません。");
      return;
    }

    const operationData = operationSheet.getRange(2, 1, lastRow - 1, 29).getValues();

    // 空白行をフィルタリング
    const requestValues = operationData.filter(row => row.some(cell => cell !== ""));

    // APIリクエスト
    for (let i = 0; i < requestValues.length; i++) {
      // R列を参照し、"ok"の場合はスキップ
      if (requestValues[i][17] === "ok") {
        Logger.log('行 ' + (i + 2) + ' をスキップ: R列が "ok"');
        continue;
      }

      const api_url = 'https://api-cleaning.m2msystems.cloud/v3/cleanings/create_with_placement';

      let placement;
      if (requestValues[i][5]) {
        placement = "commonArea";
      } else if (requestValues[i][3]) {
        placement = "listing";
      } else {
        placement = "";
      }

      const submissionId = requestValues[i][0];
      const handoverid = requestValues[i][1];

      let note = "【解決してほしい人】\n" +
        (requestValues[i][11] ? requestValues[i][11].replace(/\n/g, " ") : "") + "\n" +
        "【トラブルの内容】\n" +
        (requestValues[i][10] ? requestValues[i][10].replace(/\n/g, " ") : "") + "\n" +
        "【やってほしいこと】\n" +
        (requestValues[i][9] ? requestValues[i][9].replace(/\n/g, " ") : "") + "\n" +
        "【フォームID】\n" +
        submissionId;

      if (handoverid) {
        note += "\n【このツアーは引き継ぎタスクです。 前回のフォームID】\n" + handoverid;
      }
      note = note.trim();

      const cleaningDate = convertDate(requestValues[i][2]);
      Logger.log('cleaningDate: ' + cleaningDate);

      let cleaners = [];
      if (requestValues[i][16] !== '') {
        cleaners.push(requestValues[i][16]);
      }

      let error = false;
      if (cleaners.length === 0) {
        Logger.log('エラー: cleanersが空です。行 ' + (i + 2));
        operationSheet.getRange(i + 2, 18).setValue('error');
        error = true;
      }

      if (requestValues[i][14] === '') {
        Logger.log('エラー: photoTourIdが空です。行 ' + (i + 2));
        operationSheet.getRange(i + 2, 18).setValue('error');
        error = true;
      }

      if (error) {
        continue;
      }

      const payload = {
        "placement": placement,
        "commonAreaId": requestValues[i][5],
        "listingId": requestValues[i][3],
        "cleaningDate": cleaningDate.replace(/\//g, '-'),
        "note": note,
        "cleaners": cleaners,
        "submissionId": submissionId,
        "photoTourId": requestValues[i][14] !== '' ? requestValues[i][14] : null
      };

      Logger.log('payload: ' + JSON.stringify(payload));

      const options = {
        'method': 'post',
        'contentType': 'application/json',
        'headers': {
          'Authorization': 'Bearer ' + token
        },
        'payload': JSON.stringify(payload),
        'muteHttpExceptions': true
      };

      Logger.log('API request options: ' + JSON.stringify(options));

      Logger.log('Sending API request for row ' + (i + 2));

      let response = UrlFetchApp.fetch(api_url, options);
      response = response.getContentText();
      Logger.log('response: ' + response);

      let result;
      if (response.includes('error')) {
        result = 'error';
      } else {
        result = 'ok';
      }

      Logger.log('result: ' + result);

      operationSheet.getRange(i + 2, 18).setValue(result);
    }
  } catch (e) {
    Logger.log('エラーが発生しました: ' + e.message);
  } finally {
    lock.releaseLock(); // ロックを解除
    // setupTrigger();
  }
}
