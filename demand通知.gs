function sendReportToSlack() {
  // 排他制御のロックを取得
  const lock = LockService.getScriptLock();
  try {
    // 他のプロセスがロックを取得している場合は10秒待ち、それでも取得できなければエラーを返す
    if (lock.tryLock(10000)) {
      const spreadsheetId = "1ExSiRfy4df9yJafRvrMRdKFPw8vmUUzvJlpdSQHtdrQ";
      const sheetName = "通知情報";
      const startRow = 1;
      const lastRow = getLastDataRow4(spreadsheetId, sheetName);

      const sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(sheetName);
      const range = sheet.getRange("A" + startRow + ":X" + lastRow);
      const values = range.getValues();

      for (let i = 0; i < values.length; i++) {
        try {
          const row = values[i];
          const reportStatus = row[0];
          const rColumnValue = row[20];
          const madeTourStatus = row[21];
          const classification = row[2];
          const sColumnValue = row[19];

          // 物件名の参照（空白でない方を使用）
          const propertyName = row[8] !== "" ? row[8] : row[11];

          // 滞在期間のフォーマット変更
          const stayPeriod = `${formatDate(row[16])}~${formatDate(row[17])}`;

          // 入力日時のフォーマット変更
          const entryTime = convertISOToCustomFormat(row[1]);

          if (reportStatus !== "" && rColumnValue !== "済み" && madeTourStatus == "ok") {
            const user1 = "<!subteam^SM53BKPR8>";
            const user2 = "<!subteam^S05NVPXMSNP>";

            let color;

            // 特定の分類に対する色の設定
            if (classification === "自火報トラブル" || classification === "物理鍵トラブル" || classification === "TTlockトラブル") {
              color = "#ED1A3D"; // 赤
            } else if (sColumnValue === "CX") {
              color = "#f2c744"; // 黄色
            } else {
              color = "#0000ff"; // 青
            }

            const message = {
              "text": (sColumnValue === "CX") ? user1 : user2,
              "attachments": [
                {
                  "color": color,
                  "blocks": [
                    {
                      "type": "header",
                      "text": {
                        "type": "plain_text",
                        "text": "❗️トラブル報告❗️",
                        "emoji": true
                      }
                    },
                    {
                      "type": "section",
                      "fields": [
                        {
                          "type": "mrkdwn",
                          "text": `*物件名:*\n${propertyName}`
                        },
                        {
                          "type": "mrkdwn",
                          "text": `*契約属性:*\n${row[9]}`
                        },
                        {
                          "type": "mrkdwn",
                          "text": `*分類:*\n${classification}`
                        },
                        {
                          "type": "mrkdwn",
                          "text": `*フォームID:*\n${row[0]}`
                        }
                      ]
                    },
                    {
                      "type": "section",
                      "text": {
                        "type": "mrkdwn",
                        "text": `*誰から（予約コード）:*\n${row[3]}`
                      }
                    },
                    {
                      "type": "section",
                      "text": {
                        "type": "mrkdwn",
                        "text": `*何が起きた:*\n${row[4]}`
                      }
                    },
                    {
                      "type": "section",
                      "text": {
                        "type": "mrkdwn",
                        "text": `*何をして欲しい:*\n${row[5]}`
                      }
                    },
                    {
                      "type": "section",
                      "fields": [
                        {
                          "type": "mrkdwn",
                          "text": `*予約経路:*\n${row[15]}`
                        },
                        {
                          "type": "mrkdwn",
                          "text": `*滞在期間:*\n${stayPeriod}`
                        },
                        {
                          "type": "mrkdwn",
                          "text": `*入力日時:*\n${entryTime}`
                        },
                        {
                          "type": "mrkdwn",
                          "text": `*入力者:*\n${row[13]}`
                        },
                        {
                          "type": "mrkdwn",
                          "text": `*トラブルURL:*\n${row[14]}`
                        },
                        {
                          "type": "mrkdwn",
                          "text": `*引き継ぎフォームID:*\n${row[6]}`
                        }
                      ]
                    }
                  ]
                }
              ]
            };

            sendToSlack4(message);

            const statusCell = sheet.getRange("U" + (startRow + i));
            statusCell.setValue("済み");

          } else if (madeTourStatus == "error") {
            const user1 = "<!subteam^SM53BKPR8>";
            const user2 = "<!subteam^S05NVPXMSNP>";
            const message = {
              "text": (sColumnValue === "CX") ? user1 : user2,
              "attachments": [
                {
                  "color": "#000000",
                  "blocks": [
                    {
                      "type": "header",
                      "text": {
                        "type": "plain_text",
                        "text": "☠️ツアー作成失敗☠️",
                        "emoji": true
                      }
                    },
                    {
                      "type": "section",
                      "fields": [
                        {
                          "type": "mrkdwn",
                          "text": `*物件名:*\n${propertyName}`
                        },
                        {
                          "type": "mrkdwn",
                          "text": `*契約属性:*\n${row[9]}`
                        },
                        {
                          "type": "mrkdwn",
                          "text": `*分類:*\n${row[2]}`
                        },
                        {
                          "type": "mrkdwn",
                          "text": `*フォームID:*\n${row[0]}`
                        }
                      ]
                    },
                    {
                      "type": "section",
                      "text": {
                        "type": "mrkdwn",
                        "text": `*誰から（予約コード）:*\n${row[3]}`
                      }
                    },
                    {
                      "type": "section",
                      "text": {
                        "type": "mrkdwn",
                        "text": `*何が起きた:*\n${row[4]}`
                      }
                    },
                    {
                      "type": "section",
                      "text": {
                        "type": "mrkdwn",
                        "text": `*何をして欲しい:*\n${row[5]}`
                      }
                    },
                    {
                      "type": "section",
                      "fields": [
                        {
                          "type": "mrkdwn",
                          "text": `*予約経路:*\n${row[15]}`
                        },
                        {
                          "type": "mrkdwn",
                          "text": `*滞在期間:*\n${stayPeriod}`
                        },
                        {
                          "type": "mrkdwn",
                          "text": `*入力日時:*\n${entryTime}`
                        },
                        {
                          "type": "mrkdwn",
                          "text": `*入力者:*\n${row[13]}`
                        },
                        {
                          "type": "mrkdwn",
                          "text": `*トラブルURL:*\n${row[14]}`
                        },
                        {
                          "type": "mrkdwn",
                          "text": `*引き継ぎフォームID:*\n${row[6]}`
                        }
                      ]
                    }
                  ]
                }
              ]
            };
            sendToSlack4(message);
            const statusCell = sheet.getRange("U" + (startRow + i));
            statusCell.setValue("済み");
            changeError();
          }
        } catch (error) {
          Logger.log(`行 ${i + 1} でエラーが発生しました。エラーメッセージ: ${error}`);
          continue; // エラーが発生した場合はスキップして次の行に進む
        }
      }
    }
  } finally {
    // ロックを解除
    lock.releaseLock();
  }
}

function sendToSlack4(message) {
  let webhookUrl = "https://hooks.slack.com/services/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"; // あなたのSlackのWebhook URLに置き換えてください

  let payload = JSON.stringify(message);

  let options = {
    "method": "post",
    "contentType": "application/json",
    "payload": payload
  };

  UrlFetchApp.fetch(webhookUrl, options);
}

function changeError() {
  const sheetId = "1ExSiRfy4df9yJafRvrMRdKFPw8vmUUzvJlpdSQHtdrQ";
  const ss = SpreadsheetApp.openById(sheetId);
  const operationSheet = ss.getSheetByName("ツアー作成用");

  // R列のデータを取得
  const lastRow = getLastNonEmptyRowByColumn(operationSheet, 18); // R列を基準に最終行を取得
  const statusColumn = operationSheet.getRange(2, 18, lastRow - 1).getValues(); // データ範囲取得

  for (let i = 0; i < statusColumn.length; i++) {
    if (statusColumn[i][0] === 'error') {
      Logger.log('行 ' + (i + 2) + ' のステータスを "error" から "ok" に変更します');
      operationSheet.getRange(i + 2, 18).setValue('ok');
    }
  }
}

function getLastDataRow4(spreadsheetId, sheetName) {
  let sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(sheetName);
  let lastRow = sheet.getLastRow();
  let range = sheet.getRange("A1:A" + lastRow);
  let values = range.getValues();

  for (let i = values.length - 1; i >= 0; i--) {
    if (values[i][0] !== "") {
      return i + 1;
    }
  }

  return lastRow;
}

function formatDate(date) {
  if (date instanceof Date) {
    return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  } else {
    return date; // Dateオブジェクトでない場合は元の値を返します
  }
}

function formatDateTime(dateTime) {
  if (dateTime instanceof Date) {
    return Utilities.formatDate(dateTime, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  } else {
    return dateTime; // Dateオブジェクトでない場合は元の値を返します
  }
}

function convertISOToCustomFormat(isoString) {
  // ISO 8601形式の日付文字列をDateオブジェクトに変換
  let date = new Date(isoString);

  // 年、月、日、時、分、秒を取得
  let year = date.getFullYear();
  let month = ('0' + (date.getMonth() + 1)).slice(-2); // 月は0から始まるので +1
  let day = ('0' + date.getDate()).slice(-2);
  let hours = ('0' + date.getHours()).slice(-2);
  let minutes = ('0' + date.getMinutes()).slice(-2);
  let seconds = ('0' + date.getSeconds()).slice(-2);

  // yyyy/mm/dd hh:mm:ss形式に変換
  return `${year}/${month}/${day} ${hours}:${minutes}:${seconds}`;
}

