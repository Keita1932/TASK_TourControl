function sendReportToSlack() {
  // 排他制御のロックを取得
  var lock = LockService.getScriptLock();
  try {
    // 他のプロセスがロックを取得している場合は10秒待ち、それでも取得できなければエラーを返す
    if (lock.tryLock(10000)) {
      let spreadsheetId = "1ExSiRfy4df9yJafRvrMRdKFPw8vmUUzvJlpdSQHtdrQ";
      let sheetName = "通知情報";
      let startRow = 1;
      let lastRow = getLastDataRow4(spreadsheetId, sheetName);

      let sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(sheetName);
      let range = sheet.getRange("A" + startRow + ":X" + lastRow);
      let values = range.getValues();

      for (let i = 0; i < values.length; i++) {
        let row = values[i];
        let reportStatus = row[0];
        let rColumnValue = row[20];
        let madeTourStatus = row[21];
        let classification = row[2];
        let sColumnValue = row[19];

         // 物件名の参照（空白でない方を使用）
        let propertyName = row[8] !== "" ? row[8] : row[11];

        // 滞在期間のフォーマット変更
        let stayPeriod = `${formatDate(row[16])}~${formatDate(row[17])}`;

        // 入力日時のフォーマット変更
        let entryTime = formatDateTime(row[1]);


        if (reportStatus !== "" && rColumnValue !== "済み" && madeTourStatus == "ok") {
          let user1 = "<!subteam^SM53BKPR8>";
          let user2 = "<!subteam^S05NVPXMSNP>";

          let color;

          // 特定の分類に対する色の設定
          if (classification === "自火報トラブル" || classification === "物理鍵トラブル" || classification === "TTlockトラブル") {
            color = "#ED1A3D"; // 赤
          } else if (sColumnValue === "CX") {
            color = "#f2c744"; // 黄色
          } else {
            color = "#0000ff"; // 青
          }

          let message = {
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

          let statusCell = sheet.getRange("U" + (startRow + i));
          statusCell.setValue("済み");
        } else if (madeTourStatus == "error") {
          let user1 = "<!subteam^SM53BKPR8>";
          let user2 = "<!subteam^S05NVPXMSNP>";
          let message = {
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
          let statusCell = sheet.getRange("U" + (startRow + i));
          statusCell.setValue("済み");
          changeError()
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