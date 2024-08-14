function createDailyTrigger() {
  // 既存の executeDailyTask トリガーを削除してから新しいトリガーを設定
  deleteExecuteDailyTaskTrigger();

  // 明日の日付を取得
  let now = new Date();
  let tomorrow = new Date(now.getFullYear(), now.getMonth(), now.getDate() + 1);
  tomorrow.setHours(23);
  tomorrow.setMinutes(50);
  tomorrow.setSeconds(0);
  tomorrow.setMilliseconds(0);

  // 翌日23:50に実行するトリガーを作成
  ScriptApp.newTrigger('syncSheetToBigQuery')
    .timeBased()
    .at(tomorrow)
    .create();
}

function deleteExecuteDailyTaskTrigger() {
  // 特定の関数に関連するトリガーを削除
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'syncSheetToBigQuery') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

function executeDailyTask() {
  // ここに毎日実行したいタスクのロジックを記述
  Logger.log("executeDailyTask が実行されました");

  // トリガーを設定（これは初回のみ必要です）
  createDailyTrigger();
}




function syncSheetToBigQuery() {
  var sheetName = "WORK TASK"; // シート名を更新
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var data = sheet.getDataRange().getValues();
  var mergedRanges = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getMergedRanges();

  var headers1 = data[0]; // ヘッダー1行目
  var headers2 = data[1]; // ヘッダー2行目
  
  var projectId = "m2m-core";
  var datasetId = "su_wo";
  var tableId = "tour_control_sheet_log"; // テーブルIDを更新

  // 結合セルを考慮してヘッダー1行目を取得
  var actualHeaders1 = [];
  for (var i = 0; i < headers1.length; i++) {
    var header = headers1[i];
    if (mergedRanges.some(range => range.getColumn() == (i + 1))) {
      var mergedRange = mergedRanges.find(range => range.getColumn() == (i + 1));
      var mergedWidth = mergedRange.getWidth();
      for (var j = 0; j < mergedWidth; j++) {
        actualHeaders1.push(header);
      }
      i += mergedWidth - 1;
    } else {
      actualHeaders1.push(header);
    }
  }
  
var combinedFieldNames = actualHeaders1.map(function(header1, index) {
  var combinedFieldName = header1 + "_" + headers2[index];
  combinedFieldName = combinedFieldName.replace(/[^a-zA-Z0-9_]/g, '_');
  combinedFieldName = combinedFieldName.replace(/_+/g, '_');
  combinedFieldName = combinedFieldName.replace(/^_+|_+$/g, '');
  
  if (/^\d/.test(combinedFieldName)) {
    combinedFieldName = 'F_' + combinedFieldName;
  }
  
  // 無効なフィールド名を空にしない
  if (combinedFieldName.trim() === '') {
    combinedFieldName = 'field_' + index;
  }
  
  return combinedFieldName;
});


  // BigQueryのテーブルスキーマを取得
  var table;
  var currentSchema = [];
  try {
    table = BigQuery.Tables.get(projectId, datasetId, tableId);
    currentSchema = table.schema.fields.map(function(field) {
      return field.name.toLowerCase(); // フィールド名を小文字に変換して格納
    });
  } catch (e) {
    Logger.log('テーブルの取得に失敗しました: ' + e.message);
    return;
  }
  
  // 新しいカラムをスキーマに追加（既存のフィールドをチェック）
  var newFields = [];
  combinedFieldNames.forEach(function(fieldName) {
    var lowerFieldName = fieldName.toLowerCase();
    if (!currentSchema.includes(lowerFieldName)) { // 存在しないフィールドのみ追加
      newFields.push({name: fieldName, type: "STRING"});
    }
  });
  
  if (newFields.length > 0) {
    table.schema.fields = table.schema.fields.concat(newFields);
    try {
      BigQuery.Tables.update(table, projectId, datasetId, tableId, table);
      Logger.log("スキーマが更新されました: " + newFields.map(f => f.name).join(", "));
    } catch (e) {
      Logger.log('スキーマの更新に失敗しました: ' + e.message);
      return;
    }
  } else {
    Logger.log("新しいフィールドはありません。");
  }
  
  // データをCSV形式で準備（改行を削除し、すべての列がnullまたは空の行を除外）
  var csvData = data.filter(function(row) {
    // 全ての列がnullまたは空文字列でないか確認
    return row.some(function(cell) {
      return cell !== null && cell !== '';
    });
  }).map(function(row) {
    return row.map(function(cell) {
      // 改行を削除
      if (typeof cell === 'string') {
        return cell.replace(/\n/g, ' ');
      }
      return cell;
    }).join(",");
  }).join("\n");
  
  // CSVデータをログに出力して確認
  Logger.log('Generated CSV Data: \n' + csvData);
  
  var blob = Utilities.newBlob(csvData, "application/octet-stream");
  
  // BigQueryにロード
  var jobConfig = {
    configuration: {
      load: {
        destinationTable: {
          projectId: projectId,
          datasetId: datasetId,
          tableId: tableId
        },
        skipLeadingRows: 2, // ヘッダー行の数に応じて調整
        writeDisposition: 'WRITE_APPEND', // 既存のデータを保持しつつ新しいデータを追加
        sourceFormat: 'CSV',
        autodetect: false,
        schema: {
          fields: combinedFieldNames.map(function(name) {
            return {name: name, type: "STRING"};
          })
        }
      }
    }
  }
 
  
  ;

  //  executeDailyTask()
  


  try {
    var job = BigQuery.Jobs.insert(jobConfig, projectId, blob);
    Logger.log('BigQuery Job ID: ' + job.id);
  } catch (e) {
    Logger.log('データのロードに失敗しました: ' + e.message);
  }
}
