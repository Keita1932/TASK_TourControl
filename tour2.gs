function reload() {
  // LockServiceを使用してロックを取得
  var lock = LockService.getScriptLock();
  
  try {
    // ロックを取得できるまで待機（最大30秒）
    lock.waitLock(30000);
    
    // 必要な関数を実行
    get_tour();
    extractData3();
  } catch (e) {
    Logger.log('エラーが発生しました: ' + e.message);
  } finally {
    // 最後にロックを解放
    lock.releaseLock();
  }
}


function deleteReloadTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'reload') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

function setupReloadTrigger() {
  // 1分後に'reload'関数を再実行するトリガーを設定
  ScriptApp.newTrigger('reload')
           .timeBased()
           .after(60 * 1000)  // 60秒後に実行
           .create();
}


function get_tour() {
  const project_id = "m2m-core";
  const query_execute = `
    WITH filtered_data AS (
      SELECT cleaning_id, listing_id, room_name_common_area_name, status, status_number, work_date,
            work_start_time, photo_tour_id, work_name, worker_name, submission_id, work_cleated_time,
            self_agency, attribute, prefecture
        FROM \`m2m-core.su_wo.wo_cleaning_tour\`
      WHERE work_date BETWEEN DATE_SUB(CURRENT_DATE(), INTERVAL 4 MONTH) AND CURRENT_DATE("Asia/Tokyo")
        AND submission_id <> ''
        AND status_number IN (1, 2, 3, 4, 5, 6)
    ),
    excluded_submission_ids AS (
      SELECT submission_id
        FROM filtered_data
      GROUP BY submission_id
      HAVING COUNTIF(status_number != 6) = 0
    ),
    today_data AS (
      SELECT cleaning_id, listing_id, room_name_common_area_name, status, status_number, work_date,
            work_start_time, photo_tour_id, work_name, worker_name, submission_id, work_cleated_time,
            self_agency, attribute, prefecture
        FROM \`m2m-core.su_wo.wo_cleaning_tour\`
      WHERE work_date = CURRENT_DATE("Asia/Tokyo")
    )
    SELECT *
      FROM filtered_data
    WHERE submission_id NOT IN (SELECT submission_id FROM excluded_submission_ids)

    UNION ALL

    SELECT *
      FROM today_data
    WHERE submission_id NOT IN (SELECT submission_id FROM filtered_data)

    ORDER BY work_date;



  `;
 
  Logger.log(query_execute);
 
  const ss = SpreadsheetApp.openById("1ECLNE2D8AptKFWZfu8RSSMYMo5mSVb-WoFqdek7qKls");
  const output_sheet = ss.getSheetByName("tour2");
  output_sheet.setFrozenRows(1);
 
  const headers = [["cleaning_id", "listing_id", "room_name_common_area_name", "status", "status_number", "work_date",
                    "work_start_time", "photo_tour_id", "work_name", "worker_name", "submission_id", "work_cleated_time",
                    "self_agency", "attribute", "prefecture"]];
  output_sheet.getRange(1, 1, 1, headers[0].length).setValues(headers);
 
  const request = {
    query: query_execute,
    useLegacySql: false,
    location: 'asia-northeast1'
  };
  const result = BigQuery.Jobs.query(request, project_id);    

  Logger.log(result);


  if (result.jobComplete && result.rows) {
    const rows = result.rows.map(row => row.f.map(cell => cell.v));

    const updatedRows = groupAndAssignNumbers(rows);

    Logger.log(updatedRows);
   
    if (updatedRows.length > 0) {
      output_sheet.getRange(2, 1, output_sheet.getLastRow() - 1, 16).clearContent();
      output_sheet.getRange(2, 1, updatedRows.length, 16).setValues(updatedRows);
    }
  } else {
    Logger.log("No data returned or the job is not complete.");
  }
}

function groupAndAssignNumbers(filteredRows) {
  const grouped = {};

  // submission_idごとにグループ化
  filteredRows.forEach(row => {
    const submission_id = row[10]; // submission_idのインデックスが10であると仮定
    if (!grouped[submission_id]) {
      grouped[submission_id] = [];
    }
    grouped[submission_id].push(row);
  });

  // 各グループ内でwork_cleated_timeの早い順に番号を振る
  const result = [];
  for (const submission_id in grouped) {
    const group = grouped[submission_id];
    group.sort((a, b) => new Date(a[11]) - new Date(b[11])); // work_cleated_timeでソート、インデックスが11であると仮定

    // グループ内の各行に番号を振る
    group.forEach((row, index) => {
      row.push(index + 1); // 1から始まる番号を振る
      result.push(row);
    });
  }

  return result;
}


function extractData3() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tourSheet = ss.getSheetByName('tour2');
    const targetSheet = ss.getSheetByName('WORK TASK'); // データを出力したいシート名

    Logger.log('Sheets loaded successfully');

    // A列の空白でない最終行を取得
    const lastRow = tourSheet.getRange('A:A').getValues().filter(row => row[0].toString().trim() !== '').length;

    // A2:M列のデータを最終行まで取得
    const data = tourSheet.getRange(2, 1, lastRow - 1, 16).getValues(); // 2行目から始まる範囲を取得
    const filteredData = [];

    Logger.log('Data fetched successfully');

    data.forEach((row, index) => {
        if ((row[9] === '(WO)TASKチーム' || row[9] === '(WO)小笠原渓太') && row[10] && row[15] === 1) {
            filteredData.push([row[5], row[2], row[4], row[10], row[11]]); // 必要なデータをfilteredDataに追加
        }
        Logger.log('Row %s processed: %s', index + 2, JSON.stringify(row));
    });

    if (filteredData.length > 0) {
        Logger.log('Filtered data: %s', JSON.stringify(filteredData));

        // filteredDataを日付（F列, filteredData[0]）が早い順にソート
        filteredData.sort((a, b) => new Date(a[0]) - new Date(b[0]));

        // WORK TASKシートのD列の既存データを取得
        const targetData = targetSheet.getRange('D3:D' + targetSheet.getLastRow()).getValues().flat();

        // 新しいデータを格納する配列
        const newData = filteredData.filter(row => !targetData.includes(row[3]));

        if (newData.length > 0) {
            Logger.log('New data: %s', JSON.stringify(newData));

            // 各列の新しいデータを抽出
            const newDataA = newData.map(row => [row[0]]);
            const newDataB = newData.map(row => [row[1]]);
            const newDataD = newData.map(row => [row[3]]);

            // WORK TASKシートのA列の文字列のみを数える
            const lastRowTarget = targetSheet.getRange('A:A').getValues().filter(row => row[0] !== '').length + 1;

            // データを設定
            targetSheet.getRange(lastRowTarget + 1, 1, newDataA.length, 1).setValues(newDataA); // A列
            targetSheet.getRange(lastRowTarget + 1, 2, newDataB.length, 1).setValues(newDataB); // B列
            targetSheet.getRange(lastRowTarget + 1, 4, newDataD.length, 1).setValues(newDataD); // D列

            Logger.log('New data set successfully');
        } else {
            Logger.log('No new data to add');
        }
    } else {
        Logger.log('条件を満たすデータがありません');
    }
}


function resetData3() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheet = ss.getSheetByName('WORK TASK'); // データを出力したいシート名
  
  // A3:Fのコンテンツをクリア
  targetSheet.getRange('A3:B').clearContent();
  targetSheet.getRange('D3:D').clearContent();
  targetSheet.getRange('K3:N').clearContent();
  // targetSheet.getRange('T3:Y').clearContent();


  executeDailyTask2()


}


function createDailyTrigger2() {
  // 既存の executeDailyTask トリガーを削除してから新しいトリガーを設定
  deleteExecuteDailyTaskTrigger2();

  // 明日の日付を取得
  let now = new Date();
  let tomorrow = new Date(now.getFullYear(), now.getMonth(), now.getDate() + 1);
  tomorrow.setHours(0);
  tomorrow.setMinutes(5);
  tomorrow.setSeconds(0);
  tomorrow.setMilliseconds(0);

  // 翌日00:05に実行するトリガーを作成
  ScriptApp.newTrigger('resetData3')
    .timeBased()
    .at(tomorrow)
    .create();
}

function deleteExecuteDailyTaskTrigger2() {
  // 特定の関数に関連するトリガーを削除
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'resetData3') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

function executeDailyTask2() {
  // ここに毎日実行したいタスクのロジックを記述
  Logger.log("executeDailyTask が実行されました");

  // トリガーを設定（これは初回のみ必要です）
  createDailyTrigger2();
}

