// function tour() {
//   // 読み込みたいプロジェクトの名前
//   const project_id = "m2m-core";
  
//   // 実行するクエリ
//   const query_execute = `
//     SELECT cleaning_id, listing_id, room_name_common_area_name, status, status_number, work_date,
//            work_start_time, photo_tour_id, work_name, worker_name, submission_id, work_created_time, 
//            self_agency, attribute, prefecture
//     FROM \`m2m-core.su_wo.wo_cleaning_tour\`
//     WHERE work_date BETWEEN DATE_SUB(CURRENT_DATE(), INTERVAL 4 MONTH) AND CURRENT_DATE()
//       AND submission_id <> ''
//       AND status_number NOT IN (0, 1, 2)
//     ORDER BY work_date
//   `;
  
//   // 実行するクエリの確認
//   Logger.log(query_execute);
  
//   // 出力先シート名
//   const ss = SpreadsheetApp.openById("1ECLNE2D8AptKFWZfu8RSSMYMo5mSVb-WoFqdek7qKls");
//   const output_sheet = ss.getSheetByName("tour");
  
//   // シートの1行目を固定
//   output_sheet.setFrozenRows(1);
  
//   // ヘッダーを設定
//   const headers = [["cleaning_id", "listing_id", "room_name_common_area_name", "status", "status_number", "work_date",
//                     "work_start_time", "photo_tour_id", "work_name", "worker_name", "submission_id", "work_created_time", 
//                     "self_agency", "attribute", "prefecture", "group_number"]];
//   output_sheet.getRange(1, 1, 1, headers[0].length).setValues(headers);
  
//   // BigQueryにクエリを実行
//   const request = {
//     query: query_execute,
//     useLegacySql: false,
//     location: 'asia-northeast1'  // データセットが存在する場所を指定
//   };
//   const result = BigQuery.Jobs.query(request, project_id);
  
//   // 結果のチェック
//   if (result.jobComplete && result.rows) {
//     const rows = result.rows.map(row => row.f.map(cell => cell.v));
//     Logger.log(rows);
    
//     // submission_idごとにグループ化し、work_created_timeの早い順に番号を振る
//     const groupedData = groupAndAssignNumbers(rows);

//     // 結果をシートに設定
//     if (groupedData.length > 0) {
//       output_sheet
//         .getRange(
//           2, 
//           1, 
//           groupedData.length, 
//           groupedData[0].length
//         )
//         .setValues(groupedData);
//     }
//   } else {
//     Logger.log("No data returned or the job is not complete.");
//   }
// }

function groupAndAssignNumbers(rows) {
  const grouped = {};

  // submission_idごとにグループ化
  rows.forEach(row => {
    const submission_id = row[10];
    if (!grouped[submission_id]) {
      grouped[submission_id] = [];
    }
    grouped[submission_id].push(row);
  });

  // 各グループ内でwork_created_timeの早い順に番号を振る
  const result = [];
  for (const submission_id in grouped) {
    const group = grouped[submission_id];
    group.sort((a, b) => new Date(a[11]) - new Date(b[11])); // work_created_timeでソート

    group.forEach((row, index) => {
      row.push(index + 1); // 1から始まる番号を振る
      result.push(row);
    });
  }

  return result;
}

