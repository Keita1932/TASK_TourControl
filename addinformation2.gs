function addinformation2() {
  try {
    // 読み込みたいプロジェクトの名前
    const project_id = "m2m-core";
    
    // 実行するクエリ
    const query_execute = `
      SELECT room_id, operation_type_ja
      FROM \`m2m-core.dx_001_room.room_operation_type_ja\`
    `;
    
    // 実行するクエリの確認
    Logger.log(query_execute);
    
    // 出力先シート名
    const ss = SpreadsheetApp.openById("1ExSiRfy4df9yJafRvrMRdKFPw8vmUUzvJlpdSQHtdrQ");
    const output_sheet = ss.getSheetByName("addinformation2");
    
    // シートの1行目を固定
    output_sheet.setFrozenRows(1);
    
    // ヘッダーを設定
    const headers = [["room_id", "operation_type_ja"]];
    output_sheet.getRange(1, 1, 1, headers[0].length).setValues(headers);
    
    // BigQueryにクエリを実行
    const request = {
      query: query_execute,
      useLegacySql: false
    };
    
    const queryResults = BigQuery.Jobs.query(request, project_id);
    
    // 結果のチェック
    if (queryResults.jobComplete) {
      if (queryResults.rows && queryResults.rows.length > 0) {
        const rows = queryResults.rows.map(row => row.f.map(cell => cell.v));
        Logger.log(rows);
        
        // 結果をシートに設定
        output_sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
      } else {
        Logger.log("No data returned.");
      }
    } else {
      Logger.log("The job is not complete.");
    }
  } catch (error) {
    Logger.log("Error: " + error.message);
  }
}
