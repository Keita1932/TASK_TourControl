function addinformation() {
  // 読み込みたいプロジェクトの名前
  const project_id = "m2m-core";
  
  // 実行するクエリ
  const query_execute = "SELECT id, company_id, name FROM `m2m_users_prod.user`"; // クエリの最後に引用符を追加
  
  // 実行するクエリの確認
  Logger.log(query_execute);
  
  // 出力先シート名
  const ss = SpreadsheetApp.openById("1ECLNE2D8AptKFWZfu8RSSMYMo5mSVb-WoFqdek7qKls");
  const output_sheet = ss.getSheetByName("addinformation");
  
  // シートの1行目を固定
  output_sheet.setFrozenRows(1);
  
  // ヘッダーを設定
  const headers = [["id", "company_id", "name"]];
  output_sheet.getRange(1, 1, 1, headers[0].length).setValues(headers);
  
  // BigQueryにクエリを実行
  const request = {
    query: query_execute,
    useLegacySql: false,
    location: 'asia-northeast1'  // データセットが存在する場所を指定
  };
  const result = BigQuery.Jobs.query(request, project_id);
  
  // 結果のチェック
  if (result.jobComplete && result.rows) {
    const rows = result.rows.map(row => {
      return row.f.map(cell => cell.v);
    });
    Logger.log(rows);
    
    // 結果をシートに設定
    if (rows.length > 0) {
      output_sheet
        .getRange(
          2, 
          1, 
          rows.length, 
          rows[0].length
        )
        .setValues(rows);
    }
  } else {
    Logger.log("No data returned or the job is not complete.");
  }
}
