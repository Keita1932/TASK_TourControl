function extractTodayData() {
  // スプレッドシートIDとシート名
  const sourceSpreadsheetId = '1RzSScYg7WuJ8NHvc5k71H5djXN4y_0DV597-IfnF2sw';
  const sourceSheetName = 'row';
  
  const targetSpreadsheetId = '1ExSiRfy4df9yJafRvrMRdKFPw8vmUUzvJlpdSQHtdrQ';
  const targetSheetName = 'troubleform';
  
  // 今日の日付を取得（日本時間）
  const today = new Date();
  const todayString = Utilities.formatDate(today, 'Asia/Tokyo', 'yyyy-MM-dd');

  // 元のスプレッドシートとシートを開く
  const sourceSpreadsheet = SpreadsheetApp.openById(sourceSpreadsheetId);
  const sourceSheet = sourceSpreadsheet.getSheetByName(sourceSheetName);
  
  // 目的のスプレッドシートとシートを開く
  const targetSpreadsheet = SpreadsheetApp.openById(targetSpreadsheetId);
  const targetSheet = targetSpreadsheet.getSheetByName(targetSheetName);

  // 元のシートのデータ範囲を取得
  const dataRange = sourceSheet.getDataRange();
  const data = dataRange.getValues();

  // ヘッダーを保持しつつ、本日の日付に一致する行のみフィルタリング
  const headers = data.slice(0, 2); // 最初の2行をヘッダーとして保持
  const filteredData = data.slice(2).filter(row => {
    const createdAt = row[1]; // B列の日時文字列を取得
    const createdAtDate = new Date(createdAt); // 日時文字列をDateオブジェクトに変換
    const dateString = Utilities.formatDate(createdAtDate, 'Asia/Tokyo', 'yyyy-MM-dd');
    return dateString === todayString; // 日付部分だけを比較
  });

  // filteredData内の日時列を文字列として扱う
  const filteredDataString = filteredData.map(row => {
    row[16] = row[16].toString(); // M列を文字列に変換
    row[17] = row[17].toString(); // N列を文字列に変換
    return row;
  });

  // フィルタリング結果にヘッダーを追加
  const resultData = headers.concat(filteredDataString);

  // 目的のシートにデータを書き込む
  targetSheet.clear(); // 既存のデータをクリア
  if (resultData.length > 2) { // データがある場合
    targetSheet.getRange(1, 1, resultData.length, resultData[0].length).setValues(resultData);
  } else {
    targetSheet.getRange(1, 1, 2, headers[0].length).setValues(headers); // データがない場合はヘッダーのみ書き込む
  }

  // ログに結果を出力
  Logger.log('Data has been successfully written to the target sheet.');
}
