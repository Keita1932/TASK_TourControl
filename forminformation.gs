function forminformation() {
  // スプレッドシートIDとシート名
  const sourceSpreadsheetId = '1RzSScYg7WuJ8NHvc5k71H5djXN4y_0DV597-IfnF2sw';
  const sourceSheetName = 'row';
  
  const targetSpreadsheetId = '1ECLNE2D8AptKFWZfu8RSSMYMo5mSVb-WoFqdek7qKls';
  const targetSheetName = 'forminformation';
  
  // 今日の日付を取得（日本時間）
  const today = new Date();
  const sevenMonthsAgo = new Date(today);
  sevenMonthsAgo.setMonth(today.getMonth() - 7);
  
  const todayString = Utilities.formatDate(today, 'Asia/Tokyo', 'yyyy-MM-dd');
  const sevenMonthsAgoString = Utilities.formatDate(sevenMonthsAgo, 'Asia/Tokyo', 'yyyy-MM-dd');

  // 元のスプレッドシートとシートを開く
  const sourceSpreadsheet = openSpreadsheetWithRetry(sourceSpreadsheetId);
  const sourceSheet = sourceSpreadsheet.getSheetByName(sourceSheetName);
  
  // 目的のスプレッドシートとシートを開く
  const targetSpreadsheet = openSpreadsheetWithRetry(targetSpreadsheetId);
  const targetSheet = targetSpreadsheet.getSheetByName(targetSheetName);

  // 元のシートのデータ範囲を取得
  const dataRange = sourceSheet.getDataRange();
  const data = dataRange.getValues();

  // ヘッダーを保持しつつ、7ヶ月前から今日までの範囲に一致する行のみフィルタリング
  const headers = [[data[0][0], data[0][8], data[0][16]]]; // 最初の行のA列, I列, Q列をヘッダーとして保持
  const filteredData = data.slice(1).filter(row => {
    const createdAt = row[1]; // B列の日時文字列を取得
    const createdAtDate = new Date(createdAt); // 日時文字列をDateオブジェクトに変換
    const dateString = Utilities.formatDate(createdAtDate, 'Asia/Tokyo', 'yyyy-MM-dd');
    return dateString >= sevenMonthsAgoString && dateString <= todayString && row[20] === 'WO'; // 7ヶ月前から今日までの日付範囲を比較し、かつU列が"WO"であることを確認
  }).map(row => [row[0], row[8], row[16]]); // A列, I列, Q列のみを抽出

  // フィルタリング結果にヘッダーを追加
  const resultData = headers.concat(filteredData);

  // 目的のシートにデータを書き込む
  if (resultData.length > 1) { // データがある場合
    targetSheet.getRange(1, 1, resultData.length, resultData[0].length).setValues(resultData);
  } else {
    targetSheet.getRange(1, 1, 1, headers[0].length).setValues(headers); // データがない場合はヘッダーのみ書き込む
  }

  // ログに結果を出力
  Logger.log('Data has been successfully written to the target sheet.');
}

function openSpreadsheetWithRetry(spreadsheetId, retries = 3, delay = 1000) {
  let spreadsheet;
  while (retries > 0) {
    try {
      spreadsheet = SpreadsheetApp.openById(spreadsheetId);
      return spreadsheet;
    } catch (e) {
      if (--retries === 0) {
        throw e;
      }
      Utilities.sleep(delay);
    }
  }
}


