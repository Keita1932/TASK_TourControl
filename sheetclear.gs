function clearRangeAndSetTrigger() {
  // スプレッドシートを開く
  const spreadsheet = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1ExSiRfy4df9yJafRvrMRdKFPw8vmUUzvJlpdSQHtdrQ/edit');
  
  // // 「ツアー作成用」タブを取得
   const tourSheet = spreadsheet.getSheetByName('ツアー作成用');
  
  // // 「通知情報」タブを取得
   const notifySheet = spreadsheet.getSheetByName('通知情報');

  
  // 「ツアー作成用」タブの R2:R の範囲をクリア
  tourSheet.getRange('R2:R').clearContent();
  
  // 「通知情報」タブの U2:U の範囲をクリア
  notifySheet.getRange('U2:U').clearContent();
  
  // clearRangeAndSetTrigger 関数のトリガーを削除
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'clearRangeAndSetTrigger') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  
  // 翌日の0:05にこの関数をトリガーする
  const tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);
  tomorrow.setHours(0, 5, 0, 0);
  
  ScriptApp.newTrigger('clearRangeAndSetTrigger')
    .timeBased()
    .at(tomorrow)
    .create();
}
