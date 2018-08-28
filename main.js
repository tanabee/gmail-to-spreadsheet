// 設定情報
function getConfig() {
  return {
    spreadSheetId: '',// スプレッドシートの ID
    spreadSheetTabName: '',// スプレッドシートのタブ名
    searchText: ''// Gmail の検索ワード
  };
}

// 実行関数
function main() {
  var config   = getConfig();
  var sheet    = SpreadsheetApp.openById(config.spreadSheetId)
                               .getSheetByName(config.spreadSheetTabName);
  var messages = [['Subject', 'From', 'To']];

  GmailApp
    .search(config.searchText, 1, 500)
    .forEach(function (thread) {
      thread.getMessages().forEach(function (message) {
        var subject = message.getSubject();
        var to      = message.getTo();
        var from    = message.getFrom();
        messages.push([subject, from, to]);
      });
    });

  if (messages.length === 0) return;
  sheet.getRange('A1:C' + messages.length ).setValues(messages);
}
