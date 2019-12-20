function onOpen() {
  SpreadsheetApp
    .getActiveSpreadsheet()
    .addMenu('メール', [
      {name: '一斉送信', functionName: 'showUI'},
      {name: 'テスト送信', functionName: 'showTestUI'},
    ]);
}

function showUI() {
  const html = HtmlService.createHtmlOutputFromFile('form');
  SpreadsheetApp.getUi().showModalDialog(html, "メール一斉送信");
}

function showTestUI() {
  const html = HtmlService.createHtmlOutputFromFile('testForm');
  SpreadsheetApp.getUi().showModalDialog(html, "メールテスト送信");
}

function processForm(value) {
  sendMails(value.subject, value.template);
  Browser.msgBox("メールを一斉送信しました。");
}

function processTestForm(value) {
  const data = SpreadsheetApp.getActiveSheet().getDataRange().getValues();
  const header = data[0];
  const row = data[1];
  const address = value.toAddress;
  const subject = value.subject;
  const template = value.template;
  sendMail(address, header, row, subject, template);
  Browser.msgBox("メールをテスト送信しました。");
}

function sendMails(subject, template) {
  const data = SpreadsheetApp.getActiveSheet().getDataRange().getValues();
  const header = data[0];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var mail = {};
    var address = row[0];
    sendMail(address, header, row, subject, template);
  }
}

function sendMail(address, header, row, subject, template) {
  const t = HtmlService.createTemplateFromFile(template);
  for (var j = 1; j < row.length && j < header.length; j++) {
    t[header[j]] = row[j];
  }
  const body = t.evaluate().getContent();
  MailApp.sendEmail(address, subject, body);
  Logger.log(address + ":" + subject);
}