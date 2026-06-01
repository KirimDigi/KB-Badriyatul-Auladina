function doGet(e) {
  var action = e.parameter.action;
  var callback = e.parameter.callback;

  if (action === 'get') {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = sheet.getDataRange().getValues();
    var result = [];

    for (var i = 1; i < data.length; i++) {
      result.push({
        name: data[i][0],
        message: data[i][1],
        attendance: data[i][2],
        datetime: data[i][3] + ' | ' + data[i][4]
      });
    }

    result.reverse();

    if (callback) {
      return ContentService.createTextOutput(callback + '(' + JSON.stringify(result) + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService.createTextOutput('OK');
}

function doPost(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var name = e.parameter.name || '';
  var message = e.parameter.message || '';
  var attendance = e.parameter.attendance || '';

  var now = new Date();
  var months = ['Januari','Februari','Maret','April','Mei','Juni','Juli','Agustus','September','Oktober','November','Desember'];
  var tanggal = now.getDate() + ' ' + months[now.getMonth()] + ' ' + now.getFullYear();
  var jam = ('0' + now.getHours()).slice(-2) + ':' + ('0' + now.getMinutes()).slice(-2) + ':' + ('0' + now.getSeconds()).slice(-2);

  sheet.appendRow([name, message, attendance, tanggal, jam]);
  return ContentService.createTextOutput('OK');
}
