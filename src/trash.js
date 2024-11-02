

// function onOpen() {
//   SpreadsheetApp.getUi()
//     .createMenu('カスタムメニュー')
//     .addItem('稼働登録', 'openRegistrationForm')
//     .addToUi();
// }

// function openRegistrationForm() {
//   var html = HtmlService.createHtmlOutputFromFile('RegistrationForm')
//     .setWidth(400)
//     .setHeight(300);
//   SpreadsheetApp.getUi().showModalDialog(html, '稼働登録');
// }

// function getEmployeeList() {
//   var ss = SpreadsheetApp.getActiveSpreadsheet();
//   var sheet = ss.getSheetByName('配送業者マスタ');
//   var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
//   var employees = data.map(function(row) {
//     return { id: row[0], name: row[1] };
//   });
//   return employees;
// }

// function registerOperation(data) {
//   var employeeId = data.employeeId;
//   var operationStatus = data.operationStatus;
//   var dates = data.dates;

//   var ss = SpreadsheetApp.getActiveSpreadsheet();
//   var sheet = ss.getSheetByName('稼働マスタ');

//   var lastRow = sheet.getLastRow()
//   var existingEntries = sheet.getRange(2, 1, lastRow, 2).getValues();
//   var existingEntriesSet = new Set(existingEntries.map(function(row) {
//     return row[0] + '_' + Utilities.formatDate(new Date(row[1]), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd');
//   }));

//   var newEntries = [];
//   var skippedDates = [];

//   dates.forEach(function(dateStr) {
//     var dateObj = new Date(dateStr);
//     var dateFormatted = Utilities.formatDate(dateObj, ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd');
//     var key = employeeId + '_' + dateFormatted;
//     if (existingEntriesSet.has(key)) {
//       skippedDates.push(dateFormatted);
//     } else {
//       newEntries.push([employeeId, dateObj, operationStatus]);
//       existingEntriesSet.add(key);
//     }
//   });

//   if (newEntries.length > 0) {
//     sheet.getRange(lastRow + 1, 1, newEntries.length, 3).setValues(newEntries);
//   }

//   return {
//     success: true,
//     skippedDates: skippedDates
//   };
// }