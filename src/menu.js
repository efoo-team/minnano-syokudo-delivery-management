function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('カスタムメニュー')
    .addItem('スケジュールを出力', 'generateDeliverySchedule')
    .addToUi();
}
