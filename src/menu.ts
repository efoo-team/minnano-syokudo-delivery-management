const onOpen = () => {
    const ui = SpreadsheetApp.getUi()
    ui.createMenu('カスタムメニュー')
        .addItem('予約スケジュールを生成', 'generateDeliverySchedule')
        .addItem('スケジュールを確定処理する', 'confirmDeliverySchedule')
        .addToUi()
}
