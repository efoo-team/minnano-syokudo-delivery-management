// 編集時のトリガー関数
function onEdit(e) {
    // 編集されたレンジを取得
    const range = e.range
    const sheet = range.getSheet()

    // 本当はmatch的に分岐して処理したほうが良い
    // 「次回配送予定スケジュール」シート以外の変更は無視
    if (sheet.getName() === '次回配送予定スケジュール' && range.getColumn() === 4) onChangeRestaurant(range.getRow(), range.getValue())
}
