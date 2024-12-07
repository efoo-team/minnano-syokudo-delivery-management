namespace UseCase.View.ChangeRestaurant {
    export const execute = (row: number, restaurant: string | null): void => {
        if (!restaurant) return

        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('次回配送予定スケジュール')
        const restaurantSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('飲食店マスター')
        const restaurantData = restaurantSheet.getDataRange().getValues().slice(1)
        const restaurantHeaders = restaurantSheet.getDataRange().getValues()[0]

        // 一致する店舗名の行を取得
        const restaurantRow = restaurantData.findIndex((row) => {
            console.log(row[Helper.Basic.getColumnIndex(restaurantHeaders, '店舗名')], restaurant)
            return row[Helper.Basic.getColumnIndex(restaurantHeaders, '店舗名')] === restaurant
        })

        console.log(restaurantRow)
        // 一致する店舗名が存在しない場合は処理を終了
        if (restaurantRow === -1) return
        SpreadsheetApp.getActiveSpreadsheet().toast(`${restaurant}  の備品を取得中...`, '情報取得中')

        const equipments = Helper.Basic.getColumnRangeData(restaurantHeaders, restaurantData[restaurantRow], '貸与備品→', '←EOC')
        console.log(equipments, Object.entries(equipments).length)

        // 飲食店備品→ 以降にデータを挿入していく
        const targetRow = sheet.getRange(
            row,
            Helper.Basic.getColumnIndex(sheet.getDataRange().getValues()[0], '飲食店備品→') + 2,
            1,
            Object.keys(equipments).length,
        )
        targetRow.setValues([Object.entries(equipments).map(([, value]) => value)])
        SpreadsheetApp.getActiveSpreadsheet().toast(`更新完了`, '情報取得完了')
    }
}
