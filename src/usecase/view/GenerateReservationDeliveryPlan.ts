namespace UseCase.View.GenerateReservationDeliveryPlan {
    export const execute = (): void => {
        const sheet = SpreadsheetApp.getActiveSpreadsheet()
        const masterSheet = sheet.getSheetByName('案件マスター')
        const scheduleSheet = sheet.getSheetByName('次回配送予定スケジュール')
        const restaurants = sheet.getSheetByName('飲食店マスター')

        // シートの存在チェック
        if (!masterSheet || !scheduleSheet) {
            SpreadsheetApp.getUi().alert('必要なシートが見つかりません。\nシートが存在することを確認してください。')
            return
        }

        try {
            // 最終行を取得
            const lastRow = masterSheet.getLastRow()
            if (lastRow < 2) {
                SpreadsheetApp.getUi().alert('案件マスターにデータが存在しません。')
                return
            }

            // データ範囲を取得（A2から開始）
            const masterData = masterSheet.getRange(2, 1, lastRow - 1, masterSheet.getLastColumn()).getValues()

            // 既存のスケジュールデータを取得（A3から開始）
            const scheduleLastRow = Math.max(scheduleSheet.getLastRow(), 2)
            const existingSchedules =
                scheduleLastRow > 2 ? scheduleSheet.getRange(3, 1, scheduleLastRow - 2, scheduleSheet.getLastColumn()).getValues() : []

            // 既存のスケジュールから配送予定日と顧客IDの組み合わせを保存
            const existingDeliveries = new Set()
            existingSchedules.forEach((row) => {
                if (row[0]) {
                    // 配送予約IDが存在する行のみ処理
                    const key = `${formatDate(new Date(row[2]))}_${row[4]}` // 実施予定日_顧客ID
                    existingDeliveries.add(key)
                }
            })

            // 新しいスケジュールデータを格納する配列
            let newSchedules = []

            // 次の配送予約IDを取得
            const nextScheduleId =
                existingSchedules.length > 0 ? existingSchedules.reduce((maxId, row) => Math.max(maxId, row[0] || 0), 0) + 1 : 1

            // masterDataで初回配送が完了しているクライアントの名前
            const firstDeliveredCompanies = masterData.filter((datum) => datum[13] === '初回配送').map((row) => row[3])
            console.log('初回配送済み', firstDeliveredCompanies)

            masterData.forEach((row, index) => {
                const clientId = row[0]
                if (!clientId) return // 空の行をスキップ

                const contractStatus = row[2]
                const companyName = row[3]
                const mealCount = row[10]
                const deliveryDay = row[11]
                const firstDeliveryDate = row[14]
                const trialDeliveryDate = row[15]
                const deliveryGroup = row[16]
                const deliveryPriority = row[17]
                const deliveryTime = row[18]
                const pickupTime = row[19]
                const deliveryNotes = row[20]
                const pickupNotes = row[21]

                // 初回配送を行ったことがあるか
                const isFirstDelivered = firstDeliveredCompanies.some((company) => company === companyName)
                console.log(isFirstDelivered)

                // 備品情報のインデックス
                const equipmentStartCol = 24 // のぼり旗から始まる列
                const equipmentData = row.slice(equipmentStartCol, equipmentStartCol + 12)

                // 配送備品情報のインデックス
                const deliveryEquipmentStartCol = 37 // 配送備品から始まる列
                const deliveryEquipmentData = row.slice(deliveryEquipmentStartCol, deliveryEquipmentStartCol + 9)

                if (contractStatus === '試食会' && trialDeliveryDate) {
                    // 試食会の場合は1回のみのスケジュール
                    const key = `${formatDate(new Date(trialDeliveryDate))}_${clientId}`
                    if (!existingDeliveries.has(key)) {
                        newSchedules.push(
                            createScheduleRow(
                                '試食会',
                                nextScheduleId + newSchedules.length,
                                trialDeliveryDate,
                                clientId,
                                companyName,
                                mealCount,
                                deliveryGroup,
                                deliveryPriority,
                                deliveryTime,
                                pickupTime,
                                deliveryNotes,
                                pickupNotes,
                                equipmentData,
                                deliveryEquipmentData,
                            ),
                        )
                    }
                } else if (contractStatus === '本導入' && firstDeliveryDate && deliveryDay) {
                    // 本導入の場合は1ヶ月分のスケジュール
                    const monthlySchedules = generateMonthlySchedule(
                        isFirstDelivered,
                        firstDeliveryDate,
                        deliveryDay,
                        clientId,
                        companyName,
                        mealCount,
                        deliveryGroup,
                        deliveryPriority,
                        deliveryTime,
                        pickupTime,
                        deliveryNotes,
                        pickupNotes,
                        equipmentData,
                        deliveryEquipmentData,
                        nextScheduleId + newSchedules.length,
                        existingDeliveries,
                    )
                    newSchedules = newSchedules.concat(monthlySchedules)
                }
            })

            // 新しいスケジュールがある場合のみ追加
            if (newSchedules.length > 0) {
                const insertRow = Math.max(scheduleSheet.getLastRow() + 1, 2) // 最低でも2行目から開始
                // スケジュールの順番変更（実施予定日順）
                newSchedules.sort((a, b) => new Date(a[2]).getTime() - new Date(b[2]).getTime())
                scheduleSheet.getRange(insertRow, 1, newSchedules.length, newSchedules[0].length).setValues(newSchedules)
                SpreadsheetApp.getUi().alert(`${newSchedules.length}件のスケジュールを追加しました。`)
                setDropdownFromSheet(restaurants, 'B2:B', scheduleSheet, 'D2:D')
            } else {
                SpreadsheetApp.getUi().alert('新しく追加するスケジュールはありませんでした。')
            }
        } catch (error) {
            console.error('エラーが発生しました:', error)
            SpreadsheetApp.getUi().alert(
                'エラーが発生しました。\nスプレッドシートの形式を確認してください。\n\nエラー詳細: ' + error.toString(),
            )
        }
    }

    const createScheduleRow = (
        deliveryType,
        id,
        date,
        clientId,
        companyName,
        mealCount,
        deliveryGroup,
        deliveryPriority,
        deliveryTime,
        pickupTime,
        deliveryNotes,
        pickupNotes,
        equipmentData,
        deliveryEquipmentData,
    ) => {
        // スケジュール行の作成
        return [
            id, // 配送予約ID
            '未確定', // ステータス
            date, // 実施予定日
            '', // 店舗名
            clientId, // 顧客ID
            companyName, // 顧客名
            '', // ジャンル
            mealCount, // デフォルト食数
            '', // 変更後食数
            '', // →
            deliveryGroup, // 配送グループ
            deliveryType, // 配送種別
            deliveryPriority, // 配送優先番号
            deliveryTime, // 納品時間目安
            pickupTime, // 集荷時間目安
            deliveryNotes, // 納品時備考
            pickupNotes, // 集荷時備考
            '', // →
            ...equipmentData, // 備品情報
            '', // →
            ...deliveryEquipmentData,
        ]
    }

    const generateMonthlySchedule = (
        isFirstDelivered,
        firstDeliveryDate,
        deliveryDay,
        clientId,
        companyName,
        mealCount,
        deliveryGroup,
        deliveryPriority,
        deliveryTime,
        pickupTime,
        deliveryNotes,
        pickupNotes,
        equipmentData,
        deliveryEquipmentData,
        startId,
        existingDeliveries,
    ) => {
        const schedules = []
        const startDate = new Date(firstDeliveryDate)
        const endDate = new Date(startDate)
        endDate.setMonth(endDate.getMonth() + 1)

        // 配送曜日を数値に変換（0:日曜, 1:月曜, ...）
        const dayMapping = {
            日: 0,
            月: 1,
            火: 2,
            水: 3,
            木: 4,
            金: 5,
            土: 6,
        }
        const deliveryDayNum = dayMapping[deliveryDay]

        let currentDate = new Date(startDate)
        let idCounter = 0

        while (currentDate < endDate) {
            if (currentDate.getDay() === deliveryDayNum) {
                const key = `${formatDate(currentDate)}_${clientId}`
                if (!existingDeliveries.has(key)) {
                    schedules.push(
                        createScheduleRow(
                            // 初回配送済み or index が0より大きいの場合は継続配送、そうでないなら初回配送
                            isFirstDelivered || idCounter > 0 ? '継続配送' : '初回配送',
                            startId + idCounter,
                            formatDate(currentDate),
                            clientId,
                            companyName,
                            mealCount,
                            deliveryGroup,
                            deliveryPriority,
                            deliveryTime,
                            pickupTime,
                            deliveryNotes,
                            pickupNotes,
                            equipmentData,
                            deliveryEquipmentData,
                        ),
                    )
                    idCounter++
                }
            }
            currentDate.setDate(currentDate.getDate() + 1)
        }

        return schedules
    }

    const formatDate = (date) => {
        const year = date.getFullYear()
        const month = (date.getMonth() + 1).toString().padStart(2, '0')
        const day = date.getDate().toString().padStart(2, '0')
        return `${year}/${month}/${day}`
    }

    const setDropdownFromSheet = (
        sourceSheet: GoogleAppsScript.Spreadsheet.Sheet,
        sourceRangeString: string,
        targetSheet: GoogleAppsScript.Spreadsheet.Sheet,
        targetRangeString: string,
    ): void => {
        const sourceRange = sourceSheet.getRange(sourceRangeString) // 参照するデータがある範囲
        const targetRange = targetSheet.getRange(targetRangeString) // ドロップダウンを設定する範囲
        const rule = SpreadsheetApp.newDataValidation().requireValueInRange(sourceRange).setAllowInvalid(false).build()

        // 範囲に入力規則を適用
        targetRange.setDataValidation(rule)
    }
}
