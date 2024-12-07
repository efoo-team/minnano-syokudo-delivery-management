namespace UseCase.View.ConfirmDeliverySchedule {
    export const execute = (): void => {
        const date = '2024-11-25'
        const sheet = SpreadsheetApp.getActiveSpreadsheet()
        const sourceSheet = sheet.getSheetByName('次回配送予定スケジュール')
        const restaurants = sheet.getSheetByName('飲食店マスター').getDataRange().getValues().slice(1)
        const projectClients = sheet.getSheetByName('案件マスター').getDataRange().getValues().slice(2)
        const data = sourceSheet.getDataRange().getValues().slice(1)
        const headers = sourceSheet.getDataRange().getValues()[0]
        // 各配送グループのデータを処理
        const set = data.map((row) => row[Helper.Basic.getColumnIndex(headers, '配送グループ')])
        const groups = Array.from(new Set(set))

        groups.forEach((group) => {
            if (!group) return

            const targetSheet = refreshDeliverySheet(sheet, date, group)
            const targetHeaders = setHeader(targetSheet, ['順路', '積/降', '時間', '名称', '連絡先', '住所', '配送物', '備考'])
            // 指定した日付と配送確定ステータスのデータを抽出
            const filteredData = getFilteredData(data, headers, group, date)

            if (!filteredData.length) {
                SpreadsheetApp.getUi().alert('確定可能なスケジュールが存在しません')
                return
            }

            // 納品・集荷データを作成
            const scheduleData = filteredData
                .flatMap((row) => handleRow(row, headers, projectClients, restaurants))
                .sort((a, b) => {
                    const timeA = parseTime(a[1])
                    const timeB = parseTime(b[1])
                    return timeA - timeB
                })
                .map((row, index) => [index + 1, ...row]) // 順路を追加
            console.log(scheduleData, scheduleData[0].length, targetHeaders.length)

            // データを書き込み
            if (scheduleData.length > 0) {
                targetSheet.getRange(2, 1, scheduleData.length, targetHeaders.length).setValues(scheduleData)
            }

            // シートの体裁を整える
            setMultipleColumnWidths(targetSheet, [
                {
                    column: 1,
                    width: 100,
                },
                {
                    column: 2,
                    width: 100,
                },
                {
                    column: 3,
                    width: 100,
                },
                {
                    column: 4,
                    width: 150,
                },
                {
                    column: 5,
                    width: 100,
                },
                {
                    column: 6,
                    width: 300,
                },
                {
                    column: 7,
                    width: 300,
                },
                {
                    column: 8,
                    width: 300,
                },
            ])
            targetSheet.getRange('C:C').setNumberFormat('HH:mm')
        })
    }

    const parseTime = (timeStr: string) => {
        const [hours, minutes] = Utilities.formatDate(new Date(timeStr), 'JST', 'HH:mm').split(':').map(Number)
        return hours * 60 + minutes
    }

    const refreshDeliverySheet = (sheet: GoogleAppsScript.Spreadsheet.Spreadsheet, date: string, group: string) => {
        const sheetName = `${date} 配送グループ「${group}」`
        const targetSheet = sheet.getSheetByName(sheetName)
        if (targetSheet) {
            sheet.deleteSheet(targetSheet)
        }
        return sheet.insertSheet(sheetName)
    }

    const setHeader = (sheet: GoogleAppsScript.Spreadsheet.Sheet, headers: string[]): string[] => {
        sheet.getRange(1, 1, 1, headers.length).setValues([headers])
        return headers
    }

    const getFilteredData = (data: any[][], headers: string[], group: string, date: string): any[][] => {
        return data
            .filter((row) => row[Helper.Basic.getColumnIndex(headers, '実施予定日')])
            .filter((row) => row[Helper.Basic.getColumnIndex(headers, '配送グループ')] == group)
            .filter(
                (row) =>
                    Utilities.formatDate(new Date(row[Helper.Basic.getColumnIndex(headers, '実施予定日')]), 'JST', 'yyyy-MM-dd') == date &&
                    row[Helper.Basic.getColumnIndex(headers, 'ステータス')] == '確定',
            )
    }

    const handleRow = (row: any[], headers: string[], projectClients: string[][], restaurants: string[][]) => {
        const schedule = []
        // 将来、飲食店を経由する配送ルートを辿る場合はここで処理を追加
        // const restaurantName = row[Helper.Basic.getColumnIndex(headers, '店舗名')]
        // const restaurant = restaurants.find((r) => r[1] == restaurantName)
        const clientName = row[Helper.Basic.getColumnIndex(headers, '顧客名')]
        const client = projectClients.find((c) => c[3] == clientName)
        const clientHeader = [
            '顧客ID',
            '自社担当者',
            '契約ステータス',
            '企業名',
            '→',
            '担当者名',
            '担当者電話番号',
            '担当者メールアドレス',
            '配送先住所',
        ]

        // 納品データ
        schedule.push([
            '納品',
            row[Helper.Basic.getColumnIndex(headers, '納品時間目安')],
            row[Helper.Basic.getColumnIndex(headers, '顧客名')],
            client[Helper.Basic.getColumnIndex(clientHeader, '担当者電話番号')],
            client[Helper.Basic.getColumnIndex(clientHeader, '配送先住所')],
            handleDeliveryEquipmentsText('納品', row, headers), // 配送物
            row[Helper.Basic.getColumnIndex(headers, '納品時備考')],
        ])

        // 集荷データ
        schedule.push([
            '集荷',
            row[Helper.Basic.getColumnIndex(headers, '集荷時間目安')],
            row[Helper.Basic.getColumnIndex(headers, '顧客名')],
            client[Helper.Basic.getColumnIndex(clientHeader, '担当者電話番号')],
            client[Helper.Basic.getColumnIndex(clientHeader, '配送先住所')],
            handleDeliveryEquipmentsText('集荷', row, headers), // 配送物
            row[Helper.Basic.getColumnIndex(headers, '集荷時備考')],
        ])

        return schedule
    }

    const handleDeliveryEquipmentsText = (status: string, row: string[], headers: string[]) => {
        const deliveryType = row[Helper.Basic.getColumnIndex(headers, '配送種別')]
        const deliveryEquipments = Helper.Basic.getColumnRangeData(headers, row, '配送備品→', '配送消耗品→')
        const deliveryEquipmentsText = `【配送備品】\n\n${formatOutput(deliveryEquipments)}`
        const restaurantEquipments = Helper.Basic.getColumnRangeData(headers, row, '飲食店備品→', '←EOC')
        const restaurantEquipmentsText = `【飲食店備品】\n\n${formatOutput(restaurantEquipments)}`
        let text = ''

        switch (status) {
            case '納品':
                // 試食会
                if (deliveryType == '試食会') {
                    text = `${deliveryEquipmentsText}\n\n${restaurantEquipmentsText}`
                }
                // 初回配送
                if (deliveryType == '初回配送') {
                    text = `${deliveryEquipmentsText}\n\n${restaurantEquipmentsText}`
                }
                // 継続配送
                if (deliveryType == '継続配送') {
                    text = `${restaurantEquipmentsText}`
                }
                break
            case '集荷':
                if (deliveryType == '試食会') {
                    text = `${deliveryEquipmentsText}\n\n${restaurantEquipmentsText}`
                }

                if (deliveryType == '初回配送') {
                    text = `${restaurantEquipmentsText}`
                }

                if (deliveryType == '継続配送') {
                    text = `${restaurantEquipmentsText}`
                }
                break
            default:
                break
        }

        // 雑だけど、一旦情報がない時は削除
        if (text.trim() === '【飲食店備品】' || text.trim() === '【配送備品】\n\n\n\n【飲食店備品】') {
            text = ''
        }

        return text.trim()
    }

    const formatOutput = (data) => {
        let output = ''

        Object.entries(data).forEach(([key, value]) => {
            // 空文字列でない値のみを処理
            if (value !== '') {
                output += `${key}: ${value}\n`
            }
        })

        return output.trim() // 最後の余分な改行を削除
    }

    /**
     * 複数の列の幅を一括で設定する
     * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 対象のシート
     * @param {Array<{column: number, width: number}>} columnsConfig - 列の設定配列
     */
    function setMultipleColumnWidths(sheet, columnsConfig) {
        columnsConfig.forEach((config) => {
            sheet.setColumnWidth(config.column, config.width)
        })
    }
}
