namespace Helper.Basic {
    export const getColumnIndex = (headers: string[], columnName: string) => {
        return headers.indexOf(columnName)
    }

    export const getColumnRangeData = (headers, row, startColumn, endColumn) => {
        // ヘッダーから列インデックスを取得（検索に利用した列は含まない）
        const startIndex = Helper.Basic.getColumnIndex(headers, startColumn) + 1
        const endIndex = Helper.Basic.getColumnIndex(headers, endColumn) - 1

        // 開始列と終了列が見つからない場合はnullを返す
        if (startIndex === -1 || endIndex === -1) {
            return null
        }

        // 指定された範囲のデータを抽出してオブジェクトを作成
        const result = {}
        for (let i = startIndex; i <= endIndex; i++) {
            // headerが存在し、かつrowのデータも存在する場合のみ追加
            if (headers[i] && row[i] !== undefined) {
                result[headers[i]] = row[i]
            }
        }

        return result
    }
}
