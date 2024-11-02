function generateDeliverySchedule() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName('案件マスター'); // シート名を修正
  const scheduleSheet = ss.getSheetByName('次回配送予定スケジュール');

  // シートの存在チェック
  if (!masterSheet || !scheduleSheet) {
    SpreadsheetApp.getUi().alert('必要なシートが見つかりません。\n「案件マスター」と「次回配送予定スケジュール」シートが存在することを確認してください。');
    return;
  }

  try {
    // 最終行を取得
    const lastRow = masterSheet.getLastRow();
    if (lastRow < 2) {
      SpreadsheetApp.getUi().alert('案件マスターにデータが存在しません。');
      return;
    }

    // データ範囲を取得（A2から開始）
    const masterData = masterSheet.getRange(2, 1, lastRow - 1, masterSheet.getLastColumn()).getValues();

    // 既存のスケジュールデータを取得（A3から開始）
    const scheduleLastRow = Math.max(scheduleSheet.getLastRow(), 2);
    const existingSchedules = scheduleLastRow > 2 ?
      scheduleSheet.getRange(3, 1, scheduleLastRow - 2, scheduleSheet.getLastColumn()).getValues() :
      [];

    // 既存のスケジュールから配送予定日と顧客IDの組み合わせを保存
    const existingDeliveries = new Set();
    existingSchedules.forEach(row => {
      if (row[0]) { // 配送予約IDが存在する行のみ処理
        const key = `${formatDate(new Date(row[1]))}_${row[3]}`; // 実施予定日_顧客ID
        existingDeliveries.add(key);
      }
    });

    // 新しいスケジュールデータを格納する配列
    let newSchedules = [];

    // 次の配送予約IDを取得
    const nextScheduleId = existingSchedules.length > 0 ?
      existingSchedules.reduce((maxId, row) => Math.max(maxId, row[0] || 0), 0) + 1 : 1;

    masterData.forEach((row, index) => {
      const clientId = row[0];
      if (!clientId) return; // 空の行をスキップ

      const contractStatus = row[2];
      const companyName = row[3];
      const mealCount = row[10];
      const deliveryDay = row[11];
      const firstDeliveryDate = row[12];
      const trialDeliveryDate = row[13];
      const deliveryGroup = row[14];
      const deliveryPriority = row[15];
      const deliveryTime = row[16];
      const pickupTime = row[17];
      const deliveryNotes = row[18];
      const pickupNotes = row[19];

      // 備品情報のインデックス
      const equipmentStartCol = 23; // のぼり旗から始まる列
      const equipmentData = row.slice(equipmentStartCol, equipmentStartCol + 13);

      if (contractStatus === '試食会' && trialDeliveryDate) {
        // 試食会の場合は1回のみのスケジュール
        const key = `${formatDate(new Date(trialDeliveryDate))}_${clientId}`;
        if (!existingDeliveries.has(key)) {
          newSchedules.push(createScheduleRow(
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
            equipmentData
          ));
        }
      }
      else if (contractStatus === '本導入' && firstDeliveryDate && deliveryDay) {
        // 本導入の場合は1ヶ月分のスケジュール
        const monthlySchedules = generateMonthlySchedule(
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
          nextScheduleId + newSchedules.length,
          existingDeliveries
        );
        newSchedules = newSchedules.concat(monthlySchedules);
      }
    });

    // 新しいスケジュールがある場合のみ追加
    if (newSchedules.length > 0) {
      const insertRow = Math.max(scheduleSheet.getLastRow() + 1, 3); // 最低でも3行目から開始
      scheduleSheet.getRange(insertRow, 1, newSchedules.length, newSchedules[0].length)
        .setValues(newSchedules);
      SpreadsheetApp.getUi().alert(`${newSchedules.length}件のスケジュールを追加しました。`);
    } else {
      SpreadsheetApp.getUi().alert('新しく追加するスケジュールはありませんでした。');
    }

  } catch (error) {
    console.error('エラーが発生しました:', error);
    SpreadsheetApp.getUi().alert('エラーが発生しました。\nスプレッドシートの形式を確認してください。\n\nエラー詳細: ' + error.toString());
  }
}

function createScheduleRow(
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
  equipmentData
) {
  // スケジュール行の作成
  return [
    id,                    // 配送予約ID
    date,                  // 実施予定日
    '',                    // 店舗名
    clientId,              // 顧客ID
    companyName,           // 顧客名
    '',                    // ジャンル
    mealCount,             // デフォルト食数
    '',                    // 確定食数
    '',                    // →
    deliveryGroup,         // 配送グループ
    '',                    // ステータス
    deliveryPriority,      // 配送優先番号
    deliveryTime,          // 納品時間目安
    pickupTime,            // 回収時間目安
    deliveryNotes,         // 納品時備考
    pickupNotes,           // 回収時備考
    '',                    // →
    ...equipmentData,      // 備品情報
    '',                    // →
    // 残りの列を空白で埋める
    ...Array(20).fill('')  // 追加の列用の空白
  ];
}

function generateMonthlySchedule(
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
  startId,
  existingDeliveries
) {
  const schedules = [];
  const startDate = new Date(firstDeliveryDate);
  const endDate = new Date(startDate);
  endDate.setMonth(endDate.getMonth() + 1);

  // 配送曜日を数値に変換（0:日曜, 1:月曜, ...）
  const dayMapping = {
    '日': 0, '月': 1, '火': 2, '水': 3,
    '木': 4, '金': 5, '土': 6
  };
  const deliveryDayNum = dayMapping[deliveryDay];

  let currentDate = new Date(startDate);
  let idCounter = 0;

  while (currentDate < endDate) {
    if (currentDate.getDay() === deliveryDayNum) {
      const key = `${formatDate(currentDate)}_${clientId}`;
      if (!existingDeliveries.has(key)) {
        schedules.push(createScheduleRow(
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
          equipmentData
        ));
        idCounter++;
      }
    }
    currentDate.setDate(currentDate.getDate() + 1);
  }

  return schedules;
}

function formatDate(date) {
  const year = date.getFullYear();
  const month = (date.getMonth() + 1).toString().padStart(2, '0');
  const day = date.getDate().toString().padStart(2, '0');
  return `${year}/${month}/${day}`;
}
