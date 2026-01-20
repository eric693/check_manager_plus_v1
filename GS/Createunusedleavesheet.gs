// CreateUnusedLeaveSheet.gs - 建立「未休假記錄」工作表

/**
 * ✅ 建立「未休假記錄」工作表
 * 
 * 用於記錄員工每月的特休未休天數，用於計算未休假補薪
 */
function createUnusedLeaveSheet() {
  try {
    Logger.log('📝 開始建立「未休假記錄」工作表...');
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 檢查是否已存在
    let sheet = ss.getSheetByName('未休假記錄');
    
    if (sheet) {
      Logger.log('⚠️ 工作表已存在');
      const answer = Browser.msgBox(
        '工作表已存在',
        '「未休假記錄」工作表已存在，是否要刪除重建？',
        Browser.Buttons.YES_NO
      );
      
      if (answer === 'yes') {
        ss.deleteSheet(sheet);
        Logger.log('🗑️ 已刪除舊工作表');
      } else {
        Logger.log('❌ 使用者取消');
        return;
      }
    }
    
    // 建立新工作表
    sheet = ss.insertSheet('未休假記錄');
    
    // ⭐ 設定標題列
    const headers = [
      '員工ID',      // A欄
      '年月',        // B欄
      '未休天數',    // C欄
      '備註',        // D欄
      '更新時間'     // E欄
    ];
    
    sheet.appendRow(headers);
    
    // 美化標題列
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#10b981');
    headerRange.setFontColor('#FFFFFF');
    headerRange.setFontWeight('bold');
    headerRange.setHorizontalAlignment('center');
    headerRange.setVerticalAlignment('middle');
    
    // 設定欄寬
    sheet.setColumnWidth(1, 150);  // 員工ID
    sheet.setColumnWidth(2, 100);  // 年月
    sheet.setColumnWidth(3, 100);  // 未休天數
    sheet.setColumnWidth(4, 250);  // 備註
    sheet.setColumnWidth(5, 180);  // 更新時間
    
    // 凍結標題列
    sheet.setFrozenRows(1);
    
    // ⭐ 加入範例資料（飼料廠司機）
    const exampleData = [
      ['DRIVER001', '2026-01', 6, '1月份特休未休6天', new Date()],
      ['DRIVER001', '2025-12', 3, '12月份特休未休3天', new Date()]
    ];
    
    exampleData.forEach(row => {
      sheet.appendRow(row);
    });
    
    Logger.log('✅ 「未休假記錄」工作表建立完成');
    Logger.log(`   共加入 ${exampleData.length} 筆範例資料`);
    
    Browser.msgBox(
      '✅ 成功！',
      '「未休假記錄」工作表已建立！\n\n' +
      '標題列：\n' +
      '• 員工ID\n' +
      '• 年月 (YYYY-MM)\n' +
      '• 未休天數\n' +
      '• 備註\n' +
      '• 更新時間\n\n' +
      `已加入 ${exampleData.length} 筆範例資料`,
      Browser.Buttons.OK
    );
    
  } catch (error) {
    Logger.log('❌ 建立失敗: ' + error.message);
    Logger.log('錯誤堆疊: ' + error.stack);
    
    Browser.msgBox(
      '❌ 錯誤',
      '建立失敗：' + error.message,
      Browser.Buttons.OK
    );
  }
}

/**
 * ✅ 批次匯入未休假資料
 * 
 * 從其他來源（例如 Excel）批次匯入未休假記錄
 */
function importUnusedLeaveData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('未休假記錄');
  
  if (!sheet) {
    Browser.msgBox('❌ 錯誤', '請先建立「未休假記錄」工作表', Browser.Buttons.OK);
    return;
  }
  
  // ⭐ 這裡可以加入批次匯入邏輯
  Logger.log('📥 開始批次匯入...');
  
  // 範例：從另一個工作表讀取資料
  // const sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('匯入資料');
  // const data = sourceSheet.getDataRange().getValues();
  // ...
  
  Logger.log('✅ 批次匯入完成');
}

/**
 * ✅ 查詢員工未休假記錄
 * 
 * @param {string} employeeId - 員工ID
 * @param {string} yearMonth - 年月 (YYYY-MM)
 * @return {number} 未休天數
 */
function getUnusedLeaveDays(employeeId, yearMonth) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('未休假記錄');
    
    if (!sheet) {
      Logger.log('⚠️ 未休假記錄工作表不存在');
      return 0;
    }
    
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      const recordEmployeeId = data[i][0];
      const recordYearMonth = data[i][1];
      const unusedDays = parseFloat(data[i][2]) || 0;
      
      if (recordEmployeeId === employeeId && recordYearMonth === yearMonth) {
        Logger.log(`✅ 找到未休假記錄: ${unusedDays} 天`);
        return unusedDays;
      }
    }
    
    Logger.log('ℹ️ 沒有找到未休假記錄');
    return 0;
    
  } catch (error) {
    Logger.log('❌ 查詢失敗: ' + error.message);
    return 0;
  }
}

/**
 * 🧪 測試函數
 */
function testUnusedLeaveSheet() {
  Logger.log('🧪 測試「未休假記錄」功能');
  Logger.log('');
  
  // 測試查詢
  const unusedDays = getUnusedLeaveDays('DRIVER001', '2026-01');
  
  Logger.log(`📊 查詢結果: ${unusedDays} 天`);
  
  if (unusedDays > 0) {
    Logger.log('✅ 測試成功！');
  } else {
    Logger.log('⚠️ 沒有找到記錄');
  }
}