// LeaveManagement.gs - 小時制請假系統（完整修正版）

/**
 * ✅ 提交請假申請（修正版 - 正確處理日期時間）
 * 
 * 修正內容：
 * 1. 正確解析前端傳來的 ISO 8601 日期時間字串
 * 2. 確保日期時間正確寫入 Sheet
 * 3. 修正請假原因欄位
 */
function submitLeaveRequest(sessionToken, leaveType, startDateTime, endDateTime, reason) {
  try {
    Logger.log('═══════════════════════════════════════');
    Logger.log('📋 開始處理請假申請（0.5小時制）');
    Logger.log('═══════════════════════════════════════');
    
    // 驗證 Session
    const employee = checkSession_(sessionToken);
    
    if (!employee.ok || !employee.user) {
      return { 
        ok: false, 
        code: "ERR_SESSION_INVALID",
        msg: "未授權或 session 已過期" 
      };
    }
    
    const user = employee.user;
    Logger.log(`   員工ID: ${user.userId}`);
    Logger.log(`   員工姓名: ${user.name}`);
    
    // 解析日期時間
    let start, end;
    
    try {
      start = new Date(startDateTime);
      end = new Date(endDateTime);
      
      if (isNaN(start.getTime()) || isNaN(end.getTime())) {
        return {
          ok: false,
          code: "ERR_INVALID_DATETIME",
          msg: "日期時間格式無效"
        };
      }
      
    } catch (parseError) {
      return {
        ok: false,
        code: "ERR_DATETIME_PARSE",
        msg: "無法解析日期時間"
      };
    }
    
    // 驗證時間順序
    if (end <= start) {
      return {
        ok: false,
        code: "ERR_INVALID_TIME_RANGE",
        msg: "結束時間必須晚於開始時間"
      };
    }

    // ⭐ 修改：檢查是否為 30 分鐘的倍數
    Logger.log('🔍 檢查時間格式（30 分鐘倍數）...');

    const startMinutes = start.getMinutes();
    const startSeconds = start.getSeconds();
    const endMinutes = end.getMinutes();
    const endSeconds = end.getSeconds();

    if (startMinutes % 30 !== 0 || startSeconds !== 0) {
      return { msg: "開始時間必須是 30 分鐘的倍數（例如：08:00, 08:30, 09:00）" };
    }
    if (endMinutes % 30 !== 0 || endSeconds !== 0) {
      return { msg: "結束時間必須是 30 分鐘的倍數（例如：08:00, 08:30, 09:00）" };
    }

    Logger.log('✅ 時間格式檢查通過（30 分鐘倍數）');

    // 計算工作時數和天數
    const { workHours, days } = calculateWorkHoursAndDays(start, end);
    
    Logger.log(`   工作時數: ${workHours} 小時`);
    Logger.log(`   天數: ${days} 天`);
    
    // ⭐ 新增：檢查是否小於最小值
    if (workHours < 0.5) {
      return {
        ok: false,
        code: "ERR_HOURS_TOO_LOW",
        msg: "請假時數最少為 0.5 小時"
      };
    }
    
    // ⭐ 新增：檢查是否為 0.5 的倍數
    if ((workHours * 2) % 1 !== 0) {
      return {
        ok: false,
        code: "ERR_INVALID_HOURS",
        msg: `請假時數必須是 0.5 小時的倍數（目前為 ${workHours}）`
      };
    }
    
    // 檢查假期餘額
    const balance = getLeaveBalance(sessionToken);
    
    if (!balance.ok) {
      return {
        ok: false,
        code: "ERR_BALANCE_CHECK",
        msg: "無法取得假期餘額"
      };
    }
    
    // 格式化日期時間
    const formattedStartDateTime = Utilities.formatDate(
      start,
      Session.getScriptTimeZone(),
      'yyyy-MM-dd HH:mm:ss'
    );
    
    const formattedEndDateTime = Utilities.formatDate(
      end,
      Session.getScriptTimeZone(),
      'yyyy-MM-dd HH:mm:ss'
    );
    
    // 取得工作表
    const sheet = getLeaveRecordsSheet();
    
    if (!sheet) {
      return {
        ok: false,
        code: "ERR_SHEET_ACCESS",
        msg: "無法存取請假記錄工作表"
      };
    }
    
    // 寫入資料
    const row = [
      new Date(),                  // A: 申請時間
      user.userId || '',           // B: 員工ID
      user.name || '',             // C: 姓名
      user.dept || '',             // D: 部門
      leaveType || '',             // E: 假別
      formattedStartDateTime,      // F: 開始時間
      formattedEndDateTime,        // G: 結束時間
      workHours,                   // H: 工作時數
      days,                        // I: 天數
      reason || '',                // J: 原因
      'PENDING',                   // K: 狀態
      '',                          // L: 審核人
      '',                          // M: 審核時間
      ''                           // N: 審核意見
    ];
    
    try {
      sheet.appendRow(row);
      Logger.log('✅ 資料寫入成功');
    } catch (writeError) {
      return {
        ok: false,
        code: "ERR_WRITE_FAILED",
        msg: "無法寫入請假記錄"
      };
    }
    
    Logger.log('✅✅✅ 請假申請提交成功');
    
    return {
      ok: true,
      code: "LEAVE_SUBMIT_SUCCESS",
      msg: "請假申請已提交",
      data: {
        leaveType: leaveType,
        startDateTime: formattedStartDateTime,
        endDateTime: formattedEndDateTime,
        workHours: workHours,
        days: days,
        reason: reason
      }
    };
    
  } catch (error) {
    Logger.log('❌❌❌ submitLeaveRequest 發生錯誤');
    Logger.log('錯誤訊息: ' + error.message);
    
    return {
      ok: false,
      code: "ERR_INTERNAL_ERROR",
      msg: "系統錯誤：" + error.message
    };
  }
}

/**
 * 計算工作時數（排除午休時間 12:00-13:00）
 * 修正版：確保正確計算跨午休時段
 */
function calculateWorkHours(startTime, endTime) {
    if (!startTime || !endTime) {
        return 0;
    }
    
    const start = new Date(startTime);
    const end = new Date(endTime);
    
    // 檢查日期是否有效
    if (isNaN(start.getTime()) || isNaN(end.getTime())) {
        console.error('❌ 無效的日期格式');
        return 0;
    }
    
    // 檢查結束時間是否早於開始時間
    if (end <= start) {
        console.error('❌ 結束時間必須晚於開始時間');
        return 0;
    }
    
    // 計算總時長（毫秒）
    const totalMs = end - start;
    
    // 轉換為小時
    let totalHours = totalMs / (1000 * 60 * 60);
    
    // 如果是同一天，檢查是否跨越午休時間 12:00-13:00
    if (start.toDateString() === end.toDateString()) {
        const startHour = start.getHours() + start.getMinutes() / 60;
        const endHour = end.getHours() + end.getMinutes() / 60;
        
        const lunchStart = 12; // 12:00
        const lunchEnd = 13;   // 13:00
        
        // 判斷是否跨越午休時間
        if (startHour < lunchEnd && endHour > lunchStart) {
            // 計算重疊的午休時間
            const overlapStart = Math.max(startHour, lunchStart);
            const overlapEnd = Math.min(endHour, lunchEnd);
            const lunchOverlap = Math.max(0, overlapEnd - overlapStart);
            
            totalHours -= lunchOverlap;
            
            console.log('🍱 扣除午休時間:', lunchOverlap.toFixed(2), '小時');
        }
    } else {
        // 跨日請假：每天都要扣除 1 小時午休
        const startDate = new Date(start);
        startDate.setHours(0, 0, 0, 0);
        
        const endDate = new Date(end);
        endDate.setHours(0, 0, 0, 0);
        
        const daysDiff = Math.floor((endDate - startDate) / (1000 * 60 * 60 * 24)) + 1;
        
        // 每天扣除 1 小時午休
        totalHours -= daysDiff;
        
        console.log('📅 跨日請假，扣除', daysDiff, '天的午休時間');
    }
    
    // 確保不會是負數
    totalHours = Math.max(0, totalHours);
    
    // 四捨五入到小數點後 2 位
    return Math.round(totalHours * 100) / 100;
}
/**
 * ✅ 取得或建立請假記錄工作表
 */
function getLeaveRecordsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('請假紀錄');
  
  if (!sheet) {
    Logger.log('📝 請假紀錄工作表不存在，自動建立...');
    
    sheet = ss.insertSheet('請假紀錄');
    
    // 建立標題列（14個欄位）
    sheet.appendRow([
      '申請時間', '員工ID', '姓名', '部門', '假別',
      '開始時間', '結束時間', '工作時數', '天數', '原因',
      '狀態', '審核人', '審核時間', '審核意見'
    ]);
    
    // 美化標題列
    const headerRange = sheet.getRange(1, 1, 1, 14);
    headerRange.setBackground('#4A90E2');
    headerRange.setFontColor('#FFFFFF');
    headerRange.setFontWeight('bold');
    headerRange.setHorizontalAlignment('center');
    
    // 凍結標題列
    sheet.setFrozenRows(1);
    
    Logger.log('✅ 請假紀錄工作表已建立');
  }
  
  return sheet;
}

/**
 * ✅ 修正：取得假期餘額（適配 19 欄結構）
 */
function getLeaveBalance(sessionToken) {
  try {
    const employee = checkSession_(sessionToken);
    
    if (!employee.ok || !employee.user) {
      return {
        ok: false,
        code: "ERR_SESSION_INVALID"
      };
    }
    
    const user = employee.user;
    Logger.log('🔍 查詢員工: ' + user.userId);
    
    const sheet = getLeaveBalanceSheet();
    
    if (!sheet) {
      Logger.log('❌ 工作表不存在，嘗試建立...');
      initializeEmployeeLeave(sessionToken);
      return getLeaveBalance(sessionToken);
    }
    
    const values = sheet.getDataRange().getValues();
    Logger.log('📊 工作表行數: ' + values.length);
    
    for (let i = 1; i < values.length; i++) {
      Logger.log(`   檢查第 ${i} 行: ${values[i][0]}`);
      
      if (values[i][0] === user.userId) {
        Logger.log('✅ 找到員工資料');
        
        // ⭐⭐⭐ 修正：所有索引 +1（因為 C 欄是到職日）
        const balance = {
          employeeName: values[i][1] || user.name,  // B: 姓名
          hireDate: values[i][2] || null,           // C: 到職日 ⭐ 新增
          ANNUAL_LEAVE: values[i][3] || 0,          // D: 特休假 ⭐
          SICK_LEAVE: values[i][4] || 0,            // E: 未住院病假 ⭐
          PERSONAL_LEAVE: values[i][5] || 0,        // F: 事假 ⭐
          BEREAVEMENT_LEAVE: values[i][6] || 0,     // G: 喪假 ⭐
          MARRIAGE_LEAVE: values[i][7] || 0,        // H: 婚假 ⭐
          MATERNITY_LEAVE: values[i][8] || 0,       // I: 產假 ⭐
          PATERNITY_LEAVE: values[i][9] || 0,       // J: 陪產檢及陪產假 ⭐
          HOSPITALIZATION_LEAVE: values[i][10] || 0, // K: 住院病假 ⭐
          MENSTRUAL_LEAVE: values[i][11] || 0,      // L: 生理假 ⭐
          FAMILY_CARE_LEAVE: values[i][12] || 0,    // M: 家庭照顧假 ⭐
          OFFICIAL_LEAVE: values[i][13] || 0,       // N: 公假 ⭐
          WORK_INJURY_LEAVE: values[i][14] || 0,    // O: 公傷假 ⭐
          NATURAL_DISASTER_LEAVE: values[i][15] || 0, // P: 天然災害停班 ⭐
          COMP_TIME_OFF: values[i][16] || 0,        // Q: 加班補休假 ⭐
          ABSENCE_WITHOUT_LEAVE: values[i][17] || 0 // R: 曠工 ⭐
        };
        
        // ⭐ 格式化到職日
        if (balance.hireDate) {
          balance.hireDate = formatDate(balance.hireDate);
        }
        
        Logger.log('📋 假期餘額（小時）:');
        Logger.log(JSON.stringify(balance, null, 2));
        
        return {
          ok: true,
          balance: balance
        };
      }
    }
    
    // 如果找不到，自動初始化
    Logger.log('⚠️ 找不到員工資料，嘗試初始化...');
    initializeEmployeeLeave(sessionToken);
    return getLeaveBalance(sessionToken);
    
  } catch (error) {
    Logger.log('❌ getLeaveBalance 錯誤: ' + error);
    Logger.log('錯誤堆疊: ' + error.stack);
    return {
      ok: false,
      code: "ERR_INTERNAL_ERROR",
      msg: error.message
    };
  }
}

/**
 * ✅ 取得或建立假期餘額工作表（新增到職日欄位）
 * 
 * 新結構（19個欄位）：
 * A - 員工ID
 * B - 姓名
 * C - 到職日 ⭐ 新增
 * D - 特休假
 * E - 未住院病假
 * F - 事假
 * G - 喪假
 * H - 婚假
 * I - 產假
 * J - 陪產檢及陪產假
 * K - 住院病假
 * L - 生理假
 * M - 家庭照顧假
 * N - 公假(含兵役假)
 * O - 公傷假
 * P - 天然災害停班
 * Q - 加班補休假
 * R - 曠工
 * S - 更新時間
 */
function getLeaveBalanceSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('假期餘額');
  
  if (!sheet) {
    Logger.log('📝 假期餘額工作表不存在，自動建立...');
    
    sheet = ss.insertSheet('假期餘額');
    
    // ⭐ 修改：建立標題列（19個欄位，新增到職日）
    sheet.appendRow([
      '員工ID',           // A
      '姓名',             // B
      '到職日',           // C ⭐ 新增
      '特休假',           // D
      '未住院病假',       // E
      '事假',             // F
      '喪假',             // G
      '婚假',             // H
      '產假',             // I
      '陪產檢及陪產假',   // J
      '住院病假',         // K
      '生理假',           // L
      '家庭照顧假',       // M
      '公假(含兵役假)',   // N
      '公傷假',           // O
      '天然災害停班',     // P
      '加班補休假',       // Q
      '曠工',             // R
      '更新時間'          // S
    ]);
    
    // 美化標題列
    const headerRange = sheet.getRange(1, 1, 1, 19);
    headerRange.setBackground('#4A90E2');
    headerRange.setFontColor('#FFFFFF');
    headerRange.setFontWeight('bold');
    headerRange.setHorizontalAlignment('center');
    
    // 凍結標題列
    sheet.setFrozenRows(1);
    
    Logger.log('✅ 假期餘額工作表已建立（含到職日欄位）');
  }
  
  return sheet;
}

function testGetLeaveBalance() {
  // ⚠️ 替換成你的 sessionToken
  const token = '7dac1161-bbac-487d-900b-3e06c1acab8d';
  
  Logger.log('🧪 開始測試 getLeaveBalance');
  Logger.log('');
  
  const result = getLeaveBalance(token);
  
  Logger.log('📤 測試結果:');
  Logger.log(JSON.stringify(result, null, 2));
  
  if (result.ok) {
    Logger.log('');
    Logger.log('✅ 測試成功！');
    Logger.log('');
    Logger.log('假期餘額:');
    for (const [key, value] of Object.entries(result.balance)) {
      Logger.log(`   ${key}: ${value}`);
    }
  } else {
    Logger.log('');
    Logger.log('❌ 測試失敗');
    Logger.log('錯誤碼: ' + result.code);
  }
}

/**
 * ✅ 初始化或更新員工特休（考量即時年資與到期）
 * 
 * 使用場景：
 * 1. 新員工首次登入 → 計算目前應得特休
 * 2. 每年週年日自動觸發 → 重置為新週期特休
 * 
 * @param {string} sessionToken
 */
function initializeEmployeeLeave(sessionToken) {
  try {
    const employee = checkSession_(sessionToken);

    if (!employee.ok || !employee.user) {
      return { ok: false, code: "ERR_SESSION_INVALID" };
    }

    const user = employee.user;
    const sheet = getLeaveBalanceSheet();
    const values = sheet.getDataRange().getValues();

    // 檢查是否已存在
    for (let i = 1; i < values.length; i++) {
      if (values[i][0] === user.userId) {
        Logger.log(`ℹ️ 員工 ${user.name} 的假期餘額已存在，跳過初始化`);
        return { ok: true, msg: "假期餘額已存在" };
      }
    }

    // ⭐ 新員工：記錄到職日為今天，並計算即時特休
    const hireDate = new Date();
    const leaveInfo = getCurrentAnnualLeaveInfo(hireDate);
    const annualLeaveHours = leaveInfo.currentHours;

    Logger.log('═══════════════════════════════════════');
    Logger.log('🎉 新員工首次登入');
    Logger.log(`   員工ID: ${user.userId}`);
    Logger.log(`   姓名: ${user.name}`);
    Logger.log(`   到職日: ${formatDate(hireDate)}`);
    Logger.log(`   初始特休: ${annualLeaveHours} 小時 (${annualLeaveHours / 8} 天)`);
    Logger.log('═══════════════════════════════════════');

    const defaultBalance = [
      user.userId,         // A: 員工ID
      user.name,           // B: 姓名
      hireDate,            // C: 到職日
      annualLeaveHours,    // D: 特休假（即時計算）
      240,                 // E: 未住院病假（30天）
      112,                 // F: 事假（14天）
      40,                  // G: 喪假（5天）
      64,                  // H: 婚假（8天）
      448,                 // I: 產假（56天）
      56,                  // J: 陪產假（7天）
      240,                 // K: 住院病假（30天）
      96,                  // L: 生理假（12天）
      56,                  // M: 家庭照顧假（7天）
      0,                   // N: 公假
      0,                   // O: 公傷假
      0,                   // P: 天然災害停班
      0,                   // Q: 加班補休假
      0,                   // R: 曠工
      new Date()           // S: 更新時間
    ];

    sheet.appendRow(defaultBalance);

    Logger.log('✅ 已初始化假期餘額');

    return {
      ok: true,
      msg: "假期餘額已初始化",
      hireDate: formatDate(hireDate),
      annualLeaveHours: annualLeaveHours
    };

  } catch (error) {
    Logger.log('❌ initializeEmployeeLeave 錯誤: ' + error);
    return { ok: false, msg: error.message };
  }
}

function calculateYearsOfService(hireDate, baseDate) {
  if (!hireDate) return 0;
  
  const hire = new Date(hireDate);
  const base = baseDate || new Date();
  
  if (isNaN(hire.getTime())) return 0;
  
  let years = base.getFullYear() - hire.getFullYear();
  
  const thisYearAnniversary = new Date(hire);
  thisYearAnniversary.setFullYear(base.getFullYear());
  
  if (base < thisYearAnniversary) {
    years -= 1;
  }
  
  const lastAnniversary = new Date(hire);
  lastAnniversary.setFullYear(hire.getFullYear() + years);
  
  const nextAnniversary = new Date(hire);
  nextAnniversary.setFullYear(lastAnniversary.getFullYear() + 1);
  
  const fraction = (base - lastAnniversary) / (nextAnniversary - lastAnniversary);
  
  return years + fraction;
}

/**
 * ✅ 根據勞基法計算特休天數（依年資級距）
 * 
 * 台灣勞基法規定（天數）：
 * - 未滿 6 個月：0 天
 * - 滿 6 個月未滿 1 年：3 天
 * - 滿 1 年未滿 2 年：7 天
 * - 滿 2 年未滿 3 年：10 天
 * - 滿 3 年未滿 5 年：14 天
 * - 滿 5 年未滿 10 年：15 天
 * - 滿 10 年以上：每年加 1 天，最高 30 天
 * 
 * @param {number} yearsOfService - 年資（小數點，例如 1.5 代表 1.5 年）
 * @returns {number} - 特休小時數（天數 × 8）
 */
function calculateAnnualLeave(completedYears) {
  // ⭐ completedYears 必須是整數（Math.floor 後的值）
  let days = 0;

  if (completedYears < 1) {
    days = 0; // 未滿1年（滿6個月的3天由 getCurrentAnnualLeaveInfo 單獨處理）
  } else if (completedYears < 2) {
    days = 7;
  } else if (completedYears < 3) {
    days = 10;
  } else if (completedYears < 5) {
    days = 14;
  } else if (completedYears < 10) {
    days = 15;
  } else {
    const extraYears = completedYears - 10;
    days = Math.min(15 + extraYears + 1, 30);
  }

  return days * 8;
}


/**
 * ✅ 即時計算員工當前「應享有」的特休時數（考量到期）
 * 
 * 邏輯說明：
 * - 特休以「到職週年日」為週期，每年重新給予
 * - 每個週期的特休必須在「下一個週年日」前使用，否則到期歸零（或依公司規定）
 * - 此函數計算「當前週期」應有的特休時數
 * 
 * 範例：
 *   到職日：2023-07-01
 *   今天：2025-03-04
 *   → 年資 = 1.67 年 → 應得 7 天 = 56 小時（進入第 2 個週期）
 *   → 當前週期：2024-07-01 ~ 2025-06-30
 *   → 本週期已給予 56 小時
 * 
 * @param {Date} hireDate - 到職日
 * @param {Date} [baseDate] - 計算基準日（預設今天）
 * @returns {{currentHours: number, periodStart: Date, periodEnd: Date, yearsOfService: number}}
 */
function getCurrentAnnualLeaveInfo(hireDate, baseDate) {
  const hire = new Date(hireDate);
  const today = baseDate || new Date();

  if (isNaN(hire.getTime())) {
    return { currentHours: 0, periodStart: null, periodEnd: null, yearsOfService: 0 };
  }

  const yearsOfService = calculateYearsOfService(hire, today);
  
  // ⭐ 關鍵修正：加入微小容差防止浮點數邊界問題
  const TOLERANCE = 1e-6;
  const completedYears = Math.floor(yearsOfService + TOLERANCE);

  if (yearsOfService < 0.5 - TOLERANCE) {
    return {
      currentHours: 0,
      periodStart: null,
      periodEnd: null,
      yearsOfService: yearsOfService
    };
  } else if (yearsOfService < 1 - TOLERANCE) {
    // 滿6個月未滿1年
    const periodStart = new Date(hire);
    periodStart.setMonth(periodStart.getMonth() + 6);

    const firstAnniversary = new Date(hire);
    firstAnniversary.setFullYear(hire.getFullYear() + 1);
    const periodEnd = new Date(firstAnniversary);
    periodEnd.setDate(periodEnd.getDate() - 1);

    return {
      currentHours: 3 * 8, // 24 小時
      periodStart: periodStart,
      periodEnd: periodEnd,
      yearsOfService: yearsOfService
    };
  } else {
    // 滿1年以上
    const periodStart = new Date(hire);
    periodStart.setFullYear(hire.getFullYear() + completedYears);

    const nextAnniversary = new Date(hire);
    nextAnniversary.setFullYear(hire.getFullYear() + completedYears + 1);
    const periodEnd = new Date(nextAnniversary);
    periodEnd.setDate(periodEnd.getDate() - 1);

    const hoursForCurrentPeriod = calculateAnnualLeave(completedYears);

    Logger.log(`📅 特休週期計算:`);
    Logger.log(`   到職日: ${formatDate(hire)}`);
    Logger.log(`   今天: ${formatDate(today)}`);
    Logger.log(`   年資: ${yearsOfService.toFixed(6)} 年`);
    Logger.log(`   完整年數: ${completedYears}`);
    Logger.log(`   當前週期: ${formatDate(periodStart)} ~ ${formatDate(periodEnd)}`);
    Logger.log(`   本週期應得: ${hoursForCurrentPeriod} 小時 (${hoursForCurrentPeriod / 8} 天)`);

    return {
      currentHours: hoursForCurrentPeriod,
      periodStart: periodStart,
      periodEnd: periodEnd,
      yearsOfService: yearsOfService
    };
  }
}

/**
 * ⭐ 輔助函數：格式化日期
 */
function formatDate(date) {
  if (!date) return '';
  try {
    return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  } catch (e) {
    return String(date);
  }
}

/**
 * ✅ 取得員工請假記錄
 */
function getEmployeeLeaveRecords(sessionToken) {
  try {
    const employee = checkSession_(sessionToken);
    
    if (!employee.ok || !employee.user) {
      return {
        ok: false,
        code: "ERR_SESSION_INVALID"
      };
    }
    
    const user = employee.user;
    const sheet = getLeaveRecordsSheet();
    const values = sheet.getDataRange().getValues();
    
    if (values.length <= 1) {
      return {
        ok: true,
        records: []
      };
    }
    
    const records = [];
    
    for (let i = 1; i < values.length; i++) {
      if (values[i][1] === user.userId) {
        const record = {
          applyTime: formatDateTime(values[i][0]),      // A
          employeeName: values[i][2],                   // C
          dept: values[i][3],                           // D
          leaveType: values[i][4],                      // E
          startDateTime: values[i][5],                  // F
          endDateTime: values[i][6],                    // G
          workHours: values[i][7],                      // H
          days: values[i][8],                           // I
          reason: values[i][9] || '',                   // J
          status: values[i][10] || 'PENDING',           // K
          reviewer: values[i][11] || '',                // L
          reviewTime: values[i][12] ? formatDateTime(values[i][12]) : '', // M
          reviewComment: values[i][13] || ''            // N
        };
        
        records.push(record);
      }
    }
    
    // 按申請時間排序（最新的在前）
    records.sort((a, b) => new Date(b.applyTime) - new Date(a.applyTime));
    
    return {
      ok: true,
      records: records
    };
    
  } catch (error) {
    Logger.log('❌ getEmployeeLeaveRecords 錯誤: ' + error);
    return {
      ok: false,
      msg: error.message
    };
  }
}

/**
 * ✅ 修正版：取得待審核請假申請（管理員用）
 */
function getPendingLeaveRequests(sessionToken) {
  try {
    const employee = checkSession_(sessionToken);
    
    if (!employee.ok || !employee.user) {
      return {
        ok: false,
        code: "ERR_SESSION_INVALID"
      };
    }
    
    // 檢查是否為管理員
    if (employee.user.dept !== '管理員') {
      return {
        ok: false,
        code: "ERR_PERMISSION_DENIED",
        msg: "需要管理員權限"
      };
    }
    
    const sheet = getLeaveRecordsSheet();
    const values = sheet.getDataRange().getValues();
    
    if (values.length <= 1) {
      return {
        ok: true,
        requests: []
      };
    }
    
    const requests = [];
    
    for (let i = 1; i < values.length; i++) {
      if (values[i][10] === 'PENDING') {  // K 欄：狀態
        
        // ⭐⭐⭐ 關鍵修正：重新計算正確的工作時數
        const startDateTime = values[i][5];  // F 欄：開始時間
        const endDateTime = values[i][6];    // G 欄：結束時間
        
        let correctWorkHours = 0;
        let correctDays = 0;
        
        try {
          const start = new Date(startDateTime);
          const end = new Date(endDateTime);
          
          if (!isNaN(start.getTime()) && !isNaN(end.getTime())) {
            // 使用修正後的計算邏輯
            const result = calculateWorkHoursAndDays(start, end);
            correctWorkHours = result.workHours;
            correctDays = result.days;
          }
        } catch (err) {
          Logger.log('⚠️ 計算工時失敗:', err);
          // 如果計算失敗，使用原始值
          correctWorkHours = values[i][7] || 0;
          correctDays = values[i][8] || 0;
        }
        
        const request = {
          rowNumber: i + 1,
          applyTime: formatDateTime(values[i][0]),
          employeeId: values[i][1],
          employeeName: values[i][2],
          dept: values[i][3],
          leaveType: values[i][4],
          startDateTime: startDateTime,
          endDateTime: endDateTime,
          workHours: correctWorkHours,    // ⭐ 使用重新計算的值
          days: correctDays,              // ⭐ 使用重新計算的值
          reason: values[i][9] || ''
        };
        
        requests.push(request);
      }
    }
    
    return {
      ok: true,
      requests: requests
    };
    
  } catch (error) {
    Logger.log('❌ getPendingLeaveRequests 錯誤: ' + error);
    return {
      ok: false,
      msg: error.message
    };
  }
}


/**
 * ✅ 格式化日期時間
 */
function formatDateTime(date) {
  if (!date) return '';
  try {
    return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  } catch (e) {
    return String(date);
  }
}

/**
 * 🧪 測試函數
 */
function testSubmitLeaveWithHours() {
  Logger.log('🧪 測試小時制請假申請');
  Logger.log('');
  
  const testParams = {
    token: '7dac1161-bbac-487d-900b-3e06c1acab8d',  // ⚠️ 替換成有效 token
    leaveType: 'BEREAVEMENT_LEAVE',
    startDateTime: '2025-12-18T09:00',
    endDateTime: '2025-12-18T12:00',
    reason: '測試請假（小時制）'
  };
  
  Logger.log('📥 測試參數:');
  Logger.log(JSON.stringify(testParams, null, 2));
  Logger.log('');
  
  const result = submitLeaveRequest(
    testParams.token,
    testParams.leaveType,
    testParams.startDateTime,
    testParams.endDateTime,
    testParams.reason
  );
  
  Logger.log('');
  Logger.log('📤 測試結果:');
  Logger.log(JSON.stringify(result, null, 2));
  
  if (result.ok) {
    Logger.log('');
    Logger.log('✅✅✅ 測試成功！');
    Logger.log('請檢查 Google Sheet 的「請假紀錄」工作表');
  } else {
    Logger.log('');
    Logger.log('❌ 測試失敗');
  }
}

/**
 * 🔄 遷移假期餘額工作表（8欄 → 17欄）
 * 
 * 使用方式：
 * 1. 在 Apps Script 編輯器中執行此函數
 * 2. 會自動備份舊資料
 * 3. 重建新結構並遷移資料
 */
function migrateLeaveBalanceSheet() {
  Logger.log('═══════════════════════════════════════');
  Logger.log('🔄 開始遷移假期餘額工作表');
  Logger.log('═══════════════════════════════════════');
  Logger.log('');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const oldSheet = ss.getSheetByName('假期餘額');
  
  if (!oldSheet) {
    Logger.log('❌ 找不到「假期餘額」工作表');
    return;
  }
  
  // 📋 步驟 1：備份舊工作表
  Logger.log('📋 步驟 1：備份舊工作表...');
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
  const backupSheet = oldSheet.copyTo(ss);
  backupSheet.setName('假期餘額_備份_' + timestamp);
  Logger.log('✅ 已備份: ' + backupSheet.getName());
  Logger.log('');
  
  // 📋 步驟 2：讀取舊資料
  Logger.log('📋 步驟 2：讀取舊資料...');
  const oldData = oldSheet.getDataRange().getValues();
  const recordCount = oldData.length - 1; // 扣除標題列
  
  Logger.log(`   找到 ${recordCount} 筆員工資料`);
  Logger.log('');
  
  if (recordCount <= 0) {
    Logger.log('⚠️ 沒有資料需要遷移');
    return;
  }
  
  // 📋 步驟 3：刪除舊工作表
  Logger.log('📋 步驟 3：刪除舊工作表...');
  ss.deleteSheet(oldSheet);
  Logger.log('✅ 已刪除舊的「假期餘額」工作表');
  Logger.log('');
  
  // 📋 步驟 4：建立新工作表（17 個欄位）
  Logger.log('📋 步驟 4：建立新工作表（17 個欄位）...');
  const newSheet = ss.insertSheet('假期餘額');
  
  // 建立標題列
  const headers = [
    '員工ID',           // A
    '特休假',           // B - ANNUAL_LEAVE
    '未住院病假',       // C - SICK_LEAVE
    '事假',             // D - PERSONAL_LEAVE
    '喪假',             // E - BEREAVEMENT_LEAVE
    '婚假',             // F - MARRIAGE_LEAVE
    '產假',             // G - MATERNITY_LEAVE
    '陪產檢及陪產假',   // H - PATERNITY_LEAVE
    '住院病假',         // I - HOSPITALIZATION_LEAVE
    '生理假',           // J - MENSTRUAL_LEAVE
    '家庭照顧假',       // K - FAMILY_CARE_LEAVE
    '公假(含兵役假)',   // L - OFFICIAL_LEAVE
    '公傷假',           // M - WORK_INJURY_LEAVE
    '天然災害停班',     // N - NATURAL_DISASTER_LEAVE
    '加班補休假',       // O - COMP_TIME_OFF
    '曠工',             // P - ABSENCE_WITHOUT_LEAVE
    '更新時間'          // Q
  ];
  
  newSheet.appendRow(headers);
  
  // 美化標題列
  const headerRange = newSheet.getRange(1, 1, 1, 17);
  headerRange.setBackground('#4A90E2');
  headerRange.setFontColor('#FFFFFF');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
  
  // 凍結標題列
  newSheet.setFrozenRows(1);
  
  Logger.log('✅ 新工作表已建立（17 個欄位）');
  Logger.log('');
  
  // 📋 步驟 5：遷移資料
  Logger.log('📋 步驟 5：遷移資料...');
  Logger.log('');
  
  for (let i = 1; i < oldData.length; i++) {
    const oldRow = oldData[i];
    
    // 對應關係：
    // 舊: [員工ID, 特休假, 病假, 事假, 喪假, 婚假, 產假, 陪產假, 更新時間]
    // 新: [員工ID, 特休假, 未住院病假, 事假, 喪假, 婚假, 產假, 陪產檢及陪產假, 住院病假, 生理假, 家庭照顧假, 公假, 公傷假, 天然災害停班, 加班補休假, 曠工, 更新時間]
    
    const newRow = [
      oldRow[0] || '',      // A: 員工ID（保留）
      oldRow[1] || 7,       // B: 特休假（保留）
      oldRow[2] || 30,      // C: 未住院病假（舊的「病假」）
      oldRow[3] || 14,      // D: 事假（保留）
      oldRow[4] || 5,       // E: 喪假（保留）
      oldRow[5] || 8,       // F: 婚假（保留）
      oldRow[6] || 56,      // G: 產假（保留）
      oldRow[7] || 7,       // H: 陪產檢及陪產假（保留）
      30,                   // I: 住院病假（新增，預設 30 天）
      12,                   // J: 生理假（新增，預設 12 天）
      7,                    // K: 家庭照顧假（新增，預設 7 天）
      0,                  // L: 公假（新增，無上限）
      0,                  // M: 公傷假（新增，無上限）
      0,                  // N: 天然災害停班（新增，無上限）
      0,                    // O: 加班補休假（新增，初始 0）
      0,                    // P: 曠工（新增，初始 0）
      new Date()            // Q: 更新時間（更新為當前時間）
    ];
    
    newSheet.appendRow(newRow);
    
    Logger.log(`   ✅ [${i}/${recordCount}] 已遷移: ${oldRow[0]}`);
  }
  
  Logger.log('');
  Logger.log('═══════════════════════════════════════');
  Logger.log('✅✅✅ 遷移完成！');
  Logger.log('═══════════════════════════════════════');
  Logger.log('');
  Logger.log('📊 遷移摘要:');
  Logger.log(`   - 舊結構: 8 個欄位`);
  Logger.log(`   - 新結構: 17 個欄位`);
  Logger.log(`   - 遷移記錄數: ${recordCount} 筆`);
  Logger.log(`   - 備份工作表: ${backupSheet.getName()}`);
  Logger.log('');
  Logger.log('📝 新增假別:');
  Logger.log('   - 住院病假（30天）');
  Logger.log('   - 生理假（12天）');
  Logger.log('   - 家庭照顧假（7天）');
  Logger.log('   - 公假（無上限）');
  Logger.log('   - 公傷假（無上限）');
  Logger.log('   - 天然災害停班（無上限）');
  Logger.log('   - 加班補休假（初始0）');
  Logger.log('   - 曠工（初始0）');
  Logger.log('');
  
  // 顯示成功訊息給使用者
  Browser.msgBox(
    '✅ 遷移完成！',
    '已成功將 ' + recordCount + ' 筆假期餘額遷移到新結構！\n\n' +
    '舊結構：8 個欄位\n' +
    '新結構：17 個欄位（新增 8 種假別）\n\n' +
    '備份工作表: ' + backupSheet.getName() + '\n\n' +
    '請檢查「假期餘額」工作表確認資料正確。',
    Browser.Buttons.OK
  );
}

// LeaveManagement.gs - 小時制請假系統（完整修正版 + 餘額扣除）

/**
 * ✅ 審核請假申請（完全修正版：審核時更新正確工時）
 */
function reviewLeaveRequest(sessionToken, rowNumber, reviewAction, comment) {
  try {
    Logger.log('═══════════════════════════════════════');
    Logger.log('📋 開始審核請假');
    Logger.log('═══════════════════════════════════════');
    
    const employee = checkSession_(sessionToken);
    
    if (!employee.ok || !employee.user) {
      return {
        ok: false,
        code: "ERR_SESSION_INVALID"
      };
    }
    
    if (employee.user.dept !== '管理員') {
      return {
        ok: false,
        code: "ERR_PERMISSION_DENIED",
        msg: "需要管理員權限"
      };
    }
    
    const sheet = getLeaveRecordsSheet();
    const record = sheet.getRange(rowNumber, 1, 1, 14).getValues()[0];
    
    const userId = record[1];
    const employeeName = record[2];
    const leaveType = record[4];
    const workHours = record[7];  // 直接使用計算好的工時
    
    Logger.log(`   員工: ${employeeName}`);
    Logger.log(`   假別: ${leaveType}`);
    Logger.log(`   時數: ${workHours} 小時`);
    
    // 更新狀態
    const status = (reviewAction === 'approve') ? 'APPROVED' : 'REJECTED';
    
    sheet.getRange(rowNumber, 11).setValue(status);
    sheet.getRange(rowNumber, 12).setValue(employee.user.name);
    sheet.getRange(rowNumber, 13).setValue(new Date());
    sheet.getRange(rowNumber, 14).setValue(comment || '');
    
    if (reviewAction === 'approve') {
      Logger.log('💰 開始扣除假期餘額...');
      
      const deductResult = deductLeaveBalance(userId, leaveType, workHours);
      
      if (!deductResult.ok) {
        // 回滾狀態
        sheet.getRange(rowNumber, 11).setValue('PENDING');
        
        return {
          ok: false,
          code: "ERR_DEDUCT_FAILED",
          msg: "扣除餘額失敗: " + deductResult.msg
        };
      }
      
      Logger.log(`   剩餘: ${deductResult.remaining} 小時`);

      // ⭐ 審核通過後自動重算薪資（只在 approve 時執行）
      const yearMonth = (() => {
        const startDateTime = record[5];
        const d = startDateTime instanceof Date ? startDateTime : new Date(startDateTime);
        return Utilities.formatDate(d, 'Asia/Taipei', 'yyyy-MM');
      })();
      
      const recalcResult = _recalcOne(userId, yearMonth);
      if (recalcResult.success) {
        Logger.log(`✅ 已自動重算 ${userId} ${yearMonth} 薪資`);
      } else {
        Logger.log(`⚠️ 自動重算失敗: ${recalcResult.message}`);
      }
    }

    Logger.log('✅✅✅ 審核完成');

    return {
      ok: true,
      msg: "審核完成"
    };
    
  } catch (error) {
    Logger.log('❌ reviewLeaveRequest 錯誤: ' + error);
    return {
      ok: false,
      msg: error.message
    };
  }
}
/**
 * ✅ 完全修正版：計算工作時數和天數（09:00-18:00，扣除 12:00-13:00）
 */
function calculateWorkHoursAndDays(start, end) {
  try {
    Logger.log('💡 計算工作時數和天數（0.5小時精度）');
    Logger.log(`   開始: ${start.toISOString()}`);
    Logger.log(`   結束: ${end.toISOString()}`);
    
    // 工作時間設定
    const WORK_START_HOUR = 8;      // 上班 08:00
    const WORK_END_HOUR = 17;       // 下班 17:00
    const LUNCH_START = 12;         // 午休開始 12:00
    const LUNCH_END = 13;           // 午休結束 13:00
    const DAILY_WORK_HOURS = 8;     // 每日工作時數
    
    // 判斷是否同一天
    const startDate = new Date(start.getFullYear(), start.getMonth(), start.getDate());
    const endDate = new Date(end.getFullYear(), end.getMonth(), end.getDate());
    const isSameDay = startDate.getTime() === endDate.getTime();
    
    // 1️⃣ 同一天請假
    if (isSameDay) {
      Logger.log('   ℹ️ 同日請假');
      
      // 限制在工作時間內，支持分鐘精度
      const startHour = Math.max(
        start.getHours() + start.getMinutes() / 60,
        WORK_START_HOUR
      );
      const endHour = Math.min(
        end.getHours() + end.getMinutes() / 60,
        WORK_END_HOUR
      );
      
      if (startHour >= endHour) {
        Logger.log('   ⚠️ 請假時間不在工作時段內');
        return { workHours: 0, days: 0 };
      }
      
      let workHours = endHour - startHour;
      
      // 扣除午休時間
      if (startHour < LUNCH_END && endHour > LUNCH_START) {
        const lunchOverlapStart = Math.max(startHour, LUNCH_START);
        const lunchOverlapEnd = Math.min(endHour, LUNCH_END);
        const lunchOverlap = Math.max(0, lunchOverlapEnd - lunchOverlapStart);
        workHours -= lunchOverlap;
        
        Logger.log(`   🍱 扣除午休時間: ${lunchOverlap.toFixed(2)} 小時`);
      }
      
      workHours = Math.max(0, workHours);
      
      // ⭐ 四捨五入到 0.5 小時精度
      const finalHours = Math.round(workHours * 2) / 2;
      const days = Math.round((finalHours / 8) * 100) / 100;
      
      Logger.log(`   ✅ 同日請假：${finalHours} 小時 = ${days} 天`);
      
      return {
        workHours: finalHours,
        days: days
      };
    }
    
    // 2️⃣ 跨日請假
    else {
      Logger.log('   ℹ️ 跨日請假');
      
      let totalWorkHours = 0;
      
      // 第一天
      const firstDayStartHour = Math.max(
        start.getHours() + start.getMinutes() / 60,
        WORK_START_HOUR
      );
      
      const firstDayEndHour = WORK_END_HOUR;
      
      let firstDayHours = Math.max(0, firstDayEndHour - firstDayStartHour);
      
      if (firstDayStartHour < LUNCH_END && firstDayEndHour > LUNCH_START) {
        const lunchStart = Math.max(firstDayStartHour, LUNCH_START);
        const lunchEnd = Math.min(firstDayEndHour, LUNCH_END);
        const lunchOverlap = Math.max(0, lunchEnd - lunchStart);
        firstDayHours -= lunchOverlap;
      }
      
      firstDayHours = Math.max(0, firstDayHours);
      totalWorkHours += firstDayHours;
      
      Logger.log(`   📅 第一天: ${firstDayHours.toFixed(2)} 小時`);
      
      // 中間完整工作日
      const daysDiff = Math.floor((endDate - startDate) / (1000 * 60 * 60 * 24));
      
      if (daysDiff > 1) {
        const middleDays = daysDiff - 1;
        const middleHours = middleDays * DAILY_WORK_HOURS;
        totalWorkHours += middleHours;
        
        Logger.log(`   📅 中間 ${middleDays} 天: ${middleHours} 小時`);
      }
      
      // 最後一天
      const lastDayEndHour = Math.min(
        Math.max(
          end.getHours() + end.getMinutes() / 60,
          WORK_START_HOUR
        ),
        WORK_END_HOUR
      );
      
      const lastDayStartHour = WORK_START_HOUR;
      
      let lastDayHours = Math.max(0, lastDayEndHour - lastDayStartHour);
      
      if (lastDayStartHour < LUNCH_END && lastDayEndHour > LUNCH_START) {
        const lunchStart = Math.max(lastDayStartHour, LUNCH_START);
        const lunchEnd = Math.min(lastDayEndHour, LUNCH_END);
        const lunchOverlap = Math.max(0, lunchEnd - lunchStart);
        lastDayHours -= lunchOverlap;
      }
      
      lastDayHours = Math.max(0, lastDayHours);
      totalWorkHours += lastDayHours;
      
      Logger.log(`   📅 最後一天: ${lastDayHours.toFixed(2)} 小時`);
      
      // ⭐ 四捨五入到 0.5 小時精度
      const finalHours = Math.round(totalWorkHours * 2) / 2;
      const days = Math.round((finalHours / 8) * 100) / 100;
      
      Logger.log(`   ✅ 總工時: ${finalHours} 小時 = ${days} 天`);
      
      return {
        workHours: finalHours,
        days: days
      };
    }
    
  } catch (error) {
    Logger.log(`❌ calculateWorkHoursAndDays 錯誤: ${error.message}`);
    return { workHours: 0, days: 0 };
  }
}
/**
 * ✅ 修正：取得已核准的請假記錄（適配小時制請假系統）
 * 
 * 實際欄位順序（新結構）：
 * A - 申請時間
 * B - 員工ID
 * C - 姓名
 * D - 部門
 * E - 假別
 * F - 開始時間 (datetime)
 * G - 結束時間 (datetime)
 * H - 工作時數 (hours)
 * I - 天數 (days)
 * J - 原因
 * K - 狀態
 * L - 審核人
 * M - 審核時間
 * N - 審核意見
 */
function getApprovedLeaveRecords(monthParam, userIdParam) {
  try {
    Logger.log('═══════════════════════════════════════');
    Logger.log('📋 getApprovedLeaveRecords 開始（小時制版本）');
    Logger.log('═══════════════════════════════════════');
    Logger.log(`   monthParam: ${monthParam}`);
    Logger.log(`   userIdParam: ${userIdParam}`);
    Logger.log('');
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('請假紀錄');
    
    if (!sheet) {
      Logger.log('⚠️ 找不到請假紀錄工作表');
      return [];
    }
    
    const values = sheet.getDataRange().getValues();
    
    if (values.length <= 1) {
      Logger.log('⚠️ 工作表只有標題，沒有資料');
      return [];
    }
    
    Logger.log(`✅ 工作表有 ${values.length - 1} 筆資料`);
    Logger.log('');
    
    // ✅ 根據新的欄位順序（14 個欄位）
    const leaveRecords = [];
    
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      
      // 固定欄位索引（從 0 開始）
      const applyTime = row[0];            // A 欄 (索引 0)
      const employeeId = row[1];           // B 欄 (索引 1)
      const employeeName = row[2];         // C 欄 (索引 2)
      const dept = row[3];                 // D 欄 (索引 3)
      const leaveType = row[4];            // E 欄 (索引 4)
      const startDateTime = row[5];        // F 欄 (索引 5) - 開始時間
      const endDateTime = row[6];          // G 欄 (索引 6) - 結束時間
      const workHours = row[7];            // H 欄 (索引 7) - 工作時數
      const days = row[8];                 // I 欄 (索引 8) - 天數
      const reason = row[9];               // J 欄 (索引 9)
      const status = row[10];              // K 欄 (索引 10) ⭐ 關鍵修正
      const reviewer = row[11];            // L 欄 (索引 11)
      const reviewTime = row[12];          // M 欄 (索引 12)
      const reviewComment = row[13];       // N 欄 (索引 13)
      
      Logger.log(`═══════════════════════════════════════`);
      Logger.log(`📋 第 ${i + 1} 行:`);
      Logger.log(`   員工ID: ${employeeId}`);
      Logger.log(`   員工姓名: ${employeeName}`);
      Logger.log(`   狀態: "${status}"`);
      Logger.log(`   開始時間: ${startDateTime}`);
      Logger.log(`   結束時間: ${endDateTime}`);
      Logger.log(`   工作時數: ${workHours} 小時`);
      Logger.log(`   天數: ${days} 天`);
      
      // ⭐ 檢查狀態（只取已核准的）
      if (String(status).trim() !== 'APPROVED') {
        Logger.log(`   ⏭️ 狀態不是 APPROVED (實際: "${status}")，跳過`);
        Logger.log('');
        continue;
      }
      
      // 格式化日期時間
      let formattedStartDate, formattedEndDate;
      
      try {
        // 處理可能的日期格式
        const startDate = new Date(startDateTime);
        const endDate = new Date(endDateTime);
        
        if (isNaN(startDate.getTime()) || isNaN(endDate.getTime())) {
          Logger.log(`   ⚠️ 日期格式無效，跳過`);
          Logger.log('');
          continue;
        }
        
        formattedStartDate = formatDate(startDate);
        formattedEndDate = formatDate(endDate);
        
        Logger.log(`   格式化開始日期: ${formattedStartDate}`);
        Logger.log(`   格式化結束日期: ${formattedEndDate}`);
        
      } catch (dateError) {
        Logger.log(`   ⚠️ 日期解析錯誤: ${dateError.message}`);
        Logger.log('');
        continue;
      }
      
      // ⭐ 檢查月份（使用開始日期的月份）
      if (!formattedStartDate.startsWith(monthParam)) {
        Logger.log(`   ⏭️ 月份不符 (需要: ${monthParam}, 實際: ${formattedStartDate})，跳過`);
        Logger.log('');
        continue;
      }
      
      // ⭐ 檢查員工ID（如果有指定）
      if (userIdParam && employeeId !== userIdParam) {
        Logger.log(`   ⏭️ 員工ID不符 (需要: ${userIdParam}, 實際: ${employeeId})，跳過`);
        Logger.log('');
        continue;
      }
      
      Logger.log(`   ✅ 符合所有條件！`);
      
      // ⭐ 生成請假期間的每一天
      const start = new Date(startDateTime);
      const end = new Date(endDateTime);
      
      // 計算跨越了幾天
      const startDay = new Date(start.getFullYear(), start.getMonth(), start.getDate());
      const endDay = new Date(end.getFullYear(), end.getMonth(), end.getDate());
      const totalDays = Math.floor((endDay - startDay) / (1000 * 60 * 60 * 24)) + 1;
      
      Logger.log(`   📅 請假天數範圍: ${totalDays} 天`);
      
      // 為每一天生成記錄
      for (let d = new Date(startDay); d <= endDay; d.setDate(d.getDate() + 1)) {
        const dateStr = formatDate(d);
        
        // 確保日期在查詢月份內
        if (dateStr.startsWith(monthParam)) {
          leaveRecords.push({
            employeeId: employeeId,
            employeeName: employeeName,
            date: dateStr,
            leaveType: leaveType,
            workHours: parseFloat(workHours) || 0,      // ⭐ 新增：工作時數
            days: parseFloat(days) || 0,
            status: status,
            reason: reason || '',
            startDateTime: startDateTime,                // ⭐ 新增：完整開始時間
            endDateTime: endDateTime,                    // ⭐ 新增：完整結束時間
            reviewer: reviewer || '',                    // ⭐ 新增：審核人
            reviewTime: reviewTime || '',                // ⭐ 新增：審核時間
            reviewComment: reviewComment || ''           // ⭐ 新增：審核意見
          });
          
          Logger.log(`      ➕ 加入日期: ${dateStr}`);
        }
      }
      
      Logger.log('');
    }
    
    Logger.log('═══════════════════════════════════════');
    Logger.log(`✅ getApprovedLeaveRecords 完成`);
    Logger.log(`   共找到 ${leaveRecords.length} 筆已核准的請假記錄`);
    Logger.log('═══════════════════════════════════════');
    
    return leaveRecords;
    
  } catch (error) {
    Logger.log('═══════════════════════════════════════');
    Logger.log('❌ getApprovedLeaveRecords 錯誤');
    Logger.log('   錯誤訊息: ' + error.message);
    Logger.log('   錯誤堆疊: ' + error.stack);
    Logger.log('═══════════════════════════════════════');
    return [];
  }
}
function deductLeaveBalance(userId, leaveType, hours) {
  try {
    Logger.log('📊 扣除假期餘額');
    Logger.log(`   員工ID: ${userId}`);
    Logger.log(`   假別: ${leaveType}`);
    Logger.log(`   小時數: ${hours}`);
    Logger.log('');
    
    const sheet = getLeaveBalanceSheet();
    const values = sheet.getDataRange().getValues();
    
    // ⭐⭐⭐ 修正：所有欄位 +1（因為新增了到職日）
    const leaveTypeColumnMap = {
      'ANNUAL_LEAVE': 4,              // D 欄（索引 +1）⭐
      'SICK_LEAVE': 5,                // E 欄 ⭐
      'PERSONAL_LEAVE': 6,            // F 欄 ⭐
      'BEREAVEMENT_LEAVE': 7,         // G 欄 ⭐
      'MARRIAGE_LEAVE': 8,            // H 欄 ⭐
      'MATERNITY_LEAVE': 9,           // I 欄 ⭐
      'PATERNITY_LEAVE': 10,          // J 欄 ⭐
      'HOSPITALIZATION_LEAVE': 11,    // K 欄 ⭐
      'MENSTRUAL_LEAVE': 12,          // L 欄 ⭐
      'FAMILY_CARE_LEAVE': 13,        // M 欄 ⭐
      'OFFICIAL_LEAVE': 14,           // N 欄 ⭐
      'WORK_INJURY_LEAVE': 15,        // O 欄 ⭐
      'NATURAL_DISASTER_LEAVE': 16,   // P 欄 ⭐
      'COMP_TIME_OFF': 17,            // Q 欄 ⭐
      'ABSENCE_WITHOUT_LEAVE': 18     // R 欄 ⭐
    };
    
    const columnIndex = leaveTypeColumnMap[leaveType];
    
    if (!columnIndex) {
      Logger.log('❌ 無效的假別: ' + leaveType);
      return {
        ok: false,
        msg: "無效的假別"
      };
    }
    
    // 尋找員工記錄
    for (let i = 1; i < values.length; i++) {
      if (values[i][0] === userId) {
        Logger.log(`✅ 找到員工記錄（第 ${i + 1} 行）`);
        Logger.log(`   姓名: ${values[i][1]}`);
        
        const currentBalance = values[i][columnIndex - 1]; // 因為陣列從 0 開始
        
        Logger.log(`   目前餘額: ${currentBalance} 小時`);
        
        // 檢查餘額是否足夠
        if (currentBalance < hours) {
          Logger.log(`   ⚠️ 餘額不足：需要 ${hours} 小時，只剩 ${currentBalance} 小時`);
          return {
            ok: false,
            msg: `${leaveType} 餘額不足（需要 ${hours} 小時，只剩 ${currentBalance} 小時）`
          };
        }
        
        // 扣除餘額
        const newBalance = currentBalance - hours;
        
        Logger.log(`   扣除 ${hours} 小時後: ${newBalance} 小時`);
        
        sheet.getRange(i + 1, columnIndex).setValue(newBalance);
        sheet.getRange(i + 1, 19).setValue(new Date()); // S 欄: 更新時間（索引 +1）⭐
        
        Logger.log('✅ 餘額已更新');
        
        return {
          ok: true,
          remaining: newBalance
        };
      }
    }
    
    Logger.log('❌ 找不到員工記錄');
    return {
      ok: false,
      msg: "找不到員工記錄"
    };
    
  } catch (error) {
    Logger.log('❌ deductLeaveBalance 錯誤: ' + error);
    return {
      ok: false,
      msg: error.message
    };
  }
}

function updateAllEmployeesAnnualLeave() {
  Logger.log('═══════════════════════════════════════');
  Logger.log('🔄 開始批次更新特休餘額');
  Logger.log('═══════════════════════════════════════');

  const balanceSheet = getLeaveBalanceSheet();
  const balanceValues = balanceSheet.getDataRange().getValues();
  const today = new Date();
  const todayStr = Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyy-MM-dd');

  const todayMonth = today.getMonth();
  const todayDate = today.getDate();

  let updateCount = 0;

  for (let i = 1; i < balanceValues.length; i++) {
    const userId = balanceValues[i][0];
    const employeeName = balanceValues[i][1];
    const hireDate = balanceValues[i][2];

    if (!hireDate) {
      Logger.log(`⚠️ ${employeeName} 沒有到職日，跳過`);
      continue;
    }

    const hire = new Date(hireDate);
    const yearsOfService = calculateYearsOfService(hire, today);
    const currentAnnualLeave = parseFloat(balanceValues[i][3]) || 0;

    // ⭐ 補漏機制：特休是 0 但年資已超過 6 個月
    if (currentAnnualLeave === 0 && yearsOfService >= 0.5) {
      const leaveInfo = getCurrentAnnualLeaveInfo(hire, today);
      const correctHours = leaveInfo.currentHours;

      if (correctHours > 0) {
        balanceSheet.getRange(i + 1, 4).setValue(correctHours);
        balanceSheet.getRange(i + 1, 19).setValue(new Date());
        Logger.log(`🔧 補漏 ${employeeName}: 0 → ${correctHours} 小時`);
        updateCount++;
        continue; // 補漏完就跳到下一個員工
      }
    }

    // ⭐ 今天已更新過則跳過（防止每小時重複觸發）
    const lastUpdateDate = balanceValues[i][18];
    if (lastUpdateDate) {
      const lastUpdateStr = Utilities.formatDate(
        new Date(lastUpdateDate),
        Session.getScriptTimeZone(),
        'yyyy-MM-dd'
      );
      if (lastUpdateStr === todayStr) {
        Logger.log(`⏭️ ${employeeName}: 今天已更新過，跳過`);
        continue;
      }
    }

    // 判斷今天是否為週年日或半年日
    const halfYearDate = new Date(hire);
    halfYearDate.setMonth(halfYearDate.getMonth() + 6);

    const isAnniversary = (todayMonth === hire.getMonth() && todayDate === hire.getDate());
    const isHalfYear = (todayMonth === halfYearDate.getMonth() && todayDate === halfYearDate.getDate());
    const isHalfYearTrigger = isHalfYear && yearsOfService < 1;
    const isAnniversaryTrigger = isAnniversary && yearsOfService >= 1;

    if (!isHalfYearTrigger && !isAnniversaryTrigger) {
      Logger.log(`📊 ${employeeName}: 今天非觸發日，跳過`);
      continue;
    }

    const leaveInfo = getCurrentAnnualLeaveInfo(hire, today);
    const newAnnualHours = leaveInfo.currentHours;

    Logger.log(`🎂 ${employeeName} 週年更新: ${newAnnualHours} 小時`);

    if (isHalfYearTrigger) {
      balanceSheet.getRange(i + 1, 4).setValue(24);
      Logger.log(`   ✅ 滿半年，設定 24 小時`);
    } else {
      const oldRemaining = balanceValues[i][3];
      balanceSheet.getRange(i + 1, 4).setValue(oldRemaining + newAnnualHours);
      Logger.log(`   ✅ 週年日，累加後共 ${oldRemaining + newAnnualHours} 小時`);
    }

    balanceSheet.getRange(i + 1, 19).setValue(new Date());
    updateCount++;
  }

  Logger.log(`\n✅ 完成，共更新 ${updateCount} 筆`);
  return { ok: true, updateCount: updateCount };
}
/**
 * 🔄 遷移工具：為現有員工加入到職日欄位（18欄 → 19欄）
 * 
 * 使用方式：
 * 1. 在 Apps Script 編輯器中執行此函數
 * 2. 會自動備份舊資料
 * 3. 重建新結構並遷移資料
 */
function migrateAddHireDateColumn() {
  Logger.log('═══════════════════════════════════════');
  Logger.log('🔄 開始遷移：新增到職日欄位');
  Logger.log('═══════════════════════════════════════');
  Logger.log('');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const oldSheet = ss.getSheetByName('假期餘額');
  
  if (!oldSheet) {
    Logger.log('❌ 找不到「假期餘額」工作表');
    return;
  }
  
  // 📋 步驟 1：備份舊工作表
  Logger.log('📋 步驟 1：備份舊工作表...');
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
  const backupSheet = oldSheet.copyTo(ss);
  backupSheet.setName('假期餘額_備份_' + timestamp);
  Logger.log('✅ 已備份: ' + backupSheet.getName());
  Logger.log('');
  
  // 📋 步驟 2：讀取舊資料
  Logger.log('📋 步驟 2：讀取舊資料...');
  const oldData = oldSheet.getDataRange().getValues();
  const recordCount = oldData.length - 1;
  
  Logger.log(`   找到 ${recordCount} 筆員工資料`);
  Logger.log('');
  
  if (recordCount <= 0) {
    Logger.log('⚠️ 沒有資料需要遷移');
    return;
  }
  
  // 📋 步驟 3：刪除舊工作表
  Logger.log('📋 步驟 3：刪除舊工作表...');
  ss.deleteSheet(oldSheet);
  Logger.log('✅ 已刪除舊的「假期餘額」工作表');
  Logger.log('');
  
  // 📋 步驟 4：建立新工作表（19 個欄位）
  Logger.log('📋 步驟 4：建立新工作表（19 個欄位）...');
  const newSheet = ss.insertSheet('假期餘額');
  
  // 建立標題列
  const headers = [
    '員工ID',           // A
    '姓名',             // B
    '到職日',           // C ⭐ 新增
    '特休假',           // D
    '未住院病假',       // E
    '事假',             // F
    '喪假',             // G
    '婚假',             // H
    '產假',             // I
    '陪產檢及陪產假',   // J
    '住院病假',         // K
    '生理假',           // L
    '家庭照顧假',       // M
    '公假(含兵役假)',   // N
    '公傷假',           // O
    '天然災害停班',     // P
    '加班補休假',       // Q
    '曠工',             // R
    '更新時間'          // S
  ];
  
  newSheet.appendRow(headers);
  
  // 美化標題列
  const headerRange = newSheet.getRange(1, 1, 1, 19);
  headerRange.setBackground('#4A90E2');
  headerRange.setFontColor('#FFFFFF');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
  
  // 凍結標題列
  newSheet.setFrozenRows(1);
  
  Logger.log('✅ 新工作表已建立（19 個欄位）');
  Logger.log('');
  
  // 📋 步驟 5：遷移資料
  Logger.log('📋 步驟 5：遷移資料（新增到職日，預設為今天）...');
  Logger.log('');
  
  const today = new Date();
  
  for (let i = 1; i < oldData.length; i++) {
    const oldRow = oldData[i];
    const userId = oldRow[0];
    const employeeName = oldRow[1];
    
    // 對應關係：
    // 舊: [員工ID, 姓名, 特休假, 未住院病假, 事假, 喪假, 婚假, 產假, 陪產檢及陪產假, 住院病假, 生理假, 家庭照顧假, 公假, 公傷假, 天然災害停班, 加班補休假, 曠工, 更新時間]
    // 新: [員工ID, 姓名, 到職日⭐, 特休假, 未住院病假, 事假, 喪假, 婚假, 產假, 陪產檢及陪產假, 住院病假, 生理假, 家庭照顧假, 公假, 公傷假, 天然災害停班, 加班補休假, 曠工, 更新時間]
    
    const newRow = [
      oldRow[0] || '',      // A: 員工ID（保留）
      oldRow[1] || '',      // B: 姓名（保留）
      today,                // C: 到職日 ⭐ 新增（預設為今天）
      oldRow[2] || 0,       // D: 特休假（保留）
      oldRow[3] || 240,     // E: 未住院病假（保留）
      oldRow[4] || 112,     // F: 事假（保留）
      oldRow[5] || 40,      // G: 喪假（保留）
      oldRow[6] || 64,      // H: 婚假（保留）
      oldRow[7] || 448,     // I: 產假（保留）
      oldRow[8] || 56,      // J: 陪產檢及陪產假（保留）
      oldRow[9] || 240,     // K: 住院病假（保留）
      oldRow[10] || 96,     // L: 生理假（保留）
      oldRow[11] || 56,     // M: 家庭照顧假（保留）
      oldRow[12] || 0,      // N: 公假（保留）
      oldRow[13] || 0,      // O: 公傷假（保留）
      oldRow[14] || 0,      // P: 天然災害停班（保留）
      oldRow[15] || 0,      // Q: 加班補休假（保留）
      oldRow[16] || 0,      // R: 曠工（保留）
      new Date()            // S: 更新時間（更新為當前時間）
    ];
    
    newSheet.appendRow(newRow);
    
    Logger.log(`   ✅ [${i}/${recordCount}] 已遷移: ${userId} - ${employeeName}`);
  }
  
  Logger.log('');
  Logger.log('═══════════════════════════════════════');
  Logger.log('✅✅✅ 遷移完成！');
  Logger.log('═══════════════════════════════════════');
  Logger.log('');
  Logger.log('📊 遷移摘要:');
  Logger.log(`   - 舊結構: 18 個欄位`);
  Logger.log(`   - 新結構: 19 個欄位`);
  Logger.log(`   - 遷移記錄數: ${recordCount} 筆`);
  Logger.log(`   - 備份工作表: ${backupSheet.getName()}`);
  Logger.log('');
  Logger.log('📝 新增欄位:');
  Logger.log('   - C 欄: 到職日（預設為今天）');
  Logger.log('');
  Logger.log('⚠️ 注意：');
  Logger.log('   現有員工的到職日預設為今天');
  Logger.log('   如需修正，請手動編輯 Google Sheet');
  Logger.log('');
  
  // 顯示成功訊息給使用者
  Browser.msgBox(
    '✅ 遷移完成！',
    '已成功將 ' + recordCount + ' 筆假期餘額加上到職日欄位！\n\n' +
    '舊結構：18 個欄位\n' +
    '新結構：19 個欄位（新增到職日欄位）\n\n' +
    '⚠️ 注意：現有員工的到職日預設為今天\n' +
    '如需修正，請手動編輯 Google Sheet 的「到職日」欄位\n\n' +
    '備份工作表: ' + backupSheet.getName() + '\n\n' +
    '請檢查「假期餘額」工作表確認資料正確。',
    Browser.Buttons.OK
  );
}

/**
 * 🧪 測試扣除餘額功能
 */
function testDeductLeaveBalance() {
  Logger.log('🧪 測試扣除假期餘額');
  Logger.log('');
  
  // ⚠️ 請替換成實際的員工ID
  const testUserId = 'U7854bd6965d1c25b1c79d00c1dce001b'; // 從 LINE 取得的 userId
  
  Logger.log('📋 測試參數:');
  Logger.log(`   員工ID: ${testUserId}`);
  Logger.log(`   假別: ANNUAL_LEAVE (特休假)`);
  Logger.log(`   天數: 0.25 (2 小時)`);
  Logger.log('');
  
  const result = deductLeaveBalance(testUserId, 'ANNUAL_LEAVE', 0.25);
  
  Logger.log('📤 測試結果:');
  Logger.log(JSON.stringify(result, null, 2));
  
  if (result.ok) {
    Logger.log('');
    Logger.log('✅ 測試成功！');
    Logger.log(`   剩餘餘額: ${result.remaining} 天`);
  } else {
    Logger.log('');
    Logger.log('❌ 測試失敗');
  }
}

/**
 * 🧪 完整測試：提交 → 審核 → 扣除餘額
 */
function testCompleteLeaveFlow() {
  Logger.log('🧪 測試完整請假流程');
  Logger.log('');
  
  const token = '7dac1161-bbac-487d-900b-3e06c1acab8d'; // ⚠️ 替換成有效 token
  
  // 步驟 1：提交請假
  Logger.log('📋 步驟 1：提交請假申請');
  const submitResult = submitLeaveRequest(
    token,
    'ANNUAL_LEAVE',
    '2025-12-19T09:00',
    '2025-12-19T11:00',
    '測試完整流程'
  );
  
  Logger.log('   結果: ' + JSON.stringify(submitResult));
  
  if (!submitResult.ok) {
    Logger.log('❌ 提交失敗，測試終止');
    return;
  }
  
  Logger.log('');
  
  // 步驟 2：查詢餘額（扣除前）
  Logger.log('📋 步驟 2：查詢餘額（扣除前）');
  const balanceBefore = getLeaveBalance(token);
  Logger.log('   特休假餘額: ' + balanceBefore.balance.ANNUAL_LEAVE + ' 天');
  Logger.log('');
  
  // 步驟 3：審核請假（需要手動指定 rowNumber）
  Logger.log('📋 步驟 3：審核請假申請');
  Logger.log('   ⚠️ 請手動查看「請假紀錄」工作表的最後一行行號');
  Logger.log('   然後修改下面的 rowNumber');
  
  const rowNumber = 2; // ⚠️ 替換成實際行號
  
  const reviewResult = reviewLeaveRequest(token, rowNumber, 'approve', '核准測試');
  Logger.log('   結果: ' + JSON.stringify(reviewResult));
  Logger.log('');
  
  // 步驟 4：查詢餘額（扣除後）
  Logger.log('📋 步驟 4：查詢餘額（扣除後）');
  const balanceAfter = getLeaveBalance(token);
  Logger.log('   特休假餘額: ' + balanceAfter.balance.ANNUAL_LEAVE + ' 天');
  Logger.log('');
  
  // 比較
  Logger.log('📊 比較結果:');
  Logger.log(`   扣除前: ${balanceBefore.balance.ANNUAL_LEAVE} 天`);
  Logger.log(`   扣除後: ${balanceAfter.balance.ANNUAL_LEAVE} 天`);
  Logger.log(`   差異: ${balanceBefore.balance.ANNUAL_LEAVE - balanceAfter.balance.ANNUAL_LEAVE} 天`);
  Logger.log('');
  
  if (balanceBefore.balance.ANNUAL_LEAVE > balanceAfter.balance.ANNUAL_LEAVE) {
    Logger.log('✅✅✅ 測試成功！餘額已正確扣除');
  } else {
    Logger.log('❌ 測試失敗：餘額未扣除');
  }
}


/**
 * 🔄 遷移工具：將現有假期餘額加上姓名欄位（17欄 → 18欄）
 * 
 * 使用方式：
 * 1. 在 Apps Script 編輯器中執行此函數
 * 2. 會自動備份舊資料
 * 3. 重建新結構並遷移資料
 */
function migrateAddNameColumn() {
  Logger.log('═══════════════════════════════════════');
  Logger.log('🔄 開始遷移：新增姓名欄位');
  Logger.log('═══════════════════════════════════════');
  Logger.log('');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const oldSheet = ss.getSheetByName('假期餘額');
  
  if (!oldSheet) {
    Logger.log('❌ 找不到「假期餘額」工作表');
    return;
  }
  
  // 📋 步驟 1：備份舊工作表
  Logger.log('📋 步驟 1：備份舊工作表...');
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
  const backupSheet = oldSheet.copyTo(ss);
  backupSheet.setName('假期餘額_備份_' + timestamp);
  Logger.log('✅ 已備份: ' + backupSheet.getName());
  Logger.log('');
  
  // 📋 步驟 2：讀取舊資料
  Logger.log('📋 步驟 2：讀取舊資料...');
  const oldData = oldSheet.getDataRange().getValues();
  const recordCount = oldData.length - 1; // 扣除標題列
  
  Logger.log(`   找到 ${recordCount} 筆員工資料`);
  Logger.log('');
  
  if (recordCount <= 0) {
    Logger.log('⚠️ 沒有資料需要遷移');
    return;
  }
  
  // 📋 步驟 3：取得員工資料工作表（用於查找姓名）
  Logger.log('📋 步驟 3：準備查找員工姓名...');
  const employeeSheet = ss.getSheetByName('員工資料');
  let employeeMap = {};
  
  if (employeeSheet) {
    const empData = employeeSheet.getDataRange().getValues();
    for (let i = 1; i < empData.length; i++) {
      const userId = empData[i][0];  // A: 員工ID
      const name = empData[i][1];    // B: 姓名
      if (userId && name) {
        employeeMap[userId] = name;
      }
    }
    Logger.log(`   已載入 ${Object.keys(employeeMap).length} 筆員工姓名對照`);
  } else {
    Logger.log('   ⚠️ 找不到「員工資料」工作表，將使用預設姓名');
  }
  Logger.log('');
  
  // 📋 步驟 4：刪除舊工作表
  Logger.log('📋 步驟 4：刪除舊工作表...');
  ss.deleteSheet(oldSheet);
  Logger.log('✅ 已刪除舊的「假期餘額」工作表');
  Logger.log('');
  
  // 📋 步驟 5：建立新工作表（18 個欄位）
  Logger.log('📋 步驟 5：建立新工作表（18 個欄位）...');
  const newSheet = ss.insertSheet('假期餘額');
  
  // 建立標題列
  const headers = [
    '員工ID',           // A
    '姓名',             // B ⭐ 新增
    '特休假',           // C
    '未住院病假',       // D
    '事假',             // E
    '喪假',             // F
    '婚假',             // G
    '產假',             // H
    '陪產檢及陪產假',   // I
    '住院病假',         // J
    '生理假',           // K
    '家庭照顧假',       // L
    '公假(含兵役假)',   // M
    '公傷假',           // N
    '天然災害停班',     // O
    '加班補休假',       // P
    '曠工',             // Q
    '更新時間'          // R
  ];
  
  newSheet.appendRow(headers);
  
  // 美化標題列
  const headerRange = newSheet.getRange(1, 1, 1, 18);
  headerRange.setBackground('#4A90E2');
  headerRange.setFontColor('#FFFFFF');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
  
  // 凍結標題列
  newSheet.setFrozenRows(1);
  
  Logger.log('✅ 新工作表已建立（18 個欄位）');
  Logger.log('');
  
  // 📋 步驟 6：遷移資料
  Logger.log('📋 步驟 6：遷移資料（新增姓名）...');
  Logger.log('');
  
  for (let i = 1; i < oldData.length; i++) {
    const oldRow = oldData[i];
    const userId = oldRow[0];
    
    // 從員工資料工作表查找姓名，如果找不到則使用預設值
    const employeeName = employeeMap[userId] || `員工${i}`;
    
    // 對應關係：
    // 舊: [員工ID, 特休假, 未住院病假, 事假, 喪假, 婚假, 產假, 陪產檢及陪產假, 住院病假, 生理假, 家庭照顧假, 公假, 公傷假, 天然災害停班, 加班補休假, 曠工, 更新時間]
    // 新: [員工ID, 姓名⭐, 特休假, 未住院病假, 事假, 喪假, 婚假, 產假, 陪產檢及陪產假, 住院病假, 生理假, 家庭照顧假, 公假, 公傷假, 天然災害停班, 加班補休假, 曠工, 更新時間]
    
    const newRow = [
      oldRow[0] || '',      // A: 員工ID（保留）
      employeeName,         // B: 姓名 ⭐ 新增
      oldRow[1] || 7,       // C: 特休假（保留）
      oldRow[2] || 30,      // D: 未住院病假（保留）
      oldRow[3] || 14,      // E: 事假（保留）
      oldRow[4] || 5,       // F: 喪假（保留）
      oldRow[5] || 8,       // G: 婚假（保留）
      oldRow[6] || 56,      // H: 產假（保留）
      oldRow[7] || 7,       // I: 陪產檢及陪產假（保留）
      oldRow[8] || 30,      // J: 住院病假（保留）
      oldRow[9] || 12,      // K: 生理假（保留）
      oldRow[10] || 7,      // L: 家庭照顧假（保留）
      oldRow[11] || 0,      // M: 公假（保留）
      oldRow[12] || 0,      // N: 公傷假（保留）
      oldRow[13] || 0,      // O: 天然災害停班（保留）
      oldRow[14] || 0,      // P: 加班補休假（保留）
      oldRow[15] || 0,      // Q: 曠工（保留）
      new Date()            // R: 更新時間（更新為當前時間）
    ];
    
    newSheet.appendRow(newRow);
    
    Logger.log(`   ✅ [${i}/${recordCount}] 已遷移: ${userId} - ${employeeName}`);
  }
  
  Logger.log('');
  Logger.log('═══════════════════════════════════════');
  Logger.log('✅✅✅ 遷移完成！');
  Logger.log('═══════════════════════════════════════');
  Logger.log('');
  Logger.log('📊 遷移摘要:');
  Logger.log(`   - 舊結構: 17 個欄位`);
  Logger.log(`   - 新結構: 18 個欄位`);
  Logger.log(`   - 遷移記錄數: ${recordCount} 筆`);
  Logger.log(`   - 備份工作表: ${backupSheet.getName()}`);
  Logger.log('');
  Logger.log('📝 新增欄位:');
  Logger.log('   - B 欄: 姓名');
  Logger.log('');
  
  // 顯示成功訊息給使用者
  Browser.msgBox(
    '✅ 遷移完成！',
    '已成功將 ' + recordCount + ' 筆假期餘額加上姓名欄位！\n\n' +
    '舊結構：17 個欄位\n' +
    '新結構：18 個欄位（新增姓名欄位）\n\n' +
    '備份工作表: ' + backupSheet.getName() + '\n\n' +
    '請檢查「假期餘額」工作表確認資料正確。',
    Browser.Buttons.OK
  );
}

function checkLeaveRecordInSheet() {
  Logger.log('🔍 檢查 Sheet 中的請假記錄');
  Logger.log('');
  
  const sheet = getLeaveRecordsSheet();
  const values = sheet.getDataRange().getValues();
  
  Logger.log('📊 所有請假記錄:');
  Logger.log('');
  
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    
    Logger.log(`═══ 第 ${i + 1} 行 ═══`);
    Logger.log(`   員工ID: ${row[1]}`);
    Logger.log(`   姓名: ${row[2]}`);
    Logger.log(`   假別: ${row[4]}`);
    Logger.log(`   開始時間: ${row[5]}`);
    Logger.log(`   結束時間: ${row[6]}`);
    Logger.log(`   工作時數 (H欄): ${row[7]}`);  // ⭐ 關鍵
    Logger.log(`   天數 (I欄): ${row[8]}`);      // ⭐ 關鍵
    Logger.log(`   狀態: ${row[10]}`);
    Logger.log('');
  }
}


/**
 * 🔄 遷移腳本：將現有的天數資料轉換為小時數
 */
function migrateLeaveBalanceToHours() {
  Logger.log('═══════════════════════════════════════');
  Logger.log('🔄 開始遷移假期餘額：天數 → 小時');
  Logger.log('═══════════════════════════════════════');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('假期餘額');
  
  if (!sheet) {
    Logger.log('❌ 找不到「假期餘額」工作表');
    return;
  }
  
  const data = sheet.getDataRange().getValues();
  
  Logger.log(`📊 找到 ${data.length - 1} 筆員工資料`);
  Logger.log('');
  
  // 從第 2 行開始（跳過標題）
  for (let i = 1; i < data.length; i++) {
    const employeeId = data[i][0];
    
    Logger.log(`🔄 處理員工: ${employeeId}`);
    
    // B-P 欄（索引 1-15）：將天數 × 8 轉換為小時
    for (let col = 1; col <= 15; col++) {
      const days = parseFloat(data[i][col]) || 0;
      const hours = days * 8;
      
      sheet.getRange(i + 1, col + 1).setValue(hours);
      
      if (col === 1) {  // 只記錄第一個欄位（特休假）
        Logger.log(`   特休假: ${days} 天 → ${hours} 小時`);
      }
    }
    
    // 更新時間
    sheet.getRange(i + 1, 17).setValue(new Date());
    
    Logger.log(`   ✅ 完成`);
  }
  
  Logger.log('');
  Logger.log('═══════════════════════════════════════');
  Logger.log('✅ 遷移完成！');
  Logger.log('═══════════════════════════════════════');
  
  Browser.msgBox(
    '✅ 遷移完成！',
    `已成功將 ${data.length - 1} 筆員工的假期餘額從天數轉換為小時數！\n\n` +
    '請重新載入前端頁面以查看結果。',
    Browser.Buttons.OK
  );
}

function checkLeaveBalanceStructure() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('假期餘額');
  
  if (!sheet) {
    Logger.log('❌ 找不到工作表');
    return;
  }
  
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  Logger.log('📊 假期餘額工作表結構:');
  Logger.log(`   總欄位數: ${headers.length}`);
  Logger.log('');
  Logger.log('欄位列表:');
  
  headers.forEach((header, index) => {
    const columnLetter = String.fromCharCode(65 + index); // A, B, C...
    Logger.log(`   ${columnLetter} 欄: ${header}`);
  });
}

/**
 * 🔍 檢查現有假期餘額結構
 */
function checkCurrentStructure() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('假期餘額');
  
  if (!sheet) {
    Logger.log('❌ 找不到「假期餘額」工作表');
    return;
  }
  
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  Logger.log('📊 假期餘額工作表結構:');
  Logger.log(`   總欄位數: ${headers.length}`);
  Logger.log('');
  Logger.log('欄位列表:');
  
  headers.forEach((header, index) => {
    const columnLetter = String.fromCharCode(65 + index);
    Logger.log(`   ${columnLetter} 欄 (索引 ${index}): ${header}`);
  });
  
  Logger.log('');
  
  // 判斷需要執行哪個遷移
  if (headers.length === 18 && headers[2] !== '到職日') {
    Logger.log('✅ 需要執行：migrateAddHireDateColumn()');
    Logger.log('   （新增到職日欄位）');
  } else if (headers.length === 19 && headers[2] === '到職日') {
    Logger.log('✅ 結構正確！無需遷移');
  } else {
    Logger.log('⚠️ 結構異常，請檢查');
  }
}



/**
 * ✅ 取得假期餘額（含即時年資顯示）
 * 在 getLeaveBalance() 中可選擇性附加特休資訊
 */
function getAnnualLeaveStatus(sessionToken) {
  try {
    const employee = checkSession_(sessionToken);
    if (!employee.ok || !employee.user) return { ok: false };

    const balanceResult = getLeaveBalance(sessionToken);
    if (!balanceResult.ok) return balanceResult;

    const sheet = getLeaveBalanceSheet();
    const values = sheet.getDataRange().getValues();

    for (let i = 1; i < values.length; i++) {
      if (values[i][0] === employee.user.userId) {
        const hireDate = values[i][2];
        const leaveInfo = getCurrentAnnualLeaveInfo(hireDate);

        return {
          ok: true,
          balance: balanceResult.balance,
          annualLeaveStatus: {
            yearsOfService: leaveInfo.yearsOfService.toFixed(2),
            currentPeriodHours: leaveInfo.currentHours,
            currentPeriodDays: leaveInfo.currentHours / 8,
            periodStart: leaveInfo.periodStart ? formatDate(leaveInfo.periodStart) : null,
            periodEnd: leaveInfo.periodEnd ? formatDate(leaveInfo.periodEnd) : null,
            remainingHours: balanceResult.balance.ANNUAL_LEAVE
          }
        };
      }
    }

    return { ok: false, msg: '找不到員工資料' };
  } catch (e) {
    return { ok: false, msg: e.message };
  }
}


/**
 * 🧪 測試特休即時計算
 */
function testAnnualLeaveCalculation() {
  Logger.log('🧪 測試特休即時計算');
  Logger.log('');

  const testCases = [
    { label: '未滿 6 個月', months: 3 },
    { label: '滿 6 個月',   months: 6 },
    { label: '滿 1 年',     months: 12 },
    { label: '滿 2 年',     months: 24 },
    { label: '滿 3 年',     months: 36 },
    { label: '滿 5 年',     months: 60 },
    { label: '滿 10 年',    months: 120 },
    { label: '滿 15 年',    months: 180 },
    { label: '滿 25 年',    months: 300 },
  ];

  const today = new Date();

  testCases.forEach(tc => {
    const hireDate = new Date(today);
    hireDate.setMonth(hireDate.getMonth() - tc.months);

    const info = getCurrentAnnualLeaveInfo(hireDate, today);

    Logger.log(`📋 ${tc.label}:`);
    Logger.log(`   到職日: ${formatDate(hireDate)}`);
    Logger.log(`   年資: ${info.yearsOfService.toFixed(2)} 年`);
    Logger.log(`   特休: ${info.currentHours} 小時 (${info.currentHours / 8} 天)`);
    if (info.periodStart) {
      Logger.log(`   有效期: ${formatDate(info.periodStart)} ~ ${formatDate(info.periodEnd)}`);
    }
    Logger.log('');
  });
}


function fixAnnualLeaveByHireDate() {
  const balanceSheet = getLeaveBalanceSheet();
  const values = balanceSheet.getDataRange().getValues();
  const today = new Date();

  for (let i = 1; i < values.length; i++) {
    const hireDate = values[i][2]; // C 欄：到職日
    if (!hireDate) continue;

    const leaveInfo = getCurrentAnnualLeaveInfo(new Date(hireDate), today);
    const correctHours = leaveInfo.currentHours;

    // 只有當特休是 0 且年資已超過 6 個月時才修正
    if (values[i][3] === 0 && correctHours > 0) {
      balanceSheet.getRange(i + 1, 4).setValue(correctHours);
      balanceSheet.getRange(i + 1, 19).setValue(new Date());
      Logger.log(`✅ 修正 ${values[i][1]}: 0 → ${correctHours} 小時`);
    }
  }
  Logger.log('完成修正');
}


function debugLeaveBalance() {
  const token = 'f41e2a20-c165-4469-8e89-bea43a7f7baa';
  
  // 檢查工作表結構
  checkCurrentStructure();
  
  // 直接查詢
  const result = getLeaveBalance(token);
  Logger.log('結果: ' + JSON.stringify(result));
}