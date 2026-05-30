// OvertimeOperations.gs - 加班功能後端（完全修正版）

// ==================== 常數定義 ====================
const SHEET_OVERTIME = "加班申請";

// ==================== 資料庫操作 ====================

/**
 * 初始化加班申請工作表（修改版 - 加入補休時數欄位）
 */
function initOvertimeSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_OVERTIME);
  
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_OVERTIME);
    // ✅ 加入「補休時數」欄位
    const headers = [
      "申請ID", "員工ID", "員工姓名", "加班日期", 
      "開始時間", "結束時間", "加班時數", "申請原因",
      "申請時間", "審核狀態", "審核人ID", "審核人姓名",
      "審核時間", "審核意見", "補休時數"
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
    Logger.log("✅ 加班申請工作表已建立");
  }
  
  return sheet;
}

/**
 * 提交加班申請（修改版 - 加入補休時數）
 */
function submitOvertimeRequest(sessionToken, overtimeDate, startTime, endTime, hours, reason, compensatoryHours) {
  const employee = checkSession_(sessionToken);
  const user = employee.user;
  if (!user) return { ok: false, code: "ERR_SESSION_INVALID" };
  
  const sheet = initOvertimeSheet();
  const requestId = "OT" + new Date().getTime();
  
  // 組合完整的日期時間格式
  const startDateTime = new Date(`${overtimeDate}T${startTime}:00`);
  const endDateTime = new Date(`${overtimeDate}T${endTime}:00`);
  
  // ✅ 處理補休時數（預設為 0）
  const compHours = parseFloat(compensatoryHours) || 0;
  
  Logger.log(`📝 提交加班: ${user.name}, 日期=${overtimeDate}, 時數=${hours}, 補休=${compHours}`);
  
  const row = [
    requestId,
    user.userId,
    user.name,
    overtimeDate,
    startDateTime,
    endDateTime,
    parseFloat(hours),
    reason,
    new Date(),
    "pending",
    "", "", "", "",
    compHours  // ✅ 補休時數
  ];
  
  sheet.appendRow(row);
  
  return { 
    ok: true, 
    code: "OVERTIME_SUBMIT_SUCCESS",
    requestId: requestId
  };
}


/**
 * Handler - 接收補休時數參數
 */
function handleSubmitOvertime(params) {
  const { token, overtimeDate, startTime, endTime, hours, reason, compensatoryHours } = params;
  
  Logger.log(`📥 收到加班申請: 日期=${overtimeDate}, 時數=${hours}, 補休=${compensatoryHours || 0}`);
  
  return submitOvertimeRequest(
    token, 
    overtimeDate, 
    startTime, 
    endTime, 
    parseFloat(hours), 
    reason,
    parseFloat(compensatoryHours) || 0  // ✅ 補休時數，預設 0
  );
}

/**
 * 查詢員工的加班申請記錄
 */
function getEmployeeOvertimeRequests(sessionToken) {
  const employee = checkSession_(sessionToken);
  const user = employee.user;
  if (!user) return { ok: false, code: "ERR_SESSION_INVALID" };
  
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_OVERTIME);
  if (!sheet) return { ok: true, requests: [] };
  
  const values = sheet.getDataRange().getValues();
  
  const formatTime = (dateTime) => {
    if (!dateTime) return "";
    if (typeof dateTime === "string") {
      if (dateTime.includes(':')) return dateTime.substring(0, 5);
      return dateTime;
    }
    return Utilities.formatDate(dateTime, "Asia/Taipei", "HH:mm");
  };
  
  const requests = values.slice(1).filter(row => {
    return row[1] === user.userId;
  }).map(row => {
    return {
      requestId: row[0],
      overtimeDate: formatDate(row[3]),
      startTime: formatTime(row[4]),
      endTime: formatTime(row[5]),
      hours: parseFloat(row[6]) || 0,
      reason: row[7],
      applyDate: formatDate(row[8]),
      status: String(row[9]).trim().toLowerCase(),
      reviewerName: row[11] || "",
      reviewComment: row[13] || "",
      compensatoryHours: parseFloat(row[14]) || 0  // ✅ 加入這行
    };
  });
  
  Logger.log(`👤 員工 ${user.name} 的加班記錄: ${requests.length} 筆`);
  return { ok: true, requests: requests };
}

/**
 * 取得所有待審核的加班申請（管理員用）
 */
function getPendingOvertimeRequests(sessionToken) {
  const employee = checkSession_(sessionToken);
  const user = employee.user;
  if (!user) return { ok: false, code: "ERR_SESSION_INVALID" };
  
  if (user.dept !== "管理員") {
    return { ok: false, code: "ERR_NO_PERMISSION" };
  }
  
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_OVERTIME);
  if (!sheet) return { ok: true, requests: [] };
  
  const values = sheet.getDataRange().getValues();
  
  const formatTime = (dateTime) => {
    if (!dateTime) return "";
    if (typeof dateTime === "string") {
      if (dateTime.includes(':')) return dateTime.substring(0, 5);
      return dateTime;
    }
    return Utilities.formatDate(dateTime, "Asia/Taipei", "HH:mm");
  };
  
  const requests = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const status = String(row[9]).trim().toLowerCase();
    
    if (status === "pending") {
      requests.push({
      rowNumber: i + 1,
      requestId: row[0],
      employeeId: row[1],
      employeeName: row[2],
      overtimeDate: formatDate(row[3]),
      startTime: formatTime(row[4]),
      endTime: formatTime(row[5]),
      hours: parseFloat(row[6]) || 0,
      reason: row[7],
      applyDate: formatDate(row[8]),
      compensatoryHours: parseFloat(row[14]) || 0  // ✅ 加入這行
    });
    }
  }
  
  Logger.log(`📊 共 ${requests.length} 筆待審核加班申請`);
  return { ok: true, requests: requests };
}

// OvertimeOperations.gs - 加班審核功能（含 LINE 通知）

/**
 * 審核加班申請（完整版 - 含 LINE 通知）
 * @param {string} sessionToken - Session Token
 * @param {number} rowNumber - 試算表行號
 * @param {string} action - 審核動作 (approve/reject)
 * @param {string} comment - 審核意見
 */
function reviewOvertimeRequest(sessionToken, rowNumber, action, comment) {
  const employee = checkSession_(sessionToken);
  const user = employee.user;
  if (!user) return { ok: false, code: "ERR_SESSION_INVALID" };
  
  if (user.dept !== "管理員") {
    return { ok: false, code: "ERR_NO_PERMISSION" };
  }
  
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_OVERTIME);
  if (!sheet) return { ok: false, msg: "找不到加班申請工作表" };
  
  // 🔧 關鍵修正：嚴格處理 action 參數
  const actionStr = String(action).trim().toLowerCase();
  const isApprove = (actionStr === "approve");
  const status = isApprove ? "approved" : "rejected";
  const reviewTime = new Date();
  
  Logger.log(`📥 審核請求: rowNumber=${rowNumber}, action="${action}", 處理後="${actionStr}", isApprove=${isApprove}, 目標狀態="${status}"`);
  
  try {
    // 👉 取得加班申請的完整資訊（用於通知）
    const record = sheet.getRange(rowNumber, 1, 1, 14).getValues()[0];
    const requestId = record[0];       // 申請ID
    const employeeId = record[1];      // 員工ID
    const employeeName = record[2];    // 員工姓名
    const overtimeDate = record[3];    // 加班日期
    const startTime = record[4];       // 開始時間
    const endTime = record[5];         // 結束時間
    const hours = record[6];           // 加班時數
    const reason = record[7];          // 申請原因
    
    Logger.log(`📋 審核對象: ${employeeName}, 日期: ${formatDate(overtimeDate)}, 時數: ${hours}`);
    
    // 更新審核資訊
    sheet.getRange(rowNumber, 10).setValue(status);           // 審核狀態
    sheet.getRange(rowNumber, 11).setValue(user.userId);      // 審核人ID
    sheet.getRange(rowNumber, 12).setValue(user.name);        // 審核人姓名
    sheet.getRange(rowNumber, 13).setValue(reviewTime);       // 審核時間
    sheet.getRange(rowNumber, 14).setValue(comment || "");    // 審核意見
    
    SpreadsheetApp.flush();
    // 在 reviewOvertimeRequest 裡，審核成功寫入後加入這段
    if (isApprove) {
      // ⭐ 審核通過後自動重算薪資
      try {
        const overtimeDateStr = formatDate(overtimeDate);
        const yearMonth = overtimeDateStr.substring(0, 7); // 取 YYYY-MM
        const recalcResult = _recalcOne(employeeId, yearMonth);
        if (recalcResult.success) {
          Logger.log(`✅ 已自動重算 ${employeeId} ${yearMonth} 薪資`);
        } else {
          Logger.log(`⚠️ 自動重算失敗: ${recalcResult.message}`);
        }
      } catch (err) {
        Logger.log(`⚠️ 自動重算發生錯誤: ${err.message}`);
      }
    }
    // 驗證寫入狀態
    const actualStatus = String(sheet.getRange(rowNumber, 10).getValue()).trim().toLowerCase();
    Logger.log(`✅ 審核完成: 預期=${status}, 實際=${actualStatus}`);
    
    if (actualStatus !== status) {
      Logger.log(`❌ 狀態不符！`);
      return {
        ok: false,
        msg: `狀態寫入異常：預期 ${status}，實際 ${actualStatus}`
      };
    }
    
    // 👉 發送 LINE 通知
    try {
      notifyOvertimeReview(
        employeeId,
        employeeName,
        formatDate(overtimeDate),
        hours,
        user.name,           // 審核人姓名
        isApprove,
        comment || ""
      );
      Logger.log(`📤 已發送加班審核通知給 ${employeeName} (${employeeId})`);
    } catch (err) {
      Logger.log(`⚠️ LINE 通知發送失敗: ${err.message}`);
      // 通知失敗不影響審核流程
    }
    
    // 🔧 關鍵修正：根據 isApprove 決定回傳碼
    const resultCode = isApprove ? "OVERTIME_APPROVED" : "OVERTIME_REJECTED";
    Logger.log(`✅ 返回結果碼: ${resultCode}`);
    
    return { 
      ok: true, 
      code: resultCode
    };
    
  } catch (error) {
    Logger.log(`❌ 審核失敗: ${error.message}`);
    return { 
      ok: false, 
      msg: `審核失敗: ${error.message}` 
    };
  }
}

/**
 * 格式化日期
 */
function formatDate(date) {
  if (!date) return "";
  if (typeof date === "string") return date;
  return Utilities.formatDate(date, "Asia/Taipei", "yyyy-MM-dd");
}

/**
 * 🔧 升級工具：為現有工作表新增補休時數欄位
 * 只需執行一次
 */
function upgradeOvertimeSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_OVERTIME);
  
  if (!sheet) {
    Logger.log("⚠️ 工作表不存在，將建立新表");
    initOvertimeSheet();
    return;
  }
  
  // 檢查是否已有第 15 欄
  const lastCol = sheet.getLastColumn();
  
  if (lastCol >= 15) {
    const header15 = sheet.getRange(1, 15).getValue();
    if (header15 === "補休時數") {
      Logger.log("✅ 已存在補休時數欄位，無需升級");
      return;
    }
  }
  
  // 新增第 15 欄標題
  sheet.getRange(1, 15).setValue("補休時數").setFontWeight("bold");
  sheet.setColumnWidth(15, 80);
  
  // 為所有現有記錄填入預設值 0
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const defaultValues = Array(lastRow - 1).fill([0]);
    sheet.getRange(2, 15, lastRow - 1, 1).setValues(defaultValues);
  }
  
  Logger.log(`✅ 升級完成！已為 ${lastRow - 1} 筆記錄新增補休時數欄位`);
}
// ==================== Handlers ====================

function handleGetEmployeeOvertime(params) {
  Logger.log(`📥 查詢員工加班記錄`);
  return getEmployeeOvertimeRequests(params.token);
}

function handleGetPendingOvertime(params) {
  Logger.log(`📥 查詢待審核加班申請`);
  return getPendingOvertimeRequests(params.token);
}

/**
 * 審核加班申請
 * 🔧 修正：接收 reviewAction 參數
 */
function handleReviewOvertime(params) {
  const { token, rowNumber, reviewAction, comment } = params;  // ✅ 改用 reviewAction
  
  Logger.log(`📥 handleReviewOvertime 收到參數:`);
  Logger.log(`   - rowNumber: ${rowNumber}`);
  Logger.log(`   - reviewAction: "${reviewAction}"`);  // ✅ 改用 reviewAction
  Logger.log(`   - comment: "${comment}"`);
  
  return reviewOvertimeRequest(
    token, 
    parseInt(rowNumber), 
    reviewAction,  // ✅ 改用 reviewAction
    comment || ""
  );
}

function testDriverOvertime() {
  const employeeId = 'Uf664a35632b736301d674d8b2cc3f8c0'; // ⚠️ 換成實際司機ID
  const yearMonth = '2026-03';
  
  const result = calculateMonthlySalary(employeeId, yearMonth);
  
  Logger.log('加班費: ' + (result.data.weekdayOvertimePay + result.data.restdayOvertimePay + result.data.holidayOvertimePay));
  Logger.log('總加班時數: ' + result.data.totalOvertimeHours);
  
  const totalPay = result.data.weekdayOvertimePay + result.data.restdayOvertimePay + result.data.holidayOvertimePay;
  const rate = totalPay / result.data.totalOvertimeHours;
  
  Logger.log('實際每小時費率: $' + rate.toFixed(2));
  
  if (Math.abs(rate - 200) < 1) {
    Logger.log('✅ 確認！加班費固定 $200/小時');
  } else {
    Logger.log('❌ 費率不對，實際是 $' + rate.toFixed(2) + '/小時');
  }
}