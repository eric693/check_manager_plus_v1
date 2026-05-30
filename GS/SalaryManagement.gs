// SalaryManagement-Enhanced.gs - 薪資管理系統（完整版 - 修正版）

// ==================== 常數定義 ====================

const SHEET_SALARY_CONFIG_ENHANCED = "員工薪資設定";
const SHEET_MONTHLY_SALARY_ENHANCED = "月薪資記錄";

// 加班費率
const OVERTIME_RATES = {
  weekday: 1.34,      // 平日加班（前2小時）
  weekdayExtra: 1.67, // 平日加班（第3小時起）
  restday: 1.34,      // 休息日前2小時
  restdayExtra: 1.67, // 休息日第3小時起
  holiday: 2.0        // 國定假日
};

/**
 * ✅ 判斷日期是平日/休息日/例假日
 * @param {string} dateStr - 日期字串 (YYYY-MM-DD)
 * @returns {string} 'weekday' | 'restday' | 'holiday'
 */
function getDateType(dateStr) {
  try {
    const date = new Date(dateStr);
    const dayOfWeek = date.getDay(); // 0=週日, 1=週一, ..., 6=週六
    
    // 週日 = 例假日
    if (dayOfWeek === 0) {
      return 'holiday';
    }
    
    // 週六 = 休息日
    if (dayOfWeek === 6) {
      return 'restday';
    }
    
    // ⭐ TODO: 可以再加上國定假日判斷
    // 例如：if (isNationalHoliday(dateStr)) return 'holiday';
    
    // 週一~週五 = 平日
    return 'weekday';
    
  } catch (error) {
    Logger.log('❌ 判斷日期類型失敗: ' + error);
    return 'weekday'; // 預設為平日
  }
}

/**
 * ✅ 計算加班費（根據日期類型）
 * @param {number} hours - 加班時數
 * @param {number} hourlyRate - 時薪
 * @param {string} dateType - 日期類型
 * @returns {Object} { firstPay, secondPay, thirdPay }
 */
function calculateOvertimePay(hours, hourlyRate, dateType) {
  let firstPay = 0;   // 前2小時
  let secondPay = 0;  // 3-8小時
  let thirdPay = 0;   // 9小時起
  
  if (dateType === 'weekday') {
    // 平日加班：前2h ×1.34，3h起 ×1.67
    const first = Math.min(hours, 2);
    firstPay = hourlyRate * first * 1.34;
    
    if (hours > 2) {
      const rest = Math.min(hours - 2, 2); // 最多再算2小時（總共4h）
      secondPay = hourlyRate * rest * 1.67;
    }
    
  } else if (dateType === 'restday') {
    // 休息日（週六）：前2h ×1.34，3-8h ×1.67，9h起 ×2.67
    const first = Math.min(hours, 2);
    firstPay = hourlyRate * first * 1.34;
    
    if (hours > 2) {
      const second = Math.min(hours - 2, 6); // 3-8h
      secondPay = hourlyRate * second * 1.67;
    }
    
    if (hours > 8) {
      const third = hours - 8; // 9h起
      thirdPay = hourlyRate * third * 2.67;
    }
    
  } else if (dateType === 'holiday') {
    // 例假日/國定假日（週日）：全天 ×2.0
    firstPay = hourlyRate * hours * 2.0;
  }
  
  return {
    firstPay: Math.round(firstPay),
    secondPay: Math.round(secondPay),
    thirdPay: Math.round(thirdPay)
  };
}
/**
 * ✅ 統一的 JSON 回應格式
 */
function jsonResponse(ok, data, message, code) {
  const response = {
    ok: ok,
    success: ok,
    data: data,
    records: data,
    msg: message,
    message: message,
    code: code || (ok ? 'SUCCESS' : 'ERROR')
  };
  
  return ContentService
    .createTextOutput(JSON.stringify(response))
    .setMimeType(ContentService.MimeType.JSON);
}
// ==================== 試算表管理 ====================

// SalaryManagement-Enhanced.gs - 修正 Sheet 結構

/**
 * ✅ 取得或建立員工薪資設定試算表（修正版 - 新增工時類型）
 */
function getEmployeeSalarySheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_SALARY_CONFIG_ENHANCED);
  
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_SALARY_CONFIG_ENHANCED);
    
    const headers = [
      // ⭐ 基本資訊 (7欄: A-G) - 新增「工時類型」
      "員工ID", "員工姓名", "身分證字號", "員工類型", "薪資類型", "工時類型", "基本薪資",
      
      // 固定津貼項目 (8欄: H-O)
      "職務加給", "伙食費", "交通補助", "全勤獎金", "績效獎金", 
      "其他津貼1", "其他津貼2", "其他津貼3",
      
      // 銀行資訊 (4欄: P-S)
      "銀行代碼", "銀行帳號", "到職日期", "發薪日",
      
      // 法定扣款 (6欄: T-Y)
      "勞退自提率(%)", "勞保費", "健保費", "就業保險費", "勞退自提", "所得稅",
      
      // 其他扣款 (5欄: Z-AD)
      "福利金扣款", "宿舍費用", "團保費用", 
      "其他扣款1", "其他扣款2",
      
      // 系統欄位 (3欄: AE-AG)
      "狀態", "備註", "備註歷程", "最後更新時間"
    ];
    
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
    sheet.getRange(1, 1, 1, headers.length).setBackground("#10b981");
    sheet.getRange(1, 1, 1, headers.length).setFontColor("#ffffff");
    sheet.setFrozenRows(1);
    
    Logger.log("✅ 建立員工薪資設定試算表（修正版，共 " + headers.length + " 欄）");
  }
  
  return sheet;
}

/**
 * ✅ 設定員工薪資資料（修正版 - 包含工時類型）
 */
function setEmployeeSalaryTW(salaryData) {
  try {
    Logger.log('💰 開始設定員工薪資（修正版）');
    
    const sheet = getEmployeeSalarySheet();
    const data = sheet.getDataRange().getValues();
    
    // 驗證必填欄位
    if (!salaryData.employeeId || !salaryData.employeeName || !salaryData.baseSalary || salaryData.baseSalary <= 0) {
      return { success: false, message: "缺少必填欄位或基本薪資無效" };
    }
    
    // 檢查是否已存在
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(salaryData.employeeId).trim()) {
        rowIndex = i + 1;
        break;
      }
    }
    
    const now = new Date();
    
    const row = [
      // A-G: 基本資訊 (7欄) ⭐ 新增工時類型
      String(salaryData.employeeId).trim(),
      String(salaryData.employeeName).trim(),
      String(salaryData.idNumber || "").trim(),
      String(salaryData.employeeType || "正職").trim(),
      String(salaryData.salaryType || "月薪").trim(),
      String(salaryData.workTimeType || "標準工時").trim(),  // ⭐ 新增
      parseFloat(salaryData.baseSalary) || 0,
      
      // H-O: 固定津貼項目 (8欄)
      parseFloat(salaryData.positionAllowance) || 0,
      parseFloat(salaryData.mealAllowance) || 0,
      parseFloat(salaryData.transportAllowance) || 0,
      parseFloat(salaryData.attendanceBonus) || 0,
      parseFloat(salaryData.performanceBonus) || 0,
      parseFloat(salaryData.otherAllowance1) || 0,
      parseFloat(salaryData.otherAllowance2) || 0,
      parseFloat(salaryData.otherAllowance3) || 0,
      
      // P-S: 銀行資訊 (4欄)
      String(salaryData.bankCode || "").trim(),
      String(salaryData.bankAccount || "").trim(),
      salaryData.hireDate || "",
      String(salaryData.paymentDay || "5").trim(),
      
      // T-Y: 法定扣款 (6欄)
      parseFloat(salaryData.pensionSelfRate) || 0,
      parseFloat(salaryData.laborFee) || 0,
      parseFloat(salaryData.healthFee) || 0,
      parseFloat(salaryData.employmentFee) || 0,
      parseFloat(salaryData.pensionSelf) || 0,
      parseFloat(salaryData.incomeTax) || 0,
      
      // Z-AD: 其他扣款 (5欄)
      parseFloat(salaryData.welfareFee) || 0,
      parseFloat(salaryData.dormitoryFee) || 0,
      parseFloat(salaryData.groupInsurance) || 0,
      parseFloat(salaryData.otherDeduction1) || 0,
      parseFloat(salaryData.otherDeduction2) || 0,
      
      // AE-AG: 系統欄位 (3欄)
      "在職",
      String(salaryData.note || "").trim(),
      buildNoteHistory(salaryData.employeeId, salaryData.note || ""),
      now
    ];
    
    if (rowIndex > 0) {
      sheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);
      Logger.log(`✅ 更新員工薪資設定: ${salaryData.employeeName}`);
    } else {
      sheet.appendRow(row);
      Logger.log(`✅ 新增員工薪資設定: ${salaryData.employeeName}`);
    }
    
    Logger.log('');
    Logger.log('🔄 開始自動同步到月薪資記錄...');

    const currentYearMonth = Utilities.formatDate(now, "Asia/Taipei", "yyyy-MM");
    Logger.log(`   當月: ${currentYearMonth}`);

    const recalculated = calculateMonthlySalary(salaryData.employeeId, currentYearMonth);

    if (recalculated.success || recalculated.ok) {
      Logger.log('   ✅ 薪資計算成功');
      
      const saveResult = saveMonthlySalary(recalculated.data);
      
      if (saveResult.success) {
        Logger.log(`   ✅ 已同步到月薪資記錄: ${saveResult.salaryId}`);
      } else {
        Logger.log(`   ⚠️ 同步失敗: ${saveResult.message}`);
      }
    } else {
      Logger.log(`   ⚠️ 薪資計算失敗: ${recalculated.message || recalculated.msg}`);
    }

    Logger.log('');
    return { success: true, message: "薪資設定成功" };
    
  } catch (error) {
    Logger.log("❌ 設定薪資失敗: " + error);
    return { success: false, message: error.toString() };
  }
}

function rebuildMonthlySalarySheet() {
     // 刪除舊表（如果存在）
     const ss = SpreadsheetApp.getActiveSpreadsheet();
     const oldSheet = ss.getSheetByName('月薪資記錄');
     if (oldSheet) {
       ss.deleteSheet(oldSheet);
     }
     
     // 建立新表
     getMonthlySalarySheetEnhanced();
     
     Logger.log('✅ 月薪資記錄試算表已重建');
   }

function getMonthlySalarySheetEnhanced() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_MONTHLY_SALARY_ENHANCED);
  
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_MONTHLY_SALARY_ENHANCED);
    
    const headers = [
      // A-I: 基本資訊 (9 欄)
      '薪資單ID', '員工ID', '員工姓名', '年月', '薪資類型', '工時類型', 
      '時薪', '工作時數', '總加班時數',

      // J-Y: 收入項目 (16 欄)
      '基本薪資', '職務加給', '伙食費', '交通補助', '全勤獎金', '績效獎金', 
      '其他津貼1', '其他津貼2', '其他津貼3',
      '平日加班費', '休息日加班費', '國定假日加班費',
      '未休假補薪', '未休假天數', '月休補薪', '未休月休天數',

      // Z-AE: 法定扣款 (6 欄)
      '勞保費', '健保費', '就業保險費', '勞退自提率(%)', '勞退自提', '所得稅',

      // AF-AN: 其他扣款 (8 欄) ⭐⭐⭐ 新增 4 欄
      '請假扣款',           // AF（總計）
      '病假時數',           // AG ⭐ 新增
      '病假扣款',           // AH ⭐ 新增
      '事假時數',           // AI ⭐ 新增
      '事假扣款',           // AJ ⭐ 新增
      '福利金扣款',         // AK
      '宿舍費用',           // AL
      '團保費用',           // AM
      '其他扣款1',          // AN
      '其他扣款2',          // AO

      // AP-AQ: 總額 (2 欄)
      '應發總額', '實發金額',

      // AR-AS: 銀行 (2 欄)
      '銀行代碼', '銀行帳號',

      // AT-AV: 系統 (3 欄)
      '狀態', '備註', '建立時間'
    ];
    
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
    sheet.getRange(1, 1, 1, headers.length).setBackground("#10b981");
    sheet.getRange(1, 1, 1, headers.length).setFontColor("#ffffff");
    sheet.setFrozenRows(1);
    
    Logger.log(`✅ 建立月薪資記錄試算表（修正版），共 ${headers.length} 個欄位`);
  }
  
  return sheet;
}

// ==================== 薪資設定功能 ====================

/**
 * ✅ 取得員工薪資設定（完整版）
 */
function getEmployeeSalaryTW(employeeId) {
  try {
    const sheet = getEmployeeSalarySheet();
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(employeeId).trim()) {
        const salaryConfig = {};
        headers.forEach((header, index) => {
          salaryConfig[header] = data[i][index];
        });
        
        return { success: true, data: salaryConfig };
      }
    }
    
    return { success: false, message: "找不到該員工薪資資料" };
    
  } catch (error) {
    Logger.log("❌ 取得薪資設定失敗: " + error);
    return { success: false, message: error.toString() };
  }
}

/**
 * ✅ 同步薪資到月薪資記錄（完整版 - 修正）
 */
function syncSalaryToMonthlyRecord(employeeId, yearMonth) {
  try {
    const salaryConfig = getEmployeeSalaryTW(employeeId);
    
    if (!salaryConfig.success) {
      return { success: false, message: "找不到員工薪資設定" };
    }
    
    const config = salaryConfig.data;
    const calculatedSalary = calculateMonthlySalary(employeeId, yearMonth);
    
    if (!calculatedSalary.success) {
      // ⭐⭐⭐ 關鍵修正：先判斷薪資類型
      const salaryType = String(config['薪資類型'] || '月薪').trim();
      const isHourly = salaryType === '時薪';
      
      // 建立基本薪資記錄
      const totalAllowances = 
        (parseFloat(config['職務加給']) || 0) +
        (parseFloat(config['伙食費']) || 0) +
        (parseFloat(config['交通補助']) || 0) +
        (parseFloat(config['全勤獎金']) || 0) +
        (parseFloat(config['績效獎金']) || 0) +
        (parseFloat(config['其他津貼1']) || 0) +  // ⭐ 改名
        (parseFloat(config['其他津貼2']) || 0) +  // ⭐ 改名
        (parseFloat(config['其他津貼3']) || 0);   // ⭐ 改名
      
      const totalDeductions = 
        (parseFloat(config['勞保費']) || 0) +
        (parseFloat(config['健保費']) || 0) +
        (parseFloat(config['就業保險費']) || 0) +
        (parseFloat(config['勞退自提']) || 0) +
        (parseFloat(config['所得稅']) || 0) +
        (parseFloat(config['福利金扣款']) || 0) +
        (parseFloat(config['宿舍費用']) || 0) +
        (parseFloat(config['團保費用']) || 0) +
        (parseFloat(config['其他扣款1']) || 0) +  // ⭐ 改名
        (parseFloat(config['其他扣款2']) || 0);   // ⭐ 改名
      
      // ⭐ 現在可以安全使用 isHourly 了
      const baseAmount = isHourly ? 0 : parseFloat(config['基本薪資']);
      const grossSalary = baseAmount + totalAllowances;
      
      const basicSalary = {
        employeeId: employeeId,
        employeeName: config['員工姓名'],
        yearMonth: yearMonth,
        
        // ⭐⭐⭐ 新增：薪資類型相關欄位
        salaryType: salaryType,
        hourlyRate: isHourly ? parseFloat(config['基本薪資']) : 0,
        totalWorkHours: 0,
        totalOvertimeHours: 0,
        
        baseSalary: isHourly ? 0 : parseFloat(config['基本薪資']),
        positionAllowance: config['職務加給'] || 0,
        mealAllowance: config['伙食費'] || 0,
        transportAllowance: config['交通補助'] || 0,
        attendanceBonus: config['全勤獎金'] || 0,
        performanceBonus: config['績效獎金'] || 0,
        otherAllowance1: config['其他津貼1'] || 0,  // ⭐ 改名
        otherAllowance2: config['其他津貼2'] || 0,  // ⭐ 改名
        otherAllowance3: config['其他津貼3'] || 0,  // ⭐ 改名
        weekdayOvertimePay: 0,
        restdayOvertimePay: 0,
        holidayOvertimePay: 0,
        laborFee: config['勞保費'] || 0,
        healthFee: config['健保費'] || 0,
        employmentFee: config['就業保險費'] || 0,
        pensionSelf: config['勞退自提'] || 0,
        incomeTax: config['所得稅'] || 0,
        leaveDeduction: 0,
        welfareFee: config['福利金扣款'] || 0,
        dormitoryFee: config['宿舍費用'] || 0,
        groupInsurance: config['團保費用'] || 0,
        otherDeduction1: config['其他扣款1'] || 0,  // ⭐ 改名
        otherDeduction2: config['其他扣款2'] || 0,  // ⭐ 改名
        grossSalary: grossSalary,
        netSalary: grossSalary - totalDeductions,
        bankCode: config['銀行代碼'] || "",
        bankAccount: config['銀行帳號'] || "",
        status: "已設定",
        note: "自動建立"
      };
      
      return saveMonthlySalary(basicSalary);
    }
    
    return saveMonthlySalary(calculatedSalary.data);
    
  } catch (error) {
    Logger.log(`❌ 同步失敗: ${error}`);
    return { success: false, message: error.toString() };
  }
}

// ==================== 薪資計算功能 ====================
/**
 * ✅ 取得員工該月份的加班記錄（完整修正版）
 */
function getEmployeeMonthlyOvertime(employeeId, yearMonth) {
  try {
    Logger.log('📋 取得加班記錄');
    Logger.log('   員工ID: ' + employeeId);
    Logger.log('   年月: ' + yearMonth);
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('加班申請');
    
    if (!sheet) {
      Logger.log('⚠️ 找不到「加班申請」工作表');
      return [];
    }
    
    const data = sheet.getDataRange().getValues();
    
    if (data.length < 2) {
      Logger.log('⚠️ 「加班申請」工作表無資料');
      return [];
    }
    
    const headers = data[0];
    Logger.log('📊 加班申請欄位: ' + headers.join(', '));
    
    // ⭐⭐⭐ 欄位索引（根據實際 Sheet）
    const userIdIndex = 1;      // 員工ID (B欄)
    const dateIndex = 3;        // 加班日期 (D欄)
    const hoursIndex = 6;       // 加班時數 (G欄)
    const statusIndex = 9;      // 審核狀態 (J欄)
    
    const records = [];
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      const rowUserId = String(row[userIdIndex] || '').trim();
      const overtimeDate = row[dateIndex];
      const hours = row[hoursIndex];
      const status = String(row[statusIndex] || '').trim().toLowerCase();
      
      // ⭐⭐⭐ 關鍵修正：先檢查員工ID
      if (rowUserId !== employeeId) {
        continue;
      }
      
      // ⭐⭐⭐ 關鍵修正：只計算已核准的
      if (status !== 'approved') {
        Logger.log(`   ⏭️ 跳過未核准: ${overtimeDate} (狀態: ${status})`);
        continue;
      }
      
      // ⭐⭐⭐ 關鍵修正：處理 Date 物件
      let dateStr = '';
      
      if (overtimeDate instanceof Date) {
        // 如果是 Date 物件，轉換為 YYYY-MM-DD
        dateStr = Utilities.formatDate(overtimeDate, 'Asia/Taipei', 'yyyy-MM-dd');
      } else if (typeof overtimeDate === 'string') {
        // 如果已經是字串，直接使用
        dateStr = overtimeDate;
      } else {
        // 其他情況，跳過
        Logger.log(`   ⚠️ 無法解析日期: ${overtimeDate}`);
        continue;
      }
      
      // ⭐⭐⭐ 關鍵修正：提取年月 (YYYY-MM)
      const recordYearMonth = dateStr.substring(0, 7);
      
      // 比對年月
      if (recordYearMonth !== yearMonth) {
        Logger.log(`   ⏭️ 跳過不符月份: ${dateStr} (${recordYearMonth} ≠ ${yearMonth})`);
        continue;
      }
      
      // ⭐⭐⭐ 加入記錄
      const hoursNum = parseFloat(hours) || 0;
      
      records.push({
        date: dateStr,
        hours: hoursNum
      });
      
      Logger.log(`   ✅ ${dateStr}: ${hoursNum}h (狀態: ${status})`);
    }
    
    Logger.log(`✅ 找到 ${records.length} 筆加班記錄`);
    
    return records;
    
  } catch (error) {
    Logger.log('❌ 取得加班記錄失敗: ' + error);
    Logger.log('❌ 錯誤堆疊: ' + error.stack);
    return [];
  }
}
// ==================================================================================
// saveMonthlySalary - 完整修正版（支援情況一、二、三）
// ==================================================================================

/**
 * ✅ 儲存月薪資記錄（完整版 - 支援三種情況）
 * 
 * @param {Object} salaryData - 薪資資料物件
 * @returns {Object} 儲存結果
 */
function saveMonthlySalary(salaryData) {
  try {
    const sheet = getMonthlySalarySheetEnhanced();
    
    // 正規化年月格式
    let normalizedYearMonth = salaryData.yearMonth;
    
    if (salaryData.yearMonth instanceof Date) {
      normalizedYearMonth = Utilities.formatDate(salaryData.yearMonth, "Asia/Taipei", "yyyy-MM");
    } else if (typeof salaryData.yearMonth === 'string') {
      normalizedYearMonth = salaryData.yearMonth.substring(0, 7);
    }
    
    const salaryId = `SAL-${normalizedYearMonth}-${salaryData.employeeId}`;
    const salaryType = salaryData.salaryType || '月薪';
    const workTimeType = salaryData.workTimeType || '標準工時';
    
    // ⭐⭐⭐ 關鍵修正：銀行代碼處理
    let bankCode = salaryData.bankCode || '';
    
    // 1. 移除所有非數字字符
    bankCode = String(bankCode).replace(/[^0-9]/g, '');
    
    // 2. 補零到 3 位數
    if (bankCode.length > 0) {
      bankCode = bankCode.padStart(3, '0');
    }
    
    // 3. 強制加上單引號前綴（防止被轉成日期或數字）
    bankCode = "'" + bankCode;  // '007
    
    // ⭐⭐⭐ 完整的 row 陣列（支援所有情況）
    const row = [
      // A-I: 基本資訊 (9欄)
      salaryId,
      salaryData.employeeId,
      salaryData.employeeName,
      normalizedYearMonth,
      salaryType,
      workTimeType,
      salaryData.hourlyRate || 0,
      salaryData.totalWorkHours || 0,
      salaryData.totalOvertimeHours || 0,
      
      // J-Y: 應發項目 (16欄)
      salaryData.baseSalary || 0,
      salaryData.positionAllowance || 0,
      salaryData.mealAllowance || 0,
      salaryData.transportAllowance || 0,
      salaryData.attendanceBonus || 0,
      salaryData.performanceBonus || 0,
      salaryData.otherAllowance1 || 0,
      salaryData.otherAllowance2 || 0,
      salaryData.otherAllowance3 || 0,
      
      // ⭐⭐⭐ 加班費（根據情況使用不同欄位）
      // 情況一、二：平日/休息日/國定假日
      // 情況三：延長工時（前2h）/延長工時（後2h）/休息日/國定假日
      salaryData.weekdayOvertimePay || salaryData.extendedOvertimeFirst2h || 0,  // S欄
      salaryData.restdayOvertimePay || salaryData.extendedOvertimeAfter2h || 0,   // T欄（情況三改用途）
      salaryData.holidayOvertimePay || 0,                                          // U欄
      
      salaryData.unusedLeavePay || 0,
      salaryData.unusedLeaveDays || 0,
      salaryData.monthlyRestPay || 0,        // ⭐ 情況一專用
      salaryData.monthlyRestMissedDays || 0, // ⭐ 情況一專用
      
      // Z-AE: 法定扣款 (6欄)
      salaryData.laborFee || 0,
      salaryData.healthFee || 0,
      salaryData.employmentFee || 0,
      salaryData.pensionSelfRate || 0,
      salaryData.pensionSelf || 0,
      salaryData.incomeTax || 0,
      
      // AF-AO: 其他扣款 (10欄)
      salaryData.leaveDeduction || 0,
      salaryData.sickLeaveHours || 0,
      salaryData.sickLeaveDeduction || 0,
      salaryData.personalLeaveHours || 0,
      salaryData.personalLeaveDeduction || 0,
      salaryData.welfareFee || 0,
      salaryData.dormitoryFee || 0,
      salaryData.groupInsurance || 0,
      salaryData.otherDeduction1 || 0,
      salaryData.otherDeduction2 || 0,
      
      // AP-AQ: 總計 (2欄)
      salaryData.grossSalary || 0,
      salaryData.netSalary || 0,
      
      // AR-AS: 銀行資訊 (2欄)
      bankCode,                    // AR: 銀行代碼（強制文字格式）
      salaryData.bankAccount || "",  // AS: 銀行帳號
      
      // AT-AV: 系統欄位 (3欄)
      salaryData.status || "已計算",
      salaryData.note || "",
      new Date()
    ];
    
    Logger.log('📝 準備寫入的 row 長度: ' + row.length);
    Logger.log('   薪資單ID: ' + salaryId);
    Logger.log('   員工: ' + salaryData.employeeName);
    Logger.log('   薪資類型: ' + salaryType);
    Logger.log('   工時類型: ' + workTimeType);
    
    const data = sheet.getDataRange().getValues();
    let found = false;
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === salaryId) {
        sheet.getRange(i + 1, 1, 1, row.length).setValues([row]);
        found = true;
        Logger.log(`✅ 更新薪資單: ${salaryId}`);
        break;
      }
    }
    
    if (!found) {
      const lastRow = sheet.getLastRow();
      sheet.getRange(lastRow + 1, 1, 1, row.length).setValues([row]);
      Logger.log(`✅ 新增薪資單: ${salaryId}`);
    }
    
    return { success: true, salaryId: salaryId, message: "薪資單儲存成功" };
    
  } catch (error) {
    Logger.log("❌ 儲存薪資單失敗: " + error);
    Logger.log("❌ 錯誤堆疊: " + error.stack);
    return { success: false, message: error.toString() };
  }
}

/**
 * ⚠️ 重要提醒：欄位對應關係
 * 
 * S欄（平日加班費 / 延長工時前2h）：
 * - 情況一、二：平日加班費
 * - 情況三：延長工時加班費（前2h）
 * 
 * T欄（休息日加班費 / 延長工時後2h）：
 * - 情況一、二：休息日加班費
 * - 情況三：延長工時加班費（後2h）+ 休息日加班費（需合併）
 * 
 * U欄（國定假日加班費）：
 * - 所有情況通用
 */

console.log('✅ saveMonthlySalary 已更新（支援情況一、二、三）');
// function saveMonthlySalary(salaryData) {
//   try {
//     const sheet = getMonthlySalarySheetEnhanced();
    
//     let normalizedYearMonth = salaryData.yearMonth;
    
//     if (salaryData.yearMonth instanceof Date) {
//       normalizedYearMonth = Utilities.formatDate(salaryData.yearMonth, "Asia/Taipei", "yyyy-MM");
//     } else if (typeof salaryData.yearMonth === 'string') {
//       normalizedYearMonth = salaryData.yearMonth.substring(0, 7);
//     }
    
//     const salaryId = `SAL-${normalizedYearMonth}-${salaryData.employeeId}`;
//     const salaryType = salaryData.salaryType || '月薪';
//     const workTimeType = salaryData.workTimeType || '標準工時';
    
//     const row = [
//       // A-I: 基本資訊 (9欄)
//       salaryId,
//       salaryData.employeeId,
//       salaryData.employeeName,
//       normalizedYearMonth,
//       salaryType,
//       workTimeType,
//       salaryData.hourlyRate || 0,
//       salaryData.totalWorkHours || 0,
//       salaryData.totalOvertimeHours || 0,
      
//       // J-Y: 應發項目 (16欄)
//       salaryData.baseSalary || 0,
//       salaryData.positionAllowance || 0,
//       salaryData.mealAllowance || 0,
//       salaryData.transportAllowance || 0,
//       salaryData.attendanceBonus || 0,
//       salaryData.performanceBonus || 0,
//       salaryData.otherAllowance1 || 0,
//       salaryData.otherAllowance2 || 0,
//       salaryData.otherAllowance3 || 0,
//       salaryData.weekdayOvertimePay || 0,
//       salaryData.restdayOvertimePay || 0,
//       salaryData.holidayOvertimePay || 0,
//       salaryData.unusedLeavePay || 0,
//       salaryData.unusedLeaveDays || 0,
//       salaryData.monthlyRestPay || 0,
//       salaryData.monthlyRestMissedDays || 0,
      
//       // Z-AE: 法定扣款 (6欄)
//       salaryData.laborFee || 0,
//       salaryData.healthFee || 0,
//       salaryData.employmentFee || 0,
//       salaryData.pensionSelfRate || 0,
//       salaryData.pensionSelf || 0,
//       salaryData.incomeTax || 0,
      
//       // AF-AO: 其他扣款 (10欄) ⭐⭐⭐ 新增 4 欄
//       salaryData.leaveDeduction || 0,           // AF: 請假扣款總計
//       salaryData.sickLeaveHours || 0,           // AG ⭐
//       salaryData.sickLeaveDeduction || 0,       // AH ⭐
//       salaryData.personalLeaveHours || 0,       // AI ⭐
//       salaryData.personalLeaveDeduction || 0,   // AJ ⭐
//       salaryData.welfareFee || 0,               // AK
//       salaryData.dormitoryFee || 0,             // AL
//       salaryData.groupInsurance || 0,           // AM
//       salaryData.otherDeduction1 || 0,          // AN
//       salaryData.otherDeduction2 || 0,          // AO
      
//       // AP-AQ: 總計 (2欄)
//       salaryData.grossSalary || 0,
//       salaryData.netSalary || 0,
      
//       // AR-AS: 銀行資訊 (2欄)
//       salaryData.bankCode || "",
//       salaryData.bankAccount || "",
      
//       // AT-AV: 系統欄位 (3欄)
//       salaryData.status || "已計算",
//       salaryData.note || "",
//       new Date()
//     ];
    
//     Logger.log('📝 準備寫入的 row 長度: ' + row.length);
//     Logger.log('   前5個欄位: ' + row.slice(0, 5).join(', '));
    
//     const data = sheet.getDataRange().getValues();
//     let found = false;
    
//     for (let i = 1; i < data.length; i++) {
//       if (data[i][0] === salaryId) {
//         sheet.getRange(i + 1, 1, 1, row.length).setValues([row]);
//         found = true;
//         Logger.log(`✅ 更新薪資單: ${salaryId}`);
//         break;
//       }
//     }
    
//     if (!found) {
//       // ⭐⭐⭐ 關鍵修正：改用 setValues 寫入
//       const lastRow = sheet.getLastRow();
//       sheet.getRange(lastRow + 1, 1, 1, row.length).setValues([row]);
//       Logger.log(`✅ 新增薪資單: ${salaryId}（使用 setValues）`);
//     }
    
//     return { success: true, salaryId: salaryId, message: "薪資單儲存成功" };
    
//   } catch (error) {
//     Logger.log("❌ 儲存薪資單失敗: " + error);
//     return { success: false, message: error.toString() };
//   }
// }
/**
 * ✅ 查詢我的薪資（完整版）
 */
function getMySalary(userId, yearMonth) {
  try {
    const employeeId = userId;
    const sheet = getMonthlySalarySheetEnhanced();
    const data = sheet.getDataRange().getValues();
    
    if (data.length < 2) {
      return { success: false, message: "薪資記錄表中沒有資料" };
    }
    
    const headers = data[0];
    const employeeIdIndex = headers.indexOf('員工ID');
    const yearMonthIndex = headers.indexOf('年月');
    
    if (employeeIdIndex === -1 || yearMonthIndex === -1) {
      return { success: false, message: "試算表缺少必要欄位" };
    }
    
    for (let i = 1; i < data.length; i++) {
      const rowEmployeeId = String(data[i][employeeIdIndex]).trim();
      const rawYearMonth = data[i][yearMonthIndex];
      
      let normalizedYearMonth = '';
      
      if (rawYearMonth instanceof Date) {
        normalizedYearMonth = Utilities.formatDate(rawYearMonth, 'Asia/Taipei', 'yyyy-MM');
      } else if (typeof rawYearMonth === 'string') {
        normalizedYearMonth = rawYearMonth.substring(0, 7);
      } else {
        normalizedYearMonth = String(rawYearMonth).substring(0, 7);
      }
      
      if (rowEmployeeId === employeeId && normalizedYearMonth === yearMonth) {
        const salary = {};
        headers.forEach((header, index) => {
          if (header === '年月' && data[i][index] instanceof Date) {
            salary[header] = Utilities.formatDate(data[i][index], 'Asia/Taipei', 'yyyy-MM');
          } else {
            salary[header] = data[i][index];
          }
        });
        
        return { success: true, data: salary };
      }
    }
    
    return { success: false, message: "查無薪資記錄" };
    
  } catch (error) {
    Logger.log('❌ 查詢薪資失敗: ' + error);
    return { success: false, message: error.toString() };
  }
}

/**
 * ✅ 查詢我的薪資歷史（完整版）
 */
function getMySalaryHistory(userId, limit = 12) {
  try {
    const employeeId = userId;
    const sheet = getMonthlySalarySheetEnhanced();
    const data = sheet.getDataRange().getValues();
    
    if (data.length < 2) {
      return { success: true, data: [], total: 0 };
    }
    
    const headers = data[0];
    const employeeIdIndex = headers.indexOf('員工ID');
    
    if (employeeIdIndex === -1) {
      return { success: false, message: "試算表缺少「員工ID」欄位" };
    }
    
    const salaries = [];
    
    for (let i = 1; i < data.length; i++) {
      const rowEmployeeId = String(data[i][employeeIdIndex]).trim();
      
      if (rowEmployeeId === employeeId) {
        const salary = {};
        headers.forEach((header, index) => {
          if (header === '年月' && data[i][index] instanceof Date) {
            salary[header] = Utilities.formatDate(data[i][index], "Asia/Taipei", "yyyy-MM");
          } else {
            salary[header] = data[i][index];
          }
        });
        salaries.push(salary);
      }
    }
    
    salaries.sort((a, b) => {
      const yearMonthA = String(a['年月'] || '');
      const yearMonthB = String(b['年月'] || '');
      return yearMonthB.localeCompare(yearMonthA);
    });
    
    const result = salaries.slice(0, limit);
    
    return { success: true, data: result, total: salaries.length };
    
  } catch (error) {
    Logger.log("❌ 查詢薪資歷史失敗: " + error);
    return { success: false, message: error.toString() };
  }
}

/**
 * ✅ 查詢所有員工的月薪資列表（完整版）
 */
function getAllMonthlySalary(yearMonth) {
  try {
    const sheet = getMonthlySalarySheetEnhanced();
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    const salaries = [];
    
    for (let i = 1; i < data.length; i++) {
      const rawYearMonth = data[i][3];
      
      let normalizedYearMonth = '';
      
      if (rawYearMonth instanceof Date) {
        normalizedYearMonth = Utilities.formatDate(rawYearMonth, "Asia/Taipei", "yyyy-MM");
      } else if (typeof rawYearMonth === 'string') {
        normalizedYearMonth = rawYearMonth.substring(0, 7);
      }
      
      if (!yearMonth || normalizedYearMonth === yearMonth) {
        const salary = {};
        headers.forEach((header, index) => {
          if (header === '年月') {
            salary[header] = normalizedYearMonth;
          } else {
            salary[header] = data[i][index];
          }
        });
        salaries.push(salary);
      }
    }
    
    return { success: true, data: salaries };
    
  } catch (error) {
    Logger.log("❌ 查詢薪資列表失敗: " + error);
    return { success: false, message: error.toString() };
  }
}

// ==================== 輔助函數 ====================

/**
 * ✅ 取得員工加班記錄
 */
function getEmployeeOvertimeRecords(employeeId, yearMonth) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("加班申請");
    
    if (!sheet) {
      return { success: true, data: [] };
    }
    
    const values = sheet.getDataRange().getValues();
    const records = [];
    
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      
      if (!row[1] || !row[3]) continue;
      
      const rowEmployeeId = String(row[1]).trim();
      const overtimeDate = row[3];
      
      if (rowEmployeeId !== employeeId) continue;
      
      let dateStr = "";
      if (overtimeDate instanceof Date) {
        dateStr = Utilities.formatDate(overtimeDate, "Asia/Taipei", "yyyy-MM");
      } else if (typeof overtimeDate === "string") {
        dateStr = overtimeDate.substring(0, 7);
      }
      
      if (dateStr !== yearMonth) continue;
      
      const status = String(row[9] || "").trim().toLowerCase();
      if (status !== "approved") continue;
      
      records.push({
        overtimeDate: dateStr,
        overtimeHours: parseFloat(row[6]) || 0,
        overtimeType: "平日加班",
        reviewStatus: "核准"
      });
    }
    
    return { success: true, data: records };
    
  } catch (error) {
    Logger.log("❌ 取得加班記錄失敗: " + error);
    return { success: false, message: error.toString(), data: [] };
  }
}

/**
 * ✅ 取得員工請假記錄
 */
function getEmployeeMonthlySalary(employeeId, yearMonth) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("請假記錄");
    
    if (!sheet) {
      return { success: true, data: [] };
    }
    
    const values = sheet.getDataRange().getValues();
    const records = [];
    
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      
      if (!row[1] || !row[5]) continue;
      
      const rowEmployeeId = String(row[1]).trim();
      const startDate = row[5];
      
      if (rowEmployeeId !== employeeId) continue;
      
      let dateStr = "";
      if (startDate instanceof Date) {
        dateStr = Utilities.formatDate(startDate, "Asia/Taipei", "yyyy-MM");
      } else if (typeof startDate === "string") {
        dateStr = startDate.substring(0, 7);
      }
      
      if (dateStr !== yearMonth) continue;
      
      const status = String(row[9] || "").trim().toUpperCase();
      if (status !== "APPROVED") continue;
      
      records.push({
        leaveType: row[4] || "",
        startDate: startDate,
        leaveDays: parseFloat(row[7]) || 0,
        reviewStatus: "核准"
      });
    }
    
    return { success: true, data: records };
    
  } catch (error) {
    Logger.log("❌ 取得請假記錄失敗: " + error);
    return { success: false, message: error.toString(), data: [] };
  }
}

// ==================== 時薪計算功能 ====================

/**
 * ✅ 計算時薪員工的月薪資（完整修正版）
 * 
 * @param {string} employeeId - 員工ID
 * @param {string} yearMonth - 年月 (YYYY-MM)
 * @returns {Object} 薪資計算結果
 */
function calculateHourlySalary(employeeId, yearMonth) {
  try {
    Logger.log(`⏰ 開始計算時薪薪資: ${employeeId}, ${yearMonth}`);
    
    // 1. 取得員工薪資設定
    const salaryConfig = getEmployeeSalaryTW(employeeId);
    if (!salaryConfig.success) {
      Logger.log('❌ 找不到員工薪資設定');
      return { success: false, message: "找不到員工薪資設定" };
    }
    
    const config = salaryConfig.data;
    const hourlyRate = parseFloat(config['基本薪資']) || 0; // 時薪
    
    Logger.log(`💵 時薪: $${hourlyRate}`);
    
    // 2. ⭐ 取得該月份的打卡記錄
    const attendanceRecords = getEmployeeMonthlyAttendanceInternal(employeeId, yearMonth);
    Logger.log(`📋 找到 ${attendanceRecords.length} 筆打卡記錄`);
    
    // 3. 計算工作時數
    let totalWorkHours = 0;
    
    attendanceRecords.forEach(record => {
      if (record.workHours > 0) {
        totalWorkHours += record.workHours;
        Logger.log(`   ${record.date}: ${record.punchIn} ~ ${record.punchOut} = ${record.workHours.toFixed(2)}h`);
      }
    });
    
    const totalWorkHoursInt = Math.floor(totalWorkHours);

    Logger.log(`⏱️ 總工作時數: ${totalWorkHoursInt}h`);
    
    // 4. 計算基本薪資（工作時數 × 時薪）
    const basePay = totalWorkHours * hourlyRate;
    
    Logger.log(`💰 基本薪資 = ${hourlyRate} × ${totalWorkHours.toFixed(2)} = $${Math.round(basePay)}`);
    
    // 5. ⭐ 取得加班記錄
    const overtimeRecords = getEmployeeMonthlyOvertime(employeeId, yearMonth);
    Logger.log(`📋 找到 ${overtimeRecords.length} 筆加班記錄`);
    
    // 6. ⭐⭐⭐ 計算加班費（區分平日/休息日/例假日）
    let totalOvertimeHours = 0;
    let weekdayOvertimePay = 0;   // 平日加班費
    let restdayOvertimePay = 0;   // 休息日加班費（週六）
    let holidayOvertimePay = 0;   // 例假日加班費（週日）
    
    // 按日期分組計算
    const overtimeByDate = {};
    
    overtimeRecords.forEach(record => {
      const date = record.date;
      if (!overtimeByDate[date]) {
        overtimeByDate[date] = 0;
      }
      overtimeByDate[date] += parseFloat(record.hours) || 0;
    });
    
    // ⭐⭐⭐ 修正：遍歷每天的加班記錄（區分平日/休息日/例假日）
    Object.keys(overtimeByDate).forEach(date => {
      let dailyHours = overtimeByDate[date];
      
      // 判斷日期類型
      const dateType = getDateType(date);
      const dateTypeName = {
        'weekday': '平日',
        'restday': '休息日（週六）',
        'holiday': '例假日（週日）'
      }[dateType];
      
      Logger.log(`\n📅 ${date} (${dateTypeName}): ${dailyHours.toFixed(1)}h`);
      
      // ⭐ 根據日期類型限制加班時數
      let maxHours = 4; // 平日最多4h
      if (dateType === 'restday') maxHours = 12; // 休息日最多12h
      if (dateType === 'holiday') maxHours = 8;  // 例假日最多8h
      
      if (dailyHours > maxHours) {
        Logger.log(`   ⚠️ 超過上限 (${dailyHours}h > ${maxHours}h)，限制為 ${maxHours}h`);
        dailyHours = maxHours;
      }
      
      // ⭐ 計算加班費
      const pay = calculateOvertimePay(dailyHours, hourlyRate, dateType);
      const totalPay = pay.firstPay + pay.secondPay + pay.thirdPay;
      
      // ⭐⭐⭐ 關鍵：依日期類型分別累計
      if (dateType === 'weekday') {
        weekdayOvertimePay += totalPay;
        Logger.log(`   - 前2h: $${pay.firstPay} (×1.34)`);
        if (pay.secondPay > 0) {
          Logger.log(`   - 後2h: $${pay.secondPay} (×1.67)`);
        }
      } else if (dateType === 'restday') {
        restdayOvertimePay += totalPay;
        Logger.log(`   - 前2h: $${pay.firstPay} (×1.34)`);
        if (pay.secondPay > 0) {
          Logger.log(`   - 3-8h: $${pay.secondPay} (×1.67)`);
        }
        if (pay.thirdPay > 0) {
          Logger.log(`   - 9h起: $${pay.thirdPay} (×2.67)`);
        }
      } else if (dateType === 'holiday') {
        holidayOvertimePay += totalPay;
        Logger.log(`   - 全天: $${totalPay} (×2.0)`);
      }
      
      totalOvertimeHours += dailyHours;
      Logger.log(`   ✅ 小計: $${totalPay}`);
    });
    
    // 四捨五入
    weekdayOvertimePay = Math.round(weekdayOvertimePay);
    restdayOvertimePay = Math.round(restdayOvertimePay);
    holidayOvertimePay = Math.round(holidayOvertimePay);
    
    Logger.log(`\n✅ 加班費計算完成:`);
    Logger.log(`   - 總時數: ${totalOvertimeHours.toFixed(1)}h`);
    Logger.log(`   - 平日加班費: $${weekdayOvertimePay}`);
    Logger.log(`   - 休息日加班費: $${restdayOvertimePay}`);
    Logger.log(`   - 例假日加班費: $${holidayOvertimePay}`);
    
    // 7. 固定津貼（時薪員工通常沒有，但保留欄位）
    const positionAllowance = parseFloat(config['職務加給']) || 0;
    const mealAllowance = parseFloat(config['伙食費']) || 0;
    const transportAllowance = parseFloat(config['交通補助']) || 0;
    const attendanceBonus = parseFloat(config['全勤獎金']) || 0;
    const performanceBonus = parseFloat(config['績效獎金']) || 0;
    const otherAllowance1 = parseFloat(config['其他津貼1']) || 0;  // ⭐ 改名
    const otherAllowance2 = parseFloat(config['其他津貼2']) || 0;  // ⭐ 改名
    const otherAllowance3 = parseFloat(config['其他津貼3']) || 0;  // ⭐ 改名
    
    Logger.log(`📋 固定津貼:`);
    if (positionAllowance > 0) Logger.log(`   - 職務加給: $${positionAllowance}`);
    if (mealAllowance > 0) Logger.log(`   - 伙食費: $${mealAllowance}`);
    if (transportAllowance > 0) Logger.log(`   - 交通補助: $${transportAllowance}`);
    if (attendanceBonus > 0) Logger.log(`   - 全勤獎金: $${attendanceBonus}`);
    if (performanceBonus > 0) Logger.log(`   - 績效獎金: $${performanceBonus}`);
    if (otherAllowance1 > 0) Logger.log(`   - 其他津貼1: $${otherAllowance1}`);  // ⭐ 改名
    if (otherAllowance2 > 0) Logger.log(`   - 其他津貼2: $${otherAllowance2}`);  // ⭐ 改名
    if (otherAllowance3 > 0) Logger.log(`   - 其他津貼3: $${otherAllowance3}`);  // ⭐ 改名


    // 8. 應發總額
    const grossSalary = basePay + 
                       positionAllowance + 
                       mealAllowance + 
                       transportAllowance + 
                       attendanceBonus + 
                       performanceBonus + 
                       otherAllowance1 +   // ⭐ 改名
                       otherAllowance2 +   // ⭐ 改名
                       otherAllowance3 +   // ⭐ 改名
                       weekdayOvertimePay + 
                       restdayOvertimePay +
                       holidayOvertimePay;
    
    Logger.log(`💵 應發總額: $${Math.round(grossSalary)}`);
    
    // 9. 扣款項目（時薪若月薪未達基本工資，可能不需扣保險）
    let laborFee = 0;
    let healthFee = 0;
    let employmentFee = 0;
    let pensionSelf = 0;
    let incomeTax = 0;
    
    // ⭐ 如果月總薪資達到基本工資，才扣保險
    if (grossSalary >= 28590) {
      const insuredSalary = getInsuredSalary(grossSalary);
      laborFee = Math.round(insuredSalary * 0.115 * 0.2);
      healthFee = Math.round(insuredSalary * 0.0517 * 0.3);
      employmentFee = Math.round(insuredSalary * 0.01 * 0.2);
      
      const pensionSelfRate = parseFloat(config['勞退自提率(%)']) || 0;
      pensionSelf = Math.round(insuredSalary * (pensionSelfRate / 100));
      
      if (grossSalary > 34000) {
        incomeTax = Math.round((grossSalary - 34000) * 0.05);
      }
      
      Logger.log(`📋 月薪達基本工資，計算法定扣款 (投保薪資: ${insuredSalary})`);
      Logger.log(`   - 勞保費: $${laborFee}`);
      Logger.log(`   - 健保費: $${healthFee}`);
      Logger.log(`   - 就業保險費: $${employmentFee}`);
      Logger.log(`   - 勞退自提 (${pensionSelfRate}%): $${pensionSelf}`);
      Logger.log(`   - 所得稅: $${incomeTax}`);
    } else {
      Logger.log(`⚠️ 月薪未達基本工資 ($${Math.round(grossSalary)} < $28,590)，不扣保險`);
    }
    
    // 10. 其他扣款
    const welfareFee = parseFloat(config['福利金扣款']) || 0;
    const dormitoryFee = parseFloat(config['宿舍費用']) || 0;
    const groupInsurance = parseFloat(config['團保費用']) || 0;
    const otherDeduction1 = parseFloat(config['其他扣款1']) || 0;  // ⭐ 改名
    const otherDeduction2 = parseFloat(config['其他扣款2']) || 0;  // ⭐ 改名
    if (welfareFee > 0 || dormitoryFee > 0 || groupInsurance > 0 || otherDeduction1 > 0 || otherDeduction2 > 0) {  // ✅ 修正
      Logger.log(`📋 其他扣款:`);
      if (welfareFee > 0) Logger.log(`   - 福利金: $${welfareFee}`);
      if (dormitoryFee > 0) Logger.log(`   - 宿舍費用: $${dormitoryFee}`);
      if (groupInsurance > 0) Logger.log(`   - 團保費用: $${groupInsurance}`);
      if (otherDeduction1 > 0) Logger.log(`   - 其他扣款1: $${otherDeduction1}`);
      if (otherDeduction2 > 0) Logger.log(`   - 其他扣款2: $${otherDeduction2}`);
    }
    
    // 11. 扣款總額
    const totalDeductions = laborFee + healthFee + employmentFee + pensionSelf + incomeTax +
                          welfareFee + dormitoryFee + groupInsurance + 
                          otherDeduction1 + otherDeduction2;  // ⭐ 改名
    
    Logger.log(`💸 扣款總額: $${totalDeductions}`);
    
    // 12. 實發金額
    const netSalary = grossSalary - totalDeductions;
    
    Logger.log('');
    Logger.log('═══════════════════════════════════════');
    Logger.log('📊 時薪薪資計算結果匯總:');
    Logger.log('═══════════════════════════════════════');
    Logger.log(`   員工: ${config['員工姓名']} (${employeeId})`);
    Logger.log(`   月份: ${yearMonth}`);
    Logger.log(`   時薪: $${hourlyRate}`);
    Logger.log(`   工作時數: ${totalWorkHours.toFixed(2)}h`);
    Logger.log(`   基本薪資: $${Math.round(basePay)}`);
    Logger.log(`   加班時數: ${totalOvertimeHours.toFixed(1)}h`);
    Logger.log(`   - 平日加班費: $${weekdayOvertimePay}`);
    Logger.log(`   - 休息日加班費: $${restdayOvertimePay}`);
    Logger.log(`   - 例假日加班費: $${holidayOvertimePay}`);
    Logger.log(`   應發總額: $${Math.round(grossSalary)}`);
    Logger.log(`   扣款總額: $${totalDeductions}`);
    Logger.log(`   實發金額: $${Math.round(netSalary)}`);
    Logger.log('═══════════════════════════════════════');
    Logger.log('');
    
    // 13. 返回結果
    const result = {
      employeeId: employeeId,
      employeeName: config['員工姓名'],
      yearMonth: yearMonth,
      salaryType: '時薪',
      hourlyRate: hourlyRate,
      totalWorkHours: totalWorkHoursInt,
      baseSalary: Math.round(basePay),
      positionAllowance: positionAllowance,
      mealAllowance: mealAllowance,
      transportAllowance: transportAllowance,
      attendanceBonus: attendanceBonus,
      performanceBonus: performanceBonus,
      otherAllowance1: otherAllowance1,
      otherAllowance2: otherAllowance2,
      otherAllowance3: otherAllowance3,
      weekdayOvertimePay: weekdayOvertimePay,      // ⭐ 只有平日
      restdayOvertimePay: restdayOvertimePay,      // ⭐ 只有休息日（週六）
      holidayOvertimePay: holidayOvertimePay,      // ⭐ 只有例假日（週日）
      totalOvertimeHours: totalOvertimeHours,
      laborFee: laborFee,
      healthFee: healthFee,
      employmentFee: employmentFee,
      pensionSelf: pensionSelf,
      pensionSelfRate: parseFloat(config['勞退自提率(%)']) || 0,
      incomeTax: incomeTax,
      leaveDeduction: 0,
      welfareFee: welfareFee,
      dormitoryFee: dormitoryFee,
      groupInsurance: groupInsurance,
      otherDeduction1: otherDeduction1,
      otherDeduction2: otherDeduction2,
      grossSalary: Math.round(grossSalary),
      netSalary: Math.round(netSalary),
      bankCode: config['銀行代碼'] || "",
      bankAccount: config['銀行帳號'] || "",
      status: "已計算",
      note: `工作${totalWorkHours.toFixed(1)}h，加班${totalOvertimeHours.toFixed(1)}h`
    };
    
    Logger.log('✅ 時薪計算完成');
    
    return { success: true, data: result };
    
  } catch (error) {
    Logger.log("❌ 計算時薪薪資失敗: " + error);
    Logger.log("❌ 錯誤堆疊: " + error.stack);
    return { success: false, message: error.toString() };
  }
}

/**
 * ✅ 最終修正版：取得員工月份打卡記錄
 * 
 * 根據實際 Sheet 結構（從截圖確認）：
 * 欄 A (索引0): 打卡時間 (Date 物件，例如：2026/1/2 上午 9:00:00)
 * 欄 B (索引1): userId
 * 欄 C (索引2): 部門
 * 欄 D (索引3): 打卡人員
 * 欄 E (索引4): 打卡類別 (上班/下班)
 * 欄 F (索引5): GPS位置
 * 欄 G (索引6): 地點
 * 欄 H (索引7): 備註
 * 欄 I (索引8): 管理員審核
 * 欄 J (索引9): 使用裝置詳細訊息
 */
function getEmployeeMonthlyAttendanceInternal(employeeId, yearMonth) {
  try {
    Logger.log('📋 取得員工月份打卡記錄（最終修正版）');
    Logger.log(`   員工ID: ${employeeId}`);
    Logger.log(`   年月: ${yearMonth}`);
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('打卡紀錄');
    
    if (!sheet) {
      Logger.log('❌ 找不到「打卡紀錄」工作表');
      return [];
    }
    
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      Logger.log('⚠️ 工作表無資料（只有標題列）');
      return [];
    }
    
    Logger.log(`✅ 讀取到 ${data.length - 1} 筆打卡記錄`);
    Logger.log('');
    
    // 按日期分組的打卡記錄
    const recordsByDate = {};
    
    // 從第 2 列開始處理（跳過標題）
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      // ⭐⭐⭐ 根據實際 Sheet 結構讀取欄位
      const punchTime = row[0];      // 欄 A: 打卡時間 (Date 物件)
      const recordUserId = String(row[1] || '').trim();  // 欄 B: userId
      const punchType = String(row[4] || '').trim();     // 欄 E: 打卡類別
      
      // 篩選條件 1: 員工ID 必須匹配
      if (recordUserId !== employeeId) {
        continue;
      }
      
      // 篩選條件 2: 打卡時間必須是 Date 物件
      if (!(punchTime instanceof Date)) {
        Logger.log(`⚠️ 第 ${i + 1} 列的打卡時間不是 Date 物件，跳過`);
        continue;
      }
      
      // 格式化日期為 YYYY-MM-DD 和 YYYY-MM
      const dateStr = Utilities.formatDate(punchTime, 'Asia/Taipei', 'yyyy-MM-dd');
      const recordYearMonth = Utilities.formatDate(punchTime, 'Asia/Taipei', 'yyyy-MM');
      
      // 篩選條件 3: 必須是目標月份
      if (recordYearMonth !== yearMonth) {
        continue;
      }
      
      // 格式化時間為 HH:mm
      const timeStr = Utilities.formatDate(punchTime, 'Asia/Taipei', 'HH:mm');
      
      Logger.log(`✅ 第 ${i + 1} 列符合: ${dateStr} ${timeStr} (${punchType})`);
      
      // 按日期分組
      if (!recordsByDate[dateStr]) {
        recordsByDate[dateStr] = {
          date: dateStr,
          punchIn: null,
          punchOut: null,
          records: []
        };
      }
      
      // 記錄上班/下班時間
      if (punchType === '上班') {
        // 如果有多筆上班記錄，取最早的
        if (!recordsByDate[dateStr].punchIn) {
          recordsByDate[dateStr].punchIn = timeStr;
        }
      } else if (punchType === '下班') {
        // 如果有多筆下班記錄，取最晚的
        recordsByDate[dateStr].punchOut = timeStr;
      }
      
      recordsByDate[dateStr].records.push({
        time: timeStr,
        type: punchType
      });
    }
    
    Logger.log('');
    Logger.log('🔢 開始計算工作時數...');
    Logger.log('');
    
    // 轉換為陣列並計算工作時數
    const results = [];
    
    for (const dateStr in recordsByDate) {
      const dayRecord = recordsByDate[dateStr];
      
      let workHours = 0;
      
      // 如果有上班和下班時間，計算工時
      if (dayRecord.punchIn && dayRecord.punchOut) {
        const inTime = new Date(`${dateStr} ${dayRecord.punchIn}`);
        const outTime = new Date(`${dateStr} ${dayRecord.punchOut}`);
        
        const diffMs = outTime - inTime;
        const totalHours = diffMs / (1000 * 60 * 60);
        
        // 扣除午休 1 小時
        workHours = Math.max(0, totalHours - 1);
        
        Logger.log(`   ${dateStr}:`);
        Logger.log(`      上班: ${dayRecord.punchIn}`);
        Logger.log(`      下班: ${dayRecord.punchOut}`);
        Logger.log(`      總時: ${totalHours.toFixed(2)}h`);
        Logger.log(`      工時: ${workHours.toFixed(2)}h (扣除午休1h)`);
      } else {
        Logger.log(`   ${dateStr}: ⚠️ 缺少上班或下班記錄`);
        if (dayRecord.punchIn) Logger.log(`      有上班: ${dayRecord.punchIn}`);
        if (dayRecord.punchOut) Logger.log(`      有下班: ${dayRecord.punchOut}`);
      }
      
      results.push({
        date: dateStr,
        punchIn: dayRecord.punchIn,
        punchOut: dayRecord.punchOut,
        workHours: workHours
      });
    }
    
    // 按日期排序
    results.sort((a, b) => a.date.localeCompare(b.date));
    
    Logger.log('');
    Logger.log('📊 統計結果:');
    Logger.log(`   找到 ${results.length} 天的打卡記錄`);
    
    const totalHours = results.reduce((sum, r) => sum + r.workHours, 0);
    Logger.log(`   總工作時數: ${totalHours.toFixed(2)}h`);
    
    return results;
    
  } catch (error) {
    Logger.log('❌ 取得打卡記錄失敗: ' + error.message);
    Logger.log('   錯誤堆疊: ' + error.stack);
    return [];
  }
}

/**
 * 🧪 測試函數
 */
function testGetAttendanceFinal() {
  const employeeId = 'Uf664a35632b736301d674d8b2cc3f8c0';
  const yearMonth = '2026-01';
  
  Logger.log('🧪 測試打卡記錄讀取（最終版）');
  Logger.log('='.repeat(50));
  Logger.log('');
  
  const records = getEmployeeMonthlyAttendanceInternal(employeeId, yearMonth);
  
  Logger.log('');
  Logger.log('='.repeat(50));
  Logger.log('✅ 測試完成！');
  Logger.log('');
  Logger.log('📋 返回的記錄:');
  Logger.log(JSON.stringify(records, null, 2));
}
/**
 * ✅ 取得員工該月份的打卡記錄並計算工時（修正版）
 * 
 * @param {string} employeeId - 員工ID
 * @param {string} yearMonth - 年月 (YYYY-MM)
 * @returns {Array} 打卡記錄陣列
 */
// function getEmployeeMonthlyAttendanceInternal(employeeId, yearMonth) {
//   try {
//     Logger.log('📋 開始取得員工打卡記錄');
//     Logger.log('   員工ID: ' + employeeId);
//     Logger.log('   年月: ' + yearMonth);
    
//     const ss = SpreadsheetApp.getActiveSpreadsheet();
//     const sheet = ss.getSheetByName(SHEET_ATTENDANCE);
    
//     if (!sheet) {
//       Logger.log('⚠️ 找不到「打卡紀錄」工作表');
//       return [];
//     }
    
//     const data = sheet.getDataRange().getValues();
    
//     if (data.length < 2) {
//       Logger.log('⚠️ 「打卡紀錄」工作表無資料');
//       return [];
//     }
    
//     const headers = data[0];
//     Logger.log('📊 打卡紀錄欄位: ' + headers.join(', '));
    
//     const punchTimeIndex = headers.indexOf('打卡時間');
//     const userIdIndex = headers.indexOf('userId');
//     const typeIndex = headers.indexOf('打卡類別');
//     const noteIndex = headers.indexOf('備註');
//     const auditIndex = headers.indexOf('管理員審核');
    
//     Logger.log('🔍 欄位索引:');
//     Logger.log('   打卡時間: ' + punchTimeIndex);
//     Logger.log('   userId: ' + userIdIndex);
//     Logger.log('   打卡類別: ' + typeIndex);
    
//     if (punchTimeIndex === -1 || userIdIndex === -1 || typeIndex === -1) {
//       Logger.log('⚠️ 「打卡紀錄」工作表缺少必要欄位');
//       return [];
//     }
    
//     // ⭐ 按日期分組打卡記錄（改用陣列儲存所有打卡）
//     const recordsByDate = {};
    
//     for (let i = 1; i < data.length; i++) {
//       const row = data[i];
      
//       const rowUserId = String(row[userIdIndex] || '').trim();
//       const punchTime = row[punchTimeIndex];
//       const punchType = String(row[typeIndex] || '').trim();
//       const note = row[noteIndex] || '';
//       const audit = row[auditIndex] || '';
      
//       if (rowUserId !== employeeId) continue;
      
//       // 解析打卡時間
//       let punchDate = null;
//       let timeStr = '';
//       let fullDateTime = null;
      
//       if (punchTime instanceof Date) {
//         punchDate = Utilities.formatDate(punchTime, 'Asia/Taipei', 'yyyy-MM-dd');
//         timeStr = Utilities.formatDate(punchTime, 'Asia/Taipei', 'HH:mm');
//         fullDateTime = punchTime;
//       } else if (typeof punchTime === 'string') {
//         const parts = punchTime.split(' ');
//         if (parts.length >= 2) {
//           punchDate = parts[0];
//           timeStr = parts[1].substring(0, 5);
//           try {
//             fullDateTime = new Date(punchTime);
//           } catch (e) {
//             continue;
//           }
//         }
//       } else {
//         continue;
//       }
      
//       const dateStr = punchDate.substring(0, 7);
//       if (dateStr !== yearMonth) continue;
      
//       // 只計算正常打卡或已核准的補打卡
//       const isNormalPunch = (note !== '補打卡');
//       const isApprovedAdjustment = (note === '補打卡' && audit === 'v');
      
//       if (!isNormalPunch && !isApprovedAdjustment) {
//         Logger.log(`   ⏭️ 跳過 ${punchDate} ${timeStr} 的未核准補打卡`);
//         continue;
//       }
      
//       // ⭐ 改用陣列儲存所有打卡（支援同一天多次打卡）
//       if (!recordsByDate[punchDate]) {
//         recordsByDate[punchDate] = [];
//       }
      
//       recordsByDate[punchDate].push({
//         type: punchType,
//         time: timeStr,
//         fullDateTime: fullDateTime,
//         note: note
//       });
//     }
    
//     Logger.log(`📊 找到 ${Object.keys(recordsByDate).length} 天的打卡記錄`);
    
//     // ⭐⭐⭐ 關鍵修正：配對上下班記錄並計算工時
//     const records = [];
    
//     Object.keys(recordsByDate).forEach(date => {
//       const dayPunches = recordsByDate[date];
      
//       // 按時間排序
//       dayPunches.sort((a, b) => a.fullDateTime - b.fullDateTime);
      
//       // 找出上班和下班打卡
//       const punchIns = dayPunches.filter(p => p.type === '上班');
//       const punchOuts = dayPunches.filter(p => p.type === '下班');
      
//       let punchIn = null;
//       let punchOut = null;
//       let workHours = 0;
      
//       // ⭐ 配對邏輯：取第一個上班和最後一個下班
//       if (punchIns.length > 0) {
//         punchIn = punchIns[0].time;
//       }
      
//       if (punchOuts.length > 0) {
//         punchOut = punchOuts[punchOuts.length - 1].time;
//       }
      
//       // 計算工時
//       if (punchIn && punchOut) {
//         try {
//           const inTime = new Date(`${date} ${punchIn}`);
//           const outTime = new Date(`${date} ${punchOut}`);
//           const diffMs = outTime - inTime;
          
//           if (diffMs > 0) {
//             const totalHours = diffMs / (1000 * 60 * 60);
//             const lunchBreak = 1;
//             // workHours = Math.max(0, totalHours - lunchBreak);
//             workHours = Math.floor(Math.max(0, totalHours - lunchBreak));
//             Logger.log(`   ${date}: ${punchIn} ~ ${punchOut} = ${workHours.toFixed(2)}h (原始: ${totalHours.toFixed(2)}h)`);
//           } else {
//             Logger.log(`   ⚠️ ${date}: ${punchIn} ~ ${punchOut} 時間異常（下班早於上班）`);
//           }
//         } catch (e) {
//           Logger.log(`   ⚠️ 無法計算 ${date} 的工時: ` + e);
//         }
//       } else {
//         Logger.log(`   ⚠️ ${date}: 打卡不完整 (上班: ${punchIn || '無'}, 下班: ${punchOut || '無'})`);
//       }
      
//       records.push({
//         date: date,
//         punchIn: punchIn,
//         punchOut: punchOut,
//         workHours: workHours
//       });
//     });
    
//     // 按日期排序
//     records.sort((a, b) => a.date.localeCompare(b.date));
    
//     Logger.log(`✅ 成功處理 ${records.length} 筆打卡記錄`);
    
//     return records;
    
//   } catch (error) {
//     Logger.log('❌ 取得打卡記錄失敗: ' + error);
//     Logger.log('❌ 錯誤堆疊: ' + error.stack);
//     return [];
//   }
// }

/**
 * ✅ 新增 API：取得員工該月份的加班記錄
 */
function getEmployeeMonthlyOvertimeAPI() {
  try {
    const session = checkSessionInternal();
    if (!session.ok) {
      return jsonResponse({ ok: false, msg: 'SESSION_INVALID', code: 'SESSION_INVALID' });
    }
    
    const employeeId = session.user.userId;
    const yearMonth = getParam('yearMonth');
    
    if (!yearMonth) {
      return jsonResponse({ ok: false, msg: 'MISSING_YEAR_MONTH', code: 'MISSING_YEAR_MONTH' });
    }
    
    Logger.log(`📋 取得 ${employeeId} 在 ${yearMonth} 的加班記錄`);
    
    const records = getEmployeeMonthlyOvertime(employeeId, yearMonth);
    
    return jsonResponse({ ok: true, records: records });
    
  } catch (error) {
    Logger.log('❌ getEmployeeMonthlyOvertimeAPI 錯誤: ' + error);
    return jsonResponse({ ok: false, msg: error.toString(), code: 'ERROR' });
  }
}

/**
 * 🧪 測試打卡工時計算
 */
function testGetEmployeeMonthlyAttendance() {
  Logger.log('═══════════════════════════════════════');
  Logger.log('🧪 測試 getEmployeeMonthlyAttendance');
  Logger.log('═══════════════════════════════════════');
  Logger.log('');
  
  const employeeId = 'U68e0ca9d516e63ed15bf9387fad174ac'; // CSF
  const yearMonth = '2025-12';
  
  const records = getEmployeeMonthlyAttendanceInternal(employeeId, yearMonth);
  
  Logger.log('');
  Logger.log('📊 測試結果：找到 ' + records.length + ' 筆記錄');
  Logger.log('');
  
  let totalHours = 0;
  
  records.forEach(record => {
    Logger.log(`   ${record.date}: ${record.punchIn || '--'} ~ ${record.punchOut || '--'}, 工時: ${record.workHours.toFixed(2)}h`);
    totalHours += record.workHours;
  });
  
  Logger.log('');
  Logger.log('✅ 總工時: ' + totalHours.toFixed(2) + ' 小時');
  Logger.log('');
  Logger.log('═══════════════════════════════════════');
}

/**
 * ✅ 計算午休時間（12:00-13:00）
 * 
 * @param {Date} startTime - 上班時間
 * @param {Date} endTime - 下班時間
 * @returns {number} 午休時間（毫秒）
 */
function calculateLunchBreak(startTime, endTime) {
  const lunchStart = new Date(startTime);
  lunchStart.setHours(12, 0, 0, 0);
  
  const lunchEnd = new Date(startTime);
  lunchEnd.setHours(13, 0, 0, 0);
  
  // 如果工作時段包含午休時間，扣除1小時
  if (startTime < lunchEnd && endTime > lunchStart) {
    return 60 * 60 * 1000; // 1小時 = 3600000毫秒
  }
  
  return 0;
}

/**
 * ✅ 投保薪資級距對照表（供時薪使用）
 */
function getInsuredSalary(salary) {
  const brackets = [
    { min: 0, max: 26400, insured: 26400 },
    { min: 26401, max: 27600, insured: 27600 },
    { min: 27601, max: 28800, insured: 28800 },
    { min: 28801, max: 30300, insured: 30300 },
    { min: 30301, max: 31800, insured: 31800 },
    { min: 31801, max: 33300, insured: 33300 },
    { min: 33301, max: 34800, insured: 34800 },
    { min: 34801, max: 36300, insured: 36300 },
    { min: 36301, max: 38200, insured: 38200 },
    { min: 38201, max: 40100, insured: 40100 },
    { min: 40101, max: 42000, insured: 42000 },
    { min: 42001, max: 43900, insured: 43900 },
    { min: 43901, max: 45800, insured: 45800 },
    { min: 45801, max: Infinity, insured: 45800 }
  ];
  
  for (const bracket of brackets) {
    if (salary >= bracket.min && salary <= bracket.max) {
      return bracket.insured;
    }
  }
  
  return 26400;
}

// ==================================================================================
// 薪資管理系統 - 情況一完整版
// ==================================================================================
// 
// ✅ 完整支援情況一（飼料廠司機 + 不定時工作）：
//    1. 基本薪資：投保薪資底薪
//    2. 津貼：職務加給、伙食津貼、交通津貼
//    3. 加班費：固定 $200/小時（平日1.34倍、休息日1.34/1.67/2.67倍）
//    4. 特休未休補薪：按日薪計算
//    5. 月休補薪：應休6天，一天沒休補1,500元
//    6. 請假扣款：病假半薪、事假全薪（小時制）
//
// ==================================================================================

/**
 * ✅ 計算月薪資（完整版 - 支援情況一 v3.0）
 */
function calculateDriverSalaryCase1(employeeId, yearMonth) {
  try {
    Logger.log('═══════════════════════════════════════');
    Logger.log('🚛 開始計算飼料廠司機薪資（情況一）');
    Logger.log('   員工ID: ' + employeeId);
    Logger.log('   年月: ' + yearMonth);
    Logger.log('═══════════════════════════════════════');
    
    // 步驟 1：取得員工薪資設定
    const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('員工薪資設定');
    if (!configSheet) {
      return { ok: false, msg: "找不到員工薪資設定工作表" };
    }
    
    const configData = configSheet.getDataRange().getValues();
    let employeeConfig = null;
    
    for (let i = 1; i < configData.length; i++) {
      if (configData[i][0] === employeeId) {
        employeeConfig = configData[i];
        break;
      }
    }
    
    if (!employeeConfig) {
      return { ok: false, msg: "找不到該員工的薪資設定" };
    }
    
    Logger.log('✅ 找到員工薪資設定');
    
    // 步驟 2：解析員工薪資設定
    const employeeName = employeeConfig[1] || '';
    const employeeType = employeeConfig[3] || '正職';
    const salaryType = employeeConfig[4] || '月薪';
    const workTimeType = employeeConfig[5] || '標準工時';
    const baseSalary = parseFloat(employeeConfig[6]) || 0;
    
    // 固定津貼
    const positionAllowance = parseFloat(employeeConfig[7]) || 0;
    const mealAllowance = parseFloat(employeeConfig[8]) || 0;
    const transportAllowance = parseFloat(employeeConfig[9]) || 0;
    const attendanceBonus = parseFloat(employeeConfig[10]) || 0;
    const performanceBonus = parseFloat(employeeConfig[11]) || 0;
    const otherAllowance1 = parseFloat(employeeConfig[12]) || 0;
    const otherAllowance2 = parseFloat(employeeConfig[13]) || 0;
    const otherAllowance3 = parseFloat(employeeConfig[14]) || 0;
    
    // 銀行資訊
    const bankCode = employeeConfig[15] || '';
    const bankAccount = employeeConfig[16] || '';
    
    // 法定扣款
    const pensionSelfRate = parseFloat(employeeConfig[19]) || 0;
    const laborFee = parseFloat(employeeConfig[20]) || 0;
    const healthFee = parseFloat(employeeConfig[21]) || 0;
    const employmentFee = parseFloat(employeeConfig[22]) || 0;
    const pensionSelf = parseFloat(employeeConfig[23]) || 0;
    const incomeTax = parseFloat(employeeConfig[24]) || 0;
    
    // 其他扣款
    const welfareFee = parseFloat(employeeConfig[25]) || 0;
    const dormitoryFee = parseFloat(employeeConfig[26]) || 0;
    const groupInsurance = parseFloat(employeeConfig[27]) || 0;
    const otherDeduction1 = parseFloat(employeeConfig[28]) || 0;
    const otherDeduction2 = parseFloat(employeeConfig[29]) || 0;
    
    Logger.log('📋 基本薪資資料:');
    Logger.log(`   員工姓名: ${employeeName}`);
    Logger.log(`   員工類型: ${employeeType}`);
    Logger.log(`   薪資類型: ${salaryType}`);
    Logger.log(`   工時類型: ${workTimeType}`);
    Logger.log(`   基本薪資: ${baseSalary}`);
    Logger.log('');
    
    // 步驟 3：計算固定津貼總額
    const totalFixedAllowance = 
      positionAllowance + mealAllowance + transportAllowance +
      attendanceBonus + performanceBonus +
      otherAllowance1 + otherAllowance2 + otherAllowance3;
    
    Logger.log(`💰 固定津貼總額: ${totalFixedAllowance}`);
    Logger.log('');
    
    // ⭐⭐⭐ 步驟 4：計算加班費（修正版 - 支援情況一固定費率）
    let weekdayOvertimePay = 0;
    let restdayOvertimePay = 0;
    let holidayOvertimePay = 0;
    let totalOvertimeHours = 0;
    
    if (workTimeType === '不定時') {
      // ⭐⭐⭐ 關鍵判斷：情況一（飼料廠司機）使用固定費率
      if (employeeType === '飼料廠司機') {
        Logger.log('💼 情況一：飼料廠司機 + 不定時');
        Logger.log('   使用固定加班費率：$200/小時');
        
        const overtimeResult = calculateMonthlyOvertimePay(
          employeeId, 
          yearMonth, 
          baseSalary,
          true,   // ⭐ 使用固定費率
          200     // ⭐ $200/小時
        );
        
        weekdayOvertimePay = overtimeResult.weekdayOvertimePay;
        restdayOvertimePay = overtimeResult.restdayOvertimePay;
        holidayOvertimePay = overtimeResult.holidayOvertimePay;
        totalOvertimeHours = overtimeResult.totalOvertimeHours;
        
        Logger.log('💼 加班費計算結果:');
        Logger.log(`   平日加班費: ${weekdayOvertimePay}`);
        Logger.log(`   休息日加班費: ${restdayOvertimePay}`);
        Logger.log(`   例假日加班費: ${holidayOvertimePay}`);
        Logger.log(`   總加班時數: ${totalOvertimeHours}`);
      } else {
        // 其他不定時員工：跳過加班費
        Logger.log('⏭️ 不定時工作（非飼料廠司機），跳過加班費計算');
      }
    } else {
      // 標準工時：使用計算時薪（基本薪資/30/8）
      Logger.log('💼 標準工時，使用一般加班費計算');
      
      const overtimeResult = calculateMonthlyOvertimePay(
        employeeId, 
        yearMonth, 
        baseSalary,
        false   // 不使用固定費率
      );
      
      weekdayOvertimePay = overtimeResult.weekdayOvertimePay;
      restdayOvertimePay = overtimeResult.restdayOvertimePay;
      holidayOvertimePay = overtimeResult.holidayOvertimePay;
      totalOvertimeHours = overtimeResult.totalOvertimeHours;
      
      Logger.log('💼 加班費計算結果:');
      Logger.log(`   平日加班費: ${weekdayOvertimePay}`);
      Logger.log(`   休息日加班費: ${restdayOvertimePay}`);
      Logger.log(`   例假日加班費: ${holidayOvertimePay}`);
      Logger.log(`   總加班時數: ${totalOvertimeHours}`);
    }
    Logger.log('');
    
    // 步驟 5：計算未休假補薪（特休）
    const unusedLeavePayResult = calculateUnusedLeavePay(employeeId, yearMonth, baseSalary, totalFixedAllowance);
    const unusedLeavePay = unusedLeavePayResult.amount;
    const unusedLeaveDays = unusedLeavePayResult.days;
    
    Logger.log('🏖️ 未休假補薪（特休）:');
    Logger.log(`   未休天數: ${unusedLeaveDays} 天`);
    Logger.log(`   補薪金額: ${unusedLeavePay} 元`);
    Logger.log('');
    
    // ⭐⭐⭐ 步驟 5.5：計算月休補薪（情況一專用）
    const monthlyRestPayResult = calculateMonthlyRestDayPay(employeeId, yearMonth, employeeType);
    const monthlyRestPay = monthlyRestPayResult.payAmount;
    const monthlyRestMissedDays = monthlyRestPayResult.missedDays;
    const monthlyRestRequiredDays = monthlyRestPayResult.requiredDays;
    const monthlyRestActualDays = monthlyRestPayResult.actualDays;
    
    Logger.log('📅 月休補薪（情況一專用）:');
    Logger.log(`   應休天數: ${monthlyRestRequiredDays} 天`);
    Logger.log(`   實際休假: ${monthlyRestActualDays} 天`);
    Logger.log(`   未休天數: ${monthlyRestMissedDays} 天`);
    Logger.log(`   補薪金額: ${monthlyRestPay} 元`);
    Logger.log('');
    
    // 步驟 6：計算請假扣款（小時制）
    const leaveDeductionResult = calculateLeaveDeduction(
      employeeId, 
      yearMonth, 
      baseSalary, 
      mealAllowance,        // ✅ 伙食費
      transportAllowance    // ✅ 交通津貼
    );
    const sickLeaveDeduction = leaveDeductionResult.sickLeaveDeduction;
    const personalLeaveDeduction = leaveDeductionResult.personalLeaveDeduction;
    const totalLeaveDeduction = leaveDeductionResult.totalDeduction;
    const sickLeaveHours = leaveDeductionResult.sickLeaveHours;
    const personalLeaveHours = leaveDeductionResult.personalLeaveHours;
    
    Logger.log('🏥 請假扣款（小時制）:');
    Logger.log(`   病假: ${sickLeaveHours} 小時 → 扣款 ${sickLeaveDeduction} 元`);
    Logger.log(`   事假: ${personalLeaveHours} 小時 → 扣款 ${personalLeaveDeduction} 元`);
    Logger.log(`   扣款總計: ${totalLeaveDeduction} 元`);
    Logger.log('');
    
    // 步驟 7：計算應發總額（⭐ 加入月休補薪）
    const grossSalary = 
      baseSalary + 
      totalFixedAllowance + 
      weekdayOvertimePay + 
      restdayOvertimePay + 
      holidayOvertimePay +
      unusedLeavePay +      // 特休未休補薪
      monthlyRestPay;       // ⭐⭐⭐ 月休補薪
    
    Logger.log(`💵 應發總額: ${grossSalary}`);
    Logger.log('');
    
    // 步驟 8：計算扣款總額
    const totalDeductions = 
      laborFee + 
      healthFee + 
      employmentFee + 
      pensionSelf + 
      incomeTax + 
      totalLeaveDeduction +
      welfareFee +
      dormitoryFee +
      groupInsurance +
      otherDeduction1 +
      otherDeduction2;
    
    Logger.log(`💸 扣款總額: ${totalDeductions}`);
    Logger.log('');
    
    // 步驟 9：計算實發金額
    const netSalary = grossSalary - totalDeductions;
    
    Logger.log(`✅ 實發金額: ${netSalary}`);
    Logger.log('');
    
    // 步驟 10：返回完整薪資資料
    const salaryData = {
      employeeId: employeeId,
      employeeName: employeeName,
      yearMonth: yearMonth,
      employeeType: employeeType,
      salaryType: salaryType,
      workTimeType: workTimeType,
      
      // 基本薪資
      baseSalary: baseSalary,
      
      // 固定津貼
      positionAllowance: positionAllowance,
      mealAllowance: mealAllowance,
      transportAllowance: transportAllowance,
      attendanceBonus: attendanceBonus,
      performanceBonus: performanceBonus,
      otherAllowance1: otherAllowance1,
      otherAllowance2: otherAllowance2,
      otherAllowance3: otherAllowance3,
      
      // 補薪項目
      unusedLeavePay: unusedLeavePay,
      unusedLeaveDays: unusedLeaveDays,
      monthlyRestPay: monthlyRestPay,              // ⭐⭐⭐ 月休補薪
      monthlyRestMissedDays: monthlyRestMissedDays,
      monthlyRestRequiredDays: monthlyRestRequiredDays,
      monthlyRestActualDays: monthlyRestActualDays,
      
      // 加班費
      weekdayOvertimePay: weekdayOvertimePay,
      restdayOvertimePay: restdayOvertimePay,
      holidayOvertimePay: holidayOvertimePay,
      totalOvertimeHours: totalOvertimeHours,
      
      // 法定扣款
      laborFee: laborFee,
      healthFee: healthFee,
      employmentFee: employmentFee,
      pensionSelfRate: pensionSelfRate,
      pensionSelf: pensionSelf,
      incomeTax: incomeTax,
      
      // 請假扣款明細
      sickLeaveHours: sickLeaveHours,
      sickLeaveDeduction: sickLeaveDeduction,
      personalLeaveHours: personalLeaveHours,
      personalLeaveDeduction: personalLeaveDeduction,
      leaveDeduction: totalLeaveDeduction,
      
      // 其他扣款
      welfareFee: welfareFee,
      dormitoryFee: dormitoryFee,
      groupInsurance: groupInsurance,
      otherDeduction1: otherDeduction1,
      otherDeduction2: otherDeduction2,
      
      // 總計
      grossSalary: grossSalary,
      totalDeductions: totalDeductions,
      netSalary: netSalary,
      
      // 銀行資訊
      bankCode: bankCode,
      bankAccount: bankAccount
    };
    
    Logger.log('═══════════════════════════════════════');
    Logger.log('✅✅✅ 薪資計算完成（情況一 v3.0）');
    Logger.log('═══════════════════════════════════════');
    Logger.log('📊 薪資結構摘要:');
    Logger.log(`   基本薪資: ${baseSalary}`);
    Logger.log(`   固定津貼: ${totalFixedAllowance}`);
    Logger.log(`   加班費: ${weekdayOvertimePay + restdayOvertimePay + holidayOvertimePay}`);
    Logger.log(`   特休未休補薪: ${unusedLeavePay}`);
    Logger.log(`   月休補薪: ${monthlyRestPay} ⭐`);
    Logger.log(`   應發總額: ${grossSalary}`);
    Logger.log(`   扣款總額: ${totalDeductions}`);
    Logger.log(`   實發金額: ${netSalary}`);
    Logger.log('═══════════════════════════════════════');
    
    return {
      ok: true,
      success: true,
      data: salaryData
    };
    
  } catch (error) {
    Logger.log('');
    Logger.log('❌❌❌ calculateMonthlySalary 發生錯誤');
    Logger.log('錯誤訊息: ' + error.message);
    Logger.log('錯誤堆疊: ' + error.stack);
    Logger.log('═══════════════════════════════════════');
    
    return {
      ok: false,
      success: false,
      msg: "計算失敗：" + error.message
    };
  }
}

/**
 * ✅ 月薪計算（內部函數 - 完整修正版）
 */
function calculateMonthlySalaryInternal(employeeId, yearMonth) {
  try {
    Logger.log(`💰 開始計算月薪: ${employeeId}, ${yearMonth}`);
    
    // 1. 取得員工薪資設定
    const salaryConfig = getEmployeeSalaryTW(employeeId);
    if (!salaryConfig.success) {
      return { success: false, message: "找不到員工薪資設定" };
    }
    
    const config = salaryConfig.data;
    
    // 2. 取得加班記錄
    const overtimeRecords = getEmployeeMonthlyOvertime(employeeId, yearMonth);
    Logger.log(`📋 找到 ${overtimeRecords.length} 筆加班記錄`);
    
    // 3. 取得請假記錄
    const leaveRecords = getEmployeeMonthlySalary(employeeId, yearMonth);
    
    // 4. 基本薪資
    const baseSalary = parseFloat(config['基本薪資']) || 0;
    const hourlyRate = Math.round(baseSalary / 30 / 8); // 平日時薪
    
    Logger.log(`💵 基本薪資: ${baseSalary}, 時薪: ${hourlyRate}`);
    
    // 5. 固定津貼
    const positionAllowance = parseFloat(config['職務加給']) || 0;
    const mealAllowance = parseFloat(config['伙食費']) || 0;
    const transportAllowance = parseFloat(config['交通補助']) || 0;
    let attendanceBonus = parseFloat(config['全勤獎金']) || 0;
    const performanceBonus = parseFloat(config['績效獎金']) || 0;
    const otherAllowance1 = parseFloat(config['其他津貼1']) || 0;  // ⭐ 改名
    const otherAllowance2 = parseFloat(config['其他津貼2']) || 0;  // ⭐ 改名
    const otherAllowance3 = parseFloat(config['其他津貼3']) || 0;  // ⭐ 改名
    
    // 6. ⭐⭐⭐ 計算加班費（區分平日/休息日/例假日）
    let totalOvertimeHours = 0;
    let weekdayOvertimePay = 0;   // 平日加班費
    let restdayOvertimePay = 0;   // 休息日加班費（週六）
    let holidayOvertimePay = 0;   // 例假日加班費（週日）
    
    // 按日期分組計算
    const overtimeByDate = {};
    
    overtimeRecords.forEach(record => {
      const date = record.date;
      if (!overtimeByDate[date]) {
        overtimeByDate[date] = 0;
      }
      overtimeByDate[date] += parseFloat(record.hours) || 0;
    });
    
    Logger.log(`📊 每日加班統計: ${JSON.stringify(overtimeByDate)}`);
    
    // ⭐⭐⭐ 修正：遍歷每天的加班記錄（區分平日/休息日/例假日）
    Object.keys(overtimeByDate).forEach(date => {
      let dailyHours = overtimeByDate[date];
      
      const dateType = getDateType(date);
      const dateTypeName = {
        'weekday': '平日',
        'restday': '休息日（週六）',
        'holiday': '例假日（週日）'
      }[dateType];
      
      Logger.log(`\n📅 ${date} (${dateTypeName}): ${dailyHours.toFixed(1)}h`);
      
      let maxHours = 4;
      if (dateType === 'restday') maxHours = 12;
      if (dateType === 'holiday') maxHours = 8;
      
      if (dailyHours > maxHours) {
        Logger.log(`   ⚠️ 超過上限，限制為 ${maxHours}h`);
        dailyHours = maxHours;
      }
      
      const pay = calculateOvertimePay(dailyHours, hourlyRate, dateType);
      const totalPay = pay.firstPay + pay.secondPay + pay.thirdPay;
      
      // ⭐⭐⭐ 關鍵：依日期類型分別累計
      if (dateType === 'weekday') {
        weekdayOvertimePay += totalPay;
        Logger.log(`   - 前2h: $${pay.firstPay} (×1.34)`);
        if (pay.secondPay > 0) {
          Logger.log(`   - 後2h: $${pay.secondPay} (×1.67)`);
        }
      } else if (dateType === 'restday') {
        restdayOvertimePay += totalPay;
        Logger.log(`   - 前2h: $${pay.firstPay} (×1.34)`);
        if (pay.secondPay > 0) {
          Logger.log(`   - 3-8h: $${pay.secondPay} (×1.67)`);
        }
        if (pay.thirdPay > 0) {
          Logger.log(`   - 9h起: $${pay.thirdPay} (×2.67)`);
        }
      } else if (dateType === 'holiday') {
        holidayOvertimePay += totalPay;
        Logger.log(`   - 全天: $${totalPay} (×2.0)`);
      }
      
      totalOvertimeHours += dailyHours;
      Logger.log(`   ✅ 小計: $${totalPay}`);
    });
    
    // 四捨五入
    weekdayOvertimePay = Math.round(weekdayOvertimePay);
    restdayOvertimePay = Math.round(restdayOvertimePay);
    holidayOvertimePay = Math.round(holidayOvertimePay);
    
    Logger.log(`\n✅ 加班費計算完成:`);
    Logger.log(`   - 總時數: ${totalOvertimeHours.toFixed(1)}h`);
    Logger.log(`   - 平日加班費: $${weekdayOvertimePay}`);
    Logger.log(`   - 休息日加班費: $${restdayOvertimePay}`);
    Logger.log(`   - 例假日加班費: $${holidayOvertimePay}`);
    
    // 7. 請假扣款（使用小時制 calculateLeaveDeduction）
    const leaveDeductionResult = calculateLeaveDeduction(
        employeeId,
        yearMonth,
        baseSalary,
        mealAllowance,
        transportAllowance
    );

    const sickLeaveHours       = leaveDeductionResult.sickLeaveHours;
    const sickLeaveDeduction   = leaveDeductionResult.sickLeaveDeduction;
    const personalLeaveHours   = leaveDeductionResult.personalLeaveHours;
    const personalLeaveDeduction = leaveDeductionResult.personalLeaveDeduction;
    const leaveDeduction       = leaveDeductionResult.totalDeduction;

    // 有請假則取消全勤獎金
    if (leaveDeduction > 0) {
        attendanceBonus = 0;
        Logger.log('⚠️ 有請假記錄，取消全勤獎金');
    }
    
    // 8. 法定扣款
    const laborFee = parseFloat(config['勞保費']) || 0;
    const healthFee = parseFloat(config['健保費']) || 0;
    const employmentFee = parseFloat(config['就業保險費']) || 0;
    const pensionSelf = parseFloat(config['勞退自提']) || 0;
    const pensionSelfRate = parseFloat(config['勞退自提率(%)']) || 0;
    const incomeTax = parseFloat(config['所得稅']) || 0;
    
    // 9. 其他扣款
    const welfareFee = parseFloat(config['福利金扣款']) || 0;
    const dormitoryFee = parseFloat(config['宿舍費用']) || 0;
    const groupInsurance = parseFloat(config['團保費用']) || 0;
    const otherDeduction1 = parseFloat(config['其他扣款1']) || 0;  // ⭐ 改名
    const otherDeduction2 = parseFloat(config['其他扣款2']) || 0;  // ⭐ 改名
    
    // 10. 應發總額
    const grossSalary = baseSalary + 
                       positionAllowance + 
                       mealAllowance + 
                       transportAllowance + 
                       attendanceBonus + 
                       performanceBonus + 
                       otherAllowance1 +  // ⭐ 改名
                       otherAllowance2 +  // ⭐ 改名
                       otherAllowance3 +  // ⭐ 改名
                       weekdayOvertimePay + 
                       restdayOvertimePay +
                       holidayOvertimePay;
    
    // 11. 扣款總額
    const totalDeductions = laborFee + 
                           healthFee + 
                           employmentFee + 
                           pensionSelf + 
                           incomeTax +
                           leaveDeduction + 
                           welfareFee + 
                           dormitoryFee + 
                           groupInsurance + 
                           otherDeduction1 +  // ⭐ 改名
                           otherDeduction2 +  // ⭐ 改名
                           0;  // 移除原本的 otherDeductions
    
    // 12. 實發金額
    const netSalary = grossSalary - totalDeductions;
    
    Logger.log('');
    Logger.log('═══════════════════════════════════════');
    Logger.log('📊 月薪薪資計算結果匯總:');
    Logger.log('═══════════════════════════════════════');
    Logger.log(`   員工: ${config['員工姓名']} (${employeeId})`);
    Logger.log(`   月份: ${yearMonth}`);
    Logger.log(`   基本薪資: $${baseSalary}`);
    Logger.log(`   加班時數: ${totalOvertimeHours.toFixed(1)}h`);
    Logger.log(`   - 平日加班費: $${weekdayOvertimePay}`);
    Logger.log(`   - 休息日加班費: $${restdayOvertimePay}`);
    Logger.log(`   - 例假日加班費: $${holidayOvertimePay}`);
    Logger.log(`   應發總額: $${Math.round(grossSalary)}`);
    Logger.log(`   扣款總額: $${totalDeductions}`);
    Logger.log(`   實發金額: $${Math.round(netSalary)}`);
    Logger.log('═══════════════════════════════════════');
    Logger.log('');
    
    // 13. 返回結果
    const result = {
      employeeId: employeeId,
      employeeName: config['員工姓名'],
      yearMonth: yearMonth,
      salaryType: '月薪',
      baseSalary: baseSalary,
      positionAllowance: positionAllowance,
      mealAllowance: mealAllowance,
      transportAllowance: transportAllowance,
      attendanceBonus: attendanceBonus,
      performanceBonus: performanceBonus,
      otherAllowance1: otherAllowance1,
      otherAllowance2: otherAllowance2,
      otherAllowance3: otherAllowance3,
      weekdayOvertimePay: weekdayOvertimePay,      // ⭐ 只有平日
      restdayOvertimePay: restdayOvertimePay,      // ⭐ 只有休息日（週六）
      holidayOvertimePay: holidayOvertimePay,      // ⭐ 只有例假日（週日）
      totalOvertimeHours: totalOvertimeHours,
      laborFee: laborFee,
      healthFee: healthFee,
      employmentFee: employmentFee,
      pensionSelf: pensionSelf,
      pensionSelfRate: pensionSelfRate,
      incomeTax: incomeTax,
      leaveDeduction: Math.round(leaveDeduction),
      // ⭐ 新增：請假明細
      sickLeaveHours:          sickLeaveHours,
      sickLeaveDeduction:      sickLeaveDeduction,
      personalLeaveHours:      personalLeaveHours,
      personalLeaveDeduction:  personalLeaveDeduction,
      welfareFee: welfareFee,
      dormitoryFee: dormitoryFee,
      groupInsurance: groupInsurance,
      otherDeduction1: otherDeduction1,
      otherDeduction2: otherDeduction2,
      grossSalary: Math.round(grossSalary),
      netSalary: Math.round(netSalary),
      bankCode: config['銀行代碼'] || "",
      bankAccount: config['銀行帳號'] || "",
      status: "已計算",
      note: `本月加班${totalOvertimeHours.toFixed(1)}小時`
    };
    
    Logger.log('✅ 月薪計算完成');
    
    return { success: true, data: result };
    
  } catch (error) {
    Logger.log("❌ 計算月薪失敗: " + error);
    Logger.log("❌ 錯誤堆疊: " + error.stack);
    return { success: false, message: error.toString() };
  }
}


/**
 * ✅ API：取得員工該月份的打卡記錄
 */
function getEmployeeMonthlyAttendance() {
  try {
    const session = checkSessionInternal();
    if (!session.ok) {
      return jsonResponse({ ok: false, msg: 'SESSION_INVALID', code: 'SESSION_INVALID' });
    }
    
    const employeeId = session.user.userId;
    const yearMonth = getParam('yearMonth');
    
    if (!yearMonth) {
      return jsonResponse({ ok: false, msg: 'MISSING_YEAR_MONTH', code: 'MISSING_YEAR_MONTH' });
    }
    
    Logger.log(`📋 API: 取得 ${employeeId} 在 ${yearMonth} 的打卡記錄`);
    
    // 呼叫 SalaryManagement-Enhanced.gs 中的內部函數
    const records = getEmployeeMonthlyAttendanceInternal(employeeId, yearMonth);
    
    return jsonResponse({ ok: true, records: records });
    
  } catch (error) {
    Logger.log('❌ getEmployeeMonthlyAttendance API 錯誤: ' + error);
    return jsonResponse({ ok: false, msg: error.toString(), code: 'ERROR' });
  }
}


// ==================== 薪資匯出功能（管理員專用） ====================

/**
 * ✅ 匯出所有員工薪資總表為 Excel（修正版）
 */
function exportAllSalaryExcel() {
  try {
    // 從全域變數取得參數
    const e = globalThis.currentRequest;
    
    if (!e || !e.parameter) {
      return jsonResponse(false, null, '無法取得請求參數', 'NO_REQUEST');
    }
    
    const params = e.parameter;
    const yearMonth = params.yearMonth;
    
    Logger.log('📥 exportAllSalaryExcel 收到參數:');
    Logger.log('   yearMonth: ' + yearMonth);
    
    // 驗證參數
    if (!yearMonth) {
      return jsonResponse(false, null, '缺少 yearMonth 參數', 'MISSING_YEAR_MONTH');
    }
    
    // ⭐⭐⭐ 移除 Session 驗證（已在 Main.gs 中驗證過）
    
    Logger.log('✅ 開始匯出薪資總表: ' + yearMonth);
    
    // 取得薪資記錄
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const salarySheet = ss.getSheetByName('月薪資記錄');
    
    if (!salarySheet) {
      return jsonResponse(false, null, '找不到月薪資記錄工作表', 'SHEET_NOT_FOUND');
    }
    
    const lastRow = salarySheet.getLastRow();
    
    if (lastRow <= 1) {
      return jsonResponse(false, null, '沒有薪資記錄', 'NO_RECORDS');
    }
    
    const allData = salarySheet.getRange(2, 1, lastRow - 1, salarySheet.getLastColumn()).getValues();
    
    Logger.log(`📊 原始資料筆數: ${allData.length}`);
    
    // 篩選指定月份的記錄
    const records = [];
    
    allData.forEach((row, index) => {
      const rowYearMonth = row[3]; // 第4欄是年月
      
      let normalizedYearMonth = '';
      
      if (rowYearMonth instanceof Date) {
        normalizedYearMonth = Utilities.formatDate(rowYearMonth, 'Asia/Taipei', 'yyyy-MM');
      } else if (typeof rowYearMonth === 'string') {
        normalizedYearMonth = rowYearMonth.substring(0, 7);
      } else {
        return;
      }
      
      if (normalizedYearMonth === yearMonth) {
        records.push(row);
        Logger.log(`✅ 找到符合記錄: 員工 ${row[2]}, 年月 ${normalizedYearMonth}`);
      }
    });
    
    Logger.log(`📊 找到 ${records.length} 筆 ${yearMonth} 的記錄`);
    
    if (records.length === 0) {
      return jsonResponse(false, null, `${yearMonth} 沒有薪資記錄`, 'NO_RECORDS_FOR_MONTH');
    }
    
    // 建立新的試算表
    const spreadsheet = SpreadsheetApp.create(`薪資總表_${yearMonth}`);
    const sheet = spreadsheet.getActiveSheet();
    sheet.setName('薪資明細');
    
    // 設定標題列
    const headers = [
      '薪資單ID', '員工ID', '員工姓名', '年月', '薪資類型', '時薪', '工作時數', '總加班時數',
      '基本薪資', '職務加給', '伙食費', '交通補助', '全勤獎金', '績效獎金', '其他津貼',
      '平日加班費', '休息日加班費', '國定假日加班費',
      '未休假補薪',      // V ⭐ NEW
      '未休假天數',      // W ⭐ NEW
      '月休補薪',        // X ⭐ NEW
      '未休月休天數',     // Y ⭐ NEW
      '勞保費', '健保費', '就業保險費', '勞退自提', '所得稅',
      '請假扣款', '福利金扣款', '宿舍費用', '團保費用', '其他扣款',
      '應發總額', '實發金額',
      '銀行代碼', '銀行帳號',
      '狀態', '備註', '建立時間'
    ];
    
    // 寫入標題列
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // 格式化標題列
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#4a5568');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    headerRange.setHorizontalAlignment('center');
    
    // 寫入資料
    if (records.length > 0) {
      const dataToWrite = records.map(row => {
        while (row.length < headers.length) {
          row.push('');
        }
        return row.slice(0, headers.length);
      });
      
      sheet.getRange(2, 1, dataToWrite.length, headers.length).setValues(dataToWrite);
      Logger.log(`✅ 已寫入 ${dataToWrite.length} 筆資料`);
    }
    
    // 自動調整欄寬
    for (let i = 1; i <= headers.length; i++) {
      sheet.autoResizeColumn(i);
    }
    
    // 凍結標題列
    sheet.setFrozenRows(1);
    
    // 設定檔案權限
    const file = DriveApp.getFileById(spreadsheet.getId());
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    // 取得下載連結
    const fileId = spreadsheet.getId();
    const downloadUrl = `https://docs.google.com/spreadsheets/d/${fileId}/export?format=xlsx`;
    
    Logger.log('✅ Excel 已生成');
    Logger.log('📊 檔案 ID: ' + fileId);
    Logger.log('🔗 下載連結: ' + downloadUrl);
    
    return jsonResponse(true, {
      fileUrl: downloadUrl,
      fileId: fileId,
      fileName: `薪資總表_${yearMonth}`,
      recordCount: records.length
    }, '薪資總表已生成');
    
  } catch (error) {
    Logger.log('❌ exportAllSalaryExcel 錯誤: ' + error.toString());
    Logger.log('❌ 錯誤堆疊: ' + error.stack);
    return jsonResponse(false, null, '匯出失敗: ' + error.toString(), 'EXPORT_ERROR');
  }
}
/**
 * ✅ 取得或建立資料夾
 * 
 * @param {string} folderName - 資料夾名稱
 * @param {Folder} parentFolder - 父資料夾（可選）
 * @returns {Folder} 資料夾物件
 */
function getOrCreateFolder(folderName, parentFolder) {
  const parent = parentFolder || DriveApp.getRootFolder();
  
  const folders = parent.getFoldersByName(folderName);
  
  if (folders.hasNext()) {
    return folders.next();
  } else {
    return parent.createFolder(folderName);
  }
}

/**
 * ✅ 取得銀行名稱（重複使用現有函數）
 */
function getBankName(code) {
  if (!code || code === '') {
    return '未設定';
  }
  
  // 自動補零到 3 位數
  const bankCode = String(code).padStart(3, '0');
  
  const banks = {
    // 公股銀行
    "004": "臺灣銀行",
    "005": "臺灣土地銀行",
    "006": "合作金庫商業銀行",
    "007": "第一商業銀行",
    "008": "華南商業銀行",
    "009": "彰化商業銀行",
    "011": "上海商業儲蓄銀行",
    "012": "台北富邦商業銀行",
    "013": "國泰世華商業銀行",
    "016": "高雄銀行",
    "017": "兆豐國際商業銀行",
    "050": "臺灣中小企業銀行",
    
    // 民營銀行
    "103": "臺灣新光商業銀行",
    "108": "陽信商業銀行",
    "118": "板信商業銀行",
    "147": "三信商業銀行",
    "803": "聯邦商業銀行",
    "805": "遠東國際商業銀行",
    "806": "元大商業銀行",
    "807": "永豐商業銀行",
    "808": "玉山商業銀行",
    "809": "凱基商業銀行",
    "810": "星展（台灣）商業銀行",
    "812": "台新國際商業銀行",
    "816": "安泰商業銀行",
    "822": "中國信託商業銀行",
    "826": "樂天國際商業銀行",
    
    // 外商銀行
    "052": "渣打國際商業銀行",
    "081": "匯豐（台灣）商業銀行",
    "101": "瑞興商業銀行",
    "102": "華泰商業銀行",
    "815": "日盛國際商業銀行",
    "824": "連線商業銀行",
    
    // 郵局
    "700": "中華郵政"
  };
  
  return banks[bankCode] || `未知銀行 (${bankCode})`;
}

console.log('✅ 薪資匯出功能已載入（管理員專用）');


function testExportSalaryDirect() {
  Logger.log('🧪 开始测试汇出功能');
  
  // 模拟请求参数
  const mockParams = {
    action: 'exportAllSalaryExcel',
    token: '48c4c025-f8fa-4528-9429-910b507c6774',  // ⚠️ 替换成真实的 token
    yearMonth: '2025-12',
    callback: 'callback'
  };
  
  // 模拟 doGet 请求
  const mockEvent = {
    parameter: mockParams
  };
  
  const result = doGet(mockEvent);
  Logger.log('📤 测试结果:');
  Logger.log(result.getContent());
}

/**
 * ✅ 修正：計算請假扣款（直接使用工作時數）
 * 
 * 修正內容：
 * 1. 直接從「請假紀錄」工作表的 H 欄（工作時數）讀取 ✅
 * 2. 不再需要計算，直接使用已核准的時數 ✅
 * 3. 支援病假扣半薪、事假扣全薪 ✅
 * 
 * @param {string} employeeId - 員工ID
 * @param {string} yearMonth - 年月 (YYYY-MM)
 * @param {number} baseSalary - 基本薪資
 * @param {number} mealAllowance - 伙食費
 * @param {number} transportAllowance - 交通津貼
 * @return {object} 扣款明細
 */
function calculateLeaveDeduction(employeeId, yearMonth, baseSalary, mealAllowance, transportAllowance) {
  try {
    Logger.log('🏥 開始計算請假扣款（使用工作時數）...');
    Logger.log(`   員工ID: ${employeeId}`);
    Logger.log(`   年月: ${yearMonth}`);
    
    // ⭐⭐⭐ 修正：只計算基本薪資 + 伙食費 + 交通津貼
    const hourlyRate = (baseSalary + mealAllowance + transportAllowance) / 30 / 8;
    
    Logger.log(`   ⏰ 時薪: ${hourlyRate.toFixed(2)} 元/小時`);
    Logger.log(`      基本薪資: ${baseSalary}`);
    Logger.log(`      伙食費: ${mealAllowance}`);
    Logger.log(`      交通津貼: ${transportAllowance}`);
    Logger.log('');
    
    // ⭐⭐⭐ 修正：直接從「請假紀錄」讀取工作時數
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('請假紀錄');
    
    if (!sheet) {
      Logger.log('   ⚠️ 找不到「請假紀錄」工作表');
      return {
        sickLeaveHours: 0,
        sickLeaveDeduction: 0,
        personalLeaveHours: 0,
        personalLeaveDeduction: 0,
        totalDeduction: 0
      };
    }
    
    const data = sheet.getDataRange().getValues();
    
    let sickLeaveHours = 0;
    let personalLeaveHours = 0;
    
    Logger.log('📋 開始統計請假時數...');
    Logger.log('');
    
    // ⭐⭐⭐ 遍歷所有記錄，抓取已核准的請假
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      // 欄位對應：
      // B (索引1): 員工ID
      // E (索引4): 假別
      // F (索引5): 開始時間
      // H (索引7): 工作時數 ⭐⭐⭐ 關鍵欄位
      // K (索引10): 狀態
      
      const recordEmployeeId = String(row[1]).trim();
      const leaveType = row[4];
      const startDateTime = row[5];
      const workHours = parseFloat(row[7]) || 0;  // ⭐⭐⭐ 直接讀取 H 欄
      const status = String(row[10]).trim();
      
      // 檢查員工ID
      if (recordEmployeeId !== employeeId) continue;
      
      // 檢查狀態（只計算已核准的）
      if (status !== 'APPROVED') continue;
      
      // 檢查年月
      let recordYearMonth = '';
      
      if (startDateTime instanceof Date) {
        recordYearMonth = Utilities.formatDate(startDateTime, 'Asia/Taipei', 'yyyy-MM');
      } else if (typeof startDateTime === 'string') {
        recordYearMonth = String(startDateTime).substring(0, 7);
      } else {
        continue;
      }
      
      if (recordYearMonth !== yearMonth) continue;
      
      // ⭐⭐⭐ 累計病假和事假時數
      if (leaveType === 'SICK_LEAVE') {
        sickLeaveHours += workHours;
        Logger.log(`   ✅ 病假: ${workHours} 小時 (${startDateTime})`);
      } else if (leaveType === 'PERSONAL_LEAVE') {
        personalLeaveHours += workHours;
        Logger.log(`   ✅ 事假: ${workHours} 小時 (${startDateTime})`);
      }
    }
    
    Logger.log('');
    Logger.log('📊 請假統計結果:');
    Logger.log(`   病假總時數: ${sickLeaveHours} 小時`);
    Logger.log(`   事假總時數: ${personalLeaveHours} 小時`);
    Logger.log('');
    
    // ⭐⭐⭐ 計算扣款金額
    const sickLeaveDeduction = Math.round(hourlyRate * sickLeaveHours * 0.5);  // 病假扣半薪
    const personalLeaveDeduction = Math.round(hourlyRate * personalLeaveHours);  // 事假扣全薪
    const totalDeduction = sickLeaveDeduction + personalLeaveDeduction;
    
    Logger.log('💰 扣款計算:');
    Logger.log(`   病假扣款: ${hourlyRate.toFixed(2)} × ${sickLeaveHours} × 0.5 = ${sickLeaveDeduction} 元`);
    Logger.log(`   事假扣款: ${hourlyRate.toFixed(2)} × ${personalLeaveHours} × 1.0 = ${personalLeaveDeduction} 元`);
    Logger.log(`   總扣款: ${totalDeduction} 元`);
    Logger.log('');
    
    return {
      sickLeaveHours: sickLeaveHours,
      sickLeaveDeduction: sickLeaveDeduction,
      personalLeaveHours: personalLeaveHours,
      personalLeaveDeduction: personalLeaveDeduction,
      totalDeduction: totalDeduction
    };
    
  } catch (error) {
    Logger.log(`   ❌ 計算請假扣款失敗: ${error.message}`);
    Logger.log(`   ❌ 錯誤堆疊: ${error.stack}`);
    return {
      sickLeaveHours: 0,
      sickLeaveDeduction: 0,
      personalLeaveHours: 0,
      personalLeaveDeduction: 0,
      totalDeduction: 0
    };
  }
}

/**
 * ⭐⭐⭐ 修正：計算整月加班費總額（支援固定費率）
 * 
 * @param {string} employeeId - 員工ID
 * @param {string} yearMonth - 年月
 * @param {number} baseSalary - 基本薪資
 * @param {boolean} isFixedRate - 是否使用固定費率（情況一專用）
 * @param {number} fixedRate - 固定費率（預設 200 元/小時）
 */
function calculateMonthlyOvertimePay(employeeId, yearMonth, baseSalary, isFixedRate = false, fixedRate = 200) {
  try {
    Logger.log('💼 開始計算加班費');
    Logger.log(`   員工ID: ${employeeId}`);
    Logger.log(`   年月: ${yearMonth}`);
    Logger.log(`   基本薪資: ${baseSalary}`);
    Logger.log(`   使用固定費率: ${isFixedRate ? '是' : '否'}`);
    
    // ⭐ 加在這裡
    let actualRate = fixedRate;
    if (isFixedRate) {
      const storedRate = PropertiesService.getScriptProperties().getProperty('DRIVER_OVERTIME_RATE');
      if (storedRate) actualRate = parseInt(storedRate);
      Logger.log(`   ⭐ 加班費率: $${actualRate}/小時（來源: ${storedRate ? 'Script Properties' : '預設值'}）`);
    }

    // ⭐ 這行改用 actualRate，不用原本的 fixedRate
    const hourlyRate = isFixedRate ? actualRate : Math.round(baseSalary / 30 / 8);

    
    Logger.log(`   時薪: $${hourlyRate} ${isFixedRate ? '（固定）' : '（計算）'}`);
    Logger.log('');
    
    const overtimeRecords = getEmployeeMonthlyOvertime(employeeId, yearMonth);
    
    if (overtimeRecords.length === 0) {
      Logger.log('   ⚠️ 無加班記錄');
      return {
        weekdayOvertimePay: 0,
        restdayOvertimePay: 0,
        holidayOvertimePay: 0,
        totalOvertimeHours: 0
      };
    }
    
    Logger.log(`   📋 找到 ${overtimeRecords.length} 筆加班記錄`);
    
    const overtimeByDate = {};
    overtimeRecords.forEach(record => {
      const date = record.date;
      if (!overtimeByDate[date]) {
        overtimeByDate[date] = 0;
      }
      overtimeByDate[date] += parseFloat(record.hours) || 0;
    });
    
    let weekdayOvertimePay = 0;
    let restdayOvertimePay = 0;
    let holidayOvertimePay = 0;
    let totalOvertimeHours = 0;
    
    Object.keys(overtimeByDate).forEach(date => {
      let dailyHours = overtimeByDate[date];
      const dateType = getDateType(date);
      
      let maxHours = 4;
      if (dateType === 'restday') maxHours = 12;
      if (dateType === 'holiday') maxHours = 8;
      
      if (dailyHours > maxHours) {
        dailyHours = maxHours;
      }
      
      let totalPay = 0;
      if (isFixedRate) {
        // ⭐ 不定時制：固定費率，不套倍率
        totalPay = Math.round(actualRate * dailyHours);
        Logger.log(`   ${date}: ${dailyHours}h × $${actualRate} = $${totalPay}`);
      } else {
        // 標準工時：套倍率
        const pay = calculateOvertimePay(dailyHours, hourlyRate, dateType);
        totalPay = pay.firstPay + pay.secondPay + pay.thirdPay;
      }
      
      if (dateType === 'weekday') {
        weekdayOvertimePay += totalPay;
      } else if (dateType === 'restday') {
        restdayOvertimePay += totalPay;
      } else if (dateType === 'holiday') {
        holidayOvertimePay += totalPay;
      }
      
      totalOvertimeHours += dailyHours;
    });
    
    Logger.log('');
    Logger.log('✅ 加班費計算完成:');
    Logger.log(`   - 總時數: ${totalOvertimeHours.toFixed(1)}h`);
    Logger.log(`   - 平日加班費: $${Math.round(weekdayOvertimePay)}`);
    Logger.log(`   - 休息日加班費: $${Math.round(restdayOvertimePay)}`);
    Logger.log(`   - 例假日加班費: $${Math.round(holidayOvertimePay)}`);
    
    return {
      weekdayOvertimePay: Math.round(weekdayOvertimePay),
      restdayOvertimePay: Math.round(restdayOvertimePay),
      holidayOvertimePay: Math.round(holidayOvertimePay),
      totalOvertimeHours: totalOvertimeHours
    };
    
  } catch (error) {
    Logger.log('❌ calculateMonthlyOvertimePay 錯誤: ' + error.message);
    return {
      weekdayOvertimePay: 0,
      restdayOvertimePay: 0,
      holidayOvertimePay: 0,
      totalOvertimeHours: 0
    };
  }
}

function calculateMonthlyRestDayPay(employeeId, yearMonth, employeeType) {
  try {
    Logger.log('📅 開始計算月休補薪...');
    
    if (employeeType !== '飼料廠司機') {
      Logger.log('   ⏭️ 非飼料廠司機，跳過月休補薪');
      return {
        requiredDays: 0,
        actualDays: 0,
        missedDays: 0,
        payAmount: 0
      };
    }
    
    const REQUIRED_REST_DAYS = 6;

    // ⭐ 從 Script Properties 讀取補薪金額
    const storedCompensation = PropertiesService.getScriptProperties().getProperty('DRIVER_MONTHLY_REST_COMPENSATION');
    const REST_DAY_PAY = storedCompensation ? parseInt(storedCompensation) : 1500;
    
    Logger.log(`   📋 應休天數: ${REQUIRED_REST_DAYS} 天`);
    Logger.log(`   💰 未休補薪: ${REST_DAY_PAY} 元/天（來源: ${storedCompensation ? 'Script Properties' : '預設值'}）`);
    
    const actualRestDays = getActualMonthlyRestDays(employeeId, yearMonth);
    
    const missedDays = Math.max(0, REQUIRED_REST_DAYS - actualRestDays);
    const payAmount = missedDays * REST_DAY_PAY;
    
    Logger.log(`   ✅ 月休補薪計算完成:`);
    Logger.log(`      實際休假: ${actualRestDays} 天`);
    Logger.log(`      未休天數: ${missedDays} 天`);
    Logger.log(`      補薪金額: ${payAmount} 元`);
    Logger.log('');
    
    return {
      requiredDays: REQUIRED_REST_DAYS,
      actualDays: actualRestDays,
      missedDays: missedDays,
      payAmount: payAmount
    };
    
  } catch (error) {
    Logger.log(`   ❌ 計算月休補薪失敗: ${error.message}`);
    return {
      requiredDays: 0,
      actualDays: 0,
      missedDays: 0,
      payAmount: 0
    };
  }
}
/**
 * ⭐ 取得員工實際月休天數（最終修正版）
 * 
 * 修正內容：
 * 1. 處理 Date 對象 → 字符串轉換 ✅
 * 2. 添加詳細的調試日誌 ✅
 * 3. 統一年月格式為 YYYY-MM ✅
 * 
 * @param {string} employeeId - 員工ID
 * @param {string} yearMonth - 年月 (YYYY-MM)
 * @return {number} 實際休假天數
 */
function getActualMonthlyRestDays(employeeId, yearMonth) {
  try {
    Logger.log('   🔍 查詢實際月休天數...');
    Logger.log(`   🔍 查詢參數: employeeId="${employeeId}", yearMonth="${yearMonth}"`);
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const restDaySheet = ss.getSheetByName('月休記錄');
    
    if (!restDaySheet) {
      Logger.log('   ⚠️ 找不到「月休記錄」工作表');
      return 0;
    }
    
    Logger.log('   ✅ 找到「月休記錄」工作表');
    
    const data = restDaySheet.getDataRange().getValues();
    
    Logger.log(`   📊 「月休記錄」工作表數據筆數: ${data.length - 1}`);
    
    if (data.length < 2) {
      Logger.log('   ⚠️ 工作表沒有數據（只有標題列）');
      return 0;
    }
    
    Logger.log(`   📋 標題列: ${data[0].join(', ')}`);
    
    // 查找該員工該月份的記錄
    for (let i = 1; i < data.length; i++) {
      Logger.log(`\n   --- 檢查第 ${i + 1} 行 (索引 ${i}) ---`);
      Logger.log(`   原始數據 [0]: "${data[i][0]}" (類型: ${typeof data[i][0]})`);
      Logger.log(`   原始數據 [1]: "${data[i][1]}" (類型: ${typeof data[i][1]})`);
      Logger.log(`   原始數據 [2]: "${data[i][2]}" (類型: ${typeof data[i][2]})`);
      
      // ⭐⭐⭐ 關鍵修正：處理員工ID
      const recordEmployeeId = String(data[i][0]).trim();
      
      // ⭐⭐⭐ 關鍵修正：處理年月（可能是 Date 對象）
      let recordYearMonth = '';
      
      if (data[i][1] instanceof Date) {
        // 如果是 Date 對象，轉換為 YYYY-MM 格式
        recordYearMonth = Utilities.formatDate(data[i][1], 'Asia/Taipei', 'yyyy-MM');
        Logger.log(`   ⭐ 檢測到 Date 對象，轉換為: "${recordYearMonth}"`);
      } else if (typeof data[i][1] === 'string') {
        // 如果是字符串，提取前 7 個字符（YYYY-MM）
        recordYearMonth = String(data[i][1]).trim().substring(0, 7);
        Logger.log(`   ⭐ 字符串格式，提取為: "${recordYearMonth}"`);
      } else {
        // 其他情況，嘗試轉換為字符串
        recordYearMonth = String(data[i][1]).trim();
        Logger.log(`   ⚠️ 未知格式，轉換為: "${recordYearMonth}"`);
      }
      
      const restDays = parseFloat(data[i][2]) || 0;
      
      Logger.log(`   轉換後 recordEmployeeId: "${recordEmployeeId}"`);
      Logger.log(`   轉換後 recordYearMonth: "${recordYearMonth}"`);
      Logger.log(`   轉換後 restDays: ${restDays}`);
      
      // 比對
      const employeeIdMatch = (recordEmployeeId === employeeId);
      const yearMonthMatch = (recordYearMonth === yearMonth);
      
      Logger.log(`   比對 employeeId: "${recordEmployeeId}" === "${employeeId}" ? ${employeeIdMatch}`);
      Logger.log(`   比對 yearMonth: "${recordYearMonth}" === "${yearMonth}" ? ${yearMonthMatch}`);
      
      if (employeeIdMatch && yearMonthMatch) {
        Logger.log(`   ✅✅✅ 找到匹配記錄！`);
        Logger.log(`   ✅ 從「月休記錄」讀取: ${restDays} 天`);
        return restDays;
      } else {
        Logger.log(`   ⏭️ 不匹配，繼續下一筆...`);
      }
    }
    
    Logger.log('\n   ⚠️ 「月休記錄」中無該月資料，視為 0 天');
    return 0;
    
  } catch (error) {
    Logger.log(`\n   ❌ 查詢月休天數失敗: ${error.message}`);
    Logger.log(`   ❌ 錯誤堆疊: ${error.stack}`);
    return 0;
  }
}

/**
 * 🔧 初始化「月休記錄」工作表
 * 
 * 工作表結構：
 * A: 員工ID
 * B: 年月 (YYYY-MM)
 * C: 休假天數
 * D: 備註
 * E: 更新時間
 */
function initMonthlyRestDaySheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('月休記錄');
  
  if (!sheet) {
    Logger.log('📝 建立「月休記錄」工作表...');
    
    sheet = ss.insertSheet('月休記錄');
    
    const headers = ['員工ID', '年月', '休假天數', '備註', '更新時間'];
    sheet.appendRow(headers);
    
    // 美化標題列
    const headerRange = sheet.getRange(1, 1, 1, 5);
    headerRange.setBackground('#4A90E2');
    headerRange.setFontColor('#FFFFFF');
    headerRange.setFontWeight('bold');
    headerRange.setHorizontalAlignment('center');
    
    sheet.setFrozenRows(1);
    
    Logger.log('✅ 「月休記錄」工作表已建立');
  }
  
  return sheet;
}

/**
 * ⭐ 設定員工月休天數（修正版 - 避免 Date 自動轉換）
 * 
 * 修正內容：
 * 1. 將年月存為字符串，前面加單引號 ' 避免被轉換為 Date ✅
 * 2. 或者使用 setNumberFormat 設定為文本格式 ✅
 * 
 * @param {string} employeeId - 員工ID
 * @param {string} yearMonth - 年月 (YYYY-MM)
 * @param {number} restDays - 休假天數
 * @param {string} note - 備註
 */
function setMonthlyRestDays(employeeId, yearMonth, restDays, note = '') {
  try {
    const sheet = initMonthlyRestDaySheet();
    const data = sheet.getDataRange().getValues();
    
    // 查找是否已存在
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      const existingId = String(data[i][0]).trim();
      
      // 處理現有記錄的年月（可能是 Date）
      let existingYearMonth = '';
      if (data[i][1] instanceof Date) {
        existingYearMonth = Utilities.formatDate(data[i][1], 'Asia/Taipei', 'yyyy-MM');
      } else {
        existingYearMonth = String(data[i][1]).trim().substring(0, 7);
      }
      
      if (existingId === employeeId && existingYearMonth === yearMonth) {
        rowIndex = i + 1;
        break;
      }
    }
    
    // ⭐⭐⭐ 關鍵：在年月前加單引號，強制為文本格式
    const yearMonthText = "'" + yearMonth;  // '2026-01
    
    const row = [
      employeeId,
      yearMonthText,  // ⭐ 使用加了單引號的版本
      restDays,
      note,
      new Date()
    ];
    
    if (rowIndex > 0) {
      // 更新現有記錄
      sheet.getRange(rowIndex, 1, 1, 5).setValues([row]);
      Logger.log(`✅ 更新月休記錄: ${employeeId}, ${yearMonth}, ${restDays}天`);
    } else {
      // 新增記錄
      sheet.appendRow(row);
      Logger.log(`✅ 新增月休記錄: ${employeeId}, ${yearMonth}, ${restDays}天`);
    }
    
    // ⭐⭐⭐ 額外保險：設定 B 欄為文本格式
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 2).setNumberFormat('@');  // @ = 文本格式
    
    return { success: true, message: '月休記錄已更新' };
    
  } catch (error) {
    Logger.log(`❌ 設定月休記錄失敗: ${error.message}`);
    return { success: false, message: error.message };
  }
}

/**
 * 🧪 測試月休補薪計算
 */
function testMonthlyRestDayPay() {
  Logger.log('🧪 測試月休補薪計算');
  Logger.log('');
  
  // 先初始化工作表
  initMonthlyRestDaySheet();
  
  // 設定測試資料（假設員工休了4天，少休2天）
  const testEmployeeId = 'Uf664a35632b736301d674d8b2cc3f8c0';
  const testYearMonth = '2026-01';
  
  setMonthlyRestDays(testEmployeeId, testYearMonth, 4, '測試資料');
  
  Logger.log('');
  Logger.log('📋 開始測試計算...');
  
  const result = calculateMonthlyRestDayPay(testEmployeeId, testYearMonth, '飼料廠司機');
  
  Logger.log('');
  Logger.log('📤 測試結果:');
  Logger.log(`   應休天數: ${result.requiredDays} 天`);
  Logger.log(`   實際休假: ${result.actualDays} 天`);
  Logger.log(`   未休天數: ${result.missedDays} 天`);
  Logger.log(`   補薪金額: ${result.payAmount} 元`);
  Logger.log('');
  
  if (result.missedDays === 2 && result.payAmount === 3000) {
    Logger.log('✅✅✅ 測試成功！');
    Logger.log('   未休2天 × 1,500 = 3,000元 ✅');
  } else {
    Logger.log('❌ 測試失敗');
  }
}

// ==================================================================================
// 未休假補薪計算 - 完整修正版
// ==================================================================================

/**
 * ✅ 計算未休假補薪（Date 處理修正版）
 * 
 * 修正內容：
 * 1. 處理年月欄位可能是 Date 對象的情況 ✅
 * 2. 統一轉換為 YYYY-MM 格式再比對 ✅
 * 
 * @param {string} employeeId - 員工ID
 * @param {string} yearMonth - 年月 (YYYY-MM)
 * @param {number} baseSalary - 基本薪資
 * @param {number} totalFixedAllowance - 固定津貼總額
 * @return {object} { amount, days, dailyRate, employeeName }
 */
function calculateUnusedLeavePay(employeeId, yearMonth, baseSalary, totalFixedAllowance) {
  try {
    Logger.log('🏖️ 開始計算未休假補薪');
    Logger.log('   員工ID: ' + employeeId);
    Logger.log('   年月: ' + yearMonth);
    Logger.log('   基本薪資: ' + baseSalary);
    Logger.log('   固定津貼: ' + totalFixedAllowance);
    Logger.log('');
    
    // ⭐⭐⭐ 步驟 1：從「員工薪資設定」獲取員工姓名
    let employeeName = '未知員工';
    
    try {
      const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('員工薪資設定');
      
      if (configSheet) {
        const configData = configSheet.getDataRange().getValues();
        
        // 查找該員工
        for (let i = 1; i < configData.length; i++) {
          if (configData[i][0] === employeeId) {
            employeeName = configData[i][1] || '未知員工';
            Logger.log('✅ 從「員工薪資設定」讀取姓名: ' + employeeName);
            break;
          }
        }
      } else {
        Logger.log('⚠️ 找不到「員工薪資設定」工作表');
      }
    } catch (nameError) {
      Logger.log('⚠️ 讀取員工姓名失敗: ' + nameError.message);
    }
    
    // ⭐⭐⭐ 步驟 2：從「未休假記錄」讀取未休天數（修正版 - 處理 Date）
    let unusedDays = 0;
    
    try {
      const unusedLeaveSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('未休假記錄');
      
      if (unusedLeaveSheet) {
        const data = unusedLeaveSheet.getDataRange().getValues();
        
        Logger.log('📋 「未休假記錄」工作表資料筆數: ' + (data.length - 1));
        
        // 查找該員工該月份的記錄
        for (let i = 1; i < data.length; i++) {
          const recordEmployeeId = String(data[i][0]).trim();  // A欄：員工ID
          
          // ⭐⭐⭐ 關鍵修正：處理年月欄位（可能是 Date 對象）
          let recordYearMonth = '';
          
          if (data[i][1] instanceof Date) {
            // 如果是 Date 對象，轉換為 YYYY-MM 格式
            recordYearMonth = Utilities.formatDate(data[i][1], 'Asia/Taipei', 'yyyy-MM');
            Logger.log(`   檢查第 ${i + 1} 行: ${recordEmployeeId}, ${recordYearMonth} (Date → 文本), ${data[i][2]}天`);
          } else if (typeof data[i][1] === 'string') {
            // 如果是字符串，提取前 7 個字符
            recordYearMonth = String(data[i][1]).trim().substring(0, 7);
            Logger.log(`   檢查第 ${i + 1} 行: ${recordEmployeeId}, ${recordYearMonth} (文本), ${data[i][2]}天`);
          } else {
            // 其他情況，嘗試轉換
            recordYearMonth = String(data[i][1]).trim();
            Logger.log(`   檢查第 ${i + 1} 行: ${recordEmployeeId}, ${recordYearMonth} (未知格式), ${data[i][2]}天`);
          }
          
          const recordDays = parseFloat(data[i][2]) || 0;  // C欄：未休天數
          
          // ⭐⭐⭐ 比對（現在兩邊都是 YYYY-MM 格式）
          if (recordEmployeeId === employeeId && recordYearMonth === yearMonth) {
            unusedDays = recordDays;
            Logger.log('✅✅✅ 找到未休假記錄: ' + unusedDays + ' 天');
            break;
          }
        }
        
        if (unusedDays === 0) {
          Logger.log('⚠️ 「未休假記錄」中無該員工該月資料');
        }
      } else {
        Logger.log('⚠️ 找不到「未休假記錄」工作表');
      }
    } catch (leaveError) {
      Logger.log('⚠️ 讀取未休假記錄失敗: ' + leaveError.message);
    }
    
    // ⭐⭐⭐ 步驟 3：計算日薪
    const dailyRate = Math.round((baseSalary + totalFixedAllowance) / 30);
    
    Logger.log('');
    Logger.log('📊 計算公式:');
    Logger.log(`   日薪 = (基本薪資 ${baseSalary} + 固定津貼 ${totalFixedAllowance}) / 30`);
    Logger.log(`   日薪 = ${dailyRate} 元`);
    
    // ⭐⭐⭐ 步驟 4：計算補薪金額
    const payAmount = unusedDays * dailyRate;
    
    Logger.log('');
    Logger.log('💰 未休假補薪計算:');
    Logger.log(`   未休天數: ${unusedDays} 天`);
    Logger.log(`   日薪: ${dailyRate} 元`);
    Logger.log(`   補薪金額: ${unusedDays} × ${dailyRate} = ${payAmount} 元`);
    Logger.log('');
    
    return {
      amount: payAmount,
      days: unusedDays,
      dailyRate: dailyRate,
      employeeName: employeeName
    };
    
  } catch (error) {
    Logger.log('❌ calculateUnusedLeavePay 錯誤: ' + error.message);
    Logger.log('   錯誤堆疊: ' + error.stack);
    
    return {
      amount: 0,
      days: 0,
      dailyRate: 0,
      employeeName: '未知員工'
    };
  }
}

/**
 * 🔧 設定員工未休假天數
 * 
 * @param {string} employeeId - 員工ID
 * @param {string} yearMonth - 年月 (YYYY-MM)
 * @param {number} unusedDays - 未休天數
 * @param {string} note - 備註
 */
function setUnusedLeaveDays(employeeId, yearMonth, unusedDays, note = '') {
  try {
    const sheet = initUnusedLeaveSheet();
    const data = sheet.getDataRange().getValues();
    
    // 查找是否已存在
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === employeeId && data[i][1] === yearMonth) {
        rowIndex = i + 1;
        break;
      }
    }
    
    const row = [
      employeeId,
      yearMonth,
      unusedDays,
      note,
      new Date()
    ];
    
    if (rowIndex > 0) {
      // 更新現有記錄
      sheet.getRange(rowIndex, 1, 1, 5).setValues([row]);
      Logger.log(`✅ 更新未休假記錄: ${employeeId}, ${yearMonth}, ${unusedDays}天`);
    } else {
      // 新增記錄
      sheet.appendRow(row);
      Logger.log(`✅ 新增未休假記錄: ${employeeId}, ${yearMonth}, ${unusedDays}天`);
    }
    
    return { success: true, message: '未休假記錄已更新' };
    
  } catch (error) {
    Logger.log(`❌ 設定未休假記錄失敗: ${error.message}`);
    return { success: false, message: error.message };
  }
}

/**
 * 🧪 測試未休假補薪計算
 */
function testCalculateUnusedLeavePay() {
  Logger.log('═══════════════════════════════════════');
  Logger.log('🧪 測試未休假補薪計算');
  Logger.log('═══════════════════════════════════════');
  Logger.log('');
  
  // 1. 初始化工作表
  initUnusedLeaveSheet();
  
  // 2. 設定測試資料（假設員工未休6天特休）
  const testEmployeeId = 'Uf664a35632b736301d674d8b2cc3f8c0';  // ⚠️ 替換成實際員工ID
  const testYearMonth = '2026-01';
  
  setUnusedLeaveDays(testEmployeeId, testYearMonth, 6, '測試資料：未休6天特休');
  
  Logger.log('📋 測試資料已設定');
  Logger.log('');
  
  // 3. 執行計算
  const baseSalary = 38000;
  const totalFixedAllowance = 9000;
  
  Logger.log('📊 測試參數:');
  Logger.log(`   員工ID: ${testEmployeeId}`);
  Logger.log(`   年月: ${testYearMonth}`);
  Logger.log(`   基本薪資: ${baseSalary}`);
  Logger.log(`   固定津貼: ${totalFixedAllowance}`);
  Logger.log('');
  
  const result = calculateUnusedLeavePay(
    testEmployeeId,
    testYearMonth,
    baseSalary,
    totalFixedAllowance
  );
  
  Logger.log('');
  Logger.log('📤 測試結果:');
  Logger.log(`   員工姓名: ${result.employeeName}`);
  Logger.log(`   未休天數: ${result.days} 天`);
  Logger.log(`   日薪: ${result.dailyRate} 元`);
  Logger.log(`   補薪金額: ${result.amount} 元`);
  Logger.log('');
  
  // 驗證結果
  const expectedDailyRate = Math.round((baseSalary + totalFixedAllowance) / 30);
  const expectedAmount = 6 * expectedDailyRate;
  
  if (result.days === 6 && result.amount === expectedAmount) {
    Logger.log('✅✅✅ 測試成功！');
    Logger.log(`   未休6天 × ${expectedDailyRate} = ${expectedAmount}元 ✅`);
  } else {
    Logger.log('❌ 測試失敗');
    Logger.log(`   預期：未休6天，補薪${expectedAmount}元`);
    Logger.log(`   實際：未休${result.days}天，補薪${result.amount}元`);
  }
  
  Logger.log('');
  Logger.log('═══════════════════════════════════════');
}

/**
 * 🔧 批次設定測試資料
 */
function setupUnusedLeaveTestData() {
  Logger.log('📋 批次設定未休假測試資料...');
  Logger.log('');
  
  initUnusedLeaveSheet();
  
  // 測試資料
  const testData = [
    ['Uf664a35632b736301d674d8b2cc3f8c0', '2026-01', 6, '測試：未休6天特休'],
    ['Uf664a35632b736301d674d8b2cc3f8c0', '2026-01', 3, '測試：未休3天特休'],
    ['Uf664a35632b736301d674d8b2cc3f8c0', '2026-01', 0, '測試：已休完所有特休']
  ];
  
  testData.forEach(data => {
    setUnusedLeaveDays(data[0], data[1], data[2], data[3]);
  });
  
  Logger.log('');
  Logger.log('✅ 測試資料設定完成！');
  Logger.log('');
  Logger.log('🔍 請檢查「未休假記錄」工作表');
}

/**
 * 🔧 初始化「未休假記錄」工作表
 */
function initUnusedLeaveSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('未休假記錄');
  
  if (!sheet) {
    Logger.log('📝 建立「未休假記錄」工作表...');
    
    sheet = ss.insertSheet('未休假記錄');
    
    const headers = ['員工ID', '年月', '未休天數', '備註', '更新時間'];
    sheet.appendRow(headers);
    
    // 美化標題列
    const headerRange = sheet.getRange(1, 1, 1, 5);
    headerRange.setBackground('#4A90E2');
    headerRange.setFontColor('#FFFFFF');
    headerRange.setFontWeight('bold');
    headerRange.setHorizontalAlignment('center');
    
    sheet.setFrozenRows(1);
    
    Logger.log('✅ 「未休假記錄」工作表已建立');
  }
  
  return sheet;
}


/**
 * ⭐ 取得員工實際月休天數（完整除錯版）
 * 
 * 修正內容：
 * 1. 添加 String() 和 trim() 類型轉換 ✅
 * 2. 添加完整的調試日誌 ✅
 * 3. 顯示每一行的實際數據 ✅
 * 
 * @param {string} employeeId - 員工ID
 * @param {string} yearMonth - 年月 (YYYY-MM)
 * @return {number} 實際休假天數
 */
function getActualMonthlyRestDays(employeeId, yearMonth) {
  try {
    Logger.log('   🔍 查詢實際月休天數...');
    Logger.log(`   🔍 查詢參數: employeeId="${employeeId}", yearMonth="${yearMonth}"`);
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const restDaySheet = ss.getSheetByName('月休記錄');
    
    if (!restDaySheet) {
      Logger.log('   ⚠️ 找不到「月休記錄」工作表');
      return 0;
    }
    
    Logger.log('   ✅ 找到「月休記錄」工作表');
    
    const data = restDaySheet.getDataRange().getValues();
    
    // ⭐⭐⭐ 添加詳細的調試日誌
    Logger.log(`   📊 「月休記錄」工作表數據筆數: ${data.length - 1}`);
    
    if (data.length < 2) {
      Logger.log('   ⚠️ 工作表沒有數據（只有標題列）');
      return 0;
    }
    
    // 顯示標題列
    Logger.log(`   📋 標題列: ${data[0].join(', ')}`);
    
    // 查找該員工該月份的記錄
    for (let i = 1; i < data.length; i++) {
      // ⭐⭐⭐ 先顯示原始數據
      Logger.log(`\n   --- 檢查第 ${i + 1} 行 (索引 ${i}) ---`);
      Logger.log(`   原始數據 [0]: "${data[i][0]}" (類型: ${typeof data[i][0]})`);
      Logger.log(`   原始數據 [1]: "${data[i][1]}" (類型: ${typeof data[i][1]})`);
      Logger.log(`   原始數據 [2]: "${data[i][2]}" (類型: ${typeof data[i][2]})`);
      
      // ⭐⭐⭐ 修正：添加 String() 和 trim() 類型轉換
      const recordEmployeeId = String(data[i][0]).trim();  // A欄：員工ID
      const recordYearMonth = String(data[i][1]).trim();   // B欄：年月
      const restDays = parseFloat(data[i][2]) || 0;        // C欄：休假天數
      
      // ⭐⭐⭐ 顯示轉換後的數據
      Logger.log(`   轉換後 recordEmployeeId: "${recordEmployeeId}"`);
      Logger.log(`   轉換後 recordYearMonth: "${recordYearMonth}"`);
      Logger.log(`   轉換後 restDays: ${restDays}`);
      
      // ⭐⭐⭐ 顯示比對過程
      const employeeIdMatch = (recordEmployeeId === employeeId);
      const yearMonthMatch = (recordYearMonth === yearMonth);
      
      Logger.log(`   比對 employeeId: "${recordEmployeeId}" === "${employeeId}" ? ${employeeIdMatch}`);
      Logger.log(`   比對 yearMonth: "${recordYearMonth}" === "${yearMonth}" ? ${yearMonthMatch}`);
      
      // ⭐ 比較員工ID和年月
      if (employeeIdMatch && yearMonthMatch) {
        Logger.log(`   ✅✅✅ 找到匹配記錄！`);
        Logger.log(`   ✅ 從「月休記錄」讀取: ${restDays} 天`);
        return restDays;
      } else {
        Logger.log(`   ⏭️ 不匹配，繼續下一筆...`);
      }
    }
    
    Logger.log('\n   ⚠️ 「月休記錄」中無該月資料，視為 0 天');
    return 0;
    
  } catch (error) {
    Logger.log(`\n   ❌ 查詢月休天數失敗: ${error.message}`);
    Logger.log(`   ❌ 錯誤堆疊: ${error.stack}`);
    return 0;
  }
}


/**
 * 🧪 測試請假扣款計算
 */
function testCalculateLeaveDeduction() {
  Logger.log('═══════════════════════════════════════');
  Logger.log('🧪 測試請假扣款計算');
  Logger.log('═══════════════════════════════════════');
  Logger.log('');
  
  // 測試參數
  const testEmployeeId = 'Uf664a35632b736301d674d8b2cc3f8c0';  // ⚠️ 替換成實際員工ID
  const testYearMonth = '2026-01';
  const testBaseSalary = 38000;
  const testMealAllowance = 2400;
  const testTransportAllowance = 2000;
  
  Logger.log('📋 測試參數:');
  Logger.log(`   員工ID: ${testEmployeeId}`);
  Logger.log(`   年月: ${testYearMonth}`);
  Logger.log(`   基本薪資: ${testBaseSalary}`);
  Logger.log(`   伙食費: ${testMealAllowance}`);
  Logger.log(`   交通津貼: ${testTransportAllowance}`);
  Logger.log('');
  
  // 執行計算
  const result = calculateLeaveDeduction(
    testEmployeeId,
    testYearMonth,
    testBaseSalary,
    testMealAllowance,
    testTransportAllowance
  );
  
  Logger.log('');
  Logger.log('📤 測試結果:');
  Logger.log(`   病假時數: ${result.sickLeaveHours} 小時`);
  Logger.log(`   病假扣款: ${result.sickLeaveDeduction} 元`);
  Logger.log(`   事假時數: ${result.personalLeaveHours} 小時`);
  Logger.log(`   事假扣款: ${result.personalLeaveDeduction} 元`);
  Logger.log(`   總扣款: ${result.totalDeduction} 元`);
  Logger.log('');
  
  // 驗證計算
  const hourlyRate = (testBaseSalary + testMealAllowance + testTransportAllowance) / 30 / 8;
  
  Logger.log('🔍 驗證計算:');
  Logger.log(`   時薪 = (${testBaseSalary} + ${testMealAllowance} + ${testTransportAllowance}) / 30 / 8 = ${hourlyRate.toFixed(2)}`);
  
  if (result.sickLeaveHours > 0) {
    const expectedSickDeduction = Math.round(hourlyRate * result.sickLeaveHours * 0.5);
    Logger.log(`   病假扣款 = ${hourlyRate.toFixed(2)} × ${result.sickLeaveHours} × 0.5 = ${expectedSickDeduction} 元`);
    Logger.log(`   實際病假扣款: ${result.sickLeaveDeduction} 元 ${result.sickLeaveDeduction === expectedSickDeduction ? '✅' : '❌'}`);
  }
  
  if (result.personalLeaveHours > 0) {
    const expectedPersonalDeduction = Math.round(hourlyRate * result.personalLeaveHours);
    Logger.log(`   事假扣款 = ${hourlyRate.toFixed(2)} × ${result.personalLeaveHours} × 1.0 = ${expectedPersonalDeduction} 元`);
    Logger.log(`   實際事假扣款: ${result.personalLeaveDeduction} 元 ${result.personalLeaveDeduction === expectedPersonalDeduction ? '✅' : '❌'}`);
  }
  
  Logger.log('');
  Logger.log('═══════════════════════════════════════');
}


function testSalaryWithLeaveDeduction() {
  const employeeId = 'Uf664a35632b736301d674d8b2cc3f8c0';
  const yearMonth = '2026-01';
  
  const result = calculateMonthlySalary(employeeId, yearMonth);
  
  Logger.log('📤 薪資計算結果:');
  Logger.log(JSON.stringify(result.data, null, 2));
  
  // 檢查是否包含請假扣款明細
  if (result.data.sickLeaveHours !== undefined) {
    Logger.log('✅ 包含病假時數');
  } else {
    Logger.log('❌ 缺少病假時數');
  }
}

function finalTest() {
  const employeeId = 'Uf664a35632b736301d674d8b2cc3f8c0';
  const yearMonth = '2026-01';
  
  // 1. 計算薪資
  const calcResult = calculateMonthlySalary(employeeId, yearMonth);
  Logger.log('📊 計算結果:');
  Logger.log(`   病假時數: ${calcResult.data.sickLeaveHours}`);
  Logger.log(`   病假扣款: ${calcResult.data.sickLeaveDeduction}`);
  
  // 2. 存檔
  const saveResult = saveMonthlySalary(calcResult.data);
  Logger.log(`\n💾 存檔結果: ${saveResult.success ? '成功' : '失敗'}`);
  
  // 3. 讀取驗證
  const readResult = getMySalary(employeeId, yearMonth);
  Logger.log('\n📖 讀取結果:');
  Logger.log(`   病假時數: ${readResult.data['病假時數']}`);
  Logger.log(`   病假扣款: ${readResult.data['病假扣款']}`);
  Logger.log(`   事假時數: ${readResult.data['事假時數']}`);
  Logger.log(`   事假扣款: ${readResult.data['事假扣款']}`);
  
  // 4. 檢查
  if (readResult.data['病假時數'] === 7 && 
      readResult.data['病假扣款'] === 618) {
    Logger.log('\n✅✅✅ 測試完全成功！');
  } else {
    Logger.log('\n❌ 仍有問題，請檢查');
  }
}





/**
 * 🔧 重建月薪資記錄工作表（完整修正版）
 * 
 * 執行步驟：
 * 1. 備份現有資料
 * 2. 刪除舊工作表
 * 3. 建立新工作表（包含正確的標題列）
 * 4. 恢復資料
 */
function rebuildMonthlySalarySheetComplete() {
  try {
    Logger.log('🔧 開始重建月薪資記錄工作表...');
    Logger.log('');
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const oldSheet = ss.getSheetByName('月薪資記錄');
    
    // ========================================
    // 步驟 1: 備份現有資料
    // ========================================
    let backupData = [];
    
    if (oldSheet) {
      Logger.log('📋 步驟 1: 備份現有資料...');
      const lastRow = oldSheet.getLastRow();
      
      if (lastRow > 1) {
        // 讀取所有資料（包含標題列）
        backupData = oldSheet.getRange(1, 1, lastRow, oldSheet.getLastColumn()).getValues();
        Logger.log(`   ✅ 已備份 ${lastRow - 1} 筆薪資記錄`);
      } else {
        Logger.log('   ⚠️ 工作表沒有資料，跳過備份');
      }
      
      Logger.log('');
    }
    
    // ========================================
    // 步驟 2: 刪除舊工作表
    // ========================================
    if (oldSheet) {
      Logger.log('🗑️ 步驟 2: 刪除舊工作表...');
      ss.deleteSheet(oldSheet);
      Logger.log('   ✅ 已刪除舊工作表');
      Logger.log('');
    }
    
    // ========================================
    // 步驟 3: 建立新工作表（正確的標題列）
    // ========================================
    Logger.log('📝 步驟 3: 建立新工作表...');
    
    const newSheet = ss.insertSheet('月薪資記錄');
    
    // ⭐⭐⭐ 完整正確的標題列（48 個欄位）
    const headers = [
      // A-I: 基本資訊 (9 欄)
      '薪資單ID',        // A
      '員工ID',          // B
      '員工姓名',        // C
      '年月',            // D
      '薪資類型',        // E
      '工時類型',        // F
      '時薪',            // G
      '工作時數',        // H
      '總加班時數',      // I

      // J-Y: 收入項目 (16 欄)
      '基本薪資',        // J
      '職務加給',        // K
      '伙食費',          // L
      '交通補助',        // M
      '全勤獎金',        // N
      '績效獎金',        // O
      '其他津貼1',       // P
      '其他津貼2',       // Q
      '其他津貼3',       // R
      '平日加班費',      // S
      '休息日加班費',    // T
      '國定假日加班費',  // U
      '未休假補薪',      // V
      '未休假天數',      // W
      '月休補薪',        // X
      '未休月休天數',    // Y

      // Z-AE: 法定扣款 (6 欄)
      '勞保費',          // Z
      '健保費',          // AA
      '就業保險費',      // AB
      '勞退自提率(%)',   // AC
      '勞退自提',        // AD
      '所得稅',          // AE

      // AF-AO: 其他扣款 (10 欄) ⭐⭐⭐ 關鍵部分
      '請假扣款',        // AF（總計）
      '病假時數',        // AG ⭐ 新增
      '病假扣款',        // AH ⭐ 新增
      '事假時數',        // AI ⭐ 新增
      '事假扣款',        // AJ ⭐ 新增
      '福利金扣款',      // AK
      '宿舍費用',        // AL
      '團保費用',        // AM
      '其他扣款1',       // AN
      '其他扣款2',       // AO

      // AP-AQ: 總額 (2 欄)
      '應發總額',        // AP
      '實發金額',        // AQ

      // AR-AS: 銀行 (2 欄)
      '銀行代碼',        // AR
      '銀行帳號',        // AS

      // AT-AV: 系統 (3 欄)
      '狀態',            // AT
      '備註',            // AU
      '建立時間'         // AV
    ];
    
    Logger.log(`   ✅ 建立標題列（共 ${headers.length} 個欄位）`);
    
    // 寫入標題列
    newSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // 美化標題列
    const headerRange = newSheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#10b981');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    headerRange.setHorizontalAlignment('center');
    headerRange.setVerticalAlignment('middle');
    headerRange.setWrap(true);
    
    // 凍結標題列
    newSheet.setFrozenRows(1);
    
    // 設定欄寬
    newSheet.setColumnWidth(1, 250);  // 薪資單ID
    newSheet.setColumnWidth(2, 200);  // 員工ID
    newSheet.setColumnWidth(3, 120);  // 員工姓名
    newSheet.setColumnWidth(4, 80);   // 年月
    
    Logger.log('   ✅ 標題列格式化完成');
    Logger.log('');
    
    // ========================================
    // 步驟 4: 恢復資料（如果有備份）
    // ========================================
    if (backupData.length > 1) {
      Logger.log('📥 步驟 4: 恢復資料...');
      
      // 跳過舊的標題列，只恢復資料行
      const dataRows = backupData.slice(1);
      
      Logger.log(`   📊 準備恢復 ${dataRows.length} 筆資料`);
      
      // ⚠️ 注意：舊資料可能欄位數不同，需要調整
      const restoredRows = dataRows.map(oldRow => {
        // 建立新的 48 欄位陣列
        const newRow = new Array(headers.length).fill('');
        
        // 複製舊資料到對應位置
        for (let i = 0; i < Math.min(oldRow.length, headers.length); i++) {
          newRow[i] = oldRow[i];
        }
        
        return newRow;
      });
      
      if (restoredRows.length > 0) {
        newSheet.getRange(2, 1, restoredRows.length, headers.length).setValues(restoredRows);
        Logger.log(`   ✅ 已恢復 ${restoredRows.length} 筆資料`);
      }
      
      Logger.log('');
    } else {
      Logger.log('⏭️ 步驟 4: 沒有資料需要恢復');
      Logger.log('');
    }
    
    // ========================================
    // 完成
    // ========================================
    Logger.log('═══════════════════════════════════════');
    Logger.log('✅✅✅ 月薪資記錄工作表重建完成！');
    Logger.log('═══════════════════════════════════════');
    Logger.log('');
    Logger.log('📋 新工作表資訊:');
    Logger.log(`   名稱: 月薪資記錄`);
    Logger.log(`   欄位數: ${headers.length}`);
    Logger.log(`   資料筆數: ${newSheet.getLastRow() - 1}`);
    Logger.log('');
    Logger.log('🔍 重要欄位位置:');
    Logger.log('   AG (第 32 欄): 病假時數');
    Logger.log('   AH (第 33 欄): 病假扣款');
    Logger.log('   AI (第 34 欄): 事假時數');
    Logger.log('   AJ (第 35 欄): 事假扣款');
    Logger.log('');
    Logger.log('⚠️ 下一步：重新執行薪資計算，讓資料寫入新 Sheet');
    Logger.log('');
    
    return {
      success: true,
      message: '月薪資記錄工作表重建成功',
      totalColumns: headers.length,
      restoredRows: backupData.length > 1 ? backupData.length - 1 : 0
    };
    
  } catch (error) {
    Logger.log('❌ 重建工作表失敗: ' + error.message);
    Logger.log('❌ 錯誤堆疊: ' + error.stack);
    
    return {
      success: false,
      message: '重建失敗: ' + error.message
    };
  }
}

/**
 * 🔍 檢查月薪資記錄工作表結構
 */
function checkSalarySheetStructure() {
  try {
    Logger.log('🔍 檢查月薪資記錄工作表結構...');
    Logger.log('');
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('月薪資記錄');
    
    if (!sheet) {
      Logger.log('❌ 找不到「月薪資記錄」工作表');
      return;
    }
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    Logger.log(`📊 工作表資訊:`);
    Logger.log(`   總欄位數: ${headers.length}`);
    Logger.log(`   總資料筆數: ${sheet.getLastRow() - 1}`);
    Logger.log('');
    
    Logger.log('📋 標題列清單:');
    headers.forEach((header, index) => {
      const colLetter = getColumnLetter(index + 1);
      Logger.log(`   ${colLetter} (第 ${index + 1} 欄): ${header}`);
    });
    
    Logger.log('');
    Logger.log('🔍 檢查關鍵欄位:');
    
    const checkFields = [
      { name: '病假時數', expectedCol: 'AG' },
      { name: '病假扣款', expectedCol: 'AH' },
      { name: '事假時數', expectedCol: 'AI' },
      { name: '事假扣款', expectedCol: 'AJ' }
    ];
    
    checkFields.forEach(field => {
      const index = headers.indexOf(field.name);
      if (index !== -1) {
        const actualCol = getColumnLetter(index + 1);
        const status = actualCol === field.expectedCol ? '✅' : '⚠️';
        Logger.log(`   ${status} ${field.name}: 在第 ${actualCol} 欄 (預期: ${field.expectedCol})`);
      } else {
        Logger.log(`   ❌ ${field.name}: 找不到`);
      }
    });
    
    Logger.log('');
    
  } catch (error) {
    Logger.log('❌ 檢查失敗: ' + error.message);
  }
}

/**
 * 輔助函數：數字轉欄位字母
 */
function getColumnLetter(columnNumber) {
  let temp;
  let letter = '';
  
  while (columnNumber > 0) {
    temp = (columnNumber - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    columnNumber = (columnNumber - temp - 1) / 26;
  }
  
  return letter;
}

/**
 * 🧪 完整測試流程
 */
function testCompleteFlow() {
  Logger.log('🧪 開始完整測試流程...');
  Logger.log('');
  
  // 1. 重建工作表
  Logger.log('📝 步驟 1: 重建工作表');
  const rebuildResult = rebuildMonthlySalarySheetComplete();
  
  if (!rebuildResult.success) {
    Logger.log('❌ 重建失敗，中止測試');
    return;
  }
  
  Logger.log('');
  
  // 2. 檢查結構
  Logger.log('📝 步驟 2: 檢查工作表結構');
  checkSalarySheetStructure();
  
  Logger.log('');
  
  // 3. 重新計算薪資
  Logger.log('📝 步驟 3: 重新計算薪資');
  
  const employeeId = 'Uf664a35632b736301d674d8b2cc3f8c0';
  const yearMonth = '2026-01';
  
  const calcResult = calculateMonthlySalary(employeeId, yearMonth);
  
  if (!calcResult.ok) {
    Logger.log('❌ 薪資計算失敗');
    return;
  }
  
  Logger.log('   ✅ 薪資計算成功');
  Logger.log(`   病假時數: ${calcResult.data.sickLeaveHours}`);
  Logger.log(`   病假扣款: ${calcResult.data.sickLeaveDeduction}`);
  
  Logger.log('');
  
  // 4. 存檔
  Logger.log('📝 步驟 4: 存檔到工作表');
  
  const saveResult = saveMonthlySalary(calcResult.data);
  
  if (!saveResult.success) {
    Logger.log('❌ 存檔失敗');
    return;
  }
  
  Logger.log('   ✅ 存檔成功');
  
  Logger.log('');
  
  // 5. 讀取驗證
  Logger.log('📝 步驟 5: 讀取驗證');
  
  const readResult = getMySalary(employeeId, yearMonth);
  
  if (!readResult.success) {
    Logger.log('❌ 讀取失敗');
    return;
  }
  
  Logger.log('   ✅ 讀取成功');
  Logger.log(`   病假時數: ${readResult.data['病假時數']}`);
  Logger.log(`   病假扣款: ${readResult.data['病假扣款']}`);
  Logger.log(`   事假時數: ${readResult.data['事假時數']}`);
  Logger.log(`   事假扣款: ${readResult.data['事假扣款']}`);
  
  Logger.log('');
  
  // 6. 最終驗證
  if (readResult.data['病假時數'] === 7 && 
      readResult.data['病假扣款'] === 773 &&
      readResult.data['事假時數'] === 6 &&
      readResult.data['事假扣款'] === 1325) {
    Logger.log('═══════════════════════════════════════');
    Logger.log('✅✅✅ 完整測試成功！所有數據正確！');
    Logger.log('═══════════════════════════════════════');
  } else {
    Logger.log('❌ 數據仍有問題，請檢查');
  }
}

/**
 * ============================================================
 * 📋 增強版打卡記錄除錯工具
 * ============================================================
 * 用途：診斷打卡記錄 Sheet 的資料結構問題
 * 
 * 使用方法：
 * 1. 將此程式碼貼到 Google Apps Script 編輯器
 * 2. 執行 debugPunchRecordsEnhanced()
 * 3. 查看執行日誌，找出問題所在
 */

function debugPunchRecordsEnhanced() {
  const employeeId = 'Uf664a35632b736301d674d8b2cc3f8c0';
  const yearMonth = '2026-01';
  
  Logger.log('🔍 ===== 開始除錯打卡記錄 =====');
  Logger.log('   員工ID: ' + employeeId);
  Logger.log('   年月: ' + yearMonth);
  Logger.log('');
  
  // ==================== 步驟 1：檢查 Sheet 是否存在 ====================
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('打卡紀錄');
  
  if (!sheet) {
    Logger.log('❌ 錯誤：找不到「打卡紀錄」工作表！');
    return;
  }
  
  Logger.log('✅ 找到「打卡紀錄」工作表');
  Logger.log('');
  
  // ==================== 步驟 2：讀取並顯示標題列 ====================
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  Logger.log('📋 標題列（共 ' + headers.length + ' 欄）:');
  headers.forEach((header, index) => {
    Logger.log('   欄位 ' + (index + 1) + ': "' + header + '"');
  });
  Logger.log('');
  
  // ==================== 步驟 3：找出關鍵欄位的位置 ====================
  const colIndex = {
    userId: headers.indexOf('員工ID') + 1,
    date: headers.indexOf('日期') + 1,
    punchType: headers.indexOf('打卡類別') + 1,
    time: headers.indexOf('時間') + 1,
    location: headers.indexOf('地點') + 1,
    note: headers.indexOf('備註') + 1
  };
  
  Logger.log('🔎 關鍵欄位位置:');
  for (const [field, col] of Object.entries(colIndex)) {
    if (col === 0) {
      Logger.log('   ❌ 找不到「' + field + '」欄位');
    } else {
      Logger.log('   ✅ ' + field + ': 第 ' + col + ' 欄');
    }
  }
  Logger.log('');
  
  // ==================== 步驟 4：讀取所有資料並篩選 ====================
  const dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
  const data = dataRange.getValues();
  
  Logger.log('📊 工作表總資料筆數: ' + data.length);
  Logger.log('');
  
  // ==================== 步驟 5：顯示前 5 筆資料樣本 ====================
  Logger.log('📝 前 5 筆資料樣本:');
  data.slice(0, 5).forEach((row, index) => {
    Logger.log('   第 ' + (index + 2) + ' 列:');
    row.forEach((cell, colNum) => {
      Logger.log('      欄 ' + (colNum + 1) + ' (' + headers[colNum] + '): ' + cell);
    });
    Logger.log('');
  });
  
  // ==================== 步驟 6：篩選符合條件的記錄 ====================
  Logger.log('🔍 開始篩選記錄...');
  Logger.log('   篩選條件:');
  Logger.log('   - 員工ID = "' + employeeId + '"');
  Logger.log('   - 日期包含 "' + yearMonth + '"');
  Logger.log('');
  
  const matchedRecords = [];
  
  data.forEach((row, index) => {
    const rowUserId = String(row[colIndex.userId - 1] || '').trim();
    const rowDate = row[colIndex.date - 1];
    const rowPunchType = String(row[colIndex.punchType - 1] || '').trim();
    
    // 將 Date 物件轉換為字串
    let rowDateStr = '';
    if (rowDate instanceof Date) {
      rowDateStr = Utilities.formatDate(rowDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    } else {
      rowDateStr = String(rowDate || '');
    }
    
    // 詳細日誌
    if (index < 10 || rowDateStr.indexOf(yearMonth) >= 0) { // 顯示前10筆或符合月份的
      Logger.log('   第 ' + (index + 2) + ' 列檢查:');
      Logger.log('      員工ID: "' + rowUserId + '" (符合: ' + (rowUserId === employeeId) + ')');
      Logger.log('      日期: "' + rowDateStr + '" (包含 ' + yearMonth + ': ' + (rowDateStr.indexOf(yearMonth) >= 0) + ')');
      Logger.log('      打卡類別: "' + rowPunchType + '"');
    }
    
    // 符合條件
    if (rowUserId === employeeId && rowDateStr.indexOf(yearMonth) >= 0) {
      matchedRecords.push({
        row: index + 2,
        userId: rowUserId,
        date: rowDateStr,
        punchType: rowPunchType,
        time: row[colIndex.time - 1],
        location: row[colIndex.location - 1],
        note: row[colIndex.note - 1]
      });
    }
  });
  
  // ==================== 步驟 7：顯示結果 ====================
  Logger.log('');
  Logger.log('✅ 篩選完成！');
  Logger.log('   符合條件的記錄數: ' + matchedRecords.length);
  Logger.log('');
  
  if (matchedRecords.length > 0) {
    Logger.log('📋 符合條件的記錄:');
    matchedRecords.forEach((record, index) => {
      Logger.log('   第 ' + (index + 1) + ' 筆 (Sheet 第 ' + record.row + ' 列):');
      Logger.log('      日期: ' + record.date);
      Logger.log('      打卡類別: ' + record.punchType);
      Logger.log('      時間: ' + record.time);
      Logger.log('      地點: ' + record.location);
      Logger.log('      備註: ' + record.note);
      Logger.log('');
    });
  } else {
    Logger.log('❌ 沒有找到符合條件的記錄！');
    Logger.log('');
    Logger.log('🔍 可能的原因:');
    Logger.log('   1. 員工ID 不匹配（注意大小寫、空格）');
    Logger.log('   2. 日期格式不符（需要包含 "2026-01"）');
    Logger.log('   3. 資料在不同的工作表');
    Logger.log('   4. 欄位名稱不正確');
  }
  
  Logger.log('');
  Logger.log('🔍 ===== 除錯完成 =====');
}

/**
 * ✅ 修正「月薪資記錄」工作表的銀行代碼
 * 
 * 執行此函數會：
 * 1. 讀取 AR 欄（第 45 欄）的所有銀行代碼
 * 2. 將每個代碼標準化為 '007 格式
 * 3. 寫回工作表並設定為文字格式
 */
function fixBankCodeInSalarySheet() {
  try {
    Logger.log('🔧 開始修正「月薪資記錄」的銀行代碼...');
    Logger.log('');
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('月薪資記錄');
    
    if (!sheet) {
      Logger.log('❌ 找不到「月薪資記錄」工作表');
      return { success: false, message: '找不到工作表' };
    }
    
    const lastRow = sheet.getLastRow();
    
    if (lastRow <= 1) {
      Logger.log('⚠️ 工作表沒有資料');
      return { success: true, message: '沒有資料需要修正' };
    }
    
    Logger.log(`📊 工作表資料筆數: ${lastRow - 1}`);
    
    const bankCodeCol = 45;  // AR 欄
    const range = sheet.getRange(2, bankCodeCol, lastRow - 1, 1);
    const values = range.getValues();
    
    Logger.log('');
    Logger.log('📋 開始處理銀行代碼...');
    Logger.log('');
    
    let fixedCount = 0;
    let skippedCount = 0;
    
    const fixedValues = values.map((row, index) => {
      let bankCode = row[0];
      
      Logger.log(`處理第 ${index + 2} 行:`);
      Logger.log(`   原始值: "${bankCode}" (類型: ${typeof bankCode})`);
      
      // 如果是空值，保持空白
      if (!bankCode || bankCode === '') {
        Logger.log(`   → 保持空白`);
        skippedCount++;
        return [""];
      }
      
      // ⭐⭐⭐ 處理日期物件：1902/3/20 → 提取年份 → 902
      if (bankCode instanceof Date) {
        const year = bankCode.getFullYear();  // 1902
        bankCode = String(year).substring(1);  // "902"
        Logger.log(`   → 檢測到 Date 對象，提取年份: "${bankCode}"`);
      } else {
        // 移除所有非數字字符
        bankCode = String(bankCode).replace(/[^0-9]/g, '');
        Logger.log(`   → 移除非數字字符: "${bankCode}"`);
      }
      
      // 補零到 3 位數
      if (bankCode.length > 0) {
        bankCode = bankCode.padStart(3, '0');
        Logger.log(`   → 補零: "${bankCode}"`);
      }
      
      // 加上單引號前綴
      const finalCode = "'" + bankCode;
      Logger.log(`   → 最終值: "${finalCode}"`);
      Logger.log('');
      
      fixedCount++;
      return [finalCode];
    });
    
    // 寫回工作表
    Logger.log('💾 寫回工作表...');
    
    // 先設定為文字格式
    range.setNumberFormat('@');
    
    // 再寫入數值
    range.setValues(fixedValues);
    
    Logger.log('');
    Logger.log('═══════════════════════════════════════');
    Logger.log('✅ 修正完成！');
    Logger.log('═══════════════════════════════════════');
    Logger.log(`   處理筆數: ${fixedCount}`);
    Logger.log(`   跳過筆數: ${skippedCount}`);
    Logger.log('');
    Logger.log('📋 範例結果:');
    Logger.log('   "1902/3/20" → \'902');
    Logger.log('   "7" → \'007');
    Logger.log('   "810 wgsgsgs" → \'810');
    Logger.log('');
    
    return {
      success: true,
      message: `修正完成（${fixedCount} 筆）`,
      fixedCount: fixedCount,
      skippedCount: skippedCount
    };
    
  } catch (error) {
    Logger.log('❌ 修正失敗: ' + error.message);
    Logger.log('   錯誤堆疊: ' + error.stack);
    return { success: false, message: error.message };
  }
}

/**
 * ✅ 修正「員工薪資設定」工作表的銀行代碼
 * 
 * 執行此函數會：
 * 1. 讀取 P 欄（第 16 欄）的所有銀行代碼
 * 2. 將每個代碼標準化為 '007 格式
 * 3. 寫回工作表並設定為文字格式
 */
function fixBankCodeInConfigSheet() {
  try {
    Logger.log('🔧 開始修正「員工薪資設定」的銀行代碼...');
    Logger.log('');
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('員工薪資設定');
    
    if (!sheet) {
      Logger.log('❌ 找不到「員工薪資設定」工作表');
      return { success: false, message: '找不到工作表' };
    }
    
    const lastRow = sheet.getLastRow();
    
    if (lastRow <= 1) {
      Logger.log('⚠️ 工作表沒有資料');
      return { success: true, message: '沒有資料需要修正' };
    }
    
    Logger.log(`📊 工作表資料筆數: ${lastRow - 1}`);
    
    const bankCodeCol = 16;  // P 欄
    const range = sheet.getRange(2, bankCodeCol, lastRow - 1, 1);
    const values = range.getValues();
    
    Logger.log('');
    Logger.log('📋 開始處理銀行代碼...');
    Logger.log('');
    
    let fixedCount = 0;
    let skippedCount = 0;
    
    const fixedValues = values.map((row, index) => {
      let bankCode = row[0];
      
      Logger.log(`處理第 ${index + 2} 行:`);
      Logger.log(`   原始值: "${bankCode}" (類型: ${typeof bankCode})`);
      
      // 如果是空值，保持空白
      if (!bankCode || bankCode === '') {
        Logger.log(`   → 保持空白`);
        skippedCount++;
        return [""];
      }
      
      // ⭐⭐⭐ 處理日期物件：1902/3/20 → 提取年份 → 902
      if (bankCode instanceof Date) {
        const year = bankCode.getFullYear();  // 1902
        bankCode = String(year).substring(1);  // "902"
        Logger.log(`   → 檢測到 Date 對象，提取年份: "${bankCode}"`);
      } else {
        // 移除所有非數字字符
        bankCode = String(bankCode).replace(/[^0-9]/g, '');
        Logger.log(`   → 移除非數字字符: "${bankCode}"`);
      }
      
      // 補零到 3 位數
      if (bankCode.length > 0) {
        bankCode = bankCode.padStart(3, '0');
        Logger.log(`   → 補零: "${bankCode}"`);
      }
      
      // 加上單引號前綴
      const finalCode = "'" + bankCode;
      Logger.log(`   → 最終值: "${finalCode}"`);
      Logger.log('');
      
      fixedCount++;
      return [finalCode];
    });
    
    // 寫回工作表
    Logger.log('💾 寫回工作表...');
    
    // 先設定為文字格式
    range.setNumberFormat('@');
    
    // 再寫入數值
    range.setValues(fixedValues);
    
    Logger.log('');
    Logger.log('═══════════════════════════════════════');
    Logger.log('✅ 修正完成！');
    Logger.log('═══════════════════════════════════════');
    Logger.log(`   處理筆數: ${fixedCount}`);
    Logger.log(`   跳過筆數: ${skippedCount}`);
    Logger.log('');
    Logger.log('📋 範例結果:');
    Logger.log('   "1902/3/20" → \'902');
    Logger.log('   "7" → \'007');
    Logger.log('   "810 wgsgsgs" → \'810');
    Logger.log('');
    
    return {
      success: true,
      message: `修正完成（${fixedCount} 筆）`,
      fixedCount: fixedCount,
      skippedCount: skippedCount
    };
    
  } catch (error) {
    Logger.log('❌ 修正失敗: ' + error.message);
    Logger.log('   錯誤堆疊: ' + error.stack);
    return { success: false, message: error.message };
  }
}

/**
 * ✅ API：更新月薪資記錄的工時欄位（前端員工薪酬 tab 專用）
 * 
 * 接收參數：
 *   employeeId       - 員工ID（從 session 取得，不接受前端傳入）
 *   yearMonth        - 年月 (YYYY-MM)
 *   totalWorkHours   - 總工作時數
 *   totalOvertimeHours - 加班時數
 *   attendanceDays   - 出勤天數（選填）
 */
function updateMonthlySalaryWorkHoursAPI() {
  try {
    const session = checkSessionInternal();
    if (!session.ok) {
      return jsonResponse(false, null, 'SESSION_INVALID', 'SESSION_INVALID');
    }

    // ⭐ 員工ID 從 session 取，不信任前端傳入
    const employeeId    = session.user.userId;
    const yearMonth     = getParam('yearMonth');
    const totalWorkHours    = parseFloat(getParam('totalWorkHours')) || 0;
    const totalOvertimeHours = parseFloat(getParam('totalOvertimeHours')) || 0;
    const attendanceDays     = parseFloat(getParam('attendanceDays')) || 0;

    if (!yearMonth) {
      return jsonResponse(false, null, '缺少 yearMonth 參數', 'MISSING_PARAM');
    }

    Logger.log(`📝 updateMonthlySalaryWorkHours`);
    Logger.log(`   員工: ${employeeId}`);
    Logger.log(`   年月: ${yearMonth}`);
    Logger.log(`   工作時數: ${totalWorkHours}h`);
    Logger.log(`   加班時數: ${totalOvertimeHours}h`);
    Logger.log(`   出勤天數: ${attendanceDays}`);

    const sheet = getMonthlySalarySheetEnhanced();
    const data  = sheet.getDataRange().getValues();
    const headers = data[0];

    const salaryId = `SAL-${yearMonth}-${employeeId}`;

    // 找對應欄位索引
    const colWorkHours  = headers.indexOf('工作時數');    // H欄 (索引7)
    const colOtHours    = headers.indexOf('總加班時數');  // I欄 (索引8)

    if (colWorkHours === -1 || colOtHours === -1) {
      return jsonResponse(false, null, '找不到工時欄位', 'COLUMN_NOT_FOUND');
    }

    // 找到對應薪資單
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === salaryId) {
        sheet.getRange(i + 1, colWorkHours + 1).setValue(totalWorkHours);
        sheet.getRange(i + 1, colOtHours + 1).setValue(totalOvertimeHours);

        Logger.log(`✅ 已更新薪資單 ${salaryId} 的工時欄位`);

        return jsonResponse(true, {
          salaryId: salaryId,
          totalWorkHours: totalWorkHours,
          totalOvertimeHours: totalOvertimeHours
        }, '工時更新成功');
      }
    }

    // 找不到薪資單 → 先建立再更新
    Logger.log(`⚠️ 找不到薪資單 ${salaryId}，嘗試先同步建立...`);
    const syncResult = syncSalaryToMonthlyRecord(employeeId, yearMonth);

    if (!syncResult.success) {
      return jsonResponse(false, null, '找不到薪資記錄，且建立失敗：' + syncResult.message, 'RECORD_NOT_FOUND');
    }

    // 建立後再找一次並更新
    const data2 = sheet.getDataRange().getValues();
    for (let i = 1; i < data2.length; i++) {
      if (data2[i][0] === salaryId) {
        sheet.getRange(i + 1, colWorkHours + 1).setValue(totalWorkHours);
        sheet.getRange(i + 1, colOtHours + 1).setValue(totalOvertimeHours);

        Logger.log(`✅ 建立並更新薪資單 ${salaryId}`);

        return jsonResponse(true, {
          salaryId: salaryId,
          totalWorkHours: totalWorkHours,
          totalOvertimeHours: totalOvertimeHours,
          created: true
        }, '工時更新成功（已自動建立薪資記錄）');
      }
    }

    return jsonResponse(false, null, '建立薪資記錄後仍找不到對應列', 'UNEXPECTED_ERROR');

  } catch (error) {
    Logger.log('❌ updateMonthlySalaryWorkHoursAPI 錯誤: ' + error);
    return jsonResponse(false, null, error.toString(), 'ERROR');
  }
}


function recalculateMonthlySalaryAPI(params) {
  try {
    // ⭐ 改用傳入的 params，不再依賴 getParam() / checkSessionInternal()
    if (!params || !params.token) {
      return { ok: false, msg: '缺少 token', code: 'MISSING_TOKEN' };
    }

    if (!validateSession(params.token)) {
      return { ok: false, msg: '未授權或 session 已過期', code: 'UNAUTHORIZED' };
    }

    const sessionResult = handleCheckSession(params.token);
    if (!sessionResult.ok || !sessionResult.user) {
      return { ok: false, msg: 'Session 資料無效', code: 'SESSION_INVALID' };
    }

    if (sessionResult.user.dept !== '管理員') {
      return { ok: false, msg: '需要管理員權限', code: 'PERMISSION_DENIED' };
    }

    const yearMonth  = params.yearMonth;
    const employeeId = params.employeeId;

    if (!yearMonth) {
      return { ok: false, msg: '缺少 yearMonth 參數', code: 'MISSING_PARAM' };
    }

    Logger.log(`🔄 重新計算薪資: ${employeeId || '未指定'}, ${yearMonth}`);

    if (employeeId) {
      const result = _recalcOne(employeeId, yearMonth);
      if (!result.success) {
        return { ok: false, msg: result.message || '計算失敗', code: 'CALC_ERROR' };
      }
      return { ok: true, data: result.data, msg: '重新計算成功' };
    }

    return { ok: false, msg: '請提供 employeeId', code: 'MISSING_EMPLOYEE_ID' };

  } catch (error) {
    Logger.log('❌ recalculateMonthlySalaryAPI 錯誤: ' + error);
    return { ok: false, msg: error.toString(), code: 'ERROR' };
  }
}
// ✅ 修正後的寫法
function _recalcOne(employeeId, yearMonth) {
  // 直接用統一路由，不再自己判斷類型
  const calcResult = calculateMonthlySalary(employeeId, yearMonth);

  if (!calcResult.success && !calcResult.ok) {
    return { success: false, message: calcResult.message || calcResult.msg };
  }

  const saveResult = saveMonthlySalary(calcResult.data);
  if (!saveResult.success) {
    return { success: false, message: '計算成功但儲存失敗：' + saveResult.message };
  }

  return { success: true, data: calcResult.data };
}

function getAllEmployeeSalaryConfigAPI(params) {
  try {
    if (!params.token || !validateSession(params.token)) {
      return { ok: false, msg: '未授權' };
    }
    
    const sheet = getEmployeeSalarySheet();
    const data = sheet.getDataRange().getValues();
    
    const employees = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][30] === '在職') { // A欄有ID且狀態為在職
        employees.push({
          employeeId: data[i][0],
          employeeName: data[i][1]
        });
      }
    }
    
    return { ok: true, data: employees };
    
  } catch (error) {
    return { ok: false, msg: error.toString() };
  }
}
/**
 * ✅ 建立備註歷程（保留一年內的記錄）
 */
function buildNoteHistory(employeeId, newNote) {
  try {
    const sheet = getEmployeeSalarySheet();
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const noteHistoryColIndex = headers.indexOf("備註歷程");

    // 取得舊的歷程
    let history = [];

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(employeeId).trim()) {
        if (noteHistoryColIndex !== -1 && data[i][noteHistoryColIndex]) {
          try {
            history = JSON.parse(data[i][noteHistoryColIndex]);
          } catch (e) {
            history = [];
          }
        }
        break;
      }
    }

    // 加入新備註（只有非空才記錄）
    if (newNote && newNote.trim() !== "") {
      history.push({
        note: newNote.trim(),
        time: Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy-MM-dd HH:mm")
      });
    }

    // 過濾掉一年前的記錄
    const oneYearAgo = new Date();
    oneYearAgo.setFullYear(oneYearAgo.getFullYear() - 1);

    history = history.filter(entry => {
      try {
        return new Date(entry.time) >= oneYearAgo;
      } catch (e) {
        return true;
      }
    });

    return JSON.stringify(history);
  } catch (e) {
    Logger.log("⚠️ buildNoteHistory 失敗: " + e);
    return "[]";
  }
}

function verifyPersonalLeaveDeduction() {
  const employeeId = 'Uf664a35632b736301d674d8b2cc3f8c0'; // ⚠️ 換成實際員工ID
  const yearMonth = '2026-03'; // ⚠️ 換成有請假的月份

  // 1. 直接呼叫請假扣款計算
  const config = getEmployeeSalaryTW(employeeId);
  const baseSalary = parseFloat(config.data['基本薪資']) || 0;
  const mealAllowance = parseFloat(config.data['伙食費']) || 0;
  const transportAllowance = parseFloat(config.data['交通補助']) || 0;

  const leaveResult = calculateLeaveDeduction(employeeId, yearMonth, baseSalary, mealAllowance, transportAllowance);
  
  Logger.log('=== 請假扣款計算結果 ===');
  Logger.log('病假時數: ' + leaveResult.sickLeaveHours);
  Logger.log('病假扣款: ' + leaveResult.sickLeaveDeduction);
  Logger.log('事假時數: ' + leaveResult.personalLeaveHours);
  Logger.log('事假扣款: ' + leaveResult.personalLeaveDeduction);
  Logger.log('總扣款: ' + leaveResult.totalDeduction);

  // 2. 確認已寫入月薪資記錄
  const salary = getMySalary(employeeId, yearMonth);
  Logger.log('');
  Logger.log('=== 月薪資記錄驗證 ===');
  Logger.log('事假時數(AI欄): ' + salary.data['事假時數']);
  Logger.log('事假扣款(AJ欄): ' + salary.data['事假扣款']);
  Logger.log('請假扣款合計(AF欄): ' + salary.data['請假扣款']);
  Logger.log('實發金額: ' + salary.data['實發金額']);

  // 3. 驗證
  if (leaveResult.personalLeaveHours > 0 && salary.data['事假扣款'] > 0) {
    Logger.log('');
    Logger.log('✅ 驗證成功！事假扣款已正確納入薪資');
  } else if (leaveResult.personalLeaveHours > 0 && salary.data['事假扣款'] === 0) {
    Logger.log('');
    Logger.log('❌ 問題：有事假但薪資記錄扣款為0，需要重新執行薪資計算');
  } else {
    Logger.log('');
    Logger.log('⚠️ 該月份無事假記錄');
  }
}


function fixAndVerify() {
  const employeeId = 'Uf664a35632b736301d674d8b2cc3f8c0';
  const yearMonth = '2026-03';
  
  // ⭐ 重算並儲存
  const result = _recalcOne(employeeId, yearMonth);
  Logger.log('重算: ' + (result.success ? '成功' : result.message));
  
  if (!result.success) return;
  
  // 讀回驗證
  const salary = getMySalary(employeeId, yearMonth);
  Logger.log('事假時數: ' + salary.data['事假時數']);
  Logger.log('事假扣款: ' + salary.data['事假扣款']);
  Logger.log('實發金額: ' + salary.data['實發金額']);
}
function batchRecalcAllForMarch() {
  const yearMonth = '2026-03';
  const result = batchCalculateSalary(yearMonth);
  Logger.log('成功: ' + result.data.success.length + ' 人');
  Logger.log('失敗: ' + result.data.failed.length + ' 人');
}
