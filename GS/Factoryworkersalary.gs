// ==================================================================================
// 情況二：食品廠移工（時薪制-標準工時）薪資計算系統
// ==================================================================================
//
// 🏭 適用對象：食品廠移工
// 💰 薪資類型：時薪制
// ⏰ 工時類型：標準工時
//
// ==================================================================================

/**
 * ✅ 情況二薪資計算規則
 * 
 * 1. 基本薪資：時薪 × 工作時數（從打卡記錄計算）
 * 
 * 2. 津貼：
 *    - 底薪（大）：固定津貼
 *    - 工作津貼：固定津貼
 * 
 * 3. 加班費（標準工時制）：
 *    - 延長工時加班費：底薪/30/8 × 1.34 × 時數（前2h加班至實際分鐘數）
 *    - 休息日（週六）+ 國定假日：底薪/30/8 × 時數（前8h勞動加班至實際分鐘數）
 * 
 * 4. 請假扣款：
 *    - 請假扣：底薪/30/8 × 小時數
 *    - 請病假：底薪/30/8 × 小時數/2
 * 
 * 5. 特休未休補薪：底薪/30 × 未休天數
 * 
 * 6. 其他扣款：
 *    - 代扣所得稅
 *    - 代扣仲介服務費
 *    - 代扣健檢費
 *    - 代扣代辦費
 * 
 * ==================================================================================
 */

/**
 * 🏭 計算食品廠移工月薪資（情況二專用）
 * 
 * @param {string} employeeId - 員工ID
 * @param {string} yearMonth - 年月 (YYYY-MM)
 * @returns {Object} 薪資計算結果
 */
function calculateFactoryWorkerSalary(employeeId, yearMonth) {
  try {
    Logger.log('═══════════════════════════════════════');
    Logger.log('🏭 開始計算食品廠移工薪資（情況二）');
    Logger.log(`   員工ID: ${employeeId}`);
    Logger.log(`   年月: ${yearMonth}`);
    Logger.log('═══════════════════════════════════════');
    Logger.log('');
    
    // ========================================
    // 步驟 1：取得員工薪資設定
    // ========================================
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
    
    // ========================================
    // 步驟 2：解析員工薪資設定
    // ========================================
    const employeeName = employeeConfig[1] || '';
    const employeeType = employeeConfig[3] || '正職';
    const salaryType = employeeConfig[4] || '時薪';
    const workTimeType = employeeConfig[5] || '標準工時';
    const hourlyRate = parseFloat(employeeConfig[6]) || 0;  // 時薪（非基本薪資）
    
    // 固定津貼
    const baseAllowanceLarge = parseFloat(employeeConfig[7]) || 0;  // 底薪（大）
    const workAllowance = parseFloat(employeeConfig[8]) || 0;       // 工作津貼
    
    // 銀行資訊
    const bankCode = employeeConfig[15] || '';
    const bankAccount = employeeConfig[16] || '';
    
    // 法定扣款
    const laborFee = parseFloat(employeeConfig[20]) || 0;
    const healthFee = parseFloat(employeeConfig[21]) || 0;
    const employmentFee = parseFloat(employeeConfig[22]) || 0;
    
    // 其他扣款
    const withholdingTax = parseFloat(employeeConfig[25]) || 0;        // 代扣所得稅
    const agencyFee = parseFloat(employeeConfig[26]) || 0;             // 代扣仲介服務費
    const healthCheckFee = parseFloat(employeeConfig[27]) || 0;        // 代扣健檢費
    const agencyHandlingFee = parseFloat(employeeConfig[28]) || 0;     // 代扣代辦費
    
    Logger.log('📋 基本薪資資料:');
    Logger.log(`   員工姓名: ${employeeName}`);
    Logger.log(`   員工類型: ${employeeType}`);
    Logger.log(`   薪資類型: ${salaryType}`);
    Logger.log(`   工時類型: ${workTimeType}`);
    Logger.log(`   時薪: ${hourlyRate}`);
    Logger.log(`   底薪（大）: ${baseAllowanceLarge}`);
    Logger.log(`   工作津貼: ${workAllowance}`);
    Logger.log('');
    
    // ========================================
    // 步驟 3：計算工作時數（從打卡記錄）
    // ========================================
    const attendanceRecords = getEmployeeMonthlyAttendanceInternal(employeeId, yearMonth);
    
    let totalWorkHours = 0;
    
    attendanceRecords.forEach(record => {
      if (record.workHours > 0) {
        totalWorkHours += record.workHours;
        Logger.log(`   ${record.date}: ${record.punchIn} ~ ${record.punchOut} = ${record.workHours.toFixed(2)}h`);
      }
    });
    
    Logger.log('');
    Logger.log(`⏱️ 總工作時數: ${totalWorkHours.toFixed(2)}h`);
    Logger.log('');
    
    // ========================================
    // 步驟 4：計算基本薪資（時薪 × 工時）
    // ========================================
    const baseSalary = Math.round(hourlyRate * totalWorkHours);
    
    Logger.log(`💰 基本薪資計算:`);
    Logger.log(`   時薪 ${hourlyRate} × 工時 ${totalWorkHours.toFixed(2)} = ${baseSalary} 元`);
    Logger.log('');
    
    // ========================================
    // 步驟 5：計算加班費（情況二專用規則）
    // ========================================
    const overtimeResult = calculateFactoryWorkerOvertime(
      employeeId, 
      yearMonth, 
      baseAllowanceLarge
    );
    
    const extendedOvertimePay = overtimeResult.extendedOvertimePay;      // 延長工時加班費
    const restdayOvertimePay = overtimeResult.restdayOvertimePay;        // 休息日加班費
    const holidayOvertimePay = overtimeResult.holidayOvertimePay;        // 國定假日加班費
    const totalOvertimeHours = overtimeResult.totalOvertimeHours;
    
    Logger.log('💼 加班費計算結果:');
    Logger.log(`   延長工時加班費: ${extendedOvertimePay}`);
    Logger.log(`   休息日加班費: ${restdayOvertimePay}`);
    Logger.log(`   國定假日加班費: ${holidayOvertimePay}`);
    Logger.log(`   總加班時數: ${totalOvertimeHours}`);
    Logger.log('');
    
    // ========================================
    // 步驟 6：計算請假扣款（情況二專用規則）
    // ========================================
    const leaveDeductionResult = calculateFactoryWorkerLeaveDeduction(
      employeeId, 
      yearMonth, 
      baseAllowanceLarge
    );
    
    const sickLeaveHours = leaveDeductionResult.sickLeaveHours;
    const sickLeaveDeduction = leaveDeductionResult.sickLeaveDeduction;
    const personalLeaveHours = leaveDeductionResult.personalLeaveHours;
    const personalLeaveDeduction = leaveDeductionResult.personalLeaveDeduction;
    const totalLeaveDeduction = leaveDeductionResult.totalDeduction;
    
    Logger.log('🏥 請假扣款:');
    Logger.log(`   病假: ${sickLeaveHours}h → 扣款 ${sickLeaveDeduction} 元`);
    Logger.log(`   事假: ${personalLeaveHours}h → 扣款 ${personalLeaveDeduction} 元`);
    Logger.log(`   總扣款: ${totalLeaveDeduction} 元`);
    Logger.log('');
    
    // ========================================
    // 步驟 7：計算特休未休補薪（情況二專用規則）
    // ========================================
    const unusedLeavePayResult = calculateFactoryWorkerUnusedLeavePay(
      employeeId, 
      yearMonth, 
      baseAllowanceLarge
    );
    
    const unusedLeavePay = unusedLeavePayResult.amount;
    const unusedLeaveDays = unusedLeavePayResult.days;
    
    Logger.log('🏖️ 特休未休補薪:');
    Logger.log(`   未休天數: ${unusedLeaveDays} 天`);
    Logger.log(`   補薪金額: ${unusedLeavePay} 元`);
    Logger.log('');
    
    // ========================================
    // 步驟 8：計算應發總額
    // ========================================
    const grossSalary = 
      baseSalary +                  // 基本薪資（時薪×工時）
      baseAllowanceLarge +          // 底薪（大）
      workAllowance +               // 工作津貼
      extendedOvertimePay +         // 延長工時加班費
      restdayOvertimePay +          // 休息日加班費
      holidayOvertimePay +          // 國定假日加班費
      unusedLeavePay;               // 特休未休補薪
    
    Logger.log(`💵 應發總額: ${grossSalary}`);
    Logger.log('');
    
    // ========================================
    // 步驟 9：計算扣款總額
    // ========================================
    const totalDeductions = 
      laborFee +                    // 勞保費
      healthFee +                   // 健保費
      employmentFee +               // 就業保險費
      totalLeaveDeduction +         // 請假扣款
      withholdingTax +              // 代扣所得稅
      agencyFee +                   // 代扣仲介服務費
      healthCheckFee +              // 代扣健檢費
      agencyHandlingFee;            // 代扣代辦費
    
    Logger.log(`💸 扣款總額: ${totalDeductions}`);
    Logger.log('');
    
    // ========================================
    // 步驟 10：計算實發金額
    // ========================================
    const netSalary = grossSalary - totalDeductions;
    
    Logger.log(`✅ 實發金額: ${netSalary}`);
    Logger.log('');
    
    // ========================================
    // 步驟 11：返回完整薪資資料
    // ========================================
    const salaryData = {
      employeeId: employeeId,
      employeeName: employeeName,
      yearMonth: yearMonth,
      employeeType: employeeType,
      salaryType: salaryType,
      workTimeType: workTimeType,
      
      // 時薪與工時
      hourlyRate: hourlyRate,
      totalWorkHours: Math.floor(totalWorkHours),
      
      // 基本薪資與津貼
      baseSalary: baseSalary,
      baseAllowanceLarge: baseAllowanceLarge,      // 底薪（大）
      workAllowance: workAllowance,                // 工作津貼
      
      // 加班費
      extendedOvertimePay: extendedOvertimePay,    // 延長工時加班費（情況二原始欄位）
      weekdayOvertimePay: extendedOvertimePay,     // 標準欄位名稱別名，供通用顯示/儲存使用
      restdayOvertimePay: restdayOvertimePay,      // 休息日加班費
      holidayOvertimePay: holidayOvertimePay,      // 國定假日加班費
      totalOvertimeHours: totalOvertimeHours,
      
      // 補薪項目
      unusedLeavePay: unusedLeavePay,
      unusedLeaveDays: unusedLeaveDays,
      
      // 法定扣款
      laborFee: laborFee,
      healthFee: healthFee,
      employmentFee: employmentFee,
      
      // 請假扣款明細
      sickLeaveHours: sickLeaveHours,
      sickLeaveDeduction: sickLeaveDeduction,
      personalLeaveHours: personalLeaveHours,
      personalLeaveDeduction: personalLeaveDeduction,
      leaveDeduction: totalLeaveDeduction,
      
      // 其他扣款（情況二原始欄位）
      withholdingTax: withholdingTax,              // 代扣所得稅
      agencyFee: agencyFee,                        // 代扣仲介服務費
      healthCheckFee: healthCheckFee,              // 代扣健檢費
      agencyHandlingFee: agencyHandlingFee,        // 代扣代辦費
      // 標準欄位名稱別名，供通用顯示/儲存使用
      incomeTax: withholdingTax,                   // 對應所得稅欄位
      otherDeduction1: agencyFee,                  // 對應其他扣款1（仲介費）
      otherDeduction2: healthCheckFee + agencyHandlingFee, // 對應其他扣款2（健檢+代辦）
      
      // 總計
      grossSalary: grossSalary,
      totalDeductions: totalDeductions,
      netSalary: netSalary,
      
      // 銀行資訊
      bankCode: bankCode,
      bankAccount: bankAccount
    };
    
    Logger.log('═══════════════════════════════════════');
    Logger.log('✅✅✅ 食品廠移工薪資計算完成（情況二）');
    Logger.log('═══════════════════════════════════════');
    Logger.log('📊 薪資結構摘要:');
    Logger.log(`   基本薪資（時薪×工時）: ${baseSalary}`);
    Logger.log(`   底薪（大）: ${baseAllowanceLarge}`);
    Logger.log(`   工作津貼: ${workAllowance}`);
    Logger.log(`   延長工時加班費: ${extendedOvertimePay}`);
    Logger.log(`   休息日加班費: ${restdayOvertimePay}`);
    Logger.log(`   國定假日加班費: ${holidayOvertimePay}`);
    Logger.log(`   特休未休補薪: ${unusedLeavePay}`);
    Logger.log(`   應發總額: ${grossSalary}`);
    Logger.log(`   扣款總額: ${totalDeductions}`);
    Logger.log(`   實發金額: ${netSalary}`);
    Logger.log('═══════════════════════════════════════');
    Logger.log('');
    
    return {
      ok: true,
      success: true,
      data: salaryData
    };
    
  } catch (error) {
    Logger.log('');
    Logger.log('❌❌❌ calculateFactoryWorkerSalary 發生錯誤');
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
 * 🏭 計算食品廠移工加班費（情況二專用規則）
 * 
 * 規則：
 * 1. 延長工時加班費：底薪/30/8 × 1.34 × 時數（前2h加班至實際分鐘數）
 * 2. 休息日（週六）：底薪/30/8 × 時數（前8h勞動加班至實際分鐘數）
 * 3. 國定假日：底薪/30/8 × 時數（前8h勞動加班至實際分鐘數）
 * 
 * @param {string} employeeId - 員工ID
 * @param {string} yearMonth - 年月
 * @param {number} baseAllowanceLarge - 底薪（大）
 * @returns {Object} 加班費計算結果
 */
function calculateFactoryWorkerOvertime(employeeId, yearMonth, baseAllowanceLarge) {
  try {
    Logger.log('💼 開始計算食品廠移工加班費');
    
    // 計算時薪：底薪/30/8
    const hourlyRate = baseAllowanceLarge / 30 / 8;
    
    Logger.log(`   時薪: ${hourlyRate.toFixed(2)} 元/小時`);
    Logger.log(`   底薪（大）: ${baseAllowanceLarge}`);
    Logger.log('');
    
    // 取得加班記錄
    const overtimeRecords = getEmployeeMonthlyOvertime(employeeId, yearMonth);
    
    if (overtimeRecords.length === 0) {
      Logger.log('   ⚠️ 無加班記錄');
      return {
        extendedOvertimePay: 0,
        restdayOvertimePay: 0,
        holidayOvertimePay: 0,
        totalOvertimeHours: 0
      };
    }
    
    Logger.log(`   📋 找到 ${overtimeRecords.length} 筆加班記錄`);
    
    // 按日期分組
    const overtimeByDate = {};
    overtimeRecords.forEach(record => {
      const date = record.date;
      if (!overtimeByDate[date]) {
        overtimeByDate[date] = 0;
      }
      overtimeByDate[date] += parseFloat(record.hours) || 0;
    });
    
    let extendedOvertimePay = 0;   // 延長工時加班費（平日前2h）
    let restdayOvertimePay = 0;    // 休息日加班費（週六前8h）
    let holidayOvertimePay = 0;    // 國定假日加班費（前8h）
    let totalOvertimeHours = 0;
    
    // 遍歷每天的加班記錄
    Object.keys(overtimeByDate).forEach(date => {
      let dailyHours = overtimeByDate[date];
      const dateType = getDateType(date);
      
      Logger.log(`\n📅 ${date} (${dateType}): ${dailyHours.toFixed(2)}h`);
      
      if (dateType === 'weekday') {
        // ⭐ 平日延長工時：前2h × 1.34
        const maxHours = 2;
        const actualHours = Math.min(dailyHours, maxHours);
        const pay = Math.round(hourlyRate * actualHours * 1.34);
        
        extendedOvertimePay += pay;
        totalOvertimeHours += actualHours;
        
        Logger.log(`   延長工時: ${actualHours.toFixed(2)}h × 1.34 = ${pay} 元`);
        
      } else if (dateType === 'restday') {
        // ⭐ 休息日（週六）：前8h × 1.0
        const maxHours = 8;
        const actualHours = Math.min(dailyHours, maxHours);
        const pay = Math.round(hourlyRate * actualHours);
        
        restdayOvertimePay += pay;
        totalOvertimeHours += actualHours;
        
        Logger.log(`   休息日: ${actualHours.toFixed(2)}h × 1.0 = ${pay} 元`);
        
      } else if (dateType === 'holiday') {
        // ⭐ 國定假日：前8h × 1.0
        const maxHours = 8;
        const actualHours = Math.min(dailyHours, maxHours);
        const pay = Math.round(hourlyRate * actualHours);
        
        holidayOvertimePay += pay;
        totalOvertimeHours += actualHours;
        
        Logger.log(`   國定假日: ${actualHours.toFixed(2)}h × 1.0 = ${pay} 元`);
      }
    });
    
    Logger.log('');
    Logger.log('✅ 加班費計算完成:');
    Logger.log(`   延長工時加班費: ${extendedOvertimePay} 元`);
    Logger.log(`   休息日加班費: ${restdayOvertimePay} 元`);
    Logger.log(`   國定假日加班費: ${holidayOvertimePay} 元`);
    Logger.log(`   總時數: ${totalOvertimeHours.toFixed(2)}h`);
    
    return {
      extendedOvertimePay: extendedOvertimePay,
      restdayOvertimePay: restdayOvertimePay,
      holidayOvertimePay: holidayOvertimePay,
      totalOvertimeHours: totalOvertimeHours
    };
    
  } catch (error) {
    Logger.log('❌ calculateFactoryWorkerOvertime 錯誤: ' + error.message);
    return {
      extendedOvertimePay: 0,
      restdayOvertimePay: 0,
      holidayOvertimePay: 0,
      totalOvertimeHours: 0
    };
  }
}

/**
 * 🏭 計算食品廠移工請假扣款（情況二專用規則）
 * 
 * 規則：
 * - 請假扣：底薪/30/8 × 小時數
 * - 請病假：底薪/30/8 × 小時數/2
 * 
 * @param {string} employeeId - 員工ID
 * @param {string} yearMonth - 年月
 * @param {number} baseAllowanceLarge - 底薪（大）
 * @returns {Object} 請假扣款明細
 */
function calculateFactoryWorkerLeaveDeduction(employeeId, yearMonth, baseAllowanceLarge) {
  try {
    Logger.log('🏥 開始計算食品廠移工請假扣款...');
    
    // 計算時薪：底薪/30/8
    const hourlyRate = baseAllowanceLarge / 30 / 8;
    
    Logger.log(`   ⏰ 時薪: ${hourlyRate.toFixed(2)} 元/小時`);
    Logger.log(`      底薪（大）: ${baseAllowanceLarge}`);
    Logger.log('');
    
    // 讀取請假記錄
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
    
    // 遍歷所有記錄
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      const recordEmployeeId = String(row[1]).trim();
      const leaveType = row[4];
      const startDateTime = row[5];
      const workHours = parseFloat(row[7]) || 0;
      const status = String(row[10]).trim();
      
      if (recordEmployeeId !== employeeId) continue;
      if (status !== 'APPROVED') continue;
      
      let recordYearMonth = '';
      
      if (startDateTime instanceof Date) {
        recordYearMonth = Utilities.formatDate(startDateTime, 'Asia/Taipei', 'yyyy-MM');
      } else if (typeof startDateTime === 'string') {
        recordYearMonth = String(startDateTime).substring(0, 7);
      } else {
        continue;
      }
      
      if (recordYearMonth !== yearMonth) continue;
      
      // 累計時數
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
    
    // 計算扣款
    const sickLeaveDeduction = Math.round(hourlyRate * sickLeaveHours * 0.5);
    const personalLeaveDeduction = Math.round(hourlyRate * personalLeaveHours);
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
 * 🏭 計算食品廠移工特休未休補薪（情況二專用規則）
 * 
 * 規則：底薪/30 × 未休天數
 * 
 * @param {string} employeeId - 員工ID
 * @param {string} yearMonth - 年月
 * @param {number} baseAllowanceLarge - 底薪（大）
 * @returns {Object} 補薪明細
 */
function calculateFactoryWorkerUnusedLeavePay(employeeId, yearMonth, baseAllowanceLarge) {
  try {
    Logger.log('🏖️ 開始計算食品廠移工特休未休補薪');
    
    // 計算日薪：底薪/30
    const dailyRate = Math.round(baseAllowanceLarge / 30);
    
    Logger.log(`   日薪: ${dailyRate} 元`);
    Logger.log(`   底薪（大）: ${baseAllowanceLarge}`);
    Logger.log('');
    
    // 讀取未休假記錄
    let unusedDays = 0;
    
    try {
      const unusedLeaveSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('未休假記錄');
      
      if (unusedLeaveSheet) {
        const data = unusedLeaveSheet.getDataRange().getValues();
        
        for (let i = 1; i < data.length; i++) {
          const recordEmployeeId = String(data[i][0]).trim();
          
          let recordYearMonth = '';
          
          if (data[i][1] instanceof Date) {
            recordYearMonth = Utilities.formatDate(data[i][1], 'Asia/Taipei', 'yyyy-MM');
          } else if (typeof data[i][1] === 'string') {
            recordYearMonth = String(data[i][1]).trim().substring(0, 7);
          } else {
            recordYearMonth = String(data[i][1]).trim();
          }
          
          const recordDays = parseFloat(data[i][2]) || 0;
          
          if (recordEmployeeId === employeeId && recordYearMonth === yearMonth) {
            unusedDays = recordDays;
            Logger.log(`✅ 找到未休假記錄: ${unusedDays} 天`);
            break;
          }
        }
      }
    } catch (error) {
      Logger.log('⚠️ 讀取未休假記錄失敗: ' + error.message);
    }
    
    // 計算補薪
    const payAmount = unusedDays * dailyRate;
    
    Logger.log('');
    Logger.log('💰 補薪計算:');
    Logger.log(`   未休天數: ${unusedDays} 天`);
    Logger.log(`   日薪: ${dailyRate} 元`);
    Logger.log(`   補薪金額: ${unusedDays} × ${dailyRate} = ${payAmount} 元`);
    Logger.log('');
    
    return {
      amount: payAmount,
      days: unusedDays,
      dailyRate: dailyRate
    };
    
  } catch (error) {
    Logger.log('❌ calculateFactoryWorkerUnusedLeavePay 錯誤: ' + error.message);
    return {
      amount: 0,
      days: 0,
      dailyRate: 0
    };
  }
}

/**
 * 🧪 測試食品廠移工薪資計算（情況二）
 */
function testFactoryWorkerSalary() {
  Logger.log('🧪 測試食品廠移工薪資計算（情況二）');
  Logger.log('');
  
  const employeeId = 'U測試移工ID';  // ⚠️ 替換成實際員工ID
  const yearMonth = '2026-01';
  
  const result = calculateFactoryWorkerSalary(employeeId, yearMonth);
  
  Logger.log('');
  Logger.log('📤 測試結果:');
  Logger.log(JSON.stringify(result.data, null, 2));
}

console.log('✅ 情況二：食品廠移工薪資計算系統已載入');