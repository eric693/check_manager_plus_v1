// ==================================================================================
// 情況三：管理部行政（月薪制-標準工時）薪資計算系統
// ==================================================================================
//
// 💼 適用對象：管理部行政人員
// 💰 薪資類型：月薪制
// ⏰ 工時類型：標準工時
//
// ==================================================================================

/**
 * ✅ 情況三薪資計算規則
 * 
 * 1. 基本薪資：固定月薪
 * 
 * 2. 固定津貼：
 *    - 職務加給、伙食費、交通補助等
 * 
 * 3. 加班費（標準工時制 + 分鐘四捨五入）：
 *    - 延長工時加班費（前2小時）：底薪/30/8 × 1.34 × 時數（滿20分鐘算半小時，滿50分鐘算1小時）
 *    - 延長工時加班費（後2小時）：底薪/30/8 × 1.67 × 時數（滿20分鐘算半小時，滿50分鐘算1小時）
 *    - 休息日（週六）加班：
 *      * 前2小時：底薪/30/8 × 1.34 × 時數
 *      * 後6小時：底薪/30/8 × 1.67 × 時數
 *      * 超過8小時：底薪/30/8 × 2.67 × 時數
 *    - 國定假日出勤加班費：底薪/30 × 天數（出勤1小時也算1天）
 * 
 * 4. 請假扣款：
 *    - 請病假：底薪/30/8 × 小時數/2
 *    - 請事假：底薪/30/8 × 小時數
 * 
 * 5. 特休未休補薪：底薪/30 × 未休天數
 * 
 * 6. 不固定獎金：中秋、過年跟特殊計劃推動的短期津貼
 * 
 * ==================================================================================
 */
/**
 * 💼 計算管理部行政月薪資（情況三專用）
 * 
 * @param {string} employeeId - 員工ID
 * @param {string} yearMonth - 年月 (YYYY-MM)
 * @returns {Object} 薪資計算結果
 */
function calculateAdminSalary(employeeId, yearMonth) {
  try {
    Logger.log('═══════════════════════════════════════');
    Logger.log('💼 開始計算管理部行政薪資（情況三）');
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
    
    // ========================================
    // 步驟 3：計算固定津貼總額
    // ========================================
    const totalFixedAllowance = 
      positionAllowance + mealAllowance + transportAllowance +
      attendanceBonus + performanceBonus +
      otherAllowance1 + otherAllowance2 + otherAllowance3;
    
    Logger.log(`💰 固定津貼總額: ${totalFixedAllowance}`);
    Logger.log('');
    
    // ========================================
    // 步驟 4：計算加班費（情況三專用規則）
    // ========================================
    const overtimeResult = calculateAdminOvertimePay(
      employeeId, 
      yearMonth, 
      baseSalary
    );
    
    const extendedOvertimeFirst2h = overtimeResult.extendedOvertimeFirst2h;    // 延長工時前2h
    const extendedOvertimeAfter2h = overtimeResult.extendedOvertimeAfter2h;    // 延長工時後2h
    const restdayOvertimePay = overtimeResult.restdayOvertimePay;              // 休息日加班費
    const holidayOvertimePay = overtimeResult.holidayOvertimePay;              // 國定假日加班費
    const totalOvertimeHours = overtimeResult.totalOvertimeHours;
    
    Logger.log('💼 加班費計算結果:');
    Logger.log(`   延長工時（前2h）: ${extendedOvertimeFirst2h}`);
    Logger.log(`   延長工時（後2h）: ${extendedOvertimeAfter2h}`);
    Logger.log(`   休息日加班費: ${restdayOvertimePay}`);
    Logger.log(`   國定假日加班費: ${holidayOvertimePay}`);
    Logger.log(`   總加班時數: ${totalOvertimeHours}`);
    Logger.log('');
    
    // ========================================
    // 步驟 5：計算請假扣款（情況三專用規則）
    // ========================================
    const leaveDeductionResult = calculateAdminLeaveDeduction(
      employeeId, 
      yearMonth, 
      baseSalary
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
    // 步驟 6：計算特休未休補薪（情況三專用規則）
    // ========================================
    const unusedLeavePayResult = calculateAdminUnusedLeavePay(
      employeeId, 
      yearMonth, 
      baseSalary
    );
    
    const unusedLeavePay = unusedLeavePayResult.amount;
    const unusedLeaveDays = unusedLeavePayResult.days;
    
    Logger.log('🏖️ 特休未休補薪:');
    Logger.log(`   未休天數: ${unusedLeaveDays} 天`);
    Logger.log(`   補薪金額: ${unusedLeavePay} 元`);
    Logger.log('');
    
    // ========================================
    // 步驟 7：計算不固定獎金
    // ========================================
    const bonusResult = getMonthlyBonus(employeeId, yearMonth);
    const monthlyBonus = bonusResult.amount;
    const bonusNote = bonusResult.note;
    
    Logger.log('🎁 不固定獎金:');
    Logger.log(`   獎金金額: ${monthlyBonus} 元`);
    if (bonusNote) {
      Logger.log(`   獎金說明: ${bonusNote}`);
    }
    Logger.log('');
    
    // ========================================
    // 步驟 8：計算應發總額
    // ========================================
    const grossSalary = 
      baseSalary +                      // 基本薪資
      totalFixedAllowance +             // 固定津貼
      extendedOvertimeFirst2h +         // 延長工時（前2h）
      extendedOvertimeAfter2h +         // 延長工時（後2h）
      restdayOvertimePay +              // 休息日加班費
      holidayOvertimePay +              // 國定假日加班費
      unusedLeavePay +                  // 特休未休補薪
      monthlyBonus;                     // 不固定獎金
    
    Logger.log(`💵 應發總額: ${grossSalary}`);
    Logger.log('');
    
    // ========================================
    // 步驟 9：計算扣款總額
    // ========================================
    const totalDeductions = 
      laborFee +                        // 勞保費
      healthFee +                       // 健保費
      employmentFee +                   // 就業保險費
      pensionSelf +                     // 勞退自提
      incomeTax +                       // 所得稅
      totalLeaveDeduction +             // 請假扣款
      welfareFee +                      // 福利金扣款
      dormitoryFee +                    // 宿舍費用
      groupInsurance +                  // 團保費用
      otherDeduction1 +                 // 其他扣款1
      otherDeduction2;                  // 其他扣款2
    
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
      
      // 基本薪資與津貼
      baseSalary: baseSalary,
      positionAllowance: positionAllowance,
      mealAllowance: mealAllowance,
      transportAllowance: transportAllowance,
      attendanceBonus: attendanceBonus,
      performanceBonus: performanceBonus,
      otherAllowance1: otherAllowance1,
      otherAllowance2: otherAllowance2,
      otherAllowance3: otherAllowance3,
      
      // 加班費（情況三專用欄位）
      extendedOvertimeFirst2h: extendedOvertimeFirst2h,
      extendedOvertimeAfter2h: extendedOvertimeAfter2h,
      restdayOvertimePay: restdayOvertimePay,
      holidayOvertimePay: holidayOvertimePay,
      totalOvertimeHours: totalOvertimeHours,
      
      // 補薪項目
      unusedLeavePay: unusedLeavePay,
      unusedLeaveDays: unusedLeaveDays,
      monthlyBonus: monthlyBonus,          // 不固定獎金
      bonusNote: bonusNote,                // 獎金說明
      
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
    Logger.log('✅✅✅ 管理部行政薪資計算完成（情況三）');
    Logger.log('═══════════════════════════════════════');
    Logger.log('📊 薪資結構摘要:');
    Logger.log(`   基本薪資: ${baseSalary}`);
    Logger.log(`   固定津貼: ${totalFixedAllowance}`);
    Logger.log(`   延長工時（前2h）: ${extendedOvertimeFirst2h}`);
    Logger.log(`   延長工時（後2h）: ${extendedOvertimeAfter2h}`);
    Logger.log(`   休息日加班費: ${restdayOvertimePay}`);
    Logger.log(`   國定假日加班費: ${holidayOvertimePay}`);
    Logger.log(`   特休未休補薪: ${unusedLeavePay}`);
    Logger.log(`   不固定獎金: ${monthlyBonus}`);
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
    Logger.log('❌❌❌ calculateAdminSalary 發生錯誤');
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
 * 💼 計算管理部行政加班費（情況三專用規則 - 含分鐘四捨五入）
 * 
 * 規則：
 * 1. 延長工時（前2小時）：底薪/30/8 × 1.34 × 時數（滿20分鐘算半小時，滿50分鐘算1小時）
 * 2. 延長工時（後2小時）：底薪/30/8 × 1.67 × 時數（滿20分鐘算半小時，滿50分鐘算1小時）
 * 3. 休息日（週六）加班：
 *    - 前2小時：底薪/30/8 × 1.34 × 時數
 *    - 後6小時：底薪/30/8 × 1.67 × 時數
 *    - 超過8小時：底薪/30/8 × 2.67 × 時數
 * 4. 國定假日出勤：底薪/30 × 天數（出勤1小時也算1天）
 * 
 * @param {string} employeeId - 員工ID
 * @param {string} yearMonth - 年月
 * @param {number} baseSalary - 基本薪資（底薪）
 * @returns {Object} 加班費計算結果
 */
function calculateAdminOvertimePay(employeeId, yearMonth, baseSalary) {
  try {
    Logger.log('💼 開始計算管理部行政加班費（含分鐘四捨五入）');
    
    // 計算時薪：底薪/30/8
    const hourlyRate = baseSalary / 30 / 8;
    
    Logger.log(`   時薪: ${hourlyRate.toFixed(2)} 元/小時`);
    Logger.log(`   底薪: ${baseSalary}`);
    Logger.log('');
    
    // 取得加班記錄
    const overtimeRecords = getEmployeeMonthlyOvertime(employeeId, yearMonth);
    
    if (overtimeRecords.length === 0) {
      Logger.log('   ⚠️ 無加班記錄');
      return {
        extendedOvertimeFirst2h: 0,
        extendedOvertimeAfter2h: 0,
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
    
    let extendedOvertimeFirst2h = 0;     // 延長工時（前2h）
    let extendedOvertimeAfter2h = 0;     // 延長工時（後2h）
    let restdayOvertimePay = 0;          // 休息日加班費
    let holidayOvertimePay = 0;          // 國定假日加班費（以天計算）
    let totalOvertimeHours = 0;
    
    const holidayDays = new Set();       // 記錄國定假日出勤天數
    
    // 遍歷每天的加班記錄
    Object.keys(overtimeByDate).forEach(date => {
      let dailyHours = overtimeByDate[date];
      const dateType = getDateType(date);
      
      Logger.log(`\n📅 ${date} (${dateType}): ${dailyHours.toFixed(2)}h`);
      
      if (dateType === 'weekday') {
        // ⭐ 平日延長工時：前2h × 1.34，後2h × 1.67
        
        // ⭐⭐⭐ 關鍵：分鐘四捨五入（滿20分鐘算半小時，滿50分鐘算1小時）
        const roundedHours = roundOvertimeMinutes(dailyHours);
        
        Logger.log(`   原始時數: ${dailyHours.toFixed(2)}h → 四捨五入: ${roundedHours}h`);
        
        if (roundedHours <= 2) {
          // 前2小時
          const pay = Math.round(hourlyRate * roundedHours * 1.34);
          extendedOvertimeFirst2h += pay;
          Logger.log(`   前2h: ${roundedHours}h × 1.34 = ${pay} 元`);
        } else {
          // 前2小時
          const first2h = 2;
          const payFirst = Math.round(hourlyRate * first2h * 1.34);
          extendedOvertimeFirst2h += payFirst;
          Logger.log(`   前2h: 2h × 1.34 = ${payFirst} 元`);
          
          // 後2小時（最多再算2小時）
          const after2h = Math.min(roundedHours - 2, 2);
          const payAfter = Math.round(hourlyRate * after2h * 1.67);
          extendedOvertimeAfter2h += payAfter;
          Logger.log(`   後2h: ${after2h}h × 1.67 = ${payAfter} 元`);
        }
        
        totalOvertimeHours += roundedHours;
        
      } else if (dateType === 'restday') {
        // ⭐ 休息日（週六）：前2h (×1.34)，3-8h (×1.67)，9h起 (×2.67)
        
        const roundedHours = roundOvertimeMinutes(dailyHours);
        Logger.log(`   原始時數: ${dailyHours.toFixed(2)}h → 四捨五入: ${roundedHours}h`);
        
        let pay = 0;
        
        if (roundedHours <= 2) {
          // 前2小時
          pay = Math.round(hourlyRate * roundedHours * 1.34);
          Logger.log(`   前2h: ${roundedHours}h × 1.34 = ${pay} 元`);
        } else if (roundedHours <= 8) {
          // 前2h + 3-8h
          const first2h = 2;
          const after2h = roundedHours - 2;
          
          const payFirst = Math.round(hourlyRate * first2h * 1.34);
          const payAfter = Math.round(hourlyRate * after2h * 1.67);
          
          pay = payFirst + payAfter;
          
          Logger.log(`   前2h: 2h × 1.34 = ${payFirst} 元`);
          Logger.log(`   3-8h: ${after2h}h × 1.67 = ${payAfter} 元`);
        } else {
          // 前2h + 3-8h + 9h起
          const first2h = 2;
          const middle6h = 6;
          const after8h = roundedHours - 8;
          
          const payFirst = Math.round(hourlyRate * first2h * 1.34);
          const payMiddle = Math.round(hourlyRate * middle6h * 1.67);
          const payAfter = Math.round(hourlyRate * after8h * 2.67);
          
          pay = payFirst + payMiddle + payAfter;
          
          Logger.log(`   前2h: 2h × 1.34 = ${payFirst} 元`);
          Logger.log(`   3-8h: 6h × 1.67 = ${payMiddle} 元`);
          Logger.log(`   9h起: ${after8h}h × 2.67 = ${payAfter} 元`);
        }
        
        restdayOvertimePay += pay;
        totalOvertimeHours += roundedHours;
        
      } else if (dateType === 'holiday') {
        // ⭐ 國定假日：出勤1小時也算1天（底薪/30）
        
        if (dailyHours > 0) {
          holidayDays.add(date);
          Logger.log(`   國定假日出勤（計1天）`);
        }
        
        totalOvertimeHours += dailyHours;
      }
    });
    
    // 計算國定假日加班費（以天計算）
    const dailyRate = Math.round(baseSalary / 30);
    holidayOvertimePay = holidayDays.size * dailyRate;
    
    if (holidayDays.size > 0) {
      Logger.log(`\n🎌 國定假日出勤天數: ${holidayDays.size} 天`);
      Logger.log(`   日薪: ${dailyRate} 元`);
      Logger.log(`   加班費: ${holidayDays.size} × ${dailyRate} = ${holidayOvertimePay} 元`);
    }
    
    Logger.log('');
    Logger.log('✅ 加班費計算完成:');
    Logger.log(`   延長工時（前2h）: ${extendedOvertimeFirst2h} 元`);
    Logger.log(`   延長工時（後2h）: ${extendedOvertimeAfter2h} 元`);
    Logger.log(`   休息日加班費: ${restdayOvertimePay} 元`);
    Logger.log(`   國定假日加班費: ${holidayOvertimePay} 元`);
    Logger.log(`   總時數: ${totalOvertimeHours.toFixed(2)}h`);
    
    return {
      extendedOvertimeFirst2h: extendedOvertimeFirst2h,
      extendedOvertimeAfter2h: extendedOvertimeAfter2h,
      restdayOvertimePay: restdayOvertimePay,
      holidayOvertimePay: holidayOvertimePay,
      totalOvertimeHours: totalOvertimeHours
    };
    
  } catch (error) {
    Logger.log('❌ calculateAdminOvertimePay 錯誤: ' + error.message);
    return {
      extendedOvertimeFirst2h: 0,
      extendedOvertimeAfter2h: 0,
      restdayOvertimePay: 0,
      holidayOvertimePay: 0,
      totalOvertimeHours: 0
    };
  }
}

/**
 * ⭐⭐⭐ 加班時數分鐘四捨五入（情況三專用）
 * 
 * 規則：
 * - 滿20分鐘算半小時（0.5h）
 * - 滿50分鐘算1小時（1.0h）
 * 
 * 例如：
 * - 1.2h (1小時12分) → 1.0h
 * - 1.4h (1小時24分) → 1.5h
 * - 1.7h (1小時42分) → 1.5h
 * - 1.9h (1小時54分) → 2.0h
 * 
 * @param {number} hours - 原始小時數
 * @returns {number} 四捨五入後的小時數
 */
function roundOvertimeMinutes(hours) {
  const wholePart = Math.floor(hours);           // 整數小時
  const decimalPart = hours - wholePart;         // 小數部分
  const minutes = Math.round(decimalPart * 60);  // 分鐘數
  
  if (minutes < 20) {
    // 不滿20分鐘，不計算
    return wholePart;
  } else if (minutes < 50) {
    // 滿20分鐘但不滿50分鐘，算半小時
    return wholePart + 0.5;
  } else {
    // 滿50分鐘，算1小時
    return wholePart + 1.0;
  }
}

/**
 * 💼 計算管理部行政請假扣款（情況三專用規則）
 * 
 * 規則：
 * - 請病假：底薪/30/8 × 小時數/2
 * - 請事假：底薪/30/8 × 小時數
 * 
 * @param {string} employeeId - 員工ID
 * @param {string} yearMonth - 年月
 * @param {number} baseSalary - 基本薪資（底薪）
 * @returns {Object} 請假扣款明細
 */
function calculateAdminLeaveDeduction(employeeId, yearMonth, baseSalary) {
  try {
    Logger.log('🏥 開始計算管理部行政請假扣款...');
    
    // 計算時薪：底薪/30/8
    const hourlyRate = baseSalary / 30 / 8;
    
    Logger.log(`   ⏰ 時薪: ${hourlyRate.toFixed(2)} 元/小時`);
    Logger.log(`      底薪: ${baseSalary}`);
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
    const sickLeaveDeduction = Math.round(hourlyRate * sickLeaveHours * 0.5);      // 病假扣半薪
    const personalLeaveDeduction = Math.round(hourlyRate * personalLeaveHours);    // 事假扣全薪
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
 * 💼 計算管理部行政特休未休補薪（情況三專用規則）
 * 
 * 規則：底薪/30 × 未休天數
 * 
 * @param {string} employeeId - 員工ID
 * @param {string} yearMonth - 年月
 * @param {number} baseSalary - 基本薪資（底薪）
 * @returns {Object} 補薪明細
 */
function calculateAdminUnusedLeavePay(employeeId, yearMonth, baseSalary) {
  try {
    Logger.log('🏖️ 開始計算管理部行政特休未休補薪');
    
    // 計算日薪：底薪/30
    const dailyRate = Math.round(baseSalary / 30);
    
    Logger.log(`   日薪: ${dailyRate} 元`);
    Logger.log(`   底薪: ${baseSalary}`);
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
    Logger.log('❌ calculateAdminUnusedLeavePay 錯誤: ' + error.message);
    return {
      amount: 0,
      days: 0,
      dailyRate: 0
    };
  }
}

/**
 * 🎁 取得當月不固定獎金
 * 
 * 規則：
 * - 中秋獎金（9月）
 * - 過年獎金（2月）
 * - 特殊計劃推動的短期津貼（不固定）
 * 
 * @param {string} employeeId - 員工ID
 * @param {string} yearMonth - 年月 (YYYY-MM)
 * @returns {Object} { amount, note }
 */
function getMonthlyBonus(employeeId, yearMonth) {
  try {
    Logger.log('🎁 查詢當月不固定獎金');
    
    const bonusSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('不固定獎金記錄');
    
    if (!bonusSheet) {
      Logger.log('   ⚠️ 找不到「不固定獎金記錄」工作表');
      Logger.log('   💡 提示：首次使用時，請先執行 initBonusSheet() 建立工作表');
      return { amount: 0, note: '' };
    }
    
    const data = bonusSheet.getDataRange().getValues();
    
    // 查找該員工該月份的獎金記錄
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
      
      if (recordEmployeeId === employeeId && recordYearMonth === yearMonth) {
        const amount = parseFloat(data[i][2]) || 0;
        const note = String(data[i][3] || '');
        
        Logger.log(`   ✅ 找到獎金記錄: ${amount} 元 (${note})`);
        
        return { amount: amount, note: note };
      }
    }
    
    Logger.log('   ⚠️ 該月份無獎金記錄');
    return { amount: 0, note: '' };
    
  } catch (error) {
    Logger.log('❌ getMonthlyBonus 錯誤: ' + error.message);
    return { amount: 0, note: '' };
  }
}

/**
 * 🔧 初始化「不固定獎金記錄」工作表
 */
function initBonusSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('不固定獎金記錄');
  
  if (!sheet) {
    Logger.log('📝 建立「不固定獎金記錄」工作表...');
    
    sheet = ss.insertSheet('不固定獎金記錄');
    
    const headers = ['員工ID', '年月', '獎金金額', '獎金說明', '更新時間'];
    sheet.appendRow(headers);
    
    // 美化標題列
    const headerRange = sheet.getRange(1, 1, 1, 5);
    headerRange.setBackground('#FFD700');
    headerRange.setFontColor('#000000');
    headerRange.setFontWeight('bold');
    headerRange.setHorizontalAlignment('center');
    
    sheet.setFrozenRows(1);
    
    Logger.log('✅ 「不固定獎金記錄」工作表已建立');
  }
  
  return sheet;
}

/**
 * 🎁 設定員工當月獎金
 * 
 * @param {string} employeeId - 員工ID
 * @param {string} yearMonth - 年月 (YYYY-MM)
 * @param {number} amount - 獎金金額
 * @param {string} note - 獎金說明
 */
function setMonthlyBonus(employeeId, yearMonth, amount, note = '') {
  try {
    const sheet = initBonusSheet();
    const data = sheet.getDataRange().getValues();
    
    // 查找是否已存在
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      const existingId = String(data[i][0]).trim();
      
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
    
    // ⭐ 年月前加單引號，強制為文本格式
    const yearMonthText = "'" + yearMonth;
    
    const row = [
      employeeId,
      yearMonthText,
      amount,
      note,
      new Date()
    ];
    
    if (rowIndex > 0) {
      // 更新現有記錄
      sheet.getRange(rowIndex, 1, 1, 5).setValues([row]);
      Logger.log(`✅ 更新獎金記錄: ${employeeId}, ${yearMonth}, ${amount}元`);
    } else {
      // 新增記錄
      sheet.appendRow(row);
      Logger.log(`✅ 新增獎金記錄: ${employeeId}, ${yearMonth}, ${amount}元`);
    }
    
    // 設定 B 欄為文本格式
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 2).setNumberFormat('@');
    
    return { success: true, message: '獎金記錄已更新' };
    
  } catch (error) {
    Logger.log(`❌ 設定獎金記錄失敗: ${error.message}`);
    return { success: false, message: error.message };
  }
}

/**
 * 🧪 測試管理部行政薪資計算（情況三）
 */
function testAdminSalary() {
  Logger.log('🧪 測試管理部行政薪資計算（情況三）');
  Logger.log('');
  
  const employeeId = 'Uf664a35632b736301d674d8b2cc3f8c0';  // ⚠️ 替換成實際員工ID
  const yearMonth = '2026-01';
  
  const result = calculateAdminSalary(employeeId, yearMonth);
  
  Logger.log('');
  Logger.log('📤 測試結果:');
  Logger.log(JSON.stringify(result.data, null, 2));
}

/**
 * 🧪 測試加班時數四捨五入
 */
function testRoundOvertimeMinutes() {
  Logger.log('🧪 測試加班時數四捨五入');
  Logger.log('');
  
  const testCases = [
    { input: 1.2, expected: 1.0, desc: '1小時12分 → 1.0h' },
    { input: 1.4, expected: 1.5, desc: '1小時24分 → 1.5h' },
    { input: 1.7, expected: 1.5, desc: '1小時42分 → 1.5h' },
    { input: 1.9, expected: 2.0, desc: '1小時54分 → 2.0h' },
    { input: 2.1, expected: 2.0, desc: '2小時6分 → 2.0h' },
    { input: 2.35, expected: 2.5, desc: '2小時21分 → 2.5h' },
    { input: 2.85, expected: 3.0, desc: '2小時51分 → 3.0h' }
  ];
  
  testCases.forEach(testCase => {
    const result = roundOvertimeMinutes(testCase.input);
    const status = (result === testCase.expected) ? '✅' : '❌';
    Logger.log(`${status} ${testCase.desc}: ${testCase.input}h → ${result}h`);
  });
  
  Logger.log('');
  Logger.log('測試完成！');
}

console.log('✅ 情況三：管理部行政薪資計算系統已載入');


// ==================================================================================
// 月薪資記錄 Sheet 結構更新指引（支援情況三）
// ==================================================================================

/**
 * 🔧 更新月薪資記錄工作表結構（支援情況三）
 * 
 * 執行步驟：
 * 1. 備份現有資料
 * 2. 刪除舊工作表
 * 3. 建立新工作表（包含情況三的欄位）
 * 4. 恢復資料
 */
function rebuildMonthlySalarySheetForCase3() {
  try {
    Logger.log('🔧 開始更新月薪資記錄工作表（支援情況三）...');
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
        backupData = oldSheet.getRange(1, 1, lastRow, oldSheet.getLastColumn()).getValues();
        Logger.log(`   ✅ 已備份 ${lastRow - 1} 筆薪資記錄`);
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
    // 步驟 3: 建立新工作表（支援情況三）
    // ========================================
    Logger.log('📝 步驟 3: 建立新工作表...');
    
    const newSheet = ss.insertSheet('月薪資記錄');
    
    // ⭐⭐⭐ 完整標題列（支援情況一、二、三）
    const headers = [
      // A-I: 基本資訊 (9欄)
      '薪資單ID',        // A
      '員工ID',          // B
      '員工姓名',        // C
      '年月',            // D
      '薪資類型',        // E
      '工時類型',        // F
      '時薪',            // G
      '工作時數',        // H
      '總加班時數',      // I

      // J-Y: 收入項目 (16欄)
      '基本薪資',        // J
      '職務加給',        // K
      '伙食費',          // L
      '交通補助',        // M
      '全勤獎金',        // N
      '績效獎金',        // O
      '其他津貼1',       // P
      '其他津貼2',       // Q
      '其他津貼3',       // R
      
      // ⭐⭐⭐ 加班費（多用途欄位）
      '平日加班費/延長工時(前2h)',    // S（情況一二 / 情況三）
      '休息日加班費/延長工時(後2h)',  // T（情況一二 / 情況三）
      '國定假日加班費',               // U（通用）
      
      '未休假補薪',      // V
      '未休假天數',      // W
      '月休補薪',        // X（情況一專用）
      '未休月休天數',    // Y（情況一專用）

      // Z-AE: 法定扣款 (6欄)
      '勞保費',          // Z
      '健保費',          // AA
      '就業保險費',      // AB
      '勞退自提率(%)',   // AC
      '勞退自提',        // AD
      '所得稅',          // AE

      // AF-AO: 其他扣款 (10欄)
      '請假扣款',        // AF（總計）
      '病假時數',        // AG
      '病假扣款',        // AH
      '事假時數',        // AI
      '事假扣款',        // AJ
      '福利金扣款',      // AK
      '宿舍費用',        // AL
      '團保費用',        // AM
      '其他扣款1',       // AN
      '其他扣款2',       // AO

      // AP-AQ: 總額 (2欄)
      '應發總額',        // AP
      '實發金額',        // AQ

      // AR-AS: 銀行 (2欄)
      '銀行代碼',        // AR
      '銀行帳號',        // AS

      // AT-AV: 系統 (3欄)
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
    
    // 設定關鍵欄寬
    newSheet.setColumnWidth(1, 250);  // 薪資單ID
    newSheet.setColumnWidth(2, 200);  // 員工ID
    newSheet.setColumnWidth(3, 120);  // 員工姓名
    newSheet.setColumnWidth(4, 80);   // 年月
    newSheet.setColumnWidth(19, 180); // S欄（加班費）
    newSheet.setColumnWidth(20, 180); // T欄（加班費）
    
    Logger.log('   ✅ 標題列格式化完成');
    Logger.log('');
    
    // ========================================
    // 步驟 4: 恢復資料
    // ========================================
    if (backupData.length > 1) {
      Logger.log('📥 步驟 4: 恢復資料...');
      
      const dataRows = backupData.slice(1);
      
      const restoredRows = dataRows.map(oldRow => {
        const newRow = new Array(headers.length).fill('');
        
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
    }
    
    // ========================================
    // 完成
    // ========================================
    Logger.log('═══════════════════════════════════════');
    Logger.log('✅✅✅ 月薪資記錄工作表更新完成！');
    Logger.log('═══════════════════════════════════════');
    Logger.log('');
    Logger.log('📋 新工作表資訊:');
    Logger.log(`   名稱: 月薪資記錄`);
    Logger.log(`   欄位數: ${headers.length}`);
    Logger.log(`   資料筆數: ${newSheet.getLastRow() - 1}`);
    Logger.log('');
    Logger.log('🔍 情況三專用欄位:');
    Logger.log('   S欄 (第19欄): 平日加班費/延長工時(前2h)');
    Logger.log('   T欄 (第20欄): 休息日加班費/延長工時(後2h)');
    Logger.log('   U欄 (第21欄): 國定假日加班費');
    Logger.log('');
    Logger.log('⚠️ 下一步：');
    Logger.log('   1. 初始化不固定獎金記錄表: initBonusSheet()');
    Logger.log('   2. 測試情況三計算: testAdminSalary()');
    Logger.log('   3. 重新計算所有員工薪資');
    Logger.log('');
    
    return {
      success: true,
      message: '月薪資記錄工作表更新成功',
      totalColumns: headers.length,
      restoredRows: backupData.length > 1 ? backupData.length - 1 : 0
    };
    
  } catch (error) {
    Logger.log('❌ 更新工作表失敗: ' + error.message);
    Logger.log('❌ 錯誤堆疊: ' + error.stack);
    
    return {
      success: false,
      message: '更新失敗: ' + error.message
    };
  }
}

/**
 * 📊 欄位對應說明
 * 
 * 由於三種情況的加班費計算方式不同，S、T 欄採用多用途設計：
 * 
 * ┌─────────────────────────────────────────────────────────────┐
 * │ S欄（第19欄）：平日加班費 / 延長工時加班費（前2h）           │
 * ├─────────────────────────────────────────────────────────────┤
 * │ 情況一：飼料廠司機 → 平日加班費（固定$200/h × 1.34）        │
 * │ 情況二：食品廠移工 → 延長工時加班費（底薪/30/8 × 1.34）     │
 * │ 情況三：管理部行政 → 延長工時加班費（前2h，底薪/30/8 × 1.34）│
 * └─────────────────────────────────────────────────────────────┘
 * 
 * ┌─────────────────────────────────────────────────────────────┐
 * │ T欄（第20欄）：休息日加班費 / 延長工時加班費（後2h）         │
 * ├─────────────────────────────────────────────────────────────┤
 * │ 情況一：飼料廠司機 → 休息日加班費（固定$200/h × 1.34/1.67） │
 * │ 情況二：食品廠移工 → 休息日加班費（底薪/30/8 × 時數）       │
 * │ 情況三：管理部行政 → 延長工時加班費（後2h，底薪/30/8 × 1.67）│
 * │                    + 休息日加班費（合併顯示）                │
 * └─────────────────────────────────────────────────────────────┘
 * 
 * ⚠️ 注意：情況三的休息日加班費會與延長工時（後2h）合併顯示在 T 欄
 * 
 * 如果需要分開顯示，可以考慮增加新欄位：
 * - S欄：延長工時（前2h）
 * - T欄：延長工時（後2h）
 * - U欄：休息日加班費
 * - V欄：國定假日加班費
 */

console.log('✅ 月薪資記錄 Sheet 更新工具已載入');


// ==================================================================================
// 薪資計算系統完整整合測試指引
// ==================================================================================

/**
 * 📋 系統整合測試總覽
 * 
 * 本指引涵蓋三種情況的完整測試流程：
 * 
 * 情況一：飼料廠司機（月薪制-不定時）
 * 情況二：食品廠移工（時薪制-標準工時）
 * 情況三：管理部行政（月薪制-標準工時）
 */

// ==================================================================================
// 步驟 0：初始化系統（首次執行）
// ==================================================================================

/**
 * 🔧 初始化所有必要的工作表
 */
function initializeAllSheets() {
  Logger.log('🔧 開始初始化系統...');
  Logger.log('');
  
  // 1. 初始化員工薪資設定
  Logger.log('📝 步驟 1: 檢查員工薪資設定工作表');
  const configSheet = getEmployeeSalarySheet();
  Logger.log('   ✅ 員工薪資設定工作表已就緒');
  Logger.log('');
  
  // 2. 初始化月薪資記錄（支援情況三）
  Logger.log('📝 步驟 2: 更新月薪資記錄工作表');
  const sheetResult = rebuildMonthlySalarySheetForCase3();
  if (sheetResult.success) {
    Logger.log('   ✅ 月薪資記錄工作表已更新');
  } else {
    Logger.log('   ❌ 月薪資記錄工作表更新失敗: ' + sheetResult.message);
    return;
  }
  Logger.log('');
  
  // 3. 初始化未休假記錄
  Logger.log('📝 步驟 3: 初始化未休假記錄工作表');
  initUnusedLeaveSheet();
  Logger.log('   ✅ 未休假記錄工作表已就緒');
  Logger.log('');
  
  // 4. 初始化月休記錄（情況一專用）
  Logger.log('📝 步驟 4: 初始化月休記錄工作表');
  initMonthlyRestDaySheet();
  Logger.log('   ✅ 月休記錄工作表已就緒');
  Logger.log('');
  
  // 5. 初始化不固定獎金記錄（情況三專用）
  Logger.log('📝 步驟 5: 初始化不固定獎金記錄工作表');
  initBonusSheet();
  Logger.log('   ✅ 不固定獎金記錄工作表已就緒');
  Logger.log('');
  
  Logger.log('═══════════════════════════════════════');
  Logger.log('✅✅✅ 系統初始化完成！');
  Logger.log('═══════════════════════════════════════');
  Logger.log('');
  Logger.log('📋 已建立以下工作表：');
  Logger.log('   1. 員工薪資設定');
  Logger.log('   2. 月薪資記錄');
  Logger.log('   3. 未休假記錄');
  Logger.log('   4. 月休記錄');
  Logger.log('   5. 不固定獎金記錄');
  Logger.log('');
  Logger.log('⚠️ 下一步：設定測試資料');
  Logger.log('   執行: setupTestData()');
}

// ==================================================================================
// 步驟 1：設定測試資料
// ==================================================================================

/**
 * 🧪 設定三種情況的測試員工資料
 */
function setupTestData() {
  Logger.log('🧪 開始設定測試資料...');
  Logger.log('');
  
  const yearMonth = '2026-01';
  
  // ========================================
  // 情況一：飼料廠司機測試資料
  // ========================================
  Logger.log('📝 情況一：飼料廠司機');
  
  const driver = {
    employeeId: 'TEST_DRIVER_001',
    employeeName: '測試司機_張三',
    idNumber: 'A123456789',
    employeeType: '飼料廠司機',
    salaryType: '月薪',
    workTimeType: '不定時',
    baseSalary: 38000,
    positionAllowance: 2400,    // 職務加給
    mealAllowance: 2400,        // 伙食費
    transportAllowance: 2000,   // 交通津貼
    attendanceBonus: 1000,      // 全勤獎金
    bankCode: '004',
    bankAccount: '1234567890',
    laborFee: 955,
    healthFee: 592,
    employmentFee: 0,
    pensionSelf: 0
  };
  
  setEmployeeSalaryTW(driver);
  Logger.log('   ✅ 司機資料已設定');
  
  // 設定月休記錄（應休6天，實際休4天）
  setMonthlyRestDays('TEST_DRIVER_001', yearMonth, 4, '測試：實際休4天');
  Logger.log('   ✅ 月休記錄已設定（實際休4天，應補薪2天）');
  
  // 設定未休假記錄
  setUnusedLeaveDays('TEST_DRIVER_001', yearMonth, 3, '測試：未休3天特休');
  Logger.log('   ✅ 未休假記錄已設定（未休3天）');
  
  Logger.log('');
  
  // ========================================
  // 情況二：食品廠移工測試資料
  // ========================================
  Logger.log('📝 情況二：食品廠移工');
  
  const worker = {
    employeeId: 'TEST_WORKER_001',
    employeeName: '測試移工_阮文明',
    idNumber: 'AB12345678',
    employeeType: '食品廠移工',
    salaryType: '時薪',
    workTimeType: '標準工時',
    baseSalary: 190,            // 時薪 $190
    positionAllowance: 5000,    // 底薪（大）
    mealAllowance: 2000,        // 工作津貼
    bankCode: '700',
    bankAccount: '9876543210',
    laborFee: 500,
    healthFee: 300,
    employmentFee: 100
  };
  
  setEmployeeSalaryTW(worker);
  Logger.log('   ✅ 移工資料已設定');
  
  // 設定未休假記錄
  setUnusedLeaveDays('TEST_WORKER_001', yearMonth, 2, '測試：未休2天特休');
  Logger.log('   ✅ 未休假記錄已設定（未休2天）');
  
  Logger.log('');
  
  // ========================================
  // 情況三：管理部行政測試資料
  // ========================================
  Logger.log('📝 情況三：管理部行政');
  
  const admin = {
    employeeId: 'TEST_ADMIN_001',
    employeeName: '測試行政_李小美',
    idNumber: 'C234567890',
    employeeType: '管理部行政',
    salaryType: '月薪',
    workTimeType: '標準工時',
    baseSalary: 35000,
    positionAllowance: 3000,    // 職務加給
    mealAllowance: 2400,        // 伙食費
    transportAllowance: 1500,   // 交通補助
    attendanceBonus: 1000,      // 全勤獎金
    performanceBonus: 2000,     // 績效獎金
    bankCode: '822',
    bankAccount: '5555666677',
    laborFee: 800,
    healthFee: 500,
    employmentFee: 100,
    pensionSelfRate: 6,
    pensionSelf: 2100
  };
  
  setEmployeeSalaryTW(admin);
  Logger.log('   ✅ 行政資料已設定');
  
  // 設定未休假記錄
  setUnusedLeaveDays('TEST_ADMIN_001', yearMonth, 5, '測試：未休5天特休');
  Logger.log('   ✅ 未休假記錄已設定（未休5天）');
  
  // 設定不固定獎金（中秋獎金）
  setMonthlyBonus('TEST_ADMIN_001', yearMonth, 3000, '中秋獎金');
  Logger.log('   ✅ 不固定獎金已設定（中秋獎金 $3,000）');
  
  Logger.log('');
  
  Logger.log('═══════════════════════════════════════');
  Logger.log('✅✅✅ 測試資料設定完成！');
  Logger.log('═══════════════════════════════════════');
  Logger.log('');
  Logger.log('📋 已建立3位測試員工：');
  Logger.log('   1. TEST_DRIVER_001 - 測試司機_張三（飼料廠司機）');
  Logger.log('   2. TEST_WORKER_001 - 測試移工_阮文明（食品廠移工）');
  Logger.log('   3. TEST_ADMIN_001 - 測試行政_李小美（管理部行政）');
  Logger.log('');
  Logger.log('⚠️ 下一步：執行路由測試');
  Logger.log('   執行: testAllCases()');
}

// ==================================================================================
// 步驟 2：測試三種情況的薪資計算
// ==================================================================================

/**
 * 🧪 測試所有情況的薪資計算
 */
function testAllCases() {
  Logger.log('🧪 開始測試三種情況的薪資計算...');
  Logger.log('');
  
  const yearMonth = '2026-01';
  
  // ========================================
  // 測試情況一：飼料廠司機
  // ========================================
  Logger.log('═══════════════════════════════════════');
  Logger.log('📝 測試情況一：飼料廠司機');
  Logger.log('═══════════════════════════════════════');
  
  const result1 = calculateMonthlySalary('TEST_DRIVER_001', yearMonth);
  
  if (result1.ok || result1.success) {
    Logger.log('✅ 計算成功');
    Logger.log(`   員工: ${result1.data.employeeName}`);
    Logger.log(`   薪資類型: ${result1.data.salaryType}`);
    Logger.log(`   工時類型: ${result1.data.workTimeType}`);
    Logger.log(`   基本薪資: $${result1.data.baseSalary}`);
    Logger.log(`   平日加班費: $${result1.data.weekdayOvertimePay}`);
    Logger.log(`   月休補薪: $${result1.data.monthlyRestPay}`);
    Logger.log(`   特休未休補薪: $${result1.data.unusedLeavePay}`);
    Logger.log(`   應發總額: $${result1.data.grossSalary}`);
    Logger.log(`   實發金額: $${result1.data.netSalary}`);
    
    // 儲存
    const saveResult1 = saveMonthlySalary(result1.data);
    Logger.log(`   儲存結果: ${saveResult1.success ? '✅ 成功' : '❌ 失敗'}`);
  } else {
    Logger.log('❌ 計算失敗: ' + (result1.msg || result1.message));
  }
  
  Logger.log('');
  
  // ========================================
  // 測試情況二：食品廠移工
  // ========================================
  Logger.log('═══════════════════════════════════════');
  Logger.log('📝 測試情況二：食品廠移工');
  Logger.log('═══════════════════════════════════════');
  
  const result2 = calculateMonthlySalary('TEST_WORKER_001', yearMonth);
  
  if (result2.ok || result2.success) {
    Logger.log('✅ 計算成功');
    Logger.log(`   員工: ${result2.data.employeeName}`);
    Logger.log(`   薪資類型: ${result2.data.salaryType}`);
    Logger.log(`   工時類型: ${result2.data.workTimeType}`);
    Logger.log(`   時薪: $${result2.data.hourlyRate}`);
    Logger.log(`   工作時數: ${result2.data.totalWorkHours}h`);
    Logger.log(`   基本薪資: $${result2.data.baseSalary}`);
    Logger.log(`   延長工時加班費: $${result2.data.extendedOvertimePay}`);
    Logger.log(`   特休未休補薪: $${result2.data.unusedLeavePay}`);
    Logger.log(`   應發總額: $${result2.data.grossSalary}`);
    Logger.log(`   實發金額: $${result2.data.netSalary}`);
    
    // 儲存
    const saveResult2 = saveMonthlySalary(result2.data);
    Logger.log(`   儲存結果: ${saveResult2.success ? '✅ 成功' : '❌ 失敗'}`);
  } else {
    Logger.log('❌ 計算失敗: ' + (result2.msg || result2.message));
  }
  
  Logger.log('');
  
  // ========================================
  // 測試情況三：管理部行政
  // ========================================
  Logger.log('═══════════════════════════════════════');
  Logger.log('📝 測試情況三：管理部行政');
  Logger.log('═══════════════════════════════════════');
  
  const result3 = calculateMonthlySalary('TEST_ADMIN_001', yearMonth);
  
  if (result3.ok || result3.success) {
    Logger.log('✅ 計算成功');
    Logger.log(`   員工: ${result3.data.employeeName}`);
    Logger.log(`   薪資類型: ${result3.data.salaryType}`);
    Logger.log(`   工時類型: ${result3.data.workTimeType}`);
    Logger.log(`   基本薪資: $${result3.data.baseSalary}`);
    Logger.log(`   延長工時（前2h）: $${result3.data.extendedOvertimeFirst2h}`);
    Logger.log(`   延長工時（後2h）: $${result3.data.extendedOvertimeAfter2h}`);
    Logger.log(`   特休未休補薪: $${result3.data.unusedLeavePay}`);
    Logger.log(`   不固定獎金: $${result3.data.monthlyBonus}`);
    Logger.log(`   應發總額: $${result3.data.grossSalary}`);
    Logger.log(`   實發金額: $${result3.data.netSalary}`);
    
    // 儲存
    const saveResult3 = saveMonthlySalary(result3.data);
    Logger.log(`   儲存結果: ${saveResult3.success ? '✅ 成功' : '❌ 失敗'}`);
  } else {
    Logger.log('❌ 計算失敗: ' + (result3.msg || result3.message));
  }
  
  Logger.log('');
  
  // ========================================
  // 總結
  // ========================================
  Logger.log('═══════════════════════════════════════');
  Logger.log('📊 測試總結');
  Logger.log('═══════════════════════════════════════');
  
  const success1 = result1.ok || result1.success;
  const success2 = result2.ok || result2.success;
  const success3 = result3.ok || result3.success;
  
  Logger.log(`   情況一（飼料廠司機）: ${success1 ? '✅ 成功' : '❌ 失敗'}`);
  Logger.log(`   情況二（食品廠移工）: ${success2 ? '✅ 成功' : '❌ 失敗'}`);
  Logger.log(`   情況三（管理部行政）: ${success3 ? '✅ 成功' : '❌ 失敗'}`);
  Logger.log('');
  
  if (success1 && success2 && success3) {
    Logger.log('✅✅✅ 所有測試通過！系統運作正常！');
  } else {
    Logger.log('⚠️ 部分測試失敗，請檢查錯誤訊息');
  }
  
  Logger.log('');
  Logger.log('⚠️ 下一步：檢查月薪資記錄工作表');
  Logger.log('   請手動開啟「月薪資記錄」工作表，確認資料正確');
}

// ==================================================================================
// 步驟 3：批次計算所有員工薪資
// ==================================================================================

/**
 * 🔄 批次計算所有在職員工的薪資
 */
function batchCalculateAllEmployees() {
  const yearMonth = '2026-01';
  
  Logger.log('🔄 開始批次計算所有員工薪資...');
  Logger.log(`   年月: ${yearMonth}`);
  Logger.log('');
  
  const result = batchCalculateSalary(yearMonth);
  
  if (result.ok) {
    Logger.log('✅ 批次計算完成');
    Logger.log(`   成功: ${result.data.success.length} 位`);
    Logger.log(`   失敗: ${result.data.failed.length} 位`);
  } else {
    Logger.log('❌ 批次計算失敗: ' + result.msg);
  }
}

// ==================================================================================
// 完整執行流程（一鍵執行）
// ==================================================================================

/**
 * ⭐ 一鍵執行完整測試流程
 */
function runFullTest() {
  Logger.log('⭐⭐⭐ 開始執行完整測試流程 ⭐⭐⭐');
  Logger.log('');
  
  try {
    // 步驟 0：初始化
    Logger.log('🔧 步驟 0: 初始化系統');
    initializeAllSheets();
    Logger.log('');
    
    // 步驟 1：設定測試資料
    Logger.log('🧪 步驟 1: 設定測試資料');
    setupTestData();
    Logger.log('');
    
    // 步驟 2：測試所有情況
    Logger.log('🧪 步驟 2: 測試所有情況');
    testAllCases();
    Logger.log('');
    
    Logger.log('═══════════════════════════════════════');
    Logger.log('✅✅✅ 完整測試流程執行完成！');
    Logger.log('═══════════════════════════════════════');
    Logger.log('');
    Logger.log('📋 請檢查以下工作表：');
    Logger.log('   1. 員工薪資設定 - 確認3位測試員工已建立');
    Logger.log('   2. 月薪資記錄 - 確認3筆薪資單已生成');
    Logger.log('   3. 未休假記錄 - 確認未休假資料');
    Logger.log('   4. 月休記錄 - 確認月休資料（情況一）');
    Logger.log('   5. 不固定獎金記錄 - 確認獎金資料（情況三）');
    
  } catch (error) {
    Logger.log('❌ 測試流程發生錯誤: ' + error.message);
    Logger.log('❌ 錯誤堆疊: ' + error.stack);
  }
}

console.log('✅ 薪資計算系統完整整合測試指引已載入');
console.log('');
console.log('📋 使用說明：');
console.log('   1. 執行 runFullTest() - 一鍵執行完整測試');
console.log('   2. 或分步執行：');
console.log('      - initializeAllSheets() - 初始化系統');
console.log('      - setupTestData() - 設定測試資料');
console.log('      - testAllCases() - 測試所有情況');



/**
 * ✅ 正確的測試方式：使用路由系統自動判斷
 */
function testSalaryWithRouter() {
  Logger.log('═══════════════════════════════════════');
  Logger.log('🧪 測試薪資計算（使用路由系統）');
  Logger.log('═══════════════════════════════════════');
  Logger.log('');
  
  const employeeId = 'Uf664a35632b736301d674d8b2cc3f8c0';
  const yearMonth = '2026-01';
  
  Logger.log(`📝 測試員工: ${employeeId}`);
  Logger.log(`   年月: ${yearMonth}`);
  Logger.log('');
  
  // ⭐⭐⭐ 使用路由系統，自動判斷員工類型
  const result = calculateMonthlySalary(employeeId, yearMonth);
  
  if (result.ok || result.success) {
    Logger.log('✅ 計算成功！');
    Logger.log('');
    Logger.log('📊 基本資訊:');
    Logger.log(`   員工姓名: ${result.data.employeeName}`);
    Logger.log(`   員工類型: ${result.data.employeeType}`);
    Logger.log(`   薪資類型: ${result.data.salaryType}`);
    Logger.log(`   工時類型: ${result.data.workTimeType}`);
    Logger.log('');
    
    // ⭐ 根據薪資類型顯示不同資訊
    if (result.data.salaryType === '時薪') {
      Logger.log('💰 時薪資訊:');
      Logger.log(`   時薪: $${result.data.hourlyRate}`);
      Logger.log(`   工作時數: ${result.data.totalWorkHours}h`);
      Logger.log(`   基本薪資（時薪×工時）: $${result.data.baseSalary}`);
      Logger.log('');
      
      if (result.data.extendedOvertimePay || result.data.restdayOvertimePay || result.data.holidayOvertimePay) {
        Logger.log('💼 加班費:');
        if (result.data.extendedOvertimePay) {
          Logger.log(`   延長工時加班費: $${result.data.extendedOvertimePay}`);
        }
        if (result.data.restdayOvertimePay) {
          Logger.log(`   休息日加班費: $${result.data.restdayOvertimePay}`);
        }
        if (result.data.holidayOvertimePay) {
          Logger.log(`   國定假日加班費: $${result.data.holidayOvertimePay}`);
        }
        Logger.log('');
      }
    } else {
      Logger.log('💰 月薪資訊:');
      Logger.log(`   基本薪資: $${result.data.baseSalary}`);
      Logger.log('');
      
      if (result.data.extendedOvertimeFirst2h || result.data.extendedOvertimeAfter2h) {
        Logger.log('💼 加班費（情況三）:');
        if (result.data.extendedOvertimeFirst2h) {
          Logger.log(`   延長工時（前2h）: $${result.data.extendedOvertimeFirst2h}`);
        }
        if (result.data.extendedOvertimeAfter2h) {
          Logger.log(`   延長工時（後2h）: $${result.data.extendedOvertimeAfter2h}`);
        }
        if (result.data.restdayOvertimePay) {
          Logger.log(`   休息日加班費: $${result.data.restdayOvertimePay}`);
        }
        if (result.data.holidayOvertimePay) {
          Logger.log(`   國定假日加班費: $${result.data.holidayOvertimePay}`);
        }
        Logger.log('');
      } else if (result.data.weekdayOvertimePay || result.data.restdayOvertimePay || result.data.holidayOvertimePay) {
        Logger.log('💼 加班費（情況一）:');
        if (result.data.weekdayOvertimePay) {
          Logger.log(`   平日加班費: $${result.data.weekdayOvertimePay}`);
        }
        if (result.data.restdayOvertimePay) {
          Logger.log(`   休息日加班費: $${result.data.restdayOvertimePay}`);
        }
        if (result.data.holidayOvertimePay) {
          Logger.log(`   國定假日加班費: $${result.data.holidayOvertimePay}`);
        }
        if (result.data.monthlyRestPay) {
          Logger.log(`   月休補薪: $${result.data.monthlyRestPay}`);
        }
        Logger.log('');
      }
    }
    
    Logger.log('📊 薪資總計:');
    Logger.log(`   應發總額: $${result.data.grossSalary}`);
    Logger.log(`   扣款總額: $${result.data.grossSalary - result.data.netSalary}`);
    Logger.log(`   實發金額: $${result.data.netSalary}`);
    Logger.log('');
    
    Logger.log('═══════════════════════════════════════');
    Logger.log('✅✅✅ 測試完成！');
    Logger.log('═══════════════════════════════════════');
    
  } else {
    Logger.log('❌ 計算失敗: ' + (result.msg || result.message));
  }
}

/**
 * 🧪 對比測試：錯誤方式 vs 正確方式
 */
function compareTestMethods() {
  Logger.log('🧪 對比測試：錯誤方式 vs 正確方式');
  Logger.log('');
  
  const employeeId = 'Uf664a35632b736301d674d8b2cc3f8c0';
  const yearMonth = '2026-01';
  
  // ========================================
  // ❌ 錯誤方式：直接呼叫情況三函數
  // ========================================
  Logger.log('═══════════════════════════════════════');
  Logger.log('❌ 錯誤方式：直接呼叫 calculateAdminSalary()');
  Logger.log('═══════════════════════════════════════');
  
  const wrongResult = calculateAdminSalary(employeeId, yearMonth);
  
  if (wrongResult.ok) {
    Logger.log(`   員工類型: ${wrongResult.data.employeeType}`);
    Logger.log(`   ❌ 強制使用情況三計算`);
    Logger.log(`   ❌ 時薪被誤算為: ${wrongResult.data.baseSalary / 30 / 8} 元/小時`);
    Logger.log(`   實發金額: $${wrongResult.data.netSalary}`);
  }
  
  Logger.log('');
  
  // ========================================
  // ✅ 正確方式：使用路由系統
  // ========================================
  Logger.log('═══════════════════════════════════════');
  Logger.log('✅ 正確方式：使用 calculateMonthlySalary() 路由');
  Logger.log('═══════════════════════════════════════');
  
  const correctResult = calculateMonthlySalary(employeeId, yearMonth);
  
  if (correctResult.ok) {
    Logger.log(`   員工類型: ${correctResult.data.employeeType}`);
    Logger.log(`   ✅ 自動路由到正確的計算函數`);
    
    if (correctResult.data.salaryType === '時薪') {
      Logger.log(`   ✅ 時薪: ${correctResult.data.hourlyRate} 元/小時`);
      Logger.log(`   ✅ 工作時數: ${correctResult.data.totalWorkHours}h`);
      Logger.log(`   ✅ 基本薪資: $${correctResult.data.baseSalary}`);
    }
    
    Logger.log(`   實發金額: $${correctResult.data.netSalary}`);
  }
  
  Logger.log('');
  Logger.log('═══════════════════════════════════════');
  Logger.log('📊 結果對比:');
  Logger.log('═══════════════════════════════════════');
  
  if (wrongResult.ok && correctResult.ok) {
    const diff = Math.abs(wrongResult.data.netSalary - correctResult.data.netSalary);
    
    Logger.log(`   錯誤方式實發: $${wrongResult.data.netSalary}`);
    Logger.log(`   正確方式實發: $${correctResult.data.netSalary}`);
    Logger.log(`   差額: $${diff}`);
    Logger.log('');
    
    if (diff > 0) {
      Logger.log('⚠️⚠️⚠️ 兩種方式計算結果不同！');
      Logger.log('⚠️ 請務必使用路由系統 calculateMonthlySalary()');
    } else {
      Logger.log('✅ 兩種方式結果一致');
    }
  }
}

/**
 * 📋 使用指引
 */
function showUsageGuide() {
  Logger.log('');
  Logger.log('═══════════════════════════════════════');
  Logger.log('📋 薪資計算函數使用指引');
  Logger.log('═══════════════════════════════════════');
  Logger.log('');
  Logger.log('✅ 正確方式（推薦）：');
  Logger.log('   calculateMonthlySalary(employeeId, yearMonth)');
  Logger.log('   → 自動判斷員工類型，路由到正確的計算函數');
  Logger.log('');
  Logger.log('❌ 錯誤方式（不推薦）：');
  Logger.log('   calculateDriverSalaryCase1(employeeId, yearMonth)   // 情況一專用');
  Logger.log('   calculateFactoryWorkerSalary(employeeId, yearMonth) // 情況二專用');
  Logger.log('   calculateAdminSalary(employeeId, yearMonth)         // 情況三專用');
  Logger.log('   → 強制使用特定計算邏輯，可能導致錯誤結果');
  Logger.log('');
  Logger.log('═══════════════════════════════════════');
  Logger.log('');
  Logger.log('🎯 測試函數推薦：');
  Logger.log('');
  Logger.log('   1. testSalaryWithRouter()    // 單一員工測試（使用路由）');
  Logger.log('   2. testAllCases()            // 測試所有情況');
  Logger.log('   3. runFullTest()             // 完整系統測試');
  Logger.log('   4. compareTestMethods()      // 對比錯誤vs正確方式');
  Logger.log('');
  Logger.log('═══════════════════════════════════════');
}


function testCase3WithOvertime() {
  const employeeId = 'Uf664a35632b736301d674d8b2cc3f8c0';
  const yearMonth = '2026-01';
  
  Logger.log('🧪 測試情況三加班費計算');
  Logger.log('');
  
  // 1. 先測試 getEmployeeMonthlyOvertime
  Logger.log('📝 步驟 1: 測試取得加班記錄');
  const overtimeRecords = getEmployeeMonthlyOvertime(employeeId, yearMonth);
  Logger.log(`   結果：找到 ${overtimeRecords.length} 筆記錄`);
  Logger.log('');
  
  // 2. 測試完整薪資計算
  Logger.log('📝 步驟 2: 測試薪資計算');
  const result = calculateMonthlySalary(employeeId, yearMonth);
  
  if (result.ok || result.success) {
    Logger.log('✅ 薪資計算成功！');
    Logger.log('');
    Logger.log('📊 加班費資訊：');
    Logger.log(`   總加班時數: ${result.data.totalOvertimeHours}h`);
    Logger.log(`   延長工時（前2h）: $${result.data.extendedOvertimeFirst2h}`);
    Logger.log(`   延長工時（後2h）: $${result.data.extendedOvertimeAfter2h}`);
    Logger.log(`   休息日加班費: $${result.data.restdayOvertimePay}`);
    Logger.log(`   國定假日加班費: $${result.data.holidayOvertimePay}`);
    Logger.log('');
    Logger.log('💰 薪資總計：');
    Logger.log(`   應發總額: $${result.data.grossSalary}`);
    Logger.log(`   實發金額: $${result.data.netSalary}`);
  } else {
    Logger.log('❌ 薪資計算失敗: ' + (result.msg || result.message));
  }
}


function saveSalaryAndVerify() {
  const employeeId = 'Uf664a35632b736301d674d8b2cc3f8c0';
  const yearMonth = '2026-01';
  
  Logger.log('📝 步驟 1: 計算薪資');
  const calcResult = calculateMonthlySalary(employeeId, yearMonth);
  
  if (!calcResult.ok && !calcResult.success) {
    Logger.log('❌ 計算失敗: ' + (calcResult.msg || calcResult.message));
    return;
  }
  
  Logger.log('✅ 計算成功');
  Logger.log('');
  
  Logger.log('📝 步驟 2: 儲存到月薪資記錄');
  const saveResult = saveMonthlySalary(calcResult.data);
  
  if (!saveResult.success) {
    Logger.log('❌ 儲存失敗: ' + saveResult.message);
    return;
  }
  
  Logger.log('✅ 儲存成功: ' + saveResult.salaryId);
  Logger.log('');
  
  Logger.log('📝 步驟 3: 讀取驗證');
  const readResult = getMySalary(employeeId, yearMonth);
  
  if (!readResult.success) {
    Logger.log('❌ 讀取失敗: ' + readResult.message);
    return;
  }
  
  Logger.log('✅ 讀取成功');
  Logger.log('');
  Logger.log('═══════════════════════════════════════');
  Logger.log('📊 薪資單驗證');
  Logger.log('═══════════════════════════════════════');
  Logger.log(`員工姓名: ${readResult.data['員工姓名']}`);
  Logger.log(`年月: ${readResult.data['年月']}`);
  Logger.log('');
  Logger.log('💰 加班費明細:');
  Logger.log(`   總加班時數: ${readResult.data['總加班時數']}h`);
  Logger.log(`   平日加班費/延長工時(前2h): $${readResult.data['平日加班費/延長工時(前2h)']}`);
  Logger.log(`   休息日加班費/延長工時(後2h): $${readResult.data['休息日加班費/延長工時(後2h)']}`);
  Logger.log(`   國定假日加班費: $${readResult.data['國定假日加班費']}`);
  Logger.log('');
  Logger.log('💵 薪資總計:');
  Logger.log(`   應發總額: $${readResult.data['應發總額']}`);
  Logger.log(`   實發金額: $${readResult.data['實發金額']}`);
  Logger.log('═══════════════════════════════════════');
  
  // 驗證金額是否正確
  const isCorrect = (
    readResult.data['總加班時數'] === 6 &&
    readResult.data['平日加班費/延長工時(前2h)'] === 782 &&
    readResult.data['休息日加班費/延長工時(後2h)'] === 782 &&
    readResult.data['國定假日加班費'] === 2333 &&
    readResult.data['應發總額'] === 90440 &&
    readResult.data['實發金額'] === 83983
  );
  
  Logger.log('');
  if (isCorrect) {
    Logger.log('✅✅✅ 所有數據驗證成功！');
  } else {
    Logger.log('⚠️ 部分數據不符，請檢查');
  }
}