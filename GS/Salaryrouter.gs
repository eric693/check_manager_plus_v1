// ==================================================================================
// 薪資計算路由系統（整合情況一、二、三）
// ==================================================================================

/**
 * ✅ 統一的薪資計算入口（自動路由）
 * 
 * 根據員工類型自動選擇對應的計算函數：
 * - 情況一：飼料廠司機 + 月薪制 + 不定時 → calculateDriverSalaryCase1
 * - 情況二：食品廠移工 + 時薪制 + 標準工時 → calculateFactoryWorkerSalary
 * - 情況三：管理部行政 + 月薪制 + 標準工時 → calculateAdminSalary
 * 
 * @param {string} employeeId - 員工ID
 * @param {string} yearMonth - 年月 (YYYY-MM)
 * @returns {Object} 薪資計算結果
 */
function calculateMonthlySalary(employeeId, yearMonth) {
  try {
    Logger.log('═══════════════════════════════════════');
    Logger.log('🎯 薪資計算路由系統');
    Logger.log(`   員工ID: ${employeeId}`);
    Logger.log(`   年月: ${yearMonth}`);
    Logger.log('═══════════════════════════════════════');
    
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
    
    // ========================================
    // 步驟 2：判斷員工類型並路由
    // ========================================
    const employeeType = employeeConfig[3] || '';  // 第4欄：員工類型
    const salaryType = employeeConfig[4] || '';    // 第5欄：薪資類型
    const workTimeType = employeeConfig[5] || '';  // 第6欄：工時類型
    
    Logger.log(`📋 員工資訊:`);
    Logger.log(`   員工類型: ${employeeType}`);
    Logger.log(`   薪資類型: ${salaryType}`);
    Logger.log(`   工時類型: ${workTimeType}`);
    Logger.log('');
    
    // ⭐⭐⭐ 路由邏輯（三種情況）
    
    // 情況一：飼料廠司機 + 月薪制 + 不定時
    if (employeeType === '飼料廠司機' && salaryType === '月薪' && workTimeType === '不定時') {
      Logger.log('🚛 路由到：情況一 - 飼料廠司機計算');
      return calculateDriverSalaryCase1(employeeId, yearMonth);
    }
    
    // 情況二：食品廠移工 + 時薪制 + 標準工時
    if (employeeType === '食品廠移工' && salaryType === '時薪' && workTimeType === '標準工時') {
      Logger.log('🏭 路由到：情況二 - 食品廠移工計算');
      return calculateFactoryWorkerSalary(employeeId, yearMonth);
    }
    
    // 情況三：管理部行政 + 月薪制 + 標準工時
    if (employeeType === '管理部行政' && salaryType === '月薪' && workTimeType === '標準工時') {
      Logger.log('💼 路由到：情況三 - 管理部行政計算');
      return calculateAdminSalary(employeeId, yearMonth);
    }
    
    // ========================================
    // 預設邏輯：未知類型
    // ========================================
    Logger.log('⚠️ 未知員工類型組合，使用預設計算');
    Logger.log(`   ${employeeType} + ${salaryType} + ${workTimeType}`);
    
    return {
      ok: false,
      msg: `該員工類型組合尚未設定專屬計算邏輯：${employeeType} + ${salaryType} + ${workTimeType}`,
      code: 'UNKNOWN_EMPLOYEE_TYPE'
    };
    
  } catch (error) {
    Logger.log('❌ 薪資計算路由失敗: ' + error.message);
    Logger.log('❌ 錯誤堆疊: ' + error.stack);
    
    return {
      ok: false,
      msg: "計算失敗：" + error.message
    };
  }
}


/**
 * 🧪 測試路由功能（完整版）
 */
function testSalaryRouterComplete() {
  Logger.log('🧪 測試薪資計算路由（三種情況）');
  Logger.log('');
  
  // 測試情況一：飼料廠司機
  Logger.log('📝 測試 1: 飼料廠司機（月薪制-不定時）');
  const result1 = calculateMonthlySalary('Uf664a35632b736301d674d8b2cc3f8c0', '2026-01');
  Logger.log(`   結果: ${result1.ok ? '✅ 成功' : '❌ 失敗'}`);
  if (result1.ok) {
    Logger.log(`   員工: ${result1.data.employeeName}`);
    Logger.log(`   實發: $${result1.data.netSalary}`);
  }
  Logger.log('');
  
  // 測試情況二：食品廠移工（需要實際的移工ID）
  Logger.log('📝 測試 2: 食品廠移工（時薪制-標準工時）');
  Logger.log('   （請在「員工薪資設定」中新增測試員工）');
  Logger.log('');
  
  // 測試情況三：管理部行政（需要實際的行政ID）
  Logger.log('📝 測試 3: 管理部行政（月薪制-標準工時）');
  Logger.log('   （請在「員工薪資設定」中新增測試員工）');
  Logger.log('');
  
  Logger.log('✅ 路由測試完成');
}

/**
 * 🔧 批次計算所有員工薪資（管理員專用）
 * 
 * @param {string} yearMonth - 年月 (YYYY-MM)
 * @returns {Object} 批次計算結果
 */
function batchCalculateSalary(yearMonth) {
  try {
    Logger.log('═══════════════════════════════════════');
    Logger.log('🔄 開始批次計算薪資');
    Logger.log(`   年月: ${yearMonth}`);
    Logger.log('═══════════════════════════════════════');
    Logger.log('');
    
    const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('員工薪資設定');
    if (!configSheet) {
      return { ok: false, msg: "找不到員工薪資設定工作表" };
    }
    
    const configData = configSheet.getDataRange().getValues();
    
    const results = {
      success: [],
      failed: [],
      total: 0
    };
    
    // 遍歷所有員工（跳過標題列）
    for (let i = 1; i < configData.length; i++) {
      const employeeId = configData[i][0];
      const employeeName = configData[i][1];
      const status = configData[i][30]; // 狀態欄位
      
      // 只計算在職員工
      if (status !== '在職') {
        Logger.log(`⏭️ 跳過非在職員工: ${employeeName} (${employeeId})`);
        continue;
      }
      
      results.total++;
      
      Logger.log(`\n📝 計算 ${results.total}: ${employeeName} (${employeeId})`);
      
      // 計算薪資
      const calcResult = calculateMonthlySalary(employeeId, yearMonth);
      
      if (calcResult.ok || calcResult.success) {
        // 儲存薪資
        const saveResult = saveMonthlySalary(calcResult.data);
        
        if (saveResult.success) {
          results.success.push({
            employeeId: employeeId,
            employeeName: employeeName,
            salaryId: saveResult.salaryId
          });
          Logger.log(`   ✅ 成功: ${saveResult.salaryId}`);
        } else {
          results.failed.push({
            employeeId: employeeId,
            employeeName: employeeName,
            error: saveResult.message
          });
          Logger.log(`   ❌ 儲存失敗: ${saveResult.message}`);
        }
      } else {
        results.failed.push({
          employeeId: employeeId,
          employeeName: employeeName,
          error: calcResult.msg || calcResult.message
        });
        Logger.log(`   ❌ 計算失敗: ${calcResult.msg || calcResult.message}`);
      }
    }
    
    Logger.log('');
    Logger.log('═══════════════════════════════════════');
    Logger.log('📊 批次計算結果統計');
    Logger.log('═══════════════════════════════════════');
    Logger.log(`   總計: ${results.total} 位員工`);
    Logger.log(`   成功: ${results.success.length} 位`);
    Logger.log(`   失敗: ${results.failed.length} 位`);
    
    if (results.failed.length > 0) {
      Logger.log('');
      Logger.log('❌ 失敗清單:');
      results.failed.forEach(item => {
        Logger.log(`   - ${item.employeeName} (${item.employeeId}): ${item.error}`);
      });
    }
    
    Logger.log('═══════════════════════════════════════');
    
    return {
      ok: true,
      data: results,
      message: `批次計算完成：成功 ${results.success.length}/${results.total}`
    };
    
  } catch (error) {
    Logger.log('❌ 批次計算失敗: ' + error.message);
    return {
      ok: false,
      msg: error.message
    };
  }
}

/**
 * 🧪 測試批次計算
 */
function testBatchCalculate() {
  const yearMonth = '2026-01';
  const result = batchCalculateSalary(yearMonth);
  
  Logger.log('');
  Logger.log('📤 批次計算結果:');
  Logger.log(`   ok: ${result.ok}`);
  Logger.log(`   message: ${result.message}`);
}

console.log('✅ 薪資計算路由系統已載入（含情況一、二、三）');

/**
 * ✅ 預設薪資計算（適用於未知類型）
 */
function calculateDefaultSalary(employeeId, yearMonth) {
  Logger.log('⚠️ 使用預設計算邏輯（基本功能）');
  
  // TODO: 實作基本的薪資計算
  // 這裡可以放一個最簡單的計算邏輯
  
  return {
    ok: false,
    msg: "該員工類型尚未設定專屬計算邏輯，請聯繫系統管理員"
  };
}

/**
 * 🧪 測試路由功能
 */
function testSalaryRouter() {
  Logger.log('🧪 測試薪資計算路由');
  Logger.log('');
  
  // 測試情況一
  Logger.log('📝 測試 1: 飼料廠司機');
  const result1 = calculateMonthlySalary('Uf664a35632b736301d674d8b2cc3f8c0', '2026-01');
  Logger.log(`   結果: ${result1.ok ? '成功' : '失敗'}`);
  Logger.log('');
  
  // 測試情況二（需要實際的食品廠移工ID）
  Logger.log('📝 測試 2: 食品廠移工');
  Logger.log('   （請替換成實際的移工員工ID）');
  Logger.log('');
  
  Logger.log('✅ 路由測試完成');
}

console.log('✅ 薪資計算路由系統已載入');