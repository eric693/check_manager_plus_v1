// ==================================================================================
// 情況二：食品廠移工 - 輔助函數
// ==================================================================================

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
 * ✅ 取得員工月度加班記錄
 * 
 * @param {string} employeeId - 員工ID
 * @param {string} yearMonth - 年月 (YYYY-MM)
 * @returns {Array} 加班記錄陣列
 */
function getEmployeeMonthlyOvertime(employeeId, yearMonth) {
  try {
    Logger.log('📋 取得加班記錄');
    Logger.log(`   員工ID: ${employeeId}`);
    Logger.log(`   年月: ${yearMonth}`);
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('加班申請');
    
    if (!sheet) {
      Logger.log('⚠️ 找不到「加班申請」工作表');
      return [];
    }
    
    const data = sheet.getDataRange().getValues();
    const records = [];
    
    // 遍歷加班記錄
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      const recordEmployeeId = String(row[1] || '').trim();
      const overtimeDate = row[4];
      const hours = parseFloat(row[7]) || 0;
      const status = String(row[9] || '').trim();
      
      // 只統計已核准的加班
      if (status !== 'APPROVED') continue;
      
      // 判斷是否為目標月份
      let recordYearMonth = '';
      
      if (overtimeDate instanceof Date) {
        recordYearMonth = Utilities.formatDate(overtimeDate, 'Asia/Taipei', 'yyyy-MM');
      } else if (typeof overtimeDate === 'string') {
        recordYearMonth = overtimeDate.substring(0, 7);
      }
      
      // 篩選條件
      if (recordEmployeeId === employeeId && recordYearMonth === yearMonth) {
        records.push({
          date: overtimeDate instanceof Date 
            ? Utilities.formatDate(overtimeDate, 'Asia/Taipei', 'yyyy-MM-dd') 
            : overtimeDate,
          hours: hours
        });
      }
    }
    
    Logger.log(`✅ 找到 ${records.length} 筆加班記錄`);
    
    return records;
    
  } catch (error) {
    Logger.log('❌ 取得加班記錄失敗: ' + error.message);
    return [];
  }
}

/**
 * ✅ 判斷日期類型（平日/休息日/假日）
 * 
 * @param {string} dateStr - 日期字串 (YYYY-MM-DD)
 * @returns {string} 'weekday' | 'restday' | 'holiday'
 */
function getDateType(dateStr) {
  try {
    let date;
    
    if (typeof dateStr === 'string') {
      date = new Date(dateStr);
    } else if (dateStr instanceof Date) {
      date = dateStr;
    } else {
      Logger.log('⚠️ 無效的日期格式: ' + dateStr);
      return 'weekday';
    }
    
    // 取得星期幾 (0=週日, 6=週六)
    const dayOfWeek = date.getDay();
    
    // ⭐ 判斷規則
    if (dayOfWeek === 0) {
      // 週日 = 例假日
      return 'holiday';
    } else if (dayOfWeek === 6) {
      // 週六 = 休息日
      return 'restday';
    } else {
      // 週一～五 = 平日
      return 'weekday';
    }
    
  } catch (error) {
    Logger.log('❌ 判斷日期類型失敗: ' + error.message);
    return 'weekday';  // 預設為平日
  }
}

console.log('✅ 情況二輔助函數已載入');




function testCase1() {
  const employeeId = 'Uf664a35632b736301d674d8b2cc3f8c0';  // 飼料廠司機
  const yearMonth = '2026-01';
  
  const result = calculateMonthlySalary(employeeId, yearMonth);
  
  Logger.log('情況一測試結果:');
  Logger.log('   ok: ' + result.ok);
  Logger.log('   員工: ' + result.data.employeeName);
  Logger.log('   工時類型: ' + result.data.workTimeType);
  Logger.log('   實發金額: ' + result.data.netSalary);
}

/**
 * 🧪 測試完整的時薪薪資計算（情況二）
 */
function testHourlySalaryComplete() {
  const employeeId = 'Uf664a35632b736301d674d8b2cc3f8c0';
  const yearMonth = '2026-01';
  
  Logger.log('═══════════════════════════════════════');
  Logger.log('🧪 測試完整時薪薪資計算（情況二）');
  Logger.log('═══════════════════════════════════════');
  Logger.log('');
  
  // 步驟 1: 計算薪資
  Logger.log('📝 步驟 1: 計算時薪薪資...');
  Logger.log('');
  
  const calcResult = calculateHourlySalary(employeeId, yearMonth);
  
  if (!calcResult.success) {
    Logger.log('❌ 計算失敗: ' + calcResult.message);
    return;
  }
  
  Logger.log('✅ 計算成功！');
  Logger.log('');
  
  // 步驟 2: 顯示計算結果摘要
  Logger.log('📊 計算結果摘要:');
  Logger.log('   員工: ' + calcResult.data.employeeName);
  Logger.log('   時薪: $' + calcResult.data.hourlyRate);
  Logger.log('   工作時數: ' + calcResult.data.totalWorkHours + 'h');
  Logger.log('   基本薪資: $' + calcResult.data.baseSalary);
  Logger.log('   加班時數: ' + calcResult.data.totalOvertimeHours.toFixed(1) + 'h');
  Logger.log('   - 平日加班費: $' + calcResult.data.weekdayOvertimePay);
  Logger.log('   - 休息日加班費: $' + calcResult.data.restdayOvertimePay);
  Logger.log('   - 例假日加班費: $' + calcResult.data.holidayOvertimePay);
  Logger.log('   應發總額: $' + calcResult.data.grossSalary);
  Logger.log('   扣款總額: $' + (calcResult.data.grossSalary - calcResult.data.netSalary));
  Logger.log('   實發金額: $' + calcResult.data.netSalary);
  Logger.log('');
  
  // 步驟 3: 存檔到月薪資記錄
  Logger.log('📝 步驟 2: 儲存到月薪資記錄...');
  Logger.log('');
  
  const saveResult = saveMonthlySalary(calcResult.data);
  
  if (!saveResult.success) {
    Logger.log('❌ 儲存失敗: ' + saveResult.message);
    return;
  }
  
  Logger.log('✅ 儲存成功！薪資單ID: ' + saveResult.salaryId);
  Logger.log('');
  
  // 步驟 4: 讀取驗證
  Logger.log('📝 步驟 3: 讀取驗證...');
  Logger.log('');
  
  const readResult = getMySalary(employeeId, yearMonth);
  
  if (!readResult.success) {
    Logger.log('❌ 讀取失敗: ' + readResult.message);
    return;
  }
  
  Logger.log('✅ 讀取成功！');
  Logger.log('');
  Logger.log('📋 驗證關鍵欄位:');
  Logger.log('   薪資類型: ' + readResult.data['薪資類型']);
  Logger.log('   時薪: $' + readResult.data['時薪']);
  Logger.log('   工作時數: ' + readResult.data['工作時數'] + 'h');
  Logger.log('   基本薪資: $' + readResult.data['基本薪資']);
  Logger.log('   平日加班費: $' + readResult.data['平日加班費']);
  Logger.log('   應發總額: $' + readResult.data['應發總額']);
  Logger.log('   實發金額: $' + readResult.data['實發金額']);
  Logger.log('');
  
  // 步驟 5: 最終驗證
  const isCorrect = (
    readResult.data['薪資類型'] === '時薪' &&
    readResult.data['時薪'] === calcResult.data.hourlyRate &&
    readResult.data['工作時數'] === calcResult.data.totalWorkHours &&
    readResult.data['基本薪資'] === calcResult.data.baseSalary &&
    readResult.data['應發總額'] === calcResult.data.grossSalary &&
    readResult.data['實發金額'] === calcResult.data.netSalary
  );
  
  Logger.log('');
  Logger.log('═══════════════════════════════════════');
  
  if (isCorrect) {
    Logger.log('✅✅✅ 測試完全成功！所有數據正確！');
    Logger.log('');
    Logger.log('🎯 下一步：前端測試');
    Logger.log('   1. 前往薪資管理頁面');
    Logger.log('   2. 選擇「我的薪資」Tab');
    Logger.log('   3. 選擇 2026-01 月份');
    Logger.log('   4. 檢查顯示的薪資資料');
  } else {
    Logger.log('⚠️ 資料有差異，請檢查：');
    Logger.log('');
    Logger.log('計算結果 vs 讀取結果:');
    Logger.log('   薪資類型: ' + calcResult.data.salaryType + ' vs ' + readResult.data['薪資類型']);
    Logger.log('   時薪: $' + calcResult.data.hourlyRate + ' vs $' + readResult.data['時薪']);
    Logger.log('   工作時數: ' + calcResult.data.totalWorkHours + 'h vs ' + readResult.data['工作時數'] + 'h');
    Logger.log('   基本薪資: $' + calcResult.data.baseSalary + ' vs $' + readResult.data['基本薪資']);
    Logger.log('   應發總額: $' + calcResult.data.grossSalary + ' vs $' + readResult.data['應發總額']);
    Logger.log('   實發金額: $' + calcResult.data.netSalary + ' vs $' + readResult.data['實發金額']);
  }
  
  Logger.log('═══════════════════════════════════════');
}


function diagnoseEmployeeType() {
  const employeeId = 'Uf664a35632b736301d674d8b2cc3f8c0';  // ⚠️ 替換成你的員工ID
  
  const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('員工薪資設定');
  const configData = configSheet.getDataRange().getValues();
  
  for (let i = 1; i < configData.length; i++) {
    if (configData[i][0] === employeeId) {
      Logger.log('找到員工資料：');
      Logger.log('   員工ID: ' + configData[i][0]);
      Logger.log('   員工姓名: ' + configData[i][1]);
      Logger.log('   員工類型: "' + configData[i][3] + '"');
      Logger.log('   薪資類型: "' + configData[i][4] + '"');
      Logger.log('   工時類型: "' + configData[i][5] + '"');
      
      // 判斷會路由到哪個計算函數
      const employeeType = configData[i][3] || '';
      const salaryType = configData[i][4] || '';
      const workTimeType = configData[i][5] || '';
      
      if (employeeType === '飼料廠司機' && salaryType === '月薪' && workTimeType === '不定時') {
        Logger.log('   ✅ 會路由到：情況一');
      } else if (employeeType === '食品廠移工' && salaryType === '時薪' && workTimeType === '標準工時') {
        Logger.log('   ✅ 會路由到：情況二');
      } else if (employeeType === '管理部行政' && salaryType === '月薪' && workTimeType === '標準工時') {
        Logger.log('   ✅ 會路由到：情況三');
      } else {
        Logger.log('   ❌ 不符合任何情況！會返回錯誤！');
        Logger.log('   → 建議使用「方案一」的修正程式碼');
      }
      
      break;
    }
  }
}


function checkOvertimeStatus() {
  const employeeId = 'Uf664a35632b736301d674d8b2cc3f8c0';
  const yearMonth = '2026-01';
  
  Logger.log('🔍 檢查加班申請狀態...');
  Logger.log('');
  
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('加班申請');
  
  if (!sheet) {
    Logger.log('❌ 找不到「加班申請」工作表');
    return;
  }
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  Logger.log('📋 欄位清單：' + headers.join(', '));
  Logger.log('');
  
  Logger.log('📊 該員工的所有加班記錄：');
  Logger.log('');
  
  let foundCount = 0;
  let approvedCount = 0;
  
  for (let i = 1; i < data.length; i++) {
    const recordEmployeeId = String(data[i][1] || '').trim();  // B欄：員工ID
    
    if (recordEmployeeId !== employeeId) continue;
    
    foundCount++;
    
    const overtimeDate = data[i][3];  // D欄：加班日期
    const hours = data[i][6];         // G欄：加班時數
    const status = String(data[i][9] || '').trim().toLowerCase();  // J欄：審核狀態
    
    // 處理 Date 對象
    let dateStr = '';
    if (overtimeDate instanceof Date) {
      dateStr = Utilities.formatDate(overtimeDate, 'Asia/Taipei', 'yyyy-MM-dd');
    } else {
      dateStr = String(overtimeDate);
    }
    
    const recordYearMonth = dateStr.substring(0, 7);
    
    Logger.log(`第 ${i + 1} 列：`);
    Logger.log(`   日期: ${dateStr}`);
    Logger.log(`   年月: ${recordYearMonth}`);
    Logger.log(`   時數: ${hours}`);
    Logger.log(`   狀態: "${status}" ${status === 'approved' ? '✅' : '❌'}`);
    
    if (status === 'approved' && recordYearMonth === yearMonth) {
      approvedCount++;
      Logger.log('   → 會被計入加班費 ✅');
    } else {
      Logger.log(`   → 不會計入 (${status !== 'approved' ? '未核准' : '不同月份'})`);
    }
    
    Logger.log('');
  }
  
  Logger.log('═══════════════════════════════════════');
  Logger.log('📊 統計結果：');
  Logger.log(`   找到記錄數: ${foundCount}`);
  Logger.log(`   ${yearMonth} 已核准: ${approvedCount}`);
  Logger.log('═══════════════════════════════════════');
  
  if (approvedCount === 0) {
    Logger.log('');
    Logger.log('⚠️ 可能原因：');
    Logger.log('   1. 加班申請狀態不是 "approved"（注意小寫）');
    Logger.log('   2. 加班記錄不在 2026-01 月份');
    Logger.log('   3. 該員工該月份確實沒有加班');
  }
}


function checkEmployeeRouting() {
  const employeeId = 'Uf664a35632b736301d674d8b2cc3f8c0';
  
  const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('員工薪資設定');
  const configData = configSheet.getDataRange().getValues();
  
  for (let i = 1; i < configData.length; i++) {
    if (configData[i][0] === employeeId) {
      const employeeType = configData[i][3] || '';
      const salaryType = configData[i][4] || '';
      const workTimeType = configData[i][5] || '';
      
      Logger.log('═══════════════════════════════════════');
      Logger.log('📋 員工薪資設定：');
      Logger.log('═══════════════════════════════════════');
      Logger.log(`   員工ID: ${employeeId}`);
      Logger.log(`   員工姓名: ${configData[i][1]}`);
      Logger.log(`   員工類型: "${employeeType}"`);
      Logger.log(`   薪資類型: "${salaryType}"`);
      Logger.log(`   工時類型: "${workTimeType}"`);
      Logger.log('');
      
      // 判斷會路由到哪裡
      Logger.log('🎯 路由判斷：');
      
      if (employeeType === '飼料廠司機' && salaryType === '月薪' && workTimeType === '不定時') {
        Logger.log('   ✅ 符合情況一：飼料廠司機');
        Logger.log('   → 會執行 calculateDriverSalaryCase1()');
      } else if (employeeType === '食品廠移工' && salaryType === '時薪' && workTimeType === '標準工時') {
        Logger.log('   ✅ 符合情況二：食品廠移工');
        Logger.log('   → 會執行 calculateFactoryWorkerSalary()');
      } else if (employeeType === '管理部行政' && salaryType === '月薪' && workTimeType === '標準工時') {
        Logger.log('   ✅ 符合情況三：管理部行政');
        Logger.log('   → 會執行 calculateAdminSalary()');
      } else {
        Logger.log('   ❌ 不符合任何情況！');
        Logger.log('   → 會返回錯誤，不會計算加班費！');
        Logger.log('');
        Logger.log('💡 解決方案：');
        Logger.log('   方案 1：修改員工類型為以下之一：');
        Logger.log('          - 飼料廠司機 + 月薪 + 不定時');
        Logger.log('          - 食品廠移工 + 時薪 + 標準工時');
        Logger.log('          - 管理部行政 + 月薪 + 標準工時');
        Logger.log('');
        Logger.log('   方案 2：使用我提供的修正程式碼（加入預設邏輯）');
      }
      
      Logger.log('═══════════════════════════════════════');
      break;
    }
  }
}