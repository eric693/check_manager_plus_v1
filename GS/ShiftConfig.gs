// ShiftConfig.gs - 班別設定管理
// ================================================================
// 【使用說明】
// 1. 將此檔案新增到 Google Apps Script 專案
// 2. 執行一次 initShiftConfigSheet() 建立班別工作表
// 3. 執行一次 migrateEmployeeShiftColumn() 為員工資料表加入班別欄位
// ================================================================

const SHEET_SHIFTS = '班別設定';

// ==================== 預設班別清單 ====================
// 可在 Google Sheet「班別設定」工作表中直接新增/修改
const DEFAULT_SHIFTS = [
  // [班別ID,    班別名稱,   上班,    下班,    午休開始, 午休結束, 每日工時]
  ['STANDARD', '一般',     '09:00', '18:00', '12:00', '13:00', 8],
  ['ADMIN',    '行政',     '08:00', '17:00', '12:00', '13:00', 8],
  ['STORE',    '門市',     '09:30', '18:30', '12:00', '13:00', 8],
  ['DRIVER_A', '司機早班', '07:00', '16:00', '12:00', '13:00', 8],
  ['DRIVER_B', '司機中班', '10:00', '19:00', '14:00', '15:00', 8],
  ['DRIVER_C', '司機午班', '12:00', '21:00', '17:00', '18:00', 8],
  ['DRIVER_D', '司機晚班', '14:00', '23:00', '18:00', '19:00', 8],
];

// ==================== 初始化 ====================

/**
 * 建立班別設定工作表（執行一次即可）
 */
function initShiftConfigSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_SHIFTS);

  if (sheet) {
    Logger.log('⚠️ 班別設定工作表已存在，跳過建立');
    return sheet;
  }

  sheet = ss.insertSheet(SHEET_SHIFTS);

  const headers = ['班別ID', '班別名稱', '上班時間', '下班時間', '午休開始', '午休結束', '每日工時'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#4A90E2')
    .setFontColor('#FFFFFF')
    .setHorizontalAlignment('center');
  sheet.setFrozenRows(1);

  // 寫入預設班別
  DEFAULT_SHIFTS.forEach(row => sheet.appendRow(row));

  // 自動調整欄寬
  sheet.autoResizeColumns(1, headers.length);

  Logger.log('✅ 班別設定工作表已建立，共 ' + DEFAULT_SHIFTS.length + ' 個班別');
  return sheet;
}

/**
 * 為員工資料表新增班別欄位（I 欄）（執行一次即可）
 */
function migrateEmployeeShiftColumn() {
  Logger.log('═══════════════════════════════════════');
  Logger.log('🔄 開始新增員工班別欄位');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_EMPLOYEES);

  if (!sheet) {
    Logger.log('❌ 找不到員工工作表');
    return;
  }

  // 確認 I 欄標題是否已存在
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  if (headers[8] === '班別') {
    Logger.log('✅ 班別欄位已存在，無需遷移');
    return;
  }

  // 在 I 欄（第 9 欄）加入「班別」標題
  sheet.getRange(1, 9).setValue('班別').setFontWeight('bold');

  // 為所有現有員工填入預設班別 STANDARD
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    for (let i = 2; i <= lastRow; i++) {
      const currentVal = sheet.getRange(i, 9).getValue();
      if (!currentVal) {
        sheet.getRange(i, 9).setValue('STANDARD');
      }
    }
  }

  Logger.log(`✅ 已為 ${lastRow - 1} 位員工設定預設班別 STANDARD`);
  Logger.log('═══════════════════════════════════════');
  Logger.log('⚠️ 請至員工資料工作表手動調整各員工的班別');
  Logger.log('   可用班別: STANDARD / ADMIN / STORE / DRIVER_A / DRIVER_B / DRIVER_C / DRIVER_D');
}

// ==================== 班別查詢 ====================

/**
 * 將時間字串「HH:mm」轉換為小數時間（例如 08:30 → 8.5）
 */
function parseTimeToHour_(timeStr) {
  if (!timeStr) return 9;
  const str = String(timeStr).trim();
  const parts = str.split(':');
  if (parts.length < 2) return parseInt(str) || 9;
  return parseInt(parts[0]) + parseInt(parts[1]) / 60;
}

/**
 * 將工作表一列資料解析為班別設定物件
 */
function parseShiftRow_(row) {
  return {
    shiftId:        String(row[0] || 'STANDARD'),
    shiftName:      String(row[1] || '一般'),
    startTime:      String(row[2] || '09:00'),   // "09:00" 字串，給前端顯示
    endTime:        String(row[3] || '18:00'),
    lunchStart:     String(row[4] || '12:00'),
    lunchEnd:       String(row[5] || '13:00'),
    dailyHours:     parseFloat(row[6]) || 8,
    // 數字版本，給後端計算用
    startHour:      parseTimeToHour_(row[2]),
    endHour:        parseTimeToHour_(row[3]),
    lunchStartHour: parseTimeToHour_(row[4]),
    lunchEndHour:   parseTimeToHour_(row[5]),
  };
}

/**
 * 取得預設班別設定（找不到時的 fallback）
 */
function getDefaultShiftConfig_() {
  return {
    shiftId:        'STANDARD',
    shiftName:      '一般',
    startTime:      '09:00',
    endTime:        '18:00',
    lunchStart:     '12:00',
    lunchEnd:       '13:00',
    dailyHours:     8,
    startHour:      9,
    endHour:        18,
    lunchStartHour: 12,
    lunchEndHour:   13,
  };
}

/**
 * 根據班別ID取得班別設定
 * @param {string} shiftId
 * @returns {Object} 班別設定物件
 */
function getShiftConfig(shiftId) {
  if (!shiftId) return getDefaultShiftConfig_();

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SHIFTS);

  if (!sheet) {
    Logger.log('⚠️ 找不到班別設定工作表，使用預設');
    return getDefaultShiftConfig_();
  }

  const values = sheet.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]).trim() === String(shiftId).trim()) {
      return parseShiftRow_(values[i]);
    }
  }

  Logger.log(`⚠️ 找不到班別 "${shiftId}"，使用預設`);
  return getDefaultShiftConfig_();
}

/**
 * 根據員工 userId 取得其班別設定
 * @param {string} userId
 * @returns {Object} 班別設定物件
 */
function getUserShiftConfig(userId) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_EMPLOYEES);
  if (!sheet) return getDefaultShiftConfig_();

  const values = sheet.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]).trim() === userId) {
      const shiftId = values[i][8] || 'STANDARD'; // I 欄（index 8）
      return getShiftConfig(shiftId);
    }
  }

  return getDefaultShiftConfig_();
}

/**
 * 取得所有班別清單（管理員後台用、前端下拉選單用）
 */
function getAllShifts(sessionToken) {
  // sessionToken 可選，不傳也可以回傳（公開資料）
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SHIFTS);

  if (!sheet) {
    initShiftConfigSheet();
    return getAllShifts(sessionToken);
  }

  const values = sheet.getDataRange().getValues();
  const shifts = [];

  for (let i = 1; i < values.length; i++) {
    if (values[i][0]) {
      shifts.push(parseShiftRow_(values[i]));
    }
  }

  Logger.log('📋 取得班別清單: ' + shifts.length + ' 個');
  return { ok: true, shifts: shifts };
}

// ==================== 班別設定（管理員）====================

/**
 * 設定員工班別（管理員專用）
 * @param {string} sessionToken
 * @param {string} targetUserId - 要設定的員工 userId
 * @param {string} shiftId - 班別ID
 */
function setUserShift(sessionToken, targetUserId, shiftId) {
  const employee = checkSession_(sessionToken);
  if (!employee.ok || !employee.user) {
    return { ok: false, code: 'ERR_SESSION_INVALID' };
  }
  if (employee.user.dept !== '管理員') {
    return { ok: false, code: 'ERR_NO_PERMISSION' };
  }

  // 確認班別存在
  const shift = getShiftConfig(shiftId);
  if (shift.shiftId !== shiftId) {
    return { ok: false, msg: `找不到班別 "${shiftId}"` };
  }

  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_EMPLOYEES);
  const values = sheet.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === targetUserId) {
      sheet.getRange(i + 1, 9).setValue(shiftId);
      Logger.log(`✅ 已設定員工 ${values[i][2]}（${targetUserId}）的班別為 ${shiftId}`);
      return {
        ok: true,
        msg: `班別已更新為「${shift.shiftName}」（${shift.startTime} ~ ${shift.endTime}）`
      };
    }
  }

  return { ok: false, msg: '找不到該員工' };
}

/**
 * 取得員工目前的班別（前端顯示用）
 */
function getUserShift(sessionToken, targetUserId) {
  const employee = checkSession_(sessionToken);
  if (!employee.ok || !employee.user) {
    return { ok: false, code: 'ERR_SESSION_INVALID' };
  }

  // 員工只能查自己，管理員可以查任何人
  const queryUserId = (employee.user.dept === '管理員' && targetUserId)
    ? targetUserId
    : employee.user.userId;

  const shiftConfig = getUserShiftConfig(queryUserId);
  return { ok: true, shift: shiftConfig };
}

// ==================== 測試函數 ====================

function testShiftConfig() {
  Logger.log('🧪 測試班別設定');
  Logger.log('');

  // 測試各班別
  const shiftIds = ['STANDARD', 'ADMIN', 'STORE', 'DRIVER_A', 'DRIVER_B', 'NOTEXIST'];

  shiftIds.forEach(id => {
    const config = getShiftConfig(id);
    Logger.log(`📋 ${id}:`);
    Logger.log(`   班別名稱: ${config.shiftName}`);
    Logger.log(`   時間: ${config.startTime} ~ ${config.endTime}`);
    Logger.log(`   午休: ${config.lunchStart} ~ ${config.lunchEnd}`);
    Logger.log(`   startHour: ${config.startHour}, endHour: ${config.endHour}`);
    Logger.log('');
  });
}

function testGetAllShifts() {
  const result = getAllShifts();
  Logger.log('📋 所有班別:');
  Logger.log(JSON.stringify(result, null, 2));
}