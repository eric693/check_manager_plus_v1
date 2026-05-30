// LineBotPunch.gs - LINE Bot 打卡處理

/**
 * 處理 LINE 文字訊息
 */
function handleLineMessage(event) {
  const userId = event.source.userId;
  const messageText = event.message.text.trim();
  
  Logger.log(`📥 收到訊息: ${messageText} from ${userId}`);
  
  // 判斷指令
  if (messageText === '打卡' || messageText === '上班' || messageText === '下班') {
    // 要求傳送位置
    replyRequestLocation(event.replyToken, messageText);
  } 
  else if (messageText === '查詢打卡記錄') {
    replyPunchRecords(event.replyToken, userId);
  }
  else if (messageText === '補打卡') {
    replyAdjustPunchGuide(event.replyToken);
  }
  else {
    // 預設回覆
    replyHelpMessage(event.replyToken);
  }
}

/**
 * 處理 LINE 位置訊息（實際打卡）
 */
function handleLineLocation(event) {
  const userId = event.source.userId;
  const lat = event.message.latitude;
  const lng = event.message.longitude;
  const address = event.message.address || '';
  
  Logger.log(`📍 收到位置: ${lat}, ${lng} from ${userId}`);
  
  // 取得員工資料
  const employee = findEmployeeByLineUserId_(userId);
  
  if (!employee || !employee.ok) {
    replyMessage(event.replyToken, '❌ 找不到您的員工資料，請先完成註冊！');
    return;
  }
  
  // 判斷打卡類型（根據時間自動判斷）
  const now = new Date();
  const hour = now.getHours();
  let punchType;
  
  if (hour >= 6 && hour < 14) {
    punchType = '上班';
  } else {
    punchType = '下班';
  }
  
  // 驗證位置是否在允許範圍內
  const locationCheck = checkPunchLocation_(lat, lng);
  
  if (!locationCheck.ok) {
    replyMessage(
      event.replyToken, 
      `❌ ${locationCheck.msg}\n\n您目前的位置距離最近的打卡地點約 ${Math.round(locationCheck.distance)} 公尺。\n\n如需補打卡，請在網頁系統申請。`
    );
    return;
  }
  
  // 執行打卡
  const result = linePunch_(userId, punchType, lat, lng, locationCheck.locationName);
  
  if (result.ok) {
    replyPunchSuccess(event.replyToken, employee.name, punchType, locationCheck.locationName, now);
  } else {
    replyMessage(event.replyToken, `❌ 打卡失敗: ${result.msg}`);
  }
}

/**
 * 驗證打卡位置
 */
function checkPunchLocation_(lat, lng) {
  const shLoc = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LOCATIONS);
  const lastRow = shLoc.getLastRow();
  
  if (lastRow < 2) {
    return { ok: false, msg: "系統尚未設定打卡地點" };
  }
  
  const values = shLoc.getRange(2, 1, lastRow - 1, 5).getValues();
  let locationName = null;
  let minDistance = Infinity;
  
  for (let [, name, locLat, locLng, radius] of values) {
    if (!name || !locLat || !locLng) continue;
    
    const dist = getDistanceMeters_(lat, lng, Number(locLat), Number(locLng));
    
    if (dist <= Number(radius) && dist < minDistance) {
      locationName = name;
      minDistance = dist;
    }
  }
  
  if (!locationName) {
    return { 
      ok: false, 
      msg: "您不在任何打卡地點範圍內",
      distance: minDistance
    };
  }
  
  return { 
    ok: true, 
    locationName: locationName,
    distance: minDistance
  };
}

/**
 * 執行打卡（LINE Bot 專用）
 */
function linePunch_(userId, type, lat, lng, locationName) {
  try {
    const employee = findEmployeeByLineUserId_(userId);
    
    if (!employee || !employee.ok) {
      return { ok: false, msg: "找不到員工資料" };
    }
    
    const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_ATTENDANCE);
    const row = [
      new Date(),
      employee.userId,
      employee.dept,
      employee.name,
      type,
      `(${lat},${lng})`,
      locationName,
      "LINE打卡",  // 備註欄標示為 LINE 打卡
      "",
      "LINE Bot"
    ];
    
    sh.getRange(sh.getLastRow() + 1, 1, 1, row.length).setValues([row]);
    
    Logger.log(`✅ LINE 打卡成功: ${employee.name} - ${type}`);
    
    return { ok: true };
    
  } catch (error) {
    Logger.log('❌ LINE 打卡錯誤: ' + error);
    return { ok: false, msg: error.message };
  }
}


/**
 * 要求傳送位置
 */
function replyRequestLocation(replyToken, punchType) {
  const message = {
    type: 'text',
    text: `請分享您的位置以完成${punchType}打卡 📍`,
    quickReply: {
      items: [
        {
          type: 'action',
          action: {
            type: 'location',
            label: '📍 傳送位置'
          }
        }
      ]
    }
  };
  
  sendLineReply_(replyToken, [message]);
}

/**
 * 回覆打卡成功
 */
function replyPunchSuccess(replyToken, employeeName, punchType, location, time) {
  const timeStr = Utilities.formatDate(time, 'Asia/Taipei', 'HH:mm:ss');
  
  const message = {
    type: 'flex',
    altText: `✅ ${punchType}打卡成功`,
    contents: {
      type: 'bubble',
      header: {
        type: 'box',
        layout: 'vertical',
        contents: [
          {
            type: 'text',
            text: `✅ ${punchType}打卡成功`,
            weight: 'bold',
            size: 'xl',
            color: '#FFFFFF'
          }
        ],
        backgroundColor: '#4CAF50'
      },
      body: {
        type: 'box',
        layout: 'vertical',
        contents: [
          {
            type: 'text',
            text: `${employeeName}，您已成功打卡！`,
            margin: 'md'
          },
          {
            type: 'separator',
            margin: 'lg'
          },
          {
            type: 'box',
            layout: 'vertical',
            margin: 'lg',
            spacing: 'sm',
            contents: [
              {
                type: 'box',
                layout: 'baseline',
                contents: [
                  { type: 'text', text: '時間', color: '#999999', flex: 2 },
                  { type: 'text', text: timeStr, flex: 5, weight: 'bold' }
                ]
              },
              {
                type: 'box',
                layout: 'baseline',
                contents: [
                  { type: 'text', text: '地點', color: '#999999', flex: 2 },
                  { type: 'text', text: location, flex: 5, wrap: true }
                ]
              }
            ]
          }
        ]
      }
    }
  };
  
  sendLineReply_(replyToken, [message]);
}

/**
 * 回覆說明訊息
 */
function replyHelpMessage(replyToken) {
  const message = {
    type: 'text',
    text: `📋 打卡系統使用說明\n\n請輸入以下指令：\n• 「打卡」或「上班」或「下班」\n• 「查詢打卡記錄」\n• 「補打卡」\n\n輸入後請分享您的位置以完成打卡。`
  };
  
  sendLineReply_(replyToken, [message]);
}

/**
 * 發送 LINE 回覆訊息
 */
function sendLineReply_(replyToken, messages) {
  const url = 'https://api.line.me/v2/bot/message/reply';
  const accessToken = PropertiesService.getScriptProperties().getProperty('LINE_CHANNEL_ACCESS_TOKEN');
  
  const payload = {
    replyToken: replyToken,
    messages: messages
  };
  
  const options = {
    method: 'post',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${accessToken}`
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    Logger.log('✅ LINE 回覆已發送');
  } catch (error) {
    Logger.log('❌ LINE 回覆失敗: ' + error);
  }
}


/**
 * 🧪 測試 LINE Messaging API 連線
 */
function testLineMessagingAPI() {
  Logger.log('🧪 測試 LINE Messaging API');
  Logger.log('');
  
  const accessToken = PropertiesService.getScriptProperties().getProperty('LINE_CHANNEL_ACCESS_TOKEN');
  
  if (!accessToken) {
    Logger.log('❌ 找不到 LINE_CHANNEL_ACCESS_TOKEN');
    Logger.log('');
    Logger.log('📋 請設定：');
    Logger.log('   1. 前往 LINE Developers Console');
    Logger.log('   2. 選擇你的 Channel → Messaging API');
    Logger.log('   3. 找到「Channel access token (long-lived)」');
    Logger.log('   4. 點擊「Issue」（如果還沒建立）');
    Logger.log('   5. 複製 Token');
    Logger.log('   6. 在 Apps Script → 專案設定 → 指令碼屬性');
    Logger.log('   7. 新增：LINE_CHANNEL_ACCESS_TOKEN = (貼上 Token)');
    return;
  }
  
  Logger.log('✅ Access Token 已設定');
  Logger.log('   長度: ' + accessToken.length + ' 字元');
  Logger.log('');
  
  // 測試發送訊息到指定 User ID
  const testUserId = 'U68e0ca9d516e63ed15bf9387fad174ac'; // ⚠️ 替換成你的
  
  try {
    Logger.log('📤 嘗試發送測試訊息...');
    
    const url = 'https://api.line.me/v2/bot/message/push';
    const payload = {
      to: testUserId,
      messages: [
        {
          type: 'text',
          text: '🧪 LINE Bot 測試成功！\n\n如果你收到這則訊息，代表 LINE Bot 已經可以正常運作了。\n\n現在可以試試輸入「打卡」來測試打卡功能！'
        }
      ]
    };
    
    const options = {
      method: 'post',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${accessToken}`
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(url, options);
    const result = JSON.parse(response.getContentText());
    
    Logger.log('');
    Logger.log('📤 API 回應:');
    Logger.log('   狀態碼: ' + response.getResponseCode());
    
    if (response.getResponseCode() === 200) {
      Logger.log('   ✅ 訊息發送成功！');
      Logger.log('');
      Logger.log('📱 請檢查你的 LINE 是否收到測試訊息');
    } else {
      Logger.log('   ❌ 發送失敗');
      Logger.log('   錯誤: ' + JSON.stringify(result));
    }
    
  } catch (error) {
    Logger.log('');
    Logger.log('❌ 測試失敗: ' + error);
  }
  
  Logger.log('');
  Logger.log('═══════════════════════════════════════');
}

/**
 * 🔍 檢查最近的 Webhook 記錄
 */
function checkRecentWebhookLogs() {
  Logger.log('═══════════════════════════════════════');
  Logger.log('🔍 檢查最近的 LINE Bot 活動');
  Logger.log('═══════════════════════════════════════');
  Logger.log('');
  
  Logger.log('📋 請檢查以下項目：');
  Logger.log('');
  Logger.log('1. Apps Script → 執行作業');
  Logger.log('   查看是否有最新的執行記錄');
  Logger.log('');
  Logger.log('2. 檢查是否有錯誤訊息');
  Logger.log('');
  Logger.log('3. 確認 doPost 函數有被觸發');
  Logger.log('');
  
  // 檢查必要設定
  const accessToken = PropertiesService.getScriptProperties().getProperty('LINE_CHANNEL_ACCESS_TOKEN');
  Logger.log('✅ Access Token: ' + (accessToken ? '已設定' : '❌ 未設定'));
  
  // 檢查工作表
  try {
    const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_EMPLOYEES);
    Logger.log('✅ 員工工作表: 存在');
  } catch (e) {
    Logger.log('❌ 員工工作表: 錯誤 - ' + e);
  }
  
  try {
    const sheetLoc = SpreadsheetApp.getActive().getSheetByName(SHEET_LOCATIONS);
    const lastRow = sheetLoc.getLastRow();
    Logger.log('✅ 打卡地點: ' + (lastRow - 1) + ' 個');
    
    if (lastRow > 1) {
      Logger.log('');
      Logger.log('📍 已設定的打卡地點:');
      const data = sheetLoc.getRange(2, 1, lastRow - 1, 5).getValues();
      data.forEach((row, idx) => {
        Logger.log(`   ${idx + 1}. ${row[1]}`);
        Logger.log(`      座標: (${row[2]}, ${row[3]})`);
        Logger.log(`      範圍: ${row[4]} 公尺`);
      });
    }
  } catch (e) {
    Logger.log('❌ 打卡地點: 錯誤 - ' + e);
  }
  
  Logger.log('');
  Logger.log('═══════════════════════════════════════');
}


/**
 * 🧪 主動推送測試訊息（不需要你輸入任何東西）
 */
function pushTestMessageToMe() {
  Logger.log('📤 準備推送測試訊息...');
  
  const accessToken = PropertiesService.getScriptProperties().getProperty('LINE_CHANNEL_ACCESS_TOKEN');
  const myUserId = 'U68e0ca9d516e63ed15bf9387fad174ac'; // 你的 User ID
  
  if (!accessToken) {
    Logger.log('❌ 找不到 Access Token');
    return;
  }
  
  const url = 'https://api.line.me/v2/bot/message/push';
  
  const payload = {
    to: myUserId,
    messages: [
      {
        type: 'text',
        text: '🧪 推送測試\n\n如果你收到這則訊息，代表 LINE Messaging API 運作正常。\n\n請試著回覆「打卡」來測試 Webhook。',
        quickReply: {
          items: [
            {
              type: 'action',
              action: {
                type: 'message',
                label: '🎯 打卡',
                text: '打卡'
              }
            },
            {
              type: 'action',
              action: {
                type: 'message',
                label: '📊 查詢',
                text: '查詢'
              }
            }
          ]
        }
      }
    ]
  };
  
  const options = {
    method: 'post',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${accessToken}`
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const result = JSON.parse(response.getContentText());
    
    Logger.log('');
    Logger.log('📤 推送結果:');
    Logger.log('   狀態碼: ' + response.getResponseCode());
    
    if (response.getResponseCode() === 200) {
      Logger.log('   ✅ 訊息已推送！請檢查你的 LINE');
    } else {
      Logger.log('   ❌ 推送失敗');
      Logger.log('   回應: ' + JSON.stringify(result));
    }
    
  } catch (error) {
    Logger.log('❌ 推送錯誤: ' + error);
  }
}