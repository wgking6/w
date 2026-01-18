// ----------------------------------------------------------------------------
// 設定與全域變數 (CONFIGURATION & GLOBALS)
// ----------------------------------------------------------------------------
const APP_NAME = '遊戲方程式-自動發卡系統';
const API_VERSION = 'v1.8.0'; // Updated with Order-First Flow
const SPREADSHEET_ID = '1ywQDGsxE-lO5B3lxTJlozi0armhJb2m3cUIbjvwPuaM';

// 安全性設定
const ADMIN_PASSWORD = '8888'; // ★★★ 請在此修改您的管理密碼 ★★★

// 郵件對帳設定
// 指定轉寄來源 (您的手機轉發信箱)
const TRUSTED_FORWARDER = 'pei710514@gmail.com'; 
const BANK_EMAIL_SUBJECT = '入帳通知'; // 包含將來銀行或轉寄的標題

// 定義分頁名稱
const SHEET_USERS = '用戶資訊';
const SHEET_ORDERS = '訂單紀錄';
const SHEET_INVENTORY = '卡號資訊'; 
const SHEET_ISSUES = '問題回報';
const SHEET_PRODUCTS = '商品設定'; 

// 敏感分頁列表 (需要密碼保護的分頁)
const SENSITIVE_SHEETS = [SHEET_ORDERS, SHEET_INVENTORY];

// ----------------------------------------------------------------------------
// 核心：自動化資料庫連接 (CORE: AUTO DATABASE CONNECTION)
// ----------------------------------------------------------------------------

function getDB() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  ensureSheet(ss, SHEET_USERS, ['登入時間', 'User ID', '顯示名稱', '頭貼網址', '系統資訊']);
  ensureSheet(ss, SHEET_ORDERS, ['訂單編號', '下單時間', 'User ID', '用戶名稱', '商品名稱', '金額', '數量', '卡號', '密碼', '狀態', '付款備註', '手動發貨']);
  ensureSheet(ss, SHEET_INVENTORY, ['商品ID', '類型', '遊戲種類', '卡號', '密碼', '有效期', '狀態']);
  ensureSheet(ss, SHEET_ISSUES, ['回報時間', 'User ID', '用戶名稱', '問題類型', '詳細描述', '處理狀態']);
  ensureSheet(ss, SHEET_PRODUCTS, ['商品ID', '商品名稱', '描述', '價格', '圖片連結', '分類']);

  return ss;
}

function ensureSheet(ss, sheetName, headers) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
    
    // 初始化預設商品
    if (sheetName === SHEET_PRODUCTS) {
      sheet.appendRow(['pc-1', '新武林同萌傳輔助 - 30天月卡', '暢遊武林！30天尊榮會員', 150, 'https://i.ibb.co/WvYhdmc7/30.jpg', '電腦遊戲']);
      sheet.appendRow(['pc-2', '新武林同萌傳輔助 - 360天年卡', '年度超值方案！加贈坐騎', 1050, 'https://i.ibb.co/Bvh6JKn/360.jpg', '電腦遊戲']);
      sheet.appendRow(['pc-3', '艾爾之光輔助 - 月卡', '艾里奧斯大陸冒險必備，每日領取K-Ching', 800, 'https://i.ibb.co/xKKmw5ZQ/image.jpg', '電腦遊戲']);
      sheet.appendRow(['mob-chaos', '卡厄思夢境輔助 - 月卡', '夢境冒險，每日領取鑽石', 250, 'https://i.ibb.co/F447LW0Z/image.jpg', '手機遊戲']);
      sheet.appendRow(['mob-ro', 'RO仙境傳說輔助 - 月卡', '重返普隆德拉，月卡福利加倍', 250, 'https://i.ibb.co/sdddNjrb/RO.jpg', '手機遊戲']);
      sheet.appendRow(['mob-hot', '熱血江湖：福利加強版輔助 - 月卡', '熱血重燃，福利滿滿', 250, 'https://i.ibb.co/7tyRH2Hr/image.jpg', '手機遊戲']);
    }
    // 初始化庫存範例
    if (sheetName === SHEET_INVENTORY) {
      for(let i=0; i<5; i++) {
        sheet.appendRow(['pc-1', '月卡', '新武林同萌傳', `CODE-TEST-${i}`, `PASS-${i}`, '2025-12-31', 'Available']);
      }
    }
  }
  return sheet;
}

// ----------------------------------------------------------------------------
// 試算表選單與管理功能 (ADMIN MENU)
// ----------------------------------------------------------------------------

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🎮 遊戲方程式管理')
      .addItem('📥 立即執行 Gmail 對帳', 'checkGmailDeposits')
      .addItem('⚙️ 設定自動對帳頻率', 'setupGmailTrigger')
      .addSeparator()
      .addItem('🔓 解鎖查看敏感資料', 'unlockSensitiveSheets')
      .addItem('🔒 立即鎖定隱藏資料', 'lockSensitiveSheets')
      .addSeparator()
      .addItem('🧪 測試 Gmail 解析器 (Debug)', 'testEmailParser')
      .addItem('🧹 清理訂單頁空白列', 'cleanEmptyOrderRows')
      .addItem('🔄 強制檢查發貨', 'forceCheckPendingOrders')
      .addToUi();

  // 自動執行鎖定 (打開試算表時自動隱藏，保護資料)
  lockSensitiveSheets(true);
}

/**
 * 鎖定敏感資料 (隱藏分頁)
 * @param {boolean} silent 是否靜默執行 (不顯示 Toast)
 */
function lockSensitiveSheets(silent = false) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let count = 0;
  
  SENSITIVE_SHEETS.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (sheet && !sheet.isSheetHidden()) {
      sheet.hideSheet();
      count++;
    }
  });

  if (!silent && count > 0) {
    ss.toast(`已隱藏 ${count} 個敏感分頁。`, '安全鎖定');
  }
}

/**
 * 解鎖敏感資料 (輸入密碼後顯示)
 */
function unlockSensitiveSheets() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('身份驗證', '請輸入管理員密碼以查看敏感資料：', ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() == ui.Button.OK) {
    const input = response.getResponseText();
    
    if (input === ADMIN_PASSWORD) {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      let count = 0;
      
      SENSITIVE_SHEETS.forEach(sheetName => {
        const sheet = ss.getSheetByName(sheetName);
        if (sheet && sheet.isSheetHidden()) {
          sheet.showSheet();
          count++;
        }
      });
      
      // 自動切換到訂單頁，方便查看
      const orderSheet = ss.getSheetByName(SHEET_ORDERS);
      if (orderSheet) ss.setActiveSheet(orderSheet);

      ui.alert('驗證成功', `已解鎖顯示訂單與庫存分頁。\n\n⚠️ 注意：關閉視窗或重新整理後將自動重新鎖定。`, ui.ButtonSet.OK);
    } else {
      ui.alert('驗證失敗', '密碼錯誤，拒絕存取。', ui.ButtonSet.OK);
    }
  }
}


/**
 * 自動設定觸發器 (允許管理員自訂時間)
 */
function setupGmailTrigger() {
  const ui = SpreadsheetApp.getUi();
  const triggerName = 'checkGmailDeposits';
  const triggers = ScriptApp.getProjectTriggers();
  
  // 1. 檢查並詢問是否更新
  let existingTrigger = null;
  for (const t of triggers) {
    if (t.getHandlerFunction() === triggerName) {
      existingTrigger = t;
      break;
    }
  }

  let promptMsg = '請輸入自動檢查頻率 (分鐘)\n\n建議設定：\n- 5 分鐘 (推薦，省電穩定)\n- 1 分鐘 (最快，耗費配額)\n\n支援數值：1, 5, 10, 15, 30';
  if (existingTrigger) {
    promptMsg = '⚠️ 目前已啟用自動對帳。\n\n若要修改頻率，請重新輸入分鐘數 (1, 5, 10, 15, 30)：';
  }

  const response = ui.prompt('設定自動對帳', promptMsg, ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() == ui.Button.OK) {
    const input = response.getResponseText().trim();
    const minutes = parseInt(input, 10);
    const validIntervals = [1, 5, 10, 15, 30];

    if (!validIntervals.includes(minutes)) {
      ui.alert('輸入錯誤', 'Google 系統僅支援以下頻率 (分鐘)：\n1, 5, 10, 15, 30\n\n請重新操作。', ui.ButtonSet.OK);
      return;
    }

    try {
      // 刪除舊的觸發器 (避免重複)
      if (existingTrigger) {
        ScriptApp.deleteTrigger(existingTrigger);
      }

      // 建立新的觸發器
      ScriptApp.newTrigger(triggerName)
        .timeBased()
        .everyMinutes(minutes)
        .create();
      
      ui.alert('設定成功', `✅ 已啟用自動對帳！\n頻率：每 ${minutes} 分鐘檢查一次。\n\n系統將自動在背景運行。`, ui.ButtonSet.OK);
      
    } catch (e) {
      ui.alert('設定失敗', '無法建立觸發器，原因：' + e.toString(), ui.ButtonSet.OK);
    }
  }
}

/**
 * 測試 Email 解析功能
 */
function testEmailParser() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt('Email 解析測試', '請貼上銀行通知郵件的內容(純文字):', ui.ButtonSet.OK_CANCEL);
  
  if (result.getSelectedButton() == ui.Button.OK) {
    const text = result.getResponseText();
    const parsed = parseBankEmailContent(text, ''); // Pass empty sender for generic test
    
    let debugInfo = `解析結果：\n`;
    debugInfo += `類型 (Type): ${parsed.type}\n`;
    debugInfo += `金額 (Amount): ${parsed.amount}\n`;
    
    if (parsed.type === 'FORWARDER_TIME_MATCH') {
      debugInfo += `交易時間 (Time): ${parsed.paymentTime}\n`;
      debugInfo += `(使用時間+金額比對模式)`;
    } else {
      debugInfo += `帳號末碼 (Code): ${parsed.code}\n`;
      debugInfo += `(使用傳統末五碼比對模式)`;
    }
    
    ui.alert('解析結果', debugInfo, ui.ButtonSet.OK);
  }
}

function cleanEmptyOrderRows() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_ORDERS);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('確認清理', `即將檢查 ${lastRow} 列資料，刪除僅有核取方塊的空列，確定嗎？`, ui.ButtonSet.YES_NO);
  if (response !== ui.Button.YES) return;
  const data = sheet.getDataRange().getValues();
  let rowsToDelete = [];
  for (let i = data.length - 1; i >= 1; i--) {
    const row = data[i];
    if ((!row[0] || row[0] === '') && (!row[2] || row[2] === '')) {
      rowsToDelete.push(i + 1);
    }
  }
  if (rowsToDelete.length === 0) {
    ss.toast('沒有發現需要清理的空白列。', '完成');
    return;
  }
  rowsToDelete.forEach(rowIndex => sheet.deleteRow(rowIndex));
  ss.toast(`已清理 ${rowsToDelete.length} 列空白資料。`, '清理完成');
}

function forceCheckPendingOrders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_ORDERS);
  const data = sheet.getDataRange().getValues();
  let processedCount = 0;
  for (let i = 1; i < data.length; i++) {
    if (data[i][11] === true || data[i][11] === 'TRUE') {
       processManualFulfillment(sheet, i + 1);
       processedCount++;
    }
  }
  if (processedCount === 0) ss.toast('目前沒有打勾且未處理的訂單。', '系統提示');
}

// ----------------------------------------------------------------------------
// Gmail 自動對帳機器人 (AUTO RECONCILIATION)
// ----------------------------------------------------------------------------

function checkGmailDeposits() {
  console.log('開始執行 Gmail 對帳...');
  
  // 搜尋包含 "入帳" 或 "轉入" 的未讀郵件
  // 移除 from 限定，因為我們現在接受多種來源，但在 loop 中會判斷
  let query = `is:unread subject:("${BANK_EMAIL_SUBJECT}" OR "轉入通知")`;
  
  try {
    const threads = GmailApp.search(query, 0, 10);
    if (threads.length === 0) {
      console.log('沒有新的銀行通知郵件。');
      return;
    }

    const ss = getDB();
    const orderSheet = ss.getSheetByName(SHEET_ORDERS);
    const orderData = orderSheet.getDataRange().getValues();
    
    for (const thread of threads) {
      const messages = thread.getMessages();
      for (const message of messages) {
        if (!message.isUnread()) continue;
        
        const sender = message.getFrom();
        let body = message.getPlainBody();
        if (!body || body.length < 50) {
           body = message.getBody().replace(/<[^>]*>?/gm, ''); 
        }
        
        // 傳入 sender 以決定解析邏輯
        const parsed = parseBankEmailContent(body, sender);
        
        console.log(`郵件解析結果 (${parsed.type}): 金額=${parsed.amount}, 時間=${parsed.paymentTime}, 末碼=${parsed.code}`);

        if (parsed.amount > 0) {
           matchAndFulfill(orderSheet, orderData, parsed);
        } else {
           console.log('郵件解析失敗 (找不到金額)，跳過。');
        }
        
        message.markRead();
      }
    }
  } catch (e) {
    console.error('Gmail 對帳發生錯誤: ' + e.toString());
  }
}

/**
 * 智慧解析器
 * 根據寄件人不同，採用不同的解析策略
 */
function parseBankEmailContent(text, sender) {
  let result = {
    type: 'UNKNOWN',
    amount: 0,
    code: null,
    paymentTime: null
  };

  // 策略 A: 指定的轉發者 (pei710514@gmail.com) -> 使用「時間+金額」比對
  if (sender && sender.includes(TRUSTED_FORWARDER)) {
     // 格式: 你的帳戶在2026/01/18 21:18有 NT$1存入
     const forwarderRegex = /你的帳戶在\s*(\d{4}\/\d{2}\/\d{2}\s+\d{2}:\d{2})\s*有\s*(?:NT\$|\$)\s*([0-9,]+)\s*存入/i;
     const match = text.match(forwarderRegex);

     if (match) {
       result.type = 'FORWARDER_TIME_MATCH';
       result.paymentTime = new Date(match[1]); // 將字串轉為 Date 物件
       result.amount = parseInt(match[2].replace(/,/g, ''), 10);
       return result;
     }
  }

  // 策略 B: 一般銀行通知 -> 使用「帳號末五碼+金額」比對
  // 1. 解析金額
  const amountPatterns = [
    /(?:存入|交易|轉入)?金額(?:\(TWD\))?\s*[:：]?\s*(?:TWD|NT\$|\$)?\s*([0-9,]+)/i,
    /(?:TWD|NT\$|\$)\s*([0-9,]+)/i
  ];
  for (let pattern of amountPatterns) {
    const match = text.match(pattern);
    if (match) {
      result.amount = parseInt(match[1].replace(/,/g, ''), 10);
      break; 
    }
  }

  // 2. 解析帳號末五碼
  const codePatterns = [
    /(?:轉出|對方|來源|扣款)帳號\s*[:：]?\s*.*?([0-9]{5})(?![0-9])/i,
    /(?:末[五5]碼)\s*[:：]?\s*([0-9]{5})/i,
    /([0-9]{5})\s*(?:入帳|轉入)/,
    /(?:備註|摘要)\s*[:：]?\s*([0-9]{5})/
  ];
  for (let pattern of codePatterns) {
    const match = text.match(pattern);
    if (match) {
      result.code = match[1];
      result.type = 'STANDARD_CODE_MATCH';
      break;
    }
  }

  return result;
}

function matchAndFulfill(sheet, allData, parsedData) {
  let matchedOrderIndex = -1;
  let matchedOrderRow = null;
  let matchCount = 0;

  // 遍歷所有訂單尋找匹配者
  for (let i = 1; i < allData.length; i++) {
    const row = allData[i];
    const status = row[9]; // 狀態
    const orderPrice = Number(row[5]); // 金額
    
    if (status !== 'Pending') continue;

    let isMatch = false;

    // 邏輯分支：根據解析類型決定比對方式
    if (parsedData.type === 'FORWARDER_TIME_MATCH') {
       // --- 方案 B: 時間 + 金額 + 唯一性 ---
       if (orderPrice === parsedData.amount) {
         const orderTime = new Date(row[1]); // B欄: 下單時間
         const paymentTime = parsedData.paymentTime;
         
         // 計算時間差 (分鐘)
         const diffMinutes = (paymentTime - orderTime) / (1000 * 60);

         // 條件：入帳時間必須在下單時間之後，且在 30 分鐘內
         // 允許一點點誤差 (例如 -1 分鐘，防止伺服器時間些微不同步)，設定 >= -2
         if (diffMinutes >= -2 && diffMinutes <= 30) {
            isMatch = true;
            console.log(`[候選訂單] ID: ${row[0]}, 時間差: ${diffMinutes.toFixed(1)}分`);
         }
       }

    } else {
       // --- 方案 A: 傳統末五碼 + 金額 ---
       const paymentNote = String(row[10]).trim();
       const code = parsedData.code;
       // 寬鬆比對末五碼
       const codeMatch = (code && (paymentNote === String(code) || (String(code).endsWith(paymentNote) && paymentNote.length >= 4)));
       
       if (codeMatch && orderPrice === parsedData.amount) {
         isMatch = true;
       }
    }

    if (isMatch) {
      matchCount++;
      matchedOrderIndex = i + 1;
      matchedOrderRow = row;
    }
  }

  // 決策執行
  if (matchCount === 1) {
    // 只有唯一一筆符合 -> 發貨
    console.log(`找到唯一匹配訂單！Row: ${matchedOrderIndex}, User: ${matchedOrderRow[3]}`);
    const result = executeFulfillment(sheet, matchedOrderIndex, matchedOrderRow);
    if (result.success) {
      console.log(`訂單 ${matchedOrderRow[0]} 自動發貨成功。`);
    } else {
      console.error(`訂單 ${matchedOrderRow[0]} 發貨失敗: ${result.message}`);
    }
  } else if (matchCount > 1) {
    // 危險：有多筆符合條件 -> 不發貨，避免發錯人
    console.warn(`[安全警示] 發現 ${matchCount} 筆符合條件的訂單 (金額 $${parsedData.amount})，為避免誤發，系統已略過自動處理，請人工確認。`);
  } else {
    console.log(`找不到金額 $${parsedData.amount} 的匹配訂單 (模式: ${parsedData.type})。`);
  }
}

// ----------------------------------------------------------------------------
// API ROUTING
// ----------------------------------------------------------------------------

function doGet(e) {
  if (e.parameter && e.parameter.action) {
    return handleApiGet(e);
  }
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('遊戲方程式 Game Equation')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  try {
    return handleApiPost(e);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ 
      success: false, 
      message: 'Critical Server Error: ' + err.toString(),
      _version: API_VERSION
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ----------------------------------------------------------------------------
// API HANDLERS
// ----------------------------------------------------------------------------

function handleApiGet(e) {
  const action = e.parameter.action;
  let result = {};

  try {
    if (action === 'getProducts') {
      result = getProducts();
    } else if (action === 'getUserOrders') {
      result = getUserOrders(e.parameter.userId);
    } else {
      result = { error: 'Unknown action' };
    }
  } catch (err) {
    result = { error: err.toString() };
  }
  result._version = API_VERSION; 
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

function handleApiPost(e) {
  let data;
  try {
    if (e.postData && e.postData.contents) {
       data = JSON.parse(e.postData.contents);
    } else if (e.parameter) {
       data = e.parameter;
    }
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: 'Invalid JSON body: ' + err.toString() })).setMimeType(ContentService.MimeType.JSON);
  }

  if (!data || !data.action) {
     return ContentService.createTextOutput(JSON.stringify({ success: false, message: 'No action specified' })).setMimeType(ContentService.MimeType.JSON);
  }

  const action = data.action;
  let result = { success: false, message: 'Unknown action' };

  try {
    if (action === 'logUserAccess') {
      result = logUserAccess(data.data);
    } else if (action === 'processCartOrder') {
      result = processCartOrder(data.user, data.paymentNote, data.cartItems);
    } else if (action === 'updateOrderPayment') {
      // New Action for Update Payment
      result = updateOrderPayment(data.userId, data.orderId, data.paymentNote);
    } else if (action === 'submitIssue') {
      result = submitIssue(data.data);
    }
  } catch (err) {
    result = { success: false, message: 'Handler Error: ' + err.toString() };
  }
  result._version = API_VERSION;
  result._serverTime = new Date().toISOString();
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

// ----------------------------------------------------------------------------
// INTERNAL BUSINESS LOGIC
// ----------------------------------------------------------------------------

function logUserAccess(userProfile) {
  const ss = getDB();
  const sheet = ss.getSheetByName(SHEET_USERS);
  sheet.appendRow([new Date(), userProfile.userId, userProfile.displayName, userProfile.pictureUrl, userProfile.os || 'Unknown']);
  return { success: true };
}

function getProducts() {
  const ss = getDB();
  const sheet = ss.getSheetByName(SHEET_PRODUCTS);
  const data = sheet.getDataRange().getValues();
  const products = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[0]) {
      products.push({
        id: String(row[0]).trim(),
        name: String(row[1]).trim(),
        description: String(row[2]),
        price: Number(row[3]),
        imageUrl: String(row[4]).trim(),
        category: String(row[5] || '').trim() 
      });
    }
  }
  return products;
}

function getNextRealEmptyRow(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 2;
  const range = sheet.getRange(1, 1, lastRow, 1).getValues();
  for (let i = range.length - 1; i >= 0; i--) {
    if (range[i][0] && String(range[i][0]).trim() !== "") {
      return i + 2; 
    }
  }
  return 2; 
}

function processCartOrder(userObj, paymentNote, cartItems) {
  const lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(10000)) return { success: false, message: '系統忙碌' };
    
    const ss = getDB();
    const orderSheet = ss.getSheetByName(SHEET_ORDERS);
    const orderId = 'ORD-' + Date.now();
    let resultItems = []; 
    let nextRow = getNextRealEmptyRow(orderSheet);
    
    // 如果沒有 paymentNote，設為空字串 (允許先建單)
    const note = paymentNote ? String(paymentNote) : '';

    for (let item of cartItems) {
      const rowData = [
        orderId,
        new Date(),
        userObj.userId,
        userObj.displayName,
        item.name,
        Number(item.price) * Number(item.quantity),
        item.quantity,
        '', // Code 
        '', // Password
        'Pending', 
        note,
        false // Checkbox placeholder
      ];
      orderSheet.getRange(nextRow, 1, 1, rowData.length).setValues([rowData]);
      orderSheet.getRange(nextRow, 12).insertCheckboxes();
      resultItems.push({ name: item.name, quantity: item.quantity });
      nextRow++; 
    }
    
    SpreadsheetApp.flush();
    return { success: true, message: '訂單已提交', orderId: orderId, items: resultItems };

  } catch (err) {
    return { success: false, message: 'Process Error: ' + err.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 新增功能：更新訂單的付款資訊 (後五碼)
 */
function updateOrderPayment(userId, orderId, paymentNote) {
  const lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(10000)) return { success: false, message: '系統忙碌' };
    
    const ss = getDB();
    const sheet = ss.getSheetByName(SHEET_ORDERS);
    const data = sheet.getDataRange().getValues();
    let updatedCount = 0;

    // 搜尋符合 OrderID 與 UserID 的訂單 (確保安全性)
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(orderId) && String(data[i][2]) === String(userId)) {
        // 更新第 11 欄 (Index 10) 為付款備註
        sheet.getRange(i + 1, 11).setValue(String(paymentNote));
        updatedCount++;
      }
    }

    if (updatedCount > 0) {
      return { success: true, message: '付款資訊已更新' };
    } else {
      return { success: false, message: '找不到訂單或權限不足' };
    }

  } catch (err) {
    return { success: false, message: 'Update Error: ' + err.toString() };
  } finally {
    lock.releaseLock();
  }
}


function getUserOrders(userId) {
  const ss = getDB();
  const sheet = ss.getSheetByName(SHEET_ORDERS);
  const data = sheet.getDataRange().getValues();
  const myOrders = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i] && String(data[i][2]) === String(userId)) {
      let dateStr = "";
      try { dateStr = data[i][1] instanceof Date ? data[i][1].toISOString() : String(data[i][1]); } catch(e) { dateStr = String(data[i][1]); }
      myOrders.push({
        orderId: data[i][0],
        date: dateStr,
        productName: data[i][4],
        price: data[i][5],
        quantity: data[i][6],
        codes: data[i][7],
        passwords: data[i][8]
      });
    }
  }
  return myOrders.reverse();
}

function submitIssue(issueData) {
  try {
    const ss = getDB();
    const sheet = ss.getSheetByName(SHEET_ISSUES);
    sheet.appendRow([new Date(), issueData.userId, issueData.displayName, issueData.type, issueData.description, '待處理']);
    return { success: true, message: '回報已收到' };
  } catch(e) {
    return { success: false, message: '回報失敗' };
  }
}

// ----------------------------------------------------------------------------
// TRIGGER & FULFILLMENT LOGIC
// ----------------------------------------------------------------------------

function onEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();
  if (sheet.getName() !== SHEET_ORDERS) return;
  // 如果是 L 欄 (12) 被勾選
  if (range.getColumn() === 12 && (e.value === 'TRUE' || e.value === true)) {
    const row = range.getRow();
    if (row === 1) return; 
    
    // UI 版本發貨：會顯示 Toast 提示
    const result = executeFulfillment(sheet, row, null);
    if (result.success) {
      SpreadsheetApp.getActive().toast(result.message, '成功');
    } else {
      // 失敗則取消勾選
      sheet.getRange(row, 12).uncheck();
      SpreadsheetApp.getActive().toast(result.message, '發貨失敗');
    }
  }
}

function processManualFulfillment(orderSheet, rowIndex) {
  // Legacy Wrapper for older menu calls
  const result = executeFulfillment(orderSheet, rowIndex, null);
  if (result.success) {
     SpreadsheetApp.getActive().toast(result.message, '成功');
  } else {
     orderSheet.getRange(rowIndex, 12).uncheck();
     SpreadsheetApp.getActive().toast(result.message, '失敗');
  }
}

/**
 * 通用發貨邏輯 (不依賴 UI 互動)
 * 供 onEdit (手動) 和 checkGmailDeposits (自動) 共用
 */
function executeFulfillment(orderSheet, rowIndex, providedRowData) {
  const rowData = providedRowData || orderSheet.getRange(rowIndex, 1, 1, 12).getValues()[0];
  const orderId = rowData[0];
  const productName = rowData[4];
  const qtyNeeded = rowData[6] || 1;
  const currentCode = rowData[7];

  // 1. 檢查是否已發過貨
  if (currentCode && currentCode.toString().trim() !== '') {
    // 雖然已有卡號，但為了讓 checkbox 狀態正確，還是回傳 success
    orderSheet.getRange(rowIndex, 12).uncheck(); 
    return { success: true, message: `訂單 ${orderId} 已有卡號，無需補發` };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const invSheet = ss.getSheetByName(SHEET_INVENTORY);
  const prodSheet = ss.getSheetByName(SHEET_PRODUCTS);

  // 2. 找商品 ID
  const prodData = prodSheet.getDataRange().getValues();
  let productId = null;
  for(let p=1; p<prodData.length; p++){
    if(prodData[p][1] === productName) {
      productId = prodData[p][0];
      break;
    }
  }

  if (!productId) {
    return { success: false, message: `找不到商品 [${productName}] 的 ID` };
  }

  // 3. 找庫存
  const invData = invSheet.getDataRange().getValues();
  let foundIndices = [];
  let codes = [];
  let passwords = [];

  for (let i = 1; i < invData.length; i++) {
    if (String(invData[i][0]) === String(productId) && String(invData[i][6]).toLowerCase() === 'available') {
      foundIndices.push(i + 1);
      codes.push(invData[i][3]);
      passwords.push(invData[i][4]);
      if (foundIndices.length === qtyNeeded) break;
    }
  }

  // 4. 庫存不足處理
  if (foundIndices.length < qtyNeeded) {
    return { success: false, message: `庫存不足！商品 [${productName}] 需要 ${qtyNeeded}，僅剩 ${foundIndices.length}` };
  }

  // 5. 更新庫存狀態為 Sold
  foundIndices.forEach(idx => {
    invSheet.getRange(idx, 7).setValue('Sold');
  });

  // 6. 寫入訂單
  const finalCodes = codes.join('\n');
  const finalPass = passwords.join('\n');
  
  orderSheet.getRange(rowIndex, 8).setValue(finalCodes);
  orderSheet.getRange(rowIndex, 9).setValue(finalPass);
  orderSheet.getRange(rowIndex, 10).setValue('Completed'); 
  
  // 保持 Checkbox unchecked (因為已經處理完了，不需要打勾留在哪)
  orderSheet.getRange(rowIndex, 12).uncheck(); 
  
  SpreadsheetApp.flush();
  return { success: true, message: `訂單 ${orderId} 發貨成功` };
}
