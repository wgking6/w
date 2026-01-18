// ----------------------------------------------------------------------------
// 設定與全域變數 (CONFIGURATION & GLOBALS)
// ----------------------------------------------------------------------------
const APP_NAME = '遊戲方程式-自動發卡系統';
const API_VERSION = 'v2.1.0'; // Updated: Member Profile & New Categories
const SPREADSHEET_ID = '1ywQDGsxE-lO5B3lxTJlozi0armhJb2m3cUIbjvwPuaM';

// 安全性設定
const ADMIN_PASSWORD = '8888'; // 試算表選單用的密碼
const ADMIN_LINE_ID = 'Ua66fd77f72e4524075afd856cae91587'; // ★★★ 超級管理員 LINE ID ★★★

// 郵件對帳設定
// 指定轉寄來源 (您的手機轉發信箱)
const TRUSTED_FORWARDER = 'pei710514@gmail.com'; 
const BANK_EMAIL_SUBJECT = '入帳通知'; // 包含將來銀行或轉寄的標題

// 定義分頁名稱
const SHEET_USERS = '用戶資訊';
const SHEET_MEMBERS = '會員資料'; // ★ New Sheet
const SHEET_ORDERS = '訂單紀錄';
const SHEET_INVENTORY = '卡號資訊'; 
const SHEET_ISSUES = '問題回報';
const SHEET_PRODUCTS = '商品設定'; 

// 敏感分頁列表 (需要密碼保護的分頁)
const SENSITIVE_SHEETS = [SHEET_ORDERS, SHEET_INVENTORY, SHEET_MEMBERS];

// ----------------------------------------------------------------------------
// 核心：自動化資料庫連接 (CORE: AUTO DATABASE CONNECTION)
// ----------------------------------------------------------------------------

function getDB() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  ensureSheet(ss, SHEET_USERS, ['登入時間', 'User ID', '顯示名稱', '頭貼網址', '系統資訊']);
  ensureSheet(ss, SHEET_MEMBERS, ['User ID', '顯示名稱', '電話', '信箱', '性別', '最後更新']); // ★ New Sheet
  // Header Config:
  // 0:訂單編號, 1:下單時間, 2:User ID, 3:用戶名稱, 4:商品名稱, 5:金額, 6:數量, 
  // 7:卡號, 8:密碼, 9:發貨時間, 10:狀態, 11:付款備註, 12:手動發貨
  ensureSheet(ss, SHEET_ORDERS, ['訂單編號', '下單時間', 'User ID', '用戶名稱', '商品名稱', '金額', '數量', '卡號', '密碼', '發貨時間', '狀態', '付款備註', '手動發貨']);
  ensureSheet(ss, SHEET_INVENTORY, ['商品ID', '類型', '遊戲種類', '卡號', '密碼', '有效期', '狀態']);
  // Updated Issues Header to include '相關商品'
  ensureSheet(ss, SHEET_ISSUES, ['回報時間', 'User ID', '用戶名稱', '相關商品', '問題類型', '詳細描述', '處理狀態']);
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
      sheet.appendRow(['web-1', '網頁遊戲通用助手', '釋放雙手，自動掛機', 100, 'https://placehold.co/600x400/1e293b/fbbf24?text=Web+Game', '網頁遊戲']); // Sample Web Game
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
      .addItem('⚙️ 啟用自動對帳 (5分鐘/次)', 'setupGmailTrigger')
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
  if (!silent && count > 0) ss.toast(`已隱藏 ${count} 個敏感分頁。`, '安全鎖定');
}

function unlockSensitiveSheets() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('身份驗證', '請輸入管理員密碼以查看敏感資料：', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() == ui.Button.OK) {
    if (response.getResponseText() === ADMIN_PASSWORD) {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      SENSITIVE_SHEETS.forEach(sheetName => {
        const sheet = ss.getSheetByName(sheetName);
        if (sheet && sheet.isSheetHidden()) sheet.showSheet();
      });
      const orderSheet = ss.getSheetByName(SHEET_ORDERS);
      if (orderSheet) ss.setActiveSheet(orderSheet);
      ui.alert('驗證成功', `已解鎖顯示訂單與庫存分頁。\n\n⚠️ 注意：關閉視窗或重新整理後將自動重新鎖定。`, ui.ButtonSet.OK);
    } else {
      ui.alert('驗證失敗', '密碼錯誤，拒絕存取。', ui.ButtonSet.OK);
    }
  }
}

function setupGmailTrigger() {
  const ui = SpreadsheetApp.getUi();
  const triggerName = 'checkGmailDeposits';
  const triggers = ScriptApp.getProjectTriggers();
  for (const t of triggers) {
    if (t.getHandlerFunction() === triggerName) {
      const response = ui.alert('自動對帳已啟用', '是否重設為「每 5 分鐘」檢查一次？', ui.ButtonSet.YES_NO);
      if (response == ui.Button.YES) ScriptApp.deleteTrigger(t);
      else return;
      break;
    }
  }
  try {
    ScriptApp.newTrigger(triggerName).timeBased().everyMinutes(5).create();
    ui.alert('設定成功', `✅ 已啟用自動對帳！\n頻率：每 5 分鐘檢查一次。`, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('設定失敗', '無法建立觸發器：' + e.toString(), ui.ButtonSet.OK);
  }
}

function testEmailParser() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt('Email 解析測試', '請貼上內容:', ui.ButtonSet.OK_CANCEL);
  if (result.getSelectedButton() == ui.Button.OK) {
    const text = result.getResponseText();
    const parsed = parseBankEmailContent(text, '');
    ui.alert('解析結果', `類型: ${parsed.type}\n金額: ${parsed.amount}\n時間: ${parsed.paymentTime}\n末碼: ${parsed.code}`, ui.ButtonSet.OK);
  }
}

function cleanEmptyOrderRows() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_ORDERS);
  const data = sheet.getDataRange().getValues();
  let rowsToDelete = [];
  for (let i = data.length - 1; i >= 1; i--) {
    const row = data[i];
    if ((!row[0] || row[0] === '') && (!row[2] || row[2] === '')) rowsToDelete.push(i + 1);
  }
  if (rowsToDelete.length === 0) return ss.toast('無空白列。', '完成');
  rowsToDelete.forEach(rowIndex => sheet.deleteRow(rowIndex));
  ss.toast(`已清理 ${rowsToDelete.length} 列。`, '完成');
}

function forceCheckPendingOrders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_ORDERS);
  const data = sheet.getDataRange().getValues();
  let processedCount = 0;
  for (let i = 1; i < data.length; i++) {
    // Checkbox is at index 12 (Col M)
    if (data[i][12] === true || data[i][12] === 'TRUE') {
       processManualFulfillment(sheet, i + 1);
       processedCount++;
    }
  }
  if (processedCount === 0) ss.toast('無待處理訂單。', '系統提示');
}

// ----------------------------------------------------------------------------
// Gmail 自動對帳
// ----------------------------------------------------------------------------

function checkGmailDeposits() {
  console.log('執行對帳...');
  let query = `is:unread subject:("${BANK_EMAIL_SUBJECT}" OR "轉入通知")`;
  try {
    const threads = GmailApp.search(query, 0, 10);
    if (threads.length === 0) return;
    const ss = getDB();
    const orderSheet = ss.getSheetByName(SHEET_ORDERS);
    const orderData = orderSheet.getDataRange().getValues();
    for (const thread of threads) {
      const messages = thread.getMessages();
      for (const message of messages) {
        if (!message.isUnread()) continue;
        let body = message.getPlainBody();
        if (!body || body.length < 50) body = message.getBody().replace(/<[^>]*>?/gm, '');
        const parsed = parseBankEmailContent(body, message.getFrom());
        console.log(`解析: $${parsed.amount}, 末碼:${parsed.code}, 時間:${parsed.paymentTime}`);
        if (parsed.amount > 0) matchAndFulfill(orderSheet, orderData, parsed);
        message.markRead();
      }
    }
  } catch (e) { console.error('對帳錯誤: ' + e.toString()); }
}

function parseBankEmailContent(text, sender) {
  let result = { type: 'UNKNOWN', amount: 0, code: null, paymentTime: null };
  if (sender && sender.includes(TRUSTED_FORWARDER)) {
     const match = text.match(/你的帳戶在\s*(\d{4}\/\d{2}\/\d{2}\s+\d{2}:\d{2})\s*有\s*(?:NT\$|\$)\s*([0-9,]+)\s*存入/i);
     if (match) {
       result.type = 'FORWARDER_TIME_MATCH';
       result.paymentTime = new Date(match[1]);
       result.amount = parseInt(match[2].replace(/,/g, ''), 10);
       return result;
     }
  }
  const amtMatch = text.match(/(?:金額|NT\$|\$)\s*[:：]?\s*(?:TWD)?\s*([0-9,]+)/i);
  if (amtMatch) result.amount = parseInt(amtMatch[1].replace(/,/g, ''), 10);
  
  const codeMatch = text.match(/(?:末[五5]碼|帳號|備註).*?([0-9]{5})/i);
  if (codeMatch) {
    result.code = codeMatch[1];
    result.type = 'STANDARD_CODE_MATCH';
  }
  return result;
}

function matchAndFulfill(sheet, allData, parsedData) {
  let matchedIdx = -1;
  let matchedRow = null;
  let count = 0;

  for (let i = 1; i < allData.length; i++) {
    const row = allData[i];
    // Indices: 5:Price, 10:Status, 11:Note
    if (row[10] !== '待處理') continue; 
    
    let isMatch = false;
    if (parsedData.type === 'FORWARDER_TIME_MATCH') {
       if (Number(row[5]) === parsedData.amount) {
         const diff = (parsedData.paymentTime - new Date(row[1])) / 60000;
         if (diff >= -2 && diff <= 30) isMatch = true;
       }
    } else {
       const note = String(row[11]).trim();
       if ((note === String(parsedData.code) || (parsedData.code && note.endsWith(parsedData.code))) && Number(row[5]) === parsedData.amount) {
         isMatch = true;
       }
    }

    if (isMatch) {
      count++;
      matchedIdx = i + 1;
      matchedRow = row;
    }
  }

  if (count === 1) executeFulfillment(sheet, matchedIdx, matchedRow);
  else if (count > 1) console.warn(`多筆匹配，略過。金額: ${parsedData.amount}`);
}

// ----------------------------------------------------------------------------
// API ROUTING
// ----------------------------------------------------------------------------

function doGet(e) {
  if (e.parameter && e.parameter.action) return handleApiGet(e);
  return HtmlService.createTemplateFromFile('index').evaluate()
    .setTitle('遊戲方程式 Game Equation')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  try { return handleApiPost(e); } 
  catch (err) { return ContentService.createTextOutput(JSON.stringify({ success: false, message: err.toString() })).setMimeType(ContentService.MimeType.JSON); }
}

function handleApiGet(e) {
  const action = e.parameter.action;
  let result = {};
  try {
    if (action === 'getProducts') result = getProducts();
    else if (action === 'getUserOrders') result = getUserOrders(e.parameter.userId);
    else if (action === 'getMemberProfile') result = getMemberProfile(e.parameter.userId); // ★ New
    else result = { error: 'Unknown action' };
  } catch (err) { result = { error: err.toString() }; }
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

function handleApiPost(e) {
  const data = JSON.parse(e.postData.contents);
  const action = data.action;
  let result = { success: false, message: 'Unknown action' };
  try {
    if (action === 'logUserAccess') result = logUserAccess(data.data);
    else if (action === 'processCartOrder') result = processCartOrder(data.user, data.paymentNote, data.cartItems);
    else if (action === 'updateOrderPayment') result = updateOrderPayment(data.userId, data.orderId, data.paymentNote);
    else if (action === 'cancelOrder') result = cancelOrder(data.userId, data.orderId);
    else if (action === 'submitIssue') result = submitIssue(data.data);
    else if (action === 'updateMemberProfile') result = updateMemberProfile(data); // ★ New
    else if (action === 'adminAction') result = handleAdminAction(data);
  } catch (err) { result = { success: false, message: err.toString() }; }
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

// ----------------------------------------------------------------------------
// ADMIN & BUSINESS LOGIC
// ----------------------------------------------------------------------------

function handleAdminAction(data) {
  if (data.adminId !== ADMIN_LINE_ID) return { success: false, message: 'Unauthorized' };
  const ss = getDB();
  if (data.subAction === 'getDashboardData') return getAdminDashboardData(ss);
  if (data.subAction === 'addInventory') return adminAddInventory(ss, data.payload);
  if (data.subAction === 'deleteInventory') return adminDeleteInventory(ss, data.payload);
  if (data.subAction === 'manualFulfill') return adminManualFulfill(ss, data.payload);
  return { success: false };
}

function getAdminDashboardData(ss) {
  const invSheet = ss.getSheetByName(SHEET_INVENTORY);
  const prodSheet = ss.getSheetByName(SHEET_PRODUCTS);
  const orderSheet = ss.getSheetByName(SHEET_ORDERS);
  const userSheet = ss.getSheetByName(SHEET_USERS);

  const prodData = prodSheet.getDataRange().getValues();
  const productMap = {};
  const productList = [];
  for(let i=1; i<prodData.length; i++) {
    productMap[prodData[i][0]] = prodData[i][1];
    productList.push({ id: prodData[i][0], name: prodData[i][1] });
  }

  const invData = invSheet.getDataRange().getValues();
  const inventory = [];
  for(let i=invData.length-1; i>=1; i--) {
     if(inventory.length > 200) break;
     inventory.push({ rowIndex: i + 1, productId: invData[i][0], productName: productMap[invData[i][0]] || invData[i][0], code: invData[i][3], password: invData[i][4], status: invData[i][6] });
  }

  const orderData = orderSheet.getDataRange().getValues();
  const pendingOrders = [];
  for(let i=1; i<orderData.length; i++) {
    const row = orderData[i];
    if (row[10] === '待處理') { // Status Index 10
       pendingOrders.push({ rowIndex: i + 1, orderId: row[0], date: new Date(row[1]).toLocaleString(), userName: row[3], productName: row[4], price: row[5], paymentNote: row[11] });
    }
  }

  const userData = userSheet.getDataRange().getValues();
  const users = [];
  for(let i=userData.length-1; i>=1; i--) {
     if(users.length > 20) break;
     users.push({ date: new Date(userData[i][0]).toLocaleString(), name: userData[i][2], uid: userData[i][1] });
  }

  return { success: true, products: productList, inventory: inventory, orders: pendingOrders, users: users };
}

function adminAddInventory(ss, payload) {
  const sheet = ss.getSheetByName(SHEET_INVENTORY);
  payload.items.forEach(item => sheet.appendRow([payload.productId, 'AdminAdd', 'Manual', item.code, item.pass || '', '2099-12-31', 'Available']));
  return { success: true };
}

function adminDeleteInventory(ss, payload) {
  ss.getSheetByName(SHEET_INVENTORY).deleteRow(payload.rowIndex);
  return { success: true };
}

function adminManualFulfill(ss, payload) {
  return executeFulfillment(ss.getSheetByName(SHEET_ORDERS), payload.rowIndex, null);
}

// ----------------------------------------------------------------------------
// USER LOGIC
// ----------------------------------------------------------------------------

function logUserAccess(user) {
  // 1. Log to Users Sheet
  getDB().getSheetByName(SHEET_USERS).appendRow([new Date(), user.userId, user.displayName, user.pictureUrl, user.os || 'Unknown']);
  
  // 2. Upsert to Members Sheet (Ensure member exists)
  const ss = getDB();
  const memSheet = ss.getSheetByName(SHEET_MEMBERS);
  const data = memSheet.getDataRange().getValues();
  let found = false;
  // Use a quick check on latest rows or iterate. Since user base might be small, iterating is okay for now.
  // Optimization: In a real DB, this is a unique key upsert.
  for(let i=1; i<data.length; i++) {
    if (String(data[i][0]) === String(user.userId)) {
      found = true;
      break;
    }
  }
  if (!found) {
    memSheet.appendRow([user.userId, user.displayName, '', '', '', new Date()]);
  }
  
  return { success: true };
}

function getMemberProfile(userId) {
  const ss = getDB();
  const sheet = ss.getSheetByName(SHEET_MEMBERS);
  const data = sheet.getDataRange().getValues();
  for(let i=1; i<data.length; i++) {
    if(String(data[i][0]) === String(userId)) {
      // 0:ID, 1:Name, 2:Phone, 3:Email, 4:Gender, 5:LastUpdate
      return { success: true, phone: data[i][2], email: data[i][3], gender: data[i][4] };
    }
  }
  return { success: true, phone: '', email: '', gender: '' };
}

function updateMemberProfile(data) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return { success: false, message: '系統忙碌' };
  try {
    const ss = getDB();
    const sheet = ss.getSheetByName(SHEET_MEMBERS);
    const rows = sheet.getDataRange().getValues();
    let found = false;
    for(let i=1; i<rows.length; i++) {
      if(String(rows[i][0]) === String(data.userId)) {
        sheet.getRange(i+1, 3).setValue(data.phone);
        sheet.getRange(i+1, 4).setValue(data.email);
        sheet.getRange(i+1, 5).setValue(data.gender);
        sheet.getRange(i+1, 6).setValue(new Date());
        found = true;
        break;
      }
    }
    // If not found (should be rare if logged in), append
    if(!found) {
      sheet.appendRow([data.userId, data.displayName, data.phone, data.email, data.gender, new Date()]);
    }
    return { success: true };
  } catch(e) {
    return { success: false, message: e.toString() };
  } finally { lock.releaseLock(); }
}

function getProducts() {
  const data = getDB().getSheetByName(SHEET_PRODUCTS).getDataRange().getValues();
  const products = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) products.push({ id: String(data[i][0]).trim(), name: String(data[i][1]).trim(), description: String(data[i][2]), price: Number(data[i][3]), imageUrl: String(data[i][4]).trim(), category: String(data[i][5] || '').trim() });
  }
  return products;
}

function getNextRealEmptyRow(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 2;
  const range = sheet.getRange(1, 1, lastRow, 1).getValues();
  for (let i = range.length - 1; i >= 0; i--) {
    if (range[i][0] && String(range[i][0]).trim() !== "") return i + 2; 
  }
  return 2; 
}

function processCartOrder(userObj, paymentNote, cartItems) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return { success: false, message: '系統忙碌' };
  try {
    const ss = getDB();
    const sheet = ss.getSheetByName(SHEET_ORDERS);
    const orderId = 'ORD-' + Date.now();
    let resultItems = []; 
    let nextRow = getNextRealEmptyRow(sheet);
    const note = paymentNote ? String(paymentNote) : '';

    for (let item of cartItems) {
      // Columns: 0:ID, 1:Time, 2:UID, 3:Name, 4:Prod, 5:Price, 6:Qty, 7:Code, 8:Pass, 9:DelivTime, 10:Status, 11:Note, 12:Check
      const rowData = [orderId, new Date(), userObj.userId, userObj.displayName, item.name, Number(item.price) * Number(item.quantity), item.quantity, '', '', '', '待處理', note, false];
      sheet.getRange(nextRow, 1, 1, rowData.length).setValues([rowData]);
      sheet.getRange(nextRow, 13).insertCheckboxes();
      resultItems.push({ name: item.name, quantity: item.quantity });
      nextRow++; 
    }
    SpreadsheetApp.flush();
    return { success: true, message: '訂單已提交', orderId: orderId, items: resultItems };
  } finally { lock.releaseLock(); }
}

function updateOrderPayment(userId, orderId, paymentNote) {
  const ss = getDB();
  const sheet = ss.getSheetByName(SHEET_ORDERS);
  const data = sheet.getDataRange().getValues();
  let updatedCount = 0;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(orderId) && String(data[i][2]) === String(userId)) {
      sheet.getRange(i + 1, 12).setValue(String(paymentNote)); // Index 11 is Note
      updatedCount++;
    }
  }
  return updatedCount > 0 ? { success: true } : { success: false, message: '找不到訂單' };
}

function cancelOrder(userId, orderId) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return { success: false, message: '系統忙碌' };
  try {
    const ss = getDB();
    const sheet = ss.getSheetByName(SHEET_ORDERS);
    const data = sheet.getDataRange().getValues();
    let updatedCount = 0;
    
    for (let i = 1; i < data.length; i++) {
      // 驗證 UserID, OrderID 且狀態必須是 '待處理' (Pending)
      if (String(data[i][0]) === String(orderId) && String(data[i][2]) === String(userId)) {
        if (data[i][10] === '待處理' || data[i][10] === 'Pending') {
           sheet.getRange(i + 1, 11).setValue('已取消'); // Index 10 is Status
           updatedCount++;
        } else {
           return { success: false, message: '訂單狀態不可取消 (已發貨或已取消)' };
        }
      }
    }
    
    return updatedCount > 0 ? { success: true, message: '訂單已取消' } : { success: false, message: '找不到符合條件的訂單' };
  } finally { lock.releaseLock(); }
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
        passwords: data[i][8],
        status: data[i][10],      // ★ Add Status
        paymentNote: data[i][11]  // ★ Add Note
      });
    }
  }
  return myOrders.reverse();
}

// Updated to include productName
function submitIssue(issueData) {
  // New Header: ['回報時間', 'User ID', '用戶名稱', '相關商品', '問題類型', '詳細描述', '處理狀態']
  getDB().getSheetByName(SHEET_ISSUES).appendRow([new Date(), issueData.userId, issueData.displayName, issueData.productName || '無', issueData.type, issueData.description, '待處理']);
  return { success: true };
}

function onEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();
  if (sheet.getName() !== SHEET_ORDERS) return;
  // Index 12 (Col M) is Checkbox
  if (range.getColumn() === 13 && (e.value === 'TRUE' || e.value === true)) {
    const row = range.getRow();
    if (row === 1) return; 
    const result = executeFulfillment(sheet, row, null);
    if (result.success) SpreadsheetApp.getActive().toast(result.message, '成功');
    else {
      sheet.getRange(row, 13).uncheck();
      SpreadsheetApp.getActive().toast(result.message, '發貨失敗');
    }
  }
}

function executeFulfillment(orderSheet, rowIndex, providedRowData) {
  const rowData = providedRowData || orderSheet.getRange(rowIndex, 1, 1, 13).getValues()[0];
  const orderId = rowData[0];
  const productName = rowData[4];
  const qtyNeeded = rowData[6] || 1;
  const currentCode = rowData[7];

  if (currentCode && currentCode.toString().trim() !== '') {
    orderSheet.getRange(rowIndex, 13).uncheck(); 
    return { success: true, message: `訂單 ${orderId} 已有卡號` };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const invSheet = ss.getSheetByName(SHEET_INVENTORY);
  const prodSheet = ss.getSheetByName(SHEET_PRODUCTS);

  const prodData = prodSheet.getDataRange().getValues();
  let productId = null;
  for(let p=1; p<prodData.length; p++){
    if(prodData[p][1] === productName) { productId = prodData[p][0]; break; }
  }

  if (!productId) return { success: false, message: `無此商品ID [${productName}]` };

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

  if (foundIndices.length < qtyNeeded) return { success: false, message: `庫存不足 ${qtyNeeded}` };

  foundIndices.forEach(idx => invSheet.getRange(idx, 7).setValue('Sold'));

  orderSheet.getRange(rowIndex, 8).setValue(codes.join('\n')); // Col H (Index 7+1)
  orderSheet.getRange(rowIndex, 9).setValue(passwords.join('\n')); // Col I
  orderSheet.getRange(rowIndex, 10).setValue(new Date()); // Col J (DelivTime)
  orderSheet.getRange(rowIndex, 11).setValue('已發貨'); // Col K (Status)
  orderSheet.getRange(rowIndex, 13).uncheck(); // Col M
  
  SpreadsheetApp.flush();
  return { success: true, message: `發貨成功 ${orderId}` };
}
