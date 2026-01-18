// ----------------------------------------------------------------------------
// 設定與全域變數 (CONFIGURATION & GLOBALS)
// ----------------------------------------------------------------------------
const APP_NAME = '遊戲方程式-自動發卡系統';
// Updated Spreadsheet ID
const SPREADSHEET_ID = '1ywQDGsxE-lO5B3lxTJlozi0armhJb2m3cUIbjvwPuaM';

// 定義分頁名稱
const SHEET_USERS = '用戶資訊';
const SHEET_ORDERS = '訂單紀錄';
const SHEET_INVENTORY = '卡號資訊'; 
const SHEET_ISSUES = '問題回報';
const SHEET_PRODUCTS = '商品設定'; 

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
    
    if (sheetName === SHEET_PRODUCTS) {
      sheet.appendRow(['pc-1', '新武林同萌傳輔助 - 30天月卡', '暢遊武林！30天尊榮會員', 150, 'https://i.ibb.co/WvYhdmc7/30.jpg', '電腦遊戲']);
      sheet.appendRow(['pc-2', '新武林同萌傳輔助 - 360天年卡', '年度超值方案！加贈坐騎', 1050, 'https://i.ibb.co/Bvh6JKn/360.jpg', '電腦遊戲']);
      sheet.appendRow(['pc-3', '艾爾之光輔助 - 月卡', '艾里奧斯大陸冒險必備，每日領取K-Ching', 800, 'https://i.ibb.co/xKKmw5ZQ/image.jpg', '電腦遊戲']);
      sheet.appendRow(['mob-chaos', '卡厄思夢境輔助 - 月卡', '夢境冒險，每日領取鑽石', 250, 'https://i.ibb.co/F447LW0Z/image.jpg', '手機遊戲']);
      sheet.appendRow(['mob-ro', 'RO仙境傳說輔助 - 月卡', '重返普隆德拉，月卡福利加倍', 250, 'https://i.ibb.co/sdddNjrb/RO.jpg', '手機遊戲']);
      sheet.appendRow(['mob-hot', '熱血江湖：福利加強版輔助 - 月卡', '熱血重燃，福利滿滿', 250, 'https://i.ibb.co/7tyRH2Hr/image.jpg', '手機遊戲']);
    }
    if (sheetName === SHEET_INVENTORY) {
      for(let i=0; i<5; i++) {
        sheet.appendRow(['pc-1', '月卡', '新武林同萌傳', `CODE-TEST-${i}`, `PASS-${i}`, '2025-12-31', 'Available']);
      }
    }
    // 移除：不再預先插入 L2:L1000 的核取方塊，避免佔用列數導致 appendRow 寫入到千行之後
  }
  return sheet;
}

// ----------------------------------------------------------------------------
// API ROUTING (WEB APP ENTRY POINTS)
// ----------------------------------------------------------------------------

function doGet(e) {
  // Check if it's an API request
  if (e.parameter && e.parameter.action) {
    return handleApiGet(e);
  }

  // Otherwise serve the HTML template (Legacy/GAS Mode)
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('遊戲方程式 Game Equation')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  return handleApiPost(e);
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

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function handleApiPost(e) {
  let data;
  try {
    // 優先解析 contents，如果是透過 text/plain 送來的
    if (e.postData && e.postData.contents) {
       data = JSON.parse(e.postData.contents);
    } else if (e.parameter) {
       data = e.parameter;
    }
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: 'Invalid JSON: ' + err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (!data || !data.action) {
     return ContentService.createTextOutput(JSON.stringify({ success: false, message: 'No action specified' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const action = data.action;
  let result = { success: false, message: 'Unknown action' };

  try {
    if (action === 'logUserAccess') {
      result = logUserAccess(data.data);
    } else if (action === 'processCartOrder') {
      result = processCartOrder(data.user, data.paymentNote, data.cartItems);
    } else if (action === 'submitIssue') {
      result = submitIssue(data.data);
    }
  } catch (err) {
    result = { success: false, message: err.toString() };
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
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

/**
 * 輔助函數：尋找下一個真正的空白列（忽略僅有核取方塊的列）
 */
function getNextRealEmptyRow(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 2;
  
  // 讀取 A 欄 (訂單編號) 來判斷真正的最後一筆資料在哪
  // 這樣即使 L 欄有核取方塊，我們也能找到正確的寫入點
  const range = sheet.getRange(1, 1, lastRow, 1).getValues();
  
  for (let i = range.length - 1; i >= 0; i--) {
    if (range[i][0] && String(range[i][0]).trim() !== "") {
      return i + 2; // 資料在 index i (row i+1)，所以下一列是 i+2
    }
  }
  return 2; // 如果整欄都空，從第 2 列開始
}

/**
 * 處理購物車訂單
 */
function processCartOrder(userObj, paymentNote, cartItems) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); 
  } catch (e) {
    return { success: false, message: '系統忙碌中，請稍後重試' };
  }

  try {
    const ss = getDB();
    const orderSheet = ss.getSheetByName(SHEET_ORDERS);
    const orderId = 'ORD-' + Date.now();
    let resultItems = []; 

    // 使用智慧偵測找到正確的寫入列
    let nextRow = getNextRealEmptyRow(orderSheet);

    for (let item of cartItems) {
      const rowData = [
        orderId,
        new Date(),
        userObj.userId,
        userObj.displayName,
        item.name,
        Number(item.price) * Number(item.quantity),
        item.quantity,
        '', // Code 留空
        '', // Password 留空
        'Pending', 
        String(paymentNote || ''),
        false // 手動發貨 Checkbox
      ];

      // 使用 setValues 指定寫入特定列，而非 appendRow
      orderSheet.getRange(nextRow, 1, 1, rowData.length).setValues([rowData]);
      
      // 確保該列有核取方塊
      orderSheet.getRange(nextRow, 12).insertCheckboxes();

      resultItems.push({
        name: item.name,
        quantity: item.quantity,
        codes: '', 
        passwords: ''
      });
      
      nextRow++; // 下一筆商品往下寫
    }

    return {
      success: true,
      message: '訂單已提交',
      orderId: orderId,
      items: resultItems
    };

  } catch (err) {
    return { success: false, message: '伺服器錯誤: ' + err.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 取得用戶的所有訂單
 */
function getUserOrders(userId) {
  const ss = getDB();
  const sheet = ss.getSheetByName(SHEET_ORDERS);
  const data = sheet.getDataRange().getValues();
  
  const myOrders = [];
  // 從第2列開始 (跳過標題)
  for (let i = 1; i < data.length; i++) {
    // 檢查欄位是否存在，避免讀取到空白列報錯
    if (data[i] && String(data[i][2]) === String(userId)) {
      let dateStr = "";
      try {
         dateStr = data[i][1] instanceof Date ? data[i][1].toISOString() : String(data[i][1]);
      } catch(e) {
         dateStr = String(data[i][1]);
      }

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
// SPREADSHEET TRIGGER (手動發貨功能)
// ----------------------------------------------------------------------------

function onEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();
  
  if (sheet.getName() !== SHEET_ORDERS) return;
  
  if (range.getColumn() === 12 && e.value === 'TRUE') {
    const row = range.getRow();
    if (row === 1) return; 

    processManualFulfillment(sheet, row);
  }
}

function processManualFulfillment(orderSheet, rowIndex) {
  const rowData = orderSheet.getRange(rowIndex, 1, 1, 12).getValues()[0];
  const productName = rowData[4];
  const qtyNeeded = rowData[6] || 1;
  const currentCode = rowData[7];

  if (currentCode && currentCode.toString().trim() !== '') {
    orderSheet.getRange(rowIndex, 12).uncheck();
    SpreadsheetApp.getActive().toast(`訂單 ${rowData[0]} 已有卡號，無需補發`, '系統提示');
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const invSheet = ss.getSheetByName(SHEET_INVENTORY);
  const prodSheet = ss.getSheetByName(SHEET_PRODUCTS);

  const prodData = prodSheet.getDataRange().getValues();
  let productId = null;
  for(let p=1; p<prodData.length; p++){
    if(prodData[p][1] === productName) {
      productId = prodData[p][0];
      break;
    }
  }

  if (!productId) {
    SpreadsheetApp.getActive().toast('找不到對應商品ID', '錯誤');
    orderSheet.getRange(rowIndex, 12).uncheck();
    return;
  }

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

  if (foundIndices.length < qtyNeeded) {
    SpreadsheetApp.getActive().toast(`庫存不足！需要 ${qtyNeeded}，僅剩 ${foundIndices.length}`, '發貨失敗');
    orderSheet.getRange(rowIndex, 12).uncheck();
    return;
  }

  foundIndices.forEach(idx => {
    invSheet.getRange(idx, 7).setValue('Sold');
  });

  const finalCodes = codes.join('\n');
  const finalPass = passwords.join('\n');
  
  orderSheet.getRange(rowIndex, 8).setValue(finalCodes);
  orderSheet.getRange(rowIndex, 9).setValue(finalPass);
  orderSheet.getRange(rowIndex, 10).setValue('Manual Filled'); 
  
  orderSheet.getRange(rowIndex, 12).uncheck();
  
  SpreadsheetApp.getActive().toast(`訂單 ${rowData[0]} 手動發貨成功！`, '成功');
}