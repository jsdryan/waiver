// ============================================================
// Google Apps Script - 部署為 Web App
//
// 使用步驟：
// 1. 開啟 Google Sheet，點選「擴充功能」>「Apps Script」
// 2. 將此檔案內容貼上到 Apps Script 編輯器
// 3. 修改下方 SHEET_ID、FOLDER_ID、TURNSTILE_SECRET
// 4. 點選「部署」>「管理部署作業」>「編輯」> 版本選「新版本」> 部署
// 5. 複製部署後的 URL，貼到 index.html 的 SCRIPT_URL
// ============================================================

const SHEET_ID = '1wgQGr9TSMcCPbDkkSUHOR_atrW6iUs0Wv5oqMtCybNU';
const FOLDER_ID = '165bN7W2mN2FRSuyBfHg0MbhYev8N6clH';
const TURNSTILE_SECRET = '0x4AAAAAABjosJBBxTZHaYQXHK9E_kqSVxo';

// --- Turnstile 人機驗證 ---
function verifyTurnstile(token) {
  const response = UrlFetchApp.fetch('https://challenges.cloudflare.com/turnstile/v0/siteverify', {
    method: 'post',
    payload: {
      secret: TURNSTILE_SECRET,
      response: token
    }
  });
  const result = JSON.parse(response.getContentText());
  return result.success === true;
}

// --- 後端資料驗證 ---
function validateData(data) {
  if (!data.studentId || !/^[A-Z][12]\d{8}$/.test(data.studentId)) {
    return '身分證字號格式不正確';
  }
  if (!data.studentName || !/^[\u4e00-\u9fff]{2,10}$/.test(data.studentName)) {
    return '學生姓名格式不正確';
  }
  if (!data.waiverFileData || !data.waiverFileName) {
    return '缺少切結書檔案';
  }
  // 檢查 base64 大小（約為原檔 1.37 倍），限制 10MB
  if (data.waiverFileData.length > 14 * 1024 * 1024) {
    return '切結書檔案過大';
  }
  if (data.siblingFileData && data.siblingFileData.length > 14 * 1024 * 1024) {
    return '隨兄妹證明檔案過大';
  }
  return null;
}

// --- 簡易速率限制（同一身分證 5 分鐘內不可重複提交）---
function checkRateLimit(sheet, studentId) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return null;

  const data = sheet.getRange(lastRow, 1, 1, 2).getValues()[0];
  const lastTime = new Date(data[0]);
  const lastId = data[1];

  if (lastId === studentId && (new Date() - lastTime) < 5 * 60 * 1000) {
    return '同一身分證字號 5 分鐘內請勿重複提交';
  }
  return null;
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    // 1. Turnstile 人機驗證
    if (!data.turnstileToken || !verifyTurnstile(data.turnstileToken)) {
      return ContentService
        .createTextOutput(JSON.stringify({ status: 'error', message: '人機驗證失敗' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 2. 後端資料驗證
    const validationError = validateData(data);
    if (validationError) {
      return ContentService
        .createTextOutput(JSON.stringify({ status: 'error', message: validationError }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const sheet = SpreadsheetApp.openById(SHEET_ID).getActiveSheet();
    const folder = DriveApp.getFolderById(FOLDER_ID);

    // 3. 速率限制
    const rateLimitError = checkRateLimit(sheet, data.studentId);
    if (rateLimitError) {
      return ContentService
        .createTextOutput(JSON.stringify({ status: 'error', message: rateLimitError }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 確認表頭（首次自動建立）
    if (sheet.getLastRow() === 0) {
      const headers = [
        '提交時間',
        '學生身分證字號',
        '學生姓名',
        '切結書檔名',
        '切結書連結',
        '隨兄妹證明檔名',
        '隨兄妹證明連結',
        '處理狀態'
      ];
      sheet.appendRow(headers);

      // 設定表頭樣式
      const headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#4F46E5');
      headerRange.setFontColor('#FFFFFF');
      headerRange.setHorizontalAlignment('center');
      sheet.setFrozenRows(1);

      // 設定欄寬
      sheet.setColumnWidth(1, 160);  // 提交時間
      sheet.setColumnWidth(2, 130);  // 身分證字號
      sheet.setColumnWidth(3, 80);   // 姓名
      sheet.setColumnWidth(4, 160);  // 切結書檔名
      sheet.setColumnWidth(5, 100);  // 切結書連結
      sheet.setColumnWidth(6, 160);  // 隨兄妹證明檔名
      sheet.setColumnWidth(7, 100);  // 隨兄妹證明連結
      sheet.setColumnWidth(8, 100);  // 處理狀態
    }

    // 檔案命名：身分證字號_姓名_原始檔名
    const prefix = data.studentId + '_' + data.studentName;

    // 儲存切結書檔案到 Google Drive（不設公開分享，僅擁有者可存取）
    const waiverBlob = Utilities.newBlob(
      Utilities.base64Decode(data.waiverFileData),
      data.waiverFileType,
      prefix + '_切結書_' + data.waiverFileName
    );
    const waiverFileObj = folder.createFile(waiverBlob);
    const waiverUrl = waiverFileObj.getUrl();

    // 儲存隨兄妹就讀證明文件（如有）
    let siblingFileName = '';
    let siblingUrl = '';
    if (data.siblingFileData) {
      const siblingBlob = Utilities.newBlob(
        Utilities.base64Decode(data.siblingFileData),
        data.siblingFileType,
        prefix + '_隨兄妹證明_' + data.siblingFileName
      );
      const siblingFileObj = folder.createFile(siblingBlob);
      siblingFileName = data.siblingFileName;
      siblingUrl = siblingFileObj.getUrl();
    }

    // 寫入 Google Sheet
    const newRow = sheet.getLastRow() + 1;
    sheet.appendRow([
      Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss'),
      data.studentId,
      data.studentName,
      data.waiverFileName,
      '',  // 切結書連結
      siblingFileName,
      '',  // 隨兄妹證明連結
      '待處理'
    ]);

    // 用 RichTextValue 設定可點擊的超連結
    const waiverLinkCell = sheet.getRange(newRow, 5);
    const waiverRichText = SpreadsheetApp.newRichTextValue()
      .setText('開啟檔案')
      .setLinkUrl(0, 4, waiverUrl)
      .build();
    waiverLinkCell.setRichTextValue(waiverRichText);

    if (siblingUrl) {
      const siblingLinkCell = sheet.getRange(newRow, 7);
      const siblingRichText = SpreadsheetApp.newRichTextValue()
        .setText('開啟檔案')
        .setLinkUrl(0, 4, siblingUrl)
        .build();
      siblingLinkCell.setRichTextValue(siblingRichText);
    }

    // 設定「待處理」底色為黃色
    sheet.getRange(newRow, 8).setBackground('#FEF3C7');

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: error.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', message: 'Web App is running' }))
    .setMimeType(ContentService.MimeType.JSON);
}
