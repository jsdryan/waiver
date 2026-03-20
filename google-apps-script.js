// ============================================================
// Google Apps Script - 部署為 Web App
//
// 使用步驟：
// 1. 開啟 Google Sheet，點選「擴充功能」>「Apps Script」
// 2. 將此檔案內容貼上到 Apps Script 編輯器
// 3. 修改下方 SHEET_ID 為你的 Google Sheet ID
// 4. 修改 FOLDER_ID 為你想存檔案的 Google Drive 資料夾 ID
// 5. 點選「部署」>「新增部署作業」
//    - 類型選「網頁應用程式」
//    - 執行身分選「我」
//    - 誰可以存取選「所有人」
// 6. 複製部署後的 URL，貼到 index.html 的 SCRIPT_URL
// ============================================================

const SHEET_ID = 'YOUR_GOOGLE_SHEET_ID';
const FOLDER_ID = 'YOUR_GOOGLE_DRIVE_FOLDER_ID';

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = SpreadsheetApp.openById(SHEET_ID).getActiveSheet();
    const folder = DriveApp.getFolderById(FOLDER_ID);

    // 確認表頭（首次自動建立）
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        '時間戳記',
        '學生身分證字號',
        '學生姓名',
        '切結書檔案連結',
        '隨兄妹就讀證明文件連結'
      ]);
    }

    // 儲存切結書檔案到 Google Drive
    const waiverBlob = Utilities.newBlob(
      Utilities.base64Decode(data.waiverFileData),
      data.waiverFileType,
      data.waiverFileName
    );
    const waiverFileObj = folder.createFile(waiverBlob);
    waiverFileObj.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const waiverUrl = waiverFileObj.getUrl();

    // 儲存隨兄妹就讀證明文件（如有）
    let siblingUrl = '';
    if (data.siblingFileData) {
      const siblingBlob = Utilities.newBlob(
        Utilities.base64Decode(data.siblingFileData),
        data.siblingFileType,
        data.siblingFileName
      );
      const siblingFileObj = folder.createFile(siblingBlob);
      siblingFileObj.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      siblingUrl = siblingFileObj.getUrl();
    }

    // 寫入 Google Sheet
    sheet.appendRow([
      new Date().toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' }),
      data.studentId,
      data.studentName,
      waiverUrl,
      siblingUrl
    ]);

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
