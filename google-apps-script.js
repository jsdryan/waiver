// ============================================================
// Google Apps Script - 部署為 Web App
//
// 使用步驟：
// 1. 開啟 Google Sheet，點選「擴充功能」>「Apps Script」
// 2. 將此檔案內容貼上到 Apps Script 編輯器
// 3. 修改下方 SHEET_ID 為你的 Google Sheet ID
// 4. 修改 FOLDER_ID 為你想存檔案的 Google Drive 資料夾 ID
// 5. 點選「部署」>「管理部署作業」>「編輯」> 版本選「新版本」> 部署
// 6. 複製部署後的 URL，貼到 index.html 的 SCRIPT_URL
// ============================================================

const SHEET_ID = '1wgQGr9TSMcCPbDkkSUHOR_atrW6iUs0Wv5oqMtCybNU';
const FOLDER_ID = '165bN7W2mN2FRSuyBfHg0MbhYev8N6clH';

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = SpreadsheetApp.openById(SHEET_ID).getActiveSheet();
    const folder = DriveApp.getFolderById(FOLDER_ID);

    // 確認表頭（首次自動建立）
    if (sheet.getLastRow() === 0) {
      const headers = [
        '序號',
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
      sheet.setColumnWidth(1, 50);   // 序號
      sheet.setColumnWidth(2, 160);  // 提交時間
      sheet.setColumnWidth(3, 130);  // 身分證字號
      sheet.setColumnWidth(4, 80);   // 姓名
      sheet.setColumnWidth(5, 160);  // 切結書檔名
      sheet.setColumnWidth(6, 100);  // 切結書連結
      sheet.setColumnWidth(7, 160);  // 隨兄妹證明檔名
      sheet.setColumnWidth(8, 100);  // 隨兄妹證明連結
      sheet.setColumnWidth(9, 100);  // 處理狀態
    }

    // 檔案命名：身分證字號_姓名_原始檔名
    const prefix = data.studentId + '_' + data.studentName;

    // 儲存切結書檔案到 Google Drive
    const waiverBlob = Utilities.newBlob(
      Utilities.base64Decode(data.waiverFileData),
      data.waiverFileType,
      prefix + '_切結書_' + data.waiverFileName
    );
    const waiverFileObj = folder.createFile(waiverBlob);
    waiverFileObj.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
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
      siblingFileObj.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      siblingFileName = data.siblingFileName;
      siblingUrl = siblingFileObj.getUrl();
    }

    // 計算序號
    const rowNum = sheet.getLastRow();

    // 寫入 Google Sheet
    const newRow = sheet.getLastRow() + 1;
    sheet.appendRow([
      rowNum,
      Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss'),
      data.studentId,
      data.studentName,
      data.waiverFileName,
      '',  // 切結書連結（用超連結公式）
      siblingFileName,
      '',  // 隨兄妹證明連結（用超連結公式）
      '待處理'
    ]);

    // 用 RichTextValue 設定可點擊的超連結（顯示「開啟檔案」而非長 URL）
    const waiverLinkCell = sheet.getRange(newRow, 6);
    const waiverRichText = SpreadsheetApp.newRichTextValue()
      .setText('開啟檔案')
      .setLinkUrl(0, 4, waiverUrl)
      .build();
    waiverLinkCell.setRichTextValue(waiverRichText);

    if (siblingUrl) {
      const siblingLinkCell = sheet.getRange(newRow, 8);
      const siblingRichText = SpreadsheetApp.newRichTextValue()
        .setText('開啟檔案')
        .setLinkUrl(0, 4, siblingUrl)
        .build();
      siblingLinkCell.setRichTextValue(siblingRichText);
    }

    // 設定「待處理」底色為黃色
    sheet.getRange(newRow, 9).setBackground('#FEF3C7');

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
