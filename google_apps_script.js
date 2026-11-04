/**
 * 新竹縣消防局救災救護指揮中心 - 院前啟動輸血調度管制系統
 * Google Apps Script 後端串接腳本 (Google Sheets + Google Drive)
 * 
 * 【設定說明】
 * 1. 請在 Google 雲端硬碟建立一個新的「Google 試算表」。
 * 2. 點擊試算表上方選單「擴充功能」->「Apps Script」。
 * 3. 清空裡面的內容，將此檔案全部內容貼上。
 * 4. 將下方的 FOLDER_ID 替換成您的「Google 雲端資料夾 ID」。
 * 5. 點擊右上角「部署」->「新部署」：
 *    - 種類選「網路應用程式」(Web App)
 *    - 執行身分選「我 (您的帳號)」
 *    - 誰可以存取選「所有人 (Anyone)」
 * 6. 點選「部署」後，複製產生的「網路應用程式網址 (Web App URL)」，將其提供給系統即可！
 */

// 請替換為您 Google 雲端硬碟要儲存附件或資料庫備份的「資料夾 ID」
// (資料夾網址 https://drive.google.com/drive/folders/xxxxxxxxxx 中的 xxxxxxxxxx)
const FOLDER_ID = "請在此填入您的雲端資料夾ID";

function doGet(e) {
  try {
    const action = e.parameter.action || "getCases";
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    if (action === "getCases") {
      let sheet = ss.getSheetByName("案件紀錄");
      if (!sheet) {
        return createResponse({ status: "success", data: [] });
      }
      const rows = sheet.getDataRange().getValues();
      if (rows.length <= 1) {
        return createResponse({ status: "success", data: [] });
      }
      const headers = rows[0];
      const records = [];
      for (let i = 1; i < rows.length; i++) {
        let item = {};
        for (let j = 0; j < headers.length; j++) {
          item[headers[j]] = rows[i][j];
        }
        // 如果 logs 欄位是 JSON 字串，解析回陣列
        if (item.logs) {
          try { item.logs = JSON.parse(item.logs); } catch(err) { item.logs = [item.logs]; }
        }
        records.push(item);
      }
      return createResponse({ status: "success", data: records });
    }
    
    if (action === "getHospitalPhones") {
      let sheet = ss.getSheetByName("醫院電話表");
      if (!sheet) {
        return createResponse({ status: "success", data: {} });
      }
      const rows = sheet.getDataRange().getValues();
      let phoneMap = {};
      for (let i = 1; i < rows.length; i++) {
        if (rows[i][0]) {
          phoneMap[rows[i][0]] = rows[i][1] || "";
        }
      }
      return createResponse({ status: "success", data: phoneMap });
    }

    return createResponse({ status: "error", message: "未知的請求動作" });
  } catch (error) {
    return createResponse({ status: "error", message: error.toString() });
  }
}

function doPost(e) {
  try {
    const postData = JSON.parse(e.postData.contents);
    const action = postData.action;
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // 儲存結案案件
    if (action === "saveCase") {
      const c = postData.caseData;
      let sheet = ss.getSheetByName("案件紀錄");
      if (!sheet) {
        sheet = ss.insertSheet("案件紀錄");
        sheet.appendRow([
          "caseId", "startTime", "closedTime", "dispatcher", "dispatchUnit",
          "location", "hasTP", "backupTPUnit", "bloodBankHospital", "bloodBankPhone",
          "pickupSquad", "pickupCarCode", "destHospital", "destHospitalPhone",
          "status", "logs", "savedAt"
        ]);
        sheet.getRange(1, 1, 1, 17).setFontWeight("bold").setBackground("#f1f5f9");
      }

      sheet.appendRow([
        c.caseId || "",
        c.startTime || "",
        c.closedTime || "",
        c.dispatcher || "",
        c.dispatchUnit || "",
        c.location || "",
        c.hasTP ? "是" : "否",
        c.backupTPUnit || "",
        c.bloodBankHospital || "",
        c.bloodBankPhone || "",
        c.pickupSquad || "",
        c.pickupCarCode || "",
        c.destHospital || "",
        c.destHospitalPhone || "",
        c.status || "已結案",
        JSON.stringify(c.logs || []),
        new Date().toISOString()
      ]);

      // 若有夾帶檔案要上傳至雲端資料夾
      if (postData.fileData && FOLDER_ID && FOLDER_ID !== "請在此填入您的雲端資料夾ID") {
        try {
          const folder = DriveApp.getFolderById(FOLDER_ID);
          const contentType = postData.fileData.mimeType || "application/octet-stream";
          const bytes = Utilities.base64Decode(postData.fileData.base64);
          const blob = Utilities.newBlob(bytes, contentType, postData.fileData.filename || (c.caseId + "_attachment"));
          folder.createFile(blob);
        } catch(err) {
          console.error("檔案存入雲端資料夾失敗:", err);
        }
      }

      return createResponse({ status: "success", message: "案件已成功存入 Google 試算表與雲端硬碟" });
    }

    // 更新醫院電話表
    if (action === "saveHospitalPhones") {
      const phoneMap = postData.phoneMap || {};
      let sheet = ss.getSheetByName("醫院電話表");
      if (!sheet) {
        sheet = ss.insertSheet("醫院電話表");
      }
      sheet.clearContents();
      sheet.appendRow(["醫院名稱", "電話號碼", "最後更新時間"]);
      sheet.getRange(1, 1, 1, 3).setFontWeight("bold").setBackground("#f1f5f9");

      for (let name in phoneMap) {
        sheet.appendRow([name, phoneMap[name], new Date().toISOString()]);
      }
      return createResponse({ status: "success", message: "醫院電話表已更新" });
    }

    return createResponse({ status: "error", message: "未知的 POST 動作" });
  } catch (error) {
    return createResponse({ status: "error", message: error.toString() });
  }
}

function createResponse(result) {
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}
