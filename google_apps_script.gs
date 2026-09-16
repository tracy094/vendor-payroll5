/**
 * 廠商工資試算與 Excel 核對系統 - Google Apps Script (GAS) 部署腳本
 * 
 * 使用方式：
 * 1. 開啟目標 Google 試算表 (Spreadsheet)。
 * 2. 點擊頂部選單「擴充功能」 -> 「Google Apps Script」。
 * 3. 將本檔案全文複製貼入 `Code.gs`。
 * 4. 點選右上角「部署」 -> 「新增部署」。
 * 5. 類型選擇「Web 應用程式 (Web App)」。
 * 6. 執行身份選擇：「我 (Me)」，誰可以存取：「任何人 (Anyone)」。
 * 7. 點選「部署」並複製產出的 Web App URL 網址貼回系統中使用。
 */

function doPost(e) {
  try {
    const contents = e.postData.contents;
    const data = JSON.parse(contents);

    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    // 生成當天日期命名（例如：2026-09-16_請款試算）
    const todayStr = data.date || Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd");
    const sheetName = `${todayStr}_請款試算`;

    let sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      sheet = spreadsheet.insertSheet(sheetName);
    } else {
      sheet.clear(); // 若已存在則清空覆寫
    }

    // 寫入標頭欄位
    const headers = [
      "日期", "廠商名稱", "工人姓名", "工時(小時)", 
      "前8h費率($275)", "9-10h加班($280)", ">10h加班($345)", 
      "階梯工費小計", "額外加項費用", "未稅金額", "5% 營業稅", 
      "系統含稅總額", "廠商請款金額", "差異金額", "備註", "異常狀態"
    ];

    const rows = [headers];

    if (data.records && Array.isArray(data.records)) {
      data.records.forEach(function(r) {
        rows.push([
          r.date || todayStr,
          r.vendorName || "",
          r.workerName || "",
          r.hoursWorked || 0,
          r.regularWage || 0,
          r.ot1Wage || 0,
          r.ot2Wage || 0,
          r.steppedWageSubtotal || 0,
          r.extraFees || 0,
          r.subtotalNoTax || 0,
          r.tax || 0,
          r.totalWithTax || 0,
          r.claimedAmount !== null && r.claimedAmount !== undefined ? r.claimedAmount : "",
          r.diffAmount || 0,
          r.note || "",
          r.hasAnomaly ? (r.anomalies ? r.anomalies.map(a => a.message).join("; ") : "異常") : "正常"
        ]);
      });
    }

    // 寫入資料陣列
    sheet.getRange(1, 1, rows.length, headers.length).setValues(rows);

    // 套用 Basic 樣式格式
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground("#0284c7"); // 品牌藍
    headerRange.setFontColor("#ffffff");
    headerRange.setFontWeight("bold");
    headerRange.setHorizontalAlignment("center");

    // 設定內文格線與對齊
    const dataRange = sheet.getRange(1, 1, rows.length, headers.length);
    dataRange.setBorder(true, true, true, true, true, true, "#cbd5e1", SpreadsheetApp.BorderStyle.SOLID);
    
    // 自動調整欄寬
    for (let i = 1; i <= headers.length; i++) {
      sheet.autoResizeColumn(i);
    }

    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      sheetName: sheetName,
      message: `成功建立並寫入工作表 [${sheetName}]，共 ${data.records ? data.records.length : 0} 筆紀錄`
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService.createTextOutput("廠商工資試算與 Excel 核對系統 - GAS Web App API 正常運作中。請使用 POST 請求。");
}
