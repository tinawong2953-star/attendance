/**
 * ============================================================
 * 禾騰技術股份有限公司 考勤出缺勤系統後端 Google Apps Script
 * 【兩階段階層式簽核機制】
 *  - 申請送出：先寄信給「部門主管」（此時執行長不收信）
 *  - ≤ 8 小時：部門主管核准 ➔ 直接完成【已核准】
 *  - > 8 小時：部門主管核准 ➔ 狀態轉為【待執行長審核】➔ 系統自動發信給「執行長」➔ 執行長核准 ➔ 完成【已核准】
 *  - 任何階段退回 ➔ 直接轉為【退回修正】並中止流程
 * ============================================================
 */

// 正式發布的 Web App 網址
const WEB_APP_URL = "https://script.google.com/macros/s/AKfycbxSjVI3LZH7W3SNNFzSTsL9O-NJ4AB6YYIuCsC2K8N-Ni6wWj9twRC4NxmNUG-QtnuC/exec";

// 主管與執行長信箱
const DEFAULT_MANAGER_EMAIL = "tinawong@hetengtech.com";
const CEO_EMAIL = "nicolin@hetengtech.com";

// 部門專屬主管信箱對應表
const DEPT_MANAGERS = {
  "農水部": "tinawong@hetengtech.com",
  "流域部": "tinawong@hetengtech.com",
  "景觀部": "tinawong@hetengtech.com",
  "技研處": "tinawong@hetengtech.com",
  "行政處": "tinawong@hetengtech.com"
};

/**
 * 自動相容「考勤紀錄」或「工作表1」
 */
function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("考勤紀錄");
  if (!sheet) {
    sheet = ss.getSheetByName("工作表1");
  }
  if (!sheet) {
    sheet = ss.getActiveSheet();
  }
  return sheet;
}

/**
 * 試算表頂部管理選單
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("禾騰考勤管理 ⚙️")
    .addItem("📧 補寄【目前所選列】的審核通知信", "resendSelectedRowEmail")
    .addItem("⚡ 一鍵檢查並補發【所有待審核】通知", "resendAllPendingEmails")
    .addToUi();
}

/**
 * 1. POST 接收前端表單送出
 */
function doPost(e) {
  try {
    const rawData = e.postData.contents;
    const payload = JSON.parse(rawData);

    const sheet = getOrCreateSheet();
    const timestamp = Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy-MM-dd HH:mm:ss");

    const formNo = payload.formNo || ("ATT-" + Utilities.formatDate(new Date(), "Asia/Taipei", "yyyyMMdd") + "-001");
    const applyDate = payload.applyDate || Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy-MM-dd");
    const applicantName = payload.applicantName || "";
    const department = payload.department || "";
    const grandTotalHours = parseFloat(payload.grandTotalHours) || 0;
    const details = payload.details || [];

    // 寫入試算表
    details.forEach(item => {
      sheet.appendRow([
        timestamp,
        formNo,
        applyDate,
        applicantName,
        department,
        item.category || "",
        item.subType || "",
        item.startTime || "",
        item.endTime || "",
        item.hours || 0,
        item.reason || "",
        "待審核",      // 初始狀態：待部門主管審核
        ""             // 簽核歷程
      ]);
    });

    // 第一階段：只發送給部門主管（執行長此時不收信）
    sendManagerApprovalEmail(formNo, applyDate, applicantName, department, grandTotalHours, details);

    return ContentService.createTextOutput(JSON.stringify({ status: "success", formNo: formNo }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * 【階段一】發送給「部門主管」的審核信（執行長不收信）
 */
function sendManagerApprovalEmail(formNo, applyDate, applicantName, department, totalHours, details) {
  const managerEmail = DEPT_MANAGERS[department] || DEFAULT_MANAGER_EMAIL;
  const isOver8Hours = totalHours > 8;

  const approveUrl = `${WEB_APP_URL}?action=review&stage=manager&formNo=${encodeURIComponent(formNo)}&decision=approve`;
  const rejectUrl  = `${WEB_APP_URL}?action=review&stage=manager&formNo=${encodeURIComponent(formNo)}&decision=reject`;

  let detailsHtml = "";
  details.forEach(item => {
    detailsHtml += `
      <tr style="border-bottom: 1px solid #e2e8f0;">
        <td style="padding: 10px; font-weight: bold; color: #1e293b;">${item.category}（${item.subType}）</td>
        <td style="padding: 10px; color: #475569;">${item.startTime} ～ ${item.endTime}</td>
        <td style="padding: 10px; color: #4f46e5; font-weight: bold; text-align: center;">${item.hours} hr</td>
        <td style="padding: 10px; color: #64748b;">${item.reason || '—'}</td>
      </tr>
    `;
  });

  const emailSubject = `【主管簽核通知】${applicantName} - ${department}（單號：${formNo}，時數：${totalHours} 小時）`;

  const emailBodyHtml = `
    <div style="font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; max-width: 620px; margin: 0 auto; background: #ffffff; border-radius: 12px; border: 1px solid #e2e8f0; overflow: hidden;">
      
      <div style="background: linear-gradient(135deg, #4f46e5, #4338ca); padding: 24px 28px; color: #ffffff;">
        <div style="font-size: 13px; letter-spacing: 1px; opacity: 0.85; margin-bottom: 4px;">禾騰技術股份有限公司 · 部門主管審核</div>
        <h2 style="margin: 0; font-size: 20px; font-weight: 700;">出缺勤 / 請假申請單（第一階段審核）</h2>
      </div>

      <div style="padding: 24px 28px;">
        <div style="background: #f8fafc; border-radius: 8px; padding: 16px 20px; margin-bottom: 20px; border: 1px solid #e2e8f0;">
          <table style="width: 100%; font-size: 14px; color: #334155;">
            <tr><td style="padding: 4px 0; width: 90px; color: #64748b;">單據編號：</td><td style="padding: 4px 0; font-weight: 700; font-family: monospace;">${formNo}</td></tr>
            <tr><td style="padding: 4px 0; color: #64748b;">申請同仁：</td><td style="padding: 4px 0; font-weight: 700;">${applicantName}（${department}）</td></tr>
            <tr><td style="padding: 4px 0; color: #64748b;">填單日期：</td><td style="padding: 4px 0;">${applyDate}</td></tr>
            <tr><td style="padding: 4px 0; color: #64748b;">申請總時數：</td><td style="padding: 4px 0; font-size: 16px; font-weight: 800; color: #4f46e5;">${totalHours} 小時</td></tr>
            ${isOver8Hours ? `<tr><td colspan="2" style="padding-top: 6px; font-size: 12px; color: #0284c7; font-weight: bold;">ℹ️ 此申請時數超過 8 小時，您核准後系統將自動轉呈執行長進行第二階段核准。</td></tr>` : ''}
          </table>
        </div>

        <h3 style="font-size: 15px; color: #0f172a; margin: 0 0 10px 0;">申請項目明細</h3>
        <table style="width: 100%; border-collapse: collapse; font-size: 13px; margin-bottom: 24px;">
          <thead>
            <tr style="background: #f1f5f9; text-align: left; color: #475569;">
              <th style="padding: 8px 10px;">類別 / 細項</th>
              <th style="padding: 8px 10px;">申請時段</th>
              <th style="padding: 8px 10px; text-align: center;">時數</th>
              <th style="padding: 8px 10px;">備註事由</th>
            </tr>
          </thead>
          <tbody>${detailsHtml}</tbody>
        </table>

        <div style="background: #faf5ff; border: 1px dashed #d8b4fe; border-radius: 10px; padding: 20px; text-align: center; margin-top: 24px;">
          <div style="font-size: 14px; font-weight: bold; color: #6b21a8; margin-bottom: 14px;">部門主管線上審核批示</div>
          <div style="display: inline-block;">
            <a href="${approveUrl}" target="_blank" style="background: #16a34a; color: #ffffff; text-decoration: none; padding: 12px 28px; border-radius: 8px; font-weight: bold; font-size: 15px; display: inline-block; margin-right: 12px; box-shadow: 0 2px 4px rgba(22, 163, 74, 0.25);">
              ✅ 部門主管 核准
            </a>
            <a href="${rejectUrl}" target="_blank" style="background: #dc2626; color: #ffffff; text-decoration: none; padding: 12px 28px; border-radius: 8px; font-weight: bold; font-size: 15px; display: inline-block; box-shadow: 0 2px 4px rgba(220, 38, 38, 0.25);">
              ❌ 退回修正
            </a>
          </div>
          <div style="font-size: 11px; color: #94a3b8; margin-top: 12px;">點選上方按鈕後即時生效並更新試算表紀錄。</div>
        </div>

      </div>

      <div style="background: #f8fafc; border-top: 1px solid #e2e8f0; padding: 14px 28px; font-size: 11px; color: #94a3b8; text-align: center;">
        此為禾騰技術股份有限公司考勤系統自動發送之信件，請勿直接回覆。
      </div>

    </div>
  `;

  MailApp.sendEmail({
    to: managerEmail,
    subject: emailSubject,
    htmlBody: emailBodyHtml
  });
}

/**
 * 【階段二】發送給「執行長」的審核信（僅在主管核准且 >8 小時觸發）
 */
function sendCeoApprovalEmail(formNo, applyDate, applicantName, department, totalHours, details, managerApproveTime) {
  const approveUrl = `${WEB_APP_URL}?action=review&stage=ceo&formNo=${encodeURIComponent(formNo)}&decision=approve`;
  const rejectUrl  = `${WEB_APP_URL}?action=review&stage=ceo&formNo=${encodeURIComponent(formNo)}&decision=reject`;

  let detailsHtml = "";
  details.forEach(item => {
    detailsHtml += `
      <tr style="border-bottom: 1px solid #e2e8f0;">
        <td style="padding: 10px; font-weight: bold; color: #1e293b;">${item.category}（${item.subType}）</td>
        <td style="padding: 10px; color: #475569;">${item.startTime} ～ ${item.endTime}</td>
        <td style="padding: 10px; color: #4f46e5; font-weight: bold; text-align: center;">${item.hours} hr</td>
        <td style="padding: 10px; color: #64748b;">${item.reason || '—'}</td>
      </tr>
    `;
  });

  const emailSubject = `【呈報執行長簽核】${applicantName} - ${department}（單號：${formNo}，時數：${totalHours} 小時）`;

  const emailBodyHtml = `
    <div style="font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; max-width: 620px; margin: 0 auto; background: #ffffff; border-radius: 12px; border: 1px solid #e2e8f0; overflow: hidden;">
      
      <div style="background: linear-gradient(135deg, #0f172a, #334155); padding: 24px 28px; color: #ffffff;">
        <div style="font-size: 13px; letter-spacing: 1px; opacity: 0.85; margin-bottom: 4px;">禾騰技術股份有限公司 · 執行長最終簽核</div>
        <h2 style="margin: 0; font-size: 20px; font-weight: 700;">出缺勤申請呈報（超過8小時覆核）</h2>
      </div>

      <div style="padding: 24px 28px;">
        <div style="background: #f8fafc; border-radius: 8px; padding: 16px 20px; margin-bottom: 20px; border: 1px solid #e2e8f0;">
          <table style="width: 100%; font-size: 14px; color: #334155;">
            <tr><td style="padding: 4px 0; width: 90px; color: #64748b;">單據編號：</td><td style="padding: 4px 0; font-weight: 700; font-family: monospace;">${formNo}</td></tr>
            <tr><td style="padding: 4px 0; color: #64748b;">申請同仁：</td><td style="padding: 4px 0; font-weight: 700;">${applicantName}（${department}）</td></tr>
            <tr><td style="padding: 4px 0; color: #64748b;">填單日期：</td><td style="padding: 4px 0;">${applyDate}</td></tr>
            <tr><td style="padding: 4px 0; color: #64748b;">申請總時數：</td><td style="padding: 4px 0; font-size: 16px; font-weight: 800; color: #ea580c;">${totalHours} 小時</td></tr>
            <tr><td style="padding: 4px 0; color: #64748b;">主管初審：</td><td style="padding: 4px 0; font-weight: bold; color: #16a34a;">已於 ${managerApproveTime} 核准通過</td></tr>
          </table>
        </div>

        <h3 style="font-size: 15px; color: #0f172a; margin: 0 0 10px 0;">申請項目明細</h3>
        <table style="width: 100%; border-collapse: collapse; font-size: 13px; margin-bottom: 24px;">
          <thead>
            <tr style="background: #f1f5f9; text-align: left; color: #475569;">
              <th style="padding: 8px 10px;">類別 / 細項</th>
              <th style="padding: 8px 10px;">申請時段</th>
              <th style="padding: 8px 10px; text-align: center;">時數</th>
              <th style="padding: 8px 10px;">備註事由</th>
            </tr>
          </thead>
          <tbody>${detailsHtml}</tbody>
        </table>

        <div style="background: #fff7ed; border: 1px dashed #fdba74; border-radius: 10px; padding: 20px; text-align: center; margin-top: 24px;">
          <div style="font-size: 14px; font-weight: bold; color: #c2410c; margin-bottom: 14px;">執行長線上審核批示</div>
          <div style="display: inline-block;">
            <a href="${approveUrl}" target="_blank" style="background: #ea580c; color: #ffffff; text-decoration: none; padding: 12px 28px; border-radius: 8px; font-weight: bold; font-size: 15px; display: inline-block; margin-right: 12px; box-shadow: 0 2px 4px rgba(234, 88, 12, 0.25);">
              ✅ 執行長 核准
            </a>
            <a href="${rejectUrl}" target="_blank" style="background: #dc2626; color: #ffffff; text-decoration: none; padding: 12px 28px; border-radius: 8px; font-weight: bold; font-size: 15px; display: inline-block; box-shadow: 0 2px 4px rgba(220, 38, 38, 0.25);">
              ❌ 退回修正
            </a>
          </div>
          <div style="font-size: 11px; color: #94a3b8; margin-top: 12px;">點選上方按鈕後，單據將完成最終簽核。</div>
        </div>

      </div>

      <div style="background: #f8fafc; border-top: 1px solid #e2e8f0; padding: 14px 28px; font-size: 11px; color: #94a3b8; text-align: center;">
        此為禾騰技術股份有限公司考勤系統自動發送之信件，請勿直接回覆。
      </div>

    </div>
  `;

  MailApp.sendEmail({
    to: CEO_EMAIL,
    subject: emailSubject,
    htmlBody: emailBodyHtml
  });
}

/**
 * 2. GET 請求處理：支援 (1) 階層式審核 (2) 員工即時查詢
 */
function doGet(e) {
  const action = e.parameter.action;

  // ===== 審核動作 (主管或執行長點選) =====
  if (action === "review") {
    const stage = e.parameter.stage || "manager"; // "manager" 或 "ceo"
    const formNo = e.parameter.formNo;
    const decision = e.parameter.decision; // "approve" 或 "reject"
    const reviewTime = Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy-MM-dd HH:mm:ss");

    const sheet = getOrCreateSheet();
    const data = sheet.getDataRange().getValues();

    let applicantName = "";
    let department = "";
    let applyDate = "";
    let totalHours = 0;
    const details = [];
    const matchedRowIndices = [];

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]).trim() === String(formNo).trim()) {
        matchedRowIndices.push(i + 1);
        applicantName = String(data[i][3] || "");
        department = String(data[i][4] || "");
        applyDate = data[i][2] ? Utilities.formatDate(new Date(data[i][2]), "Asia/Taipei", "yyyy-MM-dd") : "";
        const h = parseFloat(data[i][9]) || 0;
        totalHours += h;
        details.push({
          category: String(data[i][5] || ""),
          subType: String(data[i][6] || ""),
          startTime: String(data[i][7] || ""),
          endTime: String(data[i][8] || ""),
          hours: h,
          reason: String(data[i][10] || "")
        });
      }
    }

    let finalStatus = "";
    let finalLog = "";
    let isApproved = decision === "approve";
    let isNextToCeo = false;

    if (!isApproved) {
      // 任何階段退回
      finalStatus = "退回修正";
      const who = (stage === "ceo") ? "執行長" : "部門主管";
      finalLog = `由【${who}】於 ${reviewTime} 退回修正`;
    } else {
      // 核准情境
      if (stage === "manager") {
        if (totalHours > 8) {
          // 超過 8 小時：進入第二階段
          finalStatus = "待執行長審核";
          finalLog = `由【部門主管】於 ${reviewTime} 核准通過，待執行長覆核`;
          isNextToCeo = true;
        } else {
          // 8 小時以內：直接完結
          finalStatus = "已核准";
          finalLog = `由【部門主管】於 ${reviewTime} 核准通過（簽核完畢）`;
        }
      } else if (stage === "ceo") {
        // 執行長核准：直接完結
        finalStatus = "已核准";
        finalLog = `由【執行長】於 ${reviewTime} 最終核准通過（簽核完畢）`;
      }
    }

    // 回寫試算表
    matchedRowIndices.forEach(rowIdx => {
      sheet.getRange(rowIdx, 12).setValue(finalStatus);
      const prevLog = String(sheet.getRange(rowIdx, 13).getValue() || "");
      sheet.getRange(rowIdx, 13).setValue(prevLog ? `${prevLog} ➔ ${finalLog}` : finalLog);
    });

    // 若需要轉呈執行長，立刻發信給執行長
    if (isNextToCeo) {
      sendCeoApprovalEmail(formNo, applyDate, applicantName, department, totalHours, details, reviewTime);
    }

    // 顯示漂亮的結果網頁
    const primaryColor = isApproved ? (isNextToCeo ? "#ea580c" : "#16a34a") : "#dc2626";
    const icon = isApproved ? (isNextToCeo ? "⏳" : "✅") : "⚠️";
    const titleText = isApproved ? (isNextToCeo ? "主管初審已完成（已轉呈執行長）" : "簽核作業已完成") : "單據已退回修正";

    const responseHtml = `
      <!DOCTYPE html>
      <html>
      <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>考勤簽核結果 - 禾騰技術股份有限公司</title>
        <style>
          body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; background: #f8fafc; display: flex; align-items: center; justify-content: center; min-height: 100vh; margin: 0; padding: 20px; box-sizing: border-box; }
          .card { background: white; border-radius: 16px; box-shadow: 0 10px 25px rgba(0,0,0,0.08); padding: 40px; max-width: 480px; width: 100%; text-align: center; border-top: 6px solid ${primaryColor}; }
          .icon { font-size: 56px; margin-bottom: 16px; }
          h2 { margin: 0 0 8px 0; color: #1e293b; font-size: 22px; }
          p { color: #64748b; font-size: 14px; line-height: 1.6; margin: 8px 0; }
          .badge { display: inline-block; background: ${isApproved ? (isNextToCeo ? '#fff7ed' : '#dcfce7') : '#fee2e2'}; color: ${primaryColor}; padding: 6px 16px; border-radius: 999px; font-weight: bold; font-size: 14px; margin: 16px 0; }
          .info-box { background: #f1f5f9; border-radius: 8px; padding: 14px; margin: 20px 0; text-align: left; font-size: 13px; color: #334155; }
          .footer { color: #94a3b8; font-size: 12px; margin-top: 24px; }
        </style>
      </head>
      <body>
        <div class="card">
          <div class="icon">${icon}</div>
          <h2>${titleText}</h2>
          <div class="badge">目前單據狀態：${finalStatus}</div>
          <p>單據編號 <strong>${formNo}</strong> 的批示紀錄已即時回寫至 Google 試算表。${isNextToCeo ? '<br><strong style="color:#ea580c;">系統已自動寄發簽核通知至執行長信箱 (' + CEO_EMAIL + ')。</strong>' : ''}</p>
          <div class="info-box">
            <div><strong>單據編號：</strong> ${formNo}</div>
            <div><strong>申請同仁：</strong> ${applicantName} (${department})</div>
            <div><strong>申請時數：</strong> ${totalHours} 小時</div>
            <div><strong>簽核歷程：</strong> ${finalLog}</div>
          </div>
          <div class="footer">感謝您的批示，您現在可以關閉此視窗。</div>
        </div>
      </body>
      </html>
    `;

    return HtmlService.createHtmlOutput(responseHtml);
  }

  // ===== 員工即時查詢進度 =====
  if (action === "query") {
    const keyword = (e.parameter.keyword || "").trim();
    const sheet = getOrCreateSheet();
    const data = sheet.getDataRange().getValues();
    const results = [];

    if (data.length > 1) {
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const rowFormNo = String(row[1] || "").trim();
        const rowApplicant = String(row[3] || "").trim();

        if (keyword === "" || rowFormNo.includes(keyword) || rowApplicant.includes(keyword)) {
          results.push({
            timestamp: row[0] ? Utilities.formatDate(new Date(row[0]), "Asia/Taipei", "yyyy-MM-dd HH:mm") : "",
            formNo: rowFormNo,
            applyDate: row[2] ? Utilities.formatDate(new Date(row[2]), "Asia/Taipei", "yyyy-MM-dd") : "",
            applicantName: rowApplicant,
            department: String(row[4] || ""),
            category: String(row[5] || ""),
            subType: String(row[6] || ""),
            startTime: String(row[7] || ""),
            endTime: String(row[8] || ""),
            hours: row[9] || 0,
            reason: String(row[10] || ""),
            status: String(row[11] || "待審核"),
            comment: String(row[12] || "")
          });
        }
      }
    }

    return ContentService.createTextOutput(JSON.stringify({ status: "success", data: results }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService.createTextOutput(JSON.stringify({ status: "running", message: "禾騰技術股份有限公司考勤系統後端 API 正常運行中" }))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * 手動補發目前所選列
 */
function resendSelectedRowEmail() {
  const ui = SpreadsheetApp.getUi();
  const sheet = getOrCreateSheet();
  const activeRange = sheet.getActiveRange();
  
  if (!activeRange) {
    ui.alert("請先用滑鼠點選要補發的那一列！");
    return;
  }

  const startRow = activeRange.getRow();
  const numRows = activeRange.getNumRows();

  if (startRow <= 1) {
    ui.alert("請點選第 2 列之後的申請資料！");
    return;
  }

  const data = sheet.getRange(startRow, 1, numRows, 13).getValues();
  const formsMap = {};

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const formNo = String(row[1] || "").trim();
    const status = String(row[11] || "").trim();
    if (!formNo) continue;

    if (!formsMap[formNo]) {
      formsMap[formNo] = {
        formNo: formNo,
        status: status,
        applyDate: row[2] ? Utilities.formatDate(new Date(row[2]), "Asia/Taipei", "yyyy-MM-dd") : "",
        applicantName: String(row[3] || ""),
        department: String(row[4] || ""),
        totalHours: 0,
        details: []
      };
    }

    const hours = parseFloat(row[9]) || 0;
    formsMap[formNo].totalHours += hours;
    formsMap[formNo].details.push({
      category: String(row[5] || ""),
      subType: String(row[6] || ""),
      startTime: String(row[7] || ""),
      endTime: String(row[8] || ""),
      hours: hours,
      reason: String(row[10] || "")
    });
  }

  const formKeys = Object.keys(formsMap);
  if (formKeys.length === 0) {
    ui.alert("選取的範圍內沒有有效的單據編號！");
    return;
  }

  formKeys.forEach(fNo => {
    const item = formsMap[fNo];
    if (item.status === "待執行長審核") {
      // 補寄給執行長
      sendCeoApprovalEmail(item.formNo, item.applyDate, item.applicantName, item.department, item.totalHours, item.details, "手動補發");
    } else {
      // 補寄給部門主管
      sendManagerApprovalEmail(item.formNo, item.applyDate, item.applicantName, item.department, item.totalHours, item.details);
    }
  });

  ui.alert(`✅ 補發成功！已寄出相應階段的審核通知信。`);
}

/**
 * 手動一鍵補發待審核單據
 */
function resendAllPendingEmails() {
  const ui = SpreadsheetApp.getUi();
  const sheet = getOrCreateSheet();
  const data = sheet.getDataRange().getValues();

  if (data.length <= 1) {
    ui.alert("試算表內目前無任何紀錄！");
    return;
  }

  const pendingForms = {};

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const status = String(row[11] || "").trim();
    const formNo = String(row[1] || "").trim();

    if ((status === "待審核" || status === "待執行長審核" || status === "") && formNo) {
      if (!pendingForms[formNo]) {
        pendingForms[formNo] = {
          formNo: formNo,
          status: status,
          applyDate: row[2] ? Utilities.formatDate(new Date(row[2]), "Asia/Taipei", "yyyy-MM-dd") : "",
          applicantName: String(row[3] || ""),
          department: String(row[4] || ""),
          totalHours: 0,
          details: []
        };
      }
      const hours = parseFloat(row[9]) || 0;
      pendingForms[formNo].totalHours += hours;
      pendingForms[formNo].details.push({
        category: String(row[5] || ""),
        subType: String(row[6] || ""),
        startTime: String(row[7] || ""),
        endTime: String(row[8] || ""),
        hours: hours,
        reason: String(row[10] || "")
      });
    }
  }

  const pendingKeys = Object.keys(pendingForms);
  if (pendingKeys.length === 0) {
    ui.alert("目前沒有任何待審核的單據需補寄！");
    return;
  }

  pendingKeys.forEach(fNo => {
    const item = pendingForms[fNo];
    if (item.status === "待執行長審核") {
      sendCeoApprovalEmail(item.formNo, item.applyDate, item.applicantName, item.department, item.totalHours, item.details, "手動補發");
    } else {
      sendManagerApprovalEmail(item.formNo, item.applyDate, item.applicantName, item.department, item.totalHours, item.details);
    }
  });

  ui.alert(`🎉 已完成補發作業！共補發 ${pendingKeys.length} 封通知。`);
}
