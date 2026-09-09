/**
 * ============================================================
 * 禾騰技術股份有限公司 考勤出缺勤系統後端 Google Apps Script
 * 功能包含：
 *  1. onOpen: 試算表頂部加入自訂管理選單（可手動補發 Email）
 *  2. doPost: 接收網頁送單，寫入試算表並寄信
 *  3. doGet(?action=review): 主管信內一鍵核准/退回
 *  4. doGet(?action=query): 員工於前台即時查詢進度
 *  5. 手動補發機制：選取列補發 或 一鍵補發全部待審核單據
 * ============================================================
 */

// 主管信箱設定（依公司規定）：
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
 * 當打開試算表時，自動於頂部選單列加入管理工具
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("禾騰考勤管理 ⚙️")
    .addItem("📧 補寄【目前所選列】的主管審核信", "resendSelectedRowEmail")
    .addItem("⚡ 一鍵補寄【所有待審核】的單據通知", "resendAllPendingEmails")
    .addSeparator()
    .addItem("🔄 重新整理標題列格式", "formatHeaderStyles")
    .addToUi();
}

/**
 * 取得或建立「考勤紀錄」工作表，並確保首行標題欄位完整
 */
function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("考勤紀錄");
  if (!sheet) {
    sheet = ss.insertSheet("考勤紀錄");
    const headers = [
      "時間戳記",       // Col 1 (A)
      "單據編號",       // Col 2 (B)
      "填單日期",       // Col 3 (C)
      "申請人",         // Col 4 (D)
      "所屬部門",       // Col 5 (E)
      "申請類別",       // Col 6 (F)
      "假別細項",       // Col 7 (G)
      "開始時間",       // Col 8 (H)
      "結束時間",       // Col 9 (I)
      "申請時數",       // Col 10 (J)
      "事由備註",       // Col 11 (K)
      "審核狀態",       // Col 12 (L)
      "簽核歷程與備註"   // Col 13 (M)
    ];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setBackground("#4f46e5").setFontColor("#ffffff").setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function formatHeaderStyles() {
  const sheet = getOrCreateSheet();
  sheet.getRange(1, 1, 1, 13).setBackground("#4f46e5").setFontColor("#ffffff").setFontWeight("bold");
  SpreadsheetApp.getUi().alert("標題格式已更新完成！");
}

/**
 * 【手動應急功能 1】補發「滑鼠目前所選取之資料列」的審核通知
 */
function resendSelectedRowEmail() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const activeRange = sheet.getActiveRange();
  
  if (!activeRange) {
    ui.alert("請先用滑鼠點選要補發通知的那一列（可選多列）！");
    return;
  }

  const startRow = activeRange.getRow();
  const numRows = activeRange.getNumRows();

  if (startRow <= 1) {
    ui.alert("請選擇第 2 列之後的申請資料（第 1 列為標題）！");
    return;
  }

  const data = sheet.getRange(startRow, 1, numRows, 13).getValues();
  const formsMap = {};

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const formNo = String(row[1] || "").trim();
    if (!formNo) continue;

    if (!formsMap[formNo]) {
      formsMap[formNo] = {
        formNo: formNo,
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
    ui.alert("選取的範圍內沒有有效的單據資料！");
    return;
  }

  let sentCount = 0;
  formKeys.forEach(fNo => {
    const item = formsMap[fNo];
    sendApprovalEmail(item.formNo, item.applyDate, item.applicantName, item.department, item.totalHours, item.details);
    sentCount++;
  });

  ui.alert(`✅ 補發成功！共已寄出 ${sentCount} 筆單據的主管審核信。`);
}

/**
 * 【手動應急功能 2】一鍵檢查所有「待審核」的單據並發信
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

    // 只抓取「待審核」狀態
    if (status === "待審核" && formNo) {
      if (!pendingForms[formNo]) {
        pendingForms[formNo] = {
          formNo: formNo,
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
    ui.alert("目前試算表中沒有任何「待審核」的單據需補寄！");
    return;
  }

  const confirm = ui.alert("補寄確認", `發現共有 ${pendingKeys.length} 筆待審核單據，確定要立刻發送簽核信給主管嗎？`, ui.ButtonSet.YES_NO);
  if (confirm !== ui.Button.YES) return;

  let count = 0;
  pendingKeys.forEach(fNo => {
    const item = pendingForms[fNo];
    sendApprovalEmail(item.formNo, item.applyDate, item.applicantName, item.department, item.totalHours, item.details);
    count++;
  });

  ui.alert(`🎉 已完成補寄作業！共補發 ${count} 封主管審核信件。`);
}

/**
 * 1. 接收前端 POST 申請資料
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

    // 逐筆寫入明細
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
        "待審核",      // 初始審核狀態
        ""             // 簽核備註
      ]);
    });

    // 發送主管簽核 Email
    sendApprovalEmail(formNo, applyDate, applicantName, department, grandTotalHours, details);

    return ContentService.createTextOutput(JSON.stringify({ status: "success", formNo: formNo }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * 寄送簽核 Email 給主管（支援信內直接點擊核准/退回）
 */
function sendApprovalEmail(formNo, applyDate, applicantName, department, totalHours, details) {
  const managerEmail = DEPT_MANAGERS[department] || DEFAULT_MANAGER_EMAIL;
  let gasDeploymentUrl = "";
  try {
    gasDeploymentUrl = ScriptApp.getService().getUrl();
  } catch(e) {}

  // 超過 8 小時需另外通知執行長（CC）
  const shouldCcCeo = totalHours > 8;

  // 生成點擊審核 URL
  const approveUrl = gasDeploymentUrl ? `${gasDeploymentUrl}?action=review&formNo=${encodeURIComponent(formNo)}&decision=approve` : "#";
  const rejectUrl  = gasDeploymentUrl ? `${gasDeploymentUrl}?action=review&formNo=${encodeURIComponent(formNo)}&decision=reject` : "#";

  // 組合項目清單 HTML
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

  const emailSubject = `【待審核考勤申請單】${applicantName} - ${department}（單號：${formNo}，合計 ${totalHours} 小時）`;

  const emailBodyHtml = `
    <div style="font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; max-width: 620px; margin: 0 auto; background: #ffffff; border-radius: 12px; border: 1px solid #e2e8f0; overflow: hidden;">
      
      <div style="background: linear-gradient(135deg, #4f46e5, #4338ca); padding: 24px 28px; color: #ffffff;">
        <div style="font-size: 13px; letter-spacing: 1px; opacity: 0.85; margin-bottom: 4px;">禾騰技術股份有限公司 考勤審核通知</div>
        <h2 style="margin: 0; font-size: 20px; font-weight: 700;">出缺勤 / 請假申請單簽核</h2>
      </div>

      <div style="padding: 24px 28px;">
        <div style="background: #f8fafc; border-radius: 8px; padding: 16px 20px; margin-bottom: 20px; border: 1px solid #e2e8f0;">
          <table style="width: 100%; font-size: 14px; color: #334155;">
            <tr>
              <td style="padding: 4px 0; width: 90px; color: #64748b;">單據編號：</td>
              <td style="padding: 4px 0; font-weight: 700; font-family: monospace;">${formNo}</td>
            </tr>
            <tr>
              <td style="padding: 4px 0; color: #64748b;">申請同仁：</td>
              <td style="padding: 4px 0; font-weight: 700;">${applicantName}（${department}）</td>
            </tr>
            <tr>
              <td style="padding: 4px 0; color: #64748b;">填單日期：</td>
              <td style="padding: 4px 0;">${applyDate}</td>
            </tr>
            <tr>
              <td style="padding: 4px 0; color: #64748b;">申請總時數：</td>
              <td style="padding: 4px 0; font-size: 16px; font-weight: 800; color: #4f46e5;">${totalHours} 小時</td>
            </tr>
            ${shouldCcCeo ? `<tr><td colspan="2" style="padding-top: 6px; font-size: 12px; color: #ea580c; font-weight: bold;">⚠️ 申請時數超過 8 小時，已依規定同步副本通知執行長。</td></tr>` : ''}
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
          <tbody>
            ${detailsHtml}
          </tbody>
        </table>

        <div style="background: #faf5ff; border: 1px dashed #d8b4fe; border-radius: 10px; padding: 20px; text-align: center; margin-top: 24px;">
          <div style="font-size: 14px; font-weight: bold; color: #6b21a8; margin-bottom: 14px;">主管線上快速審核（點擊直接生效）</div>
          <div style="display: inline-block;">
            <a href="${approveUrl}" style="background: #16a34a; color: #ffffff; text-decoration: none; padding: 12px 28px; border-radius: 8px; font-weight: bold; font-size: 15px; display: inline-block; margin-right: 12px; box-shadow: 0 2px 4px rgba(22, 163, 74, 0.25);">
              ✅ 核准通過
            </a>
            <a href="${rejectUrl}" style="background: #dc2626; color: #ffffff; text-decoration: none; padding: 12px 28px; border-radius: 8px; font-weight: bold; font-size: 15px; display: inline-block; box-shadow: 0 2px 4px rgba(220, 38, 38, 0.25);">
              ❌ 退回修正
            </a>
          </div>
          <div style="font-size: 11px; color: #94a3b8; margin-top: 12px;">點選上方按鈕後，系統將自動回寫 Google 試算表並更新審核進度。</div>
        </div>

      </div>

      <div style="background: #f8fafc; border-top: 1px solid #e2e8f0; padding: 14px 28px; font-size: 11px; color: #94a3b8; text-align: center;">
        此為禾騰技術股份有限公司考勤差勤系統自動寄發之信件，請勿直接回覆。
      </div>

    </div>
  `;

  const mailOptions = {
    to: managerEmail,
    subject: emailSubject,
    htmlBody: emailBodyHtml
  };

  if (shouldCcCeo) {
    mailOptions.cc = CEO_EMAIL;
  }

  MailApp.sendEmail(mailOptions);
}

/**
 * 2. GET 請求處理：支援 (1) 主管郵件簽核 (2) 員工即時查詢
 */
function doGet(e) {
  const action = e.parameter.action;

  // ===== 功能 A: 主管由 Email 點選按鈕進行簽核 =====
  if (action === "review") {
    const formNo = e.parameter.formNo;
    const decision = e.parameter.decision;
    const newStatus = (decision === "approve") ? "已核准" : "退回修正";
    const actionText = (decision === "approve") ? "核准" : "退回";

    const sheet = getOrCreateSheet();
    const data = sheet.getDataRange().getValues();
    let updatedCount = 0;
    const reviewTime = Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy-MM-dd HH:mm:ss");

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]).trim() === String(formNo).trim()) {
        sheet.getRange(i + 1, 12).setValue(newStatus);
        sheet.getRange(i + 1, 13).setValue(`由主管於 ${reviewTime} 簽核【${newStatus}】`);
        updatedCount++;
      }
    }

    const isApproved = decision === "approve";
    const primaryColor = isApproved ? "#16a34a" : "#dc2626";
    const icon = isApproved ? "✅" : "⚠️";

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
          .badge { display: inline-block; background: ${isApproved ? '#dcfce7' : '#fee2e2'}; color: ${primaryColor}; padding: 6px 16px; border-radius: 999px; font-weight: bold; font-size: 14px; margin: 16px 0; }
          .info-box { background: #f1f5f9; border-radius: 8px; padding: 14px; margin: 20px 0; text-align: left; font-size: 13px; color: #334155; }
          .footer { color: #94a3b8; font-size: 12px; margin-top: 24px; }
        </style>
      </head>
      <body>
        <div class="card">
          <div class="icon">${icon}</div>
          <h2>簽核作業已完成</h2>
          <div class="badge">單據狀態已更新為：${newStatus}</div>
          <p>單據編號 <strong>${formNo}</strong> 的審核結果已即時回寫至 Google 試算表，申請員工亦可於前台查詢進度。</p>
          <div class="info-box">
            <div><strong>單據編號：</strong> ${formNo}</div>
            <div><strong>簽核動作：</strong> ${actionText}</div>
            <div><strong>完成時間：</strong> ${reviewTime}</div>
            <div><strong>更新筆數：</strong> 共 ${updatedCount} 筆項目</div>
          </div>
          <div class="footer">感謝您的批示，您現在可以關閉此視窗。</div>
        </div>
      </body>
      </html>
    `;

    return HtmlService.createHtmlOutput(responseHtml);
  }

  // ===== 功能 B: 員工即時查詢進度 =====
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
