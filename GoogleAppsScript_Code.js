/**
 * ============================================================
 * 禾騰技術股份有限公司 考勤出缺勤系統後端 Google Apps Script
 * 【兩階段階層式簽核機制】
 *  - 申請送出：先寄信給「部門主管」（此時執行長不收信）
 *  - 時數 < 8 小時：部門主管核准 ➔ 直接完成【已核准】（流程結束）
 *  - 時數 >= 8 小時：部門主管核准 ➔ 狀態更新 ➔ 系統自動發信給「執行長」➔ 執行長核准 ➔ 完成【已核准】
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
 * 規則：時數 >= 8 小時提示需呈報執行長
 */
function sendManagerApprovalEmail(formNo, applyDate, applicantName, department, totalHours, details) {
  const managerEmail = DEPT_MANAGERS[department] || DEFAULT_MANAGER_EMAIL;
  const isRequireCeo = totalHours >= 8;

  const approveUrl = `${WEB_APP_URL}?action=review&stage=manager&formNo=${encodeURIComponent(formNo)}&decision=approve`;
  const rejectUrl  = `${WEB_APP_URL}?action=review&stage=manager&formNo=${encodeURIComponent(formNo)}&decision=reject`;

  let detailsHtml = "";
  details.forEach(item => {
    detailsHtml += `
      <tr style="border-bottom: 1px solid #f1f5f9;">
        <td style="padding: 12px 14px; font-weight: 700; color: #1e293b;">${item.category} <span style="font-weight: normal; color: #64748b; font-size: 12px;">(${item.subType})</span></td>
        <td style="padding: 12px 14px; color: #475569; font-size: 13px;">${item.startTime} ～ ${item.endTime}</td>
        <td style="padding: 12px 14px; color: #4f46e5; font-weight: 800; text-align: center; font-size: 14px;">${item.hours} hr</td>
        <td style="padding: 12px 14px; color: #64748b; font-size: 12px;">${item.reason || '—'}</td>
      </tr>
    `;
  });

  const emailSubject = `【主管簽核通知】${applicantName} - ${department}（單號：${formNo}，時數：${totalHours} 小時）`;

  const emailBodyHtml = `
    <div style="font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Noto Sans TC', sans-serif; max-width: 620px; margin: 20px auto; background: #ffffff; border-radius: 16px; border: 1px solid #e2e8f0; box-shadow: 0 10px 25px rgba(0,0,0,0.04); overflow: hidden;">
      
      <!-- 頂部品牌 Banner -->
      <div style="background: linear-gradient(135deg, #3b82f6 0%, #4f46e5 100%); padding: 28px 32px; color: #ffffff;">
        <div style="display: inline-block; font-size: 11px; font-weight: 800; letter-spacing: 1.5px; text-transform: uppercase; background: rgba(255,255,255,0.2); padding: 4px 10px; border-radius: 20px; margin-bottom: 8px;">
          禾騰技術股份有限公司 · 考勤審核
        </div>
        <h2 style="margin: 0; font-size: 22px; font-weight: 800; letter-spacing: -0.5px;">出缺勤 / 請假申請單</h2>
        <div style="font-size: 13px; opacity: 0.9; margin-top: 4px;">第一階段：部門主管審核</div>
      </div>

      <div style="padding: 28px 32px;">
        
        <!-- 資訊卡片 -->
        <div style="background: #f8fafc; border-radius: 12px; padding: 20px; margin-bottom: 24px; border: 1px solid #edf2f7;">
          <table style="width: 100%; border-collapse: collapse; font-size: 14px;">
            <tr>
              <td style="padding: 6px 0; width: 90px; color: #64748b; font-size: 13px;">單據編號</td>
              <td style="padding: 6px 0; font-weight: 800; font-family: monospace; color: #0f172a; font-size: 15px;">${formNo}</td>
            </tr>
            <tr>
              <td style="padding: 6px 0; color: #64748b; font-size: 13px;">申請同仁</td>
              <td style="padding: 6px 0; font-weight: 700; color: #1e293b;">${applicantName} <span style="font-weight: normal; color: #64748b;">(${department})</span></td>
            </tr>
            <tr>
              <td style="padding: 6px 0; color: #64748b; font-size: 13px;">填單日期</td>
              <td style="padding: 6px 0; color: #334155;">${applyDate}</td>
            </tr>
            <tr>
              <td style="padding: 6px 0; color: #64748b; font-size: 13px;">申請總時數</td>
              <td style="padding: 6px 0; font-size: 18px; font-weight: 900; color: #4f46e5;">${totalHours} <span style="font-size: 13px; font-weight: bold; color: #64748b;">小時</span></td>
            </tr>
            ${isRequireCeo ? `
            <tr>
              <td colspan="2" style="padding-top: 10px;">
                <div style="background: #eff6ff; border-left: 4px solid #3b82f6; padding: 8px 12px; border-radius: 4px; font-size: 12px; color: #1e40af; font-weight: 600;">
                  ⚡ 本單申請時數滿 8 小時（含）以上，依公司規定您核准後將自動轉呈執行長覆核。
                </div>
              </td>
            </tr>` : ''}
          </table>
        </div>

        <!-- 申請明細表 -->
        <div style="font-size: 14px; font-weight: 800; color: #0f172a; margin-bottom: 10px;">申請項目明細</div>
        <table style="width: 100%; border-collapse: collapse; border: 1px solid #e2e8f0; border-radius: 8px; overflow: hidden; margin-bottom: 28px;">
          <thead>
            <tr style="background: #f8fafc; text-align: left; font-size: 12px; color: #64748b; border-bottom: 1px solid #e2e8f0;">
              <th style="padding: 10px 14px;">類別 / 細項</th>
              <th style="padding: 10px 14px;">申請時段</th>
              <th style="padding: 10px 14px; text-align: center;">時數</th>
              <th style="padding: 10px 14px;">備註事由</th>
            </tr>
          </thead>
          <tbody>${detailsHtml}</tbody>
        </table>

        <!-- 精美現代化 SaaS 簽核按鈕區 -->
        <div style="background: #fdfefe; border: 1px solid #e0e7ff; border-radius: 14px; padding: 24px 20px; text-align: center; box-shadow: inset 0 2px 4px rgba(0,0,0,0.01);">
          <div style="font-size: 14px; font-weight: 800; color: #1e1b4b; margin-bottom: 18px; letter-spacing: 0.2px;">
            請點選下方按鈕進行線上批示
          </div>
          
          <table style="margin: 0 auto; border-collapse: separate; border-spacing: 16px 0;">
            <tr>
              <td>
                <a href="${approveUrl}" target="_blank" style="background: linear-gradient(180deg, #10b981 0%, #059669 100%); color: #ffffff; text-decoration: none; padding: 13px 32px; border-radius: 10px; font-weight: 800; font-size: 15px; display: inline-block; box-shadow: 0 4px 14px rgba(16, 185, 129, 0.35); text-shadow: 0 1px 2px rgba(0,0,0,0.2); letter-spacing: 0.5px;">
                  ✔&nbsp;&nbsp;部門主管 核准
                </a>
              </td>
              <td>
                <a href="${rejectUrl}" target="_blank" style="background: linear-gradient(180deg, #f43f5e 0%, #e11d48 100%); color: #ffffff; text-decoration: none; padding: 13px 32px; border-radius: 10px; font-weight: 800; font-size: 15px; display: inline-block; box-shadow: 0 4px 14px rgba(225, 29, 72, 0.3); text-shadow: 0 1px 2px rgba(0,0,0,0.2); letter-spacing: 0.5px;">
                  ✕&nbsp;&nbsp;退回修正
                </a>
              </td>
            </tr>
          </table>

          <div style="font-size: 11px; color: #94a3b8; margin-top: 16px;">
            點選按鈕後將立即開啟確認頁面，並同步回寫 Google 試算表。
          </div>
        </div>

      </div>

      <!-- 頁尾 -->
      <div style="background: #f8fafc; border-top: 1px solid #f1f5f9; padding: 16px; font-size: 11px; color: #94a3b8; text-align: center;">
        此為禾騰技術股份有限公司考勤系統自動通知，請勿直接回覆此信件。
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
 * 【階段二】發送給「執行長」的審核信（僅在主管核准且 >= 8 小時觸發）
 */
function sendCeoApprovalEmail(formNo, applyDate, applicantName, department, totalHours, details, managerApproveTime) {
  const approveUrl = `${WEB_APP_URL}?action=review&stage=ceo&formNo=${encodeURIComponent(formNo)}&decision=approve`;
  const rejectUrl  = `${WEB_APP_URL}?action=review&stage=ceo&formNo=${encodeURIComponent(formNo)}&decision=reject`;

  let detailsHtml = "";
  details.forEach(item => {
    detailsHtml += `
      <tr style="border-bottom: 1px solid #f1f5f9;">
        <td style="padding: 12px 14px; font-weight: 700; color: #1e293b;">${item.category} <span style="font-weight: normal; color: #64748b; font-size: 12px;">(${item.subType})</span></td>
        <td style="padding: 12px 14px; color: #475569; font-size: 13px;">${item.startTime} ～ ${item.endTime}</td>
        <td style="padding: 12px 14px; color: #ea580c; font-weight: 800; text-align: center; font-size: 14px;">${item.hours} hr</td>
        <td style="padding: 12px 14px; color: #64748b; font-size: 12px;">${item.reason || '—'}</td>
      </tr>
    `;
  });

  const emailSubject = `【呈報執行長簽核】${applicantName} - ${department}（單號：${formNo}，時數：${totalHours} 小時）`;

  const emailBodyHtml = `
    <div style="font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Noto Sans TC', sans-serif; max-width: 620px; margin: 20px auto; background: #ffffff; border-radius: 16px; border: 1px solid #e2e8f0; box-shadow: 0 10px 25px rgba(0,0,0,0.04); overflow: hidden;">
      
      <!-- 頂部黑金高階 Banner -->
      <div style="background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%); padding: 28px 32px; color: #ffffff;">
        <div style="display: inline-block; font-size: 11px; font-weight: 800; letter-spacing: 1.5px; text-transform: uppercase; background: rgba(255,255,255,0.15); padding: 4px 10px; border-radius: 20px; margin-bottom: 8px; color: #fdba74;">
          禾騰技術股份有限公司 · 執行長核決
        </div>
        <h2 style="margin: 0; font-size: 22px; font-weight: 800; letter-spacing: -0.5px;">考勤申請 最終呈報覆核</h2>
        <div style="font-size: 13px; opacity: 0.85; margin-top: 4px;">第二階段：滿 8 小時（含）以上管理階層覆核</div>
      </div>

      <div style="padding: 28px 32px;">
        
        <!-- 資訊卡片 -->
        <div style="background: #f8fafc; border-radius: 12px; padding: 20px; margin-bottom: 24px; border: 1px solid #edf2f7;">
          <table style="width: 100%; border-collapse: collapse; font-size: 14px;">
            <tr>
              <td style="padding: 6px 0; width: 90px; color: #64748b; font-size: 13px;">單據編號</td>
              <td style="padding: 6px 0; font-weight: 800; font-family: monospace; color: #0f172a; font-size: 15px;">${formNo}</td>
            </tr>
            <tr>
              <td style="padding: 6px 0; color: #64748b; font-size: 13px;">申請同仁</td>
              <td style="padding: 6px 0; font-weight: 700; color: #1e293b;">${applicantName} <span style="font-weight: normal; color: #64748b;">(${department})</span></td>
            </tr>
            <tr>
              <td style="padding: 6px 0; color: #64748b; font-size: 13px;">填單日期</td>
              <td style="padding: 6px 0; color: #334155;">${applyDate}</td>
            </tr>
            <tr>
              <td style="padding: 6px 0; color: #64748b; font-size: 13px;">申請總時數</td>
              <td style="padding: 6px 0; font-size: 18px; font-weight: 900; color: #ea580c;">${totalHours} <span style="font-size: 13px; font-weight: bold; color: #64748b;">小時</span></td>
            </tr>
            <tr>
              <td style="padding: 6px 0; color: #64748b; font-size: 13px;">主管初審紀錄</td>
              <td style="padding: 6px 0; font-weight: bold; color: #16a34a; font-size: 13px;">✔ 部門主管已於 ${managerApproveTime} 初審通過</td>
            </tr>
          </table>
        </div>

        <!-- 申請明細表 -->
        <div style="font-size: 14px; font-weight: 800; color: #0f172a; margin-bottom: 10px;">申請項目明細</div>
        <table style="width: 100%; border-collapse: collapse; border: 1px solid #e2e8f0; border-radius: 8px; overflow: hidden; margin-bottom: 28px;">
          <thead>
            <tr style="background: #f8fafc; text-align: left; font-size: 12px; color: #64748b; border-bottom: 1px solid #e2e8f0;">
              <th style="padding: 10px 14px;">類別 / 細項</th>
              <th style="padding: 10px 14px;">申請時段</th>
              <th style="padding: 10px 14px; text-align: center;">時數</th>
              <th style="padding: 10px 14px;">備註事由</th>
            </tr>
          </thead>
          <tbody>${detailsHtml}</tbody>
        </table>

        <!-- 執行長核准按鈕區 -->
        <div style="background: #fffbf5; border: 1px solid #fed7aa; border-radius: 14px; padding: 24px 20px; text-align: center;">
          <div style="font-size: 14px; font-weight: 800; color: #7c2d12; margin-bottom: 18px;">
            請執行長點選批示（最終核決）
          </div>
          
          <table style="margin: 0 auto; border-collapse: separate; border-spacing: 16px 0;">
            <tr>
              <td>
                <a href="${approveUrl}" target="_blank" style="background: linear-gradient(180deg, #ea580c 0%, #c2410c 100%); color: #ffffff; text-decoration: none; padding: 13px 34px; border-radius: 10px; font-weight: 800; font-size: 15px; display: inline-block; box-shadow: 0 4px 14px rgba(234, 88, 12, 0.35); text-shadow: 0 1px 2px rgba(0,0,0,0.25); letter-spacing: 0.5px;">
                  ✔&nbsp;&nbsp;執行長 核准通過
                </a>
              </td>
              <td>
                <a href="${rejectUrl}" target="_blank" style="background: linear-gradient(180deg, #f43f5e 0%, #e11d48 100%); color: #ffffff; text-decoration: none; padding: 13px 34px; border-radius: 10px; font-weight: 800; font-size: 15px; display: inline-block; box-shadow: 0 4px 14px rgba(225, 29, 72, 0.3); text-shadow: 0 1px 2px rgba(0,0,0,0.25); letter-spacing: 0.5px;">
                  ✕&nbsp;&nbsp;退回修正
                </a>
              </td>
            </tr>
          </table>

          <div style="font-size: 11px; color: #9a3412; margin-top: 16px; opacity: 0.8;">
            此為最終核決階段，點選核准後單據將正式完成審批並備存於試算表。
          </div>
        </div>

      </div>

      <!-- 頁尾 -->
      <div style="background: #f8fafc; border-top: 1px solid #f1f5f9; padding: 16px; font-size: 11px; color: #94a3b8; text-align: center;">
        此為禾騰技術股份有限公司考勤系統自動通知，請勿直接回覆此信件。
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
      // 核准情境：滿 8 小時（含）以上需轉呈執行長
      if (stage === "manager") {
        if (totalHours >= 8) {
          // 滿 8 小時（含）：進入第二階段
          finalStatus = "待審核"; // 保持相容「待審核」，同時在備註標明
          finalLog = `由【部門主管】於 ${reviewTime} 初審通過，待執行長覆核`;
          isNextToCeo = true;
        } else {
          // 未滿 8 小時：直接完結
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

    // 顯示現代化簽核完成回饋頁面
    const primaryColor = isApproved ? (isNextToCeo ? "#ea580c" : "#10b981") : "#f43f5e";
    const icon = isApproved ? (isNextToCeo ? "⏳" : "✔") : "✕";
    const titleText = isApproved ? (isNextToCeo ? "部門主管初審已完成（已轉呈執行長覆核）" : "簽核作業已順利完成") : "單據已退回修正";

    const responseHtml = `
      <!DOCTYPE html>
      <html>
      <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>考勤簽核結果 - 禾騰技術股份有限公司</title>
        <style>
          body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; background: #f8fafc; display: flex; align-items: center; justify-content: center; min-height: 100vh; margin: 0; padding: 20px; box-sizing: border-box; }
          .card { background: white; border-radius: 20px; box-shadow: 0 15px 35px rgba(0,0,0,0.06); padding: 44px 36px; max-width: 480px; width: 100%; text-align: center; border-top: 6px solid ${primaryColor}; }
          .icon-wrap { width: 64px; height: 64px; border-radius: 50%; background: ${primaryColor}15; color: ${primaryColor}; font-size: 32px; display: flex; align-items: center; justify-content: center; margin: 0 auto 16px auto; font-weight: bold; }
          h2 { margin: 0 0 10px 0; color: #0f172a; font-size: 22px; font-weight: 800; }
          p { color: #64748b; font-size: 14px; line-height: 1.6; margin: 8px 0; }
          .badge { display: inline-block; background: ${isApproved ? (isNextToCeo ? '#fff7ed' : '#ecfdf5') : '#fff1f2'}; color: ${primaryColor}; padding: 7px 20px; border-radius: 999px; font-weight: 800; font-size: 14px; margin: 16px 0; border: 1px solid ${primaryColor}30; }
          .info-box { background: #f8fafc; border-radius: 12px; padding: 18px; margin: 20px 0; text-align: left; font-size: 13px; color: #334155; border: 1px solid #e2e8f0; line-height: 1.8; }
          .footer { color: #94a3b8; font-size: 12px; margin-top: 24px; }
        </style>
      </head>
      <body>
        <div class="card">
          <div class="icon-wrap">${icon}</div>
          <h2>${titleText}</h2>
          <div class="badge">${isNextToCeo ? '單據已送交執行長覆核' : '單據狀態：' + finalStatus}</div>
          <p>單據編號 <strong>${formNo}</strong> 的批示結果已即時回寫至 Google 試算表。${isNextToCeo ? '<br><strong style="color:#ea580c;">系統已自動發信通知執行長進行最終審核。</strong>' : ''}</p>
          <div class="info-box">
            <div><strong>單據編號：</strong> ${formNo}</div>
            <div><strong>申請同仁：</strong> ${applicantName} (${department})</div>
            <div><strong>申請時數：</strong> ${totalHours} 小時</div>
            <div><strong>最新進度：</strong> ${finalLog}</div>
          </div>
          <div class="footer">感謝您的批示，您現在可以安心關閉此分頁。</div>
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
    const comment = String(row[12] || "").trim();
    if (!formNo) continue;

    if (!formsMap[formNo]) {
      formsMap[formNo] = {
        formNo: formNo,
        comment: comment,
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
    if (item.comment.includes("待執行長覆核")) {
      sendCeoApprovalEmail(item.formNo, item.applyDate, item.applicantName, item.department, item.totalHours, item.details, "主管已審核");
    } else {
      sendManagerApprovalEmail(item.formNo, item.applyDate, item.applicantName, item.department, item.totalHours, item.details);
    }
  });

  ui.alert(`✅ 補發成功！已發送相應審核信件。`);
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
    const comment = String(row[12] || "").trim();
    const formNo = String(row[1] || "").trim();

    if ((status === "待審核" || status === "") && formNo) {
      if (!pendingForms[formNo]) {
        pendingForms[formNo] = {
          formNo: formNo,
          comment: comment,
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
    if (item.comment.includes("待執行長覆核")) {
      sendCeoApprovalEmail(item.formNo, item.applyDate, item.applicantName, item.department, item.totalHours, item.details, "主管已審核");
    } else {
      sendManagerApprovalEmail(item.formNo, item.applyDate, item.applicantName, item.department, item.totalHours, item.details);
    }
  });

  ui.alert(`🎉 已完成補發作業！共補發 ${pendingKeys.length} 封通知。`);
}
