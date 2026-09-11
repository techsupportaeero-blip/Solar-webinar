/**
 * =========================================================================
 * Solar Career Webinar Registration
 * Google Apps Script + Google Sheets + Email + WhatsApp
 * =========================================================================
 */

// ========================================================================
// 1. WHATSAPP API SETTINGS - ImTelMsgHub
// ========================================================================

const WHATSAPP_API_KEY = "c6a41076-70d9-406d-9824-e63a4003137b";

const WHATSAPP_API_URL = "https://panel.omtelmsghub.com/api/v1/whatsapp/single";

// Connected ImTelMsgHub WhatsApp number
const WHATSAPP_SENDER_NO = "+919310413724";

// Solar Webinar WhatsApp Template ID
const WHATSAPP_TEMPLATE_ID = "753563377798566";

// ========================================================================
// 2. EMAIL SETTINGS
// ========================================================================

const RESEND_API_KEY = "";

const RESEND_FROM_EMAIL = "Webinar Team <onboarding@resend.dev>";

// ========================================================================
// 3. GOOGLE SHEET HEADERS
// ========================================================================

const HEADERS = [
  "Timestamp",
  "Full Name",
  "Mobile Number",
  "Email Address",
  "Qualification",
  "Interest Area",
];

// ========================================================================
// 4. FORM SUBMISSION HANDLER
// ========================================================================

function doPost(e) {
  const lock = LockService.getScriptLock();

  lock.tryLock(10000);

  try {
    const doc = SpreadsheetApp.getActiveSpreadsheet();

    const sheet = doc.getActiveSheet();

    // --------------------------------------------------------------------
    // CREATE HEADERS IF SHEET IS EMPTY
    // --------------------------------------------------------------------

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);

      formatHeader(sheet);
    }

    // --------------------------------------------------------------------
    // GET FORM DATA
    // --------------------------------------------------------------------

    let data = {};

    if (e && e.postData && e.postData.contents) {
      try {
        data = JSON.parse(e.postData.contents);
      } catch (err) {
        data = e.parameter || {};
      }
    } else if (e && e.parameter) {
      data = e.parameter;
    }

    // --------------------------------------------------------------------
    // FORMAT DATE
    // --------------------------------------------------------------------

    const formattedDate = Utilities.formatDate(
      new Date(),
      "Asia/Kolkata",
      "dd-MM-yyyy hh:mm:ss a",
    );

    // --------------------------------------------------------------------
    // FORMAT MOBILE NUMBER
    // --------------------------------------------------------------------

    const rawMobile = data.Mobile ? String(data.Mobile).trim() : "";

    const formattedMobile = rawMobile ? "'" + rawMobile : "";

    // --------------------------------------------------------------------
    // CREATE SHEET ROW
    // --------------------------------------------------------------------

    const row = [
      formattedDate,

      data.Name ? String(data.Name).trim() : "",

      formattedMobile,

      data.Email ? String(data.Email).trim() : "",

      data.Qualification ? String(data.Qualification).trim() : "",

      data.Interest ? String(data.Interest).trim() : "",
    ];

    // --------------------------------------------------------------------
    // SAVE LEAD TO GOOGLE SHEET
    // --------------------------------------------------------------------

    sheet.appendRow(row);

    const newRowIndex = sheet.getLastRow();

    formatSheetRow(sheet, newRowIndex);

    // ====================================================================
    // SEND EMAIL
    // ====================================================================

    if (data.Email && String(data.Email).includes("@")) {
      sendApplicantEmail(data);
    }

    // ====================================================================
    // SEND WHATSAPP
    // ====================================================================

    if (data.Mobile) {
      sendWhatsAppMessage(data.Mobile);
    }

    // ====================================================================
    // SUCCESS RESPONSE
    // ====================================================================

    return ContentService.createTextOutput(
      JSON.stringify({
        result: "success",
        row: newRowIndex,
      }),
    ).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    console.error("Submission Error: " + error.toString());

    return ContentService.createTextOutput(
      JSON.stringify({
        result: "error",
        message: error.toString(),
      }),
    ).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// ========================================================================
// 5. WHATSAPP MESSAGE FUNCTION
// ========================================================================

function sendWhatsAppMessage(mobileNumber) {
  try {
    // --------------------------------------------------------------------
    // CHECK API KEY
    // --------------------------------------------------------------------

    if (!WHATSAPP_API_KEY) {
      console.error("WhatsApp Error: API Key is missing.");

      return;
    }

    // --------------------------------------------------------------------
    // CHECK SENDER NUMBER
    // --------------------------------------------------------------------

    if (!WHATSAPP_SENDER_NO) {
      console.error("WhatsApp Error: Sender number is missing.");

      return;
    }

    // --------------------------------------------------------------------
    // CHECK TEMPLATE ID
    // --------------------------------------------------------------------

    if (!WHATSAPP_TEMPLATE_ID) {
      console.error("WhatsApp Error: Template ID is missing.");

      return;
    }

    // --------------------------------------------------------------------
    // CLEAN CUSTOMER MOBILE NUMBER
    // --------------------------------------------------------------------

    const cleanNumber = String(mobileNumber).replace(/\D/g, "").slice(-10);

    // --------------------------------------------------------------------
    // VALIDATE NUMBER
    // --------------------------------------------------------------------

    if (cleanNumber.length !== 10) {
      console.error("WhatsApp Error: Invalid mobile number - " + mobileNumber);

      return;
    }

    // --------------------------------------------------------------------
    // CUSTOMER WHATSAPP NUMBER
    // --------------------------------------------------------------------

    const finalNumber = "+91" + cleanNumber;

    // --------------------------------------------------------------------
    // API REQUEST BODY
    //
    // This matches the ImTelMsgHub /single API schema:
    // message_type ("template" — required because we are sending
    //   a pre-approved template, not a free-form text message)
    // sender
    // to
    // template_id
    // --------------------------------------------------------------------

    const payload = {
      message_type: "template",

      sender: WHATSAPP_SENDER_NO,

      to: finalNumber,

      template_id: WHATSAPP_TEMPLATE_ID,
    };

    // --------------------------------------------------------------------
    // API REQUEST OPTIONS
    // --------------------------------------------------------------------

    const options = {
      method: "post",

      contentType: "application/json",

      payload: JSON.stringify(payload),

      muteHttpExceptions: true,
    };

    // --------------------------------------------------------------------
    // API URL WITH API KEY
    // --------------------------------------------------------------------

    const url =
      WHATSAPP_API_URL + "?api_key=" + encodeURIComponent(WHATSAPP_API_KEY);

    // --------------------------------------------------------------------
    // SEND WHATSAPP API REQUEST
    // --------------------------------------------------------------------

    const response = UrlFetchApp.fetch(url, options);

    // --------------------------------------------------------------------
    // GET RESPONSE
    // --------------------------------------------------------------------

    const responseCode = response.getResponseCode();

    const responseText = response.getContentText();

    // --------------------------------------------------------------------
    // LOG RESPONSE FOR DEBUGGING
    // --------------------------------------------------------------------

    console.log("====================================");

    console.log("WhatsApp Status: " + responseCode);

    console.log("WhatsApp Response: " + responseText);

    console.log("WhatsApp To: " + finalNumber);

    console.log("WhatsApp Sender: " + WHATSAPP_SENDER_NO);

    console.log("WhatsApp Template: " + WHATSAPP_TEMPLATE_ID);

    console.log("====================================");

    // --------------------------------------------------------------------
    // ERROR LOG
    // --------------------------------------------------------------------

    if (responseCode < 200 || responseCode >= 300) {
      console.error("WhatsApp API FAILED: " + responseText);
    } else {
      console.log("WhatsApp message request sent successfully.");
    }
  } catch (error) {
    console.error("WhatsApp Error: " + error.toString());
  }
}

// ========================================================================
// 6. SHEET HEADER FORMATTING
// ========================================================================

function formatHeader(sheet) {
  const headerRange = sheet.getRange(1, 1, 1, HEADERS.length);

  headerRange
    .setFontFamily("Segoe UI")
    .setFontSize(10)
    .setFontWeight("bold")
    .setBackground("#991B1B")
    .setFontColor("#FFFFFF")
    .setVerticalAlignment("middle")
    .setHorizontalAlignment("center");

  sheet.setRowHeight(1, 38);

  sheet.setFrozenRows(1);

  sheet.setColumnWidth(1, 175);

  sheet.setColumnWidth(2, 180);

  sheet.setColumnWidth(3, 135);

  sheet.setColumnWidth(4, 230);

  sheet.setColumnWidth(5, 160);

  sheet.setColumnWidth(6, 210);
}

// ========================================================================
// 7. SHEET ROW FORMATTING
// ========================================================================

function formatSheetRow(sheet, rowIndex) {
  const rowRange = sheet.getRange(rowIndex, 1, 1, HEADERS.length);

  rowRange
    .setFontFamily("Segoe UI")
    .setFontSize(10)
    .setVerticalAlignment("middle");

  sheet.setRowHeight(rowIndex, 32);

  // --------------------------------------------------------------------
  // ZEBRA STRIPING
  // --------------------------------------------------------------------

  if (rowIndex % 2 === 0) {
    rowRange.setBackground("#F9FAFB");
  } else {
    rowRange.setBackground("#FFFFFF");
  }

  // --------------------------------------------------------------------
  // BORDERS
  // --------------------------------------------------------------------

  rowRange.setBorder(
    true,
    true,
    true,
    true,
    true,
    true,
    "#E5E7EB",
    SpreadsheetApp.BorderStyle.SOLID,
  );

  // --------------------------------------------------------------------
  // TIMESTAMP
  // --------------------------------------------------------------------

  sheet
    .getRange(rowIndex, 1)
    .setHorizontalAlignment("center")
    .setFontColor("#6B7280");

  // --------------------------------------------------------------------
  // NAME
  // --------------------------------------------------------------------

  sheet
    .getRange(rowIndex, 2)
    .setHorizontalAlignment("left")
    .setFontWeight("bold")
    .setFontColor("#111827");

  // --------------------------------------------------------------------
  // MOBILE
  // --------------------------------------------------------------------

  sheet
    .getRange(rowIndex, 3)
    .setHorizontalAlignment("center")
    .setNumberFormat("@")
    .setFontColor("#1F2937");

  // --------------------------------------------------------------------
  // EMAIL
  // --------------------------------------------------------------------

  sheet
    .getRange(rowIndex, 4)
    .setHorizontalAlignment("left")
    .setFontColor("#1D4ED8");

  // --------------------------------------------------------------------
  // QUALIFICATION
  // --------------------------------------------------------------------

  sheet
    .getRange(rowIndex, 5)
    .setHorizontalAlignment("center")
    .setFontColor("#374151");

  // --------------------------------------------------------------------
  // INTEREST
  // --------------------------------------------------------------------

  sheet
    .getRange(rowIndex, 6)
    .setHorizontalAlignment("center")
    .setFontWeight("bold")
    .setFontColor("#991B1B");
}

// ========================================================================
// 8. FORMAT ENTIRE SHEET
// ========================================================================

function formatEntireSheet() {
  const doc = SpreadsheetApp.getActiveSpreadsheet();

  const sheet = doc.getActiveSheet();

  const lastRow = sheet.getLastRow();

  if (lastRow === 0) {
    return;
  }

  formatHeader(sheet);

  for (let r = 2; r <= lastRow; r++) {
    formatSheetRow(sheet, r);
  }

  SpreadsheetApp.flush();
}

// ========================================================================
// 9. SMART EMAIL DISPATCHER
// ========================================================================

function sendSmartEmail(toEmail, subject, htmlBody) {
  let resendSuccess = false;

  // --------------------------------------------------------------------
  // RESEND API
  // --------------------------------------------------------------------

  if (RESEND_API_KEY && RESEND_API_KEY.trim() !== "") {
    try {
      const payload = {
        from: RESEND_FROM_EMAIL,

        to: [toEmail],

        subject: subject,

        html: htmlBody,
      };

      const options = {
        method: "post",

        contentType: "application/json",

        headers: {
          Authorization: "Bearer " + RESEND_API_KEY.trim(),
        },

        payload: JSON.stringify(payload),

        muteHttpExceptions: true,
      };

      const response = UrlFetchApp.fetch(
        "https://api.resend.com/emails",
        options,
      );

      if (
        response.getResponseCode() === 200 ||
        response.getResponseCode() === 201
      ) {
        resendSuccess = true;
      }
    } catch (error) {
      console.error("Resend API Error: " + error.toString());
    }
  }

  // --------------------------------------------------------------------
  // GMAIL FALLBACK
  // --------------------------------------------------------------------

  if (!resendSuccess) {
    try {
      if (MailApp.getRemainingDailyQuota() > 0) {
        MailApp.sendEmail({
          to: toEmail,

          subject: subject,

          htmlBody: htmlBody,
        });
      }
    } catch (error) {
      console.error("MailApp Error: " + error.toString());
    }
  }
}

// ========================================================================
// 10. WEBINAR CONFIRMATION EMAIL
// ========================================================================

function sendApplicantEmail(data) {
  const subject = "Seat Confirmed! 🚀 FREE Solar Career Webinar";

  const htmlBody = `

<!DOCTYPE html>

<html>

<body
style="
margin:0;
padding:20px 10px;
background-color:#f1f5f9;
font-family:'Segoe UI',sans-serif;
"
>

<div
style="
max-width:580px;
margin:0 auto;
background:#ffffff;
border:1px solid #e2e8f0;
border-radius:18px;
overflow:hidden;
"
>

<div
style="
background:#b91c1c;
padding:24px 20px;
text-align:center;
color:#ffffff;
"
>

<h2
style="
margin:0;
font-size:22px;
font-weight:bold;
"
>
Solar Career Webinar
</h2>

</div>
<div
style="
padding:24px 20px;
color:#1f2937;
"
>

<h3 style="margin-top:0;">

Hello
${data.Name || "Future Solar Expert"},

</h3>


<p>

Your seat for the
<strong>
FREE Solar Career Webinar
</strong>
has been successfully booked.

</p>


<div
style="
background:#f8fafc;
padding:16px;
border-radius:14px;
border:1px solid #e2e8f0;
"
>

<p>
📅
<strong>Date:</strong>
14th September, Monday
</p>


<p>
⏰
<strong>Time:</strong>
2:00 PM - 4:00 PM
</p>


<p>
📍
<strong>Venue:</strong>
Live on GOOGLE MEET
</p>


<p>
🌐
<strong>Language:</strong>
Hindi + English
</p>

</div>


<p style="margin-top:16px;">

The webinar joining link will be
shared with you on WhatsApp and
Email shortly before the session.

</p>


<div
style="
background:#eff6ff;
border-left:4px solid #3b82f6;
border-radius:12px;
padding:16px;
color:#1e3a8a;
font-size:13px;
line-height:1.6;
"
>

<strong>
💡 Pro Tip:
</strong>

Please join 5 minutes early
to secure your spot and grab
a notepad!

</div>


</div>

</div>

</body>

</html>

`;

  sendSmartEmail(data.Email, subject, htmlBody);
}

// ========================================================================
// 11. WEB APP STATUS
// ========================================================================

function doGet(e) {
  return ContentService.createTextOutput(
    "✅ Solar Webinar Form + Email + WhatsApp is Active.",
  );
}
