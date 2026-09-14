// Google Apps Script — Web App backend for the Solar Career Webinar registration form.
// Deploy: Extensions > Apps Script > paste this file as Code.gs > Deploy > Web App
//   Execute as: Me
//   Who has access: Anyone

const SHEET_NAME = "Leads"; // change if your sheet tab has a different name

const WHATSAPP_API_KEY = "c6a41076-70d9-406d-9824-e63a4003137b";
const WHATSAPP_API_URL = "https://panel.omtelmsghub.com/api/v1/whatsapp/single";
const WHATSAPP_SENDER_NO = "+919310413724";
const WHATSAPP_TEMPLATE_ID = "753563377798566";

function doPost(e) {
  try {
    if (!e || !e.parameter) {
      throw new Error("No form data received.");
    }

    const name = e.parameter.Name || "";
    const mobile = e.parameter.Mobile || "";
    const email = e.parameter.Email || "";
    const qualification = e.parameter.Qualification || "";
    const interest = e.parameter.Interest || "";

    saveToSheet(name, mobile, email, qualification, interest);
    sendConfirmationEmail(name, email);
    sendWhatsAppConfirmation(mobile, name);

    return ContentService
      .createTextOutput(JSON.stringify({ result: "success" }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    console.error("doPost Error: " + err.message);
    return ContentService
      .createTextOutput(JSON.stringify({ result: "error", message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function saveToSheet(name, mobile, email, qualification, interest) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME) || ss.getActiveSheet();
  sheet.appendRow([new Date(), name, mobile, email, qualification, interest]);
}

function sendConfirmationEmail(name, email) {
  if (!email) {
    console.error("Email Error: No email received.");
    return;
  }
  try {
    MailApp.sendEmail({
      to: email,
      subject: "You're Registered! Free Solar Career Webinar",
      body:
        "Hi " + name + ",\n\n" +
        "Thank you for registering for the Free Solar Career Webinar. " +
        "We will send you the joining link shortly on email and WhatsApp.\n\n" +
        "See you there!\nTeam AEERO Academy"
    });
  } catch (err) {
    console.error("Email Error: " + err.message);
  }
}

function sendWhatsAppConfirmation(mobileNumber, name) {
  if (!mobileNumber) {
    console.error("WhatsApp Error: Mobile number not received.");
    return;
  }

  const cleanNumber = String(mobileNumber).replace(/\D/g, "").slice(-10);

  if (cleanNumber.length !== 10) {
    console.error("WhatsApp Error: Invalid mobile number - " + mobileNumber);
    return;
  }

  const finalNumber = "+91" + cleanNumber;

  const payload = {
    message_type: "text",
    sender: WHATSAPP_SENDER_NO,
    to: finalNumber,
    template_id: WHATSAPP_TEMPLATE_ID
  };

  const url = WHATSAPP_API_URL + "?api_key=" + encodeURIComponent(WHATSAPP_API_KEY);

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    console.log("WhatsApp Response: " + response.getContentText());
  } catch (err) {
    console.error("WhatsApp Error: " + err.message);
  }
}
