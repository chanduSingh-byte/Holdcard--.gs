/*************************************************
 * PROJECT : HOLD CARD MANAGEMENT SYSTEM
 * AUTHOR  : Optimized & documented
 * SHEETS
 * -----------------------------------------------
 * Quality       → Master Data
 * HoldCard      → Hold Records
 * ProcessEntry  → Process History
 *************************************************/


/* ==============================================
   CONFIG & CONSTANTS
============================================== */

// Email
const REMINDER_EMAIL = "itsupport@scherdel.in";

// Hold Number Prefix
const HOLD_PREFIX = "2025-26/JRC/";

// Process Flow (used for Next Process logic)
const PROCESS_FLOW = [
  "PR-Coiling(SETUP+Coiling+HT-1)",
  "PR-Centerless and H.T-(Straight)",
  "PR-Length Making",
  "RET-Cutting & Checking",
  "PR-Ring Forming O/S",
  "PR-Shop floor(RF Insp+RHT)",
  "QC-Inspection Plating",
  "Tension Checking",
  "PR-Paint(Paint+Rust Prev.+Shop to Insp)",
  "QC-FINAL INSPECTION"
];

/* ==============================================
   HOLD CARD COLUMN INDEX (1-BASED)
============================================== */

const COL_HOLD_NO          = 1;
const COL_CREATED_AT       = 2;
const COL_PROCESS          = 3;
const COL_JOBCARD          = 4;
const COL_PART             = 5;
const COL_DECISION         = 6;
const COL_DEFECT           = 7;
const COL_EXPECTED_ACTION  = 8;
const COL_DECISION_DT      = 9;
const COL_STATUS           = 10;


/* ==============================================
   COMMON RESPONSE (JSON + CORS SAFE)
============================================== */
function corsResponse_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}


/* ==============================================
   GET ROUTER
============================================== */
function doGet(e) {
  try {
    const action = e?.parameter?.action || "master";

    switch (action) {

      case "master":
        return corsResponse_(getProcessMaster());

      case "getHoldCards":
        return corsResponse_(getHoldCardNumbers());

      case "getNextProcess":
        return corsResponse_(getNextProcesses(e.parameter.holdNo));

      case "filterHold":
        return corsResponse_(getAllWaitHold());

      case "getDefectItems":
        return corsResponse_(getDefectItems());

      case "getOperators":
        return corsResponse_(getOperatorsFromMaster());

      case "qualityColumnB":
        return corsResponse_(getQualityColumnB());

      default:
        return corsResponse_([]);

    }
  } catch (err) {
    return corsResponse_({ status: "error", message: err.message });
  }
}


/* ==============================================
   POST ROUTER
============================================== */
function doPost(e) {
  try {
    if (!e?.postData?.contents) throw new Error("No POST data");

    const data = JSON.parse(e.postData.contents);

    // WAIT → Update existing hold
    if (data.mode === "WAIT" && data.hold === "Yes") {
      return corsResponse_(updateHoldDecision(data));
    }

    // NEW → Create new hold
    if (data.mode === "NEW") {
      return corsResponse_(saveHoldCard(data));
    }

    throw new Error("Invalid mode");

  } catch (err) {
    return corsResponse_({ status: "error", message: err.message });
  }
}


/* ==============================================
   MASTER DATA APIs
============================================== */

// Operators → Master!B
function getOperatorsFromMaster() {
  const sh = SpreadsheetApp.getActive().getSheetByName("Master");
  if (!sh || sh.getLastRow() < 2) return [];

  return sh.getRange(2, 2, sh.getLastRow() - 1, 1)
    .getValues()
    .flat()
    .filter(v => v && v.toString().trim());
}


// Defect List → Quality!L
function getDefectItems() {
  const sh = SpreadsheetApp.getActive().getSheetByName("Quality");
  if (!sh || sh.getLastRow() < 2) return [];

  return sh.getRange(2, 12, sh.getLastRow() - 1, 1)
    .getValues()
    .flat()
    .filter(String);
}


// Quality Column B (Third Page usage)
function getQualityColumnB() {
  const sh = SpreadsheetApp.getActive().getSheetByName("Quality");
  if (!sh || sh.getLastRow() < 2) return [];

  return sh.getRange(2, 2, sh.getLastRow() - 1, 1)
    .getValues()
    .flat()
    .filter(String);
}


/* ==============================================
   HOLD FETCH
============================================== */

// All WAIT + OPEN records
function getAllWaitHold() {
  const sh = SpreadsheetApp.getActive().getSheetByName("HoldCard");
  if (!sh || sh.getLastRow() < 2) return [];

  const data = sh.getRange(2, 1, sh.getLastRow() - 1, 10).getValues();

  return data
    .filter(r =>
      String(r[COL_DECISION - 1]).toUpperCase() === "WAIT" &&
      (!r[COL_STATUS - 1] || String(r[COL_STATUS - 1]).toUpperCase() === "OPEN")
    )
    .map(r => ({
      holdNo: r[0],
      process: r[2],
      jobCard: r[3],
      part: r[4],
      decision: r[5],
      defect: r[6],
      expectedDecision: r[7],
      decisionDate: r[8],
      status: r[9] || "OPEN"
    }));
}


// OPEN Hold Numbers only
function getHoldCardNumbers() {
  const sh = SpreadsheetApp.getActive().getSheetByName("HoldCard");
  if (!sh || sh.getLastRow() < 2) return [];

  return sh.getRange(2, 1, sh.getLastRow() - 1, 10)
    .getValues()
    .filter(r => String(r[9]).toUpperCase() === "OPEN")
    .map(r => r[0]);
}


/* ==============================================
   SAVE / UPDATE HOLD CARD
============================================== */

// Update existing hold (WAIT → YES)
function updateHoldDecision(data) {
  const sh = SpreadsheetApp.getActive().getSheetByName("HoldCard");
  const rows = sh.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.holdCardNumber) {

      sh.getRange(i + 1, COL_DECISION).setValue("Yes");
      sh.getRange(i + 1, COL_DEFECT).setValue(data.defectItems || "");
      sh.getRange(i + 1, COL_EXPECTED_ACTION).setValue(data.expectedDecision || "");
      sh.getRange(i + 1, COL_DECISION_DT).setValue(new Date());

      return { status: "success", holdNo: data.holdCardNumber };
    }
  }
  throw new Error("Hold Card not found");
}


// Create new Hold Card
function saveHoldCard(data) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName("HoldCard");

  if (!sh) {
    sh = ss.insertSheet("HoldCard");
    sh.appendRow([
      "Hold Number", "Created At", "Process",
      "Job Card", "Part", "Decision",
      "Defect", "Expected Action",
      "Decision Date", "Status"
    ]);
  }

  // Prevent duplicate WAIT + OPEN
  if (isDuplicateHold_(sh, data.process, data.part, data.jobCard)) {
    return { status: "error", message: "Duplicate OPEN WAIT Hold exists" };
  }

  const holdNo = generateHoldNo(sh);

  sh.appendRow([
    holdNo,
    new Date(),
    data.process,
    data.jobCard,
    data.part,
    data.hold,
    data.defectItems || "",
    data.expectedDecision || "",
    data.dateTime ? new Date(data.dateTime) : "",
    "OPEN"
  ]);

  return { status: "success", holdNo };
}


// Duplicate checker
function isDuplicateHold_(sh, process, part, jobCard) {
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, 10).getValues();

  return data.some(r =>
    r[2] === process &&
    r[3] === jobCard &&
    r[4] === part &&
    String(r[5]).toUpperCase() === "WAIT" &&
    String(r[9]).toUpperCase() === "OPEN"
  );
}


// Generate sequential Hold No
function generateHoldNo(sh) {
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return HOLD_PREFIX + "00001";

  const lastNo = sh.getRange(lastRow, 1).getValue();
  const seq = Number(lastNo.split("/").pop()) + 1;
  return HOLD_PREFIX + seq.toString().padStart(5, "0");
}


/* ==============================================
   NEXT PROCESS
============================================== */
function getNextProcesses(holdNo) {
  const sh = SpreadsheetApp.getActive().getSheetByName("HoldCard");
  if (!sh) return [];

  const data = sh.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === holdNo) {
      const idx = PROCESS_FLOW.indexOf(String(data[i][2]).trim());
      return idx >= 0 ? PROCESS_FLOW.slice(idx + 1) : [];
    }
  }
  return [];
}




























// ReminderCode


/*************************************************
 * AUTO HOLD ITEMS EMAIL REPORT
 * Sheet  : HoldCard
 * Status : OPEN
 * Time   : Daily at 4:30 PM
 *************************************************/


// ⏰ Convert Date → 12 Hour format with AM/PM
function format12Hour(date) {
  if (!(date instanceof Date)) return "-";
  return Utilities.formatDate(
    date,
    Session.getScriptTimeZone(),
    "dd/MM/yyyy hh:mm a"
  );
}


// ⌛ Calculate duration from given date → now
function getDuration(fromDate) {
  if (!(fromDate instanceof Date)) return "-";

  const now = new Date();
  const diffMs = now - fromDate;

  const mins = Math.floor(diffMs / 60000);
  const hrs = Math.floor(mins / 60);
  const days = Math.floor(hrs / 24);

  if (days > 0) return `${days} Day(s) ${hrs % 24} Hr`;
  if (hrs > 0) return `${hrs} Hr ${mins % 60} Min`;
  return `${mins} Min`;
}


// 📧 Send HOLD / OPEN data mail
function sendHoldDataMail() {

  const SHEET_NAME = "HoldCard";
  const STATUS_COL = 6; // Column I (Status)
  const TO = "itsupport@scherdel.in";
  const SUBJECT = "Hold Numbers Open For Decision As On ";

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet || sheet.getLastRow() < 2) {
    MailApp.sendEmail(TO, SUBJECT, "No records found in HoldCard sheet.");
    return;
  }

  const data = sheet.getDataRange().getValues();
  const headers = [...data[0], "Hold Duration"];
  const rows = data.slice(1);

  // 🔍 Filter OPEN rows
  const holdRows = rows.filter(r =>
    String(r[STATUS_COL - 1]).toUpperCase() === "WAIT"
  );

  if (holdRows.length === 0) {
    MailApp.sendEmail(TO, SUBJECT, "No OPEN / HOLD records found.");
    return;
  }

  let html = `
    <p>Hello Team,</p>
    <p>Below is the list of <b>OPEN / HOLD</b> items:</p>

    <table border="1" cellpadding="6" cellspacing="0"
      style="border-collapse:collapse;font-family:Arial;font-size:12px">
      <thead style="background:#f2f2f2;font-weight:bold">
        <tr>${headers.map(h => `<th>${h}</th>`).join("")}</tr>
      </thead>
      <tbody>
  `;

  holdRows.forEach(row => {
    html += "<tr>";

    row.forEach((cell, colIndex) => {

      // 🕒 Column B (Timestamp)
      if (colIndex === 1) {
        html += `<td>${format12Hour(cell)}</td>`;
      }
      // 🕒 Column G (Hold DateTime)
      else if (colIndex === 6) {
        html += `<td>${format12Hour(cell)}</td>`;
      }
      else {
        html += `<td>${cell === "" ? "-" : cell}</td>`;
      }

    });

    // ⌛ Hold Duration (from Column G)
    html += `<td>${getDuration(row[6])}</td>`;
    html += "</tr>";
  });

  html += `
      </tbody>
    </table>
    <br>
    <p>Regards,<br><b>IT Support</b></p>
  `;

  MailApp.sendEmail({
    to: TO,
    subject: SUBJECT,
    htmlBody: html
  });
}


// ⏰ Create daily trigger at 4:30 PM
function create430PMTrigger() {

  // Delete old triggers
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "sendHoldDataMail") {
      ScriptApp.deleteTrigger(t);
    }
  });

  // Create new trigger
  ScriptApp.newTrigger("sendHoldDataMail")
    .timeBased()
    .everyDays(1)
    .atHour(13)      // 4 PM
    .nearMinute(10)  // 4:30 PM
    .create();
}

