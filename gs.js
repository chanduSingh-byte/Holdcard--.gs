/*************************************************
 * SHEETS USED
 * Quality       → Master Data
 * HoldCard      → Hold Records
 * ProcessEntry  → Process Entry Records
 *************************************************/

/* ==============================================
   CONFIG
============================================== */
const REMINDER_EMAIL = "itsupport@scherdel.in";
// const HOLD_PREFIX = "2025-26/JRC/";
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
   CORS RESPONSE
============================================== */
function corsResponse_(obj) {
  return ContentService.createTextOutput(
    JSON.stringify(obj)
  ).setMimeType(ContentService.MimeType.JSON);
}




/* ==============================================
   doGet ROUTER
============================================== */
function doGet(e) {
  console.log(e)
  try {
    var action = "master";
    if (e && e.parameter && e.parameter.action) {
      action = e.parameter.action;
    }

    if (action === "waitHoldAll") {
      return corsResponse_(getAllWaitHold());
    }

    switch (action) {
      case "master":
        return corsResponse_(getProcessMaster());

      case "getHoldCards":
        return corsResponse_(getHoldCardNumbers());

      case "getNextProcess":
        return corsResponse_(getNextProcesses(e.parameter.holdNo));

      case "filterHold":
        return corsResponse_(
          getHoldCardsByStatus(e.parameter.value));

      case "waitHoldSingle":
        return corsResponse_(getLatestWaitHold());

      case "getQualityColF":   // 👈 NEW
        return corsResponse_(ThirdPage());

      case "getDefectItems":
        return corsResponse_(getDefectItems());

      case "getOperators":
        return corsResponse_(getOperatorsFromMaster());



      default:
        return corsResponse_([]);
    }

  } catch (err) {
    return corsResponse_({ status: "error", message: err.message });
  }
}



const COL_HOLD_NO = 1;
const COL_TIMESTAMP = 2;
const COL_PROCESS = 3;
const COL_JOBCARD = 4;
const COL_PART = 5;
const COL_HOLD_REASON = 6;
const COL_HOLD_DATETIME = 7;
const COL_EXPECTED_DECISION = 8;
const COL_DEFECT = 9;
const COL_STATUS = 10;   // ✅ FIX
const COL_REMINDER = 11;





// ===== OPERATOR MASTER API =====
// ===== MASTER SHEET COLUMN B =====
function getOperatorsFromMaster() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("Master");

  if (!sh) {
    console.log("❌ Master sheet not found");
    return [];
  }

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  const data = sh.getRange(2, 2, lastRow - 1, 1)
    .getValues()
    .flat()
    .filter(v => v && v.toString().trim() !== "");

  console.log("✅ Operators:", data);
  return data;
}




function getAllWaitHold() {
  const sh = SpreadsheetApp.getActive().getSheetByName("HoldCard");
  if (!sh || sh.getLastRow() < 2) return [];

  // 🔥 Read FULL row (11 columns)
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, 11).getValues();

  const waits = data
    .filter(r => {
      const decision = String(r[COL_HOLD_REASON - 1]).trim().toUpperCase();
      const status = String(r[COL_STATUS - 1]).trim().toUpperCase();

      // ✅ FINAL RULE
      return decision === "WAIT" && (status === "" || status === "OPEN");
    })
    .map(r => ({
      holdNo: r[COL_HOLD_NO - 1],
      process: r[COL_PROCESS - 1],
      jobCard: r[COL_JOBCARD - 1],
      part: r[COL_PART - 1],
      decision: r[COL_HOLD_REASON - 1],
      dateTime: r[COL_HOLD_DATETIME - 1],
      userdateTime: r[COL_DEFECT - 1],
      expectedDecision: r[COL_EXPECTED_DECISION - 1],
      status: r[COL_STATUS - 1] || "OPEN",
    })
    );
  console.log(waits)

  return waits;
}




// function getAllWaitHold() {
//   const sh = SpreadsheetApp.getActive().getSheetByName("HoldCard");
//   if (!sh || sh.getLastRow() < 2) return [];

//   const data = sh.getRange(2, 1, sh.getLastRow() - 1, 10).getValues();

//   // 🔑 WAIT + OPEN only
//   const waits = data
//     // .filter(r =>
//     //   String(r[5]).trim().toUpperCase() === "WAIT" &&
//     //   String(r[8]).trim().toUpperCase() === "OPEN"
//     // )
//     .filter(r => {
//       const decision = String(r[5]).trim().toUpperCase();
//       const status = String(r[8]).trim().toUpperCase();

//       return decision.includes("WAIT") &&
//         status.includes("OPEN");
//     })


//    return.map(r => ({
//       holdNo: r[0],
//       process: r[2],
//       jobCard: r[3],
//       part: r[4],
//       decision: r[5],   // WAIT
//       dateTime: r[6],
//       expectedDecision: r[7],
//       status: r[8]
//     }));

//   // console.log(waits)
//   // return waits; // ✅ ALL records
// }


// ===== NEW : DEFECT MASTER API =====
// ===== DEFECT REASON FROM QUALITY COLUMN L =====
function getDefectItems() {
  const sh = SpreadsheetApp.getActive()
    .getSheetByName("Quality");

  if (!sh || sh.getLastRow() < 2) return [];

  return sh
    .getRange(2, 12, sh.getLastRow() - 1, 1) // ⭐ Column L
    .getValues()
    .flat()
    .filter(String);
}



// third page code 
function ThirdPage() {
  const sh = SpreadsheetApp.getActive()
    .getSheetByName("Quality");

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  console.log(lastRow)
  const data = sh
    .getRange(2, 2, lastRow - 1, 1) // Column B
    .getValues()
    .flat()
    .filter(String);

  // console.log(data)
  return data
}

// third page code 

/* ==============================================
   doPost ROUTER
============================================== */
// function doPost(e) {
//   try {
//     if (!e || !e.postData || !e.postData.contents) {
//       throw new Error("No POST data received");
//     }

//     const data = JSON.parse(e.postData.contents);

//     const result = data.holdCardNumber
//       ? saveProcessEntry(data)
//       : saveHoldCard(data);

//     return corsResponse_(result);

//   } catch (err) {
//     return corsResponse_({ status: "error", message: err.message });
//   }
// }


// function doPost(e) {
//   try {
//     if (!e || !e.postData || !e.postData.contents) {
//       throw new Error("No POST data received");
//     }

//     const data = JSON.parse(e.postData.contents);

//     let result;

//     // ✅ WAIT MODE → YES (UPDATE SAME HOLD)
//     if (data.mode === "WAIT" && data.hold === "Yes") {
//       result = updateHoldToYes(data);
//     }
//     // ✅ NEW MODE → CREATE NEW HOLD
//     else {
//       result = saveHoldCard(data);
//     }

//     return corsResponse_(result);

//   } catch (err) {
//     return corsResponse_({
//       status: "error",
//       message: err.message
//     });
//   }
// }


function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error("No POST data received");
    }

    const data = JSON.parse(e.postData.contents);

    let result;

    // 🔹 WAIT MODE → UPDATE EXISTING HOLD
    if (data.mode === "WAIT" && data.hold === "Yes") {

      result = updateHoldDecision(data);

      // 🔹 NEW MODE → CREATE NEW HOLD
    } else if (data.mode === "NEW") {
      result = saveHoldCard(data);

    } else {
      throw new Error("Invalid request");
    }

    return corsResponse_(result);

  } catch (err) {
    return corsResponse_({
      status: "error",
      message: err.message
    });
  }
}
function updateHoldDecision(data) {
  const sh = SpreadsheetApp.getActive().getSheetByName("HoldCard");
  const rows = sh.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.holdCardNumber) { // Hold No column
      sh.getRange(i + 1, 6).setValue(data.hold);              // Decision
      sh.getRange(i + 1, 7).setValue(data.defectItems);       // Defect
      sh.getRange(i + 1, 8).setValue(data.expectedDecision);  // Expected Decision
      sh.getRange(i + 1, 9).setValue(data.dateTime);          // DateTime

      return {
        status: "success",
        holdNo: data.holdCardNumber
      };
    }
  }

  throw new Error("Hold Card not found");
}




function updateHoldToYes(data) {
  const sh = SpreadsheetApp.getActive()
    .getSheetByName("HoldCard");

  if (!sh) throw new Error("HoldCard sheet not found");

  const dataRange = sh.getDataRange().getValues();

  for (let i = 1; i < dataRange.length; i++) {
    if (dataRange[i][0] === data.holdNo) {

      // Update Hold Reason
      sh.getRange(i + 1, COL_HOLD_REASON).setValue("Yes");

      // Update DateTime
      sh.getRange(i + 1, COL_HOLD_DATETIME)
        .setValue(data.dateTime ? new Date(data.dateTime) : new Date());

      // Update Expected Decision
      sh.getRange(i + 1, COL_EXPECTED_DECISION)
        .setValue(data.expectedDecision || "");

      // Update Defect
      sh.getRange(i + 1, COL_DEFECT)
        .setValue(data.defectItems || "");

      return {
        status: "success",
        holdNo: data.holdNo
      };
    }
  }

  throw new Error("Hold Number not found");
}



/* ==============================================
   MASTER DATA
============================================== */
function getProcessMaster() {
  const sh = SpreadsheetApp.getActive().getSheetByName("Quality");
  if (!sh || sh.getLastRow() < 2) return {};

  const data = sh.getRange(2, 1, sh.getLastRow() - 1, 11).getValues();
  const map = {};

  data.forEach(function (r) {
    const jobCard = String(r[1]).trim();
    const part = String(r[5]).trim();
    const process = String(r[10]).trim();
    if (!process || !part || !jobCard) return;

    if (!map[process]) map[process] = {};
    if (!map[process][part]) map[process][part] = [];

    if (map[process][part].indexOf(jobCard) === -1) {
      map[process][part].push(jobCard);
    }
  });

  console.log(map)
  return map;
}



// function getProcessMaster() {

//   const ss = SpreadsheetApp.getActive();

//   /* ---------- HOLD CARD : YES LIST ---------- */
//   const holdSh = ss.getSheetByName("HoldCard");
//   const holdYesSet = new Set();

//   if (holdSh && holdSh.getLastRow() > 1) {
//     const holdData = holdSh.getRange(2, 1, holdSh.getLastRow() - 1, 10).getValues();
//     // A = Hold Number
//     // D = Job Card No
//     // F = Decision  (Yes / Wait)

//     holdData.forEach(r => {
//       const jobCard = String(r[3]).trim(); // column D
//       const decision = String(r[5]).trim(); // column F

//       if (decision === "Yes" && jobCard) {
//         holdYesSet.add(jobCard);
//       }
//     });
//   }

//   /* ---------- QUALITY MASTER ---------- */
//   const sh = ss.getSheetByName("Quality");
//   if (!sh || sh.getLastRow() < 2) return {};

//   const data = sh.getRange(2, 1, sh.getLastRow() - 1, 11).getValues();
//   const map = {};

//   data.forEach(r => {
//     const jobCard = String(r[1]).trim();
//     const part = String(r[5]).trim();
//     const process = String(r[10]).trim();

//     // 🔴 NEW CONDITION → HoldCard me Yes hai to skip
//     if (holdYesSet.has(jobCard)) return;

//     if (!process || !part || !jobCard) return;

//     if (!map[process]) map[process] = {};
//     if (!map[process][part]) map[process][part] = [];

//     if (!map[process][part].includes(jobCard)) {
//       map[process][part].push(jobCard);
//     }
//   });

//   console.log(holdYesSet);
//   return map;
// }


/* ==============================================
   SAVE HOLD CARD
============================================== */

const HOLD_PREFIX = "2025-26/JRC/";

// Duplicate Entries Lock
function isDuplicateHold_(sh, process, part, jobCard) {
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return false;

  // ✅ READ ALL 11 COLUMNS
  const data = sh.getRange(2, 1, lastRow - 1, 11).getValues();

  return data.some(r =>
    String(r[COL_PROCESS - 1]).trim() === String(process).trim() &&
    String(r[COL_PART - 1]).trim() === String(part).trim() &&
    String(r[COL_JOBCARD - 1]).trim() === String(jobCard).trim() &&
    String(r[COL_HOLD_REASON - 1]).trim().toUpperCase() === "WAIT" &&
    String(r[COL_STATUS - 1]).trim().toUpperCase() === "OPEN"
  );
}



// function saveHoldCard(data) {
//   const ss = SpreadsheetApp.getActive();
//   let sh = ss.getSheetByName("HoldCard");

//   if (!sh) {
//     sh = ss.insertSheet("HoldCard");
//     sh.appendRow([
//       "Hold No", "Timestamp", "Process", "Job Card", "Part",
//       "Hold Reason", "Hold DateTime", "Expected Decision",
//       "Status", "ReminderSent"
//     ]);
//   }

//   const holdNo = generateHoldNo(sh);

//   sh.appendRow([
//     holdNo,
//     new Date(),
//     data.process,
//     data.jobCard,
//     data.part,
//     data.hold,
//     data.dateTime ? new Date(data.dateTime) : "",
//     data.expectedDecision,
//     "OPEN",
//     ""
//   ]);

//   return { status: "success", holdNo: holdNo };
// }


function saveHoldCard(data) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName("HoldCard");

  if (!sh) {
    sh = ss.insertSheet("HoldCard");
    sh.appendRow([
      "Hold Number",
      "HC_No-Generated At",
      "Process Name",
      "Job Card NO",
      "Part",
      "Decision",
      "Defect Item",
      "Expected Decision",
      "Decision-DateTime",
      "Status"
    ]);
  }

  // 🔒 DUPLICATE CHECK
  const duplicate = isDuplicateHold_(
    sh,
    data.process,
    data.part,
    data.jobCard
  );

  if (duplicate) {
    return {
      status: "error",
      message: "Hold already exists for same Process, Part & Job Card"
    };
  }

  // ✅ Generate Hold Number & Timestamp
  const holdNo = generateHoldNo(sh);
  const generatedAt = new Date();   // ✅ FIX
  const decisionDT = data.dateTime ? new Date(data.dateTime) : "";

  // ✅ EXACT SAME FORMAT AS ROW 2 & 3
  sh.appendRow([
    holdNo,                 // A Hold Number
    generatedAt,            // B HC_No-Generated At
    data.process,           // C Process Name
    data.jobCard,           // D Job Card NO
    data.part,              // E Part
    data.hold,              // F Decision (Yes / Wait)
    data.defectItems,       // G Defect Item
    data.expectedDecision,  // H Expected Decision
    decisionDT,             // I Decision-DateTime
    "OPEN"                  // J Status
  ]);

  return { status: "success", holdNo: holdNo };
}




// function saveHoldCard(data) {
//   const ss = SpreadsheetApp.getActive();
//   let sh = ss.getSheetByName("HoldCard");

//   if (!sh) {
//     sh = ss.insertSheet("HoldCard");
//     sh.appendRow([
//       "Hold No", "Timestamp", "Process", "Job Card", "Part",
//       "Hold Reason", "Hold DateTime", "Expected Decision",
//       "Status", "ReminderSent"
//     ]);
//   }

//   // 🔒 DUPLICATE CHECK
//   const duplicate = isDuplicateHold_(
//     sh,
//     data.process,
//     data.part,
//     data.jobCard
//   );

//   if (duplicate) {
//     return {
//       status: "error",
//       message: "Hold already exists for same Process, Part & Job Card"
//     };
//   }

//   // ✅ Generate Hold Number
//   const holdNo = generateHoldNo(sh);

//   // sh.appendRow([
//   //   holdNo,
//   //   new Date(),
//   //   data.process,
//   //   data.jobCard,
//   //   data.part,
//   //   data.hold,
//   //   data.dateTime ? new Date(data.dateTime) : "",
//   //   data.expectedDecision,
//   //   data.defectItems || "",
//   //   "OPEN",
//   //   ""
//   // ]);


//   sh.appendRow([
//   holdNo,                 // A Hold Number
//   generatedAt,            // B HC_No-Generated At
//   data.process,           // C Process Name
//   data.jobCard,           // D Job Card NO
//   data.part,              // E Part
//   data.hold,              // F Decision (Yes / Wait)
//   data.defectItems,       // G Defect Item
//   data.expectedDecision,  // H Expected Decision
//   data.dateTime,          // I Decision-DateTime
//   "OPEN"                  // J Status
// ]);


//   return { status: "success", holdNo: holdNo };
// }


/* ==============================================
   HOLD NUMBER GENERATOR
============================================== */
function generateHoldNo(sh) {
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return HOLD_PREFIX + "00001";

  const lastNo = sh.getRange(lastRow, 1).getValue();
  const seq = Number(lastNo.split("/").pop()) + 1;
  return HOLD_PREFIX + seq.toString().padStart(5, "0");
}



/* ==============================================
   GET HOLD CARD NUMBERS
============================================== */
function getHoldCardNumbers() {
  const sh = SpreadsheetApp.getActive().getSheetByName("HoldCard");
  if (!sh || sh.getLastRow() < 2) return [];

  const data = sh.getRange(2, 1, sh.getLastRow() - 1, 9).getValues();
  console.log(data)
  return data
    .filter(r => String(r[8]).toUpperCase() === "OPEN")
    .map(r => r[0]);
}

/* ==============================================
   SAVE PROCESS ENTRY
============================================== */
function saveProcessEntry(data) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName("ProcessEntry");

  if (!sh) {
    sh = ss.insertSheet("ProcessEntry");
    sh.appendRow([
      "Timestamp", "Hold Card Number",
      "Process", "Operator",
      "Duration (Min)", "Remark"
    ]);
  }

  sh.appendRow([
    new Date(),
    data.holdCardNumber,
    data.process,
    data.operator,
    data.duration,
    data.remark
  ]);

  autoCloseHold_(data.holdCardNumber);
  return { status: "success" };
}

/* ==============================================
   AUTO CLOSE HOLD
============================================== */
function autoCloseHold_(holdNo) {
  const sh = SpreadsheetApp.getActive().getSheetByName("HoldCard");
  if (!sh) return;

  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === holdNo) {
      sh.getRange(i + 1, 9).setValue("CLOSED");
      sh.getRange(i + 1, 10).setValue("DONE");
      break;
    }
  }
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

