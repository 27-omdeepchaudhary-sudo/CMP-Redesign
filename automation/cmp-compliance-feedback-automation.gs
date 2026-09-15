
// ============================================================
// CMP M4 - COMPLIANCE + FEEDBACK ALERT AUTOMATION
// FINAL CLEAN VERSION
// ============================================================
//
// Feedback_Log FINAL STRUCTURE:
// A = Pair_ID
// B = Interaction_ID
// C = Mentee_ID
// D = Mentor_ID
// E = Mentee Name
// F = Mentee Email
// G = Mentor Name
// H = Mentor Email
// I = Mentee Feedback
// J = Mentor Feedback
// K = Feedback Status
// L = Reason
//
// Compliance_Master FINAL STRUCTURE:
// A = Pair_ID
// B = Mentor_ID
// C = Mentee_ID
// D = Pair_Status
// E = Session_Completion
// F = Feedback_Completion
// G = Pending_Actions
// H = Overdue_Actions
// I = Compliance_Status
//
// IMPORTANT:
// This script READS Compliance_Master formulas/results.
// It does NOT overwrite columns E:I.
// ============================================================


// ============================================================
// MAIN FUNCTION
// ============================================================

function checkCMPCompliance() {

  // ==========================================================
  // TEMPORARY EMAIL FREEZE
  // ==========================================================

  const EMAIL_AUTOMATION_ENABLED = true;

  if (!EMAIL_AUTOMATION_ENABLED) {
    Logger.log("CMP email automation is FROZEN.");
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const complianceSheet = ss.getSheetByName("Compliance_Master");
  const interactionSheet = ss.getSheetByName("Interaction_Log");
  const feedbackSheet = ss.getSheetByName("Feedback_Log");
  const configSheet = ss.getSheetByName("Config");

  // ----------------------------------------------------------
  // VALIDATE REQUIRED SHEETS
  // ----------------------------------------------------------

  if (!complianceSheet) {
    throw new Error("Compliance_Master sheet not found.");
  }

  if (!interactionSheet) {
    throw new Error("Interaction_Log sheet not found.");
  }

  if (!feedbackSheet) {
    throw new Error("Feedback_Log sheet not found.");
  }


  // ==========================================================
  // PROGRAM OFFICE EMAIL
  // ==========================================================

  let programOfficeEmail =
    Session.getEffectiveUser().getEmail();

  if (configSheet) {

    const configData =
      configSheet.getDataRange().getValues();

    for (let i = 1; i < configData.length; i++) {

      const parameter =
        String(configData[i][0]).trim().toLowerCase();

      const value =
        String(configData[i][1]).trim();

      if (
        parameter === "program office email" &&
        value
      ) {
        programOfficeEmail = value;
        break;
      }
    }
  }


  // ==========================================================
  // READ SHEET DATA
  // ==========================================================

  const complianceData =
    complianceSheet.getDataRange().getValues();

  const interactionData =
    interactionSheet.getDataRange().getValues();

  const feedbackData =
    feedbackSheet.getDataRange().getValues();


  if (complianceData.length < 2) {

    Logger.log(
      "No Compliance_Master records found."
    );

    return;
  }


  // ==========================================================
  // BUILD INTERACTION SUMMARY BY PAIR_ID
  // ==========================================================

  const interactionSummary = {};

  for (let i = 1; i < interactionData.length; i++) {

    const row = interactionData[i];

    // Interaction_Log structure
    // B = Pair_ID
    // H = Session_Status
    // L = Action_Due_Date
    // M = Action_Status

    const pairId =
      String(row[1]).trim();

    const sessionStatus =
      String(row[7]).trim();

    const actionDueDate =
      row[11];

    const actionStatus =
      String(row[12]).trim();


    if (!pairId) {
      continue;
    }


    if (!interactionSummary[pairId]) {

      interactionSummary[pairId] = {

        totalSessions: 0,

        completedSessions: 0,

        pendingActions: 0,

        overdueActions: 0,

        relevantInteraction: null
      };
    }


    // --------------------------------------------------------
    // SESSION
    // --------------------------------------------------------

    interactionSummary[pairId].totalSessions++;


    if (
      sessionStatus.toLowerCase() ===
      "completed"
    ) {

      interactionSummary[pairId]
        .completedSessions++;
    }


    // --------------------------------------------------------
    // ACTION
    // --------------------------------------------------------

    if (
      actionStatus.toLowerCase() !==
      "completed" &&
      actionStatus !== ""
    ) {

      interactionSummary[pairId]
        .pendingActions++;


      // Check overdue action

      if (
        actionDueDate instanceof Date &&
        actionDueDate < new Date()
      ) {

        interactionSummary[pairId]
          .overdueActions++;
      }


      // Keep latest unresolved interaction

      interactionSummary[pairId]
        .relevantInteraction = row;
    }
  }


  // ==========================================================
  // BUILD FEEDBACK SUMMARY BY PAIR_ID
  // ==========================================================

  const feedbackSummary = {};

  for (let i = 1; i < feedbackData.length; i++) {

    const row = feedbackData[i];

    // Feedback_Log FINAL STRUCTURE
    //
    // A = Pair_ID
    // B = Interaction_ID
    // C = Mentee_ID
    // D = Mentor_ID
    // E = Mentee Name
    // F = Mentee Email
    // G = Mentor Name
    // H = Mentor Email
    // I = Mentee Feedback
    // J = Mentor Feedback
    // K = Feedback Status
    // L = Reason


    const pairId =
      String(row[0]).trim();

    const interactionId =
      String(row[1]).trim();

    const menteeId =
      String(row[2]).trim();

    const mentorId =
      String(row[3]).trim();

    const menteeName =
      String(row[4]).trim();

    const menteeEmail =
      String(row[5]).trim();

    const mentorName =
      String(row[6]).trim();

    const mentorEmail =
      String(row[7]).trim();

    const menteeFeedback =
      String(row[8]).trim();

    const mentorFeedback =
      String(row[9]).trim();


    if (!pairId) {
      continue;
    }


    if (!feedbackSummary[pairId]) {

      feedbackSummary[pairId] = {

        totalInteractions: 0,

        completeFeedback: 0,

        menteePending: 0,

        mentorPending: 0,

        feedbackRows: []
      };
    }


    feedbackSummary[pairId]
      .totalInteractions++;


    // --------------------------------------------------------
    // CHECK MENTEE FEEDBACK
    // --------------------------------------------------------

    const menteeOK =
      menteeFeedback.toLowerCase() === "ok";


    // --------------------------------------------------------
    // CHECK MENTOR FEEDBACK
    // --------------------------------------------------------

    const mentorOK =
      mentorFeedback.toLowerCase() === "ok";


    // --------------------------------------------------------
    // FEEDBACK COMPLETION
    // --------------------------------------------------------

    if (
      menteeOK &&
      mentorOK
    ) {

      feedbackSummary[pairId]
        .completeFeedback++;

    } else {

      if (!menteeOK) {

        feedbackSummary[pairId]
          .menteePending++;
      }


      if (!mentorOK) {

        feedbackSummary[pairId]
          .mentorPending++;
      }
    }


    // --------------------------------------------------------
    // STORE FEEDBACK ROW
    // --------------------------------------------------------

    feedbackSummary[pairId]
      .feedbackRows.push({

        interactionId: interactionId,

        menteeId: menteeId,

        mentorId: mentorId,

        menteeName: menteeName,

        menteeEmail: menteeEmail,

        mentorName: mentorName,

        mentorEmail: mentorEmail,

        menteeFeedback: menteeFeedback,

        mentorFeedback: mentorFeedback
      });
  }


  // ==========================================================
  // PROCESS EACH COMPLIANCE RECORD
  // ==========================================================

  for (let i = 1; i < complianceData.length; i++) {

    const row = complianceData[i];


    // Compliance_Master
    //
    // A = Pair_ID
    // B = Mentor_ID
    // C = Mentee_ID
    // D = Pair_Status
    // E = Session_Completion
    // F = Feedback_Completion
    // G = Pending_Actions
    // H = Overdue_Actions
    // I = Compliance_Status


    const pairId =
      String(row[0]).trim();

    const mentorId =
      String(row[1]).trim();

    const menteeId =
      String(row[2]).trim();

    const pairStatus =
      String(row[3]).trim();

    const sessionCompletion =
      Number(row[4]) || 0;

    const feedbackCompletion =
      Number(row[5]) || 0;

    const pendingActions =
      Number(row[6]) || 0;

    const overdueActions =
      Number(row[7]) || 0;

    const complianceStatus =
      String(row[8]).trim();


    if (!pairId) {
      continue;
    }


    // ========================================================
    // DETERMINE COMPLIANCE REASON FOR ALERT PURPOSE
    // ========================================================

    let complianceReason =
      "Incomplete CMP Requirements";


    if (overdueActions > 0) {

      complianceReason =
        "Overdue Action";

    }

    else if (
      feedbackSummary[pairId] &&
      feedbackSummary[pairId].menteePending > 0 &&
      feedbackSummary[pairId].mentorPending > 0
    ) {

      complianceReason =
        "Both Feedback Pending";

    }

    else if (
      feedbackSummary[pairId] &&
      feedbackSummary[pairId].menteePending > 0
    ) {

      complianceReason =
        "Mentee Feedback Pending";

    }

    else if (
      feedbackSummary[pairId] &&
      feedbackSummary[pairId].mentorPending > 0
    ) {

      complianceReason =
        "Mentor Feedback Pending";

    }

    else if (pendingActions > 0) {

      complianceReason =
        "Pending Action";

    }

    else if (
      pairStatus.toLowerCase() === "active" &&
      sessionCompletion === 1 &&
      feedbackCompletion === 1
    ) {

      complianceReason =
        "—";
    }


    // ========================================================
    // FEEDBACK ALERTS
    // ========================================================

    if (feedbackSummary[pairId]) {

      const feedbackRows =
        feedbackSummary[pairId].feedbackRows;


      for (
        let f = 0;
        f < feedbackRows.length;
        f++
      ) {

        const feedback =
          feedbackRows[f];


        // ----------------------------------------------------
        // MENTEE FEEDBACK MISSING
        // ----------------------------------------------------

        if (
          feedback.menteeFeedback
            .toLowerCase() !== "ok"
        ) {

          sendFeedbackAlert(
            "MENTEE",
            pairId,
            feedback,
            programOfficeEmail
          );
        }


        // ----------------------------------------------------
        // MENTOR FEEDBACK MISSING
        // ----------------------------------------------------

        if (
          feedback.mentorFeedback
            .toLowerCase() !== "ok"
        ) {

          sendFeedbackAlert(
            "MENTOR",
            pairId,
            feedback,
            programOfficeEmail
          );
        }
      }
    }


    // ========================================================
    // ACTION ALERT
    // ========================================================

    if (
      (
        complianceStatus === "At Risk" ||
        complianceStatus === "Non-Compliant"
      ) &&
      pendingActions > 0
    ) {

      sendActionAlert(
        pairId,
        mentorId,
        menteeId,
        complianceStatus,
        complianceReason,
        interactionSummary[pairId],
        programOfficeEmail
      );
    }
  }


  // ==========================================================
  // LOG COMPLETION
  // ==========================================================

  Logger.log(
    "CMP Compliance check completed successfully."
  );
}



// ============================================================
// FEEDBACK ALERT FUNCTION
// ============================================================

function sendFeedbackAlert(
  responsiblePerson,
  pairId,
  feedback,
  programOfficeEmail
) {

  const properties =
    PropertiesService.getScriptProperties();


  let recipientEmail = "";

  let recipientName = "";

  let otherPersonEmail = "";

  let pendingType = "";


  // ==========================================================
  // MENTEE RESPONSIBILITY
  // ==========================================================

  if (
    responsiblePerson === "MENTEE"
  ) {

    // PRIMARY SOURCE:
    // Feedback_Log Mentee Email

    recipientEmail =
      feedback.menteeEmail;

    recipientName =
      feedback.menteeName;

    // CC Mentor

    otherPersonEmail =
      feedback.mentorEmail;

    pendingType =
      "Mentee Feedback";
  }


  // ==========================================================
  // MENTOR RESPONSIBILITY
  // ==========================================================

  if (
    responsiblePerson === "MENTOR"
  ) {

    // PRIMARY SOURCE:
    // Feedback_Log Mentor Email

    recipientEmail =
      feedback.mentorEmail;

    recipientName =
      feedback.mentorName;

    // CC Mentee

    otherPersonEmail =
      feedback.menteeEmail;

    pendingType =
      "Mentor Feedback";
  }


  // ==========================================================
  // EMAIL VALIDATION
  // ==========================================================

  if (!recipientEmail) {

    Logger.log(
      "No email found for " +
      responsiblePerson +
      " | Pair: " +
      pairId +
      " | Interaction: " +
      feedback.interactionId
    );

    return;
  }


  // ==========================================================
  // DUPLICATE PREVENTION
  // ==========================================================

  const alertKey =
    "FEEDBACK_" +
    pairId +
    "_" +
    feedback.interactionId +
    "_" +
    responsiblePerson;


  if (
    properties.getProperty(alertKey)
  ) {

    Logger.log(
      "Feedback alert already sent: " +
      alertKey
    );

    return;
  }


  // ==========================================================
  // EMAIL SUBJECT
  // ==========================================================

  const subject =
    "CMP Feedback Reminder | " +
    pairId +
    " | " +
    feedback.interactionId +
    " | " +
    pendingType;


  // ==========================================================
  // EMAIL BODY
  // ==========================================================

  let body = "";


  body +=
    "Respected " +
    recipientName +
    ",\n\n";


  body +=
    "This is a reminder regarding your pending CMP feedback.\n\n";


  body +=
    "----------------------------------------\n";

  body +=
    "CMP FEEDBACK DETAILS\n";

  body +=
    "----------------------------------------\n\n";


  body +=
    "Pair ID: " +
    pairId +
    "\n";


  body +=
    "Interaction ID: " +
    feedback.interactionId +
    "\n";


  body +=
    "Mentee ID: " +
    feedback.menteeId +
    "\n";


  body +=
    "Mentee Name: " +
    feedback.menteeName +
    "\n";


  body +=
    "Mentor ID: " +
    feedback.mentorId +
    "\n";


  body +=
    "Mentor Name: " +
    feedback.mentorName +
    "\n\n";


  body +=
    "Pending Responsibility: " +
    pendingType +
    "\n\n";


  body +=
    "Please submit the required feedback at the earliest to keep the CMP record compliant.\n\n";


  body +=
    "Regards,\n";

  body +=
    "CMP Program Office";


  // ==========================================================
  // BUILD CC LIST
  // ==========================================================

  let ccList = [];


  if (
    otherPersonEmail &&
    otherPersonEmail.toLowerCase() !==
    recipientEmail.toLowerCase()
  ) {

    ccList.push(
      otherPersonEmail
    );
  }


  if (
    programOfficeEmail &&
    programOfficeEmail.toLowerCase() !==
    recipientEmail.toLowerCase() &&
    programOfficeEmail.toLowerCase() !==
    String(otherPersonEmail).toLowerCase()
  ) {

    ccList.push(
      programOfficeEmail
    );
  }


  // ==========================================================
  // SEND EMAIL
  // ==========================================================

  const emailOptions = {

    to: recipientEmail,

    subject: subject,

    body: body
  };


  if (ccList.length > 0) {

    emailOptions.cc =
      ccList.join(",");
  }


  MailApp.sendEmail(
    emailOptions
  );


  // ==========================================================
  // SAVE ALERT RECORD
  // ==========================================================

  properties.setProperty(
    alertKey,
    new Date().toISOString()
  );


  Logger.log(
    "Feedback alert sent to: " +
    recipientEmail +
    " | " +
    pendingType +
    " | " +
    pairId +
    " | " +
    feedback.interactionId
  );
}



// ============================================================
// ACTION ALERT FUNCTION
// ============================================================

function sendActionAlert(
  pairId,
  mentorId,
  menteeId,
  complianceStatus,
  complianceReason,
  interactionSummary,
  programOfficeEmail
) {

  const properties =
    PropertiesService.getScriptProperties();


  if (
    !interactionSummary ||
    !interactionSummary.relevantInteraction
  ) {

    return;
  }


  const interaction =
    interactionSummary.relevantInteraction;


  // ==========================================================
  // INTERACTION DETAILS
  // ==========================================================

  const interactionId =
    String(interaction[0]).trim();

  const sessionNumber =
    String(interaction[4]).trim();

  const scheduledDate =
    formatDate(interaction[5]);

  const actualDate =
    formatDate(interaction[6]);

  const sessionStatus =
    String(interaction[7]).trim();

  const nextAction =
    String(interaction[10]).trim();

  const actionDueDate =
    formatDate(interaction[11]);

  const actionStatus =
    String(interaction[12]).trim();


  // ==========================================================
  // LOOKUP EMAILS FROM MASTER SHEETS
  // ==========================================================

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();


  const menteeSheet =
    ss.getSheetByName("Mentee_Master");

  const mentorSheet =
    ss.getSheetByName("Mentor_Master");


  let menteeEmail = "";

  let mentorEmail = "";


  // ----------------------------------------------------------
  // MENTEE EMAIL
  // ----------------------------------------------------------

  if (menteeSheet) {

    const menteeData =
      menteeSheet.getDataRange().getValues();


    for (
      let i = 1;
      i < menteeData.length;
      i++
    ) {

      if (
        String(menteeData[i][0]).trim() ===
        menteeId
      ) {

        menteeEmail =
          String(
            menteeData[i][2]
          ).trim();

        break;
      }
    }
  }


  // ----------------------------------------------------------
  // MENTOR EMAIL
  // ----------------------------------------------------------

  if (mentorSheet) {

    const mentorData =
      mentorSheet.getDataRange().getValues();


    for (
      let i = 1;
      i < mentorData.length;
      i++
    ) {

      if (
        String(mentorData[i][0]).trim() ===
        mentorId
      ) {

        mentorEmail =
          String(
            mentorData[i][2]
          ).trim();

        break;
      }
    }
  }


  if (
    !menteeEmail &&
    !mentorEmail
  ) {

    Logger.log(
      "No email found for action alert: " +
      pairId
    );

    return;
  }


  // ==========================================================
  // DUPLICATE PREVENTION
  // ==========================================================

  const alertKey =
    "ACTION_" +
    pairId +
    "_" +
    interactionId +
    "_" +
    complianceStatus;


  if (
    properties.getProperty(alertKey)
  ) {

    Logger.log(
      "Action alert already sent: " +
      alertKey
    );

    return;
  }


  // ==========================================================
  // RECIPIENT
  // ==========================================================

  let recipientEmail =
    menteeEmail;

  let otherPersonEmail =
    mentorEmail;


  // If reason specifically relates to mentor

  if (
    complianceReason
      .toLowerCase()
      .includes("mentor")
  ) {

    recipientEmail =
      mentorEmail;

    otherPersonEmail =
      menteeEmail;
  }


  if (!recipientEmail) {
    return;
  }


  // ==========================================================
  // SUBJECT
  // ==========================================================

  const subject =
    "CMP Compliance Alert | " +
    pairId +
    " | " +
    complianceStatus;


  // ==========================================================
  // EMAIL BODY
  // ==========================================================

  let body = "";


  body +=
    "Respected CMP Participant,\n\n";


  body +=
    "Your CMP interaction record requires attention.\n\n";


  body +=
    "----------------------------------------\n";

  body +=
    "CMP COMPLIANCE DETAILS\n";

  body +=
    "----------------------------------------\n\n";


  body +=
    "Pair ID: " +
    pairId +
    "\n";


  body +=
    "Mentee ID: " +
    menteeId +
    "\n";


  body +=
    "Mentor ID: " +
    mentorId +
    "\n";


  body +=
    "Interaction ID: " +
    interactionId +
    "\n";


  body +=
    "Compliance Status: " +
    complianceStatus +
    "\n";


  body +=
    "Reason: " +
    complianceReason +
    "\n\n";


  body +=
    "----------------------------------------\n";

  body +=
    "ACTIVITY DETAILS\n";

  body +=
    "----------------------------------------\n\n";


  body +=
    "Session Number: " +
    sessionNumber +
    "\n";


  body +=
    "Scheduled Date: " +
    scheduledDate +
    "\n";


  body +=
    "Actual Date: " +
    actualDate +
    "\n";


  body +=
    "Session Status: " +
    sessionStatus +
    "\n";


  body +=
    "Next Action: " +
    nextAction +
    "\n";


  body +=
    "Action Due Date: " +
    actionDueDate +
    "\n";


  body +=
    "Action Status: " +
    actionStatus +
    "\n\n";


  body +=
    "Pending Actions: " +
    interactionSummary.pendingActions +
    "\n";


  body +=
    "Overdue Actions: " +
    interactionSummary.overdueActions +
    "\n\n";


  body +=
    "Please complete the required activity within the applicable timeline or contact the CMP Program Office if clarification is required.\n\n";


  body +=
    "Regards,\n";

  body +=
    "CMP Program Office";


  // ==========================================================
  // CC
  // ==========================================================

  let ccList = [];


  if (
    otherPersonEmail &&
    otherPersonEmail.toLowerCase() !==
    recipientEmail.toLowerCase()
  ) {

    ccList.push(
      otherPersonEmail
    );
  }


  if (
    programOfficeEmail &&
    programOfficeEmail.toLowerCase() !==
    recipientEmail.toLowerCase() &&
    programOfficeEmail.toLowerCase() !==
    String(otherPersonEmail).toLowerCase()
  ) {

    ccList.push(
      programOfficeEmail
    );
  }


  // ==========================================================
  // SEND
  // ==========================================================

  const emailOptions = {

    to: recipientEmail,

    subject: subject,

    body: body
  };


  if (ccList.length > 0) {

    emailOptions.cc =
      ccList.join(",");
  }


  MailApp.sendEmail(
    emailOptions
  );


  // ==========================================================
  // SAVE ALERT RECORD
  // ==========================================================

  properties.setProperty(
    alertKey,
    new Date().toISOString()
  );


  Logger.log(
    "Action alert sent to: " +
    recipientEmail +
    " | " +
    pairId
  );
}



// ============================================================
// DATE FORMAT FUNCTION
// ============================================================

function formatDate(value) {

  if (!value) {
    return "";
  }


  if (value instanceof Date) {

    return Utilities.formatDate(
      value,
      Session.getScriptTimeZone(),
      "dd-MMM-yyyy"
    );
  }


  return String(value);
}
