const CONFIG = {
  sheetName: "Proposals",
  defaultTemplateDocId: "1QSABUNCmnlo8-u2zG8PSULIj0KW2MHQCniAgMcT6-Rk",
  defaultSpreadsheetId: "1BcKbUgpbeXWEa4ODeflD4CVogO1Wj57cnrK3LMm6XDU",
  requiredTemplateProperty: "TEMPLATE_DOC_ID",
  spreadsheetProperty: "SPREADSHEET_ID",
  outputFolderProperty: "OUTPUT_FOLDER_ID",
  fromNameProperty: "FROM_NAME",
};

/**
 * Serves the sales form as a web app.
 */
function doGet() {
  return HtmlService.createTemplateFromFile("Index")
    .evaluate()
    .setTitle("Automated Proposal Generator")
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

/**
 * Allows splitting CSS/JS into extra HTML files.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * One-time setup helper: creates template/doc storage assets and shares with an email.
 * Also stores created IDs in Script Properties.
 *
 * @param {string} ownerEmail
 * @return {{templateDocId: string, templateDocUrl: string, spreadsheetId: string, spreadsheetUrl: string, outputFolderId: string, outputFolderUrl: string}}
 */
function setupProjectAssets(ownerEmail) {
  const email = (ownerEmail || "").trim();
  if (!email || !isValidEmail_(email)) {
    throw new Error(
      'Pass a valid email, e.g. setupProjectAssets("you@example.com")',
    );
  }

  const timestamp = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "yyyyMMdd-HHmmss",
  );
  const folder = DriveApp.createFolder(
    "Proposal Generator Assets " + timestamp,
  );

  const templateDoc = DocumentApp.create("Proposal Template - " + timestamp);
  seedTemplateDoc_(templateDoc);
  templateDoc.saveAndClose();

  const templateFile = DriveApp.getFileById(templateDoc.getId());
  folder.addFile(templateFile);
  safelyRemoveFromRoot_(templateFile);

  const spreadsheet = SpreadsheetApp.create(
    "Proposal Generator Database - " + timestamp,
  );
  const proposalSheet = spreadsheet.getSheets()[0];
  proposalSheet.setName(CONFIG.sheetName);
  proposalSheet.clear();
  proposalSheet.appendRow([
    "Timestamp",
    "Proposal Number",
    "Client Name",
    "Client Company",
    "Recipient Emails",
    "Project Title",
    "Budget",
    "Currency",
    "Sales Rep",
    "Doc URL",
    "PDF URL",
    "Status",
  ]);

  const spreadsheetFile = DriveApp.getFileById(spreadsheet.getId());
  folder.addFile(spreadsheetFile);
  safelyRemoveFromRoot_(spreadsheetFile);

  folder.addEditor(email);
  templateFile.addEditor(email);
  spreadsheetFile.addEditor(email);

  PropertiesService.getScriptProperties().setProperties(
    {
      TEMPLATE_DOC_ID: templateDoc.getId(),
      SPREADSHEET_ID: spreadsheet.getId(),
      OUTPUT_FOLDER_ID: folder.getId(),
    },
    true,
  );

  return {
    templateDocId: templateDoc.getId(),
    templateDocUrl: templateFile.getUrl(),
    spreadsheetId: spreadsheet.getId(),
    spreadsheetUrl: spreadsheet.getUrl(),
    outputFolderId: folder.getId(),
    outputFolderUrl: folder.getUrl(),
  };
}

/**
 * Runs a full end-to-end check inside Apps Script:
 * 1) creates setup assets, 2) generates proposal doc+pdf, 3) sends email, 4) logs to sheet.
 *
 * @param {string=} testEmail Optional recipient email for smoke test.
 * @return {{assets: Object, submission: Object}}
 */
function runEndToEndSmokeTest(testEmail) {
  const recipient =
    (testEmail || "").trim() || Session.getActiveUser().getEmail();

  if (!recipient || !isValidEmail_(recipient)) {
    throw new Error(
      'Provide a valid email, e.g. runEndToEndSmokeTest("you@example.com")',
    );
  }

  const assets = setupProjectAssets(recipient);

  const payload = {
    clientName: "Labib Test Client",
    clientCompany: "Demo Company",
    clientEmails: recipient,
    projectTitle: "Automated Proposal Generator Pilot",
    projectScope:
      "Create a proposal automatically from form data and send PDF by email.",
    deliverables: "Proposal document, generated PDF, and delivery email",
    timeline: "2 weeks",
    budget: "5000",
    currency: "USD",
    salesRep: "Sales Automation Bot",
    emailSubject: "Smoke Test Proposal Delivery",
    emailMessage:
      "This is an automated end-to-end smoke test email with proposal PDF attachment.",
  };

  const submission = submitProposal(payload);
  return {
    assets: assets,
    submission: submission,
  };
}

/**
 * Handles form submission from the frontend.
 * @param {Object} payload
 * @return {{success: boolean, proposalNumber: string, docUrl: string, pdfUrl: string, message: string}}
 */
function submitProposal(payload) {
  validatePayload_(payload);

  const proposalNumber = buildProposalNumber_();
  const templateDocId = getTemplateDocId_();
  const recipients = parseRecipients_(payload.clientEmails);

  const docFile = createProposalDoc_(templateDocId, payload, proposalNumber);
  const pdfFile = createPdfFromDoc_(docFile, proposalNumber);
  const logRow = logSubmission_(
    payload,
    proposalNumber,
    docFile,
    pdfFile,
    recipients,
    "GENERATED",
  );

  try {
    sendProposalEmail_(payload, recipients, pdfFile, proposalNumber);
    updateSubmissionStatus_(logRow, "EMAILED");
  } catch (error) {
    updateSubmissionStatus_(logRow, "EMAIL_FAILED");
    throw error;
  }

  return {
    success: true,
    proposalNumber: proposalNumber,
    docUrl: docFile.getUrl(),
    pdfUrl: pdfFile.getUrl(),
    message: "Proposal generated and emailed successfully.",
  };
}

function validatePayload_(payload) {
  const requiredFields = [
    "clientName",
    "clientCompany",
    "clientEmails",
    "projectTitle",
    "projectScope",
    "budget",
    "salesRep",
  ];

  requiredFields.forEach(function (field) {
    if (!payload[field] || String(payload[field]).trim() === "") {
      throw new Error("Missing required field: " + field);
    }
  });

  parseRecipients_(payload.clientEmails);
}

function getTemplateDocId_() {
  const props = PropertiesService.getScriptProperties();
  const templateDocId =
    props.getProperty(CONFIG.requiredTemplateProperty) ||
    CONFIG.defaultTemplateDocId;

  if (!templateDocId || templateDocId.trim() === "") {
    throw new Error(
      "Missing Script Property: " +
        CONFIG.requiredTemplateProperty +
        ". Set it in Project Settings > Script Properties.",
    );
  }

  props.setProperty(CONFIG.requiredTemplateProperty, templateDocId);
  return templateDocId;
}

function buildProposalNumber_() {
  const now = new Date();
  const datePart = Utilities.formatDate(
    now,
    Session.getScriptTimeZone(),
    "yyyyMMdd-HHmmss",
  );
  const randomPart = Math.floor(100 + Math.random() * 900);
  return "PROP-" + datePart + "-" + randomPart;
}

function createProposalDoc_(templateDocId, payload, proposalNumber) {
  const folder = getOutputFolder_();
  const copyName = proposalNumber + " - " + payload.clientCompany + " Proposal";
  const templateFile = DriveApp.getFileById(templateDocId);
  const copiedFile = folder
    ? templateFile.makeCopy(copyName, folder)
    : templateFile.makeCopy(copyName);

  const doc = DocumentApp.openById(copiedFile.getId());
  const body = doc.getBody();
  const generatedDate = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "dd MMM yyyy",
  );

  const replacements = {
    "{{PROPOSAL_NUMBER}}": proposalNumber,
    "{{DATE}}": generatedDate,
    "{{CLIENT_NAME}}": payload.clientName,
    "{{CLIENT_COMPANY}}": payload.clientCompany,
    "{{PROJECT_TITLE}}": payload.projectTitle,
    "{{PROJECT_SCOPE}}": payload.projectScope,
    "{{DELIVERABLES}}": payload.deliverables || "N/A",
    "{{TIMELINE}}": payload.timeline || "To be finalized",
    "{{BUDGET}}": (payload.currency || "USD") + " " + payload.budget,
    "{{SALES_REP}}": payload.salesRep,
  };

  Object.keys(replacements).forEach(function (token) {
    body.replaceText(token, sanitizeForReplace_(replacements[token]));
  });

  doc.saveAndClose();
  return copiedFile;
}

function sanitizeForReplace_(value) {
  return String(value).replace(/\$/g, "\\$");
}

function createPdfFromDoc_(docFile, proposalNumber) {
  const pdfBlob = docFile
    .getBlob()
    .getAs(MimeType.PDF)
    .setName(proposalNumber + ".pdf");
  const folder = getOutputFolder_();
  return folder ? folder.createFile(pdfBlob) : DriveApp.createFile(pdfBlob);
}

function getOutputFolder_() {
  const folderId = PropertiesService.getScriptProperties().getProperty(
    CONFIG.outputFolderProperty,
  );
  if (!folderId) {
    return null;
  }
  return DriveApp.getFolderById(folderId);
}

function parseRecipients_(emailString) {
  const emails = String(emailString)
    .split(/[;,]/)
    .map(function (item) {
      return item.trim();
    })
    .filter(function (item) {
      return item !== "";
    });

  if (emails.length === 0) {
    throw new Error("At least one recipient email is required.");
  }

  emails.forEach(function (email) {
    if (!isValidEmail_(email)) {
      throw new Error("Invalid email address: " + email);
    }
  });

  return emails;
}

function isValidEmail_(email) {
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(email);
}

function sendProposalEmail_(payload, recipients, pdfFile, proposalNumber) {
  const fromName =
    PropertiesService.getScriptProperties().getProperty(
      CONFIG.fromNameProperty,
    ) || "Sales Team";
  const subject =
    payload.emailSubject ||
    "Proposal " + proposalNumber + " - " + payload.projectTitle;
  const plainBody =
    payload.emailMessage || buildDefaultEmailBody_(payload, proposalNumber);

  MailApp.sendEmail({
    to: recipients.join(","),
    subject: subject,
    body: plainBody,
    name: fromName,
    attachments: [pdfFile.getBlob()],
  });
}

function buildDefaultEmailBody_(payload, proposalNumber) {
  return [
    "Hi " + payload.clientName + ",",
    "",
    "Please find attached our proposal (" +
      proposalNumber +
      ") for " +
      payload.projectTitle +
      ".",
    "We are excited to support " + payload.clientCompany + ".",
    "",
    "Best regards,",
    payload.salesRep,
  ].join("\n");
}

function logSubmission_(
  payload,
  proposalNumber,
  docFile,
  pdfFile,
  recipients,
  status,
) {
  const sheet = getOrCreateSheet_();

  sheet.appendRow([
    new Date(),
    proposalNumber,
    payload.clientName,
    payload.clientCompany,
    recipients.join(", "),
    payload.projectTitle,
    payload.budget,
    payload.currency || "USD",
    payload.salesRep,
    docFile.getUrl(),
    pdfFile.getUrl(),
    status,
  ]);

  return {
    sheetName: sheet.getName(),
    row: sheet.getLastRow(),
  };
}

function updateSubmissionStatus_(logRow, status) {
  const sheet = getOrCreateSheet_();
  sheet.getRange(logRow.row, 12).setValue(status);
}

function writeTestRowToSheet() {
  const sheet = getOrCreateSheet_();
  sheet.appendRow([
    new Date(),
    "TEST",
    "Test Client",
    "Test Company",
    Session.getActiveUser().getEmail() || "test@example.com",
    "Spreadsheet Connection Test",
    "0",
    "USD",
    "Apps Script",
    "",
    "",
    "TEST",
  ]);

  return {
    spreadsheetId: PropertiesService.getScriptProperties().getProperty(
      CONFIG.spreadsheetProperty,
    ),
    sheetName: sheet.getName(),
    row: sheet.getLastRow(),
  };
}

function getOrCreateSheet_() {
  const props = PropertiesService.getScriptProperties();
  let spreadsheetId =
    props.getProperty(CONFIG.spreadsheetProperty) || CONFIG.defaultSpreadsheetId;
  let spreadsheet;

  if (spreadsheetId) {
    spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    props.setProperty(CONFIG.spreadsheetProperty, spreadsheetId);
  } else {
    spreadsheet = SpreadsheetApp.create("Proposal Generator Database");
    props.setProperty(CONFIG.spreadsheetProperty, spreadsheet.getId());
  }

  let sheet = spreadsheet.getSheetByName(CONFIG.sheetName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(CONFIG.sheetName);
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      "Timestamp",
      "Proposal Number",
      "Client Name",
      "Client Company",
      "Recipient Emails",
      "Project Title",
      "Budget",
      "Currency",
      "Sales Rep",
      "Doc URL",
      "PDF URL",
      "Status",
    ]);
  }

  return sheet;
}

/**
 * Returns all important runtime configuration values for quick inspection.
 * Run this from Apps Script editor and inspect the returned object in Execution log.
 */
function inspectProjectConfig() {
  const props = PropertiesService.getScriptProperties();
  const allProps = props.getProperties();
  const templateDocId =
    allProps[CONFIG.requiredTemplateProperty] || CONFIG.defaultTemplateDocId;
  const spreadsheetId =
    allProps[CONFIG.spreadsheetProperty] || CONFIG.defaultSpreadsheetId;
  const outputFolderId = allProps[CONFIG.outputFolderProperty] || null;

  return {
    scriptId: ScriptApp.getScriptId(),
    configConstants: CONFIG,
    scriptProperties: allProps,
    resolvedAssets: {
      templateDoc: templateDocId
        ? {
            id: templateDocId,
            url: 'https://docs.google.com/document/d/' + templateDocId + '/edit'
          }
        : null,
      spreadsheet: spreadsheetId
        ? {
            id: spreadsheetId,
            url: 'https://docs.google.com/spreadsheets/d/' + spreadsheetId + '/edit'
          }
        : null,
      outputFolder: outputFolderId
        ? {
            id: outputFolderId,
            url: 'https://drive.google.com/drive/folders/' + outputFolderId
          }
        : null
    }
  };
}

function seedTemplateDoc_(doc) {
  const body = doc.getBody();
  body.clear();
  body
    .appendParagraph("PROPOSAL")
    .setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph("Proposal Number: {{PROPOSAL_NUMBER}}");
  body.appendParagraph("Date: {{DATE}}");
  body.appendParagraph("");
  body.appendParagraph("Client Name: {{CLIENT_NAME}}");
  body.appendParagraph("Client Company: {{CLIENT_COMPANY}}");
  body.appendParagraph("");
  body
    .appendParagraph("Project Title")
    .setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph("{{PROJECT_TITLE}}");
  body.appendParagraph("");
  body
    .appendParagraph("Project Scope")
    .setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph("{{PROJECT_SCOPE}}");
  body.appendParagraph("");
  body
    .appendParagraph("Deliverables")
    .setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph("{{DELIVERABLES}}");
  body.appendParagraph("");
  body.appendParagraph("Timeline: {{TIMELINE}}");
  body.appendParagraph("Budget: {{BUDGET}}");
  body.appendParagraph("Sales Representative: {{SALES_REP}}");
}

function safelyRemoveFromRoot_(file) {
  try {
    DriveApp.getRootFolder().removeFile(file);
  } catch (error) {
    // Keep going if root removal is not allowed in this Drive context.
  }
}
