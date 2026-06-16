# Automated Proposal Generator (Google Apps Script)

This project provides an end-to-end proposal workflow:

1. Sales team fills in client/project details from a web form.
2. Apps Script duplicates a Google Docs template and replaces placeholders.
3. The generated proposal is converted to PDF.
4. PDF is emailed to one or multiple client recipients.
5. Submission metadata is logged into Google Sheets.

## Tech Stack

- Frontend: HTML + CSS (Apps Script HTML Service)
- Backend: Google Apps Script
- Database: Google Sheets
- Document generation: Google Docs template + Drive PDF conversion

## Files

- `Code.gs` - main backend logic (validation, template fill, PDF generation, email, sheet logging)
- `Index.html` - frontend form UI
- `Styles.html` - CSS styling for the form
- `appsscript.json` - Apps Script manifest

## Setup

1. Create a Google Doc template and include these exact placeholders where generated values should appear:

   - `{{PROPOSAL_NUMBER}}`
   - `{{DATE}}`
   - `{{CLIENT_NAME}}`
   - `{{CLIENT_COMPANY}}`
   - `{{PROJECT_TITLE}}`
   - `{{PROJECT_SCOPE}}`
   - `{{DELIVERABLES}}`
   - `{{TIMELINE}}`
   - `{{BUDGET}}`
   - `{{SALES_REP}}`

2. Create a new Apps Script project (or use clasp) and add these files.
3. Set Script Properties in Apps Script Project Settings:

   - `TEMPLATE_DOC_ID` (optional): ID of your Google Docs template.

     - Defaults to `1QSABUNCmnlo8-u2zG8PSULIj0KW2MHQCniAgMcT6-Rk`.

   - `SPREADSHEET_ID` (optional): existing Google Sheet ID for logs.

     - Defaults to `1BcKbUgpbeXWEa4ODeflD4CVogO1Wj57cnrK3LMm6XDU`.
     - If no default or property is present, script auto-creates one and stores its ID.

   - `OUTPUT_FOLDER_ID` (optional): Drive folder where generated docs/PDFs are saved.
   - `FROM_NAME` (optional): sender display name for outgoing emails.

4. Deploy as a Web App:

   - Execute as: `User accessing the web app`
   - Access: `Anyone with the link` (or your org policy)

5. Open the deployed URL and submit a sample proposal.

## Notes

- Recipient emails can be comma-separated or semicolon-separated.
- Email subject and body are optional in UI; defaults are auto-generated.
- Ensure deployment account has permission to access template doc, Drive folder, and MailApp.

## Review Checklist (Friday 8:30pm)

- Form submission works for complete and valid input.
- Template placeholders are correctly replaced.
- PDF is generated and attached in sent email.
- Sheet logging captures each submission row.
- Drive links for generated Doc and PDF are accessible.
