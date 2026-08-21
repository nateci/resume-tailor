// Paste this into Extensions > Apps Script on EACH of the two tracker
// sheets (new-grad and intern), replace SYNC_TOKEN with the same random
// string you set as SHEET_SYNC_TOKEN for scripts/sync_sheets.py, then
// Deploy > New deployment > Web app > Execute as: Me > Who has access:
// Anyone > Deploy. Copy the resulting URL into SHEET_WEBHOOK_URL_NEWGRAD
// or SHEET_WEBHOOK_URL_INTERN.
//
// Column K holds this script's own job-id tag so re-runs can find and
// update a row instead of duplicating it. Columns A-J are never touched
// beyond Status (E) and Date Applied (A, only if blank) -- everything
// else (Contact Name, Resume Ver., Interview Dates, Notes) is left for
// hand-editing and this script will never overwrite it.

const SYNC_TOKEN = "REPLACE_WITH_A_RANDOM_STRING";

function doPost(e) {
  const body = JSON.parse(e.postData.contents);
  if (body.token !== SYNC_TOKEN) {
    return ContentService.createTextOutput(JSON.stringify({ error: "bad token" }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  const data = sheet.getDataRange().getValues(); // includes header row
  const ID_COL = 10;      // column K, 0-indexed
  const COMPANY_COL = 1;  // column B, 0-indexed

  let created = 0, updated = 0, skipped = 0;
  const skippedCompanies = [];

  (body.jobs || []).forEach(job => {
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][ID_COL] === job.id) { rowIndex = i; break; }
    }

    if (rowIndex !== -1) {
      const sheetRow = rowIndex + 1;
      sheet.getRange(sheetRow, 5).setValue(job.status); // E: Status
      if (!data[rowIndex][0] && job.date_applied) {
        sheet.getRange(sheetRow, 1).setValue(job.date_applied); // A: Date Applied
      }
      updated++;
      return;
    }

    const companyExists = data.slice(1).some(r =>
      String(r[COMPANY_COL]).trim().toLowerCase() === String(job.company).trim().toLowerCase());
    if (companyExists) {
      skipped++;
      skippedCompanies.push(job.company);
      return;
    }

    sheet.appendRow([
      job.date_applied, job.company, job.title, job.location,
      job.status, job.url, "", "", "", "", job.id,
    ]);
    created++;
  });

  return ContentService.createTextOutput(JSON.stringify({
    created, updated, skipped, skipped_companies: skippedCompanies,
  })).setMimeType(ContentService.MimeType.JSON);
}
