/**
 * Handles incoming POST requests from your website form.
 */
function doPost(e) {
  try {
    // Acquire a script lock to prevent duplicate rows if multiple people submit simultaneously
    const lock = LockService.getScriptLock();
    lock.waitLock(10000); // Wait up to 10 seconds for concurrent submissions

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Community Submissions");
    
    if (!sheet) {
      throw new Error("Tab 'Community Submissions' not found.");
    }

    // Parse the JSON payload sent from your website
    const data = JSON.parse(e.postData.contents);
    const idea = data.idea || "";
    const impact = data.impact || "";

    // Calculate the next Submission ID dynamically
    // Looks at the last row with data in Column A
    const lastRow = sheet.getLastRow();
    let nextNum = 1001;
    
    if (lastRow >= 6) {
      const lastIdValue = sheet.getRange(lastRow, 1).getValue().toString();
      const match = lastIdValue.match(/\d+/);
      if (match) {
        nextNum = parseInt(match[0], 10) + 1;
      } else {
        nextNum = 1000 + (lastRow - 5 + 1);
      }
    }
    const newId = "SUB-" + nextNum;

    // Current date in YYYY-MM-DD
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");

    // Append the row matching your sheet's column structure:
    // [Col A: Submission ID, Col B: Submission Date, Col C: Idea Title & Description, 
    //  Col D: Community Impact & Implementation, Col E: Status, Col F: Feasibility, Col G: Gautham's Notes!]
    sheet.appendRow([
      newId,
      today,
      idea,
      impact,
      "New",   // Default Status chip
      "",      // Feasibility left blank for your review
      ""       // Notes left blank for you
    ]);

    // Release the script lock
    lock.releaseLock();

    // Return success response to the website
    return ContentService
      .createTextOutput(JSON.stringify({ 
        status: "success", 
        submissionId: newId 
      }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ 
        status: "error", 
        message: error.toString() 
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
