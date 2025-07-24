// This script automatically renames a Google Drive file based on your spreadsheet.

// --- CONFIGURATION ---
// Set your sheet name and column numbers here.
const SHEET_NAME = "Creative Master";
const LINK_COL = 2; // Column B for the Drive Link
const NEW_NAME_COL = 13; // Column M for the new file name
const TRIGGER_COL = 14; // Column N for the "Rename?" dropdown
const STATUS_COL = 15; // Column O for the status feedback

// --- SCRIPT LOGIC ---
// This function runs automatically every time you edit the sheet.
function onEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();
  const editedRow = range.getRow();
  const editedCol = range.getColumn();
  const newValue = e.value;

  // 1. Check if the edit was on the correct sheet, in the trigger column, and the value is "Yes".
  if (sheet.getName() === SHEET_NAME && editedCol === TRIGGER_COL && newValue === "Yes") {
    
    // Get the cells we need for this row.
    const linkCell = sheet.getRange(editedRow, LINK_COL);
    const newNameCell = sheet.getRange(editedRow, NEW_NAME_COL);
    const statusCell = sheet.getRange(editedRow, STATUS_COL);
    const triggerCell = range; // The cell that was just edited to "Yes"

    const driveLink = getLinkUrl(linkCell);
    const newName = newNameCell.getValue();

    // 2. Make sure the link and new name cells are not empty.
    if (!driveLink || !newName) {
      statusCell.setValue("ERROR: Drive Link or New Name is empty.");
      return;
    }

    try {
      // 3. Extract the File ID from the Google Drive link.
      const fileId = extractFileIdFromLink(driveLink);
      if (!fileId) {
        throw new Error("Invalid Google Drive link.");
      }

      // 4. Get the file from Google Drive.
      const file = DriveApp.getFileById(fileId);
      const originalName = file.getName();
      
      // 5. Get the original file extension.
      const extension = originalName.includes('.') ? originalName.substring(originalName.lastIndexOf('.')) : '';
      
      // 6. Combine the new name with the original extension and rename the file.
      const finalNewName = newName + extension;
      file.setName(finalNewName);
      
      // 7. Update the sheet to show the task is complete.
      triggerCell.setValue("DONE");
      statusCell.setValue("Renamed on " + new Date().toLocaleString());

    } catch (error) {
      // If any errors happen, write the error message to the status column.
      statusCell.setValue("ERROR: " + error.message);
      Logger.log(error); // Log the full error for debugging.
    }
  }
}

/**
 * A helper function to get the file ID from various Google Drive URL formats.
 * @param {string} url The Google Drive URL.
 * @returns {string|null} The extracted file ID or null if not found.
 */
function extractFileIdFromLink(url) {
  if (!url || typeof url !== 'string') {
    return null;
  }
  let id = null;
  // Standard 'files/d/' URL
  let match = url.match(/[-\w]{25,}/);
  if (match) {
    id = match[0];
  }
  return id;
}

/**
 * Returns the first hyperlink found in a cell's rich text content.
 * Falls back to the plain cell value if no link is present.
 * @param {GoogleAppsScript.Spreadsheet.Range} cell The cell to inspect.
 * @returns {string} The URL or an empty string if none found.
 */
function getLinkUrl(cell) {
  const rich = cell.getRichTextValue();
  if (rich) {
    const runs = rich.getRuns();
    for (let i = 0; i < runs.length; i++) {
      const url = runs[i].getLinkUrl();
      if (url) {
        return url;
      }
    }
    if (rich.getLinkUrl()) {
      return rich.getLinkUrl();
    }
  }
  const plain = cell.getValue();
  return typeof plain === 'string' ? plain : '';
}

/**
 * Creates an installable onEdit trigger so the script has permission
 * to access Drive and rename files. Run this once manually.
 */
function setupTrigger() {

  ScriptApp.newTrigger('onEdit')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();
}

// Add a custom menu so the user can easily install the trigger
function onOpen(e) {
  SpreadsheetApp.getUi()
    .createMenu('Rename Tools')
    .addItem('Install Rename Trigger', 'setupTrigger')
    .addToUi();

  const ss = SpreadsheetApp.getActive();
  const triggers = ScriptApp.getProjectTriggers();
  const alreadyInstalled = triggers.some(t =>
    t.getHandlerFunction() === 'onEdit' &&
    t.getTriggerSourceId() === ss.getId()
  );
  if (!alreadyInstalled) {
    ScriptApp.newTrigger('onEdit')
      .forSpreadsheet(ss)
      .onEdit()
      .create();
  }
}
