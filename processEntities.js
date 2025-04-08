function processEntitiesForPlanning() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName('raw-data');
  const planningSheet = ss.getSheetByName('planning');

  if (!rawSheet || !planningSheet) {
    throw new Error("One or both sheets ('raw-data' or 'planning') are missing.");
  }

  // Clear only values from previous import, not data validation or formatting
  const lastRow = planningSheet.getLastRow();
  const lastCol = planningSheet.getLastColumn();
  if (lastRow > 1) {
    archivePlanningSheet();
    planningSheet.getRange(2, 1, lastRow - 1, lastCol).clear({ contentsOnly: true});
  }

  const rawData = rawSheet.getDataRange().getValues();

  // Column indicies for columns:
  // A, D, E, F, I, J, K, L, N, W, X, Y
  const selectedCols = [0, 3, 4, 5, 8, 9, 10, 11, 13, 22, 23, 24];

  // Filter out rows that are blank or obviously broken (e.g., if column A is blank)
  const filteredData = rawData.filter((row, i) => {
    if (i === 0) return true; // Keep headers
    return row[0] && row[0].toString().trim() !== "";
  });

  const processedData = filteredData.map(row => selectedCols.map(i => row[i]));

  // Write to planning sheet starting at A2 to preserve headers
  planningSheet.getRange(2, 1, processedData.length, selectedCols.length).setValues(processedData);
}

function archivePlanningSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const planningSheet = ss.getSheetByName('planning');

  if (!planningSheet) {
    throw new Error("Sheet 'planning' does not exist.");
  }

  // Get yesterday's date
  const yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);

  const formattedDate = Utilities.formatDate(yesterday, ss.getSpreadsheetTimeZone(), 'MMdd');
  const newSheetName = `planning-${formattedDate}`;

  // Delete the sheet if it already exists
  const existing = ss.getSheetByName(newSheetName);
  if (existing) {
    ss.deleteSheet(existing);
  }

  const archivedSheet = planningSheet.copyTo(ss);
  archivedSheet.setName(newSheetName);
  archivedSheet.hideSheet();

  
}


function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Planning Tools')
    .addItem('Process Entities', 'processEntitiesForPlanning')
    .addToUi();
}
