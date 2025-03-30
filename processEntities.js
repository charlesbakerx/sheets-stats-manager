function processEntitiesForPlanning() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName('raw-data');
  const planningSheet = ss.getSheetByName('planning');

  // Clear only values from previous import, not data validation or formatting
  planningSheet.getRange('A2:Z').clear({ contentsOnly: true });

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
