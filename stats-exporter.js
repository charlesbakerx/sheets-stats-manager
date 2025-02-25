function exportStatsToJson() {

    // Constants
    var PS = ["P1", "P2", "P3", "P4", "P5", "P6"];
    var sourceSheetName = "<sourceSheetName>"
    var statsSheetName = "<statsSheetName>"
    //
  
    var sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sourceSheetName);
    var date = sourceSheet.getRange(1, 1).getValue(); // Assuming date is in A1
  
    var jsonData = {
      "date": date,
      "groups": []
    };
  
    for (var i = 0; i < PS.length; i++) {
      let P = PS[i];
  
      // Calculate column positions based on initial column
      let baseCol = 2 + (i * 2);
  
      let O1 = sourceSheet.getRange(1, baseCol).getValue() || "";
      let O2 = sourceSheet.getRange(1, baseCol + 1).getValue() || "";
  
      let statsRange = sourceSheet.getRange(2, baseCol, 10, 1).getValues();
  
      // Convert stats to structured JSON
      let statsArray = statsRange.map((row, index) => {
        let value = row[0] || 0; // Ensure !null
        return {[index + 1]: value}; 
        });
  
      // Add to JSON structure
      jsonData.groups.push({
        "P": P,
        "O1": O1,
        "O2": O2,
        "values": statsArray
      });
    }
  
    // Convert to JSON string
    var jsonString = JSON.stringify(jsonData, null, 2);
  
    // Save JSON in stats sheet
    var statsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(statsSheetName);
    if (!statsSheet) {
      statsSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(statsSheetName);
    }
  
    var nextRow = statsSheet.getLastRow() + 1;
    statsSheet.getRange(nextRow,1).setValue(date)
    statsSheet.getRange(nextRow,2).setValue(jsonString)
  }