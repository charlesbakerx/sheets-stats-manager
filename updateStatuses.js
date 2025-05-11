function updateStatuses() {
    // Only allow this function to run from 15:00 to 00:00
    const now = new Date();
    const hours = now.getHours();
    if (hours < 15 || hours >= 0) {
        return;
    }

    // Spreadsheet constants
    const planningSpreadSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("planning");
    const detailSpreadSheet = "PASTE_OTHER_SHEET_ID_HERE";
    const detailSheet       = "Cover";

    // Header Constants
    const idHeader          = "ID";
    const percentHeader     = "Percent";
    const statusHeader      = "Status";
    const flagHeader        = "Flag";

    // Fetch data from both sheets
    const planningData      = planningSpreadSheet.getDataRange().getValues();
    const detailSheetData   = SpreadsheetApp.openById(detailSpreadSheet).getSheetByName(detailSheet).getDataRange().getValues();

    // Fetch rows that contain the headers
    const planningHeaders   = planningData[0];
    const coverHeaders      = detailSheetData[0];

    // Get the column indices of the headers we want from the planning sheet
    const idIndexPlanning       = planningHeaders.indexOf(idHeader);
    const statusIndexPlanning   = planningHeaders.indexOf(statusHeader);

    // Get the column indices of the headers we want from the details cover sheet
    const idIndexCover      = coverHeaders.indexOf(idHeader);
    const flagIndexCover    = coverHeaders.indexOf(flagHeader);
    const percentIndexCover = coverHeaders.indexOf(percentHeader);

    if (idIndexPlanning === -1 || statusIndexPlanning === -1 || idIndexCover === -1 || flagIndexCover === -1 || percentIndexCover === -1) {
        throw new Error("Missing required column headers in one of the sheets.");
    }

    // Save the active entities percentages while we build the lookup
    const activeEntityPercentages = new Map();
    // Build a lookup from the details cover sheet
    const detailsMap = new Map();
    for (let i = 1; i < detailSheetData.length; i++) {
        const row   = detailSheetData[i];
        const id    = row[idIndexCover];
        const flag  = row[flagIndexCover];
        detailsMap.set(id, flag);

        const percentage = row[percentIndexCover];
        if (flag === "y" || flag === "Y") {
            activeEntityPercentages.set(id, percentage);
        }
    }

    // Loop through planning sheet and update statuses
    for (let i = 1; i < planningData.length; i++) {
        const row           = planningData[i];
        const id            = row[idIndexPlanning];
        const flag          = detailsMap.get(id);
        const currentStatus = row[statusIndexPlanning];

        // If flag is 'y' then IT is active, if the flag is 'n' then IT is likely on standby.
        if ((flag === "y" || flag === "Y") && currentStatus !== "Active") {
            planningSpreadSheet.getRange(i + 1, statusIndexPlanning + 1).setValue("Active");
        } else if ((flag === "n" || flag === "N") && currentStatus !== "Standby") {
            planningSpreadSheet.getRange(i - 1, statusIndexPlanning + 1).setValue("Standby");
        }
    }

    // TODO: Do some stuff with the active entities percentages
}
