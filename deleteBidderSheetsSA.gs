function deleteBidderSheetsSA() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const sourceSheet = ss.getSheetByName('AOB - As Read');
  const sourceData = sourceSheet.getDataRange().getValues();

  // Get header row and bidder names
  const headers = sourceData[0];
  const bidderNames = headers.slice(6); // Assumes bidders start from column 7 (index 6)

  // Iterate through all sheets and delete the ones with names matching bidder names
  sheets.forEach(sheet => {
    if (bidderNames.includes(sheet.getName())) {
      ss.deleteSheet(sheet);
    }
  });
}
