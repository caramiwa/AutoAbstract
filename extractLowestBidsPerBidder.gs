function extractLowestBidsPerBidder() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName('AOB - As Calculated');
if (!sourceSheet) {
  throw new Error('Sheet "AOB - As Calculated" not found. Please verify the sheet name.');
}


  // Check if the source sheet exists
  if (!sourceSheet) {
    SpreadsheetApp.getUi().alert(
      'Error: "AOB - As Calculated" sheet does not exist.\n\nPlease generate the calculated AOB first by duplicating the As Read sheet and removing disqualified bidders.'
    );
    return; // Stop execution
  }

  const sourceData = sourceSheet.getDataRange().getValues();

  // Get header row and bidder names
  const headers = sourceData[0];
  const bidderNames = headers.slice(6); // Assumes bidders start from column 7 (index 6)
  
  deleteBidderSheets(bidderNames);

  // Track which bidders have valid sheets
  const bidderSheets = {};

  const lowestBidsPerBidder = {};

  // Process each row from the source sheet
  sourceData.slice(1).forEach(row => {
    const itemNo = row[0];
    const totalQuantity = row[1];
    const unit = row[2];
    const itemDescription = row[3];
    const unitCost = parseFloat(row[4]);
    const totalCost = totalQuantity * unitCost;

    // Find the lowest bid for the item
    let lowestBid = null;
    let lowestBidderIndex = -1;

    for (let j = 6; j < row.length; j++) {
      const bidPrice = parseFloat(row[j]);
      if (!isNaN(bidPrice)) {
        if (lowestBid === null || bidPrice < lowestBid) {
          lowestBid = bidPrice;
          lowestBidderIndex = j - 6;
        }
      }
    }

    // If a valid lowest bid was found, create or use the corresponding bidder's sheet
    if (lowestBidderIndex !== -1) {

      const bidderName = bidderNames[lowestBidderIndex];
      if (!lowestBidsPerBidder[bidderName]) {
        lowestBidsPerBidder[bidderName] = [];
      }

      lowestBidsPerBidder[bidderName].push({
        itemNo,
        itemDescription,
        unitCost,
        bidPrice: lowestBid
      });


      if (!bidderSheets[bidderName]) {
        let bidderSheet = ss.getSheetByName(bidderName);
        if (bidderSheet) {
          bidderSheet.clear(); // Clear if exists
        } else {
          bidderSheet = ss.insertSheet(bidderName); // Create new if doesn't exist
        }
        // Set headers for each bidder sheet
        bidderSheet.appendRow(['ITEM NO', 'TOTAL QUANTITY', 'UNIT', 'ITEM DESCRIPTION', 'UNIT COST', 'TOTAL COST', 'BID PRICE', 'TOTAL BID PRICE']);
        formatHeader(bidderSheet);
        bidderSheets[bidderName] = bidderSheet;
      }

      const totalBidPrice = totalQuantity * lowestBid;
      const bidderSheet = bidderSheets[bidderName];
      bidderSheet.appendRow([itemNo, totalQuantity, unit, itemDescription, unitCost.toFixed(2), totalCost.toFixed(2), lowestBid.toFixed(2), totalBidPrice.toFixed(2)]);
    }
  });

  // Adjust column widths and apply accounting format
  for (const bidder in bidderSheets) {
    const sheet = bidderSheets[bidder];
    sheet.autoResizeColumns(1, 3); // Adjust ITEM NO, TOTAL QUANTITY, UNIT
    sheet.setColumnWidth(4, 300); // ITEM DESCRIPTION column width
    sheet.getRange('D:D').setWrap(true); // Wrap text in ITEM DESCRIPTION
    sheet.autoResizeColumns(5, 8); // Adjust UNIT COST, TOTAL COST, BID PRICE, TOTAL BID PRICE
    sheet.getRange('E:H').setNumberFormat('#,##0.00_);(#,##0.00)'); // Accounting format
    applyBorders(sheet);
  }
  return lowestBidsPerBidder;

}

function formatHeader(sheet) {
  const headerRange = sheet.getRange(1, 1, 1, 8);
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
}

function applyBorders(sheet) {
  const range = sheet.getDataRange();
  range.setBorder(true, true, true, true, true, true);
}
