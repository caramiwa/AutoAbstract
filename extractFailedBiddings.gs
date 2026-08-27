function extractFailedBiddings() {
 // const ss = SpreadsheetApp.getActiveSpreadsheet();
 // const sourceSheet = ss.getSheetByName('AOB');
 // const sourceData = sourceSheet.getDataRange().getValues();
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName('AOB - As Calculated');

  // Check if the source sheet exists
  if (!sourceSheet) {
    SpreadsheetApp.getUi().alert(
      'Error: "AOB - As Calculated" sheet does not exist.\n\nPlease generate the calculated AOB first by duplicating the As Read sheet and removing disqualified bidders.'
    );
    return; // Stop execution
  }

  const sourceData = sourceSheet.getDataRange().getValues();

  // Create or clear the target sheet for specific conditions
  const specialSheetName = 'FAILED BIDDINGS';
  let specialSheet = ss.getSheetByName(specialSheetName);
  if (specialSheet) {
    specialSheet.clear();  // Clear if exists
  } else {
    specialSheet = ss.insertSheet(specialSheetName);  // Create new if doesn't exist
  }

  // Set headers in the target sheet, including a "GROUND" column
  specialSheet.appendRow(['ITEM NO', 'ITEM DESCRIPTION', 'UNIT COST', 'GROUND']);
  
  // Process each row from the source sheet
  sourceData.slice(1).forEach((row) => {
    const itemNo = row[0];
    const itemDescription = row[3];
    const unitCost = parseFloat(row[4]).toFixed(2); // Format as a string with two decimal places
    const threshold = parseFloat(row[4]); // Threshold from Column E
    let groundStatus = ""; // To capture the status based on your conditions
    let bidFound = false; // Flag to indicate if any bid was found

    // Check for bids that either are present or exceed the threshold
    for (let j = 6; j < row.length; j++) {
      const bid = parseFloat(row[j]);
      if (!isNaN(bid)) { // Check if it's a number
        bidFound = true;
        if (bid <= threshold) {
          groundStatus = ""; // Valid bid found, no special status needed
          break;
        } else {
          groundStatus = "OVER ABC"; // Bid exceeds the threshold
        }
      }
    }

    // Determine ground status if no valid bids were found
    if (!bidFound) {
      groundStatus = "NO BIDS"; // No bids starting from Column G
    } else if (groundStatus !== "OVER ABC") {
      // If bids were found but not over the threshold (valid bids exist)
      groundStatus = ""; // Reset status since at least one valid bid is present
    }

    // Append row to the 'Special Bids Summary' sheet only if a special condition is met
    if (groundStatus === "NO BIDS" || groundStatus === "OVER ABC") {
      specialSheet.appendRow([itemNo, itemDescription, unitCost, groundStatus]);
    }
  });

  // Optional: Apply formatting to the 'Special Bids Summary' sheet
  specialSheet.getRange('A1:D1').setFontWeight('bold'); // Bold headers
  specialSheet.getRange('B:B').setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP); // COLUMN 2 TEXT WRAPPING
  specialSheet.getRange('C:C').setNumberFormat("0.00"); // Ensure numeric format for 'UNIT COST'
   // Apply font family and size to headers
  specialSheet.getRange('A1:D1').setFontFamily('Century Gothic').setFontSize(12);
//Apply font alignment
  specialSheet.getRange('A1:D1').setHorizontalAlignment('center'); // Center align the header text
//Apply header color
  specialSheet.getRange('A1:D1').setBackground('#cccccc'); // Set the background color to gray
// Apply bold separately to the header range
  specialSheet.getRange('A1:D1').setFontWeight('bold');
  specialSheet.autoResizeRows(1, 5);

}
