function extractSortBids() { //FINAL
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName('AOB - As Calculated');

  const ui = SpreadsheetApp.getUi();

  // 📌 Instruction alert before anything else
  const initialPrompt = ui.alert(
    'Ensure that the "Disqualified bidders have been removed from the AOB - As Calculated to generate accurate Ranking',
    ui.ButtonSet.OK
  );
  

  // Check if the source sheet exists
  if (!sourceSheet) {
    SpreadsheetApp.getUi().alert(
      'Error: "AOB - As Calculated" sheet does not exist.\n\nPlease generate the calculated AOB first by duplicating the As Read sheet and removing disqualified bidders.'
    );
    return; // Stop execution
  }

  const sourceData = sourceSheet.getDataRange().getValues();

  // Determine the maximum number of bids across all rows
  let maxBids = 0;
  sourceData.slice(1).forEach((row) => {
    let bidCount = 0;
    for (let j = 6; j < row.length; j++) {
      const bid = parseFloat(row[j]);
      if (!isNaN(bid)) {
        bidCount++;
      }
    }
    if (bidCount > maxBids) {
      maxBids = bidCount;
    }
  });

  // Create or clear the target sheet
  const targetSheetName = 'Bid Ranking';
  let targetSheet = ss.getSheetByName(targetSheetName);
  if (targetSheet) {
    targetSheet.clear();  // Clear if exists
  } else {
    targetSheet = ss.insertSheet(targetSheetName);  // Create new if doesn't exist
  }
  targetSheet.activate(); // Ensure it's the active sheet
  ss.moveActiveSheet(3);  // Always moves the correct one

  // Create headers based on the maximum number of bids
  const headers = ['ITEM NO', 'ITEM DESCRIPTION', 'UNIT COST'];
  for (let i = 1; i <= maxBids; i++) {
    headers.push(`RANK ${i}`);
  }
  targetSheet.appendRow(headers);

  // Set text wrapping for 'ITEM DESCRIPTION' and LCB columns
  targetSheet.getRange('B:B').setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP); // For ITEM DESCRIPTION
  targetSheet.getRange(1, 4, targetSheet.getMaxRows(), maxBids).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP); // For LCB columns

  const columnColors = ['#80b463', '#27AAE1', '#FFD400', '#F79552', '#01B9BB', '#C08497','#89b1bc','#bcf4de','#ffd6e0','#b3d2a1','#7dcced','#ffee99','#fabf97','#ccf1f1','#d9b5c1','#b8d0d7','#d7f8eb','#ffe6ec'];

  // Process each row from the source sheet
  sourceData.slice(1).forEach((row) => {
    const itemNo = row[0];
    const itemDescription = row[3];
    const unitCost = parseFloat(row[4]).toFixed(2);
    const threshold = parseFloat(row[4]);
    let bids = [];
    let hasValidBid = false;

    // Extract bids that do not exceed the threshold in Column E and are not empty
    for (let j = 6; j < row.length; j++) {
      const bid = parseFloat(row[j]);
      if (!isNaN(bid) && bid <= threshold) {
        hasValidBid = true; // Mark as having at least one valid bid
        bids.push({ text: `${sourceData[0][j]} - \n${bid.toFixed(2)}`, value: bid });
      }
    }

    if (hasValidBid) {
      // Sort bids by their value if there's at least one valid bid
      bids.sort((a, b) => a.value - b.value);

      // Format sorted bids for insertion
      //const formattedBids = bids.map(bid => bid.text);
      let formattedBids = [];
let rank = 1;
for (let i = 0; i < bids.length; ) {
  let tiedGroup = [bids[i]];
  
  // Group all tied bids
  while (
    i + tiedGroup.length < bids.length &&
    bids[i].value === bids[i + tiedGroup.length].value
  ) {
    tiedGroup.push(bids[i + tiedGroup.length]);
  }

  // Join tied bids in one rank column, line-separated
  const combinedText = tiedGroup.map(bid => bid.text).join('\n\n');
  formattedBids.push(combinedText);

  // Move to next group
  i += tiedGroup.length;
  rank++;
}


      // Insert into the target sheet
      targetSheet.appendRow([itemNo, itemDescription, unitCost, ...formattedBids]);
    }
  });
// Remove autocolors for printout - June 3, 2025
//Revived autocolors - June 6, 2024
  // Apply colors to the columns starting from the 4th column
  for (let i = 0; i < maxBids; i++) {
    const color = columnColors[i % columnColors.length];
    targetSheet.getRange(2, 4 + i, targetSheet.getLastRow()-1, 1).setBackground(color);
  }

  // TEXT FORMATTING

  // Apply font family and size to headers
  targetSheet.getRange('A1:' + getColumnLetter(3 + maxBids) + '1').setFontFamily('Century Gothic').setFontSize(12);
  // Apply font alignment
  targetSheet.getRange('A1:' + getColumnLetter(3 + maxBids) + '1').setHorizontalAlignment('center'); // Center align the header text
  // Apply header color
  targetSheet.getRange('A1:' + getColumnLetter(3 + maxBids) + '1').setBackground('#cccccc'); // Set the background color to gray
  // Apply bold separately to the header range
  targetSheet.getRange('A1:' + getColumnLetter(3 + maxBids) + '1').setFontWeight('bold');
  //Freeze first row
  targetSheet.setFrozenRows(1);


  // Apply font family and size to data rows
  targetSheet.getRange('A2:' + getColumnLetter(3 + maxBids) + targetSheet.getLastRow()).setFontFamily('Century Gothic').setFontSize(10);
}

// Helper function to get column letter from index
function getColumnLetter(columnIndex) {
  let temp, letter = '';
  while (columnIndex > 0) {
    temp = (columnIndex - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    columnIndex = (columnIndex - temp - 1) / 26;
  }
  return letter;
}
