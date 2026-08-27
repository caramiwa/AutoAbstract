function HighlightLCRB() {
//  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
//  const dataRange = sheet.getDataRange();
//  const data = dataRange.getValues();


  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName('AOB - Responsive');

  // Check if the source sheet exists
  if (!sourceSheet) {
    SpreadsheetApp.getUi().alert(
      'Error: "AOB - As Responsive" sheet does not exist.\n\nPlease generate the calculated AOB first by duplicating the As Read sheet and removing disqualified bidders.'
    );
    return;
  }

  const sheet = sourceSheet;
  const dataRange = sheet.getDataRange();
  const data = dataRange.getValues();

  //Clear formatting
  // sheet.getRange(2, 1, sheet.getMaxRows() - 1, sheet.getMaxColumns()).clearFormat();
  
  // Clear background colors from row 2 onwards
  sheet.getRange(2, 1, sheet.getMaxRows() - 1, sheet.getMaxColumns()).setBackground(null);


  // Define colors for the ranking
  const rankColors = ['#80b463', '#27AAE1', '#FFD400', '#F79552', '#01B9BB', '#C08497','#89b1bc','#bcf4de','#ffd6e0','#b3d2a1','#7dcced','#ffee99','#fabf97','#ccf1f1','#d9b5c1','#b8d0d7','#d7f8eb','#ffe6ec'];

 

  // Loop through each row, skipping the header row
  for (let i = 1; i < data.length; i++) {
    const quantity = data[i][1];
    const unitCost = data[i][4];
    const totalCost = quantity * unitCost;
    sheet.getRange(i + 1, 6).setValue(totalCost); // Set Total Cost in Column 6

    let bids = [];
    let noEntries = true;

    // Collect bids for ranking
    for (let j = 6; j < data[i].length; j++) {
      const bid = data[i][j];
      if (bid) {
        noEntries = false;
        bids.push({ value: bid, index: j });
      }
    }
   
    // Apply red cell highlight if bid value exceeds Unit Cost
    bids.forEach(bid => {
      if (bid.value > unitCost) {
        sheet.getRange(i + 1, bid.index + 1).setBackground('#EF404A');
      }
    });

    // Sort bids for ranking
    bids.sort((a, b) => a.value - b.value);

    // Apply ranking colors
   /* for (let k = 0; k < bids.length && k < rankColors.length; k++) {
      sheet.getRange(i + 1, bids[k].index + 1).setBackground(rankColors[k]);
    }*/
    let colorIndex = 0;
for (let k = 0; k < bids.length && colorIndex < rankColors.length; ) {
  let tiedBids = [bids[k]];
  
  // Group tied bids
  while (
    k + tiedBids.length < bids.length &&
    bids[k].value === bids[k + tiedBids.length].value
  ) {
    tiedBids.push(bids[k + tiedBids.length]);
  }

  // Assign the same color to all tied bids
  tiedBids.forEach(bid => {
    sheet.getRange(i + 1, bid.index + 1).setBackground(rankColors[colorIndex]);
  });

  // Move to next group
  k += tiedBids.length;
  colorIndex++;
}


    // Apply row highlighting based on conditions
    if (noEntries) {
      // Highlight the entire row in red if there are no entries from Column G onward
      sheet.getRange(i + 1, 1, 1, sheet.getLastColumn()).setBackground('#EF404A');
    } else if (bids.length > 0 && bids.every(bid => bid.value > unitCost)) {
      // Highlight the entire row in purple if all bids exceed the Unit Cost
      sheet.getRange(i + 1, 1, 1, sheet.getLastColumn()).setBackground('#BBB8DC');
    }
  }
}
