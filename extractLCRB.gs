function extractLCRB() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('AOB - Responsive');

  if (!sheet) {
    SpreadsheetApp.getUi().alert('Sheet "AOB - Responsive" not found.');
    return {};
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);

  const bidderNames = headers.slice(6); // Bidder columns start at index 6
  const lcrbData = {};

  rows.forEach(row => {
    const itemNo = parseInt(row[0]); // Whole number
    const totalQuantity = parseInt(row[1]); // Whole number
    const unit = row[2];
    const itemDescription = row[3];
    const unitCost = parseFloat(row[4]);

    let lowestBid = null;
    let winningBidderIndex = -1;

    for (let i = 6; i < row.length; i++) {
      const bid = parseFloat(row[i]);
      if (!isNaN(bid)) {
        if (lowestBid === null || bid < lowestBid) {
          lowestBid = bid;
          winningBidderIndex = i - 6;
        }
      }
    }

    if (winningBidderIndex !== -1) {
      const bidderName = bidderNames[winningBidderIndex];
      if (!lcrbData[bidderName]) {
        lcrbData[bidderName] = [];
      }

      lcrbData[bidderName].push({
        itemNo,
        totalQuantity,
        unit,
        itemDescription,
        unitCost,
        bidPrice: lowestBid,
        totalBidPrice: (lowestBid * totalQuantity)
      });
    }
  });

  return lcrbData;
}
