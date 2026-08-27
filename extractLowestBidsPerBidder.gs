function extractLowestBidsPerBidder() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sourceSheetName = 'AOB - As Calculated';
  const sourceSheet = ss.getSheetByName(sourceSheetName);

  if (!sourceSheet) {
    ui.alert(
      `Error: "${sourceSheetName}" sheet does not exist.\n\n` +
      'Please generate the calculated AOB first.'
    );
    return null;
  }

  const data = sourceSheet.getDataRange().getValues();
  if (data.length < 2) {
    ui.alert(`No bid data found in "${sourceSheetName}".`);
    return null;
  }

  const headers = data[0];
  const bidderNames = headers.slice(6);

  // Remove previously generated bidder sheets before rebuilding them.
  deleteBidderSheetsSA();

  // Build the LCB result first, then write each bidder's complete result in one batch.
  const lowestBidsPerBidder = {};

  data.slice(1).forEach(row => {
    const itemNo = row[0];
    const totalQuantity = toNumber(row[1]);
    const unit = row[2];
    const itemDescription = row[3];
    const unitCost = toNumber(row[4]);
    const totalCost = totalQuantity * unitCost;

    if (isNaN(unitCost)) return;

    // Only bids at or below the Unit Cost are eligible, consistent with Bid Ranking.
    const eligibleBids = [];

    for (let j = 6; j < row.length; j++) {
      const bidPrice = toNumber(row[j]);

      if (!isNaN(bidPrice) && bidPrice <= unitCost) {
        eligibleBids.push({
          bidderName: bidderNames[j - 6],
          bidPrice: bidPrice
        });
      }
    }

    if (eligibleBids.length === 0) return;

    // Find the lowest eligible bid.
    const lowestBid = Math.min(...eligibleBids.map(bid => bid.bidPrice));

    // IMPORTANT: retain ALL bidders whose bid equals the lowest bid.
    const tiedLowestBids = eligibleBids.filter(
      bid => bid.bidPrice === lowestBid
    );

    tiedLowestBids.forEach(({ bidderName, bidPrice }) => {
      if (!bidderName) return;

      if (!lowestBidsPerBidder[bidderName]) {
        lowestBidsPerBidder[bidderName] = [];
      }

      lowestBidsPerBidder[bidderName].push({
        itemNo: itemNo,
        totalQuantity: totalQuantity,
        unit: unit,
        itemDescription: itemDescription,
        unitCost: unitCost,
        totalCost: totalCost,
        bidPrice: bidPrice,
        totalBidPrice: totalQuantity * bidPrice
      });
    });
  });

  // Create one sheet per bidder represented in the LCB result.
  Object.keys(lowestBidsPerBidder).forEach(bidderName => {
    const sheet = ss.insertSheet(bidderName);
    const rows = lowestBidsPerBidder[bidderName];

    const output = [[
      'ITEM NO',
      'TOTAL QUANTITY',
      'UNIT',
      'ITEM DESCRIPTION',
      'UNIT COST',
      'TOTAL COST',
      'BID PRICE',
      'TOTAL BID PRICE'
    ]];

    rows.forEach(item => {
      output.push([
        item.itemNo,
        item.totalQuantity,
        item.unit,
        item.itemDescription,
        item.unitCost,
        item.totalCost,
        item.bidPrice,
        item.totalBidPrice
      ]);
    });

    sheet.getRange(1, 1, output.length, output[0].length).setValues(output);
    formatBidderSheet(sheet);
  });

  return lowestBidsPerBidder;
}

function toNumber(value) {
  if (typeof value === 'number') return value;
  if (value === null || value === '') return NaN;

  const parsed = parseFloat(String(value).replace(/,/g, ''));
  return isNaN(parsed) ? NaN : parsed;
}

function formatBidderSheet(sheet) {
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();

  if (lastRow === 0 || lastColumn === 0) return;

  const headerRange = sheet.getRange(1, 1, 1, lastColumn);
  headerRange
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  sheet.autoResizeColumns(1, 3);
  sheet.setColumnWidth(4, 300);
  sheet.getRange('D:D').setWrap(true);
  sheet.autoResizeColumns(5, 8);

  if (lastRow > 1) {
    sheet.getRange(2, 2, lastRow - 1, 1).setNumberFormat('#,##0.##');
    sheet.getRange(2, 5, lastRow - 1, 4)
      .setNumberFormat('#,##0.00_);(#,##0.00)');
  }

  sheet.getDataRange().setBorder(true, true, true, true, true, true);
  sheet.setFrozenRows(1);
}
