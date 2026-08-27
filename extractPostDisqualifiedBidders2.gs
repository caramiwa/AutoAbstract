function extractPostDisqualifiedBidders2() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetCalc = ss.getSheetByName('AOB - As Calculated');
  const sheetResp = ss.getSheetByName('AOB - Responsive');

  const calcData = sheetCalc.getDataRange().getValues();
  const respData = sheetResp.getDataRange().getValues();

  const headers = calcData[0];
  const bidderNames = headers.slice(6); // bidder columns start at column G

  const postDisqualified = {};

  for (let i = 1; i < calcData.length; i++) {
    const rowCalc = calcData[i];
    const rowResp = respData[i];

    const itemNo = rowCalc[0];
    const itemDesc = rowCalc[3];

    let lowestBid = null;
    let lcbIndex = -1;

    // Find LCB from Calculated sheet
    for (let j = 6; j < rowCalc.length; j++) {
      const bid = parseFloat(rowCalc[j]);
      if (!isNaN(bid) && (lowestBid === null || bid < lowestBid)) {
        lowestBid = bid;
        lcbIndex = j;
      }
    }

    if (lcbIndex !== -1) {
      const lcbBidder = headers[lcbIndex];

      // Check if this bidder is blank in Responsive sheet = Post-disqualified
      const respBid = rowResp[lcbIndex];
      if (respBid === "" || isNaN(parseFloat(respBid))) {
        if (!postDisqualified[lcbBidder]) postDisqualified[lcbBidder] = [];

        postDisqualified[lcbBidder].push([
          itemNo,
          itemDesc,
          "", // Ground
          "", // RR
          ""  // Resolution
        ]);
      }
    }
  }

  const tables = [];
  for (const bidder in postDisqualified) {
//Logger.log(`Mapping post-disqualified: Row ${i + 1}, Bidder column ${j}, Bidder name: "${bidder}"`);

    tables.push({
      header: bidder,
      subHeaders: ["Item No.", "Item Description", "Ground", "RR", "Resolution"],
      rows: postDisqualified[bidder]
    });
  }

  return tables.length > 0 ? tables : null;
}


function extractPostDisqualifiedBidders2() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('AOB - Responsive');
  const data = sheet.getDataRange().getValues();

  const headers = data[0];
  const bidderNames = headers.slice(6); // Assume bidders start at column G
  const postDisqualifiedData = {};

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const itemNo = row[0];
    const itemDescription = row[3];

    // Track if the entire row is blank for all bidders (post-disqualified)
    let allBlank = true;
    for (let j = 6; j < row.length; j++) {
      if (row[j] !== "" && row[j] !== null && !isNaN(row[j])) {
        allBlank = false;
        break;
      }
    }

    if (allBlank) {
      // Row is post-disqualified; record it for each bidder
      bidderNames.forEach(bidder => {
        if (!postDisqualifiedData[bidder]) {
          postDisqualifiedData[bidder] = [];
        }

        postDisqualifiedData[bidder].push([
          itemNo,
          itemDescription,
          "",  // Ground (manual)
          "",  // RR (manual)
          ""   // Resolution (manual)
        ]);
      });
    }
  }

  // Format the structure for generateBACResolution
  const tables = [];
  for (const bidder in postDisqualifiedData) {
    tables.push({
      header: bidder,
      subHeaders: ["Item No.", "Item Description", "Ground", "RR", "Resolution"],
      rows: postDisqualifiedData[bidder]
    });
  }

  return tables.length > 0 ? tables : null;
}
