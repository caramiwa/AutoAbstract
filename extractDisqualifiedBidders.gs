function extractDisqualifiedBidders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetRead = ss.getSheetByName("AOB - As Read");
  const sheetCalc = ss.getSheetByName("AOB - As Calculated");

  if (!sheetRead || !sheetCalc) {
    Logger.log("Missing sheet(s) for Disqualified extraction.");
    return null;
  }

  const dataRead = sheetRead.getDataRange().getValues();
  const dataCalc = sheetCalc.getDataRange().getValues();

  const headers = dataRead[0];
  const bidderNames = headers.slice(6); // Bidders start from col G

  const disqualifiedData = {};

  for (let i = 1; i < dataRead.length; i++) {
    const rowRead = dataRead[i];
    const rowCalc = dataCalc[i];

    const itemNo = rowRead[0];
    const itemDesc = rowRead[3];
    const unitCost = parseFloat(rowRead[4]);

    for (let j = 6; j < rowRead.length; j++) {
      const bidRead = parseFloat(rowRead[j]);
      const bidCalc = rowCalc[j];

      if (!isNaN(bidRead) && (bidCalc === "" || bidCalc === null)) {
        const bidderName = bidderNames[j - 6];
        if (!disqualifiedData[bidderName]) disqualifiedData[bidderName] = [];

        disqualifiedData[bidderName].push([
          itemNo,
          itemDesc,
          unitCost.toFixed(2),
          bidRead.toFixed(2),
          "", "", ""
        ]);

        Logger.log(`Disqualified → Bidder: ${bidderName}, Item No: ${itemNo}, Bid: ${bidRead}`);
      }
    }
  }

  if (Object.keys(disqualifiedData).length === 0) {
    Logger.log("No disqualified bidders found.");
    return null;
  }

  Logger.log("✅ Disqualified data collected:");
  Logger.log(disqualifiedData);

  return disqualifiedData;
}
