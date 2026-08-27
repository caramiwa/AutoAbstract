function generateNoticeOfAward() {

  // 🔴 STEP 1: CHECK FOR TIED LCRBs FIRST
  const tiedData = getTiedLCRBData();

  if (tiedData.length > 0) {
    generateNoticeToDrawLots(tiedData);

    SpreadsheetApp.getUi().alert(
      'Tied LCRBs detected.\n\n' +
      'Notice of Award generation has been BLOCKED.\n' +
      'A Notice to Draw Lots has been created.'
    );

    return; // ⛔ STOP EXECUTION
  }

  // 🟢 STEP 2: PROCEED WITH NOA
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('AOB - Responsive');
  const data = sheet.getDataRange().getValues();

  const headers = data[0];
  const bidderNames = headers.slice(6);

  const folderName = 'Notice of Award';
  const parentFolder = DriveApp.getRootFolder();

  let folder = parentFolder.getFoldersByName(folderName).hasNext()
    ? parentFolder.getFoldersByName(folderName).next()
    : parentFolder.createFolder(folderName);

  // 🧠 LCRB GROUPING
  const lowestBids = {};

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    let lowestBid = null;
    let lowestBidderIndex = -1;

    for (let j = 6; j < row.length; j++) {
      const bid = parseFloat(row[j]);

      if (!isNaN(bid) && (lowestBid === null || bid < lowestBid)) {
        lowestBid = bid;
        lowestBidderIndex = j - 6;
      }
    }

    if (lowestBidderIndex >= 0) {
      const bidderName = bidderNames[lowestBidderIndex];

      if (!lowestBids[bidderName]) lowestBids[bidderName] = [];

      lowestBids[bidderName].push({
        itemNo: row[0],
        totalQty: parseInt(row[1]),
        unit: row[2],
        itemDesc: row[3],
        unitCost: parseFloat(row[4]),
        totalCost: parseFloat(row[5]),
        bidPrice: parseFloat(row[6 + lowestBidderIndex])
      });
    }
  }

  const templateId = '18IPnTrz45KqaPTFNH-P7y5FhUx8mQDIw24aMpZUUeK4';

  // 📄 GENERATE NOA PER BIDDER
  for (const bidder in lowestBids) {
    const targetFileName = `${bidder} - Notice of Award`;

    // Remove existing file with same name
    const existingFiles = folder.getFilesByName(targetFileName);
    while (existingFiles.hasNext()) {
      existingFiles.next().setTrashed(true);
    }

    const docFile = DriveApp.getFileById(templateId).makeCopy(targetFileName, folder);

    Utilities.sleep(500); // ✅ prevent "document inaccessible"

    const docId = docFile.getId();
    const body = DocumentApp.openById(docId).getBody();

    // Replace bidder name
    const nameMatch = body.findText('BIDDER_NAME_PLACEHOLDER');
    if (nameMatch) {
      nameMatch.getElement().asText()
        .replaceText('BIDDER_NAME_PLACEHOLDER', bidder);
    }

    // Build table
    let totalBidPriceSum = 0;

    const table = [[
      'ITEM NO', 'TOTAL QUANTITY', 'UNIT',
      'ITEM DESCRIPTION', 'UNIT COST',
      'TOTAL COST', 'BID PRICE', 'TOTAL BID PRICE'
    ]];

    lowestBids[bidder].forEach(item => {
      const totalBidPrice = item.totalQty * item.bidPrice;
      totalBidPriceSum += totalBidPrice;

      table.push([
        Math.round(item.itemNo).toLocaleString(),
        Math.round(item.totalQty).toLocaleString(),
        item.unit,
        item.itemDesc,
        item.unitCost.toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2}),
        item.totalCost.toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2}),
        item.bidPrice.toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2}),
        totalBidPrice.toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2})
      ]);
    });

    // Insert table
    const tablePlaceholder = body.findText('{{LCRB Table}}');
    if (tablePlaceholder) {
      const paragraph = tablePlaceholder.getElement().getParent();
      const index = body.getChildIndex(paragraph);
      body.removeChild(paragraph);
      body.insertTable(index, table);
    }

    // Replace total amount
    const totalFormatted = `₱${totalBidPriceSum.toLocaleString(undefined, { minimumFractionDigits: 2 })}`;
    const totalMatch = body.findText('TOTALCONTRACTAMOUNT_PLACEHOLDER');
    if (totalMatch) {
      totalMatch.getElement().asText()
        .replaceText('TOTALCONTRACTAMOUNT_PLACEHOLDER', totalFormatted);
    }

    // Replace amount in words
    const amountInWords = convertNumberToWords(totalBidPriceSum);
    const wordsMatch = body.findText('AMOUNTINWORDS_PLACEHOLDER');
    if (wordsMatch) {
      wordsMatch.getElement().asText()
        .replaceText('AMOUNTINWORDS_PLACEHOLDER', amountInWords + ' Only');
    }

    Logger.log(`💌 Notice of Award for ${bidder}: https://docs.google.com/document/d/${docId}`);
  }
}

function generateNoticeToDrawLots(tiedData) {
  const templateId = '1k8B1YvGs73UShRV7HLyOJwVgDnFO0as9edwOTA9_ZFI';
  const folderName = 'Notice of Award';

  const parentFolder = DriveApp.getRootFolder();
  let folder = parentFolder.getFoldersByName(folderName).hasNext()
    ? parentFolder.getFoldersByName(folderName).next()
    : parentFolder.createFolder(folderName);

  const fileName = 'Notice to Draw Lots';

  // Remove existing file with same name
  const existingFiles = folder.getFilesByName(fileName);
  while (existingFiles.hasNext()) {
    existingFiles.next().setTrashed(true);
  }

  const copy = DriveApp.getFileById(templateId).makeCopy(fileName, folder);
  const docId = copy.getId();

  // Retry open (stable)
  let doc;
  for (let i = 0; i < 5; i++) {
    try {
      Utilities.sleep(1000);
      doc = DocumentApp.openById(docId);
      break;
    } catch (e) {
      if (i === 4) throw e;
    }
  }

  const body = doc.getBody();

  // 🧾 Build table data
  const tableData = [[
    'ITEM NO.',
    'ITEM DESCRIPTION',
    'TIED BIDDERS',
    'BID PRICE'
  ]];

  tiedData.forEach(item => {
    tableData.push([
      item.itemNo,
      item.description,
      item.bidders.join(',\n'),
      '₱' + item.amount.toLocaleString(undefined, {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2
      })
    ]);
  });

  // 🔁 Replace {{ITEMS}} with table
  const placeholder = body.findText('{{ITEMS}}');

  if (placeholder) {
    const element = placeholder.getElement();
    const paragraph = element.getParent();
    const index = body.getChildIndex(paragraph);

    body.removeChild(paragraph); // remove placeholder
    body.insertTable(index, tableData); // insert table
  }

  doc.saveAndClose();
}


function getTiedLCRBData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('AOB - Responsive');
  const data = sheet.getDataRange().getValues();

  const headers = data[0];
  const bidderNames = headers.slice(6);

  let tiedResults = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const itemNo = row[0];
    const description = row[3];

    let bids = [];

    for (let j = 6; j < row.length; j++) {
      const bid = row[j];
      if (bid !== '' && !isNaN(bid)) {
        bids.push({
          value: bid,
          bidder: bidderNames[j - 6]
        });
      }
    }

    if (bids.length === 0) continue;

    const minBid = Math.min(...bids.map(b => b.value));
    const tied = bids.filter(b => b.value === minBid);

    if (tied.length > 1) {
      tiedResults.push({
        itemNo,
        description,
        amount: minBid,
        bidders: tied.map(t => t.bidder)
      });
    }
  }

  return tiedResults;
}
