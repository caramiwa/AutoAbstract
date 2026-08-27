function generateLCBDocuments() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const bidderSheets = ss.getSheets().filter(sheet => {
    const name = sheet.getName();
    return name !== 'Sheet1' && name !== 'AOB - As Read' && name !== 'AOB - As Calculated' && name !== 'Bid Ranking' && name !== 'FAILED BIDDINGS';
  });

  if (bidderSheets.length === 0) {
    ui.alert('No bidder sheets found. Please Extract LCB Per Bidder from the AOB Tools first to generate bidder data sheets.');
    return;
  }

  const projectTitle = "Notices of LCB"; // 🔁 Update this manually

  const folder = getOrCreateProjectFolder(projectTitle);
  const templateId = '1M9iQu_McK1p9hV1ejcVYZb93TRhs0UWARfYaA2GlOnQ'; // replace accordingly

  bidderSheets.forEach(sheet => {
    const bidderName = sheet.getName();
    const data = sheet.getDataRange().getValues();
    const bidTable = generateHtmlTable(data);

    const docId = copyAndFillTemplate(templateId, {
      '{{BidTable}}': bidTable,
      '{{BiddersName}}': bidderName
    });


   const fileName = `LCB – ${bidderName}`;

  // Check if file already exists in the folder
    const existingFiles = folder.getFilesByName(fileName);
    while (existingFiles.hasNext()) {
     const existingFile = existingFiles.next();
     existingFile.setTrashed(true); // send to trash instead of hard delete
  }

    const file = DriveApp.getFileById(docId);
    folder.addFile(file);
    DriveApp.getRootFolder().removeFile(file);
    file.setName(`LCB – ${bidderName}`);

   /* if (existingFiles.hasNext()) {
     ui.alert(`File "${fileName}" already exists. Skipping this bidder.`);
    return; // or continue to skip
    }*/

  });

  ui.alert('LCB documents generated successfully.');
}



function getOrCreateProjectFolder(projectTitle) {
  const folders = DriveApp.getFoldersByName(projectTitle);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(projectTitle);
}

function generateHtmlTable(data) {
  const tableRows = data.map(row => 
    '<tr>' + row.map(cell => `<td>${cell}</td>`).join('') + '</tr>'
  ).join('');
  return `<table border="1" cellspacing="0" cellpadding="4">${tableRows}</table>`;
}

function copyAndFillTemplate(templateId, replacements) {
  const template = DriveApp.getFileById(templateId);
  const copy = template.makeCopy();
  const doc = DocumentApp.openById(copy.getId());
  const body = doc.getBody();

  // Replace placeholders in the document body text first (except the table placeholder)
  for (const key in replacements) {
    if (key !== '{{BidTable}}') {
      body.replaceText(key, replacements[key]);
    }
  }

  // Insert table at placeholder {{BidTable}} location
  const placeholder = '{{BidTable}}';
  const foundElement = findTextElement(body, placeholder);

  if (!foundElement) {
    throw new Error('Placeholder {{BidTable}} not found in the document.');
  }

  // Remove the placeholder text
  const textElement = foundElement.getElement().asText();
  const startOffset = foundElement.getStartOffset();
  const endOffset = foundElement.getEndOffsetInclusive();
  textElement.deleteText(startOffset, endOffset);

  // Parse your HTML table string to 2D array (assuming replacements['{{BidTable}}'] is HTML string)
  const data = parseHtmlTableToArray(replacements['{{BidTable}}']);

  // Insert table into the document at the location of the placeholder
  const parent = textElement.getParent();
  const table = body.insertTable(body.getChildIndex(parent) + 1, data);

  // Format table: font, size, column width, wrapping, and number formatting
  formatTable(table);

  doc.saveAndClose();
  return doc.getId();
}

// Helper: Find text element containing placeholder text
function findTextElement(body, searchText) {
  const searchResult = body.findText(searchText);
  return searchResult;
}

// Helper: Parse HTML table string to 2D array of values
function parseHtmlTableToArray(htmlTable) {
  // Simplified parsing for table rows and cells
  // This can be replaced with a proper parser if needed

  // Remove outer <table> tags and split rows
  const rows = htmlTable.match(/<tr>(.*?)<\/tr>/g);
  if (!rows) return [[]];

  return rows.map(row => {
    const cells = row.match(/<t[dh]>(.*?)<\/t[dh]>/g);
    if (!cells) return [];
    return cells.map(cell => {
      // Remove tags to get raw cell content
      return cell.replace(/<\/?t[dh]>/g, '').trim();
    });
  });
}

// Helper: Format the inserted table as per your needs
function formatTable(table) {
  const numRows = table.getNumRows();

  // Set font and size for whole table
  for (let r = 0; r < numRows; r++) {
    const row = table.getRow(r);
    for (let c = 0; c < row.getNumCells(); c++) {
      const cell = row.getCell(c);
      cell.editAsText().setFontFamily('Century Gothic').setFontSize(10);
    }
  }

  // Adjust column widths (Google Docs API is limited so we'll mimic by setting min width)
  // Wider ITEM DESCRIPTION column (index 3)
  for (let r = 0; r < numRows; r++) {
    const cell = table.getRow(r).getCell(3);
    // Wrap text by setting the cell's text style to allow wrapping (Docs auto wraps)
    cell.setWidth(100); // approximate width in points
  }

  // Format numeric columns (indices 1,4,5,6,7)
  // Convert text to numbers and apply currency formatting
  for (let r = 1; r < numRows; r++) { // skip header
    const row = table.getRow(r);

    // TOTAL QUANTITY (index 1) as number (no currency)
    formatCellAsNumber(row.getCell(1));

    // UNIT COST (4), TOTAL COST (5), BID PRICE (6), TOTAL BID PRICE (7) as currency
    formatCellAsCurrency(row.getCell(4));
    formatCellAsCurrency(row.getCell(5));
    formatCellAsCurrency(row.getCell(6));
    formatCellAsCurrency(row.getCell(7));
  }
}

function formatCellAsNumber(cell) {
  const text = cell.getText();
  const num = parseFloat(text);
  if (!isNaN(num)) {
    cell.clear();
    cell.appendParagraph(num.toString());
  }
}

function formatCellAsCurrency(cell) {
  const text = cell.getText();
  let num = parseFloat(text);
  if (!isNaN(num)) {
    // Format as PHP currency or generic currency with two decimals
  //  const formatted = `₱${num.toFixed(2)}`;
    const formatted = num.toFixed(2);
    cell.clear();
    cell.appendParagraph(formatted);
  }
}

