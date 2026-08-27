function generateLCBDocuments() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const projectTitle = 'Notices of LCB';
  const templateId = '1M9iQu_McK1p9hV1ejcVYZb93TRhs0UWARfYaA2GlOnQ';

  // Identify bidder sheets using the same bidder list that drives LCB extraction.
  const bidderNames = getBidderNamesFromSource();
  if (bidderNames.length === 0) {
    ui.alert('No bidder names found. Please check the AOB - As Calculated sheet.');
    return;
  }

  const bidderSheets = bidderNames
    .map(name => ss.getSheetByName(name))
    .filter(sheet => sheet && sheet.getLastRow() > 1);

  if (bidderSheets.length === 0) {
    ui.alert(
      'No bidder sheets found. Please run "Extract LCB per Bidder" first.'
    );
    return;
  }

  const folder = getOrCreateProjectFolder(projectTitle);
  const generatedFiles = [];

  bidderSheets.forEach(sheet => {
    const bidderName = sheet.getName();
    const data = sheet.getDataRange().getValues();
    const fileName = `LCB – ${bidderName}`;

    // Remove an existing document with the same name.
    const existingFiles = folder.getFilesByName(fileName);
    while (existingFiles.hasNext()) {
      existingFiles.next().setTrashed(true);
    }

    const docId = copyAndFillLCBTemplate(templateId, bidderName, data);
    const file = DriveApp.getFileById(docId);

    file.setName(fileName);
    folder.addFile(file);

    // Remove the file from My Drive root when possible.
    try {
      DriveApp.getRootFolder().removeFile(file);
    } catch (e) {
      // Ignore if the file is not directly contained in the root folder.
    }

    generatedFiles.push(fileName);
  });

  ui.alert(
    `LCB documents generated successfully.\n\n${generatedFiles.length} document(s) created.`
  );
}

function getBidderNamesFromSource() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName('AOB - As Calculated');

  if (!sourceSheet || sourceSheet.getLastColumn() < 7) return [];

  const headers = sourceSheet
    .getRange(1, 7, 1, sourceSheet.getLastColumn() - 6)
    .getValues()[0];

  return headers
    .map(name => String(name).trim())
    .filter(name => name !== '');
}

function getOrCreateProjectFolder(projectTitle) {
  const folders = DriveApp.getFoldersByName(projectTitle);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(projectTitle);
}

function copyAndFillLCBTemplate(templateId, bidderName, data) {
  const template = DriveApp.getFileById(templateId);
  const copy = template.makeCopy();

  // A newly copied Google Doc can occasionally take a moment before it is accessible.
  const doc = openDocumentWithRetry(copy.getId());
  const body = doc.getBody();

  // Replace bidder placeholder.
  replaceTextIfFound(body, '{{BiddersName}}', bidderName);

  // Insert the actual sheet data directly as a Google Docs table.
  insertLCBTable(body, data);

  formatLCBTable(body);

  doc.saveAndClose();
  return copy.getId();
}

function openDocumentWithRetry(docId) {
  let lastError;

  for (let attempt = 0; attempt < 5; attempt++) {
    try {
      Utilities.sleep(500);
      return DocumentApp.openById(docId);
    } catch (error) {
      lastError = error;
      Utilities.sleep(500);
    }
  }

  throw lastError;
}

function replaceTextIfFound(body, searchText, replacement) {
  const match = body.findText(searchText);
  if (!match) return false;

  match.getElement().asText().replaceText(searchText, replacement);
  return true;
}

function insertLCBTable(body, data) {
  const placeholder = '{{BidTable}}';
  const match = body.findText(placeholder);

  if (!match) {
    throw new Error('Placeholder {{BidTable}} not found in the LCB template.');
  }

  const textElement = match.getElement().asText();
  const parent = textElement.getParent();
  const index = body.getChildIndex(parent);

  // Remove only the placeholder paragraph.
  body.removeChild(parent);

  const tableData = data.map((row, rowIndex) => {
    return row.map((value, columnIndex) => {
      if (rowIndex === 0) return String(value);

      // Keep quantities readable while consistently formatting monetary values.
      if (columnIndex === 1) return formatQuantity(value);
      if (columnIndex >= 4) return formatCurrencyNumber(value);

      return value === null || value === undefined ? '' : String(value);
    });
  });

  body.insertTable(index, tableData);
}

function formatQuantity(value) {
  const number = Number(value);
  if (isNaN(number)) return value === null || value === undefined ? '' : String(value);
  return number.toLocaleString(undefined, { maximumFractionDigits: 2 });
}

function formatCurrencyNumber(value) {
  const number = Number(value);
  if (isNaN(number)) return value === null || value === undefined ? '' : String(value);
  return number.toLocaleString(undefined, {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  });
}

function formatLCBTable(body) {
  const tables = body.getTables();
  if (tables.length === 0) return;

  // The LCB table is the table inserted by insertLCBTable().
  const table = tables[tables.length - 1];
  const numRows = table.getNumRows();

  for (let r = 0; r < numRows; r++) {
    const row = table.getRow(r);

    for (let c = 0; c < row.getNumCells(); c++) {
      const cell = row.getCell(c);
      cell.editAsText()
        .setFontFamily('Century Gothic')
        .setFontSize(10);
    }

    // Keep the ITEM DESCRIPTION column reasonably wide.
    if (row.getNumCells() > 3) {
      row.getCell(3).setWidth(100);
    }
  }
}
