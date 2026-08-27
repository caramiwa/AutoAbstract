function generateLCBDocuments() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const projectTitle = 'Notices of LCB';
  const templateId = '1M9iQu_McK1p9hV1ejcVYZb93TRhs0UWARfYaA2GlOnQ';

  const bidderNames = getBidderNamesFromSource();
  if (bidderNames.length === 0) {
    ui.alert('No bidder names found. Please check the AOB - As Calculated sheet.');
    return;
  }

  const bidderSheets = bidderNames
    .map(name => ss.getSheetByName(name))
    .filter(sheet => sheet && sheet.getLastRow() > 1);

  if (bidderSheets.length === 0) {
    ui.alert('No bidder sheets found. Please run "Extract LCB per Bidder" first.');
    return;
  }

  const folder = getOrCreateProjectFolder(projectTitle);
  const generatedFiles = [];

  bidderSheets.forEach(sheet => {
    const bidderName = sheet.getName();
    const data = sheet.getDataRange().getValues();
    const fileName = `LCB – ${bidderName}`;

    const existingFiles = folder.getFilesByName(fileName);
    while (existingFiles.hasNext()) {
      existingFiles.next().setTrashed(true);
    }

    const docId = copyAndFillLCBTemplate(templateId, bidderName, data);
    const file = DriveApp.getFileById(docId);

    file.setName(fileName);
    folder.addFile(file);

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

  return sourceSheet
    .getRange(1, 7, 1, sourceSheet.getLastColumn() - 6)
    .getValues()[0]
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
  const doc = openDocumentWithRetry(copy.getId());
  const body = doc.getBody();

  // Replace only the dynamic bidder placeholder; all template formatting remains intact.
  replaceTextIfFound(body, '{{BiddersName}}', bidderName);

  // Replace the BidTable placeholder with the exact LCB result from the bidder sheet.
  insertLCBTableAtPlaceholder(body, data);

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

function insertLCBTableAtPlaceholder(body, data) {
  const placeholder = '{{BidTable}}';
  const match = body.findText(placeholder);

  if (!match) {
    throw new Error('Placeholder {{BidTable}} not found in the LCB template.');
  }

  const textElement = match.getElement().asText();
  const parent = textElement.getParent();
  const index = body.getChildIndex(parent);

  // Remove only the paragraph containing {{BidTable}}.
  // Everything else in the template remains untouched.
  body.removeChild(parent);

  // Use the bidder sheet result exactly as supplied by extractLowestBidsPerBidder().
  // No ranking or filtering is performed here, so tied LCB items are preserved.
  const tableData = data.map(row =>
    row.map(value => value === null || value === undefined ? '' : String(value))
  );

  body.insertTable(index, tableData);
}
