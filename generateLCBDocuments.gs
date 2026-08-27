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

  showLCBFolderDialog(generatedFiles.length, folder.getUrl());
}

function showLCBFolderDialog(documentCount, folderUrl) {
  const html = HtmlService.createHtmlOutput(`
    <div style="font-family: Arial, sans-serif; padding: 20px; text-align: center;">
      <div style="font-size: 18px; font-weight: bold; margin-bottom: 12px;">
        LCB documents generated successfully.
      </div>
      <div style="margin-bottom: 20px;">
        ${documentCount} document(s) created.
      </div>
      <a href="${folderUrl}" target="_blank"
         style="display: inline-block; padding: 10px 18px; background: #1a73e8;
                color: white; text-decoration: none; border-radius: 4px;">
        Open Notices of LCB Folder
      </a>
      <div style="margin-top: 16px;">
        <button onclick="google.script.host.close()"
                style="padding: 7px 18px; cursor: pointer;">
          Close
        </button>
      </div>
    </div>
  `)
    .setWidth(420)
    .setHeight(230);

  SpreadsheetApp.getUi().showModalDialog(html, 'AutoAbstract');
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

  replaceTextIfFound(body, '{{BiddersName}}', bidderName);
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

  body.removeChild(parent);

  const tableData = data.map((row, rowIndex) =>
    row.map((value, columnIndex) =>
      formatLCBDocumentValue(value, columnIndex, rowIndex)
    )
  );

  body.insertTable(index, tableData);
}

function formatLCBDocumentValue(value, columnIndex, rowIndex) {
  if (value === null || value === undefined) return '';

  // Header row remains unchanged.
  if (rowIndex === 0) return String(value);

  // Column B: quantity. Add thousands separators, but no currency symbol.
  if (columnIndex === 1) {
    const quantity = Number(value);
    if (!isNaN(quantity)) {
      return quantity.toLocaleString('en-PH', { maximumFractionDigits: 2 });
    }
    return String(value);
  }

  // Columns E-H are monetary values. Format as Philippine pesos with
  // thousands separators and exactly two decimal places. Rounding here
  // also prevents floating-point artifacts such as 241199.99999999997.
  if (columnIndex >= 4 && columnIndex <= 7) {
    const amount = Number(value);
    if (!isNaN(amount)) {
      return '₱' + amount.toLocaleString('en-PH', {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2
      });
    }
  }

  return String(value);
}
