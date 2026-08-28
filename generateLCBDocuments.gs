function generateLCBDocuments() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const inputs = promptForLCBResources(ui);
  if (!inputs) return;

  const templateId = extractGoogleId(inputs.templateUrl, 'document');
  const folderId = extractGoogleId(inputs.folderUrl, 'folder');

  if (!templateId) {
    ui.alert('Invalid Notice Template URL. Please paste the complete Google Docs URL.');
    return;
  }

  if (!folderId) {
    ui.alert('Invalid destination Folder URL. Please paste the complete Google Drive folder URL.');
    return;
  }

  let template;
  let folder;

  try {
    template = DriveApp.getFileById(templateId);
    if (template.getMimeType() !== MimeType.GOOGLE_DOCS) {
      ui.alert('The selected Notice Template is not a Google Docs document.');
      return;
    }

    folder = DriveApp.getFolderById(folderId);
  } catch (error) {
    ui.alert('Unable to access the template or destination folder. Please check the URLs and your access permissions.');
    return;
  }

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

function promptForLCBResources(ui) {
  const templateResponse = ui.prompt(
    'Prepare Notice of LCB',
    'Paste the Google Docs URL of the Notice Template:',
    ui.ButtonSet.OK_CANCEL
  );

  if (templateResponse.getSelectedButton() !== ui.Button.OK) return null;

  const templateUrl = templateResponse.getResponseText().trim();
  if (!templateUrl) {
    ui.alert('No Notice Template URL was provided.');
    return null;
  }

  const folderResponse = ui.prompt(
    'Prepare Notice of LCB',
    'Paste the Google Drive URL of the destination folder:',
    ui.ButtonSet.OK_CANCEL
  );

  if (folderResponse.getSelectedButton() !== ui.Button.OK) return null;

  const folderUrl = folderResponse.getResponseText().trim();
  if (!folderUrl) {
    ui.alert('No destination Folder URL was provided.');
    return null;
  }

  return { templateUrl, folderUrl };
}

function extractGoogleId(url, type) {
  const text = String(url || '').trim();

  if (type === 'document') {
    const match = text.match(/\/document\/d\/([a-zA-Z0-9_-]+)/);
    return match ? match[1] : null;
  }

  if (type === 'folder') {
    const match = text.match(/\/folders\/([a-zA-Z0-9_-]+)/);
    return match ? match[1] : null;
  }

  return null;
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

function copyAndFillLCBTemplate(templateId, bidderName, data) {
  const template = DriveApp.getFileById(templateId);
  const copy = template.makeCopy();
  const doc = openDocumentWithRetry(copy.getId());
  const body = doc.getBody();

  replaceTextIfFound(body, '{{BiddersName}}', bidderName);
  insertLCBTableAtPlaceholder(body, data);
  replaceFooterPlaceholder(doc, '{{BiddersName}}', bidderName);

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

function replaceFooterPlaceholder(doc, placeholder, replacement) {
  const footer = doc.getFooter();
  if (!footer) return false;

  const table = findFirstTableInFooter_(footer);
  if (!table) return false;

  // The LCB template places {{BiddersName}} in Row 1, Column 1
  // (zero-based: row 0, column 0) of the footer table.
  if (table.getNumRows() < 1 || table.getRow(0).getNumCells() < 1) {
    return false;
  }

  const cell = table.getCell(0, 0);
  const escapedPlaceholder = escapeRegExp_(placeholder);

  const match = cell.findText(escapedPlaceholder);
  if (!match) return false;

  match.getElement().asText().replaceText(escapedPlaceholder, replacement);
  return true;
}

function findFirstTableInFooter_(footer) {
  for (let i = 0; i < footer.getNumChildren(); i++) {
    const child = footer.getChild(i);
    if (child.getType() === DocumentApp.ElementType.TABLE) {
      return child.asTable();
    }
  }

  return null;
}

function escapeRegExp_(text) {
  return String(text).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
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

  if (rowIndex === 0) return String(value);

  if (columnIndex === 1) {
    const quantity = Number(value);
    if (!isNaN(quantity)) {
      return quantity.toLocaleString('en-PH', { maximumFractionDigits: 2 });
    }
    return String(value);
  }

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
