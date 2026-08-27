function generateAOBAsCalculated() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheetName = 'AOB - As Read';
  const targetSheetName = 'AOB - As Calculated';

  const ui = SpreadsheetApp.getUi();

  const sourceSheet = ss.getSheetByName(sourceSheetName);
  if (!sourceSheet) {
    ui.alert(`Error: "${sourceSheetName}" does not exist.`);
    return;
  }

  const existingTarget = ss.getSheetByName(targetSheetName);
  if (existingTarget) {
    const response = ui.alert(
      `"${targetSheetName}" already exists. Do you want to overwrite it?`,
      ui.ButtonSet.YES_NO
    );
    if (response === ui.Button.NO) return;

    ss.deleteSheet(existingTarget);
  }

  const newSheet = sourceSheet.copyTo(ss);
  newSheet.setName(targetSheetName);

  const calculatedSheet = ss.getSheetByName('AOB - As Calculated');
    if (calculatedSheet) {
        calculatedSheet.getDataRange().setBackground(null);
}


  // Move to the end and activate
  ss.setActiveSheet(newSheet);
  //ss.moveActiveSheet(ss.getSheets().length);
  ss.moveActiveSheet(2); 
  ui.alert(`"${targetSheetName}" has been created. You may now remove disqualified bids manually before running the ranking.`);
}
