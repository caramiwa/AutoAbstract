function generateAOBPostQualified() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheetName = 'AOB - As Calculated';
  const targetSheetName = 'AOB - Responsive';

  const ui = SpreadsheetApp.getUi();

  // 📌 Instruction alert before anything else
  const initialPrompt = ui.alert(
    'Ensure that the "Disqualified bidders have been removed from the AOB - As Calculated to generate accurate Abstract of Post-qualified bids',
    ui.ButtonSet.OK
  );
  
  if (initialPrompt !== ui.Button.OK) return;

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

  // Move to the second sheet position and activate
  ss.setActiveSheet(newSheet);
  ss.moveActiveSheet(2);

  ui.alert(`"${targetSheetName}" has been created.\n\nYou may now remove disqualified bids manually before extracting the list of Responsive Bids.`);
}
