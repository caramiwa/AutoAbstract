function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('AOB Tools')
      .addItem('Higlight AOB - As Read ','HighlightRowsCells')
      .addSeparator()
      .addSubMenu(
          SpreadsheetApp.getUi()
          .createMenu('Detailed Evaluation')
          .addItem('Generate AOB - As Calculated', 'generateAOBAsCalculated')      
          .addItem('Higlight AOB - As Calculated ','HighlightRowsCells2')
          .addItem('Generate Bid Rankings', 'extractSortBids')      
          .addItem('Extract LCB per Bidder', 'extractLowestBidsPerBidder')
          .addItem("Prepare Notice of LCB", 'generateLCBDocuments')          
      )  
      .addSeparator()
      .addSubMenu(
       SpreadsheetApp.getUi()
        .createMenu('Post-Qualification')
        .addItem('Generate Abstract of Post-Qualified Bids', 'generateAOBPostQualified') 
        .addItem('Higlight LCRB ','HighlightLCRB')
  //    .createMenu('📄 BAC Resolution')
        .addItem('Generate BAC Resolution Document', 'promptAndGenerateBACResolution')
  //      .addItem("Generate BAC RESOLUTION", 'generateBACResolution')        
        .addItem("Generate Notice of Award", 'generateNoticeOfAward')

      )
      .addSeparator()
       .addItem('Extract Failed Biddings','extractFailedBiddings')     
      .addItem('Delete Bidders Sheets', 'deleteBidderSheetsSA')
//    .addItem('Extract Bids per Bidder', 'extractAllBidsPerBidder')
      .addToUi();
}

function generateLCBDocuments() {
  SpreadsheetApp.getUi().alert('Function is found and running!');
}
