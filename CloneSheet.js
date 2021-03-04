// copy data from Google Sheet A to Google Sheet B
// Credit: @chrislkeller [modified by CKR September 10, 2020]

function cloneGoogleSheet() {

  var srcID = '1T6-t26yWyLGlmurPtmzSW7ems3I34dm0MrEEaZOrJC8' ; // source: Caroline's playground    
  //var srcID = '1o2YvAWd5i5_tGxeinpA2ZSCCqohQVUSZd82e--Mvtv4' ; // source: BlocksDBEXtraTabs
  //var srcID = '1qnCxA6FPIh1Y5w-cG3LFzdPnkVu2b0p14_viVjkDldg' ; // source: Blocks DB (the master)
  
  //var dstID = '16VYh9xAYIak7oOVEQlnLXFA275TouBhjptUzYYFZDO8' ; // destination: DummyTest
  //var dstID = '1T6-t26yWyLGlmurPtmzSW7ems3I34dm0MrEEaZOrJC8' ; // destination: Caroline's playground
  var dstID = '1o2YvAWd5i5_tGxeinpA2ZSCCqohQVUSZd82e--Mvtv4' ; // destination: BlocksDBEXtraTabs
      
  // Tab name
  var tabname = 'Blocks DB';
  
  // source spreadsheet (doc)
  var source = SpreadsheetApp.openById(srcID);

  // source sheet ("tab")
  var sourcetab = source.getSheetByName(tabname);

  // Get full range of data
  var SRange = sourcetab.getDataRange();

  // get A1 notation identifying the range
  var A1Range = SRange.getA1Notation();

  // get the data values in range
  var SData = SRange.getValues();

  // target spreadsheet (doc)
  var target = SpreadsheetApp.openById(dstID);

  // target sheet ("tab")
  var targettab = target.getSheetByName(tabname);

  // Clear the target tab before copy:
  targettab.clear({contentsOnly: true});

  // set the target range to the values of the source data
  targettab.getRange(A1Range).setValues(SData);

};
