function myFunction() {
  
  //1.Change 'Sheet1' to be matching your sheet name
  var sheet = SpreadsheetApp.getActive().getSheetByName('5r_test');
  sheet.appendRow(['test1', 'test2']);  
}
