function myFunction2() {
  function myFunction() {
  
  //1.Change 'Sheet1' to be matching your sheet name
  var sheet2 = SpreadsheetApp.getActive().getSheetByName('Weight');
  var sheet = SpreadsheetApp.getActive().getSheetByName('5r_test');
  
  
  
  var data = sheet.getDataRange().getValues();
  for (var i = 0; i < data.length; i++) {
    sheet.appendRow([data[i][0], data[i][1]]);    
    
    
  }

}
  
}
