function copyOnDateMatch() {
  var sheet1 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data");
  var sheet2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Final");
  
  // Get today's date
  var today = new Date();
  
  // Format today's date to dd-MM-yyyy
  var todayFormatted = Utilities.formatDate(today, Session.getScriptTimeZone(), "dd-MM-yyyy");

  // Get all data in column A
  var data = sheet1.getRange(1, 1, sheet1.getLastRow(), sheet1.getLastColumn()).getValues();

  // Loop through each row
  for (var i = 0; i < data.length; i++) {
    var dateInCell = new Date(data[i][0]); // Get date from column A
    
    // Subtract 6 days from the date in the cell
    var adjustedDate = new Date(dateInCell);
    adjustedDate.setDate(adjustedDate.getDate() - 6);

    // Format the adjusted date
    var formattedAdjustedDate = Utilities.formatDate(adjustedDate, Session.getScriptTimeZone(), "dd-MM-yyyy");

    // Check if the adjusted date matches today's date
    if (formattedAdjustedDate === todayFormatted) {
      // Copy the entire row to Sheet2
      sheet2.appendRow(data[i]);
    }
  }
}
