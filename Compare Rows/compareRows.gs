function compareRows(sheetName, row1, row2) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    var rowData1 = sheet.getRange(row1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var rowData2 = sheet.getRange(row2, 1, 1, sheet.getLastColumn()).getValues()[0];
    var differences = [];
  
    for (var i = 0; i < rowData1.length; i++) {
        if (rowData1[i] !== rowData2[i]) {
            differences.push("Column " + String.fromCharCode(65 + i) + ": " + rowData1[i] + " vs " + rowData2[i]);
        }
    }
  
    if (differences.length === 0) {
        return "No differences";
    } else {
        return differences.join(", ");
    }
}
