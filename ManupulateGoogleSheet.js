function myAllFunction() {
  var activeSheetName = "";
  var googleSheetIdOSA = "1rLuqdq8fcjLVjFjjln1zOeX3_pezQap1mAwzr2rHMpGA9_9f45E4CC6t";
  var googleSheetName = "8 Sep";
  activeSheetName = this.getSheetByUrl(googleSheetIdOSA,googleSheetName)
  Logger.log(`✅ '${activeSheetName}' has been found...`);
  // this.createNewSheetAtFirstPosition(activeSheetName);
  // this.copyDataAndFormattingFromSecondToFirstSheet();
  // this.setManualColumnWidths();
  // this.freezeTheColumn(2);
  // this.freezeTheRows(2);
}

function getSheetByUrl(spreadSheetId, sheetName) {
  var url = `https://docs.google.com/spreadsheets/d/${spreadSheetId}/edit`;
  var spreadsheet = SpreadsheetApp.openByUrl(urls);
  var sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    Logger.log('❌ Sheet with name "' + sheetName + '" not found..');
    return null;
  }
  Logger.log(`✅ Sheet '${sheet}' has been found successfully...`);
  return sheet;
}

function createNewSheet(sheetName) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  // Check if sheet already exists
  var existingSheet = spreadsheet.getSheetByName(sheetName);
  if (!existingSheet) {
    spreadsheet.insertSheet(sheetName);
    Logger.log(`✅ New sheet '${sheetName}' has been created successfully...`);
  } else {
    Logger.log('❌ Sheet with this name already exists.');
  }
}

function createNewSheetAtFirstPosition(sheetName) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var existingSheet = spreadsheet.getSheetByName(sheetName);
  if (!existingSheet) {
    var newSheet = spreadsheet.insertSheet(sheetName, 0);
    Logger.log(`✅ New sheet '${sheetName}' has been created successfully...`);
  } else {
    Logger.log('❌ Sheet with this name already exists.');
  }
}

function copyDataFromSecondToFirstSheet() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = spreadsheet.getSheets()[1]; // 2nd sheet (index 1)
  var targetSheet = spreadsheet.getSheets()[0]; // 1st sheet (index 0)

  // Clear the target sheet before copying
  targetSheet.clearContents();

  // Get the data range from the source sheet
  var sourceRange = sourceSheet.getDataRange();
  var data = sourceRange.getValues();

  // Set the data into the target sheet starting from A1
  targetSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
}

function copyDataAndFormattingFromSecondToFirstSheet() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = spreadsheet.getSheets()[1]; // 2nd sheet (index 1)
  var targetSheet = spreadsheet.getSheets()[0]; // 1st sheet (index 0)

  // Clear the target sheet before copying
  targetSheet.clear();

  // Get the data range from the source sheet
  var sourceRange = sourceSheet.getDataRange();

  // Copy the values
  var data = sourceRange.getValues();
  targetSheet.getRange(1, 1, data.length, data[0].length).setValues(data);

  // Copy background colors
  var backgrounds = sourceRange.getBackgrounds();
  targetSheet.getRange(1, 1, backgrounds.length, backgrounds[0].length).setBackgrounds(backgrounds);

  // Copy font colors
  var fontColors = sourceRange.getFontColors();
  targetSheet.getRange(1, 1, fontColors.length, fontColors[0].length).setFontColors(fontColors);

  // Copy font weights (bold)
  var fontWeights = sourceRange.getFontWeights();
  targetSheet.getRange(1, 1, fontWeights.length, fontWeights[0].length).setFontWeights(fontWeights);

  // Copy font styles (italic)
  var fontStyles = sourceRange.getFontStyles();
  targetSheet.getRange(1, 1, fontStyles.length, fontStyles[0].length).setFontStyles(fontStyles);

  // Copy horizontal alignments
  var horizontalAlignments = sourceRange.getHorizontalAlignments();
  targetSheet.getRange(1, 1, horizontalAlignments.length, horizontalAlignments[0].length).setHorizontalAlignments(horizontalAlignments);

  // Copy vertical alignments
  var verticalAlignments = sourceRange.getVerticalAlignments();
  targetSheet.getRange(1, 1, verticalAlignments.length, verticalAlignments[0].length).setVerticalAlignments(verticalAlignments);

  // Copy font sizes
  var fontSizes = sourceRange.getFontSizes();
  targetSheet.getRange(1, 1, fontSizes.length, fontSizes[0].length).setFontSizes(fontSizes);

  // Copy borders (if needed, borders can be copied using copyTo)
  sourceRange.copyTo(targetSheet.getRange(1, 1), {formatOnly:true});
}

function setManualColumnWidthsPOS() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0]; // 1st sheet

  // Set widths as per your specification
  sheet.setColumnWidth(1, 8);  // Column A
  sheet.setColumnWidth(2, 500); // Column B
  sheet.setColumnWidth(3, 320); // Column C
  sheet.setColumnWidth(4, 320); // Column D
  sheet.setColumnWidth(5, 320); // Column E
  sheet.setColumnWidth(6, 8);  // Column F
  sheet.setColumnWidth(7, 70);  // Column G
  sheet.setColumnWidth(8, 80);  // Column H
  sheet.setColumnWidth(9, 80);  // Column H
  sheet.setColumnWidth(10, 80);  // Column H
}

function setManualColumnWidths() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0]; // 1st sheet

  // Set widths as per your specification
  sheet.setColumnWidth(1, 10);  // Column A
  sheet.setColumnWidth(2, 500); // Column B
  sheet.setColumnWidth(3, 300); // Column C
  sheet.setColumnWidth(4, 10); // Column D
  sheet.setColumnWidth(5, 70); // Column E
  sheet.setColumnWidth(6, 280);  // Column F
}


function setDynamicColumnWidths() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0]; // 1st sheet

  // Define widths for the first 8 columns, you can customize these values
  var columnWidths = [120, 150, 100, 130, 110, 140, 160, 125];

  // Loop through the columns and set their widths
  for (var i = 0; i < columnWidths.length; i++) {
    sheet.setColumnWidth(i + 1, columnWidths[i]); // Columns are 1-indexed
  }
}

function freezeTheColumn(columnIndex) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0]; // 1st sheet
  sheet.setFrozenColumns(columnIndex); // Freeze first 2 columns, which includes Column B
}

function freezeTheRows(rowIndex) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0]; // 1st sheet
  sheet.setFrozenRows(rowIndex); // Freeze first 2 columns, which includes Column B
}

function setRowsTextMiddle(spreadSheetId, sheetName) {
  var url = `https://docs.google.com/spreadsheets/d/${spreadSheetId}/edit`;
  var sheetName = sheetName;
  var sheet = getSheetByUrl(url, sheetName);
  // var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0]; // 1st sheet
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  sheet.getRange(1, 1, lastRow, lastColumn).setVerticalAlignment("middle");
}







