function myAllFunction() {
  Logger.log(`\n▶️ Executing All functions...\n`);
  var activeSheetName = "";
  var googleSheetIdOSA = "1MQ1xwZ1PQm0E2J3WEjoN9mXf2a_HVUpp5S_u5NQdVn4";
  var googleSheetName = this.getSprintDate();
  this.sheetCreationFunctions(googleSheetIdOSA, googleSheetName);
}

function sheetCreationFunctions(googleSheetIdOSA, googleSheetName) {
  Logger.log(`\n▶️ Executing Sheet Creation All functions...\n`);
  var activeSheetName = "";
  Logger.log(`✅ '${googleSheetName}' has been found...`);
  var sheetExists= this.isSheetPresentByUrl(googleSheetIdOSA, googleSheetName);
  if (sheetExists){
    // this.createNewSheetAtFirstPosition(googleSheetIdOSA, googleSheetName);
  } else{
    this.createNewSheetAtFirstPosition(googleSheetIdOSA, googleSheetName);
  }
  this.copyDataAndFormattingFromSecondToFirstSheet(googleSheetIdOSA, googleSheetName);
  this.setManualColumnWidths(googleSheetIdOSA, googleSheetName);
  this.freezeTheColumn(googleSheetIdOSA, googleSheetName, 3);
  this.freezeTheRows(googleSheetIdOSA, googleSheetName, 2);
  this.setFontNameStyleSize(googleSheetIdOSA, googleSheetName);
  this.setRowsTextMiddle(googleSheetIdOSA, googleSheetName);
}

function getSheetByUrl(spreadSheetId, sheetName) {
  
  var url = `https://docs.google.com/spreadsheets/d/${spreadSheetId}/edit`;
  Logger.log(`Url: ${spreadSheetId} and Sheet Name: ${sheetName} & Final Url: ${url} `);
  var spreadsheet = SpreadsheetApp.openByUrl(url);
  var sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    Logger.log('❌ Sheet with name "' + sheetName + '" not found..');
    return null;
  }
  Logger.log(`\n➡️ Sheet '${sheetName}' has been found successfully...`);
  return sheet;
}

function createNewSheet(sheetName) {
  Logger.log(`\n▶️ Executing Create New Sheet functions...`);
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  // Check if sheet already exists
  var existingSheet = spreadsheet.getSheetByName(sheetName);
  if (!existingSheet) {
    spreadsheet.insertSheet(sheetName);
    Logger.log(`\n✅ New sheet '${sheetName}' has been created successfully...`);
  } else {
    Logger.log('\n❌ Sheet with this name already exists.');
  }
}

function createNewSheetAtFirstPosition(spreadSheetId, sheetName) {
  Logger.log(`\n▶️ Executing Create New Sheet functions...\n`);
  var url = `https://docs.google.com/spreadsheets/d/${spreadSheetId}/edit`;
  Logger.log(`Url: ${spreadSheetId} and Sheet Name: ${sheetName} & Final Url: ${url} `);
  var spreadsheet = SpreadsheetApp.openByUrl(url);
  var existingSheet = spreadsheet.getSheetByName(sheetName);
  if (!existingSheet) {
    var newSheet = spreadsheet.insertSheet(sheetName, 0);
    Logger.log(`\n✅ New sheet '${sheetName}' has been created successfully...\n`);
  } else {
    Logger.log('\n❌ Sheet with this name already exists.\n\n');
  }
}

function copyDataFromSecondToFirstSheet(spreadSheetId, sheetName) {
  Logger.log(`\n▶️ Executing Copy Data fron 2nd to New Sheet functions...`);
  var url = `https://docs.google.com/spreadsheets/d/${spreadSheetId}/edit`;
  Logger.log(`Url: ${spreadSheetId} and Sheet Name: ${sheetName} & Final Url: ${url} `);
  var spreadsheet = SpreadsheetApp.openByUrl(url);
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

function copyDataAndFormattingFromSecondToFirstSheet(spreadSheetId, sheetName) {
  Logger.log(`\n▶️ Executing Copy Data fron 2nd to New Sheet with format functions...\n`);
  var url = `https://docs.google.com/spreadsheets/d/${spreadSheetId}/edit`;
  Logger.log(`Url: ${spreadSheetId} and Sheet Name: ${sheetName} & Final Url: ${url} `);
  var spreadsheet = SpreadsheetApp.openByUrl(url);
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

function setManualColumnWidthsPOS(spreadSheetId, sheetName) {
  Logger.log(`▶️ Executing Set Columns Width in New Sheet functions...`);
  var url = `https://docs.google.com/spreadsheets/d/${spreadSheetId}/edit`;
  Logger.log(`Url: ${spreadSheetId} and Sheet Name: ${sheetName} & Final Url: ${url} `);
  var spreadsheet = SpreadsheetApp.openByUrl(url);
  var sheet = spreadsheet.getSheetByName(sheetName).getSheets()[0]; // 1st sheet
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

function setManualColumnWidths(spreadSheetId, sheetName) {
  Logger.log(`\n▶️ Executing Set Columns Width in New Sheet functions...`);
  var url = `https://docs.google.com/spreadsheets/d/${spreadSheetId}/edit`;
  Logger.log(`Url: ${spreadSheetId} and Sheet Name: ${sheetName} & Final Url: ${url} `);
  var spreadsheet = SpreadsheetApp.openByUrl(url);
  var sheet = spreadsheet.getSheets()[0]; // 1st sheet

  // var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0]; // 1st sheet

  // Set widths as per your specification
  sheet.setColumnWidth(1, 15);  // Column A
  sheet.setColumnWidth(2, 550); // Column B
  sheet.setColumnWidth(3, 330); // Column C
  sheet.setColumnWidth(4, 15); // Column D
  sheet.setColumnWidth(5, 70); // Column E
  sheet.setColumnWidth(6, 280);  // Column F
}

function setDynamicColumnWidths(spreadSheetId, sheetName) {
  Logger.log(`\n▶️ Executing Set Dynamic Width in New Sheet functions...\n`);
  var url = `https://docs.google.com/spreadsheets/d/${spreadSheetId}/edit`;
  Logger.log(`Url: ${spreadSheetId} and Sheet Name: ${sheetName} & Final Url: ${url} `);
  var spreadsheet = SpreadsheetApp.openByUrl(url);
  var sheet = spreadsheet.getSheetByName(sheetName).getSheets()[0]; // 1st sheet


  // Define widths for the first 8 columns, you can customize these values
  var columnWidths = [120, 150, 100, 130, 110, 140, 160, 125];

  // Loop through the columns and set their widths
  for (var i = 0; i < columnWidths.length; i++) {
    sheet.setColumnWidth(i + 1, columnWidths[i]); // Columns are 1-indexed
  }
}

function freezeTheColumn(spreadSheetId, sheetName, columnIndex) {
  Logger.log(`\n▶️ Executing Set Freeze Columns in New Sheet functions...`);
  var url = `https://docs.google.com/spreadsheets/d/${spreadSheetId}/edit`;
  Logger.log(`Url: ${spreadSheetId} and Sheet Name: ${sheetName} & Final Url: ${url} `);
  var spreadsheet = SpreadsheetApp.openByUrl(url);
  var sheet = spreadsheet.getSheets()[0]; // 1st sheet

  sheet.setFrozenColumns(columnIndex); // Freeze first 2 columns, which includes Column B
}

function freezeTheRows(spreadSheetId, sheetName, rowIndex) {
  Logger.log(`\n▶️ Executing Set Freeze Rows in New Sheet functions...`);
  var url = `https://docs.google.com/spreadsheets/d/${spreadSheetId}/edit`;
  Logger.log(`Url: ${spreadSheetId} and Sheet Name: ${sheetName} & Final Url: ${url} `);
  var spreadsheet = SpreadsheetApp.openByUrl(url);
  var sheet = spreadsheet.getSheets()[0]; // 1st sheet
  sheet.setFrozenRows(rowIndex); // Freeze first 2 columns, which includes Column B
}

function setRowsTextMiddle(spreadSheetId, sheetName) {
  Logger.log(`\n▶️ Executing Seting Text format as Middle in New Sheet functions...\n`);
  var url = `https://docs.google.com/spreadsheets/d/${spreadSheetId}/edit`;
  var spreadsheet = SpreadsheetApp.openByUrl(url);
  var sheet = spreadsheet.getSheets()[0]; // 1st sheet
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  sheet.getRange(1, 1, lastRow, lastColumn).setVerticalAlignment("middle");
}

function setFontNameStyleSize(spreadSheetId, sheetName) {
  Logger.log(`\n▶️ Executing Seting Text format as Middle in New Sheet functions...\n`);
  var url = `https://docs.google.com/spreadsheets/d/${spreadSheetId}/edit`;
  var spreadsheet = SpreadsheetApp.openByUrl(url);
  var sheet = spreadsheet.getSheets()[0]; // 1st sheet
  
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  var range = sheet.getRange(1, 1, lastRow, lastColumn);

  range.setFontFamily("Roboto");  // Set font family as Roboto
  // range.setFontFamily("Space Grotesk");  // Set font family as Space Grotesk
  // range.setFontStyle("italic");           // Set font style: italic or normal
  // range.setFontWeight("normal");          // Set font weight: normal or bold
  range.setFontSize(12);                  // Set font size in points
}







function getNextBiweeklyMondaySprint(startDateStr, currentDateStr) {
  var startDate = new Date(startDateStr);
  var currentDate = new Date(currentDateStr);

  // Normalize startDate to Monday if needed
  if (startDate.getDay() !== 1) { // 1 = Monday
    var daysUntilMonday = (8 - startDate.getDay()) % 7;
    startDate.setDate(startDate.getDate() + daysUntilMonday);
  }

  // Calculate days difference from startDate to currentDate
  var diffDays = Math.floor((currentDate - startDate) / (1000 * 60 * 60 * 24));

  // Compute sprint index — number of 14-day periods completed before currentDate
  var sprintIndex = Math.floor(diffDays / 14);

  // Calculate this sprint Monday date
  var currentSprintMonday = new Date(startDate);
  currentSprintMonday.setDate(startDate.getDate() + sprintIndex * 14);

  // Calculate sprint week Sunday (6 days after Monday)
  var sprintWeekSunday = new Date(currentSprintMonday);
  sprintWeekSunday.setDate(currentSprintMonday.getDate() + 6);

  // Calculate next sprint Monday (14 days after current sprint Monday)
  var nextSprintMonday = new Date(currentSprintMonday);
  nextSprintMonday.setDate(currentSprintMonday.getDate() + 14);

  var resultSprintStart;

  // Logic:
  // If the current date is before the start date, return the start date
  if (diffDays < 0) {
    resultSprintStart = startDate;
  }
  // If current date is within sprint week (Monday to Sunday)
  else if (currentDate >= currentSprintMonday && currentDate <= sprintWeekSunday) {
    resultSprintStart = currentSprintMonday;
  }
  // If current date is after sprint Sunday, return next sprint Monday
  else if (currentDate > sprintWeekSunday) {
    resultSprintStart = nextSprintMonday;
  }
  else {
    // Fallback (should not come here)
    resultSprintStart = currentSprintMonday;
  }

  // Format YYYY-MM-DD
  var yyyy = resultSprintStart.getFullYear();
  var mm = (resultSprintStart.getMonth() + 1).toString().padStart(2, '0');
  var dd = resultSprintStart.getDate().toString().padStart(2, '0');

  return `${yyyy}-${mm}-${dd}`;
}

// Test function that logs expected results
function testSprintDate() {
  var sprintStart = "2025-08-25";
  var testDates = [
    "2025-09-07",
    "2025-09-08",
    "2025-09-09",
    "2025-09-10",
    "2025-09-20",
    "2025-09-22",
    "2025-09-23",
    "2025-09-24"
  ];

  testDates.forEach(function(date_) {
    var nextSprint = getNextBiweeklyMondaySprint(sprintStart, date_);
    var dates = this.formatDateToDayMonth(nextSprint)
    Logger.log(`Current date: ${date_} --> Sprint start: ${dates}`);

  });
}

function getSprintDate(){
  var sprintStart = "2025-08-25";
  var toDate = this.getTodayAsYyyyMmDd();
  var nextSprint = getNextBiweeklyMondaySprint(sprintStart, toDate);
  var dates = this.formatDateToDayMonth(nextSprint)
  Logger.log(`Current date: ${toDate} --> Sprint start: ${dates}`);
  return dates;
}

function formatDateToDayMonth(date) {
  // If the input is a string, convert it to a Date object
  if (!(date instanceof Date)) {
    date = new Date(date);
  }
  
  var day = date.getDate();
  var monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
                    "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  var month = monthNames[date.getMonth()];
  return day + " " + month;
}

function getTodayAsYyyyMmDd() {
  var today = new Date();
  var yyyy = today.getFullYear();
  var mm = (today.getMonth() + 1).toString().padStart(2, '0');
  var dd = today.getDate().toString().padStart(2, '0');
  return `${yyyy}-${mm}-${dd}`;
}

function isSheetPresentByUrl(spreadSheetId, sheetName) {
  var url = `https://docs.google.com/spreadsheets/d/${spreadSheetId}/edit`;
  var spreadsheet = SpreadsheetApp.openByUrl(url);
  var sheet = spreadsheet.getSheetByName(sheetName);
  return sheet !== null;
}

// creating data updation functions




