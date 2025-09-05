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

