class M02_GeneralMarcos {
  constructor(spreadsheet = SpreadsheetApp.getActive()) {
    this.spreadsheet = spreadsheet;
  }

  deleteUnimportantSheets() {
    const sheets = this.spreadsheet.getSheets();

    for (let i = 0; i < sheets.length; i++) {
      let sheetNameToTest = sheets[i].getSheetName();
      if (!M01_functions.testIfSheetIsImportant(sheetNameToTest)) {
        this.spreadsheet.deleteSheet(
          this.spreadsheet.getSheetByName(sheetNameToTest)
        );
        Logger.log("Sheet '" + sheetNameToTest + "' is deleted.");
      }
    }
  }

  createSheetIfNonExist(sheetName) {
    let sheet = this.spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      sheet = this.spreadsheet.insertSheet(sheetName);
      Logger.log(sheet.getName() + "created.");
    }
    return sheet;
  }
}
