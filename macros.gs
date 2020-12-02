function G021_Import_All_CSV() {
  S02_importCSVExcuteAll();
}

function G022_Reset_File() {
  var spreadsheet = SpreadsheetApp.getActive();
  var mainSheet = spreadsheet.getSheetByName(name_importantSheets.mainSheet);
  // if mainSheet doesn't exist create it
  if (!mainSheet) {
    spreadsheet.insertSheet(name_importantSheets.mainSheet);
    mainSheet = spreadsheet.getSheetByName(name_importantSheets.mainSheet);
    Logger.log("mainSheet created.");
  }
  S02_resetFile();
}

function G03_Apply_Actions_To_Sheets() {
  S03_ApplyActionToAllSheets();
}

function G041_Update_FileList() {
  S04_updateFileList();
}

function G042_Change_FileName() {
  S04_changeFileName();
}

function G051_Update_SheetList() {
  S05_updateSheetList();
}

function G052_Change_SheetName() {
  S05_changeSheetsName();
}

function GZZ2_ResetFile() {
  const spreadsheet = SpreadsheetApp.getActive();
  const mainSheet = spreadsheet.getSheetByName(name_importantSheets.mainSheet);

  // if mainSheet doesn't exist create it
  if (!mainSheet) {
    spreadsheet.insertSheet(name_importantSheets.mainSheet);
    mainSheet = spreadsheet.getSheetByName(name_importantSheets.mainSheet);
    Logger.log("mainSheet created.");
  }

  SZZ_resetFile();
}

function SZZ_resetFile() {
  const spreadsheet = SpreadsheetApp.getActive();
  const mainSheet = spreadsheet.getSheetByName(name_importantSheets.mainSheet);
  SZZ_Delete_NonImportant_Sheets();
  mainSheet.clearContents().clearFormats();
  Logger.log("File is reset.");

  var list = FN_makeFirst2ArrayOfLists(
    address_firstCell_A1_Style.httpRequestList.columns,
    address_firstCell_A1_Style.httpRequestList.title
  );

  //Set array values to the range
  var firstCell = mainSheet.getRange(
    address_firstCell_A1_Style.httpRequestList.firstCell
  );
  mainSheet
    .getRange(
      firstCell.getRow(),
      firstCell.getColumn(),
      list.length,
      list[0].length
    )
    .setValues(list);
}

function SZZ_getHttpRequests() {
  const spreadsheet = SpreadsheetApp.getActive();
  const mainSheet = spreadsheet.getSheetByName(name_importantSheets.mainSheet);

  const rangeApiList = FN_returnListRangeExcludeTopRows(
    mainSheet,
    address_firstCell_A1_Style.httpRequestList.firstCell,
    2
  );

  let rangeApiListData = rangeApiList.getValues();

  Logger.log(rangeApiListData);
}

function SZZ_Delete_NonImportant_Sheets() {
  const spreadsheet = SpreadsheetApp.getActive();
  const sheets = spreadsheet.getSheets();

  for (let i = 0; i < sheets.length; i++) {
    let sheetNameToTest = sheets[i].getSheetName();
    if (!TestIfSheetIsImportant(sheetNameToTest)) {
      spreadsheet.deleteSheet(spreadsheet.getSheetByName(sheetNameToTest));
      Logger.log("Sheet '" + sheetNameToTest + "' is deleted.");
    }
  }

  function TestIfSheetIsImportant(sheetName) {
    const importantSheetsArray = FN_changeObjectValueToArray(
      name_importantSheets
    );
    if (importantSheetsArray.indexOf(sheetName) > -1) {
      return true;
    } else {
      return false;
    }
  }
}
