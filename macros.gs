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

function GZZ1_MonotarouExcuteAll() {
  var spreadsheet = SpreadsheetApp.getActive();
  var targetSheet = spreadsheet.getActiveSheet();
  Monotarou_DoSthAfterPaste1();
  Monotarou_DoSthAfterPaste2();
}

function GZZ2_ResetFile() {
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

function GZZ3_JointShts() {
  Monotaro_JointShts();
}
