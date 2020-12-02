function S02_Delete_NonImportant_Sheets() {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheets = spreadsheet.getSheets();

  for (var i = 0; i < sheets.length; i++) {
    var sheetNameToTest = sheets[i].getSheetName();
    if (!TestIfSheetIsImportant(sheetNameToTest)) {
      spreadsheet.deleteSheet(spreadsheet.getSheetByName(sheetNameToTest));
      Logger.log("Sheet '" + sheetNameToTest + "' is deleted.");
    }
  }

  function TestIfSheetIsImportant(sheetName) {
    var importantSheetsArray = FN_changeObjectValueToArray(
      name_importantSheets
    );
    if (importantSheetsArray.indexOf(sheetName) > -1) {
      return true;
    } else {
      return false;
    }
  }
}

function S02_Make_list_For_all_sheets(sheet, firstCellAddress_In_A1Style) {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheetListFirstRange = sheet.getRange(firstCellAddress_In_A1Style);

  //Make sheetName array
  var sheetNames = [];
  for (var i = 0; i < spreadsheet.getSheets().length; i++) {
    sheetNames.push(spreadsheet.getSheets()[i].getSheetName());
  }

  //Make finalList array
  var finalList = FN_makeFirst2ArrayOfLists(
    address_firstCell_A1_Style.sheetListTitleColIndex,
    address_firstCell_A1_Style.sheetList_MainTitleText
  ); //First 2 rows

  //Push every sheet infos into array
  for (i = 0; i < sheetNames.length; i++) {
    var row = [];
    row.push(i, sheetNames[i], "-");
    finalList.push(row);
  }

  //Set array values to the range
  sheet
    .getRange(
      sheetListFirstRange.getRow(),
      sheetListFirstRange.getColumn(),
      finalList.length,
      finalList[0].length
    )
    .setValues(finalList);

  Logger.log(
    "sheetList is made in sheet '" +
      sheet.getName() +
      "' on cell '" +
      firstCellAddress_In_A1Style +
      "'."
  );
}

function S02_list_all_files_inside_folder(
  folderID,
  sheet,
  cellAddress_In_A1Style
) {
  var folder = DriveApp.getFolderById(folderID);

  //First 2 rows
  var list = FN_makeFirst2ArrayOfLists(
    address_firstCell_A1_Style.csvFileListTitleColIndex,
    address_firstCell_A1_Style.csvFileList_MainTitleText
  );

  //Push every file infos into array
  var files = folder.getFiles();
  while (files.hasNext()) {
    var file = files.next();
    var row = [];
    row.push(file.getName(), file.getId(), file.getSize(), "-");
    list.push(row);
  }

  //Set array values to the range
  var firstCell = sheet.getRange(cellAddress_In_A1Style);
  sheet
    .getRange(
      firstCell.getRow(),
      firstCell.getColumn(),
      list.length,
      list[0].length
    )
    .setValues(list);

  Logger.log(
    "csvFileList is made in sheet '" +
      sheet.getName() +
      "' on cell '" +
      cellAddress_In_A1Style +
      "'."
  );
}

function S02_importCSVFromGoogleDrive(fileID) {
  var spreadsheet = SpreadsheetApp.getActive();
  var mainSheet = spreadsheet.getSheetByName(name_importantSheets.mainSheet);

  var file = DriveApp.getFileById(fileID);
  var csvData = Utilities.parseCsv(file.getBlob().getDataAsString("sjis"));

  spreadsheet.insertSheet(SpreadsheetApp.getActive().getNumSheets());
  var newSheet = spreadsheet.getActiveSheet();
  var nameExcludeExtension = file.getName().split(".").slice(0, -1).join(".");
  newSheet.setName(nameExcludeExtension);
  newSheet.getRange(1, 1, csvData.length, csvData[0].length).setValues(csvData);
  Logger.log(
    "File '" +
      file.getName() +
      "' is imported to sheet '" +
      nameExcludeExtension +
      "'"
  );
}

function S02_importAll_CSV_Files() {
  var spreadsheet = SpreadsheetApp.getActive();
  var mainSheet = spreadsheet.getSheetByName(name_importantSheets.mainSheet);

  //Read file list and make array
  var rangeCSV_FileList = FN_returnListRangeExcludeTopRows(
    mainSheet,
    address_firstCell_A1_Style.csvFileList,
    2
  );
  var rangeCSV_FileListData = rangeCSV_FileList.getValues();

  //Read col index for googleDriveID column
  var fileID_ColIndex =
    address_firstCell_A1_Style.csvFileListTitleColIndex.googleDriveID[0];

  //Make fileID list
  var csvFileID_List = [];

  for (i = 0; i < rangeCSV_FileListData.length; i++) {
    csvFileID_List.push(rangeCSV_FileListData[i][fileID_ColIndex]);
  }

  //Import every file
  for (i = 0; i < csvFileID_List.length; i++) {
    S02_importCSVFromGoogleDrive(csvFileID_List[i]);
  }
}

function S02_Sort_CSV_FileList() {
  var spreadsheet = SpreadsheetApp.getActive();
  var mainSheet = spreadsheet.getSheetByName(name_importantSheets.mainSheet);
  var rangeToSort = FN_returnListRangeExcludeTopRows(
    mainSheet,
    address_firstCell_A1_Style.csvFileList,
    2
  );

  var columnToSort =
    address_firstCell_A1_Style.csvFileListTitleColIndex.name[0] + 1;
  //Sort the list
  rangeToSort.sort({ column: columnToSort, ascending: true });
  Logger.log("csvFileList is sorted.");
}

function S02_resetFile() {
  var spreadsheet = SpreadsheetApp.getActive();
  var mainSheet = spreadsheet.getSheetByName(name_importantSheets.mainSheet);
  S02_Delete_NonImportant_Sheets();
  mainSheet.clearContents().clearFormats();
  Logger.log("File is reset.");
}

function S02_importCSVExcuteAll() {
  var spreadsheet = SpreadsheetApp.getActive();
  var mainSheet = spreadsheet.getSheetByName(name_importantSheets.mainSheet);

  // if mainSheet doesn't exist create it
  if (!mainSheet) {
    spreadsheet.insertSheet(name_importantSheets.mainSheet);
    mainSheet = spreadsheet.getSheetByName(name_importantSheets.mainSheet);
    Logger.log("mainSheet created.");
  }

  S02_resetFile();
  S02_list_all_files_inside_folder(
    folderGoogleDriveID,
    mainSheet,
    address_firstCell_A1_Style.csvFileList
  );

  FN_returnListRangeExcludeTopRows(
    mainSheet,
    address_firstCell_A1_Style.csvFileList,
    1
  ).applyRowBanding(); //Banding rows for csvFileList

  S02_Sort_CSV_FileList();
  S02_importAll_CSV_Files();

  S02_Make_list_For_all_sheets(mainSheet, address_firstCell_A1_Style.sheetList);

  FN_returnListRangeExcludeTopRows(
    mainSheet,
    address_firstCell_A1_Style.sheetList,
    1
  ).applyRowBanding(); //Banding rows for sheetList

  mainSheet.activate();
}
