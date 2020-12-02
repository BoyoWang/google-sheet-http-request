function S03_ApplyActionToAllSheets() {
  var spreadsheet = SpreadsheetApp.getActive();
  var mainSheet = spreadsheet.getSheetByName(name_importantSheets.mainSheet);

  var sheetList;

  //Aquire sheet list
  var rangesheetList = FN_returnListRangeExcludeTopRows(
    mainSheet,
    address_firstCell_A1_Style.sheetList,
    2
  );
  var rangesheetListData = rangesheetList.getValues();
  var sheetList = [];

  var colIndex = address_firstCell_A1_Style.sheetListTitleColIndex.name[0];

  for (i = 1; i < rangesheetListData.length; i++) {
    sheetList.push(rangesheetListData[i][colIndex]);
  }

  //Do something from here

  //Delete usless rows
  for (i = 0; i < sheetList.length; i++) {
    var targetSheet = spreadsheet.getSheetByName(sheetList[i]);
    var DataStart = FN_findCellByText_ReturnRange(targetSheet, "番号");
    var rowsToDelete = DataStart.getRow() - 1;
    targetSheet.deleteRows(1, rowsToDelete);

    //Create new Sheet
    spreadsheet.insertSheet(SpreadsheetApp.getActive().getNumSheets());
    var newSheet = spreadsheet.getActiveSheet();
    var targetSheetTempName = sheetList[i] + "_t";
    var newSheetName = sheetList[i];

    targetSheet.setName(targetSheetTempName);
    newSheet.setName(newSheetName);

    //copy needed columns
    var columnsToKeepText = [
      ["us", "General"],
      ["CH2", "#,##0.000"],
    ];
    for (j = 0; j < columnsToKeepText.length; j++) {
      var rangeCopyFrom = FN_get_ColRange_In_TitleRow(
        columnsToKeepText[j][0],
        targetSheet
      );
      rangeCopyFrom.setNumberFormat(columnsToKeepText[j][1]);
      //Set format before copy
      rangeCopyFrom.copyTo(newSheet.getRange(1, j + 1));
    }

    S03_CreateAndAdjustTimeColumn(newSheet);

    //Delete old and rename new sheet
    spreadsheet.deleteSheet(targetSheet);
  }
}

function S03_CreateAndAdjustTimeColumn(sheet) {
  var spreadsheet = SpreadsheetApp.getActive().getActiveSheet();
  var spreadsheet = sheet;

  spreadsheet.insertColumnsBefore(spreadsheet.getRange("A:A").getColumn(), 1);
  spreadsheet.getRange("A1").setValue("Time");
  spreadsheet.getRange("A2").setValue("s");

  //Create the first two timeStamp and auto fill
  var A3_Value = spreadsheet.getRange("B3").getValue() / 1000000;
  var A4_Value = spreadsheet.getRange("B4").getValue() / 1000000;
  spreadsheet.getRange("A3").setValue(A3_Value);
  spreadsheet.getRange("A4").setValue(A4_Value);
  spreadsheet
    .getRange("A3:A4")
    .autoFillToNeighbor(SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);

  //Delete Col B
  spreadsheet.deleteColumns(2, 1);
  spreadsheet.getRange("B1").setValue(spreadsheet.getName());
}
