const S01_functions = {
  transformObjectValuesToArray: function (object) {
    return Object.values(object);
  },
  makeFirst2ArrayOfTable: function (
    /*arrayOfStrings*/ arrayColumns,
    /*string*/ tableTitle
  ) {
    // read column amount
    const colAmt = arrayColumns.length;

    // values for first row, title, "", "" ...
    let firstRowArray = [];
    firstRowArray.push(tableTitle);
    for (i = 1; i < colAmt; i++) {
      firstRowArray.push("");
    }

    // values for second row, col1Title, col2Title, col3Title ...
    var secondRowArray = [];
    for (i = 0; i < colAmt; i++) {
      secondRowArray.push(arrayColumns[i]);
    }

    // push fisrt and second row into new array
    var arrayReturn = [];
    arrayReturn.push(firstRowArray);
    arrayReturn.push(secondRowArray);

    return arrayReturn;
  },
  getCellRangeByText: function (
    /*sheet*/ sheet,
    /*string*/ textToFind,
    /*boolean*/ exactMatch = true
  ) {
    var allDataRange = sheet.getDataRange();
    var allDataRangeData = allDataRange.getValues();

    for (var i = 0; i < allDataRange.getNumRows(); i++) {
      for (var j = 0; j < allDataRange.getNumColumns(); j++) {
        if (
          exactMatch
            ? allDataRangeData[i][j] == textToFind
            : typeof allDataRangeData[i][j] == "string" &&
              allDataRangeData[i][j].indexOf(textToFind) > -1
        ) {
          return sheet.getRange(i + 1, j + 1);
        }
      }
    }
  },
  getColRangeTitleRow: function (/*sheet*/ sheet, /*string*/ textToFind) {
    var dataRange = sheet.getDataRange();
    var data = dataRange.getValues(); //data[row][col]

    for (var i = 0; i < dataRange.getNumColumns(); i++) {
      if (data[0][i] == textToFind) {
        return sheet.getRange(1, i + 1, dataRange.getNumRows(), 1);
      }
    }
  },
  getTableRangeExcludeTopRows: function (
    /*sheet*/ sheet,
    /*string*/ firstCellAddressA1Style,
    /*integer*/ intExcludeRowAmt
  ) {
    var dataRegion = sheet.getRange(firstCellAddressA1Style).getDataRegion();
    var rangeToReturn = sheet.getRange(
      dataRegion.getRow() + intExcludeRowAmt,
      dataRegion.getColumn(),
      dataRegion.getNumRows() - intExcludeRowAmt,
      dataRegion.getNumColumns()
    );
    return rangeToReturn;
  },
  testIfSheetIsImportant: function (/*string*/ sheetName) {
    const importantSheetsArray = this.transformObjectValuesToArray(
      S00_importantSheets
    );
    if (importantSheetsArray.indexOf(sheetName) > -1) {
      return true;
    } else {
      return false;
    }
  },
};
