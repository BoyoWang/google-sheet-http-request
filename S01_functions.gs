const M01_functions = {
  transformObjectValuesToArray: function (object) {
    return Object.values(object);
  },
  makeFirst2ArrayOfTable: function (columnInfoObject, tableTitle) {
    // read table column infos
    var indexInfoArray = this.transformObjectValuesToArray(columnInfoObject);

    // sort the indexInfoArray by index
    indexInfoArray.sort(function (a, b) {
      return a[0] - b[0];
    });

    // read column amount
    var colAmt = indexInfoArray.length;

    // values for first row, title, "", "" ...
    var firstRowArray = [];
    firstRowArray.push(tableTitle);
    for (var i = 1; i < colAmt; i++) {
      firstRowArray.push("");
    }

    // values for second row, col1Title, col2Title, col3Title ...
    var secondRowArray = [];
    for (var i = 0; i < colAmt; i++) {
      secondRowArray.push(indexInfoArray[i][1]);
    }

    // push fisrt and second row into new array
    var arrayReturn = [];
    arrayReturn.push(firstRowArray);
    arrayReturn.push(secondRowArray);

    return arrayReturn;
  },
  getCellRangeByText: function (/*sheet*/ sheet, /*string*/ textToFind) {
    var allDataRange = sheet.getDataRange();
    var allDataRangeData = allDataRange.getValues();

    for (var i = 0; i < allDataRange.getNumRows(); i++) {
      for (var j = 0; j < allDataRange.getNumColumns(); j++) {
        if (allDataRangeData[i][j] == textToFind) {
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
  testIfSheetIsImportant: function (sheetName) {
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
