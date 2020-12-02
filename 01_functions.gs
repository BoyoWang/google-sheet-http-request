function FN_findCellByText_ReturnRange(/*sheet*/ sheet, /*string*/ textToFind) {
  var spreadsheet = SpreadsheetApp.getActive();
  var allDataRange = sheet.getDataRange();
  var allDataRangeData = allDataRange.getValues();

  for (var i = 0; i < allDataRange.getNumRows(); i++) {
    for (var j = 0; j < allDataRange.getNumColumns(); j++) {
      if (allDataRangeData[i][j] == textToFind) {
        return sheet.getRange(i + 1, j + 1);
      }
    }
  }
}

function FN_get_ColRange_In_TitleRow(/*string*/ textToFind, /*sheet*/ sheet) {
  var dataRange = sheet.getDataRange();
  var data = dataRange.getValues(); //data[row][col]

  for (var i = 0; i < dataRange.getNumColumns(); i++) {
    if (data[0][i] == textToFind) {
      return sheet.getRange(1, i + 1, dataRange.getNumRows(), 1);
    }
  }
}

function FN_transformObjectValuesToArray(object) {
  return Object.values(object);
}

function FN_makeFirst2ArrayOfLists(columnInfoObject, tableTitle) {
  // read table column infos
  var indexInfoArray = FN_transformObjectValuesToArray(columnInfoObject);

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
}

function FN_returnListRangeExcludeTopRows(
  sheet,
  firstCellAddress_In_A1Style,
  intExcludeRowNum
) {
  var spreadsheet = SpreadsheetApp.getActive();
  //  var sheet = spreadsheet.getActiveSheet();
  var dataRegion = sheet.getRange(firstCellAddress_In_A1Style).getDataRegion();
  var rangeToReturn = sheet.getRange(
    dataRegion.getRow() + intExcludeRowNum,
    dataRegion.getColumn(),
    dataRegion.getNumRows() - intExcludeRowNum,
    dataRegion.getNumColumns()
  );
  return rangeToReturn;
}
