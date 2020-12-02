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

function FN_changeObjectValueToArray(object) {
  //  var newArray = [];
  //  for ([key, val] in object){
  //    newArray.push(val);
  //  };
  //  return newArray
  return Object.values(object); // 2020.02.27 : fixed
}

function FN_makeFirst2ArrayOfLists(
  /*object*/ ListTitleColIndexObject,
  /*string*/ listMainTitle
) {
  var indexInfoArray = FN_changeObjectValueToArray(ListTitleColIndexObject);
  indexInfoArray.sort(function (a, b) {
    return a[0] - b[0];
  }); //sort the indexInfoArray by index

  var indexAmt = indexInfoArray.length;

  var firstRowArray = [];
  firstRowArray.push(listMainTitle);
  for (var i = 1; i < indexAmt; i++) {
    firstRowArray.push("");
  }
  var secondRowArray = [];
  for (var i = 0; i < indexAmt; i++) {
    secondRowArray.push(indexInfoArray[i][1]);
  }

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
