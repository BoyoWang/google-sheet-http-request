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

  const indexColIndex =
    address_firstCell_A1_Style.httpRequestList.columns.index[0];
  const apiColIndex =
    address_firstCell_A1_Style.httpRequestList.columns.apiAdderss[0];
  const resColIndex =
    address_firstCell_A1_Style.httpRequestList.columns.response[0];
  let testApi = "";
  let result = "";

  for (i = 0; i < rangeApiListData.length; i++) {
    testApi = rangeApiListData[i][apiColIndex];
    result = sendGetRequest(testApi);
    if (result.id) result = result.id;
    if (result.message) result = result.message;
    rangeApiListData[i][resColIndex] = result;
    rangeApiListData[i][indexColIndex] = i + 1;
  }

  rangeApiList.setValues(rangeApiListData);

  function sendGetRequest(apiAddress) {
    try {
      const response = UrlFetchApp.fetch(apiAddress, {
        muteHttpExceptions: true,
      });
      const json = response.getContentText();
      const data = JSON.parse(json);
      return data;
    } catch (error) {
      return error;
    }
  }
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
