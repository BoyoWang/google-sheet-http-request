function SZZ_resetFile() {
  const generalMarcos = new M02_GeneralMarcos();

  const mainSheet = generalMarcos.createSheetIfNonExist(
    M00_importantSheets.mainSheet
  );

  generalMarcos.deleteUnimportantSheets();

  mainSheet.clearContents().clearFormats();
  Logger.log("File is reset.");

  var arrayTableData = M01_functions.makeFirst2ArrayOfTable(
    M00_tablesInfo.httpRequestList.columns,
    M00_tablesInfo.httpRequestList.title
  );

  generalMarcos.setValuesToSheet(
    mainSheet,
    M00_tablesInfo.httpRequestList.firstCell,
    arrayTableData
  );
}

function SZZ_getHttpRequests() {
  const spreadsheet = SpreadsheetApp.getActive();
  const mainSheet = spreadsheet.getSheetByName(M00_importantSheets.mainSheet);

  const rangeApiList = M01_functions.getTableRangeExcludeTopRows(
    mainSheet,
    M00_tablesInfo.httpRequestList.firstCell,
    2
  );

  let rangeApiListData = rangeApiList.getValues();

  Logger.log(rangeApiListData);

  const indexColIndex = M00_tablesInfo.httpRequestList.columns.index[0];
  const apiColIndex = M00_tablesInfo.httpRequestList.columns.apiAdderss[0];
  const resColIndex = M00_tablesInfo.httpRequestList.columns.response[0];
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
