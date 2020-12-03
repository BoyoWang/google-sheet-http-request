function SZZ_resetFile() {
  const generalMarcos = new M02_GeneralMarcos();

  const mainSheet = generalMarcos.createSheetIfNonExist(
    M00_importantSheets.mainSheet
  );

  generalMarcos.deleteUnimportantSheets();

  mainSheet.clearContents().clearFormats();
  Logger.log("File is reset.");

  const arrayTableData = M01_functions.makeFirst2ArrayOfTable(
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

  const rangeApiListTable = M01_functions.getTableRangeExcludeTopRows(
    mainSheet,
    M00_tablesInfo.httpRequestList.firstCell,
    2
  );

  let arrayApiListTableData = rangeApiListTable.getValues();

  const indexColIndex = M00_tablesInfo.httpRequestList.columns.index[0];
  const apiColIndex = M00_tablesInfo.httpRequestList.columns.apiAdderss[0];
  const resColIndex = M00_tablesInfo.httpRequestList.columns.response[0];
  let apiAddress = "";
  let res = "";

  for (i = 0; i < arrayApiListTableData.length; i++) {
    apiAddress = arrayApiListTableData[i][apiColIndex];
    res = sendGetRequest(apiAddress);
    if (res.id) res = res.id;
    if (res.message) res = res.message;
    arrayApiListTableData[i][resColIndex] = res;
    arrayApiListTableData[i][indexColIndex] = i + 1;
  }

  rangeApiListTable.setValues(arrayApiListTableData);

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
