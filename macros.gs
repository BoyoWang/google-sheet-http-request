function SZZ_resetFile() {
  const generalMarcos = new S02_GeneralMarcos();

  const mainSheet = generalMarcos.createSheetIfNonExist(
    S00_importantSheets.mainSheet
  );

  generalMarcos.deleteUnimportantSheets();

  mainSheet.clearContents().clearFormats();
  Logger.log("File is reset.");

  const arrayTableData = M01_functions.makeFirst2ArrayOfTable(
    S00_httpRequestTable.columns,
    S00_httpRequestTable.title
  );

  generalMarcos.setValuesToSheet(
    mainSheet,
    S00_httpRequestTable.firstCell,
    arrayTableData
  );
}

function SZZ_getHttpRequests() {
  const spreadsheet = SpreadsheetApp.getActive();
  const mainSheet = spreadsheet.getSheetByName(S00_importantSheets.mainSheet);

  const rangeApiListTable = M01_functions.getTableRangeExcludeTopRows(
    mainSheet,
    S00_httpRequestTable.firstCell,
    2
  );

  let arrayApiListTableData = rangeApiListTable.getValues();

  const indexColIndex = S00_httpRequestTable.columns.indexOf("index");
  const apiColIndex = S00_httpRequestTable.columns.indexOf("apiAddress");
  const resColIndex = S00_httpRequestTable.columns.indexOf("response");
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
