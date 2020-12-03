class S00_Table {
  constructor(title, firstCellAddress, columns) {
    this.title = title;
    this.firstCell = firstCellAddress;
    this.columns = columns;
  }
}

const S00_tablesInfo = new S00_Table("httpRequestList", "A1", [
  "index",
  "apiAddress",
  "response",
]);

const S00_importantSheets = {
  mainSheet: "Main",
};
