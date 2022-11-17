function makeExcelFiles() {
  const monthFolder = getMonthFolder();
  const excelFolder = monthFolder.createFolder(`excel`);
  const spreadsheetFiles = monthFolder.getFilesByType(MimeType.GOOGLE_SHEETS);
  while (spreadsheetFiles.hasNext()) {
    const spreadsheetFile = spreadsheetFiles.next();
    convertToExcel(spreadsheetFile, excelFolder);
    spreadsheetFile.setTrashed(true);
  };
}

function convertToExcel(spreadsheetFile, parentFolder) {
  const url = `https://docs.google.com/feeds/download/spreadsheets/Export?key=${spreadsheetFile.getId()}&exportFormat=xlsx`;
  const params = {
    method: "get",
    headers: { "Authorization": `Bearer ${ScriptApp.getOAuthToken()}` },
    muteHttpExceptions: true
  };
  const blob = UrlFetchApp.fetch(url, params).getBlob();
  blob.setName(`${spreadsheetFile.getName()}.xlsx`);
  parentFolder.createFile(blob);
}