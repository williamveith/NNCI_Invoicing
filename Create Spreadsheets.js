class Invoice {
  constructor(file) {
    this.invFile = file;
    this.fileName = file.getName().replace(`.inv`, '');
    this.spreadsheet = SpreadsheetApp.create(this.fileName);
    this.sheet = this.spreadsheet.getSheetByName(`Sheet1`);
    this.data = file.getBlob().getDataAsString().split(/\r?\n/).map(line => {
      const lineArray = line.split(/\t/);
      const arrayLength = lineArray.length;
      if (arrayLength < 9) {
        lineArray.push(...new Array(9 - arrayLength).fill(""));
      };
      return lineArray;
    });
  }
};

function generateInventorySpreadsheets() {
  const monthFolder = getMonthFolder();
  const invFiles = getInventoryFiles(monthFolder);
  const invoices = invFiles.map(file => new Invoice(file));
  createSpreadsheets(invoices, monthFolder);
  invoices.forEach(invoice => setItemizedCellTypes(invoice.sheet));
  organizeFiles(invoices, monthFolder)
}

function getInventoryFiles(monthFolder) {
  const monthFolderFiles = monthFolder.getFilesByType(`application/octet-stream`);
  const monthFiles = [];
  while (monthFolderFiles.hasNext()) {
    monthFiles.push(monthFolderFiles.next());
  };
  return monthFiles;
}

function createSpreadsheets(invoices) {
  invoices.forEach(invoice => {
    invoice.sheet.setName(invoice.fileName)
      .getRange(1, 1, invoice.data.length, invoice.data[0].length)
      .setValues(invoice.data);
  });
}

function setItemizedCellTypes(sheet) {
  const numberOfHeaderRows = 4;
  const maxRows = sheet.getMaxRows() - numberOfHeaderRows;
  const dateRange = sheet.getRange(numberOfHeaderRows, 1, maxRows, 1);
  const timeRange = sheet.getRange(numberOfHeaderRows, 2, maxRows, 1);
  const profUserTooolRange = sheet.getRange(numberOfHeaderRows, 3, maxRows, 3);
  const useRange = sheet.getRange(numberOfHeaderRows, 6, maxRows, 1);
  const rateRange = sheet.getRange(numberOfHeaderRows, 7, maxRows, 1);
  const costAppliedCostRange = sheet.getRange(numberOfHeaderRows, 8, maxRows, 2);
  dateRange.setNumberFormat("@");
  timeRange.setNumberFormat('hh":"mm":"ss').setNumberFormat("@");
  profUserTooolRange.setNumberFormat("@");
  useRange.setNumberFormat("0.00000");
  rateRange.setNumberFormat("0");
  costAppliedCostRange.setNumberFormat("0.00");
}

function organizeFiles(invoices, monthFolder) {
  const invFolder = monthFolder.createFolder(`inv`);
  invoices.forEach(invoice => DriveApp.getFileById(invoice.spreadsheet.getId()).moveTo(monthFolder));
  invoices.forEach(invoice => invoice.invFile.moveTo(invFolder));
}