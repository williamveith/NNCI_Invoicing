var invoicingUIFunctions = {
  createTheSheets: (function () {
    generateInventorySpreadsheets();
  }),
  addFees: (function () {
    addFees();
  }),
  makePdfFiles: (function () {
    createInvoices();
  }),
  makeSheetsExcel: (function () {
    makeExcelFiles();
  }),
  sendSheets: (function () {
    mailInventorySpreadsheets();
  }),
  openReadMe: (function () {
    openModalBox("https://docs.google.com/document/d/e/2PACX-1vQimVfLCOViHQAOjUoC0LPjOzLP0wRMh91fNdJtRx6E3vYcSq_ni4biKRmnAz-gLRKIXb6vZPQEwriH/pub")
  })
}

function invoicingUI() {
  SpreadsheetApp.getUi().createMenu(`Invoice`)
    .addItem(`Generate .Inv Spreadsheets`, `NNCIInvoicing.invoicingUIFunctions.createTheSheets`)
    .addSeparator()
    .addItem(`Add Fees`, `NNCIInvoicing.invoicingUIFunctions.addFees`)
    .addSeparator()
    .addItem(`Generate PDF Invoices`, `NNCIInvoicing.invoicingUIFunctions.makePdfFiles`)
    .addSeparator()
    .addItem(`Generate Excel Files`, `NNCIInvoicing.invoicingUIFunctions.makeSheetsExcel`)
    .addSeparator()
    .addItem(`Mail Invoices`, `NNCIInvoicing.invoicingUIFunctions.sendSheets`)
    .addSeparator()
    .addItem(`Help`, `NNCIInvoicing.invoicingUIFunctions.openReadMe`)
    .addToUi();
}