function createInvoices() {
  const monthFolder = getMonthFolder();
  const spreadsheetInvoices = getSpreadsheetFiles();
  spreadsheetInvoices.forEach((invoice, index) => getInvoiceValues(invoice, index));
  const invoiceDataOfService = getDateOfService();
  spreadsheetInvoices.forEach(invoice => createHTML(invoice, invoiceDataOfService));
  const pdfFolder = monthFolder.createFolder(`pdf`);
  spreadsheetInvoices.forEach(invoice => turnHTMLIntoPDF(invoice, pdfFolder));
}

function getInvoiceValues(invoice, index) {
  const fees = invoice.sheet.getRange(2, 12, 6, 1).getValues().flat();
  const itemizedExpenses = invoice.sheet.getRange(4, 1, invoice.sheet.getLastRow() - 3, 9).getValues();
  invoice.invoiceNumber = getInvoiceNumber(index);
  invoice.itemizedExpenses = itemizedExpenses.length !== 4 ? itemizedExpenses : itemizedExpenses.filter(row => row[0] !== "");
  invoice.totalCost = formatCurrency(fees[0]);
  invoice.appliedCost = formatCurrency(fees[1]);
  invoice.surcharge = formatCurrency(fees[2]);
  invoice.salesTax = formatCurrency(fees[3]);
  invoice.balenceDue = formatCurrency(fees[4]);
  invoice.totalTime = fees[5].toFixed(5);
}

function createHTML(invoice, invoiceDataOfService) {
  const template = HtmlService.createTemplateFromFile("Invoice Template");
  template.contact = invoice.contact;
  template.invoiceNumber = invoice.invoiceNumber;
  template.dateOfService = invoiceDataOfService;
  template.appliedCost = invoice.appliedCost;
  template.surcharge = invoice.surcharge;
  template.salesTax = invoice.salesTax;
  template.balenceDue = invoice.balenceDue;
  template.itemizedExpenses = invoice.itemizedExpenses;
  template.totalTime = invoice.totalTime;
  template.totalCost = invoice.totalCost;
  invoice.htmlContent = template.evaluate().getContent();
}

function turnHTMLIntoPDF(invoice, parentFolder) {
  const blob = Utilities.newBlob(invoice.htmlContent, MimeType.HTML);
  blob.setName(`${invoice.user}.pdf`);
  parentFolder.createFile(blob.getAs(MimeType.PDF));
}
