/* ----------- Get and Format Month & Year ----------- */
function getMonthYear() {
  const date = new Date();
  let month = date.getMonth();
  let year = date.getFullYear();
  if (month === 0) {
    month = 12;
    year = year - 1;
  };
  if (month < 10) {
    month = `0${month}`;
  };
  return {
    month: month,
    year: year
  };
}

function getMonthFolder(nnciInvoicingFolderId = `1ip09FMvDwpmnCAykdi5N5fjs-OHLaXnr`) {
  const date = getMonthYear()
  return DriveApp.getFolderById(nnciInvoicingFolderId)
    .getFoldersByName(`${date.year}-${date.month}`)
    .next();
}

// Used in Add Sales Tax.gs
function getCustomChargeDate() {
  const date = getMonthYear();
  return `${date.year}-${date.month}-01`;
}

// Used in Create Invoices.gs
function getDateOfService() {
  const date = getMonthYear()
  return `${date.year}-${date.month}-01 to ${date.year}-${date.month}-${new Date(date.year, date.month, 0).getDate()}`;
}

function getInvoiceNumber(index) {
  const invoiceIndex = index + 1;
  const date = getMonthYear()
  const padding = '0'.repeat(4 - invoiceIndex.toString().length)
  return `${date.year}${date.month}${padding}${invoiceIndex}`.toString();
}

/* ----------- Get and Process Invoices in Spreadsheet Form ----------- */
class SpreadsheetInvoice {
  constructor(spreadsheetFile) {
    this.user = spreadsheetFile.getName();
    this.sheet = SpreadsheetApp.openById(spreadsheetFile.getId()).getSheets()[0];
    this.contact = (() => {
      try {
        const contactFile = DriveApp.getFolderById("1edBks0lsVpI6CLA5XMHILOU-IkWHNb5f").getFilesByName(`${this.user}.txt`).next();
        return contactFile.getBlob().getDataAsString().split(/\r?\n/);
      } catch (error) {
        Logger.log(`${this.user} does not have a contact file`);
        return [""];
      }
    })();
    this.invoiceNumber = "";
    this.itemizedExpenses = "";
    this.totalCost = "";
    this.appliedCost = "";
    this.surcharge = "";
    this.salesTax = "";
    this.balenceDue = "";
    this.totalTime = "";
  }
}

function getSpreadsheetFiles() {
  const monthFolder = getMonthFolder();
  const spreadsheetFiles = monthFolder.getFilesByType(MimeType.GOOGLE_SHEETS);
  const invoices = [];
  while (spreadsheetFiles.hasNext()) {
    invoices.push(new SpreadsheetInvoice(spreadsheetFiles.next()));
  }
  return invoices;
}

/* ----------- Format Numbers  ----------- */
function formatCurrency(currencyValue) {
  const formatter = new Intl.NumberFormat('en-US', {
    style: 'currency',
    currency: 'USD',
  });
  return formatter.format(currencyValue);
}

/* ----------- Open Document ----------- */
function openModalBox(url) {
  const htmlTemplate = HtmlService.createTemplateFromFile(`open`);
  htmlTemplate.url = url;
  const html = htmlTemplate.evaluate().getContent();
  SpreadsheetApp.getUi()
    .showModalDialog(
      HtmlService.createHtmlOutput(html).setHeight(1),
      `Opening...`,
    );
}