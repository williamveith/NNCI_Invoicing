function doGet() {
  const htmlOutput = HtmlService.createHtmlOutput(createInvoices());
  return htmlOutput;
}

function previewHTMLOutput(invoice, parentFolder) {
  parentFolder.createFile(`${invoice.user}.html`, invoice.htmlContent, MimeType.HTML)
}