class InvoiceSummary {
  constructor(taxed) {
    this.taxes = taxed ?
      {
        appliedCost: "=ROUNDUP(F1,2)",
        surcharge: "=ROUNDUP(L3*0.265,2)",
        saleTax: "=ROUNDUP(sum(L3:L4)*0.0825,2)"
      } :
      {
        appliedCost: "=IF(ROUNDUP(F1,2)>3500,3500,ROUNDUP(SUM(F1),2))",
        surcharge: "0.00",
        saleTax: "0.00"
      };
    this.summary = [["Catagory", "Amount"],
    ["Total Cost", "=ROUNDUP(D1,2)"],
    ["Applied Cost", this.taxes.appliedCost],
    ["Surcharge (26.5%)", this.taxes.surcharge],
    ["Sales Tax (8.25%)", this.taxes.saleTax],
    ["Payment Due", "=ROUNDUP(SUM(L3:L5),2)"],
    ["Total Time", "=ROUNDUP(SUM(F4:F),5)"]];
  };
}

function addFees() {
  const spreadsheetInvoices = getSpreadsheetFiles();
  spreadsheetInvoices.forEach(invoice => addCustomFees(invoice));
  spreadsheetInvoices.forEach(invoice => addInvoiceSummary(invoice));
  spreadsheetInvoices.forEach(invoice => styleGoogleSheet(invoice.sheet));
}

function addCustomFees(invoice) {
  let userFee = undefined;
  switch (invoice.user) {
    case "AND (Applied Novel Devices)":
      userFee = 500;
      break;
    case "GraphAudio, Inc":
      userFee = 1000;
      break;
  };
  if (userFee !== undefined) {
    const fee = [[getCustomChargeDate(), "00:00:00", invoice.user, invoice.user, "Lab use", "1.00000", userFee, userFee, userFee]];
    invoice.sheet.getRange(invoice.sheet.getLastRow() + 1, 1, 1, 9)
      .setValues(fee);
  };
  invoice.sheet.getRange(1, 6, 1, 1)
    .setValues([["=ROUNDUP(SUM(I4:I),2)"]]);
  invoice.sheet.getRange(1, 4, 1, 1)
    .setValues([["=ROUNDUP(SUM(H4:H),2)"]]);
}

function addInvoiceSummary(invoice) {
  const taxed = !invoice.user.includes(...["Prof", "Dr. Neal Hall", "Bureau of Economic Geology"]);
  const charges = new InvoiceSummary(taxed);
  invoice.sheet.getRange(1, 11, 7, 2)
    .setValues(charges.summary);
}

function styleGoogleSheet(sheet) {
  sheet.getDataRange()
    .setFontFamily("Times New Roman")
    .setFontSize(12);
  sheet.getRange(1, 11, 1, 2)
    .setFontWeight("bold");
  sheet.getRange(2, 12, 5, 1)
    .setNumberFormat("#,##0.00");
  sheet.getRange(7, 12, 1, 1)
    .setNumberFormat("0.00000");
  sheet.autoResizeColumns(1, 12);
}
