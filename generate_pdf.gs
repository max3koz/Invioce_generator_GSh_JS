function exportInvoicesToPDF() {
  const originalSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const fileName = "Invoices_" + new Date().toISOString().slice(0,10);
  const invoiceSheets = originalSpreadsheet.getSheets().filter(s => s.getName().startsWith("Invoice_"));

  if (invoiceSheets.length === 0) {
    SpreadsheetApp.getUi().alert("The letter 'Invoice_' was not found.");
    return;
  }

  // Copy only with Invoice_ letters
  const tempSpreadsheet = SpreadsheetApp.create("Temp for PDF");
  const tempFile = DriveApp.getFileById(tempSpreadsheet.getId());
  invoiceSheets.forEach(sheet => {
    sheet.copyTo(tempSpreadsheet).setName(sheet.getName());
  });

  // Delete the blank sheet created automatically
  const defaultSheet = tempSpreadsheet.getSheets()[0];
  tempSpreadsheet.deleteSheet(defaultSheet);

  const url = `https://docs.google.com/spreadsheets/d/${tempSpreadsheet.getId()}/export?format=pdf` +
              `&size=A4&portrait=true&fitw=true&sheetnames=false&printtitle=false` +
              `&pagenumbers=false&gridlines=false&fzr=false`;

  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(url, {
    headers: { Authorization: `Bearer ${token}` }
  });

  const blob = response.getBlob().setName(fileName + ".pdf");

  // Saving PDF
  const folder = DriveApp.getRootFolder(); // Або вкажи свій ID папки
  const file = folder.createFile(blob);

  // Removing the temporary table
  tempFile.setTrashed(true);

  SpreadsheetApp.getUi().alert("✅ PDF created: " + file.getUrl());
}
