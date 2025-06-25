function generateInvoices() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet(); 
  var dataSheet = spreadsheet.getSheetByName('Data'); 
  var calculateSheet = spreadsheet.getSheetByName('Calculate'); 
  var invoiceDraft = spreadsheet.getSheetByName('InvoiceDraft'); 

  var startDate = dataSheet.getRange("B18").getValue();
  var endDate = dataSheet.getRange("B19").getValue();

  var dataDict = createDictionary(dataSheet);
  var calculateDict = createDictionary(calculateSheet);

  let invoiceCount = 1;

  Object.keys(calculateDict).forEach(studentKey => {
    // Skip the student if there is no data or no contract date
    if (!dataDict[studentKey] || !dataDict[studentKey]["Data umowy"]) {
      console.log(`Skip: missing key data for '${studentKey}'.`);
      return;
    }

    let newSheetName = `Invoice_${invoiceCount}`;
    let newSheet = spreadsheet.getSheetByName(newSheetName);
    if (!newSheet) {
      newSheet = invoiceDraft.copyTo(spreadsheet);
      newSheet.setName(newSheetName);
    }

    let currentDate = new Date();
    let formattedDate = `${String(currentDate.getDate()).padStart(2, '0')}.${String(currentDate.getMonth() + 1).padStart(2, '0')}.${currentDate.getFullYear()}`;
    let invoiceNumber = `UL/${currentDate.getFullYear()}/${String(currentDate.getMonth() + 1).padStart(2, '0')}/${invoiceCount}`;

    newSheet.getRange("J1").setValue(formattedDate);
    newSheet.getRange("F3").setValue(invoiceNumber);

    invoiceCount++;

    writeStudentData(newSheet, studentKey, dataDict, calculateDict, startDate, endDate);
  });

  function writeStudentData(sheet, studentKey, dataDict, calculateDict, startDate, endDate) {
    function formatDate(date) {
      let day = String(date.getDate()).padStart(2, '0');
      let month = String(date.getMonth() + 1).padStart(2, '0'); 
      let year = date.getFullYear();
      return `${day}.${month}.${year}`;
    }

    let formattedContractDate = formatDate(new Date(dataDict[studentKey]["Data umowy"]));
    let formattedStartDate = formatDate(new Date(startDate));
    let formattedEndDate = formatDate(new Date(endDate));

    let contactName = dataDict[studentKey]["Contact name"];
    sheet.getRange("B23").setValue(contactName);
    sheet.getRange("C24").setValue(dataDict[studentKey]["Adresa_1"]);
    sheet.getRange("C25").setValue(dataDict[studentKey]["Adresa_2"]);
    sheet.getRange("C26").setValue(dataDict[studentKey]["PESEL"]);
    sheet.getRange("C27").setValue(dataDict[studentKey]["NIP"]);
    sheet.getRange("C28").setValue(dataDict[studentKey]["REGON"]);
    sheet.getRange("C29").setValue(dataDict[studentKey]["telefon"]);
    sheet.getRange("C30").setValue(dataDict[studentKey]["e-mail"]);

    sheet.getRange("C35").setValue("Za wykonanie prace zgodne z umową od " + formattedContractDate + " (okres " + formattedStartDate + " - " + formattedEndDate + ")");
    sheet.getRange("H35").setValue(calculateDict[studentKey]["Тривалість"]);
    sheet.getRange("I35").setValue(calculateDict[studentKey]["Ціна"]);
    sheet.getRange("J35").setValue(calculateDict[studentKey]["Вартість"]);
  }

  function createDictionary(sheet) {
    let data = sheet.getDataRange().getValues();
    if (data.length < 2) return {};

    let headers = data[0];
    let dictionary = {};

    for (let i = 1; i < data.length; i++) {
      let row = data[i];
      let studentName = row[0];

      if (!studentName || studentName.toString().trim() === "") continue; // Skipping empty names

      let studentData = {};
      headers.forEach((header, index) => {
        if (index > 0) studentData[header] = row[index];
      });

      dictionary[studentName] = studentData;
    }

    return dictionary;
  }

  SpreadsheetApp.getUi().alert('✅ Invoices generated!!! You can check!!!');
}
