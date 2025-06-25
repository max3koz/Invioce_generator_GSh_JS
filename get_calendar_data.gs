function getCalendarEvents() {
  var calendarId = "e_mail@gmail.com"; // !!! Add your google accout e_mail !!!
  var calendar = CalendarApp.getCalendarById(calendarId);
  
  if (!calendar) {
    Logger.log("Calendar not found. Check ID!");
    return;
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var priceSheet = ss.getSheetByName("Data");
  if (!priceSheet) {
    Logger.log("Sheet 'Data' not found!");
    return;
  }

  var startDate = priceSheet.getRange("B18").getValue();
  var endDate = priceSheet.getRange("B19").getValue();
  if (!(startDate instanceof Date) || !(endDate instanceof Date)) {
    Logger.log("Error: Invalid values ​​in B18 or B19!");
    return;
  }

  var sheet = ss.getSheetByName("Calculate");
  sheet.clear();

  var events = calendar.getEvents(startDate, endDate);
  var durationMap = {};

  events.forEach(function(event) {
    var startTime = event.getStartTime();
    var endTime = event.getEndTime();
    var duration = (endTime - startTime) / (1000 * 60 * 60);
    var title = event.getTitle();

    if (title.includes(":")) {
      var category = title.split(":")[0].trim();
      durationMap[category] = (durationMap[category] || 0) + duration;
    }
  });

  sheet.appendRow(["Учень", "Тривалість", "Ціна", "Вартість"]);
  for (var cat in durationMap) {
    sheet.appendRow([cat, durationMap[cat].toFixed(2)]);
  }
  sheet.appendRow([" ", " "]);

  var priceData = priceSheet.getRange(2, 1, priceSheet.getLastRow(), 2).getValues();
  var priceMap = {};
  priceData.forEach(row => {
    priceMap[row[0]] = row[1];
  });

  // Get all rows from A2:C and filter
  var allData = sheet.getRange("A2:C").getValues();
  var filteredData = allData.filter(row => row[0] && row[0].toString().trim() !== "");

  var updatedData = filteredData.map(row => {
    var studentName = row[0];
    var duration = parseFloat(row[1]) || 0;
    var pricePerHour = priceMap[studentName] || "No price";
    return [studentName, duration, pricePerHour];
  });

  sheet.getRange(2, 1, updatedData.length, 3).setValues(updatedData);

  // Add the calculated cost
  var finalData = updatedData.map(row => {
    var studentName = row[0];
    var duration = parseFloat(row[1]) || 0;
    var price = parseFloat(row[2]) || 0;
    var cost = (duration * price).toFixed(2).replace(".", ",");
    return [studentName, duration, price, cost];
  });

  sheet.getRange(2, 1, finalData.length, 4).setValues(finalData);
  sheet.getRange(2, 4, finalData.length, 1).setHorizontalAlignment("right");
}
