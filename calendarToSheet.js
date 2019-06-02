function calendarToSheet() {
  
  // Reset the spreadsheet, no information from previous version is stored
  var ssFile = DriveApp.getFilesByName('Automated National Calendar').next();
  var ss = SpreadsheetApp.open(ssFile);
  var sheet = ss.getSheets()[0];
  sheet.clear();
  
  // Constants
  const BOLD_TEXT_STYLE = SpreadsheetApp.newTextStyle().setBold(true).build();
  const WEEKEND_COLOR = '#b7e1cd';
  const MS_PER_DAY = 1000*60*60*24;
    
  // Get calendars, only include visible and self-owned calendars
  var allCalendars = CalendarApp.getAllCalendars();
  var calendars = [];
  for each (c in allCalendars) {
    if (!c.isHidden() && c.isOwnedByMe()) {
      calendars.push(c);
    }
  }
  
  // Create headers 'DATE - CALENDAR 1 - CALENDAR 2 - ...'
  var headerRow = 1;
  var headerCol = 1;
  sheet.getRange(headerRow,headerCol).setValue('DATE');
  for each (c in calendars) {
    headerCol++;
    sheet.getRange(headerRow,headerCol).setValue(c.getName().toUpperCase());
    sheet.getRange(headerRow,headerCol).setBackground(c.getColor());
  }
  sheet.getRange(headerRow,1,1,headerCol).setFontWeight('bold');
  sheet.setFrozenRows(headerRow);
  
  
  // Include events happening between calendarStart and calendarEnd, inclusive
  var calendarStart = new Date(2019,4,23);
  Logger.log(calendarStart.getDay());
  var calendarEnd = new Date();
  calendarEnd.setDate(new Date().getDate() + 60);
  
  // Create the date column
  var calendarStartRow = headerRow + 1;
  var currentRow = calendarStartRow;
  var currentDate = new Date(calendarStart.getTime());
  while (currentDate.getTime() <= calendarEnd.getTime()) {
    sheet.getRange(currentRow,1).setValue(currentDate.getDate() + '/' + (currentDate.getMonth()+1));
    if (currentDate.getDay() === 0 || currentDate.getDay() === 6) {
      sheet.getRange(currentRow,1).setBackground(WEEKEND_COLOR);
    }
    currentDate.setDate(currentDate.getDate() + 1);
    currentRow++;
  }
  sheet.getRange(calendarStartRow,1,currentRow - calendarStartRow).setHorizontalAlignment('right');
  sheet.setFrozenColumns(1);
  
  // Go through each calendar and add its events
  var col = 1;
  for each (c in calendars) {
    col++;
    for each (e in c.getEvents(calendarStart,calendarEnd)) {
      
      // Get row range, i.e. the number of days the event spans
      var startTime = e.getStartTime();
      var diffStartTime = Math.abs(startTime.getTime() - calendarStart.getTime());
      var diffStartDays = Math.floor(diffStartTime / (MS_PER_DAY));
      var startRow = calendarStartRow+diffStartDays;
      
      var endTime = e.getEndTime();
      var diffEndTime = Math.abs(endTime.getTime() - calendarStart.getTime());
      var diffEndDays = Math.floor(diffEndTime / (MS_PER_DAY));
      var endRow = calendarStartRow+diffEndDays;
      
      // Don't count the last day if it ends at 00:00
      if (endTime.getHours() == 0 && endTime.getMinutes() == 0) {
        endRow--;
      }
      
      // Create the event details
      var details;
      var time;
      
      if (e.isAllDayEvent()) {
        time = 'Heldag';
      } else {
        time = startTime.getHours();
        if (startTime.getMinutes() != 0) {
          time += ':' + startTime.getMinutes();
        }
        time += '-' + endTime.getHours();
        if (endTime.getMinutes() != 0) {
          time += ':' + endTime.getMinutes();
        }
      }
      
      details = e.getTitle() + ', ' + time;
      if (e.getLocation().length > 0) {
        details += ', ' + e.getLocation();
      }
      
      // Insert details after any existing event details in the cell
      for (var row = startRow; row <= endRow; row++) {
        var prevDetails = sheet.getRange(row,col).getValue();
        if (prevDetails.length > 0) {
          details = prevDetails + '. ' + details;
        }
        // Make the event title stand out
        details = SpreadsheetApp.newRichTextValue()
          .setText(details)
          .setTextStyle(0,e.getTitle().length,BOLD_TEXT_STYLE)
          .build();
        sheet.getRange(row,col).setRichTextValue(details)
      }
    }
  }
  
  // Hide dates older than today
  var today = new Date();
  var diffTime = Math.abs(today.getTime() - calendarStart.getTime());
  var diffDays = Math.floor(diffTime / (MS_PER_DAY));
  var todayRow = calendarStartRow+diffDays;
  sheet.showRows(1,sheet.getDataRange().getLastRow());
  sheet.hideRows(calendarStartRow,todayRow-calendarStartRow);
}
