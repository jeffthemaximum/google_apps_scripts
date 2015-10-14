REVIEW_COLUMN_NAME = "REVIEW DATE";
PP_COLUMN_NAME = "Point Person";
LAST_NAME_COLUMN = "Last Name";
FIRST_NAME_COLUMN = "First Name";


// spreadSheet object constructor function
function spreadSheet(googleSpreadSheet){
  this.sheets = {}
  this.googleSpreadSheet = googleSpreadSheet;

  this.addSheet = function(sheet) {
    
    sheetName = sheet.sheetName;
    this.sheets[sheetName] = sheet;
  }
}

// sheet object constructor function
function shit(sheet, sheetName, spreadSheet) {
  this.sheetName = sheetName;
  this.spreadSheet = spreadSheet;
  this.sheet = sheet;
  this.numCols = sheet.getLastColumn();
  this.numRows = sheet.getLastRow();
  this.ppColumnNumber;
  this.meetingColumnNumber;
  this.studentLastNameColumnNumber;
  this.studentFirstNameColumnNumber;
  
  
  this.getColumnNumberByColumnTitle = function(title) {
    for (var i = 1; i <= this.numCols; i++) {
      //get first row
      var x = this.sheet.getRange(1, i, 1, this.numCols)
      // get cells
      var data = x.getValue();
      if (data == title) {
        return i;
      }
    }
  }
  
  this.checkUpcomingMeetings = function() {
    var range, 
        cell, 
        emailConfirmationCell, 
        emailConfirmationRange, 
        ppemail, firstName, 
        lastName, 
        calendarConfirmationRange, 
        calendarConfirmationCell,
        calendars,
        iepCalendar,
        calendarTitle,
        calendarDate;
    var dateNow = new Date();
    //get whole column of review dates
    for (var i = 2; i <= this.numRows; i++) {
      range = this.sheet.getRange(i, this.meetingColumnNumber, 1, 1);
      cell = range.getValue();
      
      //check if time between today and upcoming meeting is less than 2 weeks
      var dateMeeting = new Date(cell);
      var timeDiff =  dateMeeting.getTime() - dateNow.getTime();
      var diffDays = Math.ceil(timeDiff / (1000 * 3600 * 24)); 
      if (!(cell instanceof Date)) {
        diffDays = "No date entered for meeting";
      }
      //set value of cell to the right with days countdown
      this.sheet.getRange(i, this.meetingColumnNumber+1, 1, 1).setValue(diffDays);
      
      //get relevant values
      ppemail = this.sheet.getRange(i, this.ppColumnNumber, 1, 1).getValue();
      firstName = this.sheet.getRange(i, this.studentFirstNameColumnNumber, 1, 1).getValue();
      lastName = this.sheet.getRange(i, this.studentLastNameColumnNumber, 1, 1).getValue();
      
      //check value of countdown to see if it's less than 14
      if (diffDays < 14) {
        //check value of last column to see if email has not been sent
        emailConfirmationRange = this.sheet.getRange(i, this.numCols, 1, 1);
        emailConfirmationCell = emailConfirmationRange.getValue();
        //if email hasn't been sent, send email
        if (emailConfirmationCell != 'YES') {
          //send email
          
          //update confirmationCell
          emailConfirmationRange.setValue('YES');
        }
      }
      //make calendar event if it isn't made yet
      debugger;
      if (cell instanceof Date) { 
        calendarConfirmationRange = this.sheet.getRange(i, this.numCols-1, 1, 1);
        calendarConfirmationCell = calendarConfirmationRange.getValue();
        if (calendarConfirmationCell != 'YES') {
          calendars = CalendarApp.getCalendarsByName('IEP Calendar');
          iepCalendar = calendars[0];
          calendarTitle = 'IEP Meeting for ' + firstName + ' ' + lastName;
          calendarDate = dateMeeting;
          iepCalendar.createAllDayEvent(calendarTitle, calendarDate)
          calendarConfirmationRange.setValue('YES');
        }
      }
    }
  }
}

// makes sheet objects given an array of sheets
function instantiateSheets(sheetsArray, spreadSheet) {
  var newSheet, sheetName;
  for (var i = 0; i < sheetsArray.length; i++) {
    sheetName = sheetsArray[i].getSheetName();
    // make new sheet object
    newSheet = new shit(sheetsArray[i], sheetName, spreadSheet);
    // add sheet to list of sheets in spreadSheet object
    spreadSheet.addSheet(newSheet);
    // get IEP meeting column number
    newSheet.meetingColumnNumber = newSheet.getColumnNumberByColumnTitle(REVIEW_COLUMN_NAME);
    // get pointperson column number
    newSheet.ppColumnNumber = newSheet.getColumnNumberByColumnTitle(PP_COLUMN_NAME);
    //get student last name
    newSheet.studentLastNameColumnNumber = newSheet.getColumnNumberByColumnTitle(LAST_NAME_COLUMN);
    //get student first name
    newSheet.studentFirstNameColumnNumber = newSheet.getColumnNumberByColumnTitle(FIRST_NAME_COLUMN);
    // only check upcoming meetings on sheets with meetings entered
    if (typeof newSheet.meetingColumnNumber == 'number') {
      newSheet.checkUpcomingMeetings();
    }
  }
}

function myFunction() {
  var currentSheet;
  // get active ss
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // instantiate new spreadSheet object to store sheets
  var lucySheet = new spreadSheet(ss);
  // get all sheets
  var sheets = ss.getSheets();
  // instantiate sheet objects
  
  instantiateSheets(sheets, lucySheet);

}
