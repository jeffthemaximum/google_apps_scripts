REVIEW_COLUMN_NAME = "REVIEW DATE";
PP_COLUMN_NAME = "Point Person";
LAST_NAME_COLUMN = "Last Name";
FIRST_NAME_COLUMN = "First Name";
EMAIL_SENT_COLUMN = "Email Sent?";
MEETING_SCHEDULED_CONFIRMATION_COLUMN = "Meeting Scheduled Calendar event made?";
REVIEW_DATE_CONFIRMATION_COLUMN = "Review Date Calendar event made?";
MEETING_CONFIRMED_COLUMN = "Meeting Scheduled";
DAYS_TILL_MEETING_COLUMN = "Days Till Meeting";


// spreadSheet object constructor function
function spreadSheet(googleSpreadSheet){
  this.sheets = {};
  this.googleSpreadSheet = googleSpreadSheet;

  this.addSheet = function(sheet) {
    
    sheetName = sheet.sheetName;
    this.sheets[sheetName] = sheet;
  };
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
  this.meetingConfirmedColumnNumber;
  this.studentLastNameColumnNumber;
  this.studentFirstNameColumnNumber;
  this.emailConfirmationColumnNumber;
  this.meetingScheduledConfirmationColumnNumber;
  this.reviewDateConfirmationColumnNumber;
  this.daysTillMeetingColumnNumber;
  
  
  this.getColumnNumberByColumnTitle = function(title) {
    for (var i = 1; i <= this.numCols; i++) {
      //get first row
      var x = this.sheet.getRange(1, i, 1, this.numCols);
      // get cells
      var data = x.getValue();
      if (data == title) {
        return i;
      }
    }
  };
  
  this.checkUpcomingMeetings = function() {
    var range,
        cell,
        emailConfirmationCell,
        emailConfirmationRange,
        ppemail, 
        firstName,
        lastName,
        calendarMeetingScheduledConfirmationRange,
        calendarMeetingScheduledConfirmationCell,
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
      var countDown = this.sheet.getRange(i, this.daysTillMeetingColumnNumber, 1, 1);
      countDown.setValue(diffDays);
      
      //get relevant values
      ppemail = this.sheet.getRange(i, this.ppColumnNumber, 1, 1).getValue();
      firstName = this.sheet.getRange(i, this.studentFirstNameColumnNumber, 1, 1).getValue();
      lastName = this.sheet.getRange(i, this.studentLastNameColumnNumber, 1, 1).getValue();
      
      //check value of countdown to see if it's less than 14
      if (diffDays < 14) {
        //check value of last column to see if email has not been sent
        emailConfirmationRange = this.sheet.getRange(i, this.emailConfirmationColumnNumber, 1, 1);
        emailConfirmationCell = emailConfirmationRange.getValue();
        //if email hasn't been sent, send email
        if (emailConfirmationCell != 'YES') {
          //send email
          
          //update confirmationCell
          emailConfirmationRange.setValue('YES');
        }
      }

      //make calendar event for meeting scheduled if it isn't made yet
      if (cell instanceof Date) {
        calendarMeetingScheduledConfirmationRange = this.sheet.getRange(i, this.reviewDateConfirmationColumnNumber, 1, 1);
        calendarMeetingScheduledConfirmationCell = calendarMeetingScheduledConfirmationRange.getValue();
        if (calendarMeetingScheduledConfirmationCell != 'YES') {
          calendars = CalendarApp.getCalendarsByName('IEP Calendar');
          iepCalendar = calendars[0];
          calendarTitle = firstName + ' ' + lastName + ' SCHEDULED for IEP meeting';
          calendarDate = dateMeeting;
          iepCalendar.createAllDayEvent(calendarTitle, calendarDate);
          calendarMeetingScheduledConfirmationRange.setValue('YES');
        }
      }
    }
  };

  this.checkUpcomingScheduledMeetings = function() {


    var range,
        cell,
        emailConfirmationCell,
        emailConfirmationRange,
        ppemail, 
        firstName,
        lastName,
        calendarMeetingConfirmedConfirmationRange,
        calendarMeetingConfirmedConfirmationCell,
        calendars,
        iepCalendar,
        calendarTitle,
        calendarDate,
        countdown,
        ev;
    var dateNow = new Date();
    //get whole column of review dates
    for (var i = 2; i <= this.numRows; i++) {
      range = this.sheet.getRange(i, this.meetingConfirmedColumnNumber, 1, 1);
      cell = range.getValue();
      
      //get relevant values
      firstName = this.sheet.getRange(i, this.studentFirstNameColumnNumber, 1, 1).getValue();
      lastName = this.sheet.getRange(i, this.studentLastNameColumnNumber, 1, 1).getValue();
      
      var dateMeeting = new Date(cell);
      //make calendar event for meeting scheduled if it isn't made yet
      if (cell instanceof Date) {
        calendarMeetingConfirmedConfirmationRange = this.sheet.getRange(i, this.meetingScheduledConfirmationColumnNumber, 1, 1);
        calendarMeetingConfirmedConfirmationCell = calendarMeetingConfirmedConfirmationRange.getValue();
        if (calendarMeetingConfirmedConfirmationCell != 'YES') {
          calendars = CalendarApp.getCalendarsByName('IEP Calendar');
          iepCalendar = calendars[0];
          calendarTitle = firstName + ' ' + lastName + ' CONFIRMED for IEP meeting';
          calendarDate = dateMeeting;
          ev = iepCalendar.createAllDayEvent(calendarTitle, calendarDate);     
          //update ss
          calendarMeetingConfirmedConfirmationRange.setValue('YES');
        }
      }
    }

  };
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
    //get IEP confirmed meeting column number
    newSheet.meetingConfirmedColumnNumber = newSheet.getColumnNumberByColumnTitle(MEETING_CONFIRMED_COLUMN);
    // get pointperson column number
    newSheet.ppColumnNumber = newSheet.getColumnNumberByColumnTitle(PP_COLUMN_NAME);
    //get student last name
    newSheet.studentLastNameColumnNumber = newSheet.getColumnNumberByColumnTitle(LAST_NAME_COLUMN);
    //get student first name
    newSheet.studentFirstNameColumnNumber = newSheet.getColumnNumberByColumnTitle(FIRST_NAME_COLUMN);
    //get email confirmation column
    newSheet.emailConfirmationColumnNumber = newSheet.getColumnNumberByColumnTitle(EMAIL_SENT_COLUMN);
    //get meeting scheduled confirmation column number
    newSheet.meetingScheduledConfirmationColumnNumber = newSheet.getColumnNumberByColumnTitle(MEETING_SCHEDULED_CONFIRMATION_COLUMN);
    //get review date confirmation column number
    newSheet.reviewDateConfirmationColumnNumber = newSheet.getColumnNumberByColumnTitle(REVIEW_DATE_CONFIRMATION_COLUMN);
    //get days till meeting column number
    newSheet.daysTillMeetingColumnNumber = newSheet.getColumnNumberByColumnTitle(DAYS_TILL_MEETING_COLUMN);
    // only check upcoming meetings on sheets with meetings entered
    if (typeof newSheet.meetingColumnNumber == 'number') {
      newSheet.checkUpcomingMeetings();
      newSheet.checkUpcomingScheduledMeetings();
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
