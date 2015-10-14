COLUMN_NAME = "REVIEW DATE";

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
  this.meetingColumnNumber;
  
  this.getColumnNumberByColumnTitle = function(title) {
    for (var i = 1; i <= this.numCols; i++) {
      //get first row
      var x = this.sheet.getRange(1, i, 1, this.numCols)
      // get cells
      var data = x.getValue();
      if (data == title) {
        this.meetingColumnNumber = i;
        return i;
      }
    }
  }
  
  this.checkUpcomingMeetings = function() {
    var range, cell;
    var dateNow = new Date();
    //get whole column of review dates
    for (var i = 2; i <= this.numRows; i++) {
      range = this.sheet.getRange(i, this.meetingColumnNumber, 1, 1);
      cell = range.getValue();
      
      //check if time between today and upcoming meeting is less than 2 weeks
      var dateMeeting = new Date(cell);
      var timeDiff = Math.abs(dateNow.getTime() - dateMeeting.getTime());
      var diffDays = Math.ceil(timeDiff / (1000 * 3600 * 24)); 
      if (!(cell instanceof Date)) {
        diffDays = "No date entered for meeting";
      }
      //set value of cell to the right with days countdown
      this.sheet.getRange(i, this.meetingColumnNumber+1, 1, 1).setValue(diffDays);
    }
    
    
    
    //check if email has been sent
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
    var col = newSheet.getColumnNumberByColumnTitle(COLUMN_NAME);
    // only check upcoming meetings on sheets with meetings entered
    if (typeof col == 'number') {
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
