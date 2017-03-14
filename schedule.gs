// last updated for: Fall 2017
// by: MichaelH13
// note: don't forget to update column values in reverse alphabetical order.
var Event = function(start, end, isTask) {
  this.start = start;
  this.end = end;
  this.isTask = isTask;

  this.overLapsWithOtherEvent = function(otherEvent) {

    if (otherEvent.start == this.start ||                                                              // if we start at the same time, we overlap
        otherEvent.end == this.end ||                                                                  // if we end at the same time, we overlap
        (this.start < otherEvent.start && this.end < otherEvent.end && otherEvent.end > this.start) || // if we start after the otherEvent and end before it and we end after the other event has started
        (this.end < otherEvent.end && this.start > otherEvent.start) ||                                // if we end before the otherEvent and start after it
        (this.start > otherEvent.start && this.end > otherEvent.end && otherEvent.start < this.end)    // if we start after the otherEvent and end after it and it starts before we end
      ) {
      return true;
    }

    return false;
  }
};

function onOpen() {
  setupStudentAvailability();
  setupEventTaskList();
}

function setupEventTaskList() {

  var debugging = false;
  var numberOfTasks = 1000;
  var sheetName = "EventTaskList";
  var ss = SpreadsheetApp.getActive().getSheetByName(sheetName);
  var validationFormulas = new Array()
  var studentDropdownRange = ss.getRange("L2:L" + numberOfTasks);
  var listOfStudentsRange = ss.getRange("N2:N" + numberOfTasks);

  // ROW N
  // set list students available before setting up dropdown selection
  for (var i = 2; i <= numberOfTasks; ++i) {

    // this must put an array at formulas[i - 2]
    validationFormulas[i - 2] = [("=IFERROR(TRANSPOSE(IF($A" + i + "<>\"GSCI120\", QUERY(StudentSchedules,\"SELECT \" & QUERY(StudentColumnNameMapping!$A$1:$B$100,\"SELECT A WHERE B = 'FullName'\") & \" WHERE \" & QUERY(StudentColumnNameMapping!$A$1:$B$100,\"SELECT A WHERE B = '\" & A" + i + " & \"'\") & \" = 'Yes' AND \" & QUERY(StudentColumnNameMapping!$A$1:$B$100,\"SELECT A WHERE B = 'HoursRemaining'\") & \" > 0\"), QUERY(StudentSchedules,\"SELECT \" & QUERY(StudentColumnNameMapping!$A$1:$B$100,\"SELECT A WHERE B = 'FullName'\") & \" WHERE \" & QUERY(StudentColumnNameMapping!$A$1:$B$100,\"SELECT A WHERE B = 'FullName'\") & \" <> '' AND \" & QUERY(StudentColumnNameMapping!$A$1:$B$100,\"SELECT A WHERE B = 'HoursRemaining'\") & \" > 0\"))),\"NA\")")];
    if (debugging) {Logger.log(validationFormulas[i - 2]);}
  }

  listOfStudentsRange.setFormulas(validationFormulas);

  validationFormulas = new Array();

  // set data validation for the EventTaskList (student selection)
  for (var i = 2; i <= numberOfTasks; ++i) {
    validationFormulas[i - 2] = [SpreadsheetApp.newDataValidation().requireValueInRange(ss.getRange("N" + i + ":LJ" + i), true).setAllowInvalid(false).build()];
    if (debugging) {Logger.log("Setting L" + i);}
  }

  studentDropdownRange.setDataValidations(validationFormulas);
}

// columns:
// A  = ID;               = Row()
// F  = FullName          = Trim(D) & " " & TRIM(E)
// AZ = BIO-STOCKROOM     = Z
// BA = CHEM-STOCKROOM    = Z
// BB = MAIL              = V
// BC = INVENTORY         = Z
// BD = ANIMAL ROOM       = Z
// BE = SPECIAL-ASSISTANT = Yes
// BF = LAUNDRY           = W
// BG = Custodial         = Yes
// BH = Admin-Assistant   = AA
// BI = DeptAssistant     = AA
// BJ = HoursUsed         = IFERROR(IF($D50 <> "", INDEX(QUERY(EventTaskList!$B$1:$L$1068, "SELECT SUM(" & QUERY(EventTaskColumnNameMapping!$A$1:$B$1000,"SELECT A WHERE B = 'Hours'") & ") WHERE " & QUERY(EventTaskColumnNameMapping!$A$1:$B$1000, "SELECT A WHERE B = 'Student'") & " = '" & D & "'"),2,0),""),0)
// BK = HoursRemaining    = K - BJ
function setupStudentAvailability() {
  var numberOfStudents = 300;
  var sheetName = "StudentAvailability";
  var ss = SpreadsheetApp.getActive().getSheetByName(sheetName);
  var formulas = new Array();
  var debugging = false;

  // ROW A
  // set IDs column
  ss.getRange("A" + 2 + ":A" + numberOfStudents).setValue("=Row() - 2");


  // ROW F
  // build the FullName formulas
  for (var i = 2; i <= numberOfStudents; ++i) {

    // this must put an array at formulas[i - 2]
    formulas[i - 2] = [("=Trim(D" + i + ") & \" \" & Trim(E" + i + ")")];
    if (debugging) {Logger.log(formulas[i - 2]);}
  }

  ss.getRange("F2:F" + numberOfStudents).setFormulas(formulas);

  formulas = new Array();

  // build the CHEM110 = CHEM21(1/2) formulas
  for (var i = 2; i <= numberOfStudents; ++i) {

    // this must put an array at formulas[i - 2]
    formulas[i - 2] = [("=AO" + i)];
    if (debugging) {Logger.log(formulas[i - 2]);}
  }

  // ROW AM (CHEM 110)
  ss.getRange("AM2:AM" + numberOfStudents).setFormulas(formulas);


  formulas = new Array();

  // build the Stockroom/Animal Room, Inventory formulas
  for (var i = 2; i <= numberOfStudents; ++i) {

    // this must put an array at formulas[i - 2]
    formulas[i - 2] = [("=Z" + i)];
    if (debugging) {Logger.log(formulas[i - 2]);}
  }

  // ROW AZ
  ss.getRange("AZ2:AZ" + numberOfStudents).setFormulas(formulas);

  // ROW BA
  ss.getRange("BA2:BA" + numberOfStudents).setFormulas(formulas);

  // ROW BC
  ss.getRange("BC2:BC" + numberOfStudents).setFormulas(formulas);

  // ROW BD
  ss.getRange("BD2:BD" + numberOfStudents).setFormulas(formulas);

  formulas = new Array();

  // ROW BB
  // build the Mail formulas
  for (var i = 2; i <= numberOfStudents; ++i) {

    // this must put an array at formulas[i - 2]
    formulas[i - 2] = [("=V" + i)];
    if (debugging) {Logger.log(formulas[i - 2]);}
  }

  // ROW BB
  ss.getRange("BB2:BB" + numberOfStudents).setFormulas(formulas);

  formulas = new Array();

  // build the Special Assistant and Custodial array values
  for (var i = 2; i <= numberOfStudents; ++i) {

    // this must put an array at formulas[i - 2]
    formulas[i - 2] = [("Yes")];
    if (debugging) {Logger.log(formulas[i - 2]);}
  }

  // ROW BE
  ss.getRange("BE2:BE" + numberOfStudents).setValues(formulas);

  // ROW BG
  ss.getRange("BG2:BG" + numberOfStudents).setValues(formulas);

  formulas = new Array();

  // build the ADMIN-ASSISTANT DeptAssistant array values
  for (var i = 2; i <= numberOfStudents; ++i) {

    // this must put an array at formulas[i - 2]
    formulas[i - 2] = [("=AA" + i)];
    if (debugging) {Logger.log(formulas[i - 2]);}
  }

  // ROW BH
  ss.getRange("BH2:BH" + numberOfStudents).setValues(formulas);

  // ROW BI
  ss.getRange("BI2:BI" + numberOfStudents).setValues(formulas);

  formulas = new Array();

  // ROW BF
  // build the Laundry formulas
  for (var i = 2; i <= numberOfStudents; ++i) {

    // this must put an array at formulas[i - 2]
    formulas[i - 2] = [("=W" + i)];
    if (debugging) {Logger.log(formulas[i - 2]);}
  }

  // ROW BF
  ss.getRange("BF2:BF" + numberOfStudents).setFormulas(formulas);

  formulas = new Array();

  // ROW BJ
  // build the HoursUsed formulas
  for (var i = 2; i <= numberOfStudents; ++i) {

    // this must put an array at formulas[i - 2]
    formulas[i - 2] = [("=IFERROR(IF($F" + i + " <> \"\", INDEX(QUERY(EventTaskList!$A$1:$L$1000, \"SELECT SUM(\" & QUERY(EventTaskListColumnNameMapping!$A$1:$B$100,\"SELECT A WHERE B = 'Hours'\") & \") WHERE \" & QUERY(EventTaskListColumnNameMapping!$A$1:$B$100, \"SELECT A WHERE B = 'Student'\") & \" = '\" & $F" + i + " & \"'\"),2,0),\"\"),0)")];
    if (debugging) {Logger.log(formulas[i - 2]);}
  }

  // ROW BJ
  ss.getRange("BJ2:BJ" + numberOfStudents).setFormulas(formulas);

  formulas = new Array();

  // ROW BK
  // build the HoursRemaining formulas
  for (var i = 2; i <= numberOfStudents; ++i) {

    // this must put an array at formulas[i - 2]
    formulas[i - 2] = [("=IF($BJ" + i + " <> \"\", $K" + i + " - $BJ" + i + ", $K" + i + ")")];
    if (debugging) {Logger.log(formulas[i - 2]);}
  }

  // ROW BK
  ss.getRange("BK2:BK" + numberOfStudents).setFormulas(formulas);
}
  
