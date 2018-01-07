// last updated for: Spring 2018
// by: MichaelH13
// note: don't forget to update column values in reverse alphabetical order.
var debugging = false;

var Event = function(start, end, isTask) {
    this.start = start;
    this.end = end;
    this.isTask = isTask;
    
    this.overLapsWithOtherEvent = function(otherEvent) {
        
        if (otherEvent.start == this.start ||                                                              // if we start at the same time, we overlap
            otherEvent.end == this.end ||                                                                  // if we end at the same time, we overlap
            (this.start < otherEvent.start && this.end < otherEvent.end && otherEvent.end > this.start) || // if we start after the otherEvent and end before it and we end after the other event has started
            (this.end < otherEvent.end && this.start > otherEvent.start) ||                                // if we end before the other event and start after it
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
    var sheetName = "EventTaskList";
    var ss = SpreadsheetApp.getActive().getSheetByName(sheetName);
    var numberOfTasks = ss.getRange("A2:A").getLastRow();
    var validationFormulas = new Array();
    var studentDropdownRange = ss.getRange("M2:M");
    var listOfStudentsRange = ss.getRange("P2:P");
    var query = ""
    var whereClause = ""
    var overrideColumnResult = ""
    
    // ROW P
    // set list students available before setting up dropdown selection
    for (var i = 2; i <= numberOfTasks; ++i) {
        overrideColumnResult = "IF($B" + i + "= \"\", $A" + i + ", $B" + i + ")"
        whereClause = "IF(" + overrideColumnResult + "=\"GSCI120\", \
        QUERY(StudentColumnNameMapping!$A$1:$B$100,\"SELECT A WHERE B = 'FullName'\") & \" <> ''\", \
        QUERY(StudentColumnNameMapping!$A$1:$B$100,\"SELECT A WHERE B = '\" & " + overrideColumnResult + " & \"'\") & \" = 'Yes'\")"
        query = "QUERY(StudentSchedules,\
        \"SELECT \" & QUERY(StudentColumnNameMapping!$A$1:$B$100,\"SELECT A WHERE B = 'FullName'\") & \" \
        WHERE \" & " + whereClause + " & \"  AND \" & QUERY(StudentColumnNameMapping!$A$1:$B$100,\"SELECT A WHERE B = 'HoursRemaining'\") & \" > 0\")"
        
        // this must put an array at formulas[i - 2]
        validationFormulas[i - 2] = [("=IFERROR(TRANSPOSE(" + query + "), \"NA\")")];
        
        whereClause = ""
        overrideColumnResult = ""
        query = ""
        Log(validationFormulas[i - 2]);
    }
    
    listOfStudentsRange.setFormulas(validationFormulas);
    
    validationFormulas = new Array();
    
    // set data validation for the EventTaskList (student selection)
    for (var i = 2; i <= numberOfTasks; ++i) {
        validationFormulas[i - 2] = [SpreadsheetApp.newDataValidation().requireValueInRange(ss.getRange("P" + i + ":LK" + i), true).setAllowInvalid(false).build()];
        Log("Setting M" + i);
    }
    
    studentDropdownRange.setDataValidations(validationFormulas);
}

// All Columns here must be created before formulas are set.
// columns:
// A  = ID;               = Row() - 2
// G  = FullName          = Trim(D) & " " & TRIM(E)
// AW = BIO-STOCKROOM     = Z
// AX = CHEM-STOCKROOM    = Z
// AY = MAIL              = V
// AZ = INVENTORY         = Z
// BA = ANIMAL ROOM       = Z
// BB = SPECIAL-ASSISTANT = Yes
// BC = LAUNDRY           = W
// BD = Custodial         = Yes
// BE = Admin-Assistant   = AA
// BF = DeptAssistant     = AA
// BG = HoursUsed         = IFERROR(IF($DG50 <> "", INDEX(QUERY(EventTaskList!$B$1:$L$1068, "SELECT SUM(" & QUERY(EventTaskColumnNameMapping!$A$1:$B$1000,"SELECT A WHERE B = 'Hours'") & ") WHERE " & QUERY(EventTaskColumnNameMapping!$A$1:$B$1000, "SELECT A WHERE B = 'Student'") & " = '" & D & "'"),2,0),""),0)
// BH = HoursRemaining    = K - BG
function setupStudentAvailability() {
    var ssProj = SpreadsheetApp.getActive();
    var sheets = ssProj.getSheets();
    var ss = null;
    
    for (var i = 0; i < ssProj.getNumSheets(); ++i) {
        var sheet = sheets[i];
        Log(sheet);
        
        if (sheet.getFormUrl() != null) {
            Log(sheet.getName());
            ss = sheet;
        }
    }
    
    // Don't blow up if the form isn't linked yet.
    if (ss == null) {
        alert("Form Not Linked", "The form is not linked to this spreadsheet project, please link the form to the sheet that contains the student responses before continuing.")
        return;
    }
    
    var numberOfStudents = ss.getRange("A2:A").getLastRow();
    var columnHeaders = ss.getRange(1, 1, 1, ss.getLastColumn());
    var headers = columnHeaders.getDisplayValues();
    
    Log("headers: " + headers);
    
    for (var i = 0; i < headers.count; ++i) {
        Log("header: " + headers[i]);
    }
    
    if (ss.getRange(1, 1).getDisplayValue() != "ID") {
        Log("ID COLUMN DOES NOT EXIST");
        ss.insertColumnBefore(1); // https://developers.google.com/apps-script/reference/spreadsheet/sheet
        ss.getRange(1, 1).setValue("ID");
        
        // Set data formatting.
        // https://stackoverflow.com/questions/38490800/set-cell-format-to-text-with-google-apps-script
        var column = ss.getRange("A2:A");
        column.setNumberFormat("@");
    } else {
        Log("ID COLUMN EXISTS");
    }
    
    // ROW A
    // set IDs column
    ss.getRange("A2:A").setValue("=Row() - 2");
    
    var formulas = new Array();
    // ROW G
    // build the FullName formulas
    for (var i = 2; i <= numberOfStudents; ++i) {
        
        // this must put an array at formulas[i - 2]
        formulas[i - 2] = [("=Trim(D" + i + ") & \" \" & Trim(E" + i + ")")];
        Log(formulas[i - 2]);
    }
    
    ss.getRange("G2:G").setFormulas(formulas);
    
    //formulas = new Array();
    //
    // build the CHEM110 = CHEM21(1/2) formulas
    //for (var i = 2; i <= numberOfStudents; ++i) {
    //
    // this must put an array at formulas[i - 2]
    //formulas[i - 2] = [("=AO" + i)];
    //if (debugging) {Logger.log(formulas[i - 2]);}
    //}
    //
    // ROW AM (CHEM 110)
    //ss.getRange("AM2:AM" + numberOfStudents).setFormulas(formulas);
    
    
    formulas = new Array();
    
    // build the Stockroom/Animal Room, Inventory formulas
    for (var i = 2; i <= numberOfStudents; ++i) {
        
        // this must put an array at formulas[i - 2]
        formulas[i - 2] = [("=Z" + i)];
        Log(formulas[i - 2]);
    }
    
    // ROW AW
    ss.getRange("AW2:AW").setFormulas(formulas);
    
    // ROW AX
    ss.getRange("AX2:AX").setFormulas(formulas);
    
    // ROW AZ
    ss.getRange("AZ2:AZ").setFormulas(formulas);
    
    // ROW BA
    ss.getRange("BA2:BA").setFormulas(formulas);
    
    formulas = new Array();
    
    // ROW AY
    // build the Mail formulas
    for (var i = 2; i <= numberOfStudents; ++i) {
        
        // this must put an array at formulas[i - 2]
        formulas[i - 2] = [("=V" + i)];
        Log(formulas[i - 2]);
    }
    
    // ROW AY
    ss.getRange("AY2:AY" + numberOfStudents).setFormulas(formulas);
    
    formulas = new Array();
    
    // ROW AW
    // build the Mail formulas
    for (var i = 2; i <= numberOfStudents; ++i) {
        
        // this must put an array at formulas[i - 2]
        formulas[i - 2] = [("=V" + i)];
        Log(formulas[i - 2]);
    }
    
    // ROW AW
    ss.getRange("AW2:AW" + numberOfStudents).setFormulas(formulas);
    
    formulas = new Array();
    
    // build the Special Assistant and Custodial array values
    for (var i = 2; i <= numberOfStudents; ++i) {
        
        // this must put an array at formulas[i - 2]
        formulas[i - 2] = [("Yes")];
        Log(formulas[i - 2]);
    }
    
    // ROW BB
    ss.getRange("BB2:BB").setValues(formulas);
    
    // ROW BD
    ss.getRange("BD2:BD").setValues(formulas);
    
    formulas = new Array();
    
    // build the ADMIN-ASSISTANT DeptAssistant array values
    for (var i = 2; i <= numberOfStudents; ++i) {
        
        // this must put an array at formulas[i - 2]
        formulas[i - 2] = [("=AA" + i)];
        Log(formulas[i - 2]);
    }
    
    // ROW BE
    ss.getRange("BE2:BE").setValues(formulas);
    
    // ROW BF
    ss.getRange("BF2:BF").setValues(formulas);
    
    formulas = new Array();
    
    // ROW BC
    // build the Laundry formulas
    for (var i = 2; i <= numberOfStudents; ++i) {
        
        // this must put an array at formulas[i - 2]
        formulas[i - 2] = [("=W" + i)];
        Log(formulas[i - 2]);
    }
    
    // ROW BC
    ss.getRange("BC2:BC" + numberOfStudents).setFormulas(formulas);
    
    formulas = new Array();
    
    // ROW BG
    // build the HoursUsed formulas
    for (var i = 2; i <= numberOfStudents; ++i) {
        
        // this must put an array at formulas[i - 2]
        formulas[i - 2] = [("=IFERROR(IF($G" + i + " <> \"\", INDEX(QUERY(EventTaskList!$A$1:$M$1000, \"SELECT SUM(\" & QUERY(EventTaskListColumnNameMapping!$A$1:$B$100,\"SELECT A WHERE B = 'Hours'\") & \") WHERE \" & QUERY(EventTaskListColumnNameMapping!$A$1:$B$100, \"SELECT A WHERE B = 'Student'\") & \" = '\" & $G" + i + " & \"'\"),2,0),\"\"),0)")];
        Log(formulas[i - 2]);
    }
    
    // ROW BG
    ss.getRange("BG2:BG" + numberOfStudents).setFormulas(formulas);
    
    formulas = new Array();
    
    // ROW BH
    // build the HoursRemaining formulas
    for (var i = 2; i <= numberOfStudents; ++i) {
        
        // this must put an array at formulas[i - 2]
        formulas[i - 2] = [("=IF($BG" + i + " <> \"\", $K" + i + " - $BG" + i + ", $K" + i + ")")];
        Log(formulas[i - 2]);
    }
    
    // ROW BH
    ss.getRange("BH2:BH" + numberOfStudents).setFormulas(formulas);
}

function Log(message) {
    if (debugging) { Logger.log(message); }
}

function alert(title, message) {
    var ui = SpreadsheetApp.getUi();
    var response = ui.prompt(title, message, ui.ButtonSet.OK);
}

