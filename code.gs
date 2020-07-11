// Sheets names
var SHEET_DATA = '_Data';
var SHEET_TASKS = 'Tasks';
var SHEET_CALENDAR = 'Calendar';
// _Data variables
var DATA_MEMBER_COL = 1;
var DATA_EMAIL_COL = 2;
var DATA_CAL_ID = [2,6];
// Tasks variables
var TASKS_MEMBER_INCREMENT = 5;
var TASKS_TITLES_COL = 1;
var TASKS_VALUES_COL = 2;
var TASKS_TASK_ROW = 1;
var TASKS_ROUTINE_ROW = 2;
var TASKS_MEMBER_ROW = 3;
var TASKS_DATE_ROW = 4;
var TASKS_START_ROW = 5;
var TASKS_END_ROW = 6;
var TASKS_COLLABORATORS_ROW = 7
var TASKS_EMAILS_COLL_COL = 5;
var TASKS_DESCRIPTION_ROW = 8;
var TASKS_LOCATION_ROW = 9;
var TASKS_SWITCH = [1,3];
var TASKS_DAYS = [4,3];
var TASKS_DURATION = [5,4];
var TASKS_ADD_TASK_BUTTON = [11,1];
var TASKS_ADD_ROUTINE_BUTTON = [12,1];

var DATE_CAPTION = 'Double click to pop calendar up';

/*
  Common functions
*/

function searchRowMember(member) {
  var data = SpreadsheetApp.getActive().getSheetByName(SHEET_DATA).getDataRange().getValues();

  for (var i = 0; i < data.length; i++)
    if (data[i][0] == member)
      return i+1;

  return -1;
}

function addToCalendar(event, date, start, end, member, collaborators, description, location) {
  var ui = SpreadsheetApp.getUi(); // gets user interface
  var ssCalendar = SpreadsheetApp.getActive().getSheetByName(SHEET_CALENDAR);
  var ssData = SpreadsheetApp.getActive().getSheetByName(SHEET_DATA);

  // Task
  if (date != null) {
    // Converting parameters into Date objects
    var d = new Date(date);
    var s = new Date(start);
    var e = new Date(end);

    // Creating actual start and end dates
    var startDate = new Date(d.getFullYear(), d.getMonth(), d.getDate(), s.getHours(), s.getMinutes(), s.getSeconds(), s.getMilliseconds());
    var endDate = new Date(d.getFullYear(), d.getMonth(), d.getDate(), e.getHours(), e.getMinutes(), e.getSeconds(), e.getMilliseconds());

    // Getting cell coordinates
    var day = (startDate.getDay() != 0) ? startDate.getDay()+1 : 8;
    var hour = startDate.getHours()+2;

    // One hour events
    if (endDate.getHours() - startDate.getHours() == 1)
      ssCalendar.getRange(hour, day).setValue(event);
    // More than one hour events
    else {
      var eventRange = ssCalendar.getRange(hour, day, endDate.getHours()-startDate.getHours());
      eventRange.mergeVertically();
      eventRange.setValue(event);
      eventRange.setHorizontalAlignment("center");
      eventRange.setVerticalAlignment("middle");
    }

    // creating Google Calendar event
    var rowMember = searchRowMember(member);
    var email = ssData.getRange(rowMember, DATA_EMAIL_COL).getValue();
    var calendarId = ssData.getRange(DATA_CAL_ID[0], DATA_CAL_ID[1]).getValue();
    var options = {
      'location': location,
      'description': description,
      'guests': email
    };

    var eventCal = CalendarApp.getCalendarById(calendarId);
    eventCal.createEvent(event, startDate, endDate, options);
  }
  // Routine
  else {

  }
}

function getLastRowRange(range) {
  return range.getValues().filter(String).length;
}

/*
  Button scripts
*/

function addNewMember() {
  var ui = SpreadsheetApp.getUi(); // gets user interface
  var ss = SpreadsheetApp.getActive(); // assign active spreadsheet to variable
  var ssData = ss.getSheetByName(SHEET_DATA);
  var ssTasks = ss.getSheetByName(SHEET_TASKS);
  var memberRange = ssData.getRange("A:A");

  // retrieving data from prompts
  var uiMember = ui.prompt('Insert member name');
  // validate data
  var member = uiMember.getResponseText();
  do {
    if (uiMember.getSelectedButton() != ui.Button.OK)
      return;

    if (searchRowMember(member) != -1) {
      ui.alert('Member "' + member + '" already exists, please choose a different name');
      uiMember = ui.prompt('Insert member name');
      member = uiMember.getResponseText();
    }
    else
      break;
  } while (true);

  uiMember = ui.prompt('Insert member email');
  if (uiMember.getSelectedButton() != ui.Button.OK)
    return;
  var email = uiMember.getResponseText();

  // once data has been retrieved and validated, insert data in sheets

  // creating task table
  var row = (10*(getLastRowRange(memberRange)))+TASKS_MEMBER_INCREMENT;
  var headers = [['Member', 'Task', 'Value']];
  var headersRange = ssTasks.getRange(row,1,1,3);

  // inserting and formatting header data
  headersRange.setValues(headers);
  headersRange.setHorizontalAlignment("center");
  headersRange.setFontWeight("bold");

  // inserting and formatting member data
  var nameRange = ssTasks.getRange(row+1,1,8);
  nameRange.mergeVertically();
  nameRange.setValue(member);
  nameRange.setHorizontalAlignment("center");
  nameRange.setVerticalAlignment("middle");
  nameRange.setFontWeight("bold");

  // inserting member in _Data
  var rowIndex = getLastRowRange(memberRange)+1;
  ssData.getRange(rowIndex,DATA_MEMBER_COL).setValue(member);
  ssData.getRange(rowIndex,DATA_EMAIL_COL).setValue(email);
}

function switchTaskRoutine() {
  var ui = SpreadsheetApp.getUi(); // gets user interface
  var ssTasks = SpreadsheetApp.getActive().getSheetByName(SHEET_TASKS);

  var switchCaption = ssTasks.getRange(TASKS_SWITCH[0], TASKS_SWITCH[1])
  var disable = SpreadsheetApp.newDataValidation().requireTextEqualTo('null').setAllowInvalid(false).setHelpText("Do not edit in print sheet").build();
  // switch to routine
  if (switchCaption.getValue().includes('routine')) {
    // tasks controls
    ssTasks.getRange(TASKS_TASK_ROW, TASKS_TITLES_COL, 1, 2).setBackground('#dedede').setDataValidation(disable);
    ssTasks.getRange(TASKS_DATE_ROW, TASKS_TITLES_COL, 1, 2).setBackground('#dedede');
    ssTasks.getRange(TASKS_ADD_TASK_BUTTON[0], TASKS_ADD_TASK_BUTTON[1]).setBackground('#dedede');
    // routine controls
    ssTasks.getRange(TASKS_ROUTINE_ROW, TASKS_TITLES_COL, 1, 2).setBackground('white').setDataValidation(null);
    ssTasks.getRange(TASKS_DAYS[0], TASKS_DAYS[1], 1, 4).setBackground('white');
    ssTasks.getRange(TASKS_DURATION[0], TASKS_DURATION[1]-1, 1, 3).setBackground('white');
    ssTasks.getRange(TASKS_ADD_ROUTINE_BUTTON[0], TASKS_ADD_ROUTINE_BUTTON[1]).setBackground('white');

    switchCaption.setValue('Switch to task');
  }
  // switch to tasks
  else {
    // tasks controls
    ssTasks.getRange(TASKS_TASK_ROW, TASKS_TITLES_COL, 1, 2).setBackground('white').setDataValidation(null);
    ssTasks.getRange(TASKS_DATE_ROW, TASKS_TITLES_COL, 1, 2).setBackground('white');
    ssTasks.getRange(TASKS_ADD_TASK_BUTTON[0], TASKS_ADD_TASK_BUTTON[1]).setBackground('white');
    // routine controls
    ssTasks.getRange(TASKS_ROUTINE_ROW, TASKS_TITLES_COL, 1, 2).setBackground('#dedede').setDataValidation(disable);
    ssTasks.getRange(TASKS_DAYS[0], TASKS_DAYS[1], 1, 4).setBackground('#dedede');
    ssTasks.getRange(TASKS_DURATION[0], TASKS_DURATION[1]-1, 1, 3).setBackground('#dedede');
    ssTasks.getRange(TASKS_ADD_ROUTINE_BUTTON[0], TASKS_ADD_ROUTINE_BUTTON[1]).setBackground('#dedede');

    switchCaption.setValue('Switch to routine');
  }
}

function deleteMember() {
  var ui = SpreadsheetApp.getUi(); // gets user interface
  var ss = SpreadsheetApp.getActive(); // assign active spreadsheet to variable
  var ssData = ss.getSheetByName(SHEET_DATA);
  var ssTasks = ss.getSheetByName(SHEET_TASKS);
  var member = ssTasks.getRange(TASKS_MEMBER_ROW,TASKS_VALUES_COL).getValue();

  // Validating data
  var found = false;
  if (member == '') {
    ui.alert('No member selected', 'Choose a member from cell B3', ui.ButtonSet.OK);
    return;
  }

  var rowMember = searchRowMember(member)-1;
  if (rowMember != -1) {
    if (ui.alert('Do you really want to delete "' + member + '"?', '', ui.ButtonSet.YES_NO) == ui.Button.YES) {
      // deletes data in _Data
      ssData.getRange(rowMember+1,1,1,3).deleteCells(SpreadsheetApp.Dimension.ROWS);

      // deletes data in Tasks
      ssTasks.getRange(TASKS_MEMBER_ROW,TASKS_VALUES_COL).setValue('');
      ssTasks.getRange(10*rowMember+TASKS_MEMBER_INCREMENT,1,10,3).deleteCells(SpreadsheetApp.Dimension.ROWS);
    }
  }
  else
    ui.prompt('No member found', 'Make sure you choose the member within the options the dropdown list gives you', ui.ButtonSet.OK);
}

function addTask() {
  var ui = SpreadsheetApp.getUi(); // gets user interface
  // MODIFIED
  // var ss = ; // assign active spreadsheet to variable
  var ssTasks = SpreadsheetApp.getActive().getSheetByName(SHEET_TASKS);

  var task = ssTasks.getRange(TASKS_TASK_ROW,TASKS_VALUES_COL).getValue();
  var member = ssTasks.getRange(TASKS_MEMBER_ROW,TASKS_VALUES_COL).getValue();
  var date = ssTasks.getRange(TASKS_DATE_ROW,TASKS_VALUES_COL).getValue();
  var start = ssTasks.getRange(TASKS_START_ROW,TASKS_VALUES_COL).getValue();
  var end = ssTasks.getRange(TASKS_END_ROW,TASKS_VALUES_COL).getValue();
  var collaborators = ssTasks.getRange(TASKS_COLLABORATORS_ROW, TASKS_EMAILS_COLL_COL).getValue();
  var description = ssTasks.getRange(TASKS_DESCRIPTION_ROW, TASKS_VALUES_COL).getValue();
  var location = ssTasks.getRange(TASKS_LOCATION_ROW, TASKS_VALUES_COL).getValue();

  var isValid = true;
  if (task == '') {
    ui.alert('Missing task');
    isValid = false;
  }
  if (member == '') {
    ui.alert('Missing member');
    isValid = false;
  }
  if (date == '' || date == DATE_CAPTION) {
    ui.alert('Missing date');
    isValid = false;
  }
  if (start == '') {
    ui.alert('Missing start time');
    isValid = false;
  }
  if (end == '') {
    ui.alert('Missing end time');
    isValid = false;
  }
  if (new Date(start).getHours() > new Date(end).getHours()) {
    ui.alert('Start hour greater than end hour');
    isValid = false;
  }

  if (!isValid)
    return;

  // All data is valid, proceed to manage it
  var rowMember = searchRowMember(member)-1;
  if (rowMember != -1) {
    var row = 10*rowMember+TASKS_MEMBER_INCREMENT;
    var tasksRange = ssTasks.getRange(row,TASKS_VALUES_COL,9);
    var noRows = getLastRowRange(tasksRange);

    // setting task
    ssTasks.getRange(row+noRows,TASKS_VALUES_COL).setValue(task);

    // Google Calendar
    addToCalendar(task, date, start, end, member, collaborators, description, location);

    // Resetting controls
    ssTasks.getRange(TASKS_TASK_ROW,TASKS_VALUES_COL).setValue('');
    ssTasks.getRange(TASKS_ROUTINE_ROW,TASKS_VALUES_COL).setValue('');
    ssTasks.getRange(TASKS_DATE_ROW,TASKS_VALUES_COL).setValue(DATE_CAPTION);
    ssTasks.getRange(TASKS_START_ROW,TASKS_VALUES_COL).setValue('');
    ssTasks.getRange(TASKS_END_ROW,TASKS_VALUES_COL).setValue('');
    ssTasks.getRange(TASKS_COLLABORATORS_ROW,TASKS_VALUES_COL).setValue('');
    ssTasks.getRange(TASKS_COLLABORATORS_ROW,TASKS_EMAILS_COLL_COL).setValue('');
    ssTasks.getRange(TASKS_DESCRIPTION_ROW,TASKS_VALUES_COL).setValue('');
    ssTasks.getRange(TASKS_LOCATION_ROW,TASKS_VALUES_COL).setValue('');
  }
  else
    ui.prompt('No member found', 'Make sure you choose the member within the options the dropdown list gives you', ui.ButtonSet.OK);
}
