//<editor-fold> Constants
// sheets names
const SHEET_DATA = '_Data';
const SHEET_TASKS = 'Tasks';
const SHEET_CALENDAR = 'Calendar';
const SHEET_HISTORY = 'History';
// SHEET_DATA constants
const DATA_MEMBER = "A2";
const DATA_HISTORY = "C2";
const DATA_MEMBER_COL = 1;
const DATA_EMAIL_COL = 2;
const DATA_HISTORY_COL = 3;
const DATA_CAL_ID = 'F3';
const DATA_APPROVE_VAL = 'F5';
const DATA_EVENT = 'H2';
const DATA_TASKROUTINE = 'I2';
const DATA_WEEKS = 'R2';
const DATA_EVENT_COL = 8;
const DATA_INITIAL_EVENT_ROW = 3;
const DATA_DEFAULT_APPROVE_VAL = 60;
// SHEET_TASKS constants
const TASKS_MEMBER_INCREMENT = 5;
const TASKS_TITLES_COL = 1;
const TASKS_VALUES_COL = 2;
const TASKS_TASK = 'B1';
const TASKS_TASK_ROW = 1;
const TASKS_ROUTINE = 'B2';
const TASKS_ROUTINE_ROW = 2;
const TASKS_MEMBER = 'B3';
const TASKS_DATE = 'B4';
const TASKS_DATE_ROW = 4;
const TASKS_START = 'B5';
const TASKS_END = 'B6';
const TASKS_COLLABORATORS = 'B7';
const TASKS_EMAILS_COLLABORATORS = 'E7';
const TASKS_DESCRIPTION = 'B8';
const TASKS_LOCATION = 'B9';
const TASKS_SWITCH = 'C1';
const TASKS_DAYS = [4,3];
const TASKS_DAYS_DROPDOWN = 'D4';
const TASKS_DAYS_CHOSEN = 'F4';
const TASKS_DURATION = [5,4];
const TASKS_ADD_TASK_BUTTON = 'A11';
const TASKS_ADD_ROUTINE_BUTTON = 'A12';
const DATE_CAPTION = 'Double click to pop calendar up';
const DAYS = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
const TASKS_VALUE_COLUMN = 'C';
const TASKS_CHECKBOX_COLUMN = 'D';
const TASKS_CHECKBOX_COL = 4;
const TASKS_H_M_COLUMN = 'E';
const TASKS_ACHIEVEMENT_COLUMN = 'F';
const TASKS_TOTAL_COLUMN = 'G';
const NUM_TASKS = 8;
const TASKS_VALUES_TASKS_COL = 3;
const TASKS_NON_FIX_VALUES = 'A13';
// SHEET_CALENDAR constants
const CALENDAR_INITIAL_DATE = 'B2';
const CALENDAR_FINAL_DATE = 'G25';
// SHEET_HISTORY constants
const HISTORY_MEMBER_COL = 1;
// calendar options
const SEND_INVITES = true;
// coder info
const EMAIL = 'eduardo.mendozamartinez@aiesec.net';
//</editor-fold>

//<editor-fold> Common functions

  //<editor-fold> Search
function searchRowMember(member) {
  var data = SpreadsheetApp.getActive().getSheetByName(SHEET_DATA).getDataRange().getValues();

  for (var i = 0; i < data.length; i++)
    if (data[i][0] == member)
      return i+1;

  return -1;
}

function searchRowEmail(email) {
  var data = SpreadsheetApp.getActive().getSheetByName(SHEET_DATA).getDataRange().getValues();

  for (var i = 0; i < data.length; i++)
    if (data[i][1] == email)
      return i+1;

  return -1;
}

function searchEvent(event, tr) {
  var ssData = SpreadsheetApp.getActive().getSheetByName(SHEET_DATA);
  var eventsRange = ssData.getRange(DATA_EVENT + ":" + DATA_TASKROUTINE[0]);

  var data = eventsRange.getValues();
  for (var i = 0; i < data.length; i++)
    if (data[i][0] == event && data[i][1] == tr)
      return i+DATA_INITIAL_EVENT_ROW-1;

  return -1;
}

function getIndexOf(element, collection) {
  for (var i = 0; i < collection.length; i++)
    if (collection[i] == element)
      return i;

  return -1;
}

  //</editor-fold>

  //<editor-fold> Calendars
function addToCalendar(event, startDate, endDate) {
  var ssCalendar = SpreadsheetApp.getActive().getSheetByName(SHEET_CALENDAR);

  // Getting cell coordinates
  var day = (startDate.getDay() != 0) ? startDate.getDay()+1 : 8;
  var hour = startDate.getHours()+2;

  var eventRange = null;
  // One hour events
  if (endDate.getHours() - startDate.getHours() == 1) {
    eventRange = ssCalendar.getRange(hour, day);
    eventRange.setValue((eventRange.getValue() == '') ? event : eventRange.getValue() + ";" + event);
  }
  // More than one hour events
  else {
    eventRange = ssCalendar.getRange(hour, day, endDate.getHours()-startDate.getHours());
    eventRange.mergeVertically();
    eventRange.setValue((eventRange.getValue() == '') ? event : eventRange.getValue() + ";" + event);
    eventRange.setHorizontalAlignment("center");
    eventRange.setVerticalAlignment("middle");
  }

  return eventRange;
}

function addToGoogleCalendar(event, date, start, end, member, collaborators, description, location) {
  var ui = SpreadsheetApp.getUi(); // gets user interface
  var ssData = SpreadsheetApp.getActive().getSheetByName(SHEET_DATA);

  var calendarId = ssData.getRange(DATA_CAL_ID).getValue();
  if (calendarId == '') {
    ui.alert("🤔", "There is no calendar ID in " + SHEET_DATA + "!" + DATA_CAL_ID + "\nMake sure to set this up in order to arrange the tasks you give in Google Calendar(:", ui.ButtonSet.OK);
    return;
  }

  // Converting parameters into Date objects
  var d = new Date(date);
  var s = new Date(start);
  var e = new Date(end);

  // Creating actual start and end dates
  var startDate = new Date(d.getFullYear(), d.getMonth(), d.getDate(), s.getHours(), s.getMinutes(), s.getSeconds(), s.getMilliseconds());
  var endDate = new Date(d.getFullYear(), d.getMonth(), d.getDate(), e.getHours(), e.getMinutes(), e.getSeconds(), e.getMilliseconds());

  addToCalendar(event, startDate, endDate);

  // creating Google Calendar event
  var rowMember = searchRowMember(member);
  var email = ssData.getRange(rowMember, DATA_EMAIL_COL).getValue();

  // checking again collaborators
  if (collaborators.includes(email)) {
    ui.alert('🙃', 'did ya really insisted on havin the same member as his own collaborator???\nit\'s okay, nevermind I gotcha', ui.ButtonSet.OK);
    email = collaborators;
    collaborators = '';
  }

  var options = {
    'location': location,
    // 'description': (description == '') ? 'No description' : description,
    'description': description,
    'guests': (collaborators == '') ? email : email + ',' + collaborators,
    'sendInvites': (SEND_INVITES) ? 'True' : 'False'
  };

  var eventCal = CalendarApp.getCalendarById(calendarId);
  eventCal.createEvent(event, startDate, endDate, options);
}
  //</editor-fold>

  //<editor-fold> Other
function getLastRowRange(range) {
  return range.getValues().filter(String).length;
}

function getNumElements(collection, separator) {
  if (collection == '')
    return 0;

  var num = 1;

  for (var i = 0; i < collection.length; i++)
    if (collection[i] == separator)
      num++;

  return num;
}

function getMembershipNumber() {
  var data = SpreadsheetApp.getActive().getSheetByName(SHEET_DATA).getDataRange().getValues();

  for (var i = 0; i < data.length; i++)
    if (data[i][0] == '' && data[i][1] == '')
      return i-1;

  return 0;
}

function getColumnNumber(chr) {
  return chr.toLowerCase().charCodeAt(0) - 97 + 1;
}
  //</editor-fold>

function setValues(row, noRows) {
  var ssTasks = SpreadsheetApp.getActive().getSheetByName(SHEET_TASKS);

  if (noRows == 1) {
    ssTasks.getRange(row+noRows, TASKS_VALUES_TASKS_COL).setValue(1);
  }
  else if (noRows == 2 && ssTasks.getRange(row+noRows-1, TASKS_VALUES_TASKS_COL).getValue() == 1) {
    ssTasks.getRange(row+noRows-1, TASKS_VALUES_TASKS_COL).setValue(0.5);
    ssTasks.getRange(row+noRows, TASKS_VALUES_TASKS_COL).setValue(0.5);
  }
  else if (noRows == 4 && ssTasks.getRange(row+noRows-1, TASKS_VALUES_TASKS_COL).getValue() == 0 && ssTasks.getRange(row+noRows-2, TASKS_VALUES_TASKS_COL).getValue() == 0.5) {
    ssTasks.getRange(row+noRows-3, TASKS_VALUES_TASKS_COL).setValue(0.25);
    ssTasks.getRange(row+noRows-2, TASKS_VALUES_TASKS_COL).setValue(0.25);
    ssTasks.getRange(row+noRows-1, TASKS_VALUES_TASKS_COL).setValue(0.25);
    ssTasks.getRange(row+noRows, TASKS_VALUES_TASKS_COL).setValue(0.25);
  }
  else if (noRows == 5 && ssTasks.getRange(row+noRows-1, TASKS_VALUES_TASKS_COL).getValue() == 0.25 && ssTasks.getRange(row+noRows-2, TASKS_VALUES_TASKS_COL).getValue() == 0.25) {
    ssTasks.getRange(row+noRows-4, TASKS_VALUES_TASKS_COL).setValue(0.20);
    ssTasks.getRange(row+noRows-3, TASKS_VALUES_TASKS_COL).setValue(0.20);
    ssTasks.getRange(row+noRows-2, TASKS_VALUES_TASKS_COL).setValue(0.20);
    ssTasks.getRange(row+noRows-1, TASKS_VALUES_TASKS_COL).setValue(0.20);
    ssTasks.getRange(row+noRows, TASKS_VALUES_TASKS_COL).setValue(0.20);
  }
  else
    ssTasks.getRange(row+noRows, TASKS_VALUES_TASKS_COL).setValue(0);
}

function setTask(collaborators, task) {
  var ui = SpreadsheetApp.getUi(); // gets user interface
  var ssTasks = SpreadsheetApp.getActive().getSheetByName(SHEET_TASKS);
  var data = SpreadsheetApp.getActive().getSheetByName(SHEET_DATA).getDataRange().getValues();

  var members = [];
  collaborators.forEach(email => {
    var rowMember = searchRowEmail(email)

    if (rowMember != -1)
      members.push(data[rowMember-1][0])
    else
      ui.prompt(':(', 'There was an error trying to retrieve member data with the following email: "' + email + '"')
  });

  members.forEach(member => {
    var rowMember = searchRowMember(member)-1

    if (rowMember != -1) {
      // creating task
      var row = 10*rowMember+TASKS_MEMBER_INCREMENT
      var tasksRange = ssTasks.getRange(row,TASKS_VALUES_COL,9)
      var noRows = getLastRowRange(tasksRange)

      if (noRows == 9)
        ui.prompt('parece ser que el cuerpo aieseco solo resiste 8 tareas')
      else { // setting task
        ssTasks.getRange(row+noRows,TASKS_VALUES_COL).setValue(task)
        setValues(row, noRows)
      }
    }
  });
}

//</editor-fold>

/*
  Button scripts
*/

//<editor-fold> Button scripts

  //<editor-fold> Members
function addNewMember() {
  var ui = SpreadsheetApp.getUi(); // gets user interface
  var ss = SpreadsheetApp.getActive(); // assign active spreadsheet to variable
  var ssData = ss.getSheetByName(SHEET_DATA);
  var ssTasks = ss.getSheetByName(SHEET_TASKS);
  var ssHistory = ss.getSheetByName(SHEET_HISTORY);
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

  //<editor-fold> Tasks
  // creating task table
  var row = (10*(getLastRowRange(memberRange)))+TASKS_MEMBER_INCREMENT;
  var headers = [['Member', 'Task', 'Value', 'Fully done?', 'If not, how much?', 'Achievement', 'Total']];
  var headersRange = ssTasks.getRange(row,1,1,7);

  // inserting and formatting header data
  headersRange.setValues(headers);
  headersRange.setHorizontalAlignment("center");
  headersRange.setFontWeight("bold");

  // inserting and formatting member data
  var nameRange = ssTasks.getRange(row+1,1,NUM_TASKS);
  nameRange.mergeVertically();
  nameRange.setValue(member);
  nameRange.setHorizontalAlignment("center");
  nameRange.setVerticalAlignment("middle");
  nameRange.setFontWeight("bold");

  // inserting value number format
  ssTasks.getRange(row+1,3,NUM_TASKS).setNumberFormat('0.00%');

  // inserting checkboxes
  var checkboxesRange = ssTasks.getRange(row+1,4,NUM_TASKS);
  var enforceCheckbox = SpreadsheetApp.newDataValidation();
  enforceCheckbox.requireCheckbox();
  enforceCheckbox.setAllowInvalid(false);
  enforceCheckbox.build();
  checkboxesRange.setDataValidation(enforceCheckbox);

  // inserting 100% how much
  var hmRange = ssTasks.getRange(row+1,5,NUM_TASKS);
  hmRange.setValue('0');
  hmRange.setNumberFormat('0.00%');

  // inserting achievement
  for (var i = 0; i < NUM_TASKS; i++) {
    ssTasks.getRange(row+1+i,getColumnNumber(TASKS_ACHIEVEMENT_COLUMN)).setFormula('=IF(' + TASKS_CHECKBOX_COLUMN + (row+1+i).toString() + '=TRUE, ' + TASKS_VALUE_COLUMN + (row+1+i).toString() + ', ' + TASKS_H_M_COLUMN + (row+1+i).toString() + ')');
  }
  var achievementRange = ssTasks.getRange(row+1,6,NUM_TASKS);
  achievementRange.setNumberFormat('0.00%');

  // inserting total
  var totalRange = ssTasks.getRange(row+1,7,NUM_TASKS);
  totalRange.mergeVertically();
  totalRange.setFormula('=SUM(' + TASKS_ACHIEVEMENT_COLUMN + (row+1).toString() + ':' + TASKS_ACHIEVEMENT_COLUMN + (row+9).toString() + ')');
  totalRange.setNumberFormat('0.00%');
  totalRange.setHorizontalAlignment("center");
  totalRange.setVerticalAlignment("middle");
  totalRange.setFontWeight("bold");

  // creating ConditionalFormatting
  var appVal = 0.6;
  try {
    approves = ssData.getRange(DATA_APPROVE_VAL).getValue();
    appVal = (approves < 0) ? approves : approves/100;
  } catch (e) {
    ui.alert('Exception thrown when trying to retrieve value to approve from ' + SHEET_DATA + ' in cell ' + DATA_APPROVE_VAL + '\nThe value could be not written as a percentage value must be written. The value to approve will be 60% as default\n\nException: ' + e);
    appVal = 0.6;
  }
  var failedRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(0.6)
    .setBackground('#ea4335')
    .setRanges([totalRange])
    .build();
  var warningRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(0.6, appVal)
    .setBackground('#fbbc04')
    .setRanges([totalRange])
    .build();
  var approvedRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(appVal, 1)
    .setBackground('#34a853')
    .setRanges([totalRange])
    .build();
  var excellenceRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(1)
    .setBackground('#4285f4')
    .setRanges([totalRange])
    .build();

  var rules = ssTasks.getConditionalFormatRules();
  rules.push(failedRule);
  rules.push(warningRule);
  rules.push(approvedRule);
  rules.push(excellenceRule);
  ssTasks.setConditionalFormatRules(rules);
  //</editor-fold>

  // inserting member in _Data
  var rowIndex = getLastRowRange(memberRange)+1;
  ssData.getRange(rowIndex,DATA_MEMBER_COL).setValue(member);
  ssData.getRange(rowIndex,DATA_EMAIL_COL).setValue(email);

  // inserting member in History
  ssHistory.getRange(rowIndex-1, HISTORY_MEMBER_COL).setValue(member);
}

function deleteMember() {
  var ui = SpreadsheetApp.getUi(); // gets user interface
  var ss = SpreadsheetApp.getActive(); // assign active spreadsheet to variable
  var ssData = ss.getSheetByName(SHEET_DATA);
  var ssTasks = ss.getSheetByName(SHEET_TASKS);
  var member = ssTasks.getRange(TASKS_MEMBER).getValue();

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
      ssData.getRange(rowMember+1,1,1,2).deleteCells(SpreadsheetApp.Dimension.ROWS);

      // deletes data in Tasks
      ssTasks.getRange(TASKS_MEMBER).setValue('');
      ssTasks.getRange(10*rowMember+TASKS_MEMBER_INCREMENT,1,10,7).deleteCells(SpreadsheetApp.Dimension.ROWS);
    }
  }
  else
    ui.prompt('No member found', 'Make sure you choose the member within the options the dropdown list gives you', ui.ButtonSet.OK);
}
  //</editor-fold>

  //<editor-fold> Days
function addDay() {
  var ui = SpreadsheetApp.getUi(); // gets user interface
  var ssTasks = SpreadsheetApp.getActive().getSheetByName(SHEET_TASKS);

  var dayOption = ssTasks.getRange(TASKS_DAYS_DROPDOWN).getValue();
  var daysChosenRange = ssTasks.getRange(TASKS_DAYS_CHOSEN);
  var daysChosen = daysChosenRange.getValue();

  // validating
  if (dayOption == '') {
    ui.alert(':(', 'You haven\'t choose a day from the dropdown list', ui.ButtonSet.OK);
    return;
  }
  if (daysChosen.includes(dayOption))
    return;

  // entering data
  var today = new Date();

  if (dayOption == 'Everyday' || dayOption == 'Once every two days') {
    if (dayOption == 'Everyday')
      daysChosenRange.setValue(DAYS.join());
    else if (dayOption == 'Once every two days') {

      var days = '';
      for (var i = 0; i < 3; i++) {
        var index = (2*i + today.getDay()) % 6;
        days += (days == '') ? DAYS[index] : ',' + DAYS[index];
      }

      daysChosenRange.setValue(days);
    }
    else
      ui.alert('Wat?', 'This doesn\'t even make sense in the code, how did you do it tho?\nPlease tell me how you did it, I\'m impressed lol\n' + EMAIL, ui.ButtonSet.OK);
  }
  else {
    var days = (daysChosen == '') ? dayOption : daysChosen + "," + dayOption;
    ssTasks.getRange(TASKS_DAYS_CHOSEN).setValue(days);
  }
}

function removeDay() {
  var ui = SpreadsheetApp.getUi(); // gets user interface
  var ssTasks = SpreadsheetApp.getActive().getSheetByName(SHEET_TASKS);

  var dayOption = ssTasks.getRange(TASKS_DAYS_DROPDOWN).getValue();
  var daysChosenRange = ssTasks.getRange(TASKS_DAYS_CHOSEN);

  // validating
  if (dayOption == '') {
    ui.alert(':(', 'You haven\'t choose a day from the dropdown list', ui.ButtonSet.OK);
    return;
  }
  if (!daysChosenRange.getValue().includes(dayOption))
    return;


  daysChosenRange.setValue(daysChosenRange.getValue().replace(dayOption, '').replace(',,', ','));
  var daysChosen = daysChosenRange.getValue();
  if (daysChosen[0] == ',')
    daysChosenRange.setValue(daysChosen.substring(1, daysChosen.length));
  if (daysChosen[daysChosen.length-1] == ',')
    daysChosenRange.setValue(daysChosen.substring(0, daysChosen.length-1));
}

function clearDays() {
  SpreadsheetApp.getActive().getSheetByName(SHEET_TASKS).getRange(TASKS_DAYS_CHOSEN).setValue('');
}
  //</editor-fold>

  //<editor-fold> Collaborators
function addCollaborator() {
  var ui = SpreadsheetApp.getUi(); // gets user interface
  var ssTasks = SpreadsheetApp.getActive().getSheetByName(SHEET_TASKS);
  var ssData = SpreadsheetApp.getActive().getSheetByName(SHEET_DATA);

  // checking if it is not the same person
  var collaborator = ssTasks.getRange(TASKS_COLLABORATORS).getValue();
  if (collaborator == ssTasks.getRange(TASKS_MEMBER).getValue()) {
    ui.alert('🤨', 'You can\'t choose the same member as his collaborator', ui.ButtonSet.OK);
    ssTasks.getRange(TASKS_COLLABORATORS).setValue('');
    return;
  }

  // verifying if member has been chosen
  if (collaborator == '') {
    ui.alert(':(', 'You didn\'t choose a member from ' + TASKS_COLLABORATORS, ui.ButtonSet.OK);
    return;
  }

  var rowCollaborator = searchRowMember(collaborator);
  if (rowCollaborator == -1) {
    ui.alert(':(', 'An error has ocurred while trying to choose a member from ' + TASKS_COLLABORATORS + '\nMake sure data is not corruputed in ' + SHEET_DATA, ui.ButtonSet.OK);
    return;
  }

  var email = ssData.getRange(rowCollaborator, DATA_EMAIL_COL).getValue();
  var emailColRange = ssTasks.getRange(TASKS_EMAILS_COLLABORATORS);
  if (emailColRange.getValue().includes(email))
    ui.alert('🤨', 'You\'ve already chosen ' + collaborator, ui.ButtonSet.OK);
  else {
    if (emailColRange.getValue() == '')
      ssTasks.getRange(TASKS_EMAILS_COLLABORATORS).setValue(email);
    else {
      var emRangeVal = emailColRange.getValue();
      ssTasks.getRange(TASKS_EMAILS_COLLABORATORS).setValue(emRangeVal + ',' + email);
    }
  }

  ssTasks.getRange(TASKS_COLLABORATORS).setValue('');
}

function removeCollaborator() {
  var ui = SpreadsheetApp.getUi(); // gets user interface
  var ssTasks = SpreadsheetApp.getActive().getSheetByName(SHEET_TASKS);
  var ssData = SpreadsheetApp.getActive().getSheetByName(SHEET_DATA);

  // verifying if member has been chosen
  var collaborator = ssTasks.getRange(TASKS_COLLABORATORS).getValue();
  if (collaborator == '') {
    ui.alert(':(', 'You didn\'t choose a member from ' + TASKS_COLLABORATORS, ui.ButtonSet.OK);
    return;
  }

  var rowCollaborator = searchRowMember(collaborator);
  if (rowCollaborator == -1) {
    ui.alert(':(', 'An error has ocurred while trying to choose a member from ' + TASKS_COLLABORATORS + '\nMake sure data is not corruputed in ' + SHEET_DATA, ui.ButtonSet.OK);
    return;
  }

  var email = ssData.getRange(rowCollaborator, DATA_EMAIL_COL).getValue();
  var emailColRange = ssTasks.getRange(TASKS_EMAILS_COLLABORATORS);

  emailColRange.setValue(emailColRange.getValue().replace(email, '').replace(',,', ','));
  var emails = emailColRange.getValue();
  if (emails[0] == ',')
    emailColRange.setValue(emails.substring(1, emails.length));
  if (emails[emails.length-1] == ',')
    emailColRange.setValue(emails.substring(0, emails.length-1));

  ssTasks.getRange(TASKS_COLLABORATORS).setValue('');
}

function clearEmailsColl() {
  SpreadsheetApp.getActive().getSheetByName(SHEET_TASKS).getRange(TASKS_EMAILS_COLLABORATORS).setValue('');
}
  //</editor-fold>

function setDataEvent(event, tr, members, from, to, description, location, date, days, weeks) {
  var ui = SpreadsheetApp.getUi(); // gets user interface
  var ssData = SpreadsheetApp.getActive().getSheetByName(SHEET_DATA);

  var addedRows = getLastRowRange(ssData.getRange(DATA_EVENT + ":" + DATA_EVENT[0])) - 1;

  ssData.getRange(DATA_INITIAL_EVENT_ROW + addedRows, DATA_EVENT_COL).setValue(event);
  ssData.getRange(DATA_INITIAL_EVENT_ROW + addedRows, DATA_EVENT_COL+1).setValue(tr);
  ssData.getRange(DATA_INITIAL_EVENT_ROW + addedRows, DATA_EVENT_COL+2+1).setValue(members);
  ssData.getRange(DATA_INITIAL_EVENT_ROW + addedRows, DATA_EVENT_COL+3+1).setValue(from);
  ssData.getRange(DATA_INITIAL_EVENT_ROW + addedRows, DATA_EVENT_COL+4+1).setValue(to);
  ssData.getRange(DATA_INITIAL_EVENT_ROW + addedRows, DATA_EVENT_COL+5+1).setValue(description);
  ssData.getRange(DATA_INITIAL_EVENT_ROW + addedRows, DATA_EVENT_COL+6+1).setValue(location);
  ssData.getRange(DATA_INITIAL_EVENT_ROW + addedRows, DATA_EVENT_COL+7+1).setValue(date);
  ssData.getRange(DATA_INITIAL_EVENT_ROW + addedRows, DATA_EVENT_COL+8+1).setValue(days);
  ssData.getRange(DATA_INITIAL_EVENT_ROW + addedRows, DATA_EVENT_COL+9+1).setValue(weeks);

  // Converting parameters into Date objects
  var d = new Date(date);
  var s = new Date(from);
  var e = new Date(to);

  // Creating actual start and end dates
  var startDate = new Date(d.getFullYear(), d.getMonth(), d.getDate(), s.getHours(), s.getMinutes(), s.getSeconds(), s.getMilliseconds());
  var endDate = new Date(d.getFullYear(), d.getMonth(), d.getDate(), e.getHours(), e.getMinutes(), e.getSeconds(), e.getMilliseconds());

  var cell = addToCalendar(event, startDate, endDate).getA1Notation();
  ssData.getRange(DATA_INITIAL_EVENT_ROW + addedRows, DATA_EVENT_COL+2).setValue(cell);
}

function resetRoutineControls() {
  var ssTasks = SpreadsheetApp.getActive().getSheetByName(SHEET_TASKS);
  ssTasks.getRange(TASKS_ROUTINE).setValue('');
  ssTasks.getRange(TASKS_DAYS[0], TASKS_DAYS[1]+1).setValue('');
  ssTasks.getRange(TASKS_START).setValue('');
  ssTasks.getRange(TASKS_END).setValue('');
  ssTasks.getRange(TASKS_DURATION[0], TASKS_DURATION[1]).setValue('');
  ssTasks.getRange(TASKS_COLLABORATORS).setValue('');
  ssTasks.getRange(TASKS_DESCRIPTION).setValue('');
  ssTasks.getRange(TASKS_LOCATION).setValue('');
}

function addRoutine() {
  var ui = SpreadsheetApp.getUi(); // gets user interface
  var ssTasks = SpreadsheetApp.getActive().getSheetByName(SHEET_TASKS);

  // verifies if it is in routine mode
  if (ssTasks.getRange(TASKS_SWITCH).getValue().includes('routine')) {
    ui.alert('Ups! Wrong button', 'You are in task mode, you need to click ' + TASKS_SWITCH + ' button to change to routine mode if you want to create a routine', ui.ButtonSet.OK);
    return;
  }

  var routine = ssTasks.getRange(TASKS_ROUTINE).getValue();
  var member = ssTasks.getRange(TASKS_MEMBER).getValue();
  var days = ssTasks.getRange(TASKS_DAYS_CHOSEN).getValue();
  var duration = ssTasks.getRange(TASKS_DURATION[0], TASKS_DURATION[1]).getValue();
  var start = ssTasks.getRange(TASKS_START).getValue();
  var end = ssTasks.getRange(TASKS_END).getValue();
  var collaborators = ssTasks.getRange(TASKS_EMAILS_COLLABORATORS).getValue();
  var description = ssTasks.getRange(TASKS_DESCRIPTION).getValue();
  var location = ssTasks.getRange(TASKS_LOCATION).getValue();

  var isValid = true;
  if (routine == '') {
    ui.alert('Missing routine [' + TASKS_ROUTINE + ']');
    isValid = false;
  }
  if (member == '') {
    ui.alert('Missing member [' + TASKS_MEMBER + ']');
    isValid = false;
  }
  if (days == '') {
    ui.alert('Missing days [' + TASKS_DAYS_CHOSEN + ']');
    isValid = false;
  }
  if (duration == '') {
    ui.alert('Missing duration [' + TASKS_DURATION + ']');
    isValid = false;
  }
  if (start == '') {
    ui.alert('Missing start time [' + TASKS_START + ']');
    isValid = false;
  }
  if (end == '') {
    ui.alert('Missing end time [' + TASKS_END + ']');
    isValid = false;
  }
  if (new Date(start).getHours() > new Date(end).getHours()) {
    ui.alert('Start hour greater than end hour');
    isValid = false;
  }

  if (!isValid)
    return;

  // all data is valid, proceed to manage it
  var rowMember = searchRowMember(member)-1;
  if (rowMember == -1) {
    ui.prompt('😢 No member found', 'Make sure the member is in the sheet "' + SHEET_DATA + '"\n(Or that you have properly chosen within the dropdown list of ' + TASKS_MEMBER + ')', ui.ButtonSet.OK);
    return;
  }

  // creating routine
  var today = new Date();
  var numDay = today.getDay() - 1;
  var arrDays = days.split(',');
  var arrNumDays = [];

  // getting difference from today's date
  for (var i = 0; i < arrDays.length; i++) {
    var dif = getIndexOf(arrDays[i], DAYS) - numDay;
    arrNumDays.push( (dif <= 0) ?  7+dif : dif );
  }
  arrNumDays.sort();

  // getting next days from today on
  var nextDates = [];
  for (var i = 0; i < duration; i++) {
    for (var j = 0; j < arrNumDays.length; j++) {
      var nextDate = new Date(today);
      nextDate.setDate(today.getDate() + arrNumDays[j] + 7*i)
      nextDates.push(nextDate);
    }
  }

  // creating calendar events
  for (var i = 0; i < nextDates.length; i++) {
    addToGoogleCalendar(routine, nextDates[i], start, end, member, collaborators, description, location);
  }

  // resetting controls
  resetRoutineControls();
}

function resetTaskControls() {
  var ssTasks = SpreadsheetApp.getActive().getSheetByName(SHEET_TASKS);
  ssTasks.getRange(TASKS_TASK).setValue('');
  ssTasks.getRange(TASKS_DATE).setValue(DATE_CAPTION);
  ssTasks.getRange(TASKS_START).setValue('');
  ssTasks.getRange(TASKS_END).setValue('');
  ssTasks.getRange(TASKS_COLLABORATORS).setValue('');
  ssTasks.getRange(TASKS_DESCRIPTION).setValue('');
  ssTasks.getRange(TASKS_LOCATION).setValue('');
}

function addTask() {
  var ui = SpreadsheetApp.getUi(); // gets user interface
  var ssTasks = SpreadsheetApp.getActive().getSheetByName(SHEET_TASKS);
  var ssData = SpreadsheetApp.getActive().getSheetByName(SHEET_DATA);

  // verifies if it is in task mode
  if (ssTasks.getRange(TASKS_SWITCH).getValue().includes('task')) {
    ui.alert('Ups! Wrong button', 'You are in routine mode, you need to click ' + TASKS_SWITCH + ' button to change to task mode if you want to create a task', ui.ButtonSet.OK);
    return;
  }

  //<editor-fold> Retrieves and validates data
  var task = ssTasks.getRange(TASKS_TASK).getValue();
  var member = ssTasks.getRange(TASKS_MEMBER).getValue();
  var date = ssTasks.getRange(TASKS_DATE).getValue();
  var start = ssTasks.getRange(TASKS_START).getValue();
  var end = ssTasks.getRange(TASKS_END).getValue();
  var collaborators = ssTasks.getRange(TASKS_EMAILS_COLLABORATORS).getValue();
  var description = ssTasks.getRange(TASKS_DESCRIPTION).getValue();
  var location = ssTasks.getRange(TASKS_LOCATION).getValue();

  if (searchEvent(task, 'T') != -1) {
    ui.alert(':(', 'Another task exists witht that same name. Please choose another one', ui.ButtonSet.OK);
    return;
  }

  var isValid = true;
  if (task == '') {
    ui.alert('Missing task');
    isValid = false;
  }
  if (member == '') {
    ui.alert('Missing member [' + TASKS_MEMBER + ']');
    isValid = false;
  }
  if (date == '' || date == DATE_CAPTION) {
    ui.alert('Missing date');
    isValid = false;
  }
  if (start == '') {
    ui.alert('Missing start time [' + TASKS_START + ']');
    isValid = false;
  }
  if (end == '') {
    ui.alert('Missing end time [' + TASKS_END + ']');
    isValid = false;
  }
  if (new Date(start).getHours() > new Date(end).getHours()) {
    ui.alert('Start hour greater than end hour');
    isValid = false;
  }

  if (!isValid)
    return;
  //</editor-fold>

  // all data is valid, proceed to manage it
  var rowMember = searchRowMember(member)-1;
  if (rowMember == -1) {
    ui.prompt('😢 No member found', 'Make sure the member is in the sheet "' + SHEET_DATA + '"\n(Or that you have properly chosen within the dropdown list of ' + TASKS_MEMBER + ')', ui.ButtonSet.OK);
    return;
  }

  // creating task
  var row = 10*rowMember+TASKS_MEMBER_INCREMENT;
  var tasksRange = ssTasks.getRange(row,TASKS_VALUES_COL,9);
  var noRows = getLastRowRange(tasksRange);

  if (noRows == 9) {
    ui.prompt('parece ser que el cuerpo aieseco solo resiste 8 tareas');
    return;
  }

  // setting task
  ssTasks.getRange(row+noRows,TASKS_VALUES_COL).setValue(task);
  if (collaborators != '')
    setTask(collaborators.split(','), task);

  //<editor-fold> Giving value percentages
  setValues(row, noRows);
  //</editor-fold>

  // placing info in _Data
  var rowMember = searchRowMember(member);
  var email = ssData.getRange(rowMember, DATA_EMAIL_COL).getValue();
  setDataEvent(task, 'T', (collaborators == '') ? email : email + "," + collaborators, start, end, description, location, date, null, null);
  // Google Calendar
  addToGoogleCalendar(task, date, start, end, member, collaborators, description, location);

  // resetting controls
  resetTaskControls();
}

function switchTaskRoutine() {
  var ui = SpreadsheetApp.getUi(); // gets user interface
  var ssTasks = SpreadsheetApp.getActive().getSheetByName(SHEET_TASKS);

  var switchCaption = ssTasks.getRange(TASKS_SWITCH)
  var disable = SpreadsheetApp.newDataValidation().requireTextEqualTo('null').setAllowInvalid(false).setHelpText('You cannot edit this cell').build();
  // switch to routine
  if (switchCaption.getValue().includes('routine')) {
    // tasks controls
    ssTasks.getRange(TASKS_TASK_ROW, TASKS_TITLES_COL, 1, 2).setBackground('#dedede').setDataValidation(disable);
    ssTasks.getRange(TASKS_DATE_ROW, TASKS_TITLES_COL, 1, 2).setBackground('#dedede');
    ssTasks.getRange(TASKS_ADD_TASK_BUTTON).setBackground('#dedede');
    // routine controls
    ssTasks.getRange(TASKS_ROUTINE_ROW, TASKS_TITLES_COL, 1, 2).setBackground('white').setDataValidation(null);
    ssTasks.getRange(TASKS_DAYS[0], TASKS_DAYS[1], 1, 4).setBackground('white');
    ssTasks.getRange(TASKS_DURATION[0], TASKS_DURATION[1]-1, 1, 3).setBackground('white');
    ssTasks.getRange(TASKS_ADD_ROUTINE_BUTTON).setBackground('white');

    switchCaption.setValue('Switch to task');
  }
  // switch to tasks
  else {
    // tasks controls
    ssTasks.getRange(TASKS_TASK_ROW, TASKS_TITLES_COL, 1, 2).setBackground('white').setDataValidation(null);
    ssTasks.getRange(TASKS_DATE_ROW, TASKS_TITLES_COL, 1, 2).setBackground('white');
    ssTasks.getRange(TASKS_ADD_TASK_BUTTON).setBackground('white');
    // routine controls
    ssTasks.getRange(TASKS_ROUTINE_ROW, TASKS_TITLES_COL, 1, 2).setBackground('#dedede').setDataValidation(disable);
    ssTasks.getRange(TASKS_DAYS[0], TASKS_DAYS[1], 1, 4).setBackground('#dedede');
    ssTasks.getRange(TASKS_DURATION[0], TASKS_DURATION[1]-1, 1, 3).setBackground('#dedede');
    ssTasks.getRange(TASKS_ADD_ROUTINE_BUTTON).setBackground('#dedede');

    switchCaption.setValue('Switch to routine');
  }
}

function clearCompleted() {
  var ui = SpreadsheetApp.getUi(); // gets user interface
  var ssTasks = SpreadsheetApp.getActive().getSheetByName(SHEET_TASKS);

  for (var rowMember = 1; rowMember <= getMembershipNumber(); rowMember++) {
    var row = 10*rowMember+TASKS_MEMBER_INCREMENT;
    var noRows = getLastRowRange(ssTasks.getRange(row,TASKS_VALUES_COL,NUM_TASKS+1));
    var tasksRange = ssTasks.getRange(row,TASKS_VALUES_COL,NUM_TASKS+1,3);

    var deletedCells = 0;
    var index = 0;

    var rowsToDelete = []

    // retrieving row's indexes to delete according to the checbox true
    for (var i = 1; i < noRows; i++)
      if (tasksRange.getValues()[i][2] == true)
        rowsToDelete.push(row+i);

    // deleting cells
    for (var i = rowsToDelete.length-1; i >= 0; i--) {
      ssTasks.getRange(rowsToDelete[i], TASKS_VALUES_COL, 1, 2).deleteCells(SpreadsheetApp.Dimension.ROWS);
      ssTasks.getRange(rowsToDelete[i], getColumnNumber(TASKS_CHECKBOX_COLUMN)).setValue(false);
    }

    // inserting deleted cells
    for (var i = 0; i < rowsToDelete.length; i++)
      ssTasks.getRange(row+9-rowsToDelete.length, TASKS_VALUES_COL, 1, 2).insertCells(SpreadsheetApp.Dimension.ROWS);

    // format values
    ssTasks.getRange(row+1, getColumnNumber(TASKS_VALUE_COLUMN), NUM_TASKS).setNumberFormat('0.00%').setFontWeight("normal");

    // achievements
    for (var i = 1; i <= NUM_TASKS; i++)
      ssTasks.getRange(row+i,getColumnNumber(TASKS_ACHIEVEMENT_COLUMN)).setFormula('=IF(' + TASKS_CHECKBOX_COLUMN + (row+i).toString() + '=TRUE, ' + TASKS_VALUE_COLUMN + (row+i).toString() + ', ' + TASKS_H_M_COLUMN + (row+i).toString() + ')');
  }
}

// MISSING
function clearAll() {
  // resetCompleted();
  var ssCalendar = SpreadsheetApp.getActive().getSheetByName(SHEET_CALENDAR);
  var ssData = SpreadsheetApp.getActive().getSheetByName(SHEET_DATA);

  ssCalendar.getRange(CALENDAR_INITIAL_DATE + ":" + CALENDAR_FINAL_DATE).deleteCells(SpreadsheetApp.Dimension.ROWS);
  ssData.getRange(DATA_EVENT[0] + DATA_INITIAL_EVENT_ROW.toString() + ":" + DATA_WEEKS[0]).deleteCells(SpreadsheetApp.Dimension.ROWS);
}

function hardReset() {
  var ui = SpreadsheetApp.getUi(); // gets user interface
  var ssTasks = SpreadsheetApp.getActive().getSheetByName(SHEET_TASKS);
  var ssData = SpreadsheetApp.getActive().getSheetByName(SHEET_DATA);
  var ssCalendar = SpreadsheetApp.getActive().getSheetByName(SHEET_CALENDAR);
  var ssHistory = SpreadsheetApp.getActive().getSheetByName(SHEET_HISTORY)

  var confirmation = ui.alert('Wowwowowowwoah slow down', 'You\'re about to end this tracker\'s whole carrer\nAre you sure you want to delete all the information?\n\n(Click "yes" to proceed)', ui.ButtonSet.YES_NO);

  if (confirmation != ui.Button.YES)
    return;

  ssData.getRange(DATA_MEMBER + ":" + DATA_HISTORY[0]).deleteCells(SpreadsheetApp.Dimension.ROWS);
  ssData.getRange(DATA_EVENT[0] + DATA_INITIAL_EVENT_ROW.toString() + ":" + DATA_WEEKS[0]).deleteCells(SpreadsheetApp.Dimension.ROWS);
  ssData.getRange(DATA_CAL_ID).setValue('');
  ssTasks.getRange(TASKS_NON_FIX_VALUES + ":" + TASKS_TOTAL_COLUMN).deleteCells(SpreadsheetApp.Dimension.ROWS);
  resetTaskControls();
  resetRoutineControls();
  clearDays();
  clearEmailsColl();
  ssTasks.getRange(TASKS_MEMBER).setValue('');
  ssCalendar.getRange(CALENDAR_INITIAL_DATE + ":" + CALENDAR_FINAL_DATE).deleteCells(SpreadsheetApp.Dimension.ROWS);
  ssHistory.getDataRange().deleteCells(SpreadsheetApp.Dimension.ROWS);
}

//</editor-fold>
