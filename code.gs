// sheets names
var SHEET_DATA = '_Data';
var SHEET_TASKS = 'Tasks';
var SHEET_CALENDAR = 'Calendar';
// SHEET_DATA variables
var DATA_MEMBER_COL = 1;
var DATA_EMAIL_COL = 2;
var DATA_CAL_ID = 'F3';
// SHEET_TASKS variables
var TASKS_MEMBER_INCREMENT = 5;
var TASKS_TITLES_COL = 1;
var TASKS_VALUES_COL = 2;
var TASKS_TASK = 'B1';
var TASKS_TASK_ROW = 1;
var TASKS_ROUTINE = 'B2';
var TASKS_ROUTINE_ROW = 2;
var TASKS_MEMBER = 'B3';
var TASKS_DATE = 'B4';
var TASKS_DATE_ROW = 4;
var TASKS_START = 'B5';
var TASKS_END = 'B6';
var TASKS_COLLABORATORS = 'B7';
var TASKS_EMAILS_COLLABORATORS = 'E7';
var TASKS_DESCRIPTION = 'B8';
var TASKS_LOCATION = 'B9';
var TASKS_SWITCH = 'C1';
var TASKS_DAYS = [4,3];
var TASKS_DAYS_DROPDOWN = 'D4';
var TASKS_DAYS_CHOSEN = 'F4';
var TASKS_DURATION = [5,4];
var TASKS_ADD_TASK_BUTTON = 'A11';
var TASKS_ADD_ROUTINE_BUTTON = 'A12';
var DATE_CAPTION = 'Double click to pop calendar up';
var DAYS = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
// calendar options
var SEND_INVITES = true;
// coder info
var EMAIL = 'eduardo.mendozamartinez@aiesec.net';

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

function getIndexOf(element, collection) {
  for (var i = 0; i < collection.length; i++)
    if (collection[i] == element)
      return i;

  return -1;
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
      ssData.getRange(rowMember+1,1,1,3).deleteCells(SpreadsheetApp.Dimension.ROWS);

      // deletes data in Tasks
      ssTasks.getRange(TASKS_MEMBER).setValue('');
      ssTasks.getRange(10*rowMember+TASKS_MEMBER_INCREMENT,1,10,3).deleteCells(SpreadsheetApp.Dimension.ROWS);
    }
  }
  else
    ui.prompt('No member found', 'Make sure you choose the member within the options the dropdown list gives you', ui.ButtonSet.OK);
}

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
}

function clearEmailsColl() {
  SpreadsheetApp.getActive().getSheetByName(SHEET_TASKS).getRange(TASKS_EMAILS_COLLABORATORS).setValue('');
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
    addToCalendar(routine, nextDates[i], start, end, member, collaborators, description, location);
  }

  // resetting controls
  ssTasks.getRange(TASKS_ROUTINE).setValue('');
  ssTasks.getRange(TASKS_DAYS[0], TASKS_DAYS[1]+1).setValue('');
  ssTasks.getRange(TASKS_START).setValue('');
  ssTasks.getRange(TASKS_END).setValue('');
  ssTasks.getRange(TASKS_DURATION[0], TASKS_DURATION[1]).setValue('');
  ssTasks.getRange(TASKS_COLLABORATORS).setValue('');
  ssTasks.getRange(TASKS_DESCRIPTION).setValue('');
  ssTasks.getRange(TASKS_LOCATION).setValue('');
}

function addTask() {
  var ui = SpreadsheetApp.getUi(); // gets user interface
  var ssTasks = SpreadsheetApp.getActive().getSheetByName(SHEET_TASKS);

  // verifies if it is in task mode
  if (ssTasks.getRange(TASKS_SWITCH).getValue().includes('task')) {
    ui.alert('Ups! Wrong button', 'You are in routine mode, you need to click ' + TASKS_SWITCH + ' button to change to task mode if you want to create a task', ui.ButtonSet.OK);
    return;
  }

  var task = ssTasks.getRange(TASKS_TASK).getValue();
  var member = ssTasks.getRange(TASKS_MEMBER).getValue();
  var date = ssTasks.getRange(TASKS_DATE).getValue();
  var start = ssTasks.getRange(TASKS_START).getValue();
  var end = ssTasks.getRange(TASKS_END).getValue();
  var collaborators = ssTasks.getRange(TASKS_EMAILS_COLLABORATORS).getValue();
  var description = ssTasks.getRange(TASKS_DESCRIPTION).getValue();
  var location = ssTasks.getRange(TASKS_LOCATION).getValue();

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

  // setting task
  ssTasks.getRange(row+noRows,TASKS_VALUES_COL).setValue(task);

  // Google Calendar
  addToCalendar(task, date, start, end, member, collaborators, description, location);

  // resetting controls
  ssTasks.getRange(TASKS_TASK).setValue('');
  ssTasks.getRange(TASKS_DATE).setValue(DATE_CAPTION);
  ssTasks.getRange(TASKS_START).setValue('');
  ssTasks.getRange(TASKS_END).setValue('');
  ssTasks.getRange(TASKS_COLLABORATORS).setValue('');
  ssTasks.getRange(TASKS_DESCRIPTION).setValue('');
  ssTasks.getRange(TASKS_LOCATION).setValue('');
}

function switchTaskRoutine() {
  var ui = SpreadsheetApp.getUi(); // gets user interface
  var ssTasks = SpreadsheetApp.getActive().getSheetByName(SHEET_TASKS);

  var switchCaption = ssTasks.getRange(TASKS_SWITCH)
  var disable = SpreadsheetApp.newDataValidation().requireTextEqualTo('null').setAllowInvalid(false).setHelpText('Do not edit in print sheet').build();
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
