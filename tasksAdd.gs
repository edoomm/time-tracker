function addRoutine() {
  // <editor-fold> Retrieving and validating data
  // verifies if it is in routine mode
  if (SS_TASKS.getRange(TASKS_SWITCH).getValue().includes('routine')) {
    UI.alert('Ups! Wrong button', 'You are in task mode, you need to click ' + TASKS_SWITCH + ' button to change to routine mode if you want to create a routine', UI.ButtonSet.OK);
    return;
  }

  // collaborator not added notification
  if (!isEmptyValue(SS_TASKS.getRange(TASKS_COLLABORATORS).getValue(), null))
    if (UI.alert('😯', 'It seems that you wanted to add a collaborator but you didn\'t click the "Add" button next to the cell in which you chose the collaborator\'s name\n\nIf you click "Ok", that collaborator will be ignored, otherwise you can click "Cancel" and go click the "Add" button to add that collaborator', UI.ButtonSet.OK_CANCEL) != UI.Button.OK)
      return;

  // days not added notification
  if (!isEmptyValue(SS_TASKS.getRange(TASKS_DAYS_DROPDOWN).getValue(), null))
    if (UI.alert('😲', 'Did you wanted to add days to the routine?\nYou didn\'t click the "Add" button next to the dropdown menu, the days you choose will be displayed in cell [' + TASKS_DAYS_CHOSEN + ']', UI.ButtonSet.OK_CANCEL) != UI.Button.OK)
      return;

  var routine = SS_TASKS.getRange(TASKS_ROUTINE).getValue();

  if (searchEvent(routine, 'R') != -1) {
    UI.alert(':(', 'It already exists a routine with the name "' + routine + '", you have to choose a different one, in order to create the routine', UI.ButtonSet.OK);
    return;
  }

  var member = SS_TASKS.getRange(TASKS_MEMBER).getValue();
  var days = SS_TASKS.getRange(TASKS_DAYS_CHOSEN).getValue();
  var duration = SS_TASKS.getRange(TASKS_DURATION[0], TASKS_DURATION[1]).getValue();
  var start = SS_TASKS.getRange(TASKS_START).getValue();
  var end = SS_TASKS.getRange(TASKS_END).getValue();
  var collaborators = SS_TASKS.getRange(TASKS_EMAILS_COLLABORATORS).getValue();
  var description = SS_TASKS.getRange(TASKS_DESCRIPTION).getValue();
  var location = SS_TASKS.getRange(TASKS_LOCATION).getValue();

  var isValid = true;
  if (routine == '') {
    UI.alert('Missing routine [' + TASKS_ROUTINE + ']');
    isValid = false;
  }
  if (member == '') {
    UI.alert('Missing member [' + TASKS_MEMBER + ']');
    isValid = false;
  }
  if (days == '') {
    UI.alert('Missing days [' + TASKS_DAYS_CHOSEN + ']');
    isValid = false;
  }
  if (duration == '') {
    UI.alert('Missing duration [' + TASKS_DURATION + ']');
    isValid = false;
  }
  if (start == '') {
    UI.alert('Missing start time [' + TASKS_START + ']');
    isValid = false;
  }
  if (end == '') {
    UI.alert('Missing end time [' + TASKS_END + ']');
    isValid = false;
  }
  if (new Date(start).getHours() > new Date(end).getHours()) {
    UI.alert('Start hour greater than end hour');
    isValid = false;
  }

  if (!isValid)
    return;

  if (routine.includes(';')) {
    UI.alert(':(', 'The name of the routine "' + routine + '" includes an illegal character ";"\n\nPlease use a comma (,) or a period (.) instead', UI.ButtonSet.OK);
    return;
  }
  // </editor-fold>

  // all data is valid, proceed to manage it
  var rowMember = searchRowMember(member) - 1;
  if (rowMember == -1) {
    UI.prompt('😢 No member found', 'Make sure the member is in the sheet "' + SHEET_DATA + '"\n(Or that you have properly chosen within the dropdown list of ' + TASKS_MEMBER + ')', UI.ButtonSet.OK);
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
    arrNumDays.push((dif <= 0) ? 7 + dif : dif);
  }
  arrNumDays.sort();

  // getting next dates from today on
  var nextDates = [];
  for (var i = 0; i < duration; i++) {
    for (var j = 0; j < arrNumDays.length; j++) {
      var nextDate = new Date(today);
      nextDate.setDate(today.getDate() + arrNumDays[j] + 7 * i)
      nextDates.push(nextDate);
    }
  }

  var rowMember = searchRowMember(member);
  var email = SS_DATA.getRange(rowMember, DATA_EMAIL_COL).getValue();

  // creating calendar events
  for (var i = 0; i < nextDates.length; i++) {
    addToGoogleCalendar(routine, nextDates[i], start, end, member, collaborators, description, location);
  }

  setDataEvent(routine, 'R', (collaborators == '') ? email : email + "," + collaborators, start, end, description, location, nextDates, days, duration);

  // resetting controls
  resetRoutineControls();
}

function addTask() {
  // verifies if it is in task mode
  if (SS_TASKS.getRange(TASKS_SWITCH).getValue().includes('task')) {
    UI.alert('Ups! Wrong button', 'You are in routine mode, you need to click ' + TASKS_SWITCH + ' button to change to task mode if you want to create a task', UI.ButtonSet.OK);
    return;
  }

  // collaborator not added notification
  if (!isEmptyValue(SS_TASKS.getRange(TASKS_COLLABORATORS).getValue(), null))
    if (UI.alert('😯', 'It seems that you wanted to add a collaborator but you didn\'t click the "Add" button next to the cell in which you chose the collaborator\'s name\n\nIf you click "Ok", that collaborator will be ignored, otherwise you can click "Cancel" and go click the "Add" button to add that collaborator', UI.ButtonSet.OK_CANCEL) != UI.Button.OK)
      return;

  //<editor-fold> Retrieves and validates data
  var task = SS_TASKS.getRange(TASKS_TASK).getValue();
  var member = SS_TASKS.getRange(TASKS_MEMBER).getValue();
  var date = SS_TASKS.getRange(TASKS_DATE).getValue();
  var start = SS_TASKS.getRange(TASKS_START).getValue();
  var end = SS_TASKS.getRange(TASKS_END).getValue();
  var collaborators = SS_TASKS.getRange(TASKS_EMAILS_COLLABORATORS).getValue();
  var description = SS_TASKS.getRange(TASKS_DESCRIPTION).getValue();
  var location = SS_TASKS.getRange(TASKS_LOCATION).getValue();

  if (searchEvent(task, 'T') != -1) {
    UI.alert(':(', 'Another task exists with that same name. Please choose another one', UI.ButtonSet.OK);
    return;
  }

  var isValid = true;
  if (task == '') {
    UI.alert('Missing task');
    isValid = false;
  }
  if (member == '') {
    UI.alert('Missing member [' + TASKS_MEMBER + ']');
    isValid = false;
  }
  if (date == '' || date == DATE_CAPTION) {
    UI.alert('Missing date');
    isValid = false;
  }
  if (start == '') {
    UI.alert('Missing start time [' + TASKS_START + ']');
    isValid = false;
  }
  if (end == '') {
    UI.alert('Missing end time [' + TASKS_END + ']');
    isValid = false;
  }
  if (new Date(start).getHours() > new Date(end).getHours()) {
    UI.alert('Start hour greater than end hour');
    isValid = false;
  }

  if (!isValid)
    return;

  if (task.includes(';')) {
    UI.alert(':(', 'The name of the task "' + task + '" includes an illegal character ";"\n\nPlease use a comma (,) or a period (.) instead', UI.ButtonSet.OK);
    return;
  }
  //</editor-fold>

  // all data is valid, proceed to manage it
  var rowMember = searchRowMember(member) - 1;
  if (rowMember == -1) {
    UI.prompt('😢 No member found', 'Make sure the member is in the sheet "' + SHEET_DATA + '"\n(Or that you have properly chosen within the dropdown list of ' + TASKS_MEMBER + ')', UI.ButtonSet.OK);
    return;
  }

  // creating task
  var row = 10 * rowMember + TASKS_MEMBER_INCREMENT;
  var tasksRange = SS_TASKS.getRange(row, TASKS_VALUES_COL, 9);
  var noRows = getLastRow(tasksRange);

  if (noRows == 9) {
    UI.prompt('parece ser que el cuerpo aieseco solo resiste 8 tareas');
    return;
  }

  // setting task
  SS_TASKS.getRange(row + noRows, TASKS_VALUES_COL).setValue(task);
  if (collaborators != '')
    setTask(collaborators.split(','), task);

  // Giving value percentages
  setValues(row, noRows);

  // placing info in _Data
  var rowMember = searchRowMember(member);
  var email = SS_DATA.getRange(rowMember, DATA_EMAIL_COL).getValue();
  setDataEvent(task, 'T', (collaborators == '') ? email : email + "," + collaborators, start, end, description, location, date, null, null);
  // Google Calendar
  addToGoogleCalendar(task, date, start, end, member, collaborators, description, location);

  // resetting controls
  resetTaskControls();
}

function switchTaskRoutine() {
  var switchCaption = SS_TASKS.getRange(TASKS_SWITCH);
  var disable = SpreadsheetApp.newDataValidation().requireTextEqualTo('null').setAllowInvalid(false).setHelpText('You cannot edit this cell').build();
  // switch to routine
  if (switchCaption.getValue().includes('routine')) {
    // tasks controls
    SS_TASKS.getRange(TASKS_TASK_ROW, TASKS_TITLES_COL, 1, 2).setBackground('#dedede').setDataValidation(disable);
    SS_TASKS.getRange(TASKS_DATE_ROW, TASKS_TITLES_COL, 1, 2).setBackground('#dedede');
    SS_TASKS.getRange(TASKS_ADD_TASK_BUTTON).setBackground('#dedede');
    // routine controls
    SS_TASKS.getRange(TASKS_ROUTINE_ROW, TASKS_TITLES_COL, 1, 2).setBackground('white').setDataValidation(null);
    SS_TASKS.getRange(TASKS_DAYS[0], TASKS_DAYS[1], 1, 4).setBackground('white');
    SS_TASKS.getRange(TASKS_DURATION[0], TASKS_DURATION[1] - 1, 1, 3).setBackground('white');
    SS_TASKS.getRange(TASKS_ADD_ROUTINE_BUTTON).setBackground('white');

    switchCaption.setValue('Switch to task');
  }
  // switch to tasks
  else {
    // tasks controls
    SS_TASKS.getRange(TASKS_TASK_ROW, TASKS_TITLES_COL, 1, 2).setBackground('white').setDataValidation(null);
    SS_TASKS.getRange(TASKS_DATE_ROW, TASKS_TITLES_COL, 1, 2).setBackground('white');
    SS_TASKS.getRange(TASKS_ADD_TASK_BUTTON).setBackground('white');
    // routine controls
    SS_TASKS.getRange(TASKS_ROUTINE_ROW, TASKS_TITLES_COL, 1, 2).setBackground('#dedede').setDataValidation(disable);
    SS_TASKS.getRange(TASKS_DAYS[0], TASKS_DAYS[1], 1, 4).setBackground('#dedede');
    SS_TASKS.getRange(TASKS_DURATION[0], TASKS_DURATION[1] - 1, 1, 3).setBackground('#dedede');
    SS_TASKS.getRange(TASKS_ADD_ROUTINE_BUTTON).setBackground('#dedede');

    switchCaption.setValue('Switch to routine');
  }
}