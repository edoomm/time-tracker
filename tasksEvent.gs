function deleteEvent() {
  // retrieving and validating
  var event = SS_TASKS.getRange(TASKS_EVENT_CHOSEN).getValue();
  if (isEmptyValue(event, '): you did not choose an event from [' + TASKS_EVENT_CHOSEN + ']')) // THIS
    return;

  var rowEvent = searchEvent(event, 'T');
  // task
  if (rowEvent != -1) {
    var calendarCell = SS_DATA.getRange(rowEvent, getColumnNumber(DATA_CALENDAR_CELL)).getValue();

    // deletes from SHEET_CALENDAR the event that is in SHEET_CALENDAR
    removeFromCalendar(calendarCell, event);

    var emails = SS_DATA.getRange(DATA_MEMBERS_CELL[0] + rowEvent.toString()).getValue().toString().split(',');
    for (var email of emails)
      deleteTask(event, email)
  }
  // event
  else {
    // validating
    rowEvent = searchEvent(event, 'R');
    if (rowEvent == -1) {
      UI.alert('An error has ocurred):', 'The routine "' + event + '" could not be found in "' + SHEET_DATA + '"\nYou will have to delete it manually by selecting the cells of "' + event + '" from column [' + DATA_EVENT[0] + '] to column [' + DATA_WEEKS[0] + '] and shifting them up', UI.ButtonSet.OK);
      return;
    }

    var calendarCells = SS_DATA.getRange(rowEvent, getColumnNumber(DATA_CALENDAR_CELL)).getValue().split(',');

    // deletes from SHEET_CALENDAR each event that is in SHEET_CALENDAR
    for (var cell of calendarCells)
      removeFromCalendar(cell, event);
  }
  // deletes from SHEET_DATA
  SS_DATA.getRange(DATA_EVENT[0] + rowEvent.toString() + ":" + DATA_WEEKS[0] + rowEvent.toString()).deleteCells(SpreadsheetApp.Dimension.ROWS);
  // resets in SHEET_TASKS
  SS_TASKS.getRange(TASKS_EVENT_CHOSEN).setValue('');
}

/**
 * Deletes a task from SHEET_DATA through its name and the type of event
 *
 * @param  {string} event The name of the task or routine to delete
 * @param  {string} tr 'T' for task, 'R' for routine
 */
function deleteEvent(event, tr) {
  var rowEvent = searchEvent(event, tr);
  if (rowEvent == -1)
    return;

  // deleting in SHEET_CALENDAR
  var scheduleRange = SS_CALENDAR.getRange(SS_DATA.getRange(rowEvent, getColumnNumber(DATA_CALENDAR_CELL[0])).getValue());
  if (!scheduleRange.getValue().includes(';'))
    scheduleRange.setValue('');

  // deleting in SHEET_DATA
  SS_DATA.getRange(rowEvent, getColumnNumber(DATA_EVENT[0]), 1, getColumnNumber(DATA_WEEKS[0])).deleteCells(SpreadsheetApp.Dimension.ROWS);
}