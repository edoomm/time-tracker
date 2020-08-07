/**
 * Deletes a task from SHEET_DATA through its name and the type of event
 *
 * @param  {string} event The name of the task or routine to delete
 * @param  {string} tr 'T' for task, 'R' for routine
 */
function deleteEvent() {
  // retrieving and validating
  var event = SS_TASKS.getRange(TASKS_EVENT_CHOSEN).getValue();
  if (isEmptyValue(event, '): you did not choose an event from [' + TASKS_EVENT_CHOSEN + ']'))
    return;

  var rowEvent = searchEvent(event, 'T');
  // task
  if (rowEvent != -1) {
    UI.alert(rowEvent);

    // TODO: shift up task in SHEET_TASKS for each member that has the task
  }
  // event
  else {
    // validating
    rowEvent = searchEvent(event, 'R');
    if (rowEvent == -1) {
      UI.alert('An error has ocurred):', 'The routine "' + event + '" could not be found in "' + SHEET_DATA + '"\nYou will have to delete it manually by selecting the cells of "' + event + '" from column [' + DATA_EVENT[0] + '] to column [' + DATA_WEEKS[0] + '] and shifting them up', UI.ButtonSet.OK);
      return;
    }

    // TODO: Delete from SHEET_DATA

    var calendarCells = SS_DATA.getRange(rowEvent, getColumnNumber(DATA_CALENDAR_CELL)).getValue().split(',');
    UI.alert(calendarCells);

    // TODO: DELETE from SHEET_CALENDAR with removeFromCalendar()
  }
}
