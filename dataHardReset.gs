
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
  SS_HISTORY.getDataRange().deleteCells(SpreadsheetApp.Dimension.ROWS);
}
