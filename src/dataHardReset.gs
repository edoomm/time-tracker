function hardReset() {
  var confirmation = UI.alert('Wowwowowowwoah slow down', 'You\'re about to end this tracker\'s whole carrer\nAre you sure you want to delete all the information?\n\n(Click "yes" to proceed)', UI.ButtonSet.YES_NO);

  if (confirmation != UI.Button.YES)
    return;

  SS_DATA.getRange(DATA_MEMBER + ":" + DATA_HISTORY[0]).deleteCells(SpreadsheetApp.Dimension.ROWS);
  SS_DATA.getRange(DATA_EVENT[0] + DATA_INITIAL_EVENT_ROW.toString() + ":" + DATA_WEEKS[0]).deleteCells(SpreadsheetApp.Dimension.ROWS);
  SS_DATA.getRange(DATA_CAL_ID).setValue('');
  SS_TASKS.getRange(TASKS_NON_FIX_VALUES + ":" + TASKS_TOTAL_COLUMN).deleteCells(SpreadsheetApp.Dimension.ROWS);
  resetTaskControls();
  resetRoutineControls();
  clearDays();
  clearEmailsColl();
  SS_TASKS.getRange(TASKS_MEMBER).setValue('');
  SS_CALENDAR.getRange(CALENDAR_INITIAL_DATE + ":" + CALENDAR_FINAL_DATE).deleteCells(SpreadsheetApp.Dimension.ROWS);
  SS_HISTORY.getDataRange().deleteCells(SpreadsheetApp.Dimension.ROWS);
}