/**
 * Removes an event from SHEET_DATA through its name and the type of event
 *
 * @param  {string} event The name of the task or routine to delete
 * @param  {string} tr 'T' for task, 'R' for routine
 */
function removeEvent(event, tr) {
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

function weeklyCut() {
  if (getMembershipNumber() == 0)
    return;

  // confirmation
  if (UI.alert("?", "Are you sure you want to make the weekly cut?\n\nThis will reset all the finished events and will let those who were not completed\n(Click 'YES' to continue)", UI.ButtonSet.YES_NO) != UI.Button.YES)
    return;

  // setting history up
  var noWeekColumn = getLastColumn(SS_HISTORY.getDataRange());
  SS_HISTORY.getRange(HISTORY_WEEKS_ROW, noWeekColumn + 1).setValue('w' + noWeekColumn.toString());

  for (var rowMember = 1; rowMember <= getMembershipNumber(); rowMember++) {
    var row = 10 * rowMember + TASKS_MEMBER_INCREMENT;
    var noRows = getLastRow(SS_TASKS.getRange(row, TASKS_VALUES_COL, NUM_TASKS + 1));
    var tasksRange = SS_TASKS.getRange(row, TASKS_VALUES_COL, NUM_TASKS + 1, 3);

    // adding to history
    SS_HISTORY.getRange(rowMember + 1, noWeekColumn + 1).setValue(SS_TASKS.getRange(row + 1, getColumnNumber(TASKS_TOTAL_COLUMN)).getValue());

    var rowsToDelete = []
    // retrieving row's indexes to delete according to the checbox
    for (var i = 1; i < noRows; i++)
      if (tasksRange.getValues()[i][2] == true) {
        rowsToDelete.push(row + i);

        // deleting value assigned to task
        SS_TASKS.getRange(row + i, getColumnNumber(TASKS_VALUE_COLUMN)).setValue('');
        // deleting from _Data
        removeEvent(SS_TASKS.getRange(row + i, TASKS_VALUES_COL).getValue(), 'T');
      }

    // deleting cells
    for (var i = rowsToDelete.length - 1; i >= 0; i--) {
      SS_TASKS.getRange(rowsToDelete[i], TASKS_VALUES_COL, 1, 2).deleteCells(SpreadsheetApp.Dimension.ROWS);
      SS_TASKS.getRange(rowsToDelete[i], getColumnNumber(TASKS_CHECKBOX_COLUMN)).setValue(false);
    }

    // inserting deleted cells
    for (var i = 0; i < rowsToDelete.length; i++)
      SS_TASKS.getRange(row + 9 - rowsToDelete.length, TASKS_VALUES_COL, 1, 2).insertCells(SpreadsheetApp.Dimension.ROWS);

    if (rowsToDelete.length > 0) {
      // format values
      SS_TASKS.getRange(row + 1, getColumnNumber(TASKS_VALUE_COLUMN), NUM_TASKS).setNumberFormat('0.00%').setFontWeight("normal");
      // achievements
      setAchievementRange(row);
    }

    // recalculating Value of tasks
    for (var i = 1; i <= NUM_TASKS; i++) {
      hmRange = SS_TASKS.getRange(row + i, getColumnNumber(TASKS_H_M_COLUMN));
      if (SS_TASKS.getRange(row + i, TASKS_CHECKBOX_COL).getValue() == false && hmRange.getValue() != 0) {
        valueRange = SS_TASKS.getRange(row + i, getColumnNumber(TASKS_VALUE_COLUMN));
        valueRange.setValue(valueRange.getValue() - SS_TASKS.getRange(row + i, getColumnNumber(TASKS_ACHIEVEMENT_COLUMN)).getValue());
      }
    }

    // resetting 'if not, how much' values
    SS_TASKS.getRange(row + 1, getColumnNumber(TASKS_H_M_COLUMN), NUM_TASKS).setValue(0);
  }

  // Creating conditional formatting for achieved values in history
  var achievedRange = SS_HISTORY.getRange(2, noWeekColumn + 1, rowMember - 1);
  achievedRange.setNumberFormat('0.00%');
  achievedRange.setHorizontalAlignment("center");
  achievedRange.setVerticalAlignment("middle");
  achievedRange.setFontWeight("bold");

  var appVal = getValueToApprove();
  var rules = SS_HISTORY.getConditionalFormatRules();
  rules.push(createRule(false, 0.6, COLOR_FAIL, achievedRange));
  rules.push(createRuleInInterval(0.6, appVal, COLOR_WARNING, achievedRange));
  rules.push(createRuleInInterval(appVal, 1, COLOR_APPROVED, achievedRange));
  rules.push(createRule(true, 1, COLOR_EXCELLENCE, achievedRange));
  SS_HISTORY.setConditionalFormatRules(rules);
}