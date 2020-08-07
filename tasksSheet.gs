//Functions that changing the tasks sheet

function setValues(row, noRows) {
  var ssTasks = SpreadsheetApp.getActive().getSheetByName(SHEET_TASKS);

  if (noRows == 1) {
    ssTasks.getRange(row + noRows, TASKS_VALUES_TASKS_COL).setValue(1);
  } else if (noRows == 2 && ssTasks.getRange(row + noRows - 1, TASKS_VALUES_TASKS_COL).getValue() == 1) {
    ssTasks.getRange(row + noRows - 1, TASKS_VALUES_TASKS_COL).setValue(0.5);
    ssTasks.getRange(row + noRows, TASKS_VALUES_TASKS_COL).setValue(0.5);
  } else if (noRows == 4 && ssTasks.getRange(row + noRows - 1, TASKS_VALUES_TASKS_COL).getValue() == 0 && ssTasks.getRange(row + noRows - 2, TASKS_VALUES_TASKS_COL).getValue() == 0.5) {
    ssTasks.getRange(row + noRows - 3, TASKS_VALUES_TASKS_COL).setValue(0.25);
    ssTasks.getRange(row + noRows - 2, TASKS_VALUES_TASKS_COL).setValue(0.25);
    ssTasks.getRange(row + noRows - 1, TASKS_VALUES_TASKS_COL).setValue(0.25);
    ssTasks.getRange(row + noRows, TASKS_VALUES_TASKS_COL).setValue(0.25);
  } else if (noRows == 5 && ssTasks.getRange(row + noRows - 1, TASKS_VALUES_TASKS_COL).getValue() == 0.25 && ssTasks.getRange(row + noRows - 2, TASKS_VALUES_TASKS_COL).getValue() == 0.25) {
    ssTasks.getRange(row + noRows - 4, TASKS_VALUES_TASKS_COL).setValue(0.20);
    ssTasks.getRange(row + noRows - 3, TASKS_VALUES_TASKS_COL).setValue(0.20);
    ssTasks.getRange(row + noRows - 2, TASKS_VALUES_TASKS_COL).setValue(0.20);
    ssTasks.getRange(row + noRows - 1, TASKS_VALUES_TASKS_COL).setValue(0.20);
    ssTasks.getRange(row + noRows, TASKS_VALUES_TASKS_COL).setValue(0.20);
  } else
    ssTasks.getRange(row + noRows, TASKS_VALUES_TASKS_COL).setValue(0);
}

function setTask(collaborators, task) {
  var ui = SpreadsheetApp.getUi(); // gets user interface
  var ssTasks = SpreadsheetApp.getActive().getSheetByName(SHEET_TASKS);
  var data = SpreadsheetApp.getActive().getSheetByName(SHEET_DATA).getDataRange().getValues();

  var members = [];
  collaborators.forEach(email => {
    var rowMember = searchRowEmail(email)

    if (rowMember != -1)
      members.push(data[rowMember - 1][0])
    else
      ui.prompt(':(', 'There was an error trying to retrieve member data with the following email: "' + email + '"')
  });

  members.forEach(member => {
    var rowMember = searchRowMember(member) - 1

    if (rowMember != -1) {
      // creating task
      var row = 10 * rowMember + TASKS_MEMBER_INCREMENT
      var tasksRange = ssTasks.getRange(row, TASKS_VALUES_COL, 9)
      var noRows = getLastRow(tasksRange)

      if (noRows == 9)
        ui.prompt('parece ser que el cuerpo aieseco solo resiste 8 tareas')
      else { // setting task
        ssTasks.getRange(row + noRows, TASKS_VALUES_COL).setValue(task)
        setValues(row, noRows)
      }
    }
  });
}

/**
 * Sets the achievement formula for the according range to calculate that
 *
 * @param  {number} row The number of the row where this will be set up
 */
function setAchievementRange(row) {
  for (var i = 1; i <= NUM_TASKS; i++) {
    rowStr = (row + i).toString();
    SS_TASKS.getRange(row + i, getColumnNumber(TASKS_ACHIEVEMENT_COLUMN)).setFormula('=IF(' + TASKS_CHECKBOX_COLUMN + rowStr + '=TRUE, ' + TASKS_VALUE_COLUMN + rowStr + ', ' + TASKS_H_M_COLUMN + rowStr + '*' + TASKS_VALUE_COLUMN + rowStr + ')');
  }

  var achievementRange = SS_TASKS.getRange(row + 1, 6, NUM_TASKS);
  achievementRange.setNumberFormat('0.00%');
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
  SS_TASKS.getRange(TASKS_EMAILS_COLLABORATORS).setValue('');
}

function resetRoutineControls() {
  var ssTasks = SpreadsheetApp.getActive().getSheetByName(SHEET_TASKS);
  ssTasks.getRange(TASKS_ROUTINE).setValue('');
  ssTasks.getRange(TASKS_DAYS[0], TASKS_DAYS[1] + 1).setValue('');
  ssTasks.getRange(TASKS_START).setValue('');
  ssTasks.getRange(TASKS_END).setValue('');
  ssTasks.getRange(TASKS_DURATION[0], TASKS_DURATION[1]).setValue('');
  ssTasks.getRange(TASKS_COLLABORATORS).setValue('');
  ssTasks.getRange(TASKS_DESCRIPTION).setValue('');
  ssTasks.getRange(TASKS_LOCATION).setValue('');
  SS_TASKS.getRange(TASKS_DAYS_CHOSEN).setValue('');
  SS_TASKS.getRange(TASKS_EMAILS_COLLABORATORS).setValue('');
}
