//Functions for changing the tasks sheet

function setValues(row, noRows) {
  if (noRows == 1) {
    SS_TASKS.getRange(row + noRows, TASKS_VALUES_TASKS_COL).setValue(1);
  } else if (noRows == 2 && SS_TASKS.getRange(row + noRows - 1, TASKS_VALUES_TASKS_COL).getValue() == 1) {
    SS_TASKS.getRange(row + noRows - 1, TASKS_VALUES_TASKS_COL).setValue(0.5);
    SS_TASKS.getRange(row + noRows, TASKS_VALUES_TASKS_COL).setValue(0.5);
  } else if (noRows == 4 && SS_TASKS.getRange(row + noRows - 1, TASKS_VALUES_TASKS_COL).getValue() == 0 && SS_TASKS.getRange(row + noRows - 2, TASKS_VALUES_TASKS_COL).getValue() == 0.5) {
    SS_TASKS.getRange(row + noRows - 3, TASKS_VALUES_TASKS_COL).setValue(0.25);
    SS_TASKS.getRange(row + noRows - 2, TASKS_VALUES_TASKS_COL).setValue(0.25);
    SS_TASKS.getRange(row + noRows - 1, TASKS_VALUES_TASKS_COL).setValue(0.25);
    SS_TASKS.getRange(row + noRows, TASKS_VALUES_TASKS_COL).setValue(0.25);
  } else if (noRows == 5 && SS_TASKS.getRange(row + noRows - 1, TASKS_VALUES_TASKS_COL).getValue() == 0.25 && SS_TASKS.getRange(row + noRows - 2, TASKS_VALUES_TASKS_COL).getValue() == 0.25) {
    SS_TASKS.getRange(row + noRows - 4, TASKS_VALUES_TASKS_COL).setValue(0.20);
    SS_TASKS.getRange(row + noRows - 3, TASKS_VALUES_TASKS_COL).setValue(0.20);
    SS_TASKS.getRange(row + noRows - 2, TASKS_VALUES_TASKS_COL).setValue(0.20);
    SS_TASKS.getRange(row + noRows - 1, TASKS_VALUES_TASKS_COL).setValue(0.20);
    SS_TASKS.getRange(row + noRows, TASKS_VALUES_TASKS_COL).setValue(0.20);
  } else
    SS_TASKS.getRange(row + noRows, TASKS_VALUES_TASKS_COL).setValue(0);
}

function setTask(collaborators, task) {
  var data = SS_DATA.getDataRange().getValues();

  var members = [];
  collaborators.forEach(email => {
    var rowMember = searchRowEmail(email)

    if (rowMember != -1)
      members.push(data[rowMember - 1][0])
    else
      UI.prompt(':(', 'There was an error trying to retrieve member data with the following email: "' + email + '"')
  });

  members.forEach(member => {
    var rowMember = searchRowMember(member) - 1

    if (rowMember != -1) {
      // creating task
      var row = 10 * rowMember + TASKS_MEMBER_INCREMENT
      var tasksRange = SS_TASKS.getRange(row, TASKS_VALUES_COL, 9)
      var noRows = getLastRow(tasksRange)

      if (noRows == 9)
        UI.prompt('parece ser que el cuerpo aieseco solo resiste 8 tareas')
      else { // setting task
        SS_TASKS.getRange(row + noRows, TASKS_VALUES_COL).setValue(task)
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
  SS_TASKS.getRange(TASKS_TASK).setValue('');
  SS_TASKS.getRange(TASKS_DATE).setValue(DATE_CAPTION);
  SS_TASKS.getRange(TASKS_START).setValue('');
  SS_TASKS.getRange(TASKS_END).setValue('');
  SS_TASKS.getRange(TASKS_COLLABORATORS).setValue('');
  SS_TASKS.getRange(TASKS_DESCRIPTION).setValue('');
  SS_TASKS.getRange(TASKS_LOCATION).setValue('');
  SS_TASKS.getRange(TASKS_EMAILS_COLLABORATORS).setValue('');
}

function resetRoutineControls() {
  SS_TASKS.getRange(TASKS_ROUTINE).setValue('');
  SS_TASKS.getRange(TASKS_DAYS[0], TASKS_DAYS[1] + 1).setValue('');
  SS_TASKS.getRange(TASKS_START).setValue('');
  SS_TASKS.getRange(TASKS_END).setValue('');
  SS_TASKS.getRange(TASKS_DURATION[0], TASKS_DURATION[1]).setValue('');
  SS_TASKS.getRange(TASKS_COLLABORATORS).setValue('');
  SS_TASKS.getRange(TASKS_DESCRIPTION).setValue('');
  SS_TASKS.getRange(TASKS_LOCATION).setValue('');
  SS_TASKS.getRange(TASKS_DAYS_CHOSEN).setValue('');
  SS_TASKS.getRange(TASKS_EMAILS_COLLABORATORS).setValue('');
}