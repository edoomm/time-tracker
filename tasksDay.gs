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
  if (daysChosen.includes(dayOption)) {
    // clearing day chosen
    SS_TASKS.getRange(TASKS_DAYS_DROPDOWN).setValue('');
    return;
  }

  // entering data
  var today = new Date();

  if (dayOption == 'Everyday' || dayOption == 'Once every two days') {
    if (dayOption == 'Everyday')
      daysChosenRange.setValue(DAYS.join());
    else if (dayOption == 'Once every two days') {

      var days = '';
      for (var i = 0; i < 3; i++) {
        var index = (2 * i + today.getDay()) % 6;
        days += (days == '') ? DAYS[index] : ',' + DAYS[index];
      }

      daysChosenRange.setValue(days);
    } else
      ui.alert('Wat?', 'This doesn\'t even make sense in the code, how did you do it tho?\nPlease tell me how you did it, I\'m impressed lol\n' + EMAIL, ui.ButtonSet.OK);
  } else {
    var days = (daysChosen == '') ? dayOption : daysChosen + "," + dayOption;
    ssTasks.getRange(TASKS_DAYS_CHOSEN).setValue(days);
  }

  // clearing day chosen
  SS_TASKS.getRange(TASKS_DAYS_DROPDOWN).setValue('');
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
  if (daysChosen[daysChosen.length - 1] == ',')
    daysChosenRange.setValue(daysChosen.substring(0, daysChosen.length - 1));
}

function clearDays() {
  SpreadsheetApp.getActive().getSheetByName(SHEET_TASKS).getRange(TASKS_DAYS_CHOSEN).setValue('');
}
