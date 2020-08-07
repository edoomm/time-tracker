//Functions that changing data sheet
function searchRowMember(member) {
  var data = SpreadsheetApp.getActive().getSheetByName(SHEET_DATA).getDataRange().getValues();

  for (var i = 0; i < data.length; i++)
    if (data[i][0] == member)
      return i + 1;

  return -1;
}

function searchRowEmail(email) {
  var data = SpreadsheetApp.getActive().getSheetByName(SHEET_DATA).getDataRange().getValues();

  for (var i = 0; i < data.length; i++)
    if (data[i][1] == email)
      return i + 1;

  return -1;
}


/**
 * Searches an Event in SHEET_DATA that's been previously assigned and returns row index
 *
 * @param  {string} event The name of the event to search
 * @param  {string} tr    'T' will stand for 'Task' and 'R' for routine
 * @return {number}       The row index if found, if not returns -1
 */
function searchEvent(event, tr) {
  var ssData = SpreadsheetApp.getActive().getSheetByName(SHEET_DATA);
  var eventsRange = ssData.getRange(DATA_EVENT + ":" + DATA_TASKROUTINE[0]);

  var data = eventsRange.getValues();
  for (var i = 0; i < data.length; i++)
    if (data[i][0] == event && data[i][1] == tr)
      return i + DATA_INITIAL_EVENT_ROW - 1;

  return -1;
}

/**
 * Gets the value minumum to approve for the members
 *
 * @return {number} The value to approve from [0,1] or 0.6 if occurs an error
 */
function getValueToApprove() {
  try {
    approves = SS_DATA.getRange(DATA_APPROVE_VAL).getValue();
    return appVal = (approves < 0) ? approves : approves / 100;
  } catch (e) {
    UI.alert('Exception thrown when trying to retrieve value to approve from ' + SHEET_DATA + ' in cell ' + DATA_APPROVE_VAL + '\nThe value could be not written as a percentage value must be written. The value to approve will be 60% as default\n\nException: ' + e);
    return 0.6;
  }
}


/**
 * Gets the total number of members registered in the current tracker through counting members' names
 * @return {number} The number of members registered
 */
function getMembershipNumber() {
  var data = SpreadsheetApp.getActive().getSheetByName(SHEET_DATA).getDataRange().getValues();

  for (var i = 0; i < data.length; i++)
    if (data[i][0] == '' && data[i][1] == '')
      return i - 1;

  return 0;
}

/**
 * Sets the data event in SHEET_DATA and in SHEET_CALENDAR
 *
 * @param  {string} event       The name of the task or routine
 * @param  {string} tr          'T' stands for task and 'R' for routine
 * @param  {string} members     The emails of the members assigned
 * @param  {string} from        The start hour
 * @param  {string} to          The end hour
 * @param  {string} description The description of the event
 * @param  {string} location    The location where the event will take place
 * @param  {(string|Array.<string>)} date        The date of the event, for routines this value should be the following dates
 * @param  {string} days        A csv for the days ['Mon', 'Tue', ...] that the event will repeat (for tasks this value should be null)
 * @param  {number} weeks       The number of weeks that the event will be repeating (for tasks this value should be null)
 */
function setDataEvent(event, tr, members, from, to, description, location, date, days, weeks) {
  var ui = SpreadsheetApp.getUi(); // gets user interface
  var ssData = SpreadsheetApp.getActive().getSheetByName(SHEET_DATA);

  var addedRows = getLastRow(ssData.getRange(DATA_EVENT + ":" + DATA_EVENT[0])) - 1;

  ssData.getRange(DATA_INITIAL_EVENT_ROW + addedRows, DATA_EVENT_COL).setValue(event);
  ssData.getRange(DATA_INITIAL_EVENT_ROW + addedRows, DATA_EVENT_COL + 1).setValue(tr);
  ssData.getRange(DATA_INITIAL_EVENT_ROW + addedRows, DATA_EVENT_COL + 2 + 1).setValue(members);
  ssData.getRange(DATA_INITIAL_EVENT_ROW + addedRows, DATA_EVENT_COL + 3 + 1).setValue(from);
  ssData.getRange(DATA_INITIAL_EVENT_ROW + addedRows, DATA_EVENT_COL + 4 + 1).setValue(to);
  ssData.getRange(DATA_INITIAL_EVENT_ROW + addedRows, DATA_EVENT_COL + 5 + 1).setValue(description);
  ssData.getRange(DATA_INITIAL_EVENT_ROW + addedRows, DATA_EVENT_COL + 6 + 1).setValue(location);
  ssData.getRange(DATA_INITIAL_EVENT_ROW + addedRows, DATA_EVENT_COL + 7 + 1).setValue(date);
  ssData.getRange(DATA_INITIAL_EVENT_ROW + addedRows, DATA_EVENT_COL + 8 + 1).setValue(days);
  ssData.getRange(DATA_INITIAL_EVENT_ROW + addedRows, DATA_EVENT_COL + 9 + 1).setValue(weeks);

  if (tr == 'T') {
    // Converting parameters into Date objects
    var d = new Date(date);
    var s = new Date(from);
    var e = new Date(to);

    // Creating actual start and end dates
    var startDate = new Date(d.getFullYear(), d.getMonth(), d.getDate(), s.getHours(), s.getMinutes(), s.getSeconds(), s.getMilliseconds());
    var endDate = new Date(d.getFullYear(), d.getMonth(), d.getDate(), e.getHours(), e.getMinutes(), e.getSeconds(), e.getMilliseconds());

    var cell = addToCalendar(event, startDate, endDate).getA1Notation();
    ssData.getRange(DATA_INITIAL_EVENT_ROW + addedRows, DATA_EVENT_COL + 2).setValue(cell);
  } else if (tr == 'R') {
    // creating all necessary events in SHEET_CALENDAR
    var arrDays = days.split(',');
    var cells = '';
    for (var i = 0; i < arrDays.length; i++) {
      // Converting parameters into Date objects
      var d = new Date(date[i]);
      var s = new Date(from);
      var e = new Date(to);

      // Creating actual start and end dates
      var startDate = new Date(d.getFullYear(), d.getMonth(), d.getDate(), s.getHours(), s.getMinutes(), s.getSeconds(), s.getMilliseconds());
      var endDate = new Date(d.getFullYear(), d.getMonth(), d.getDate(), e.getHours(), e.getMinutes(), e.getSeconds(), e.getMilliseconds());

      cells += addToCalendar(event, startDate, endDate).getA1Notation();
      if (i != arrDays.length - 1)
        cells += ',';
    }
    SS_DATA.getRange(DATA_INITIAL_EVENT_ROW + addedRows, getColumnNumber(DATA_CALENDAR_CELL[0])).setValue(cells);
  } else {
    throw 'No correct usage for "tr" variable in setDataEvent(), use "T" for Tasks and "R" for Routines';
  }
}
