//Functions for changing Calendar Sheet

/**
 * Adds an event to the internal calendar of the SpreadSheet (SHEET_CALENDAR)
 *
 * @param  {string} event     The name of the task or routine
 * @param  {string} startDate The start date including hours
 * @param  {string} endDate   The end date including hours
 * @return {Range}           The range where the event was stored in SHEET_CALENDAR
 */
function addToCalendar(event, startDate, endDate) {
  // Getting cell coordinates
  var day = (startDate.getDay() != 0) ? startDate.getDay() + 1 : 8;
  var hour = startDate.getHours() + 2;

  var eventRange = null;
  // One hour events
  if (endDate.getHours() - startDate.getHours() == 1) {
    eventRange = SS_CALENDAR.getRange(hour, day);
    eventRange.setValue((eventRange.getValue() == '') ? event : eventRange.getValue() + ";" + event);
  }
  // More than one hour events
  else {
    eventRange = SS_CALENDAR.getRange(hour, day, endDate.getHours() - startDate.getHours());
    eventRange.mergeVertically();
    eventRange.setValue((eventRange.getValue() == '') ? event : eventRange.getValue() + ";" + event);
    eventRange.setHorizontalAlignment("center");
    eventRange.setVerticalAlignment("middle");
  }

  return eventRange;
}

/**
 * Removes an event in a specified cell
 *
 * @param  {string} cell  The cell where the event is contained in SHEET_CALENDAR
 * @param  {string} event The name of the task or routine
 */
function removeFromCalendar(cell, event) {
  try {
    var range = SS_CALENDAR.getRange(cell);
  } catch (e) {
    UI.alert('Cannot delete from "!' + SHEET_CALENDAR + '" range "' + range + '"\n\nException: ' + e);
    return;
  }

  // removing
  range.setValue(deleteSubstring(range.getValue(), event, ";"));
}