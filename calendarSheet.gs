//Functions that changing data sheets

/**
 * Adds an event to the internal calendar of the SpreadSheet (SHEET_CALENDAR)
 *
 * @param  {string} event     The name of the task or routine
 * @param  {string} startDate The start date including hours
 * @param  {string} endDate   The end date including hours
 * @return {Range}           The range where the event was stored in SHEET_CALENDAR
 */
function addToCalendar(event, startDate, endDate) {
  var ssCalendar = SpreadsheetApp.getActive().getSheetByName(SHEET_CALENDAR);

  // Getting cell coordinates
  var day = (startDate.getDay() != 0) ? startDate.getDay() + 1 : 8;
  var hour = startDate.getHours() + 2;

  var eventRange = null;
  // One hour events
  if (endDate.getHours() - startDate.getHours() == 1) {
    eventRange = ssCalendar.getRange(hour, day);
    eventRange.setValue((eventRange.getValue() == '') ? event : eventRange.getValue() + ";" + event);
  }
  // More than one hour events
  else {
    eventRange = ssCalendar.getRange(hour, day, endDate.getHours() - startDate.getHours());
    eventRange.mergeVertically();
    eventRange.setValue((eventRange.getValue() == '') ? event : eventRange.getValue() + ";" + event);
    eventRange.setHorizontalAlignment("center");
    eventRange.setVerticalAlignment("middle");
  }

  return eventRange;
}

function removeFromCalendar(range) {

}
