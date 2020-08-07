
/**
 * Creates an event in Google Calendar
 *
 * @param  {string} event         The name of the task or routine
 * @param  {string} date          The date when it is going to take place the event
 * @param  {string} start         The start hour of the event
 * @param  {string} end           The end hour of the event
 * @param  {string} member        The name of the member, who has been assigned to
 * @param  {string} collaborators A csv with the mails of the collaborators
 * @param  {string} description   The description of the event
 * @param  {string} location      The location of the event
 */
function addToGoogleCalendar(event, date, start, end, member, collaborators, description, location) {
  var ui = SpreadsheetApp.getUi(); // gets user interface
  var ssData = SpreadsheetApp.getActive().getSheetByName(SHEET_DATA);

  var calendarId = ssData.getRange(DATA_CAL_ID).getValue();
  if (calendarId == '') {
    ui.alert("🤔", "There is no calendar ID in " + SHEET_DATA + "!" + DATA_CAL_ID + "\nMake sure to set this up in order to arrange the tasks you give in Google Calendar(:", ui.ButtonSet.OK);
    return;
  }

  // Converting parameters into Date objects
  var d = new Date(date);
  var s = new Date(start);
  var e = new Date(end);

  // Creating actual start and end dates
  var startDate = new Date(d.getFullYear(), d.getMonth(), d.getDate(), s.getHours(), s.getMinutes(), s.getSeconds(), s.getMilliseconds());
  var endDate = new Date(d.getFullYear(), d.getMonth(), d.getDate(), e.getHours(), e.getMinutes(), e.getSeconds(), e.getMilliseconds());

  // creating Google Calendar event
  var rowMember = searchRowMember(member);
  var email = ssData.getRange(rowMember, DATA_EMAIL_COL).getValue();

  // checking again collaborators
  if (collaborators.includes(email)) {
    ui.alert('🙃', 'did ya really insisted on havin the same member as his own collaborator???\nit\'s okay, nevermind I gotcha', ui.ButtonSet.OK);
    email = collaborators;
    collaborators = '';
  }

  var options = {
    'location': location,
    // 'description': (description == '') ? 'No description' : description,
    'description': description,
    'guests': (collaborators == '') ? email : email + ',' + collaborators,
    'sendInvites': (SEND_INVITES) ? 'True' : 'False'
  };

  var eventCal = CalendarApp.getCalendarById(calendarId);
  eventCal.createEvent(event, startDate, endDate, options);
}
