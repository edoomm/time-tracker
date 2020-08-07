//Constants for programs
// sheets names
const SHEET_DATA = '_Data';
const SHEET_TASKS = 'Tasks';
const SHEET_CALENDAR = 'Calendar';
const SHEET_HISTORY = 'History';
// SHEET_DATA constants
const DATA_MEMBER = "A2";
const DATA_HISTORY = "C2";
const DATA_MEMBER_COL = 1;
const DATA_EMAIL_COL = 2;
const DATA_HISTORY_COL = 3;
const DATA_CAL_ID = 'F3';
const DATA_APPROVE_VAL = 'F5';
const DATA_EVENT = 'H2';
const DATA_TASKROUTINE = 'I2';
/** @constant {String} - Represents the header of the column where a cell from SHEET_CALENDAR is storing an event */
const DATA_CALENDAR_CELL = 'J2';
/** @constant {String} - Represents the header of the column where the members of the event are stored */
const DATA_MEMBERS_CELL = 'K2';
const DATA_WEEKS = 'R2';
const DATA_EVENT_COL = 8;
const DATA_INITIAL_EVENT_ROW = 3;
const DATA_DEFAULT_APPROVE_VAL = 60;
// SHEET_TASKS constants
const TASKS_MEMBER_INCREMENT = 5;
const TASKS_TITLES_COL = 1;
const TASKS_VALUES_COL = 2;
/** @constant {String} - Represents the column where most of the values for tasks & routines are stored */
const TASKS_VALUES_COLUMN = 'B';
const TASKS_TASK = 'B1';
const TASKS_TASK_ROW = 1;
const TASKS_ROUTINE = 'B2';
const TASKS_ROUTINE_ROW = 2;
const TASKS_MEMBER = 'B3';
const TASKS_DATE = 'B4';
const TASKS_DATE_ROW = 4;
const TASKS_START = 'B5';
const TASKS_END = 'B6';
const TASKS_COLLABORATORS = 'B7';
const TASKS_EMAILS_COLLABORATORS = 'E7';
const TASKS_DESCRIPTION = 'B8';
const TASKS_LOCATION = 'B9';
const TASKS_SWITCH = 'C1';
const TASKS_DAYS = [4, 3];
const TASKS_DAYS_DROPDOWN = 'D4';
const TASKS_DAYS_CHOSEN = 'F4';
const TASKS_DURATION = [5, 4];
const TASKS_ADD_TASK_BUTTON = 'A11';
const TASKS_ADD_ROUTINE_BUTTON = 'A12';
const DATE_CAPTION = 'Double click to pop calendar up';
const DAYS = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
const TASKS_VALUE_COLUMN = 'C';
const TASKS_CHECKBOX_COLUMN = 'D';
const TASKS_CHECKBOX_COL = 4;
const TASKS_H_M_COLUMN = 'E';
const TASKS_ACHIEVEMENT_COLUMN = 'F';
const TASKS_TOTAL_COLUMN = 'G';
const TASKS_VALUES_TASKS_COL = 3;
const TASKS_NON_FIX_VALUES = 'A13';

const NUM_TASKS = 8;

/** @constant {string} - Cell where user chooses an event from SHEET_DATA to delete it */
const TASKS_EVENT_CHOSEN = 'D11';
// SHEET_CALENDAR constants
const CALENDAR_INITIAL_DATE = 'B2';
const CALENDAR_FINAL_DATE = 'G25';
// SHEET_HISTORY constants
const HISTORY_MEMBER_COL = 1;
const HISTORY_WEEKS_ROW = 1;
// calendar options
const SEND_INVITES = true;
// coder info
const EMAIL = 'eduardo.mendozamartinez@aiesec.net';

// SpreadSheets and User interface
/** @constant {Sheet} - Data Sheet where all the important data for the correct use of the SpreadSheet */
const SS_DATA = SpreadsheetApp.getActive().getSheetByName(SHEET_DATA);
/** @constant {Sheet} - Tasks Sheet where tasks or routines will be assigned */
const SS_TASKS = SpreadsheetApp.getActive().getSheetByName(SHEET_TASKS);
/** @constant {Sheet} - Calendar Sheet where everyone can see the activities of the week of all members */
const SS_CALENDAR = SpreadsheetApp.getActive().getSheetByName(SHEET_CALENDAR);
/** @constant {Sheet} - History Sheet for recording members achievement through weeks */
const SS_HISTORY = SpreadsheetApp.getActive().getSheetByName(SHEET_HISTORY);
/** @constant {Sheet} - An instance of the user-interface environment for a Google App that allows the script to add features like menus, dialogs, and sidebars. */
const UI = SpreadsheetApp.getUi();

// Colors for conditional formatting
const COLOR_FAIL = '#ea4335';
const COLOR_WARNING = '#fbbc04';
const COLOR_APPROVED = '#34a853';
const COLOR_EXCELLENCE = '#4285f4';