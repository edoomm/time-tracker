//Function for general purpose
function getIndexOf(element, collection) {
  for (var i = 0; i < collection.length; i++)
    if (collection[i] == element)
      return i;

  return -1;
}

/**
 * Gets the last row number within a given range
 * @param {Range} range - A range can be a single cell in a sheet or a group of adjacent cells in a sheet
 * @returns {number} The index of the last row with data in it
 */
function getLastRow(range) {
  return range.getValues().filter(String).length;
}

/**
 * Gets the last column number within a given range in a maximum of 9000 rows which is the minimum recursion in web apps
 * @see {@link https://bestwebhostingaustralia.org/browserscope-org-joins-aussie-hosting/?v=3&layout=simple}
 *
 * @param  {Range} range A range can be a single cell in a sheet or a group of adjacent cells in a sheet
 * @return {number} The last column number that contained data
 */
function getLastColumn(range) {
  return getLastColumn(range, 9000);
}

/**
 * Gets the last column number within a given range in a given row with a given maximum recursion
 * @see {@link https://bestwebhostingaustralia.org/browserscope-org-joins-aussie-hosting/?v=3&layout=simple}
 *
 * @param  {Range} range A range can be a single cell in a sheet or a group of adjacent cells in a sheet
 * @param  {number} limit The number of rows that can search through, this can call exceptions if number is too big
 * @return {number} The last column number that contained data
 */
function getLastColumn(range, limit) {
  var lastRow = getLastRow(range);
  if (lastRow == 0)
    return 0;

  var max = 0;
  var values = range.getValues();
  var uppLmt = (lastRow > limit) ? limit : lastRow;

  for (var i = 0; i < uppLmt; i++)
    if (max < values[i].length)
      max = values[i].length;

  return max;
}

function getNumElements(collection, separator) {
  if (collection == '')
    return 0;

  var num = 1;

  for (var i = 0; i < collection.length; i++)
    if (collection[i] == separator)
      num++;

  return num;
}

/**
 * Transforms A1Notation column [A,B,C,...] to number [1,2,3,...]
 *
 * @param  {string} chr The column given
 * @return {number} The number of the column
 */
function getColumnNumber(chr) {
  return chr.toLowerCase().charCodeAt(0) - 97 + 1;
}

/**
 * Creates a new conditional format rule within a interval of numbers
 *
 * @param  {number} bottom The bottom limit of the interval
 * @param  {number} upper The upper limit of the interval
 * @param  {string} background The color for the background in hexadecimal format
 * @param  {Range} range A range can be a single cell in a sheet or a group of adjacent cells in a sheet
 * @return {ConditionalFormatRuleBuilder} The rule ready to be pushed
 */
function createRuleInInterval(bottom, upper, background, range) {
  return SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(bottom, upper)
    .setBackground(background)
    .setRanges([range])
    .build();
}

/**
 * Creates a new conditional format rule for numbers greater or lower than a limit
 *
 * @param  {boolean} isGreater If it is true the function will use "whenNumberGreaterThan(limit)", otherwise "whenNumberLessThan(limit)"
 * @param  {number} limit The greater or less than number
 * @param  {string} background The color for the background in hexadecimal format
 * @param  {Range} range A range can be a single cell in a sheet or a group of adjacent cells in a sheet
 * @return {ConditionalFormatRuleBuilder} The rule ready to be pushed
 */
function createRule(isGreater, limit, background, range) {
  if (isGreater)
    return SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(limit)
      .setBackground(background)
      .setRanges([range])
      .build();
  else
    return SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(limit)
      .setBackground(background)
      .setRanges([range])
      .build();
}

/**
 * Verifies if a value is empty or not, and shows a message when value is empty if desired
 *
 * @param  {string} value   The variable that will be tested out
 * @param  {string} message The message which will be displayed if value is empty
 * @return {boolean} true when the value is empty, otherwise false
 */
function isEmptyValue(value, message) {
  if (value === "") {
    if (message != null)
      UI.alert(message);

    return true;
  }

  return false;
}