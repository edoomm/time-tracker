function addNewMember() {
  var ui = SpreadsheetApp.getUi(); // gets user interface
  var ss = SpreadsheetApp.getActive(); // assign active spreadsheet to variable
  var ssData = ss.getSheetByName(SHEET_DATA);
  var ssTasks = ss.getSheetByName(SHEET_TASKS);
  var memberRange = ssData.getRange("A:A");

  // retrieving data from prompts
  var uiMember = ui.prompt('Insert member name');
  // validate data
  var member = uiMember.getResponseText();
  do {
    if (uiMember.getSelectedButton() != ui.Button.OK)
      return;

    if (searchRowMember(member) != -1) {
      ui.alert('Member "' + member + '" already exists, please choose a different name');
      uiMember = ui.prompt('Insert member name');
      member = uiMember.getResponseText();
    } else
      break;
  } while (true);

  uiMember = ui.prompt('Insert member email');
  if (uiMember.getSelectedButton() != ui.Button.OK)
    return;
  var email = uiMember.getResponseText();

  // once data has been retrieved and validated, insert data in sheets

  //<editor-fold> Tasks
  // creating task table
  var row = (10 * (getLastRow(memberRange))) + TASKS_MEMBER_INCREMENT;
  var headers = [
    ['Member', 'Task', 'Value', 'Fully done?', 'If not, how much?', 'Achievement', 'Total']
  ];
  var headersRange = ssTasks.getRange(row, 1, 1, 7);

  // inserting and formatting header data
  headersRange.setValues(headers);
  headersRange.setHorizontalAlignment("center");
  headersRange.setFontWeight("bold");

  // inserting and formatting member data
  var nameRange = ssTasks.getRange(row + 1, 1, NUM_TASKS);
  nameRange.mergeVertically();
  nameRange.setValue(member);
  nameRange.setHorizontalAlignment("center");
  nameRange.setVerticalAlignment("middle");
  nameRange.setFontWeight("bold");

  // inserting value number format
  ssTasks.getRange(row + 1, 3, NUM_TASKS).setNumberFormat('0.00%');

  // inserting checkboxes
  var checkboxesRange = ssTasks.getRange(row + 1, 4, NUM_TASKS);
  var enforceCheckbox = SpreadsheetApp.newDataValidation();
  enforceCheckbox.requireCheckbox();
  enforceCheckbox.setAllowInvalid(false);
  enforceCheckbox.build();
  checkboxesRange.setDataValidation(enforceCheckbox);

  // inserting 100% how much
  var hmRange = ssTasks.getRange(row + 1, 5, NUM_TASKS);
  hmRange.setValue('0');
  hmRange.setNumberFormat('0.00%');

  // inserting achievement
  setAchievementRange(row);

  // inserting total
  var totalRange = ssTasks.getRange(row + 1, 7, NUM_TASKS);
  totalRange.mergeVertically();
  totalRange.setFormula('=SUM(' + TASKS_ACHIEVEMENT_COLUMN + (row + 1).toString() + ':' + TASKS_ACHIEVEMENT_COLUMN + (row + 9).toString() + ')');
  totalRange.setNumberFormat('0.00%');
  totalRange.setHorizontalAlignment("center");
  totalRange.setVerticalAlignment("middle");
  totalRange.setFontWeight("bold");

  // creating ConditionalFormatting
  var appVal = getValueToApprove();

  var rules = ssTasks.getConditionalFormatRules();
  rules.push(createRule(false, 0.6, COLOR_FAIL, totalRange));
  rules.push(createRuleInInterval(0.6, appVal, COLOR_WARNING, totalRange));
  rules.push(createRuleInInterval(appVal, 1, COLOR_APPROVED, totalRange));
  rules.push(createRule(true, 1, COLOR_EXCELLENCE, totalRange));
  ssTasks.setConditionalFormatRules(rules);
  //</editor-fold>

  // inserting member in _Data
  var rowIndex = getLastRow(memberRange) + 1;
  ssData.getRange(rowIndex, DATA_MEMBER_COL).setValue(member);
  ssData.getRange(rowIndex, DATA_EMAIL_COL).setValue(email);

  // inserting member in History
  SS_HISTORY.getRange(rowIndex, HISTORY_MEMBER_COL).setValue(member);
}

function deleteMember() {
  var ui = SpreadsheetApp.getUi(); // gets user interface
  var ss = SpreadsheetApp.getActive(); // assign active spreadsheet to variable
  var ssData = ss.getSheetByName(SHEET_DATA);
  var ssTasks = ss.getSheetByName(SHEET_TASKS);
  var member = ssTasks.getRange(TASKS_MEMBER).getValue();

  // Validating data
  var found = false;
  if (member == '') {
    ui.alert('No member selected', 'Choose a member from cell B3', ui.ButtonSet.OK);
    return;
  }

  var rowMember = searchRowMember(member) - 1;
  if (rowMember != -1) {
    if (ui.alert('Do you really want to delete "' + member + '"?', '', ui.ButtonSet.YES_NO) == ui.Button.YES) {
      // deletes data in _Data
      ssData.getRange(rowMember + 1, 1, 1, 2).deleteCells(SpreadsheetApp.Dimension.ROWS);

      // deletes data in Tasks
      ssTasks.getRange(TASKS_MEMBER).setValue('');
      ssTasks.getRange(10 * rowMember + TASKS_MEMBER_INCREMENT, 1, 10, 7).deleteCells(SpreadsheetApp.Dimension.ROWS);

      // deletes history
      SS_HISTORY.getRange(rowMember + 1, HISTORY_MEMBER_COL, 1, 500).deleteCells(SpreadsheetApp.Dimension.ROWS);
      if (getLastRow(SS_HISTORY.getDataRange()) == 1)
        deleteAllHistory();
    }
  } else
    ui.prompt('No member found', 'Make sure you choose the member within the options the dropdown list gives you', ui.ButtonSet.OK);
}
