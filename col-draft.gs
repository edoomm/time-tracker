function addCollaborator() {
<<<<<<< HEAD
  var ui = SpreadsheetApp.getUi(); // gets user interface
  var ssTasks = SpreadsheetApp.getActive().getSheetByName(SHEET_TASKS);
  var ssData = SpreadsheetApp.getActive().getSheetByName(SHEET_DATA);

  // checking if it is not the same person
  var collaborator = ssTasks.getRange(TASKS_COLLABORATORS).getValue();
  if (collaborator == ssTasks.getRange(TASKS_MEMBER).getValue()) {
    ui.alert('🤨', 'You can\'t choose the same member as his collaborator', ui.ButtonSet.OK);
    ssTasks.getRange(TASKS_COLLABORATORS).setValue('');
=======
  // checking if it is not the same person
  var collaborator = SS_TASKS.getRange(TASKS_COLLABORATORS).getValue();
  if (collaborator == SS_TASKS.getRange(TASKS_MEMBER).getValue()) {
    UI.alert('🤨', 'You can\'t choose the same member as his collaborator', UI.ButtonSet.OK);
    SS_TASKS.getRange(TASKS_COLLABORATORS).setValue('');
>>>>>>> 16b5a17c4c1a403153bec06f234d85cead9fb783
    return;
  }

  // verifying if member has been chosen
  if (collaborator == '') {
<<<<<<< HEAD
    ui.alert(':(', 'You didn\'t choose a member from ' + TASKS_COLLABORATORS, ui.ButtonSet.OK);
=======
    UI.alert(':(', 'You didn\'t choose a member from ' + TASKS_COLLABORATORS, UI.ButtonSet.OK);
>>>>>>> 16b5a17c4c1a403153bec06f234d85cead9fb783
    return;
  }

  var rowCollaborator = searchRowMember(collaborator);
  if (rowCollaborator == -1) {
<<<<<<< HEAD
    ui.alert(':(', 'An error has ocurred while trying to choose a member from ' + TASKS_COLLABORATORS + '\nMake sure data is not corruputed in ' + SHEET_DATA, ui.ButtonSet.OK);
    return;
  }

  var email = ssData.getRange(rowCollaborator, DATA_EMAIL_COL).getValue();
  var emailColRange = ssTasks.getRange(TASKS_EMAILS_COLLABORATORS);
  if (emailColRange.getValue().includes(email))
    ui.alert('🤨', 'You\'ve already chosen ' + collaborator, ui.ButtonSet.OK);
  else {
    if (emailColRange.getValue() == '')
      ssTasks.getRange(TASKS_EMAILS_COLLABORATORS).setValue(email);
    else {
      var emRangeVal = emailColRange.getValue();
      ssTasks.getRange(TASKS_EMAILS_COLLABORATORS).setValue(emRangeVal + ',' + email);
    }
  }

  ssTasks.getRange(TASKS_COLLABORATORS).setValue('');
}

function removeCollaborator() {
  var ui = SpreadsheetApp.getUi(); // gets user interface
  var ssTasks = SpreadsheetApp.getActive().getSheetByName(SHEET_TASKS);
  var ssData = SpreadsheetApp.getActive().getSheetByName(SHEET_DATA);

  // verifying if member has been chosen
  var collaborator = ssTasks.getRange(TASKS_COLLABORATORS).getValue();
  if (collaborator == '') {
    ui.alert(':(', 'You didn\'t choose a member from ' + TASKS_COLLABORATORS, ui.ButtonSet.OK);
=======
    UI.alert(':(', 'An error has ocurred while trying to choose a member from ' + TASKS_COLLABORATORS + '\nMake sure data is not corruputed in ' + SHEET_DATA, UI.ButtonSet.OK);
    return;
  }

  var email = SS_DATA.getRange(rowCollaborator, DATA_EMAIL_COL).getValue();
  var emailColRange = SS_TASKS.getRange(TASKS_EMAILS_COLLABORATORS);
  if (emailColRange.getValue().includes(email))
    UI.alert('🤨', 'You\'ve already chosen ' + collaborator, UI.ButtonSet.OK);
  else {
    if (emailColRange.getValue() == '')
      SS_TASKS.getRange(TASKS_EMAILS_COLLABORATORS).setValue(email);
    else {
      var emRangeVal = emailColRange.getValue();
      SS_TASKS.getRange(TASKS_EMAILS_COLLABORATORS).setValue(emRangeVal + ',' + email);
    }
  }

  SS_TASKS.getRange(TASKS_COLLABORATORS).setValue('');
}

function removeCollaborator() {
  // verifying if member has been chosen
  var collaborator = SS_TASKS.getRange(TASKS_COLLABORATORS).getValue();
  if (collaborator == '') {
    UI.alert(':(', 'You didn\'t choose a member from ' + TASKS_COLLABORATORS, UI.ButtonSet.OK);
>>>>>>> 16b5a17c4c1a403153bec06f234d85cead9fb783
    return;
  }

  var rowCollaborator = searchRowMember(collaborator);
  if (rowCollaborator == -1) {
<<<<<<< HEAD
    ui.alert(':(', 'An error has ocurred while trying to choose a member from ' + TASKS_COLLABORATORS + '\nMake sure data is not corruputed in ' + SHEET_DATA, ui.ButtonSet.OK);
    return;
  }

  var email = ssData.getRange(rowCollaborator, DATA_EMAIL_COL).getValue();
  var emailColRange = ssTasks.getRange(TASKS_EMAILS_COLLABORATORS);
=======
    UI.alert(':(', 'An error has ocurred while trying to choose a member from ' + TASKS_COLLABORATORS + '\nMake sure data is not corruputed in ' + SHEET_DATA, UI.ButtonSet.OK);
    return;
  }

  var email = SS_DATA.getRange(rowCollaborator, DATA_EMAIL_COL).getValue();
  var emailColRange = SS_TASKS.getRange(TASKS_EMAILS_COLLABORATORS);
>>>>>>> 16b5a17c4c1a403153bec06f234d85cead9fb783

  emailColRange.setValue(emailColRange.getValue().replace(email, '').replace(',,', ','));
  var emails = emailColRange.getValue();
  if (emails[0] == ',')
    emailColRange.setValue(emails.substring(1, emails.length));
  if (emails[emails.length - 1] == ',')
    emailColRange.setValue(emails.substring(0, emails.length - 1));

<<<<<<< HEAD
  ssTasks.getRange(TASKS_COLLABORATORS).setValue('');
}

function clearEmailsColl() {
  SpreadsheetApp.getActive().getSheetByName(SHEET_TASKS).getRange(TASKS_EMAILS_COLLABORATORS).setValue('');
}
=======
  SS_TASKS.getRange(TASKS_COLLABORATORS).setValue('');
}

function clearEmailsColl() {
  SS_TASKS.getRange(TASKS_EMAILS_COLLABORATORS).setValue('');
}
>>>>>>> 16b5a17c4c1a403153bec06f234d85cead9fb783
