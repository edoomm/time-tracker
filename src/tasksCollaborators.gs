function addCollaborator() {
  // checking if it is not the same person
  var collaborator = SS_TASKS.getRange(TASKS_COLLABORATORS).getValue();
  if (collaborator == SS_TASKS.getRange(TASKS_MEMBER).getValue()) {
    UI.alert('🤨', 'You can\'t choose the same member as his collaborator', UI.ButtonSet.OK);
    SS_TASKS.getRange(TASKS_COLLABORATORS).setValue('');
    return;
  }

  // verifying if member has been chosen
  if (collaborator == '') {
    UI.alert(':(', 'You didn\'t choose a member from ' + TASKS_COLLABORATORS, UI.ButtonSet.OK);
    return;
  }

  var rowCollaborator = searchRowMember(collaborator);
  if (rowCollaborator == -1) {
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
    return;
  }

  var rowCollaborator = searchRowMember(collaborator);
  if (rowCollaborator == -1) {
    UI.alert(':(', 'An error has ocurred while trying to choose a member from ' + TASKS_COLLABORATORS + '\nMake sure data is not corruputed in ' + SHEET_DATA, UI.ButtonSet.OK);
    return;
  }

  var email = SS_DATA.getRange(rowCollaborator, DATA_EMAIL_COL).getValue();
  var emailColRange = SS_TASKS.getRange(TASKS_EMAILS_COLLABORATORS);

  emailColRange.setValue(emailColRange.getValue().replace(email, '').replace(',,', ','));
  var emails = emailColRange.getValue();
  if (emails[0] == ',')
    emailColRange.setValue(emails.substring(1, emails.length));
  if (emails[emails.length - 1] == ',')
    emailColRange.setValue(emails.substring(0, emails.length - 1));

  SS_TASKS.getRange(TASKS_COLLABORATORS).setValue('');
}

function clearEmailsColl() {
  SS_TASKS.getRange(TASKS_EMAILS_COLLABORATORS).setValue('');
}