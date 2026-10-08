function doGet(e) {
  return ContentService.createTextOutput(
    JSON.stringify({
      success: true,
      message: 'Operations Portal API is running',
    }),
  ).setMimeType(ContentService.MimeType.JSON);
}

function setupTestData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(
    '1eKOh70td9eLf9UaoPC36mmb2H4rQDzEr1CqlvYI_xVE',
  );

  const tasksSheet = ss.getSheetByName('Tasks');
  const usersSheet = ss.getSheetByName('Users');
  const settingsLogSheet = ss.getSheetByName('SettingsLog');
  const emailQueueSheet = ss.getSheetByName('EmailQueue');

  // clear old data
  tasksSheet.clearContents();
  usersSheet.clearContents();
  settingsLogSheet.clearContents();
  emailQueueSheet.clearContents();

  //tasks

  tasksSheet
    .getRange(1, 1, 1, 5)
    .setValues([['id', 'assignee', 'desc', 'status', 'completedAt']]);

  const tasks = [];

  for (let i = 1; i <= 500; i++) {
    tasks.push([
      `t${i}`,
      `u${((i - 1) % 5) + 1}`,
      `Sample task ${i}`,
      'Pending',
      '',
    ]);
  }

  tasksSheet.getRange(2, 1, tasks.length, 5).setValues(tasks);

  //users

  usersSheet
    .getRange(1, 1, 1, 4)
    .setValues([['id', 'name', 'role', 'attendanceRequired']]);

  const users = [
    ['u1', 'Arun', 'user', true],
    ['u2', 'Rahul', 'user', false],
    ['u3', 'Anu', 'admin', true],
    ['u4', 'Meera', 'user', true],
    ['u5', 'John', 'user', false],
  ];

  usersSheet.getRange(2, 1, users.length, 4).setValues(users);

  //settings

  settingsLogSheet
    .getRange(1, 1, 1, 5)
    .setValues([['changedAt', 'changedBy', 'target', 'oldValue', 'newValue']]);

  //email

  emailQueueSheet
    .getRange(1, 1, 1, 7)
    .setValues([
      ['queuedAt', 'taskId', 'email', 'subject', 'message', 'status', 'sentAt'],
    ]);

  Logger.log('Test data created successfully');
}
