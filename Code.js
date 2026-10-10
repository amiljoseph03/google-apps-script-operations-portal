function doGet(e) {
  return ContentService.createTextOutput(
    JSON.stringify({
      success: true,
      message: 'Operations Portal API is running',
    }),
  ).setMimeType(ContentService.MimeType.JSON);
}

// -----------post

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    if (data.action === 'updateTask') {
      return jsonResponse(updateTask(data));
    }

    return jsonResponse({
      success: false,
      message: 'Unknown action',
    });
  } catch (error) {
    return jsonResponse({
      success: false,
      message: error.message,
    });
  }
}

// ...

const SPREADSHEET_ID = '1eKOh70td9eLf9UaoPC36mmb2H4rQDzEr1CqlvYI_xVE';

function getOperationsSpreadsheet() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(
    ContentService.MimeType.JSON,
  );
}

function setupTestData() {
  const ss = getOperationsSpreadsheet();

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

// ------------------------------------------------------------ update

function updateTask(data) {
  const id = data.id;
  const status = data.status;

  if (!id || !status) {
    return {
      success: false,
      message: 'id and status are required',
    };
  }

  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(5000);

    const cache = CacheService.getScriptCache();
    const duplicateKey = `task_update_${id}_${status}`;

    if (cache.get(duplicateKey)) {
      return {
        success: true,
        duplicate: true,
        message: 'Duplicate update ignored',
      };
    }

    cache.put(duplicateKey, '1', 10);

    const ss = getOperationsSpreadsheet();
    const sheet = ss.getSheetByName('Tasks');

    const idCell = sheet
      .getRange('A:A')
      .createTextFinder(id)
      .matchEntireCell(true)
      .findNext();

    if (!idCell) {
      return {
        success: false,
        message: 'Task not found',
      };
    }

    const row = idCell.getRow();

    const currentValues = sheet.getRange(row, 1, 1, 5).getValues()[0];

    const oldStatus = currentValues[3];

    let completedAt = currentValues[4];

    if (status === 'Complete' && oldStatus !== 'Complete') {
      completedAt = new Date();
    }

    if (status !== 'Complete') {
      completedAt = '';
    }

    const newValues = [
      currentValues[0],
      currentValues[1],
      currentValues[2],
      status,
      completedAt,
    ];

    sheet.getRange(row, 1, 1, 5).setValues([newValues]);

    if (status === 'Complete' && oldStatus !== 'Complete') {
      addEmailToQueue(id, currentValues[1]);
    }

    return {
      success: true,
      message: 'Task updated successfully',
      id: id,
      status: status,
    };
  } finally {
    lock.releaseLock();
  }
}

// ----------------email////////

function addEmailToQueue(taskId, assignee) {
  const ss = getOperationsSpreadsheet();
  const sheet = ss.getSheetByName('EmailQueue');

  const email = '...';

  sheet.appendRow([
    new Date(),
    taskId,
    email,
    `Task ${taskId} completed`,
    `Task ${taskId} has been marked as Complete.`,
    'Pending',
    '',
  ]);
}


