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

    if (data.action === 'setAttendanceRequired') {
      return jsonResponse(setAttendanceRequired(data));
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


// ------ 
// function doPost(e) {
//   try {
//     const data = JSON.parse(e.postData.contents);

//     if (data.action === 'updateTask') {
//       return jsonResponse(updateTask(data));
//     }

//     return jsonResponse({
//       success: false,
//       message: 'Unknown action',
//     });
//   } catch (error) {
//     return jsonResponse({
//       success: false,
//       message: error.message,
//     });
//   }
// }

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

  const email = '';

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


// ---------------------------------------

function processEmailQueue() {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(5000);

    const ss = SpreadsheetApp.openById(
      '1eKOh70td9eLf9UaoPC36mmb2H4rQDzEr1CqlvYI_xVE',
    );
    const sheet = ss.getSheetByName('EmailQueue');

    const lastRow = sheet.getLastRow();

    if (lastRow < 2) {
      Logger.log('No emails in queue');
      return;
    }

    const data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();

    for (let i = 0; i < data.length; i++) {
      const row = data[i];

      const email = row[2];
      const subject = row[3];
      const message = row[4];
      const status = row[5];

      if (status !== 'Pending') {
        continue;
      }

      const sheetRow = i + 2;

      try {
        MailApp.sendEmail(email, subject, message);

        sheet.getRange(sheetRow, 6, 1, 2).setValues([['Sent', new Date()]]);
      } catch (error) {
        sheet.getRange(sheetRow, 6).setValue('Failed');

        Logger.log(
          `Failed to send email for row ${sheetRow}: ${error.message}`,
        );
      }
    }
  } finally {
    lock.releaseLock();
  }
}
// --------------- 

function setAttendanceRequired(data) {
  const actorId = data.actorId;
  const targetId = data.targetId;
  const required = data.required;

  if (!actorId || !targetId || typeof required !== 'boolean') {
    return {
      success: false,
      message: 'actorId, targetId and boolean required are required',
    };
  }

  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(5000);

    const ss = SpreadsheetApp.openById(
      '1eKOh70td9eLf9UaoPC36mmb2H4rQDzEr1CqlvYI_xVE',
    );

    const usersSheet = ss.getSheetByName('Users');
    const logSheet = ss.getSheetByName('SettingsLog');

    // Find the actor without looping through all users
    const actorCell = usersSheet
      .getRange('A:A')
      .createTextFinder(actorId)
      .matchEntireCell(true)
      .findNext();

    if (!actorCell) {
      return {
        success: false,
        message: 'Actor not found',
      };
    }

    const actorRow = actorCell.getRow();
    const actor = usersSheet.getRange(actorRow, 1, 1, 4).getValues()[0];

    if (actor[2] !== 'admin') {
      return {
        success: false,
        message: 'Access denied: admin only',
      };
    }

    // Find the target user
    const targetCell = usersSheet
      .getRange('A:A')
      .createTextFinder(targetId)
      .matchEntireCell(true)
      .findNext();

    if (!targetCell) {
      return {
        success: false,
        message: 'Target user not found',
      };
    }

    const targetRow = targetCell.getRow();
    const target = usersSheet.getRange(targetRow, 1, 1, 4).getValues()[0];

    const oldValue = target[3];

    // 
    if (oldValue === required) {
      return {
        success: true,
        message: 'Attendance setting is already set to this value',
        changed: false,
      };
    }

    // Update only the target user's attendanceRequired cell
    usersSheet.getRange(targetRow, 4).setValue(required);

    // Record the successful change
    logSheet.appendRow([new Date(), actorId, targetId, oldValue, required]);

    return {
      success: true,
      message: 'Attendance setting updated successfully',
      targetId: targetId,
      oldValue: oldValue,
      newValue: required,
    };
  } finally {
    lock.releaseLock();
  }
}


