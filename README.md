Google Apps Script Operations Portal

1. What I Built

This project is a simple operations portal using Google Apps Script and Google Sheets. It helps manage tasks, prevent repeated saves, send task-completion email notifications separately, and control attendance settings with a record of changes.

The project uses made-up test data only.

2. Technologies Used

- Google Apps Script
- Google Sheets
- JavaScript
- CacheService for temporary duplicate-request protection
- LockService for controlling simultaneous script operations
- MailApp for email notifications
- Git and GitHub for source-code management

3. Google Sheets Structure

The project uses these sheets:

- Tasks: Stores task IDs, assignees, descriptions, statuses, and completion timestamps.
- Users: Stores user IDs, names, roles, and attendance requirements.
- EmailQueue: Stores email notifications waiting to be processed, their status, and their send timestamps.
- SettingsLog: Records attendance-setting changes, including who made the change, which user was affected, and the old and new values.

The Tasks sheet contains 500 dummy tasks, from `t1` to `t500`. The Users sheet contains five dummy users, including one administrator.

 4. Features and How to Test Them

  Feature 1: Update a Task

The `updateTask` action finds a task by its ID without looping through every row. It updates only the matching row. When a task becomes Complete, the script records its completion date.

How to test:

1. Open the deployed Google Apps Script Web App URL ending in `/exec`.
2. Send a POST request with the action `updateTask`, a task ID such as `t101`, and a status such as `In Progress`.
3. Check the response for a successful update.
4. Open the Tasks sheet and verify that the task's status changed.
5. Mark a test task Complete and verify that its `completedAt` field is populated.

Feature 2: Prevent Duplicate Saves

The script uses CacheService to identify an identical task and status update received within the configured 10-second window. A repeated request is ignored and reported as a duplicate.

How to test:

1. Send an `updateTask` request for a task.
2. Send the identical request again within 10 seconds.
3. Check that the second response indicates that the duplicate update was ignored.
4. Confirm that the task was not processed again.

Cache-based duplicate protection is temporary and is not a permanent record of every request.

Feature 3: Queue Email Notifications

When a task becomes Complete, the script adds an email notification to the EmailQueue sheet rather than sending the email during the task-save request.

The `processEmailQueue` function processes pending notifications through a time-driven trigger.

How to test:

1. Complete a test task.
2. Open the EmailQueue sheet and confirm that a notification was queued.
3. Run `processEmailQueue` manually in Apps Script to test email processing.
4. Verify the queue status and sent timestamp.
5. Confirm that the time-driven trigger is configured to run the processor automatically.

Test emails should use an email address you control.

Feature 4: Admin-Only Attendance Setting and Audit Log

The `setAttendanceRequired` action changes a user's attendance requirement only when the role check accepts the requesting actor as an administrator. Successful changes are recorded in SettingsLog.

How to test:

1. Send a request using a normal user's ID as `actorId`.
2. Confirm that the request is rejected as non-admin.
3. Send a request using the test administrator's ID and a different attendance value.
4. Confirm that the setting changes in the Users sheet.
5. Open SettingsLog and verify the timestamp, actor, target user, old value, and new value.

5. How the Admin Check Works and Its Weakness

The current implementation looks up the supplied `actorId` in the Users sheet and checks whether that user's role is `admin`.

However, `actorId` is supplied by the caller. Someone could change the request and submit an administrator's ID. Therefore, this implementation demonstrates a basic role check but does not provide secure production authentication.

For production use, I would verify the user's identity through a trusted authentication mechanism and perform authorization checks on the server using that verified identity.

6. Execution Time

The screenshot at `screenshots/execution-time.png` shows the Apps Script execution information collected during testing.

The actual execution duration should be read from the Apps Script Executions page. Performance can vary depending on spreadsheet size, service delays, and the operations performed.

 7. What I Would Improve With More Time

- Replace the caller-supplied administrator ID with trusted authentication.
- Improve validation and error handling.
- Add reliable email retry handling and duplicate-notification protection.
- Add automated tests for task updates, duplicate requests, and access control.
- Test concurrent requests and monitor execution performance.
- Improve deployment management and maintain separate test and production environments.

8. Deployment and Security Notes

The Apps Script Web App is deployed separately from the GitHub repository. The Web App URL and its access settings must be configured for the reviewer to test it.

The repository should contain the script source files, `appsscript.json`, this README, and the execution screenshot. It must not contain `.clasp.json`, credentials, private tokens, or sensitive personal data.

All data used for this assignment is fictional.