function doGet(e) {
  var action = e.parameter.action;
  try {

    if (action === "refresh_hardcoded_data") {
    return runShortTask('insert_hardcoded_categorised_Data_phone_email', 'Refresh data');
    } 
    else if (action === "duplicate_sheet") {
      return runShortTask('duplicate_sus_player_without_percentile_lt', 'Duplicate sus data');
    }

    // if (action === "refresh_data_phone_email") {
    //   const ss = SpreadsheetApp.getActiveSpreadsheet();
    //   const sheet = ss.getSheetByName("Deposit_Report_RT_v2");
    //   const statusCell = sheet.getRange("C4");

    //   const result = startAsyncTask('Run_Deposit_Raw_phone_email_RT', 'Task fetching data from BQ');
    //   statusCell.setValue(result);

    //   return ContentService.createTextOutput(result).setMimeType(ContentService.MimeType.TEXT);
    // // }
    // // if (action === "refresh_deposit_raw") {
    // //   const ss = SpreadsheetApp.getActiveSpreadsheet();
    // //   const sheet = ss.getSheetByName("Deposit_Report");
    // //   const statusCell = sheet.getRange("C3");

    // //   const result = startAsyncTask('Run_Deposit_Raw_v1', 'Deposit task fetching data from BQ');
    // //   statusCell.setValue(result);

    // //   return ContentService.createTextOutput(result).setMimeType(ContentService.MimeType.TEXT);
    // } else if (action === "refresh_hardcoded_data") {
    // return runShortTask('insert_hardcoded_categorised_Data_phone_email', 'Refresh data');
    // }  else if (action === "refresh_extreme_case") {
    // return runShortTask('insert_extreme_cases_Data_RT_phone_email', 'Refresh extreme data');
    // } 
    // else {
    //   return ContentService.createTextOutput("No action specified");
    // }
  } catch(err) {
    return ContentService.createTextOutput("Error: " + err.message);
  }
}

// -------------------- Status Checker --------------------
function getTaskStatusMessage() {
  const props = PropertiesService.getScriptProperties();
  const status = props.getProperty('task_status') || 'IDLE';
  const taskName = props.getProperty('task_name') || 'None';
  const startTime = props.getProperty('task_start_time');
  const endTime = props.getProperty('task_end_time');
  const error = props.getProperty('task_error');
  
  let message = '';
  
  if (status === 'RUNNING' && startTime) {
    const elapsed = Math.floor((Date.now() - parseInt(startTime)) / 1000);
    const minutes = Math.floor(elapsed / 60);
    const seconds = elapsed % 60;
    message = `🔄 RUNNING: ${taskName}\n` +
              `Time elapsed: ${minutes}m ${seconds}s\n` +
              `Status: Processing in background...\n\n` +
              `Check execution logs for detailed progress.`;
  } else if (status === 'COMPLETED' && endTime) {
    const totalTime = Math.floor((parseInt(endTime) - parseInt(startTime)) / 1000);
    const minutes = Math.floor(totalTime / 60);
    const seconds = totalTime % 60;
    message = `✅ COMPLETED: ${taskName}\n` +
              `Total time: ${minutes}m ${seconds}s\n` +
              `Status: Task finished successfully`;
  } else if (status === 'ERROR') {
    message = `❌ ERROR in ${taskName}\n` +
              `Error: ${error || 'Unknown error'}\n` +
              `Check execution logs for details.`;
  } else {
    message = `💤 IDLE\n` +
              `No task currently running.\n` +
              `Last task: ${taskName || 'None'}`;
  }
  
  return message;
}
// -------------------- Wrapper Functions for Buttons --------------------
function runrefresh_hardcoded_data() {
  var url = "https://script.google.com/macros/s/AKfycbwIoLXEaWwji4k9o4fFXi1yKt32fiZqU3RX6GRWjgSxTnnh9esWffFJJmV3pR8v_LOkZg/exec?action=refresh_hardcoded_data";
  var response = UrlFetchApp.fetch(url, {muteHttpExceptions: true});
  Logger.log(response.getContentText());
}

function runDuplicateSheet() {
  var url = "https://script.google.com/macros/s/AKfycbwIoLXEaWwji4k9o4fFXi1yKt32fiZqU3RX6GRWjgSxTnnh9esWffFJJmV3pR8v_LOkZg/exec?action=duplicate_sheet";
  var response = UrlFetchApp.fetch(url, {muteHttpExceptions: true});
  Logger.log(response.getContentText());
}

// function runrefresh_data_phone_email() {
//   var url = "https://script.google.com/macros/s/AKfycbwIoLXEaWwji4k9o4fFXi1yKt32fiZqU3RX6GRWjgSxTnnh9esWffFJJmV3pR8v_LOkZg/exec?action=refresh_data_phone_email";
//   var response = UrlFetchApp.fetch(url, {muteHttpExceptions: true});
//   Logger.log(response.getContentText());
// }

// function runrefresh_extreme_data() {
//   var url = "https://script.google.com/macros/s/AKfycbwIoLXEaWwji4k9o4fFXi1yKt32fiZqU3RX6GRWjgSxTnnh9esWffFJJmV3pR8v_LOkZg/exec?action=refresh_extreme_case";
//   var response = UrlFetchApp.fetch(url, {muteHttpExceptions: true});
//   Logger.log(response.getContentText());
// }

// -------------------- Async Task Starter --------------------
function startAsyncTask(functionName, taskName) {
  const props = PropertiesService.getScriptProperties();
  
  const currentStatus = props.getProperty('task_status');

  if (currentStatus === 'RUNNING') {
    const currentTask = props.getProperty('task_name') || 'Unknown';
    const startTime = props.getProperty('task_start_time');
    const elapsed = startTime ? Math.floor((Date.now() - parseInt(startTime)) / 1000) : 0;

    
    const textMessage =  `⚠️ Task already running: ${currentTask}\n` +
      `Started: ${elapsed}s ago\n` +
      `Please wait for completion and try again later.`;

    return textMessage;
  } 
  
  props.setProperty('task_status', 'RUNNING');
  props.setProperty('task_name', taskName);
  props.setProperty('task_function', functionName);
  props.setProperty('task_start_time', Date.now().toString());
  props.deleteProperty('task_error');
  props.deleteProperty('task_end_time');
  
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === functionName) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  Logger.log("Creating trigger for: " + functionName);

  const trigger = ScriptApp.newTrigger(functionName)
    .timeBased()
    .after(1000)
    .create();

  Logger.log("Trigger created: " + trigger.getUniqueId());
    
  const textMessage = `✅ ${taskName} queued!\n` +
    `The task is now queueing in the background.\n` +
    `To check progress, click the button again.\n` +
    `Note: Long date range may take 10-20 minutes.`;

  return textMessage;
}

// -------------------- Short Task Runner --------------------
function runShortTask(functionName, taskName) {
  const props = PropertiesService.getScriptProperties();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Decide sheet based on task
  let sheetName = "Suspicious_Player";

  const sheet = ss.getSheetByName(sheetName);

  const statusCell = sheet.getRange("C6");

  if (props.getProperty('task_status') === 'RUNNING') {
  
    const currentTask = props.getProperty('task_name') || 'Unknown Task';
    const startTime = props.getProperty('task_start_time');
    const elapsed = startTime 
      ? Math.floor((Date.now() - parseInt(startTime)) / 1000)
      : 0;

    statusCell.setValue(
    `⚠️ Currently Running: ${currentTask} (${elapsed}s) Please wait until it finish and try again later\n Checked time: ${new Date().toLocaleString()}`
    );

    return;
  }

  props.setProperty('task_status', 'RUNNING');
  props.setProperty('task_name', taskName);
  props.setProperty('task_start_time', Date.now().toString());

  // statusCell.setValue(`⏳ ${taskName} is running...`);

  try {
    this[functionName]();
    markTaskComplete();
  } catch (e) {
    markTaskError(e);
    statusCell.setValue(`❌ ${taskName} failed.\n${e.message}`);
  }
}


// -------------------- Task Helpers --------------------
function markTaskStart(taskName) {
  const props = PropertiesService.getScriptProperties();
  props.setProperty('task_status', 'RUNNING');
  props.setProperty('task_name', taskName);
  props.setProperty('task_start_time', Date.now().toString());
  props.deleteProperty('task_error');
  props.deleteProperty('task_end_time');
  Logger.log("🚀 Task started: " + taskName);
}

function markTaskComplete() {
  const props = PropertiesService.getScriptProperties();
  props.setProperty('task_status', 'COMPLETED');
  props.setProperty('task_end_time', Date.now().toString());
  Logger.log("✅ Task marked as complete");
}

function markTaskError(error) {
  const props = PropertiesService.getScriptProperties();
  props.setProperty('task_status', 'ERROR');
  props.setProperty('task_error', error.toString());
  props.setProperty('task_end_time', Date.now().toString());
  Logger.log("❌ Task marked as error: " + error.toString());
}

function resetTaskStatus() {
  const props = PropertiesService.getScriptProperties();
  props.setProperty('task_status', 'IDLE');
  props.deleteProperty('task_error');
  props.deleteProperty('task_end_time');
  Logger.log("✅ Status reset to IDLE");
}