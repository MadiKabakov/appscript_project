function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Task Manager')
      .addItem('Add Task', 'showAddTaskDialog')
      .addItem('Update Task', 'showUpdateTaskDialog')
      .addItem('Complete Task', 'completeTask')
      .addToUi();
  }
  
  function showAddTaskDialog() {
    var html = HtmlService.createHtmlOutputFromFile('AddTask')
      .setWidth(400)
      .setHeight(300);
    SpreadsheetApp.getUi().showModalDialog(html, 'Add New Task');
  }
  
  function addTask(task, status, startDate, endDate) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var lastRow = sheet.getLastRow();
    var taskID = lastRow + 1;
    sheet.appendRow([taskID, task, status, startDate, endDate]);
  }
  
  function completeTask() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var range = sheet.getDataRange();
    var values = range.getValues();
    
    var ui = SpreadsheetApp.getUi();
    var response = ui.prompt('Enter task ID to mark as complete:');
    var taskID = response.getResponseText();
    
    for (var i = 1; i < values.length; i++) {
      if (values[i][0] == taskID) {
        sheet.getRange(i + 1, 3).setValue('Completed');
        break;
      }
    }
  }
  