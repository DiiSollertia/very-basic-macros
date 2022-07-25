function CreateTriggers() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var currentTriggers = ScriptApp.getProjectTriggers();

  while (currentTriggers.length > 0) {
    ScriptApp.deleteTrigger(currentTriggers[0])
    Logger.log('Removed old trigger.')
    const currentTriggers = ScriptApp.getProjectTriggers();
  }

  ScriptApp.newTrigger("onEdit").forSpreadsheet(sheet).onEdit().create();
  ScriptApp.newTrigger("onOpen").forSpreadsheet(sheet).onOpen().create();
}