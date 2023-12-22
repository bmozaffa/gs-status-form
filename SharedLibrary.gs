function isPaused(trigger) {
  let sheet = SpreadsheetApp.openById('1vrAk1qITGj72Vm4yVcsYpu3rHWZhALgA0uW7tCiX1sc').getSheetByName('main');
  for (let row = 2; row <= sheet.getLastRow(); row++) {
    let name = sheet.getRange(row, 1, 1, 1).getValue();
    if (name === trigger) {
      let now = new Date();
      let start = sheet.getRange(row, 2, 1, 1).getValue();
      if (start && start > now) {
        return false;
      }
      let end = sheet.getRange(row, 3, 1, 1).getValue();
      if (end && end < now) {
        return false;
      }
      //Script is paused
      console.log("The trigger called %s is paused with a start date of %s and end date of %s", trigger, start, end);
      return true;
    } else if (!name) {
      return false;
    }
  }
}
