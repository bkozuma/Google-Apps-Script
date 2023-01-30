/**
* Gets the default working times from the 2023 Ginkgo Hours Tracking Google Sheet
*
* @param {pSpreadsheetApp, pValuesSheetName} Spreadsheet app, name of the sheet with values
* @return {defaultTimes [defaultStartTime, defaultStopTime]}
*/
//
// Name: getDefaultWorkingTime
// Author: Bruce Kozuma
//
//
// Version: 0.1
// Date: 2023/01/28
// - Initial release
//
//
function getDefaultWorkingTime(pSpreadsheetApp, pValuesSheetName) {
  // Get default start and stop times
  let cValuesSheet = pSpreadsheetApp.getSheetByName(pValuesSheetName);
  let cDefaultStartTimeRow = 3;
  let cDefaultStopTimeRow = 4;
  let cDefaultTimeCol = 12;
  let range = 'R' +  cDefaultStartTimeRow + 'C' + cDefaultTimeCol;
  let defaultStartTime = cValuesSheet.getRange(range).getValue().getHours();
  let minutes = cValuesSheet.getRange(range).getValue().getMinutes();
  // Is minute a single digit?
  if (10 > minutes)
  {
    // Yes, so add a leading zero
    defaultStartTime += ':0' + minutes;

  } else {
    // Longer than one digit so just add it
    defaultStartTime += ':' + minutes;
    
  } // Is minute a single digit?
  
  range = 'R' +  cDefaultStopTimeRow + 'C' + cDefaultTimeCol;
  let defaultStopTime = cValuesSheet.getRange(range).getValue().getHours();
  minutes = cValuesSheet.getRange(range).getValue().getMinutes();
  // Is minute a single digit?
  if (10 > minutes)
  {
    // Yes, so add a leading zero
    defaultStopTime += ':0' + minutes;

  } else {
    // Longer than one digit so just add it
    defaultStopTime += ':' + minutes;
    
  } // Is minute a single digit?

  let defaultTimes = [defaultStartTime, defaultStopTime];
  return defaultTimes;

} // function getDefaultWorkingTime
