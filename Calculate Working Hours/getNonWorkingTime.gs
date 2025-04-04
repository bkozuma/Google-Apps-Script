/**
* Gets the indicators of non-working time for the 2023 Ginkgo Hours Tracking Google Sheet
*
* @param {pSpreadsheetApp, pValuesSheetName} Spreadsheet app, name of the sheet with values
* @return {nonWorkingIndicators}
*/
//
// Name: getNonWorkingTimeIndicators
// Author: Bruce Kozuma
//
//
// Version: 0.2
// Date: 2023/05/19
// - Change name of file to match name of function
//
//
// Version: 0.1
// Date: 2023/01/24
// - Initial release
//
//
function getNonWorkingTimeIndicators(pSpreadsheetApp, pValuesSheetName) {
  // Get non-working indicators
  let cValuesSheet = pSpreadsheetApp.getSheetByName(pValuesSheetName);
  let cNonWorkingIndicatorsStartRow = 3;
  let cNonWorkingIndicatorStartCol = 7;
  let nonWorkingIndicatorIndex = 0;
  let nonWorkingIndicators = new Array();
  let nonWorkingIndicatorRange = 'R' +  cNonWorkingIndicatorsStartRow + 'C' + cNonWorkingIndicatorStartCol;
  let nonWorkingIndicator = cValuesSheet.getRange(nonWorkingIndicatorRange).getValue();
  while ('' != nonWorkingIndicator) {

    // The cell isn't empty, so write reserved sheet names to array
    nonWorkingIndicators[nonWorkingIndicatorIndex] = nonWorkingIndicator;


    // Prepare for next loop
    nonWorkingIndicatorIndex++;
    nonWorkingIndicatorRange = 'R' +  (cNonWorkingIndicatorsStartRow + nonWorkingIndicatorIndex) + 'C' + cNonWorkingIndicatorStartCol;
    nonWorkingIndicator = cValuesSheet.getRange(nonWorkingIndicatorRange).getValue();

  } // Get non-working indicators

  return nonWorkingIndicators;

} // function getNonWorkingTimeIndicators
