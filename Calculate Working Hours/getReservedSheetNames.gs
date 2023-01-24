/**
* Gets the names of reserved sheets from the 2023 Ginkgo Hours Tracking Google Sheet
*
* @param {pSpreadsheetApp, pValuesSheetName} Spreadsheet app, name of the sheet with values
* @return {reservedSheetNames}
*/
//
// Name: getReservedSheetNames
// Author: Bruce Kozuma
//
//
// Version: 0.1
// Date: 2023/01/24
// - Initial release
//
//
function getReservedSheetNames(pSpreadsheetApp, pValuesSheetName) {
  // Get reserved sheetnames
  let cValuesSheet = pSpreadsheetApp.getSheetByName(pValuesSheetName);
  let cReservedSheetStartRow = 3;
  let cReservedSheetStartCol = 5;
  let reservedSheetIndex = 0;
  let reservedSheetNames = new Array();
  let reservedSheetRange = 'R' +  cReservedSheetStartRow + 'C' + cReservedSheetStartCol;
  let reservedSheetName = cValuesSheet.getRange(reservedSheetRange).getValue();
  while ('' != reservedSheetName) {

    // The cell isn't empty, so write reserved sheet names to array
    reservedSheetNames[reservedSheetIndex] = reservedSheetName;


    // Prepare for next loop
    reservedSheetIndex++;
    reservedSheetRange = 'R' +  (cReservedSheetStartRow + reservedSheetIndex) + 'C' + cReservedSheetStartCol;
    reservedSheetName = cValuesSheet.getRange(reservedSheetRange).getValue();

  } // Get reserved sheetnames

  return reservedSheetNames;

} // function getReservedSheetNames
