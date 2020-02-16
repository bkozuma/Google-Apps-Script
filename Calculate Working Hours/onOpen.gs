/**
* Adds a menu to the 2020 Ginkgo Hours Tracking Google Sheet
*
* @param {} No input parameters
* @return {} No return code
*/
//
// Name: onOpen
// Author: Bruce Kozuma
//
//
// Version: 0.1
// Date: 2020/02/15
// - Initial release
//
//
function onOpen() {

  // Create custom menus
  const cUI = SpreadsheetApp.getUi();
  cUI.createMenu('Hours tracking')
      .addItem('Update hours', 'getWorkingTimeFromCalendar')
      .addToUi();

} // onOpen
