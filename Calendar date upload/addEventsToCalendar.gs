// Copyright 2020 Ginkgo Bioworks
// SPDX-License-Identifier: BSD-3-Clause
// https://opensource.org/licenses/BSD-3-Clause

/**
* Add events to a Google Calendar from the attached Google Sheet
*
* Referenece: https://towardsdatascience.com/creating-calendar-events-using-google-sheets-data-with-appscript-203b26446ce9
*
* @param {} No input parameters
* @return {} No return code
*/
//
// Name: addEventsToCalendar
// Author: Bruce Kozuma
//
// Version: 0.1
// Date: 2020/01/24
// - Initial release
//
//
function addEventsToCalendar()
{
  // Just avoiding silly pitfalls
  "use strict";


  // Function name
  const cScriptName = 'AddEventsToCalendar';


  // The Google Sheet and UI
  const cSs = SpreadsheetApp.getActiveSpreadsheet();
  const cSheet = cSs.getActiveSheet();
  const cUI = SpreadsheetApp.getUi();


  // Get email of user invoking this script
  // Production
  const cUserEmail = Session.getActiveUser().getEmail();
  // Debug
//  const cUserEmail = 'd' + Session.getActiveUser().getEmail();


  // Get the associated calendar
  var eventCal = CalendarApp.getDefaultCalendar();


  // ColumnscreateEvent(title, startTime, endTime, options)  // Columns for spreadsheet data
  var cTitleCol = 1;
  var title = '';
  var cStartDateCol = 2;
  var startDate = '';
  var cStartTimeCol = 3;
  var startTime = '';
  var cEndDateCol = 4;
  var endDate = '';
  var cEndTimeCol = 5;
  var endTime = '';
  var cDescriptionCol = 6;
  var description = '';


  // Header row
  const cHeaderRow = 1;
  var currentRow = cHeaderRow + 1;
  var eventData = '';
  var eventOptions = '';
  var event = [];


  // Get event data for a row
  var range = 'R' +  currentRow + 'C' + cTitleCol + ':' + 'R' + currentRow + 'C' + cEndTimeCol;
  eventData = cSheet.getRange(range).getValues();
  title = eventData[0][cTitleCol - 1];
  while ('' != title) {

    // Get rest of data for the event
    // getValues returns a two dimentional array of the data in row co9lumn format
    // Since we're only getting one row at a time, the row is always [0]
    startTime = constructDateTime(eventData[0][cStartDateCol - 1], eventData[0][cStartTimeCol - 1]);
    endTime = constructDateTime(eventData[0][cEndDateCol - 1], eventData[0][cEndTimeCol - 1]);
    description = eventData[0][cDescriptionCol - 1];
    eventOptions = {
      'description': description,
      'sendInvites': 'False'
    }
    event = eventCal.createEvent(String(title), new Date(startTime), new Date(endTime), eventOptions);

    cUI.alert('Created event: ' + event.getTitle() + ' ' + event.getStartTime() + ' ' + event.getEndTime());

 
    // Prepare for the next row
    ++currentRow;
    range = 'R' +  currentRow + 'C' + cTitleCol + ':' + 'R' + currentRow + 'C' + cEndTimeCol;
    eventData = cSheet.getRange(range).getValues();
    title = eventData[0][cTitleCol - 1];
    
  } // Get event data for a row

  return;
  
} // addEventsToCalendar
