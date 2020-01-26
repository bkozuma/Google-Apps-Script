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
// Version: 0.2
// Date: 2020/01/25
// - Removed End Date as I do multiday things as single day events
// - Added ability to create all day events
//
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


  // Get the associated calendar
  var eventCal = CalendarApp.getDefaultCalendar();


  // Columns for spreadsheet data
  var cTitleCol = 1;
  var title = '';
  var cStartDateCol = 2;
  var startDate = '';
  var cStartTimeCol = 3;
  var startTime = '';
  var cEndTimeCol = 4;
  var endTime = '';
  var cAllDayCol = 5;
  var allDay = '';
  var cDescriptionCol = 6;
  var description = '';


  // Track number of events
  var numEventsCreated = 0;


  // Header row
  const cHeaderRow = 1;
  var currentRow = cHeaderRow + 1;
  var eventOptions = '';
  var event = [];


  // Get event data for a row
  var range = 'R' +  currentRow + 'C' + cTitleCol + ':' + 'R' + currentRow + 'C' + cDescriptionCol;
  var eventData = cSheet.getRange(range).getValues();
  title = eventData[0][cTitleCol - 1];
  allDay = eventData[0][cAllDayCol - 1];
  while ('' != title) {

    // Get rest of data for the event
    // getValues returns a two dimentional array of the data in row column format
    // Since we're only getting one row at a time, the row is always [0]
    if ('' != allDay) {
      // Create an all day event
      // https://developers.google.com/apps-script/reference/calendar/calendar#createalldayeventtitle,-date,-options
      startTime = eventData[0][cStartDateCol - 1];
      description = eventData[0][cDescriptionCol - 1];
      eventOptions = {
        'description': description,
        'sendInvites': 'False'
      }
      event = eventCal.createAllDayEvent(String(title), new Date(startTime), eventOptions);

    } else {
      // Create a timed event
      // https://developers.google.com/apps-script/reference/calendar/calendar#createeventtitle,-starttime,-endtime,-options
      startTime = constructDateTime(eventData[0][cStartDateCol - 1], eventData[0][cStartTimeCol - 1]);
      endTime = constructDateTime(eventData[0][cStartDateCol - 1], eventData[0][cEndTimeCol - 1]);
      description = eventData[0][cDescriptionCol - 1];
      eventOptions = {
        'description': description,
        'sendInvites': 'False'
      }
      event = eventCal.createEvent(String(title), new Date(startTime), new Date(endTime), eventOptions);

    } // Get rest of data for the event


    // Increment number of events created
    ++numEventsCreated;

 
    // Prepare for the next row
    ++currentRow;
    range = 'R' +  currentRow + 'C' + cTitleCol + ':' + 'R' + currentRow + 'C' + cEndTimeCol;
    eventData = cSheet.getRange(range).getValues();
    title = eventData[0][cTitleCol - 1];
    allDay = eventData[0][cAllDayCol - 1];
    
  } // Get event data for a row

  
  cUI.alert('# events created: ' + numEventsCreated);
  return;
  
} // addEventsToCalendar
