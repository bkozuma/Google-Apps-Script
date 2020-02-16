// Copyright 2020 Ginkgo Bioworks
// SPDX-License-Identifier: BSD-3-Clause
// https://opensource.org/licenses/BSD-3-Clause

/**
* Gathers work time from Google Calendar and puts into the a Google Sheet
*
* Referenece: https://towardsdatascience.com/creating-calendar-events-using-google-sheets-data-with-appscript-203b26446ce9
*
* @param {} No input parameters
* @return {} No return code
*/
//
// Name: getWorkingTimeFromCalendar
// Author: Bruce Kozuma
//
// Version: 0.1
// Date: 2020/02/15
// - Initial release
//
// To do
// Get start and end times
// Do duration calculations
// Deposit into Google Sheet
//
//
function getWorkingTimeFromCalendar()
{
  // Just avoiding silly pitfalls
  "use strict";


  // Function name
  const cScriptName = 'getWorkingTimeFromCalendar';


  // The Google Sheet and UI
  const cSs = SpreadsheetApp.getActiveSpreadsheet();
  const cSheet = cSs.getActiveSheet();
  const cUI = SpreadsheetApp.getUi();


  // Get the associated calendar
  let eventCal = CalendarApp.getDefaultCalendar();


  // Columns for spreadsheet data
  const cWeekStartingCol = 1;
  let weekStarting = '';
  const cWeekEndingCol = 2;
  let weekEnding = '';
  const cMonCol = 3;
  const cTueCol = 4;
  const cWedCol = 5;
  const cThuCol = 6;
  const cFriCol = 7;
  const cSatCol = 8;
  const cSunCol = 9;
  const cNumDaysInWeek = cSunCol - cMonCol + 1;
  const cWeekendDay = [cSatCol, cSunCol];

  // This offset needs to be 3 to account for the difference between the day offset
  // and the position of the Saturday column 
  const cWeekendDayOffset = 3;


  // End of working day in miliseconds from 00:00
  const cMilisecondsPerHour = 1000*60*60;
  const cWorkingDayBegins = '07:30';
  const cWorkingDayEnds = '17:00';


  // Track number of weeks handled
  let numWeeks = 0;


  // Header row
  const cHeaderRow = 1;
  let currentRow = cHeaderRow + 1;


  // Get event data for a row while Week starting isn't blank
  // getValues returns a two dimentional array of the data in row column format
  // Since we're only getting one row at a time, the row is always [0]
  let range = 'R' +  currentRow + 'C' + cWeekStartingCol + ':' + 'R' + currentRow + 'C' + cSunCol;
  let eventData = cSheet.getRange(range).getValues();
  weekStarting = eventData[0][cWeekStartingCol - 1];
  weekEnding = eventData[0][cWeekEndingCol - 1];
  let dayEvents = {};
  let numDayEvents = 0;
  let eventIndex = 0;
//  let eventTitle = '';
//  let eventStartTime = '';
//  let eventEndTime = '';
//  let workingTime = 0;
//  let nonWorkingTime = 0;
  const cNonWorkingTime = [ 'OoO', 'Out', 'Transit'];
  const cDefaultEarliestStartTime = new Date('12/31/3000');
//  let earliestEventStartTime = cDefaultEarliestStartTime;
  const cDefaultLatestEndTime = new Date('01/01/1970');
//  let latestEventEndTime = cDefaultLatestEndTime;
  let dayOffset = 0;
  let currentDay = '';
  const cNotFound = -1;
//  let afterHoursTime = 0;
//  let dayWorkingTime = 0;
//  let beginWorkingDay = '';
//  let endWorkingDay = '';
  let isAllDayEvent = false;
  let wasEventAccepted = false;
//
  let eventTitle = '';
  let eventStartTime = '';
  let eventEndTime = '';
  let earliestEventStartTime = cDefaultEarliestStartTime;
  let latestEventEndTime = cDefaultLatestEndTime;
  let beginWorkingDay = '';
  let endWorkingDay = '';
  let workingTime = 0;
  let nonWorkingTime = 0;
  let afterHoursTime = 0;
  let dayWorkingTime = 0;
//
  while ('' != weekStarting) {

    // Get data for Monday (same as week starting) then loop through days of week
    while (dayOffset < cNumDaysInWeek) {

      // If cell is empty, calculate working time
      range = 'R' + currentRow + 'C' + (cMonCol + dayOffset);
      if ('' == cSheet.getRange(range).getValue()) {
        // Cell is empty, so write result


        // Get events for the day
        currentDay = weekStarting.getFullYear() + '/' + (weekStarting.getMonth() + 1) + '/' + (weekStarting.getDate() + dayOffset);
        dayEvents = eventCal.getEventsForDay(new Date(currentDay));
        beginWorkingDay = new Date(currentDay + ' ' + cWorkingDayBegins);
        endWorkingDay = new Date(currentDay + ' ' + cWorkingDayEnds);


        // Loop through day events
        numDayEvents = dayEvents.length;
        while (numDayEvents > eventIndex) {
          // Algorithm is to calculate working hours as follows:
          // - Note: Javascript dates are calculated as miliseconds from January 1, 1970
          // - Days that have values are ignored
          // - Usual working hours for a normal working day:
          //   - Earliest event start time is 07:30
          //   - Latest event end time is 17:00
          // - Working hours are calculated as follows:
          //   - Take the end time of the last appointment in the day
          //   - Subtract the start time from the earliest appointment in the day
          //   - This approach accounts for gaps between events
          // - Events before or after usual working hours
          //   - Working time is calculated as event end time - event start time
          //   - A running total of such working time per day is kept and added to normal working time
          // - All day events are ignored
          // - Non-accepted events are ignored
          // - Non-working time is ignored (e.g., Transit, Out; there's a list)
          // - Saturday and Sunday are not usual working days (there's a list)
          // - Edge cases not handled by this code
          //   - Working time on holidays
          //   - Time spent working before or after normal working hours not in a gap between events


          // Is the event an All Day event?
          isAllDayEvent = dayEvents[eventIndex].isAllDayEvent();


          // Was the event accepted?
          wasEventAccepted = dayEvents[eventIndex].getMyStatus();


          // Skip the event if it is an all day event or was not accepted
          // https://stackoverflow.com/questions/44102840/google-apps-script-for-calendar-searchfilter
          if ((!isAllDayEvent) && (CalendarApp.GuestStatus.Yes != wasEventAccepted)) {

            // Get the title of the event, and the start and end time
            eventTitle = dayEvents[eventIndex].getTitle();
            eventStartTime = dayEvents[eventIndex].getStartTime();
            eventEndTime = dayEvents[eventIndex].getEndTime();


            // Earlier event?
            if (eventStartTime < beginWorkingDay) {
              // Start of event is before beginning of working day, so don't reset earliest appointment,
              // but add event duration to tally
              afterHoursTime += eventEndTime - eventStartTime;

            } else if (earliestEventStartTime > eventStartTime) {
              // Yup, current event's start time is before the earliest
              earliestEventStartTime = eventStartTime;

            } // Earlier event?


            // Later end time?
            if (eventEndTime > endWorkingDay) {
              // End of event is after end of working day, so don't reset latest appointment,
              // but add event duration to tally
              afterHoursTime += eventEndTime - eventStartTime;

            } else if (latestEventEndTime < eventEndTime) {
              // Yup, current event's end time is after the current latest and this is not an event
              // after normal working hours
              latestEventEndTime = eventEndTime;

            } // Earlier event?


            // Is the event a non-working event?
            if (cNotFound != cNonWorkingTime.indexOf(eventTitle)) {
              // Event is a non-working event, so subtract time from the daily running total
              nonWorkingTime += Math.abs(eventEndTime.getTime() - eventStartTime.getTime()) / cMilisecondsPerHour;

            } // Is the evnet a non-working event?

          } // Skip the event if it is an all day event


          // Check next event
          ++eventIndex;

        } // Loop through events


        // Check for case where latest event date/time end wasn't set
        if (cDefaultLatestEndTime == latestEventEndTime) {
          // Probably a weekend event
          // Set start time to earliest start time of the day
          latestEventEndTime = eventEndTime;

        } // Check for case where earliest event date/time isn't set


        // Calculate working hours
        // https://stackoverflow.com/questions/19225414/how-to-get-the-hours-difference-between-two-date-objects
        if (0 != numDayEvents) {
          // There were events during a day

          // Is this a weekend day?
// HERE
          if (cNotFound == cWeekendDay.indexOf(dayOffset + cWeekendDayOffset)) {
            // Not a weekend day, so working time includes after hours time
            workingTime = (latestEventEndTime - earliestEventStartTime + afterHoursTime) / cMilisecondsPerHour;

          } else {
            // It's a weekend day, so working time is only after hours time
            workingTime = afterHoursTime / cMilisecondsPerHour;

          } // Is this a weekend day?


          // Account for any non-working time
          workingTime -= nonWorkingTime;

        } else {
          // Empty day
          workingTime = 0;

        } // Calculate working hours


        // Write result to sheet
        cSheet.getRange(range).setValue(workingTime);

      } // If cell is empty, calculate working time


      // Prepare for next day
      earliestEventStartTime = cDefaultEarliestStartTime;
      latestEventEndTime = cDefaultLatestEndTime;
      workingTime = 0;
      nonWorkingTime = 0;
      eventIndex = 0;
      afterHoursTime = 0;
      ++dayOffset;

    } // Loop through days of week


    // Prepare for the next row
    ++numWeeks;
    ++currentRow;
    range = 'R' +  currentRow + 'C' + cWeekStartingCol + ':' + 'R' + currentRow + 'C' + cSunCol;
    eventData = cSheet.getRange(range).getValues();
    weekStarting = eventData[0][cWeekStartingCol - 1];
    weekEnding = eventData[0][cWeekEndingCol - 1];
    dayOffset = 0;

  } // Get event data for a row


  // Finish up
  cUI.alert(cScriptName, 'Processed weeks: ' + numWeeks, cUI.ButtonSet.OK);
  return;

} // getWorkingTimeFromCalendar
