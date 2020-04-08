// Copyright 2020 Ginkgo Bioworks
// SPDX-License-Identifier: BSD-3-Clause
// https://opensource.org/licenses/BSD-3-Clause

/**
* Gathers work time from Google Calendar and puts into the a Google Sheet
* Note: Javascript dates are calculated as miliseconds from January 1, 1970
* Algorithm is to calculate working hours as follows:
* - Weeks that have been submitted are ignored
* - Days that have values are ignored
* - Events that are ignored
*   - All day events
*   - Non-accepted events
*   - Non-working time (e.g., Transit, Out; there's a list)
* - Usual working hours for a normal working day:
*   - Default earliest event start time is 07:30
*   - Default latest event end time is 17:00
* - Events before or after usual working hours
*   - Working time is calculated as event end time - event start time
*   - A running total of such working time per day is kept and added to normal working time
* - Saturday and Sunday are not usual working days (there's a list)
*   - Working time calculated like events before or after usual working hours
* - Algorithm (accounts for gaps between events)
*   - Scan the day's events for Start and Stop events
*   - If found, set bounds of working day appropriately
*   - If not found, use defaults
*   - Assumption is that a day's events are sorted by time
*   - For each non-working event, accumulate that time if the event is during working hours
*   - If the end of the non-working event is before the start of the work day,
      set the start of the work day to the end of the non-working event,
      e.g., a Transit calendar entry where I arrive before 07:30
*   - If the non-working event is after the end of the working day, nothing needs to be done as
      the event won't be included in the calculation for working time
*   - For each working event, if the end of the event is after the stop time, event time is only day stop time - event start time
*   - Working hours are calculated by day stop time - day start time - non-working time +
*   - Work time outside of working hours is calculated by duration of the event and added to the total working time
*   - Day working time is calculated by working hours + work time outside of working hours
* - Edge cases not handled by this code
*   - Working time on holidays
*   - Time spent working before or after normal working hours not in a gap between events
*
* Reference: https://towardsdatascience.com/creating-calendar-events-using-google-sheets-data-with-appscript-203b26446ce9
*
* @param {} No input parameters
* @return {} No return code
*/
//
// Name: getWorkingTimeFromCalendar
// Author: Bruce Kozuma
//
// Version: 0.6
// Date: 2020/04/08
// - Re-wrote again to have more explicit rules
//
//
// Version: 0.5
// Date: 2020/04/01
// - Accounted for start of work day after default
//
//
// Version: 0.4
// Date: 2020/03/29
// - New algorithm added
// - Accounted for getCurrentDay
//
//
// Version: 0.3
// Date: 2020/02/20
// - Added 'Out of Office' to the list of time out of office
//
//
// Version: 0.2
// Date: 2020/02/16
// - Adjust for new column postions
// - Weeks that have been submitted are ignored
// - Typos in comments fixed
//
//
// Version: 0.1
// Date: 2020/02/15
// - Initial release
//
// To do
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
  const cSubmittedCol = 3;
  let hasBeenSubmitted = '';
  const cMonCol = 5;
  const cTueCol = 6;
  const cWedCol = 7;
  const cThuCol = 8;
  const cFriCol = 9;
  const cSatCol = 10;
  const cSunCol = 11;
  const cNumDaysInWeek = cSunCol - cMonCol + 1;
  const cWeekendDay = [cSatCol, cSunCol];
  let dayOffset = 0;
  let currentDay = '';


  // This offset needs to be 3 to account for the difference between the day offset
  // and the position of the Saturday column
  const cWeekendDayOffset = 5;


  // End of working day in miliseconds from 00:00
  const cMilisecondsPerHour = 1000*60*60;
  const cWorkingDayBegins = '07:15';
  const cWorkingDayEnds = '17:00';


  // Track number of weeks handled
  let numWeeksProcessed = 0;


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
  let beginWorkingDay = 0;
  let defaultBeginWorkingDay = 0;
  let endWorkingDay = 0;
  let eventIndex = 0;
  let eventTitle = '';
  let eventStartTime = 0;
  let eventEndTime = 0;
  let workingTime = 0;
  let nonWorkingTime = 0;
  let beforeHoursTime = 0;
  let afterHoursTime = 0;
  // Potentially get non-working time values from a separate Sheet?
  const cNonWorkingTime = [ 'OoO', 'Out', 'Out of office', 'Transit'];
  const cStartEvent = 'Start';
  const cStopEvent = 'Stop';
  const cNotFound = -1;
  let isAllDayEvent = false;
  let wasEventAccepted = false;
  let weekendWorkingTime = 0;
  while ('' != weekStarting) {

    // Has week been submitted
    hasBeenSubmitted = cSheet.getRange('R' + currentRow + 'C' + cSubmittedCol).getValue();
    if ('' == hasBeenSubmitted) {
      // Week has NOT been submitted
      ++numWeeksProcessed;

    } // Has week been submitted


    // Get data for Monday (same as week starting) then loop through days of week
    while ((dayOffset < cNumDaysInWeek) && ('' == hasBeenSubmitted)) {

      // If cell is empty, calculate working time
      range = 'R' + currentRow + 'C' + (cMonCol + dayOffset);
      if ('' == cSheet.getRange(range).getValue()) {
        // Cell is empty, so calculate day's hours


        // Set the current day
        currentDay = getCurrentDay(weekStarting, dayOffset);


        // Set default beginning and ending of the day
        beginWorkingDay = new Date(currentDay + ' ' + cWorkingDayBegins);
        defaultBeginWorkingDay = beginWorkingDay;
        endWorkingDay = new Date(currentDay + ' ' + cWorkingDayEnds);


        // Get the events for the day
        dayEvents = eventCal.getEventsForDay(new Date(currentDay));


        // Loop through day events
        numDayEvents = dayEvents.length;
        while (numDayEvents > eventIndex) {
          // Is the event an All Day event?
          isAllDayEvent = dayEvents[eventIndex].isAllDayEvent();


          // Was the event accepted?
          wasEventAccepted = dayEvents[eventIndex].getMyStatus();


          // Skip the event if it is an all day event or was not accepted
          // https://stackoverflow.com/questions/44102840/google-apps-script-for-calendar-searchfilter
          if ((!isAllDayEvent) &&
              ((CalendarApp.GuestStatus.YES == wasEventAccepted) || (CalendarApp.GuestStatus.OWNER == wasEventAccepted))) {


            // Get the title of the event, and the start and end time
            eventTitle = dayEvents[eventIndex].getTitle();
            eventStartTime = dayEvents[eventIndex].getStartTime();
            eventEndTime = dayEvents[eventIndex].getEndTime();


            // Is the event a Start event?
            // If it is, use the start time of the event to set the start of the working day
            // If no Start event found, default start of working day time is used
            if (cStartEvent == eventTitle) {
              beginWorkingDay = eventStartTime;


            // Is the event a Stop event?
            // If it is, use the start time of the event to set the end of the working day
            // If no Stop event found, the default end of working day time is used
            } else if (cStopEvent == eventTitle) {
              endWorkingDay = eventStartTime;


              // Also blank out after hours time since we set a new end of the working day
              afterHoursTime = 0;


            // Is event on a weekend day?
            // Assumption is that Start and Stop events don't occur on weekends
            } else if (cNotFound != cWeekendDay.indexOf(dayOffset + cWeekendDayOffset)) {
              // For weekend days, working time is just the time of the appointments
              weekendWorkingTime += (eventEndTime - eventStartTime) / cMilisecondsPerHour;


            // Is the event a non-working event?
            } else if (cNotFound != cNonWorkingTime.indexOf(eventTitle)) {
              // Event is a non-working event, so subtract time from the daily running total

              // Is the end of the non-working event before the normal start of the work day?
              // This can happen when I commute to work and don't have meetings until Stand up,
              // so I want time from the end of my commute included in the working hours
              if (eventEndTime < beginWorkingDay) {
                // Non-working event occurs before the start of the working day,
                // so reset beginning of working day to end of the event
                beginWorkingDay = eventEndTime;


              // Non-working event during the working day?
              } else if ((eventStartTime < endWorkingDay) && (eventEndTime < endWorkingDay)) {
                // Non-working event start time falls within the working day (e.g., time for a personal event),
                // so count the event duration as non-working time
                nonWorkingTime += Math.abs(eventEndTime.getTime() - eventStartTime.getTime()) / cMilisecondsPerHour;


              // Non-working event spans the end of the working day
              } else if ((eventStartTime < endWorkingDay) && (eventEndTime > endWorkingDay)) {
               // Non-working event begins before the end of the working day and the event
               // ends after the working day ends, then reset the end of the working day
               // to the beginning of the event
               // Events found after this one will be treated like working events after the
               // end of the normal working day
               endWorkingDay = eventStartTime;

              } // Is the end of the non-working event after the normal start of the work day?


              // If the non-working event is after the end of the working day, nothing needs to be done as
              // the event won't be included in the calculation for working time


              // Not a non-working event, i.e., a working event
              // Event ends before the working day?
            } else if  (eventEndTime < beginWorkingDay) {
              // End of the event is before beginning of working day,
              // so add event duration to after hours time
              beforeHoursTime += (eventEndTime - eventStartTime) / cMilisecondsPerHour;


              // Event ends at the beginning of the working day?
            } else if (eventEndTime == beginWorkingDay) {
              // Start and end of the event is before beginning of working day,
              // so reset beginning of working day to start of event
              beginWorkingDay = eventStartTime;


              // Event spans beginning of working day
            } else if ((eventStartTime < beginWorkingDay) && (eventEndTime > beginWorkingDay)) {
              // Event starts before the beginning of the working day AND
              // end of the event is after the beginning of the working day
              // so set the beginning of the working day to the start of the event
              beginWorkingDay = eventStartTime;


              // Event occurs after the beginning of working day and before the end of the working day
              // Nothing needs to be done as the will be included in the calculation for working time


            // Event starts after working day, starts after the default working day, and the default
            // working day hasn't been changed, i.e., I started work late
            } else if ((eventStartTime > beginWorkingDay) && (beginWorkingDay == defaultBeginWorkingDay)) {
              // Start of event is after the default start of the working day
              // and the default start of the day hasn't been changed
              // so reset start of working day to beginning of the event
              /*
              Case 1: Event after 07:15, default not changed already, what happens? Change default
              Case 2: Third even after first event, default changed already, what happens? Skip
              */
              beginWorkingDay = eventStartTime;


            // Event spans end of the working day
            } else if ((eventStartTime < endWorkingDay) && (eventEndTime > endWorkingDay)) {
              // Start of event is before the end of the working day and end of event is after
              // end of working day, so reset end of working day to end of event
              endWorkingDay = eventEndTime;


            // Event starts at end the working day
            } else if (eventStartTime == endWorkingDay) {
              // Start of event is at the end of the working day
              // so reset end of working day to end of event
              endWorkingDay = eventEndTime;


            // Event starts after the end of the working day
            } else if (eventStartTime > endWorkingDay) {
              // Start of event is after the end of the working day so
              // so add event duration to after hours time
              afterHoursTime += (eventEndTime - eventStartTime) / cMilisecondsPerHour;

            } // Is the event a Start event?

          } // Skip the event if it is an all day event

          // Check next event
//cUI.alert('earliestEventStartTime: ' + earliestEventStartTime + "\n" +
//         'latestEventEndTime: ' + latestEventEndTime);
          ++eventIndex;

        } // Loop through events


        // Calculate working hours
        // https://stackoverflow.com/questions/19225414/how-to-get-the-hours-difference-between-two-date-objects
        if (0 != numDayEvents) {
          // There were events during a day

          // Is this a weekend day?
          if (cNotFound != cWeekendDay.indexOf(dayOffset + cWeekendDayOffset)) {
            // It's a weekend day, so working time is only after hours time
            workingTime = weekendWorkingTime;

          } else {
            // Not a weekend day, so working time includes after hours time
            workingTime = ((endWorkingDay - beginWorkingDay) / cMilisecondsPerHour) + beforeHoursTime + afterHoursTime;

            // Account for any non-working time
            workingTime -= nonWorkingTime;

          } // Is this a weekend day?

        } else {
          // Empty day
          workingTime = 0;

        } // Calculate working hours


        // Write result to sheet
        cSheet.getRange(range).setValue(workingTime);

      } // If cell is empty, calculate working time


      // Prepare for next day
      workingTime = 0;
      nonWorkingTime = 0;
      eventIndex = 0;
      afterHoursTime = 0;
      weekendWorkingTime = 0;
      ++dayOffset;

    } // Loop through days of week


    // Prepare for the next row
    ++currentRow;
    range = 'R' +  currentRow + 'C' + cWeekStartingCol + ':' + 'R' + currentRow + 'C' + cSunCol;
    eventData = cSheet.getRange(range).getValues();
    weekStarting = eventData[0][cWeekStartingCol - 1];
    weekEnding = eventData[0][cWeekEndingCol - 1];
    dayOffset = 0;

  } // Get event data for a row


  // Finish up
  cUI.alert(cScriptName, 'Processed weeks: ' + numWeeksProcessed, cUI.ButtonSet.OK);
  return;

} // getWorkingTimeFromCalendar
