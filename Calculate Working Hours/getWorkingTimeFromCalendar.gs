// Copyright 2020, 2021, 2022, 2023 Ginkgo Bioworks
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
*   - Default earliest event start time is 07:15
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
*   - If the start of the non-working event is before the start of the work day,
      and the end of the non-working event is after the start of the work day,
      set the start of the work day to the end of the non-working event,
      e.g., I slept in
*   - If the non-working event is after the end of the working day, nothing needs to be done as
      the event won't be included in the calculation for working time
*   - For each working event, if the end of the event is after the stop time, event time is only day stop time - event start time
*   - Working hours are calculated by day stop time - day start time - non-working time +
*   - Work time outside of working hours is calculated by duration of the event and added to the total working time
*   - Day working time is calculated by working hours + work time outside of working hours
* - Edge cases not handled by this code
*   - Working time on holidays
*   - Time spent working before or after normal working hours not in a gap between events
* - Now finds time for Project Diamond
*
* Reference: https://towardsdatascience.com/creating-calendar-events-using-google-sheets-data-with-appscript-203b26446ce9
*
* @param {} No input parameters
* @return {} No return code
*
*
To do:
// - To do: Handle events for which I am the owner and then I turn down
//          https://stackoverflow.com/questions/41504049/how-to-test-if-the-owner-of-a-calendar-event-declined-it
// - To do: Handle overlapping working events
*
*/
//
// Name: getWorkingTimeFromCalendar
// Author: Bruce Kozuma
//
// Version: 0.27
// Date: 2023/01/24
// - Revamp how P # Project time is calculated
// - Time on P # Projects is calculated solely by event time, NOT by general working time
// - Removing Start and Stop event handling
//
//
// Version: 0.26
// Date: 2022/12/31
// - Updated copyright for 2023
//
//
// Version: 0.25
// Date: 2022/12/02
// - Fix bug where Sunday time calculated as negative
//
//
// Version: 0.24
// Date: 2022/10/23
// - Account for events that span days
// - Adjust for P # projects
//
//
// Version: 0.23
// Date: 2022/08/12
// - Set the precision of Project Diamond time to 2 digits
//
//
// Version: 0.22
// Date: 2022/08/03
// - Included P-Diamond as a search term for P317
//
//
// Version: 0.21
// Date: 2022/06/21
// - Account for addition of tags column per row
// - Corrected returned precision to 2 digits, was 1 actually
//
//
// Version: 0.20
// Date: 2022/05/31
// - Included Bayer as a search term for P317
//
//
// Version: 0.19
// Date: 2022/05/31
// - Write the accumulated weekly Project Diamond time to a column to ensure accurate total working hours
//
//
// Version: 0.18
// Date: 2022/05/31
// - Accounted for having to track time on Project Diamond, i.e., P317
// - Set the returned precision to 2 digits
//
//
// Version: 0.17
// Date: 2022/03/07
// - Updated copyright
//
//
// Version: 0.16
// Date: 2021/10/23
// - Added 'Transit to Drydock' and 'Transit to Wilson Rd' as non-working events
//
//
// Version: 0.15
// Date: 2021/10/05
// - Added ability to count tentatively accepted events
//
//
// Version: 0.14
// Date: 2021/07/03
// - Adjusted how Out events are handled, they can now just start with Out,
//   the event title does not have to only be "Out"
//
//
// Version: 0.13
// Date: 2021/03/06
// - Updated the copyright
// - Addressed case where non-working event starts before end of normal working day
//   and ends at end of normal working day
//
//
// Version: 0.12
// Date: 2020/09/17
// - Added ability to handle events that start before the start of the work day,
//   and the end after the start of the work day, e.g., I slept in
//
//
// Version: 0.11
// Date: 2020/08/29
// - Removed additional code related to P141
//
//
// Version: 0.10
// Date: 2020/05/25
// - Removed code related to P141
// - Handled work on holidays or days off
//
//
// Version: 0.9
// Date: 2020/05/22
// - Cleared beforeWorkTime
//
//
// Version: 0.8
// Date: 2020/05/06
// - Handled case where start event occurs after beginning of the work day has been modified
//
//
// Version: 0.7
// Date: 2020/04/08
// - Added case where an event starts at the default beginning of the work day
// - Added case where an event starts at the default ending of the work day
//
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
  let eventTitle = '';


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
  const cSatCol = 10;
  const cSunCol = 11;
  const cNumDaysInWeek = cSunCol - cMonCol + 1;
  const cWeekendDay = [cSatCol, cSunCol];
  let dayOffset = 0;
  let currentDay = '';


  // This offset needs to account for the difference between the day offset
  // and the position of the Saturday column
  const cWeekendDayOffset = 5;


  // End of working day in miliseconds from 00:00
  // Get from Google Calendar? When there is an API call for it
  const cMilisecondsPerHour = 1000*60*60;
  const cWorkingDayBegins = '07:15';
  const cWorkingDayEnds = '17:00';


  // Track number of weeks handled
  let numWeeksProcessed = 0;


  // Header row
  const cHeaderRow = 2;
  let currentRow = cHeaderRow + 1;


  // Get event data for a row while Week starting isn't blank
  // getValues returns a two dimentional array of the data in row column format
  // Since we're only getting one row at a time, the row is always [0]
  let range = 'R' +  currentRow + 'C' + cWeekStartingCol + ':' + 'R' + currentRow + 'C' + cSunCol;
  let eventData = cSheet.getRange(range).getValues();
  weekStarting = eventData[0][cWeekStartingCol - 1];
  weekEnding = eventData[0][cWeekEndingCol - 1];


  // Set up a bunch of other variables
  let dayEvents = {};
  let numDayEvents = 0;
  let beginWorkingDay = 0;
  let defaultBeginWorkingDay = 0;
  let endWorkingDay = 0;
  let beginCurrentDay = new Date();
  let endCurrentDay = new Date();
  let eventIndex = 0;
//  let eventTitle = '';
  let eventMidnight = new Date();
  let eventYear = new Date();
  let eventMonth = new Date();
  let eventDay = new Date();
  let eventStartTime = new Date();
  let eventEndTime = new Date();
  let workingTime = 0;
  let nonWorkingTime = 0;
  let beforeHoursTime = 0;
  let afterHoursTime = 0;


  // Potentially get non-working time values from a separate Sheet?
  const cOut = 'Out';
  const cNonWorkingTime = [ 'OoO', cOut, 'Out of office', 'Transit', 'Transit to Drydock', 'Transit to Wilson Rd'];
  const cNotFound = -1;
  let isAllDayEvent = false;
  let wasEventAccepted = false;
  let weekendWorkingTime = 0;
  let isWeekendDay = false;
  let isHoliday = false;
  const cHoliday = 'Holiday';


  // Get current sheet name
  let sheetName = '';
  sheetName = cSheet.getSheetName();


  // Get reserved sheetnames
  const cValuesSheetName = 'Values';
  let reservedSheetNames = new Array();
  reservedSheetNames = getReservedSheetNames(cSs, cValuesSheetName);


  // Check if active sheet is a reserved sheet
  let reservedSheetName = '';
  for (reservedSheetName of reservedSheetNames) {

    // Is the active sheet a reserved sheet?
    if (sheetName == reservedSheetName) {
      // Yup, active sheet is a reserved one, warn and exit
      cUI.alert('Sheet ' + sheetName + ' is not intended for this function; function will terminate.',cUI.ButtonSet.OK)
      return;

    } // Is the active sheet a reserved sheet?

  } // Check if active sheet is a reserved sheet


  // Get P values
  // P Number and P Code are required
  // P Other is optional
  let PValues = new Array ();
  PValues = getPValues(cSs, cValuesSheetName);
  let currentSheetIsPValueSheet = false;
  let PValueFound = false;


  // Is current sheet a P Value sheet?
  if (true == searchPValues(PValues, sheetName, sheetName)) {
    // Yup, on a P Value sheet
    currentSheetIsPValueSheet = true;

  } // Is current sheet a P Value sheet?
  let currentSheetPValueDayHours = 0;


  // Process weeks
  while ('' != weekStarting) {

    // Has week been submitted
    hasBeenSubmitted = cSheet.getRange('R' + currentRow + 'C' + cSubmittedCol).getValue();
    if ('' == hasBeenSubmitted) {
      // Week has NOT been submitted
      ++numWeeksProcessed;

    } else {
      // Week has been submitted, so skip week
      hasBeenSubmitted = '';
      ++currentRow;
      range = 'R' +  currentRow + 'C' + cWeekStartingCol + ':' + 'R' + currentRow + 'C' + cSunCol;
      eventData = cSheet.getRange(range).getValues();
      weekStarting = eventData[0][cWeekStartingCol - 1];
      weekEnding = eventData[0][cWeekEndingCol - 1];
      dayOffset = 0;
      currentSheetPValueDayHours = 0;
      continue;

    } // Has week been submitted


    // Get data for Monday (same as week starting) then loop through days of week
    while (dayOffset < cNumDaysInWeek) {

      // If cell is empty, calculate working time
      range = 'R' + currentRow + 'C' + (cMonCol + dayOffset);
      if (true == cSheet.getRange(range).isBlank()) {
        // Cell is empty, so calculate day's hours


        // Set the current day
        currentDay = getCurrentDay(weekStarting, dayOffset);


        // Set default beginning and ending of the day
        beginWorkingDay = new Date(currentDay + ' ' + cWorkingDayBegins);
        defaultBeginWorkingDay = new Date(beginWorkingDay.getTime());
        beginCurrentDay = new Date(beginWorkingDay.getTime());
        beginCurrentDay.setHours(0);
        beginCurrentDay.setMinutes(0);
        beginCurrentDay.setSeconds(0);

        endWorkingDay = new Date(currentDay + ' ' + cWorkingDayEnds);
        endCurrentDay = new Date(endWorkingDay.getTime());
        endCurrentDay.setHours(23);
        endCurrentDay.setMinutes(59);
        endCurrentDay.setSeconds(59);


        // Get the events for the day
        dayEvents = eventCal.getEventsForDay(new Date(currentDay));


        // Loop through day events
        numDayEvents = dayEvents.length;
        while (numDayEvents > eventIndex) {
          // Get the title of the event, and the start and end time
          eventTitle = dayEvents[eventIndex].getTitle();
          eventStartTime = new Date(dayEvents[eventIndex].getStartTime());
          eventEndTime = new Date(dayEvents[eventIndex].getEndTime());


          // Is the event an All Day event?
          isAllDayEvent = dayEvents[eventIndex].isAllDayEvent();


          // Is it a holiday or work on a day off?
          if ((cNotFound != eventTitle.indexOf(cHoliday)) ||
              (isAllDayEvent && (true == eventTitle.startsWith(cOut)))) {
            isHoliday = true;

          } // Is it a holiday or work on a day off?


          // Skip if event is an All Day event
          if (true == isAllDayEvent) {
            // Yup, all day event, skip it
            ++eventIndex;
            continue;

          } // Skip if event is an All Day event


          // Was the event accepted?
          wasEventAccepted = dayEvents[eventIndex].getMyStatus();


          // Was event accepted as a Yes or a Maybe and I don't own the meeting
          // https://stackoverflow.com/questions/44102840/google-apps-script-for-calendar-searchfilter
          // https://developers.google.com/apps-script/reference/calendar/guest-status
            if ((CalendarApp.GuestStatus.NO == wasEventAccepted) ||
              (CalendarApp.GuestStatus.INVITED == wasEventAccepted)) { 
              // Event can be skipped since it wasn't accepted as Yes or Maybe
              // Events to which I've been invited but to which I've declined are listed as OWNER
              ++eventIndex;
              continue;

          } // Was event accepted as a Yes or a Maybe and I don't own the meeting


          // Check for modifications to the start or end of the day
          // Handle event that started the previous day
          eventYear = eventEndTime.getFullYear();
          eventMonth = eventEndTime.getMonth();
          eventDay = eventEndTime.getDate();
          eventMidnight.setFullYear(eventYear);
          eventMidnight.setMonth(eventMonth);
          eventMidnight.setDate(eventDay);
          eventMidnight.setHours(0);
          eventMidnight.setMinutes(0);
          eventMidnight.setSeconds(0);
          if (eventStartTime < beginCurrentDay) {
            // Event starts on the previous day, so set event start time to
            // beginning of current day so we don't over count hours
            eventStartTime = beginCurrentDay;
            beginWorkingDay = eventStartTime;

          } // Handle event that started the previous day


          // Handle event that ends after the current day
          if (eventEndTime > endCurrentDay) {
            // Event ends after the current day, so set event end time to
            // midnight of the current day so we don't over count hours
            eventEndTime = endCurrentDay;

          } // Handle event that ends after the current day


          // Is event related to a P value, e.g., P317, Diamond, Bayer?
          //* @return {-1 == pItemToFind not found
          // 0 == pItemToFind found and not for pSheetName,
          // 1 == if pItemToFind found and for pSheetName}
          PValueFound = searchPValues(PValues, eventTitle, sheetName);
          switch(PValueFound) {
            case cNotFound:
              // Event title doesn't match a P value, so proceed
              break;

            case 0:
              // Yup, found a P value in eventTitle and not the same as the current sheet name,


              // Does the event span start of the working day
              if (eventEndTime < beginWorkingDay) {
                // Non-working event occurs before the start of the working day,
                // Do nothing

              // Does event spans normal beginning of the work day?
              } else if ((eventStartTime < beginWorkingDay) && (eventEndTime > beginWorkingDay)) {
                // Event spans beginning of the working day
                // Set start of working day to end of event
                beginWorkingDay == eventEndTime;

              // Does the event span end of the working day
              } else if ((eventStartTime < endWorkingDay) && (eventEndTime > endWorkingDay)) {
                // Event spans end of the working day, so set end of the working day to
                // beginning of event
                endWorkingDay == eventStartTime;

              // Does the event start at the end of the working day?
              } else if (eventStartTime - endWorkingDay == 0) {
                // Event starts at end of working day
                // Do nothing

              // Does event start after the working day
              } else if (eventStartTime > endWorkingDay) {
                // Event starts after the working day
                // Do nothing

              } else {
                // Event occurs during the working day so treat it like non-working time
                nonWorkingTime += Math.abs(eventEndTime.getTime() - eventStartTime.getTime()) / cMilisecondsPerHour;

              } // Is it after the end of the working day?

              ++eventIndex;
              continue;

            case 1:
              // Found a pValue same as the current sheet
              currentSheetPValueDayHours += Math.abs((eventEndTime - eventStartTime)) / cMilisecondsPerHour;
              ++eventIndex;
              continue;

          } // Is event related to a P value, e.g., P317, Diamond, Bayer?


          // Is current day a weekend day?
          if (cNotFound != cWeekendDay.indexOf(dayOffset + cWeekendDayOffset)) {
            // It's a weekend day
            isWeekendDay = true;

          } // Is current day a weekend day?


          // Is event on a weekend day or a holiday?
          if ((true == isWeekendDay) || (true == isHoliday)) {
            // Event is on a weekend day or a holiday


            // Is it an out event on a non-working day?
            if (false == eventTitle.startsWith(cOut)) {
              // An out event on a non-working day, so working time is just the time of the appointments
              // Out events are ignored on weekend days, holidays
              // BUG: Double counts overlapping working events on non-work days
              weekendWorkingTime += Math.abs((eventEndTime - eventStartTime)) / cMilisecondsPerHour;


            // Is the event a non-working event?
            } else if (cNotFound != cNonWorkingTime.indexOf(eventTitle)) {
              // Event is a non-working event and not a P Value for the current sheet
              // May subtract time from the daily running total

              // Is the end of the non-working event before the normal start of the work day?
              // This can happen when I commute to work and don't have meetings until Stand up,
              // so I want time from the end of my commute included in the working hours
              if (eventEndTime < beginWorkingDay) {
                // Non-working event occurs before the start of the working day,
                // so reset beginning of working day to end of the event
                beginWorkingDay = eventEndTime;

              } // Is the end of the non-working event before the normal start of the work day?


              // Does non-working event spans normal beginning of the work day?
              if ((eventStartTime < beginWorkingDay) && (eventEndTime > beginWorkingDay)) {
                // Non-working time starts before the start of the working day
                // and ends after the start of the working day (like I slept in)
                // so reset beginning of working day to end of the event
                beginWorkingDay = eventEndTime;

              } // Does non-working event spans normal beginning of the work day?


              // Non-working event during the working day?
              if ((eventStartTime < endWorkingDay) && (eventEndTime < endWorkingDay)) {
                // Non-working event start time falls within the working day (e.g., time for a personal event),
                // so count the event duration as non-working time
                nonWorkingTime += Math.abs(eventEndTime.getTime() - eventStartTime.getTime()) / cMilisecondsPerHour;

              } // Non-working event during the working day?


              // Non-working event is at or spans the end of the working day?
              if ((eventStartTime < endWorkingDay) && (eventEndTime >= endWorkingDay)) {
                // Non-working event begins before the end of the working day and the event
                // ends after the working day ends, then reset the end of the working day
                // to the beginning of the event
                // Events found after this one will be treated like working events after the
                // end of the normal working day
                endWorkingDay = eventStartTime;

              } // Non-working event is at or spans the end of the working day?


              // If the non-working event is after the end of the working day, nothing needs to be done as
              // the event won't be included in the calculation for working time

            } // Is the event a non-working event and not a P Value event not for this sheet?


          // Not a non-working day, i.e., a work day
          // Event ends before the working day?
          } else if (eventEndTime < beginWorkingDay) {
            // End of the event is before beginning of working day,
            // so add event duration to after hours time
            beforeHoursTime += Math.abs((eventEndTime - eventStartTime)) / cMilisecondsPerHour;


          // Event ends at the beginning of the working day?
          // https://stackoverflow.com/questions/492994/compare-two-dates-with-javascript
          } else if (0 == eventEndTime - beginWorkingDay) {
            // Start and end of the event is before beginning of working day,
            // so reset beginning of working day to start of event
            beginWorkingDay = eventStartTime - 1;


          // Event spans beginning of working day
          } else if ((eventStartTime < beginWorkingDay) && (eventEndTime > beginWorkingDay)) {
            // Event starts before the beginning of the working day AND
            // end of the event is after the beginning of the working day
            // so set the beginning of the working day to the start of the event
            beginWorkingDay = eventStartTime;


            // Event occurs after the beginning of working day and before the end of the working day
            // Nothing needs to be done as the will be included in the calculation for working time


          // Event starts at the beginning working day
          // https://stackoverflow.com/questions/492994/compare-two-dates-with-javascript
          } else if (eventStartTime - beginWorkingDay == 0) {
            // Event starts at the beginning of the working day, either the default or
            // another appointment with the same start time, so capture event and set the default
            // working day to some nonsense value, which prevents the working day being set too late
            // Useful in cases where I start work at the default time
            defaultBeginWorkingDay = -1;


          // Event starts after the beginning of the workday and the default hasn't changed
          // https://stackoverflow.com/questions/492994/compare-two-dates-with-javascript
          } else if ((eventStartTime > beginWorkingDay) && (0 == beginWorkingDay - defaultBeginWorkingDay)) {
              // Start of event is after the default start of the working day
              // and the default start of the day hasn't been changed
              // (i.e., there was another event even earlier than this one)
              // so reset start of working day to beginning of the event
              // i.e., I started work after the default time
              beginWorkingDay = eventStartTime;


          // Event spans end of the working day
          } else if ((eventStartTime < endWorkingDay) && (eventEndTime > endWorkingDay)) {
            // Start of event is before the end of the working day and end of event is after
            // end of working day, so reset end of working day to end of event
            endWorkingDay = eventEndTime;


          // Event starts at end the working day
          // https://stackoverflow.com/questions/492994/compare-two-dates-with-javascript
          } else if (eventStartTime - endWorkingDay == 0) {
            // Start of event is at the end of the working day
            // so reset end of working day to end of event
            endWorkingDay = eventEndTime;


          // Event starts after the end of the working day
          } else if (eventStartTime > endWorkingDay) {
            // Start of event is after the end of the working day so
            // so add event duration to after hours time
            afterHoursTime += Math.abs((eventEndTime - eventStartTime)) / cMilisecondsPerHour;

          } // Event starts after the end of the working day


          // Is it an out event not on a weekend or holiday?
          if ((true == eventTitle.startsWith(cOut)) &&
            ((true != isWeekendDay) || (true != isHoliday))) {
            // It's an out event not on a weekend or holiday
            nonWorkingTime += Math.abs((eventEndTime - eventStartTime)) / cMilisecondsPerHour;

          } // Is it an out event not on a weekend or holiday?


          // Check next event
          ++eventIndex;

        } // Is event on a weekend day, a holiday, or a for a P Value that's not for the current sheet?


        // Calculate working hours
        // https://stackoverflow.com/questions/19225414/how-to-get-the-hours-difference-between-two-date-objects
        if (0 != numDayEvents) {
          // There were events during a day


          // Is this a weekend day or a holiday?
          if ((cNotFound != cWeekendDay.indexOf(dayOffset + cWeekendDayOffset)) || (true == isHoliday)) {
            // It's a weekend day, so working time is only after hours time
            workingTime = weekendWorkingTime;

          } else {
            // Not a weekend day, so working time includes after hours time
            workingTime = (Math.abs((endWorkingDay - beginWorkingDay)) / cMilisecondsPerHour) + beforeHoursTime + afterHoursTime;


            // Account for any non-working time, which includes P Value time not for this sheet
            workingTime -= nonWorkingTime;

          } // Is this a weekend day or a holiday?


          // Is current sheet a P Value sheet?
          if (true == currentSheetIsPValueSheet) {
            // Yup, so working time is P Value time for the sheet
            // Don't care what the current value of workingTime is since this is a P Value sheet
            workingTime = currentSheetPValueDayHours;

          } // Is current sheet a P Value sheet?

        } else {
          // Empty day
          workingTime = 0;

        } // Calculate working hours


        // Write result to sheet
        // Set the precision to 2 digits
        cSheet.getRange(range).setValue(workingTime.toPrecision(3));

      } // If cell is empty, calculate working time


      // Prepare for next day
      currentSheetPValueDayHours = 0;
      currentSheetPValueTime = 0;
      workingTime = 0;
      nonWorkingTime = 0;
      eventIndex = 0;
      beforeHoursTime = 0;
      afterHoursTime = 0;
      weekendWorkingTime = 0;
      isWeekendDay = false;
      isHoliday = false;
      ++dayOffset;

    } // Get data for Monday (same as week starting) then loop through days of week


    // Prepare for the next row
    ++currentRow;
    range = 'R' +  currentRow + 'C' + cWeekStartingCol + ':' + 'R' + currentRow + 'C' + cSunCol;
    eventData = cSheet.getRange(range).getValues();
    weekStarting = eventData[0][cWeekStartingCol - 1];
    weekEnding = eventData[0][cWeekEndingCol - 1];
    dayOffset = 0;

  } // Process weeks


  // Finish up
  cUI.alert(cScriptName, 'Processed weeks: ' + numWeeksProcessed, cUI.ButtonSet.OK);
  return;

} // getWorkingTimeFromCalendar
