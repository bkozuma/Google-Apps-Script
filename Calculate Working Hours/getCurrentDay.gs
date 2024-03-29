// Copyright 2020, 2021, 2022 Ginkgo Bioworks
// SPDX-License-Identifier: BSD-3-Clause
// https://opensource.org/licenses/BSD-3-Clause

/**
* Gathers work time from Google Calendar and puts into the a Google Sheet
* @param {pWeekStartingDate, pDayOffset} pWeekStartingDate = Date of the start of the week. pOffset = offset from the given date
* @return {newDate} The date
*/
//
// Name: getCurrentDay
// Author: Bruce Kozuma
//
// Version: 0.2
// Date: 2022/03/07
// - Updated copyright
// - Set last day of the month to 28 in non-leap year Februarys
//
//
// Version: 0.1
// Date: 2021/01/03
// - Account for current day in a new year
//
//
// Date: 2020/03/29
// - Initial release
//
// To do
//
//
function getCurrentDay(pWeekStartingDate, pDayOffset) {
  // Month is January, March, May, July, August, October, December)
  const cLongMonths = [1, 3, 5, 7, 8, 10, 12];


  // Get passed date components
  let year = pWeekStartingDate.getFullYear();
  let currentMonth = pWeekStartingDate.getMonth(); // returns 0 - 11
  let monthOffset = 1;
  let currentDate = pWeekStartingDate.getDate();
  let currentDay = 0;


  // Get the last day of the month
  let lastDayOfMonth = 30;
  if (cLongMonths.includes(currentMonth + 1)) {
    // Month has 31 days
    lastDayOfMonth = 31;

  } else if (1 == currentMonth ) {
    // February, so account for leap year
    // https://www.timeanddate.com/date/leapyear.html
    lastDayOfMonth = 28;
    let isLeapYear = false;
    if ((0 == year % 4) && ((0 == year % 100) || (0 == year % 400))) {
      // Yes, it's a leap year
      lastDayOfMonth = 29;
      
    } // Is the current year a leap year
    
  } // Get last day of the month
    

  // Does the pWeekStartingDate plus day offset mean the next date is in a new month
  if (currentDate + pDayOffset > lastDayOfMonth) {
    // Advance the month as well as the day
    currentMonth += 2;


    // Is currentMonth in a new year, i.e., greater than 12?
    if (12 < currentMonth) {
      // In a new year, so set month to 1 and increment year
      currentMonth = 1;
      ++year;

    } // Is currentMonth in a new year, i.e., greater than 12?

    // Next date is in a new month, but not a new year
    currentDate += pDayOffset - lastDayOfMonth;

  } else {
    // Nope, so add the currentDate and pDayOffset
    currentMonth += 1;
    currentDate += pDayOffset;

  } // Does the pWeekStartingDate plus day offset mean the next date is in a new month

  
  // Set the current day and return it
  currentDay = year + '/' + currentMonth + '/' + currentDate;
  return currentDay;

} // function getCurrentDay
