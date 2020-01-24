// Copyright 2020 Ginkgo Bioworks
// SPDX-License-Identifier: BSD-3-Clause
// https://opensource.org/licenses/BSD-3-Clause

/**
* Converts a passed date and time into a Date object
*
* Referenece: https://stackoverflow.com/questions/25174219/convert-a-string-to-date-in-google-apps-script
*
* @param pDate Date in YYYY/MM/DD format
* @param pTime Time (in HH:MM format)
* @return New Date object
* @return Any error messages, blank if successful
*/
//
// Name: constructDateTime
// Author: Bruce Kozuma
//
// Version: 0.1
// Date: 2020/01/24
// - Initial release
//
//
// Things to implement
// - Timezone
//
function constructDateTime(pDate, pTime) {
  // Just avoiding silly pitfalls
  "use strict";


  // Function name
  const cScriptName = 'constructDateTime';


  // Things to return
  var returnDate = {};
  var returnMessage = '';
  const cERR_PASSED_DATE_EMPTY = 'Passed date is empty';
  const cERR_PASSED_TIME_EMPTY = 'Passed time is empty';

  
  // Date, time variables
  var year = '';
  var month = '';
  var day = '';
  var hour = '';
  var minute = '';


  // Verify passed date isn't empty
  if ('' == pDate) {
    returnDate = {};
    returnMessage = cERR_PASSED_DATE_EMPTY;
    return [returnDate, returnMessage];

  } // Verify passed date isn't empty
  

  // Verify passed time isn't empty
  if ('' == pTime) {
    returnDate = {};
    returnMessage = cERR_PASSED_TIME_EMPTY;
    return [returnDate, returnMessage];

  } // Verify passed time isn't empty
  
 
  // Parse the date, expecting YYYY/MM/DD format
  // Leading + converts the string to numbers
  // Subtract 1 from month since month enumeration starts at 0
  year = +pDate.substring(0,4);
  month = +pDate.substring(5,7) - 1;
  day = +pDate.substring(8,10);

  hour = +pTime.substring(0,2);
  minute = +pTime.substring(3,5);
                          
  returnDate = new Date(year, month, day, hour, minute);
  return [returnDate, returnMessage];
 
} // constructDateTime
