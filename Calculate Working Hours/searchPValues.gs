/**
* Searches the P values from the 2023 Ginkgo Hours Tracking Google Sheet
*
* @param {pValues, pItemToFind, pSheetName} Array of P Values, item to be found, sheet name
* @return {-1 == pItemToFind not found,0 == pItemToFind found and not for pSheetName, 1 == if pItemToFind found and for pSheetName}
* PValues[PValueIndex] = [ PNumber, PCodename, POther ];
*/
//
// Name: searchPValues
// Author: Bruce Kozuma
//
//
// Version: 0.1
// Date: 2023/01/24
// - Initial release
//
//
function searchPValues(pPValues, pItemToFind, pSheetName) {
  // Search for pItemToFind
  const cNotFound = -1;
  const cFoundNotForPSheetName = 0;
  const cItemToFindIsPSheetName = 1;


  // loop for sets of P Values (outer array acting as container)
  for(var outer = 0; outer < pPValues.length; outer++){

    // Is the passed sheet name different from the P Number (first element)?
    if (cNotFound != pItemToFind.indexOf(pPValues[outer][0])) {
      // PName is same as the sheet name, so indicate so
        return cItemToFindIsPSheetName;

    } // Is the passed sheet name different from the P Number (first element)?


    // Check other items for a match to a P Value
    // Don't have to check P Number since we already checked that
    for(var inner = 1; inner < pPValues[outer].length; inner++){

      // pItemToFind is as a P value?
      if ( cNotFound != pItemToFind.indexOf(pPValues[outer][inner])) {
        // pItemToFind is found as a P value
        // Is pSheetName the P Number?
        if (pSheetName == pPValues[outer][0]) {
          // Yup, pSheetName is the P Number
            return cItemToFindIsPSheetName;

        } // // Is pItemToFind for the pSheetName, i.e., is a psuedonym?

        // pItemToFind is not for the pSheetName
        return cFoundNotForPSheetName;

      } // pItemToFind is as a P value?

    } // Check other items for a match to a P Value

  } // loop for sets of P Values (outer array acting as container)


  // Didn't return within the loops, so not found
  return cNotFound;

} // function searchPValues(pPValues, pItemToFind, pSheetName)
