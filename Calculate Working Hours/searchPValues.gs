/**
* Searches the P values from the 2025 Ginkgo Hours Tracking Google Sheet
*
* @param {pValues, pItemToFind, pSheetName} Array of P Values, item to be found, sheet name
* @return {-1 == pItemToFind not found in P Value list, 0 == pItemToFind found and not for pSheetName, 1 == if pItemToFind found and for pSheetName}
* PValues[PValueIndex] = [ PNumber, PCodename, POther ];
*/
//
// Name: searchPValues
// Author: Bruce Kozuma
//
//
// Version: 0.3
// Date: 2025/01/21
// - Re-wrote for clarity of understanding
//
//
// Version: 0.2
// Date: 2024/01/08
// - Fixed comment for clarity
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
  const cItemFoundNotForPSheetName = 0;
  const cItemFoundForPSheetName = 1;


  ////////////////////////////////////////////////////////////////////
  // Cases to check for item to find
  // * Case 1: item not found in P Value list
  // * Case 2: item found in the P Value list
  //   - Case 2a: Not for passed sheet name
  //   - Case 2b: For passed sheet name
  ////////////////////////////////////////////////////////////////////
  // loop for sets of P Values (outer array acting as container)
  let numPValueSets = pPValues.length;
  for(var outer = 0; outer < numPValueSets; outer++){

    // Is item in this set of P Values
    let numPValues = pPValues[outer].length;
    for(var inner = 0; inner < numPValues; inner++){

      // item to find in the P Value list, i.e., Case 2?
      // BUG: Handle case where pValue is blank
      let pValue = pPValues[outer][inner];
      let pItemToFindFound = pItemToFind.indexOf(pValue);
      if ( cNotFound != pItemToFindFound) {
        // pItemToFind was found in PValue list

        // Is item to find for the passed sheet name
        // i.e., Case 2b
        let pItemSheetName = pPValues[outer][0];
        if (pSheetName == pItemSheetName) {
          // Yup, item to find is for passed sheet name
            return cItemFoundForPSheetName;

        } else {
          // item found in P Value list but not for the passed sheet name
          // i.e., Case 2a
          return cItemFoundNotForPSheetName;

        } // item to find find in the P Value list?

      } // item to find in the P Value list, i.e., Case 2?

    } // Is item in this set of P Values

  } // loop for sets of P Values (outer array acting as container)


  // Didn't return within the loops, so item to find not found in P value list
  // i.e., Case 1
  return cNotFound;

} // function searchPValues(pPValues, pItemToFind, pSheetName) 
