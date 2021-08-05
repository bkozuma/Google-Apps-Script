/**
* This script takes the user inputs from a Google Sheet and creates Issues
* in the specified Jira Project based on the Google Sheet contents
*
* This script is customized for Ginkgo Bioworks
*
* @param none
* @return [0] returnMessage Return message, blank if successful, non-empty otherwise
*
*/
//
// Name: CreateIssue
// Author: Bruce Kozuma
// Credits: Based on code from here: http://www.thedataplumber.net/using-google-forms-and-scripts-to-create-a-jira-issue/
// A really good reference: https://developer.atlassian.com/cloud/jira/platform/rest/v3/intro/
//
// To be implemented
//
//
// Version: 0.1
// Date: 2020/08/04
// - Initial version
//

function CreateIssue() {

  // Just avoiding silly pitfalls
  "use strict";


  // name of the script
  const cScriptName = 'CreateIssue';


  // The Google Sheet and UI
  const cSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const cUI = SpreadsheetApp.getUi();
  let uiResult = "";


  // Header information, including authorization information
  // This API call is linked to an account in Jira, and follows the Basic Authentication method
  // ("username:API token" are Base64 encoded)
  // Note: email:API token!
  // IMPORTANT: Take out Base64 encoded information before checking into GitHub!!!!
  const cHeaders = {
    "content-type": "application/json",
    "Accept": "application/json",
  //"authorization": "Basic <Base64 encoded email:API token>"
  // Use the following to test the authorization string from a command prompt
  // curl --url https://ginkgobioworks.atlassian.net/rest/api/latest/issue/createmeta --header "Content-Type: application/json" --user "bkozuma@ginkgobioworks.com:HxGyRBAhjcSYCq6n8ZV70FA4"

  }; // Header information, including authorization information


  // Items related to the Ginkgo Jira Cloud instance and REST API calls
  const cBaseURL = 'https://ginkgobioworks.atlassian.net/';
  const cAccessURL = '/rest/api/latest/issue/';
  const cBrowseURL = 'browse/';
  let issueURL = cBaseURL + cAccessURL;
  const cProjectKey = 'ITENG';


  // For reading data for the Issue
  let summary = '';
  let description = '';
  let keyContact = '';
  let note = '';


  // For creating the Issue
  let data = '';
  let payload = '';
  let options = '';
  let response = '';
  let returnMessage = '';
  let dataAll = '';
  let issueKey = '';
  let returnCode = '';


  // Google Sheet Header row
  const cHeaderRow = 2;
  let currentRow = cHeaderRow + 1;


  // Google Sheet column layout
  let cSummaryBaseCol = 1;
  let cKeyContactCol = 5;
  let cTicketCol = 7;
  let cNoteCol = 9;
  const cNA = "NA";


  // Standard issue components
  const cIssueTextComponent1 = "Onboard ";
  const cIssueTextComponent2 = " to Okta ";
  const cKeyContactComponent = "R" + cHeaderRow + "C" + cKeyContactCol;


  // Get summary
  summary = cSheet.getRange('R' + currentRow + 'C' + cSummaryBaseCol).getValue();


  // Loop through Sheet
  let range = 'R' +  currentRow + 'C' + cSummaryBaseCol + ':' +
              'R' + currentRow + 'C' + cNoteCol;
  let rowData = cSheet.getRange(range).getValues();
  while ('' != summary) {
    // Skip if ITENG Ticket # is NA or the cell is blank
    issueKey = cSheet.getRange(range).getValues()[0][cTicketCol - 1];
    if ((cNA != issueKey) && ('' == issueKey)) {

      // Assemble summary for the new Issue
      summary = cIssueTextComponent1 + summary + cIssueTextComponent2


      // Assemble description for the new Issue
      // Only add note if needed
      // This might be faster if rowData was used
      keyContact = cSheet.getRange(cKeyContactComponent).getValue() + " is ";
      keyContact += cSheet.getRange('R' +  currentRow + 'C' + cKeyContactCol).getValue();
      note = cSheet.getRange('R' +  currentRow + 'C' + cNoteCol).getValue();
      description = summary + "\n\n" +
                    keyContact;
      if ("" != note) {
        note = "\n\n" + note;
        description += note;

      } // Assemble description for the new Issue


      // data for the new Issue
      data = {
        "fields": {
          "project":{
              "key": cProjectKey

          },
          "summary": summary,
          "description": description,


          // All tickets are categorised as Tasks, and are not changed by the form submission.
          "issuetype":{
              "name": "Task"

          }

        }

      }; // data for the new Issue


      // Turn all the post data into a JSON string to be send to the API
      payload = JSON.stringify(data);


      // Options to complete the JSON string
      // For muteHttpExceptions:
      // - Use true to get less ugly notifications for end users
      // - Use false to get raw reports from Jira
      // For method:
      // - Use GET to read
      // - Use POST to create\write
      options = {
        "content-type": "application/json",
        "method": "POST",
        "headers": cHeaders,
        "payload": payload,
        "muteHttpExceptions": false

      }; // Options to complete the JSON string


      // Make the HTTP call to the JIRA API
      response = UrlFetchApp.fetch(issueURL, options);


      // Check for return code
      returnMessage = '';
      dataAll = '';
      issueKey = '';
      returnCode = response.getResponseCode();
      switch (returnCode) {
        case 201: {
          // Issue creation is successful if 201 is returned
          dataAll = JSON.parse(response.getContentText());
          issueKey = dataAll.key;

          break;

        } // Issue was updated

        default: {
          // There was an Issue with creating the Issue
          // beef this up a lot

          returnMessage += 'There was a problem creating an issue for row ' + currentRow;

          break;

        } // default

      } // Check for return code


      // Update the sheet with the Issue Key and status
      // The real way to set the link for the Issue Key is here: https://hawksey.info/blog/2020/04/everything-a-google-apps-script-developer-wanted-to-know-about-reading-hyperlinks-in-google-sheets-but-was-afraid-to-ask/
      cSheet.getRange('R' +  currentRow + 'C' + cTicketCol).setValue("=hyperlink(" + "\"" + cBaseURL + cBrowseURL + issueKey + "\";\"" + issueKey + "\")");

    } // Skip if ITENG Ticket # is NA


    // Prepare for next row
    keyContact = '';
    note = '';
    description = '';
    ++currentRow;
    summary = cSheet.getRange('R' + currentRow + 'C' + cSummaryBaseCol).getValue();
    range = 'R' +  currentRow + 'C' + cSummaryBaseCol + ':' +
            'R' + currentRow + 'C' + cNoteCol;
    rowData = cSheet.getRange(range).getValues();

  } // Loop through Sheet


  // End
  return [returnMessage];

} // CreateIssue
