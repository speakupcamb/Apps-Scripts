# Apps-Scripts
These are, well, Apps Script scripts for our Toastmasters club. They are organized as follows:
* Airtable.js -- Fetches all required Airtable tables, and collects all data required for reporting
* Sheets.js -- Gets all required Google Sheets sheets, and writes all data to sheets required for reporting
* Reports.js -- Creates an Airtable and Sheets object in order to fetch data, get sheets, collect data, and write the collected data to sheets
* Test.js -- Uses GSUnit to implement tests of Airtable and Sheets functions

Note that in this organization:
* All data manipulation resides in the Airtable object
* Airtable data is left in a form that is very close to that returned by the API
* All data presentation resides in the Sheets object
* Google Sheets is used only for presenting the data collected by the Airtable object
* No formulas are created in the Google Sheets sheets
* All report creation resides in the Reports object
