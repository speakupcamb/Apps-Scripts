'use strict';

/**
 * Fetches tables from Airtable, and sheets from Google
 * Sheets. Collects the Airtable data and writes the result to a
 * Google Sheets sheet.
 */
function Reports() {
  
  var airtable = new Airtable();
  var sheets = new Sheets();

  airtable.fetchTables();
  sheets.getSheets();

  airtable.collectData();
  sheets.writeDataToSheet(airtable.personData, sheets.personSheet);
  sheets.writeDataToSheet(airtable.meetingData, sheets.meetingSheet);
  sheets.writeDataToSheet(airtable.activityData, sheets.activitySheet);
}
