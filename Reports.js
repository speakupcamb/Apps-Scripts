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
  airtable.createMaps();
  sheets.getSheets();

  airtable.collectData();
  sheets.writeDataToSheet(airtable.personData, sheets.personSheet);
  sheets.writeDataToSheet(airtable.meetingData, sheets.meetingSheet);
  sheets.writeDataToSheet(airtable.activityData, sheets.activitySheet);
  sheets.writeDataToSheet(airtable.duesData, sheets.duesSheet);
  sheets.findAbsenteeMembers(sheets.personSheet, sheets.notAttendedSheet);
  
  //Closing ceremonies
  sheets.analyticsSheet.getRange(2, 1).setValue("Last Refresh: " + getTimestamp());
  
  //*** Ray, is there a better way to do this in a loop? ***
  //Resize columns and return cursor to first cell for all sheets
  sheets.personSheet.sort(1, true);
  sheets.personSheet.autoResizeColumns(1, sheets.personSheet.getLastColumn());
  sheets.personSheet.getRange(2, 1).activate();
  sheets.meetingSheet.sort(1, true);
  sheets.meetingSheet.autoResizeColumns(1, sheets.meetingSheet.getLastColumn());
  sheets.meetingSheet.getRange(2, 1).activate();
  sheets.meetingSheet.sort(1, false);
  sheets.activitySheet.autoResizeColumns(1, sheets.activitySheet.getLastColumn());
  sheets.activitySheet.getRange(2, 1).activate();
  sheets.activitySheet.sort(5, true);
  sheets.notAttendedSheet.autoResizeColumns(1, sheets.notAttendedSheet.getLastColumn());
  sheets.notAttendedSheet.getRange(2, 1).activate();
  sheets.duesSheet.sort(1, false);
  sheets.duesSheet.autoResizeColumns(1, sheets.duesSheet.getLastColumn());
  sheets.duesSheet.getRange(2, 1).activate();
  sheets.analyticsSheet.autoResizeColumns(1, sheets.analyticsSheet.getLastColumn());
  sheets.analyticsSheet.getRange(2, 1).activate();
  
  Browser.msgBox("Refresh complete!");
}
