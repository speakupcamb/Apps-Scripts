'use strict';

/**
 * Constructs a Sheets object, and initializes its properties.
 */
function Sheets() {
  
  Logger.log('Initializing Sheets');
  
  // Sheets tabs of interest
  this.personSheet = null;
  this.meetingSheet = null;
  this.activitySheet = null;
  this.analyticsSheet = null;
  this.notAttendedSheet = null;
  this.duesSheet = null;
}

/**
 * Write data containing a header and rows of values to a sheet.
 *
 * @param data {Object} object with properties header, the header
 * values array, and rows, an array of equal length row values arrays
 * @param sheet {Object} Google Sheets object
 *
 * @return {Undefined}
 */
Sheets.prototype.writeDataToSheet = function(data, sheet) {

  var range;

  sheet.clear();

  // Write and format the header
  range = sheet.getRange(1, 1, 1, data.header.length);
  range.setValues([data.header]);
  range.setFontWeight("bold");
  
  // Write the rows
  range = sheet.getRange(sheet.getLastRow() + 1, 1, data.rows.length, data.rows[0].length);
  range.setValues(data.rows);
};


/**
 * Get all required sheets using the Google Sheets API, if getting of
 * any of the required sheets has not been completed.
 *
 * @return {Undefined}
 */
Sheets.prototype.getSheets = function() {
  if (!this.isGetSheetsComplete()) {
    this.personSheet = this.getSheetByName('Persons');
    this.meetingSheet = this.getSheetByName('Meetings');
    this.activitySheet = this.getSheetByName('Activities');
    this.analyticsSheet = this.getSheetByName('Analytics');
    this.duesSheet = this.getSheetByName('Dues');
    this.notAttendedSheet = this.getSheetByName('Members Not Attended > Month');
  }
};

/**
 * Indicate if the get for all required sheets has been completed, or
 * not.
 *
 * @return {Boolean} true if complete
 */
Sheets.prototype.isGetSheetsComplete = function () {
  return (null !== this.personSheet &&
          null !== this.meetingSheet &&
          null !== this.analyticsSheet &&
          null !== this.notAttendedSheet &&
          null !== this.duesSheet &&
          null !== this.activitySheet);
};

/**
 * Get the sheet corresponding to the specified name.
 *
 * @param name {String} the name of the sheet
 *
 * @return a Google Sheets object
 */
Sheets.prototype.getSheetByName = function(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
};
