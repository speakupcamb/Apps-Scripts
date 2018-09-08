'use strict';

function Sheets() {

  Logger.log('Initializing Sheets');

  this.personSheet = null;
  this.analyticsSheet = null;
  this.meetingAttendanceSheet = null;
  this.meetingSpeechesSheet = null;
  this.meetingRantsSheet = null;
  this.duesSheet = null;
  this.memberNotAttendedSheet = null;
  this.meetingCountSheet = null;

}


var personSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Members/Guests');
var analyticsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Analytics');
var meetingAttendanceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Meetings-Attendance');
var meetingSpeechesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Meetings-Speeches');
var meetingRantsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Meetings-Rants');
var duesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Dues');
var memberNotAttendedSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Members Not Attended > Month');
var meetingCountSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Meeting Counts');

Sheets.prototype.getSheets = function() {
  if (!this.isGetSheetsComplete()) {
  this.personSheet = this.getSheetByName('Members/Guests');
  this.analyticsSheet = this.getSheetByName('Analytics');
  this.meetingAttendanceSheet = this.getSheetByName('Meetings-Attendance');
  this.meetingSpeechesSheet = this.getSheetByName('Meetings-Speeches');
  this.meetingRantsSheet = this.getSheetByName('Meetings-Rants');
  this.duesSheet = this.getSheetByName('Dues');
  this.memberNotAttendedSheet = this.getSheetByName('Members Not Attended > Month');
  this.meetingCountSheet = this.getSheetByName('Meeting Counts');
  }
};

Sheets.prototype.isGetSheetsComplete = function () {
  return (null === this.personSheet &&
          null === this.analyticsSheet &&
          null === this.meetingAttendanceSheet &&
          null === this.meetingSpeechesSheet &&
          null === this.meetingRantsSheet &&
          null === this.duesSheet &&
          null === this.memberNotAttendedSheet &&
          null === this.meetingCountSheet);
};

Sheets.prototype.getSheetByName = function(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
};

function Test() {

  // https://sites.google.com/site/scriptsexamples/custom-methods/gsunit

  var sheets = new Sheets();

  Logger.log('Test getSheetByName()');

  personSheet = sheets.getSheetByName('Members/Guests');

  GSUnit.assertArrayEquals('Header of Guests and Members sheet',
                           ["Name", "First Name", "Last Name", "Type", "Status", "Email Address"],
                           personSheet.getRange(1, 1, 1, 6).getDisplayValues()[0]);
}
