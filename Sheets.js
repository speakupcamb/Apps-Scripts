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
