'use strict';

function Airtable() {
  
  Logger.log('Initializing Airtable');
  
  this.baseURL = 'https://api.airtable.com/v0/';
  this.db_key = 'apppcMJJYoQygtgfZ';
  this.api_key = 'keyl4k6PFanvcnSPr';
  
  this.persons = null;
  this.types = null;
  this.statuses = null;
  this.meetings = null;
  this.dues = null;
  this.duesStatues = null;
  
  this.typeMap = {};
  this.statusMap = {};
  this.duesStatusMap = {};
  
  this.personFields = [
    'Name',
    'First Name',
    'Last Name',
    'Type',
    'Status',
    'Email Address',
    'Address',
    'Phone'
  ];
  this.meetingAttendanceFields = [
    'Attended',
    'Meeting Date'
  ];
  this.meetingSpeechesFields = [
    'Speech',
    'Meeting Date'
  ];
  this.meetingRantsFields = [
    'Rant',
    'Meeting Date'
  ];
  this.duesFields = [
    'Person',
    'Due Period',
    'Status'
  ];
};

Airtable.prototype.collectData = function() {
  
  for (var iT = 0; iT < this.types.length; iT++) {
    if (!this.typeMap.hasOwnProperty(this.types[iT].id)) {
      this.typeMap[this.types[iT].id] = this.types[iT].fields.Status;
    }
  }

  for (var iS = 0; iS < this.statuses.length; iS++) {
    if (!this.statusMap.hasOwnProperty(this.statuses[iS].id)) {
      this.statusMap[this.statuses[iS].id] = this.statuses[iS].fields.Name;
    }
  }

  for (var iD = 0; iD < this.duesStatuses.length; iD++) {
    if (!this.statusMap.hasOwnProperty(this.duesStatuses[iD].id)) {
      this.duesStatusMap[this.duesStatuses[iD].id] = this.duesStatuses[iD].fields['Paid/Unpaid'];
    }
  }

};

Airtable.prototype.fetchTables = function() {
  if (!this.isFetchComplete()) {
    Logger.log('Fetching Guests and Members');
    this.persons = this.fetch(this.baseURL + this.db_key + '/Guests%20and%20Members?api_key=' + this.api_key);
    Logger.log('Fetching Types');
    this.types = this.fetch(this.baseURL + this.db_key + '/Type?api_key=' + this.api_key);
    Logger.log('Fetching Statuses');
    this.statuses = this.fetch(this.baseURL + this.db_key + '/Status?api_key=' + this.api_key);
    Logger.log('Fetching Meetings');
    this.meetings = this.fetch(this.baseURL + this.db_key + '/Meetings?api_key=' + this.api_key);
    Logger.log('Fetching Dues');
    this.dues = this.fetch(this.baseURL + this.db_key + '/Dues?api_key=' + this.api_key);
    Logger.log('Fetching Dues Statuses');
    this.duesStatuses = this.fetch(this.baseURL + this.db_key + '/Dues%20Status?api_key=' + this.api_key);
  }
};

Airtable.prototype.isFetchComplete = function() {
  return (null !== this.persons &&
          null !== this.types &&
          null !== this.statuses &&
          null !== this.meetings &&
          null !== this.dues &&
          null !== this.duesStatuses);
};

Airtable.prototype.fetch = function(url, offset, records) {
  
  var response;
  if (offset === undefined) {
    response = UrlFetchApp.fetch(url);
    
  } else {
    response = UrlFetchApp.fetch(url + '&offset=' + offset);
  }
  
  var json = JSON.parse(response.getContentText());
  
  if (records === undefined) {
    records = json.records;
  } else {
    records = records.concat(json.records);
  }
  
  if (json.offset !== undefined) {
    records = this.fetch(url, json.offset, records);
  }
  
  return records;
};
