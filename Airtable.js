'use strict';

/**
 * Constructs an Airtable object, and initializes its properties.
 */
function Airtable() {
  
  Logger.log('Initializing Airtable');
  
  // Airtable API parameters
  this.baseURL = 'https://api.airtable.com/v0/';
  this.db_key = 'apppcMJJYoQygtgfZ';
  this.api_key = 'keyl4k6PFanvcnSPr';
  
  // Airtable data as returned
  this.persons = null;
  this.types = null;
  this.statuses = null;
  this.meetings = null;
  this.dues = null;
  this.duesStatuses = null;

  // Meeting activities
  this.meetingRoles = [
    'Toastmaster',
    'Word and Thought',
    'Humorist',
    'Round Robin',
    'Table Topics Leader',
    'General Evaluator',
    'Speech Evaluator',
    'Timer',
    'Grammarian',
    'Ah-Um Counter'
  ];
  this.meetingSpeeches = [
    'Rant',
    'Speech',
    'Table Topics'
  ];

  // Airtable data as mapped
  this.typeMap = {};
  this.statusMap = {};
  this.duesStatusMap = {};
  this.personMap = {};

  // Airtable data as collected
  this.personData = {};
  this.personData.header = [
    'Name',
    'First Name',
    'Last Name',
    'Type',
    'Status',
    'Email Address',
    'Last Attended Date',
    'Last Role Date',
    'Last Role Name',
    'Last Speech Date',
    'Last Speech Name'
  ];
  this.personData.value = [
    '',  // 'Name'
    '',  // 'First Name'
    '',  // 'Last Name'
    '',  // 'Type'
    '',  // 'Status'
    '',  // 'Email Address'
    '2000-01-01',  // 'Last Attended Date'
    '2000-01-01',  // 'Last Role Date'
    '',  // 'Last Role Name'
    '2000-01-01',  // 'Last Speech Date'
    ''  // 'Last Speech Name'
  ];
  this.personData.rows = [];

  this.meetingData = {};
  this.meetingData.header = [
    'Date',
    'Guests',
    'Members',
    'Speeches',
    'Rants',
    'Table Topics',
    'Toastmaster',
    'Word and Thought',
    'Humorist',
    'Round Robin',
    'Table Topics Leader',
    'General Evaluator',
    'Timer',
    'Grammarian',
    'Ah-Um Counter'
  ];
  this.meetingData.value = [
    '',  // 'Date'
    0,  // 'Guests'
    0,  // 'Members'
    0,  // 'Speeches'
    0,  // 'Rants'
    0,  // 'Table Topics'
    '',  // 'Toastmaster'
    '',  // 'Word and Thought'
    '',  // 'Humorist'
    '',  // 'Round Robin'
    '',  // 'Table Topics Leader'
    '',  // 'General Evaluator'
    '',  // 'Timer'
    '',  // 'Grammarian'
    ''  // 'Ah-Um Counter'
  ];
  this.meetingData.rows = [];

  this.activityData = {};
  this.activityData.header = [
    'Name',
    'Type',
    'Status',
    'Date',
    'Activity'
  ];
  this.activityData.rows = [];

}

/**
 * Collect Airtable data in a format for reporting using Google Sheets.
 *
 * @return {Undefined}
 */
Airtable.prototype.collectData = function() {
  
  var iP, iM, iT, iS, iDS, iH;
  var person, meeting;
  var row, header, value;
  
  // Create mappings from type, status, and dues status ids to values
  // for reporting
  for (iT = 0; iT < this.types.length; iT++) {
    if (!this.typeMap.hasOwnProperty(this.types[iT].id)) {
      this.typeMap[this.types[iT].id] = this.types[iT].fields['Status'];
    }
  }
  for (iS = 0; iS < this.statuses.length; iS++) {
    if (!this.statusMap.hasOwnProperty(this.statuses[iS].id)) {
      this.statusMap[this.statuses[iS].id] = this.statuses[iS].fields['Name'];
    }
  }
  for (iDS = 0; iDS < this.duesStatuses.length; iDS++) {
    if (!this.duesStatusMap.hasOwnProperty(this.duesStatuses[iDS].id)) {
      this.duesStatusMap[this.duesStatuses[iDS].id] = this.duesStatuses[iDS].fields['Paid/Unpaid'];
    }
  }

  // Create mapping from person ids to person fields object for fast
  // lookup when collecting meeting data
  for (iP = 0; iP < this.persons.length; iP++) {
    if (!this.personMap.hasOwnProperty(this.persons[iP].id)) {
      person = this.persons[iP].fields;

      // Ensure the person object has all expected fields
      for (iH = 0; iH < this.personData.header.length; iH++) {
        header = this.personData.header[iH];
        if (!person.hasOwnProperty(header)) {
          person[header] = this.personData.value[iH];
        }
      }

      // Map type, status, and dues status ids to values for reporting
      person['Type'] = this.typeMap[person['Type']];
      person['Status'] = this.statusMap[person['Status']];
      person['Dues Status'] = this.duesStatusMap[person['Dues Status']];
      
      this.personMap[this.persons[iP].id] = person;
    }
  }

  // Collect meeting activities (attended, role, or speech) for each
  // person
  for (iM = 0; iM < this.meetings.length; iM++) {
    meeting = this.meetings[iM].fields;
    
    // Ensure the meeting object has all expected fields
    for (iH = 0; iH < this.meetingData.header.length; iH++) {
      header = this.meetingData.header[iH];
      if (!meeting.hasOwnProperty(header)) {
        meeting[header] = this.meetingData.value[iH];
      }
    }

    // Assign meeting activities for each person
    this.assignActivity(meeting, ['Attended'], 'Last Attended');
    this.assignActivity(meeting, this.meetingRoles, 'Last Role');
    this.assignActivity(meeting, this.meetingSpeeches, 'Last Speech');

    // Populate meeting data for reporting
    row = [];
    for (iH = 0; iH < this.meetingData.header.length; iH++) {
      header = this.meetingData.header[iH];
      value = meeting[header];
      row.push(value);
    }
    this.meetingData.rows.push(row);
  }

  // Populate person data for reporting
  for (iP = 0; iP < this.persons.length; iP++) {
    person = this.personMap[this.persons[iP].id];
    
    row = [];
    for (iH = 0; iH < this.personData.header.length; iH++) {
      header = this.personData.header[iH];
      value = this.persons[iP].fields[header];
      row.push(value);
    }
    this.personData.rows.push(row);
  }

};

/**
 * Assign meeting activities (attended, role, or speech) for each
 * person.
 *
 * @param meeting {Object} a meeting object
 * @param meeting {Array} a list of activities
 * @param activities {String} a lable for reporting
 *
 * @return {Undefined}
 */
Airtable.prototype.assignActivity = function(meeting, activities) {

  // Consider each activity (attended, role, or speech)
  for (var iA = 0; iA < activities.length; iA++) {
    var activity = activities[iA];
    if (!meeting.hasOwnProperty(activity)) {
      continue;
    }

    // Consider each person participating in the current activity
    var personIds = meeting[activity];
    for (var iId = 0; iId < personIds.length; iId++) {
      var person = this.personMap[personIds[iId]];

      // Record the activity for the current meeting and person
      this.activityData.rows.push(
        [person['Name'], person['Type'], person['Status'], meeting['Meeting Date'], activity]);

      // Count guests, members, speeches, rants, or table topics, or
      // record a name with each role for the current meeting
      if (activity === 'Attended') {
        if (person['Type'] === 'Guest') {
          meeting['Guests'] += 1;
        } else {
          meeting['Members'] += 1;
        }
      } else if (this.meetingSpeeches.indexOf(activity) > -1) {
        meeting[activity] += 1;

      } else if (this.meetingRoles.indexOf(activity) > -1) {
        meeting[activity] = person['Name'];
      }
        
      // Assign the most recent activity for the current person. Note
      // that dates are strings formatted as YYYY-MM-DD, and so sort
      // as expected.
      var assignAs;
      if (activity === 'Attended') {
        assignAs = 'Last Attended';

      } else if (this.meetingSpeeches.indexOf(activity) > -1) {
        assignAs = 'Last Speech';

      } else if (this.meetingRoles.indexOf(activity) > -1) {
        assignAs = 'Last Role';
      }
      if (meeting['Meeting Date'] > person[assignAs + ' Date']) {
        person[assignAs + ' Date'] = meeting['Meeting Date'];
        if (assignAs !== 'Last Attended') {
          person[assignAs + ' Name'] = activity;
        }
      }
    }
  }
};

/**
 * Fetch all required tables using the Airtable API, if fetching of
 * any required table has not been completed.
 *
 * @return {Undefined}
 */
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

/**
 * Indicate if the fetch for all required tables has been completed,
 * or not.
 *
 * @return {Boolean} true if complete
 */
Airtable.prototype.isFetchComplete = function() {
  return (null !== this.persons &&
          null !== this.types &&
          null !== this.statuses &&
          null !== this.meetings &&
          null !== this.dues &&
          null !== this.duesStatuses);
};

/**
 * Fetch all records for an Airtable URL using the API.
 *
 * @param {String} url Airtable URL
 * @param {String} offset Airtable provided offset id
 * @param {Array} records Airtable records
 *
 * @return {Array} Airtable records
 */
Airtable.prototype.fetch = function(url, offset, records) {
  
  // Use the API to fetch a batch of records
  var response;
  if (offset === undefined) {
    response = UrlFetchApp.fetch(url);
    
  } else {
    response = UrlFetchApp.fetch(url + '&offset=' + offset);
  }
  
  // Parse the respone
  var json = JSON.parse(response.getContentText());
  
  // Collect the records
  if (records === undefined) {
    records = json.records;
  } else {
    records = records.concat(json.records);
  }
  
  // Fetch another batch of records, if required
  if (json.offset !== undefined) {
    records = this.fetch(url, json.offset, records);
  }
  
  return records;
};
