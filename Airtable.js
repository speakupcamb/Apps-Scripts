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

  this.personFields = [
    'Name',
    'First Name',
    'Last Name',
    'Type',
    'Status',
    'Email Address',
    'Address',
    'Phone',
  ];
  this.meetingAttendanceFields = [
    'Attended',
    'Meeting Date',
  ];
  this.meetingSpeechesFields = [
    'Speech',
    'Meeting Date',
  ];
  this.meetingRantsFields = [
    'Rant',
    'Meeting Date',
  ];
  this.duesFields = [
    'Person',
    'Due Period',
    'Status',
  ];
}

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

function TestAirtable() {

  // https://sites.google.com/site/scriptsexamples/custom-methods/gsunit

  var airtable = new Airtable();

  Logger.log('Testing fetch()');

  var duesStatus = airtable.fetch(airtable.baseURL + airtable.db_key + '/Dues%20Status?api_key=' + airtable.api_key);
  var persons = airtable.fetch(airtable.baseURL + airtable.db_key + '/Guests%20And%20Members?api_key=' + airtable.api_key);

  GSUnit.assertEquals('First row of Dues Status table', 'Unpaid', duesStatus[0].fields['Paid/Unpaid']);
  GSUnit.assertEquals('Second row of Dues Status table', 'Paid', duesStatus[1].fields['Paid/Unpaid']);
  GSUnit.assertTrue('Length of Guests and Members table', persons.length > 100);
  
  Logger.log('Testing isFetchComplete()');

  GSUnit.assertFalse('Fetch incomplete', airtable.isFetchComplete());
  
  airtable.fetchTables();
  
  GSUnit.assertTrue('Fetch complete', airtable.isFetchComplete());

  Logger.log('persons[0]:');
  Logger.log(airtable.persons[0]);
  Logger.log('types[0]:');
  Logger.log(airtable.types[0]);
  Logger.log('statuses[0]:');
  Logger.log(airtable.statuses[0]);
  Logger.log('meetings[0]:');
  Logger.log(airtable.meetings[0]);
  Logger.log('dues[0]:');
  Logger.log(airtable.dues[0]);
  Logger.log('duesStatuses[0]:');
  Logger.log(airtable.duesStatuses[0]);

  /*
    [18-09-07 21:30:20:340 EDT] persons[0]:
    [18-09-07 21:30:20:341 EDT] {createdTime=2018-05-16T00:33:57.000Z, id=rec0Jmw8dfSzlM8Ic, fields={Status=[recs1HDEhExfZInSI], Created Date/Time=2018-05-16T00:33:57.000Z, Type=[recvbtD72JdrQbOhu], First Name=Jagjeet, Last Name=Panesar, Email Address=jpanesar@gmail.com, Name=Jagjeet Panesar}}
    [18-09-07 21:30:20:341 EDT] types[0]:
    [18-09-07 21:30:20:342 EDT] {createdTime=2018-05-16T00:17:38.000Z, id=recPw9NgPLMASPFf4, fields={Status=Member/Officer, Person=[rec9glWgFrJAx7eiA, recLbyBV5QtAz0Lmp, recVos09UVdEe4cG6, recqQggqIM7YzGL61, recixycaTWRrpHHf8, rec2vJU1jPRDVxh9h, recaWXfb04g9aVnZs, recbRqvr0EQ4Ief4Y]}}
    [18-09-07 21:30:20:342 EDT] statuses[0]:
    [18-09-07 21:30:20:343 EDT] {createdTime=2018-05-16T00:14:06.000Z, id=recNue6lFsrDtVb33, fields={Person=[recTvybtY9GWEQx4R, recScybFW1Ey9qpIl, recfxNYVZdVK53ebl, recdYeBGkvpkYcYdz, recs8qf6U8gHR8F6p, recmdNXjbDWNgBbH4, reczPOdGfGKg18aYH, rec2RoNQz3vLqbWFf, recRFz3yGCyqsDYeg, rectBraz2AVnquAnf, rec1FqtRogc2yAwBe], Name=Inactive}}
    [18-09-07 21:30:20:343 EDT] meetings[0]:
    [18-09-07 21:30:20:344 EDT] {createdTime=2018-05-16T00:36:28.000Z, id=recAspf3cA5Gw2nwf, fields={Word and Thought=[recVos09UVdEe4cG6], General Evaluator=[rect9A4zUkEgIl5Ef], Timer=[recixycaTWRrpHHf8], Meeting Date=2018-05-09, Humorist=[rec9glWgFrJAx7eiA], Round Robin=[recaWXfb04g9aVnZs], Attended=[rech2u7jnyXac28ea, recVos09UVdEe4cG6, recqQggqIM7YzGL61, recaWXfb04g9aVnZs, recSNQj2tl0uBisZa, rect9A4zUkEgIl5Ef, rec2vJU1jPRDVxh9h, recbRqvr0EQ4Ief4Y, rec9glWgFrJAx7eiA], Toastmaster=[rech2u7jnyXac28ea]}}
    [18-09-07 21:30:20:344 EDT] dues[0]:
    [18-09-07 21:30:20:345 EDT] {createdTime=2018-05-16T01:56:47.000Z, id=recXtkYOcFz6DT2k0, fields={Status=Paid, Due Period=Spring 2018, Person=[recaWXfb04g9aVnZs]}}
    [18-09-07 21:30:20:345 EDT] duesStatuses[0]:
    [18-09-07 21:30:20:346 EDT] {createdTime=2018-05-16T00:42:17.000Z, id=recQ4VPT1F5WnhsGh, fields={Paid/Unpaid=Unpaid}}
  */

  Logger.log('Testing fetch()');

  // TODO: Complete
}
