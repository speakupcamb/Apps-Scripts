'use strict';

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
  Logger.log(airtable.types[0].fields);
  Logger.log(airtable.types[0].fields.Status);
  Logger.log('statuses[0]:');
  Logger.log(airtable.statuses[0]);
  Logger.log('meetings[0]:');
  Logger.log(airtable.meetings[0]);
  Logger.log('dues[0]:');
  Logger.log(airtable.dues[0]);
  Logger.log('duesStatuses[0]:');
  Logger.log(airtable.duesStatuses[0]);
  
  /*
  [18-09-08 13:03:38:397 EDT] persons[0]:
  [18-09-08 13:03:38:397 EDT] {createdTime=2018-05-16T00:33:57.000Z, id=rec0Jmw8dfSzlM8Ic, fields={Status=[recs1HDEhExfZInSI], Created Date/Time=2018-05-16T00:33:57.000Z, Type=[recvbtD72JdrQbOhu], First Name=Jagjeet, Last Name=Panesar, Email Address=jpanesar@gmail.com, Name=Jagjeet Panesar}}
  [18-09-08 13:03:38:398 EDT] types[0]:
  [18-09-08 13:03:38:398 EDT] {createdTime=2018-05-16T00:17:38.000Z, id=recPw9NgPLMASPFf4, fields={Status=Member/Officer, Person=[rec9glWgFrJAx7eiA, recLbyBV5QtAz0Lmp, recVos09UVdEe4cG6, recqQggqIM7YzGL61, recixycaTWRrpHHf8, rec2vJU1jPRDVxh9h, recaWXfb04g9aVnZs, recbRqvr0EQ4Ief4Y]}}
  [18-09-08 13:03:38:399 EDT] statuses[0]:
  [18-09-08 13:03:38:399 EDT] {createdTime=2018-05-16T00:14:06.000Z, id=recNue6lFsrDtVb33, fields={Person=[recTvybtY9GWEQx4R, recScybFW1Ey9qpIl, recfxNYVZdVK53ebl, recdYeBGkvpkYcYdz, recs8qf6U8gHR8F6p, recmdNXjbDWNgBbH4, reczPOdGfGKg18aYH, rec2RoNQz3vLqbWFf, recRFz3yGCyqsDYeg, rectBraz2AVnquAnf, rec1FqtRogc2yAwBe], Name=Inactive}}
  [18-09-08 13:03:38:400 EDT] meetings[0]:
  [18-09-08 13:03:38:400 EDT] {createdTime=2018-05-16T00:36:28.000Z, id=recAspf3cA5Gw2nwf, fields={Word and Thought=[recVos09UVdEe4cG6], General Evaluator=[rect9A4zUkEgIl5Ef], Timer=[recixycaTWRrpHHf8], Meeting Date=2018-05-09, Humorist=[rec9glWgFrJAx7eiA], Round Robin=[recaWXfb04g9aVnZs], Attended=[rech2u7jnyXac28ea, recVos09UVdEe4cG6, recqQggqIM7YzGL61, recaWXfb04g9aVnZs, recSNQj2tl0uBisZa, rect9A4zUkEgIl5Ef, rec2vJU1jPRDVxh9h, recbRqvr0EQ4Ief4Y, rec9glWgFrJAx7eiA], Toastmaster=[rech2u7jnyXac28ea]}}
  [18-09-08 13:03:38:401 EDT] dues[0]:
  [18-09-08 13:03:38:401 EDT] {createdTime=2018-05-16T01:56:47.000Z, id=recXtkYOcFz6DT2k0, fields={Status=Paid, Due Period=Spring 2018, Person=[recaWXfb04g9aVnZs]}}
  [18-09-08 13:03:38:402 EDT] duesStatuses[0]:
  [18-09-08 13:03:38:402 EDT] {createdTime=2018-05-16T00:42:17.000Z, id=recQ4VPT1F5WnhsGh, fields={Paid/Unpaid=Unpaid}}
  */
  
  Logger.log('Testing fetchTables()');
  
  var person = airtable.persons.filter(function (person) {
    return person.id === 'rec0Jmw8dfSzlM8Ic';
  })[0].fields;
  GSUnit.assertEquals('Selected person First Name', 'Jagjeet', person['First Name']);
  
  Logger.log('Testing createMaps');
  
  airtable.createMaps();
  
  GSUnit.assertEquals('Type map for Member', 'Member', airtable.typeMap['reclyIyLOaIs6YXwV']);
  
  Logger.log('Testing assignActivity');
  
  var meeting = airtable.meetings.filter(function (meeting) {
    return meeting.id === 'recAspf3cA5Gw2nwf';
  })[0].fields;
  airtable.assignActivity(meeting, ['Toastmaster']);
  
  GSUnit.assertEquals('Assign Toastmaster for selected meeting', 'Andrea Moore', meeting['Toastmaster Name']);
  
  Logger.log('Testing collectData');
  
  airtable.collectData();
  
  GSUnit.assertEquals('Collect Member Count for selected meeting', 9, meeting['Member Count']);
  
}

function TestSheets() {
  
  // https://sites.google.com/site/scriptsexamples/custom-methods/gsunit
  
  var sheets = new Sheets();
  
  Logger.log('Test getSheetByName()');
  
  personSheet = sheets.getSheetByName('Persons');
  
  GSUnit.assertArrayEquals('Header of Guests and Members sheet',
                           ["Name", "First Name", "Last Name", "Type", "Status", "Email Address"],
                           personSheet.getRange(1, 1, 1, 6).getDisplayValues()[0]);
  
  Logger.log('Testing isGetSheetsComplete()');
  
  GSUnit.assertFalse('Get incomplete', sheets.isGetSheetsComplete());
  
  sheets.getSheets();
  
  GSUnit.assertTrue('Get complete', sheets.isGetSheetsComplete());
  
  Logger.log('Testing getSheets()');
    
  GSUnit.assertEquals('Get person sheet second column header', 'First Name', sheets.personSheet.getRange(1, 2, 1, 2).getValue());
  GSUnit.assertEquals('Get meeting sheet second column header', 'Guest Count', sheets.meetingSheet.getRange(1, 2, 1, 2).getValue());
  GSUnit.assertEquals('Get activity sheet second column header', 'Type', sheets.activitySheet.getRange(1, 2, 1, 2).getValue());
  
}
