// Written by Paul Ventresca, June 2018
// For issues/maintenance, please contact: psventresca@gmail.com and/or raymond.leclair@gmail.com
// Mucked with by Ray LeClair, July 18, 2018. Sorry.

//Global arrays
var apiRecords = [];
var descriptorMap = [];
var personMap = [];

// Define worksheets
var personSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Members/Guests');
var analyticsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Analytics');
var meetingAttendanceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Meetings-Attendance');
var meetingSpeechesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Meetings-Speeches');
var meetingRantsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Meetings-Rants');
var duesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Dues');
var memberNotAttendedSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Members Not Attended > Month');
var meetingCountSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Meeting Counts');

function refreshData() {
  
  // Prompt user on click of Refresh button
  /*
  if (Browser.msgBox('Are you sure you want to refresh the workbook? All prior data will be erased.', Browser.Buttons.YES_NO) == 'no') {
    return;
  }
  */
  
  // Get all required data from API (use global array to capture API calls, including offset calls)
  var persons = requestAPI("https://api.airtable.com/v0/apppcMJJYoQygtgfZ/Guests%20and%20Members?api_key=keyl4k6PFanvcnSPr");
  var personFields = ["Name","First Name","Last Name","Type","Status","Email Address","Address","Phone"];
  apiRecords = [];
  var types = requestAPI("https://api.airtable.com/v0/apppcMJJYoQygtgfZ/Type?api_key=keyl4k6PFanvcnSPr");
  apiRecords = [];
  var statuses = requestAPI("https://api.airtable.com/v0/apppcMJJYoQygtgfZ/Status?api_key=keyl4k6PFanvcnSPr");
  apiRecords = [];
  var meetings = requestAPI("https://api.airtable.com/v0/apppcMJJYoQygtgfZ/Meetings?api_key=keyl4k6PFanvcnSPr");
  var meetingAttendanceFields = ["Attended","Meeting Date"];
  var meetingSpeechesFields = ["Speech","Meeting Date"];
  var meetingRantsFields = ["Rant","Meeting Date"];
  apiRecords = [];
  var dues = requestAPI("https://api.airtable.com/v0/apppcMJJYoQygtgfZ/Dues?api_key=keyl4k6PFanvcnSPr");
  var duesFields = ["Person","Due Period","Status"];
  apiRecords = [];
  var duesStatuses = requestAPI("https://api.airtable.com/v0/apppcMJJYoQygtgfZ/Dues%20Status?api_key=keyl4k6PFanvcnSPr");
  apiRecords = [];
  
  // Create a key-value pair of lookups
  descriptorMap = createKeyValues(types,descriptorMap);
  descriptorMap.concat(createKeyValues(statuses,descriptorMap));
  descriptorMap.concat(createKeyValues(duesStatuses,descriptorMap));
  personMap = createKeyValues(persons,personMap,true);
  
  // Write data from API to appropriate data sheets
  writeJSONtoSheet(persons,personSheet,personFields);
  writeJSONtoSheet(meetings,meetingAttendanceSheet,meetingAttendanceFields,true);
  writeJSONtoSheet(meetings,meetingSpeechesSheet,meetingSpeechesFields,true);
  writeJSONtoSheet(meetings,meetingRantsSheet,meetingRantsFields,true);
  writeJSONtoSheet(dues,duesSheet,duesFields,true);
  
  // Find last row(s)
  var lastRowPersonSheet = personSheet.getLastRow();
  var lastRowMeetingSheet = meetingAttendanceSheet.getLastRow();
  
  // Populate Analytics sheet
  analyticsSheet.clear()
  analyticsSheet.getRange(1, 1).setValue("Speak Up Cambridge - Analytics");
  analyticsSheet.getRange(1, 1).setFontSize(12);
  analyticsSheet.getRange(1, 1).setFontWeight("bold");
  analyticsSheet.getRange(2, 1).setValue("Last Refresh: " + getTimestamp());
  analyticsSheet.getRange(2, 1).setFontSize(11);
  analyticsSheet.getRange(2, 1).setFontWeight("bold");
  analyticsSheet.getRange(4, 1).setValue("Members");
  analyticsSheet.getRange(4, 1).setFontWeight("bold");
  analyticsSheet.getRange(5, 1).setValue("Total Members");
  analyticsSheet.getRange(5, 2).setNumberFormat("0");
  analyticsSheet.getRange(6, 1).setValue("# Active");
  analyticsSheet.getRange(6, 2).setNumberFormat("0");
  analyticsSheet.getRange(7, 1).setValue("% Active");
  analyticsSheet.getRange(7, 2).setNumberFormat("0.00%");
  analyticsSheet.getRange(8, 1).setValue("# Inactive");
  analyticsSheet.getRange(8, 2).setNumberFormat("0");
  analyticsSheet.getRange(9, 1).setValue("% Inactive");
  analyticsSheet.getRange(9, 2).setNumberFormat("0.00%");
  analyticsSheet.getRange(11, 1).setValue("Guests");
  analyticsSheet.getRange(11, 1).setFontWeight("bold");
  analyticsSheet.getRange(12, 1).setValue("Total Guests");
  analyticsSheet.getRange(12, 2).setNumberFormat("0");
  analyticsSheet.getRange(13, 1).setValue("# Active");
  analyticsSheet.getRange(13, 2).setNumberFormat("0");
  analyticsSheet.getRange(14, 1).setValue("% Active");
  analyticsSheet.getRange(14, 2).setNumberFormat("0.00%");
  analyticsSheet.getRange(15, 1).setValue("# Inactive");
  analyticsSheet.getRange(15, 2).setNumberFormat("0");
  analyticsSheet.getRange(16, 1).setValue("% Inactive");
  analyticsSheet.getRange(16, 2).setNumberFormat("0.00%");
  analyticsSheet.getRange(18, 1).setValue("Meetings");
  analyticsSheet.getRange(18, 1).setFontWeight("bold");
  analyticsSheet.getRange(19, 1).setValue("Average Meeting Attendance");
  analyticsSheet.getRange(19, 2).setNumberFormat("0.00");
  analyticsSheet.getRange(20, 1).setValue("Average Member Attendance");
  analyticsSheet.getRange(20, 2).setNumberFormat("0.00");
  analyticsSheet.getRange(21, 1).setValue("Average Guest Attendance");
  analyticsSheet.getRange(21, 2).setNumberFormat("0.00");
  analyticsSheet.getRange(22, 1).setValue("Average Officer Attendance");
  analyticsSheet.getRange(22, 2).setNumberFormat("0.00");
  // Populate Members Not Attended > Month sheet
  memberNotAttendedSheet.clear()
  memberNotAttendedSheet.getRange(1, 1).setValue("Member");
  memberNotAttendedSheet.getRange(1, 1).setFontWeight("bold");
  memberNotAttendedSheet.getRange(1, 2).setValue("Last Attended");
  memberNotAttendedSheet.getRange(1, 2).setFontWeight("bold");

  // VLookup and for Meeting Attendee type
  meetingAttendanceSheet.getRange(1, 3).setValue("Attendee Type")
  meetingAttendanceSheet.getRange(1, 3).setFontWeight("bold")
  meetingAttendanceSheet.getRange(2,3,lastRowMeetingSheet-1,1).setFormulaR1C1('=vlookup(R[0]C[-2],\'Members/Guests\'!C[-2]:C[1],4,false)');
  // Meeting Count Formula
  meetingCountSheet.getRange(1, 1).setFormula('=query(\'Meetings-Attendance\'!B1:C' + lastRowMeetingSheet + ',"select B,count(B),C GROUP BY C,B")');
  meetingCountSheet.getRange(1, 1, 1, 3).setFontWeight("bold");
  // Members Formulas
  analyticsSheet.getRange(5, 2).setFormula('=(countif(\'Members/Guests\'!D2:D' + lastRowPersonSheet + ',"Member")+(countif(\'Members/Guests\'!D2:D' + lastRowPersonSheet + ',"Member/Officer")))');
  analyticsSheet.getRange(6, 2).setFormula('=countifs(\'Members/Guests\'!D2:D' + lastRowPersonSheet + ',"Member",\'Members/Guests\'!E2:E' + lastRowPersonSheet + ',"Active")+countifs(\'Members/Guests\'!D2:D' + lastRowPersonSheet + ',"Member/Officer",\'Members/Guests\'!E2:E' + lastRowPersonSheet + ',"Active")');
  analyticsSheet.getRange(7, 2).setFormula('=(countifs(\'Members/Guests\'!D2:D' + lastRowPersonSheet + ',"Member",\'Members/Guests\'!E2:E' + lastRowPersonSheet + ',"Active")+countifs(\'Members/Guests\'!D2:D' + lastRowPersonSheet + ',"Member/Officer",\'Members/Guests\'!E2:E' + lastRowPersonSheet + ',"Active"))/(countif(\'Members/Guests\'!D2:D' + lastRowPersonSheet + ',"Member")+(countif(\'Members/Guests\'!D2:D' + lastRowPersonSheet + ',"Member/Officer")))');
  analyticsSheet.getRange(8, 2).setFormula('=countifs(\'Members/Guests\'!D2:D' + lastRowPersonSheet + ',"Member",\'Members/Guests\'!E2:E' + lastRowPersonSheet + ',"Inactive")+countifs(\'Members/Guests\'!D2:D' + lastRowPersonSheet + ',"Member/Officer",\'Members/Guests\'!E2:E' + lastRowPersonSheet + ',"Inactive")');
  analyticsSheet.getRange(9, 2).setFormula('=(countifs(\'Members/Guests\'!D2:D' + lastRowPersonSheet + ',"Member",\'Members/Guests\'!E2:E' + lastRowPersonSheet + ',"Inactive")+countifs(\'Members/Guests\'!D2:D' + lastRowPersonSheet + ',"Member/Officer",\'Members/Guests\'!E2:E' + lastRowPersonSheet + ',"Inactive"))/(countif(\'Members/Guests\'!D2:D' + lastRowPersonSheet + ',"Member")+(countif(\'Members/Guests\'!D2:D' + lastRowPersonSheet + ',"Member/Officer")))');
  // Guests Formulas
  analyticsSheet.getRange(12, 2).setFormula('=(countif(\'Members/Guests\'!D2:D' + lastRowPersonSheet + ',"Guest"))');
  analyticsSheet.getRange(13, 2).setFormula('=countifs(\'Members/Guests\'!D2:D' + lastRowPersonSheet + ',"Guest",\'Members/Guests\'!E2:E' + lastRowPersonSheet + ',"Active")');
  analyticsSheet.getRange(14, 2).setFormula('=countifs(\'Members/Guests\'!D2:D' + lastRowPersonSheet + ',"Guest",\'Members/Guests\'!E2:E' + lastRowPersonSheet + ',"Active")/countif(\'Members/Guests\'!D2:D' + lastRowPersonSheet + ',"Guest")');
  analyticsSheet.getRange(15, 2).setFormula('=countifs(\'Members/Guests\'!D2:D' + lastRowPersonSheet + ',"Guest",\'Members/Guests\'!E2:E' + lastRowPersonSheet + ',"Inactive")');
  analyticsSheet.getRange(16, 2).setFormula('=countifs(\'Members/Guests\'!D2:D' + lastRowPersonSheet + ',"Guest",\'Members/Guests\'!E2:E' + lastRowPersonSheet + ',"Inactive")/countif(\'Members/Guests\'!D2:D' + lastRowPersonSheet + ',"Guest")');
  // Meetings Formulas
  analyticsSheet.getRange(19, 2).setFormula('=sum(\'Meeting Counts\'!B:B)/count(unique(\'Meeting Counts\'!A:A))');
  analyticsSheet.getRange(20, 2).setFormula('=sumif(\'Meeting Counts\'!C:C,"Member",\'Meeting Counts\'!B:B)/count(unique(\'Meeting Counts\'!A:A))');
  analyticsSheet.getRange(21, 2).setFormula('=sumif(\'Meeting Counts\'!C:C,"Guest",\'Meeting Counts\'!B:B)/count(unique(\'Meeting Counts\'!A:A))');
  analyticsSheet.getRange(22, 2).setFormula('=sumif(\'Meeting Counts\'!C:C,"Member/Officer",\'Meeting Counts\'!B:B)/count(unique(\'Meeting Counts\'!A:A))');
  // Members/Guest Formulas
  personSheet.getRange(1, 9).setValue("Last Meeting Attended");
  personSheet.getRange(1, 9).setFontWeight("bold");
  personSheet.getRange(2,9,lastRowPersonSheet-1,1).setFormulaR1C1('=if(maxifs(\'Meetings-Attendance\'!C[-7]:C[-7],\'Meetings-Attendance\'!C[-8]:C[-8],R[0]C[-8])<>0,maxifs(\'Meetings-Attendance\'!C[-7]:C[-7],\'Meetings-Attendance\'!C[-8]:C[-8],R[0]C[-8]),"")');
  personSheet.getRange(1, 10).setValue("Last Speech");
  personSheet.getRange(1, 10).setFontWeight("bold");
  personSheet.getRange(2,10,lastRowPersonSheet-1,1).setFormulaR1C1('=if(maxifs(\'Meetings-Speeches\'!C[-8]:C[-8],\'Meetings-Speeches\'!C[-9]:C[-9],R[0]C[-9])<>0,maxifs(\'Meetings-Speeches\'!C[-8]:C[-8],\'Meetings-Speeches\'!C[-9]:C[-9],R[0]C[-9]),"")');
  personSheet.getRange(1, 11).setValue("Last Rant");
  personSheet.getRange(1, 11).setFontWeight("bold");
  personSheet.getRange(2,11,lastRowPersonSheet-1,1).setFormulaR1C1('=if(maxifs(\'Meetings-Rants\'!C[-9]:C[-9],\'Meetings-Rants\'!C[-10]:C[-10],R[0]C[-10])<>0,maxifs(\'Meetings-Rants\'!C[-9]:C[-9],\'Meetings-Rants\'!C[-10]:C[-10],R[0]C[-10]),"")');

  
  // Call external functions
  findAbsenteeMembers();
  
  // Sort/size the sheets
  personSheet.sort(1, true);
  meetingAttendanceSheet.sort(2, false);
  meetingSpeechesSheet.sort(2, false);
  meetingRantsSheet.sort(2, false);
  duesSheet.sort(1, true);
  analyticsSheet.autoResizeColumns(1, analyticsSheet.getLastColumn());
  
  // Put cursor in default location for each sheet
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  for (var i = 0; i < sheets.length ; i++ ) {
    sheets[i].getRange(2, 1).activate();
  }
  
  analyticsSheet.getRange(1, 1).activate();
  Browser.msgBox("Refresh complete!");

};

// Finds and populates list of Members who haven't attended a meeting (and whose attendance was recorded) in over a month
function findAbsenteeMembers() {
  
  var x = 2;
  var checkDate;
  var todayDate = new Date();
  var lastMonthDate = new Date();
  lastMonthDate.setDate(todayDate.getDate()-30);
  var rangeData = personSheet.getDataRange();
  var lastColumn = rangeData.getLastColumn();
  var lastRow = rangeData.getLastRow();
  var searchRange = personSheet.getRange(2,1, lastRow-1, lastColumn-1);
  
  // Get array of values in the search Range
  var rangeValues = searchRange.getValues();
  // Loop through array and if condition met, add it to memberNotAttendedSheet
  for (j = 0 ; j < lastRow - 1; j++){
    checkDate = rangeValues[j][8];
    if(checkDate) {
      if(checkDate < lastMonthDate){
        memberNotAttendedSheet.getRange(x,1).setValue(rangeValues[j][0]);
        memberNotAttendedSheet.getRange(x,2).setValue(rangeValues[j][8]);
        x++;
      }
    }
  }

  memberNotAttendedSheet.sort(2, true);
  memberNotAttendedSheet.autoResizeColumns(1, memberNotAttendedSheet.getLastColumn());
    
};

// Recursive function to make an API request and any applicable offsets (limit is 100 records per API call)
function requestAPI(url,offset,records) {
  
  if(offset) {
    var response = UrlFetchApp.fetch(url + "&offset=" + offset);
  }
  else {
    var response = UrlFetchApp.fetch(url);
  }
  var json = JSON.parse(response.getContentText());
  
  apiRecords = apiRecords.concat(json.records);
  
  offset = json.offset
  if(offset) {
    requestAPI(url,offset,apiRecords);
  }
  
  return apiRecords;
};

//Function to write all JSON data to worksheet
function writeJSONtoSheet(jsonObj,sheet,keys,isOneToMany) {
  
  var rowArray = [];
  var rowSubArray = [];
  var dataArray = [];
  var val;
  var staticVal;
  var isArr;
  var valList;
  //var keys = Object.keys(jsonObj[0].fields);
   
  //Populate fields on sheet (user-provided)
  sheet.clear();
  sheet.getRange(1, 1, 1, keys.length).setValues([keys]);
  sheet.getRange(1, 1, 1, keys.length).setFontWeight("bold");

  // Loop through all values and build two-dimensional data array
  // If there is a one-to-many relationship in the data, handle the "manies" first (always first key of dataset)
  if(isOneToMany) {
    for(var x = 0; x < jsonObj.length; x++) {
      valList = jsonObj[x].fields[keys[0]];
      if(valList) {
        for(var i = 0; i < valList.length; i++) {
          rowArray = [];
          rowArray.push(findVal(valList[i],personMap));
          for(var z = 1; z < keys.length; z++) {
            val = jsonObj[x].fields[keys[z]];
            rowArray = populateRowArray(rowArray,val,personMap);
          }
          dataArray.push(rowArray);
        }
      }
    }
  }
  // Otherwise handle as a one-to-one relationship
  else {
    for(var x = 0; x < jsonObj.length; x++) {
      rowArray = [];
      for(var i = 0; i < keys.length; i++) {
        val = jsonObj[x].fields[keys[i]];
        rowArray = populateRowArray(rowArray,val,descriptorMap);
      }
      dataArray.push(rowArray);
    }
  }
  
  //Populate data on worksheet
  sheet.getRange(sheet.getLastRow()+1, 1,dataArray.length,dataArray[0].length).setValues(dataArray);
  sheet.autoResizeColumns(1, sheet.getLastColumn());
};

// Creates Key-Value pairs for lookup values
function createKeyValues(lookupArray,outputArray,person) {
  var keys = Object.keys(lookupArray[0].fields);
  for(i = 0; i < lookupArray.length; i++) {
    if(person) {
      outputArray.push([lookupArray[i].id,lookupArray[i].fields['Name']]);
    }
    else {
      outputArray.push([lookupArray[i].id,lookupArray[i].fields[keys[0]]]);
    }
  }
  return outputArray;
};

// Finds value within key-value pair when needed
function findVal(val,lookIn) {
  var idx;
  for (var k in lookIn) {
    idx = lookIn[k][0].indexOf(val)
    var tmp = lookIn[k];
    if (idx == 0) {
      return lookIn[k][1];  
    }
  }
  return "";
};

// Gets current date/time and formats it accordingly
function getTimestamp() {
  var currentTime = new Date();
  var timeZone = CalendarApp.getDefaultCalendar().getTimeZone();
  var timeString = Utilities.formatDate(currentTime, timeZone,'EEEE, M/d/yy, hh:mm:ss aaa');
  return timeString;
};

// Validates and populates one row of data to be inserted into data array
function populateRowArray(rowArray,val,map) {
  if(!val) {
    rowArray.push("");
  }
  else {
    isArr = Array.isArray(val)
    if(isArr) {
      rowArray.push(findVal(val[0],map));
    }
    else {
      rowArray.push(val);
    }
  }
  return rowArray;
};