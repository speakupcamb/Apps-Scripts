// Written by Paul Ventresca, June 2018
// For issues/maintenance, please contact: psventresca@gmail.com
// Mucked with by Ray LeClair, July 18, 2018. Sorry.

function refreshData() {
  
  // Prompt user on click of Refresh button
  if (Browser.msgBox('Are you sure you want to refresh the workbook? All prior data will be erased.', Browser.Buttons.YES_NO) == 'no') {
    return;
  }
  
  // Get all required data for Members/Guests
  var persons = requestAPI("https://api.airtable.com/v0/apppcMJJYoQygtgfZ/Guests%20and%20Members?api_key=keyl4k6PFanvcnSPr")
  var types = requestAPI("https://api.airtable.com/v0/apppcMJJYoQygtgfZ/Type?api_key=keyl4k6PFanvcnSPr")
  var statuses = requestAPI("https://api.airtable.com/v0/apppcMJJYoQygtgfZ/Status?api_key=keyl4k6PFanvcnSPr")
  var meetings = requestAPI("https://api.airtable.com/v0/apppcMJJYoQygtgfZ/Meetings?api_key=keyl4k6PFanvcnSPr")
  
  // Define variables
  var lastRowPersonSheet = 0
  var lastRowMeetingSheet = 0
  var attended = ""
  var person = ""
  var personType = ""
  var personIndex = -1
  var typeIndex = -1
  var statusIndex = -1
  var personType = ""
  var personStatus = ""
  var personArray = []
  var typeArray = []
  var meetingArray = []
  var meetingAttendeeArray = []
  var statusArray = []
  var personData = persons.records
  var typeData = types.records
  var statusData = statuses.records
  var personOffset = persons.offset
  var meetingOffset = meetings.offset
  var meetingData = meetings.records
  
  // Define spreadsheet variables
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var personSheet
  var analyticsSheet
  var meetingSheet
  
  // Populate Type data
  for (var j in typeData){
    typeArray.push([typeData[j].id,typeData[j].fields.Status])
  }
  
  // Populate Status data
  for (var j in statusData){
    statusArray.push([statusData[j].id,statusData[j].fields.Name])
  }
  
  // Populate Person data
  for (var j in personData){
    
    for (var k in typeArray) {
      typeIndex = typeArray[k][0].indexOf(personData[j].fields.Type[0])
      if (typeIndex == 0) {
        personType = typeArray[k][1]
        break;
      }
    }
    for (var l in statusArray) {
      statusIndex = statusArray[l][0].indexOf(personData[j].fields.Status[0])
      if (statusIndex == 0) {
        personStatus = statusArray[l][1]
        break;
      }
    }
    
    personArray.push([personData[j].fields.Name,personType,personStatus,personData[j].id])
  }
  
  // Gather additional Persons data (if necessary). API only returns 100 records at a time
  if (personOffset) {
    while (personOffset) {
      persons = requestAPI("https://api.airtable.com/v0/apppcMJJYoQygtgfZ/Guests%20and%20Members?api_key=keyl4k6PFanvcnSPr&offset=" + personOffset)
      personData = persons.records
      
      for (var j in personData){
        
        for (var k in typeArray) {
          typeIndex = typeArray[k][0].indexOf(personData[j].fields.Type[0])
          if (typeIndex == 0) {
            personType = typeArray[k][1]
            break;
          }
        }
        for (var l in statusArray) {
          statusIndex = statusArray[l][0].indexOf(personData[j].fields.Status[0])
          if (statusIndex == 0) {
            personStatus = statusArray[l][1]
            break;
          }
        }
        
        personArray.push([personData[j].fields.Name,personType,personStatus,personData[j].id])
      }
      personOffset = personData.offset
    }
  }
  
  // Populate Meeting data
  for (var j in meetingData){
    for (var x in meetingData[j].fields.Attended) {
      attended = meetingData[j].fields.Attended[x];
      
    for (var k in personArray) {
      personIndex = personArray[k][3].indexOf(attended)
      if (personIndex == 0) {
        person = personArray[k][0]
        personType = personArray[k][1]
        break;
      }
    }
      
      meetingArray.push([meetingData[j].fields['Meeting Date'],person,personType])
    }
  }
  
  // Gather additional Meeting data (if necessary). API only returns 100 records at a time
  if (meetingOffset) {
    while (meetingOffset) {
      meetings = requestAPI("https://api.airtable.com/v0/apppcMJJYoQygtgfZ/Meetings?api_key=keyl4k6PFanvcnSPr&offset=" + meetingOffset)
      meetingData = meetings.records
      
      for (var j in meetingData){
        for (var x in meetingData[j].fields.Attended) {
          attended = meetingData[j].fields.Attended[x];
          
          for (var k in personArray) {
            personIndex = personArray[k][3].indexOf(attended)
            if (personIndex == 0) {
              person = personArray[k][0]
              personType = personArray[k][1]
              break;
            }
          }
          
          meetingArray.push([meetingData[j].fields['Meeting Date'],person,personType])
        }
      }
      meetingOffset = meetingData.offset
    }
  }
  
  // Populate Members/Guests sheet
  personSheet = ss.getSheetByName('Members/Guests');
  personSheet.clear()
  personSheet.getRange(1, 1).setValue("Name")
  personSheet.getRange(1, 2).setValue("Type")
  personSheet.getRange(1, 3).setValue("Status")
  personSheet.getRange(1, 4).setValue("Person ID")
  personSheet.getRange(1, 1, 1, personSheet.getLastColumn()).setFontWeight("bold")
  personSheet.getRange(personSheet.getLastRow()+1, 1,personArray.length,personArray[0].length).setValues(personArray)
  personSheet.autoResizeColumns(1, personSheet.getLastColumn())
  lastRowPersonSheet = personSheet.getLastRow();
  
  // Populate Analytics sheet
  analyticsSheet = ss.getSheetByName('Analytics');
  analyticsSheet.clear()
  analyticsSheet.getRange(1, 1).setValue("Speak Up Cambridge - Analytics")
  analyticsSheet.getRange(1, 1).setFontSize(12)
  analyticsSheet.getRange(1, 1).setFontWeight("bold")
  analyticsSheet.getRange(2, 1).setValue("Last Refresh: " + getTimestamp())
  analyticsSheet.getRange(2, 1).setFontSize(11)
  analyticsSheet.getRange(2, 1).setFontWeight("bold")
  analyticsSheet.getRange(4, 1).setValue("Members")
  analyticsSheet.getRange(4, 1).setFontWeight("bold")
  analyticsSheet.getRange(5, 1).setValue("Total Members")
  analyticsSheet.getRange(6, 1).setValue("# Active")
  analyticsSheet.getRange(7, 1).setValue("% Active")
  analyticsSheet.getRange(8, 1).setValue("# Inactive")
  analyticsSheet.getRange(9, 1).setValue("% Inactive")
  analyticsSheet.getRange(11, 1).setValue("Meetings")
  analyticsSheet.getRange(11, 1).setFontWeight("bold")
  analyticsSheet.getRange(12, 1).setValue("Average Meeting Attendance")
  analyticsSheet.getRange(13, 1).setValue("Guests")
  analyticsSheet.getRange(14, 1).setValue("Members")
  analyticsSheet.getRange(16, 1).setValue("Guests")
  analyticsSheet.getRange(16, 1).setFontWeight("bold")
  analyticsSheet.getRange(17, 1).setValue("Total Guests")
  analyticsSheet.getRange(18, 1).setValue("# Active")
  analyticsSheet.getRange(19, 1).setValue("% Active")
  analyticsSheet.getRange(20, 1).setValue("# Inactive")
  analyticsSheet.getRange(21, 1).setValue("% Inactive")
  analyticsSheet.getRange(23, 1).setValue("Conversion")
  analyticsSheet.getRange(23, 1).setFontWeight("bold")
  analyticsSheet.getRange(24, 1).setValue("TBD")
  
  // Populate Meetings sheet
  meetingSheet = ss.getSheetByName('Meetings');
  meetingSheet.clear()
  meetingSheet.getRange(1, 1).setValue("Meeting")
  meetingSheet.getRange(1, 2).setValue("Person")
  meetingSheet.getRange(1, 3).setValue("Person Type")
  meetingSheet.getRange(1, 1, 1, meetingSheet.getLastColumn()).setFontWeight("bold")
  meetingSheet.getRange(meetingSheet.getLastRow()+1, 1,meetingArray.length,meetingArray[0].length).setValues(meetingArray)
  meetingSheet.autoResizeColumns(1, meetingSheet.getLastColumn())
  lastRowMeetingSheet = meetingSheet.getLastRow();
  
  // Members Formulas
  analyticsSheet.getRange(5, 2).setFormula('=(countif(\'Members/Guests\'!B2:B' + lastRowPersonSheet + ',"Member")+(countif(\'Members/Guests\'!B2:B' + lastRowPersonSheet + ',"Member/Officer")))');
  analyticsSheet.getRange(6, 2).setFormula('=countifs(\'Members/Guests\'!B2:B' + lastRowPersonSheet + ',"Member",\'Members/Guests\'!C2:C' + lastRowPersonSheet + ',"Active")+countifs(\'Members/Guests\'!B2:B' + lastRowPersonSheet + ',"Member/Officer",\'Members/Guests\'!C2:C' + lastRowPersonSheet + ',"Active")');
  analyticsSheet.getRange(7, 2).setFormula('=(countifs(\'Members/Guests\'!B2:B' + lastRowPersonSheet + ',"Member",\'Members/Guests\'!C2:C' + lastRowPersonSheet + ',"Active")+countifs(\'Members/Guests\'!B2:B' + lastRowPersonSheet + ',"Member/Officer",\'Members/Guests\'!C2:C' + lastRowPersonSheet + ',"Active"))/(countif(\'Members/Guests\'!B2:B' + lastRowPersonSheet + ',"Member")+(countif(\'Members/Guests\'!B2:B' + lastRowPersonSheet + ',"Member/Officer")))');
  analyticsSheet.getRange(8, 2).setFormula('=countifs(\'Members/Guests\'!B2:B' + lastRowPersonSheet + ',"Member",\'Members/Guests\'!C2:C' + lastRowPersonSheet + ',"Inactive")+countifs(\'Members/Guests\'!B2:B' + lastRowPersonSheet + ',"Member/Officer",\'Members/Guests\'!C2:C' + lastRowPersonSheet + ',"Inactive")');
  analyticsSheet.getRange(9, 2).setFormula('=(countifs(\'Members/Guests\'!B2:B' + lastRowPersonSheet + ',"Member",\'Members/Guests\'!C2:C' + lastRowPersonSheet + ',"Inactive")+countifs(\'Members/Guests\'!B2:B' + lastRowPersonSheet + ',"Member/Officer",\'Members/Guests\'!C2:C' + lastRowPersonSheet + ',"Inactive"))/(countif(\'Members/Guests\'!B2:B' + lastRowPersonSheet + ',"Member")+(countif(\'Members/Guests\'!B2:B' + lastRowPersonSheet + ',"Member/Officer")))');
  // Meetings Formulas
  analyticsSheet.getRange(12, 2).setFormula('=count(Meetings!A2:A' + lastRowMeetingSheet + ')/count(unique(Meetings!A2:A' + lastRowMeetingSheet + '))');
  analyticsSheet.getRange(13, 2).setFormula('=countif(Meetings!C2:C' + lastRowMeetingSheet + ',"Guest")');
  analyticsSheet.getRange(14, 2).setFormula('=countif(Meetings!C2:C' + lastRowMeetingSheet + ',"Member")+countif(Meetings!C2:C' + lastRowMeetingSheet + ',"Member/Officer")');
  // Guests Formulas
  analyticsSheet.getRange(17, 2).setFormula('=(countif(\'Members/Guests\'!B2:B' + lastRowPersonSheet + ',"Guest"))');
  analyticsSheet.getRange(18, 2).setFormula('=countifs(\'Members/Guests\'!B2:B' + lastRowPersonSheet + ',"Guest",\'Members/Guests\'!C2:C' + lastRowPersonSheet + ',"Active")');
  analyticsSheet.getRange(19, 2).setFormula('=countifs(\'Members/Guests\'!B2:B' + lastRowPersonSheet + ',"Guest",\'Members/Guests\'!C2:C' + lastRowPersonSheet + ',"Active")/countif(\'Members/Guests\'!B2:B' + lastRowPersonSheet + ',"Guest")');
  analyticsSheet.getRange(20, 2).setFormula('=countifs(\'Members/Guests\'!B2:B' + lastRowPersonSheet + ',"Guest",\'Members/Guests\'!C2:C' + lastRowPersonSheet + ',"Inactive")');
  analyticsSheet.getRange(21, 2).setFormula('=countifs(\'Members/Guests\'!B2:B' + lastRowPersonSheet + ',"Guest",\'Members/Guests\'!C2:C' + lastRowPersonSheet + ',"Inactive")/countif(\'Members/Guests\'!B2:B' + lastRowPersonSheet + ',"Guest")');
  
}

// Standard function to make an API request
function requestAPI(url) {
  var response = UrlFetchApp.fetch(url);
  return JSON.parse(response.getContentText());
}

// Gets current date/time and formats it accordingly
function getTimestamp() {
  var currentTime = new Date();
  var timeZone = CalendarApp.getDefaultCalendar().getTimeZone();
  var timeString = Utilities.formatDate(currentTime, timeZone,'EEEE, M/d/yy, hh:mm:ss aaa')
  return timeString;
}
