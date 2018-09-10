function Reports() {
  
  var airtable = new Airtable();
  
  airtable.fetchTables();
  airtable.collectData();
  
  var sheets = new Sheets();
  
  sheets.getSheets();
  sheets.writeDataToSheet(airtable.personData, sheets.personSheet);
  
}
