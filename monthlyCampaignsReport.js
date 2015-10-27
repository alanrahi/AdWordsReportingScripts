/*                    
 * This script outputs a monthly campaign report report
 * to a Google spreadsheet and sends it to alan@webfordoctors.com
 */

var RECIPIENT_EMAIL = "alan@webfordoctors.com";

var SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1WqvMobtwIAVRRBhcnrEzfOTKD2QzGCyFm3UvmfZrWdg/edit#gid=0";



function main() {
  
  
  
  
  
  function getTotals(column) {
    var totalRows =0;
    var total = 0;
    var row = 3;
    var rows = report.rows();
    while(rows.hasNext()) {
     
       totalRows++; 
     
       rows.next();
     
      }
    
    
    var rowlooper = row;
    while (rowlooper < totalRows+row) {
      
      total += sheet.getRange(rowlooper, column).getValue();
      
      rowlooper++;
        
      }
    
    
    var outputRow = totalRows + 4;
    
    sheet.getRange(outputRow, 1).setValue("total");
    sheet.getRange(outputRow,column).setValue(total); 
    sheet.getRange(outputRow,4).setNumberFormat("$0.00");
    sheet.getRange(outputRow,2).setNumberFormat("$0.00");
    return total;
    
  }
  
  
  function getCtr(column) {
    var average = 0;
    var total = 0;
    var row = 3;
    var totalRows =0;
    var rows = report.rows();
    while(rows.hasNext()) {
     
       totalRows++; 
       rows.next();
     
      }
    
    var account = AdWordsApp.currentAccount();
    var stats = account.getStatsFor("LAST_MONTH");
    var ctr = stats.getCtr();
    var outputRow = totalRows + 4;
    sheet.getRange(outputRow,column).setValue(ctr);
    sheet.getRange(outputRow,column).setNumberFormat("0.00%");
   
     
    
  }
  
  function getAverageCPC() {
    
    var column = spreadsheet.getRangeByName('AverageCPC').getColumn();
    var averageCPC = getTotals(4)/getTotals(3);
    //counts the number of rows
    var rows = report.rows(); 
    var totalRows =0;
    while(rows.hasNext()) {
     
       totalRows++; 
       rows.next();
     
      }
    var outputRow = totalRows+4;
    
    sheet.getRange(outputRow, column).setValue(averageCPC);
    sheet.getRange(outputRow,column).setNumberFormat("$0.00");  
    
  }
  
  function getAveragePos() {
    
    var column = spreadsheet.getRangeByName("AveragePosition").getColumn();
    var account = AdWordsApp.currentAccount();
    var stats = account.getStatsFor("LAST_MONTH");
    var avgPos = stats.getAveragePosition();
    
    var rows = report.rows(); 
    var totalRows =0;
    while(rows.hasNext()) {
     
       totalRows++; 
       rows.next();
     
      }
    var outputRow = totalRows+4;
    sheet.getRange(outputRow, column).setValue(avgPos);
    sheet.getRange(outputRow,column).setNumberFormat("0.0");
    
    
    
    
  }
  
  
  
  
  
  
  
  

  
  var spreadsheet = SpreadsheetApp.openByUrl(SPREADSHEET_URL).copy('Monthly Campaign Report' + new Date());
  spreadsheet.addEditor("alan@webfordoctors.com");

  
  var report = AdWordsApp.report(
    'SELECT CampaignName, Amount, Clicks, Cost, Ctr, AverageCpc, AveragePosition, NumOfflineInteractions, Conversions ' +
    'FROM   CAMPAIGN_PERFORMANCE_REPORT ' +
    'WHERE  Impressions > 0 ' +
    'DURING LAST_MONTH');
 
  
  report.exportToSheet(spreadsheet.getActiveSheet());
  var sheet = spreadsheet.getActiveSheet();
  var numColumns = 9;
  for (var i = 1; i < numColumns; i++) {
  
  sheet.autoResizeColumn(i);
    
  }
  
  //sheet.insertColumnBefore(1);
  sheet.insertRows(1);
  sheet.getRange(2, 2).setValue("Budget");
  sheet.getRange(2, 8).setValue("Phone Calls");
  var date = new Date();
  Logger.log(date);
  Logger.log(typeof date);
  var month = date.getMonth();
  var day = date.getDay();
  var year = date.getYear();
  
  var months = ["january","february","march","april","may","june","july","august","september","october","november","december"];
  for (var i = 1; i < 13; i++) {
    
    if (month === i && i === 1) {
      
     month = months[0];
      
    }
   
    else if (month === i) {
    
    month = months[i-1];
    var othermonth = months[i-2];
    
  }
  }
    
  
  var dateRange = "comparing " + month + " to " + othermonth + " ";
  sheet.getRange(1,1).setValue(dateRange);
  
  
  /*
  trying to generate time range for the report dynamically. need to finish this later
 
 for each (var value in date) {
  Logger.log(value);
}
  */
  
  //populate the sheet with totals, averages and additional labels
  
  //sheet.getRange(1,1).setValue(date);
  
  getTotals(2);
  getTotals(3);
  getTotals(4);
  getTotals(8);
  getTotals(9);
  getCtr(5);
  getAverageCPC();
  getAveragePos();
  
 
  
 
  //mail the spreadsheet link
  
  if (RECIPIENT_EMAIL) {
    MailApp.sendEmail(RECIPIENT_EMAIL,
      'this is a test of the automated reporting system. you can ignore and or delete this email. ~alan',
      spreadsheet.getUrl());
  }
  
  

  
}
