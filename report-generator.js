/*                    
 * This script outputs a monthly campaign report report
 * to a Google spreadsheet and sends it to alan@yaddayadda.com
 */

var RECIPIENT_EMAIL = "alan@yaddayadda.com";

var SPREADSHEET_URL = "";

var spreadsheetC = SpreadsheetApp.openByUrl(SPREADSHEET_URL).copy('Monthly Campaign Report' + new Date());



var spreadsheetA = SpreadsheetApp.openByUrl(SPREADSHEET_URL).copy('Monthly Ad Performance Report' + new Date());
  
var reportA = AdWordsApp.report(
    'SELECT Headline, Description1, Description2,CampaignName, AdGroupName, Impressions, Clicks, Ctr, AllConversionRate ' +
    'FROM   AD_PERFORMANCE_REPORT ' +
    'WHERE  Impressions > 0 ' +
    'DURING LAST_MONTH');

  
var reportC = AdWordsApp.report(
    'SELECT CampaignName, Amount, Clicks, Cost, Ctr, AverageCpc, AveragePosition, NumOfflineInteractions, Conversions ' +
    'FROM   CAMPAIGN_PERFORMANCE_REPORT ' +
    'WHERE  Impressions > 0 ' +
    'DURING LAST_MONTH');
 
var sheetC = spreadsheetC.getActiveSheet();


function main() {
  
  
  
  
  
  function getTotals(column) {
    var totalRows =0;
    var total = 0;
    var row = 3;
    var rows = reportC.rows();
    while(rows.hasNext()) {
     
       totalRows++; 
     
       rows.next();
     
      }
    var sheet = spreadsheetC.getActiveSheet();
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
    var rows = reportC.rows();
    while(rows.hasNext()) {
     
       totalRows++; 
       rows.next();
     
      }
    
    var account = AdWordsApp.currentAccount();
    var stats = account.getStatsFor("LAST_MONTH");
    var ctr = stats.getCtr();
    var outputRow = totalRows + 4;
    var sheet = spreadsheetC.getActiveSheet();
    sheet.getRange(outputRow,column).setValue(ctr);
    sheet.getRange(outputRow,column).setNumberFormat("0.00%");
   
     
    
  }
  
  function getAverageCPC() {
    
    var column = spreadsheetC.getRangeByName('AverageCPC').getColumn();
    var averageCPC = getTotals(4)/getTotals(3);
    //counts the number of rows
    var rows = reportC.rows(); 
    var totalRows =0;
    while(rows.hasNext()) {
     
       totalRows++; 
       rows.next();
     
      }
    var outputRow = totalRows+4;
    var sheet = spreadsheetC.getActiveSheet();
    sheet.getRange(outputRow, column).setValue(averageCPC);
    sheet.getRange(outputRow,column).setNumberFormat("$0.00");  
    
  }
  
  function getAveragePos() {
    
    var column = spreadsheetC.getRangeByName("AveragePosition").getColumn();
    var account = AdWordsApp.currentAccount();
    var stats = account.getStatsFor("LAST_MONTH");
    var avgPos = stats.getAveragePosition();
    var rows = reportC.rows();
    var totalRows =0;
    while(rows.hasNext()) {
     
       totalRows++; 
       rows.next();
     
      }
    var outputRow = totalRows+4;
    var sheet = spreadsheetC.getActiveSheet();
    sheet.getRange(outputRow, column).setValue(avgPos);
    sheet.getRange(outputRow,column).setNumberFormat("0.0");
    
    
    
    
  }
  
  
  
  
  
  
  
  function prepareReport(spreadsheet,report) {

  spreadsheet.addEditor("alan@webfordoctors.com");
  report.exportToSheet(spreadsheet.getActiveSheet());  
  var sheet = spreadsheet.getActiveSheet();
  var numColumns = 9;
  for (var i = 1; i < numColumns; i++) {
  
  sheet.autoResizeColumn(i);
    
  }
   
  
  //sheet.insertColumnBefore(1);
  sheet.insertRows(1);
    if (sheet === sheetC) {
  
      sheet.getRange(2, 2).setValue("Budget");
      sheet.getRange(2, 8).setValue("Phone Calls");
      
    }
  var date = new Date();
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
 
  
  getTotals(2);
  getTotals(3);
  getTotals(4);
  getTotals(8);
  getTotals(9);
  getCtr(5);
  getAverageCPC();
  getAveragePos();
  
  }
  
  //mail the report
  
  function mailReport() {
  
   
  if (RECIPIENT_EMAIL) {
    MailApp.sendEmail(RECIPIENT_EMAIL,
      'this is a test of the automated reporting system. you can ignore and or delete this email. ~alan',
                      "campaign report: " + spreadsheetC.getUrl() + " " + "ad performance report: " + spreadsheetA.getUrl()
               );
      
  }
    
  }
  
  prepareReport(spreadsheetC,reportC);
  prepareReport(spreadsheetA,reportA);
  mailReport();

  
}
