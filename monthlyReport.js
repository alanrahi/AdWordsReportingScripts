/*                    
 * This script outputs a monthly campaign report 
 * to a Google spreadsheet and sends it to alan@webfordoctors.com
 
 This header section of code contains the global variables accessed by many of the functions below. 
 */

//add recipient emails here
var RECIPIENT_EMAIL = "alan@yaddayadda.com";

//these are the template files, these will never change.
var campSPREADSHEET_URL = "";
var adSPREADSHEET_URL = " ";
var callSPREADSHEET_URL = "";

//copy the template files, adding current date.
var campSpreadsheet = SpreadsheetApp.openByUrl(campSPREADSHEET_URL).copy('Monthly Campaign Report' + new Date());
var adSpreadsheet = SpreadsheetApp.openByUrl(adSPREADSHEET_URL).copy('Monthly Ad Performance Report' + new Date());
var callSpreadsheet = SpreadsheetApp.openByUrl(callSPREADSHEET_URL).copy('Call Details Report' + new Date()); 


//here we pull the reports and store them in global variables
var adReport = AdWordsApp.report(
    'SELECT Headline, Description1, Description2,CampaignName, AdGroupName, Impressions, Clicks, Ctr, AllConversionRate ' +
    'FROM   AD_PERFORMANCE_REPORT ' +
    'WHERE  Impressions > 0 ' +
    'DURING LAST_MONTH '
    );

  
var campReport = AdWordsApp.report(
    'SELECT CampaignName, Amount, Clicks, Cost, Ctr, AverageCpc, AveragePosition, NumOfflineInteractions, Conversions ' +
    'FROM   CAMPAIGN_PERFORMANCE_REPORT ' +
    'WHERE  Impressions > 0 ' +
    'DURING LAST_MONTH ');

var callReport = AdWordsApp.report(
    'SELECT AdGroupName, CampaignName, CallDuration, CallStartTime, CallerNationalDesignatedCode ' +
    'FROM   CALL_METRICS_CALL_DETAILS_REPORT');
 
var campSheet = campSpreadsheet.getActiveSheet();

/* 


The rest of the code is stored inside this main function


*/

function main() {
  
  
  
/*
First we define basic fuctions to work with the data contained in those global variables
and print it to the report spreadsheets. 

*/
  
  //This counts up the number of rows in the report
  //It also takes a column as parameter and adds the sum of its contents to a totals row and formats the data
  function getTotals(column) {
    var totalRows =0;
    var total = 0;
    var row = 3;
    var rows = campReport.rows();
    while(rows.hasNext()) {
     
       totalRows++; 
     
       rows.next();
     
      }
    var sheet = campSpreadsheet.getActiveSheet();
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
  
  
  
  /*
  this function gets the aggregated stats for the past month, using adwords data rather than the
  math functions to retrieve averages.   */
  function getCtr(column) {
    var average = 0;
    var total = 0;
    var row = 3;
    var totalRows =0;
    var rows = campReport.rows();
    while(rows.hasNext()) {
     
       totalRows++; 
       rows.next();
     
      }
    
    var account = AdWordsApp.currentAccount();
    var stats = account.getStatsFor("LAST_MONTH");
    var ctr = stats.getCtr();
    var outputRow = totalRows + 4;
    var sheet = campSpreadsheet.getActiveSheet();
    sheet.getRange(outputRow,column).setValue(ctr);
    sheet.getRange(outputRow,column).setNumberFormat("0.00%");
   
     
    
  }
  
  
  //uses spreadsheet math to print averageCPC to sheet
  function getAverageCPC() {
    
    var column = campSpreadsheet.getRangeByName('AverageCPC').getColumn();
    var averageCPC = getTotals(4)/getTotals(3);
    //row counter
    var rows = campReport.rows(); 
    var totalRows =0;
    while(rows.hasNext()) {
     
       totalRows++; 
       rows.next();
     
      }
    var outputRow = totalRows+4;
    var sheet = campSpreadsheet.getActiveSheet();
    sheet.getRange(outputRow, column).setValue(averageCPC);
    sheet.getRange(outputRow,column).setNumberFormat("$0.00");  
    
  }
  
  //using adwords account data rather than math functions
  function getAveragePos() {
    
    var column = campSpreadsheet.getRangeByName("AveragePosition").getColumn();
    var account = AdWordsApp.currentAccount();
    var stats = account.getStatsFor("LAST_MONTH");
    var avgPos = stats.getAveragePosition();
    var rows = campReport.rows();
    var totalRows =0;
    while(rows.hasNext()) {
     
       totalRows++; 
       rows.next();
     
      }
    var outputRow = totalRows+4;
    var sheet = campSpreadsheet.getActiveSheet();
    sheet.getRange(outputRow, column).setValue(avgPos);
    sheet.getRange(outputRow,column).setNumberFormat("0.0");
    
    
    
    
  }
  
  
  
  
  
  //this function gets called later for each type of report we want to include
  //as of now we are running it for a campaign report, ad report, and call details report 
  //on lines 290 - 293
  
  function prepareReport(spreadsheet,report) {

  spreadsheet.addEditor("alan@webfordoctors.com");
  report.exportToSheet(spreadsheet.getActiveSheet());  
  var sheet = spreadsheet.getActiveSheet();
    Logger.log(sheet.getSheetName());
  //increase numColumns if you want to run reports with more than 9 category headlines
  var numColumns = 9;
  for (var i = 1; i < numColumns; i++) {
  
  sheet.autoResizeColumn(i);
    
  }
   
  
  //sheet.insertColumnBefore(1); 
  //if you want to add some space/margin to the left hand side.
  sheet.insertRows(1);
    
    //checks the report type and if this is a campaign report 
    //it resets certain headers to make it more readable.
    if (sheet.getSheetName() === "campaign") {
  
      sheet.getRange(2, 2).setValue("Budget");
      sheet.getRange(2, 8).setValue("Phone Calls");
      
    }
  
    
  /*here we could check to see if its a call details report
  
  
  if (sheet.getSheetName() === "campaign")   
    
    
    
  and then filter out all calls except for last month's */
    
  
  
  //this gets the name of the current month and the previous month for later use in a comparison statement
  //to later insert in the spreadsheet and email as needed.
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
    
  //you can customize the comparison string here:
  var dateRange = "showing data from " + month + "as compared to " + othermonth + " ";
  //put this message in the top left cell
  sheet.getRange(1,1).setValue(dateRange);
 
  //prints summary to last row and sets number formats
  getTotals(2);
  getTotals(3);
  getTotals(4);
  getTotals(8);
  getTotals(9);
  getCtr(5);
  getAverageCPC();
  getAveragePos();
  
  return dateRange;
  }
  
  
  //call this function later to mail the report after all prepareReport()s run.
  function mailReport() {
  
    
  
   
  if (RECIPIENT_EMAIL) {
    
    var SUBJECT = "test"
    //here is where we can customize the body of the report
    var BODY = "campaign report: " + campSpreadsheet.getUrl() +
      "   ad performance report: " + adSpreadsheet.getUrl() +
        "   call details report: " + callSpreadsheet.getUrl()
   
    MailApp.sendEmail(RECIPIENT_EMAIL,
             SUBJECT,
            BODY      
               );
      
  }
    
  }
  
  //call these functions once for each type of report you want to include. 
  //Still need to add function to dynamically insert additional reports into the email.
  //for now just add them to the email in the code above for each additional report. 
  //each time you call the prepareReport function give it the spreadsheet and report name as parameters.
  prepareReport(campSpreadsheet,campReport);
  prepareReport(adSpreadsheet,adReport);
  prepareReport(callSpreadsheet, callReport);
  mailReport();

  
}
