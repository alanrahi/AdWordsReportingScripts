// URL of the default spreadsheet template.
var SPREADSHEET_URL = 'https://docs.google.com/spreadsheets/d/1vM_OZ_FnAzFa4xxHUJfq_-dqGNOAViOyD6aRpZpjg3c/edit?usp=sharing_eid'
/**                    
 * This script computes an Ad performance report
 * and outputs it to a Google spreadsheet.
 */
function main() {
  
  Logger.log('Using template spreadsheet - %s.', SPREADSHEET_URL);

  //open the template spreadsheet  
  var spreadsheet = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  
  Logger.log(spreadsheet.getName());
  
  var sheet = spreadsheet.getSheets()[0];
  sheet.getRange(1, 2, 1, 1).setValue(new Date());
  //make a copy of the template and rename it  
  var report = spreadsheet.copy('monthly ad report');
  
  // Logger.log(report.getName());
  
  //empty object for holding all values together
  var segmentMap = {};
  
  var row = 4;
  
  //retrieve the campaigns that you want to work with
  
  var campaignSelector = AdWordsApp.campaigns()
     .withCondition("Name CONTAINS_IGNORE_CASE 'campaign'");
  
  var campaignIterator = campaignSelector.get();
  
  //iterate over campaigns
  while (campaignIterator.hasNext()) {
    var campaign = campaignIterator.next();
    //retrieve the ads
    var adSelector = campaign.ads()
       .withCondition("Status = ENABLED");
     
    var adIterator = adSelector.get();
    Logger.log(adIterator.totalNumEntities());
    
    //iterate over the ads
    while (adIterator.hasNext() || row < adIterator.totalNumEntities()) {
      
      
      var ad = adIterator.next();
      //print ad headlines to column 2
      var headline = ad.getHeadline();
      var id = ad.getId();
      sheet.getRange(row, 2).setValue(headline);
      sheet.getRange(row, 3).setValue(id);
      //move down a row
      row++;
    
    }
    
  } 
 


  
  
  
}
