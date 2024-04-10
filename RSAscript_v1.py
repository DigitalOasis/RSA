// Copyright 2024. Increase BV. All Rights Reserved.
//
// Created By: Digital Oasis
// for Increase B.V.
//
// Created: 04-03-2024
// Last update: 
//
// ABOUT THE SCRIPT
//This script will exports RSA and assets performance and the corresponding fields.
//The script exports out Campaign Name, Campaign Status, AdGroup Name,AdGroup Status, Ad id, AD Strength, Ad Status, Asset Copy Text, Asset Type, Performance Label,Asset Pinned position,Impressions 
//into excel spread.
//Please specify DATERANGE. If you do not specify a fromDate and toDate then the daterange DATERANGE to last 90 days
// 
const clientCode= ''
let NumberOfDays=90 //change the number of days required to be pulled
let fromDate= '01/10/2023' // specify fromdate in dd/mm/yyyy format
let toDate= '31/12/2023'   // specify todate in dd/mm/yyyy format

var config = {
  
  LOG : true,
  
  // Make a copy of this script and copy the URL: https://docs.google.com/spreadsheets/d/1WvNSbaZi2dz3Uu74AIniKy5YLHctZ_6c6i0Gp2DQ8ns/copy
  SPREADSHEET_URL : "https://docs.google.com/spreadsheets/d/1wakQjIXaVIqQatlqGBdYDaglbQml9hMTDuZGqcDiNjU/edit#gid=980901217",
  SHEET_NAME : ["RSA Asset Performance"],
  QA_Query :[],
  date: NumberOfDays,
  fromDate: fromDate,
  toDate: toDate
  
}


function main() {
  
  if(config.SPREADSHEET_URL == "https://"){
    throw Error("Make a copy of the sheet and paste the URL in the config \nhttps://docs.google.com/spreadsheets/d/1WvNSbaZi2dz3Uu74AIniKy5YLHctZ_6c6i0Gp2DQ8ns/copy");
  }  
  
  
 
  let CurrentaccountName = AdsApp.currentAccount().getName();
  let tag = clientCode ? clientCode : CurrentaccountName;
  var ss = SpreadsheetApp.openByUrl(config.SPREADSHEET_URL);
  ss.rename(tag + ' Google Ads RSA Performance Insights - digitaloasis.com.au');
  
  let defaultSettings= {
    NumberofDays: config.date,
    fromDate: config.fromDate,
    toDate: config.toDate
  };
  
  
  //let settings= updateVariablesFromSheet(ss, defaultSettings);
  let numberofdays= defaultSettings.NumberofDays;
  let fromDate = defaultSettings.fromDate;
  let toDate = defaultSettings.toDate;
  
  
  let timeZone= AdsApp.currentAccount().getTimeZone();
  let dateCheck = fromDate !== undefined && toDate !== undefined ? 1 : 0;
  
  let today = new Date(), yesterday = new Date(), startDate = new Date();
  yesterday.setDate(today.getDate() - 1);
  startDate.setDate(today.getDate() - numberofdays);
  
  let formattedStartDate = Utilities.formatDate(startDate, timeZone, 'yyyy-MM-dd');
  let formattedYesterday = Utilities.formatDate(yesterday, timeZone, 'yyyy-MM-dd');
  
  function formatDate(dateString) {
    // Use a regular expression to extract date parts
    let dateParts = dateString.match(/(\d{2})\/(\d{2})\/(\d{4})/);
    if (!dateParts) {
      throw new Error('Date is not in a valid format. Expected format dd/mm/yyyy.');
    }
    // Rearrange the date parts to 'yyyy-MM-dd' format
    let formattedDate = `${dateParts[3]}-${dateParts[2]}-${dateParts[1]}`;
    return formattedDate;
  }

  let formattedFromDate = dateCheck ? formatDate(fromDate) : undefined;
  let formattedToDate   = dateCheck ? formatDate(toDate) : undefined;
  
  let DateRangeQuery = dateCheck ? `WHERE segments.date BETWEEN "${formattedFromDate}" AND "${formattedToDate}"` : `WHERE segments.date BETWEEN "${formattedStartDate}" AND "${formattedYesterday}"`;

  config.QA_Query = [ "SELECT campaign.name, campaign.status, ad_group.name,ad_group.status, ad_group.id,ad_group_ad.ad_strength, ad_group_ad.status, asset.text_asset.text, ad_group_ad_asset_view.field_type, ad_group_ad_asset_view.performance_label, ad_group_ad_asset_view.pinned_field, metrics.impressions FROM ad_group_ad_asset_view " + 
                     " WHERE ad_group.status != 'REMOVED' AND metrics.impressions != 0 ORDER BY ad_group_ad_asset_view.performance_label DESC"
                     ]
  
  for (let i=0; i < config.SHEET_NAME.length-1; i++){
    
    var sheet = ss.getSheetByName(config.SHEET_NAME[i]);

    var report = AdsApp.report(config.QA_Query[i]);
    
    // Export data and clean up sheet
    sheet.clearContents();
    report.exportToSheet(sheet);

    if(config.SHEET_NAME[i] == "RSA Asset Performance" ){
        var customColumnNames = ["Campaign Name", "Campaign Status", "AdGroup Name","AdGroup Status", "Ad id", "AD Strength", "Ad Status", "Asset Copy Text", "Asset Type", "Performance Label", "Asset Pinned position", "Impressions"];
        var range = sheet.getRange(1, 1, 1, customColumnNames.length);
        range.setValues([customColumnNames]);
  
  }
    sheet.autoResizeColumns(1, sheet.getLastColumn());
  
  if(sheet.getMaxColumns() - sheet.getLastColumn() != 0){
    sheet.deleteColumns(sheet.getLastColumn() + 1, sheet.getMaxColumns() - sheet.getLastColumn());
  }
  }

  
  Logger.log("Export completed");
  
} // function main()

