/**
 * Copyright 2015 Google Inc. All rights reserved.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 * 
 *     http://www.apache.org/licenses/LICENSE-2.0

 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

/**
 * @fileoverview Set credentials and import records
 */

/**
 * Create form dialog for developer to enter SAP credentials
 * @private
 */
function setCredentialsForm_() {
  var template = HtmlService.createTemplateFromFile('CredentialsUi');  // calls runReport() 
  Logger.log("USER "+USER);
  template.user = USER;
  template.password = PASSWORD;
  var output = template.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME).setTitle("Set SAP Credentials").setWidth(300);

  SpreadsheetApp.getUi().showSidebar(output);
}

/**
 * Store SAP credentials
 * @param {string} user
 * @param {string} password
 */
function setCredentials(user,password) {
  PropertiesService.getUserProperties().setProperty("SAPGatewayUser",user);
  PropertiesService.getUserProperties().setProperty("SAPGatewayPassword",password);
}

/**
 * Import Business Partners from SAP Gateway to spreadsheet
 * @private
 */
function getBusinessPartners_(){
  var sheet = createOrSetActiveSheet_('Business Partners');
  sheet.clear();
  
  // Lab -> Set SAP business partner collection url and data format here
    
    var base_url = "https://sapes1.sapdevcenter.com/sap/opu/odata/IWBEP/GWDEMO/BusinessPartnerCollection/?";
    var additionParams = "$format=json";  
    var getDataURL = base_url + additionParams;
  Logger.log(getDataURL);
  
  // Lab -> Fetch business partner data from SAP here
  
  
  var dataResponse = UrlFetchApp.fetch(getDataURL,{headers: {Authorization: HEADER}, muteHttpExceptions: true}).getContentText();

  var dataObj = JSON.parse(dataResponse);
  
  
  var headers = ['Key','Name','Role','Website','Email','Address'];
  
  var maxLetter = String.fromCharCode(64 + headers.length);
  
  sheet.appendRow(headers);
  
  sheet.getRange('A1:'+maxLetter+'1').setBackground('Grey').setFontStyle('oblique');
  
  sheet.setColumnWidth(2, 220);
  sheet.setColumnWidth(3, 89);
  
  // Lab -> add business partner data to spreadsheet here

   var rows = [];
  
  for(var i = 0;dataObj.d.results && i!=dataObj.d.results.length;i++){
    
    var row = dataObj.d.results[i];
    
    rows.push([row.BusinessPartnerKey,row.Company,row.BusinessPartnerRoleText,row.WebAddress, row.EmailAddress, 'street address']);
    
  }
  
  var rowsCount = rows.length+1;
  
  sheet.getRange('A2:'+maxLetter+rowsCount).setValues(rows);
  
  sheet.setName('Business Partners');
  
  
}

/**
 * Import Sales Orders from SAP Gateway to spreadsheet
 * @private
 */
function getSalesOrders_(){
  var sheet = createOrSetActiveSheet_('Sales Orders');
  sheet.clear();
  
  // Set SAP sales order collection url and data format
  var base_url = "https://sapes1.sapdevcenter.com/sap/opu/odata/IWBEP/GWDEMO/SalesOrderCollection/?";
  var additionParams = "$format=json";

  var getDataURL = base_url + additionParams;
  Logger.log(getDataURL);
  
  // Fetch sales order data from SAP
  var dataResponse = UrlFetchApp.fetch(getDataURL,{headers: {Authorization: HEADER}, muteHttpExceptions: true}).getContentText();  
  var dataObj = JSON.parse(dataResponse);
  var headers = ['Sales Order Key','Customer Name','Status','Currency','Total Amt','Tax Amt','Note','Created By','Changed By'];
  
  var maxLetter = String.fromCharCode(64 + headers.length);
  
  sheet.appendRow(headers);
  
  sheet.getRange('A1:'+maxLetter+'1').setBackground('Grey').setFontStyle('oblique');
  
  sheet.setColumnWidth(2, 220);
  sheet.setColumnWidth(3, 89);
  
  // add sales order data to spreadsheet
  var rows = [];
  for(var i = 0;dataObj.d.results && i<dataObj.d.results.length;i++){
    var row = dataObj.d.results[i];
    
    rows.push([row.SalesOrderKey,row.CustomerName,row.StatusDescription,row.Currency,row.TotalSum,row.Tax,row.Note,row.CreatedByEmployeeLastName,row.ChangedByEmployeeLastName]);
  }
  var rowsCount = rows.length+1;
  sheet.getRange('A2:'+maxLetter+rowsCount).setValues(rows);
}

/**
 * Import Sales Items from SAP Gateway to spreadsheet for selected Sales Order 
 * @private
 */
function loadSalesItems_(){
  
  var sheetName = SHEET.getName();
  var activeRow = SHEET.getActiveRange().getRow();
  if (sheetName != 'Sales Orders' || activeRow == 1 || activeRow > SHEET.getLastRow()) {
    //didn't select a valid row
     Browser.msgBox("Please select a valid spreadsheet row in the Sales Orders sheet");
  }
  else {
    
    var salesOrderKey = SpreadsheetApp.getActiveRange().getValue();
    var initialPageCount = 99; //because a default sheet has only 100 rows to start
    
    Logger.log("salesOrderKey "+salesOrderKey);
    var sheet = createOrSetActiveSheet_(salesOrderKey);
    sheet.clearContents;
    var base_url = "https://sapes1.sapdevcenter.com/sap/opu/odata/IWBEP/GWDEMO/SalesOrderCollection('"+salesOrderKey+"')/salesorderlineitems?";
    var getDataURL = base_url + '$format=json';
    Logger.log(getDataURL);
    
    var dataResponse = UrlFetchApp.fetch(getDataURL,{headers: {Authorization: HEADER}, muteHttpExceptions: true}).getContentText();  
    Logger.log(dataResponse);
    var dataObj = JSON.parse(dataResponse);
    var headers = ['SalesOrderItemKey','ProductName','Availability','Note','Currency','NetSum','Tax','TotalSum'];
    var maxLetter = String.fromCharCode(64 + headers.length);
    
    sheet.appendRow(headers);
    sheet.getRange('A1:'+maxLetter+'1').setBackground('Grey').setFontStyle('oblique');
    var rows = [];
    for(var i = 0;dataObj.d.results && i<dataObj.d.results.length;i++){
      var row = dataObj.d.results[i];
      rows.push([row.SalesOrderItemKey,row.ProductName,row.Availability,row.Note,row.Currency,row.NetSum,row.Tax,row.TotalSum]);
    }
    var rowsCount = rows.length+1;
    sheet.getRange('A2:'+maxLetter+rowsCount).setValues(rows);
  }
}

/**
 * Create new sheet to match SAP record type if required, otherwise set existing sheet to active 
 * @private
 */
function createOrSetActiveSheet_(sheetName){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var sheet;
  for(var i = 0;i<sheets.length;i++){
    sheet = sheets[i];
    var currentSheetName = sheet.getSheetName();
    if(sheetName===currentSheetName){
      ss.setActiveSheet(sheet);
      sheet.clear();
      return sheet;
    }
  }
  sheet = ss.insertSheet();
  sheet.setName(sheetName);
  ss.setActiveSheet(sheet);
  return sheet;
}


/**
 * Clear sheet
 * @private
 */
function clearSheet_() {
  SHEET.clearContents();
}

/**
 * Delete all sheets
 * @private
 */
function deleteSheets_(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var sheet;
  var len = sheets.length;
  for(var i=len;i--;){
    sheet = sheets[i];
    ss.deleteSheet(sheet);
  }
}
