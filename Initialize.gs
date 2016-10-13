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
 * @fileoverview Contains global variable declarations and menu initialization
 */

// Lab -> Get SAP User Id and password here
var USER = PropertiesService.getUserProperties().getProperty("UID");

var PASSWORD = PropertiesService.getUserProperties().getProperty("password");

var HEADER = 'Basic '+Utilities.base64Encode(USER+':'+PASSWORD);

try {
  var SHEET = SpreadsheetApp.getActiveSheet();
}
catch(err) {
  // Add On not yet Enabled
}

/**
 * Configure and create menu
 * @param {object} e environment
 */
function onOpen(){
  SpreadsheetApp.getActiveSpreadsheet()
  .addMenu('SAP Gateway',[{name:'Load Business Partners', functionName:'getBusinessPartners_'},  // not working
                            {name:'Load Sales Orders', functionName:'getSalesOrders_'},
                            null,
                            {name:'Load Items for Sales Order', functionName:'loadSalesItems_'},
                            null,
                            {name:'Clear Sheet', functionName:'clearSheet_'},
                            {name:'Delete Sheets', functionName:'deleteSheets_'},
                            null,
                            {name:'Set SAP Credentials', functionName:'setCredentialsForm_'}]);
}