/*GApps Mail Merge
*version 1.0
*License GPL
*Author: Brian Tong (bahrieinn@gmail.com)


/*Copying Permission
	This program is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with this program.  If not, see <http://www.gnu.org/licenses/>.

/*Description
 This is basically a mashup of many mail merge scripts already out there
 the only difference being an included editor to compose templated emails all within the spreadsheet
 Majority of script functionality straight from Google's Tutorial: Simple Mail Merge.
 Snippet for sent status and timestamp borrowed from alexdesignworks' mail merge script
 */

//Creates menu item within spreadsheet and shows template notes on open
function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menu = [
    {name: "Start App", functionName: "beginUI"}
    ];
  ss.addMenu("Mail Merge", menu);
  Browser.msgBox("Note: column for recipient emails must be titled 'Email' or 'email' for script to work");
}


//Initiates UI to compose and send templated email 
function beginUI(e) {
  //Initialize App Window
  var app = UiApp.createApplication().setHeight(320).setWidth(500).setTitle('Mail Merge Template Editor!');
  var tabPanel = app.createTabPanel().setHeight('400px').setWidth('500px');
  var flowPanel1 = app.createFlowPanel().setHeight('280').setWidth('485px').setStyleAttributes({background: 'F5F5F5', padding:'0px'});
  
  //Place flowpanels in tab structure
  app.add(tabPanel);
  tabPanel.add(flowPanel1, "Compose Email Template");  
  tabPanel.selectTab(0); //set default to first tab
  
  //Pulling text from most recent edit
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var templateSheet = ss.getSheets()[1];
  var emailTemplate = templateSheet.getRange("A1").getValue();
 
  //UI Elements
  var subjectBox = app.createTextBox().setWidth('250px').setName("subjectBox");
  var messageBox = app.createTextArea().setSize('250px', '150px').setName("messageBox").setText(emailTemplate); //Setting default text most recent edit
  var subjectLabel = app.createLabel('Subject').setStyleAttributes({textAlign: 'right', color: '858585'});
  var messageLabel = app.createLabel('Message').setStyleAttributes({textAlign: 'right', color: '858585'});
  var sendButton = app.createButton('Send Emails')
                      .setStyleAttribute('background', 'D64937')
                      .setStyleAttribute('color','white')
                      .setStyleAttribute('borderStyle', 'none')
                      .setWidth('250px');
  var panelLabel1 = app.createLabel("Create templates and send to all contacts in spreadsheet"); 
  var howtoLabel = app.createLabel("To insert variables use: $%header name%").setStyleAttributes({color: '858585', background:'F9EDBE', borderWidth: '2px', borderColor: 'DAA025', textAlign: 'left', fontSize: '0.8em', padding: '0.5em'});
  var exLabel = app.createLabel('e.g: \n Hi $%first name%, is $%phone number% your correct phone number?').setStyleAttributes({color: '858585', background:'F9EDBE', borderWidth: '1px', borderColor: 'DAA025', textAlign: 'left', fontSize: '0.8em', padding: '0.5em'});
  var meLabel = app.createLabel('mashup by bahrieinn').setStyleAttributes({fontSize: '0.6em', textAlign: 'right', verticalAlign: 'bottom'});
  
  //Layout
  var vertPanel1 = app.createVerticalPanel();
  var grid1 = app.createGrid(4, 3);
  
  grid1.setWidget(0, 1, panelLabel1);
  grid1.setWidget(1, 0, subjectLabel);
  grid1.setWidget(1, 1, subjectBox);
  grid1.setWidget(2, 0, messageLabel);
  grid1.setWidget(2, 1, messageBox);
  grid1.setWidget(1, 2, howtoLabel);
  grid1.setWidget(2, 2, exLabel);
  grid1.setWidget(3, 1, sendButton);
  grid1.setWidget(3, 2, app.createLabel(MailApp.getRemainingDailyQuota() + " messages left today \(gmail data limit\)").setStyleAttribute('fontSize', '0.8em'));
  
  flowPanel1.add(vertPanel1);
  flowPanel1.add(grid1);
  flowPanel1.add(meLabel);
  //Handlers
  var handler = app.createServerHandler('userConfirm');
  handler.addCallbackElement(subjectBox)
         .addCallbackElement(messageBox)
  sendButton.addClickHandler(handler);
  
  
  //Code to display app within spreadsheet
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  spreadsheet.show(app);
}


//Much of the following code was adapted from Google's 'Tutorial: Simple Mail Merge' located at https://developers.google.com/apps-script/articles/mail_merge
//Some minor additions made to make it work within spreadsheet UI

function sendEmails(e) { 
  var parameter = e.parameter; 
  var subjectInput = parameter.subjectBox;
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dataSheet = ss.getSheets()[0];
  var dataRange = dataSheet.getRange(2, 1, dataSheet.getMaxRows() - 1, dataSheet.getLastColumn());

  var templateSheet = ss.getSheets()[1];
  var emailTemplate = templateSheet.getRange("A1").getValue();
  
  var sentEmails = 0;
   var headersRange = dataSheet.getRange(1, 1, 1, dataSheet.getLastColumn());
  // Create one JavaScript object per row of data.
  var objects = getRowsData(dataSheet, dataRange);

  // For every row object, create a personalized email from a template and send
  // it to the appropriate person.
  for (var i = 0; i < objects.length; ++i) {
    // Get a row object
    var rowData = objects[i];

    // Generate a personalized email.
    // Given a template string, replace markers (for instance $%First Name%) with
    // the corresponding value in a row object (for instance rowData.firstName).
    var emailText = fillInTemplateFromObject(emailTemplate, rowData);
    
    
    if (typeof subjectInput === "undefined") {
      Browser.msgBox("There's something wrong with your subject!");
      return;
    }
    else if (typeof emailText === "undefined") {
      Browser.msgBox("Make sure you're separating your variables like this $%header%")
      return;
    }
    else {
      MailApp.sendEmail(rowData.email, subjectInput, emailText);
     
      // Add 2 more columns if they are not already set.
      var cellSentStatus = searchRange(dataSheet, 'Sent status', headersRange);
      if (!cellSentStatus) {
        cellSentStatus = dataSheet.getRange(1, dataSheet.getLastColumn() + 1, 1, 1);
        cellSentStatus.setValue('Sent status');
      }
      var cellSentTimestamp = searchRange(dataSheet, 'Sent timestamp', headersRange);
      if (!cellSentTimestamp) {
        cellSentTimestamp = dataSheet.getRange(1, dataSheet.getLastColumn() + 1, 1, 1);
        cellSentTimestamp.setValue('Sent timestamp');
      }
      
      
      //Populate new status columns with 'email sent' and timestamps and increment sentEmails counter
      dataSheet.getRange(i + 2, dataSheet.getLastColumn() - 1).setValue("Email sent");
      dataSheet.getRange(i + 2, dataSheet.getLastColumn()).setValue(new Date().toString());
      sentEmails++;
    }
  }
  if (sentEmails > 0) {
    ss.toast(sentEmails + ' of ' + objects.length + ' emails were sent', 'Mail Merge', 10);
  }
  else {
    ss.toast('None of ' + objects.length + ' emails were sent as they were sent before.Remove "Email sent" from "Sent status" column to resend.', 'Mail Merge', 10);
  }
  app.close();
  return app;
}


//Prompts user to confirm before sending
function userConfirm(e) {
  var parameter = e.parameter; 
  var messageInput = parameter.messageBox; 
  
  //Saves message box contents to spreadsheet even before send confirmation to allow for editing in case user cancels
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var templateSheet = ss.getSheets()[1];
  templateSheet.clearContents(); 
  templateSheet.appendRow([messageInput]); 
  
  //Confirmation popup, OK sends, cancel brings back to editing
  var confirmMessage = Browser.msgBox("Have you checked for errors? Mistyped variables will NOT populate correctly.\n Click OK to send, and Cancel to continue editing",Browser.Buttons.OK_CANCEL);
  if (confirmMessage === "ok") {
    return sendEmails(e);
  }
  else {
    return beginUI(e);
  }
}


// Replaces markers in a template string with values define in a JavaScript data object.
// Arguments:
//   - template: string containing markers, for instance ${"Column name"}
//   - data: JavaScript object with values to that will replace markers. For instance
//           data.columnName will replace marker ${"Column name"}
// Returns a string without markers. If no data is found to replace a marker, it is
// simply removed.
function fillInTemplateFromObject(template, data) {
  var email = template;
  // Search for all the variables to be replaced, for instance ${"Column name"}
  var templateVars = template.match(/\$\%[^\%]+\%/g);

  // Replace variables from the template with the actual values from the data object.
  // If no value is available, replace with the empty string.
  if (templateVars === null) {
    return email;
  }
  else {
    for (var i = 0; i < templateVars.length; ++i) {
    // normalizeHeader ignores ${"} so we can call it directly here.
    var variableData = data[normalizeHeader(templateVars[i])];
    email = email.replace(templateVars[i], variableData);
    }
    return email;
  }
}





//////////////////////////////////////////////////////////////////////////////////////////
//
// The code below is reused from the 'Reading Spreadsheet data using JavaScript Objects'
// tutorial.
//
//////////////////////////////////////////////////////////////////////////////////////////

// getRowsData iterates row by row in the input range and returns an array of objects.
// Each object contains all the data for a given row, indexed by its normalized column name.
// Arguments:
//   - sheet: the sheet object that contains the data to be processed
//   - range: the exact range of cells where the data is stored
//   - columnHeadersRowIndex: specifies the row number where the column names are stored.
//       This argument is optional and it defaults to the row immediately above range; 
// Returns an Array of objects.
function getRowsData(sheet, range, columnHeadersRowIndex) {
  columnHeadersRowIndex = columnHeadersRowIndex || range.getRowIndex() - 1;
  var numColumns = range.getLastColumn() - range.getColumn() + 1;
  var headersRange = sheet.getRange(columnHeadersRowIndex, range.getColumn(), 1, numColumns);
  var headers = headersRange.getValues()[0];
  return getObjects(range.getValues(), normalizeHeaders(headers));
  Browser.msgBox(getObjects(range.getValues(), normalizeHeaders(headers)));
}

// For every row of data in data, generates an object that contains the data. Names of
// object fields are defined in keys.
// Arguments:
//   - data: JavaScript 2d array
//   - keys: Array of Strings that define the property names for the objects to create
function getObjects(data, keys) {
  var objects = [];
  for (var i = 0; i < data.length; ++i) {
    var object = {};
    var hasData = false;
    for (var j = 0; j < data[i].length; ++j) {
      var cellData = data[i][j];
      if (isCellEmpty(cellData)) {
        continue;
      }
      object[keys[j]] = cellData;
      hasData = true;
    }
    if (hasData) {
      objects.push(object);
    }
  }
  return objects;
}


// Returns an Array of normalized Strings.
// Arguments:
//   - headers: Array of Strings to normalize
function normalizeHeaders(headers) {
  var keys = [];
  for (var i = 0; i < headers.length; ++i) {
    var key = normalizeHeader(headers[i]);
    if (key.length > 0) {
      keys.push(key);
    }
  }
  return keys;
}

// Normalizes a string, by removing all alphanumeric characters and using mixed case
// to separate words. The output will always start with a lower case letter.
// This function is designed to produce JavaScript object property names.
// Arguments:
//   - header: string to normalize
// Examples:
//   "First Name" -> "firstName"
//   "Market Cap (millions) -> "marketCapMillions
//   "1 number at the beginning is ignored" -> "numberAtTheBeginningIsIgnored"
function normalizeHeader(header) {
  var key = "";
  var upperCase = false;
  for (var i = 0; i < header.length; ++i) {
    var letter = header[i];
    if (letter == " " && key.length > 0) {
      upperCase = true;
      continue;
    }
    if (!isAlnum(letter)) {
      continue;
    }
    if (key.length == 0 && isDigit(letter)) {
      continue; // first character must be a letter
    }
    if (upperCase) {
      upperCase = false;
      key += letter.toUpperCase();
    } else {
      key += letter.toLowerCase();
    }
  }
  return key;
}

// Returns true if the cell where cellData was read from is empty.
// Arguments:
//   - cellData: string
function isCellEmpty(cellData) {
  return typeof(cellData) == "string" && cellData == "";
}

// Returns true if the character char is alphabetical, false otherwise.
function isAlnum(char) {
  return char >= 'A' && char <= 'Z' ||
    char >= 'a' && char <= 'z' ||
    isDigit(char);
}

// Returns true if the character char is a digit, false otherwise.
function isDigit(char) {
  return char >= '0' && char <= '9';
}


//Search for value in the range and return range of first occurence
function searchRange(sheet, needle, range) {
  if (!range) {
    range = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn());
  }

  var leftColumn = range.getColumnIndex();
  var rightColumn = range.getLastColumn();
  var topRow = range.getRowIndex();
  var bottomRow = range.getLastRow();

  for (var i = topRow; i <= bottomRow; i++) {
    for (var j = leftColumn; j <= rightColumn; j++) {
      if (needle == sheet.getRange(i, j, 1, 1).getValue()) {
        return sheet.getRange(i, j);
      }
    }
  }

  return false;
}