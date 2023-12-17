/**
 * Represents a reminder system integrated with Notion.
 */
class NotionReminder {
  /**
   * Creates a NotionReminder instance.
   * @param {string} sheetName - The name of the Google Sheets sheet to load configuration from.
   */
  constructor(sheetName) {
    if(!sheetName){
      console.log("Sheet Name is not input in the funciton. Check the script and modify.");
      return;
    }
    let config = this.loadConfiguration(sheetName);
    this.recipients = config['Reminder Recipient(s)'];
    this.apiKey = config['Integration API Key'];
    this.tableUrl = config['Target Table URL'];
    this.tableName = config['Target Table Name'];
    this.titleProperty = config['Title Property Name'];
    this.remindDateProperty = config['Remind Date Property Name'];
    this.completeCheckProperty = config['Complete Check Property Name'];
    this.themeColor = config['Theme Color'];
    this.ss = config['Spreadsheet'];
    this.sheet = config['Sheet'];
    this.sheetUrl = config['Sheet URL'];
    this.sheetName = config['Sheet Name'];
    this.themeColorHexCode = this.getThemeColorHexCode();
    this.sheet.setTabColor(this.themeColorHexCode);

    this.headers = {
      "Authorization": "Bearer " + this.apiKey,
      "Notion-Version": "2022-06-28",
      'Content-Type': 'application/json'
    };
    
  }

  /**
   * Checks if the sheet name provided during the test reminder setup is valid.
   * @returns {boolean} - Returns true if the sheet name is valid, otherwise false.
   */
  sheetNameCheckTestReminder(){
    let sheetNameCheckChoice = Browser.msgBox(`FOR TEST REMINDER: "${this.sheetName}" is the target sheet that includes the necessary information for Test Reminder.`,Browser.Buttons.YES_NO);
    if (sheetNameCheckChoice === 'no'){
      Browser.msgBox("Please go to the target sheet and do click 'Test Reminder' from the Custom Menu again.");
      return false;
    }
    return true;
  }
  /**
   * Loads configuration data from the specified sheet.
   * @param {string} sheetName - The name of the sheet to load configuration from.
   * @returns {Object} - An object containing configuration data.
   */
  loadConfiguration(sheetName) {
    let ss = SpreadsheetApp.getActiveSpreadsheet();
    let ssUrl = ss.getUrl();
    let sheet = ss.getSheetByName(sheetName);

    if(!sheet){
      Browser.msgBox(`${sheetName} does not exist. Enter the valid sheet name.`);
      return;
    }
    let sheetId = sheet.getSheetId;
    let sheetUrl = ssUrl+"#gid="+sheetId;
    let data = sheet.getRange("B1:B9").getValues();

    return {
      'Reminder Recipient(s)': data[1][0],
      'Integration API Key': data[2][0],
      'Target Table URL': data[3][0],
      'Target Table Name': data[4][0],
      'Title Property Name': data[5][0],
      'Remind Date Property Name': data[6][0],
      'Complete Check Property Name':data[7][0],
      'Theme Color':data[8][0],
      'Spreadsheet': ss,
      'Sheet': sheet,
      'Sheet URL':sheetUrl,
      'Sheet Name': sheetName
    };

  }
  /**
   * Retrieves the hexadecimal code of the theme color from the color list sheet.
   * @returns {string} - The hexadecimal code of the theme color.
   */
  getThemeColorHexCode(){
    // Assuming the lookup table is in a sheet called "ColorLookup"
    let colorLookupSheet = this.ss.getSheetByName("color-list");
    let colorLookupRange = colorLookupSheet.getRange("A2:B" + colorLookupSheet.getLastRow());
    let colorLookupValues = colorLookupRange.getValues();
    
    let hexCode = ""; // Default to an empty string if color name is not found
    for (let i = 0; i < colorLookupValues.length; i++) {
      if (colorLookupValues[i][0].toLowerCase() === this.themeColor.toLowerCase()) {
        hexCode = colorLookupValues[i][1]; // Found the hex code for the color name
        break;
      }
    }

    // Set the tab color using the hex code
    if (hexCode) {
      return hexCode;
    } else {
      console.log(`The color name "${this.themeColor}" does not have a corresponding hex code.`)
    }
  }
  /**
   * Extracts the Notion table ID from a given URL.
   * @param {string} url - The Notion URL to extract the table ID from.
   * @returns {string|null} - The extracted table ID, or null if not found.
   */
  // Extract table ID from URL
  extractNotionTableIdFromUrl(url) {
    let regex = /https:\/\/www\.notion\.so\/(?:[a-zA-Z0-9]+\/)?([a-zA-Z0-9-]+)/;
    let match = url.match(regex);
    return match ? match[1] : null;
  }
  /**
   * Fetches to-do data from the Notion API.
   * @returns {Object} - An object containing arrays of overdue and throughout to-do items.
   */
  getTodoData() {
    let payload;
    let options;
    if(this.completeCheckProperty !== "Property does not exist." && this.apiKey){
      //Payload for the HTTP request to the Notion API
      payload = {
        "filter": {
          "property": this.completeCheckProperty,
          "checkbox": {
            "equals":false
          }
        },
        "sorts": [{
          "property": this.remindDateProperty,
          "direction": "ascending"
        }]
      };
      
      //
      options = {
        "method": "post",
        "headers": this.headers,
        "payload": JSON.stringify(payload)
      };
    } else if (this.completeCheckProperty === "Property does not exist." && this.apiKey){
      //Payload for the HTTP request to the Notion API
      payload = {
        "sorts": [{
          "property": this.remindDateProperty,
          "direction": "ascending"
        }]
      };
      
      //
      options = {
        "method": "post",
        "headers": this.headers,
        "payload": JSON.stringify(payload)
      };
    }

    let tableId = this.extractNotionTableIdFromUrl(this.tableUrl);
    let queryUrl = `https://api.notion.com/v1/databases/${tableId}/query`;

    //Fetches a URL using optional advanced parameters; payload is expected as string
    let response = UrlFetchApp.fetch(queryUrl, options);
    let data = JSON.parse(response.getContentText());
    let results = data.results;

    let today = new Date();
    today = Utilities.formatDate(today, 'Asia/Tokyo', 'yyyy/MM/dd');

    let todoOverDue = [];
    let todoThroughout = [];
    let preDefinedValuesError = 0;

    //loop through the results and add the items to the todo_list if the remind_date is before today
    for (let i = 0; i < results.length; i++) {
      let record = results[i];
      let todoTitle;
      let remindDate;
      // Check if the title property exists
      if (record.properties[this.titleProperty]){
        if(record.properties[this.titleProperty].title) {
          todoTitle = record.properties[this.titleProperty].title[0].plain_text;
        } else {
          todoTitle = 'The title of the record is empty.';
        }
      } else {
        preDefinedValuesError++;
      }

      let todoUrl = record.url;

      if (record.properties[this.remindDateProperty]){
        if(record.properties[this.remindDateProperty].date) {
          remindDate = record.properties[this.remindDateProperty].date[0];
          remindDate = new Date(record.properties[this.remindDateProperty].date.start);
          remindDate = Utilities.formatDate(remindDate, 'Asia/Tokyo', 'yyyy/MM/dd');
        } else {
          remindDate = 'No date specified';
          console.log(`${todoTitle} does not have a remind date`);
        }
      } else {
        preDefinedValuesError++;
      }

      //if the remindDate is before today, add it to the list of items
      if (remindDate === 'No date specified'){
        let itemThroughout = {
          'Todo Title':todoTitle,
          'Remind Date':remindDate,
          'Todo Item URL':todoUrl,
        }
        todoThroughout.push(itemThroughout);
      } else if (remindDate <= today) {
        let itemOverdue = {
          'Todo Title':todoTitle,
          'Remind Date':remindDate,
          'Todo Item URL':todoUrl,
        }
        todoOverDue.push(itemOverdue);
      } else if (remindDate > today) {
        // console.log(`"${todoTitle}" will NOT be reminded this time, later`);
      }
    }

    if(preDefinedValuesError>0){
      let subject = `Notion Reminder Error: Inconsistency between ${this.sheetName} and ${this.tableName}`
      let body = `Some of the predefined values in ${this.sheetName} in Spreadsheet are not consistent with ${this.tableName} in Notion. Check from the links below.`;
      let options = {
        'htmlBody': body
      }
      GmailApp.sendEmail(this.recipients,subject,"",options);
      return;
    }

    return {
      'Todo Overdue': todoOverDue,
      'Todo Throughout': todoThroughout,
    }
  }
  /**
   * Sends a reminder email with to-do data or a test reminder.
   * @param {string} testCheck - Indicates whether this is a test reminder.
   */
  sendReminder(testCheck) {
    let template = HtmlService.createTemplateFromFile('reminder-email');
    let subject;
    if(this.apiKey === ''){
      if(testCheck === "test"){
          template.tableName = this.tableName;
          template.tableUrl = this.tableUrl;
          template.themeColor = this.themeColorHexCode;
          template.remindDatePresence = 'none';
          template.testCheck = testCheck;
          template.setInstructionSlideUrl = "https://docs.google.com/presentation/d/1QYZmRGCpvDynGNwytI1fn8Otty9uGgCT_erF84lc3Zw/edit#slide=id.p";
          subject = `Test Reminder for ${this.tableName}`;

      } else {
          template.tableName = this.tableName;
          template.tableUrl = this.tableUrl;
          template.themeColor = this.themeColorHexCode;
          template.remindDatePresence = 'none';
          template.testCheck = testCheck;
          template.setInstructionSlideUrl = "https://docs.google.com/presentation/d/1QYZmRGCpvDynGNwytI1fn8Otty9uGgCT_erF84lc3Zw/edit#slide=id.p";
          subject = `Reminder for ${this.tableName}`;
      }
    } else {
      let todoData = this.getTodoData();
      if(testCheck === "test"){
          template.tableName = this.tableName;
          template.tableUrl = this.tableUrl;
          template.themeColor = this.themeColorHexCode;
          template.remindDatePresence = 'present';
          template.todoOverdue = todoData['Todo Overdue'];
          template.todoThroughout = todoData['Todo Throughout'];
          template.testCheck = testCheck;
          template.setInstructionSlideUrl = "https://docs.google.com/presentation/d/1QYZmRGCpvDynGNwytI1fn8Otty9uGgCT_erF84lc3Zw/edit#slide=id.p";
          subject = `Test Reminder for ${this.tableName}`;

      } else {
          if(todoData['Todo Overdue'].length > 0){
              template.tableName = this.tableName;
              template.tableUrl = this.tableUrl;
              template.themeColor = this.themeColorHexCode;
              template.remindDatePresence = 'present';
              template.todoOverdue = todoData['Todo Overdue'];
              template.todoThroughout = todoData['Todo Throughout'];
              template.testCheck = testCheck;
              subject = `Reminder for ${this.tableName}`;
          } else {
            return;
          }
      }
    }
    let body = template.evaluate().getContent();
    GmailApp.sendEmail(this.recipients, subject, "", {
        htmlBody: body,
    });
  }
}

/**
 * Called when the spreadsheet is opened. Adds custom menu items.
 */
function onOpen() {
  let ui = SpreadsheetApp.getUi();
  ui.createMenu('Notion Reminder')
    .addItem('Conduct GAS Authorization', 'showAuthorization')
    .addSeparator()
    .addItem('Create New Sheet', 'showCreateNewSheetModal')
    .addSeparator()
    .addItem('Test Reminder', 'testReminder')
    .addSeparator()
    .addItem('Update Index Sheet','updateIndexSheet')
    .addToUi();
}

/**
 * Prompts Google Apps Script authorization.
 */
// Used to prompt Google Apps Script authorization
function showAuthorization() {
  SpreadsheetApp;
  GmailApp;
}

/**
 * Displays a modal dialog for creating a new sheet.
 */
function showCreateNewSheetModal() {
  let colorList = getColorList_(); // Get the color list
  let template = HtmlService.createTemplateFromFile('create-sheet-modal');
  template.colorList = JSON.stringify(colorList); // Pass the color list to the template
  let html = template.evaluate()
    .setWidth(600)
    .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, 'Create New Sheet');
}

/**
 * Retrieves a list of color names and their corresponding hexadecimal codes.
 * @returns {Array} - An array of objects containing color names and hex codes.
 */
function getColorList_() {
  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let colorListSheet = spreadsheet.getSheetByName('color-list');
  let colors = colorListSheet.getRange('A2:B' + colorListSheet.getLastRow()).getValues();
  return colors.map(function(row) { return { name: row[0], hex: row[1] }; });
}

/**
 * Creates a new sheet with specified settings.
 * @param {string} sheetName - Name of the new sheet.
 * @param {string} tableType - Type of the table (Type 1, Type 2, or Type 3).
 * @param {string} themeColor - Hexadecimal color code for the theme.
 */
function createSheetWithSettings(sheetName, tableType, themeColor) {
    let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    // Check if the sheet already exists
    if (spreadsheet.getSheetByName(sheetName) !== null) {
      throw new Error('A sheet with this name already exists. Please choose a different name.');
    }

    // Create the new sheet based on the template
    let templateSheet = spreadsheet.getSheetByName('template');
    if (!templateSheet) throw new Error('Template sheet not found');

    let newSheet = spreadsheet.insertSheet(sheetName, {template: templateSheet});
    Logger.log("New sheet created: " + newSheet.getName());
    let [colorName, colorHex] = themeColor.split('|');    
    newSheet.setTabColor(colorHex);

    // Apply modifications based on tableType
    if (tableType === 'type1') {
        // Type 1: Set the color name in B9
        newSheet.getRange('B9').setValue(colorName);
    } else if (tableType === 'type2') {
        // Type 2: Set B8 as "Property does not exist." and color in light gray 1 and set the color name in B9
        newSheet.getRange('B8').setValue('Property does not exist.').setBackground('#D3D3D3');
        newSheet.getRange('B9').setValue(colorName);
    } else if (tableType === 'type3') {
        // Type 3: Set B3, B6, B7, and B8 light gray 1 and set B8 as "Property does not exist." and set the color name in B9
        newSheet.getRange('B3').setBackground('#D3D3D3');
        newSheet.getRange('B6:B7').setBackground('#D3D3D3');
        newSheet.getRange('B8').setValue('Property does not exist.').setBackground('#D3D3D3');
        newSheet.getRange('B9').setValue(colorName);
    }

    Browser.msgBox("New Sheet was successfully created.");
}


/**
 * Updates the index sheet with a list of sheets and their descriptions.
 */
function updateIndexSheet() {
  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheets = spreadsheet.getSheets();
  let sheetNum = sheets.length;
  let indexSheet = spreadsheet.getSheetByName("index");
  if (!indexSheet) {
    indexSheet = spreadsheet.insertSheet("index");
  }
  indexSheet.clear();
  let headers = ["Sheet Name", "Description"];
  indexSheet.appendRow(headers);
  indexSheet.getRange("A1:B1").setBackground('#D3D3D3').setFontWeight("bold");
  indexSheet.setFrozenRows(1);

  sheets.forEach((sheet, index) => {
    let sheetName = sheet.getName();
    if (sheetName === "index" || sheetName === "color-list" || sheetName === "template") return;
    let sheetUrl = spreadsheet.getUrl() + "#gid=" + sheet.getSheetId();
    let note = sheet.getRange("B1").getValue();
    indexSheet.appendRow([`=HYPERLINK("${sheetUrl}", "${sheetName}")`, note]);
  });
  indexSheet.getRange(1,1,sheetNum-2,2).setBorder(true,true,true,true,true,true);
}
/**
 * Sends a test reminder and displays a message box upon completion.
 */
function testReminder() {
  let sheetName = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
  let testReminder = new NotionReminder(sheetName);
  if (testReminder.sheetNameCheckTestReminder() === false) {
    return;
  }
  testReminder.sendReminder('test');
  Browser.msgBox('Test reminder was sent. Check the reminder email in Gmail. If the contets are fine, make a function and set a trigger for the funciton by yourself.');
}

