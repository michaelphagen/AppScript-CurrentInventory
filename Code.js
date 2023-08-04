/* exported onOpen */
function onOpen() {
    const ui = SpreadsheetApp.getUi();
    setup();
    ui.createMenu('Inventory')
        .addItem('Pull Inventories', 'pullInventories')
        .addItem('Push Inventories', 'pushInventories')
        .addToUi();
}

function setup() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    // If there isn't a sheet called "Current Inventory", create one
    if (!spreadsheet.getSheetByName("Current Inventory")) {
        spreadsheet.insertSheet("Current Inventory");
    }
    // If there isn't a sheet called "Expected Responses", create one
    if (!spreadsheet.getSheetByName("Expected Responses")) {
        spreadsheet.insertSheet("Expected Responses");
    }
}

function getLatestInventory(sheetID) {
    // Get the latest inventory from the individual inventory sheet and return the latest inventory data
    // Output Format is an Object with the following properties:
    const spreadsheet = SpreadsheetApp.openById(sheetID)
    const sheets = spreadsheet.getSheets();
    if (!sheets.length) {
        Logger.log("No sheets found");
        return;
    }
    const sheet = sheets[0];
    if (sheet.getSheetName() != "Form Responses 1") {
        Logger.log("Sheet name is not Form Responses 1");
        return;
    }
    const lastRowValues = sheet.getRange(sheet.getLastRow(), 1, 1, sheet.getLastColumn()).getValues()[0];
    const firstRowValues = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const [, , ...names] = firstRowValues;
    const [time, email, ...values] = lastRowValues;
    const room = spreadsheet.getName().replace(" Weekly Inventory", "");
    return names.map((n, i) => [
        time, email, room, n, values[i]
    ]);


}

function appendCurrentInventory(latestInventory) {
    // Append the "latest inventory" (output of getLatestInventory) to the Current inventory sheet
    // Starts from the next empty row
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName("Current Inventory");
    const lastRow = sheet.getLastRow();
    const expectedQuantites = getExpectedQuantities();
    const newArr = latestInventory.map((row,index) => {row.push(getAdjustmentRequired(lastRow + index +1 , row, expectedQuantites)); return row;});
    const range = sheet.getRange(lastRow + 1, 1, newArr.length, newArr[0].length);
    range.setValues(newArr);
}

function updateInventory(sheetID) {
    const latestInventory = getLatestInventory(sheetID);
    appendCurrentInventory(latestInventory);
}


// eslint-disable-next-line no-unused-vars
function pullInventories() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName("Current Inventory");
    clearSheet(sheet, ["Timestamp", "User", "Room", "Question", "Response", "Adjustment Required"]);
    const inventories = getInventories();
    const names = Object.keys(inventories);
    names.sort();
    names.forEach((name) => updateInventory(inventories[name]));
    const expectedQuantites = getExpectedQuantities();
    //printDictionary(expectedQuantites);
    const questions = getCurrentInventory();
    for (const room of Object.keys(questions)) {
        for (const name of Object.keys(questions[room])) {
            if (room in expectedQuantites && name in expectedQuantites[room]) {
                questions[room][name] = expectedQuantites[room][name];
            }
            else {
                questions[room][name] = "";
            
            }
        }
    }

    pushExpectedQuantities(questions);
    format();
}

function pushExpectedQuantities(questions) {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName("Expected Responses");
    clearSheet(sheet, ["Room", "Question", "Expected Response"]);
    const values = [];
    for (const room of Object.keys(questions)) {
        for (const name of Object.keys(questions[room])) {
            values.push([room, name, questions[room][name]]);
        }
    }

    sheet.getRange(sheet.getLastRow()+1, 1, values.length, values[0].length).setValues(values);
}

function clearSheet(sheet, header) {
    const firstRow = sheet.getRange(1, 1, 1, sheet.getLastColumn() || 1).getValues();
    if (!firstRow.every((cell, index) => cell === header[index])) {
        sheet.getRange(1, 1, 1, header.length).setValues([header]);
        //Clear Additional columns
        //sheet.getRange(1, header.length + 1, 1, sheet.getLastColumn() - (header.length - 1)).clearContent();
    }
    sheet.getRange(2, 1, sheet.getLastRow(), sheet.getLastColumn()).clearContent();
}

/**
 * 
 * @returns {Record<string,Record<string,string|number>>}
 */
function getCurrentInventory() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName("Current Inventory");
    const [, ...values] = sheet.getDataRange().getDisplayValues();
    const inventories = {}
    for (const [, , room, name, value] of values) {
        if (!inventories[room]) {
            inventories[room] = {}
        }
        inventories[room][name] = value;
    }
    return inventories;
}

function getAdjustmentRequired(rowNumber,row, expectedQuantites) {
    // Three Adjustment types:
    // 1. No Expected Response - Red "No Response Expected" if there IS a response
    // 2. Expected response is number - Red difference if response is < expected response, blue difference if response is > expected response
    // 3. Expected response is text (non list) - Red "Expected Response was: $expected_response" if response is not equal to expected response
    // 4. Expected response is list - Red $items_missing_from_list if response is not equal to expected response
    const [, , room, name, ] = row;
    const expectedValue = "Filter('Expected Responses'!C:C,'Expected Responses'!A:A=C"+rowNumber+",'Expected Responses'!B:B=D"+rowNumber+")";
    const numberComparison = "="+expectedValue+"-E"+rowNumber;
    const listComparison = '=IFNA(Join(", ",Filter(Split(RegexReplace('+expectedValue+',", ",","),","),ArrayFormula(IF(COUNTIF(Split(E'+rowNumber+',","), Split('+expectedValue+',","))>0, False, True)))),"")';
    const stringComparison = '=if(E'+rowNumber+'='+expectedValue+', "", "Expected Response was: "+'+expectedValue+')';
    const noResponseExpected = '=if(E'+rowNumber+'="", "", "No Response Expected")';
    if (!(room in expectedQuantites) || !(name in expectedQuantites[room])) {
        return "";
    }

    if (expectedQuantites[room][name] === "") {
        return noResponseExpected;
    }
    else if (typeof expectedQuantites[room][name] === "number") {
        return numberComparison;
    }
    // If expected response is comma separated list
    else if (expectedQuantites[room][name].includes(",")) {
        return listComparison;
    }
    else {
        return stringComparison;
    }
}
/**
 * 
 * @returns {Record<string,Record<string,string|number>>}
 */
function getExpectedQuantities() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName("Expected Responses");
    const [, ...values] = sheet.getDataRange().getDisplayValues();
    const inventories = {}
    for (const [room, name, value] of values) {
        if (!inventories[room]) {
            inventories[room] = {}
        }
        if (value === "") {
            inventories[room][name] = "";
        }
        else if (!isNaN(value)) {
            inventories[room][name] = Number(value);
        }
        else {
            inventories[room][name] = value;
        }
    }
    //printDictionary(inventories);
    return inventories;
}

function printDictionary(dictionary) {
    for (const key of Object.keys(dictionary)) {
        if (typeof dictionary[key] === "object") {
            printDictionary(dictionary[key]);
        }
        else {
            Logger.log(key + ": " + dictionary[key]);
        }
    }
}


function format() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName("Current Inventory");
    sheet.clearConditionalFormatRules();
    const timeStampRange = sheet.getRange(2, 1, sheet.getLastRow() -1, 1);
    const adjustmentRange = sheet.getRange(2, 6, sheet.getLastRow() -1, 1);
    applyFormatting(sheet, timeStampRange, "=A2>today()-7", "#008000");
    applyFormatting(sheet, timeStampRange, "=A2<today()-7", "#ff0000");
    applyFormatting(sheet, adjustmentRange, "=AND(F2<0,NOT(F2=\"\"))", "#ff0000");
    applyFormatting(sheet, adjustmentRange, "=AND(F2>0,NOT(F2=\"\"))", "#00ffff");
    applyFormatting(sheet, adjustmentRange, "=AND(NOT(F2=\"\"),NOT(ISNUMBER(F2)))", "#ff0000");
}

function applyFormatting(sheet,range, functionString, color) {
    //Apply Formatting Rules
    //  1. If response is positive number, red
    //  2. If response is negative number, blue
    //  3. If response is text
    const rule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(functionString)
    .setBackground(color)
    .setRanges([range])
    .build();
    const rules = sheet.getConditionalFormatRules();
    rules.push(rule);
    sheet.setConditionalFormatRules(rules);
}

/* exported pushInventories */
function pushInventories() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName("Current Inventory");
    const currentInventories = getCurrentInventory(sheet);
    const roomInventories = getInventories();
    const sheetNames = Object.keys(roomInventories);
    for (const sheetName of sheetNames) {
        const roomName = sheetName.replace(" Weekly Inventory", "");
        const sheetID = roomInventories[sheetName];
        const inventory = currentInventories[roomName];
        pushInventory(sheetID, inventory);
    }
}

/**
 * 
 * @param {string} sheetID 
 * @param {Record<string,string|number>} inventory 
 * @returns 
 */
function pushInventory(sheetID, inventory) {
    const spreadsheet = SpreadsheetApp.openById(sheetID)
    const sheets = spreadsheet.getSheets();
    if (!sheets.length) {
        Logger.log("No sheets found");
        return;
    }
    const sheet = sheets[0];
    if (sheet.getSheetName() != "Form Responses 1") {
        Logger.log("Sheet name is not Form Responses 1");
        return;
    }
    const values = sheet.getDataRange().getValues();
    const firstRowValues = values[0];
    const newRoomInv = [new Date(), Session.getActiveUser().getEmail()]
    for (let i = 2; i < firstRowValues.length; i++) {
        newRoomInv.push(inventory[firstRowValues[i]]);
    }
    const [, , ...oldRowValues] = values[values.length -1];
    const [, , ...newRowValues] = newRoomInv;
    if (!oldRowValues.every((v, i) => v == newRowValues[i])) {
        Logger.log("Updating " + spreadsheet.getName());
        const range = sheet.getRange(sheet.getLastRow() + 1, 1, 1, newRoomInv.length);
        range.setValues([newRoomInv]);
    }
}

/**
 * 
 * @returns {Record<string,string>}
 */
function getInventories() {
    // Get the inventory sheet ids from a google drive folder
    // Return a dictionary of name: sheetID
    // Return null if no sheets found
    const folderID = PropertiesService.getScriptProperties().getProperty("INVENTORY_FOLDER_ID");
    if (!folderID) {
        Logger.log("ERROR! No folder ID found");
        Logger.log("Please set the INVENTORY_FOLDER_ID property in the script properties");
    }
    const folder = DriveApp.getFolderById(folderID);
    const files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
    const sheetIDs = {};
    while (files.hasNext()) {
        const file = files.next();
        const sheetID = file.getId();
        const sheetName = file.getName();
        sheetIDs[sheetName] = sheetID;
    }
    return sheetIDs;
}