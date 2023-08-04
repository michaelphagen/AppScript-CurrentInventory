# AppScript-CurrentInventory

Show the latest inventory report from spreadsheets or google forms using Google Apps Script

## How to use

1. Create google forms (titled "XYZ Weekly Inventory") for each space where you want to track inventory.
2. Put all the sheets that are fed by the forms in a folder.
3. Create a new Spreadsheet outside that folder and put this script in it as an Apps Script.
4. Set the folder ID of the folder where the sheets are in the script in Script Properties (in the "Project Settings" menu of the script editor).
5. Re-load the spreadsheet, you will see a new menu item "Inventory".
6. Click "Inventory" -> "Pull Inventory" to update the inventory.
7. Make any changes directly in the spreadsheet, then click "Inventory" -> "Push Inventory" to push the changes to the sheets.