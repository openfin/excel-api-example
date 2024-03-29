# OpenFin Service API Documentation

The Excel API is composition-based object model. Where Excel is the top most level that has workbooks that have worksheets and worksheets have cells.
To use the API, you will need to include `ExcelAPI.js` in your project and it will extend Openfin API with Excel API included.
Once included, you will be able to use following API calls.

## fin.desktop.ExcelService

Represents the helper service which manages OpenFin connections to running instances of Excel.

### Properties

```javascript
connected: Boolean // indicates that OpenFin is connected to the helper service
initialized: Boolean // indicates that the current window is subscribed to Excel service events
```

### Functions

```javascript
/*
init();
Returns a promise which resolves when the Excel helper service is running and initialized.
*/

await fin.desktop.ExcelService.init();
```

## fin.desktop.Excel:

Represents a single instance of an Excel application.

### Functions

```javascript

/*
getWorkbooks();
Returns a promise which resolves the currently opened workbooks from Excel.
*/

var workbooks = await fin.desktop.Excel.getWorkbooks();

/*
addWorkbook();
Asynchronously creates a new workbook in Excel and returns a promise which resolves the newly added workbook.
*/

var workbook = await fin.desktop.Excel.addWorkbook();

/*
openWorkbook(path);
Asynchronously opens workbook from the specified path and returns a promise which resolves the opened workbook.
*/

var workbook = await fin.desktop.Excel.openWorkbook(path);

/*
getWorkbookByName(name);
Synchronously returns workbook object that represents the workbook with supplied name.

*/

var workbook = fin.desktop.Excel.getWorkbookByName("Workbook1");

/*
getConnectionStatus();
Returns a promise which resolves the connection status of the current Excel application instance.
*/

var isConnected = await fin.desktop.Excel.getConnectionStatus();

/*
getCalculationMode();
Returns a promise which resolves the calculation mode information.
*/

var info = await fin.desktop.Excel.getCalculationMode();

/*
calculateAll();
Asynchronously forces calculation on all sheets
*/

await fin.desktop.Excel.calculateAll();

/*
addEventListener(type, listener);
Adds event handler to handle events from Excel
*/

fin.desktop.Excel.addEventListener("workbookAdded", function(event){...})

/*
removeEventListener(type, listener);
removes an attached event handler from Excel
*/

removeEventListener("workbookAdded", handler);
```

### Events

```javascript
{type: "connected"};
// is fired when excel connects to Openfin.
//Example:
fin.desktop.Excel.addEventListener("connected", function(){ console.log("Connected to Excel"); })

{type: "workbookAdded", workbook: ExcelWorkbook};
//is fired when a new workbook is added in excel (this includes adding workbooks using API).
//Example:
fin.desktop.Excel.addEventListener("workbookAdded",
function(event){
    console.log("New workbook added; Name:", event.workbook.name);
});

{type: "workbookClosed", workbook: ExcelWorkbook};
//is fired when a workbook is closed.
//Example:
fin.desktop.Excel.addEventListener("workbookClosed",
function(event){
    console.log("Workbook closed; Name:", event.workbook.name);
});

{type: "afterCalculation"};
//is fired when calculation is complete on any sheet.
//Example:
fin.desktop.Excel.addEventListener("afterCalculation",
function(event){
    console.log("calculation is complete.";
});
```

## fin.desktop.ExcelWorkbook:

Represents an Excel workbook.

Note: New workbooks are not supposed to be created using new or `Object.create()`.
Workbook objects can only be retrieved using API calls like `fin.desktop.Excel.getWorkbooks()`, `fin.desktop.Excel.getWorkbookByName()`,  and `fin.desktop.Excel.addWorkbook()`, etc.

### Properties

```
name: String // name of the workbook that the object represents.
```

### Functions

```javascript
/*
getWorksheets();
Returns a promise which resolves an array of worksheets in the current workbook.
*/

var workbook = fin.desktop.Excel.getWorkbookByName("Workbook1");
var worksheets = await workbook.getWorksheets();

/*
getWorksheetByName(name);
Synchronously returns worksheet object that represents the worksheet with supplied name.
*/

var workbook = fin.desktop.Excel.getWorkbookByName("Workbook1");
var sheet = workbook.getWorksheetByName("sheet1");

/*
addWorksheet();
Asynchronously creates a new worksheet in the workbook and returns a promise which resolves the newly added worksheet.
*/

var workbook = fin.desktop.Excel.getWorkbookByName("Workbook1");
var worksheet = await workbook.addWorksheet();

/*
activate();
Returns a promise which resolves when the workbook is activated
*/

var workbook = fin.desktop.Excel.getWorkbookByName("Workbook1");
await workbook.activate();

/*
save();
Asynchronously saves changes to the current workbook and returns a promise which resolves when the operation is complete.
*/

var workbook = fin.desktop.Excel.getWorkbookByName("Workbook1");
await workbook.save();

/*
close();
Asynchronously closes the current workbook and returns a promise which resolves when the operation is complete.
*/

var workbook = fin.desktop.Excel.getWorkbookByName("Workbook1");
await workbook.close();
```

### Events

```javascript
{type: "sheetAdded", target: ExcelWorkbook, worksheet: ExcelWorksheet};
//fired when a new sheet is added to the workbook
//Example:
var workbook = fin.desktop.Excel.getWorkbookByName("Workbook1");
workbook.addEventListener("sheetAdded",
function(event){
    console.log("New sheet", event.worksheet.name, "was added to the workbook", event.worksheet.workbook.name)
});

{type: "sheetRemoved", target: ExcelWorkbook, worksheet: ExcelWorksheet};
//fired when a sheet is closed/removed
//Example:
var workbook = fin.desktop.Excel.getWorkbookByName("Workbook1");
workbook.addEventListener("sheetRemoved",
function(event){
    console.log("Sheet", event.worksheet.name, "was removed from workbook", event.worksheet.workbook.name)
});

{type: "workbookActivated", target: ExcelWorkbook};
//fired when a workbook is activated/focused
//Example:
var workbook = fin.desktop.Excel.getWorkbookByName("Workbook1");
workbook.addEventListener("workbookActivated",
function(event){
    console.log("Workbook", event.target.name, "was activated");
});

{type: "workbookDeactivated", target: ExcelWorkbook};
//fired when a workbook is deactivated/blurred
//Example:
var workbook = fin.desktop.Excel.getWorkbookByName("Workbook1");
workbook.addEventListener("workbookDeactivated",
function(event){
    console.log("Workbook", event.target.name, "was deactivated");
});

```

## fin.desktop.ExcelWorksheet:

Represents a worksheet in Excel.
Note: New sheets are not supposed to be created using `new` or `Object.create()`.
new sheets can be created only using workbook.addWorksheet() or existing sheet objects can be retrieved using `workbook.getWorksheets()`  and `workbook.getWorksheetByName();`

### Properties

```
name: String // name of the worksheet
workbook: fin.desktop.ExcelWorkbook // workbook object that worksheet belongs to.
```

### Functions

```javascript
/*
setCells(values, offset);
Asynchronously populates the cells with the values starting from the provided cell reference and returns a promise which resolves when the operation is complete.
*/

var worksheet = await workbook.addWorksheet();
await worksheet.setCells([["a", "b", "c"], [1, 2, 3]], "A1");

/*
setFilter(start, offsetWidth, offsetHeight, field, criteria1[, operator, criteria2, visibleDropDown]);
Asynchronously sets a filter on selected range in a worksheet and returns a promise which resolves when the operation is complete.

arguments:
start: starting address of the range. e.g "A1"
offsetWidth: width of the range.
offsetHeight: height of the range.
field: integer representing the field number. starts with 1.
criteria1: The criteria (a string; for example, "101"). Use "=" to find blank fields, or use "<>" to find nonblank fields. If this argument is omitted, the criteria is All. If Operator is xlTop10Items, Criteria1 specifies the number of items (for example, "10").
operator: Optional. Can be one of the following:
          and
          bottom10items
          bottom10percent
          or
          top10items
          top10percent
criteria2: Optional. The second criteria (a string). Used with Criteria1 and Operator to construct compound criteria.
visibleDropDown: Optional. true to display the AutoFilter drop-down arrow for the filtered field; false to hide the AutoFilter drop-down arrow for the filtered field. true by default.
*/

var worksheet = workbook.getSheetByName("sheet1");
await worksheet.setCells([["Column1", "Column2"], ["TRUE", "1"], ["TRUE", "2"],["FALSE", ""]], "A1");
await worksheet.setFilter("A1", 2, 4, 1, "TRUE");

/*
getCells(start, offsetWidth, offsetHeight);
Returns a promise which resolves a two dimensional array of cell values starting at the specified cell reference and the specified width and length.
*/

var sheet = workbook.getSheetByName("sheet1");
var cells = await sheet.getCells("A1", 3, 2); // cells: [[{value: --, formula: --}, ...], ...]

/*
activate();
Asynchronously activates the worksheet and returns a promise which resolves when the operation is complete.
*/

var sheet = workbook.getSheetByName("sheet1");
await sheet.activate();

/*
activateCell(cellReference);
Asynchronously selects specified cell reference and returns a promise which resolves when the operation is complete.
*/

var sheet = workbook.getSheetByName("sheet1");
await sheet.activateCell("A1");

/*
clearAllCells();
Asynchronously clears all the cell values and formatting in the worksheet and returns a promise which resolves when the operation is complete.
*/

var sheet = workbook.getSheetByName("sheet1");
await sheet.clearAllCells();

/*
clearAllCellContents();
Asynchronously clears all the cell values in the worksheet and returns a promise which resolves when the operation is complete.
*/

var sheet = workbook.getSheetByName("sheet1");
await sheet.clearAllCellContents();

/*
clearAllCellFormats();
Asynchronously clears all the cell formatting in the worksheet and returns a promise which resolves when the operation is complete.
*/

var sheet = workbook.getSheetByName("sheet1");
await sheet.clearAllCellFormats();

/*
clearRange();
Asynchronously clears all the cell values and formatting in the specified range and returns a promise which resolves when the operation is complete.
*/

var sheet = workbook.getSheetByName("sheet1");
await sheet.clearRange();

/*
clearRangeContents();
Asynchronously clears all the cell values in the specified range and returns a promise which resolves when the operation is complete.
*/

var sheet = workbook.getSheetByName("sheet1");
await sheet.clearRangeContents();

/*
clearRangeFormats();
Asynchronously clears all the cell formatting in the specified range and returns a promise which resolves when the operation is complete.
*/

var sheet = workbook.getSheetByName("sheet1");
await sheet.clearRangeFormats();

/*
setCellName(cellAddress, cellName);
Asynchronously sets a name for the cell which can be referenced to get values or in formulas and returns a promise which resolves when the operation is complete.
*/

var sheet = workbook.getSheetByName("sheet1");
await sheet.setCellName("A1", "TheCellName");

/*
calculate();
Asynchronously forces calculation on the sheet and returns a promise which resolves when the operation is complete.
*/

var sheet = workbook.getSheetByName("sheet1");
await sheet.calculate();

/*
getCellByName(name);
Returns a promise which resolves the cell info of the cell with the name provided.
*/

var sheet = workbook.getSheetByName("sheet1");
var cellInfo = await sheet.getCellByName("TheCellName");

/*
protect();
Asynchronously makes all cells in the sheet read only, except the ones marked as locked:false and returns a promise which resolves when the operation is complete.
*/

var sheet = workbook.getSheetByName("sheet1");
await sheet.protect();

/*
formatRange(rangeCode, format);
Asynchronously formats the specified range and returns a promise which resolves when the operation is complete.
*/

var sheet = workbook.getSheetByName("sheet1");
await sheet.formatRange("A1:E:10", {
                                    border: {color:"0,0,0,1", style: "continuous"}, //dash, dashDot, dashDotDot, dot, double, none, slantDashDot
                                    border-right: {color:"0,0,0,1", style: "continuous"},
                                    border-left: {color:"0,0,0,1", style: "continuous"},
                                    border-top: {color:"0,0,0,1", style: "continuous"},
                                    border-bottom: {color:"0,0,0,1", style: "continuous"},
                                    horizontalLines: {color:"255,255,255,1", style: "none"}, // horizontal lines between cells
                                    verticalLines: {color:"255,255,255,1", style: "none"}, // vertical lines between cell rows
                                    font: {color: "100,100,100,1", size: 12, bold: true, italic: true, name: "Verdana"},
                                    mergeCells: true, // merges the given range into one big cell
                                    shrinkToFit: true, // the text will shrink to fit the cell
                                    locked: false // specifies if the cell is readonly or not in protect mode, default is true
                                });
```

### Events

```javascript
{type: "sheetChanged", target: ExcelWorksheet,  data: {column: int, row: int, formula: String, sheetName: String, value:String}};
//fired when any cell value in the sheet has changed.
//Example:
var sheet = workbook.getSheetByName("sheet1");
sheet.addEventListener("sheetChanged",
function(event){
    console.log("sheet values were changed. column:", event.data.column, "row:", event.data.row, "value:", event.data.value, "formula", event.data.formula);
});

{type: "selectionChanged", target: ExcelWorksheet, data: {column: int, row: int, value: String}};
//fired when a selection on the sheet has changed.
//Example:
var sheet = workbook.getSheetByName("sheet1");
sheet.addEventListener("selectionChanged",
function(event){
    console.log("sheet selection was changed. column:", event.data.column, "row:", event.data.row, "value:", event.data.value);
});

{type: "sheetActivated", target: ExcelWorksheet};
//fired when the sheet gets into focus.
//Example:
var sheet = workbook.getSheetByName("sheet1");
sheet.addEventListener("sheetActivated",
function(event){
    console.log("sheet activated. Sheet", event.target.name, "Workbook:", event.target.workbook.name);
});

{type: "sheetDeactivated", target: ExcelWorksheet};
//fired when the sheet gets out of focus due to a different sheet getting in focus.
//Example:
var sheet = workbook.getSheetByName("sheet1");
sheet.addEventListener("sheetDeactivated",
function(event){
    console.log("sheet deactivated. Sheet", event.target.name, "Workbook:", event.target.workbook.name);
});

{type: "sheetRenamed", target: ExcelWorksheet};
//fired when the sheet is renamed.
//Example:
var sheet = workbook.getSheetByName("sheet1");
sheet.addEventListener("sheetRenamed",
function(event){
    console.log("sheet", event.data.sheetName, "was renamed to: ", event.data.newName
    );
});

```

### Excel RTD Functionality (New in 3.1.1) 

You can create an RTD Connection in JavaScript:

```javascript
const rtd = await fin.desktop.ExcelService.createRtd('Provider1');

// you can then listen to when an excel sheet connects by adding the following
// example to a cell: =OpenFinRTD("Provider1", "Topic1")
rtd.addEventListener('connected', console.dir);

// the connected event will included the topic id that has been assigned

// Your application can listen to the disconnected event
// The event will have topic: "Topic1", type: "disconnected" and your application can react.
rtd.addEventListener('disconnected', console.dir);

// if the webapp refreshes or is closed and loses connection then the excel sheet wipes away the value with #N/A to ensure no stale data is shown.

// You can push values to excel cells as follows
rtd.setValue('Topic1', 'Hello');
rtd.setValue('Topic1', true);
rtd.setValue('Topic1', 123.55);

// with this functionality you can listen to multiple cells (each with a specific topic) and return a single value or apply updates on an interval as the application value changes.

```

### Excel RTD Functionality (New in 4.0.0)

From version 4.0 you can dispose the rtd instance that you created. You should call dispose if your application is closing or if you no longer need the functionality.

```javascript
const rtd = await fin.desktop.ExcelService.createRtd('Provider1');

// Will stop listening to excel and will tell excel to clear all values in the connected topics
// When the excel provider service closes we will clear the cells as well but that won't happen if another application 
// is still using the excel service so you should always try and tidy up and dispose when your application is about to close. 
 await rtd.dispose();

```

### Excel Self Hosting

You may wish to control the version of the excel service you are using and self host it.

This would involve having your own server to host the assets.

An example can be seen in our demo by looking at the **dev-app.json** file within the /demo folder.

#### Scenario: You are hosting the assets in a versioned excel folder on your server https://yourserver/excel/x.x.x/ (x.x.x represents the version you are using throughout this example)

The excel setup is made up of the following:

1. Your app references a service through it's manifest (and this is where you can specify your own service manifest url):  
  
```javascript
"services": [
   {
     "name": "excel",
     "manifestUrl": "https://yourserver/excel/x.x.x/provider/app.json"
   }
 ]
```
2. Your app references adds a script tag (or bundles) the excel client script e.g. 

```javascript
<script src="https://yourserver/excel/x.x.x/client/fin.desktop.Excel.js"></script>
```

3. The provider folder that is hosted by you has an updated app.json file to reflect your host:

```javascript
{
  "licenseKey": "Your License Key",
  "startup_app": {
    "name": "Excel-Service-Manager",
    "url": "https://yourserver/excel/x.x.x/provider/provider.html",
    "uuid": "excel-service-manager",
    "permissions": {
      "System": {
        "launchExternalProcess": true
      }
    }
  },
  "runtime": {
    "arguments": "",
    "version": "14.78.48.16"
  },
  "appAssets": [
    {
      "src": "https://yourserver/excel/x.x.x/provider/add-in.zip",
      "alias": "excel-api-addin",
      "version": "x.x.x"
    }
  ]
}
```

#### Why do this?

This is to control the version of the scripts/service you are using (rather than pointing to the latest version).

#### Anything to be aware of?

The excel service acts as a singleton so the first version that is spun up is the active version (for scenarios where you have multiple apps using excel).

#### What if I want more features/control over my excel integration?

OpenFin have a .Net Adapter that allows you to launch OpenFin applications/windows as well as communicate on the OpenFin Message Bus. This in combination with ExcelDNA will give you the two pieces you need to create your own custom experience.

* .NET Adapter: https://developers.openfin.co/docs/net-api
* Excel DNA: https://excel-dna.net/
