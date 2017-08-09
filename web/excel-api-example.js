/******/ (function(modules) { // webpackBootstrap
/******/ 	// The module cache
/******/ 	var installedModules = {};
/******/
/******/ 	// The require function
/******/ 	function __webpack_require__(moduleId) {
/******/
/******/ 		// Check if module is in cache
/******/ 		if(installedModules[moduleId]) {
/******/ 			return installedModules[moduleId].exports;
/******/ 		}
/******/ 		// Create a new module (and put it into the cache)
/******/ 		var module = installedModules[moduleId] = {
/******/ 			i: moduleId,
/******/ 			l: false,
/******/ 			exports: {}
/******/ 		};
/******/
/******/ 		// Execute the module function
/******/ 		modules[moduleId].call(module.exports, module, module.exports, __webpack_require__);
/******/
/******/ 		// Flag the module as loaded
/******/ 		module.l = true;
/******/
/******/ 		// Return the exports of the module
/******/ 		return module.exports;
/******/ 	}
/******/
/******/
/******/ 	// expose the modules object (__webpack_modules__)
/******/ 	__webpack_require__.m = modules;
/******/
/******/ 	// expose the module cache
/******/ 	__webpack_require__.c = installedModules;
/******/
/******/ 	// identity function for calling harmony imports with the correct context
/******/ 	__webpack_require__.i = function(value) { return value; };
/******/
/******/ 	// define getter function for harmony exports
/******/ 	__webpack_require__.d = function(exports, name, getter) {
/******/ 		if(!__webpack_require__.o(exports, name)) {
/******/ 			Object.defineProperty(exports, name, {
/******/ 				configurable: false,
/******/ 				enumerable: true,
/******/ 				get: getter
/******/ 			});
/******/ 		}
/******/ 	};
/******/
/******/ 	// getDefaultExport function for compatibility with non-harmony modules
/******/ 	__webpack_require__.n = function(module) {
/******/ 		var getter = module && module.__esModule ?
/******/ 			function getDefault() { return module['default']; } :
/******/ 			function getModuleExports() { return module; };
/******/ 		__webpack_require__.d(getter, 'a', getter);
/******/ 		return getter;
/******/ 	};
/******/
/******/ 	// Object.prototype.hasOwnProperty.call
/******/ 	__webpack_require__.o = function(object, property) { return Object.prototype.hasOwnProperty.call(object, property); };
/******/
/******/ 	// __webpack_public_path__
/******/ 	__webpack_require__.p = "";
/******/
/******/ 	// Load entry module and return exports
/******/ 	return __webpack_require__(__webpack_require__.s = 4);
/******/ })
/************************************************************************/
/******/ ([
/* 0 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

class RpcDispatcher {
    constructor() {
        this.listeners = {};
    }
    addEventListener(type, listener) {
        if (this.hasEventListener(type, listener)) {
            return;
        }
        if (!this.listeners[type]) {
            this.listeners[type] = [];
        }
        this.listeners[type].push(listener);
    }
    removeEventListener(type, listener) {
        if (!this.hasEventListener(type, listener)) {
            return;
        }
        var callbacksOfType = this.listeners[type];
        callbacksOfType.splice(callbacksOfType.indexOf(listener), 1);
    }
    hasEventListener(type, listener) {
        if (!this.listeners[type]) {
            return false;
        }
        if (!listener) {
            return true;
        }
        return (this.listeners[type].indexOf(listener) >= 0);
    }
    dispatchEvent(event) {
        event.target = this;
        if (!this.listeners[event.type]) {
            return false;
        }
        var callbacks = this.listeners[event.type];
        for (var i = 0; i < callbacks.length; i++) {
            callbacks[i](event);
        }
        return event.defaultPrevented;
    }
    getDefaultMessage() {
        return {};
    }
    invokeExcelCall(functionName, data, callback) {
        this.invokeRemoteCall('excelCall', functionName, data, callback);
    }
    invokeServiceCall(functionName, data, callback) {
        this.invokeRemoteCall('excelServiceCall', functionName, data, callback);
    }
    invokeRemoteCall(topic, functionName, data, callback) {
        var message = this.getDefaultMessage();
        message.messageId = RpcDispatcher.messageId;
        message.action = functionName;
        Object.assign(message, data);
        if (callback) {
            RpcDispatcher.callbacks[RpcDispatcher.messageId] = callback;
        }
        fin.desktop.InterApplicationBus.publish(topic, message);
        RpcDispatcher.messageId++;
    }
}
RpcDispatcher.messageId = 1;
RpcDispatcher.callbacks = {};
exports.RpcDispatcher = RpcDispatcher;
//# sourceMappingURL=RpcDispatcher.js.map

/***/ }),
/* 1 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

const RpcDispatcher_1 = __webpack_require__(0);
const ExcelWorkbook_1 = __webpack_require__(2);
const ExcelWorksheet_1 = __webpack_require__(3);
class Excel extends RpcDispatcher_1.RpcDispatcher {
    constructor() {
        super();
        this.workbooks = {};
        this.worksheets = {};
        this.processExcelEvent = (data, uuid) => {
            switch (data.event) {
                case "connected":
                    this.connected = true;
                    this.monitorDisconnect(uuid);
                    this.dispatchEvent({ type: data.event });
                    break;
                case "sheetChanged":
                    var sheets = this.worksheets[data.workbookName];
                    if (sheets && sheets[data.sheetName]) {
                        sheets[data.sheetName].dispatchEvent({ type: data.event, data: data });
                    }
                    break;
                case "sheetRenamed":
                    var sheets = this.worksheets[data.workbookName];
                    if (sheets && sheets[data.sheetName]) {
                        var sheet = sheets[data.sheetName];
                        sheets[data.sheetName] = null;
                        sheet.name = data.newName;
                        sheets[data.newName] = sheet;
                        sheet.dispatchEvent({ type: data.event, data: data });
                    }
                    break;
                case "selectionChanged":
                    var sheets = this.worksheets[data.workbookName];
                    if (sheets && sheets[data.sheetName]) {
                        sheets[data.sheetName].dispatchEvent({ type: data.event, data: data });
                    }
                    break;
                case "sheetActivated":
                    var sheets = this.worksheets[data.workbookName];
                    if (sheets && sheets[data.sheetName]) {
                        sheets[data.sheetName].dispatchEvent({ type: data.event });
                    }
                    break;
                case "sheetDeactivated":
                    var sheets = this.worksheets[data.workbookName];
                    if (sheets && sheets[data.sheetName]) {
                        sheets[data.sheetName].dispatchEvent({ type: data.event });
                    }
                    break;
                case "sheetAdded":
                    var workbook = this.getWorkbookByName(data.workbookName);
                    if (!this.worksheets[data.workbookName])
                        this.worksheets[data.workbookName] = {};
                    var sheets = this.worksheets[data.workbookName];
                    var sheet = sheets[data.sheetName] ? sheets[data.sheetName] : sheets[data.sheetName] = new ExcelWorksheet_1.ExcelWorksheet(data.sheetName, workbook);
                    workbook.dispatchEvent({ type: data.event, worksheet: sheet.toObject() });
                    break;
                case "sheetRemoved":
                    var workbook = this.getWorkbookByName(data.workbookName);
                    var sheet = this.worksheets[data.workbookName][data.sheetName];
                    delete this.worksheets[data.workbookName][data.sheetName];
                    workbook.dispatchEvent({ type: data.event, worksheet: sheet.toObject() });
                    break;
                case "workbookAdded":
                case "workbookOpened":
                    var workbook = new ExcelWorkbook_1.ExcelWorkbook(this, data.workbookName);
                    this.workbooks[data.workbookName] = workbook;
                    this.dispatchEvent({ type: data.event, workbook: workbook.toObject() });
                    break;
                case "afterCalculation":
                    this.dispatchEvent({ type: data.event });
                    break;
                case "workbookDeactivated":
                case "workbookActivated":
                    var workbook = this.getWorkbookByName(data.workbookName);
                    if (workbook)
                        workbook.dispatchEvent({ type: data.event });
                    break;
                case "workbookClosed":
                    var workbook = this.getWorkbookByName(data.workbookName);
                    delete this.workbooks[data.workbookName];
                    delete this.worksheets[data.workbookName];
                    workbook.dispatchEvent({ type: data.event });
                    this.dispatchEvent({ type: data.event, workbook: workbook.toObject() });
                    break;
                default:
                    this.dispatchEvent({ type: data.event });
                    break;
            }
        };
        this.processExcelResult = (data) => {
            var callbackData = {};
            switch (data.action) {
                case "getWorkbooks":
                    var workbookNames = data.data;
                    var _workbooks = [];
                    for (var i = 0; i < workbookNames.length; i++) {
                        var name = workbookNames[i];
                        if (!this.workbooks[name]) {
                            this.workbooks[name] = new ExcelWorkbook_1.ExcelWorkbook(this, name);
                        }
                        _workbooks.push(this.workbooks[name]);
                    }
                    callbackData = _workbooks.map(wb => wb.toObject());
                    break;
                case "getWorksheets":
                    var worksheetNames = data.data;
                    var _worksheets = [];
                    var worksheet = null;
                    for (var i = 0; i < worksheetNames.length; i++) {
                        if (!this.worksheets[data.workbook]) {
                            this.worksheets[data.workbook] = {};
                        }
                        worksheet = this.worksheets[data.workbook][worksheetNames[i]] ? this.worksheets[data.workbook][worksheetNames[i]] : this.worksheets[data.workbook][worksheetNames[i]] = new ExcelWorksheet_1.ExcelWorksheet(worksheetNames[i], this.workbooks[data.workbook]);
                        _worksheets.push(worksheet);
                    }
                    callbackData = _worksheets.map(ws => ws.toObject());
                    break;
                case "getCells":
                case "getCellsColumn":
                case "getCellsRow":
                    callbackData = data.data;
                    break;
                case "addWorkbook":
                case "openWorkbook":
                    if (!this.workbooks[data.workbookName]) {
                        var workbook = new ExcelWorkbook_1.ExcelWorkbook(this, data.workbook);
                        this.workbooks[data.workbook] = workbook;
                    }
                    else {
                        var workbook = this.workbooks[data.workbookName];
                    }
                    callbackData = workbook.toObject();
                case "addSheet":
                    if (!this.worksheets[data.workbookName])
                        this.worksheets[data.workbookName] = {};
                    var sheets = this.worksheets[data.workbookName];
                    var worksheet = sheets[data.sheetName] ? sheets[data.sheetName] : sheets[data.sheetName] = new ExcelWorksheet_1.ExcelWorksheet(data.sheetName, this.workbooks[data.workbookName]);
                    callbackData = worksheet.toObject();
                    break;
                case "getStatus":
                    callbackData = data.status;
                    break;
                case "getCalculationMode":
                case "getCellByName":
                    callbackData = data;
                    break;
            }
            if (RpcDispatcher_1.RpcDispatcher.callbacks[data.messageId]) {
                RpcDispatcher_1.RpcDispatcher.callbacks[data.messageId](callbackData);
                delete RpcDispatcher_1.RpcDispatcher.callbacks[data.messageId];
            }
        };
        this.processExcelServiceResult = (data) => {
            if (RpcDispatcher_1.RpcDispatcher.callbacks[data.messageId]) {
                RpcDispatcher_1.RpcDispatcher.callbacks[data.messageId](data.result);
                delete RpcDispatcher_1.RpcDispatcher.callbacks[data.messageId];
            }
        };
    }
    init() {
        fin.desktop.InterApplicationBus.subscribe("*", "excelEvent", this.processExcelEvent);
        fin.desktop.InterApplicationBus.subscribe("*", "excelResult", this.processExcelResult);
        fin.desktop.InterApplicationBus.subscribe("*", "excelServiceCallResult", this.processExcelServiceResult);
    }
    monitorDisconnect(uuid) {
        fin.desktop.ExternalApplication.wrap(uuid).addEventListener('disconnected', () => {
            this.connected = false;
            this.dispatchEvent({ type: 'disconnected' });
        });
    }
    run(callback) {
        if (this.connected) {
            callback();
        }
        else {
            var connectedCallback = () => {
                this.removeEventListener('connected', connectedCallback);
                callback();
            };
            this.addEventListener('connected', connectedCallback);
            fin.desktop.System.launchExternalProcess({ target: 'excel' });
        }
    }
    install(callback) {
        this.invokeServiceCall("install", null, callback);
    }
    getInstallationStatus(callback) {
        this.invokeServiceCall("getInstallationStatus", null, callback);
    }
    getWorkbooks(callback) {
        this.invokeExcelCall("getWorkbooks", null, callback);
    }
    getWorkbookByName(name) {
        return this.workbooks[name];
    }
    getWorksheetByName(workbookName, worksheetName) {
        if (this.worksheets[workbookName])
            return this.worksheets[workbookName][worksheetName] ? this.worksheets[workbookName][worksheetName] : null;
        return null;
    }
    addWorkbook(callback) {
        this.invokeExcelCall("addWorkbook", null, callback);
    }
    openWorkbook(path, callback) {
        this.invokeExcelCall("openWorkbook", { path: path }, callback);
    }
    getConnectionStatus(callback) {
        callback(this.connected);
    }
    getCalculationMode(callback) {
        this.invokeExcelCall("getCalculationMode", null, callback);
    }
    calculateAll(callback) {
        this.invokeExcelCall("calculateFull", null, callback);
    }
    toObject() {
        return {
            addEventListener: this.addEventListener.bind(this),
            dispatchEvent: this.dispatchEvent.bind(this),
            removeEventListener: this.removeEventListener.bind(this),
            addWorkbook: this.addWorkbook.bind(this),
            calculateAll: this.calculateAll.bind(this),
            getCalculationMode: this.getCalculationMode.bind(this),
            getConnectionStatus: this.getConnectionStatus.bind(this),
            getWorkbookByName: name => this.getWorkbookByName(name).toObject(),
            getWorkbooks: this.getWorkbooks.bind(this),
            init: this.init.bind(this),
            openWorkbook: this.openWorkbook.bind(this)
        };
    }
}
exports.Excel = Excel;
Object.defineProperty(exports, "__esModule", { value: true });
exports.default = new Excel();
//# sourceMappingURL=ExcelApplication.js.map

/***/ }),
/* 2 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

const RpcDispatcher_1 = __webpack_require__(0);
class ExcelWorkbook extends RpcDispatcher_1.RpcDispatcher {
    constructor(application, name) {
        super();
        this.application = application;
        this.name = name;
    }
    getDefaultMessage() {
        return {
            workbook: this.name
        };
    }
    getWorksheets(callback) {
        this.invokeExcelCall("getWorksheets", null, callback);
    }
    getWorksheetByName(name) {
        return this.application.getWorksheetByName(this.name, name);
    }
    addWorksheet(callback) {
        this.invokeExcelCall("addSheet", null, callback);
    }
    activate() {
        this.invokeExcelCall("activateWorkbook");
    }
    save() {
        this.invokeExcelCall("saveWorkbook");
    }
    close() {
        this.invokeExcelCall("closeWorkbook");
    }
    toObject() {
        return {
            addEventListener: this.addEventListener.bind(this),
            dispatchEvent: this.dispatchEvent.bind(this),
            removeEventListener: this.removeEventListener.bind(this),
            name: this.name,
            activate: this.activate.bind(this),
            addWorksheet: this.addWorksheet.bind(this),
            close: this.close.bind(this),
            getWorksheetByName: name => this.getWorksheetByName(name).toObject(),
            getWorksheets: this.getWorksheets.bind(this),
            save: this.save.bind(this)
        };
    }
}
exports.ExcelWorkbook = ExcelWorkbook;
//# sourceMappingURL=ExcelWorkbook.js.map

/***/ }),
/* 3 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

const RpcDispatcher_1 = __webpack_require__(0);
class ExcelWorksheet extends RpcDispatcher_1.RpcDispatcher {
    constructor(name, workbook) {
        super();
        this.name = name;
        this.workbook = workbook;
    }
    getDefaultMessage() {
        return {
            workbook: this.workbook.name,
            worksheet: this.name
        };
    }
    setCells(values, offset) {
        if (!offset)
            offset = "A1";
        this.invokeExcelCall("setCells", { offset: offset, values: values });
    }
    getCells(start, offsetWidth, offsetHeight, callback) {
        this.invokeExcelCall("getCells", { start: start, offsetWidth: offsetWidth, offsetHeight: offsetHeight }, callback);
    }
    getRow(start, width, callback) {
        this.invokeExcelCall("getCellsRow", { start: start, offsetWidth: width }, callback);
    }
    getColumn(start, offsetHeight, callback) {
        this.invokeExcelCall("getCellsColumn", { start: start, offsetHeight: offsetHeight }, callback);
    }
    activate() {
        this.invokeExcelCall("activateSheet");
    }
    activateCell(cellAddress) {
        this.invokeExcelCall("activateCell", { address: cellAddress });
    }
    addButton(name, caption, cellAddress) {
        this.invokeExcelCall("addButton", { address: cellAddress, buttonName: name, buttonCaption: caption });
    }
    setFilter(start, offsetWidth, offsetHeight, field, criteria1, op, criteria2, visibleDropDown) {
        this.invokeExcelCall("setFilter", {
            start: start,
            offsetWidth: offsetWidth,
            offsetHeight: offsetHeight,
            field: field,
            criteria1: criteria1,
            op: op,
            criteria2: criteria2,
            visibleDropDown: visibleDropDown
        });
    }
    formatRange(rangeCode, format, callback) {
        this.invokeExcelCall("formatRange", { rangeCode: rangeCode, format: format }, callback);
    }
    clearRange(rangeCode, callback) {
        this.invokeExcelCall("clearRange", { rangeCode: rangeCode }, callback);
    }
    clearRangeContents(rangeCode, callback) {
        this.invokeExcelCall("clearRangeContents", { rangeCode: rangeCode }, callback);
    }
    clearRangeFormats(rangeCode, callback) {
        this.invokeExcelCall("clearRangeFormats", { rangeCode: rangeCode }, callback);
    }
    clearAllCells(callback) {
        this.invokeExcelCall("clearAllCells", null, callback);
    }
    clearAllCellContents(callback) {
        this.invokeExcelCall("clearAllCellContents", null, callback);
    }
    clearAllCellFormats(callback) {
        this.invokeExcelCall("clearAllCellFormats", null, callback);
    }
    setCellName(cellAddress, cellName) {
        this.invokeExcelCall("setCellName", { address: cellAddress, cellName: cellName });
    }
    calculate() {
        this.invokeExcelCall("calculateSheet");
    }
    getCellByName(cellName, callback) {
        this.invokeExcelCall("getCellByName", { cellName: cellName }, callback);
    }
    protect(password) {
        this.invokeExcelCall("protectSheet", { password: password ? password : null });
    }
    toObject() {
        return {
            addEventListener: this.addEventListener.bind(this),
            dispatchEvent: this.dispatchEvent.bind(this),
            removeEventListener: this.removeEventListener.bind(this),
            name: this.name,
            activate: this.activate.bind(this),
            activateCell: this.activateCell.bind(this),
            addButton: this.addButton.bind(this),
            calculate: this.calculate.bind(this),
            clearAllCellContents: this.clearAllCellContents.bind(this),
            clearAllCellFormats: this.clearAllCellFormats.bind(this),
            clearAllCells: this.clearAllCells.bind(this),
            clearRange: this.clearRange.bind(this),
            clearRangeContents: this.clearRangeContents.bind(this),
            clearRangeFormats: this.clearRangeFormats.bind(this),
            formatRange: this.formatRange.bind(this),
            getCellByName: this.getCellByName.bind(this),
            getCells: this.getCells.bind(this),
            getColumn: this.getColumn.bind(this),
            getRow: this.getRow.bind(this),
            protect: this.protect.bind(this),
            setCellName: this.setCellName.bind(this),
            setCells: this.setCells.bind(this),
            setFilter: this.setFilter.bind(this)
        };
    }
}
exports.ExcelWorksheet = ExcelWorksheet;
//# sourceMappingURL=ExcelWorksheet.js.map

/***/ }),
/* 4 */
/***/ (function(module, exports, __webpack_require__) {

/**
 * Created by haseebriaz on 14/05/15.
 */

fin.desktop.Excel = __webpack_require__(1).default;

window.addEventListener("DOMContentLoaded", function () {

    var rowLength = 27;
    var columnLength = 12;
    var table = document.getElementById("excelExample");
    var tBody = table.getElementsByTagName("tbody")[0];
    var tHead = table.getElementsByTagName("thead")[0];

    var newWorkbookButton = document.getElementById("newWorkbookButton");
    var newWorksheetButton = document.getElementById("newSheetButton");

    var noConnectionContainer = document.getElementById("noConnection");
    var noWorkbooksContainer = document.getElementById("noWorkbooks");
    var workbooksContainer = document.getElementById("workbooksContainer");

    var displayContainers = new Map([
        [noConnectionContainer, { windowHeight: 195 }],
        [noWorkbooksContainer, { windowHeight: 195 }],
        [workbooksContainer, { windowHeight: 830 }]
    ]);

    newWorkbookButton.addEventListener("click", function () {
        fin.desktop.Excel.addWorkbook();
    });

    newWorksheetButton.addEventListener("click", function () {
        currentWorkbook.addWorksheet();
    });

    var currentWorksheet = null;
    var currentWorkbook = null;
    var currentCell = null;
    var formulaInput = document.getElementById("formulaInput");

    window.addEventListener("keydown", function (event) {

        switch (event.keyCode) {

            case 78: // N
                if (event.ctrlKey) fin.desktop.Excel.addWorkbook();
                break;
            case 37: // LEFT
                selectPreviousCell();
                break;
            case 38: // UP
                selectCellAbove();
                break;
            case 39: // RIGHT
                selectNextCell();
                break;
            case 40: //DOWN
                selectCellBelow();
                break;
        }
    });

    function setDisplayContainer(containerToDisplay) {
        if (!displayContainers.has(containerToDisplay)) {
            return;
        }

        for (var container of displayContainers.keys()) {
            container.style.display = "none";
        }
        containerToDisplay.style.display = "block";

        var windowHeight = displayContainers.get(containerToDisplay).windowHeight;

        fin.desktop.Window.getCurrent().animate({
            size: {
                height: windowHeight,
                duration: 500
            }
        });
    }

    function initTable() {

        var row = createRow(["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L"], "cellHeader", false);
        var column = createColumn("");
        column.className = "rowNumber";
        row.insertBefore(column, row.childNodes[0]);
        tHead.appendChild(row);

        for (var i = 1; i <= rowLength; i++) {

            row = createRow(columnLength, "cell", true);
            column = createColumn(i);
            column.className = "rowNumber";
            column.contentEditable = false;
            row.insertBefore(column, row.childNodes[0]);
            tBody.appendChild(row);
        }
    }

    function createRow(data, cellClassName, editable) {

        var length = data.length ? data.length : data;
        var row = document.createElement("tr");

        for (var i = 0; i < length; i++) {

            row.appendChild(createColumn(data[i], cellClassName, editable));
        }

        return row;
    }

    function createColumn(data, cellClassName, editable) {

        var column = document.createElement("td");
        column.className = cellClassName;

        if (editable) {

            column.contentEditable = true;
            //column.addEventListener("DOMCharacterDataModified", onDataChange);
            column.addEventListener("keydown", onDataChange);
            column.addEventListener("blur", onDataChange);
            column.addEventListener("mousedown", onCellClicked);
        }

        if (data) column.innerText = data;
        return column;
    }

    function onCellClicked(event) {

        selectCell(event.target);
    }

    function selectCell(cell, preventDefault) {

        if (currentCell) {

            currentCell.className = "cell";
            updateCellNumberClass(currentCell, "rowNumber", "cellHeader");
        }

        currentCell = cell;
        currentCell.className = "cellSelected";
        formulaInput.innerText = "Formula: " + cell.title;
        cell.focus();

        updateCellNumberClass(cell, "rowNumberSelected", "cellHeaderSelected");

        var address = getAddress(currentCell);

        if (!preventDefault) {
            currentWorksheet.activateCell(address.offset);
        }
    }

    function updateCellNumberClass(cell, className, headerClassName) {

        var row = cell.parentNode;
        var columnIndex = Array.prototype.indexOf.call(row.childNodes, cell);
        var rowIndex = Array.prototype.indexOf.call(row.parentNode.childNodes, cell.parentNode);
        tBody.childNodes[rowIndex].childNodes[0].className = className;
        tHead.getElementsByTagName("tr")[0].getElementsByTagName("td")[columnIndex].className = headerClassName;
    }

    function selectCellBelow() {

        if (!currentCell) return;
        var info = getAddress(currentCell);
        if (info.row >= rowLength) return;
        var cell = tBody.childNodes[info.row].childNodes[info.column];
        selectCell(cell);
    }

    function selectCellAbove() {

        if (!currentCell) return;
        var info = getAddress(currentCell);
        if (info.row <= 1) return;
        var cell = tBody.childNodes[info.row - 2].childNodes[info.column];
        selectCell(cell);
    }

    function selectNextCell() {

        if (!currentCell) return;
        var info = getAddress(currentCell);
        if (info.column >= columnLength) return;
        var cell = tBody.childNodes[info.row - 1].childNodes[info.column + 1];
        selectCell(cell);
    }

    function selectPreviousCell() {

        if (!currentCell) return;
        var info = getAddress(currentCell);
        if (info.column <= 1) return;
        var cell = tBody.childNodes[info.row - 1].childNodes[info.column - 1];
        selectCell(cell);
    }

    function onDataChange(event) {

        if (event.keyCode === 13 || event.type === "blur") {

            var update = getAddress(event.target);
            update.value = event.target.innerText;

            console.log('onDataChange ' + event.type);
            currentWorksheet.setCells([[update.value]], update.offset);
            if (event.type === "keydown") {

                selectCellBelow();
                event.preventDefault();
            }
        }
    }

    function getAddress(td) {

        var column = td.cellIndex;
        var row = td.parentElement.rowIndex;
        var offset = tHead.getElementsByTagName("td")[column].innerText.toString() + row;
        return { column: column, row: row, offset: offset };
    }

    function updateData(data) {

        var row = null;
        var currentData = null;

        for (var i = 0; i < data.length; i++) {

            row = tBody.childNodes[i];
            for (var j = 1; j < row.childNodes.length; j++) {

                currentData = data[i][j - 1];
                updateCell(row.childNodes[j], currentData.value, currentData.formula);
            }
        }
    }

    function updateCell(cell, value, formula) {

        cell.innerText = value ? value : "";
        cell.title = formula ? formula : "";
    }

    function onSheetChanged(event) {

        var cell = tBody.getElementsByTagName("tr")[event.data.row - 1].getElementsByTagName("td")[event.data.column];
        updateCell(cell, event.data.value, event.data.formula);
    }

    function onSelectionChanged(event) {
        var cell = tBody.getElementsByTagName("tr")[event.data.row - 1].getElementsByTagName("td")[event.data.column];
        selectCell(cell, true);
    }

    function onSheetActivated(event) {

        selectWorksheet(event.target);
    }

    function selectWorksheet(sheet) {

        if (currentWorksheet === sheet) {
            return;
        }

        if (currentWorksheet) {
            var tab = document.getElementById(currentWorksheet.name);
            if (tab) tab.className = "tab";
        }
        document.getElementById(sheet.name).className = "tabSelected";
        currentWorksheet = sheet;
        currentWorksheet.getCells("A1", columnLength, rowLength, updateData);
    }

    function selectWorkbook(workbook) {

        if (currentWorkbook) {

            var tab = document.getElementById(currentWorkbook.name);
            if (tab) tab.className = "workbookTab";
        }

        document.getElementById(workbook.name).className = "workbookTabSelected";
        currentWorkbook = workbook;
        currentWorkbook.getWorksheets(updateSheets);
    }

    function onWorkbookTabClicked(event) {

        var workbook = fin.desktop.Excel.getWorkbookByName(event.target.innerText);
        if (currentWorkbook === workbook) return;
        workbook.activate();
    }

    function onWorkbookActivated(event) {

        selectWorkbook(event.target);
    }

    function onWorkbookAdded(event) {

        var workbook = event.workbook;
        workbook.addEventListener("workbookActivated", onWorkbookActivated);
        workbook.addEventListener("sheetAdded", onWorksheetAdded);
        workbook.addEventListener("sheetRemoved", onWorksheetRemoved);
        addWorkbookTab(event.workbook.name);

        if (workbooksContainer.style.display === "none") {
            setDisplayContainer(workbooksContainer);
        }
    }

    function onWorkbookRemoved(event) {
        currentWorkbook = null;
        var workbook = event.workbook;
        workbook.removeEventListener("workbookActivated", onWorkbookActivated);
        workbook.removeEventListener("sheetAdded", onWorksheetAdded);
        workbook.removeEventListener("sheetRemoved", onWorksheetRemoved);

        document.getElementById("workbookTabs").removeChild(document.getElementById(event.workbook.name));

        if (document.getElementById("workbookTabs").childNodes.length < 2) {
            setDisplayContainer(noWorkbooksContainer);
        }
    }

    function onWorksheetAdded(event) {

        addWorksheetTab(event.worksheet);
    }

    function addWorksheetTab(worksheet) {

        var sheetsTabHolder = document.getElementById("sheets");
        var button = document.createElement("button");
        button.innerText = worksheet.name;
        button.className = "tab";
        button.id = worksheet.name;
        button.addEventListener("click", onSheetButtonClicked);
        sheetsTabHolder.insertBefore(button, newWorksheetButton);

        worksheet.addEventListener("sheetChanged", onSheetChanged);
        worksheet.addEventListener("selectionChanged", onSelectionChanged);
        worksheet.addEventListener("sheetActivated", onSheetActivated);
    }

    function onSheetButtonClicked(event) {

        var sheet = currentWorkbook.getWorksheetByName(event.target.innerText);
        if (currentWorksheet === sheet) return;
        sheet.activate();
    }

    function onWorksheetRemoved(event) {

        if (event.worksheet.workbook === currentWorkbook) {

            event.worksheet.removeEventListener("sheetChanged", onSheetChanged);
            event.worksheet.removeEventListener("selectionChanged", onSelectionChanged);
            event.worksheet.removeEventListener("sheetActivated", onSheetActivated);
            document.getElementById("sheets").removeChild(document.getElementById(event.worksheet.name));
            currentWorksheet = null;
        }
    }

    function updateSheets(worksheets) {

        var sheetsTabHolder = document.getElementById("sheets");
        while (sheetsTabHolder.firstChild) {

            sheetsTabHolder.removeChild(sheetsTabHolder.firstChild);
        }

        sheetsTabHolder.appendChild(newWorksheetButton);
        for (var i = 0; i < worksheets.length; i++) {

            addWorksheetTab(worksheets[i]);
        }

        selectWorksheet(worksheets[0]);
    }

    function addWorkbookTab(name) {

        var button = document.createElement("button");
        button.id = button.innerText = name;
        button.className = "workbookTab";
        button.addEventListener("click", onWorkbookTabClicked);
        document.getElementById("workbookTabs").insertBefore(button, newWorkbookButton);
    }

    function onExcelConnected(event) {

        document.getElementById("status").innerText = "Connected to Excel";
        fin.desktop.Excel.getWorkbooks(function (workbooks) {

            for (var i = 0; i < workbooks.length; i++) {

                addWorkbookTab(workbooks[i].name);
                workbooks[i].addEventListener("workbookActivated", onWorkbookActivated);
                workbooks[i].addEventListener("sheetAdded", onWorksheetAdded);
                workbooks[i].addEventListener("sheetRemoved", onWorksheetRemoved);
            }

            if (workbooks.length) {

                selectWorkbook(workbooks[0]);
                setDisplayContainer(workbooksContainer);
            }
            else {
                setDisplayContainer(noWorkbooksContainer);
            }
        });
    }

    function onExcelDisconnected(event) {
        document.getElementById("status").innerText = "Excel not connected";
        setDisplayContainer(noConnectionContainer);
    }

    function installAddIn() {
        var installFolder = '%localappdata%\\OpenFin\\shared\\assets\\excel-api-addin';
        var servicePath = 'OpenFin.ExcelService.exe';
        var addInPath = 'OpenFin.ExcelApi-AddIn.xll';

        var statusElement = document.getElementById("status");

        if (statusElement.innerText == "Connecting...") {
            return;
        }

        statusElement.innerText = "Connecting...";

        Promise.resolve()
            .then(() => deployAddIn(servicePath, installFolder))
            .then(() => startExcelService(servicePath, installFolder))
            .then(() => registerAddIn(servicePath, installFolder))
            .then(launchExcel)
            .catch(err => console.log(err));
    }

    function deployAddIn(servicePath, installFolder) {
        return new Promise((resolve, reject) => {
            fin.desktop.System.launchExternalProcess({
                alias: 'excel-api-addin',
                target: servicePath,
                arguments: '-d "' + installFolder + '"',
                listener: function (args) {
                    console.log('Installer script completed! ' + args.exitCode);
                    // (args.exitCode === 0) {
                        resolve();
                    //} else {
                    //    reject('Error deploying Add-In');
                    //}
                }
            });
        });
    }

    function registerAddIn(servicePath, installFolder) {
        return new Promise((resolve, reject) => {
            fin.desktop.Excel.install(ack => {
                if (ack.success) {
                    resolve();
                } else {
                    reject();
                }
            });
        });
    }

    function startExcelService(servicePath, installFolder) {
        return new Promise((resolve, reject) => {
            var onServiceStarted = () => {
                console.log('Service Started');
                fin.desktop.Excel.removeEventListener('excelServiceStarted', onServiceStarted);
                resolve();
            };

            chrome.desktop.getDetails(function (details) {
                fin.desktop.Excel.addEventListener('excelServiceStarted', onServiceStarted);

                fin.desktop.System.launchExternalProcess({
                    target: installFolder + '\\OpenFin.ExcelService.exe',
                    arguments: '-p ' + details.port
                }, process => {
                    console.log('Service Launced');
                }, error => {
                    reject('Error starting Excel service');
                });
            });
        });
    }

    function launchExcel() {
        return new Promise((resolve, reject) => {
            fin.desktop.Excel.run(resolve);
        });
    }

    initTable(27, 12);

    fin.desktop.main(function () {

        var Excel = fin.desktop.Excel;
        Excel.init();
        Excel.getConnectionStatus(onExcelConnected);
        Excel.addEventListener("workbookAdded", onWorkbookAdded);
        Excel.addEventListener("workbookOpened", onWorkbookAdded);
        Excel.addEventListener("workbookClosed", onWorkbookRemoved);
        Excel.addEventListener("connected", onExcelConnected);
        Excel.addEventListener("disconnected", onExcelDisconnected)

        installAddIn();
    });

    window.installAddIn = installAddIn;
});


/***/ })
/******/ ]);