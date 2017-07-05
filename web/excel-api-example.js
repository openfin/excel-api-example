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

var RpcDispatcher = (function () {
    function RpcDispatcher() {
        this.listeners = {};
    }
    RpcDispatcher.prototype.addEventListener = function (type, listener) {
        if (this.hasEventListener(type, listener)) {
            return;
        }
        if (!this.listeners[type]) {
            this.listeners[type] = [];
        }
        this.listeners[type].push(listener);
    };
    RpcDispatcher.prototype.removeEventListener = function (type, listener) {
        if (!this.hasEventListener(type, listener)) {
            return;
        }
        var callbacksOfType = this.listeners[type];
        callbacksOfType.splice(callbacksOfType.indexOf(listener), 1);
    };
    RpcDispatcher.prototype.hasEventListener = function (type, listener) {
        if (!this.listeners[type]) {
            return false;
        }
        if (!listener) {
            return true;
        }
        return (this.listeners[type].indexOf(listener) >= 0);
    };
    RpcDispatcher.prototype.dispatchEvent = function (event) {
        event.target = this;
        if (!this.listeners[event.type]) {
            return false;
        }
        var callbacks = this.listeners[event.type];
        for (var i = 0; i < callbacks.length; i++) {
            callbacks[i](event);
        }
        return event.defaultPrevented;
    };
    RpcDispatcher.prototype.getDefaultMessage = function () {
        return {};
    };
    RpcDispatcher.prototype.invokeRemote = function (functionName, data, callback) {
        var message = this.getDefaultMessage();
        message.messageId = RpcDispatcher.messageId;
        message.action = functionName;
        if (data) {
            for (var key in data) {
                message[key] = data[key];
            }
        }
        if (callback) {
            RpcDispatcher.callbacks[RpcDispatcher.messageId] = callback;
        }
        fin.desktop.InterApplicationBus.publish("excelCall", message);
        RpcDispatcher.messageId++;
    };
    RpcDispatcher.messageId = 1;
    RpcDispatcher.callbacks = {};
    return RpcDispatcher;
}());
exports.RpcDispatcher = RpcDispatcher;
//# sourceMappingURL=RpcDispatcher.js.map

/***/ }),
/* 1 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

var __extends = (this && this.__extends) || function (d, b) {
    for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p];
    function __() { this.constructor = d; }
    d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
};
var RpcDispatcher_1 = __webpack_require__(0);
var ExcelWorkbook_1 = __webpack_require__(2);
var ExcelWorksheet_1 = __webpack_require__(3);
var Excel = (function (_super) {
    __extends(Excel, _super);
    function Excel() {
        var _this = this;
        _super.call(this);
        this.workbooks = {};
        this.worksheets = {};
        this.processExcelEvent = function (data, uuid) {
            switch (data.event) {
                case "connected":
                    _this.connected = true;
                    _this.monitorDisconnect(uuid);
                    _this.dispatchEvent({ type: data.event });
                    break;
                case "sheetChanged":
                    var sheets = _this.worksheets[data.workbookName];
                    if (sheets && sheets[data.sheetName]) {
                        sheets[data.sheetName].dispatchEvent({ type: data.event, data: data });
                    }
                    break;
                case "sheetRenamed":
                    var sheets = _this.worksheets[data.workbookName];
                    if (sheets && sheets[data.sheetName]) {
                        var sheet = sheets[data.sheetName];
                        sheets[data.sheetName] = null;
                        sheet.name = data.newName;
                        sheets[data.newName] = sheet;
                        sheet.dispatchEvent({ type: data.event, data: data });
                    }
                    break;
                case "selectionChanged":
                    var sheets = _this.worksheets[data.workbookName];
                    if (sheets && sheets[data.sheetName]) {
                        sheets[data.sheetName].dispatchEvent({ type: data.event, data: data });
                    }
                    break;
                case "sheetActivated":
                    var sheets = _this.worksheets[data.workbookName];
                    if (sheets && sheets[data.sheetName]) {
                        sheets[data.sheetName].dispatchEvent({ type: data.event });
                    }
                    break;
                case "sheetDeactivated":
                    var sheets = _this.worksheets[data.workbookName];
                    if (sheets && sheets[data.sheetName]) {
                        sheets[data.sheetName].dispatchEvent({ type: data.event });
                    }
                    break;
                case "sheetAdded":
                    var workbook = _this.getWorkbookByName(data.workbookName);
                    if (!_this.worksheets[data.workbookName])
                        _this.worksheets[data.workbookName] = {};
                    var sheets = _this.worksheets[data.workbookName];
                    var sheet = sheets[data.sheetName] ? sheets[data.sheetName] : sheets[data.sheetName] = new ExcelWorksheet_1.ExcelWorksheet(data.sheetName, workbook);
                    workbook.dispatchEvent({ type: data.event, worksheet: sheet.toObject() });
                    break;
                case "sheetRemoved":
                    var workbook = _this.getWorkbookByName(data.workbookName);
                    var sheet = _this.worksheets[data.workbookName][data.sheetName];
                    delete _this.worksheets[data.workbookName][data.sheetName];
                    workbook.dispatchEvent({ type: data.event, worksheet: sheet.toObject() });
                    break;
                case "workbookAdded":
                case "workbookOpened":
                    var workbook = new ExcelWorkbook_1.ExcelWorkbook(_this, data.workbookName);
                    _this.workbooks[data.workbookName] = workbook;
                    _this.dispatchEvent({ type: data.event, workbook: workbook.toObject() });
                    break;
                case "afterCalculation":
                    _this.dispatchEvent({ type: data.event });
                    break;
                case "workbookDeactivated":
                case "workbookActivated":
                    var workbook = _this.getWorkbookByName(data.workbookName);
                    if (workbook)
                        workbook.dispatchEvent({ type: data.event });
                    break;
                case "workbookClosed":
                    var workbook = _this.getWorkbookByName(data.workbookName);
                    delete _this.workbooks[data.workbookName];
                    delete _this.worksheets[data.workbookName];
                    workbook.dispatchEvent({ type: data.event });
                    _this.dispatchEvent({ type: data.event, workbook: workbook.toObject() });
                    break;
            }
        };
        this.processExcelResult = function (data) {
            var callbackData = {};
            switch (data.action) {
                case "getWorkbooks":
                    var workbookNames = data.data;
                    var _workbooks = [];
                    for (var i = 0; i < workbookNames.length; i++) {
                        var name = workbookNames[i];
                        if (!_this.workbooks[name]) {
                            _this.workbooks[name] = new ExcelWorkbook_1.ExcelWorkbook(_this, name);
                        }
                        _workbooks.push(_this.workbooks[name]);
                    }
                    callbackData = _workbooks.map(function (wb) { return wb.toObject(); });
                    break;
                case "getWorksheets":
                    var worksheetNames = data.data;
                    var _worksheets = [];
                    var worksheet = null;
                    for (var i = 0; i < worksheetNames.length; i++) {
                        if (!_this.worksheets[data.workbook]) {
                            _this.worksheets[data.workbook] = {};
                        }
                        worksheet = _this.worksheets[data.workbook][worksheetNames[i]] ? _this.worksheets[data.workbook][worksheetNames[i]] : _this.worksheets[data.workbook][worksheetNames[i]] = new ExcelWorksheet_1.ExcelWorksheet(worksheetNames[i], _this.workbooks[data.workbook]);
                        _worksheets.push(worksheet);
                    }
                    callbackData = _worksheets.map(function (ws) { return ws.toObject(); });
                    break;
                case "getCells":
                case "getCellsColumn":
                case "getCellsRow":
                    callbackData = data.data;
                    break;
                case "addWorkbook":
                case "openWorkbook":
                    if (!_this.workbooks[data.workbookName]) {
                        var workbook = new ExcelWorkbook_1.ExcelWorkbook(_this, data.workbook);
                        _this.workbooks[data.workbook] = workbook;
                    }
                    else {
                        var workbook = _this.workbooks[data.workbookName];
                    }
                    callbackData = workbook.toObject();
                case "addSheet":
                    if (!_this.worksheets[data.workbookName])
                        _this.worksheets[data.workbookName] = {};
                    var sheets = _this.worksheets[data.workbookName];
                    var worksheet = sheets[data.sheetName] ? sheets[data.sheetName] : sheets[data.sheetName] = new ExcelWorksheet_1.ExcelWorksheet(data.sheetName, _this.workbooks[data.workbookName]);
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
        this.processExcelCustomFunction = function (data) {
            if (!window[data.functionName]) {
                console.error("function ", data.functionName, "is not defined.");
                return;
            }
            var args = data.arguments.split(",");
            for (var i = 0; i < args.length; i++) {
                var num = Number(args[i]);
                if (!isNaN(num))
                    args[i] = num;
            }
            window[data.functionName].apply(null, args);
        };
    }
    Excel.prototype.init = function () {
        fin.desktop.InterApplicationBus.subscribe("*", "excelEvent", this.processExcelEvent);
        fin.desktop.InterApplicationBus.subscribe("*", "excelResult", this.processExcelResult);
        fin.desktop.InterApplicationBus.subscribe("*", "excelCustomFunction", this.processExcelCustomFunction);
    };
    Excel.prototype.monitorDisconnect = function (uuid) {
        var _this = this;
        fin.desktop.ExternalApplication.wrap(uuid).addEventListener('disconnected', function () {
            _this.connected = false;
            _this.dispatchEvent({ type: 'disconnected' });
        });
    };
    Excel.prototype.getWorkbooks = function (callback) {
        this.invokeRemote("getWorkbooks", null, callback);
    };
    Excel.prototype.getWorkbookByName = function (name) {
        return this.workbooks[name];
    };
    Excel.prototype.getWorksheetByName = function (workbookName, worksheetName) {
        if (this.worksheets[workbookName])
            return this.worksheets[workbookName][worksheetName] ? this.worksheets[workbookName][worksheetName] : null;
        return null;
    };
    Excel.prototype.addWorkbook = function (callback) {
        this.invokeRemote("addWorkbook", null, callback);
    };
    Excel.prototype.openWorkbook = function (path, callback) {
        this.invokeRemote("openWorkbook", { path: path }, callback);
    };
    Excel.prototype.getConnectionStatus = function (callback) {
        callback(this.connected);
    };
    Excel.prototype.getCalculationMode = function (callback) {
        this.invokeRemote("getCalculationMode", null, callback);
    };
    Excel.prototype.calculateAll = function (callback) {
        this.invokeRemote("calculateFull", null, callback);
    };
    Excel.prototype.toObject = function () {
        var _this = this;
        return {
            addEventListener: this.addEventListener.bind(this),
            dispatchEvent: this.dispatchEvent.bind(this),
            removeEventListener: this.removeEventListener.bind(this),
            addWorkbook: this.addWorkbook.bind(this),
            calculateAll: this.calculateAll.bind(this),
            getCalculationMode: this.getCalculationMode.bind(this),
            getConnectionStatus: this.getConnectionStatus.bind(this),
            getWorkbookByName: function (name) { return _this.getWorkbookByName(name).toObject(); },
            getWorkbooks: this.getWorkbooks.bind(this),
            init: this.init.bind(this),
            openWorkbook: this.openWorkbook.bind(this)
        };
    };
    return Excel;
}(RpcDispatcher_1.RpcDispatcher));
exports.Excel = Excel;
Object.defineProperty(exports, "__esModule", { value: true });
exports.default = new Excel();
//# sourceMappingURL=ExcelApplication.js.map

/***/ }),
/* 2 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

var __extends = (this && this.__extends) || function (d, b) {
    for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p];
    function __() { this.constructor = d; }
    d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
};
var RpcDispatcher_1 = __webpack_require__(0);
var ExcelWorkbook = (function (_super) {
    __extends(ExcelWorkbook, _super);
    function ExcelWorkbook(application, name) {
        _super.call(this);
        this.application = application;
        this.name = name;
    }
    ExcelWorkbook.prototype.getDefaultMessage = function () {
        return {
            workbook: this.name
        };
    };
    ExcelWorkbook.prototype.getWorksheets = function (callback) {
        this.invokeRemote("getWorksheets", null, callback);
    };
    ExcelWorkbook.prototype.getWorksheetByName = function (name) {
        return this.application.getWorksheetByName(this.name, name);
    };
    ExcelWorkbook.prototype.addWorksheet = function (callback) {
        this.invokeRemote("addSheet", null, callback);
    };
    ExcelWorkbook.prototype.activate = function () {
        this.invokeRemote("activateWorkbook");
    };
    ExcelWorkbook.prototype.save = function () {
        this.invokeRemote("saveWorkbook");
    };
    ExcelWorkbook.prototype.close = function () {
        this.invokeRemote("closeWorkbook");
    };
    ExcelWorkbook.prototype.toObject = function () {
        var _this = this;
        return {
            addEventListener: this.addEventListener.bind(this),
            dispatchEvent: this.dispatchEvent.bind(this),
            removeEventListener: this.removeEventListener.bind(this),
            name: this.name,
            activate: this.activate.bind(this),
            addWorksheet: this.addWorksheet.bind(this),
            close: this.close.bind(this),
            getWorksheetByName: function (name) { return _this.getWorksheetByName(name).toObject(); },
            getWorksheets: this.getWorksheets.bind(this),
            save: this.save.bind(this)
        };
    };
    return ExcelWorkbook;
}(RpcDispatcher_1.RpcDispatcher));
exports.ExcelWorkbook = ExcelWorkbook;
//# sourceMappingURL=ExcelWorkbook.js.map

/***/ }),
/* 3 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

var __extends = (this && this.__extends) || function (d, b) {
    for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p];
    function __() { this.constructor = d; }
    d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
};
var RpcDispatcher_1 = __webpack_require__(0);
var ExcelWorksheet = (function (_super) {
    __extends(ExcelWorksheet, _super);
    function ExcelWorksheet(name, workbook) {
        _super.call(this);
        this.name = name;
        this.workbook = workbook;
    }
    ExcelWorksheet.prototype.getDefaultMessage = function () {
        return {
            workbook: this.workbook.name,
            worksheet: this.name
        };
    };
    ExcelWorksheet.prototype.setCells = function (values, offset) {
        if (!offset)
            offset = "A1";
        this.invokeRemote("setCells", { offset: offset, values: values });
    };
    ExcelWorksheet.prototype.getCells = function (start, offsetWidth, offsetHeight, callback) {
        this.invokeRemote("getCells", { start: start, offsetWidth: offsetWidth, offsetHeight: offsetHeight }, callback);
    };
    ExcelWorksheet.prototype.getRow = function (start, width, callback) {
        this.invokeRemote("getCellsRow", { start: start, offsetWidth: width }, callback);
    };
    ExcelWorksheet.prototype.getColumn = function (start, offsetHeight, callback) {
        this.invokeRemote("getCellsColumn", { start: start, offsetHeight: offsetHeight }, callback);
    };
    ExcelWorksheet.prototype.activate = function () {
        this.invokeRemote("activateSheet");
    };
    ExcelWorksheet.prototype.activateCell = function (cellAddress) {
        this.invokeRemote("activateCell", { address: cellAddress });
    };
    ExcelWorksheet.prototype.addButton = function (name, caption, cellAddress) {
        this.invokeRemote("addButton", { address: cellAddress, buttonName: name, buttonCaption: caption });
    };
    ExcelWorksheet.prototype.setFilter = function (start, offsetWidth, offsetHeight, field, criteria1, op, criteria2, visibleDropDown) {
        this.invokeRemote("setFilter", {
            start: start,
            offsetWidth: offsetWidth,
            offsetHeight: offsetHeight,
            field: field,
            criteria1: criteria1,
            op: op,
            criteria2: criteria2,
            visibleDropDown: visibleDropDown
        });
    };
    ExcelWorksheet.prototype.formatRange = function (rangeCode, format, callback) {
        this.invokeRemote("formatRange", { rangeCode: rangeCode, format: format }, callback);
    };
    ExcelWorksheet.prototype.clearRange = function (rangeCode, callback) {
        this.invokeRemote("clearRange", { rangeCode: rangeCode }, callback);
    };
    ExcelWorksheet.prototype.clearRangeContents = function (rangeCode, callback) {
        this.invokeRemote("clearRangeContents", { rangeCode: rangeCode }, callback);
    };
    ExcelWorksheet.prototype.clearRangeFormats = function (rangeCode, callback) {
        this.invokeRemote("clearRangeFormats", { rangeCode: rangeCode }, callback);
    };
    ExcelWorksheet.prototype.clearAllCells = function (callback) {
        this.invokeRemote("clearAllCells", null, callback);
    };
    ExcelWorksheet.prototype.clearAllCellContents = function (callback) {
        this.invokeRemote("clearAllCellContents", null, callback);
    };
    ExcelWorksheet.prototype.clearAllCellFormats = function (callback) {
        this.invokeRemote("clearAllCellFormats", null, callback);
    };
    ExcelWorksheet.prototype.setCellName = function (cellAddress, cellName) {
        this.invokeRemote("setCellName", { address: cellAddress, cellName: cellName });
    };
    ExcelWorksheet.prototype.calculate = function () {
        this.invokeRemote("calculateSheet");
    };
    ExcelWorksheet.prototype.getCellByName = function (cellName, callback) {
        this.invokeRemote("getCellByName", { cellName: cellName }, callback);
    };
    ExcelWorksheet.prototype.protect = function (password) {
        this.invokeRemote("protectSheet", { password: password ? password : null });
    };
    ExcelWorksheet.prototype.toObject = function () {
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
    };
    return ExcelWorksheet;
}(RpcDispatcher_1.RpcDispatcher));
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

        var statusElement = document.getElementById("status");

        if (statusElement.innerText == "Connecting...") {
            return;
        }

        statusElement.innerText = "Connecting...";

        fin.desktop.System.launchExternalProcess({
            alias: 'excel-api-addin',
            target: 'OpenFin.ExcelApi.Installer.exe',
            arguments: '-d "' + installFolder + '"',
            listener: function (args) {
                console.log('Installer script completed!');
                if (args.exitCode === 0) {
                    fin.desktop.System.launchExternalProcess({
                        target: installFolder + '\\OpenFin.ExcelApi-AddIn.xll'
                    });
                }
            }
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