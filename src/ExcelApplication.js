"use strict";
var __extends = (this && this.__extends) || function (d, b) {
    for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p];
    function __() { this.constructor = d; }
    d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
};
var RpcDispatcher_1 = require("./RpcDispatcher");
var ExcelWorkbook_1 = require("./ExcelWorkbook");
var ExcelWorksheet_1 = require("./ExcelWorksheet");
var Excel = (function (_super) {
    __extends(Excel, _super);
    function Excel() {
        var _this = _super.call(this) || this;
        _this.workbooks = {};
        _this.worksheets = {};
        _this.processExcelEvent = function (data) {
            switch (data.event) {
                case "connected":
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
        _this.processExcelResult = function (data) {
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
        _this.processExcelCustomFunction = function (data) {
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
        return _this;
    }
    Excel.prototype.init = function () {
        fin.desktop.InterApplicationBus.subscribe("*", "excelEvent", this.processExcelEvent);
        fin.desktop.InterApplicationBus.subscribe("*", "excelResult", this.processExcelResult);
        fin.desktop.InterApplicationBus.subscribe("*", "excelCustomFunction", this.processExcelCustomFunction);
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
        this.invokeRemote("getStatus", null, callback);
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