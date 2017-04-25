var __extends = (this && this.__extends) || function (d, b) {
    for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p];
    function __() { this.constructor = d; }
    d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
};
var fin;
(function (fin) {
    var desktop;
    (function (desktop) {
        // RpcDispatcher
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
            return RpcDispatcher;
        }());
        RpcDispatcher.messageId = 1;
        RpcDispatcher.callbacks = {};
        // RpcDispatcher
        // workbook
        var ExcelWorkbook = (function (_super) {
            __extends(ExcelWorkbook, _super);
            function ExcelWorkbook(application, name) {
                var _this = _super.call(this) || this;
                _this.application = application;
                _this.name = name;
                return _this;
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
        }(RpcDispatcher));
        // workbook
        // worksheet
        var ExcelWorksheet = (function (_super) {
            __extends(ExcelWorksheet, _super);
            function ExcelWorksheet(name, workbook) {
                var _this = _super.call(this) || this;
                _this.name = name;
                _this.workbook = workbook;
                return _this;
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
        }(RpcDispatcher));
        // worksheet
        // Excel
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
                            var sheet = sheets[data.sheetName] ? sheets[data.sheetName] : sheets[data.sheetName] = new ExcelWorksheet(data.sheetName, workbook);
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
                            var workbook = new ExcelWorkbook(_this, data.workbookName);
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
                                    _this.workbooks[name] = new ExcelWorkbook(_this, name);
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
                                worksheet = _this.worksheets[data.workbook][worksheetNames[i]] ? _this.worksheets[data.workbook][worksheetNames[i]] : _this.worksheets[data.workbook][worksheetNames[i]] = new ExcelWorksheet(worksheetNames[i], _this.workbooks[data.workbook]);
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
                                var workbook = new ExcelWorkbook(_this, data.workbook);
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
                            var worksheet = sheets[data.sheetName] ? sheets[data.sheetName] : sheets[data.sheetName] = new ExcelWorksheet(data.sheetName, _this.workbooks[data.workbookName]);
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
                    if (RpcDispatcher.callbacks[data.messageId]) {
                        RpcDispatcher.callbacks[data.messageId](callbackData);
                        delete RpcDispatcher.callbacks[data.messageId];
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
        }(RpcDispatcher));
        // Excel
        fin.desktop.Excel = (new Excel()).toObject();
    })(desktop = fin.desktop || (fin.desktop = {}));
})(fin || (fin = {}));
//# sourceMappingURL=ExcelApi.js.map