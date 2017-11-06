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
/******/ 	return __webpack_require__(__webpack_require__.s = 5);
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
    dispatchEvent(evtOrTypeArg, data) {
        var event;
        if (typeof evtOrTypeArg == "string") {
            event = Object.assign({
                target: this.toObject(),
                type: evtOrTypeArg,
                defaultPrevented: false
            }, data);
        }
        else {
            event = evtOrTypeArg;
        }
        var callbacks = this.listeners[event.type] || [];
        callbacks.forEach(callback => callback(event));
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
        var args = data || {};
        var invoker = this;
        Object.assign(message, {
            messageId: RpcDispatcher.messageId,
            target: {
                connectionUuid: this.connectionUuid,
                workbookName: invoker.workbookName || (invoker.workbook && invoker.workbook.workbookName) || args.workbookName || args.workbook,
                worksheetName: invoker.worksheetName || args.worksheetName || args.worksheet,
                rangeCode: args.rangeCode
            },
            action: functionName,
            data: data
        });
        if (callback) {
            RpcDispatcher.callbacks[RpcDispatcher.messageId] = callback;
        }
        fin.desktop.InterApplicationBus.send(this.connectionUuid, topic, message);
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
const ExcelApplication_1 = __webpack_require__(2);
const excelServiceUuid = "886834D1-4651-4872-996C-7B2578E953B9";
class ExcelApi extends RpcDispatcher_1.RpcDispatcher {
    constructor() {
        super();
        this.applications = {};
        this.processExcelServiceEvent = (data) => {
            var preventDefault = false;
            switch (data.event) {
                case "started":
                    break;
                case "registrationRollCall":
                    this.registerAppInstance();
                    break;
                case "excelConnected":
                    this.processExcelConnectedEvent(data);
                    break;
                case "excelDisconnected":
                    this.processExcelDisconnectedEvent(data);
                    break;
            }
            if (!preventDefault) {
                this.dispatchEvent(data.event);
            }
        };
        this.processExcelServiceResult = (data) => {
            // Internal processing
            switch (data.action) {
                case "getExcelInstances":
                    this.processGetExcelInstancesResult(data.result);
                    break;
            }
            // Dispatch result to callbacks
            if (RpcDispatcher_1.RpcDispatcher.callbacks[data.messageId]) {
                RpcDispatcher_1.RpcDispatcher.callbacks[data.messageId](data.result);
                delete RpcDispatcher_1.RpcDispatcher.callbacks[data.messageId];
            }
        };
        this.registerAppInstance = () => {
            this.invokeServiceCall("registerAppInstance", { domain: document.domain });
        };
        this.connectionUuid = excelServiceUuid;
    }
    init() {
        if (!this.initialized) {
            fin.desktop.InterApplicationBus.subscribe("*", "excelServiceEvent", this.processExcelServiceEvent);
            fin.desktop.InterApplicationBus.subscribe("*", "excelServiceCallResult", this.processExcelServiceResult);
            this.registerAppInstance();
            this.getExcelInstances();
            this.monitorDisconnect();
            this.initialized = true;
        }
    }
    monitorDisconnect() {
        fin.desktop.ExternalApplication.wrap(excelServiceUuid).addEventListener("disconnected", () => {
            this.dispatchEvent("stopped");
        });
    }
    connectLegacyApi(connectedUuid) {
        if (!ExcelApi.legacyApi) {
            ExcelApi.legacyApi = ExcelApi.instance.applications[connectedUuid].toObject();
        }
    }
    disconnectLegacyApi(disconnectedUuid) {
        if (ExcelApi.legacyApi.connectionUuid === disconnectedUuid) {
            ExcelApi.legacyApi = undefined;
            for (var connectionUuid in ExcelApi.instance.applications) {
                ExcelApi.legacyApi = ExcelApi.instance.applications[connectionUuid].toObject();
                break;
            }
        }
    }
    // Internal Event Handlers
    processExcelConnectedEvent(data) {
        var applicationInstance = this.applications[data.uuid] || new ExcelApplication_1.ExcelApplication(data.uuid);
        this.applications[data.uuid] = applicationInstance;
        applicationInstance.init();
        // Synthetically raise connected event
        applicationInstance.processExcelEvent({ event: "connected" }, data.uuid);
        this.connectLegacyApi(data.uuid);
    }
    processExcelDisconnectedEvent(data) {
        delete this.applications[data.uuid];
        this.disconnectLegacyApi(data.uuid);
    }
    // Internal API Handlers
    processGetExcelInstancesResult(connectionUuids) {
        var oldInstances = this.applications;
        this.applications = {};
        connectionUuids.forEach(connectionUuid => {
            var applicationInstance = oldInstances[connectionUuid] || new ExcelApplication_1.ExcelApplication(connectionUuid);
            this.applications[connectionUuid] = applicationInstance;
            applicationInstance.init();
            this.connectLegacyApi(connectionUuid);
        });
    }
    // API Calls
    install(callback) {
        this.invokeServiceCall("install", null, callback);
    }
    getInstallationStatus(callback) {
        this.invokeServiceCall("getInstallationStatus", null, callback);
    }
    getExcelInstances(callback) {
        this.invokeServiceCall("getExcelInstances", null, callback);
    }
    toObject() {
        return {};
    }
    // Legacy API / Single-Application Functions
    static init() {
        ExcelApi.instance.init();
    }
    static addEventListener(type, listener) {
        ExcelApi.legacyApi.addEventListener(type, listener);
    }
    static removeEventListener(type, listener) {
        ExcelApi.legacyApi.removeEventListener(type, listener);
    }
    static run(callback) {
        if (ExcelApi.legacyApi && callback) {
            callback();
        }
        else {
            var connectedCallback = () => {
                ExcelApi.instance.removeEventListener("excelConnected", connectedCallback);
                callback && callback();
            };
            ExcelApi.instance.addEventListener("excelConnected", connectedCallback);
            fin.desktop.System.launchExternalProcess({
                target: "excel"
            });
        }
    }
    static install(callback) {
        ExcelApi.instance.install(callback);
    }
    static getInstallationStatus(callback) {
        ExcelApi.instance.getInstallationStatus(callback);
    }
    static getWorkbooks(callback) {
        ExcelApi.legacyApi.getWorkbooks(callback);
    }
    static getWorkbookByName(name) {
        return ExcelApi.legacyApi.getWorkbookByName(name);
    }
    static addWorkbook(callback) {
        ExcelApi.legacyApi.addWorkbook(callback);
    }
    static openWorkbook(path, callback) {
        ExcelApi.legacyApi.openWorkbook(path, callback);
    }
    static getConnectionStatus(callback) {
        if (ExcelApi.legacyApi) {
            ExcelApi.legacyApi.getConnectionStatus(callback);
        }
        else {
            callback(false);
        }
    }
    static getCalculationMode(callback) {
        ExcelApi.legacyApi.getCalculationMode(callback);
    }
    static calculateAll(callback) {
        ExcelApi.legacyApi.calculateAll(callback);
    }
}
ExcelApi.instance = new ExcelApi();
ExcelApi.legacyApi = undefined;
exports.ExcelApi = ExcelApi;
exports.LegacyApi = ExcelApi;
Object.defineProperty(exports, "__esModule", { value: true });
exports.default = ExcelApi.instance;
//# sourceMappingURL=ExcelApi.js.map

/***/ }),
/* 2 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

const RpcDispatcher_1 = __webpack_require__(0);
const ExcelWorkbook_1 = __webpack_require__(3);
const ExcelWorksheet_1 = __webpack_require__(4);
class ExcelApplication extends RpcDispatcher_1.RpcDispatcher {
    constructor(connectionUuid) {
        super();
        this.workbooks = {};
        this.processExcelEvent = (data, uuid) => {
            var eventType = data.event;
            var workbook = this.workbooks[data.workbookName];
            var worksheets = workbook && workbook.worksheets;
            var worksheet = worksheets && worksheets[data.sheetName];
            switch (eventType) {
                case "connected":
                    this.connected = true;
                    this.dispatchEvent(eventType);
                    break;
                case "sheetChanged":
                    if (worksheet) {
                        worksheet.dispatchEvent(eventType, { data: data });
                    }
                    break;
                case "sheetRenamed":
                    var oldWorksheetName = data.oldSheetName;
                    var newWorksheetName = data.sheetName;
                    worksheet = worksheets && worksheets[oldWorksheetName];
                    if (worksheet) {
                        delete worksheets[oldWorksheetName];
                        worksheet.worksheetName = newWorksheetName;
                        worksheets[worksheet.worksheetName] = worksheet;
                        workbook.dispatchEvent(eventType, { worksheet: worksheet.toObject(), oldWorksheetName: oldWorksheetName });
                    }
                    break;
                case "selectionChanged":
                    if (worksheet) {
                        worksheet.dispatchEvent(eventType, { data: data });
                    }
                    break;
                case "sheetActivated":
                case "sheetDeactivated":
                    if (worksheet) {
                        worksheet.dispatchEvent(eventType);
                    }
                    break;
                case "workbookDeactivated":
                case "workbookActivated":
                    if (workbook) {
                        workbook.dispatchEvent(eventType);
                    }
                    break;
                case "sheetAdded":
                    var newWorksheet = worksheet || new ExcelWorksheet_1.ExcelWorksheet(data.sheetName, workbook);
                    worksheets[newWorksheet.worksheetName] = newWorksheet;
                    workbook.dispatchEvent(eventType, { worksheet: newWorksheet.toObject() });
                    break;
                case "sheetRemoved":
                    delete workbook.worksheets[worksheet.worksheetName];
                    worksheet.dispatchEvent(eventType);
                    workbook.dispatchEvent(eventType, { worksheet: worksheet.toObject() });
                    break;
                case "workbookAdded":
                case "workbookOpened":
                    var newWorkbook = workbook || new ExcelWorkbook_1.ExcelWorkbook(this, data.workbookName);
                    this.workbooks[newWorkbook.workbookName] = newWorkbook;
                    this.dispatchEvent(eventType, { workbook: newWorkbook.toObject() });
                    break;
                case "workbookClosed":
                    delete this.workbooks[workbook.workbookName];
                    workbook.dispatchEvent(eventType);
                    this.dispatchEvent(eventType, { workbook: workbook.toObject() });
                    break;
                case "workbookSaved":
                    var oldWorkbookName = data.oldWorkbookName;
                    var newWorkbookName = data.workbookName;
                    workbook = this.workbooks[oldWorkbookName];
                    if (workbook) {
                        delete this.workbooks[oldWorkbookName];
                        workbook.workbookName = newWorkbookName;
                        this.workbooks[workbook.workbookName] = workbook;
                        this.dispatchEvent(eventType, { workbook: workbook.toObject(), oldWorkbookName: oldWorkbookName });
                    }
                    break;
                case "afterCalculation":
                default:
                    this.dispatchEvent(eventType);
                    break;
            }
        };
        this.processExcelResult = (result) => {
            var callbackData = {};
            var workbook = this.workbooks[result.target.workbookName];
            var worksheets = workbook && workbook.worksheets;
            var worksheet = worksheets && worksheets[result.target.sheetName];
            var resultData = result.data;
            switch (result.action) {
                case "getWorkbooks":
                    var workbookNames = resultData;
                    var oldworkbooks = this.workbooks;
                    this.workbooks = {};
                    workbookNames.forEach(workbookName => {
                        this.workbooks[workbookName] = oldworkbooks[workbookName] || new ExcelWorkbook_1.ExcelWorkbook(this, workbookName);
                    });
                    callbackData = workbookNames.map(workbookName => this.workbooks[workbookName].toObject());
                    break;
                case "getWorksheets":
                    var worksheetNames = resultData;
                    var oldworksheets = worksheets;
                    workbook.worksheets = {};
                    worksheetNames.forEach(worksheetName => {
                        workbook.worksheets[worksheetName] = oldworksheets[worksheetName] || new ExcelWorksheet_1.ExcelWorksheet(worksheetName, workbook);
                    });
                    callbackData = worksheetNames.map(worksheetName => workbook.worksheets[worksheetName].toObject());
                    break;
                case "addWorkbook":
                case "openWorkbook":
                    var newWorkbookName = resultData;
                    var newWorkbook = this.workbooks[newWorkbookName] || new ExcelWorkbook_1.ExcelWorkbook(this, newWorkbookName);
                    this.workbooks[newWorkbook.workbookName] = newWorkbook;
                    callbackData = newWorkbook.toObject();
                    break;
                case "addSheet":
                    var newWorksheetName = resultData;
                    var newWorksheet = workbook[newWorkbookName] || new ExcelWorksheet_1.ExcelWorksheet(newWorksheetName, workbook);
                    worksheets[newWorksheet.worksheetName] = newWorksheet;
                    callbackData = newWorksheet.toObject();
                    break;
                default:
                    callbackData = resultData;
                    break;
            }
            if (RpcDispatcher_1.RpcDispatcher.callbacks[result.messageId]) {
                RpcDispatcher_1.RpcDispatcher.callbacks[result.messageId](callbackData);
                delete RpcDispatcher_1.RpcDispatcher.callbacks[result.messageId];
            }
        };
        this.connectionUuid = connectionUuid;
    }
    init() {
        if (!this.initialized) {
            fin.desktop.InterApplicationBus.subscribe("*", "excelEvent", this.processExcelEvent);
            fin.desktop.InterApplicationBus.subscribe("*", "excelResult", this.processExcelResult);
            this.monitorDisconnect();
            this.initialized = true;
        }
    }
    monitorDisconnect() {
        fin.desktop.ExternalApplication.wrap(this.connectionUuid).addEventListener('disconnected', () => {
            this.connected = false;
            this.dispatchEvent('disconnected');
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
            fin.desktop.System.launchExternalProcess({
                target: 'excel',
                uuid: this.connectionUuid
            });
        }
    }
    getWorkbooks(callback) {
        this.invokeExcelCall("getWorkbooks", null, callback);
    }
    getWorkbookByName(name) {
        return this.workbooks[name];
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
            connectionUuid: this.connectionUuid,
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
exports.ExcelApplication = ExcelApplication;
//# sourceMappingURL=ExcelApplication.js.map

/***/ }),
/* 3 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

const RpcDispatcher_1 = __webpack_require__(0);
class ExcelWorkbook extends RpcDispatcher_1.RpcDispatcher {
    constructor(application, name) {
        super();
        this.worksheets = {};
        this.connectionUuid = application.connectionUuid;
        this.application = application;
        this.workbookName = name;
    }
    getDefaultMessage() {
        return {
            workbook: this.workbookName
        };
    }
    getWorksheets(callback) {
        this.invokeExcelCall("getWorksheets", null, callback);
    }
    getWorksheetByName(name) {
        return this.worksheets[name];
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
            name: this.workbookName,
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
/* 4 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

const RpcDispatcher_1 = __webpack_require__(0);
class ExcelWorksheet extends RpcDispatcher_1.RpcDispatcher {
    constructor(name, workbook) {
        super();
        this.connectionUuid = workbook.connectionUuid;
        this.workbook = workbook;
        this.worksheetName = name;
    }
    getDefaultMessage() {
        return {
            workbook: this.workbook.workbookName,
            worksheet: this.worksheetName
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
            name: this.worksheetName,
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
/* 5 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

// This is the entry point of the Plugin script
const ExcelApi_1 = __webpack_require__(1);
window.fin.desktop.Excel = ExcelApi_1.LegacyApi;
//# sourceMappingURL=plugin.js.map

/***/ })
/******/ ]);