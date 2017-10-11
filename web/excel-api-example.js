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

/**
 * Created by haseebriaz on 14/05/15.
 */

fin.desktop.Excel = __webpack_require__(1).LegacyApi;

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
        workbook.addEventListener("sheetRenamed", onWorksheetRenamed);

        addWorkbookTab(workbook.name);

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
        workbook.removeEventListener("sheetRenamed", onWorksheetRenamed);

        document.getElementById("workbookTabs").removeChild(document.getElementById(workbook.name));

        if (document.getElementById("workbookTabs").childNodes.length < 2) {
            setDisplayContainer(noWorkbooksContainer);
        }
    }

    function onWorkbookSaved(event) {
        var workbook = event.workbook;
        var oldWorkbookName = event.oldWorkbookName;

        var button = document.getElementById(oldWorkbookName);

        button.id = workbook.name;
        button.innerText = workbook.name;
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
        var worksheet = event.worksheet;

        if (worksheet.workbook === currentWorkbook) {
            worksheet.removeEventListener("sheetChanged", onSheetChanged);
            worksheet.removeEventListener("selectionChanged", onSelectionChanged);
            worksheet.removeEventListener("sheetActivated", onSheetActivated);
            document.getElementById("sheets").removeChild(document.getElementById(worksheet.name));
            currentWorksheet = null;
        }
    }

    function onWorksheetRenamed(event) {
        var worksheet = event.worksheet;
        var oldWorksheetName = event.oldWorksheetName;

        var button = document.getElementById(oldWorksheetName);
        button.id = worksheet.name;
        button.innerText = worksheet.name;
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

    function onExcelConnected() {
        console.log("Excel Connected: " + fin.desktop.Excel.legacyApi.connectionUuid);
        document.getElementById("status").innerText = "Connected to Excel";

        fin.desktop.Excel.instance.removeEventListener("excelConnected", onExcelConnected);

        // Grab a snapshot of the current instance, it can change!
        var legacyApi = fin.desktop.Excel.legacyApi;

        var onExcelDisconnected = function () {
            console.log("Excel Disconnected: " + legacyApi.connectionUuid);

            fin.desktop.Excel.instance.removeEventListener("excelDisconnected", onExcelDisconnected);
            legacyApi.removeEventListener("workbookAdded", onWorkbookAdded);
            legacyApi.removeEventListener("workbookOpened", onWorkbookAdded);
            legacyApi.removeEventListener("workbookClosed", onWorkbookRemoved);
            legacyApi.removeEventListener("workbookSaved", onWorkbookSaved);


            if (fin.desktop.Excel.legacyApi) {
                onExcelConnected();
            } else {
                document.getElementById("status").innerText = "Excel not connected";

                fin.desktop.Excel.instance.addEventListener("excelConnected", onExcelConnected);
                setDisplayContainer(noConnectionContainer);
            }
        }

        fin.desktop.Excel.instance.addEventListener("excelDisconnected", onExcelDisconnected);
        fin.desktop.Excel.addEventListener("workbookAdded", onWorkbookAdded);
        fin.desktop.Excel.addEventListener("workbookOpened", onWorkbookAdded);
        fin.desktop.Excel.addEventListener("workbookClosed", onWorkbookRemoved);
        fin.desktop.Excel.addEventListener("workbookSaved", onWorkbookSaved);

        fin.desktop.Excel.getWorkbooks(workbooks => {
            for (var i = 0; i < workbooks.length; i++) {
                addWorkbookTab(workbooks[i].name);
                workbooks[i].addEventListener("workbookActivated", onWorkbookActivated);
                workbooks[i].addEventListener("sheetAdded", onWorksheetAdded);
                workbooks[i].addEventListener("sheetRemoved", onWorksheetRemoved);
                workbooks[i].addEventListener("sheetRenamed", onWorksheetRenamed);

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


    function simluatePluginService() {
        var installFolder = '%localappdata%\\OpenFin\\shared\\assets\\excel-api-addin';
        var servicePath = 'OpenFin.ExcelService.exe';
        var addInPath = 'OpenFin.ExcelApi-AddIn.xll';

        var statusElement = document.getElementById("status");

        if (statusElement.innerText === "Connecting...") {
            return;
        }

        statusElement.innerText = "Connecting...";

        return Promise.resolve()
            .then(() => deployAddIn(servicePath, installFolder))
            .then(() => startExcelService(servicePath, installFolder))
            .then(() => registerAddIn(servicePath, installFolder));
    }

    function deployAddIn(servicePath, installFolder) {
        return new Promise((resolve, reject) => {
            console.log('Deploying Add-In');
            fin.desktop.System.launchExternalProcess({
                alias: 'excel-api-addin',
                target: servicePath,
                arguments: '-d "' + installFolder + '"',
                listener: function (args) {
                    console.log('Installer script completed! ' + args.exitCode);
                    resolve();
                }
            });
        });
    }

    function registerAddIn(servicePath, installFolder) {
        return new Promise((resolve, reject) => {
            console.log('Registering Add-In');
            fin.desktop.Excel.install(ack => {
                resolve();
            });
        });
    }

    function startExcelService(servicePath, installFolder) {
        var serviceUuid = '886834D1-4651-4872-996C-7B2578E953B9';

        return new Promise((resolve, reject) => {
            fin.desktop.System.getAllExternalApplications(extApps => {
                var excelServiceIndex = extApps.findIndex(extApp => extApp.uuid === serviceUuid);

                if (excelServiceIndex >= 0) {
                    console.log('Service Already Running');
                    resolve();
                    return;
                }

                var onServiceStarted = () => {
                    console.log('Service Started');
                    fin.desktop.Excel.instance.removeEventListener('started', onServiceStarted);
                    resolve();
                };

                chrome.desktop.getDetails(function (details) {
                    fin.desktop.Excel.instance.addEventListener('started', onServiceStarted);

                    fin.desktop.System.launchExternalProcess({
                        target: installFolder + '\\OpenFin.ExcelService.exe',
                        arguments: '-p ' + details.port,
                        uuid: serviceUuid,
                    }, process => {
                        console.log('Service Launched: ' + process.uuid);
                    }, error => {
                        reject('Error starting Excel service');
                    });
                });
            });
        });
    }

    function connectToExcel() {
        return new Promise((resolve, reject) => {
            fin.desktop.Excel.instance.getExcelInstances(instances => {
                if (instances.length > 0) {
                    console.log("Excel Already Running");
                    resolve();
                } else {
                    console.log("Launching Excel");
                    fin.desktop.Excel.run(resolve);
                }
            });
        });
    }

    initTable(27, 12);

    fin.desktop.main(function () {
       fin.desktop.Excel.init();

        Promise.resolve()
            .then(simluatePluginService)
            .then(connectToExcel)
            .then(onExcelConnected)
            .catch(err => console.log(err));
    });
});


/***/ })
/******/ ]);