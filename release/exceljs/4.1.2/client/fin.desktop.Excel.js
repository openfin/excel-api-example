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
/******/ 	return __webpack_require__(__webpack_require__.s = 9);
/******/ })
/************************************************************************/
/******/ ([
/* 0 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

Object.defineProperty(exports, "__esModule", { value: true });
exports.RpcDispatcher = void 0;
const EventEmitter_1 = __webpack_require__(1);
const NoOpLogger_1 = __webpack_require__(2);
class RpcDispatcher extends EventEmitter_1.EventEmitter {
    constructor(logger) {
        super();
        this.logger = new NoOpLogger_1.NoOpLogger();
        if (logger !== undefined) {
            this.logger = logger;
        }
    }
    getDefaultMessage() {
        return {};
    }
    invokeExcelCall(functionName, data, callback) {
        return this.invokeRemoteCall('excelCall', functionName, data, callback);
    }
    invokeServiceCall(functionName, data, callback) {
        return this.invokeRemoteCall('excelServiceCall', functionName, data, callback);
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
        var executor;
        var promise = new Promise((resolve, reject) => {
            executor = {
                resolve,
                reject
            };
        });
        // Legacy Callback-style API
        promise = this.applyCallbackToPromise(promise, callback);
        var currentMessageId = RpcDispatcher.messageId;
        RpcDispatcher.messageId++;
        if (this.connectionUuid !== undefined) {
            RpcDispatcher.promiseExecutors[currentMessageId] = executor;
            fin.desktop.InterApplicationBus.send(this.connectionUuid, topic, message, ack => {
                // TODO: log once we support configurable logging.
            }, nak => {
                delete RpcDispatcher.promiseExecutors[currentMessageId];
                executor.reject(new Error(nak));
            });
        }
        else {
            executor.reject(new Error('The target UUID of the remote call is undefined.'));
        }
        return promise;
    }
    applyCallbackToPromise(promise, callback) {
        if (callback) {
            promise
                .then(result => {
                callback(result);
                return result;
            }).catch(err => {
                this.logger.error("unable to apply callback to promise", err);
            });
            promise = undefined;
        }
        return promise;
    }
}
exports.RpcDispatcher = RpcDispatcher;
RpcDispatcher.messageId = 1;
RpcDispatcher.promiseExecutors = {};
//# sourceMappingURL=RpcDispatcher.js.map

/***/ }),
/* 1 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

Object.defineProperty(exports, "__esModule", { value: true });
exports.EventEmitter = void 0;
class EventEmitter {
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
}
exports.EventEmitter = EventEmitter;
//# sourceMappingURL=EventEmitter.js.map

/***/ }),
/* 2 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

Object.defineProperty(exports, "__esModule", { value: true });
exports.NoOpLogger = void 0;
class NoOpLogger {
    constructor() {
    }
    trace(message, ...args) {
    }
    debug(message, ...args) {
    }
    info(message, ...args) {
    }
    warn(message, ...args) {
    }
    error(message, error, ...args) {
    }
    fatal(message, error, ...args) {
    }
}
exports.NoOpLogger = NoOpLogger;
//# sourceMappingURL=NoOpLogger.js.map

/***/ }),
/* 3 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.ExcelService = void 0;
const RpcDispatcher_1 = __webpack_require__(0);
const ExcelApplication_1 = __webpack_require__(5);
const ExcelRtd_1 = __webpack_require__(6);
const DefaultLogger_1 = __webpack_require__(4);
const NoOpLogger_1 = __webpack_require__(2);
const excelServiceUuid = "886834D1-4651-4872-996C-7B2578E953B9";
class ExcelService extends RpcDispatcher_1.RpcDispatcher {
    constructor() {
        super(new NoOpLogger_1.NoOpLogger());
        this.defaultApplicationUuid = undefined;
        this.defaultApplicationObj = undefined;
        this.logger = new NoOpLogger_1.NoOpLogger();
        this.loggerName = "ExcelService";
        this.applications = {};
        this.version = {
            buildVersion: "0.0.0.0", "providerVersion": "0.0.0"
        };
        this.processExcelServiceEvent = (data) => __awaiter(this, void 0, void 0, function* () {
            var eventType = data.event;
            this.logger.debug(this.loggerName + ": Received event for data...");
            this.logger.debug(JSON.stringify(data));
            const mainChannelCreated = yield this.mainChannelCreated;
            this.logger.debug(this.loggerName + `: Main Channel created... ${mainChannelCreated}`);
            var eventData;
            switch (data.event) {
                case "started":
                    break;
                case "registrationRollCall":
                    if (this.initialized) {
                        this.logger.debug(this.loggerName + ": Initialized, about to register window instance.");
                        this.registerWindowInstance();
                    }
                    else {
                        this.logger.debug(this.loggerName + ": NOT initialized. Window will not be registered.");
                    }
                    break;
                case "excelConnected":
                    yield this.processExcelConnectedEvent(data);
                    eventData = { connectionUuid: data.uuid };
                    break;
                case "excelDisconnected":
                    yield this.processExcelDisconnectedEvent(data);
                    eventData = { connectionUuid: data.uuid };
                    break;
            }
            this.dispatchEvent(eventType, eventData);
        });
        this.processExcelServiceResult = (result) => __awaiter(this, void 0, void 0, function* () {
            var executor = RpcDispatcher_1.RpcDispatcher.promiseExecutors[result.messageId];
            delete RpcDispatcher_1.RpcDispatcher.promiseExecutors[result.messageId];
            this.logger.debug(this.loggerName + `: Received an ExcelService result with messageId ${result.messageId}.`);
            //TODO: Somehow received a result not in the callback map
            if (!executor) {
                this.logger.debug(this.loggerName + `: Received an ExcelService result for messageId ${result.messageId} that doesnt have an associated promise executor.`);
                return;
            }
            if (result.error) {
                this.logger.debug(this.loggerName + `: Received a result with error ${result.error}.`);
                executor.reject(result.error);
                return;
            }
            // Internal processing
            switch (result.action) {
                case "getExcelInstances":
                    yield this.processGetExcelInstancesResult(result.data);
                    break;
            }
            this.logger.debug(this.loggerName + `: Calling resolver for message ${result.messageId} with data ${JSON.stringify(result.data)}.`);
            executor.resolve(result.data);
        });
        this.registerWindowInstance = (callback) => {
            return this.invokeServiceCall("registerOpenfinWindow", { domain: document.domain }, callback);
        };
        this.connectionUuid = excelServiceUuid;
        this.setMainChanelCreated();
    }
    setMainChanelCreated() {
        this.mainChannelCreated = new Promise((resolve, reject) => {
            this.mainChannelResolve = resolve;
            this.mainChannelReject = reject;
        });
    }
    init(logger) {
        return __awaiter(this, void 0, void 0, function* () {
            if (!this.initialized) {
                if (logger !== undefined) {
                    if (typeof logger === "boolean") {
                        if (logger) {
                            let defaultLogger = new DefaultLogger_1.DefaultLogger("Excel Adapter");
                            this.logger = defaultLogger;
                        }
                    }
                    else {
                        let defaultLogger = new DefaultLogger_1.DefaultLogger(logger.name || "Excel Adapter");
                        this.logger = Object.assign({}, logger);
                        if (this.logger.name === undefined) {
                            this.logger.name === defaultLogger.name;
                        }
                        if (this.logger.trace === undefined) {
                            this.logger.trace = defaultLogger.trace;
                        }
                        if (this.logger.debug === undefined) {
                            this.logger.debug = defaultLogger.debug;
                        }
                        if (this.logger.info === undefined) {
                            this.logger.info = defaultLogger.info;
                        }
                        if (this.logger.warn === undefined) {
                            this.logger.warn = defaultLogger.warn;
                        }
                        if (this.logger.error === undefined) {
                            this.logger.error = defaultLogger.error;
                        }
                        if (this.logger.fatal === undefined) {
                            this.logger.fatal = defaultLogger.fatal;
                        }
                    }
                }
                this.logger.info(this.loggerName + ": Initialised called.");
                this.logger.debug(this.loggerName + ": Subscribing to Service Messages.");
                yield this.subscribeToServiceMessages();
                this.logger.debug(this.loggerName + ": Ensuring monitor is not conencted before connecting to channel.");
                yield this.monitorDisconnect();
                try {
                    this.logger.debug(this.loggerName + ": Connecting to channel: " + excelServiceUuid);
                    let providerChannel = yield fin.desktop.InterApplicationBus.Channel.connect(excelServiceUuid);
                    this.logger.debug(this.loggerName + ": Channel connected: " + excelServiceUuid);
                    this.logger.debug(this.loggerName + ": Flagging that mainChannelIs connected: " + excelServiceUuid);
                    this.logger.debug(this.loggerName + ": Setting service provider version by requesting it from channel.");
                    this.version = yield providerChannel.dispatch('getVersion');
                    this.logger.debug(this.loggerName + `: Service provider version set to: ${JSON.stringify(this.version)}.`);
                }
                catch (err) {
                    let errorMessage;
                    if (err !== undefined && err.message !== undefined) {
                        errorMessage = "Error: " + err.message;
                    }
                    this.mainChannelReject(`${this.loggerName}: Error connecting or fetching version to/from provider. The version of the provider is likely older than the script version. ${errorMessage}`);
                    this.logger.warn(this.loggerName + ": Error connecting or fetching version to/from provider. The version of the provider is likely older than the script version.", errorMessage);
                }
                this.initialized = true;
                this.mainChannelResolve(true);
                yield this.registerWindowInstance();
                yield this.getExcelInstances();
            }
            return;
        });
    }
    subscribeToServiceMessages() {
        return Promise.all([
            new Promise(resolve => fin.desktop.InterApplicationBus.subscribe(excelServiceUuid, "excelServiceEvent", this.processExcelServiceEvent, resolve)),
            new Promise(resolve => fin.desktop.InterApplicationBus.subscribe(excelServiceUuid, "excelServiceCallResult", this.processExcelServiceResult, resolve))
        ]);
    }
    monitorDisconnect() {
        return new Promise((resolve, reject) => {
            var excelServiceConnection = fin.desktop.ExternalApplication.wrap(excelServiceUuid);
            var onDisconnect;
            excelServiceConnection.addEventListener("disconnected", onDisconnect = () => {
                excelServiceConnection.removeEventListener("disconnected", onDisconnect);
                this.dispatchEvent("stopped");
            }, resolve, reject);
        });
    }
    configureDefaultApplication() {
        return __awaiter(this, void 0, void 0, function* () {
            this.logger.debug(this.loggerName + ": Configuring Default Excel Application.");
            var defaultAppObjUuid = this.defaultApplicationObj && this.defaultApplicationObj.connectionUuid;
            var defaultAppEntry = this.applications[defaultAppObjUuid];
            var defaultAppObjConnected = defaultAppEntry ? defaultAppEntry.connected : false;
            if (defaultAppObjConnected) {
                this.logger.debug(this.loggerName + ": Already connected to Default Excel Application: " + defaultAppObjUuid);
                return;
            }
            else {
                this.logger.debug(this.loggerName + ": Default Excel Application: " + defaultAppObjUuid + " not connected.");
            }
            this.logger.debug(this.loggerName + ": As Default Excel Application not connected checking for existing connected instance.");
            var connectedAppUuid = Object.keys(this.applications).find(appUuid => this.applications[appUuid].connected);
            if (connectedAppUuid) {
                this.logger.debug(this.loggerName + ": Found connected Excel Application: " + connectedAppUuid + " setting it as default instance.");
                delete this.applications[defaultAppObjUuid];
                this.defaultApplicationObj = this.applications[connectedAppUuid].toObject();
                return;
            }
            if (defaultAppEntry === undefined) {
                var disconnectedAppUuid = fin.desktop.getUuid();
                this.logger.debug(this.loggerName + ": No default Excel Application. Creating one with id: " + disconnectedAppUuid + " and setting it as default instance.");
                var disconnectedApp = new ExcelApplication_1.ExcelApplication(disconnectedAppUuid, this.logger);
                yield disconnectedApp.init();
                this.applications[disconnectedAppUuid] = disconnectedApp;
                this.defaultApplicationObj = disconnectedApp.toObject();
                this.logger.debug(this.loggerName + ": Default Excel Application with id: " + disconnectedAppUuid + " set as default instance.");
            }
        });
    }
    // Internal Event Handlers
    processExcelConnectedEvent(data) {
        return __awaiter(this, void 0, void 0, function* () {
            var applicationInstance = this.applications[data.uuid] || new ExcelApplication_1.ExcelApplication(data.uuid, this.logger);
            yield applicationInstance.init();
            this.applications[data.uuid] = applicationInstance;
            // Synthetically raise connected event
            applicationInstance.processExcelEvent({ event: "connected" }, data.uuid);
            yield this.configureDefaultApplication();
            return;
        });
    }
    processExcelDisconnectedEvent(data) {
        return __awaiter(this, void 0, void 0, function* () {
            var applicationInstance = this.applications[data.uuid];
            if (applicationInstance === undefined) {
                return;
            }
            delete this.applications[data.uuid];
            yield this.configureDefaultApplication();
            yield applicationInstance.release();
        });
    }
    // Internal API Handlers
    processGetExcelInstancesResult(connectionUuids) {
        return __awaiter(this, void 0, void 0, function* () {
            var existingInstances = this.applications;
            this.applications = {};
            yield Promise.all(connectionUuids.map((connectionUuid) => __awaiter(this, void 0, void 0, function* () {
                var existingInstance = existingInstances[connectionUuid];
                if (existingInstance === undefined) {
                    // Assume that since the ExcelService is aware of the instance it,
                    // is connected and to simulate the the connection event
                    yield this.processExcelServiceEvent({ event: "excelConnected", uuid: connectionUuid });
                }
                else {
                    this.applications[connectionUuid] = existingInstance;
                }
                return;
            })));
            yield this.configureDefaultApplication();
        });
    }
    // API Calls
    install(callback) {
        return this.invokeServiceCall("install", null, callback);
    }
    getInstallationStatus(callback) {
        return this.invokeServiceCall("getInstallationStatus", null, callback);
    }
    getExcelInstances(callback) {
        return this.invokeServiceCall("getExcelInstances", null, callback);
    }
    createRtd(providerName, heartbeatIntervalInMilliseconds = 10000) {
        return ExcelRtd_1.ExcelRtd.create(providerName, this.logger, heartbeatIntervalInMilliseconds);
    }
    toObject() {
        return {};
    }
}
exports.ExcelService = ExcelService;
ExcelService.instance = new ExcelService();
//# sourceMappingURL=ExcelApi.js.map

/***/ }),
/* 4 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

Object.defineProperty(exports, "__esModule", { value: true });
exports.DefaultLogger = void 0;
class DefaultLogger {
    constructor(name) {
        this.name = name || "logger";
    }
    trace(message, ...args) {
        console.log(this.name + ": " + message, ...args);
    }
    debug(message, ...args) {
        console.log(this.name + ": " + message, ...args);
    }
    info(message, ...args) {
        console.info(this.name + ": " + message, ...args);
    }
    warn(message, ...args) {
        console.warn(this.name + ": " + message, ...args);
    }
    error(message, error, ...args) {
        console.error(this.name + ": " + message, error, ...args);
    }
    fatal(message, error, ...args) {
        console.error(this.name + ": " + message, error, ...args);
    }
}
exports.DefaultLogger = DefaultLogger;
//# sourceMappingURL=DefaultLogger.js.map

/***/ }),
/* 5 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.ExcelApplication = void 0;
const RpcDispatcher_1 = __webpack_require__(0);
const ExcelWorkbook_1 = __webpack_require__(7);
const ExcelWorksheet_1 = __webpack_require__(8);
class ExcelApplication extends RpcDispatcher_1.RpcDispatcher {
    constructor(connectionUuid, logger) {
        super(logger);
        this.workbooks = {};
        this.version = { clientVersion: "4.1.2", buildVersion: "4.1.2.0" };
        this.loggerName = "ExcelApplication";
        this.processExcelEvent = (data, uuid) => {
            var eventType = data.event;
            this.logger.debug(this.loggerName + `: ExcelApplication.processExcelEvent received from ${uuid} with data ${JSON.stringify(data)}.`);
            var workbook = this.workbooks[data.workbookName];
            var worksheets = workbook && workbook.worksheets;
            var worksheet = worksheets && worksheets[data.sheetName];
            switch (eventType) {
                case "connected":
                    this.logger.debug(this.loggerName + ": Received Excel Connected Event.");
                    this.connected = true;
                    this.dispatchEvent(eventType);
                    break;
                case "sheetChanged":
                    this.logger.debug(this.loggerName + ": Received Sheet Changed Event.");
                    if (worksheet) {
                        this.logger.debug(this.loggerName + ": Worksheet Exists, dispatching event.");
                        worksheet.dispatchEvent(eventType, { data: data });
                    }
                    break;
                case "sheetRenamed":
                    var oldWorksheetName = data.oldSheetName;
                    var newWorksheetName = data.sheetName;
                    worksheet = worksheets && worksheets[oldWorksheetName];
                    this.logger.debug(this.loggerName + ": Sheet renamed: Old name:" + oldWorksheetName + " New name: " + newWorksheetName);
                    if (worksheet) {
                        delete worksheets[oldWorksheetName];
                        worksheet.worksheetName = newWorksheetName;
                        worksheets[worksheet.worksheetName] = worksheet;
                        workbook.dispatchEvent(eventType, { worksheet: worksheet.toObject(), oldWorksheetName: oldWorksheetName });
                    }
                    break;
                case "selectionChanged":
                    this.logger.debug(this.loggerName + ": Selection Changed.");
                    if (worksheet) {
                        worksheet.dispatchEvent(eventType, { data: data });
                    }
                    break;
                case "sheetActivated":
                case "sheetDeactivated":
                    this.logger.debug(this.loggerName + ": Sheet Deactivated");
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
                        workbook.refreshObject();
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
            var executor = RpcDispatcher_1.RpcDispatcher.promiseExecutors[result.messageId];
            delete RpcDispatcher_1.RpcDispatcher.promiseExecutors[result.messageId];
            this.logger.debug(this.loggerName + `: Received an Excel result with messageId ${result.messageId}.`);
            //TODO: Somehow received a result not in the callback map
            if (!executor) {
                this.logger.debug(this.loggerName + `: Received an Excel result for messageId ${result.messageId} that doesnt have an associated promise executor.`);
                return;
            }
            if (result.error) {
                this.logger.debug(this.loggerName + `: Received an Excel result with error ${result.error}.`);
                executor.reject(result.error);
                return;
            }
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
            this.logger.debug(this.loggerName + `: Calling resolver for Excel message ${result.messageId} with data ${callbackData}.`);
            executor.resolve(callbackData);
        };
        this.connectionUuid = connectionUuid;
        this.connected = false;
    }
    init() {
        return __awaiter(this, void 0, void 0, function* () {
            this.logger.info(this.loggerName + ": Init called.");
            if (!this.initialized) {
                this.logger.info(this.loggerName + ": Not initialised...Initialising.");
                yield this.subscribeToExcelMessages();
                yield this.monitorDisconnect();
                this.initialized = true;
                this.logger.info(this.loggerName + ": initialised.");
            }
            return;
        });
    }
    release() {
        return __awaiter(this, void 0, void 0, function* () {
            this.logger.info(this.loggerName + ": Release called.");
            if (this.initialized) {
                this.logger.info(this.loggerName + ": Calling unsubscribe as we are currently intialised.");
                yield this.unsubscribeToExcelMessages();
                //TODO: Provide external means to stop monitoring disconnect
            }
            return;
        });
    }
    subscribeToExcelMessages() {
        return Promise.all([
            new Promise(resolve => fin.desktop.InterApplicationBus.subscribe(this.connectionUuid, "excelEvent", this.processExcelEvent, resolve)),
            new Promise(resolve => fin.desktop.InterApplicationBus.subscribe(this.connectionUuid, "excelResult", this.processExcelResult, resolve))
        ]);
    }
    unsubscribeToExcelMessages() {
        return Promise.all([
            new Promise(resolve => fin.desktop.InterApplicationBus.unsubscribe(this.connectionUuid, "excelEvent", this.processExcelEvent, resolve)),
            new Promise(resolve => fin.desktop.InterApplicationBus.unsubscribe(this.connectionUuid, "excelResult", this.processExcelResult, resolve))
        ]);
    }
    monitorDisconnect() {
        return new Promise((resolve, reject) => {
            var excelApplicationConnection = fin.desktop.ExternalApplication.wrap(this.connectionUuid);
            var onDisconnect;
            excelApplicationConnection.addEventListener('disconnected', onDisconnect = () => {
                excelApplicationConnection.removeEventListener('disconnected', onDisconnect);
                this.connected = false;
                this.dispatchEvent('disconnected');
            }, resolve, reject);
        });
    }
    run(callback) {
        var runPromise = this.connected ? Promise.resolve() : new Promise(resolve => {
            var connectedCallback = () => {
                this.removeEventListener('connected', connectedCallback);
                resolve();
            };
            if (this.connectionUuid !== undefined) {
                this.addEventListener('connected', connectedCallback);
            }
            fin.desktop.System.launchExternalProcess({
                target: 'excel',
                uuid: this.connectionUuid
            });
        });
        return this.applyCallbackToPromise(runPromise, callback);
    }
    getWorkbooks(callback) {
        return this.invokeExcelCall("getWorkbooks", null, callback);
    }
    getWorkbookByName(name) {
        return this.workbooks[name];
    }
    addWorkbook(callback) {
        return this.invokeExcelCall("addWorkbook", null, callback);
    }
    openWorkbook(path, callback) {
        return this.invokeExcelCall("openWorkbook", { path: path }, callback);
    }
    getConnectionStatus(callback) {
        return this.applyCallbackToPromise(Promise.resolve(this.connected), callback);
    }
    getCalculationMode(callback) {
        return this.invokeExcelCall("getCalculationMode", null, callback);
    }
    calculateAll(callback) {
        return this.invokeExcelCall("calculateFull", null, callback);
    }
    toObject() {
        return this.objectInstance || (this.objectInstance = {
            connectionUuid: this.connectionUuid,
            version: this.version,
            addEventListener: this.addEventListener.bind(this),
            removeEventListener: this.removeEventListener.bind(this),
            addWorkbook: this.addWorkbook.bind(this),
            calculateAll: this.calculateAll.bind(this),
            getCalculationMode: this.getCalculationMode.bind(this),
            getConnectionStatus: this.getConnectionStatus.bind(this),
            getWorkbookByName: name => this.getWorkbookByName(name).toObject(),
            getWorkbooks: this.getWorkbooks.bind(this),
            openWorkbook: this.openWorkbook.bind(this),
            run: this.run.bind(this)
        });
    }
}
exports.ExcelApplication = ExcelApplication;
ExcelApplication.defaultInstance = undefined;
//# sourceMappingURL=ExcelApplication.js.map

/***/ }),
/* 6 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.ExcelRtd = void 0;
const EventEmitter_1 = __webpack_require__(1);
class ExcelRtd extends EventEmitter_1.EventEmitter {
    constructor(providerName, logger, heartbeatIntervalInMilliseconds = 10000) {
        super();
        this.heartbeatIntervalInMilliseconds = heartbeatIntervalInMilliseconds;
        this.listeners = {};
        this.connectedTopics = {};
        this.connectedKey = 'connected';
        this.disconnectedKey = 'disconnected';
        this.loggerName = "ExcelRtd";
        this.initialized = false;
        this.disposed = false;
        var minimumDefaultHeartbeat = 10000;
        if (this.heartbeatIntervalInMilliseconds < minimumDefaultHeartbeat) {
            logger.warn(`heartbeatIntervalInMilliseconds cannot be less than ${minimumDefaultHeartbeat}. Setting heartbeatIntervalInMilliseconds to ${minimumDefaultHeartbeat}.`);
            this.heartbeatIntervalInMilliseconds = minimumDefaultHeartbeat;
        }
        this.providerName = providerName;
        this.logger = logger;
        logger.debug(this.loggerName + ": instance created for provider: " + providerName);
    }
    static create(providerName, logger, heartbeatIntervalInMilliseconds = 10000) {
        return __awaiter(this, void 0, void 0, function* () {
            logger.debug("ExcelRtd: create called to create provider: " + providerName);
            const instance = new ExcelRtd(providerName, logger, heartbeatIntervalInMilliseconds);
            yield instance.init();
            if (!instance.isInitialized) {
                return undefined;
            }
            return instance;
        });
    }
    init() {
        return __awaiter(this, void 0, void 0, function* () {
            if (this.isInitialized) {
                return;
            }
            this.logger.debug(this.loggerName + ": Initialise called for provider: " + this.providerName);
            try {
                // A channel is created to ensure it is a singleton so you don't have two apps pushing updates over each other or two windows within the same app
                this.provider = yield fin.InterApplicationBus.Channel.create(`excelRtd/${this.providerName}`);
            }
            catch (err) {
                this.logger.warn(this.loggerName + `: The excelRtd/${this.providerName} channel already exists. You can only have one instance of a connection for a provider to avoid confusion. It may be you have multiple instances or another window or application has created a provider with the same name.`, err);
                return;
            }
            this.logger.debug(this.loggerName + `: Subscribing to messages to this provider (${this.providerName}) from excel.`);
            yield fin.InterApplicationBus.subscribe({ uuid: '*' }, `excelRtd/pong/${this.providerName}`, this.onSubscribe.bind(this));
            yield fin.InterApplicationBus.subscribe({ uuid: '*' }, `excelRtd/ping-request/${this.providerName}`, this.ping.bind(this));
            yield fin.InterApplicationBus.subscribe({ uuid: '*' }, `excelRtd/unsubscribed/${this.providerName}`, this.onUnsubscribe.bind(this));
            yield this.ping();
            this.establishHeartbeat();
            this.logger.debug(this.loggerName + `: initialisation for provider (${this.providerName}) finished.`);
            this.initialized = true;
        });
    }
    get isDisposed() {
        return this.disposed;
    }
    get isInitialized() {
        return this.initialized;
    }
    setValue(topic, value) {
        this.logger.trace(this.loggerName + `: Publishing on rtdTopic: ${topic} and provider: ${this.providerName} value: ${JSON.stringify(value)}`);
        fin.InterApplicationBus.publish(`excelRtd/data/${this.providerName}/${topic}`, value);
    }
    dispose() {
        return __awaiter(this, void 0, void 0, function* () {
            if (!this.disposed) {
                this.logger.debug(this.loggerName + `: dispose called. Will send message to clear values for this provider (${this.providerName}).`);
                if (this.heartbeatToken) {
                    clearInterval(this.heartbeatToken);
                }
                this.clear();
                if (this.provider !== undefined) {
                    try {
                        yield this.provider.destroy();
                    }
                    catch (err) {
                        // without a catch the rest of the initialisation would be broken
                        this.logger.warn(this.loggerName + `: The excelRtd/${this.providerName} channel could not be destroyed during cleanup.`, err);
                    }
                }
                this.logger.debug(this.loggerName + `: UnSubscribing to messages to this provider (${this.providerName}) from excel.`);
                yield fin.InterApplicationBus.unsubscribe({ uuid: '*' }, `excelRtd/pong/${this.providerName}`, this.onSubscribe.bind(this));
                yield fin.InterApplicationBus.unsubscribe({ uuid: '*' }, `excelRtd/ping-request/${this.providerName}`, this.ping.bind(this));
                yield fin.InterApplicationBus.unsubscribe({ uuid: '*' }, `excelRtd/unsubscribed/${this.providerName}`, this.onUnsubscribe.bind(this));
                this.disposed = true;
                this.initialized = false;
            }
            else {
                this.logger.debug(this.loggerName + `: This provider (${this.providerName}) has already been disposed.`);
            }
        });
    }
    // Overriding
    addEventListener(type, listener) {
        this.logger.debug(this.loggerName + `: Event listener add requested for type ${type} received.`);
        if (super.hasEventListener(type, listener)) {
            this.logger.debug(this.loggerName + `: Event listener add requested for type ${type} received.`);
            return;
        }
        let connectedTopicIds = Object.keys(this.connectedTopics);
        let topics = this.connectedTopics;
        if (connectedTopicIds.length > 0) {
            // need to simulate async action as by default this method would return and then a listener would be called
            setTimeout(() => {
                connectedTopicIds.forEach(id => {
                    this.logger.debug(this.loggerName + `: Raising synthetic event as the event listener was added after the event for connected for rtdTopic: ${id}.`);
                    listener(topics[id]);
                });
            }, 0);
        }
        super.addEventListener(type, listener);
    }
    dispatchEvent(evtOrTypeArg, data) {
        var event;
        if (typeof evtOrTypeArg == "string" && data !== undefined) {
            this.logger.debug(this.loggerName + `: dispatch event called for type ${evtOrTypeArg} and data: ${JSON.stringify(data)}`);
            event = Object.assign({
                target: this.toObject(),
                type: evtOrTypeArg,
                defaultPrevented: false
            }, data);
            if (data.topic !== undefined) {
                if (evtOrTypeArg === this.connectedKey) {
                    this.connectedTopics[data.topic] = event;
                    this.logger.debug(this.loggerName + `: Saving connected event for rtdTopic: ${data.topic}.`);
                }
                else if (evtOrTypeArg === this.disconnectedKey) {
                    this.logger.debug(this.loggerName + `: Disconnected event for rtdTopic: ${data.topic} received.`);
                    if (this.connectedTopics[data.topic] !== undefined) {
                        // we have removed the topic so clear it from the connected list for late subscribers
                        this.logger.debug(this.loggerName + `: Clearing saved connected event for rtdTopic: ${data.topic}.`);
                        delete this.connectedTopics[data.topic];
                    }
                }
            }
            this.logger.debug(this.loggerName + `: Dispatching event.`);
            return super.dispatchEvent(event, data);
        }
        event = evtOrTypeArg;
        return super.dispatchEvent(event);
    }
    toObject() {
        return this;
    }
    ping(topic) {
        return __awaiter(this, void 0, void 0, function* () {
            if (topic !== undefined) {
                this.pingPath = `excelRtd/ping/${this.providerName}/${topic}`;
            }
            else {
                this.pingPath = `excelRtd/ping/${this.providerName}`;
            }
            this.logger.debug(this.loggerName + `: Publishing ping message for this provider (${this.providerName}) to excel on topic: ${this.pingPath}.`);
            yield fin.InterApplicationBus.publish(`${this.pingPath}`, true);
        });
    }
    establishHeartbeat() {
        this.heartbeatPath = `excelRtd/heartbeat/${this.providerName}`;
        this.heartbeatToken = setInterval(() => {
            this.logger.debug(`Heartbeating for ${this.heartbeatPath}.`);
            fin.InterApplicationBus.publish(`${this.heartbeatPath}`, this.heartbeatIntervalInMilliseconds);
        }, this.heartbeatIntervalInMilliseconds);
    }
    onSubscribe(topic) {
        this.logger.debug(this.loggerName + `: Subscription for rtdTopic ${topic} found. Dispatching connected event for rtdTopic.`);
        this.dispatchEvent(this.connectedKey, { topic });
    }
    onUnsubscribe(topic) {
        this.logger.debug(this.loggerName + `: Unsubscribe for rtdTopic ${topic}. Dispatching disconnected event for rtdTopic.`);
        this.dispatchEvent(this.disconnectedKey, { topic });
    }
    clear() {
        let path = `excelRtd/clear/${this.providerName}`;
        this.logger.debug(this.loggerName + `: Clear called. Publishing to excel on topic: ${path} `);
        fin.InterApplicationBus.publish(`excelRtd/clear/${this.providerName}`, true);
    }
}
exports.ExcelRtd = ExcelRtd;
//# sourceMappingURL=ExcelRtd.js.map

/***/ }),
/* 7 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

Object.defineProperty(exports, "__esModule", { value: true });
exports.ExcelWorkbook = void 0;
const RpcDispatcher_1 = __webpack_require__(0);
class ExcelWorkbook extends RpcDispatcher_1.RpcDispatcher {
    constructor(application, name) {
        super(application.logger);
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
        return this.invokeExcelCall("getWorksheets", null, callback);
    }
    getWorksheetByName(name) {
        return this.worksheets[name];
    }
    addWorksheet(callback) {
        return this.invokeExcelCall("addSheet", null, callback);
    }
    activate() {
        return this.invokeExcelCall("activateWorkbook");
    }
    save() {
        return this.invokeExcelCall("saveWorkbook");
    }
    close() {
        return this.invokeExcelCall("closeWorkbook");
    }
    refreshObject() {
        this.objectInstance = null;
        this.toObject();
    }
    toObject() {
        return this.objectInstance || (this.objectInstance = {
            addEventListener: this.addEventListener.bind(this),
            removeEventListener: this.removeEventListener.bind(this),
            name: this.workbookName,
            activate: this.activate.bind(this),
            addWorksheet: this.addWorksheet.bind(this),
            close: this.close.bind(this),
            getWorksheetByName: name => this.getWorksheetByName(name).toObject(),
            getWorksheets: this.getWorksheets.bind(this),
            save: this.save.bind(this)
        });
    }
}
exports.ExcelWorkbook = ExcelWorkbook;
//# sourceMappingURL=ExcelWorkbook.js.map

/***/ }),
/* 8 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

Object.defineProperty(exports, "__esModule", { value: true });
exports.ExcelWorksheet = void 0;
const RpcDispatcher_1 = __webpack_require__(0);
class ExcelWorksheet extends RpcDispatcher_1.RpcDispatcher {
    constructor(name, workbook) {
        super(workbook.logger);
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
        return this.invokeExcelCall("setCells", { offset: offset, values: values });
    }
    getCells(start, offsetWidth, offsetHeight, callback) {
        return this.invokeExcelCall("getCells", { start: start, offsetWidth: offsetWidth, offsetHeight: offsetHeight }, callback);
    }
    getRow(start, width, callback) {
        return this.invokeExcelCall("getCellsRow", { start: start, offsetWidth: width }, callback);
    }
    getColumn(start, offsetHeight, callback) {
        return this.invokeExcelCall("getCellsColumn", { start: start, offsetHeight: offsetHeight }, callback);
    }
    activate() {
        return this.invokeExcelCall("activateSheet");
    }
    activateCell(cellAddress) {
        return this.invokeExcelCall("activateCell", { address: cellAddress });
    }
    addButton(name, caption, cellAddress) {
        return this.invokeExcelCall("addButton", { address: cellAddress, buttonName: name, buttonCaption: caption });
    }
    setFilter(start, offsetWidth, offsetHeight, field, criteria1, op, criteria2, visibleDropDown) {
        return this.invokeExcelCall("setFilter", {
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
        return this.invokeExcelCall("formatRange", { rangeCode: rangeCode, format: format }, callback);
    }
    clearRange(rangeCode, callback) {
        return this.invokeExcelCall("clearRange", { rangeCode: rangeCode }, callback);
    }
    clearRangeContents(rangeCode, callback) {
        return this.invokeExcelCall("clearRangeContents", { rangeCode: rangeCode }, callback);
    }
    clearRangeFormats(rangeCode, callback) {
        return this.invokeExcelCall("clearRangeFormats", { rangeCode: rangeCode }, callback);
    }
    clearAllCells(callback) {
        return this.invokeExcelCall("clearAllCells", null, callback);
    }
    clearAllCellContents(callback) {
        return this.invokeExcelCall("clearAllCellContents", null, callback);
    }
    clearAllCellFormats(callback) {
        return this.invokeExcelCall("clearAllCellFormats", null, callback);
    }
    setCellName(cellAddress, cellName) {
        return this.invokeExcelCall("setCellName", { address: cellAddress, cellName: cellName });
    }
    calculate() {
        return this.invokeExcelCall("calculateSheet");
    }
    getCellByName(cellName, callback) {
        return this.invokeExcelCall("getCellByName", { cellName: cellName }, callback);
    }
    protect(password) {
        return this.invokeExcelCall("protectSheet", { password: password ? password : null });
    }
    renameSheet(name) {
        return this.invokeExcelCall("renameSheet", { worksheetName: name });
    }
    toObject() {
        return this.objectInstance || (this.objectInstance = {
            addEventListener: this.addEventListener.bind(this),
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
            renameSheet: this.renameSheet.bind(this),
            setCellName: this.setCellName.bind(this),
            setCells: this.setCells.bind(this),
            setFilter: this.setFilter.bind(this)
        });
    }
}
exports.ExcelWorksheet = ExcelWorksheet;
//# sourceMappingURL=ExcelWorksheet.js.map

/***/ }),
/* 9 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

Object.defineProperty(exports, "__esModule", { value: true });
// This is the entry point of the Plugin script
const ExcelApi_1 = __webpack_require__(3);
window.fin.desktop.ExcelService = ExcelApi_1.ExcelService.instance;
Object.defineProperty(window.fin.desktop, 'Excel', {
    get() { return ExcelApi_1.ExcelService.instance.defaultApplicationObj; }
});
//# sourceMappingURL=index.js.map

/***/ })
/******/ ]);