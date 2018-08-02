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

Object.defineProperty(exports, "__esModule", { value: true });
/**
 * @abstract
 * @class Top level class that communicates with the Excel application
 */
class RpcDispatcher {
    constructor() {
        /**
         * @private
         * @description Holds event listeners
         */
        this.listeners = {};
    }
    /**
     * @public
     * @function addEventListener Adds event listener to listen to events coming from Excel application
     * @param type The type of the event to listen to
     * @param listener The method to execute when the event has been fired
     */
    addEventListener(type, listener) {
        if (this.hasEventListener(type, listener)) {
            return;
        }
        if (!this.listeners[type]) {
            this.listeners[type] = [];
        }
        this.listeners[type].push(listener);
    }
    /**
     * @public
     * @function removeEventListener Removes the event from the local store
     * @param type The type of the event to listen to
     * @param listener The method to execute when the event has been fired
     */
    removeEventListener(type, listener) {
        if (!this.hasEventListener(type, listener)) {
            return;
        }
        var callbacksOfType = this.listeners[type];
        callbacksOfType.splice(callbacksOfType.indexOf(listener), 1);
    }
    /**
     * @private
     * @function hasEventListener Check whether an event listener has been registered
     * @param type The type of the event
     * @param listener The method to execute when the event has been fired
     */
    hasEventListener(type, listener) {
        if (!this.listeners[type]) {
            return false;
        }
        if (!listener) {
            return true;
        }
        return (this.listeners[type].indexOf(listener) >= 0);
    }
    /**
     * @public
     * @function dispatchEvent Sends event over to the correct entity e.g. Workbook, worksheet
     * @param evtOrTypeArg Pass either an event or event type as a string
     * @param data The data to be passed to the receiving entity
     */
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
    /**
     * @private
     * @function getDefaultMessage Get the default message when invoking a remote call
     * @returns {object} Returns an empty object to be populated
     */
    getDefaultMessage() {
        return {};
    }
    /**
     * @protected
     * @function invokeExcelCall Invokes a call in excel application via RPC
     * @param functionName The name of the function to invoke
     * @param data Any data to be sent over as part of the invocation
     */
    invokeExcelCall(functionName, data) {
        return this.invokeRemoteCall('excelCall', functionName, data);
    }
    /**
     * @protected
     * @function invokeServiceCall Invokes a call in the excel service process via RPC
     * @param functionName The name of the function to invoke
     * @param data Any data to be sent over as part of the invocation
     */
    invokeServiceCall(functionName, data) {
        return this.invokeRemoteCall('excelServiceCall', functionName, data);
    }
    /**
     * @private
     * @function invokeRemoteCall Invokes a remote procedure call
     * @param topic Topic to send on
     * @param functionName The name of the function to invoke
     * @param data The data to be sent over as part of the invocation
     * @param callback Callback to be applied to the promise
     */
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
            fin.desktop.InterApplicationBus.send(this.connectionUuid, topic, message, ack => {
                RpcDispatcher.promiseExecutors[currentMessageId] = executor;
            }, nak => {
                executor.reject(new Error(nak));
            });
        }
        else {
            executor.reject(new Error('The target UUID of the remote call is undefined.'));
        }
        return promise;
    }
    /**
     * @protected
     * @function applyCallbackToPromise Applies a callback to the promise
     * @param promise The promise to be acted on
     * @param callback THe callback to be applied to the promise
     * @returns {Promise<any>} A promise with the callback applied
     */
    applyCallbackToPromise(promise, callback) {
        if (callback) {
            promise
                .then(result => {
                callback(result);
                return result;
            }).catch(err => {
                console.error(err);
            });
            promise = undefined;
        }
        return promise;
    }
}
/**
 * @protected
 * @static
 * @description The message id of the action being sent
 */
RpcDispatcher.messageId = 1;
/**
 * @protected
 * @static
 * @description Promises to be executed
 */
RpcDispatcher.promiseExecutors = {};
exports.RpcDispatcher = RpcDispatcher;
//# sourceMappingURL=RpcDispatcher.js.map

/***/ }),
/* 1 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
Object.defineProperty(exports, "__esModule", { value: true });
const RpcDispatcher_1 = __webpack_require__(0);
const ExcelApplication_1 = __webpack_require__(2);
/**
 * @constant {string} excelServiceUuid Uuid for the excel service
 */
const excelServiceUuid = "886834D1-4651-4872-996C-7B2578E953B9";
/**
 * @class Class for interacting with the .NET ExcelService process
 */
class ExcelService extends RpcDispatcher_1.RpcDispatcher {
    /**
     * @constructor Constructor for ExcelService
     */
    constructor() {
        super();
        this.connectionUuid = excelServiceUuid;
        this.mInitialized = false;
        this.mApplications = {};
        this.mDefaultApplicationUuid = undefined;
        this.defaultApplicationObj = undefined;
        this.getInitialized();
    }
    /**
     * @public
     * @function init Initialises the ExcelService
     * @returns {Promise<void>} A promise
     */
    init() {
        return __awaiter(this, void 0, void 0, function* () {
            if (!this.mInitialized) {
                yield this.subscribeToServiceMessages();
                yield this.monitorDisconnect();
                //await fin.desktop.Service.connect({ uuid: excelServiceUuid })
                yield this.registerWindowInstance();
                yield this.getExcelInstances();
                this.mInitialized = true;
            }
            return;
        });
    }
    /**
     * @private
     * @function processExcelServiceEvent Processes events coming from the Excel application
     * @param {any} data Payload passed from the Excel Service
     * @returns {Promise<void>} A promise
     */
    processExcelServiceEvent(data) {
        return __awaiter(this, void 0, void 0, function* () {
            let eventType = data.event;
            let eventData;
            switch (eventType) {
                case "started":
                    break;
                case "registrationRollCall":
                    if (this.mInitialized) {
                        this.registerWindowInstance();
                    }
                    break;
                case "excelConnected":
                    yield this.processExcelConnectedEvent(data);
                    eventData = { connectionUuid: data.uuid };
                    break;
                case "excelDisconnected":
                    yield this.processExcelDisconnectedEvent(data).catch(console.error);
                    eventData = { connectionUuid: data.uuid };
                    break;
            }
            this.dispatchEvent(eventType, eventData);
        });
    }
    /**
     * @private
     * @function processExcelServiceResult Processes results from excel service
     * @param {any} result The result from the service
     * @returns {Promise<void>} A promise
     */
    processExcelServiceResult(result) {
        return __awaiter(this, void 0, void 0, function* () {
            var executor = RpcDispatcher_1.RpcDispatcher.promiseExecutors[result.messageId];
            if (!executor) {
                console.warn("No executors matching the messageId: " + result.messageId, result);
                return;
            }
            delete RpcDispatcher_1.RpcDispatcher.promiseExecutors[result.messageId];
            if (result.error) {
                executor.reject(result.error);
                return;
            }
            // Internal processing
            switch (result.action) {
                case "getExcelInstances":
                    yield this.processGetExcelInstancesResult(result.data);
                    break;
                case "getInitialized":
                    this.mInitialized = result.data;
                    break;
                default:
                    break;
            }
            executor.resolve(result.data);
        });
    }
    /**
     * @private
     * @function subscribeToServiceMessages function to subscribe to topics ExcelService will send to
     * @returns {Promise<[void, void]>} A list of promises
     */
    subscribeToServiceMessages() {
        return Promise.all([
            new Promise(resolve => fin.desktop.InterApplicationBus.subscribe(excelServiceUuid, "excelServiceEvent", this.processExcelServiceEvent.bind(this), resolve)),
            new Promise(resolve => fin.desktop.InterApplicationBus.subscribe(excelServiceUuid, "excelServiceCallResult", this.processExcelServiceResult.bind(this), resolve))
        ]);
    }
    /**
     * @private
     * @function monitorDisconnect Subscribes to the disconnected event and dispatches to the excel application
     * @returns {Promnise<void>} A promise
     */
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
    /**
     * @private
     * @function registerWindowInstance This registers a new Excel instance to a new workbook domain
     * @returns {Promise<void>} A promise
     */
    registerWindowInstance() {
        return __awaiter(this, void 0, void 0, function* () {
            return this.invokeServiceCall("registerOpenfinWindow", { domain: document.domain });
        });
    }
    /**
     * @private
     * @function configureDefaultApplication Configures the default application when the application first starts
     * @returns {Promise<void>} A promise
     */
    configureDefaultApplication() {
        return __awaiter(this, void 0, void 0, function* () {
            var defaultAppObjUuid = this.defaultApplicationObj && this.defaultApplicationObj.connectionUuid;
            var defaultAppEntry = this.mApplications[defaultAppObjUuid];
            var defaultAppObjConnected = defaultAppEntry ? defaultAppEntry.connected : false;
            if (defaultAppObjConnected) {
                return;
            }
            var connectedAppUuid = Object.keys(this.mApplications).find(appUuid => this.mApplications[appUuid].connected);
            if (connectedAppUuid) {
                delete this.mApplications[defaultAppObjUuid];
                this.defaultApplicationObj = this.mApplications[connectedAppUuid].toObject();
                return;
            }
            if (defaultAppEntry === undefined) {
                var disconnectedAppUuid = fin.desktop.getUuid();
                var disconnectedApp = new ExcelApplication_1.ExcelApplication(disconnectedAppUuid);
                yield disconnectedApp.init();
                this.mApplications[disconnectedAppUuid] = disconnectedApp;
                this.defaultApplicationObj = disconnectedApp.toObject();
            }
            return;
        });
    }
    // Internal Event Handlers
    /**
     * @private
     * @function processExcelConnectedEvent Process the connected event
     * @param {any} data payload that holds uuid of the connected application
     * @returns {Promise<void>} A promise
     */
    processExcelConnectedEvent(data) {
        return __awaiter(this, void 0, void 0, function* () {
            var applicationInstance = this.mApplications[data.uuid] || new ExcelApplication_1.ExcelApplication(data.uuid);
            yield applicationInstance.init();
            this.mApplications[data.uuid] = applicationInstance;
            // Synthetically raise connected event
            applicationInstance.processExcelEvent({ event: "connected" }, data.uuid);
            yield this.configureDefaultApplication();
            return;
        });
    }
    /**
     * @public
     * @function processExcelDisconnectedEvent Processes event when excel is disconnected
     * @param data The data from excel
     * @returns {Promise<void>} A promise
     */
    processExcelDisconnectedEvent(data) {
        return __awaiter(this, void 0, void 0, function* () {
            var applicationInstance = this.mApplications[data.uuid];
            if (applicationInstance === undefined) {
                return;
            }
            delete this.mApplications[data.uuid];
            console.log('configuring default application in disconnect event');
            this.configureDefaultApplication().then(applicationInstance.release).catch(console.error);
            return;
        });
    }
    // Internal API Handlers
    /**
     * @private
     * @function processGetExcelInstancesResult Get Excel instance
     * @param {string[]} connectionUuids THe connection Uuids the Excel service is holding
     * @returns {Promise<void>} A promise
     */
    processGetExcelInstancesResult(connectionUuids) {
        return __awaiter(this, void 0, void 0, function* () {
            var existingInstances = this.mApplications;
            this.mApplications = {};
            yield Promise.all(connectionUuids.map((connectionUuid) => __awaiter(this, void 0, void 0, function* () {
                var existingInstance = existingInstances[connectionUuid];
                if (existingInstance === undefined) {
                    // Assume that since the ExcelService is aware of the instance it,
                    // is connected and to simulate the the connection event
                    yield this.processExcelServiceEvent({ event: "excelConnected", uuid: connectionUuid });
                }
                else {
                    this.mApplications[connectionUuid] = existingInstance;
                }
                return;
            })));
            yield this.configureDefaultApplication();
        });
    }
    // API Calls
    /**
     * @public
     * @function install Installs the addin
     * @returns {Promise<any>} A promise
     */
    install() {
        return this.invokeServiceCall("install", null);
    }
    /**
     * @public
     * @function getInstallationStatus Checks the installation status
     * @returns {Promise<any>} A promise
     */
    getInstallationStatus() {
        return this.invokeServiceCall("getInstallationStatus", null);
    }
    /**
     * @public
     * @function getExcelInstances Returns all the excel instances that are open
     * @returns {Promise<any>} A promsie
     */
    getExcelInstances() {
        return this.invokeServiceCall("getExcelInstances", null);
    }
    /**
     * @public
     * @function getInitialized Returns whether or not the service has been initialised or not
     * @returns {Promise<any>} A promise
     */
    getInitialized() {
        return this.invokeServiceCall("getInitialized", null);
    }
    /**
     * @public
     * @function toObject Creates an empty object
     * @returns {object} An empty object
     */
    toObject() {
        return {};
    }
}
exports.ExcelService = ExcelService;
//# sourceMappingURL=ExcelApi.js.map

/***/ }),
/* 2 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
Object.defineProperty(exports, "__esModule", { value: true });
const RpcDispatcher_1 = __webpack_require__(0);
const ExcelWorkbook_1 = __webpack_require__(3);
const ExcelWorksheet_1 = __webpack_require__(4);
/**
 * @class Represents the Excel application itself
 */
class ExcelApplication extends RpcDispatcher_1.RpcDispatcher {
    /**
     * @constructor Constructor for the class
     * @param connectionUuid The connection uuid of the openfin application
     */
    constructor(connectionUuid) {
        super();
        /**
         * @private
         * @description A key value pair container that holds name of the workbook as key
         * and the workbook object itself as the value
         */
        this.workbooks = {};
        this.connectionUuid = connectionUuid;
        this.mConnected = false;
    }
    /**
     * @public
     * @property Flag to indicate whether excel is connected to openfin
     * @returns {boolean} Connected or not
     */
    get connected() {
        return this.mConnected;
    }
    /**
     * @public
     * @function init Initialises the application
     * @returns {Promise<void>} A promise
     */
    init() {
        return __awaiter(this, void 0, void 0, function* () {
            if (!this.initialized) {
                yield this.subscribeToExcelMessages();
                yield this.monitorDisconnect();
                this.initialized = true;
            }
            return;
        });
    }
    /**
     * @public
     * @function release Release all connection from the excel application to the openfin app
     * @returns {Promise<void>} A promise
     */
    release() {
        return __awaiter(this, void 0, void 0, function* () {
            if (this.initialized) {
                yield this.unsubscribeToExcelMessages();
                //TODO: Provide external means to stop monitoring disconnect
            }
            return;
        });
    }
    /**
     * @private
     * @function processExcelEvent Process events coming from excel to be handled by the openfin app
     * @param data The data being sent over from the excel app
     * @param uuid The uuid of the sender
     */
    processExcelEvent(data, uuid) {
        var eventType = data.event;
        var workbook = this.workbooks[data.workbookName];
        var worksheets = workbook && workbook.worksheets;
        var worksheet = worksheets && worksheets[data.sheetName];
        switch (eventType) {
            case "connected":
                this.mConnected = true;
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
            case "rowDeleted":
                console.log(data);
                if (!worksheet) {
                    console.error('No worksheet could be found');
                    return;
                }
                worksheet.dispatchEvent(eventType, { data: data });
                break;
            case "rowInserted":
                console.log(data);
                if (!worksheet) {
                    console.error('No worksheet could be found');
                    return;
                }
                worksheet.dispatchEvent(eventType, { data: data });
            case "afterCalculation":
            default:
                this.dispatchEvent(eventType);
                break;
        }
    }
    /**
     * @private
     * @function processExcelResult Process results coming from excel application
     * @param result The result of the call being made in the excel application
     */
    processExcelResult(result) {
        var callbackData = {};
        var executor = RpcDispatcher_1.RpcDispatcher.promiseExecutors[result.messageId];
        delete RpcDispatcher_1.RpcDispatcher.promiseExecutors[result.messageId];
        if (result.error) {
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
        executor.resolve(callbackData);
    }
    /**
     * @private
     * @function subscribeToExelMessages Subscribes to messages from Excel application
     * @returns {Promise<[void, void]>} A promise
     */
    subscribeToExcelMessages() {
        return Promise.all([
            new Promise(resolve => fin.desktop.InterApplicationBus.subscribe(this.connectionUuid, "excelEvent", this.processExcelEvent.bind(this), resolve)),
            new Promise(resolve => fin.desktop.InterApplicationBus.subscribe(this.connectionUuid, "excelResult", this.processExcelResult.bind(this), resolve))
        ]);
    }
    /**
     * @private
     * @function unsubscribeToExcelMessages Unsubscribes from Excel application
     * @returns {Promise<[void, void]>} A promise
     */
    unsubscribeToExcelMessages() {
        return Promise.all([
            new Promise(resolve => fin.desktop.InterApplicationBus.unsubscribe(this.connectionUuid, "excelEvent", this.processExcelEvent.bind(this), resolve)),
            new Promise(resolve => fin.desktop.InterApplicationBus.unsubscribe(this.connectionUuid, "excelResult", this.processExcelResult.bind(this), resolve))
        ]);
    }
    /**
     * @private
     * @function monitorDisconnect Monitors disconnection event when openfin disconnects from excel
     * @returns {Promise<void>} A promise
     */
    monitorDisconnect() {
        return new Promise((resolve, reject) => {
            var excelApplicationConnection = fin.desktop.ExternalApplication.wrap(this.connectionUuid);
            var onDisconnect;
            excelApplicationConnection.addEventListener('disconnected', onDisconnect = () => {
                excelApplicationConnection.removeEventListener('disconnected', onDisconnect);
                this.mConnected = false;
                this.dispatchEvent('disconnected');
            }, resolve, reject);
        });
    }
    /**
     * @public
     * @function run Runs Excel application
     * @param callback The callback to be applied
     */
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
    /**
     * @public
     * @function getWorkbooks Gets the workbooks within the excel application
     * @returns {Promise<any>} A promise
     */
    getWorkbooks() {
        return this.invokeExcelCall("getWorkbooks", null);
    }
    /**
     * @public
     * @function getWorkbookByName Gets the registered workbook with the specified name
     * @param name The name of the workbook
     */
    getWorkbookByName(name) {
        return this.workbooks[name];
    }
    /**
     * @function addWorkbook adds a workbook to the Excel application
     * @returns {Promise<any>} A promise with a result
     */
    addWorkbook() {
        return this.invokeExcelCall("addWorkbook", null);
    }
    /**
     * @public
     * @function openWorkbook Opens the workbook specified at the path
     * @param path The path of the workbook
     * @returns {Promise<any>} Returns a promise with a result
     */
    openWorkbook(path) {
        return this.invokeExcelCall("openWorkbook", { path: path });
    }
    /**
     * @public
     * @function getConnectionStatus Gets the connection status of of the Excel application
     * @returns {Promise<any>} A promise with a result
     */
    getConnectionStatus(callback) {
        return this.applyCallbackToPromise(Promise.resolve(this.connected), callback);
    }
    /**
     * @public
     * @function getCalculationMode Gets the calculation mode from Excel application
     * @returns {Promise<any>} A promise with a result
     */
    getCalculationMode() {
        return this.invokeExcelCall("getCalculationMode", null);
    }
    /**
     * @public
     * @function calculateAll Calculates all formulas on the workbook
     * @returns {Promise<any>} A promise with a result
     */
    calculateAll() {
        return this.invokeExcelCall("calculateFull", null);
    }
    /**
     * @public
     * @function toObject Returns an object with only the methods and properties to be exposed
     * @returns {any} An object with only the methods and properties to be exposed
     */
    toObject() {
        return this.objectInstance || (this.objectInstance = {
            connectionUuid: this.connectionUuid,
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
/**
 * @public
 * @static
 * @description The default excel application instance
 */
ExcelApplication.defaultInstance = undefined;
exports.ExcelApplication = ExcelApplication;
//# sourceMappingURL=ExcelApplication.js.map

/***/ }),
/* 3 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

Object.defineProperty(exports, "__esModule", { value: true });
const RpcDispatcher_1 = __webpack_require__(0);
/**
 * @class Class that represents a workbook
 */
class ExcelWorkbook extends RpcDispatcher_1.RpcDispatcher {
    /**
     * @constructor Constructor for the ExcelWorkbook class
     * @param application The Application this workbook belongs to
     * @param name The name of the workbook
     */
    constructor(application, name) {
        super();
        this.connectionUuid = application.connectionUuid;
        this.application = application;
        this.mWorksheets = {};
        this.mWorkbookName = name;
    }
    /**
     * @private
     * @function getDefaultMessage Gets the default message to be sent over the wire
     * @returns {any} An object with the workbook name in as default
     */
    getDefaultMessage() {
        return {
            workbook: this.mWorkbookName
        };
    }
    /**
     * @public
     * @property Worksheets tied to this workbook
     * @returns {{ [worksheetName: string]: ExcelWorksheet }}
     */
    get worksheets() {
        return this.mWorksheets;
    }
    set worksheets(worksheets) {
        this.mWorksheets = worksheets;
    }
    /**
     * @public
     * @property workbookName property
     * @returns {string} The name of the workbook
     */
    get workbookName() {
        return this.mWorkbookName;
    }
    /**
     * @public
     * @property Sets the workbook name
     */
    set workbookName(name) {
        this.mWorkbookName = name;
    }
    /**
     * @public
     * @function getWorksheets Gets the worksheets tied to this workbook
     * @returns A promise with worksheets as the result
     */
    getWorksheets() {
        return this.invokeExcelCall("getWorksheets", null);
    }
    /**
     * @public
     * @function getWorksheetByName Gets the worksheet by name
     * @param name The name of the worksheet
     * @returns {ExcelWorksheet} The excel worksheet with the specified name
     */
    getWorksheetByName(name) {
        return this.worksheets[name];
    }
    /**
     * @public
     * @function addWorksheet Adds a new worksheet to the workbook
     * @returns {Promise<any>} A promise
     */
    addWorksheet() {
        return this.invokeExcelCall("addSheet", null);
    }
    /**
     * @public
     * @function activate Activates the workbook
     * @returns {Promise<any>} A promise
     */
    activate() {
        return this.invokeExcelCall("activateWorkbook");
    }
    /**
     * @public
     * @function save Save the workbook
     * @returns {Promise<void>} A promise
     */
    save() {
        return this.invokeExcelCall("saveWorkbook");
    }
    /**
     * @public
     * @function close Closes the workbook
     * @returns {Promise<void>} A promise
     */
    close() {
        return this.invokeExcelCall("closeWorkbook");
    }
    /**
     * @public
     * @function toObject Returns only the methods exposed
     * @returns {any} Returns only the methods exposed
     */
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
/* 4 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

Object.defineProperty(exports, "__esModule", { value: true });
const RpcDispatcher_1 = __webpack_require__(0);
/**
 * @class Class that represents a worksheet
 */
class ExcelWorksheet extends RpcDispatcher_1.RpcDispatcher {
    /**
     * @constructor Constructor for the ExcelWorksheet class
     * @param name The name of the worksheet
     * @param workbook The ExcelWorkbook this worksheet is tied to
     */
    constructor(name, workbook) {
        super();
        this.connectionUuid = workbook.connectionUuid;
        this.workbook = workbook;
        this.mWorksheetName = name;
    }
    /**
     * @protected
     * @function getDefaultMessage Returns the default message
     * @returns {any} Returns the default message
     */
    getDefaultMessage() {
        return {
            workbook: this.workbook.workbookName,
            worksheet: this.mWorksheetName
        };
    }
    /**
     * @public
     * @property Returns worksheet name
     * @returns {string} The name of the worksheet
     */
    get worksheetName() {
        return this.mWorksheetName;
    }
    /**
     * @public
     * @property Returns worksheet name
     * @returns {string} The name of the worksheet
     */
    set worksheetName(name) {
        this.mWorksheetName = name;
    }
    /**
     * @public
     * @function setCells Sets the content for the cells
     * @param values values for the cell
     * @param offset The cell address
     */
    setCells(values, offset) {
        if (!offset) {
            offset = "A1";
        }
        return this.invokeExcelCall("setCells", { offset: offset, values: values });
    }
    /**
     * @public
     * @function getCells Gets cell values from the range specified
     * @param start The start cell address
     * @param offsetWidth The number of columns in the openfin app
     * @param offsetHeight The number of rows in the openfin app
     */
    getCells(start, offsetWidth, offsetHeight) {
        return this.invokeExcelCall("getCells", { start: start, offsetWidth: offsetWidth, offsetHeight: offsetHeight });
    }
    /**
     * @function activateRow This mirrors the row selected in the openfin application to Excel
     * @param {string} cellAddress THe address of the first cell of the row
     */
    activateRow(cellAddress) {
        return this.invokeExcelCall("activateRow", { address: cellAddress });
    }
    /**
     * @function insertRow This inserts a row just before the selected row
     * @param {number} rowNumber The address of the first cell in the row
     * @returns {Promise<any>} A promise
     */
    insertRow(rowNumber) {
        return this.invokeExcelCall("insertRow", { rowNumber: rowNumber });
    }
    /**
     * @function deleteRow This deletes the selected row
     * @param {number} rowNumber The address of the first cell in the row
     * @returns {Promise<any>} A promise
     */
    deleteRow(rowNumber) {
        return this.invokeExcelCall("deleteRow", { rowNumber: rowNumber });
    }
    /**
     * @public
     * @function activate Activates the current worksheet
     * @returns {Promise<any>} A promise
     */
    activate() {
        return this.invokeExcelCall("activateSheet");
    }
    /**
     * @public
     * @function activateCell Activates the selected cell
     * @param cellAddress The address of the cell
     * @returns {Promise<any>} A promise
     */
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
    /**
     * @public
     * @function formatRange Formats the range selected
     * @param rangeCode The selected range
     * @param format The formatting to be applied to the range
     */
    formatRange(rangeCode, format) {
        return this.invokeExcelCall("formatRange", { rangeCode: rangeCode, format: format });
    }
    /**
     * @public
     * @function clearRange Clear the range of formatting and content
     * @param rangeCode The range selected
     */
    clearRange(rangeCode) {
        return this.invokeExcelCall("clearRange", { rangeCode: rangeCode });
    }
    /**
     * @public
     * @function clearRangeContents Clears the contents in the specified range
     * @param rangeCode The selected range
     */
    clearRangeContents(rangeCode) {
        return this.invokeExcelCall("clearRangeContents", { rangeCode: rangeCode });
    }
    /**
     * @public
     * @function clearRangeFormats Clears the formatting in the range specified
     * @param rangeCode The selected range
     */
    clearRangeFormats(rangeCode) {
        return this.invokeExcelCall("clearRangeFormats", { rangeCode: rangeCode });
    }
    /**
     * @public
     * @function clearAllCells Clear all cells and their formatting
     * @returns {Promise<any>} A promise
     */
    clearAllCells() {
        return this.invokeExcelCall("clearAllCells", null);
    }
    /**
     * @public
     * @function clearAllCellContents Clears all the cells content
     * @returns {Promise<any>} A promise
     */
    clearAllCellContents() {
        return this.invokeExcelCall("clearAllCellContents", null);
    }
    /**
     * @public
     * @function clearAllCellFormats Clear all formatting in every cell
     * @returns {Promise<any>} A promise
     */
    clearAllCellFormats() {
        return this.invokeExcelCall("clearAllCellFormats", null);
    }
    /**
     * @public
     * @function setCellName Sets a name for the cell address
     * @param cellAddress The address of the cell e.g. A1
     * @param cellName The name of the cell
     * @returns {Promise<any>} A promise
     */
    setCellName(cellAddress, cellName) {
        return this.invokeExcelCall("setCellName", { address: cellAddress, cellName: cellName });
    }
    /**
     * @public
     * @function calculate Calculates all formula on teh sheet
     * @returns {Promise<any>} A promise
     */
    calculate() {
        return this.invokeExcelCall("calculateSheet");
    }
    /**
     * @public
     * @function getCellByName Gets a cell by its name
     * @param cellName The name of the cell
     * @returns {Promise<any>} A promise
     */
    getCellByName(cellName) {
        return this.invokeExcelCall("getCellByName", { cellName: cellName });
    }
    /**
     * @public
     * @function protect Password protects the sheet
     * @param password Password used to protect the sheet
     * @returns {Promise<any>} A promise
     */
    protect(password) {
        return this.invokeExcelCall("protectSheet", { password: password ? password : null });
    }
    /**
     * @public
     * @function toObject Returns only the functions that should be exposed by this class
     * @returns {object} Public methods in ExcelWorksheet
     */
    toObject() {
        return this.objectInstance || (this.objectInstance = {
            addEventListener: this.addEventListener.bind(this),
            removeEventListener: this.removeEventListener.bind(this),
            name: this.worksheetName,
            activate: this.activate.bind(this),
            activateCell: this.activateCell.bind(this),
            activateRow: this.activateRow.bind(this),
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
            protect: this.protect.bind(this),
            setCellName: this.setCellName.bind(this),
            setCells: this.setCells.bind(this),
            setFilter: this.setFilter.bind(this),
            insertRow: this.insertRow.bind(this),
            deleteRow: this.deleteRow.bind(this)
        });
    }
}
exports.ExcelWorksheet = ExcelWorksheet;
//# sourceMappingURL=ExcelWorksheet.js.map

/***/ }),
/* 5 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

Object.defineProperty(exports, "__esModule", { value: true });
// This is the entry point of the Plugin script
const ExcelApi_1 = __webpack_require__(1);
const excelService = new ExcelApi_1.ExcelService();
// Attach ExcelService to the window
window.fin.desktop.ExcelService = excelService;
// Attach the Excel api to the window
Object.defineProperty(window.fin.desktop, 'Excel', {
    get() { return excelService.defaultApplicationObj; }
});
fin.desktop.main(() => {
    function init(message) {
        console.log(message);
        excelService.init()
            .then(() => { fin.desktop.InterApplicationBus.unsubscribe("886834D1-4651-4872-996C-7B2578E953B9", "init", init); })
            .catch((err) => { console.log("This error might be ok", err); });
    }
    fin.desktop.InterApplicationBus.subscribe("886834D1-4651-4872-996C-7B2578E953B9", "init", init);
    fin.desktop.InterApplicationBus.send("886834D1-4651-4872-996C-7B2578E953B9", "init-multi-window", "initial fire");
});
//# sourceMappingURL=plugin.js.map

/***/ })
/******/ ]);