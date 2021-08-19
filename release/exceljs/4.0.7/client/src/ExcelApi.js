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
const RpcDispatcher_1 = require("./RpcDispatcher");
const ExcelApplication_1 = require("./ExcelApplication");
const ExcelRtd_1 = require("./ExcelRtd");
const DefaultLogger_1 = require("./DefaultLogger");
const NoOpLogger_1 = require("./NoOpLogger");
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
                    this.logger.debug(this.loggerName + ": Setting service provider version by requesting it from channel.");
                    this.version = yield providerChannel.dispatch('getVersion');
                    this.logger.debug(this.loggerName + `: Service provider version set to: ${JSON.stringify(this.version)}.`);
                }
                catch (err) {
                    let errorMessage;
                    if (err !== undefined && err.message !== undefined) {
                        errorMessage = "Error: " + err.message;
                    }
                    this.logger.warn(this.loggerName + ": Error connecting or fetching version to/from provider. The version of the provider is likely older than the script version.", errorMessage);
                }
                yield this.registerWindowInstance();
                yield this.getExcelInstances();
                this.initialized = true;
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