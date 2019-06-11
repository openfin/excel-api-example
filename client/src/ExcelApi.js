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
const RpcDispatcher_1 = require("./RpcDispatcher");
const ExcelApplication_1 = require("./ExcelApplication");
const excelServiceUuid = "886834D1-4651-4872-996C-7B2578E953B9";
class ExcelService extends RpcDispatcher_1.RpcDispatcher {
    constructor() {
        super();
        this.defaultApplicationUuid = undefined;
        this.defaultApplicationObj = undefined;
        this.applications = {};
        this.processExcelServiceEvent = (data) => __awaiter(this, void 0, void 0, function* () {
            var eventType = data.event;
            var eventData;
            switch (data.event) {
                case "started":
                    break;
                case "registrationRollCall":
                    if (this.initialized) {
                        this.registerWindowInstance();
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
            if (result.error) {
                executor.reject(result.error);
                return;
            }
            // Internal processing
            switch (result.action) {
                case "getExcelInstances":
                    yield this.processGetExcelInstancesResult(result.data);
                    break;
            }
            executor.resolve(result.data);
        });
        this.registerWindowInstance = (callback) => {
            return this.invokeServiceCall("registerOpenfinWindow", { domain: document.domain }, callback);
        };
        this.connectionUuid = excelServiceUuid;
    }
    init() {
        return __awaiter(this, void 0, void 0, function* () {
            if (!this.initialized) {
                yield this.subscribeToServiceMessages();
                yield this.monitorDisconnect();
                yield fin.desktop.InterApplicationBus.Channel.connect(excelServiceUuid);
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
            var defaultAppObjUuid = this.defaultApplicationObj && this.defaultApplicationObj.connectionUuid;
            var defaultAppEntry = this.applications[defaultAppObjUuid];
            var defaultAppObjConnected = defaultAppEntry ? defaultAppEntry.connected : false;
            if (defaultAppObjConnected) {
                return;
            }
            var connectedAppUuid = Object.keys(this.applications).find(appUuid => this.applications[appUuid].connected);
            if (connectedAppUuid) {
                delete this.applications[defaultAppObjUuid];
                this.defaultApplicationObj = this.applications[connectedAppUuid].toObject();
                return;
            }
            if (defaultAppEntry === undefined) {
                var disconnectedAppUuid = fin.desktop.getUuid();
                var disconnectedApp = new ExcelApplication_1.ExcelApplication(disconnectedAppUuid);
                yield disconnectedApp.init();
                this.applications[disconnectedAppUuid] = disconnectedApp;
                this.defaultApplicationObj = disconnectedApp.toObject();
            }
        });
    }
    // Internal Event Handlers
    processExcelConnectedEvent(data) {
        return __awaiter(this, void 0, void 0, function* () {
            var applicationInstance = this.applications[data.uuid] || new ExcelApplication_1.ExcelApplication(data.uuid);
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
    toObject() {
        return {};
    }
}
ExcelService.instance = new ExcelService();
exports.ExcelService = ExcelService;
//# sourceMappingURL=ExcelApi.js.map