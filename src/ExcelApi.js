"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
const RpcDispatcher_1 = require("./RpcDispatcher");
const ExcelApplication_1 = require("./ExcelApplication");
const excelServiceUuid = "886834D1-4651-4872-996C-7B2578E953B9";
const nullApplication = new ExcelApplication_1.ExcelApplication(undefined);
class ExcelService extends RpcDispatcher_1.RpcDispatcher {
    constructor() {
        super();
        this.applications = {};
        this.processExcelServiceEvent = (data) => __awaiter(this, void 0, void 0, function* () {
            console.log('processExcelServiceEvent', data.event);
            var eventType = data.event;
            var eventData;
            switch (data.event) {
                case "started":
                    break;
                case "registrationRollCall":
                    if (this.initialized) {
                        this.registerAppInstance();
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
        this.processExcelServiceResult = (data) => __awaiter(this, void 0, void 0, function* () {
            console.log('processExcelServiceResult', data.action);
            var executor = RpcDispatcher_1.RpcDispatcher.callbacksP[data.messageId];
            delete RpcDispatcher_1.RpcDispatcher.callbacksP[data.messageId];
            if (data.error) {
                executor.reject(data.error);
                return;
            }
            // Internal processing
            switch (data.action) {
                case "getExcelInstances":
                    yield this.processGetExcelInstancesResult(data.result);
                    break;
            }
            executor.resolve(data.result);
        });
        this.registerAppInstance = (callback) => {
            this.invokeServiceCall("registerAppInstance", { domain: document.domain }, callback);
        };
        this.connectionUuid = excelServiceUuid;
    }
    init() {
        return __awaiter(this, void 0, void 0, function* () {
            if (!this.initialized) {
                yield this.subscribeToServiceMessages();
                yield this.monitorDisconnect();
                yield fin.desktop.Service.connect({ uuid: excelServiceUuid });
                //TODO: Change these once API Calls return promises
                yield new Promise(resolve => this.registerAppInstance(resolve));
                yield new Promise(resolve => this.getExcelInstances(resolve));
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
    // Internal Event Handlers
    processExcelConnectedEvent(data) {
        return __awaiter(this, void 0, void 0, function* () {
            var applicationInstance = this.applications[data.uuid] || new ExcelApplication_1.ExcelApplication(data.uuid);
            yield applicationInstance.init();
            this.applications[data.uuid] = applicationInstance;
            // Synthetically raise connected event
            applicationInstance.processExcelEvent({ event: "connected" }, data.uuid);
            if (ExcelService.defaultApplication.connectionUuid === undefined) {
                ExcelService.defaultApplication = applicationInstance.toObject();
            }
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
            if (applicationInstance.connectionUuid === ExcelService.defaultApplication.connectionUuid) {
                var nextDefaultUuid = Object.keys(this.applications).find(() => true);
                ExcelService.defaultApplication = nextDefaultUuid && this.applications[nextDefaultUuid].toObject();
            }
            if (ExcelService.defaultApplication === undefined) {
                ExcelService.defaultApplication = nullApplication;
            }
            yield applicationInstance.release();
        });
    }
    // Internal API Handlers
    processGetExcelInstancesResult(connectionUuids) {
        return __awaiter(this, void 0, void 0, function* () {
            var oldInstances = this.applications;
            this.applications = {};
            yield Promise.all(connectionUuids.map((connectionUuid) => __awaiter(this, void 0, void 0, function* () {
                var applicationInstance = oldInstances[connectionUuid] || new ExcelApplication_1.ExcelApplication(connectionUuid);
                yield applicationInstance.init();
                // Assume since the ExcelService reported the instance
                // that it is currently subscribed and connected
                applicationInstance.connected = true;
                this.applications[connectionUuid] = applicationInstance;
                if (ExcelService.defaultApplication.connectionUuid === undefined) {
                    ExcelService.defaultApplication = applicationInstance.toObject();
                }
                return;
            })));
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
    run(callback) {
        if (ExcelService.defaultApplication.connectionUuid != undefined && callback) {
            callback();
        }
        else {
            var connectedCallback = () => {
                ExcelService.instance.removeEventListener("excelConnected", connectedCallback);
                callback && callback();
            };
            ExcelService.instance.addEventListener("excelConnected", connectedCallback);
            fin.desktop.System.launchExternalProcess({
                target: "excel"
            });
        }
    }
    toObject() {
        return {};
    }
}
ExcelService.instance = new ExcelService();
ExcelService.defaultApplication = nullApplication;
exports.ExcelService = ExcelService;
//# sourceMappingURL=ExcelApi.js.map