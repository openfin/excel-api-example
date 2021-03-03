import { RpcDispatcher } from './RpcDispatcher';
import { ExcelApplication } from './ExcelApplication';
import { ExcelRtd } from './ExcelRtd';
import { ILog } from './ILog';
export declare class ExcelService extends RpcDispatcher {
    static instance: ExcelService;
    defaultApplicationUuid: string;
    defaultApplicationObj: any;
    logger: ILog;
    loggerName: string;
    initialized: boolean;
    applications: {
        [connectionUuid: string]: ExcelApplication;
    };
    version: {
        buildVersion: string;
        providerVersion: string;
    };
    constructor();
    init(logger: ILog | boolean): Promise<void>;
    processExcelServiceEvent: (data: any) => Promise<void>;
    processExcelServiceResult: (result: any) => Promise<void>;
    subscribeToServiceMessages(): Promise<[void, void]>;
    monitorDisconnect(): Promise<unknown>;
    registerWindowInstance: (callback?: Function) => Promise<any>;
    configureDefaultApplication(): Promise<void>;
    processExcelConnectedEvent(data: any): Promise<void>;
    processExcelDisconnectedEvent(data: any): Promise<void>;
    processGetExcelInstancesResult(connectionUuids: string[]): Promise<void>;
    install(callback?: Function): Promise<any>;
    getInstallationStatus(callback?: Function): Promise<any>;
    getExcelInstances(callback?: Function): Promise<any>;
    createRtd(providerName: string): Promise<ExcelRtd>;
    toObject(): any;
}
