import { RpcDispatcher } from './RpcDispatcher';
import { ExcelApplication } from './ExcelApplication';
import { ExcelRtd2 as ExcelRtd } from './ExcelRtd';
export declare class ExcelService extends RpcDispatcher {
    static instance: ExcelService;
    defaultApplicationUuid: string;
    defaultApplicationObj: any;
    initialized: boolean;
    applications: {
        [connectionUuid: string]: ExcelApplication;
    };
    constructor();
    init(): Promise<void>;
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
