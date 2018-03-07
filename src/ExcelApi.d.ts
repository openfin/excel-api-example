import { RpcDispatcher } from './RpcDispatcher';
import { ExcelApplication } from './ExcelApplication';
export declare class ExcelService extends RpcDispatcher {
    static instance: ExcelService;
    static defaultApplication: ExcelApplication;
    initialized: boolean;
    applications: {
        [connectionUuid: string]: ExcelApplication;
    };
    constructor();
    init(): Promise<void>;
    processExcelServiceEvent: (data: any) => Promise<void>;
    processExcelServiceResult: (data: any) => Promise<void>;
    subscribeToServiceMessages(): Promise<[void, void]>;
    monitorDisconnect(): Promise<{}>;
    registerAppInstance: (callback?: Function) => void;
    processExcelConnectedEvent(data: any): Promise<void>;
    processExcelDisconnectedEvent(data: any): Promise<void>;
    processGetExcelInstancesResult(connectionUuids: string[]): Promise<void>;
    install(callback: Function): void;
    getInstallationStatus(callback?: Function): void;
    getExcelInstances(callback?: Function): void;
    run(callback?: Function): void;
    toObject(): any;
}
