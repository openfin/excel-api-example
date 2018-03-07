import { RpcDispatcher } from './RpcDispatcher';
import { ExcelWorkbook } from './ExcelWorkbook';
export declare class ExcelApplication extends RpcDispatcher {
    static defaultInstance: ExcelApplication;
    workbooks: {
        [workbookName: string]: ExcelWorkbook;
    };
    connected: boolean;
    initialized: boolean;
    constructor(connectionUuid: string);
    init(): Promise<void>;
    release(): Promise<void>;
    processExcelEvent: (data: any, uuid: string) => void;
    processExcelResult: (result: any) => void;
    subscribeToExcelMessages(): Promise<[void, void]>;
    unsubscribeToExcelMessages(): Promise<[void, void]>;
    monitorDisconnect(): Promise<{}>;
    run(callback: Function): void;
    getWorkbooks(callback: Function): void;
    getWorkbookByName(name: string): ExcelWorkbook;
    addWorkbook(callback: Function): void;
    openWorkbook(path: string, callback: Function): void;
    getConnectionStatus(callback: Function): void;
    getCalculationMode(callback: Function): void;
    calculateAll(callback: Function): void;
    toObject(): any;
}
