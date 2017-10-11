import { RpcDispatcher } from './RpcDispatcher';
import { ExcelWorkbook } from './ExcelWorkbook';
export declare class ExcelApplication extends RpcDispatcher {
    workbooks: {
        [workbookName: string]: ExcelWorkbook;
    };
    initialized: boolean;
    connected: boolean;
    constructor(connectionUuid: string);
    init(): void;
    processExcelEvent: (data: any, uuid: string) => void;
    processExcelResult: (result: any) => void;
    monitorDisconnect(): void;
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
