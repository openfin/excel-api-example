import { RpcDispatcher } from './RpcDispatcher';
import { ExcelWorkbook } from './ExcelWorkbook';
import { ExcelWorksheet } from './ExcelWorksheet';
export declare class ExcelApplication extends RpcDispatcher {
    workbooks: {
        [workbookName: string]: ExcelWorkbook;
    };
    worksheets: {
        [workbookName: string]: {
            [worksheetName: string]: ExcelWorksheet;
        };
    };
    initialized: boolean;
    connected: boolean;
    constructor(connectionUuid: string);
    init(): void;
    processExcelEvent: (data: any, uuid: string) => void;
    processExcelResult: (data: any) => void;
    monitorDisconnect(): void;
    run(callback: Function): void;
    getWorkbooks(callback: Function): void;
    getWorkbookByName(name: string): ExcelWorkbook;
    getWorksheetByName(workbookName: string, worksheetName: string): ExcelWorksheet;
    addWorkbook(callback: Function): void;
    openWorkbook(path: string, callback: Function): void;
    getConnectionStatus(callback: Function): void;
    getCalculationMode(callback: Function): void;
    calculateAll(callback: Function): void;
    toObject(): any;
}
