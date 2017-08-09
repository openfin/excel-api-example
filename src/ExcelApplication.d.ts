import { RpcDispatcher } from './RpcDispatcher';
import { ExcelWorkbook } from './ExcelWorkbook';
import { ExcelWorksheet } from './ExcelWorksheet';
export declare class Excel extends RpcDispatcher {
    workbooks: {
        [workbookName: string]: ExcelWorkbook;
    };
    worksheets: {
        [workbookName: string]: {
            [worksheetName: string]: ExcelWorksheet;
        };
    };
    connected: boolean;
    constructor();
    init(): void;
    processExcelEvent: (data: any, uuid: string) => void;
    processExcelResult: (data: any) => void;
    processExcelServiceResult: (data: any) => void;
    monitorDisconnect(uuid: string): void;
    run(callback: Function): void;
    install(callback: Function): void;
    getInstallationStatus(callback: Function): void;
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
declare var _default: Excel;
export default _default;
