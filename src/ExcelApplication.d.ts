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
    constructor();
    init(): void;
    processExcelEvent: (data: any) => void;
    processExcelResult: (data: any) => void;
    processExcelCustomFunction: (data: any) => void;
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
