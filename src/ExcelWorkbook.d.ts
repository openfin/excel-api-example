import { RpcDispatcher } from './RpcDispatcher';
import { Excel } from './ExcelApplication';
import { ExcelWorksheet } from './ExcelWorksheet';
export declare class ExcelWorkbook extends RpcDispatcher {
    application: Excel;
    name: string;
    constructor(application: Excel, name: string);
    getDefaultMessage(): any;
    getWorksheets(callback: Function): void;
    getWorksheetByName(name: string): ExcelWorksheet;
    addWorksheet(callback: Function): void;
    activate(): void;
    save(): void;
    close(): void;
    toObject(): any;
}
