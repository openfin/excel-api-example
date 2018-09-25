"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const RpcDispatcher_1 = require("./RpcDispatcher");
/**
 * @class Class that represents a worksheet
 */
class ExcelWorksheet extends RpcDispatcher_1.RpcDispatcher {
    /**
     * @constructor Constructor for the ExcelWorksheet class
     * @param name The name of the worksheet
     * @param workbook The ExcelWorkbook this worksheet is tied to
     */
    constructor(name, workbook) {
        super();
        this.connectionUuid = workbook.connectionUuid;
        this.mWorkbook = workbook;
        this.mWorksheetName = name;
    }
    /**
     * @protected
     * @function getDefaultMessage Returns the default message
     * @returns {object} Returns the default message
     */
    getDefaultMessage() {
        return { workbook: this.workbook.name, worksheet: this.mWorksheetName };
    }
    /**
     * @public
     * @property Returns worksheet name
     * @returns {string} The name of the worksheet
     */
    get name() {
        return this.mWorksheetName;
    }
    /**
     * @public
     * @property Returns worksheet name
     * @returns {string} The name of the worksheet
     */
    set name(name) {
        this.mWorksheetName = name;
    }
    get workbook() {
        return this.mWorkbook;
    }
    /**
     * @public
     * @function setCells Sets the content for the cells
     * @param values values for the cell
     * @param offset The cell address
     */
    setCells(values, offset) {
        if (!offset) {
            offset = 'A1';
        }
        const payload = { offset, values };
        return this.invokeExcelCall('setCells', payload);
    }
    /**
     * @public
     * @function getCells Gets cell values from the range specified
     * @param start The start cell address
     * @param offsetWidth The number of columns in the openfin app
     * @param offsetHeight The number of rows in the openfin app
     */
    getCells(start, offsetWidth, offsetHeight) {
        const payload = { start, offsetHeight, offsetWidth };
        return this.invokeExcelCall('getCells', payload);
    }
    /**
     * @function activateRow This mirrors the row selected in the openfin
     * application to Excel
     * @param {string} cellAddress THe address of the first cell of the row
     */
    activateRow(cellAddress) {
        const payload = { address: cellAddress };
        return this.invokeExcelCall('activateRow', payload);
    }
    /**
     * @function insertRow This inserts a row just before the selected row
     * @param {number} rowNumber The address of the first cell in the row
     * @returns {Promise<void>} A promise
     */
    insertRow(rowNumber) {
        return this.invokeExcelCall('insertRow', { rowNumber });
    }
    /**
     * @function deleteRow This deletes the selected row
     * @param {number} rowNumber The address of the first cell in the row
     * @returns {Promise<any>} A promise
     */
    deleteRow(rowNumber) {
        return this.invokeExcelCall('deleteRow', { rowNumber });
    }
    /**
     * @public
     * @function activate Activates the current worksheet
     * @returns {Promise<any>} A promise
     */
    activate() {
        return this.invokeExcelCall('activateSheet');
    }
    /**
     * @public
     * @function activateCell Activates the selected cell
     * @param cellAddress The address of the cell
     * @returns {Promise<void>} A promise
     */
    activateCell(cellAddress) {
        return this.invokeExcelCall('activateCell', { address: cellAddress });
    }
    // public addButton(name: string, caption: string, cellAddress: string):
    // Promise<void> {
    //    return this.invokeExcelCall("addButton", { address: cellAddress,
    //    buttonName: name, buttonCaption: caption });
    //}
    // public setFilter(start: string, offsetWidth: number, offsetHeight: number,
    // field: number, criteria1: string, op: string, criteria2: string,
    // visibleDropDown: string): Promise<void> {
    //    return this.invokeExcelCall("setFilter", {
    //        start,
    //        offsetWidth,
    //        offsetHeight,
    //        field,
    //        criteria1,
    //        op,
    //        criteria2,
    //        visibleDropDown
    //    });
    //}
    ///**
    // * @public
    // * @function formatRange Formats the range selected
    // * @param {string} rangeCode The selected range
    // * @param {any} format The formatting to be applied to the range
    // * @returns {Promise<void>} A promise
    // */
    // public formatRange(rangeCode: string, format: any): Promise<void> {
    //    return this.invokeExcelCall("formatRange", { rangeCode, format });
    //}
    /**
     * @public
     * @function clearRange Clear the range of formatting and content
     * @param rangeCode The range selected
     */
    clearRange(rangeCode) {
        return this.invokeExcelCall('clearRange', { rangeCode });
    }
    /**
     * @public
     * @function clearRangeContents Clears the contents in the specified range
     * @param rangeCode The selected range
     */
    clearRangeContents(rangeCode) {
        return this.invokeExcelCall('clearRangeContents', { rangeCode });
    }
    ///**
    // * @public
    // * @function clearRangeFormats Clears the formatting in the range specified
    // * @param rangeCode The selected range
    // */
    // public clearRangeFormats(rangeCode: string): Promise<void> {
    //    return this.invokeExcelCall("clearRangeFormats", { rangeCode });
    //}
    /**
     * @public
     * @function clearAllCells Clear all cells and their formatting
     * @returns {Promise<any>} A promise
     */
    clearAllCells() {
        return this.invokeExcelCall('clearAllCells', null);
    }
    /**
     * @public
     * @function clearAllCellContents Clears all the cells content
     * @returns {Promise<void>} A promise
     */
    clearAllCellContents() {
        return this.invokeExcelCall('clearAllCellContents', null);
    }
    ///**
    // * @public
    // * @function clearAllCellFormats Clear all formatting in every cell
    // * @returns {Promise<any>} A promise
    // */
    // public clearAllCellFormats(): Promise<void> {
    //    return this.invokeExcelCall("clearAllCellFormats", null);
    //}
    /**
     * @public
     * @function setCellName Sets a name for the cell address
     * @param cellAddress The address of the cell e.g. A1
     * @param cellName The name of the cell
     * @returns {Promise<any>} A promise
     */
    setCellName(cellAddress, cellName) {
        return this.invokeExcelCall('setCellName', { address: cellAddress, cellName });
    }
    /**
     * @public
     * @function calculate Calculates all formula on teh sheet
     * @returns {Promise<void>} A promise
     */
    calculate() {
        return this.invokeExcelCall('calculateSheet');
    }
    /**
     * @public
     * @function getCellByName Gets a cell by its name
     * @param cellName The name of the cell
     * @returns {Promise<any>} A promise
     */
    getCellByName(cellName) {
        return this.invokeExcelCall('getCellByName', { cellName });
    }
    /**
     * @public
     * @function protect Password protects the sheet
     * @param password Password used to protect the sheet
     * @returns {Promise<any>} A promise
     */
    protect(password) {
        return this.invokeExcelCall('protectSheet', { password: password ? password : null });
    }
    /**
     * @public
     * @function toObject Returns only the functions that should be exposed by
     * this class
     * @returns {object} Public methods in ExcelWorksheet
     */
    toObject() {
        return this.objectInstance || (this.objectInstance = {
            addEventListener: this.addEventListener.bind(this),
            removeEventListener: this.removeEventListener.bind(this),
            name: this.name,
            workbook: this.workbook,
            activate: this.activate.bind(this),
            activateCell: this.activateCell.bind(this),
            activateRow: this.activateRow.bind(this),
            calculate: this.calculate.bind(this),
            clearAllCellContents: this.clearAllCellContents.bind(this),
            clearAllCells: this.clearAllCells.bind(this),
            clearRange: this.clearRange.bind(this),
            clearRangeContents: this.clearRangeContents.bind(this),
            getCellByName: this.getCellByName.bind(this),
            getCells: this.getCells.bind(this),
            protect: this.protect.bind(this),
            setCellName: this.setCellName.bind(this),
            setCells: this.setCells.bind(this),
            insertRow: this.insertRow.bind(this),
            deleteRow: this.deleteRow.bind(this),
            toObject: this.toObject.bind(this)
        });
    }
}
exports.ExcelWorksheet = ExcelWorksheet;
//# sourceMappingURL=ExcelWorksheet.js.map