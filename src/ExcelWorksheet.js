"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const RpcDispatcher_1 = require("./RpcDispatcher");
/**
 * @class
 * @description Class that represents a worksheet
 */
class ExcelWorksheet extends RpcDispatcher_1.RpcDispatcher {
    /**
     * @constructor
     * @description Constructor for the ExcelWorksheet class
     * @param {string} name The name of the worksheet
     * @param {Workbook} workbook The ExcelWorkbook this worksheet is tied to
     */
    constructor(name, workbook) {
        super();
        this.connectionUuid = workbook.connectionUuid;
        this.mWorkbook = workbook;
        this.mWorksheetName = name;
        this.objectInstance = null;
    }
    /**
     * @protected
     * @function getDefaultMessage
     * @description Returns the default message
     * @returns {object} Returns the default message
     */
    getDefaultMessage() {
        return { workbook: this.workbook.name, worksheet: this.mWorksheetName };
    }
    /**
     * @public
     * @property
     * @description Returns worksheet name
     * @returns {string} The name of the worksheet
     */
    get name() {
        return this.mWorksheetName;
    }
    /**
     * @public
     * @property
     * @description Returns worksheet name
     * @param {string} name The name of the worksheet
     */
    set name(name) {
        this.mWorksheetName = name;
    }
    /**
     * @public
     * @property
     * @description Returns the workbook that this worksheet is tied to
     * @returns {Workbook} Returns the workbook that this worksheet is tied to
     */
    get workbook() {
        return this.mWorkbook;
    }
    /**
     * @public
     * @function setCells
     * @description Sets the content for the cells
     * @param {(string|number)[][]} values values for the cell
     * @param {string} offset The cell address
     * @returns {Promise<void>} A promise
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
     * @function getCells
     * @description Gets cell values from the range specified
     * @param {string} start The start cell address
     * @param {number} offsetWidth The number of columns in the openfin app
     * @param {number} offsetHeight The number of rows in the openfin app
     * @returns {Promise<(string | number)[][]>} A promise containing the cells
     */
    getCells(start, offsetWidth, offsetHeight) {
        const payload = { start, offsetHeight, offsetWidth };
        return this.invokeExcelCall('getCells', payload);
    }
    /**
     * @public
     * @function activateRow
     * @description This mirrors the row selected in the openfin
     * application to Excel
     * @param {string} cellAddress THe address of the first cell of the row
     * @returns {Promise<void>} A promise
     */
    activateRow(cellAddress) {
        const payload = { address: cellAddress };
        return this.invokeExcelCall('activateRow', payload);
    }
    /**
     * @public
     * @function insertRow
     * @description This inserts a row just before the selected row
     * @param {number} rowNumber The address of the first cell in the row
     * @returns {Promise<void>} A promise
     */
    insertRow(rowNumber) {
        return this.invokeExcelCall('insertRow', { rowNumber });
    }
    /**
     * @public
     * @function deleteRow
     * @description This deletes the selected row
     * @param {number} rowNumber The address of the first cell in the row
     * @returns {Promise<any>} A promise
     */
    deleteRow(rowNumber) {
        return this.invokeExcelCall('deleteRow', { rowNumber });
    }
    /**
     * @public
     * @function
     * @description activate Activates the current worksheet
     * @returns {Promise<any>} A promise
     */
    activate() {
        return this.invokeExcelCall('activateSheet');
    }
    /**
     * @public
     * @function
     * @description activateCell Activates the selected cell
     * @param {string} cellAddress The address of the cell
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
     * @function clearRange
     * @description Clear the range of formatting and content
     * @param {string} rangeCode The range selected
     * @returns {Promise<void>} A promise
     */
    clearRange(rangeCode) {
        return this.invokeExcelCall('clearRange', { rangeCode });
    }
    /**
     * @public
     * @function clearRangeContents
     * @description Clears the contents in the specified range
     * @param {string} rangeCode The selected range
     * @returns {Promise<void>} A promise
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
     * @function clearAllCells
     * @description Clear all cells and their formatting
     * @returns {Promise<any>} A promise
     */
    clearAllCells() {
        return this.invokeExcelCall('clearAllCells', null);
    }
    /**
     * @public
     * @function clearAllCellContents
     * @description Clears all the cells content
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
     * @function setCellName
     * @description Sets a name for the cell address
     * @param {string} cellAddress The address of the cell e.g. A1
     * @param {string} cellName The name of the cell
     * @returns {Promise<any>} A promise
     */
    setCellName(cellAddress, cellName) {
        return this.invokeExcelCall('setCellName', { address: cellAddress, cellName });
    }
    /**
     * @public
     * @function calculate
     * @description Calculates all formula on teh sheet
     * @returns {Promise<void>} A promise
     */
    calculate() {
        return this.invokeExcelCall('calculateSheet');
    }
    /**
     * @public
     * @function getCellByName
     * @description Gets a cell by its name
     * @param {string} cellName The name of the cell
     * @returns {Promise<any>} A promise
     */
    getCellByName(cellName) {
        return this.invokeExcelCall('getCellByName', { cellName });
    }
    /**
     * @public
     * @function protect
     * @description Password protects the sheet
     * @param {string} password Password used to protect the sheet
     * @returns {Promise<any>} A promise
     */
    protect(password) {
        return this.invokeExcelCall('protectSheet', { password });
    }
    /**
     * @public
     * @function
     * @description toObject Returns only the functions that should be exposed by
     * this class
     * @returns {Worksheet} Public methods in ExcelWorksheet
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