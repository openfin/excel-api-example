"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const RpcDispatcher_1 = require("./RpcDispatcher");
/**
 * @class Class that represents a workbook
 */
class ExcelWorkbook extends RpcDispatcher_1.RpcDispatcher {
    /**
     * @constructor Constructor for the ExcelWorkbook class
     * @param application The Application this workbook belongs to
     * @param name The name of the workbook
     */
    constructor(application, name) {
        super();
        this.connectionUuid = application.connectionUuid;
        this.application = application;
        this.mWorksheets = {};
        this.mWorkbookName = name;
    }
    /**
     * @private
     * @function getDefaultMessage Gets the default message to be sent over the
     * wire
     * @returns {any} An object with the workbook name in as default
     */
    getDefaultMessage() {
        return { workbook: this.mWorkbookName };
    }
    /**
     * @public
     * @property Worksheets tied to this workbook
     * @returns {Worksheets}
     */
    get worksheets() {
        return this.mWorksheets;
    }
    set worksheets(worksheets) {
        this.mWorksheets = worksheets;
    }
    /**
     * @public
     * @property workbookName property
     * @returns {string} The name of the workbook
     */
    get name() {
        return this.mWorkbookName;
    }
    /**
     * @public
     * @property Sets the workbook name
     */
    set name(name) {
        this.mWorkbookName = name;
    }
    /**
     * @public
     * @function getWorksheets Gets the worksheets tied to this workbook
     * @returns A promise with worksheets as the result
     */
    getWorksheets() {
        return this.invokeExcelCall('getWorksheets', null);
    }
    /**
     * @public
     * @function getWorksheetByName Gets the worksheet by name
     * @param name The name of the worksheet
     * @returns {ExcelWorksheet} The excel worksheet with the specified name
     */
    getWorksheetByName(name) {
        const worksheet = this.worksheets[name];
        if (!worksheet) {
            console.error(`No worksheet found with the name: ${name}`);
            return;
        }
        return this.worksheets[name].toObject();
    }
    /**
     * @public
     * @function addWorksheet Adds a new worksheet to the workbook
     * @returns {Promise<any>} A promise
     */
    addWorksheet() {
        return this.invokeExcelCall('addSheet', null);
    }
    /**
     * @public
     * @function activate Activates the workbook
     * @returns {Promise<any>} A promise
     */
    activate() {
        return this.invokeExcelCall('activateWorkbook');
    }
    /**
     * @public
     * @function save Save the workbook
     * @returns {Promise<void>} A promise
     */
    save() {
        return this.invokeExcelCall('saveWorkbook');
    }
    /**
     * @public
     * @function close Closes the workbook
     * @returns {Promise<void>} A promise
     */
    close() {
        return this.invokeExcelCall('closeWorkbook');
    }
    /**
     * @public
     * @function toObject Returns only the methods exposed
     * @returns {Workbook} Returns only the methods exposed
     */
    toObject() {
        return this.objectInstance || (this.objectInstance = {
            addEventListener: this.addEventListener.bind(this),
            removeEventListener: this.removeEventListener.bind(this),
            name: this.name,
            activate: this.activate.bind(this),
            addWorksheet: this.addWorksheet.bind(this),
            close: this.close.bind(this),
            getWorksheetByName: (name) => {
                return this.getWorksheetByName(name);
            },
            getWorksheets: this.getWorksheets.bind(this),
            save: this.save.bind(this),
            toObject: this.toObject.bind(this)
        });
    }
}
exports.ExcelWorkbook = ExcelWorkbook;
//# sourceMappingURL=ExcelWorkbook.js.map