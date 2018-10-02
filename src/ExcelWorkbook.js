"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const RpcDispatcher_1 = require("./RpcDispatcher");
/**
 * @class
 * @description Class that represents a workbook
 */
class ExcelWorkbook extends RpcDispatcher_1.RpcDispatcher {
    /**
     * @constructor
     * @description Constructor for the ExcelWorkbook class
     * @param {Application} application The Application this workbook belongs to
     * @param {string}name The name of the workbook
     */
    constructor(application, name) {
        super();
        this.connectionUuid = application.connectionUuid;
        this.application = application;
        this.mWorksheets = {};
        this.mWorkbookName = name;
        this.objectInstance = null;
    }
    /**
     * @private
     * @function getDefaultMessage
     * @description Gets the default message to be sent over the
     * wire
     * @returns {object} An object with the workbook name in as default
     */
    getDefaultMessage() {
        return { workbook: this.mWorkbookName };
    }
    /**
     * @public
     * @property
     * @description Worksheets tied to this workbook
     * @returns {Worksheets} The worksheets tied to this workbook
     */
    get worksheets() {
        return this.mWorksheets;
    }
    /**
     * @public
     * @property
     * @description Set the worksheets that are tied to this workbook
     */
    set worksheets(worksheets) {
        this.mWorksheets = worksheets;
    }
    /**
     * @public
     * @property
     * @description workbookName property
     * @returns {string} The name of the workbook
     */
    get name() {
        return this.mWorkbookName;
    }
    /**
     * @public
     * @property
     * @description Sets the workbook name
     * @param {string} name Set the name of the workbook
     */
    set name(name) {
        this.mWorkbookName = name;
    }
    /**
     * @public
     * @function getWorksheets
     * @description Gets the worksheets tied to this workbook
     * @returns {Promise<Worksheets>} A promise with worksheets as the result
     */
    getWorksheets() {
        return this.invokeExcelCall('getWorksheets', null);
    }
    /**
     * @public
     * @function getWorksheetByName
     * @description Gets the worksheet by name
     * @param {string} name The name of the worksheet
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
     * @description Adds a new worksheet to the workbook
     * @returns {Promise<Worksheet>} A promise
     */
    addWorksheet() {
        return this.invokeExcelCall('addSheet', null);
    }
    /**
     * @public
     * @function activate
     * @description Activates the workbook
     * @returns {Promise<void>} A promise
     */
    activate() {
        return this.invokeExcelCall('activateWorkbook');
    }
    /**
     * @public
     * @function save
     * @description Save the current workbook
     * @returns {Promise<void>} A promise
     */
    save() {
        return this.invokeExcelCall('saveWorkbook');
    }
    /**
     * @public
     * @function close
     * @description Closes the workbook
     * @returns {Promise<void>} A promise
     */
    close() {
        return this.invokeExcelCall('closeWorkbook');
    }
    /**
     * @public
     * @function toObject
     * @description Returns only the methods exposed
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