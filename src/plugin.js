"use strict";
// This is the entry point of the Plugin script
const ExcelApi_1 = require("./ExcelApi");
window.fin.desktop.ExcelService = ExcelApi_1.ExcelService.instance;
Object.defineProperty(window.fin.desktop, 'Excel', {
    get() { return ExcelApi_1.ExcelService.defaultApplication; }
});
//# sourceMappingURL=plugin.js.map