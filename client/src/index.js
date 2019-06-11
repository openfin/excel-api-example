"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
// This is the entry point of the Plugin script
const ExcelApi_1 = require("./ExcelApi");
window.fin.desktop.ExcelService = ExcelApi_1.ExcelService.instance;
Object.defineProperty(window.fin.desktop, 'Excel', {
    get() { return ExcelApi_1.ExcelService.instance.defaultApplicationObj; }
});
//# sourceMappingURL=index.js.map