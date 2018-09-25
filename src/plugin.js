"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
// This is the entry point of the Plugin script
const ExcelApi_1 = require("./ExcelApi");
const excelService = new ExcelApi_1.ExcelService();
// Attach ExcelService to the window
window.fin.desktop.ExcelService = excelService;
// Attach the Excel api to the window
Object.defineProperty(window.fin.desktop, 'Excel', {
    get() {
        return excelService.defaultApplicationObj;
    }
});
fin.desktop.main(() => {
    // For dev purposes
    fin.desktop.System.deleteCacheOnExit();
    function init(message) {
        console.log(message);
        excelService.init()
            .then(() => {
            fin.desktop.InterApplicationBus.unsubscribe('886834D1-4651-4872-996C-7B2578E953B9', 'init', init, () => {
                console.log('Successfully unsubscribed from initialisation');
            }, (reason) => {
                console.error(reason);
            });
        })
            .catch((err) => {
            console.log('This error might be ok', err);
        });
    }
    fin.desktop.InterApplicationBus.subscribe('886834D1-4651-4872-996C-7B2578E953B9', 'init', init);
    fin.desktop.InterApplicationBus.send('886834D1-4651-4872-996C-7B2578E953B9', 'init-multi-window', 'initial fire');
});
//# sourceMappingURL=plugin.js.map