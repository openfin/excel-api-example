// This script exists to mimic the functionality that will ultimately be provided
// by the OpenFin services API. It's primary purpose is to deploy the shared assets
// needed by Excel to a common location, and to start the ExcelService process
fin.desktop.main(() => {
    var excelServiceUuid = '886834D1-4651-4872-996C-7B2578E953B9';
    var installFolder = '%localappdata%\\OpenFin\\shared\\assets\\excel-api-addin';
    var servicePath = 'OpenFin.ExcelService.exe';
    var addInPath = 'OpenFin.ExcelApi-AddIn.xll';
    var excelServiceEventTopic = 'excelServiceEvent';
    // This promise resolves when the ExcelService is ready
    var excelServicePromise = Promise.resolve()
        .then(assertServiceIsNotRunning)
        .then(() => Promise.resolve()
        .then(deploySharedAssets)
        .then(startExcelService)
        .catch(err => console.error(err)))
        .catch(() => console.log('Service Already Running: Skipping Deployment and Registration'));
    // Technically there is a small window of time between when the UUID is
    // registered as an external application and when the service is ready to
    // receive commands. This edge-case will be best handled in the future 
    // with the availability of plugins and services from the fin API
    function assertServiceIsNotRunning() {
        return new Promise((resolve, reject) => {
            fin.desktop.System.getAllExternalApplications(extApps => {
                var excelServiceIndex = extApps.findIndex(extApp => extApp.uuid === excelServiceUuid);
                if (excelServiceIndex >= 0) {
                    reject();
                }
                else {
                    resolve();
                }
            });
        });
    }
    function deploySharedAssets() {
        return new Promise((resolve, reject) => {
            fin.desktop.Application.getCurrent().getManifest(manifest => {
                console.log('Deploying Shared Assets');
                fin.desktop.System.launchExternalProcess({
                    alias: 'excel-api-addin',
                    target: servicePath,
                    arguments: '-d "' + installFolder + '" -c ' + manifest.runtime.version,
                    listener: function (args) {
                        console.log('Asset Deployment completed! ' + args.exitCode);
                        resolve();
                    }
                });
            });
        });
    }
    function startExcelService() {
        return new Promise((resolve, reject) => {
            var onExcelServiceEvent;
            fin.desktop.InterApplicationBus.subscribe('*', excelServiceEventTopic, onExcelServiceEvent = () => {
                console.log('Excel Service Alive');
                fin.desktop.InterApplicationBus.unsubscribe('*', excelServiceEventTopic, onExcelServiceEvent);
                resolve();
            });
            chrome.desktop.getDetails(function (details) {
                fin.desktop.System.launchExternalProcess({
                    target: installFolder + '\\OpenFin.ExcelService.exe',
                    arguments: '-p ' + details.port,
                    uuid: excelServiceUuid,
                }, process => {
                    console.log('Service Launched: ' + process.uuid);
                }, error => {
                    reject('Error starting Excel service');
                });
            });
        });
    }
    // This is a very shallow polyfill for the services API
    window.fin.desktop.Service = {
        connect: serviceOpts => {
            var serviceUuid = serviceOpts.uuid;
            if (serviceUuid !== excelServiceUuid) {
                console.error('Unknown service UUID!');
            }
            return excelServicePromise;
        }
    };
});
//# sourceMappingURL=service-loader.js.map