fin.desktop.main(() => {
    var excelServiceUuid = '886834D1-4651-4872-996C-7B2578E953B9';
    var installFolder = '%localappdata%\\OpenFin\\shared\\assets\\excel-api-addin';
    var servicePath = 'OpenFin.ExcelService.exe';
    var addInPath = 'OpenFin.ExcelApi-AddIn.xll';

    var excelServicePromise = Promise.resolve()
        .then(assertServiceIsNotRunning)
        .then(() =>
            Promise.resolve()
                .then(deploySharedAssets)
                .then(startExcelService)
                .then(registerAddIn))
        .catch(() => console.log('Service Already Running: Skipping Deployment and Registration'));
    
    function assertServiceIsNotRunning() {
        return new Promise((resolve, reject) => {
            fin.desktop.System.getAllExternalApplications(extApps => {
                var excelServiceIndex = extApps.findIndex(extApp => extApp.uuid === excelServiceUuid);
                if (excelServiceIndex >= 0) {     
                    reject();
                } else {
                    resolve();
                }
            });
        });
    }

    function deploySharedAssets() {
        return new Promise((resolve, reject) => {
            console.log('Deploying Shared Assets');
            fin.desktop.System.launchExternalProcess({
                alias: 'excel-api-addin',
                target: servicePath,
                arguments: '-d "' + installFolder + '"',
                listener: function (args) {
                    console.log('Installer script completed! ' + args.exitCode);
                    resolve();
                }
            });
        });
    }

    function startExcelService() {
        return new Promise((resolve, reject) => {
            var onServiceStarted;
            fin.desktop.Excel.instance.addEventListener('started', onServiceStarted = () => {
                fin.desktop.Excel.instance.removeEventListener('started', onServiceStarted);
                console.log('Service Started');
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

    function registerAddIn() {
        return new Promise((resolve, reject) => {
            console.log('Installing Add-In');
            fin.desktop.Excel.install(ack => {
                resolve();
            });
        });
    }

    fin.desktop.Service = {
        connect: serviceOpts => {
            var serviceUuid = serviceOpts.uuid;

            if (serviceUuid !== excelServiceUuid) {
                console.error('Unknown service UUID!');
            }

            return excelServicePromise;
        }
    };
});