const logLevels = fin.desktop.System.logLevels;
// This script exists to mimic the functionality that will ultimately be
// provided by the OpenFin services API. It's primary purpose is to deploy the
// shared assets needed by Excel to a common location, and to start the
// ExcelService process
fin.desktop.main(() => {
    let consoleLog;
    let consoleError;
    const excelServiceUuid = '886834D1-4651-4872-996C-7B2578E953B9';
    const installFolder = '%localappdata%\\OpenFin\\shared\\assets\\excel-api-addin';
    const servicePath = 'OpenFin.ExcelService.exe';
    const excelServiceEventTopic = 'excelServiceEvent';
    // This promise resolves when the ExcelService is ready
    Promise.resolve()
        .then(configureLogger)
        .then(assertServiceIsNotRunning)
        .then(() => Promise.resolve()
        .then(deploySharedAssets)
        .then(tryInstallAddIn)
        .then(startExcelService)
        .catch(err => consoleError(err)))
        .catch(() => consoleLog('Service Already Running: Skipping Deployment and Registration'));
    function configureLogger() {
        return new Promise((resolve) => {
            fin.desktop.System.getMinLogLevel((logLevel) => {
                if (logLevel === logLevels.INFO) {
                    consoleLog = console.log;
                    consoleError = console.error;
                }
                else {
                    consoleLog = () => { };
                    consoleError = () => { };
                }
                resolve();
            });
        });
    }
    // Technically there is a small window of time between when the UUID is
    // registered as an external application and when the service is ready to
    // receive commands. This edge-case will be best handled in the future
    // with the availability of plugins and services from the fin API
    function assertServiceIsNotRunning() {
        return new Promise((resolve, reject) => {
            fin.desktop.System.getAllExternalApplications((extApps) => {
                const excelServiceIndex = extApps.findIndex((extApp) => extApp.uuid === excelServiceUuid);
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
            fin.desktop.Application.getCurrent().getManifest((manifest) => {
                fin.desktop.System.launchExternalProcess(Object.assign({
                    alias: 'excel-api-addin',
                    target: servicePath,
                    arguments: `-d "${installFolder}" -c ${manifest.runtime.version}`,
                    listener: (result) => {
                        consoleLog(`Asset Deployment completed! Exit Code: ${result.exitCode}`);
                        resolve();
                    }
                }), () => consoleLog('Deploying Shared Assets'), (err) => reject(err));
            });
        });
    }
    function tryInstallAddIn() {
        const xllInstalledCookie = 'openfin-xll-installed';
        return new Promise((resolve, reject) => {
            if (document.cookie.includes(xllInstalledCookie)) {
                consoleLog('Add-In previously installed');
                resolve();
            }
            else {
                fin.desktop.System.launchExternalProcess({
                    path: `${installFolder}\\${servicePath}`,
                    arguments: `-i "${installFolder}"`,
                    listener: result => {
                        if (result.exitCode === 0) {
                            consoleLog('Add-In Installed');
                            document.cookie = `${xllInstalledCookie}; expires=Fri, 31 Dec 9999 23:59:59 GMT`;
                            resolve();
                        }
                        else {
                            reject(new Error(`Installation failed. Exit code: ${result.exitCode}`));
                        }
                    }
                }, () => consoleLog('Installing Add-In'), err => reject(err));
            }
        });
    }
    function startExcelService() {
        return new Promise((resolve, reject) => {
            let onExcelServiceEvent;
            fin.desktop.InterApplicationBus.subscribe('*', excelServiceEventTopic, onExcelServiceEvent = () => {
                consoleLog('Excel Service Alive');
                fin.desktop.InterApplicationBus.unsubscribe('*', excelServiceEventTopic, onExcelServiceEvent);
                resolve();
            });
            chrome.desktop.getDetails((details) => {
                fin.desktop.System.launchExternalProcess(Object.assign({
                    target: `${installFolder}\\${servicePath}`,
                    arguments: '-p ' + details.port,
                    uuid: excelServiceUuid,
                }), (process) => {
                    consoleLog('Service Launched: ' + process.uuid);
                }, () => {
                    reject('Error starting Excel service');
                });
            });
        });
    }
});
//# sourceMappingURL=service-loader.js.map