/******/ (function(modules) { // webpackBootstrap
/******/ 	// The module cache
/******/ 	var installedModules = {};
/******/
/******/ 	// The require function
/******/ 	function __webpack_require__(moduleId) {
/******/
/******/ 		// Check if module is in cache
/******/ 		if(installedModules[moduleId]) {
/******/ 			return installedModules[moduleId].exports;
/******/ 		}
/******/ 		// Create a new module (and put it into the cache)
/******/ 		var module = installedModules[moduleId] = {
/******/ 			i: moduleId,
/******/ 			l: false,
/******/ 			exports: {}
/******/ 		};
/******/
/******/ 		// Execute the module function
/******/ 		modules[moduleId].call(module.exports, module, module.exports, __webpack_require__);
/******/
/******/ 		// Flag the module as loaded
/******/ 		module.l = true;
/******/
/******/ 		// Return the exports of the module
/******/ 		return module.exports;
/******/ 	}
/******/
/******/
/******/ 	// expose the modules object (__webpack_modules__)
/******/ 	__webpack_require__.m = modules;
/******/
/******/ 	// expose the module cache
/******/ 	__webpack_require__.c = installedModules;
/******/
/******/ 	// identity function for calling harmony imports with the correct context
/******/ 	__webpack_require__.i = function(value) { return value; };
/******/
/******/ 	// define getter function for harmony exports
/******/ 	__webpack_require__.d = function(exports, name, getter) {
/******/ 		if(!__webpack_require__.o(exports, name)) {
/******/ 			Object.defineProperty(exports, name, {
/******/ 				configurable: false,
/******/ 				enumerable: true,
/******/ 				get: getter
/******/ 			});
/******/ 		}
/******/ 	};
/******/
/******/ 	// getDefaultExport function for compatibility with non-harmony modules
/******/ 	__webpack_require__.n = function(module) {
/******/ 		var getter = module && module.__esModule ?
/******/ 			function getDefault() { return module['default']; } :
/******/ 			function getModuleExports() { return module; };
/******/ 		__webpack_require__.d(getter, 'a', getter);
/******/ 		return getter;
/******/ 	};
/******/
/******/ 	// Object.prototype.hasOwnProperty.call
/******/ 	__webpack_require__.o = function(object, property) { return Object.prototype.hasOwnProperty.call(object, property); };
/******/
/******/ 	// __webpack_public_path__
/******/ 	__webpack_require__.p = "";
/******/
/******/ 	// Load entry module and return exports
/******/ 	return __webpack_require__(__webpack_require__.s = 0);
/******/ })
/************************************************************************/
/******/ ([
/* 0 */
/***/ (function(module, exports) {

var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
// This script determines if the .NET OpenFin.ExcelService is running and if the
// XLL Add-In has been installed. If not, it will perform the deployment, registration,
// and start the service process
fin.desktop.main(() => __awaiter(this, void 0, void 0, function* () {
    const providerVersion = "4.0.6";
    const buildVersion = "4.0.6.0";
    const excelAssetAlias = 'excel-api-addin';
    const excelServiceUuid = '886834D1-4651-4872-996C-7B2578E953B9';
    const installFolder = '%localappdata%\\OpenFin\\shared\\assets\\excel-api-addin';
    const servicePath = 'OpenFin.ExcelService.exe';
    const addInPath = 'OpenFin.ExcelApi-AddIn.xll';
    const excelServiceEventTopic = 'excelServiceEvent';
    try {
        console.log("Starting up excel provider: " + providerVersion + " build: " + buildVersion);
        let serviceIsRunning = yield isServiceRunning();
        let assetInfo = yield getAppAssetInfo();
        console.log("Provider Configured Asset Info: " + JSON.stringify(assetInfo));
        if (serviceIsRunning) {
            console.log('Service Already Running: Skipping Deployment and Registration');
            return;
        }
        if (assetInfo.version === localStorage.installedAssetVersion && !assetInfo.forceDownload) {
            console.log('Current Add-In version previously installed: Skipping Deployment and Registration');
        }
        else {
            yield deploySharedAssets();
            yield tryInstallAddIn();
            console.log("Updating locally stored version number to: " + assetInfo.version);
            localStorage.installedAssetVersion = assetInfo.version;
        }
        yield startExcelService();
        console.log('Excel Service Started');
    }
    catch (err) {
        console.error(err);
    }
    // Technically there is a small window of time between when the UUID is
    // registered as an external application and when the service is ready to
    // receive commands. This edge-case will be best handled in the future 
    // with the availability of plugins and services from the fin API
    function isServiceRunning() {
        return new Promise((resolve, reject) => {
            console.log("Performing check to see if Excel .NET Exe is running: " + excelServiceUuid);
            fin.desktop.System.getAllExternalApplications(extApps => {
                var excelServiceIndex = extApps.findIndex(extApp => extApp.uuid === excelServiceUuid);
                if (excelServiceIndex >= 0) {
                    console.log("Excel .NET Exe uuid found in list of external applications.");
                    resolve(true);
                }
                else {
                    console.log("Excel .NET Exe uuid found in list of external applications.");
                    resolve(false);
                }
            });
        });
    }
    function getAppAssetInfo() {
        return new Promise((resolve, reject) => {
            console.log("Getting app asset info for alias: " + excelAssetAlias);
            fin.desktop.System.getAppAssetInfo({ alias: excelAssetAlias }, resolve, reject);
        });
    }
    function deploySharedAssets() {
        return new Promise((resolve, reject) => {
            console.log("Deploying Shared Assets.");
            fin.desktop.Application.getCurrent().getManifest(manifest => {
                let arguments = `-d "${installFolder}" -c ${manifest.runtime.version}`;
                console.log("Manifest retrieved: " + JSON.stringify(manifest));
                console.log(`Launching external process. Alias: ${excelAssetAlias}  target:  ${servicePath} arguments: ${arguments}`);
                fin.desktop.System.launchExternalProcess({
                    alias: excelAssetAlias,
                    target: servicePath,
                    arguments: arguments,
                    listener: result => {
                        console.log(`Asset Deployment completed! Exit Code: ${result.exitCode}`);
                        resolve();
                    }
                }, () => console.log('Deploying Shared Assets. Launch External Process executed.'), err => reject(err));
            });
        });
    }
    function tryInstallAddIn() {
        return new Promise((resolve, reject) => {
            let path = `${installFolder}\\${servicePath}`;
            let arguments = `-i "${installFolder}"`;
            console.log(`Installing Excel Addin. Path: ${path} arguments: ${arguments}`);
            fin.desktop.System.launchExternalProcess({
                path: path,
                arguments: arguments,
                listener: result => {
                    if (result.exitCode === 0) {
                        console.log('Add-In Installed');
                    }
                    else {
                        console.warn(`Installation failed. Exit code: ${result.exitCode}`);
                    }
                    resolve();
                }
            }, () => console.log('Installing Add-In. Launch External Process executed.'), err => reject(err));
        });
    }
    function startExcelService() {
        return new Promise((resolve, reject) => {
            console.log("Starting Excel .NET Exe Service");
            console.log(`Subscribing to: ${excelServiceEventTopic}`);
            let onMessageReceived;
            let connected = false;
            fin.desktop.InterApplicationBus.subscribe('*', excelServiceEventTopic, onMessageReceived = () => {
                console.log("Received message from .NET Exe Service on topic: " + excelServiceEventTopic + ". Unsubscribing now that we have received notification.");
                if (connected) {
                    console.log("We have already recieved a message from the exe indicating connected. Resolve promise.");
                    resolve();
                    return;
                }
                connected = true;
                fin.desktop.InterApplicationBus.unsubscribe('*', excelServiceEventTopic, onMessageReceived, () => {
                    console.log("Unsubscribed from topic: " + excelServiceEventTopic);
                }, err => {
                    console.log("Error while trying to unsubscribe from " + excelServiceEventTopic + " Error: ", err);
                });
                // The channel provider should eventually move into the .NET app
                // but for now it only being used for signalling and providing provider version
                console.log("Creating Channel for client script to connect to: " + excelServiceUuid + " and providing a getVersion function.");
                fin.desktop.InterApplicationBus.Channel.create(excelServiceUuid).then(channel => {
                    channel.register('getVersion', () => {
                        return {
                            providerVersion, buildVersion
                        };
                    });
                });
                resolve();
            });
            console.log("Getting system details to get port.");
            chrome.desktop.getDetails(function (details) {
                let target = `${installFolder}\\${servicePath}`;
                let arguments = '-p ' + details.port;
                console.log(`Details retrieved. Launching external process. Target: ${target} Arguments: ${arguments} UUID: ${excelServiceUuid}`);
                fin.desktop.System.launchExternalProcess({
                    target: target,
                    arguments: arguments,
                    uuid: excelServiceUuid,
                }, process => {
                    console.log('Service Launched: ' + process.uuid);
                }, error => {
                    reject('Error starting Excel service');
                });
            });
        });
    }
}));
//# sourceMappingURL=provider.js.map

/***/ })
/******/ ]);