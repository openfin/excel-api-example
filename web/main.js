/**
 * Created by haseebriaz on 14/05/15.
 */

// fin.desktop.Excel API Injected via preload script

window.addEventListener("DOMContentLoaded", function () {

    var legacyApi;

    var rowLength = 27;
    var columnLength = 12;
    var table = document.getElementById("excelExample");
    var tBody = table.getElementsByTagName("tbody")[0];
    var tHead = table.getElementsByTagName("thead")[0];

    var newWorkbookTab = document.getElementById("newWorkbookTab");
    var openWorkbookTab = document.getElementById("openWorkbookTab");
    var openWorkbookButton = document.getElementById("openWorkbookButton");

    var newWorksheetButton = document.getElementById("newSheetButton");

    var noConnectionContainer = document.getElementById("noConnection");
    var noWorkbooksContainer = document.getElementById("noWorkbooks");
    var workbooksContainer = document.getElementById("workbooksContainer");

    var openWorkbookPath = document.getElementById("openWorkbookPath");

    var dialogOverlay = document.getElementById("dialogOverlay");

    var connectionStatus = document.getElementById("status");

    var displayContainers = new Map([
        [noConnectionContainer, { windowHeight: 195 }],
        [noWorkbooksContainer, { windowHeight: 195 }],
        [workbooksContainer, { windowHeight: 830 }]
    ]);

    newWorkbookTab.addEventListener("click", function () {
        fin.desktop.Excel.addWorkbook();
    });

    openWorkbookTab.addEventListener("click", function () {
        dialogOverlay.style.visibility = "visible";
    });

    newWorksheetButton.addEventListener("click", function () {
        currentWorkbook.addWorksheet();
    });

    dialogOverlay.addEventListener("click", function (e) {
        if (e.target === dialogOverlay) {
            dialogOverlay.style.visibility = "hidden";
        } else {
            e.stopPropagation();
        }
    });

    openWorkbookButton.addEventListener("click", function (e) {
        dialogOverlay.style.visibility = "hidden";
        fin.desktop.Excel.openWorkbook(openWorkbookPath.value);
    });

    var currentWorksheet = null;
    var currentWorkbook = null;
    var currentCell = null;
    var formulaInput = document.getElementById("formulaInput");

    window.addEventListener("keydown", function (event) {

        switch (event.keyCode) {

            case 78: // N
                if (event.ctrlKey) fin.desktop.Excel.addWorkbook();
                break;
            case 37: // LEFT
                selectPreviousCell();
                break;
            case 38: // UP
                selectCellAbove();
                break;
            case 39: // RIGHT
                selectNextCell();
                break;
            case 40: //DOWN
                selectCellBelow();
                break;
        }
    });

    function setDisplayContainer(containerToDisplay) {
        if (!displayContainers.has(containerToDisplay)) {
            return;
        }

        for (var container of displayContainers.keys()) {
            container.style.display = "none";
        }
        containerToDisplay.style.display = "block";

        var windowHeight = displayContainers.get(containerToDisplay).windowHeight;

        fin.desktop.Window.getCurrent().animate({
            size: {
                height: windowHeight,
                duration: 500
            }
        });
    }

    function initTable() {

        var row = createRow(["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L"], "cellHeader", false);
        var column = createColumn("");
        column.className = "rowNumber";
        row.insertBefore(column, row.childNodes[0]);
        tHead.appendChild(row);

        for (var i = 1; i <= rowLength; i++) {

            row = createRow(columnLength, "cell", true);
            column = createColumn(i);
            column.className = "rowNumber";
            column.contentEditable = false;
            row.insertBefore(column, row.childNodes[0]);
            tBody.appendChild(row);
        }
    }

    function createRow(data, cellClassName, editable) {

        var length = data.length ? data.length : data;
        var row = document.createElement("tr");

        for (var i = 0; i < length; i++) {

            row.appendChild(createColumn(data[i], cellClassName, editable));
        }

        return row;
    }

    function createColumn(data, cellClassName, editable) {

        var column = document.createElement("td");
        column.className = cellClassName;

        if (editable) {

            column.contentEditable = true;
            //column.addEventListener("DOMCharacterDataModified", onDataChange);
            column.addEventListener("keydown", onDataChange);
            column.addEventListener("blur", onDataChange);
            column.addEventListener("mousedown", onCellClicked);
        }

        if (data) column.innerText = data;
        return column;
    }

    function onCellClicked(event) {

        selectCell(event.target);
    }

    function selectCell(cell, preventDefault) {

        if (currentCell) {

            currentCell.className = "cell";
            updateCellNumberClass(currentCell, "rowNumber", "cellHeader");
        }

        currentCell = cell;
        currentCell.className = "cellSelected";
        formulaInput.innerText = "Formula: " + cell.title;
        cell.focus();

        updateCellNumberClass(cell, "rowNumberSelected", "cellHeaderSelected");

        var address = getAddress(currentCell);

        if (!preventDefault) {
            currentWorksheet.activateCell(address.offset);
        }
    }

    function updateCellNumberClass(cell, className, headerClassName) {

        var row = cell.parentNode;
        var columnIndex = Array.prototype.indexOf.call(row.childNodes, cell);
        var rowIndex = Array.prototype.indexOf.call(row.parentNode.childNodes, cell.parentNode);
        tBody.childNodes[rowIndex].childNodes[0].className = className;
        tHead.getElementsByTagName("tr")[0].getElementsByTagName("td")[columnIndex].className = headerClassName;
    }

    function selectCellBelow() {

        if (!currentCell) return;
        var info = getAddress(currentCell);
        if (info.row >= rowLength) return;
        var cell = tBody.childNodes[info.row].childNodes[info.column];
        selectCell(cell);
    }

    function selectCellAbove() {

        if (!currentCell) return;
        var info = getAddress(currentCell);
        if (info.row <= 1) return;
        var cell = tBody.childNodes[info.row - 2].childNodes[info.column];
        selectCell(cell);
    }

    function selectNextCell() {

        if (!currentCell) return;
        var info = getAddress(currentCell);
        if (info.column >= columnLength) return;
        var cell = tBody.childNodes[info.row - 1].childNodes[info.column + 1];
        selectCell(cell);
    }

    function selectPreviousCell() {

        if (!currentCell) return;
        var info = getAddress(currentCell);
        if (info.column <= 1) return;
        var cell = tBody.childNodes[info.row - 1].childNodes[info.column - 1];
        selectCell(cell);
    }

    function onDataChange(event) {

        if (event.keyCode === 13 || event.type === "blur") {

            var update = getAddress(event.target);
            update.value = event.target.innerText;

            currentWorksheet.setCells([[update.value]], update.offset);
            if (event.type === "keydown") {

                selectCellBelow();
                event.preventDefault();
            }
        }
    }

    function getAddress(td) {

        var column = td.cellIndex;
        var row = td.parentElement.rowIndex;
        var offset = tHead.getElementsByTagName("td")[column].innerText.toString() + row;
        return { column: column, row: row, offset: offset };
    }

    function updateData(data) {

        var row = null;
        var currentData = null;

        for (var i = 0; i < data.length; i++) {

            row = tBody.childNodes[i];

            if (!row) {
                continue;
            }

            for (var j = 1; j < row.childNodes.length; j++) {

                currentData = data[i][j - 1];
                updateCell(row.childNodes[j], currentData.value, currentData.formula);
            }
        }
    }

    function updateCell(cell, value, formula) {

        cell.innerText = value ? value : "";
        cell.title = formula ? formula : "";
    }

    function onSheetChanged(event) {
        var cell = tBody.getElementsByTagName("tr")[event.data.row - 1].getElementsByTagName("td")[event.data.column];
        updateCell(cell, event.data.value, event.data.formula);
    }

    function onSelectionChanged(event) {
        var cell = tBody.getElementsByTagName("tr")[event.data.row - 1].getElementsByTagName("td")[event.data.column];
        selectCell(cell, true);
    }

    function onSheetActivated(event) {
        selectWorksheet(event.target);
    }

    function selectWorksheet(sheet) {

        if (currentWorksheet === sheet) {
            return;
        }

        if (currentWorksheet) {
            var tab = document.getElementById(currentWorksheet.name);
            if (tab) tab.className = "tab";
        }
        document.getElementById(sheet.name).className = "tabSelected";
        currentWorksheet = sheet;
        currentWorksheet.getCells("A1", columnLength, rowLength, updateData);
    }

    function selectWorkbook(workbook) {
        if (currentWorkbook) {

            var tab = document.getElementById(currentWorkbook.name);
            if (tab) tab.className = "workbookTab";
        }

        document.getElementById(workbook.name).className = "workbookTabSelected";
        currentWorkbook = workbook;
        currentWorkbook.getWorksheets(updateSheets);
    }

    function onWorkbookTabClicked(event) {
        var workbook = fin.desktop.Excel.getWorkbookByName(event.target.innerText);
        workbook.activate();
    }

    function onWorkbookActivated(event) {
        selectWorkbook(event.target);
    }

    function onWorkbookAdded(event) {
        var workbook = event.workbook;

        workbook.addEventListener("workbookActivated", onWorkbookActivated);
        workbook.addEventListener("sheetAdded", onWorksheetAdded);
        workbook.addEventListener("sheetRemoved", onWorksheetRemoved);
        workbook.addEventListener("sheetRenamed", onWorksheetRenamed);

        addWorkbookTab(workbook.name);

        if (workbooksContainer.style.display === "none") {
            setDisplayContainer(workbooksContainer);
        }
    }

    function onWorkbookRemoved(event) {
        currentWorkbook = null;
        var workbook = event.workbook;
        workbook.removeEventListener("workbookActivated", onWorkbookActivated);
        workbook.removeEventListener("sheetAdded", onWorksheetAdded);
        workbook.removeEventListener("sheetRemoved", onWorksheetRemoved);
        workbook.removeEventListener("sheetRenamed", onWorksheetRenamed);

        document.getElementById("workbookTabs").removeChild(document.getElementById(workbook.name));

        if (document.getElementById("workbookTabs").childNodes.length < 2) {
            setDisplayContainer(noWorkbooksContainer);
        }
    }

    function onWorkbookSaved(event) {
        var workbook = event.workbook;
        var oldWorkbookName = event.oldWorkbookName;

        var button = document.getElementById(oldWorkbookName);

        button.id = workbook.name;
        button.innerText = workbook.name;
    }

    function onWorksheetAdded(event) {
        addWorksheetTab(event.worksheet);
    }

    function addWorksheetTab(worksheet) {
        var sheetsTabHolder = document.getElementById("sheets");
        var button = document.createElement("button");
        button.innerText = worksheet.name;
        button.className = "tab";
        button.id = worksheet.name;
        button.addEventListener("click", onSheetButtonClicked);
        sheetsTabHolder.insertBefore(button, newWorksheetButton);

        worksheet.addEventListener("sheetChanged", onSheetChanged);
        worksheet.addEventListener("selectionChanged", onSelectionChanged);
        worksheet.addEventListener("sheetActivated", onSheetActivated);
    }

    function onSheetButtonClicked(event) {
        var sheet = currentWorkbook.getWorksheetByName(event.target.innerText);
        if (currentWorksheet === sheet) return;
        sheet.activate();
    }

    function onWorksheetRemoved(event) {
        var worksheet = event.worksheet;

        if (worksheet.workbook === currentWorkbook) {
            worksheet.removeEventListener("sheetChanged", onSheetChanged);
            worksheet.removeEventListener("selectionChanged", onSelectionChanged);
            worksheet.removeEventListener("sheetActivated", onSheetActivated);
            document.getElementById("sheets").removeChild(document.getElementById(worksheet.name));
            currentWorksheet = null;
        }
    }

    function onWorksheetRenamed(event) {
        var worksheet = event.worksheet;
        var oldWorksheetName = event.oldWorksheetName;

        var button = document.getElementById(oldWorksheetName);
        button.id = worksheet.name;
        button.innerText = worksheet.name;
    }

    function updateSheets(worksheets) {

        var sheetsTabHolder = document.getElementById("sheets");
        while (sheetsTabHolder.firstChild) {

            sheetsTabHolder.removeChild(sheetsTabHolder.firstChild);
        }

        sheetsTabHolder.appendChild(newWorksheetButton);
        for (var i = 0; i < worksheets.length; i++) {

            addWorksheetTab(worksheets[i]);
        }

        selectWorksheet(worksheets[0]);
    }

    function addWorkbookTab(name) {
        var button = document.createElement("button");
        button.id = button.innerText = name;
        button.className = "workbookTab";
        button.addEventListener("click", onWorkbookTabClicked);
        document.getElementById("workbookTabs").insertBefore(button, newWorkbookTab);
    }

    function onExcelConnected(data) {
        console.log("Excel Connected: " + data.connectionUuid);

        connectionStatus.innerText = "Connected to Excel";

        if (legacyApi) {
            return;
        }

        // Grab a snapshot of the current instance, it can change!
        legacyApi = fin.desktop.Excel.legacyApi;
       
        legacyApi.addEventListener("workbookAdded", onWorkbookAdded);
        legacyApi.addEventListener("workbookOpened", onWorkbookAdded);
        legacyApi.addEventListener("workbookClosed", onWorkbookRemoved);
        legacyApi.addEventListener("workbookSaved", onWorkbookSaved);

        // Wait for subscription to be ready. Fix after true services are available
        //setTimeout(() =>
        fin.desktop.Excel.getWorkbooks(workbooks => {
            for (var i = 0; i < workbooks.length; i++) {
                addWorkbookTab(workbooks[i].name);
                workbooks[i].addEventListener("workbookActivated", onWorkbookActivated);
                workbooks[i].addEventListener("sheetAdded", onWorksheetAdded);
                workbooks[i].addEventListener("sheetRemoved", onWorksheetRemoved);
                workbooks[i].addEventListener("sheetRenamed", onWorksheetRenamed);
            }

            if (workbooks.length) {
                selectWorkbook(workbooks[0]);
                setDisplayContainer(workbooksContainer);
            }
            else {
                setDisplayContainer(noWorkbooksContainer);
            }
        });//, 1000);
    }

    function onExcelDisconnected(data) {
        console.log("Excel Disconnected: " + data.connectionUuid);

        if (data.connectionUuid !== legacyApi.connectionUuid) {
            return;
        }

        legacyApi.removeEventListener("workbookAdded", onWorkbookAdded);
        legacyApi.removeEventListener("workbookOpened", onWorkbookAdded);
        legacyApi.removeEventListener("workbookClosed", onWorkbookRemoved);
        legacyApi.removeEventListener("workbookSaved", onWorkbookSaved);

        legacyApi = undefined;

        if (fin.desktop.Excel.legacyApi) {
            onExcelConnected(fin.desktop.Excel.legacyApi);
        } else {
            connectionStatus.innerText = "Excel not connected";
            setDisplayContainer(noConnectionContainer);
        }
    }

    function connectToExcel() {
        console.log('connectToExcel');
        connectionStatus.innerText = "Connecting...";

        return new Promise((resolve, reject) => {
            fin.desktop.Excel.instance.addEventListener("excelConnected", onExcelConnected);
            fin.desktop.Excel.instance.addEventListener("excelDisconnected", onExcelDisconnected);

            if (fin.desktop.Excel.legacyApi) {
                console.log("Excel Already Running");
                onExcelConnected(fin.desktop.Excel.legacyApi);
                resolve();
            } else {
                console.log("Launching Excel");
                fin.desktop.Excel.run(resolve);
            }
        });
    }

    initTable(27, 12);

    fin.desktop.main(function () {
        fin.desktop.Excel.init()
            .then(connectToExcel);

        fin.desktop.System.getEnvironmentVariable("userprofile", profilePath => {
            openWorkbookPath.value = profilePath + "\\Documents\\";
        });
    });
});
