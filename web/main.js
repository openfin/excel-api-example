/**
 * Created by haseebriaz on 14/05/15.
 */

// fin.desktop.Excel API Injected via preload script

fin.desktop.main(function () {
    // Initialization and startup logic for Excel is at the very bottom

    var view = {};
    [].slice.call(document.querySelectorAll('[id]')).forEach(element => view[element.id] = element);

    var displayContainers = new Map([
        [view.noConnectionContainer, { windowHeight: 195 }],
        [view.noWorkbooksContainer, { windowHeight: 195 }],
        [view.workbooksContainer, { windowHeight: 830 }]
    ]);

    var rowLength = 27;
    var columnLength = 12;

    var excelInstance;
    var currentWorksheet = null;
    var currentWorkbook = null;
    var currentCell = null;

    // Initialization

    function initializeTable() {

        for (var i = 0; i <= rowLength; i++) {
            var isHeaderRow = i === 0;
            var rowClass = isHeaderRow ? "cellHeader" : "cell";

            var row = createRow(i, columnLength, rowClass, isHeaderRow);
            var rowTarget = isHeaderRow ? view.worksheetHeader : view.worksheetBody;
            rowTarget.appendChild(row);
        }
    }

    function createRow(rowNumber, columnCount, rowClassName, isHeaderRow) {
        var row = document.createElement("tr");

        for (var i = 0; i <= columnCount; i++) {
            var isHeaderCell = i === 0;
            var cellClass = isHeaderCell ? "rowNumber" : rowClassName;
            var editable = !(isHeaderRow || isHeaderCell);

            var cellContent =
                isHeaderCell && !isHeaderRow ? rowNumber.toString() :
                    !isHeaderCell && isHeaderRow ? String.fromCharCode(64 + i) : // Only support one letter-columns for now
                        undefined;

            var cell = createCell(cellClass, cellContent, editable);
            row.appendChild(cell);
        }

        return row;
    }

    function createCell(cellClassName, cellContent, editable) {

        var cell = document.createElement("td");
        cell.className = cellClassName;

        if (cellContent !== undefined) {
            cell.innerText = cellContent;
        }

        if (editable) {

            cell.contentEditable = true;

            cell.addEventListener("keydown", onDataChange);
            cell.addEventListener("blur", onDataChange);
            cell.addEventListener("mousedown", onCellClicked);
        } else {
            cell.onmousedown = contextMenu;
            cell.addEventListener("mousedown", onRowSelected);
        }

        return cell;
    }

    function initializeUIEvents() {
        view.newWorkbookTab.addEventListener("click", function () {
            fin.desktop.Excel.addWorkbook();
        });

        view.openWorkbookTab.addEventListener("click", function () {
            view.dialogOverlay.style.visibility = "visible";
        });

        view.newWorksheetButton.addEventListener("click", function () {
            currentWorkbook.addWorksheet();
        });

        view.launchExcelLink.addEventListener("click", function () {
            connectToExcel();
        });

        view.newWorkbookLink.addEventListener("click", function () {
            fin.desktop.Excel.addWorkbook();
        });

        view.dialogOverlay.addEventListener("click", function (e) {
            if (e.target === view.dialogOverlay) {
                view.dialogOverlay.style.visibility = "hidden";
            } else {
                e.stopPropagation();
            }
        });

        view.openWorkbookButton.addEventListener("click", function (e) {
            view.dialogOverlay.style.visibility = "hidden";
            fin.desktop.Excel.openWorkbook(view.openWorkbookPath.value);
        });

        window.addEventListener("keydown", function (event) {

            switch (event.keyCode) {

                case 78: // N
                    if (event.ctrlKey) fin.desktop.Excel.addWorkbook();
                    break;
                case 37: // LEFT
                    selectAdjacentCell('left');
                    break;
                case 38: // UP
                    selectAdjacentCell('above');
                    break;
                case 39: // RIGHT
                    selectAdjacentCell('right');
                    break;
                case 40: //DOWN
                    selectAdjacentCell('below');
                    break;
            }
        });

    }

    function initializeExcelEvents() {
        fin.desktop.ExcelService.addEventListener("excelConnected", onExcelConnected);
        fin.desktop.ExcelService.addEventListener("excelDisconnected", onExcelDisconnected);
    }

    // UI Functions

    function setDisplayContainer(containerToDisplay) {
        if (!displayContainers.has(containerToDisplay)) {
            return;
        }

        // Element is already showing
        if (containerToDisplay.style.display === "block") {
            return;
        }

        // Reset the display containers to their CSS definition (none)
        for (var container of displayContainers.keys()) {
            container.style.display = null;
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

    function setStatusLabel(text) {
        view.connectionStatus.innerText = text;
    }

    function addWorkbookTab(name) {
        var button = getWorkbookTab(name);
        button.addEventListener("click", onWorkbookTabClicked);
        view.workbookTabs.insertBefore(button, view.newWorkbookTab);
    }

    function getWorkbookTab(name) {
        var elementId = 'workbook-'.concat(name);
        var element = document.getElementById(elementId) || document.createElement('button');

        element.id = elementId;
        element.className = 'workbookTab';
        element.innerHTML = name;

        return element;
    }

    function selectWorkbook(workbook) {
        if (currentWorkbook) {

            var tab = getWorkbookTab(currentWorkbook.name);
            if (tab) tab.className = "workbookTab";
        }

        getWorkbookTab(workbook.name).className = "workbookTabSelected";
        currentWorkbook = workbook;
        currentWorkbook.getWorksheets().then(updateSheets);
    }

    function addWorksheetTab(worksheet) {
        var sheetsTabHolder = view.worksheetTabs;
        var button = getWorksheetTab(worksheet.name);
        button.addEventListener("click", onSheetButtonClicked);
        sheetsTabHolder.insertBefore(button, view.newWorksheetButton);

        worksheet.addEventListener("sheetChanged", onSheetChanged);
        worksheet.addEventListener("selectionChanged", onSelectionChanged);
        worksheet.addEventListener("sheetActivated", onSheetActivated);
        worksheet.addEventListener("rowDeleted", onRowDeleted);
        worksheet.addEventListener("rowInserted", onRowInserted);
    }

    function getWorksheetTab(name) {
        var elementId = 'worksheet-'.concat(name);
        var element = document.getElementById(elementId) || document.createElement('button');

        element.id = elementId;
        element.className = 'tab';
        element.innerText = name;

        return element;
    }

    function selectWorksheet(sheet) {

        if (currentWorksheet === sheet) {
            return;
        }

        if (currentWorksheet) {
            var tab = getWorksheetTab(currentWorksheet.name);
            if (tab) tab.className = "tab";
        }
        getWorksheetTab(sheet.name).className = "tabSelected";
        currentWorksheet = sheet;
        currentWorksheet.getCells("A1", columnLength, rowLength).then(updateData);
    }

    function updateSheets(worksheets) {

        var sheetsTabHolder = view.worksheetTabs;
        while (sheetsTabHolder.firstChild) {

            sheetsTabHolder.removeChild(sheetsTabHolder.firstChild);
        }

        sheetsTabHolder.appendChild(view.newWorksheetButton);
        for (var i = 0; i < worksheets.length; i++) {

            addWorksheetTab(worksheets[i]);
        }

        selectWorksheet(worksheets[0]);
    }

    function getAddress(td) {

        var column = td.cellIndex;
        var row = td.parentElement.rowIndex;
        var offset = view.worksheetHeader.children[0].children[column].innerText.toString() + row;
        return { column: column, row: row, offset: offset };
    }

    function selectCell(cell, preventDefault) {

        clearAllSelectedCells();

        if (currentCell && currentCell.parentNode.parentNode) {

            currentCell.className = "cell";
            updateCellNumberClass(currentCell, "rowNumber", "cellHeader");
        }

        currentCell = cell;
        currentCell.className = "cellSelected";
        view.formulaInput.innerText = "Formula: " + cell.title;
        cell.focus();

        updateCellNumberClass(cell, "rowNumberSelected", "cellHeaderSelected");

        var address = getAddress(currentCell);

        if (!preventDefault) {
            currentWorksheet.activateCell(address.offset);
        }
    }

    function selectAdjacentCell(direction) {
        if (!currentCell) return;
        var info = getAddress(currentCell);

        var cell;

        switch (direction) {
            case 'above':
                if (info.row <= 1) return;
                cell = view.worksheetBody.childNodes[info.row - 2].childNodes[info.column];
                break;
            case 'below':
                if (info.row >= rowLength) return;
                cell = view.worksheetBody.childNodes[info.row].childNodes[info.column];
                break;
            case 'left':
                if (info.column <= 1) return;
                cell = view.worksheetBody.childNodes[info.row - 1].childNodes[info.column - 1];
                break;
            case 'right':
                if (info.column >= columnLength) return;
                cell = view.worksheetBody.childNodes[info.row - 1].childNodes[info.column + 1];
                break;
        }

        if (cell) {
            selectCell(cell);
        }
    }

    function updateData(data) {

        var row = null;
        var currentData = null;

        for (var i = 0; i < data.length; i++) {

            row = view.worksheetBody.childNodes[i];

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

    function updateCellNumberClass(cell, className, headerClassName) {
        var row = cell.parentNode;
        var columnIndex = Array.prototype.indexOf.call(row.childNodes, cell);
        //let rows = document.getElementById('worksheetBody');
        var rowIndex = Array.prototype.indexOf.call(row.parentNode.childNodes, row);
        view.worksheetBody.childNodes[rowIndex].childNodes[0].className = className;
        view.worksheetHeader.children[0].children[columnIndex].className = headerClassName;
    }

    // UI Event Handlers
    function clearAllSelectedCells() {
        let rowNumberSelected = document.querySelectorAll('.rowNumberSelected');
        rowNumberSelected.forEach((rowNumber) => {
            rowNumber.className = 'rowNumber';
        });

        let cellHeaderSelected = document.querySelectorAll('.cellHeaderSelected');
        cellHeaderSelected.forEach((header) => {
            header.className = 'cellHeader';
        });

        let selectedCells = document.querySelectorAll('.cellSelected');
        selectedCells.forEach((cell) => {
            cell.className = 'cell';
        });
    }

    function onCellClicked(event) {
        if (event.target.class !== 'rowNumber') {
            selectCell(event.target);
        }
    }

    function selectRow(rowHeader) {
        clearAllSelectedCells();
        let cells = rowHeader.parentElement.children;

        if (cells[0].innerHTML) {
            cells[0].className = 'rowNumberSelected';
            for (let i = 1; i < cells.length; i++) {
                cells[i].className = 'cellSelected';
            }
        }
        let currentCellAddress = `A${cells[0].innerHTML}`;
        currentWorksheet.activateRow(currentCellAddress);
    }

    function onRowSelected(event) {
        event.stopImmediatePropagation();
        event.preventDefault();
        selectRow(event.target);
    }

    function onSheetButtonClicked(event) {
        var sheet = currentWorkbook.getWorksheetByName(event.target.innerText);
        if (currentWorksheet === sheet) return;
        sheet.activate();
    }

    function onWorkbookTabClicked(event) {
        var workbook = fin.desktop.Excel.getWorkbookByName(event.target.innerText);
        workbook.activate();
    }

    function onDataChange(event) {

        if (event.keyCode === 13 || event.type === "blur") {

            var update = getAddress(event.target);
            update.value = event.target.innerText;

            currentWorksheet.setCells([[update.value]], update.offset);
            if (event.type === "keydown") {

                selectAdjacentCell('below');
                event.preventDefault();
            }
        }
    }

    // Excel Helper Functions

    function checkConnectionStatus() {
        fin.desktop.Excel.getConnectionStatus().then(connected => {
            if (connected) {
                console.log('Already connected to Excel, synthetically raising event.');
                onExcelConnected(fin.desktop.Excel);
            } else {
                setStatusLabel("Excel not connected");
                setDisplayContainer(view.noConnectionContainer);
            }
        });
    }

    function connectToExcel() {
        console.log('connectToExcel');
        setStatusLabel("Connecting...");

        return fin.desktop.Excel.run();
    }

    // Excel Event Handlers

    function onExcelConnected(data) {
        if (excelInstance) {
            return;
        }

        console.log("Excel Connected: " + data.connectionUuid);
        setStatusLabel("Connected to Excel");

        // Grab a snapshot of the current instance, it can change!
        excelInstance = fin.desktop.Excel;

        excelInstance.addEventListener("workbookAdded", onWorkbookAdded);
        excelInstance.addEventListener("workbookOpened", onWorkbookAdded);
        excelInstance.addEventListener("workbookClosed", onWorkbookRemoved);
        excelInstance.addEventListener("workbookSaved", onWorkbookSaved);

        let workbooks = fin.desktop.Excel.getWorkbooks().then(function () {
            for (var i = 0; i < workbooks.length; i++) {
                addWorkbookTab(workbooks[i].name);
                workbooks[i].addEventListener("workbookActivated", onWorkbookActivated);
                workbooks[i].addEventListener("sheetAdded", onWorksheetAdded);
                workbooks[i].addEventListener("sheetRemoved", onWorksheetRemoved);
                workbooks[i].addEventListener("sheetRenamed", onWorksheetRenamed);
            }

            if (workbooks.length) {
                selectWorkbook(workbooks[0]);
                setDisplayContainer(view.workbooksContainer);
            }
            else {
                setDisplayContainer(view.noWorkbooksContainer);
            }
        });
    }

    function onExcelDisconnected(data) {
        if (!excelInstance) {
            return;
        }

        console.log("Excel Disconnected: " + data.connectionUuid);

        if (data.connectionUuid !== excelInstance.connectionUuid) {
            return;
        }

        excelInstance.removeEventListener("workbookAdded", onWorkbookAdded);
        excelInstance.removeEventListener("workbookOpened", onWorkbookAdded);
        excelInstance.removeEventListener("workbookClosed", onWorkbookRemoved);
        excelInstance.removeEventListener("workbookSaved", onWorkbookSaved);

        excelInstance = undefined;

        checkConnectionStatus();
    }

    function onWorkbookAdded(event) {
        var workbook = event.workbook;

        workbook.addEventListener("workbookActivated", onWorkbookActivated);
        workbook.addEventListener("sheetAdded", onWorksheetAdded);
        workbook.addEventListener("sheetRemoved", onWorksheetRemoved);
        workbook.addEventListener("sheetRenamed", onWorksheetRenamed);

        addWorkbookTab(workbook.name);

        workbook.getWorksheets().then((sheets) => {
            sheets.forEach((sheet) => {
                sheet.addEventListener("rowDeleted", onRowDeleted);
                sheet.addEventListener("rowInserted", onRowInserted);
            });
        });

        setDisplayContainer(view.workbooksContainer);
    }

    function onWorkbookRemoved(event) {
        currentWorkbook = null;
        var workbook = event.workbook;
        workbook.removeEventListener("workbookActivated", onWorkbookActivated);
        workbook.removeEventListener("sheetAdded", onWorksheetAdded);
        workbook.removeEventListener("sheetRemoved", onWorksheetRemoved);
        workbook.removeEventListener("sheetRenamed", onWorksheetRenamed);

        view.workbookTabs.removeChild(getWorkbookTab(workbook.name));

        if (view.workbookTabs.children.length < 3) {
            setDisplayContainer(view.noWorkbooksContainer);
        }
    }

    function onWorkbookActivated(event) {
        selectWorkbook(event.target);
    }

    function onWorkbookSaved(event) {
        var workbook = event.workbook;
        var oldWorkbookName = event.oldWorkbookName;

        var button = getWorkbookTab(oldWorkbookName);

        button.id = workbook.name;
        button.innerText = workbook.name;
    }

    function onWorksheetAdded(event) {
        console.log('worksheetadded');
        addWorksheetTab(event.worksheet);
    }

    function onWorksheetRemoved(event) {
        var worksheet = event.worksheet;

        if (worksheet.workbook === currentWorkbook) {
            worksheet.removeEventListener("sheetChanged", onSheetChanged);
            worksheet.removeEventListener("selectionChanged", onSelectionChanged);
            worksheet.removeEventListener("sheetActivated", onSheetActivated);
            worksheet.removeEventListener("rowDeleted", onRowDeleted);
            worksheet.removeEventListener("rowInserted", onRowInserted);
            view.worksheetTabs.removeChild(getWorksheetTab(worksheet.name));
            currentWorksheet = null;
        }
    }

    function onSheetActivated(event) {
        selectWorksheet(event.target);
    }

    function onRowDeleted(event) {
        deleteRow(event.data.range);
    }

    function onRowInserted(event) {
        insertRow(event.data.range);
    }

    function onWorksheetRenamed(event) {
        var worksheet = event.worksheet;
        var oldWorksheetName = event.oldWorksheetName;

        var button = getWorksheetTab(oldWorksheetName);
        button.id = worksheet.name;
        button.innerText = worksheet.name;
    }

    function onSelectionChanged(event) {
        let target;
        if (event.data.width === 1) {
            target = view.worksheetBody.children[event.data.row - 1].children[event.data.column];
            selectCell(target, true);
        } else {
            target = view.worksheetBody.children[event.data.row - 1].children[event.data.column - 1];
            selectRow(target);
        }
    }

    function onSheetChanged(event) {
        var cell = view.worksheetBody.children[event.data.row - 1].children[event.data.column];
        updateCell(cell, event.data.value, event.data.formula);
    }

    // Right click context menu
    function contextMenu(event) {
        event.preventDefault();
        if (event.which === 3) {
            let menu = document.getElementById('excelContextMenu');

            menu.style.left = `${event.pageX}px`
            menu.style.top = `${event.pageY}px`
            menu.style.display = "block";
        }
    }

    /**
     * This function will be attached on click of the delete button and will insert a row above
     * @returns {void}
     */
    function insertRowDelegate() {
        let rowNumber = document.getElementsByClassName('rowNumberSelected')[0].innerText;
        insertRow(rowNumber);
        currentWorksheet.insertRow(rowNumber);
        let menu = document.getElementById('excelContextMenu');
        menu.style.display = 'none';
    }

    /**
     * This function will be attached on click of the delete button and will insert a row above
     * @returns {void}
     */
    function deleteRowDelegate() {
        let rowNumber = document.getElementsByClassName('rowNumberSelected')[0].innerText;
        deleteRow(rowNumber);
        currentWorksheet.deleteRow(rowNumber);
        let menu = document.getElementById('excelContextMenu');
        menu.style.display = 'none';
    }

    /**
     * Inserts a row above the currently selected row
     * @param {any} range The row number
     */
    function insertRow(range) {
        if (isNaN(range)) {
            console.error('Either no range has been passed or the range is not a number');
            return;
        }

        let rowNumber = parseInt(range);
        let rowToInsertAbove = document.getElementById('worksheetBody').children[rowNumber - 1];

        var row = createRow(rowToInsertAbove.children[0].innerText, columnLength, 'cellSelected', false);
        clearAllSelectedCells();
        rowToInsertAbove.parentNode.insertBefore(row, rowToInsertAbove);
        let rows = document.querySelectorAll('tr');

        for (let i = rowNumber; i < rows.length; i++) {
            let rowNumberElement = rows[i].children[0];
            rowNumberElement.innerHTML = rowNumber;
            rowNumber++;
        }
    }

    /**
     * Deletes the selected row
     * @param {any} range The selected row
     */
    function deleteRow(range) {
        if (isNaN(range)) {
            console.error('Either no range has been passed or the range is not a number');
            return;
        }

        let rowNumber = parseInt(range);

        document.getElementById('worksheetBody').deleteRow(rowNumber - 1);

        let rows = document.querySelectorAll('tr');

        for (let i = rowNumber; i < rows.length; i++) {
            let rowNumberElement = rows[i].children[0];
            rowNumberElement.innerHTML = rowNumber;
            rowNumber++;
        }
    }

    // Main App Start
    document.getElementById('insert').onclick = insertRowDelegate;
    document.getElementById('delete').onclick = deleteRowDelegate;
    initializeTable();
    initializeUIEvents();
    initializeExcelEvents();

    fin.desktop.ExcelService.init()
        .then(checkConnectionStatus)
        .catch(err => console.error(err));

    fin.desktop.System.getEnvironmentVariable("userprofile", profilePath => {
        view.openWorkbookPath.value = profilePath + "\\Documents\\";
    });
});