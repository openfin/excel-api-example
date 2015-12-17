/**
 * Created by haseebriaz on 14/05/15.
 */


var averageColumn, averageRow;
window.addEventListener("DOMContentLoaded", function(){

    var rowLength = 27;
    var columnLength = 12 ;
    var table = document.getElementById("excelExample");
    var tBody = table.getElementsByTagName("tbody")[0];
    var tHead = table.getElementsByTagName("thead")[0];
    var newWorkbookButton = document.getElementById("newWorkbookButton");
    var newWorksheetButton = document.getElementById("newSheetButton");
    newWorksheetButton.addEventListener("click", function(){

        currentWorkbook.addWorksheet();
    });
    var currentWorksheet = null;
    var currentWorkbook = null;
    var currentCell = null;
    var formulaInput = document.getElementById("formulaInput");

    window.addEventListener("keydown", function(event){

        switch(event.keyCode){

            case 78: // N
                if(event.ctrlKey) fin.desktop.Excel.addWorkbook();
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

        };
    });

    function initTable(){

        var row = createRow(["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L"], "cellHeader", false);
        var column = createColumn("");
        column.className = "rowNumber";
        row.insertBefore(column, row.childNodes[0]);
        tHead.appendChild(row);

        for(var i = 1; i <= rowLength; i++){

            row = createRow(columnLength, "cell", true);
            column = createColumn(i);
            column.className = "rowNumber";
            column.contentEditable = false;
            row.insertBefore(column, row.childNodes[0]);
            tBody.appendChild(row);
        }
    }

    function createRow(data, cellClassName, editable){

        var length = data.length? data.length: data;
        var row = document.createElement("tr");

        for(var i = 0; i < length; i++){

            row.appendChild(createColumn(data[i], cellClassName, editable));
        }

        return row;
    }

    function createColumn(data, cellClassName, editable){

        var column = document.createElement("td");
        column.className = cellClassName;

        if(editable){

            column.contentEditable = true;
            //column.addEventListener("DOMCharacterDataModified", onDataChange);
            column.addEventListener("keydown", onDataChange);
            column.addEventListener("blur", onDataChange);
          //  column.addEventListener("mousedown", onCellClicked);
            column.addEventListener("focus", selectCell);
        }

        if(data)column.innerText = data;
        return column;
    }

    function onCellClicked(event){

        selectCell(event.target);
    }

    function selectCell(event){

        var cell = event.target;
        if(currentCell){

            currentCell.className = "cell";
            updateCellNumberClass(currentCell, "rowNumber", "cellHeader");
        }

        currentCell = cell;
        currentCell.className = "cellSelected";
        formulaInput.innerText = "Formula: " + cell.title;

        updateCellNumberClass(cell, "rowNumberSelected", "cellHeaderSelected");

        var address = getAddress(currentCell);
        currentWorksheet.activateCell(address.offset);
    }

    function updateCellNumberClass(cell, className, headerClassName){

        var row = cell.parentNode;
        var columnIndex = Array.prototype.indexOf.call(row.childNodes, cell);
        var rowIndex = Array.prototype.indexOf.call(row.parentNode.childNodes, cell.parentNode);
        tBody.childNodes[rowIndex].childNodes[0].className = className;
        tHead.getElementsByTagName("tr")[0].getElementsByTagName("td")[columnIndex].className = headerClassName;
    }

    function selectCellBelow(){

        if(!currentCell) return;
        var info = getAddress(currentCell);
        if(info.row >= rowLength) return;
        var cell = tBody.childNodes[info.row].childNodes[info.column];
        cell.focus();
    }

    function selectCellAbove(){

        if(!currentCell) return;
        var info = getAddress(currentCell);
        if(info.row <= 1) return;
        var cell = tBody.childNodes[info.row - 2].childNodes[info.column];
        cell.focus();
    }

    function selectNextCell(){

        if(!currentCell) return;
        var info = getAddress(currentCell);
        if(info.column >= columnLength) return;
        var cell = tBody.childNodes[info.row - 1].childNodes[info.column + 1];
        cell.focus();
    }

    function selectPreviousCell(){

        if(!currentCell) return;
        var info = getAddress(currentCell);
        if(info.column <= 1) return;
        var cell = tBody.childNodes[info.row - 1].childNodes[info.column - 1];
        cell.focus();
    }

    function onDataChange(event){

        if(event.keyCode == 13 || event.type == "blur") {

            var update = getAddress(event.target);
            update.value = event.target.innerText;
            currentWorksheet.setCells([[update.value]], update.offset);
            if(event.type == "keydown"){

                selectCellBelow();
                event.preventDefault();
            }
        }
    }

    function getAddress(td){

        var column = td.cellIndex;
        var row = td.parentElement.rowIndex;
        var offset = tHead.getElementsByTagName("td")[column].innerText.toString() + row;
        return {column: column, row: row, offset: offset};
    }

    function updateData(data){

        var row = null;
        var currentData = null;

        for(var i = 0; i < data.length; i++){

            row = tBody.childNodes[i];
            for(var j = 1; j < row.childNodes.length; j++){

                currentData = data[i][j - 1];
                updateCell(row.childNodes[j], currentData.value, currentData.formula );
            }
        }
    }

    function updateCell(cell, value, formula){

        cell.innerText = value || value === 0? value: "";
        cell.title = formula? formula: "";
    }

    function onSheetChanged(event){

        var cell = tBody.getElementsByTagName("tr")[event.data.row - 1].getElementsByTagName("td")[event.data.column];
        updateCell(cell, event.data.value, event.data.formula);
    }

    function onSelectionChanged(event){

        var cell = tBody.getElementsByTagName("tr")[event.data.row - 1].getElementsByTagName("td")[event.data.column];
        cell.focus();
    }

    function onSheetActivated(event){

        selectWorksheet(event.target);
    }

    function selectWorksheet(sheet){

        if(currentWorksheet == sheet) {
            return;
        }

        if(currentWorksheet) {
            var tab = document.getElementById(currentWorksheet.name);
            if(tab)tab.className = "tab";
        }
        document.getElementById(sheet.name).className = "tabSelected";
        currentWorksheet = sheet;
        currentWorksheet.getCells("A1", columnLength, rowLength, updateData);
    }

    function selectWorkbook(workbook){

        if(currentWorkbook) {

            var tab = document.getElementById(currentWorkbook.name);
            if(tab)tab.className = "workbookTab";
        }

        document.getElementById(workbook.name).className = "workbookTabSelected";
        currentWorkbook = workbook;
        currentWorkbook.getWorksheets(updateSheets);
    }

    function onWorkbookTabClicked(event){

        var workbook = fin.desktop.Excel.getWorkbookByName(event.target.innerText);
        if(currentWorkbook == workbook) return;
        workbook.activate();
    }

    function onWorkbookActivated(event){

        selectWorkbook(event.target);
    }

    function onWorkbookAdded(event){

        var workbook = event.workbook;
        workbook.addEventListener("workbookActivated", onWorkbookActivated);
        workbook.addEventListener("sheetAdded", onWorksheetAdded);
        workbook.addEventListener("sheetRemoved", onWorksheetRemoved);
        addWorkbookTab(event.workbook.name);

        if(workbooksContainer.style.display == "none") showWorkbooksContainer();
    }

    function onWorkbookRemoved(event){

        currentWorkbook = null;
        var workbook = event.workbook;
        workbook.removeEventListener("workbookActivated", onWorkbookActivated);
        workbook.removeEventListener("sheetAdded", onWorksheetAdded);
        workbook.removeEventListener("sheetRemoved", onWorksheetRemoved);

        document.getElementById("workbookTabs").removeChild(document.getElementById(event.workbook.name));

        if(document.getElementById("workbookTabs").childNodes.length < 2){

            showNoWorkbooksMessage();
        }
    }

    function onWorksheetAdded(event){

        addWorksheetTab(event.worksheet);
    }

    function addWorksheetTab(worksheet){

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

    function onSheetButtonClicked(event){

        var sheet = currentWorkbook.getWorksheetByName(event.target.innerText);
        if(currentWorksheet == sheet) return;
        sheet.activate();
    }

    function onWorksheetRemoved(event){

        if(event.worksheet.workbook == currentWorkbook){

            event.worksheet.removeEventListener("sheetChanged", onSheetChanged);
            event.worksheet.removeEventListener("selectionChanged", onSelectionChanged);
            event.worksheet.removeEventListener("sheetActivated", onSheetActivated);
            document.getElementById("sheets").removeChild(document.getElementById(event.worksheet.name));
            currentWorksheet = null;
        }
    }

    function updateSheets(worksheets){

        var sheetsTabHolder = document.getElementById("sheets");
        while(sheetsTabHolder.firstChild){

            sheetsTabHolder.removeChild(sheetsTabHolder.firstChild);
        }

        sheetsTabHolder.appendChild(newWorksheetButton);
        for(var i = 0; i < worksheets.length; i++){

            addWorksheetTab(worksheets[i]);
        }

        selectWorksheet(worksheets[0]);
    }

    function addWorkbookTab(name){

        var button = document.createElement("button");
        button.id = button.innerText = name;
        button.className = "workbookTab";
        button.addEventListener("click", onWorkbookTabClicked);
        document.getElementById("workbookTabs").insertBefore(button, newWorkbookButton);
    }

    function showNoWorkbooksMessage(){

        fin.desktop.Window.getCurrent().animate({

            size: {
                height: 195,
                duration: 500
            }
        });
        noWorkbooks.style.display = "block";
        workbooksContainer.style.display = "none";
    }

    function showWorkbooksContainer(){

        workbooksContainer.style.display = "block";
        noWorkbooks.style.display = "none";
        fin.desktop.Window.getCurrent().animate({

            size: {
                height: 830,
                duration: 500
            }
        });
    }

    initTable(27, 12);

    fin.desktop.main(function(){

        var Excel = fin.desktop.Excel;
        Excel.init();
        Excel.getConnectionStatus(onExcelConnected);
        Excel.addEventListener("workbookAdded", onWorkbookAdded);
        Excel.addEventListener("workbookClosed", onWorkbookRemoved);
        Excel.addEventListener("connected", onExcelConnected);

    });

    function onExcelConnected(){

        document.getElementById("status").innerText = "Connected to Excel";
        fin.desktop.Excel.getWorkbooks(function(workbooks){

            for(var i = 0; i < workbooks.length; i++){

                addWorkbookTab(workbooks[i].name);
                workbooks[i].addEventListener("workbookActivated", onWorkbookActivated);
                workbooks[i].addEventListener("sheetAdded", onWorksheetAdded);
                workbooks[i].addEventListener("sheetRemoved", onWorksheetRemoved);
            };

            if(workbooks.length){

                selectWorkbook(workbooks[0]);
                showWorkbooksContainer();
            }
            else {
                showNoWorkbooksMessage();
            }
        });
    }

    averageColumn = function(start, height, output){

        currentWorksheet.getColumn(start, height, function(data){

            var sum = 0;
            for(var i = 0; i < data.length; i++){

                sum += Number(data[i].value);
            }

            currentWorksheet.setCells([[sum/data.length]], output);
        });
    }

    averageRow = function(start, width, output){

        currentWorksheet.getRow(start, width, function(data){

            var sum = 0;
            for(var i = 0; i < data.length; i++){

                sum += Number(data[i].value);
            }

            currentWorksheet.setCells([[sum/data.length]], output);
        });
    }

    checkNullValues = function(start, width, output){
        console.log("Checking null values... ", arguments);
    }

});





