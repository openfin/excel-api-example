/**
 * Created by haseebriaz on 14/05/15.
 */

window.addEventListener("DOMContentLoaded", function(){

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
        };
    });

    function initTable(rowLength, columnLength){

        var row = createRow(["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L"], "cellHeader");
        var column = createColumn("");
        column.className = "rowNumber";
        row.insertBefore(column, row.childNodes[0]);
        tHead.appendChild(row);

        for(var i = 1; i <= rowLength; i++){

            row = createRow(columnLength, "cell");
            column = createColumn(i);
            column.className = "rowNumber";
            row.insertBefore(column, row.childNodes[0]);
            tBody.appendChild(row);
        }
    }

    function createRow(data, cellClassName){

        var length = data.length? data.length: data;
        var row = document.createElement("tr");

        for(var i = 0; i < length; i++){

            row.appendChild(createColumn(data[i], cellClassName));
        }

        return row;
    }

    function createColumn(data, cellClassName){

        var column = document.createElement("td");
        column.className = cellClassName;
        column.contentEditable = true;
        column.addEventListener("DOMCharacterDataModified", onDataChange);
        column.addEventListener("click", onCellClicked);
        if(data)column.innerText = data;
        return column;
    }

    function selectCell(cell){

        if(currentCell){

            currentCell.className = "cell";
            updateCellNumberClass(currentCell, "rowNumber", "cellHeader");
        }

        currentCell = cell;
        currentCell.className = "cellSelected";
        formulaInput.innerText = "Formula: " + cell.title;

        updateCellNumberClass(cell, "rowNumberSelected", "cellHeaderSelected");
    }

    function updateCellNumberClass(cell, className, headerClassName){

        var row = cell.parentNode;
        var columnIndex = Array.prototype.indexOf.call(row.childNodes, cell);
        var rowIndex = Array.prototype.indexOf.call(row.parentNode.childNodes, cell.parentNode);
        tBody.childNodes[rowIndex].childNodes[0].className = className;
        tHead.getElementsByTagName("tr")[0].getElementsByTagName("td")[columnIndex].className = headerClassName;
    }

    function onCellClicked(event){

        selectCell(event.target);
        var address = getAddress(event.target);
        currentWorksheet.activateCell(address.offset);
    }

    function onDataChange(event){

        var update = getAddress(event.path[1]);
        update.value = event.target.nodeValue;
        currentWorksheet.setCells([[update.value]], update.offset);
        console.log(update);
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

        cell.removeEventListener("DOMCharacterDataModified", onDataChange);
        cell.innerText = value? value: "";
        cell.title = formula? formula: "";
        cell.addEventListener("DOMCharacterDataModified", onDataChange);
    }

    function onSheetChanged(event){

        var cell = tBody.getElementsByTagName("tr")[event.data.row - 1].getElementsByTagName("td")[event.data.column];
        updateCell(cell, event.data.value, event.data.formula);
    }

    function onSelectionChanged(event){

        var cell = tBody.getElementsByTagName("tr")[event.data.row - 1].getElementsByTagName("td")[event.data.column];
        selectCell(cell);
    }

    function onSheetActivated(event){

        selectWorksheet(event.target);
    }

    function selectWorksheet(sheet){

        if(currentWorksheet == sheet) {
            console.log("same");
            return;
        }
        if(currentWorksheet) {
            var tab = document.getElementById(currentWorksheet.name);
            if(tab)tab.className = "tab";
        }
        document.getElementById(sheet.name).className = "tabSelected";
        currentWorksheet = sheet;
        currentWorksheet.getCells("A1", 12, 27, updateData);
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
        selectWorkbook(event.workbook);
    }

    function onWorkbookRemoved(event){

        currentWorkbook = null;
        var workbook = event.workbook;
        workbook.workbook.removeEventListener("workbookActivated", onWorkbookActivated);
        workbook.removeEventListener("sheetAdded", onWorksheetAdded);
        workbook.removeEventListener("sheetRemoved", onWorksheetRemoved);

        document.getElementById("workbookTabs").removeChild(document.getElementById(event.workbook.name));
    }

    function onWorksheetAdded(event){

        console.log("sheetadded");
        addWorksheetTab(event.worksheet);
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

    function onSheetButtonClicked(event){

        var sheet = currentWorkbook.getWorksheetByName(event.target.innerText);
        if(currentWorksheet == sheet) return;
        sheet.activate();
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

    function addWorkbookTab(name){

        var button = document.createElement("button");
        button.id = button.innerText = name;
        button.className = "workbookTab";
        button.addEventListener("click", onWorkbookTabClicked);
        document.getElementById("workbookTabs").insertBefore(button, newWorkbookButton);
    }

    initTable(27, 12);

    fin.desktop.main(function(){

        fin.desktop.Excel.init();
        fin.desktop.Excel.addEventListener("workbookAdded", onWorkbookAdded);
        fin.desktop.Excel.addEventListener("workbookClosed", onWorkbookRemoved);
        fin.desktop.Excel.getWorkbooks(function(workbooks){

            for(var i = 0; i < workbooks.length; i++){

                addWorkbookTab(workbooks[i].name);
                workbooks[i].addEventListener("workbookActivated", onWorkbookActivated);
                workbooks[i].addEventListener("sheetAdded", onWorksheetAdded);
                workbooks[i].addEventListener("sheetRemoved", onWorksheetRemoved);
            };

            selectWorkbook(workbooks[0]);
            currentWorkbook.getWorksheets(updateSheets);
        });
    });

});


