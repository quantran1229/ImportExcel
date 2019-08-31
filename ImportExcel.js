const _ = require('lodash')
const XlsxPopulate = require('xlsx-populate');

//get row number
function getRow(cell) {
    let sheet = Object.values(cell)[0];
    let workbook = Object.values(sheet)[1];
    let a = workbook.attributes;
    return a.r;
}

//get column number
function getColumn(cell, value) {
    let sheet = Object.values(cell)[0];
    let workbook = Object.values(sheet)[1];
    let a = workbook.children;
    let b;
    a.forEach((node) => {
        if (Object.values(node)[3] === value)
            b = Object.values(node)[1];
    })
    return b;
}

//write name of cell
function toCell(column, row) {
    let c = '';
    let n = column;
    do {
        n--;
        c = String.fromCharCode(65 + (n % 26)) + c;
        n = Math.floor(n / 26);
    } while (n)
    return c + row;
}

function find(sheet, name) {
    let result = [];
    let list = sheet.find(name);
    for (let entry of list) {
        if (entry._value === name) {
            result.push(entry);
        }
    }
    return result
}

//export to sheet
let findArray = (sheet, list) => {

    //find where STT first appear in worksheet
    let firstCells = find(sheet, list[0]);
    let menubar_col = [];

    let initialList = []
    for (let cell of firstCells) {
        initialList.push(getRow(cell))
    }

    list.forEach((name) => {
        let foundCells = find(sheet, name);
        if (foundCells.length > 0) {
            let newList = [];
            for (let cell of foundCells) {
                newList.push(getRow(cell))
            }
            initialList = _.intersection(initialList, newList)
        } else {
            console.log("Can't find", name);
        }
    });

    if (initialList.length < 1) {
        throw new Error("WRONG EXCEL FORMAT, Can't find menubar")
    }

    let longestRows = 0;
    let rowsCount = 0;
    for (let choice = 0; choice < initialList.length; choice++) {
        let no = initialList[choice]
        let col = 1;
        let firstBarList = find(sheet, list[0])
        for (let firstElement of firstBarList) {
            if (getRow(firstElement) == no) {
                col = getColumn(firstElement, list[0]);
                break;
            }
        }
        while (sheet.cell(toCell(col, no)).value()) {
            no++;
        }

        if (no - initialList[choice] > rowsCount) {
            rowsCount = no - initialList[choice];
            longestRows = choice
        }
    }

    let correctRow = initialList[longestRows]
    let firstCell = null;

    list.forEach((name) => {
        let foundCells = find(sheet, name);
        if (foundCells.length > 0) {
            for (let cell of foundCells) {
                if (getRow(cell) == correctRow) {
                    if (firstCell == null)
                        firstCell = cell;
                    menubar_col.push(getColumn(cell, name));
                    break;
                }
            }
        } else {
            menubar_col.push(1000);
        }
    });

    let startRow = getRow(firstCell);

    let returnValue = {
        "startRow": startRow + 1,
        "menubar": menubar_col
    }

    return returnValue;
}

let findJson = (sheet, listMenu, callback) => {
    let value = findArray(sheet, listMenu);
    let list = value.menubar;
    let n = value.startRow;
    let result = [];
    let check = true
    while (check) {
        check = false
        let entry = {};
        for (let col = 0; col < list.length; col++) {
            entry[listMenu[col]] = ""
            entry[listMenu[col]] = sheet.cell(toCell(list[col], n)).value();
            if (!isNaN(entry[listMenu[col]])){
                
                check = true;
            }
        }
        if (check) {
            if (callback) callback(entry)
            result.push(entry);
        }
        n++;
    }
    return result;
}

let getJSON = async (filename, sheetNumber, list, callback) => {
    let workbook = await XlsxPopulate.fromFileAsync("./" + filename);
    let sheet = workbook.sheet(sheetNumber);
    let result = findJson(sheet, list, callback)
    return result;
}

module.exports = {
    getJSON
}