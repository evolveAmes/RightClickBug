﻿if (typeof RightClickBug == "undefined") {
    RightClickBug = {};
}

function obj2Arr(x) {
    var row = []
    row.push(x["Header1"])
    row.push(x["Headrr2"])
    console.log(JSON.stringify(x))
    return row
}

// XLS EXPORT METHOD
RightClickBug.ExportXls = function (exportData) {

    var fileName = "export.xls"

    // variables for excel
    // worksheet name + first row header
    var wsName = "Title"
    var header = ["Header1", "Header2"]

    // array for final data, adding header
    var final = []
    final.push(header)

    // cycle through array to create rows from objects
    // add to final array
    var array = new Array()
    array = JSON.parse(exportData)
    for (var i = 0; i < array.length; i++) {
        var row = obj2Arr(array[i])
        final.push(row)
    }

    // EXCEL -------------
    // create workbook
    var wb = XLSX.utils.book_new()
    var ws = XLSX.utils.aoa_to_sheet(final);

    // add worksheet to workbook 
    XLSX.utils.book_append_sheet(wb, ws, wsName);

    // write workbook 
    XLSX.writeFile(wb, fileName);
}


// RUNS COMMAND(S) FROM _CLIENT IN FORMIT
function RightClickBug(){
    FormItInterface.CallMethod("RightClickBug.GetArray", function (result) {
        FormItInterface.ConsoleLog("Result: " + result)
        RightClickBug.ExportXls(result)
    })
}


// test url
// https://evolveames.github.io/RightClickBug/RightClickBug