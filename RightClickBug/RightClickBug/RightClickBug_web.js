if (typeof RightClickBug == "undefined") {
    RightClickBug = {};
}

// create export div
var qDiv = document.createElement("div")
qDiv.id = "export"

// create export button
var qButton = document.createElement("input");
qButton.setAttribute("type", "button")
qButton.value = "Export";

// create download link div
var dlDiv = document.createElement("div")
dlDiv.id = "download"

// link for download
var aLink = document.createElement("a");

function obj2Arr(x) {
    var row = []
    row.push(x["Item_Number"])
    row.push(x["Quantity"])
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

function CreateBody() {

    //add export button
    window.document.body.appendChild(qDiv);
    qDiv.appendChild(qButton)
    // download div
    window.document.body.appendChild(dlDiv)
}


// INTERFACE CREATION METHOD
RightClickBug.CreateInterface = function () {
    // create body
    CreateBody()
}

// RUNS COMMAND(S) FROM _CLIENT IN FORMIT
RightClickBug.ShowBug = async function (sendBool) {
    await FormItInterface.CallMethod("RightClickBug.GetArray",sendBool, function (result) {
        FormItInterface.ConsoleLog("Result: " + result)
        RightClickBug.ExportXls(result)
    })
}


var doSend = true;
// run CollectData function on button click
qButton.onclick = function () {
    RightClickBug.ShowBug(doSend);
};

// test url
// https://evolveames.github.io/RightClickBug/RightClickBug
