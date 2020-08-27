if (typeof RightClickBug == "undefined") {
    RightClickBug = {};
}

// create loading div
var loadDiv = document.createElement("div")
loadDiv.className = "hide"

// loading gif in loading div
var loadGif = document.createElement("img")
loadGif.src = "helpers/Images/loading.gif"

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

// CSV EXPORT METHOD
RightClickBug.ExportCsv = function (array) {
    // change array back to csv format  
    var csv = Papa.unparse(array)

    // create download link div
    var dlDiv = document.createElement("div")
    dlDiv.id = "download"

    // link for download
    var aLink = document.createElement("a");

    // create full data for export - includes CSV
    var urlData = "data:text/csv;charset=UTF-8," + encodeURIComponent(csv)

    // create link element
    aLink.href = urlData;
    aLink.target = "_blank"
    aLink.download = "export.csv";
    aLink.innerHTML = "Download CSV"

    // create manual link
    dlDiv.appendChild(aLink)

    // wait before clicking link 
    setTimeout(aLink.click(), 300);

};


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

    //add quantify button
    window.document.body.appendChild(qDiv);
    qDiv.appendChild(qButton)
    // download and loading divs
    window.document.body.appendChild(loadDiv)
    loadDiv.appendChild(loadGif)
    window.document.body.appendChild(dlDiv)
}


// INTERFACE CREATION METHOD
RightClickBug.CreateInterface = function () {
    // create body
    CreateBody()
}



// RUNS COMMAND(S) FROM _CLIENT IN FORMIT
RightClickBug.ShowBug = async function () {
    await FormItInterface.CallMethod("RightClickBug.GetArray", function (result) {
        FormItInterface.ConsoleLog("Result: " + result)
        RightClickBug.ExportXls(result)
        loadDiv.className = "hide"
    })
}




// run CollectData function on button click
qButton.onclick = function () {
    loadDiv.className = "loading"
    RightClickBug.ShowBug();
};

// test url
// https://evolveames.github.io/RightClickBug/RightClickBug
