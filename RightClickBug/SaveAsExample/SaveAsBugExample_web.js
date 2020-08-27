﻿if (typeof SaveAsExample == "undefined") {
    SaveAsExample = {};
}

// for parsed LGQ catalog
var parsedCsvCatalog


// create input div, file input (hidden), and label
var inputDiv = document.createElement("div")
var fileInput = document.createElement("input")
fileInput.type = "file"
fileInput.name = "file"
fileInput.id = "file"
fileInput.className = "inputfile"

var uploadButton = document.createElement("label")
uploadButton.htmlFor = "file"
uploadButton.innerHTML = "Choose .CSV Catalog..."

// div for displaying if file is ready
var readyDiv = document.createElement("div")
readyDiv.className = "ready"
readyDiv.innerHTML = "Waiting for .CSV file..."

// create loading div
var loadDiv = document.createElement("div")
loadDiv.className = "hide"

// loading gif in loading div
var loadGif = document.createElement("img")
loadGif.src = "helpers/Images/loading.gif"

// create quantification div
var qDiv = document.createElement("div")
qDiv.id = "quantify"

// create Quantify button
var qButton = document.createElement("input");
qButton.setAttribute("type", "button")
qButton.value = "Quantify";
qButton.disabled = true;

// create download link div
var dlDiv = document.createElement("div")
dlDiv.id = "download"

// link for download
var aLink = document.createElement("a");

// CSV EXPORT METHOD
SaveAsExample.ExportCsv = function (array) {
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
SaveAsExample.ExportXls = function (exportData) {

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
    // add file input button
    window.document.body.appendChild(inputDiv)
    inputDiv.appendChild(fileInput)
    inputDiv.appendChild(uploadButton)
    window.document.body.appendChild(readyDiv)
    //add quantify button
    window.document.body.appendChild(qDiv);
    qDiv.appendChild(qButton)
    // download and loading divs
    window.document.body.appendChild(loadDiv)
    loadDiv.appendChild(loadGif)
    window.document.body.appendChild(dlDiv)
}

function parseCSV(file, delimiter, callback) {
    // convert csv data to array
    Papa.parse(file, {
        header: true,
        download: true,
        dynamicTyping: true,
        complete: function (results) {
            //console.log("Finished:", results);
            callback(results.data);
        }
    });
};

// INTERFACE CREATION METHOD
SaveAsExample.CreateInterface = function () {
    // create body
    CreateBody()

    // event listener for when file is chosen
    var input = document.getElementById("file")
    var label = input.nextElementSibling;
    var labelVal = label.innerHTML;

    input.addEventListener('change', function (e) {
        // update label depending on file chosen
        var fileName = '';
        fileName = e.target.value.split('\\').pop();
        //file was picked...
        if (fileName) {
            // update button text
            label.innerHTML = fileName;
            // does it end in .csv?
            if (fileName.endsWith(".csv")) {
                // show quantify button now that csv is loaded
                qButton.disabled = false;
                // read csv
                var files = this.files;
                for (var i = 0; i < files.length; i++) {
                    parseCSV(files[i], ',', function (result) {
                        parsedCsvCatalog = result
                    
                    });
                }
                readyDiv.textContent = "CSV loaded!"
            }
            // file loaded, but not csv
            else {
                readyDiv.textContent = "Must be a CSV file!"
                qButton.disabled = true;
            }
        }
        // no file loaded
        else {
            label.innerHTML = labelVal;
            qButton.disabled = true;
        }
    })
}



// RUNS COMMAND(S) FROM _CLIENT IN FORMIT
SaveAsExample.findMatchingGroupsAndQuantify = async function (catalog) {
    await FormItInterface.CallMethod("SaveAsExample.GetArray", catalog, function (result) {
        FormItInterface.ConsoleLog("Result: " + result)
        SaveAsExample.ExportXls(result)
        loadDiv.className = "hide"
    })
}




// run CollectData function on button click
qButton.onclick = function () {
    loadDiv.className = "loading"
    SaveAsExample.findMatchingGroupsAndQuantify(parsedCsvCatalog);
};

// test url
// https://github.com/evolveAmes/RightClickBug.git
