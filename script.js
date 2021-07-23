/**
 * This script uses SheetJS to parse excel file into js object that we can process in browser. 
 * see docs here:
 * https://github.com/SheetJS/sheetjs?utm_source=cdnjs&utm_medium=cdnjs_link&utm_campaign=cdnjs_library
 * 
 * one API used here is FileReader, see docs here:
 * https://developer.mozilla.org/en-US/docs/Web/API/FileReader
 */


//query the #upload element
let $upload = $("#upload");

//add an onChange event listener
$upload.on("change", (e) => {
    // get the file from our input element
    // @ts-ignore
    let file = $upload[0].files[0];

    //create a file reader, file reader will read file from disk asynchronously
    let fr = new FileReader();

    //add an onload event handler, this callback function executes when the file is loaded
    fr.onload = (e) => {
        //get the loaded data in binary form
        let data = e.target.result;

        //pass the data to XLSX library
        // @ts-ignore
        let workbook = XLSX.read(data, {
            type: "binary"
        })

        //create an array to hold each sheet
        let sheets = [];

        //loop through each sheet
        for(const sheetName of workbook.SheetNames) {
            // get rows of the sheet as an array, each row will be an object inside this array
            // @ts-ignore
            let rowsArray = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);
            // put rows in sheets array
            sheets.push(rowsArray);
        }

        //show the result in console
        console.log(sheets);
    };

    //say something when a file read error happens
    fr.onerror = (err) => {
        console.error(err);
    }

    //start the file reading process, previous lines were describing what to do when a file is loaded
    //this line tells the program to start reading the file.
    fr.readAsBinaryString(file);
})