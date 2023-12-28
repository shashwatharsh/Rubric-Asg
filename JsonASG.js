// used xlsx package for using creating excel file and multiple sheets
const XLSX = require("xlsx");

// fetched data of json file
const jsonObject = require('./data.json');


// creating excel workbook
const workBook = XLSX.utils.book_new();

// creating excel worksheet 
const workSheet = XLSX.utils.json_to_sheet(jsonObject);

// appending the worksheet with workbook
XLSX.utils.book_append_sheet(workBook, workSheet, "Sheet1");

// creating multiple sheets of nested object 
const workSheet2 = XLSX.utils.json_to_sheet(jsonObject[0].test)
XLSX.utils.book_append_sheet(workBook, workSheet2, "test");

const workSheet3 = XLSX.utils.json_to_sheet(jsonObject[0].test[0].test2);
XLSX.utils.book_append_sheet(workBook, workSheet3, "test2");

const workSheet4 = XLSX.utils.json_to_sheet(jsonObject[0].sections[0].books);
XLSX.utils.book_append_sheet(workBook, workSheet4, jsonObject[0].sections[0].sectionName);

const workSheet5 = XLSX.utils.json_to_sheet(jsonObject[0].sections[1].books);
XLSX.utils.book_append_sheet(workBook, workSheet5, jsonObject[0].sections[1].sectionName);


// buffer 
XLSX.write(workBook, { bookType: "xlsx", type: "buffer" });

// binary
XLSX.write(workBook, { bookType: "xlsx", type: "binary" });

// creating excel sheet 
XLSX.writeFile(workBook,"newExcel.xlsx");