const xlsx = require('xlsx'); // Requiring module to modify excel files

/* create a new blank workbook */
//var wb = xlsx.utils.book_new();
  
var ws_name = "Einkaufsliste";

/* make worksheet */
var ws_data1 = [
  [ "S", "h", "e", "e", "t", "J", "S" ],
];
var ws_data2 = [
  [ "H", "a", "l", "l", "o"],
];



/*xlsx.utils.sheet_add_json(ws, [{ elementsOfRow: ws_data1[0][1] }], {skipHeader : true, origin: -1});
xlsx.utils.sheet_add_json(ws, [{ elementsOfRow: ws_data1[0][2] }], {skipHeader : true, origin: -1});
//xlsx.utils.sheet_add_json(ws, [{test: ws_data [0][1]}], {skipHeader : true, origin: { r: 1, c: 1}});
*/



var wb = xlsx.readFile("../Einkaufsliste2.xlsx"); // parse the file
  
sheets = wb.SheetNames

var sheetOpenBom = wb.Sheets[wb.SheetNames[0]]; // get the first worksheet
var jsonDataOpenBom = xlsx.utils.sheet_to_json(sheetOpenBom, {defval:""}); //convert sheet to json

var ws = xlsx.utils.json_to_sheet(jsonDataOpenBom,{skipHeader : true});
xlsx.utils.sheet_add_json(ws, ws_data1, {origin: -1, skipHeader : true}); 

console.log(jsonDataOpenBom);


  /* Add the worksheet to the workbook */
xlsx.utils.book_append_sheet(wb, ws, ws_name);

/* write whole workbook to "testBom.xlxs" */
xlsx.writeFile(wb,"../test.xlsx")



