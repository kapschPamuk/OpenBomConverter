const xlsx = require('xlsx'); // requiring module to modify excel files
const fetch = require("node-fetch") //requiring module for HTTP requests

const openbom_appkey = "60539257ed9d7a343a900dc0" //appkey for retrieving access_token after login

/* credentials for login */
const body = {
  "username": "halilibrahim.pamuk@kapsch.net",
  "password": "Df123456!"
}

/* variables used in start() */
var token
var access_token
var orderBOM
var orderBOMcolumns
var orderBOMcells
var created=false
var sb = 0


/* entry point of application */
start()


async function start() {
  /* login */
  token = await login()
  /* get access_token */
  access_token = token.access_token

  /* get specific orderBOM document by ID */
  orderBOM = await getSpecificBOM("05642718-f3e8-4e60-a427-18f3e88e60dc")
  orderBOMc = await getSpecificCatalog("a2d685a6-3170-4239-9685-a6317092390d")
  allBOMs = await getAllBOMS()

  //console.log(allBOMs[0].name);

  /* get columns of the document */
  orderBOMcolumns = orderBOM.columns
  /* get cells of the document */
  orderBOMcells = orderBOM.cells


  /* write data to excel */
  createSupExcel()


  /* identify the names of the subboms from the document and find out their ID. Then,execute getSpecificBOM()*/
  for (var subbomIndex = 0; subbomIndex < orderBOMcells.length; subbomIndex++) {
    subbom = orderBOM.cells[subbomIndex][0]
    //console.log(subbom);
    //console.log(orderBOMc);

    /* verify that subbom is actually in catalog subboms */

    for (var k = 0; k < orderBOMc.cells.length; k++) {
      if (orderBOMc.cells[k][0] == subbom) {
        //console.log(subbom + " is a subbom");
      }
    }
 
    for (var i = 0; i < allBOMs.length; i++) { //go through all BOMs, if a name matches with a subbom name of the document, return ID of that subbom
      if (allBOMs[i].name == subbom) {
        //console.log(allBOMs[i].id)
        subbomID = allBOMs[i].id

        /* print each subbom to excel */
        excelSubBOM = await getSpecificBOM(subbomID)
        /* get columns of the subbom */
        subBOMcolumns = excelSubBOM.columns
        /* get cells of the subbom */
        subBOMcells = excelSubBOM.cells

        createSubExcel()
        if(i==allBOMs.length-1) endReached = true //not used yet
      }

    }
  }
  
  combineExcelsInOneDocument()

}

/* login to get an access_token for further API requests */
function login() {
  return fetch("https://developer-api.openbom.com/login", {
    method: "post",
    body: JSON.stringify(body),
    headers: {
      "Content-Type": "application/json",
      "x-openbom-appkey": openbom_appkey
    }
  })
    .then(res => res.json())
    .catch(err => console.log(err))
}

/* get data about a specific orderBOM document by ID */
function getSpecificBOM(documentID) {
  return fetch("https://developer-api.openbom.com/bom/" + documentID, {
    method: "get",
    headers: {

      "x-openbom-accesstoken": access_token,
      "x-openbom-appkey": openbom_appkey
    }
  })
    .then(res => res.json())
    .catch(err => console.log(err))
}

/* get data about a specific catalog document by ID */
function getSpecificCatalog(documentID) {
  return fetch("https://developer-api.openbom.com/catalog/" + documentID, {
    method: "get",
    headers: {
      "x-openbom-accesstoken": access_token,
      "x-openbom-appkey": openbom_appkey
    }
  })
    .then(res => res.json())
    .catch(err => console.log(err))
}

/* get data about a specific catalog document by ID */
function getAllBOMS() {
  return fetch("https://developer-api.openbom.com/boms", {
    method: "get",
    headers: {
      "x-openbom-accesstoken": access_token,
      "x-openbom-appkey": openbom_appkey
    }
  })
    .then(res => res.json())
    .catch(err => console.log(err))
}


/* create new excel workbook and write data of API request to excel */
function createSupExcel() {


  /* create a new blank workbook */
  var wb = xlsx.utils.book_new();

  /* assign a name to the worksheet in the workbook */
  var ws_name = "Einkaufsliste";

  /* initial header row */
  /*var ws = xlsx.utils.json_to_sheet([], {header: ['Kreditor','K,Nummer','Incoterm','KOK?',
    'KD,#','CDP BE#','Angebot #','Hersteller','HOK?','Hersteller-Artikel#','Kred,#','Bezeichnung','Stk,','Liste/Stk,','Rabatt%',
    'EK/Stk,','EK/Ges,','Fracht','TOTAL'], skipHeader: false});
  */

  /* create empty worksheet */
  var ws = xlsx.utils.json_to_sheet([]);

  
  /* set cell width */
  /*var wscols = [
    { wch: 50 }, // "characters"
    { wch: 50 }, // "characters"
    { wch: 50 }, // "characters"
    { wch: 50 }, // "characters"
    { wch: 50 }, // "characters"
    { wch: 50 }, // "characters"
    { wch: 50 }, // "characters"
    { wch: 50 }, // "characters"
    { wch: 50 }, // "characters"
    { wch: 50 }, // "characters"
    //{wpx: 500}, // "pixels"
    ,
    { hidden: true } // hide column
  ];
  ws['!cols'] = wscols;
*/


  /* variables to identify a specific cell */
  var row
  var column

  //--------------------------------------------------------------------------------------------------------

  /* fill header with orderBOMcolumns.length values */
  row = 0
  column = 0
  for (var headerIndex = 0; headerIndex < orderBOMcolumns.length; headerIndex++) {
    xlsx.utils.sheet_add_json(ws, [], { origin: { r: row, c: column }, header: [orderBOMcolumns[headerIndex]] }); //Write headers to worksheet
    column++ //move to next column in same row
  }

  /* fill body with orderBOMcells.length rows and orderBOM.columns.length values per row */
  row = 1
  column = 0
  for (var i = 0; i < orderBOMcells.length; i++) { //i = row
    for (var j = 0; j < orderBOM.columns.length; j++) { //j = column
      xlsx.utils.sheet_add_json(ws, [{ elementsOfRow: orderBOMcells[i][j] }], { skipHeader: true, origin: { r: row, c: column } }); // Write data starting from specific row and column to worksheet
      column++ //move to next column in same row
      if (column == orderBOM.columns.length) column = 0 //once column reaches orderBOM.columns.length, move to first column again in the next row
    }
    row++ //move to next row
  }


  /* add the worksheet with the name ws_name to the workbook */
  xlsx.utils.book_append_sheet(wb, ws, ws_name);

  /* write whole workbook to "Einkaufsliste.xlxs" */
  xlsx.writeFile(wb, "../SubBoms_list.xlsx")

}


/* create new excel workbook and write data of API request to excel */
function createSubExcel() {

if(created==false){
  //console.log("created");
  /* create a new blank workbook */
  var wb = xlsx.utils.book_new();
 
}
else {
  //console.log("already exists");
  var wb = xlsx.readFile("../SubBoms_seperated.xlsx"); // parse the file
}

  /* create empty worksheet */
  var ws = xlsx.utils.json_to_sheet([]);

  /* assign a name to the worksheet in the workbook */
  var ws_name = ["SubBom"+sb]

  //--------------------------------------------------------------------------------------------------------

  /* fill header with subBOMcolumns.length values */
  row = 0 
  column = 0
  for (var headerIndex = 0; headerIndex < subBOMcolumns.length; headerIndex++) {
    xlsx.utils.sheet_add_json(ws, [], { origin: { r: row, c: column }, header: [subBOMcolumns[headerIndex]] }); //Write headers to worksheet
    column++ //move to next column in same row
  }

  /* fill body with subBOMcells.length rows and subBOMcolumns.length values per row */
  row = 1 
  column = 0
  for (var i = 0; i < subBOMcells.length; i++) { //i = row
    for (var j = 0; j < subBOMcolumns.length; j++) { //j = column
      xlsx.utils.sheet_add_json(ws, [{ elementsOfRow: subBOMcells[i][j] }], { skipHeader: true, origin: { r: row, c: column } }); // Write data starting from specific row and column to worksheet
      column++ //move to next column in same row
      if (column == subBOMcolumns.length) column = 0 //once column reaches orderBOM.columns.length, move to first column again in the next row
    }
    row++ //move to next row
  }

  //--------------------------------------------------------------------------------------------------------


  /* add the worksheet with the name ws_name to the workbook */
  xlsx.utils.book_append_sheet(wb, ws, ws_name);
  sb++
  
  /* write whole workbook to "Einkaufsliste.xlxs" */
  xlsx.writeFile(wb, "../SubBoms_seperated.xlsx")

  created=true
}

/* combine created new excels in one sheet */
function combineExcelsInOneDocument(){

  var ws_name = "Einkaufsliste";
  
  var wb = xlsx.readFile("../SubBoms_seperated.xlsx"); // parse the file
    
  sheets = wb.SheetNames
  
  var ws = xlsx.utils.json_to_sheet([]);
  
  var headerSet = false
  for (var i = 0 ; i < sheets.length ; i ++ ){
  
  
  if(headerSet){ // Insert header first time
    var jsonData = xlsx.utils.sheet_to_json(wb.Sheets[wb.SheetNames[i]],{skipHeader : true, defval:""})
    xlsx.utils.sheet_add_json(ws, jsonData, {origin: -1,skipHeader : true}); 
    //console.log("ohne header");
  }
  
  else { // Skip header next times
    var jsonData = xlsx.utils.sheet_to_json(wb.Sheets[wb.SheetNames[i]],{defval:""})
    xlsx.utils.sheet_add_json(ws, jsonData, {origin: { r: 0, c: 0 }}); 
    //console.log("mit header");
    headerSet = true
  }
  }

  /* Add the worksheet to the workbook */
  xlsx.utils.book_append_sheet(wb, ws, ws_name);
  
  /* Hide all the elements of SubBoms_seperated from SubBoms_combined */
  for(var subboms = 0 ; subboms < sheets.length - 1; subboms ++){
    xlsx.utils.book_set_sheet_visibility(wb,subboms,xlsx.utils.consts.SHEET_HIDDEN);
  }

  /* write whole workbook to "Einkaufsliste3.xlxs" */
  xlsx.writeFile(wb,"../SubBoms_combined.xlsx")

}