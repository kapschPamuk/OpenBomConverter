const xlsx = require('xlsx'); // Requiring module to modify excel files
const fetch = require("node-fetch") //HTTP request

// appkey to retrieve access_token after login
const openbom_appkey = "60539257ed9d7a343a900dc0"

// credentials for login
const body = {
    "username":"halilibrahim.pamuk@kapsch.net",
    "password":"Df123456!"
}

var token
var orderBOM
var orderBOMcolumns
var orderBOMcells
var access_token


//entry point of application
start()


async function start() {
  //login
  token = await login()
  //extract access_token from json response
  access_token = token.access_token

  //get specific document by ID
  orderBOM = await getSpecificDocument("05642718-f3e8-4e60-a427-18f3e88e60dc")
  //get columns of orderBOM
  orderBOMcolumns = orderBOM.columns
  //get cells of orderBOM
  orderBOMcells = orderBOM.cells

  createExcel()

}


function login() {
  return fetch("https://developer-api.openbom.com/login", {
    method: "post",
    body: JSON.stringify(body),
    headers: { 
      "Content-Type": "application/json",
      "x-openbom-appkey": "60539257ed9d7a343a900dc0"
    }
  })
  .then(res => res.json())
  .catch(err => console.log(err))
}


function getSpecificDocument(documentID) {
  return fetch("https://developer-api.openbom.com/bom/"+documentID, {
    method: "get",
    headers: { 
      "x-openbom-accesstoken": access_token,
      "x-openbom-appkey": openbom_appkey
    }
  })
  .then(res => res.json())
  //.then(json => {console.log(json.columns);console.log(json.cells)})
  .catch(err => console.log(err))
}


//create new excel workbook and add API result
function createExcel() {


/* create a new blank workbook */
var wb = xlsx.utils.book_new();
  
var ws_name = "Einkaufsliste";

/* make worksheet */
var ws_data = [
  [ "S", "h", "e", "e", "t", "J", "S" ],
  [  1 ,  2 ,  3 ,  4 ,  5 ]
];

/* Initial header row */
/*var ws = xlsx.utils.json_to_sheet([], {header: ['Kreditor','K,Nummer','Incoterm','KOK?',
	'KD,#','CDP BE#','Angebot #','Hersteller','HOK?','Hersteller-Artikel#','Kred,#','Bezeichnung','Stk,','Liste/Stk,','Rabatt%',
	'EK/Stk,','EK/Ges,','Fracht','TOTAL'], skipHeader: false});
*/

var headerValue=0;
var ws = xlsx.utils.json_to_sheet([],{header: [
  orderBOMcolumns[headerValue++],orderBOMcolumns[headerValue++],
  orderBOMcolumns[headerValue++],orderBOMcolumns[headerValue++],
  orderBOMcolumns[headerValue++],orderBOMcolumns[headerValue++],
  orderBOMcolumns[headerValue++],orderBOMcolumns[headerValue++],
  orderBOMcolumns[headerValue++],orderBOMcolumns[headerValue++],
  orderBOMcolumns[headerValue++],orderBOMcolumns[headerValue++],
  orderBOMcolumns[headerValue++],orderBOMcolumns[headerValue++],
  orderBOMcolumns[headerValue++],orderBOMcolumns[headerValue++],
  orderBOMcolumns[headerValue++],orderBOMcolumns[headerValue++],
  orderBOMcolumns[headerValue++],orderBOMcolumns[headerValue++],
  orderBOMcolumns[headerValue++],orderBOMcolumns[headerValue++],
  orderBOMcolumns[headerValue++],orderBOMcolumns[headerValue++],
  orderBOMcolumns[headerValue++],orderBOMcolumns[headerValue++],
]});

//Find out how many columns each row has 
orderBOMcolumnsLength = orderBOMcolumns.length


for(var i=0 ; i<orderBOMcells.length ; i++){ //i = Zeile
  //for(var j=0 ; j<orderBOM.columns.length ; j++){ //j = Spalte
  var j = 0
  /* Append rows */
  xlsx.utils.sheet_add_json(ws, [{ 
    A: orderBOMcells[i][j++], B: orderBOMcells[i][j++],
    C: orderBOMcells[i][j++], D: orderBOMcells[i][j++],
    E: orderBOMcells[i][j++], F: orderBOMcells[i][j++],
    G: orderBOMcells[i][j++], H: orderBOMcells[i][j++],
    I: orderBOMcells[i][j++], J: orderBOMcells[i][j++],
    K: orderBOMcells[i][j++], L: orderBOMcells[i][j++],
    M: orderBOMcells[i][j++], N: orderBOMcells[i][j++],
    O: orderBOMcells[i][j++], P: orderBOMcells[i][j++],
    Q: orderBOMcells[i][j++], R: orderBOMcells[i][j++],
    S: orderBOMcells[i][j++], T: orderBOMcells[i][j++],
    U: orderBOMcells[i][j++], V: orderBOMcells[i][j++],
    W: orderBOMcells[i][j++], X: orderBOMcells[i][j++],
    Y: orderBOMcells[i][j++], Z: orderBOMcells[i][j++],
  }], {origin: -1, skipHeader: true}); //C: 6, D: 7, E: 8, F: 9, G: 0 
  
//}
}
  
  /* Add the worksheet to the workbook */
xlsx.utils.book_append_sheet(wb, ws, ws_name);

/* write whole workbook to "Einkaufsliste.xlxs" */
xlsx.writeFile(wb,"../Einkaufsliste.xlsx")
}