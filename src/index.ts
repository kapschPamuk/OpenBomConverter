import { QMainWindow, QWidget, QLabel, FlexLayout, QPushButton, QIcon} from '@nodegui/nodegui';
import logo from '../greentick.png';
import logoSettings from '../settings.jpg';

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


var settingsFile = null
var settingsFileSet = false

const xlsx = require('xlsx'); // Requiring module to modify excel files

// Flag if file was converted
var converted = false;

const win = new QMainWindow();
win.setWindowTitle("Convert OpenBom");

const centralWidget = new QWidget();
centralWidget.setObjectName("myroot");

const rootLayout = new FlexLayout();
centralWidget.setLayout(rootLayout);

//Settings Button
const settingsButton = new QPushButton();
settingsButton.setObjectName("mysettingsbutton");
settingsButton.setText('Settings');
settingsButton.setIcon(new QIcon(logoSettings));
settingsButton.addEventListener('clicked', () => {
	console.log('Settings');
});

//Conversion Button
const converterButton = new QPushButton();
converterButton.setObjectName("myconverterbutton");
converterButton.setText('Choose file to convert');
converterButton.addEventListener('clicked', () => {


  if(converted == true) {
    console.log('Conversion done!');
    //rootLayout.addWidget(labelConverted);
    converterButton.setText('File converted!');
    converterButton.setIcon(new QIcon(logo));
  }
  else console.log('Conversion failed!')
});


//converted label
const labelConverted = new QLabel();
labelConverted.setObjectName("mylabelConverted");
//converterButton.setIcon(new QIcon(logo));
labelConverted.setText('File converted!');

rootLayout.addWidget(settingsButton);
rootLayout.addWidget(converterButton);

win.setCentralWidget(centralWidget);
win.setStyleSheet(
  `
    #myroot {
      background-color: #009688;
      height: 250px;
      width: 300px;
      align-items: 'center';
      justify-content: 'center';
    }
    #mylabel {
      font-size: 16px;
      font-weight: bold;
      padding: 1;
    }
    #myconverterbutton {
	  font-size: 20px;
      height: 150px;
      width: 220px;
     }
	#mysettingsbutton {
		font-size: 15px;
		height: 40px;
		width: 100px;
	   }
  `
);
//win.show();

(global as any).win = win;






















//entry point of application
start()


async function start() {
  //login
  token = await login()
  //extract access_token from json response
  /*access_token = token.access_token

  //get specific document by ID
  orderBOM = await getSpecificDocument("05642718-f3e8-4e60-a427-18f3e88e60dc")
  //get columns of orderBOM
  orderBOMcolumns = orderBOM.columns
  //get cells of orderBOM
  orderBOMcells = orderBOM.cells

  createExcel()
*/
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