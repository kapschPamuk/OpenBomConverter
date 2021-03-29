import { QMainWindow, QWidget, QLabel, FlexLayout, QPushButton, QIcon} from '@nodegui/nodegui';
import logo from '../greentick.png';
import logoSettings from '../settings.jpg';

var settingsFiles = null
var settingsFile = null
var settingsFileSet = false
var fs = require('fs');

var bodyGlobal = {}
const xlsx = require('xlsx'); // Requiring module to modify excel files

const prompt = require('prompt-sync')(); //user input

const request = require('request'); //http request

// Flag if file was converted
var converted = false;

//needed for delete_cols() function
var crefregex = /(^|[^._A-Z0-9])([$]?)([A-Z]{1,2}|[A-W][A-Z]{2}|X[A-E][A-Z]|XF[A-D])([$]?)([1-9]\d{0,5}|10[0-3]\d{4}|104[0-7]\d{3}|1048[0-4]\d{2}|10485[0-6]\d|104857[0-6])(?![_.\(A-Za-z0-9])/g;

const { QFileDialog } = require("@nodegui/nodegui");
const fileDialog = new QFileDialog();

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
	settings();
});

//Conversion Button
const converterButton = new QPushButton();
converterButton.setObjectName("myconverterbutton");
converterButton.setText('Choose file to convert');
converterButton.addEventListener('clicked', () => {
  fileDialog.exec(); //show menu to choose a file
  const selectedFiles = fileDialog.selectedFiles();
  const file =  selectedFiles.join();
  //console.log(file);
  console.log('Converting ' + file + "...");

if(settingsFileSet == true) convert(file);
else { settingsFile = "conf.ini"; convert(file);}

   //else console.log('No settings found, please set settings first!')


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
win.show();

(global as any).win = win;




//Functions---------------------------------------------------------------------------------------------------------------------------------------

/*
	DELETES `ncols` cols STARTING WITH `start_col`
	- ws         = worksheet object
	- start_col  = starting col (0-indexed) | default 0
	- ncols      = number of cols to delete | default 1
*/
function delete_cols(ws, start_col, ncols) {
	if(!ws) throw new Error("operation expects a worksheet");
	var dense = Array.isArray(ws);
	if(!ncols) ncols = 1;
	if(!start_col) start_col = 0;

	/* extract original range */
	var range = xlsx.utils.decode_range(ws["!ref"]);
	var R = 0, C = 0;

	var formula_cb = function($0, $1, $2, $3, $4, $5) {
		var _R = xlsx.utils.decode_row($5), _C = xlsx.utils.decode_col($3);
		if(_C >= start_col) {
			_C -= ncols;
			if(_C < start_col) return "#REF!";
		}
		return $1+($2=="$" ? $2+$3 : xlsx.utils.encode_col(_C))+($4=="$" ? $4+$5 : xlsx.utils.encode_row(_R));
	};

	var addr, naddr;
	/* move cells and update formulae */
	if(dense) {
	} else {
		for(C = start_col + ncols; C <= range.e.c; ++C) {
			for(R = range.s.r; R <= range.e.r; ++R) {
				addr = xlsx.utils.encode_cell({r:R, c:C});
				naddr = xlsx.utils.encode_cell({r:R, c:C - ncols});
				if(!ws[addr]) { delete ws[naddr]; continue; }
				if(ws[addr].f) ws[addr].f = ws[addr].f.replace(crefregex, formula_cb);
				ws[naddr] = ws[addr];
			}
		}
		for(C = range.e.c; C > range.e.c - ncols; --C) {
			for(R = range.s.r; R <= range.e.r; ++R) {
				addr = xlsx.utils.encode_cell({r:R, c:C});
				delete ws[addr];
			}
		}
		for(C = 0; C < start_col; ++C) {
			for(R = range.s.r; R <= range.e.r; ++R) {
				addr = xlsx.utils.encode_cell({r:R, c:C});
				if(ws[addr] && ws[addr].f) ws[addr].f = ws[addr].f.replace(crefregex, formula_cb);
			}
		}
	}

	/* write new range */
	range.e.c -= ncols;
	if(range.e.c < range.s.c) range.e.c = range.s.c;
	ws["!ref"] = xlsx.utils.encode_range(clamp_range(range));

	/* merge cells */
	if(ws["!merges"]) ws["!merges"].forEach(function(merge, idx) {
		var mergerange;
		switch(typeof merge) {
			case 'string': mergerange = xlsx.utils.decode_range(merge); break;
			case 'object': mergerange = merge; break;
			default: throw new Error("Unexpected merge ref " + merge);
		}
		if(mergerange.s.c >= start_col) {
			mergerange.s.c = Math.max(mergerange.s.c - ncols, start_col);
			if(mergerange.e.c < start_col + ncols) { delete ws["!merges"][idx]; return; }
			mergerange.e.c -= ncols;
			if(mergerange.e.c < mergerange.s.c) { delete ws["!merges"][idx]; return; }
		} else if(mergerange.e.c >= start_col) mergerange.e.c = Math.max(mergerange.e.c - ncols, start_col);
		clamp_range(mergerange);
		ws["!merges"][idx] = mergerange;
	});
	if(ws["!merges"]) ws["!merges"] = ws["!merges"].filter(function(x) { return !!x; });

	/* cols */
	if(ws["!cols"]) ws["!cols"].splice(start_col, ncols);
}

function clamp_range(range) {
	if(range.e.r >= (1<<20)) range.e.r = (1<<20)-1;
	if(range.e.c >= (1<<14)) range.e.c = (1<<14)-1;
	return range;
}

//converts OpenBom file from API call to desired format
function convert(file) {
/* read the excel file */
var workbook = xlsx.readFile(file); // parse the file
var sheetOpenBom = workbook.Sheets[workbook.SheetNames[0]]; // get the first worksheet
var jsonDataOpenBom = xlsx.utils.sheet_to_json(sheetOpenBom, {range: 9 , defval:""}); //convert sheet to json and skip first 9 columns to then convert json to sheet
console.log(jsonDataOpenBom);

//make API call and read that instead of excel 
request('https://jsonplaceholder.typicode.com/users', { json: true }, (err, res, body) => {
  if (err) { return console.log(err); }
 // console.log(body);




 //jsonDataOpenBom =  body;
 console.log(jsonDataOpenBom);


fs.readFile(settingsFile, 'utf8', function(err, jsonStr) {
	if (err) throw err;

	const jsonObj = JSON.parse(jsonStr); //take string input and convert to JSON object

	const map1 = jsonObj[0].Vendor; 
	const map2 = jsonObj[0]["Quantity Required"];
	const map3 = jsonObj[0]["Quantity Gap"];
	const map4 = jsonObj[0]["Part Number"];
	const map5 = jsonObj[0]["Stk,"];
	const map6 = jsonObj[0].Bezeichnung; 
	const map7 = jsonObj[0]["Summe Materialpreis"];
	const map8 = jsonObj[0]["Summe Personal"]; 
	const map9 = jsonObj[0]["Summe Preis"]; 
	const map10 = jsonObj[0]["EK/Stk"]; 
	const map11 = jsonObj[0]["Preis/Stunde"]; 
	const map12 = jsonObj[0]["Hersteller-Artikel#"];
	const map13 = jsonObj[0].Angebotsnummer;
	const map14 = jsonObj[0]["Kreditor#"];
	const map15 = jsonObj[0].Kreditor;
	const map16 = jsonObj[0].Hersteller;
	const map17 = jsonObj[0]["VPE Einheit"];
	const map18 = jsonObj[0]["Listenpreis/VPE"];
	const map19 = jsonObj[0]["Listenpreis/Stk."];
	const map20 = jsonObj[0]["Rabatt %"];
	const map21 = jsonObj[0]["VPE Menge"];
	
	
	/*
	console.log(map1)
	console.log(map2)
	console.log(map3)
	console.log(map4)
	console.log(map5)
	console.log(map6)
	console.log(map7)
	console.log(map8)
	console.log(map9)
	console.log(map10)
	console.log(map11)
	console.log(map12)
	console.log(map13)
	console.log(map14)
	console.log(map15)
	console.log(map16)
	console.log(map17)
	console.log(map18)
	console.log(map19)
	console.log(map20)
	console.log(map21)
	*/
	
	
	var sheetdDataOpenBom = xlsx.utils.json_to_sheet(jsonDataOpenBom, {header:[]}); //convert json to sheet
	//console.log(xlsx.utils.sheet_to_json(sheetdDataOpenBom)); //json data after modification
	xlsx.utils.book_append_sheet(workbook,sheetdDataOpenBom,"Formatted sheet"); //create "Formatted sheet" and fill with content of first worksheet
	var sheetFormatted = workbook.Sheets[workbook.SheetNames[1]]; // get the second worksheet
	//console.log(sheetFormatted); //before changes to "Formatted sheet"
	
	
	//change values of a cell in a sheet
	/* loop through every cell manually */
	var range = xlsx.utils.decode_range(sheetFormatted['!ref']); // get the range
	for(var R = range.s.r; R <= range.e.r; ++R) {
	  for(var C = range.s.c; C <= range.e.c; ++C) {
		/* find the cell object */
		var cellref = xlsx.utils.encode_cell({c:C, r:R}); // construct A1 reference for cell
		if(!sheetFormatted[cellref]) continue; // if cell doesn't exist, move on
		var cell = sheetFormatted[cellref];
	
		/* if the cell is a text cell with the old string, change it */
		if(!(cell.t == 's' || cell.t == 'str')) continue; // skip if cell is not text
		if(cell.v == "Vendor") cell.v = map1; // change the cell value
		if(cell.v == "Quantity Required") cell.v = map2; // change the cell value
		if(cell.v == "Quantity Gap") cell.v = map3; // change the cell value
		if(cell.v == "Part Number") cell.v = map4; // change the cell value
		if(cell.v == "Stk,") cell.v = map5; // change the cell value
		if(cell.v == "Bezeichnung") cell.v = map6; // change the cell value
		if(cell.v == "Summe Materialpreis") cell.v = map7; // change the cell value
		if(cell.v == "Summe Personal") cell.v = map8; // change the cell value
		if(cell.v == "Summe Preis") cell.v = map9; // change the cell value
		if(cell.v == "EK/Stk") cell.v = map10; // change the cell value
		if(cell.v == "Preis/Stunde") cell.v = map11; // change the cell value
		if(cell.v == "Hersteller-Artikel#") cell.v = map12; // change the cell value
		if(cell.v == "Angebotsnummer") cell.v = map13; // change the cell value
		if(cell.v == "Kreditor#") cell.v = map14; // change the cell value
		if(cell.v == "Kreditor") cell.v = map15; // change the cell value
		if(cell.v == "Hersteller") cell.v = map16; // change the cell value
		if(cell.v == "VPE Einheit") cell.v = map17; // change the cell value
		if(cell.v == "Listenpreis/VPE") cell.v = map18; // change the cell value
		if(cell.v == "Listenpreis/Stk.") cell.v = map19; // change the cell value
		if(cell.v == "Rabatt %") cell.v = map20; // change the cell value
		if(cell.v == "VPE Menge") cell.v = map21; // change the cell value
	  }
	}
	
	
	/* order columns according to Einkaufsliste*/
	var sheetOpenBom2 = workbook.Sheets[workbook.SheetNames[1]]; // get the second worksheet
	var jsonDataOpenBom2 = xlsx.utils.sheet_to_json(sheetOpenBom2); //convert sheet to json
	//console.log(jsonDataOpenBom); //json data before modification
	var sheetdDataOpenBom2 = xlsx.utils.json_to_sheet(jsonDataOpenBom2, {header:[map15,'K,Nummer','Incoterm','KOK?',
	'KD,#','CDP BE#','Angebot #',map16,'HOK?',map12,'Kred,#',map6,map5,map19,map20,
	map10,'EK/Ges,','Fracht','TOTAL']}); //convert json to sheet
	
	
	xlsx.utils.book_append_sheet(workbook,sheetdDataOpenBom2,"Für Einkauf"); //create "Für Einkauf" and fill with content with ordered headers
	var sheetOpenBom3 = workbook.Sheets[workbook.SheetNames[2]]; // get the third worksheet
	delete_cols(sheetOpenBom3, 20, 13) //delete 13 columns with unmapped elements, starting from column 20  
	
	xlsx.utils.book_set_sheet_visibility(workbook,1,xlsx.utils.consts.SHEET_HIDDEN); //hide unordered second worksheet 
	
	xlsx.writeFile(workbook,"Einkaufsliste.xlsx") //write whole workbook to "testBom.xlxs"
});

});
converted = true;

}


function settings(){
	fileDialog.exec(); //show menu to choose a file
	settingsFile = fileDialog.selectedFiles().join(); //return absolute path of selected file as string

	if (settingsFile==""){setTimeout(settings, 2000);console.log("Please choose settings")}
	else settingsFileSet = true
	
	if(settingsFileSet)	console.log('Settings configured, now choose the file');
}