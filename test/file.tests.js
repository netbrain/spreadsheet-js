/*

This test file allows for using a filetests.json file (which is a basic javascript object key,value pair).
Where the key is the exported excel document values, and value is exported excel document formulas.

ex:

{
	"file.csv":"file-formulas.csv"
}

Configure up the filetests.json and export csv files under the folder csv in order to test exported 
excel documents.

*/
var formulas = {};
var values = {};
var csvConfig = {
		'fSep': ',',
		'trim': true
};

function log(str){
	if(console && console.log){
		console.log(str);
	}
}

function convertFromRowColArrayToSpreadsheetData(data){
	result = {};
	var row = 0;
	while(data.length > 0){
		row++;
		var cols = data.shift();
		var col = 0;
		while(cols.length > 0){
			col++;
			var v = cols.shift();
			if(v.length !== 0){
				result[Spreadsheet.getColumnNameByIndex(col)+row] = v;
			}
		}
	}
	return result;
}

QUnit.begin = function(){
	var filetests;
	//GOOGLE CHROME : --allow-file-access-from-files
	try{
		var xmlhttp=new XMLHttpRequest();
		xmlhttp.open("GET","filetests.json",false);
		xmlhttp.send();
		filetests = JSON.parse(xmlhttp.responseText);
	}catch(e){
		log("Error when fetching filetests.json, not adding any file tests. ("+e+")");
	}

	for(var key in filetests){
		try{
			var valueFile = key;
			var formulaFile = filetests[key];

			xmlhttp.open("GET","./csv/"+valueFile,false);
			xmlhttp.send();
			if(xmlhttp.responseText.length === 0){
				throw valueFile+": No data";
			}
			var csvValues = xmlhttp.responseText.csvToArray(csvConfig);

			xmlhttp.open("GET","./csv/"+formulaFile,false);
			xmlhttp.send();
			if(xmlhttp.responseText.length === 0){
				throw valueFile+": No data";
			}
			var csvFormulas = xmlhttp.responseText.csvToArray(csvConfig);

			var formula = convertFromRowColArrayToSpreadsheetData(csvFormulas);
			var value = convertFromRowColArrayToSpreadsheetData(csvValues);
			var sheet = Spreadsheet.createSheet();
			sheet.setData(formula);

			for(var pos in formula){
				if(typeof(formula[pos]) === "string" && formula[pos].indexOf('=') === 0){
					module(key);
					test(pos+' ( '+formula[pos]+' )', (function(pos,formula,value,sheet){
						return function(){
							var cell = sheet.getCell(pos);
							var cellValue = cell.getCalculatedValue();
							var expected = value[pos];
							if(expected === "0.00" || expected === ""){
								if(cellValue == null){
									ok(true);
								}else{
									equal(expected,cellValue.toFixed(2));
								}
							}else{

								if(expected && expected.indexOf('%') === expected.length-1){
									expected = expected.substring(0,expected.length-1);
									expected = ''+parseFloat(expected)/100;
								}

								if (cellValue != null && !isNaN(expected) && !isNaN(cellValue)){
									var numDecimals = expected.indexOf('.') !== -1 ? expected.split('.')[1].length : 0;
									cellValue = parseFloat(cellValue).toFixed(numDecimals);
								}

								equal(cellValue,expected == null ? "" : expected,cellValue+' === '+expected);
							}
							//Log cell references, usefull for debugging
							/*for (var p in cell.referencedCells){
								cell = sheet.getCell(p);
								cellValue = cell.getCalculatedValue();
								expected = value[p];
								if(cellValue+"" !== expected+""){
									log(p+" : "+cellValue+"  !== "+expected);	
								}
							}*/
						};
					}(pos,formula,value,sheet)));
				}
			}

		}catch(e){
			log("Error when getting/parsing ("+valueFile+" or "+formulaFile+") "+e);
		}


	}
};