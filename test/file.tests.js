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
var sheet;
var csvConfig = {
	'fSep': ',',
	'trim': true
};

function getParameterByName(name) {
    name = name.replace(/[\[]/, "\\\[").replace(/[\]]/, "\\\]");
    var regex = new RegExp("[\\?&]" + name + "=([^&#]*)"),
        results = regex.exec(location.search);
    return results == null ? "" : decodeURIComponent(results[1].replace(/\+/g, " "));
}

function log(str){
	if(console && console.log){
		console.log(str);
	}
}

function debugReferences(refs,data,indent){
	indent = typeof(indent) === 'number' ? indent : 0;
	var indentText = '';
	for(var x = 0; x < indent; x++){
		indentText = indentText.replace(/./g,' ');
		indentText += '^--';
	}

	for (var p in refs){
		cell = refs[p];
		cellValue = cell.getCalculatedValue();
		expected = data[p];
		if(cellValue+"" !== expected+""){
			log(indentText+p+" : expected: "+expected+"  but got "+cellValue+' ['+cell.formula+']');
			debugReferences(cell.referencedCells,data,indent+1);
		}
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
			var key = Spreadsheet.getColumnNameByIndex(col-1)+row;
			//if(v.length !== 0){
				result[key] = v;
			//}
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
			sheet = Spreadsheet.createSheet();
			sheet.setData(formula);


			var testList = [];
			var debug = false;
			var pos = getParameterByName('pos');
			if(pos.length > 0){
				debug = true;
				if(pos.indexOf(':') !== -1){
					var range = sheet.getCellRange(pos);
					for(var x = 0; x < range.length; x++){
						testList.push(range[x].position.toString());
					}
				}else{
					testList.push(pos);
				}
			}else{
				for(pos in formula){
					testList.push(pos);
				}				
			}

			while(testList.length > 0){
				pos = testList.shift();
				if(typeof(formula[pos]) === "string" && (debug || formula[pos].indexOf('=') === 0)){
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
									if(typeof(cellValue) === "number"){
										cellValue = cellValue.toFixed(2);
									}
									equal(expected,cellValue, expected +' === '+cellValue);
								}
							}else{

								if(expected){
									if(expected.indexOf('%') === expected.length-1){
										expected = expected.substring(0,expected.length-1);
										expected = ''+parseFloat(expected)/100;
									}

									var parts = expected.split(',');
									var isNumber = true;
									//formatted number match
									for(var x = 0; x < parts.length; x++){
										if(isNaN(parseFloat(parts[x]))){
											isNumber = false;
											break;
										}
									}

									if(isNumber){
										expected = parts.join('');
									}

								}

								if (cellValue != null && !isNaN(parseFloat(expected)) && !isNaN(parseFloat(cellValue))){
									var numDecimals = expected.indexOf('.') !== -1 ? expected.split('.')[1].length : 0;
									cellValue = parseFloat(cellValue).toFixed(numDecimals);
								}

								equal(cellValue,expected == null ? "" : expected,cellValue+' === '+expected);
							}

							//Log cell references, usefull for debugging
							if(debug){
								debugReferences(cell.referencedCells,value);
							}
						};
					}(pos,formula,value,sheet)));
				}						
			}

		}catch(e){
			log("Error when getting/parsing ("+valueFile+" or "+formulaFile+") "+e);
		}
	}
};
