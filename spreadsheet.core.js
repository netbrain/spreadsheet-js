(function( Spreadsheet, $, undefined ) {
    //Private Properties
    var charArray = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z'];
    var indexArray = {A:0,B:1,C:2,D:3,E:4,F:5,G:6,H:7,I:8,J:9,K:10,L:11,M:12,N:13,O:14,P:15,Q:16,R:17,S:18,T:19,U:20,V:21,W:22,X:23,Y:24,Z:25};

    //Public Property
    //this.public = "public";

    //Public Methods
    Spreadsheet.createSheet = function(name){
        return new Sheet(name);
    };

    Spreadsheet.parsePosition = function(input){

        if(typeof input !== "string"){
            throw "expected string as input"
        }

        var matches = input.match(/((.+)!)?\$?([A-Z]+)\$?([0-9]+)(\:\$?([A-Z]+)\$?([0-9]+))?/);
        var sheet = matches[2];
        var col1 = matches[3];
        var row1 = matches[4];
        var col2 = matches[6];
        var row2 = matches[7];

        if(!(row2 && col2)){
            return new Position(col1,row1,sheet);
        }else{
            return new PositionRange(
                new Position(col1,row1,sheet),
                new Position(col2,row2,sheet)
            );
        }

        return
    }

    Spreadsheet.getColumnNameByIndex = function(i) {
        var output = '';
        if(i > charArray.length-1){
            output += this.getColumnNameByIndex(Math.floor(i/charArray.length-1))
            output += charArray[i%charArray.length]
        }else{
            output = charArray[i];
        }
        return output;
    };

    Spreadsheet.getColumnIndexByName = function(name) {
        var index = -1;
        var charNum = charArray.length;
        var cArray = name.split('').reverse();
        for(var x = 0; x < cArray.length; x++){
            var cChar = cArray[x];
            var cIndex = indexArray[cChar];
            index += (Math.pow(charNum,x)*(cIndex+1))
        }
        return index+1;
    };


    //Private Methods
    function Sheet(name) {
        this.name = name;
        this.cells = {
            byPosition:{},
            byRowAndCol:[]
        };

        this.formulaParser = Parser.newInstance(this.cells.byPosition);

        this.setData = function(data){
            for (var pos in data){
                this.setCellData(pos,data[pos]);
            }
        };

        this.setCellData = function(pos,value){

            if(typeof(pos) === 'string'){
                pos = Spreadsheet.parsePosition(pos);
            }

            if (!(pos instanceof Position)){
                throw "Illegal argument";
            }

            pos.sheet = this;

            var cellsP = this.cells.byPosition;
            var cellsRC = this.cells.byRowAndCol;

            if(cellsP[pos] === undefined){
                var cell = new Cell(pos,value);
                cellsP[pos] = cell;
                if(cellsRC[pos.row-1] === undefined){
                    cellsRC[pos.row-1] = [];
                }
                cellsRC[pos.row-1][pos.getColumnIndex()-1] = cell;
            }else{
                cellsP[pos].setValue(value);
            }

            return cellsP[pos];
        };

        this.getCellRange = function (range){
            if(typeof range === "string"){
                range = Spreadsheet.parsePosition(range);
            }
            if(!(range instanceof PositionRange)){
                throw "Illegal argument";
            }

            var cellRange = [];
            var p1row = range.from.row;
            var p2row = range.to.row;
            var p1col = range.from.getColumnIndex();
            var p2col = range.to.getColumnIndex();
            var minr = Math.min(p1row,p2row);
            var maxr = Math.max(p1row,p2row);
            var minc = Math.min(p1col,p2col);
            var maxc = Math.max(p1col,p2col);

            for(var row = minr; row <= maxr; row++){
                var rowArray = [];
                for(var col = minc; col <= maxc; col++){
                    rowArray.push(this.getCellByRowAndCol(row,col));
                }
                cellRange.push(rowArray);
            }
            return cellRange;
        };

        this.getCellRangeValues = function(pos1,pos2){
            var values = [];
            var cells = this.getCellRange(pos1,pos2);
            for (var x = 0 ; x < cells.length; x++){
                for (var y = 0 ; y < cells[x].length; y++){
                    if(cells[x][y] !== undefined){
                        values.push(cells[x][y].value);
                    }
                }
            }
            return values;
        }

        this.getCell = function(pos){
            if(typeof(pos) === 'string'){
                pos = Spreadsheet.parsePosition(pos);
            }

            if (!(pos instanceof Position)){
                throw "Illegal argument";
            }

            return this.cells.byPosition[pos];
        };

        this.getCellByRowAndCol = function(r,c){
            if(!(typeof(r) === 'number' && typeof(c) === 'number')){
                throw "Illegal argument";
            }
            var cells = this.cells.byRowAndCol;
            r--;
            c--;
            if(r in cells && c in cells[r]){
                return this.cells.byRowAndCol[r][c];    
            }
        };

        this.getRowNum = function(){
            return this.cells.byRowAndCol.length;
        }

        this.getColNum = function(){
            var max = 0;
            var rows = this.cells.byRowAndCol;
            for (var row = 0; row < rows.length; row++){
                var rowLen = rows[row] != null ? rows[row].length : 0; 
                if(rowLen > max) max = rowLen;
            }
            return max;
        }

        this.toString = function(){
            if(this.name == null) return '';
            return this.name;
        }
    }

    function Cell(position,value){

        this.addReferingCell = function(cell){
            var alreadyAdded = false;
            for(var x = 0; x < this.referingCells.length; x++){
                if(this.referingCells[x] === cell){
                    alreadyAdded = true;
                    break;
                }   
            }
            if(!alreadyAdded){
                this.referingCells.push(cell);
            }            
        };

        this.addReferencedCell = function(cell){
            var alreadyAdded = false;
            for(var x = 0; x < this.referencedCells.length; x++){
                if(this.referencedCells[x] === cell){
                    alreadyAdded = true;
                    break;
                }   
            }
            if(!alreadyAdded){
                this.referencedCells.push(cell);
                cell.addValueChangeListener(this.referenceValueChange,this);
            }
        };
        
        this.addValueChangeListener = function(handler,context){
            if(typeof(handler) === 'function'){
                var alreadyAdded = false;
                for(var x = 0; x < this.valueListeners.length; x++){
                    var l = this.valueListeners[x];
                    if(l.handler === handler){
                        alreadyAdded = true;
                        break;
                    }   
                }
                if(!alreadyAdded){
                    this.valueListeners.push({
                        handler:handler,
                        context:context,
                    })
                }
            }
        };

        this.setValue = function(val){         
            var oldFormula = this.formula;
            this.value = val;
            this.formula = val;
            this.triggerValueChangeEvent(this,oldFormula,this.formula);
        }

        this.triggerValueChangeEvent = function(cell,from,to) {
            for (var x = 0; x < this.valueListeners.length; x++) {
                var l = this.valueListeners[x];
                l.handler.call(l.context, {
                    cell: cell,
                    from: from,
                    to: to
                });
            }
        }

        this.recalculate = function(){
            if(this.isFormula()){
                this.value = this.formula;
                this.calculateValue();           
            }         
        }        

        this.referenceValueChange = function(e){
            this.recalculate();
            this.triggerValueChangeEvent(e.cell,e.from,e.to);
        }

        this.isCalculated = function(){
            if(!this.isFormula()) return true;
            return this.value !== this.formula;
        }

        this.calculateValue = function(){
            if(!this.isCalculated()){
                this.value = formulaParser.parse(this.formula,this.position);
                if(typeof(this.value) === 'object'){
                    this.value = this.value.valueOf();
                }              
                if(formulaParser.hasReferences(this.position)){
                    var refs = formulaParser.getReferences(this.position);
                    if(refs != null){
                        for(var x = 0; x < refs.length; x++){
                            var referencedCell = this.position.sheet.getCell(refs[x]);
                            if(referencedCell == null){
                                referencedCell = this.position.sheet.setCellData(refs[x],'');
                            }
                            referencedCell.addReferingCell(this);
                            this.addReferencedCell(referencedCell);

                        }            
                    }
                }               
            }
        }

        this.getCalculatedValue = function(){
            this.calculateValue();
            return this.value;
        }

        this.toString = function(){
            return this.value != null ? ''+this.value : '';
        }
        
        //Returns the primitive calculated value of the cell
        //In it's excel format
        this.valueOf = function(){
            if(!this.isCalculated()){
                this.calculateValue();
            }
            var number = parseFloat(this.value);
            return isNaN(number) ? '"'+this.value+'"' : number;
        }

        this.isFormula = function(){
            if(typeof(this.formula) === 'string'){
                return this.formula.length > 1 && this.formula.indexOf('=') === 0;
            }
            return false
        }

        this.setMetadata = function(metadata){
            this.metadata = metadata;
        }

        if(!(position instanceof Position)){
            throw "Illegal argument";
        }
        var formulaParser = position.sheet.formulaParser;
        this.position = position;
        this.valueListeners = [];
        this.referencedCells = [];
        this.referingCells = [];

        this.setValue(value);

    }

    function Position(col,row,sheet){

        this.sheet = null;
        this.col = null;
        this.row = null;

        this.getColumnIndex = function(){
            return Spreadsheet.getColumnIndexByName(this.col);
        }

        this.isValidPosition = function(){
            return this.col && this.row;
        }

        this.toString = function(){
            return this.col+this.row;
        }

        if(sheet) this.sheet = sheet;
        if(col) this.col = col;
        if(row){
            if(typeof(row) !== "number"){
                row = parseInt(row);
            }
            if(isNaN(row)){
                throw "Expected number as argument";
            }
            this.row = row;
        }
    }

    function PositionRange(from,to){
        if(!(from instanceof Position && to instanceof Position)){
            throw "Illegal arguments";
        }

        this.from = from;
        this.to = to;

    }

}( window.Spreadsheet = window.Spreadsheet || {}, jQuery ));