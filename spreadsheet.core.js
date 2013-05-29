(function( Spreadsheet, $, undefined ) {
    //Private Properties
    var charArray = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z'];
    var indexArray = {A:0,B:1,C:2,D:3,E:4,F:5,G:6,H:7,I:8,J:9,K:10,L:11,M:12,N:13,O:14,P:15,Q:16,R:17,S:18,T:19,U:20,V:21,W:22,X:23,Y:24,Z:25};


    //Public properties
    Spreadsheet.DataValidationException = DataValidationException;

    //Public Methods
    Spreadsheet.createSheet = function(name){
        return new Sheet(name);
    };

    Spreadsheet.parsePosition = function(input,S){

        if(typeof input !== "string"){
            throw "expected string as input";
        }

        if(!(typeof(S) === "undefined" || S instanceof Sheet)){
            throw "expected second argument to be of type Sheet";
        }

        var matches = input.match(/((.+)!)?\$?([A-Z]+)\$?([0-9]+)?(\:\$?([A-Z]+)\$?([0-9]+)?)?/);

        if(!matches){
            throw "not a valid position: "+input;
        }

        var sheet = matches[2];
        var col1 = matches[3];
        var row1 = matches[4];
        var col2 = matches[6];
        var row2 = matches[7];

        if(!col2){
            return new Position(col1,row1,sheet);
        }else{
            //Column reference
            if(!!col1 && !row1){
                row1 = 1;
            }

            if(!!col2 && !row2){
                if(S){
                    row2 = S.getRowNum();
                }else{
                    row2 = 65536;
                }
            }

            return new PositionRange(
                new Position(col1,row1,sheet),
                new Position(col2,row2,sheet)
            );
        }
    };

    Spreadsheet.getColumnNameByIndex = function(i) {
        var output = '';
        if(i > charArray.length-1){
            output += this.getColumnNameByIndex(Math.floor(i/charArray.length-1));
            output += charArray[i%charArray.length];
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
            index += (Math.pow(charNum,x)*(cIndex+1));
        }
        return index;
    };


    //Private Methods
    function Sheet(name) {

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
                cellsRC[pos.row-1][pos.getColumnIndex()] = cell;
            }else{
                cellsP[pos].setValue(value);
            }
            return cellsP[pos];
        };

        this.getCellRange = function (range){
            if(typeof range === "string"){
                range = Spreadsheet.parsePosition(range,this);
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
                for(var col = minc; col <= maxc; col++){
                    cellRange.push(this.getCellByRowAndCol(row,col));
                }
            }
            return cellRange;
        };

        this.getCellRangeValues = function(range){
            var values = [];
            var cells = this.getCellRange(range);
            for (var x = 0 ; x < cells.length; x++){
                values.push(cells[x].getCalculatedValue());
            }
            return values;
        };

        this.getCellValue = function(pos){
            return this.getCell(pos).getCalculatedValue();
        };

        this.getCell = function(pos){
            if(typeof(pos) === 'string'){
                pos = Spreadsheet.parsePosition(pos);
            }

            if (!(pos instanceof Position)){
                throw "Illegal argument";
            }

            if(!(pos in this.cells.byPosition)){
                this.setCellData(pos);
            }

            return this.cells.byPosition[pos];
        };

        this.getCellByRowAndCol = function(r,c){
            if(!(typeof(r) === 'number' && typeof(c) === 'number')){
                throw "Illegal argument";
            }
            var cells = this.cells.byRowAndCol;
            r--;
            if(!(r in cells && c in cells[r])){
                var pos = Spreadsheet.getColumnNameByIndex(c)+(r+1);
                this.setCellData(pos);
            }
            return this.cells.byRowAndCol[r][c];
        };

        this.addValueChangeListener = function(rangeOrPos,listener){
            if(typeof rangeOrPos === "string"){
                rangeOrPos = Spreadsheet.parsePosition(rangeOrPos,this);
            }

            var cells;
            if(rangeOrPos instanceof PositionRange){
                cells = this.getCellRange(rangeOrPos);
            }else if(rangeOrPos instanceof Position){
                cells = [this.getCell(rangeOrPos)];
            }else{
                throw "Illegal argument";
            }

            for (var i = cells.length - 1; i >= 0; i--) {
                cells[i].addValueChangeListener(listener);
            }
        };

        this.getRowNum = function(){
            return this.cells.byRowAndCol.length;
        };

        this.getColNum = function(){
            var max = 0;
            var rows = this.cells.byRowAndCol;
            for (var row = 0; row < rows.length; row++){
                var rowLen = rows[row] ? rows[row].length : 0;
                if(rowLen > max) max = rowLen;
            }
            return max;
        };

        this.parsePosition = function(pos){
            return Spreadsheet.parsePosition(pos,this);
        };

        this.toString = function(){
            if(!this.name) {
                return '';
            }
            return this.name;
        };

        this.name = name;
        this.cells = {
            byPosition:{},
            byRowAndCol:[]
        };
        this.formulaParser = EFP.newInstance(this);

    }

    function Cell(position,value){

        this.getId = function(){
            return this.position.toString();
        };

        this.addReferingCell = function(cell){
            if (!(cell.position in this.referingCells)) {
                this.referingCells[cell.position] = cell;
            }
        };

        this.addReferencedCell = function(cell){
            if (!(cell.position in this.referencedCells)) {
                this.referencedCells[cell.position] = cell;
                cell.addValueChangeListener(this);
            }
        };

        this.addValueChangeListener = function(o){
            if(typeof(o) === 'object' && o.cellValueChangeEvent && o.getId){
                if (!(o.getId() in this.valueListeners)) {
                    this.valueListeners[o.getId()] = {
                        handler:o.cellValueChangeEvent,
                        context:o
                    };
                }
            }else if (typeof(o) === 'function'){
                this.valueListeners[Date.now()] = {
                    handler: o,
                    context: this
                };
            }else{
                throw "listener needs to a simple function or an object that implements 'getId' and 'cellValueChangeEvent'";
            }
        };

        this.setValue = function(val,force){
            var oldFormula = this.formula;
            var oldValue = this.value;

            var n = parseFloat(val);
            if (!isNaN(n) && isFinite(val)){
                //convert to number if it actually is a number
                val = n;
            }

            this.value = val;
            this.formula = val;

            if(force !== true && this.hasDataValidation()){
               try{
                    this.dataValidation.validate(this);
                }catch(e){
                    if(e instanceof DataValidationException){
                        this.value = oldValue;
                        this.formula = oldFormula;
                    }
                    throw e;
                }
            }
            this.triggerValueChangeEvent(oldFormula,this.formula);
        };

        this.triggerValueChangeEvent = function(from,to,event) {
            for(var key in this.valueListeners){
                var l = this.valueListeners[key];
                l.handler.call(l.context, {
                    cell: this,
                    from: from,
                    to: to,
                    event: event
                });
            }
        };

        this.recalculate = function(){
            if(this.isFormula()){
                this.value = this.formula;
                this.calculateValue();
            }
        };

        this.cellValueChangeEvent = function(e){
            if(e.cell.getId() in this.referencedCells){
                var from = this.value;
                this.recalculate();
                var to = this.value;
                this.triggerValueChangeEvent(from,to,e);
            }
        };


        this.isCalculated = function(){
            if(!this.isFormula()) return true;
            return this.value !== this.formula;
        };

        this.calculateValue = function(){
            if(!this.isCalculated()){
                this.value = formulaParser.parse(this.formula,this.position);
                if(typeof(this.value) === 'object'){
                    
                    if(this.value instanceof EFP.Ref){
                        if(EFP.fn.isBlank(this.value)){
                            //todo should probably add value formatting
                            //instead of hardcoded values,
                            //so that the value "blank" is interpreted
                            //as zero when the cell is set with a number
                            //format. and when formatted as a string,
                            //it should yield a blank cell.
                            this.value = 0;
                        }
                    }
                    this.value = this.value.valueOf();
                    
                }
                if(formulaParser.hasReferences(this.position)){
                    var refs = formulaParser.getReferences(this.position);
                    if(refs){
                        for(var x = 0; x < refs.length; x++){
                            var referencedCell = this.position.sheet.getCell(refs[x]);
                            if(!referencedCell){
                                referencedCell = this.position.sheet.setCellData(refs[x],'');
                            }
                            referencedCell.addReferingCell(this);
                            this.addReferencedCell(referencedCell);

                        }
                    }
                }
            }
        };

        this.getCalculatedValue = function(){
            this.calculateValue();
            if(this.value == null){
                return '';
            }
            return this.value;
        };

        this.toString = function(){
            return this.value != null ? ''+this.value : '';
        };

        //Returns the primitive calculated value of the cell
        //In it's excel format
        this.valueOf = function(){
            if(!this.isCalculated()){
                this.calculateValue();
            }
            if(this.value == null) return null;
            return typeof this.value === 'number' ? this.value : '"'+this.value+'"';
        };

        this.isFormula = function(){
            if(typeof(this.formula) === 'string'){
                return this.formula.length > 1 && this.formula.indexOf('=') === 0;
            }
            return false;
        };

        this.setMetadata = function(metadata){
            this.metadata = metadata;
        };

        this.setDataValidation = function(type,operator,args,options){
            this.dataValidation = new DataValidation(type,operator,args,options);
        };

        this.hasDataValidation = function(){
            return this.dataValidation !== undefined && this.dataValidation !== null;
        };

        if(!(position instanceof Position)){
            throw "Illegal argument";
        }
        var formulaParser = position.sheet.formulaParser;
        this.position = position;
        this.valueListeners = {};
        this.referencedCells = {};
        this.referingCells = {};

        this.setValue(value);

    }

    function DataValidation(type,operator,args,options){

        var fn = {
            between: function(value,start,end){
                return value >= start && value <= end;
            },
            notBetween:function(value,start,end){
                return !fn.between(value,start,end);
            },
            equal:function(value,formula){
                return value === formula;
            },
            notEqual:function(value,formula){
                return !fn.equal(value,formula);
            },
            greaterThan:function(value,formula){
                return value > formula;
            },
            lessThan:function(value,formula){
                return value < formula;
            },
            greaterThanOrEqual:function(value,formula){
                return value >= formula;
            },
            lessThanOrEqual:function(value,formula){
                return value <= formula;
            }
        };

        var defaultOptions = {
            allowBlank: true,
            showDropDown:true,
            showErrorMessage:true,
            showInputMessage:true,
            promptTitle:null,
            promptText:null,
            errorText:null,
            errorType:null,
            errorTitle:null
        };

        this.showInputMessage = function(){
            return this.options.showInputMessage &&
            this.options.promptText != null &&
            this.options.promptText !== '';
        };

        this.showErrorMessage = function(){
            return this.options.showErrorMessage;
        };

        this.getListItems = function(cell){
            var list = args[0];
            if(list.indexOf(':') !== -1){
                return cell.position.sheet.getCellRangeValues(list);
            }else if(list.indexOf(',') !== -1){
                return list.split(',');
            }else{
                throw "unknown format for list argument";
            }
        }

        this.validate = function(cell){
            var valid;
            var parser = cell.position.sheet.formulaParser;
            var value = cell.getCalculatedValue();

            if(this.options.allowBlank === true && value === ''){
                valid = true;
            }else{
                switch(type){
                    case 'whole':
                    case 'decimal':
                    case 'date':
                    case 'time':
                        if(!(operator in fn)){
                            throw "unkown operator: "+operator;
                        }

                        for(var x = 0; x < 2; x++){
                            if(args[x] != null){
                                args[x] = parser.parse(args[x].toString());
                                if(args[x] != null && typeof(args[x]) === 'object' ){
                                    args[x] = args[x].valueOf();
                                }
                            }
                        }

                        valid = fn[operator](value,args[0],args[1]);
                        break;
                    case 'list':
                        var list = this.getListItems(cell);
                        for(var x = 0; x < list.length; x++){
                            if (list[x] === value){
                                valid = true;
                                break;
                            }
                        }
                        break;
                    case 'textLength':
                        if(!(operator in fn)){
                            throw "unkown operator: "+operator;
                        }
                        valid = fn[operator]((''+value).length,args[0],args[1]);
                        break;
                    case 'custom':
                        valid = parser.parse(args[0]).toBool();
                        break;
                    default:
                        throw "unknown type: "+type;
                }
            }

            if(!valid && this.showErrorMessage()){
                throw new DataValidationException(options.errorText, options.errorTitle, options.errorType);
            }
        };

        if(!(args instanceof Array)){
            throw "Expected array as input argument.";
        }

        if(options == null){
            options = {};
        }

        this.options = {};
        for (var key in defaultOptions){
            if(!(key in options)){
                this.options[key] = defaultOptions[key];
            }else{
                this.options[key] = options[key];
            }
        }

        this.type = type;
        this.operator = operator;
    }

    function DataValidationException(error,title,type){


        var defaults = {
            stop: {
                error: 'The value you entered is not valid.',
                title: 'Stop'
            },
            warning: {
                error: 'The value you entered is not valid. Continue anyway?',
                title: 'Warning'
            },
            information:{
                error: 'The value you entered is not valid.',
                title: 'Information'
            }
        };

        if(type !== 'stop' && type !== 'warning' && type !== 'information'){
            type = 'stop';
        }

        this.error = error != null && error !== '' ? error : defaults[type].error;
        this.title = title != null && title !== '' ? title : defaults[type].title;
        this.type = type; //stop, warning, information

        this.toString = function(){
            return this.error;
        };
    }

    function Position(col,row,sheet){

        this.sheet = null;
        this.col = null;
        this.row = null;

        this.getColumnIndex = function(){
            return Spreadsheet.getColumnIndexByName(this.col);
        };

        this.isValidPosition = function(){
            return this.col && this.row;
        };

        this.toString = function(){
            return this.col+(this.row ? this.row : '');
        };

        if(sheet) this.sheet = sheet;
        if(col) this.col = col;
        if(row){
            if(typeof(row) !== "number"){
                row = parseInt(row,10);
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