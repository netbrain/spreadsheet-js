(function (Spreadsheet) {
  $.extend(true, window, {
    "Spreadsheet": {
      "Slick": SpreadsheetSlick,
    }
  });

  function SpreadsheetSlick(selector,data){
    var grid;
    var columns = [{
      id:'rowNum',
      field:'rowNum',
      name:'',
      resizable: true,
      width: 50   
    }];
    var dataModel = {
      spreadsheet: Spreadsheet.createSheet(),
      getItem: function(i){
        var row = this.spreadsheet.cells.byRowAndCol[i];
        if(row == null){
          row = {};
        }
        if(row.rowNum == null) row.rowNum = i+1;        
        return row;
      },
      getItemMetadata: function(i){
        var meta = {
          cssClasses: '',
          focusable: true,
          selectable: true,
            columns:  {
              'rowNum':{
                selectable:false
              }
          }
        };

        var row = this.getItem(i);
        if(row != null){
          for(var col = 0; col < row.length; col++){
            var cell = row[col];
            if(cell != null){
              var itemMeta = {};
              meta.columns[cell.position.col] = itemMeta;                
              itemMeta.formatter = function(r,c,v,metadata,row){
                if(v.isCalculated()){
                  return v.toString();
                }else{
                  return '...';
                }
              }
              if(cell.metadata != null){
                if(cell.metadata.colspan != null){
                  itemMeta.colspan = cell.metadata.colspan;
                }  
              }
              
            }
          }
        }
        return meta;
      },
      getLength: function(){
        if(this.spreadsheet != null && 
          this.spreadsheet.cells != null && 
          this.spreadsheet.cells.byRowAndCol != null){
          return this.spreadsheet.cells.byRowAndCol.length;
        }
        return 0
      }
    };
    var options = {
      enableColumnReorder: false,
      editable: true,
      enableAddRow: false,
      enableCellNavigation: true,
      asyncEditorLoading: false,
      autoEdit: false,
      enableAsyncPostRender: true,
      asyncPostRenderDelay: 10,
    };    

    this.setData = function(data){
      for(var pos in data){
         var cellFormula = data[pos].formula;
         var cell = dataModel.spreadsheet.setCellData(pos,cellFormula);
         cell.setMetadata(data[pos].metadata);
         cell.addValueChangeListener(ValueChangeListener)
      }
    }

    this.getGrid = function(){
      return grid;
    }

    this.setData(data);

    var cNum = dataModel.spreadsheet.getColNum();
    for(var c = 0; c < cNum; c++){
      var name = Spreadsheet.getColumnNameByIndex(c);
      columns.push({
        id:name,
        field:c,
        name:name,
        editor:FormulaEditor,
        asyncPostRender:AsyncRenderer
      });
    }
          
    grid = new Slick.Grid(selector, dataModel, columns, options);
    grid.setSelectionModel(new Slick.CellSelectionModel());
    grid.registerPlugin(new Slick.AutoTooltips());

    function AsyncRenderer(cellNode, row, dataContext, colDef) {
      if(dataContext != null){
        var cell = dataContext[colDef.field];
        if(cell != null){
          try{
            $(cellNode).text('...').text(cell.getCalculatedValue());
          }catch(err){
            $(cellNode).text(''+Parser.Error.NAME);
          }
        }
      }
    }

    function FormulaEditor(args) {
      var $input;
      var defaultValue;
      var scope = this;

      this.init = function () {
        $input = $("<input type=text class='editor-formula' />")
            .appendTo(args.container)
            .bind("keydown.nav", function (e) {
              if (e.keyCode === $.ui.keyCode.LEFT || e.keyCode === $.ui.keyCode.RIGHT) {
                e.stopImmediatePropagation();
              }
            })
            .focus()
            .select();
      };

      this.destroy = function () {
        $input.remove();
      };

      this.focus = function () {
        $input.focus();
      };

      this.getValue = function () {
        return $input.val();
      };

      this.setValue = function (val) {
        $input.val(val);
      };

      this.loadValue = function (item) {
        defaultValue = ""
        var cell = item[args.column.field];
        if (cell != null){
          if(cell.isFormula()){     
          defaultValue = cell.formula;
        }else{
          defaultValue = cell.value;
        }
        }
        $input.val(defaultValue);
        $input[0].defaultValue = defaultValue;
        $input.select();      
      };

      this.serializeValue = function () {
        return $input.val();
      };

      this.applyValue = function (item, state) {
        var cell = item[args.column.field];
        if (cell != null){
          cell.setValue(state)
        }else{
          var dataModel = args.grid.getData();
          var pos = args.column.id+args.item.rowNum;
          item[args.column.field] = dataModel.spreadsheet.setCellData(pos,state);
        }     
      };

      this.isValueChanged = function () {
        return (!($input.val() == "" && defaultValue == null)) && ($input.val() != defaultValue);
      };

      this.validate = function () {
        if (args.column.validator) {
          var validationResults = args.column.validator($input.val());
          if (!validationResults.valid) {
            return validationResults;
          }
        }

        return {
          valid: true,
          msg: null
        };
      };

      this.init();
    }

    function ValueChangeListener(e){
      grid.invalidateRow(this.position.row-1);
      grid.render();
    }      
  }   

})(Spreadsheet);