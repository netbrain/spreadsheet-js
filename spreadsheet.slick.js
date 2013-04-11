(function (Spreadsheet) {
  $.extend(true, window, {
    "Spreadsheet": {
      "Slick": SpreadsheetSlick
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
        if(!row){
          row = {};
        }
        if(!row.rowNum) row.rowNum = i+1;
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
        if(row){
          for(var col = 0; col < row.length; col++){
            var cell = row[col];
            if(cell){
              var itemMeta = {};
              meta.columns[cell.position.col] = itemMeta;
              itemMeta.formatter = CellFormatter;
              if(cell.metadata && cell.metadata.colspan){
                itemMeta.colspan = cell.metadata.colspan;
              }
            }
          }
        }
        return meta;
      },
      getLength: function(){
        if(this.spreadsheet &&
          this.spreadsheet.cells &&
          this.spreadsheet.cells.byRowAndCol){
          return this.spreadsheet.cells.byRowAndCol.length;
        }
        return 0;
      },
      setCellData: function(pos,value){
        var cell = this.spreadsheet.setCellData(pos,value);
        cell.addValueChangeListener(ValueChangeListener,cell);
        return cell;
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
      asyncPostRenderDelay: 10
    };

    this.setData = function(data){
      for(var pos in data){
        var formula = typeof(data[pos]) === "object" ? data[pos].formula : data[pos];
        var cell = dataModel.setCellData(pos,formula);
        if(data[pos].metadata){
          cell.setMetadata(data[pos].metadata);
        }
      }
    };

    this.getGrid = function(){
      return grid;
    };

    this.getSpreadsheet = function(){
      return dataModel.spreadsheet;
    };

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
      if(dataContext){
        var cell = dataContext[colDef.field];
        if(cell){
          try{
            $(cellNode).html(cell.getCalculatedValue());
          }catch(err){
            console.log("cell: "+cell.position+'('+cell.formula+') - '+err);
            $(cellNode).text(''+EFP.Error.NAME);
          }
        }
      }
    }

    function CellFormatter(r,c,v,metadata,row){
      if(v.isCalculated()){
        return v.toString();
      }else{
        return '...';
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
        defaultValue = "";
        var cell = item[args.column.field];
        if (cell){
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
        if (cell){
          cell.setValue(state);
        }else{
          var dataModel = args.grid.getData();
          var pos = args.column.id+args.item.rowNum;
          item[args.column.field] = dataModel.setCellData(pos,state);
        }
      };

      this.isValueChanged = function () {
        return (!($input.val() === "" && defaultValue === null)) && ($input.val() !== defaultValue);
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
      grid.invalidateRow(e.cell.position.row-1);
      grid.render();
    }
  }

})(Spreadsheet);