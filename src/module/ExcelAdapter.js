var excelAdapter = {};

function checkImport(m){if(!this[m]){WScript.Echo('not import ' + m);WScript.Quit();}}
checkImport('utility');
checkImport('logger');

(function(mod){
  // TODO self is affected by other instances
  var self;

  /**
   * @constructor
   */
  mod.ExcelAdapter = function(){
    this.msg = (function(){
      var m  ={};
      m.no_support = 'Support xls or xlsx!';
      m.error = 'Error! {0}';
      m.excel_start = 'Start up excel';
      m.excel_end = 'Quit excel';
      m.excel_edit_start = 'edit start';
      m.excel_book_open = 'open book {0}';
      m.excel_book_close = 'close book {0}';
      m.excel_name_count = 'name count={0}';
      m.excel_name_value = 'Name:{0}, Value:{1}';
      m.excel_name_hit_count = 'hit name count={0}';
      m.excel_name_delete_value = 'delete Name:{0}, Value:{1}';
      return m;
    })();

    this.excel_use_ignore_reg = false;
    this.excel_ignore_reg = [];
    this.excel_error_reg = [/#N\/A/, /#REF!/, /[a-z,A-Z]:\\(.+\\)*.+\.xlsx?/, /\[.+\.xlsx?\]/];

    self = this;
  };

  (function(p, ut){
    /**
     * @return {Excel}
     */
    function openExcel(){
      var ws_excel;
      try{
        ws_excel = WScript.CreateObject('Excel.Application');
        ws_excel.Visible = false;
        logger.trace(self.msg.excel_start);
      }catch(e){
        ut.dump(e);
      }
      return ws_excel;
    };

    /**
     * @param {Excel} ws_excel
     */
    function closeExcel(ws_excel){
      try{
        if(ws_excel !== undefined){
          ws_excel.Quit();
          logger.trace(self.msg.excel_end);
        }
      }catch(e){
        ut.dump(e);
      }
    };

    /**
     * @callback ExcelAdapter~fu_execute
     * @param {Workbook} ws_book
     */
    /**
     * @param {Array<String>} ar_files
     * @param {ExcelAdapter~fu_execute} fu_execute
     */
    p.executeExcel = function(ar_files, fu_execute){
      var ws_excel = openExcel();
      if(ws_excel !== undefined){
        logger.trace(self.msg.excel_edit_start);

        // repeat arg file
        for(var i = 0; i < ar_files.length; i++){
          var st_arg = ar_files[i];
          // ignore extention at pattern
          if(st_arg.search(/^.+\.xlsx?$/) === -1){
            logger.warn(self.msg.no_support);
            return;
          }

          // execute execl function
          var ws_book;
          try{
            ws_book = ws_excel.Workbooks.Open(
                /* FileName */ st_arg,
                /* UpdateLinks */ 0,
                /* ReadOnly */ false,
                /* Format */ null,
                /* Password */ null,
                /* WriteResPassword */ null,
                /* IgnoreReadOnlyRecommended */ true
                );
            logger.traceBuild(self.msg.excel_book_open, [st_arg]);

            fu_execute(ws_book);
          }catch(e){
            ut.dump(e);
            logger.errorBuild(self.msg.error, [st_arg]);
          }finally{
            try{
              if(ws_book !== undefined){
                ws_book.Close(true);
                logger.traceBuild(self.msg.excel_book_close, [st_arg]);
              }
            }catch(e){
              ut.dump(e);
            }
          }
        };
        closeExcel(ws_excel);
      }
    };

    /**
     * @callback ExcelAdapter~fu_execute
     * @param {Object<Excel>} ws_item
     */
    /**
     * @param {Object<Excel>} ws
     * @param {ExcelAdapter~fu_execute} fu_execute
     */
    p.eachItem = function(ws, fu_execute){
      for(var nu_ws = 1; nu_ws <= ws.Count; nu_ws++){
        fu_execute(ws.Item(nu_ws));
      }
    };

    /**
     * @callback ExcelAdapter~fu_execute
     * @param {Worksheet} ws_sheet
     */
    /**
     * @param {Workbook} ws_book
     * @param {ExcelAdapter~fu_execute} fu_execute
     */
    p.eachSheet = function(ws_book, fu_execute){
      var ws_sheets = ws_book.Worksheets;
      logger.traceBuild(self.msg.excel_name_count, [ws_sheets.Count]);
      p.eachItem(ws_sheets, function(ws_sheet){
        logger.traceBuild(self.msg.excel_name_count, [ws_sheet.Name]);
        fu_execute(ws_sheet);
      });
    };

    /**
     * @param {Workbook} ws_book
     */
    p.excelErrorNameDelete = function(ws_book){
      var ws_names = ws_book.Names;
      var ar_del_name = [];

      logger.traceBuild(self.msg.excel_name_count, [ws_names.Count]);
      p.eachItem(ws_names, function(ws_name){
        logger.traceBuild(self.msg.excel_name_value, [ws_name.Name, ws_name.Value]);

        ws_name.Visible = true;

        if(self.excel_use_ignore_reg){
          for(var i = 0; i < self.excel_ignore_reg.length; i++){
            var st_ignore = self.excel_ignore_reg[i];

            // when not found regex, add delete array
            if(ws_name.Value.search(st_ignore) === -1){
              ar_del_name.push(ws_name);
              break;
            }
          };

        }else{
          for(var i = 0; i < self.excel_error_reg.length; i++){
            var st_err = self.excel_error_reg[i];

            // when contains error, add delete array
            if(ws_name.Value.search(st_err) !== -1){
              ar_del_name.push(ws_name);
              break;
            }
          };
        }
      });

      // execute error name delete
      logger.traceBuild(self.msg.excel_name_hit_count, [ar_del_name.length]);
      while(ar_del_name.length !== 0){
        var ws_del = ar_del_name.pop();
        logger.traceBuild(self.msg.excel_name_delete_value, [ws_del.Name, ws_del.Value]);
        ws_del.Delete();
      };
    };

    /**
     * @param {Workbook} ws_book
     */
    p.excelErrorFormatDelete = function(ws_book){
      // TODO coding
      p.eachSheet(ws_book, funtion(ws_sheet){
        var ws_fcs = ws_sheet.Cells.FormatConditions;
        p.eachItem(ws_fcs, function(ws_fc){
          // ws_fc.Formula1;
          // ws_fc.Formula2;
        });
      });
    }

  })(
    mod.ExcelAdapter.prototype
    , new utility.Utility()
    );
})(excelAdapter);

