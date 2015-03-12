var excelAdapter = {};

function checkImport(m){if(!this[m]){WScript.Echo('not import ' + m);WScript.Quit();}}
checkImport('utility');
checkImport('logger');

(function(mod){
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

  (function(p, ut, lo){
    /**
     * @return {Excel}
     */
    function openExcel(){
      var ws_excel;
      try{
        ws_excel = WScript.CreateObject('Excel.Application');
        ws_excel.Visible = false;
        lo.trace(self.msg.excel_start);
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
          lo.trace(self.msg.excel_end);
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
        lo.trace(self.msg.excel_edit_start);

        // repeat arg file
        ut.each(ar_files, function(st_arg){
          // ignore extention at pattern
          if(st_arg.search(/^.+\.xlsx?$/) === -1){
            lo.warn(self.msg.no_support);
            return;
          }

          // execute execl function
          var ws_book;
          try{
            ws_book = ws_excel.Workbooks.Open(st_arg);
            lo.trace(ut.buildMsg(self.msg.excel_book_open, [st_arg]));

            fu_execute(ws_book);
          }catch(e){
            ut.dump(e);
            lo.error(ut.buildMsg(self.msg.error, [st_arg]));
          }finally{
            try{
              if(ws_book !== undefined){
                ws_book.Close(true);
                lo.trace(ut.buildMsg(self.msg.excel_book_close, [st_arg]));
              }
            }catch(e){
              ut.dump(e);
            }
          }
        });
        closeExcel(ws_excel);
      }
    };

    /**
     * @param {Workbook} ws_book
     * @throws e
     */
    p.excelErrorNameDelete = function(ws_book){
      try{
        var ws_names = ws_book.Names;
        var ar_del_name = [];
        lo.trace(ut.buildMsg(self.msg.excel_name_count, [ws_names.Count]));

        for(var nu_name = 1; nu_name <= ws_names.Count; nu_name++){
          var ws_name = ws_names.Item(nu_name);
          ws_name.Visible = true;

          lo.trace(ut.buildMsg(self.msg.excel_name_value, [ws_name.Name, ws_name.Value]));

          if(self.excel_use_ignore_reg){
            ut.each(self.excel_ignore_reg, function(st_ignore){
              // when not found regex, add delete array
              if(ws_name.Value.search(st_ignore) === -1){
                ar_del_name.push(ws_name);
              }
            });

          }else{
            ut.each(self.excel_error_reg, function(st_err){
              // when contains error, add delete array
              if(ws_name.Value.search(st_err) !== -1){
                ar_del_name.push(ws_name);
              }
            });
          }
        }

        // execute error name delete
        lo.trace(ut.buildMsg(self.msg.excel_name_hit_count, [ar_del_name.length]));
        ut.each(ar_del_name, function(ws_del){
          lo.trace(ut.buildMsg(self.msg.excel_name_delete_value, [ws_name.Name, ws_name.Value]));
          ws_del.Delete();
        });
      }catch(e){
        ut.dump(e);
        throw e
      }
    };
  })(
    mod.ExcelAdapter.prototype
    , new utility.Utility()
    , new logger.Logger()
    );
})(excelAdapter);

