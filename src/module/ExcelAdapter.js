var excelAdapter = {};

function checkImport(m){if(!this[m]){WScript.Echo('not import ' + m);WScript.Quit();}}
checkImport('utility');
checkImport('logger');

(function(mod, modu, modl){
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
  };

  (function(p, u, l){
    /**
     * @param {this} self
     * @return {Excel}
     */
    function openExcel(self){
      var ws_excel;
      try{
        ws_excel = WScript.CreateObject('Excel.Application');
        ws_excel.Visible = false;
        l.trace(self.msg.excel_start);
      }catch(e){
        u.dump(e);
      }
      return ws_excel;
    };

    /**
     * @param {this} self
     * @param {Excel} ws_excel
     */
    function closeExcel(self, ws_excel){
      try{
        if(ws_excel !== undefined){
          ws_excel.Quit();
          l.trace(self.msg.excel_end);
        }
      }catch(e){
        u.dump(e);
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
      var self = this;
      var ws_excel = openExcel(self);
      if(ws_excel !== undefined){
        l.trace(self.msg.excel_edit_start);

        // repeat arg file
        u.each(ar_files, function(st_arg){
          // ignore extention at pattern
          if(st_arg.search(/^.+\.xlsx?$/) === -1){
            l.warn(self.msg.no_support);
            return;
          }

          // execute execl function
          var ws_book;
          try{
            ws_book = ws_excel.Workbooks.Open(st_arg);
            l.trace(u.buildMsg(self.msg.excel_book_open, [st_arg]));

            fu_execute(ws_book);
          }catch(e){
            u.dump(e);
            l.error(u.buildMsg(self.msg.error, [st_arg]));
          }finally{
            try{
              if(ws_book !== undefined){
                ws_book.Close(true);
                l.trace(u.buildMsg(self.msg.excel_book_close, [st_arg]));
              }
            }catch(e){
              u.dump(e);
            }
          }
        });
        closeExcel(self, ws_excel);
      }
    };

    /**
     * @param {Workbook} ws_book
     * @throws e
     */
    p.excelErrorNameDelete = function(ws_book){
      var self = this;
      try{
        var ws_names = ws_book.Names;
        var ar_del_name = [];
        l.trace(u.buildMsg(self.msg.excel_name_count, [ws_names.Count]));

        for(var nu_name = 1; nu_name <= ws_names.Count; nu_name++){
          var ws_name = ws_names.Item(nu_name);
          ws_name.Visible = true;

          l.trace(u.buildMsg(self.msg.excel_name_value, [ws_name.Name, ws_name.Value]));

          if(self.excel_use_ignore_reg){
            u.each(self.excel_ignore_reg, function(st_ignore){
              // when not found regex, add delete array
              if(ws_name.Value.search(st_ignore) === -1){
                ar_del_name.push(ws_name);
              }
            });

          }else{
            u.each(self.excel_error_reg, function(st_err){
              // when contains error, add delete array
              if(ws_name.Value.search(st_err) !== -1){
                ar_del_name.push(ws_name);
              }
            });
          }
        }

        // execute error name delete
        l.trace(u.buildMsg(self.msg.excel_name_hit_count, [ar_del_name.length]));
        u.each(ar_del_name, function(ws_del){
          l.trace(u.buildMsg(self.msg.excel_name_delete_value, [ws_name.Name, ws_name.Value]));
          ws_del.Delete();
        });
      }catch(e){
        u.dump(e);
        throw e
      }
    };
  })(
    mod.ExcelAdapter.prototype
    , new modu.Utility()
    , new modl.Logger()
    );
})(excelAdapter, utility, logger);

