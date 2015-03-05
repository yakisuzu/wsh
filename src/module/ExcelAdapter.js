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
      m.error = 'Error!';
      return m;
    })();
  };

  (function(p, u, l){
    p.excel_error_string = ['#N/A', '#REF!'];

    /**
     * @return {Excel}
     */
    function openExcel(){
      var ws_excel;
      try{
        ws_excel = WScript.CreateObject('Excel.Application');
        ws_excel.Visible = false;
        l.trace('Start up excel');
      }catch(e){
        l.error(e);
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
          l.trace('Quit excel');
        }
      }catch(e){
        l.error(e);
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
      var ws_excel = openExcel();
      if(ws_excel !== undefined){
        l.trace('edit start');

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
            l.trace('open book ' + st_arg);

            fu_execute(ws_book);
          }catch(e){
            l.error(self.msg.error + ' ' + st_arg);
          }finally{
            try{
              if(ws_book !== undefined){
                ws_book.Close(true);
                l.trace('close book ' + st_arg);
              }
            }catch(e){
              l.error(e);
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
        l.trace('name count=' + ws_names.Count);

        for(var nu_name = 1; nu_name <= ws_names.Count; nu_name++){
          var ws_name = ws_names.Item(nu_name);
          ws_name.Visible = true;

          l.trace('Name : ' + ws_name.Name + ', Value : ' + ws_name.Value);

          u.each(p.excel_error_string, function(st_err){
            // when contains error, add delete array
            if(ws_name.Value.search(st_err) !== -1){
              ar_del_name.push(ws_name);
            }
          });
        }

        // execute error name delete
        l.trace('hit name count=' + ar_del_name.length);
        u.each(ar_del_name, function(ws_del){
          ws_del.Delete();
        });
      }catch(e){
        l.error(e);
        throw e
      }
    };
  })(
    mod.ExcelAdapter.prototype
    , new modu.Utility()
    , new modl.Logger()
    );
})(excelAdapter, utility, logger);

