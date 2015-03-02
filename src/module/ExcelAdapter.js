/**
 * @constructor
 * @param {Logger} logger
 * @param {Msg} msg
 */
function ExcelAdapter(logger, msg){
  // Required
  this.logger = logger || new Logger();
  this.msg = msg || new Msg();
};

(function(p){
  p.excel_error_string = ['#N/A', '#REF!'];

  /**
   * @param {ExcelAdapter} this
   * @return {Excel}
   */
  function openExcel(t){
    var ws_excel;
    try{
      ws_excel = WScript.CreateObject('Excel.Application');
      ws_excel.Visible = false;
    }catch(e){
      t.logger.error(e);
    }
    return ws_excel;
  };

  /**
   * @param {ExcelAdapter} this
   * @param {Excel} ws_excel
   */
  function closeExcel(t, ws_excel){
    try{
      if(ws_excel !== undefined){
        ws_excel.Quit();
      }
    }catch(e){
      t.logger.error(e);
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
    var ws_excel = openExcel(this);
    if(ws_excel !== undefined){

      // repeat arg file
      for(var nu_arg = 0; nu_arg < ar_files.length; nu_arg++){
        var st_arg = ar_files[nu_arg];

        // ignore extention at pattern
        if(st_arg.search(/^.+\.xlsx?$/) === -1){
          this.logger.warn(this.msg.no_support);
          continue;
        }

        // execute execl function
        var ws_book;
        try{
          ws_book = ws_excel.Workbooks.Open(st_arg);
          fu_execute(ws_book);
        }catch(e){
          this.logger.error(this.msg.error + ' ' + st_arg);
        }finally{
          try{
            if(ws_book !== undefined){
              ws_book.Close(true);
            }
          }catch(e){
            this.logger.error(e);
          }
        }
      }
      closeExcel(this, ws_excel);
    }
  };

  /**
   * @param {Workbook} ws_book
   * @throws e
   */
  p.excelErrorNameDelete = function(ws_book){
    var ar_err = p.excel_error_string;

    try{
      var ws_names = ws_book.Names;
      var ar_del_name = [];
      this.logger.trace('name_count=' + ws_names.Count);

      for(var nu_name = 1; nu_name <= ws_names.Count; nu_name++){
        var ws_name = ws_names.Item(nu_name);
        ws_name.Visible = true;

        this.logger.trace('Name : ' + ws_name.Name + ' , value : ' + ws_name.Value);

        for(var nu_err = 0; nu_err < ar_err.length; nu_err++){
          var st_err = ar_err[nu_err];

          // when contains error, add delete array
          if(ws_name.Value.search(st_err) !== -1){
            ar_del_name.push(ws_name);
          }
        }
      }

      // execute error name delete
      this.logger.trace('hit_name_count=' + ar_del_name.length);
      while(true){
        var ws_del = ar_del_name.shift();
        if(ws_del === undefined){break;}
        ws_del.Delete();
      }
    }catch(e){
      this.logger.error(e);
      throw e
    }
  };
})(ExcelAdapter.prototype);

