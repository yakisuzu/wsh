/**
 * @constructor
 * @param {Msg} msg
 */
function ExcelAdapter(msg){
  // Required
  this.msg = msg || new Msg();
};

(function(p){
  p.self = this;

  p.excel_error_string = [''];

  /**
   * @return {Excel}
   *
   */
  function openExcel(){
    var ws_excel;
    try{
      ws_excel = WScript.CreateObject('Excel.Application');
      ws_excel.Visible = false;
    }catch(e){
      WScript.Echo(e);
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
      }
    }catch(e){
      WScript.Echo(e);
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
      while(true){
        // repeat arg file
        var st_arg = ar_files.shift();
        if(st_arg === undefined){break;}

        // ignore extention at pattern
        if(st_arg.search(/^.+\.xlsx?$/) === -1){
          WScript.Echo(this.msg.no_support);
          continue;
        }

        // execute execl function
        var ws_book;
        try{
          ws_book = ws_excel.Workbooks.Open(st_arg);
          fu_execute(ws_book);
        }catch(e){
          WScript.Echo(this.msg.error + ' ' + st_arg);
        }finally{
          try{
            if(ws_book !== undefined){
              ws_book.Close(true);
            }
          }catch(e){
          }
        }
      }
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
      // var ws_names = this.ws_excel.Workbooks.Item(1).Names;
      var ar_del_name = [];
      for(var nu_name = 0; nu_name < ws_names.Count; nu_name++){
        var ws_name = ws_names.Item(nu_name);
        ws_name.Visible = true;

        var ar_err = JSON.parse(JSON.stringify(p.excel_error_string));
        while(true){
          var st_err = ar_err.shift();
          if(st_err === undefined){break;}

          // when contains error, add delete array
          if(ws_name.Name.contains(st_err)){
            ar_del_name.push(ws_name);
          }
        }
      }

      // execute error name delete
      while(true){
        var ws_del = ar_del_name.shift();
        if(ws_del === undefined){break;}
        ws_del.Delete();
      }
    }catch(e){
      WScript.Echo(e);
      throw e
    }
  };
})(ExcelAdapter.prototype);

