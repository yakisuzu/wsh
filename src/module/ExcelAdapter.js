// require
var msg = msg || new Msg();

// class
var ExcelAdapter = function(){
  this.ws_excel;
};

ExcelAdapter.prototype = {
  /**
   * in : Function()
   */
  open : function(run){
    this.ws_excel =  WScript.CreateObject('Excel.Application');
    try{
      this.ws_excel.Visible = false;
      run();
    }catch(e){
      throw e
    }finally{
      try{
        this.ws_excel.Quit();
      }catch(e){
      }
    }
  }
         ,edit_excel : function(st_path, re_ignore){
           if(st_path.search(/^.+\.xlsx?$/) == -1){
             WScript.Echo(msg.no_support);
             return;
           }
           var ws_book = this.ws_excel.Workbooks.Open(st_path);
           try{
             var ws_names = this.ws_excel.Workbooks.Item(1).Names;
             var ar_del_name = [];
             for(var nu_name_cnt = 0; nu_name_cnt < ws_names.Count; nu_name_cnt++){
               var ws_name = ws_names.Item(nu_name_cnt);
               ws_name.Visible = true;
               if(ws_name.Name.search(re_ignore) == -1){
                 ar_del_name.push(ws_name);
               }
             }
             while(true){
               var ws_del = ar_del_name.shift();
               if(ws_del === undefined){
                 break;
               }
               ws_del.Delete();
             }
           }catch(e){
             WScript.Echo(msg.error + ' ' + st_path);
             throw e
           }finally{
             try{
               ws_book.Close(true);
             }catch(e){
             }
           }
         }
};
