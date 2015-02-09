//TODO: Abstraction extention
//TODO: Abstraction name pattern
//
//cscript C:\shared\wsh\src\deleteErrorExcelName.js C:\shared\wsh\test\test.xls
WScript.Echo('Wait!');
var ws_excel =  WScript.CreateObject('Excel.Application');
try{
  ws_excel.Visible = false;
  var ar_args = get_args();
  while(true){
    var st_arg = ar_args.shift();
    if(st_arg === undefined){
      break;
    }
    edit_excel(ws_excel, st_arg, /^.*$/);
  }
}catch(e){
  throw e
}finally{
  try{
    ws_excel.Quit();
  }catch(e){
  }
}
WScript.Echo('Done!');

function get_args(){
  var ws_args = WScript.Arguments;
  if(ws_args.Length === 0){
    WScript.Echo('Please Drag & drop excel file!');
    WScript.Quit();
  }
  var ar_args = [];
  for(var nu_arg_cnt = 0; nu_arg_cnt < ws_args.Length; nu_arg_cnt++){
    var st_arg = ws_args.Item(nu_arg_cnt);
    ar_args.push(st_arg);
  }
  return ar_args;
}

function edit_excel(ws_excel, st_path, re_ignore){
  if(st_path.search(/^.+\.xlsx?$/) == -1){
    WScript.Echo('Suport xls or xlsx!');
    return;
  }
  var ws_book = ws_excel.Workbooks.Open(st_path);
  try{
    var ws_names = ws_excel.Workbooks.Item(1).Names;
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
    WScript.Echo('Error! ' + st_path);
    throw e
  }finally{
    try{
      ws_book.Close(true);
    }catch(e){
    }
  }
}

