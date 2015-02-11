// require
var msg = msg || new Msg();

// class
var Args = function(){
};

Args.prototype = {
  /**
   * out : Array<String>
   */
  get_args : function(){
    var ws_args = WScript.Arguments;
    if(ws_args.Length === 0){
      WScript.Echo(msg.no_args);
      WScript.Quit();
    }
    var ar_args = [];
    for(var nu_arg_cnt = 0; nu_arg_cnt < ws_args.Length; nu_arg_cnt++){
      var st_arg = ws_args.Item(nu_arg_cnt);
      ar_args.push(st_arg);
    }
    return ar_args;
  }
};

