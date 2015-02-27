/**
 * @constructor
 * @param {Msg} msg
 */
function Args(msg){
  // Required
  this.msg = msg || new Msg();
};

(function(p){
  p.self = this;

  /**
   * @return {Array<String>}
   */
  p.getArgs = function(){
    var ws_args = WScript.Arguments;
    if(ws_args.Length === 0){
      WScript.Echo(p.self.msg.no_args);
      WScript.Quit();
    }
    var ar_args = [];
    for(var nu_arg = 0; nu_arg < ws_args.Length; nu_arg++){
      var st_arg = ws_args.Item(nu_arg);
      ar_args.push(st_arg);
    }
    return ar_args;
  };
})(Args.prototype);

