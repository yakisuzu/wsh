/**
 * @constructor
 * @param {Logger} logger
 * @param {Msg} msg
 */
function Args(logger, msg){
  // Required
  this.logger = logger || new Logger();
  this.msg = msg || new Msg();
};

(function(p){
  /**
   * @return {Array<String>}
   */
  p.getArgs = function(){
    var ws_args = WScript.Arguments;
    if(ws_args.Length === 0){
      this.logger.info(this.msg.no_args);
      this.logger.println();
      WScript.Quit();
    }
    var ar_args = [];
    for(var nu_arg = 0; nu_arg < ws_args.Length; nu_arg++){
      var st_arg = ws_args.Item(nu_arg);
      this.logger.trace(st_arg);
      ar_args.push(st_arg);
    }
    return ar_args;
  };
})(Args.prototype);

