var args = {};

/**
 * @constructor
 */
args.Args = function(){
};

if(!this.logger){WScript.Echo('not import logger');WScript.Quit();}

(function(p, l){
  var msg = {};
  msg.no_args = 'Please Drag & drop excel file!';

  /**
   * @return {Array<String>}
   */
  p.getArgs = function(){
    var ws_args = WScript.Arguments;
    if(ws_args.Length === 0){
      l.info(msg.no_args);
      l.println();
      WScript.Quit();
    }

    var ar_args = [];
    for(var nu_arg = 0; nu_arg < ws_args.Length; nu_arg++){
      var st_arg = ws_args.Item(nu_arg);
      l.trace(st_arg);
      ar_args.push(st_arg);
    }
    return ar_args;
  };
})(
  args.Args.prototype
  , new logger.Logger()
  );

