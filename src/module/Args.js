var args = {};

function checkImport(m){if(!this[m]){WScript.Echo('not import ' + m);WScript.Quit();}}
checkImport('logger');

(function(mod){
  // TODO self is affected by other instances
  var self;

  /**
   * @constructor
   */
  mod.Args = function(){
    this.msg = (function(){
      var m  ={};
      m.no_args = 'Please Drag & drop excel file!';
      return m;
    })();

    self = this;
  };

  (function(p){
    /**
     * @return {Array<String>}
     */
    p.getArgs = function(){
      var ws_args = WScript.Arguments;
      if(ws_args.Length === 0){
        logger.info(self.msg.no_args);
        logger.print();
        WScript.Quit();
      }

      var ar_args = [];
      for(var nu_arg = 0; nu_arg < ws_args.Length; nu_arg++){
        var st_arg = ws_args.Item(nu_arg);
        logger.trace(st_arg);
        ar_args.push(st_arg);
      }
      return ar_args;
    };
  })(mod.Args.prototype);
})(args);

