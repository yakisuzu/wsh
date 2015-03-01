/**
 * @constructor
 * @param {Msg} msg
 */
function Logger(msg, level){
  // Required
  this.msg = msg || new Msg();
  // FIXME replace level setting
  this.level = level || this.Level.INFO;

  this.outputStock = [];
};

(function(p){
  p.self = this;

  p.Level = {};
  p.Level.OFF = 9;
  p.Level.FATAL = 6;
  p.Level.ERROR = 5;
  p.Level.WARN = 4;
  p.Level.INFO = 3;
  p.Level.DEBUG = 2;
  p.Level.TRACE = 1;
  p.Level.ALL = -1;

  function outpushPush(level, text){
    p.self.outputStock.push({'level' : level, 'text' : text});
  }

  p.fatal = function(text){
    outputPush(p.Level.FATAL, text);
  };
  p.error = function(){};
  p.warn = function(){};
  p.info = function(){};
  p.debug = function(){};
  p.trace = function(){};

  p.build = function(){
    var outputString = '';
    while(true){
      var ob_text = p.self.outputStock.shift();
      if(ob_text === undefined){break;}

      if(p.self.level <= ob_text.level){
        outputString = outputString + ob_text.text + '\n';
      }
    }
    if(outputString !== ''){
      WScript.Echo(outputString);
    }
  };

})(Logger.prototype);

