// TODO setting format
// TODO show line no
var logger = {};

/**
 * @constructor
 */
logger.Logger = function(){
};

logger.Level = {};

if(!this.utility){WScript.Echo('not import utility');WScript.Quit();}

(function(p, u){
  var ob_level_list = {
    0 : 'ALL',
    1 : 'TRACE',
    2 : 'DEBUG',
    3 : 'INFO',
    4 : 'WARN',
    5 : 'ERROR',
    6 : 'FATAL',
    9 : 'OFF'
  };

  // for logger.Level setting
  u.each(ob_level_list, function(key, val){
    logger.Level[val] = key;
  });

  var nu_output_level = logger.Level.INFO;

  var ar_output_stock = [];

  /**
   * @param {Number} nu_level
   * @param {String} st_text
   */
  function outputPush(nu_level, st_text){
    if(nu_output_level <= nu_level){
      ar_output_stock.push({'level' : nu_level, 'text' : st_text});
    }
  }

  /**
   * @param {String} st_text
   */
  p.trace = function(st_text){
    outputPush(logger.Level.TRACE, st_text);
  };

  /**
   * @param {String} st_text
   */
  p.debug = function(st_text){
    outputPush(logger.Level.DEBUG, st_text);
  };

  /**
   * @param {String} st_text
   */
  p.info = function(st_text){
    outputPush(logger.Level.INFO, st_text);
  };

  /**
   * @param {String} st_text
   */
  p.warn = function(st_text){
    outputPush(logger.Level.WARN, st_text);
  };

  /**
   * @param {String} st_text
   */
  p.error = function(st_text){
    outputPush(logger.Level.ERROR, st_text);
  };

  /**
   * @param {String} st_text
   */
  p.fatal = function(st_text){
    outputPush(logger.Level.FATAL, st_text);
  };

  /**
   * @param {Number} nu_level
   */
  p.setOutputLevel = function(nu_level){
    nu_output_level = nu_level;
  }

  /**
   *
   */
  p.println = function(){
    var st_output_string = '';
    while(true){
      var ob_output = ar_output_stock.shift();
      if(!ob_output){break;}
      st_output_string += '[' + ob_level_list[ob_output.level] + ']' + ob_output.text + '\n';
    }

    if(st_output_string !== ''){
      u.echo(st_output_string);
    }
  };
})(
  logger.Logger.prototype
  , new utility.Utility()
  );

