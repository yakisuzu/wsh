// TODO setting format
// TODO show line no
/**
 * @constructor
 */
function Logger(){
};

(function(p){
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

  p.Level = {};
  for(var key in ob_level_list){
    var val = ob_level_list[key];
    p.Level[val] = key;
  }

  var nu_output_level = p.Level.INFO;

  var ar_output_stock = [];

  function outputPush(nu_level, st_text){
    if(nu_output_level <= nu_level){
      ar_output_stock.push({'level' : nu_level, 'text' : st_text});
    }
  }

  p.trace = function(st_text){
    outputPush(p.Level.TRACE, st_text);
  };
  p.debug = function(st_text){
    outputPush(p.Level.DEBUG, st_text);
  };
  p.info = function(st_text){
    outputPush(p.Level.INFO, st_text);
  };
  p.warn = function(st_text){
    outputPush(p.Level.WARN, st_text);
  };
  p.error = function(st_text){
    outputPush(p.Level.ERROR, st_text);
  };
  p.fatal = function(st_text){
    outputPush(p.Level.FATAL, st_text);
  };

  p.setOutputLevel = function(nu_level){
    nu_output_level = nu_level;
  }

  p.println = function(){
    var st_output_string = '';
    while(true){
      var ob_output = ar_output_stock.shift();
      if(ob_output === undefined){break;}

      st_output_string = st_output_string + '[' + ob_level_list[ob_output.level] + ']' + ob_output.text + '\n';
    }
    if(st_output_string !== ''){
      WScript.Echo(st_output_string);
    }
  };

})(Logger.prototype);

