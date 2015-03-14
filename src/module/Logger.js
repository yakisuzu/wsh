// TODO setting format
// TODO show line no
var logger = {};

function checkImport(m){if(!this[m]){WScript.Echo('not import ' + m);WScript.Quit();}}
checkImport('utility');

(function(mod, ut){
  var ob_level_list = {
    1 : 'TRACE',
    2 : 'DEBUG',
    3 : 'INFO',
    4 : 'WARN',
    5 : 'ERROR',
    6 : 'FATAL'
  };

  /**
   * 
   */
  mod.Level = (function(){
    var level = {};
    level.ALL = '0';
    ut.each(ob_level_list, function(key, val){
      level[val] = key;
    });
    level.OFF = '9';
    return level;
  })();

  var ar_output_stock = [];
  var ob_output_setting = (function(){
    var o = {};
    o.output_level = mod.Level.INFO;
    o.header = function(nu_level){return '[' + ob_level_list[nu_level] + ']';};
    o.linefeed = '\n';
    o.output = function(st_msg){ut.echo(st_msg);};
    return o;
  })();

  /**
   * @param {Number} nu_level
   * @param {String} st_text
   */
  function outputPush(nu_level, st_text){
    if(ob_output_setting.output_level <= nu_level){
      ar_output_stock.push({'level' : nu_level, 'text' : st_text});
    }
  }

  /**
   * @param {String} st_text
   */
  ut.each(ob_level_list, function(key, val){
    mod[val.toLowerCase()] = function(st_text){
      outputPush(mod.Level[val], st_text);
    };
  });

  /**
   * @param {String} st_msg
   * @param {Array<Object>} st_args
   */
  ut.each(ob_level_list, function(key, val){
    mod[val.toLowerCase() + 'Build'] = function(st_msg, ar_args){
      outputPush(mod.Level[val], ut.buildMsg(st_msg, ar_args));
    };
  });

  /**
   * @param {void}
   */
  mod.print = function(){
    var st_output_string = '';
    while(true){
      var ob_output = ar_output_stock.shift();
      if(!ob_output){break;}
      st_output_string += ob_output_setting.header(ob_output.level) + ob_output.text + ob_output_setting.linefeed;
    }

    if(st_output_string !== ''){
      ob_output_setting.output(st_output_string);
    }
  };

  mod.set = {};
  (function(mods, setting){
    /**
     * @param {Number} nu_level
     */
    mods.outputLevel = function(nu_level){
      setting.output_level = nu_level;
    }

    /**
     * @callback logger.set~fu_header
     * @param {Number} nu_level
     */
    /**
     * @param {logger.set~fu_header} fu_header
     */
    mods.header = function(fu_header){
      setting.header = fu_header;
    };

    /**
     *@param {String} st_linefeed
     */
    mods.linefeed = function(st_linefeed){
      setting.linefeed = st_linefeed;
    };

    /**
     * @callback logger.set~fu_output
     * @param {String} st_msg
     */
    /**
     * @param {logger.set~fu_output} fu_output
     */
    mods.output = function(fu_output){
      setting.output = fu_output;
    };

  })(mod.set, ob_output_setting);

})(
  logger
  , new utility.Utility()
  );

