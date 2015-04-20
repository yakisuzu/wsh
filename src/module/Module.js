var module = {};
module.utility = (function(){
  var mod = {};

  /**
   * @param {Object} o
   */
  mod.echo = function(o){
    WScript.Echo(o);
  };

  /**
   * @param {String} st_msg
   * @param {Array<String>} ar_args
   */
  mod.buildMsg = function(st_msg, ar_args){
    var st_build = st_msg;
    for(var i = 0; i < ar_args.length; i++){
      st_build = st_build.replace('{' + i + '}', ar_args[i]);
    }
    return st_build;
  };

  /**
   * @param {String} st_file
   * @param {String} st_module
   */
  mod.checkImport = function(st_file, st_module){
    if(!module[st_module]){
      mod.echo(mod.buildMsg(getMsg().not_import, [st_module, st_file]));
      WScript.Quit();
    }
  };

  /**
   * @param {Object} o
   * @return {String}
   */
  mod.getClass = function(o){
    var st_class =  Object.prototype.toString.apply(o);
    return st_class.replace(/\[object /, '').replace(/\]/, '');
  };

  /**
   * @param {Object} object
   */
  mod.dump = function(object){
    (function dumpR(object, st_pac){
      var st_class = mod.getClass(object);
      var st_pac = (st_pac ? st_pac + '.' : '');

      switch(st_class){
        case 'Object':
          for(var key in object){
            var value = object[key];
            mod.echo(mod.buildMsg(getMsg().dump_object, [st_pac + key, mod.getClass(value), toString(value)]));
            dumpR(value, st_pac + key);
          }
          function toString(v){
            var isCast = ['Function'].join().search(mod.getClass(v)) !== -1;
            return (isCast ? v.toString() : v);
          }
          break;

        case 'Array':
          // TODO coding
          mod.echo(object.toString());
          break;

        case 'Function':
          dumpR(object.prototype, st_pac + 'prototype');
          break;

        case 'Error':
          mod.echo(mod.buildMsg(getMsg().dump_error, [object.name, object.message]));
          break;

        case 'String':
        case 'Number':
          break;

        default:
          mod.echo(mod.buildMsg(getMsg().not_support, [st_class]));
      }
    })(object);
  };

  /**
   * private
   * @return {Object}
   */
  function getMsg(){
    return (function(){
      var m = {};
      m.not_import = '{0} has not been imported into the {1} module';
      m.not_support = '{0} class not support';
      m.dump_object = 'key:{0}, class:{1}, value:{2}';
      m.dump_error = 'name:{0}, message:{1}';
      return m;
    })();
  }

  return mod;
})();

