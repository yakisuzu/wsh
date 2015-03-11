var utility = {};

(function(mod){
  /**
   * @constructor
   */
  mod.Utility = function(){
    this.msg = (function(){
      var m = {};
      m.not_support = '{0} class not support';
      m.dump_object = 'key:{0}, class:{1}, value:{2}';
      m.dump_error = 'name:{0}, message:{1}';
      return m;
    })();
  };

  (function(p){
    /**
     * @param {Object} o
     */
    p.echo = function(o){
      WScript.Echo(o);
    }

    /**
     * @param {Object} o
     * @return {String}
     */
    p.getClass = function(o){
      var st_class =  Object.prototype.toString.apply(o);
      return st_class.replace(/\[object /, '').replace(/\]/, '');
    };

    /**
     * @param {Object} object
     * @param {Function} func
     */
    p.each = function(object, func){
      var self = this;
      var st_class = p.getClass(object);

      switch(st_class){
        case 'Object':
          for(var key in object){
            try{
              var val = object[key];
            }catch(e){
              continue;
            }
            func(key, val);
          }
          break;
        case 'Array':
          for(var i = 0; i < object.length; i++){
            var val = object[i];
            func(val, i);
          }
          break;
        default:
          p.echo(p.buildMsg(self.msg.not_support, [st_class]));
      }
    };

    /**
     * @param {Object} object
     */
    p.dump = function(object){
      var self = this;

      (function dumpR(object, st_pac){
        var st_class = p.getClass(object);
        var st_pac = (st_pac ? st_pac + '.' : '');

        switch(st_class){
          case 'Object':
            p.each(object, function(key, val){
              p.echo(p.buildMsg(self.msg.dump_object, [key, p.getClass(val), val]));
              dumpR(val, st_pac + key);
            });
            break;
          case 'Function':
            dumpR(object.prototype, st_pac + 'prototype');
            break;
          case 'String':
            p.echo(object);
            break;
          case 'Error':
            p.echo(p.buildMsg(self.msg.dump_error, [object.name, object.message]));
            break;
          default:
            p.echo(p.buildMsg(self.msg.not_support, [st_class]));
        }
      })(object, 'root');
    };

    /**
     * @param {String} st_msg
     * @param {Array<String>} ar_args
     */
    p.buildMsg = function(st_msg, ar_args){
      var st_build = st_msg;
      p.each(ar_args, function(val ,idx){
        st_build = st_build.replace('{' + idx + '}', val);
      });
      return st_build;
    };
  })(
    mod.Utility.prototype
    );
})(utility);

