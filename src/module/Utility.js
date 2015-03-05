var utility = {};

(function(mod){
  /**
   * @constructor
   */
  mod.Utility = function(){
    this.msg = (function(){
      var m = {};
      m.not_found = 'class not found';
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
      switch(p.getClass(object)){
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
            func(val);
          }
          break;
        default:
          p.echo(this.msg.not_found);
      }
    };

    /**
     * @param {Object} object
     */
    p.dump = function(object){
      switch(p.getClass(object)){
        case 'Object':
          p.each(object, function(key, val){
            p.echo('key=' + key + ', class=' + p.getClass(val) + ', value=' + val);
            switch(p.getClass(val)){
              case 'Object':
                p.dump(val);
                break;
              case 'Function':
                break;
              default:
                p.echo(this.msg.not_found);
            }
          });
          break;
        default:
          p.echo(this.msg.not_found);
      }
    };
  })(
    mod.Utility.prototype
    );
})(utility);

