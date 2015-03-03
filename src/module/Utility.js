var utility = {};

/**
 * @constructor
 */
utility.Utility = function(){
};

(function(p){
  var msg = {};
  msg.not_found = 'class not found';

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
          var val = object[key];
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
        p.echo(msg.not_found);
    }
  };
})(
  utility.Utility.prototype
  );

