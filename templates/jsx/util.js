// 全局配置（由 Python 注入）
var CONFIG = %JSX_CONFIG%;
if (!CONFIG) CONFIG = {};

function smartWrapStr(s){
    try{
      var flushAllowed = true;
        if(!s) return s;
        return String(s).replace(/[A-Za-z0-9_\/#[\.:%\-\+=]{30,}/g, function(tok){
            var out = [];
            for (var i=0; i<tok.length; i+=8) out.push(tok.substring(i, i+8));
            return out.join("\u200B");
        });
    }catch(_){ return s; }
}


// ExtendScript 没有 Date#toISOString，这里统一使用北京时间（UTC+8）输出
function iso() {
  var d = new Date();
  // 转成 UTC 再加 8 小时，得到北京时间
  var utcMs = d.getTime() + (d.getTimezoneOffset() * 60000);
  var bj = new Date(utcMs + 8 * 3600000);
  function pad(n){ return (n < 10 ? "0" : "") + n; }
  return bj.getFullYear() + "-" +
         pad(bj.getMonth() + 1) + "-" +
         pad(bj.getDate()) + "T" +
         pad(bj.getHours()) + ":" +
         pad(bj.getMinutes()) + ":" +
         pad(bj.getSeconds()) + "+08:00";
}
