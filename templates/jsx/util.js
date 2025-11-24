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


function iso() {
  var d = new Date();
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
// entry logging helpers (shared)
var __EVENT_LINES = [];
function __sanitizeLogMessage(m){
  var txt = String(m == null ? "" : m);
  txt = txt.replace(/[\r\n]+/g, " ").replace(/\t/g, " ");
  return txt;
}
function __initEventLog(fileObj, logWriteFlag){
  __EVENT_LINES = [];
  var okFile = fileObj && fileObj.exists !== undefined ? fileObj : null;
  return { file: okFile, logWrite: !!logWriteFlag };
}
function __pushEvent(eventCtx, level, message){
  if (level === "debug" && !(eventCtx && eventCtx.logWrite)) return;
  var stamp = iso();
  __EVENT_LINES.push(level + "\t" + stamp + "\t" + __sanitizeLogMessage(message));
  var EVENT_FILE = eventCtx && eventCtx.file;
  if (EVENT_FILE) {
    try {
      if (EVENT_FILE.parent && !EVENT_FILE.parent.exists) EVENT_FILE.parent.create();
      EVENT_FILE.encoding = "UTF-8";
      EVENT_FILE.open("a");
      EVENT_FILE.writeln(__EVENT_LINES[__EVENT_LINES.length - 1]);
      EVENT_FILE.close();
    } catch(_){}
  }
}
function __flushEvents(eventCtx){
  var EVENT_FILE = eventCtx && eventCtx.file;
  if (!EVENT_FILE) return;
  try{
    if (EVENT_FILE.parent && !EVENT_FILE.parent.exists) EVENT_FILE.parent.create();
    EVENT_FILE.encoding = "UTF-8";
    EVENT_FILE.open("w");
    for (var i=0; i<__EVENT_LINES.length; i++){
      EVENT_FILE.writeln(__EVENT_LINES[i]);
    }
    EVENT_FILE.close();
  }catch(_){ }
}
