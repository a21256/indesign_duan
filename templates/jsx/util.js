var CONFIG = %JSX_CONFIG%;
if (!CONFIG) CONFIG = {};
// debug ??? entry.js ????? false
var __DEBUG_WRITE = false;

// JSON helpers (ExtendScript 可能没有内置 JSON 对象)
var __HAS_JSON = (typeof JSON !== "undefined" && JSON && typeof JSON.stringify === "function");
function __jsonStringifySafe(obj){
  if (__HAS_JSON){
    try{ return JSON.stringify(obj); }catch(_){}
  }
  try{ return String(obj); }catch(__){ return ""; }
}
function __jsonParseSafe(str){
  if (typeof JSON !== "undefined" && JSON && typeof JSON.parse === "function"){
    try{ return JSON.parse(str); }catch(_){}
  }
  try{ return eval("(" + str + ")"); }catch(__){ return null; }
}

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
// text style helpers (moved from entry.js)
function __fontInfo(r){
  var fam="NA", sty="NA";
  try{ fam = String(r.appliedFont.name || r.appliedFont.family || r.appliedFont); }catch(_){}
  try{ sty = String(r.fontStyle); }catch(_){}
  var tI="NA", tB="NA";
  try{ tI = String(!!r.trueItalic); }catch(_){}
  try{ tB = String(!!r.trueBold); }catch(_){}
  return "font="+fam+" ; style="+sty+" ; trueItalic="+tI+" ; trueBold="+tB;
}
function findCharStyleCI(doc, name){
  var lower = String(name).toLowerCase();
  var cs = doc.characterStyles;
  for (var i=0;i<cs.length;i++){
    try{
      if (String(cs[i].name).toLowerCase() === lower) return cs[i];
    }catch(_){ }
  }
  return null;
}
function _safeIP(tf){
  try{
    if (tf && tf.isValid) {
      var ip = tf.insertionPoints[-1];
      try { var _t = ip.anchoredObjectSettings; }
      catch(e1){
        try { ip.contents = "\u200B"; } catch(_){}
        try { ip = tf.insertionPoints[-1]; } catch(_){}
      }
      if (ip && ip.isValid) return ip;
    }
  } catch(_){}
  try{
    var story = (tf && tf.isValid) ? tf.parentStory : app.activeDocument.stories[0];
    var ip2 = story.insertionPoints[-1];
    try { var _t2 = ip2.anchoredObjectSettings; }
    catch(e2){
      try { ip2.contents = "\u200B"; } catch(_){}
      try { ip2 = story.insertionPoints[-1]; } catch(__){}
    }
    return ip2;
  }catch(e){ try{ log("[LOG] _safeIP fallback error"); }catch(__){} return null; }
}

function __setItalicSafe(r){
  try {
    var doc = app.activeDocument;
    var cs = findCharStyleCI(doc, "斜体") || doc.characterStyles.itemByName("斜体");
    if (!cs || !cs.isValid) {
      try { cs = doc.characterStyles.add({name:"斜体"}); } catch(e) { try { cs = doc.characterStyles.itemByName("斜体"); } catch(__){} }
    }
    if (cs && cs.isValid) {
      try { cs.fontStyle = "Italic"; } catch(_){}
      try { r.appliedCharacterStyle = cs; return "cs:斜体"; } catch(__){}
    }
  } catch(e){}
  try { r.fontStyle = "Italic"; return "fs:Italic"; } catch(_){}
  try { r.fontStyle = "Oblique"; return "fs:Oblique"; } catch(_){}
  return "noop";
}
function __setBoldSafe(r){
  try {
    var doc = app.activeDocument;
    var cs = findCharStyleCI(doc, "粗体") || doc.characterStyles.itemByName("粗体");
    if (!cs || !cs.isValid) {
      try { cs = doc.characterStyles.add({name:"粗体"}); } catch(e) { try { cs = doc.characterStyles.itemByName("粗体"); } catch(__){} }
    }
    if (cs && cs.isValid) {
      try { cs.fontStyle = "Bold"; } catch(_){}
      try { r.appliedCharacterStyle = cs; return "cs:粗体"; } catch(__){}
    }
  } catch(e){}
  try { r.fontStyle = "Bold"; return "fs:Bold"; } catch(_){}
  try { r.fontStyle = "Semibold"; return "fs:Semibold"; } catch(_){}
  return "noop";
}
// note helpers
function __applyInlineFormattingOnRange(story, startCharIndex, endCharIndex, st){
  try {
    if (endCharIndex <= startCharIndex) return;
    var r = story.characters.itemByRange(startCharIndex, endCharIndex - 1);
    var txt=""; try{ txt = String(r.contents).substr(0,50); }catch(_){}
    log("[I/B/U] range="+startCharIndex+"-"+endCharIndex+" ; flags="+__jsonStringifySafe(st)+" ; txt=\""+txt+"\"");

    try { r.underline = !!st.u; log("[U] set="+ (!!st.u)); } catch(eu){ log("[U][ERR] "+eu); }

    if (st.i) {
      try { var howI = __setItalicSafe(r); log("[I] via " + howI + " ; " + __fontInfo(r)); } catch(ei){ log("[I][ERR] "+ei); }
    }
    if (st.b) {
      try { var howB = __setBoldSafe(r);   log("[B] via " + howB + " ; " + __fontInfo(r)); } catch(eb){ log("[B][ERR] "+eb); }
    }
  } catch(e) {
    log("[IBU][ERR] "+e);
  }
}
function __processNoteMatch(m, ctx){
  // ctx: {story, tf, page, stFlags, pendingNoteId, tableTag, tableWarnTag}
  var story = ctx.story;
  var st = ctx.stFlags || {i:0,b:0,u:0};
  function on(x){ return x>0; }
  if (m[1]) {
    ctx.pendingNoteId = parseInt(m[1], 10);
    return;
  }
  if (m[2]) {
    var noteType = m[2];
    var noteContent = m[3];
    var ip = story.insertionPoints[-1];
    try {
      log("[NOTE] create " + noteType + " id=" + ctx.pendingNoteId + " len=" + (noteContent||"").length);
      if (noteType === "FN") createFootnoteAt(ip, noteContent, ctx.pendingNoteId);
      else createEndnoteAt(ip, noteContent, ctx.pendingNoteId);
    } catch(e){ log("[NOTE][ERR] " + e); }
    ctx.pendingNoteId = null;
    return;
  }
  if (m[4]) {
    // format toggles [[/I]] etc
    var closing = m[4] === "/";
    var flag = m[5];
    if (flag === "I") st.i = closing ? 0 : 1;
    else if (flag === "B") st.b = closing ? 0 : 1;
    else if (flag === "U") st.u = closing ? 0 : 1;
    return;
  }
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
  // 保留传入的 File 对象，即使文件尚不存在，后续 __pushEvent 会按需创建
  var okFile = fileObj || null;
  return { file: okFile, logWrite: !!logWriteFlag };
}
function __pushEvent(eventCtx, level, message){
  if (level === "debug" && !__DEBUG_WRITE) return;
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
// unit helpers
function __logUnitValueFail(msg, err){
  if (__UNITVALUE_FAIL_ONCE) return;
  __UNITVALUE_FAIL_ONCE = true;
  try{ log("[DBG] UnitValue unavailable: " + msg + " err=" + err); }catch(_){}
}
function unitPt(val){
  if (val && typeof val === "object") return val;
  var num = parseFloat(val);
  if (!isFinite(num)) return null;
  if (typeof UnitValue === "function"){
    try{ return new UnitValue(num, "pt"); }catch(e){ __logUnitValueFail("num+pt", e); }
    try{ return new UnitValue(num, "points"); }catch(e2){ __logUnitValueFail("num+points", e2); }
    try{ return new UnitValue(num + " pt"); }catch(e3){ __logUnitValueFail("str pt", e3); }
    try{ return new UnitValue(num + "pt"); }catch(e4){ __logUnitValueFail("strpt", e4); }
  } else {
    __logUnitValueFail("UnitValue undefined", "NA");
  }
  return null;
}
// paragraph/character style helpers
function findParaStyleCI(doc, name){
  function norm(n){ return String(n||"").toLowerCase().replace(/\s+/g,"").replace(/[_-]/g,""); }
  var target = norm(name);
  var ps = doc.paragraphStyles;
  for (var i=0;i<ps.length;i++){
    try{ if (norm(ps[i].name) === target) return ps[i]; }catch(_){}
  }
  function scanGroup(g){
    try{
      var arr = g.paragraphStyles;
      for (var i=0;i<arr.length;i++){ try{ if (norm(arr[i].name)===target) return arr[i]; }catch(_){ } }
      var subs = g.paragraphStyleGroups;
      for (var j=0;j<subs.length;j++){ var hit = scanGroup(subs[j]); if (hit) return hit; }
    }catch(_){}
    return null;
  }
  try{
    var groups = doc.paragraphStyleGroups;
    for (var k=0;k<groups.length;k++){ var hit = scanGroup(groups[k]); if (hit) return hit; }
  }catch(_){}
  return null;
}
var ENDNOTE_PS = ENDNOTE_PS || null;
var FOOTNOTE_PS = FOOTNOTE_PS || null;
function ensureFootnoteParaStyle(doc){
  var ps = findParaStyleCI(doc, "footnote");
  if (ps && ps.isValid){ return ps; }
  try { ps = doc.paragraphStyles.itemByName("FootnoteFallback"); } catch(_){}
  if (!ps || !ps.isValid){
    try { ps = doc.paragraphStyles.add({name:"FootnoteFallback"}); } catch(e){
      try { ps = doc.paragraphStyles.itemByName("FootnoteFallback"); } catch(__){}
    }
  }
  try { ps.pointSize   = %FN_FALLBACK_PT%; } catch(_){}
  try { ps.leading     = %FN_FALLBACK_LEAD%; } catch(_){}
  try { ps.spaceBefore = 0; ps.spaceAfter = 0; } catch(_){}
  return ps;
}
function ensureEndnoteParaStyle(doc){
  var ps = findParaStyleCI(doc, "endnote");
  if (ps && ps.isValid){ return ps; }
  try { ps = doc.paragraphStyles.itemByName("FootnoteFallback"); } catch(_){}
  if (!ps || !ps.isValid){
    try { ps = doc.paragraphStyles.add({name:"FootnoteFallback"}); } catch(e){
      try { ps = doc.paragraphStyles.itemByName("FootnoteFallback"); } catch(__){}
    }
  }
  return ps;
}
function createFootnoteAt(ip, content, idForDisplay){
  if(!ip || !ip.isValid) return null;
  var doc = app.activeDocument, story = ip.parentStory;
  var fn = null, ok = false;
  try { fn = story.footnotes.add(LocationOptions.AFTER, ip); ok = (fn && fn.isValid); } catch(e){}
  if (!ok) { try { fn = story.footnotes.add(ip); ok = (fn && fn.isValid); } catch(e){} }
  if (!ok) { try { fn = doc.footnotes.add(ip);   ok = (fn && fn.isValid); } catch(e){} }
  if (!ok) { return null; }
  try {
      var tgtFn = fn.texts[0];
      tgtFn.insertionPoints[-1].contents = content;
  } catch(_){
      try { fn.contents = content; } catch(__){ try { fn.insertionPoints[-1].contents = content; } catch(___) {} }
  }
  try { if (!FOOTNOTE_PS || !FOOTNOTE_PS.isValid) FOOTNOTE_PS = ensureFootnoteParaStyle(doc);
        fn.texts[0].paragraphs.everyItem().appliedParagraphStyle = FOOTNOTE_PS; } catch(_){}
  return fn;
}
function createEndnoteAt(ip, content, idForDisplay){
  if(!ip || !ip.isValid) return null;
  var doc = app.activeDocument, story = ip.parentStory;
  var en = null, ok = false;
  try { if (ip.createEndnote) { en = ip.createEndnote(); ok = (en && en.isValid); } } catch(e){ }
  if (!ok) { try { en = story.endnotes.add(ip); ok = (en && en.isValid); } catch(e){ } }
  if (!ok) { try { en = doc.endnotes.add(ip);   ok = (en && en.isValid); } catch(e){ } }
  if (!ok) { return null; }
  var target = null;
  try { target = en.endnoteText; } catch(_){}
  if (!target || !target.isValid) {
      try { target = en.texts[0]; } catch(_){}
  }
  if (!target || !target.isValid) {
      target = en;
  }
  try {
      target.insertionPoints[-1].contents = content;
  } catch(_){
      try { target.contents = content; } catch(__){}
  }
  try { if (!ENDNOTE_PS || !ENDNOTE_PS.isValid) ENDNOTE_PS = ensureEndnoteParaStyle(app.activeDocument);
        (en.endnoteText || en.texts[0] || en).paragraphs.everyItem().appliedParagraphStyle = ENDNOTE_PS; } catch(_){}
  return en;
}
// table helpers
function __mapAlign(h){ if(h=="center") return Justification.CENTER_ALIGN; if(h=="right") return Justification.RIGHT_ALIGN; return Justification.LEFT_ALIGN; }
function __mapVAlign(v){ if(v=="bottom") return VerticalJustification.BOTTOM_ALIGN; if(v=="center"||v=="middle") return VerticalJustification.CENTER_ALIGN; return VerticalJustification.TOP_ALIGN; }
function __applyTableBorders(tbl, opts, warnTag){
  try{
    opts = opts || {};
    var strokeCol = opts.strokeColor || "Black";
    var outerOn = (opts.outerOn !== false);
    var innerHOn = (opts.innerHOn !== false);
    var innerVOn = (opts.innerVOn !== false);
    var outerWeight = (typeof opts.outerWeight === "number") ? opts.outerWeight : 0.5;
    try{
        var allCells = tbl.cells.everyItem();
        var cellInset = (typeof opts.cellInset === "number") ? opts.cellInset : 0;
        if (cellInset){
            allCells.topInset    = cellInset;
            allCells.bottomInset = cellInset;
            allCells.leftInset   = cellInset;
            allCells.rightInset  = cellInset;
        }
    }catch(_){}

    if(!innerHOn){
        for(var r=0; r<tbl.rows.length-1; r++){
            try{
                var cells = tbl.rows[r].cells.everyItem();
                cells.bottomStrokeWeight = 0;
                cells.bottomEdgeStrokeWeight = 0;
            }catch(_){ }
        }
    }
    if(!innerVOn){
        for(var c=0; c<tbl.columns.length-1; c++){
            try{
                var cc = tbl.columns[c].cells.everyItem();
                cc.rightStrokeWeight = 0;
                cc.rightEdgeStrokeWeight = 0;
            }catch(_){ }
        }
    }

    var topRow    = tbl.rows[0];
    var bottomRow = tbl.rows[tbl.rows.length-1];
    var leftCol   = tbl.columns[0];
    var rightCol  = tbl.columns[tbl.columns.length-1];

    if(outerOn){
        try{
            var tr = topRow.cells.everyItem();
            tr.topStrokeWeight = outerWeight;
            tr.topEdgeStrokeWeight = outerWeight;
            tr.topStrokeColor = strokeCol;
            tr.topEdgeStrokeColor = strokeCol;
        }catch(_){}
        try{
            var br = bottomRow.cells.everyItem();
            br.bottomStrokeWeight = outerWeight;
            br.bottomEdgeStrokeWeight = outerWeight;
            br.bottomStrokeColor = strokeCol;
            br.bottomEdgeStrokeColor = strokeCol;
        }catch(_){}
        try{
            var lc = leftCol.cells.everyItem();
            lc.leftStrokeWeight = outerWeight;
            lc.leftEdgeStrokeWeight = outerWeight;
            lc.leftStrokeColor = strokeCol;
            lc.leftEdgeStrokeColor = strokeCol;
        }catch(_){}
        try{
            var rc = rightCol.cells.everyItem();
            rc.rightStrokeWeight = outerWeight;
            rc.rightEdgeStrokeWeight = outerWeight;
            rc.rightStrokeColor = strokeCol;
            rc.rightEdgeStrokeColor = strokeCol;
        }catch(_){}
    }else{
        try{
            var tr0 = topRow.cells.everyItem();
            tr0.topStrokeWeight = 0;
            tr0.topEdgeStrokeWeight = 0;
        }catch(_){}
        try{
            var br0 = bottomRow.cells.everyItem();
            br0.bottomStrokeWeight = 0;
            br0.bottomEdgeStrokeWeight = 0;
        }catch(_){}
        try{
            var lc0 = leftCol.cells.everyItem();
            lc0.leftStrokeWeight = 0;
            lc0.leftEdgeStrokeWeight = 0;
        }catch(_){}
        try{
            var rc0 = rightCol.cells.everyItem();
            rc0.rightStrokeWeight = 0;
            rc0.rightEdgeStrokeWeight = 0;
        }catch(_){}
    }

    if(opts.headerBoldBorder && tbl.headerRowCount>0){
        try{
            var w = (typeof opts.headerBorderWeight === "number") ? opts.headerBorderWeight : (outerWeight*1.2);
            for(var rr=0; rr<Math.min(tbl.headerRowCount, tbl.rows.length); rr++){
                var row = tbl.rows[rr];
                var cells = row.cells.everyItem();
                cells.bottomStrokeWeight = w;
                cells.bottomEdgeStrokeWeight = w;
                cells.bottomStrokeColor  = strokeCol;
                cells.bottomEdgeStrokeColor  = strokeCol;
            }
        }catch(_){}
    }
  }catch(e){ try{ log((warnTag||"[DBG]") + " applyTableBorders: "+e); }catch(__){} }
}
function __normalizeTableWidth(tbl, warnTag){
  try{
    if (!tbl || !tbl.isValid) return;
    var storyRef = null;
    try{ storyRef = tbl.parentStory; }catch(_){}
    var tf = null;
    if (storyRef && storyRef.isValid && storyRef.textContainers && storyRef.textContainers.length>0){
        for (var i=storyRef.textContainers.length-1; i>=0; i--){
            try{ if (storyRef.textContainers[i].isValid && !storyRef.textContainers[i].overflows){ tf = storyRef.textContainers[i]; break; } }catch(_){}
        }
        if (!tf) tf = storyRef.textContainers[storyRef.textContainers.length-1];
    }
    if (!tf || !tf.isValid){
        try{
            var pg = (app.activeWindow && app.activeWindow.activePage) ? app.activeWindow.activePage : null;
            if (pg){
                var pb = pg.bounds, mp = pg.marginPreferences;
                tf = {
                    geometricBounds: [pb[0]+mp.top, pb[1]+mp.left, pb[2]-mp.bottom, pb[3]-mp.right],
                    textFramePreferences: {
                        leftInset: 0,
                        rightInset: 0,
                        textColumnCount: 1,
                        textColumnFixedWidth: (pb[3]-pb[1])-(mp.left+mp.right)
                    }
                };
            }
        }catch(_){}
    }
    if (!tf) return;
    var gb = tf.geometricBounds;
    var insetL = 0, insetR = 0;
    try{ insetL = tf.textFramePreferences.leftInset  || 0; }catch(_){}
    try{ insetR = tf.textFramePreferences.rightInset || 0; }catch(_){}
    var colW = 0;
    try{
        var tfp = tf.textFramePreferences;
        if (tfp && tfp.textColumnCount>=1 && tfp.textColumnFixedWidth>0) {
            colW = tfp.textColumnFixedWidth;
        }
    }catch(_){}
    if (!colW || colW<=0){
        colW = (gb[3]-gb[1]) - insetL - insetR;
    }
    if (colW>0){
        try{ tbl.preferredWidth = colW; }catch(_){}
        try{ tbl.width = colW; }catch(_){}
        try{
            var C = tbl.columns.length;
            if (C>0){
              var even = colW / C;
              for (var c=0;c<C;c++){
                  try{ tbl.columns[c].width = even; }catch(__){}
              }
            }
        }catch(_){}
    }
  }catch(e){ try{ log((warnTag||"[WARN]") + " _normalizeTableWidth: "+e); }catch(__){} }
}
