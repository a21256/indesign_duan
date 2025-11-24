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
