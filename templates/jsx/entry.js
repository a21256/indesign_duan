    app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;
    var __origScriptUnit = null, __origViewH = null, __origViewV = null;
    try{
        __origScriptUnit = app.scriptPreferences.measurementUnit;
    }catch(_){}
    try{
        __origViewH = app.viewPreferences.horizontalMeasurementUnits;
        __origViewV = app.viewPreferences.verticalMeasurementUnits;
    }catch(_){}
    try{
        app.scriptPreferences.measurementUnit = MeasurementUnits.POINTS;
    }catch(_){}
    try{
        app.viewPreferences.horizontalMeasurementUnits = MeasurementUnits.POINTS;
        app.viewPreferences.verticalMeasurementUnits = MeasurementUnits.POINTS;
    }catch(_){}

    // config + selfcheck
    var EVENT_FILE = File("%EVENT_LOG_PATH%");
    var LOG_WRITE  = (CONFIG && CONFIG.flags && typeof CONFIG.flags.logWrite === "boolean")
                     ? CONFIG.flags.logWrite : %LOG_WRITE%;   // true=log debug; false=only warn/error/info
    var __EVENT_CTX = __initEventLog(EVENT_FILE, LOG_WRITE);
    function __entryLog(tag, msg){
      try{ log("[" + tag + "] " + msg); }catch(_){}
    }
    function __selfCheck(){
      try{
        if (String(EVENT_FILE || "").indexOf("%") >= 0) throw "EVENT_LOG_PATH placeholder not replaced";
        if (String(LOG_WRITE).indexOf("%") >= 0) throw "LOG_WRITE placeholder not replaced";
        var required = ["__ensureLayoutDefault","__imgAddImageAtV2","__imgAddFloatingImage","__tblAddTableHiFi"];
        for (var i=0;i<required.length;i++){
          var n = required[i];
          if (typeof eval(n) !== "function") throw ("missing function: " + n);
        }
      }catch(e){
        __entryLog("ERR","selfcheck failed: " + e);
        throw e;
      }
    }

    try{
      if (EVENT_FILE){
        EVENT_FILE.encoding = "UTF-8";
        EVENT_FILE.open("w");
        EVENT_FILE.writeln("");
        EVENT_FILE.close();
      }
    }catch(_){}

    function info(m){ __pushEvent(__EVENT_CTX, "info", m); }
    function warn(m){ __pushEvent(__EVENT_CTX, "warn", m); }
    function err(m){  __pushEvent(__EVENT_CTX, "error", m); }
    var __LAST_LAYOUT_LOG = null;
    function __logLayoutEvent(message){
      if (!__LAST_LAYOUT_LOG || __LAST_LAYOUT_LOG !== message){
        __LAST_LAYOUT_LOG = message;
        __pushEvent(__EVENT_CTX, "debug", message);
      }
    }
    function log(m){
      if (String(m||"").indexOf("[LAYOUT]") === 0){
        __logLayoutEvent(String(m));
      } else {
        __LAST_LAYOUT_LOG = null;
        __pushEvent(__EVENT_CTX, "debug", m);
    }
    }
    __selfCheck();
    var __flushEventsWrapper = function(){
      __flushEvents(__EVENT_CTX);
    };
    // alias util formatter for compatibility
    function applyInlineFormattingOnRange(story, startCharIndex, endCharIndex, st){
      return __applyInlineFormattingOnRange(story, startCharIndex, endCharIndex, st);
    }
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
      }
      else{
        __logUnitValueFail("UnitValue undefined", "NA");
      }
      return null;
    }

    function _assignColumnWidth(colObj, widthPt, idx){
      if (!colObj || !colObj.isValid) return false;
      var num = parseFloat(widthPt);
      if (!isFinite(num)) num = widthPt;
      var attempts = [];
      attempts.push({label:"unitPt", factory:function(){ return unitPt(num); }});
      if (typeof UnitValue === "function"){
        attempts.push({label:"Unit(pt)", factory:function(){ return new UnitValue(num, "pt"); }});
        attempts.push({label:"Unit(points)", factory:function(){ return new UnitValue(num, "points"); }});
        attempts.push({label:"Unit(str pt)", factory:function(){ return new UnitValue(num + " pt"); }});
        attempts.push({label:"Unit(strpt)", factory:function(){ return new UnitValue(num + "pt"); }});
      }
      attempts.push({label:"number", factory:function(){ return num; }});
      var logs = [];
      for (var i=0; i<attempts.length; i++){
        var attempt = attempts[i];
        var val = null;
        try{
          val = attempt.factory();
        }catch(factoryErr){
          logs.push(attempt.label + ":ctor=" + factoryErr);
          continue;
        }
        if (val === null || val === undefined){
          logs.push(attempt.label + ":null");
          continue;
        }
        try{
          colObj.width = val;
          return true;
        }catch(applyErr){
          logs.push(attempt.label + ":set=" + applyErr);
        }
      }
      try{
        var docRef = app && app.activeDocument;
        var curPageName = "NA";
        var curFrameId = "NA";
        try{
          if (docRef && docRef.selection && docRef.selection.length){
            var sel = docRef.selection[0];
            if (sel && sel.parentTextFrames && sel.parentTextFrames.length){
              var tf = sel.parentTextFrames[0];
              curFrameId = tf && tf.isValid ? tf.id : "NA";
              if (tf && tf.isValid && tf.parentPage && tf.parentPage.isValid){
                curPageName = tf.parentPage.name;
              }
            }
          }
        }catch(__ctx){}
        log("[DBG] width apply failed idx=" + idx + " val=" + widthPt + " page=" + curPageName + " frame=" + curFrameId + " trace=" + logs.join("|"));
      }catch(_){}
      return false;
    }


    if (!String.prototype.trim) {
      String.prototype.trim = function(){ return String(this).replace(/^\s+|\s+$/g, ""); };
    }

    function _trim(x){ 
        return String(x==null?"":x).replace(/^\s+|\s+$/g,""); 
    }

    log("[BOOT] JSX loaded");
    log("[LOG] start");

    var __DEFAULT_LAYOUT = null;
    var __CURRENT_LAYOUT = null;
    var __DEFAULT_INNER_WIDTH = null;
    var __DEFAULT_INNER_HEIGHT = null;
    var __ENABLE_TRAILING_TRIM = false;
    var __UNITVALUE_FAIL_ONCE = false;
    var __ALLOW_IMG_EXT_FALLBACK = (CONFIG && CONFIG.flags && typeof CONFIG.flags.allowImgExtFallback === "boolean")
                                   ? CONFIG.flags.allowImgExtFallback
                                   : (typeof $.global.__ALLOW_IMG_EXT_FALLBACK !== "undefined"
                                      ? !!$.global.__ALLOW_IMG_EXT_FALLBACK : true);
    var __SAFE_PAGE_LIMIT = (CONFIG && CONFIG.flags && typeof CONFIG.flags.safePageLimit === "number" && isFinite(CONFIG.flags.safePageLimit))
                             ? CONFIG.flags.safePageLimit : 2000;
    var __PARA_SEQ = 0;
    var __PROGRESS_TOTAL = %PROGRESS_TOTAL%;
    var __PROGRESS_DONE = 0;
    var __PROGRESS_LAST_PCT = -1;
    var __PROGRESS_LAST_TS = (new Date()).getTime();
    var __PROGRESS_HEARTBEAT_MS = (CONFIG && CONFIG.progress && typeof CONFIG.progress.heartbeatMs === "number" && isFinite(CONFIG.progress.heartbeatMs))
                                  ? CONFIG.progress.heartbeatMs : %PROGRESS_HEARTBEAT%;
    function __progressDetailText(detail){
      if (!detail) return "";
      try{
        if (typeof detail === "string") return detail;
        var parts = [];
        for (var key in detail){
          if (!detail.hasOwnProperty(key)) continue;
          parts.push(key + "=" + detail[key]);
        }
        return parts.join(" ");
      }catch(_){ return ""; }
    }
    function __progressBump(kind, detail, forceLog){
      if (!__PROGRESS_TOTAL || __PROGRESS_TOTAL <= 0) return;
      __PROGRESS_DONE++;
      var doneDisplay = __PROGRESS_DONE;
      if (__PROGRESS_TOTAL > 0){
        doneDisplay = Math.min(__PROGRESS_DONE, __PROGRESS_TOTAL);
      }
      var pct = Math.min(100, Math.floor((doneDisplay * 100) / __PROGRESS_TOTAL));
      var now = (new Date()).getTime();
      var shouldLog = !!forceLog;
      if (!shouldLog && pct !== __PROGRESS_LAST_PCT){
        shouldLog = true;
      }
      if (!shouldLog && (now - __PROGRESS_LAST_TS) >= __PROGRESS_HEARTBEAT_MS){
        shouldLog = true;
      }
      if (shouldLog){
        __PROGRESS_LAST_PCT = pct;
        __PROGRESS_LAST_TS = now;
        var suffix = "";
        var detailText = __progressDetailText(detail);
        if (detailText) suffix = " " + detailText;
        try{
        info("[PROGRESS][" + kind + "] done=" + doneDisplay + "/" + __PROGRESS_TOTAL + " pct=" + pct + suffix);
        }catch(_){}
      }
    }
    function __progressFinalize(detail){
      if (!__PROGRESS_TOTAL || __PROGRESS_TOTAL <= 0) return;
      var suffix = "";
      var detailText = __progressDetailText(detail);
      if (detailText) suffix = " " + detailText;
      var doneDisplay = Math.min(__PROGRESS_DONE, __PROGRESS_TOTAL);
      var pct = Math.min(100, Math.floor((doneDisplay * 100) / __PROGRESS_TOTAL));
      try{
        info("[PROGRESS][COMPLETE] done=" + doneDisplay + "/" + __PROGRESS_TOTAL + " pct=" + pct + suffix);
      }catch(_){}
    }

    function __resetParaSeq(){ __PARA_SEQ = 0; }
    function __nextParaSeq(){ __PARA_SEQ++; return __PARA_SEQ; }
    function __logSkipParagraph(seq, styleName, reason, textSample){
      try{
        var preview = "";
        if (textSample){
          preview = String(textSample).replace(/\s+/g, " ");
          if (preview.length > 80) preview = preview.substring(0, 80) + "...";
        }
        log("[SKIP][PARA " + seq + "] style=" + styleName + " reason=" + reason + (preview ? " text=\"" + preview + "\"" : ""));
      }catch(_){}
    }
    function __recoverAfterParagraph(storyObj, startIdx){
      try{
        if (storyObj && storyObj.isValid && typeof startIdx === "number"){
          var total = storyObj.characters.length;
          if (total > startIdx){
            try{ storyObj.characters.itemByRange(startIdx, total-1).remove(); }catch(_rm){}
          }
        }
      }catch(_){}
      try{
        if (storyObj && storyObj.isValid){
          var ip = storyObj.insertionPoints[-1];
          if (ip && ip.isValid) ip.contents = "\r";
          storyObj.recompose();
        }
      }catch(_r){}
      try{ __LAST_IMG_ANCHOR_IDX = -1; }catch(_){}
    }

    if (typeof curTextFrame === "undefined" && typeof tf !== "undefined") {
      var curTextFrame = tf;
    }

    function __cloneLayoutState(src){
      var out = {};
      if (!src) return out;
      if (src.pageOrientation){
        out.pageOrientation = String(src.pageOrientation).toLowerCase();
      }
      function _num(v){
        if (v === undefined || v === null) return null;
        var n = parseFloat(v);
        return isFinite(n) ? n : null;
      }
      var w = _num(src.pageWidthPt);
      if (w !== null) out.pageWidthPt = w;
      var h = _num(src.pageHeightPt);
      if (h !== null) out.pageHeightPt = h;
      var pmSrc = src.pageMarginsPt;
      if (pmSrc && typeof pmSrc === "object"){
        var pm = {};
        var has = false;
        var keys = ["top","bottom","left","right"];
        for (var i=0;i<keys.length;i++){
          var k = keys[i];
          if (pmSrc.hasOwnProperty(k)){
            var nv = _num(pmSrc[k]);
            if (nv !== null){
              pm[k] = nv;
              has = true;
            }
          }
        }
        if (has) out.pageMarginsPt = pm;
      }
      return out;
    }

    function __layoutsEqual(a, b){
      function _ori(x){ return (x && typeof x === "string") ? String(x).toLowerCase() : ""; }
      function _diff(n1, n2){
        if (n1 === undefined || n1 === null){
          return !(n2 === undefined || n2 === null);
        }
        if (n2 === undefined || n2 === null) return true;
        var v1 = parseFloat(n1), v2 = parseFloat(n2);
        if (!isFinite(v1) || !isFinite(v2)) return false;
        return Math.abs(v1 - v2) > 0.5;
      }
      a = a || {};
      b = b || {};
      if (_ori(a.pageOrientation) !== _ori(b.pageOrientation)) return false;
      if (_diff(a.pageWidthPt, b.pageWidthPt)) return false;
      if (_diff(a.pageHeightPt, b.pageHeightPt)) return false;
      var keys = ["top","bottom","left","right"];
      var am = a.pageMarginsPt || {};
      var bm = b.pageMarginsPt || {};
      for (var i=0;i<keys.length;i++){
        var k = keys[i];
        if (_diff(am[k], bm[k])) return false;
      }
      return true;
    }

    function __createLayoutFrame(layoutState, linkFromFrame, opts){
      opts = opts || {};
      var target = __cloneLayoutState(layoutState);
      try{
        if (!target.pageOrientation && __DEFAULT_LAYOUT && __DEFAULT_LAYOUT.pageOrientation){
          target.pageOrientation = __DEFAULT_LAYOUT.pageOrientation;
        }
        if ((target.pageWidthPt === undefined || target.pageWidthPt === null) && __DEFAULT_LAYOUT){
          target.pageWidthPt = __DEFAULT_LAYOUT.pageWidthPt;
        }
        if ((target.pageHeightPt === undefined || target.pageHeightPt === null) && __DEFAULT_LAYOUT){
          target.pageHeightPt = __DEFAULT_LAYOUT.pageHeightPt;
        }
        if (!target.pageMarginsPt && __DEFAULT_LAYOUT && __DEFAULT_LAYOUT.pageMarginsPt){
          target.pageMarginsPt = __cloneLayoutState({pageMarginsPt:__DEFAULT_LAYOUT.pageMarginsPt}).pageMarginsPt;
        }
        var basePage = (opts.afterPage && opts.afterPage.isValid) ? opts.afterPage : (page && page.isValid ? page : doc.pages[doc.pages.length-1]);
        var newPage = null;
        try{
          try{ doc.allowPageShuffle = true; }catch(_docShuf){}
          if (basePage && basePage.parent && basePage.parent.isValid){
            try{ basePage.parent.allowPageShuffle = true; }catch(_spShuf){}
          }
        }catch(_prep){}
        var forceSpread = !!(opts && opts.forceNewSpread);
        if (forceSpread){
          try{
            var targetSpread = null;
            try{
              var baseSpread = (basePage && basePage.parent && basePage.parent.isValid) ? basePage.parent : null;
              if (baseSpread){
                try{ log("[LAYOUT] base spread pages=" + baseSpread.pages.length + " name=" + baseSpread.name); }catch(__logBase){}
                targetSpread = doc.spreads.add(LocationOptions.AFTER, baseSpread);
              } else {
                targetSpread = doc.spreads.add(LocationOptions.AT_END);
              }
            }catch(__baseInfo){
              try{ targetSpread = doc.spreads.add(LocationOptions.AT_END); }catch(__spreadFallback){}
            }
            if (targetSpread && targetSpread.isValid){
              try{ targetSpread.allowPageShuffle = true; }catch(__spAllow){}
              try{
                while(targetSpread.pages.length > 1){
                  targetSpread.pages[1].remove();
                }
              }catch(__trimSpread){}
              if (targetSpread.pages.length > 0){
                newPage = targetSpread.pages[0];
              } else {
                newPage = targetSpread.pages.add();
              }
            }
            if (!newPage || !newPage.isValid){
              newPage = doc.pages.add(LocationOptions.AT_END);
            }
          }catch(eAddForce){
            try{ newPage = doc.pages.add(LocationOptions.AT_END); }catch(eAddForce2){ newPage = doc.pages.add(); }
          }
        } else {
          try{
            if (basePage && basePage.isValid){
              newPage = doc.pages.add(LocationOptions.AFTER, basePage);
            } else {
              newPage = doc.pages.add(LocationOptions.AT_END);
            }
          }catch(eAdd){
            try{ newPage = doc.pages.add(LocationOptions.AT_END); }catch(eAdd2){ newPage = doc.pages.add(); }
          }
        }
        try{ /* no-op restore to keep shuffle true intentionally */ }catch(_restore){}
        if (!newPage || !newPage.isValid){
          throw new Error("page add failed");
        }
        try{
          var w = target.pageWidthPt, h = target.pageHeightPt;
          if (isFinite(w) && isFinite(h) && w > 0 && h > 0){
            newPage.resize(
              CoordinateSpaces.PASTEBOARD_COORDINATES,
              AnchorPoint.TOP_LEFT_ANCHOR,
              ResizeMethods.REPLACING_CURRENT_DIMENSIONS_WITH,
              [w, h]
            );
          }
        }catch(eResize){ try{ log("[WARN] layout page resize failed: " + eResize); }catch(_){ } }
        try{
          var mp = newPage.marginPreferences;
          var margins = target.pageMarginsPt || {};
          if (mp){
            if (isFinite(margins.top)) mp.top = margins.top;
            if (isFinite(margins.bottom)) mp.bottom = margins.bottom;
            if (isFinite(margins.left)) mp.left = margins.left;
            if (isFinite(margins.right)) mp.right = margins.right;
          }
        }catch(eMargin){ try{ log("[WARN] layout margin apply failed: " + eMargin); }catch(_){ } }
        var newFrame = createTextFrameOnPage(newPage, target);
        try{
          if (newFrame && newFrame.isValid){
            log("[LAYOUT] new frame id=" + newFrame.id + " orient=" + (target.pageOrientation||"") + " page=" + (newPage && newPage.name));
          }
        }catch(_){}
        if (newFrame && newFrame.isValid && linkFromFrame && linkFromFrame.isValid){
          try{ linkFromFrame.nextTextFrame = newFrame; }catch(eLink){ try{ log("[WARN] layout frame link failed: " + eLink); }catch(_){ } }
        }
        return { page: newPage, frame: newFrame };
      }catch(e){ try{ log("[WARN] create layout frame failed: " + e); }catch(_){ } }
      return null;
    }

    function __ensureLayout(targetState){
      try{ log("[LAYOUT] ensure request orient=" + (targetState && targetState.pageOrientation) + " width=" + (targetState && targetState.pageWidthPt) + " height=" + (targetState && targetState.pageHeightPt)); }catch(_){}
      var target = targetState ? __cloneLayoutState(targetState) : __cloneLayoutState(__DEFAULT_LAYOUT);
      if ((target.pageWidthPt === undefined || target.pageWidthPt === null) && __DEFAULT_LAYOUT){
        target.pageWidthPt = __DEFAULT_LAYOUT.pageWidthPt;
      }
      if ((target.pageHeightPt === undefined || target.pageHeightPt === null) && __DEFAULT_LAYOUT){
        target.pageHeightPt = __DEFAULT_LAYOUT.pageHeightPt;
      }
      if (!target.pageMarginsPt && __DEFAULT_LAYOUT && __DEFAULT_LAYOUT.pageMarginsPt){
        target.pageMarginsPt = __cloneLayoutState({pageMarginsPt:__DEFAULT_LAYOUT.pageMarginsPt}).pageMarginsPt;
      }
      if (!__DEFAULT_LAYOUT) __DEFAULT_LAYOUT = __cloneLayoutState(target);
      if (target.pageOrientation === "landscape" && isFinite(target.pageWidthPt) && isFinite(target.pageHeightPt) && target.pageWidthPt < target.pageHeightPt){
        var tmpW = target.pageWidthPt;
        target.pageWidthPt = target.pageHeightPt;
        target.pageHeightPt = tmpW;
      }else if (target.pageOrientation === "portrait" && isFinite(target.pageWidthPt) && isFinite(target.pageHeightPt) && target.pageWidthPt > target.pageHeightPt){
        var tmpH = target.pageHeightPt;
        target.pageHeightPt = target.pageWidthPt;
        target.pageWidthPt = tmpH;
      }
      var prevOrientation = __CURRENT_LAYOUT ? __CURRENT_LAYOUT.pageOrientation : null;
      var needNewSpread = !!(target.pageOrientation && prevOrientation && target.pageOrientation !== prevOrientation);
      if (__layoutsEqual(__CURRENT_LAYOUT, target)){
        try{
          log("[LAYOUT] ensure skip orient=" + (target.pageOrientation||"") + " width=" + target.pageWidthPt + " height=" + target.pageHeightPt);
        }catch(_){}
        try{
          if (target.pageOrientation && __CURRENT_LAYOUT && __CURRENT_LAYOUT.pageOrientation !== target.pageOrientation){
            var __skipPayload = __cloneLayoutState(target);
            __skipPayload.origin = "skip";
            log("[LAYOUT] still skipping due to same state; page=" + (page && page.name) + " spec=" + __jsonStringifySafe(__skipPayload));
          }
        }catch(__skipLog){}
        return;
      }
      var prevFrame = (typeof tf !== "undefined" && tf && tf.isValid) ? tf : null;
      var pkt = __createLayoutFrame(target, prevFrame, {forceNewSpread: needNewSpread});
      if (pkt && pkt.frame && pkt.frame.isValid){
        try{ log("[LAYOUT] ensure apply orient=" + (target.pageOrientation||"") + " page=" + (pkt.page && pkt.page.name) + " frame=" + pkt.frame.id); }catch(_){}
        page = pkt.page;
        tf = pkt.frame;
        story = tf.parentStory;
        curTextFrame = tf;
        __CURRENT_LAYOUT = __cloneLayoutState(target);
        try{ story.recompose(); }catch(_){}
        try{ app.activeDocument.recompose(); }catch(_){}
      }
    }

    function __ensureLayoutDefault(){
      __ensureLayout(__DEFAULT_LAYOUT);
    }

    var IMG_DIRS = (CONFIG && CONFIG.imgDirs && CONFIG.imgDirs.length) ? CONFIG.imgDirs : %IMG_DIRS_JSON%;

function _holderInnerBounds(holder){
  var innerW = 0;
  var innerH = 0;
  try{
    if (!holder || !holder.isValid) return {innerW:0, innerH:0};
    var hb = holder.geometricBounds;
    var inset = holder.textFramePreferences && holder.textFramePreferences.insetSpacing;
    var leftInset = (inset && inset.length>=2)? inset[1] : 0;
    var rightInset = (inset && inset.length>=4)? inset[3] : 0;
    var topInset = (inset && inset.length>=1)? inset[0] : 0;
    var bottomInset = (inset && inset.length>=3)? inset[2] : 0;
    if (hb && hb.length===4){
      innerW = Math.max(0, hb[3]-hb[1] - leftInset - rightInset);
      innerH = Math.max(0, hb[2]-hb[0] - topInset - bottomInset);
    }
    if (innerH <= 0 && holder.parentPage && holder.parentPage.isValid){
      var pageBounds = holder.parentPage.bounds;
      innerH = Math.max(innerH, (pageBounds[2]-pageBounds[0]) - topInset - bottomInset);
    }
  }catch(_){}
  return {innerW:innerW, innerH:innerH};
}


    // text style helpers moved to util.js: __fontInfo, __setItalicSafe, __setBoldSafe

    function _mapAlign(h){ return __mapAlign(h); }
    function _mapVAlign(v){ return __mapVAlign(v); }
    function applyTableBorders(tbl, opts){ return __applyTableBorders(tbl, opts, __tableWarnTag); }
    function _normalizeTableWidth(tbl){ return __normalizeTableWidth(tbl, __tableWarnTag); }


    function addParaWithNotes(story, styleName, raw) {
        var paraSeq = __nextParaSeq();
        var s = app.activeDocument.paragraphStyles.itemByName(styleName);
        try { log("[PARA] style=" + styleName + " len=" + String(raw||"").length); } catch(_){}
        if (!s.isValid) { s = app.activeDocument.paragraphStyles.add({name:styleName}); }

        var text = String(raw).replace(/^\s+|\s+$/g, "");
        if (text.length === 0) return;

        var insertionStart = 0;
        try{ insertionStart = (story && story.isValid) ? story.characters.length : 0; }catch(_){ }

        try{
                var re = /\[{2,}FNI:(\d+)\]{2,}|\[{2,}(FN|EN):(.*?)\]{2,}|\[\[(\/?)(I|B|U)\]\]|\[\[IMG\s+([^\]]+)\]\]|\[\[TABLE\s+(\{[\s\S]*?\})\]\]/g;
                var last = 0, m;
                var st = {i:0, b:0, u:0};
                var noteCtx = {story: story, tf: tf, page: page, stFlags: st, pendingNoteId: null};

                while ((m = re.exec(text)) !== null) {
                    var chunk = text.substring(last, m.index);
                    if (chunk.length) {
                        var startIdx = story.characters.length;
                        story.insertionPoints[-1].contents = chunk;
                        var endIdx   = story.characters.length;
                        __applyInlineFormattingOnRange(story, startIdx, endIdx, {i:(st.i>0), b:(st.b>0), u:(st.u>0)});
                    }
                    try { story.insertionPoints[-1].appliedCharacterStyle = app.activeDocument.characterStyles.itemByName("[None]"); } catch(_){ try { story.insertionPoints[-1].appliedCharacterStyle = app.activeDocument.characterStyles[0]; } catch(__){} }


                    if (m[1] || m[2] || m[4]) {
                        __processNoteMatch(m, noteCtx);
                    } else if (m[6]) {
                try{ log("[IMGDBG] enter [[IMG]] attrs=" + m[6]); }catch(_){}
                var kv = m[6], spec = {};
                try{
                  log('[IMGDBG] enter [[IMG]] lastIdx='
                      + (typeof __LAST_IMG_ANCHOR_IDX==='number'?__LAST_IMG_ANCHOR_IDX:'NA'));
                }catch(_){}
                kv.replace(/(\w+)=['"“”]([^'"”]*)['"”]/g, function(_,k,v){ spec[k]=v; return _; });
                try{
                  var _keys = [];
                  for (var _k in spec){ if (spec.hasOwnProperty(_k)) _keys.push(_k); }
                  log('[IMGDBG] parsed spec keys='+_keys.join(','));
                  log('[IMGDBG] parsed posHref='+ (spec.posHref||'') +' posVref='+ (spec.posVref||'') +' posV='+ (spec.posV||''));
                }catch(_){}

                if (!spec.align) spec.align = "center";
                if (spec.spaceBefore == null) spec.spaceBefore = 0;
                if (spec.spaceAfter  == null) spec.spaceAfter  = 2;
                if (!spec.wrap) spec.wrap = "none"; 

                try{
                  var __ipEnd0 = story.insertionPoints[-1];
                  var __holder0 = (__ipEnd0 && __ipEnd0.isValid && __ipEnd0.parentTextFrames && __ipEnd0.parentTextFrames.length)
                                  ? __ipEnd0.parentTextFrames[0] : null;
                  var __pg0 = (__holder0 && __holder0.isValid) ? __holder0.parentPage : null;
                  var __tfId0 = (typeof tf!=="undefined" && tf && tf.isValid) ? tf.id : "NA";
                  var __cfId0 = (typeof curTextFrame!=="undefined" && curTextFrame && curTextFrame.isValid) ? curTextFrame.id : "NA";
                  log("[IMG-LOC][pre] storyEnd.tf=" + (__holder0?__holder0.id:"NA")
                      + " page=" + (__pg0?__pg0.name:"NA")
                      + " ; tf=" + __tfId0 + " ; curTF=" + __cfId0
                      + " ; over(tf)=" + (tf&&tf.isValid?tf.overflows:"NA")
                      + " ; over(curTF)=" + (curTextFrame&&curTextFrame.isValid?curTextFrame.overflows:"NA")
                      + " ; storyLen=" + story.characters.length);
                }catch(_){}
                try {
                  if (typeof flushOverflow === "function" && typeof tf !== "undefined" && tf && tf.isValid) {
                    var _rs = flushOverflow(story, page, tf);
                    if (_rs && _rs.frame && _rs.page) { page = _rs.page; tf = _rs.frame; story = tf.parentStory; curTextFrame = tf; }
                  }
                  try{
                      var _ipEnd = story.insertionPoints[-1];
                      var _holder = (_ipEnd && _ipEnd.isValid && _ipEnd.parentTextFrames && _ipEnd.parentTextFrames.length)
                                      ? _ipEnd.parentTextFrames[0] : null;
                      if (_holder && _holder.isValid) {
                        tf = _holder;
                        curTextFrame = _holder;
                        try { page = _holder.parentPage; } catch(_){}
                      }
                      try{
                        log("[IMG-LOC][after-flush] holder=" + (_holder?_holder.id:"NA")
                            + " page=" + ((page&&page.isValid)?page.name:"NA")
                            + " ; tf=" + (tf&&tf.isValid?tf.id:"NA")
                            + " ; curTF=" + (curTextFrame&&curTextFrame.isValid?curTextFrame.id:"NA"));
                      }catch(__){}
                  }catch(_){}
                  var _ipEnd = story.insertionPoints[-1];
                  var _holder = (_ipEnd && _ipEnd.isValid && _ipEnd.parentTextFrames && _ipEnd.parentTextFrames.length)
                                  ? _ipEnd.parentTextFrames[0] : null;
                  if (_holder && _holder.isValid) {
                    tf = _holder; curTextFrame = _holder;
                    try { page = _holder.parentPage; } catch(_){}
                  }
                  try{
                    log("[IMG-LOC][after-flush] holder=" + (_holder?_holder.id:"NA")
                        + " page=" + ((page&&page.isValid)?page.name:"NA")
                        + " ; tf=" + (tf&&tf.isValid?tf.id:"NA")
                        + " ; curTF=" + (curTextFrame&&curTextFrame.isValid?curTextFrame.id:"NA"));
                  }catch(__){}
                } catch(_){}
                try {
                  var lastChar = (story.characters.length>0) ? String(story.characters[-1].contents||"") : "";
                  if (lastChar !== "\r") story.insertionPoints[-1].contents = "\r";
                } catch(__){}

                var ipNow = (tf && tf.isValid) ? tf.insertionPoints[-1] : story.insertionPoints[-1];
                try{
                  var __h = (ipNow && ipNow.isValid && ipNow.parentTextFrames && ipNow.parentTextFrames.length) ? ipNow.parentTextFrames[0] : null;
                  var __pg = (__h && __h.isValid) ? __h.parentPage : null;
                  log("[IMG-LOC][ipNow] frame=" + (__h?__h.id:"NA") + " page=" + (__pg?__pg.name:"NA")
                      + " ; ip.index=" + (ipNow&&ipNow.isValid?ipNow.index:"NA"));
                }catch(_){}

                var fsrc = __imgNormPath(spec.src);
                if (fsrc && fsrc.exists) {
                  spec.src = fsrc.fsName;
                  try {
                    var fsrc = __imgNormPath(spec.src);
                    if (fsrc && fsrc.exists) {
                      spec.src = fsrc.fsName;

                      var inl = _trim(spec.inline);
                      log(__imgTag + " dispatch src="+spec.src+" inline="+inl+" posH="+(spec.posH||"")+" posV="+(spec.posV||""));
                  try{
                    var _preIP = (tf && tf.isValid) ? tf.insertionPoints[-1] : null;
                    log("[IMG-STACK] pre ip=" + (_preIP && _preIP.isValid?_preIP.index:"NA")
                        + " tf=" + (tf&&tf.isValid?tf.id:"NA")
                        + " page=" + (page?page.name:"NA"));
                  }catch(_){ }
                      try{
                        if (inl==="0" || /^false$/i.test(inl)){
                          var rect = __imgAddFloatingImage(tf, story, page, spec);
                          if (rect && rect.isValid) log("[IMG] ok (float): " + spec.src);
                        } else {
                          var rect = __imgAddImageAtV2(ipNow, spec);
                          if (rect && rect.isValid) log("[IMG] ok (inline): " + spec.src);
                        }
                      } catch(e) {
                        log("[ERR] addImage dispatch failed: " + e);
                      }
                    } else {
                      log("[IMG] missing: " + spec.src);
                    }
                    if (rect && rect.isValid) log("[IMG] ok: " + spec.src);
                  } catch(e) {
                    log("[ERR] addImageAt failed: " + e);
                  }
                } else {
                  log("[IMG] missing: " + spec.src);
                }

                try { story.insertionPoints[-1].appliedCharacterStyle = app.activeDocument.characterStyles.itemByName("[None]"); } catch(_){ try { story.insertionPoints[-1].appliedCharacterStyle = app.activeDocument.characterStyles[0]; } catch(__){} }
                last = re.lastIndex;
                continue;
            } else if (m[7]) {
                try {
                    var obj = __jsonParseSafe(m[7]);
                    __tblAddTableHiFi(obj);
                } catch(e){
                    try { var obj2 = eval("("+m[7]+")"); __tblAddTableHiFi(obj2); } catch(__){}
                }
            } else {
                var closing = !!m[4];
                var tag = (m[5] || "").toUpperCase();
                if (tag === "I") st.i += closing ? -1 : 1;
                else if (tag === "B") st.b += closing ? -1 : 1;
                else if (tag === "U") st.u += closing ? -1 : 1;
                if (st.i < 0) st.i = 0; if (st.b < 0) st.b = 0; if (st.u < 0) st.u = 0;
            }

            last = m.index + m[0].length;
        }

        var tail = text.substring(last);
        if (tail.length) {
            var sIdx = story.characters.length;
            story.insertionPoints[-1].contents = tail;
            var eIdx = story.characters.length;
            __applyInlineFormattingOnRange(story, sIdx, eIdx, {i:(st.i>0), b:(st.b>0), u:(st.u>0)});
        }
        try { story.insertionPoints[-1].appliedCharacterStyle = app.activeDocument.characterStyles.itemByName("[None]"); } catch(_){ try { story.insertionPoints[-1].appliedCharacterStyle = app.activeDocument.characterStyles[0]; } catch(__){} }


        story.insertionPoints[-1].contents = "\r";
        story.paragraphs[-1].appliedParagraphStyle = s;
        try {
            story.recompose(); app.activeDocument.recompose();
        } catch(_){}
        try {
            if (typeof __paraCounter === "undefined") __paraCounter = 0;
            __paraCounter++;
            if ((__paraCounter % 50) === 0) {
                var st = flushOverflow(story, page, tf);
                page  = st.page;
                tf    = st.frame;
                story = tf.parentStory;
                curTextFrame = tf;
            }
        } catch(_){}
        }catch(eAddPara){
            __logSkipParagraph(paraSeq, styleName, String(eAddPara||"error"), text);
            __recoverAfterParagraph(story, insertionStart);
        }
        try{
            __progressBump("PARA", "seq=" + paraSeq + " style=" + styleName);
        }catch(_){}
    }

    var templateFile = File("%TEMPLATE_PATH%");
    if (!templateFile.exists) { alert("未找到模板文件 template.idml"); return; }
    var doc = app.open(templateFile);
    try{
        doc.allowPageShuffle = true;
        try{
            var __dp = doc.documentPreferences;
            var __fpBefore = null;
            try{ __fpBefore = __dp.facingPages; }catch(__fpRead){}
            var __fpError = false;
            try{
                __dp.facingPages = false;
            }catch(__fpAssign){
                __fpError = true;
                try{ __dp.properties = { facingPages: false }; __fpError = false; }catch(__fpProp){}
            }
            try{ log("[LAYOUT] facingPages before=" + __fpBefore + " after=" + __dp.facingPages + " assignErr=" + __fpError); }catch(__faceLog){}
        }catch(__face){}
        try{
            var spreads = doc.spreads;
            try{ log("[LAYOUT] spreads init count=" + (spreads ? spreads.length : "NA")); }catch(__spreadCntLog){}
            for (var si=0; spreads && si<spreads.length; si++){
                try{ spreads[si].allowPageShuffle = true; }catch(__spreadEnable){}
            }
        }catch(__spreadLoop){}
    }catch(__allowDoc){}

    for (var pi = doc.pages.length - 1; pi >= 0; pi--) {
        var pg = doc.pages[pi];
        for (var tfi = pg.textFrames.length - 1; tfi >= 0; tfi--) {
            try { pg.textFrames[tfi].remove(); } catch(e) { try { pg.textFrames[tfi].contents = ""; } catch(_) {} }
        }
    }
    try {
        var msp = doc.masterSpreads;
        for (var mi = 0; mi < msp.length; mi++) {
            var ms = msp[mi];
            for (var it = ms.textFrames.length - 1; it >= 0; it--) {
                try { ms.textFrames[it].remove(); } catch(e) {}
            }
        }
    } catch(e) {}
    while (doc.pages.length > 1) { doc.pages[doc.pages.length - 1].remove(); }
    try{
        doc.allowPageShuffle = true;
        var __dpAfterTrim = doc.documentPreferences;
        try{ __dpAfterTrim.facingPages = false; }catch(__faceAfter){
            try{ __dpAfterTrim.properties = { facingPages: false }; }catch(__faceAfterProp){}
        }
        try{
            var __spreadsAfter = doc.spreads;
            for (var __si=0; __spreadsAfter && __si<__spreadsAfter.length; __si++){
                try{ __spreadsAfter[__si].allowPageShuffle = true; }catch(__spAllow){}
            }
        }catch(__spreadTrim){}
        try{ log("[LAYOUT] post-trim spreads=" + doc.spreads.length + " facing=" + __dpAfterTrim.facingPages); }catch(__trimLog){}
    }catch(__allowTrim){}
    __DEFAULT_LAYOUT = (function(){
        var state = {};
        try{
            var dp = doc.documentPreferences;
            if (dp){
                try{
                    var ori = "portrait";
                    if (dp.pageOrientation === PageOrientation.LANDSCAPE) ori = "landscape";
                    state.pageOrientation = ori;
                }catch(_){ }
                try{ state.pageWidthPt = parseFloat(dp.pageWidth); }catch(_){ }
                try{ state.pageHeightPt = parseFloat(dp.pageHeight); }catch(_){ }
            }
        }catch(_){ }
        try{
            var mpSource = null;
            try{ if (doc.pages.length > 0){ mpSource = doc.pages[0].marginPreferences; } }catch(_){ }
            if (!mpSource){ try{ mpSource = doc.marginPreferences; }catch(_){ } }
            if (mpSource){
                state.pageMarginsPt = {
                    top: parseFloat(mpSource.top),
                    bottom: parseFloat(mpSource.bottom),
                    left: parseFloat(mpSource.left),
                    right: parseFloat(mpSource.right)
                };
            }
        }catch(_){ }
        return __cloneLayoutState(state);
    })();
    __CURRENT_LAYOUT = __cloneLayoutState(__DEFAULT_LAYOUT);


    function ensureStyle(name, pointSize, leading, spaceBefore, spaceAfter) {
        var ps = doc.paragraphStyles.itemByName(name);
        if (!ps.isValid) {
            ps = doc.paragraphStyles.add({
                name: name,
                pointSize: pointSize,
                leading: leading,
                spaceBefore: spaceBefore,
                spaceAfter: spaceAfter
            });
        }
        return ps;
    }
    __STYLE_LINES__

    var page  = doc.pages[0];
    try{ log("[LOG] script boot ok; page="+doc.pages.length); }catch(_){}

    var tf    = createTextFrameOnPage(page, __DEFAULT_LAYOUT);
    if (__DEFAULT_INNER_WIDTH === null) __DEFAULT_INNER_WIDTH = _innerFrameWidth(tf);
    if (__DEFAULT_INNER_HEIGHT === null) __DEFAULT_INNER_HEIGHT = _innerFrameHeight(tf);
    try{ log("[LAYOUT] default inner width=" + __DEFAULT_INNER_WIDTH + " height=" + __DEFAULT_INNER_HEIGHT); }catch(_defaultLog){}
    var story = tf.parentStory;
    curTextFrame = tf; 

    var firstChapterSeen = false;
    __resetParaSeq();

    __ADD_LINES__
    var tail = flushOverflow(story, page, tf);
    page  = tail.page;
    tf    = tail.frame;
    story = tf.parentStory;
    curTextFrame = tf;
    __trimTrailingEmptyFrames(story);
    __trimTrailingEmptyPages(doc);
    try { fixAllTables(); } catch(_) {}
    try{ __progressFinalize(); }catch(_){ }

    var __AUTO_EXPORT = (CONFIG && CONFIG.flags && typeof CONFIG.flags.autoExportIdml === "boolean")
                        ? CONFIG.flags.autoExportIdml : %AUTO_EXPORT%;
    if (__AUTO_EXPORT) {
        try {
            var outFile = File("%OUT_IDML%");
            doc.exportFile(ExportFormat.INDESIGN_MARKUP, outFile, false);
        } catch(ex) { alert("导出 IDML 失败: " + ex); }
    }
    try{
        if (__origScriptUnit !== null) app.scriptPreferences.measurementUnit = __origScriptUnit;
    }catch(_){ }
    try{
        if (__origViewH !== null) app.viewPreferences.horizontalMeasurementUnits = __origViewH;
        if (__origViewV !== null) app.viewPreferences.verticalMeasurementUnits = __origViewV;
    }catch(_){ }
