    // config + selfcheck
    function __initEnvironment(){
      var state = {scriptUnit:null, viewH:null, viewV:null, userLevel:null};
      try{ state.userLevel = app.scriptPreferences.userInteractionLevel; }catch(_){}
      try{ app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT; }catch(_){}
      try{ state.scriptUnit = app.scriptPreferences.measurementUnit; }catch(_){}
      try{
        state.viewH = app.viewPreferences.horizontalMeasurementUnits;
        state.viewV = app.viewPreferences.verticalMeasurementUnits;
      }catch(_){}
      try{ app.scriptPreferences.measurementUnit = MeasurementUnits.POINTS; }catch(_){}
      try{
        app.viewPreferences.horizontalMeasurementUnits = MeasurementUnits.POINTS;
        app.viewPreferences.verticalMeasurementUnits = MeasurementUnits.POINTS;
      }catch(_){}
      return state;
    }
    function __restoreEnvironment(state){
      if (!state) return;
      try{
        if (state.scriptUnit !== null) app.scriptPreferences.measurementUnit = state.scriptUnit;
      }catch(_){}
      try{
        if (state.viewH !== null) app.viewPreferences.horizontalMeasurementUnits = state.viewH;
        if (state.viewV !== null) app.viewPreferences.verticalMeasurementUnits = state.viewV;
      }catch(_){}
      try{
        if (state.userLevel !== null) app.scriptPreferences.userInteractionLevel = state.userLevel;
      }catch(_){}
    }

    function __initEntryLogging(){
      var EVENT_FILE = File("%EVENT_LOG_PATH%");
      var LOG_WRITE  = (CONFIG && CONFIG.flags && typeof CONFIG.flags.logWrite === "boolean")
                       ? CONFIG.flags.logWrite : %LOG_WRITE%;   // true=log debug; false=only warn/error/info
      var ctx = __initEventLog(EVENT_FILE, LOG_WRITE);
      try{
        if (EVENT_FILE){
          EVENT_FILE.encoding = "UTF-8";
          EVENT_FILE.open("w");
          EVENT_FILE.writeln("");
          EVENT_FILE.close();
        }
      }catch(_){}
      var __LAST_LAYOUT_LOG = null;
      function __logLayoutEvent(message){
        if (!__LAST_LAYOUT_LOG || __LAST_LAYOUT_LOG !== message){
          __LAST_LAYOUT_LOG = message;
          __pushEvent(ctx, "debug", message);
        }
      }
      function logWrap(m){
        if (String(m||"").indexOf("[LAYOUT]") === 0){
          __logLayoutEvent(String(m));
        } else if (String(m||"").indexOf("[WARN]") === 0){
          __LAST_LAYOUT_LOG = null;
          __pushEvent(ctx, "warn", m);
        } else if (String(m||"").indexOf("[ERR]") === 0 || String(m||"").indexOf("[ERROR]") === 0){
          __LAST_LAYOUT_LOG = null;
          __pushEvent(ctx, "error", m);
        } else {
          __LAST_LAYOUT_LOG = null;
          __pushEvent(ctx, "debug", m);
        }
      }
      function selfCheck(){
        try{
          if (String(EVENT_FILE || "").indexOf("%") >= 0) throw "EVENT_LOG_PATH placeholder not replaced";
          if (String(LOG_WRITE).indexOf("%") >= 0) throw "LOG_WRITE placeholder not replaced";
          var required = ["__ensureLayoutDefault","__imgAddImageAtV2","__imgAddFloatingImage","__tblAddTableHiFi"];
          for (var i=0;i<required.length;i++){
            var n = required[i];
            if (typeof eval(n) !== "function") throw ("missing function: " + n);
          }
        }catch(e){
          try{ logWrap("[ERR] selfcheck failed: " + e); }catch(__){}
          throw e;
        }
      }
      return {
        EVENT_FILE: EVENT_FILE,
        LOG_WRITE: LOG_WRITE,
        ctx: ctx,
        info: function(m){ __pushEvent(ctx, "info", m); },
        warn: function(m){ __pushEvent(ctx, "warn", m); },
        err:  function(m){ __pushEvent(ctx, "error", m); },
        log: logWrap,
        logLayout: __logLayoutEvent,
        selfCheck: selfCheck
      };
    }

    var __LOG_CTX = __initEntryLogging();
    var EVENT_FILE = __LOG_CTX.EVENT_FILE;
    var LOG_WRITE  = __LOG_CTX.LOG_WRITE;
    try{ __DEBUG_WRITE = __LOG_CTX.LOG_WRITE; }catch(_){}
    var __EVENT_CTX = __LOG_CTX.ctx;
    var __logLayoutEvent = __LOG_CTX.logLayout;
    var info = __LOG_CTX.info;
    var warn = __LOG_CTX.warn;
    var err  = __LOG_CTX.err;
    function log(m){ __LOG_CTX.log(m); }

    var __ENV_STATE = __initEnvironment();
    __LOG_CTX.selfCheck();
    function __finalizeDocument(doc, story, page, tf){
      if (!doc || !doc.isValid) { __restoreEnvironment(__ENV_STATE); return; }
      try{ __trimTrailingEmptyFrames(story); }catch(_){}
      try{ __trimTrailingEmptyPages(doc); }catch(_){}
      try { fixAllTables(); } catch(_){}
      try{ __progressFinalize(); }catch(_){}
      var __AUTO_EXPORT = (CONFIG && CONFIG.flags && typeof CONFIG.flags.autoExportIdml === "boolean")
                          ? CONFIG.flags.autoExportIdml : %AUTO_EXPORT%;
      if (__AUTO_EXPORT) {
          try {
              var outFile = File("%OUT_IDML%");
              try{
                var outPathStr = String(outFile && outFile.absoluteURI ? outFile.absoluteURI : outFile || "");
                if (outPathStr.indexOf("%") >= 0){
                  try{ log("[ERR] OUT_IDML placeholder not replaced: " + outPathStr); }catch(_){}
                  outFile = null;
                }
              }catch(_){}
              if (!outFile){ /* skip export if placeholder missing */ }
              doc.exportFile(ExportFormat.INDESIGN_MARKUP, outFile, false);
          } catch(ex) { alert("Export IDML failed: " + ex); }
      }
      __restoreEnvironment(__ENV_STATE);
    }
    // alias util formatter for compatibility
    function applyInlineFormattingOnRange(story, startCharIndex, endCharIndex, st){
      return __applyInlineFormattingOnRange(story, startCharIndex, endCharIndex, st);
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
    try{ log("[BOOT] entry.js restore-hook version=force-call"); }catch(_){}
    function __tableRestoreLayout(){
        try{
            var mainLayout = __DEFAULT_LAYOUT || __CURRENT_LAYOUT;
            if (!mainLayout){
                try{ log("[TABLE][restore] skip: no mainLayout"); }catch(_){}
                return;
            }
            var lastLayout = __CURRENT_LAYOUT;
            // 横版模版：不做恢复，避免破坏横版布局
            var isDefaultLandscape = (__DEFAULT_LAYOUT && __DEFAULT_LAYOUT.pageOrientation === "landscape");
            if (isDefaultLandscape){
                try{ log("[TABLE][restore] skip: default layout is landscape"); }catch(_){}
                return;
            }
            if (mainLayout.pageOrientation === "portrait"
                && lastLayout && lastLayout.pageOrientation === "portrait"){
                try{ log("[TABLE][restore] skip: same portrait layout; no restore needed"); }catch(_){}
                return;
            }
            try{ log("[TABLE][restore] enter mainLayout orient=" + (mainLayout.pageOrientation||"") + " w=" + mainLayout.pageWidthPt + " h=" + mainLayout.pageHeightPt + " page=" + (page && page.name)); }catch(_){}
            try{
                var _docLen = (doc && doc.pages) ? doc.pages.length : "NA";
                var _docOff = (page && page.isValid && typeof page.documentOffset !== "undefined") ? page.documentOffset : "NA";
                log("[TABLE][restore] before create: docLen=" + _docLen + " page=" + (page&&page.name) + " docOff=" + _docOff + " tf=" + (tf&&tf.isValid?tf.id:"NA"));
            }catch(_){}
            try{ __pageBreak(); }catch(_){}
            var newFrame = null;
            var padPage = null;
            // 如果上一段布局是横版，则无条件先补一页横版空白，以保证跨页对齐
            try{
                var lastLayout = __CURRENT_LAYOUT;
                var needPadLandscape = (lastLayout && lastLayout.pageOrientation === "landscape");
                if (needPadLandscape){
                    try{ log("[TABLE][restore] add landscape pad to align spread; docOff=" + (page&&page.documentOffset)); }catch(_){}
                    try{ doc.allowPageShuffle = true; }catch(_){}
                    try{ padPage = doc.pages.add(LocationOptions.AFTER, page); }catch(_addPad){ try{ log("[TABLE][restore] pad page add failed: " + _addPad); }catch(__){} }
                    if (padPage && padPage.isValid){
                        try{
                            var wL = lastLayout.pageWidthPt, hL = lastLayout.pageHeightPt;
                            if (isFinite(wL) && isFinite(hL) && wL > 0 && hL > 0){
                                padPage.resize(CoordinateSpaces.PASTEBOARD_COORDINATES, AnchorPoint.TOP_LEFT_ANCHOR, ResizeMethods.REPLACING_CURRENT_DIMENSIONS_WITH, [wL, hL]);
                            }
                            var mpL = padPage.marginPreferences, marginsL = lastLayout.pageMarginsPt || {};
                            if (mpL){
                                if (isFinite(marginsL.top)) mpL.top = marginsL.top;
                                if (isFinite(marginsL.bottom)) mpL.bottom = marginsL.bottom;
                                if (isFinite(marginsL.left)) mpL.left = marginsL.left;
                                if (isFinite(marginsL.right)) mpL.right = marginsL.right;
                            }
                        }catch(_padApply){ try{ log("[TABLE][restore] pad apply layout failed: " + _padApply); }catch(__){} }
                    }
                }
            }catch(_padAll){ try{ log("[TABLE][restore] pad landscape failed: " + _padAll); }catch(__){} }

            // 创建竖版新跨页，并保证两页都为竖版，正文从新跨页第一页开始
            var portraitSpread = null;
            var baseSpread = (padPage && padPage.isValid) ? padPage.parent : (page && page.isValid ? page.parent : null);
            try{
                try{ doc.allowPageShuffle = true; }catch(_){}
                if (baseSpread && baseSpread.isValid){
                    portraitSpread = doc.spreads.add(LocationOptions.AFTER, baseSpread);
                } else {
                    portraitSpread = doc.spreads.add(LocationOptions.AT_END);
                }
            }catch(_addSp){ try{ log("[TABLE][restore] add portrait spread failed: " + _addSp); }catch(__){} }
            if (portraitSpread && portraitSpread.isValid){
                try{ portraitSpread.allowPageShuffle = true; }catch(_){}
                // 确保有两页
                try{
                    while(portraitSpread.pages.length > 1){ portraitSpread.pages[-1].remove(); }
                    if (portraitSpread.pages.length < 1){ portraitSpread.pages.add(); }
                }catch(_trim){}
                // 设定两页为竖版尺寸/边距
                try{
                    var wP = mainLayout.pageWidthPt, hP = mainLayout.pageHeightPt;
                    var marginsP = mainLayout.pageMarginsPt || {};
                    for (var pi=0; pi<portraitSpread.pages.length; pi++){
                        var pg = portraitSpread.pages[pi];
                        if (!pg || !pg.isValid) continue;
                        try{
                            if (isFinite(wP) && isFinite(hP) && wP>0 && hP>0){
                                pg.resize(CoordinateSpaces.PASTEBOARD_COORDINATES, AnchorPoint.TOP_LEFT_ANCHOR, ResizeMethods.REPLACING_CURRENT_DIMENSIONS_WITH, [wP, hP]);
                            }
                            var mp = pg.marginPreferences;
                            if (mp){
                                if (isFinite(marginsP.top)) mp.top = marginsP.top;
                                if (isFinite(marginsP.bottom)) mp.bottom = marginsP.bottom;
                                if (isFinite(marginsP.left)) mp.left = marginsP.left;
                                if (isFinite(marginsP.right)) mp.right = marginsP.right;
                            }
                        }catch(_pgApply){}
                    }
                }catch(_applyAll){ try{ log("[TABLE][restore] apply portrait layout failed: " + _applyAll); }catch(__){} }
                var targetPage = portraitSpread.pages.length ? portraitSpread.pages[0] : null;
                if (targetPage && targetPage.isValid){
                    try{
                        newFrame = createTextFrameOnPage(targetPage, mainLayout);
                    }catch(_cf2){ try{ log("[TABLE][restore] create frame failed: " + _cf2); }catch(__){} }
                    page = targetPage;
                }
            }
            if (newFrame && newFrame.isValid){
                try{ if (tf && tf.isValid) tf.nextTextFrame = newFrame; }catch(_){}
                tf = newFrame;
                story = tf.parentStory;
                curTextFrame = tf;
                try{ __CURRENT_LAYOUT = __cloneLayoutState(mainLayout); }catch(_){}
                try{
                    var _docLen2 = (doc && doc.pages) ? doc.pages.length : "NA";
                    var _docOff2 = (page && page.isValid && typeof page.documentOffset !== "undefined") ? page.documentOffset : "NA";
                    log("[LAYOUT] table restore new frame=" + tf.id + " page=" + (page && page.name) + " docOff=" + _docOff2 + " docLen=" + _docLen2);
                }catch(_){}
            } else {
                try{ log("[TABLE][restore] failed to create frame"); }catch(_){}
            }
        }catch(e){
            try{ log("[WARN] table restore layout failed: " + e); }catch(_){}
        }
    }
    log("[LOG] start");

    var __DEFAULT_LAYOUT = null;
    var __CURRENT_LAYOUT = null;
    var __DEFAULT_INNER_WIDTH = null;
    var __DEFAULT_INNER_HEIGHT = null;
    // shared story/page/frame for composition helpers
    var page = null;
    var tf = null;
    var story = null;
    var __ENABLE_TRAILING_TRIM = false;
    var __UNITVALUE_FAIL_ONCE = false;
    var __ALLOW_IMG_EXT_FALLBACK = (CONFIG && CONFIG.flags && typeof CONFIG.flags.allowImgExtFallback === "boolean")
                                   ? CONFIG.flags.allowImgExtFallback
                                   : (typeof $.global.__ALLOW_IMG_EXT_FALLBACK !== "undefined"
                                      ? !!$.global.__ALLOW_IMG_EXT_FALLBACK : true);
    var __SAFE_PAGE_LIMIT = (CONFIG && CONFIG.flags && typeof CONFIG.flags.safePageLimit === "number" && isFinite(CONFIG.flags.safePageLimit))
                             ? CONFIG.flags.safePageLimit : 2000;
    function __createProgressTracker(){
      var __PARA_SEQ = 0;
      var __PROGRESS_TOTAL = %PROGRESS_TOTAL%;
      var __PROGRESS_DONE = 0;
      var __PROGRESS_LAST_PCT = -1;
      var __PROGRESS_LAST_TS = (new Date()).getTime();
      var __PROGRESS_HEARTBEAT_MS = (CONFIG && CONFIG.progress && typeof CONFIG.progress.heartbeatMs === "number" && isFinite(CONFIG.progress.heartbeatMs))
                                    ? CONFIG.progress.heartbeatMs : %PROGRESS_HEARTBEAT%;
      function detailText(detail){
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
      function bump(kind, detail, forceLog){
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
          var dt = detailText(detail);
          if (dt) suffix = " " + dt;
          try{ __pushEvent(__EVENT_CTX, "progress", "[PROGRESS][" + kind + "] done=" + doneDisplay + "/" + __PROGRESS_TOTAL + " pct=" + pct + suffix); }catch(_){}
        }
      }
      function finalize(detail){
        if (!__PROGRESS_TOTAL || __PROGRESS_TOTAL <= 0) return;
        var suffix = "";
        var dt = detailText(detail);
        if (dt) suffix = " " + dt;
        var doneDisplay = Math.min(__PROGRESS_DONE, __PROGRESS_TOTAL);
        var pct = Math.min(100, Math.floor((doneDisplay * 100) / __PROGRESS_TOTAL));
        try{
          __pushEvent(__EVENT_CTX, "progress", "[PROGRESS][COMPLETE] done=" + doneDisplay + "/" + __PROGRESS_TOTAL + " pct=" + pct + suffix);
        }catch(_){}
      }
      function resetSeq(){ __PARA_SEQ = 0; }
      function nextSeq(){ __PARA_SEQ++; return __PARA_SEQ; }
      return {
        bump: bump,
        finalize: finalize,
        resetSeq: resetSeq,
        nextSeq: nextSeq
      };
    }
    var __PROGRESS = __createProgressTracker();
    function __progressBump(kind, detail, forceLog){ __PROGRESS.bump(kind, detail, forceLog); }
    function __progressFinalize(detail){ __PROGRESS.finalize(detail); }
    function __resetParaSeq(){ __PROGRESS.resetSeq(); }
    function __nextParaSeq(){ return __PROGRESS.nextSeq(); }
    function __logSkipParagraph(seq, styleName, reason, textSample, ctx){
      try{
        var preview = "";
        if (textSample){
          preview = String(textSample).replace(/\s+/g, " ");
          if (preview.length > 80) preview = preview.substring(0, 80) + "...";
        }
        var pName = "NA";
        var fId   = "NA";
        try{
          if (ctx && ctx.page && ctx.page.isValid && ctx.page.name) pName = ctx.page.name;
          else if (page && page.isValid && page.name) pName = page.name;
        }catch(_pg){}
        try{
          if (ctx && ctx.frame && ctx.frame.isValid && ctx.frame.id!=null) fId = ctx.frame.id;
          else if (tf && tf.isValid && tf.id!=null) fId = tf.id;
          else if (curTextFrame && curTextFrame.isValid && curTextFrame.id!=null) fId = curTextFrame.id;
        }catch(_fr){}
        var msg = "[SKIP][PARA " + seq + "] style=" + styleName + " page=" + pName + " frame=" + fId + " reason=" + reason;
        if (ctx && ctx.startPage) msg += " startPage=" + ctx.startPage;
        if (ctx && ctx.removed && ctx.removed.length){ msg += " removedPages=" + ctx.removed.join(","); }
        if (preview) msg += " text=\"" + preview + "\"";
        warn(msg);
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
    function __cleanupPagesAfterSkip(docObj, startPageCount){
      var removed = [];
      try{
        if (!docObj || !docObj.isValid) return removed;
        if (startPageCount === null || startPageCount === undefined) return removed;
        for (var idx = docObj.pages.length - 1; idx >= startPageCount; idx--){
          var pg = docObj.pages[idx];
          if (!pg || !pg.isValid) continue;
          var hasGraphics = false;
          try{ hasGraphics = (pg.allGraphics && pg.allGraphics.length > 0); }catch(_g){}
          var tfs = null;
          try{ tfs = pg.textFrames; }catch(_tfArr){}
          if (hasGraphics) continue;
          var keep = false;
          if (tfs && tfs.length){
            for (var ti=0; ti<tfs.length; ti++){
              var tfLocal = tfs[ti];
              if (!tfLocal || !tfLocal.isValid) continue;
              try{ if (tfLocal.tables && tfLocal.tables.length>0){ keep = true; break; } }catch(_tbl){}
              try{
                var txt = String(tfLocal.contents || "");
                if (txt.replace(/[\s\u0000-\u001f\u2028\u2029\uFFFC\uF8FF]+/g, "") !== "") { keep = true; break; }
              }catch(_txt){}
              try{ if (tfLocal.overflows){ keep = true; break; } }catch(_ov){}
            }
          }
          try{
            if (pg.pageItems && pg.pageItems.length>0 && (!tfs || !tfs.length)){
              keep = true;
            }
          }catch(_pi){}
          if (keep) continue;
          try{
            if (tfs && tfs.length){
              var tf0 = tfs[0];
              try{
                var prev = tf0.previousTextFrame;
                if (prev && prev.isValid){
                  try{ prev.nextTextFrame = null; }catch(_lnk){}
                }
              }catch(_prev){}
            }
          }catch(_dec){}
          try{ removed.push(pg.name); }catch(_nm){ removed.push(String(idx)); }
          try{ pg.remove(); }catch(_rmPg){}
        }
      }catch(_c){}
      return removed;
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
        var docRef = app.activeDocument;
        var startPageCount = null;
        var startPageName = null;
        var startFrameCtx = null;
        try{
          if (docRef && docRef.isValid){
            startPageCount = docRef.pages.length;
          }
          var ipStart = story && story.isValid ? story.insertionPoints[-1] : null;
          startFrameCtx = (ipStart && ipStart.isValid && ipStart.parentTextFrames && ipStart.parentTextFrames.length) ? ipStart.parentTextFrames[0] : null;
          var pgStart = (startFrameCtx && startFrameCtx.isValid) ? startFrameCtx.parentPage : null;
          if (pgStart && pgStart.isValid && pgStart.name) startPageName = pgStart.name;
        }catch(_ctx){}
        var s = docRef.paragraphStyles.itemByName(styleName);
        try { log("[PARA] style=" + styleName + " len=" + String(raw||"").length); } catch(_){}
        if (!s.isValid) { s = app.activeDocument.paragraphStyles.add({name:styleName}); }

        var text = String(raw).replace(/^\s+|\s+$/g, "");
        try{
          // 兜底将 <sup>/<sub> 转成标记，避免遗漏
          if (/<sup>/i.test(text) || /<sub>/i.test(text)){
            text = text.replace(/<sup>([\s\S]*?)<\/sup>/gi, "[[SUP]]$1[[/SUP]]");
            text = text.replace(/<sub>([\s\S]*?)<\/sub>/gi, "[[SUB]]$1[[/SUB]]");
          }
          if (text.indexOf("[[SUP") !== -1 || text.indexOf("[[SUB") !== -1){
            try{ log("[SUPSUB][TEXT] para="+idx+" style="+styleName+" snippet="+text.substr(0,80)); }catch(_slog){}
          }
        }catch(_p){ try{ log("[SUPSUB][MARK] preprocess failed: "+_p); }catch(__){} }
        if (text.length === 0) return;

        var insertionStart = 0;
        try{ insertionStart = (story && story.isValid) ? story.characters.length : 0; }catch(_){ }
        var __IMG_MARK_RE = /\[\[IMG\s+[^\]]+\]\]/g;
        var __imgMatchArr = text.match(__IMG_MARK_RE) || [];
        var __multiImgPara = (__imgMatchArr.length > 1);
        var __imgGroupSpecs = []; // collect when multi

        try{
                var re = /\[{2,}FNI:(\d+)\]{2,}|\[{2,}(FN|EN):(.*?)\]{2,}|\[\[(\/?)(I|B|U|SUP|SUB)\]\]|\[\[IMG\s+([^\]]+)\]\]|\[\[TABLE\s+(\{[\s\S]*?\})\]\]/g;
                var last = 0, m;
                var st = {i:0, b:0, u:0, sup:0, sub:0};
                var noteCtx = {story: story, tf: tf, page: page, stFlags: st, pendingNoteId: null};

                while ((m = re.exec(text)) !== null) {
                    var chunk = text.substring(last, m.index);
                    if (chunk.length) {
                        var startIdx = story.characters.length;
                        story.insertionPoints[-1].contents = chunk;
                        var endIdx   = story.characters.length;
                        __applyInlineFormattingOnRange(story, startIdx, endIdx, {
                          i:(st.i>0),
                          b:(st.b>0),
                          u:(st.u>0),
                          sup:(st.sup>0),
                          sub:(st.sub>0),
                          sup2:(st.sup2>0),
                          sub2:(st.sub2>0)
                        });
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

                if (__multiImgPara){
                  __imgGroupSpecs.push(spec);
                  last = re.lastIndex;
                  continue;
                }
                // single image path
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
                try{ log("[TABLE][restore] branch entered raw=" + String(m[7]).substring(0,120)); }catch(_){}
                try {
                    var obj = __jsonParseSafe(m[7]);
                    try{ log("[TABLE][restore] parse json start"); }catch(_){}
                    __tblAddTableHiFi(obj);
                    try{ log("[TABLE][restore] FORCE call after JSON table"); }catch(_){}
                } catch(e){
                    try {
                        try{ log("[TABLE][restore] json failed, eval start: " + e); }catch(_){}
                        var obj2 = eval("("+m[7]+")");
                        __tblAddTableHiFi(obj2);
                        try{ log("[TABLE][restore] FORCE call after eval table"); }catch(_){}
                    } catch(__evalErr){
                        try{ log("[TABLE][restore] parse table failed: " + __evalErr); }catch(_){}
                    }
                }
                try{ __tableRestoreLayout(); }catch(__callRestore){ try{ log("[TABLE][restore] call error: " + __callRestore); }catch(_){}} 
                try{ __ensureLayoutDefault(); }catch(__callEns){ try{ log("[TABLE][restore] ensure default error: " + __callEns); }catch(_){}} 
        } else {
            var closing = !!m[4];
                var tag = (m[5] || "").toUpperCase();
                if (tag === "I") st.i += closing ? -1 : 1;
                else if (tag === "B") st.b += closing ? -1 : 1;
                else if (tag === "U") st.u += closing ? -1 : 1;
                else if (tag === "SUP") { st.sup += closing ? -1 : 1; st.sub = 0; }
                else if (tag === "SUB") { st.sub += closing ? -1 : 1; st.sup = 0; }
                if (st.i < 0) st.i = 0; if (st.b < 0) st.b = 0; if (st.u < 0) st.u = 0; if (st.sup < 0) st.sup = 0; if (st.sub < 0) st.sub = 0;
            }

            last = m.index + m[0].length;
        }

        var tail = text.substring(last);
        if (tail.length) {
            var sIdx = story.characters.length;
            story.insertionPoints[-1].contents = tail;
            var eIdx = story.characters.length;
        __applyInlineFormattingOnRange(story, sIdx, eIdx, {i:(st.i>0), b:(st.b>0), u:(st.u>0), sup:(st.sup>0), sub:(st.sub>0)});
        }
        try { story.insertionPoints[-1].appliedCharacterStyle = app.activeDocument.characterStyles.itemByName("[None]"); } catch(_){ try { story.insertionPoints[-1].appliedCharacterStyle = app.activeDocument.characterStyles[0]; } catch(__){} }

        // place grouped images if collected
        try{
          if (__multiImgPara && __imgGroupSpecs && __imgGroupSpecs.length>1){
            try{ log("[IMG-GROUP] placing group count=" + __imgGroupSpecs.length); }catch(_){}
            try{
              if (typeof __imgPlaceImageGroup === "function"){
                var _ctx = __imgPlaceImageGroup(tf, story, page, __imgGroupSpecs);
                if (_ctx && _ctx.tf && _ctx.tf.isValid) tf = _ctx.tf;
                if (_ctx && _ctx.page && _ctx.page.isValid) page = _ctx.page;
                if (_ctx && _ctx.story && _ctx.story.isValid) story = _ctx.story;
              }
            }catch(_grp){}
          }
        }catch(_){}

        story.insertionPoints[-1].contents = "\r";
        story.paragraphs[-1].appliedParagraphStyle = s;
        try {
            story.recompose(); app.activeDocument.recompose();
        } catch(_){}
        try {
            if (typeof __paraCounter === "undefined") __paraCounter = 0;
            __paraCounter++;
            if ((__paraCounter % 50) === 0) {
                var st = flushOverflow(story, page, tf, 1);
                page  = st.page;
                tf    = st.frame;
                story = tf.parentStory;
                curTextFrame = tf;
            }
        } catch(_){}
        try{
          if (typeof flushOverflow === "function" && tf && tf.isValid){
            var st2 = flushOverflow(story, page, tf, 1);
            if (st2 && st2.frame && st2.page) { page = st2.page; tf = st2.frame; story = tf.parentStory; curTextFrame = tf; }
            if (st2 && st2.overset){
              var pageHint = "NA";
              var frameHint = null;
              try{
                var _ipHint = story && story.isValid ? story.insertionPoints[-1] : null;
                frameHint = (_ipHint && _ipHint.isValid && _ipHint.parentTextFrames && _ipHint.parentTextFrames.length) ? _ipHint.parentTextFrames[0] : null;
                var _pgHint = (frameHint && frameHint.isValid) ? frameHint.parentPage : null;
                if (_pgHint && _pgHint.isValid && _pgHint.name) pageHint = _pgHint.name;
              }catch(_ph){}
              var removedPages = __cleanupPagesAfterSkip(docRef, startPageCount);
              __logSkipParagraph(paraSeq, styleName, "overset/no-progress page="+pageHint, text, {page:(frameHint&&frameHint.isValid?frameHint.parentPage:null), frame:frameHint, startPage:startPageName, removed:removedPages});
              __recoverAfterParagraph(story, insertionStart);
              return;
            }
          }
        }catch(_skipFlush){}
        }catch(eAddPara){
            var removedErr = __cleanupPagesAfterSkip(docRef, startPageCount);
            __logSkipParagraph(paraSeq, styleName, String(eAddPara||"error"), text, {page:page, frame:tf, startPage:startPageName, removed:removedErr});
            __recoverAfterParagraph(story, insertionStart);
        }
        try{
            __progressBump("PARA", "seq=" + paraSeq + " style=" + styleName);
        }catch(_){}
    }

    function __openAndPrepareTemplate(){
      var templateFile = File("%TEMPLATE_PATH%");
      try{
        var tplPathStr = String(templateFile && templateFile.absoluteURI ? templateFile.absoluteURI : templateFile || "");
        if (tplPathStr.indexOf("%") >= 0){
          try{ log("[ERR] placeholder not replaced for template path: " + tplPathStr); }catch(_){}
          return null;
        }
      }catch(_){}
      if (!templateFile.exists) { alert("Template file not found: template.idml"); return null; }
      var doc = app.open(templateFile);
      try{
        doc.allowPageShuffle = true;
    try{
      var __dp = doc.documentPreferences;
      var __fpBefore = null;
      try{ __fpBefore = __dp.facingPages; }catch(__fpRead){}
      var __fpError = false;
      try{ __dp.facingPages = false; }
      catch(__fpAssign){
        __fpError = true;
        try{ __dp.properties = { facingPages: false }; __fpError = false; }catch(__fpProp){}
      }
      try{ log("[LAYOUT] facingPages before=" + __fpBefore + " after=" + __dp.facingPages + " assignErr=" + __fpError); }catch(__faceLog){}
    }catch(__face){}
    // spreads allow shuffle
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
  return doc;
}

    var doc = __openAndPrepareTemplate();
    if (!doc || !doc.isValid) { __restoreEnvironment(__ENV_STATE); return; }



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

    function __composeDocument(doc){
      if (!doc || !doc.isValid) { try{ log("[ERR] compose: doc invalid"); }catch(_){ } return; }
      try{
        page  = doc.pages[0];
        try{ log("[LOG] script boot ok; page="+doc.pages.length); }catch(_){}

        tf    = createTextFrameOnPage(page, __DEFAULT_LAYOUT);
        if (__DEFAULT_INNER_WIDTH === null) __DEFAULT_INNER_WIDTH = _innerFrameWidth(tf);
        if (__DEFAULT_INNER_HEIGHT === null) __DEFAULT_INNER_HEIGHT = _innerFrameHeight(tf);
        try{ log("[LAYOUT] default inner width=" + __DEFAULT_INNER_WIDTH + " height=" + __DEFAULT_INNER_HEIGHT); }catch(_defaultLog){}
        story = tf.parentStory;
        curTextFrame = tf; 

        var firstChapterSeen = false;
        __resetParaSeq();

        __ADD_LINES__
        var tail = flushOverflow(story, page, tf, 1);
        if (!tail || !tail.frame || !tail.page) { try{ log("[ERR] compose: tail invalid"); }catch(_){ } __finalizeDocument(doc, story, page, tf); return; }
        page  = tail.page;
        tf    = tail.frame;
        story = tf.parentStory;
        curTextFrame = tf;
      }catch(__composeErr){
        try{ log("[ERR] compose failed: " + __composeErr); }catch(_){}
      }
      __finalizeDocument(doc, story, page, tf);
    }

    __composeDocument(doc);
