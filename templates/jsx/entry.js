(function () {
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

    // ====== 日志收集 ======
    var EVENT_FILE = File("%EVENT_LOG_PATH%");
    var LOG_WRITE  = %LOG_WRITE%;   // true=记录 debug；false=仅保留 warn/error/info
    var __EVENT_LINES = [];

    function __sanitizeLogMessage(m){
      var txt = String(m == null ? "" : m);
      txt = txt.replace(/[\r\n]+/g, " ").replace(/\t/g, " ");
      return txt;
    }
    function __pushEvent(level, message){
      if (level === "debug" && !LOG_WRITE) return;
      var stamp = iso();
      __EVENT_LINES.push(level + "\t" + stamp + "\t" + __sanitizeLogMessage(message));
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
    function info(m){ __pushEvent("info", m); }
    function warn(m){ __pushEvent("warn", m); }
    function err(m){  __pushEvent("error", m); }
    var __LAST_LAYOUT_LOG = null;
    function __logLayoutEvent(message){
      if (!__LAST_LAYOUT_LOG || __LAST_LAYOUT_LOG !== message){
        __LAST_LAYOUT_LOG = message;
        __pushEvent("debug", message);
      }
    }
    function log(m){
      if (String(m||"").indexOf("[LAYOUT]") === 0){
        __logLayoutEvent(String(m));
      } else {
        __LAST_LAYOUT_LOG = null;
        __pushEvent("debug", m);
      }
    }
    function __flushEvents(){
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


    // 兼容 InDesign 2020：没有 String#trim
    if (!String.prototype.trim) {
      String.prototype.trim = function(){ return String(this).replace(/^\s+|\s+$/g, ""); };
    }

    function _trim(x){ 
        return String(x==null?"":x).replace(/^\s+|\s+$/g,""); 
    }

    log("[BOOT] JSX loaded");
    log("[LOG] start");

    // 全局状态：不要挂在 app 上（COM 对象不能扩展），改为脚本内私有变量
    var __FLOAT_CTX = {};               // 用于 addFloatingImage 的同段堆叠
    __FLOAT_CTX.imgAnchors = __FLOAT_CTX.imgAnchors || {};
    function __recordWordSeqPage(wordSeqVal, pageObj){
      try{
        if (!wordSeqVal || !pageObj || !pageObj.isValid) return;
        if (!__FLOAT_CTX) return;
        if (!__FLOAT_CTX.wordSeqPages) __FLOAT_CTX.wordSeqPages = {};
        __FLOAT_CTX.wordSeqPages[wordSeqVal] = {page: pageObj};
        if (__FLOAT_CTX.wordSeqBaseSeq == null){
          __FLOAT_CTX.wordSeqBaseSeq = wordSeqVal;
          __FLOAT_CTX.wordSeqBasePage = pageObj;
        }
      }catch(_){}
    }
    function __pageForWordSeq(wordSeqVal){
      try{
        if (!wordSeqVal) return null;
        var docRef = app && app.activeDocument;
        if (!docRef || !docRef.pages) return null;
        var extendGuard = 0;
        while (wordSeqVal > docRef.pages.length){
          if (__SAFE_PAGE_LIMIT && docRef.pages.length >= __SAFE_PAGE_LIMIT){
            try{ log("[ERROR] seq page request exceeds limit seq=" + wordSeqVal + " limit=" + __SAFE_PAGE_LIMIT); }catch(_){ }
            return null;
          }
          docRef.pages.add(LocationOptions.AT_END);
          extendGuard++;
          if (extendGuard > 50){
            try{ log("[ERROR] seq page request guard tripped seq=" + wordSeqVal); }catch(_){ }
            break;
          }
        }
        var pageObj = docRef.pages[wordSeqVal-1];
        if (pageObj && pageObj.isValid){
          __recordWordSeqPage(wordSeqVal, pageObj);
          return pageObj;
        }
      }catch(_pageSeq){}
      return null;
    }
    var __LAST_IMG_ANCHOR_IDX = -1;     // 用于 addImageAtV2 的“同锚点”检测
    var __DEFAULT_LAYOUT = null;
    var __CURRENT_LAYOUT = null;
    var __DEFAULT_INNER_WIDTH = null;
    var __DEFAULT_INNER_HEIGHT = null;
    var __ENABLE_TRAILING_TRIM = false;
    var __UNITVALUE_FAIL_ONCE = false;
    var __ALLOW_IMG_EXT_FALLBACK = (typeof $.global.__ALLOW_IMG_EXT_FALLBACK !== "undefined")
                                   ? !!$.global.__ALLOW_IMG_EXT_FALLBACK : true;
    var __SAFE_PAGE_LIMIT = 2000;
    var __PARA_SEQ = 0;
    var __PROGRESS_TOTAL = %PROGRESS_TOTAL%;
    var __PROGRESS_DONE = 0;
    var __PROGRESS_LAST_PCT = -1;
    var __PROGRESS_LAST_TS = (new Date()).getTime();
    var __PROGRESS_HEARTBEAT_MS = %PROGRESS_HEARTBEAT%;
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

    // 放在定义 log() 之后、其它函数之前即可
    if (typeof curTextFrame === "undefined" && typeof tf !== "undefined") {
      var curTextFrame = tf;
    }

    // —— 兼容 InDesign 2020：没有 JSON 对象 —— 
    var _HAS_JSON = (typeof JSON !== "undefined" && JSON && typeof JSON.stringify === "function");
    function _s(obj){
      // 尽量用 JSON.stringify；没有就手拼
      if (_HAS_JSON) {
        try { return JSON.stringify(obj); } catch(_){}
      }
      try {
        return "{i:" + (obj && obj.i ? 1:0) +
               ",b:" + (obj && obj.b ? 1:0) +
               ",u:" + (obj && obj.u ? 1:0) + "}";
      } catch(e) { return String(obj); }
    }

    // 在“当前文本框”末尾构造一个就地的安全插入点；仅在不可用时才退回 story 末尾
    function _safeIP(tf){
      try{
        if (tf && tf.isValid) {
          var ip = tf.insertionPoints[-1];   // 就地：当前文本框的末尾
          // 检测是否可用于锚定；不可用则在该框尾部补一个零宽空格再取一次
          try { var _t = ip.anchoredObjectSettings; }
          catch(e1){
            try { ip.contents = "\u200B"; } catch(_){}
            try { ip = tf.insertionPoints[-1]; } catch(_){}
          }
          if (ip && ip.isValid) return ip;
        }
      } catch(_){}
      // 兜底：story 末尾
      try{
        var story = (tf && tf.isValid) ? tf.parentStory : app.activeDocument.stories[0];
        var ip2 = story.insertionPoints[-1];
        try { var _t2 = ip2.anchoredObjectSettings; }
        catch(e2){ try { ip2.contents = "\u200B"; } catch(_){}
                   try { ip2 = story.insertionPoints[-1]; } catch(_){} }
        return ip2;
      }catch(e){ log("[LOG] _safeIP fallback error"); return null; }
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
            log("[LAYOUT] still skipping due to same state; page=" + (page && page.name) + " spec=" + JSON.stringify(__skipPayload));
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

    // ==== 图片路径解析（新增） ====
    // 这些目录会被依次尝试：脚本目录、脚本目录的 assets、XML 同目录、XML 同目录的 assets
    var IMG_DIRS = %IMG_DIRS_JSON%;
    function _normPath(p){
        if(!p) return null;
        p = String(p).replace(/^\s+|\s+$/g,"").replace(/\\/g,"/");
        // 直接支持 http(s) & data:，交给 InDesign 自己处理
        if (/^(https?:|data:)/i.test(p)) return File(p);
        // 先尝试原始路径
        try { var f0 = File(p); if (f0.exists) return f0; } catch(_){}
        // 仅文件名时，逐目录拼接
        var baseName = p.split("/").pop();
        function _alts(name){
            var i = name.lastIndexOf(".");
            if (i < 0) return [name, name+".png", name+".jpg", name+".jpeg"];
            if (!__ALLOW_IMG_EXT_FALLBACK) return [name];
            var stem = name.substring(0,i), ext = name.substring(i+1).toLowerCase();
            if (ext==="jpg")  return [name, stem+".jpeg", stem+".png"];
            if (ext==="jpeg") return [name, stem+".jpg",  stem+".png"];
            if (ext==="png")  return [name, stem+".jpg",  stem+".jpeg"];
            return [name];
        }
        var candNames = _alts(baseName);

        for (var i=0;i<IMG_DIRS.length;i++){
            try{
                for (var n=0;n<candNames.length;n++){
                    var f1 = File(IMG_DIRS[i]+"/"+candNames[n]);
                    if (f1.exists) {
                      if (__ALLOW_IMG_EXT_FALLBACK && candNames[n].toLowerCase() !== baseName.toLowerCase()) {
                        try{
                          log("[IMG] fallback ext hit base=" + baseName + " -> " + candNames[n] + " dir=" + IMG_DIRS[i]);
                        }catch(_){}
                      }
                      return f1;
                    }
                }
                var f2 = File(IMG_DIRS[i]+"/"+p);
                if (f2.exists) return f2;
            }catch(_){}
        }
        // …函数结尾附近
        try { p = decodeURI(p); } catch(_){}
        p = String(p).replace(/\\/g, "/");   // ← 新增：统一为正斜杠
        return File(p);
    }

    function logStep(s){ log("[IMGSTEP] " + s); }

    function addFloatingImage(tf, story, page, spec){
  log("[IMGFLOAT6] enter src="+(spec&&spec.src)+" w="+(spec&&spec.w)+" h="+(spec&&spec.h));
  function _toPtLocal(v){
    var s = String(v==null?"":v).replace(/^\s+|\s+$/g,"");
    if (/mm$/i.test(s)) return parseFloat(s)*2.83464567;
    if (/pt$/i.test(s)) return parseFloat(s);
    if (/px$/i.test(s)) return parseFloat(s)*0.75;
    if (s==="") return 0;
    var n = parseFloat(s); return isNaN(n)?0:n*0.75;
  }
  function _cloneSpec(base){
    var out = {};
    if (!base || typeof base !== "object") return out;
    for (var key in base){
      try{ out[key] = base[key]; }catch(_){}
    }
    return out;
  }
  function _fallbackToClassic(reason, anchorIndex, rectRef){
    try{ log("[IMGFLOAT6][FALLBACK] " + reason); }catch(_){}
    try{
      if (rectRef && rectRef.isValid){
        rectRef.remove();
      }
    }catch(_){}
    var ipFallback = null;
    try{
      if (tf && tf.isValid && tf.insertionPoints && tf.insertionPoints.length){
        ipFallback = tf.insertionPoints[-1];
      }
    }catch(_){}
    if ((!ipFallback || !ipFallback.isValid) && story && story.isValid){
      try{
        var safeIdx = (typeof anchorIndex === "number")
          ? Math.max(0, Math.min(anchorIndex, story.insertionPoints.length-1))
          : story.insertionPoints.length-1;
        ipFallback = story.insertionPoints[safeIdx];
      }catch(_){}
    }
    if (!ipFallback || !ipFallback.isValid){
      try{
        if (story && story.isValid && story.insertionPoints.length){
          ipFallback = story.insertionPoints[-1];
        }
      }catch(_){}
    }
    var fallbackSpec = _cloneSpec(spec);
    fallbackSpec.forceBlock = true;
    fallbackSpec.inline = "0";
    fallbackSpec.wrap = fallbackSpec.wrap || "none";
    fallbackSpec.__floatFallback = (fallbackSpec.__floatFallback || 0) + 1;
    return addImageAtV2(ipFallback, fallbackSpec);
  }

  var wordSeq = null;
  try{
    if (spec && spec.wordPageSeq){
      var tmpSeq = parseInt(spec.wordPageSeq, 10);
      if (!isNaN(tmpSeq) && isFinite(tmpSeq)) wordSeq = tmpSeq;
    }
  }catch(_){}
  try{
    if (!tf || !tf.isValid) { log("[IMGFLOAT6][ERR] tf invalid"); return null; }
    var f = _normPath(spec && spec.src);
    log("[IMGFLOAT6] resolved file="+(f?f.fsName:"NA"));
    if(!f || !f.exists){ log("[IMGFLOAT6][ERR] file missing: "+(spec&&spec.src)); return null; }

    function _lowerFlag(v){
      var s = String(v||"");
      return s ? s.toLowerCase() : "";
    }

    function _isPageAnchored(posHref, posVref){
      try{
        if (!(spec && spec.forceBlock)) return false;
      }catch(_){ return false; }
      var h = _lowerFlag(posHref);
      var v = _lowerFlag(posVref);
      var pageRefs = { "page":true, "pagearea":true, "pageedge":true, "margin":true, "spread":true };
      return !!(pageRefs[h] && pageRefs[v]);
    }

    function _placeOnPage(pageObj, stObj, anchorIdx, fileObj){
      if (!pageObj || !pageObj.isValid){
        log("[IMGFLOAT6][ERR] page invalid for page-level image");
        return null;
      }
      var pb = pageObj.bounds || [0,0,0,0];
      var mp = pageObj.marginPreferences || {};
      var pageTop = pb[0], pageLeft = pb[1], pageBottom = pb[2], pageRight = pb[3];
      var marginTop = parseFloat(mp.top)||0;
      var marginBottom = parseFloat(mp.bottom)||0;
      var marginLeft = parseFloat(mp.left)||0;
      var marginRight = parseFloat(mp.right)||0;
      var innerLeft = pageLeft + marginLeft;
      var innerRight = pageRight - marginRight;
      var innerTop = pageTop + marginTop;
      var innerBottom = pageBottom - marginBottom;
      var innerWidth = Math.max(1, innerRight - innerLeft);
      var innerHeight = Math.max(1, innerBottom - innerTop);

      var pageWidth = pageRight - pageLeft;
      var pageHeight = pageBottom - pageTop;
      var wordPageWidth = _toPtLocal(spec && spec.wordPageWidth);
      var wordPageHeight = _toPtLocal(spec && spec.wordPageHeight);

      var posHrefRaw = _lowerFlag(spec && spec.posHref);
      var posVrefRaw = _lowerFlag(spec && spec.posVref);
      var offXP = _toPtLocal(spec && spec.offX) || 0;
      var offYP = _toPtLocal(spec && spec.offY) || 0;
      if (wordPageWidth && wordPageWidth > 0){
        offXP = offXP * (pageWidth / wordPageWidth);
      }
      if (wordPageHeight && wordPageHeight > 0){
        offYP = offYP * (pageHeight / wordPageHeight);
      }
      var pageRefKeys = { "page":true, "pagearea":true, "pageedge":true, "margin":true, "spread":true };
      var useInnerH = !!pageRefKeys[posHrefRaw];
      var useInnerV = !!pageRefKeys[posVrefRaw] || posVrefRaw==="paragraph";

      var baseX = useInnerH ? innerLeft : pageLeft;
      if (posHrefRaw==="column") baseX = pageLeft + marginLeft;
      var baseY = useInnerV ? innerTop : pageTop;

      var maxWidth = useInnerH ? innerWidth : (pageRight - pageLeft);
      var maxHeight = useInnerV ? innerHeight : (pageBottom - pageTop);
      var targetW = wPt>0 ? Math.min(wPt, maxWidth) : maxWidth;
      var targetH = hPt>0 ? Math.min(hPt, maxHeight) : maxHeight;
      if (targetW <= 0) targetW = maxWidth;
      if (targetH <= 0) targetH = maxHeight;

      var guardL = Math.max(0, _toPtLocal(spec && spec.distL) || 0);
      var guardR = Math.max(0, _toPtLocal(spec && spec.distR) || 0);
      var guardTotal = guardL + guardR;
      if (guardTotal > 0){
        var availableW = Math.max(12, maxWidth - guardTotal);
        if (targetW > availableW) targetW = availableW;
      }

      var left = baseX + offXP;
      var top = baseY + offYP;
      var maxBottom = baseY + maxHeight;

      var innerLimitLeft = (useInnerH ? innerLeft : pageLeft) + guardL;
      var innerLimitRight = (useInnerH ? innerRight : pageRight) - guardR;
      if (innerLimitRight <= innerLimitLeft){
        innerLimitRight = innerLimitLeft + Math.max(10, targetW);
      }
      if (left < innerLimitLeft) left = innerLimitLeft;
      if (left > innerLimitRight - targetW) left = Math.max(innerLimitLeft, innerLimitRight - targetW);
      if (top < pageTop) top = pageTop;
      if (targetH > (maxBottom - top)) targetH = Math.max(10, maxBottom - top);
      var right = Math.min(innerLimitRight, left + targetW);
      targetW = Math.max(10, right - left);
      var bottom = top + targetH;

      var rect = pageObj.rectangles.add();
      try{ rect.strokeWeight = 0; rect.fillOpacity = 100; }catch(_){}
      rect.geometricBounds = [top, left, bottom, right];
      var placed = null;
      try{
        placed = rect.place(fileObj);
      }catch(ePlacePage){
        log("[IMGFLOAT6][ERR] page place failed: "+ePlacePage);
        try{ rect.remove(); }catch(__){}
        return null;
      }
      if (!placed || !placed.length || !(placed[0] && placed[0].isValid)){
        try{ rect.remove(); }catch(__){}
        log("[IMGFLOAT6][ERR] page place invalid result");
        return null;
      }
      try{ rect.fit(FitOptions.PROPORTIONALLY); rect.fit(FitOptions.CENTER_CONTENT); }catch(_){}
      _applyFloatTextWrap(rect);
      try{
        log("[IMGFLOAT6][PAGE] gb="+rect.geometricBounds+" w="+targetW.toFixed(2)+" h="+targetH.toFixed(2)
            +" offX="+offXP.toFixed(2)+" offY="+offYP.toFixed(2)+" page="+(pageObj.name||"NA"));
      }catch(_){}
      try{ rect.label = "PAGE-FLOAT"; }catch(_){}

      // 在 story 中依旧插入段落分隔，避免下一段落堆叠
      try{
        var aft1 = stObj && stObj.insertionPoints && stObj.insertionPoints.length
          ? stObj.insertionPoints[Math.min(stObj.insertionPoints.length-1, anchorIdx+1)]
          : null;
        if (aft1 && aft1.isValid) aft1.contents = "\r";
        var aft2 = stObj && stObj.insertionPoints && stObj.insertionPoints.length
          ? stObj.insertionPoints[Math.min(stObj.insertionPoints.length-1, anchorIdx+2)]
          : null;
        if (aft2 && aft2.isValid) aft2.contents = "\u200B";
        try{ stObj.recompose(); }catch(__re){}
      }catch(_){}

      return rect;
    }

    function _applyFloatTextWrap(rectObj){
      try{
        if (!rectObj || !rectObj.isValid) return;
        var tw = rectObj.textWrapPreferences;
        if (!tw) return;
        var wrapKey = _lowerFlag(spec && spec.wrap);
        var wrapMode = TextWrapModes.NONE;
        if (wrapKey === "wrapsquare" || wrapKey === "square"){
          wrapMode = TextWrapModes.BOUNDING_BOX_TEXT_WRAP;
        } else if (wrapKey === "wraptight" || wrapKey === "tight" || wrapKey === "wrapthrough"){
          wrapMode = TextWrapModes.OBJECT_SHAPE_TEXT_WRAP;
        } else if (wrapKey === "wraptopbottom" || wrapKey === "topbottom"){
          wrapMode = TextWrapModes.JUMP_OBJECT_TEXT_WRAP;
        } else if (wrapKey === "wrapbehind"){
          wrapMode = TextWrapModes.NONE;
        }
        tw.textWrapMode = wrapMode;
        if (wrapMode !== TextWrapModes.NONE){
          var distT = _toPtLocal(spec && spec.distT) || 0;
          var distB = _toPtLocal(spec && spec.distB) || 0;
          var distL = _toPtLocal(spec && spec.distL);
          var distR = _toPtLocal(spec && spec.distR);
          if (!distL && distL !== 0) distL = 12;
          if (!distR && distR !== 0) distR = 12;
          tw.textWrapOffset = [distT, distL, distB, distR];
        }
      }catch(_){}
    }

    var wPt=_toPtLocal(spec&&spec.w), hPt=_toPtLocal(spec&&spec.h);
    var posH=String((spec&&spec.posH)||"center").toLowerCase();
    var alignMode=String((spec&&spec.align)||"").toLowerCase();
    if (!alignMode){ alignMode = posH || "center"; }
    var wrap=String((spec&&spec.wrap)||"none").toLowerCase();
    var spB=_toPtLocal(spec&&spec.spaceBefore)||0;
    var spA=_toPtLocal(spec&&spec.spaceAfter); if (spA===0) spA = 2;
    var distT=_toPtLocal(spec&&spec.distT)||0, distB=_toPtLocal(spec&&spec.distB)||0,
        distL=_toPtLocal(spec&&spec.distL)||0, distR=_toPtLocal(spec&&spec.distR)||0;

    var st = tf.parentStory;
    try{
      var endIP=tf.insertionPoints[-1];
      var prev=(endIP&&endIP.isValid&&endIP.index>0)?st.insertionPoints[endIP.index-1]:null;
      var prevIsCR=false; try{ prevIsCR=(prev&&String(prev.contents)==="\r"); }catch(_){}
      if(!prevIsCR){ endIP.contents="\r"; try{ st.recompose(); }catch(__){} }
    }catch(_){}
      var ip = null;
      try{
        if (tf && tf.isValid){
          if (typeof _safeIP === "function"){
            ip = _safeIP(tf);
          }
          if (!ip || !ip.isValid){
            ip = tf.insertionPoints[-1];
          }
        }
      }catch(_){}
      if ((!ip || !ip.isValid) && st && st.isValid){
        try{ ip = st.insertionPoints[-1]; }catch(__){}
      }
      try{
        var ipIdx = "NA";
        if (ip && ip.isValid) ipIdx = ip.index;
        log("[IMGFLOAT6][DBG] dispatch ip.index=" + ipIdx);
      }catch(_){}
    if (!ip || !ip.isValid) { log("[IMGFLOAT6][ERR] invalid ip"); return null; }
    var anchorIndex = ip.index;
    var posHrefRaw = _lowerFlag(spec && spec.posHref);
    var posVrefRaw = _lowerFlag(spec && spec.posVref);
    var posVRaw = _lowerFlag(spec && spec.posV);
    try{
      log("[IMGFLOAT6][DBG] anchorFlags posHref="+posHrefRaw+" posVref="+posVrefRaw+" posV="+posVRaw+" inline="+(spec&&spec.inline));
    }catch(_){}
    var isPageAnchor = _isPageAnchored(posHrefRaw, posVrefRaw);
    try{
      log("[IMGFLOAT6][DBG] isPageAnchor="+isPageAnchor);
    }catch(_){}

    function _ensureNextPageFrame(basePage){
      try{
        if (!basePage || !basePage.isValid || typeof __createLayoutFrame !== "function") return null;
        var pkt = __createLayoutFrame(__CURRENT_LAYOUT, tf, {afterPage: basePage, forceBreak: true});
        if (pkt && pkt.frame && pkt.page){
          page  = pkt.page;
          tf    = pkt.frame;
          story = tf.parentStory;
          curTextFrame = tf;
          try{
            log("[IMGFLOAT6][PAGE] newFrame=" + tf.id + " page=" + (page?page.name:"NA"));
          }catch(_){}
          try{
            __FLOAT_CTX.lastTf = tf;
            __FLOAT_CTX.lastPage = page;
          }catch(_){}
          return pkt;
        }
      }catch(_){}
      return null;
    }

    if (isPageAnchor){
      var targetPage = page;
      try{
        if (!targetPage || !targetPage.isValid){
          var docPages = app.activeDocument.pages;
          if (docPages.length){
            targetPage = docPages[0];
          }
        }
        if (!targetPage || !targetPage.isValid){
          targetPage = (tf && tf.isValid && tf.parentPage && tf.parentPage.isValid) ? tf.parentPage : null;
        }
      }catch(_){}
      var pageRect = _placeOnPage(targetPage, st, anchorIndex, f);
          if (pageRect){
            try{
              var thisPage = (pageRect.parentPage && pageRect.parentPage.isValid) ? pageRect.parentPage : targetPage;
              if (thisPage && thisPage.isValid) page = thisPage;
              try{
                __recordWordSeqPage(wordSeq, thisPage);
              }catch(_){}
              try{
                if (__FLOAT_CTX){
                  if (!__FLOAT_CTX.imgAnchors) __FLOAT_CTX.imgAnchors = {};
                  var anchorHintKey = spec.anchorId || spec.docPrId || "";
                  if (anchorHintKey){
                    __FLOAT_CTX.imgAnchors[anchorHintKey] = {
                      page: page,
                      anchorX: (ip && ip.isValid) ? ip.horizontalOffset : null,
                      anchorY: (ip && ip.isValid) ? ip.baseline : null,
                      wordSeq: wordSeq
                    };
                    try{
                      log("[IMGFLOAT6][DBG] store ctx key=" + anchorHintKey + " page=" + (page?page.name:"NA"));
                    }catch(_){}
                  }
            }
          }catch(_){}
        }catch(_){}
        try{
          if (st && st.isValid){
            var afterPara = st.insertionPoints[Math.min(st.insertionPoints.length-1, anchorIndex+1)];
            if (afterPara && afterPara.isValid) afterPara.contents = "\r";
            var ztail = st.insertionPoints[Math.min(st.insertionPoints.length-1, anchorIndex+2)];
            if (ztail && ztail.isValid) ztail.contents = "\u200B";
            try{ st.recompose(); }catch(__re){}
          }
        }catch(_){}
        var nextPkt = _ensureNextPageFrame(page);
        if (nextPkt && nextPkt.frame && nextPkt.page){
          page = nextPkt.page;
          tf = nextPkt.frame;
          story = tf.parentStory;
          curTextFrame = tf;
          try{
            if (typeof _safeIP === "function"){
              ip = _safeIP(tf);
            }
            if ((!ip || !ip.isValid) && tf && tf.isValid && tf.insertionPoints && tf.insertionPoints.length){
              ip = tf.insertionPoints[-1];
            }
          }catch(_){}
        }
        try{ __LAST_IMG_ANCHOR_IDX = anchorIndex; }catch(_){}
        return pageRect;
      }
      // 若页面放置失败，继续走原浮动逻辑
    }

      var placed = null;
      try { placed = ip.place(f); } catch(ePl){ log("[IMGFLOAT6][ERR] place failed(ip): " + ePl); return null; }
      if (!placed || !placed.length || !(placed[0] && placed[0].isValid)) { log("[IMGFLOAT6][ERR] place returned invalid"); return null; }

      var item = placed[0], rect = null, cname = "";
    try { cname = String(item.constructor.name); } catch(_){}
    if (cname === "Rectangle") rect = item;
    else {
      try {
        var cur = item;
        for (var g=0; g<6 && cur && cur.isValid; g++){
          var nm=""; try{ nm=String(cur.constructor.name); }catch(__){}
          if (nm==="Rectangle"){ rect=cur; break; }
          cur = cur.parent;
        }
      } catch(_){}
    }
    if (!rect || !rect.isValid) { log("[IMGFLOAT6][ERR] no rectangle after place"); return null; }

    try {
      var _aos = rect.anchoredObjectSettings;
      if (_aos && _aos.isValid){
        _aos.anchoredPosition = AnchorPosition.ABOVE_LINE;
        _aos.anchorPoint      = AnchorPoint.TOP_LEFT_ANCHOR;
        try{ _aos.lockPosition = false; }catch(_){}
      }
    } catch(_){}
    _applyFloatTextWrap(rect);
    try{ rect.fittingOptions.autoFit=false; rect.absoluteHorizontalScale=100; rect.absoluteVerticalScale=100; }catch(_){ }
    try{
      var _imgCount = null, _gCount = null, _cid = null, _pid = null;
      try{ _imgCount = rect.images ? rect.images.length : "NA"; }catch(__){ _imgCount="ERR"; }
      try{ _gCount = rect.graphics ? rect.graphics.length : "NA"; }catch(___){ _gCount="ERR"; }
      try{ _cid = rect.id; }catch(__4){}
      try{ _pid = item && item.isValid ? item.id : "NA"; }catch(__5){}
      log("[IMGFLOAT6][DBG] container id="+_cid+" from item id="+_pid+" images="+_imgCount+" graphics="+_gCount);
    }catch(_){}

    try { var _gbF = rect.geometricBounds; var _wF = (_gbF[3]-_gbF[1]).toFixed(2), _hF = (_gbF[2]-_gbF[0]).toFixed(2);
          log("[IMGFLOAT6][FINAL] gb=["+_gbF.join(",")+"] W="+_wF+" H="+_hF); } catch(_){}
    log("[IMGFLOAT6] place() ok");

    try{
      try{ rect.fit(FitOptions.CENTER_CONTENT);}catch(_){}
      try{ rect.fittingOptions.autoFit=false; }catch(_){}

      var holder = null;
      try { if (rect.parentTextFrames && rect.parentTextFrames.length) holder = rect.parentTextFrames[0]; } catch(_){}
      if ((!holder || !holder.isValid) && tf && tf.isValid) holder = tf;

      var innerInfo = _holderInnerBounds(holder);
      var innerW = innerInfo.innerW;
      var innerH = innerInfo.innerH;
      try{
        log("[IMGFLOAT6][DBG] holderInner id=" + (holder && holder.isValid ? holder.id : "NA")
            + " innerW=" + innerW.toFixed(2) + " innerH=" + innerH.toFixed(2));
      }catch(_){}
      if (innerW <= 0){
        try{
          var hb = holder ? holder.geometricBounds : null;
          var inset = holder ? holder.textFramePreferences.insetSpacing : null;
          var li = (inset && inset.length>=2)? inset[1] : 0;
          var ri = (inset && inset.length>=4)? inset[3] : 0;
          innerW = (hb ? (hb[3]-hb[1]) : 0) - li - ri;
        }catch(_){}
      }


      // 优先使用单栏宽度（多栏情况下用 textColumnFixedWidth，保证与 Word 类似的列宽约束）
      try {
        var _colW = (holder && holder.isValid) ? holder.textFramePreferences.textColumnFixedWidth : 0;
        var _colN = (holder && holder.isValid) ? holder.textFramePreferences.textColumnCount       : 1;
        if (_colN > 1 && _colW > 0) innerW = _colW;
      } catch(_){ }
try{ st.recompose(); }catch(_){ }
var gb = null;
function _rectifyCandidate(obj){
  if (!obj || !obj.isValid) return null;
  var nm = "";
  try{ nm = String(obj.constructor.name); }catch(_){}
  if (nm === "Rectangle") return obj;
  try{
    if (obj.parent && obj.parent.isValid && String(obj.parent.constructor.name) === "Rectangle") {
      return obj.parent;
    }
  }catch(_){}
  try{
    if (obj.graphics && obj.graphics.length){
      var g = obj.graphics[0];
      if (g && g.isValid && g.parent && g.parent.isValid && String(g.parent.constructor.name) === "Rectangle") {
        return g.parent;
      }
    }
  }catch(_){}
  return null;
}

function _setRect(candidate, tag){
  if (!candidate || !candidate.isValid) return false;
  rect = candidate;
  try{ log("[IMGFLOAT6][RECT] " + tag); }catch(_){}
  return true;
}
function _ensureRectValid(_retry){
  var candidate = _rectifyCandidate(rect);
  if (candidate && _setRect(candidate, "reuse")){ return true; }

  try{
    var p0 = (placed && placed.length) ? placed[0] : null;
    candidate = _rectifyCandidate(p0);
    if (candidate && _setRect(candidate, "from placed[0]")){ return true; }
  }catch(_){}

  try{
    if (placed && placed.length){
      for (var ii=0; ii<placed.length; ii++){
        candidate = _rectifyCandidate(placed[ii]);
        if (candidate && _setRect(candidate, "from placed["+ii+"]")){ return true; }
      }
    }
  }catch(_){}

  try{
    if (st && st.isValid && typeof anchorIndex === "number"){
      var idx = Math.min(Math.max(anchorIndex, 0), st.insertionPoints.length-1);
      var anchorIP = st.insertionPoints[idx];
      if (anchorIP && anchorIP.isValid){
        try{
          var ao = anchorIP.anchoredObjects;
          if (ao && ao.length){
            for (var jj=0;jj<ao.length;jj++){
              candidate = _rectifyCandidate(ao[jj]);
              if (candidate && _setRect(candidate, "from anchoredObjects["+jj+"]")){ return true; }
            }
          }
        }catch(_){}
        try{
          var recs = anchorIP.rectangles;
          if (recs && recs.length){
            for (var kk=0;kk<recs.length;kk++){
              candidate = _rectifyCandidate(recs[kk]);
              if (candidate && _setRect(candidate, "from ip.rectangles["+kk+"]")){ return true; }
            }
          }
        }catch(_){}
      }
    }
  }catch(_){}

  try{
    if ((!rect || !rect.isValid) && st && st.isValid){
      st.recompose();
      try{ app.activeDocument.recompose(); }catch(__){}
      candidate = _rectifyCandidate(rect);
      if (candidate && _setRect(candidate, "after recompose")){ return true; }
    }
  }catch(_){}

  if ((!rect || !rect.isValid) && !_retry){
    try{
      log("[IMGFLOAT6][RECT] wait redraw");
    }catch(_){}
    try{ app.waitForRedraw(); }catch(__){}
    return _ensureRectValid(true);
  }

  return !!(rect && rect.isValid);
}
if (_ensureRectValid()){
  try { gb = rect.geometricBounds; }
  catch(eGB){
    if (_ensureRectValid()){
      try { gb = rect.geometricBounds; } catch(__){}
    }
  }
}
if (!gb){
  try{
    log('[IMGFLOAT6][DBG] gb invalid, use fallback sizing');
  }catch(_){}
  try{
    var dbgAbsW = null, dbgAbsH = null, dbgRectValid = (rect && rect.isValid);
    try{ dbgAbsW = rect.width; dbgAbsH = rect.height; }catch(__){}
    log('[IMGFLOAT6][DBG] before fallback rectValid=' + dbgRectValid
        + ' width=' + (dbgAbsW||'NA') + ' height=' + (dbgAbsH||'NA')
        + ' absScale=' + (rect?rect.absoluteHorizontalScale:'NA') + '/'
        + (rect?rect.absoluteVerticalScale:'NA'));
  }catch(___){}
}
      var curW = (gb ? Math.max(1e-6, gb[3]-gb[1]) : (wPt>0?wPt:innerW>0?innerW:(rect&&rect.isValid&&rect.width?rect.width:1)));
      var curH = (gb ? Math.max(1e-6, gb[2]-gb[0]) : (hPt>0?hPt:(rect&&rect.isValid&&rect.height?rect.height:(curW>0?curW:1))));
      var ratio = curW / Math.max(1e-6, curH);

      var targetW = curW;
      if (wPt>0){
        targetW = wPt;
        if (innerW>0) targetW = Math.min(targetW, innerW);
      } else if (innerW>0){
        targetW = Math.min(curW, innerW);
      }
      var targetH = curH;
      if (hPt>0 && wPt>0){
        targetH = hPt;
      } else if (hPt>0){
        targetH = hPt;
        if (wPt<=0) targetW = targetH * (ratio || 1);
      } else {
        targetH = targetW / (ratio || 1);
      }

      var targetRatio = targetW / Math.max(1e-6, targetH);
      if (innerH > 0 && targetH > innerH){
        targetH = innerH;
        targetW = targetH * targetRatio;
      }
      if (innerW > 0 && targetW > innerW){
        targetW = innerW;
        targetH = targetW / Math.max(1e-6, targetRatio);
      }

      var pageInnerH = 0;
      try{
        if (page && page.isValid){
          var pb = page.bounds;
          var mp = page.marginPreferences;
          if (pb && pb.length === 4){
            var topMargin = (mp && mp.top) ? parseFloat(mp.top) || 0 : 0;
            var bottomMargin = (mp && mp.bottom) ? parseFloat(mp.bottom) || 0 : 0;
            pageInnerH = (pb[2]-pb[0]) - topMargin - bottomMargin;
          }
        }
        if (pageInnerH > 0 && targetH > pageInnerH){
          targetH = pageInnerH;
          targetW = targetH * targetRatio;
        }
      }catch(_pageClamp){}
      try{
        log("[IMGFLOAT6][DBG] targetClamp W=" + targetW.toFixed(2)
            + " H=" + targetH.toFixed(2)
            + " pageInnerH=" + pageInnerH.toFixed(2)
            + " ratio=" + targetRatio.toFixed(2));
      }catch(_targetLog){}

      try{ rect.absoluteHorizontalScale=100; rect.absoluteVerticalScale=100; }catch(_){ }
      var _graphic = null;
      try{
        if (rect.images && rect.images.length) _graphic = rect.images[0];
        else if (rect.graphics && rect.graphics.length) _graphic = rect.graphics[0];
      }catch(_){}
      if (_graphic && _graphic.isValid){
        try{ _graphic.absoluteHorizontalScale = 100; }catch(__){}
        try{ _graphic.absoluteVerticalScale = 100; }catch(__){}
      }
      var _boundsApplied = false;
      var _boundsErr = null;
      if (gb){
        try{
          rect.geometricBounds = [gb[0], gb[1], gb[0] + targetH, gb[1] + targetW];
          _boundsApplied = true;
        }catch(eBounds){
          _boundsErr = eBounds;
        }
      }
      if (!_boundsApplied){
        try{
          var _holderGB = holder && holder.isValid ? holder.geometricBounds : null;
          if (_holderGB && _holderGB.length === 4){
            var topBase = _holderGB[0] + spB;
            var leftInset = (holder.textFramePreferences && holder.textFramePreferences.insetSpacing && holder.textFramePreferences.insetSpacing.length>=2)
                              ? holder.textFramePreferences.insetSpacing[1] : 0;
            var leftBase = _holderGB[1] + leftInset;
            try{
              rect.geometricBounds = [topBase, leftBase, topBase + targetH, leftBase + targetW];
              _boundsApplied = true;
            }catch(__gbManual){
              _boundsErr = __gbManual;
            }
          }
        }catch(__manualBounds){}
      }
      if (!_boundsApplied){
        if (_boundsErr){
          try{
            log("[IMGFLOAT6][FALLBACK] setBounds fallback: " + _boundsErr);
          }catch(_){}
        }
        try{
          var _aos = rect.anchoredObjectSettings;
          if (_aos && _aos.isValid){
            try{ _aos.anchoredObjectSizeOption = AnchorSize.HEIGHT_AND_WIDTH; }catch(__){}
            try{ _aos.width  = targetW; }catch(__){}
            try{ _aos.height = targetH; }catch(__){}
            _boundsApplied = true;
          }
        }catch(_){}
      }
      if (_boundsApplied){
        try{ st.recompose(); }catch(_){}
        try{ app.waitForRedraw(); }catch(_){}
      }
      if (!_boundsApplied){
        try{
          var scaleX = targetW / Math.max(1e-6, curW);
          var scaleY = targetH / Math.max(1e-6, curH);
          rect.absoluteHorizontalScale = scaleX * 100;
          rect.absoluteVerticalScale   = scaleY * 100;
        }catch(_){}
      }
      try { rect.fit(FitOptions.CENTER_CONTENT); } catch(_){ }
      try{
        var _gbPath = rect.geometricBounds;
        if (_gbPath && _gbPath.length === 4 && rect.paths && rect.paths.length){
          var _top = _gbPath[0], _left = _gbPath[1], _bottom = _gbPath[2], _right = _gbPath[3];
          rect.paths[0].entirePath = [
            [_left, _top],
            [_left, _bottom],
            [_right, _bottom],
            [_right, _top]
          ];
        }
      }catch(_){}
      try{
        log('[IMGFLOAT6][DBG] after fallback width=' + (rect.width||'NA')
            + ' height=' + (rect.height||'NA')
            + ' absScale=' + (rect.absoluteHorizontalScale||'NA') + '/'
            + (rect.absoluteVerticalScale||'NA'));
      }catch(_){}
      try{
        rect.frameFittingOptions.leftCrop   = 0;
        rect.frameFittingOptions.rightCrop  = 0;
        rect.frameFittingOptions.topCrop    = 0;
        rect.frameFittingOptions.bottomCrop = 0;
      }catch(_){}
      try{
        var host = rect.parent;
        var hop = 0;
        var gbFinal = null;
        try{ gbFinal = rect.geometricBounds; }catch(__gbFinal){}
        while (host && host.isValid && hop < 3){
          var cname="";
          try{ cname = String(host.constructor.name); }catch(__c){ cname=""; }
          try{ log("[IMGFLOAT6][DBG] host chain level="+hop+" name="+cname); }catch(__hc){}
          if (cname === "Rectangle" || cname === "Polygon"){
            try{
              if (gbFinal){
                host.geometricBounds = [gbFinal[0], gbFinal[1], gbFinal[2], gbFinal[3]];
                if (host.paths && host.paths.length){
                  host.paths[0].entirePath = [
                    [gbFinal[1], gbFinal[0]],
                    [gbFinal[1], gbFinal[2]],
                    [gbFinal[3], gbFinal[2]],
                    [gbFinal[3], gbFinal[0]]
                  ];
                }
              }
              host.frameFittingOptions.leftCrop   = 0;
              host.frameFittingOptions.rightCrop  = 0;
              host.frameFittingOptions.topCrop    = 0;
              host.frameFittingOptions.bottomCrop = 0;
            }catch(__adj){}
            try{ log("[IMGFLOAT6][DBG] shrink host level="+hop+" name="+cname); }catch(__log){}
          }
          try{ host = host.parent; }catch(__p){ host = null; }
          hop++;
        }
      }catch(__host){}

      try{
        var alignInfo = _alignFloatingRect(rect, holder, innerW, alignMode);
        if (alignInfo){
          log("[IMGFLOAT6][ALIGN] align="+alignInfo.align+" offset="+alignInfo.offset.toFixed(2)
              + " innerW="+(alignInfo.innerW||0)+" holder="+(alignInfo.holder?alignInfo.holder.id:'NA'));
        }
      }catch(_){}

      try { var _gb2 = rect.geometricBounds; var _w2 = (_gb2[3]-_gb2[1]).toFixed(2), _h2 = (_gb2[2]-_gb2[0]).toFixed(2);
            log("[IMGFLOAT6][POST] gb="+_gb2+" W="+_w2+" H="+_h2+" innerW="+(innerW||0)); } catch(_){
        try{
          app.waitForRedraw();
          var _gb3 = rect.geometricBounds;
          var _w3 = (_gb3[3]-_gb3[1]).toFixed(2), _h3 = (_gb3[2]-_gb3[0]).toFixed(2);
          log("[IMGFLOAT6][POST2] gb="+_gb3+" W="+_w3+" H="+_h3+" innerW="+(innerW||0));
        }catch(__){
          try{ log("[IMGFLOAT6][POST2] gb still invalid"); }catch(___){}
        }
      }

      log("[IMGFLOAT6] size W=" + (targetW||0).toFixed(2)
          + " H=" + (targetH||0).toFixed(2)
          + " innerW=" + (innerW||0).toFixed(2));
      var finalGb = null;
      try{ finalGb = rect.geometricBounds; }catch(_){}
      if (!finalGb || finalGb.length !== 4){
        return _fallbackToClassic("geometricBounds invalid", anchorIndex, rect);
      }
    } catch(eSz){ log("[IMGFLOAT6][DBG] size "+eSz); }

    try{
      var p = (st && st.isValid) ? st.insertionPoints[anchorIndex].paragraphs[0] : null;
      if(p && p.isValid){
        p.justification=(posH==="right")?Justification.RIGHT_ALIGN:(posH==="center"?Justification.CENTER_ALIGN:Justification.LEFT_ALIGN);
        p.spaceBefore=spB; p.spaceAfter=spA;
        p.keepOptions.keepWithNext=false; p.keepOptions.keepLinesTogether=false;
      }
    }catch(_){}

    try{
      var aft1 = st.insertionPoints[Math.min(st.insertionPoints.length-1, anchorIndex+1)];
      if(aft1 && aft1.isValid){ aft1.contents = "\r"; }
      var aft2 = st.insertionPoints[Math.min(st.insertionPoints.length-1, anchorIndex+2)];
      if(aft2 && aft2.isValid){ aft2.contents = "\u200B"; }
      try{ st.recompose(); }catch(__){}
      try{
        var holderNext = (aft2 && aft2.isValid && aft2.parentTextFrames && aft2.parentTextFrames.length)
                           ? aft2.parentTextFrames[0] : null;
        if (holderNext && holderNext.isValid){
          tf = holderNext; curTextFrame = holderNext; story = holderNext.parentStory;
        }
      }catch(__){}
    }catch(_){}

    try{
      var _tfLog = (rect.parentTextFrames && rect.parentTextFrames.length) ? rect.parentTextFrames[0] : tf;
      var _pgLog = (_tfLog && _tfLog.isValid) ? _tfLog.parentPage : null;
      var _gbNow = rect.geometricBounds;
      log("[IMGFLOAT6] placed tf="+(_tfLog?_tfLog.id:"NA")+" page="+(_pgLog?_pgLog.name:"NA")+" gb="+[_gbNow[0],_gbNow[1],_gbNow[2],_gbNow[3]].join(","));
    }catch(_){}

    try {
      if (st && st.isValid) st.recompose();
      if (rect && rect.isValid) { try { rect.recompose(); } catch(__){} }
      if (typeof flushOverflow === "function") {
        var fl = flushOverflow(story, page, tf);
        if (fl && fl.frame && fl.page) {
          page  = fl.page;
          tf    = fl.frame;
          story = tf.parentStory;
          curTextFrame = tf;
        }
        try{
          log("[IMG] after.flush  tf=" + (tf&&tf.isValid?tf.id:"NA")
              + " page=" + (page?page.name:"NA")
              + " over(tf)=" + (tf&&tf.isValid?tf.overflows:"NA")
              + " over(curTF)=" + (curTextFrame&&curTextFrame.isValid?curTextFrame.overflows:"NA"));
        }catch(_){}
      }
    } catch(eFlush){ log("[WARN] flush after image: " + eFlush); }

    try{ rect.label="ANCHOR-FLOAT"; }catch(_){}
    try{
      var finalPage = (rect && rect.isValid && rect.parentPage && rect.parentPage.isValid)
        ? rect.parentPage
        : (page && page.isValid ? page : null);
      __recordWordSeqPage(wordSeq, finalPage);
    }catch(_){}
    return rect;
  }catch(e){
    log("[IMGFLOAT6][ERR] "+e);
    return null;
  }
}

function addFloatingFrame(tf, story, page, spec){
  try{
  try{ log("[FRAMEFLOAT] enter id="+(spec&&spec.id)+" textLen="+((spec&&spec.text)||"").length); }catch(_){}
  function _toPtLocal(v){
    var s = String(v==null?"":v).replace(/^\s+|\s+$/g,"");
    if (/mm$/i.test(s)) return parseFloat(s)*2.83464567;
    if (/pt$/i.test(s)) return parseFloat(s);
    if (/px$/i.test(s)) return parseFloat(s)*0.75;
    if (s==="") return 0;
    var n = parseFloat(s); return isNaN(n)?0:n*0.75;
  }
  function _lowerFlag(v){
    if (v == null) return "";
    return String(v).replace(/^\s+|\s+$/g,"").toLowerCase();
  }
  var doc = app.activeDocument;
  var anchorIP = null;
  try{
    if (tf && tf.isValid && tf.insertionPoints && tf.insertionPoints.length){
      anchorIP = tf.insertionPoints[-1];
    }
  }catch(_){}
  if ((!anchorIP || !anchorIP.isValid) && story && story.isValid && story.insertionPoints && story.insertionPoints.length){
    try{ anchorIP = story.insertionPoints[-1]; }catch(_){}
  }
  var hintKey = spec && spec.pageHint;
  var anchorCtx = null;
  try{
    if (__FLOAT_CTX && __FLOAT_CTX.imgAnchors && hintKey && __FLOAT_CTX.imgAnchors[hintKey]){
      anchorCtx = __FLOAT_CTX.imgAnchors[hintKey];
    }
  }catch(_){}
  var anchorWordSeq = null;
  try{
    if (anchorCtx && anchorCtx.wordSeq != null){
      anchorWordSeq = anchorCtx.wordSeq;
    }
  }catch(_){}
  var wordSeq = null;
  try{
    if (spec && spec.wordPageSeq){
      var tmpSeq = parseInt(spec.wordPageSeq, 10);
      if (!isNaN(tmpSeq) && isFinite(tmpSeq)) wordSeq = tmpSeq;
    }
  }catch(_){}
  function _isPageRef(v){
    return !!{"page":1,"pagearea":1,"pageedge":1,"margin":1,"spread":1}[v];
  }
  var posHrefRaw = _lowerFlag(spec && spec.posHref);
  var posVrefRaw = _lowerFlag(spec && spec.posVref);
  var posVRaw = _lowerFlag(spec && spec.posV);
  var wantsSeqAutoPage = false;
  if (wordSeq){
    if (_isPageRef(posHrefRaw) && _isPageRef(posVrefRaw)){
      wantsSeqAutoPage = true;
    } else if (!anchorCtx || anchorWordSeq == null || anchorWordSeq !== wordSeq){
      wantsSeqAutoPage = true;
    }
  }
  var pageFromSeq = null;
  var seqPageWasApplied = false;
  try{
    if (__FLOAT_CTX && __FLOAT_CTX.wordSeqPages && wordSeq){
      var seqCtx = __FLOAT_CTX.wordSeqPages[wordSeq];
      if (seqCtx && seqCtx.page && seqCtx.page.isValid){
        pageFromSeq = seqCtx.page;
      }
    }
  }catch(_){}
  if (!pageFromSeq && wantsSeqAutoPage){
    try{
      pageFromSeq = __pageForWordSeq(wordSeq);
    }catch(_autoSeq){}
  }
  try{
    log("[FRAMEFLOAT][DBG] hintKey=" + (hintKey||"") + " ctxExists=" + (anchorCtx ? "Y":"N"));
    log("[FRAMEFLOAT][SEQ] wordSeq=" + (wordSeq!=null?wordSeq:"NA")
        + " want=" + wantsSeqAutoPage
        + " pageFromSeq=" + (pageFromSeq && pageFromSeq.isValid ? pageFromSeq.name : "NA"));
  }catch(_){}
  function _shiftPage(basePageObj, offset){
    try{
      if (!basePageObj || !basePageObj.isValid) return basePageObj;
      var docRef = basePageObj.parent;
      if (!docRef || !docRef.pages) return basePageObj;
      var targetIndex = basePageObj.documentOffset + offset;
      if (targetIndex < 0) targetIndex = 0;
      while (targetIndex >= docRef.pages.length){
        docRef.pages.add(LocationOptions.AT_END);
      }
      return docRef.pages[targetIndex];
    }catch(_shiftErr){}
    return basePageObj;
  }

  var targetPage = null;
  var enforceHintPage = false;
  try{
    if (page && page.isValid) targetPage = page;
    else if (tf && tf.isValid && tf.parentPage && tf.parentPage.isValid) targetPage = tf.parentPage;
    else if (doc && doc.pages && doc.pages.length) targetPage = doc.pages[0];
  }catch(_){}
  if (pageFromSeq && pageFromSeq.isValid){
    targetPage = pageFromSeq;
    seqPageWasApplied = true;
    enforceHintPage = true;
  }
  if (!pageFromSeq && wordSeq && __FLOAT_CTX){
    try{
      var baseSeq = __FLOAT_CTX.wordSeqBaseSeq;
      var basePage = __FLOAT_CTX.wordSeqBasePage;
      if (baseSeq != null && basePage && basePage.isValid){
        var offset = wordSeq - baseSeq;
        if (offset !== 0){
          var shiftedFromBase = _shiftPage(basePage, offset);
          if (shiftedFromBase && shiftedFromBase.isValid){
            targetPage = shiftedFromBase;
            seqPageWasApplied = true;
            enforceHintPage = true;
            if (!__FLOAT_CTX.wordSeqPages) __FLOAT_CTX.wordSeqPages = {};
            __FLOAT_CTX.wordSeqPages[wordSeq] = {page: shiftedFromBase};
          }
        } else {
          targetPage = basePage;
          seqPageWasApplied = true;
          enforceHintPage = true;
          if (!__FLOAT_CTX.wordSeqPages) __FLOAT_CTX.wordSeqPages = {};
          __FLOAT_CTX.wordSeqPages[wordSeq] = {page: basePage};
        }
      }
    }catch(_baseSeq){}
  }
  var allowAnchorOverride = true;
  if (wordSeq && anchorWordSeq != null && anchorWordSeq !== wordSeq){
    allowAnchorOverride = false;
  }
  if (anchorCtx && anchorCtx.page && anchorCtx.page.isValid && !seqPageWasApplied && allowAnchorOverride){
    targetPage = anchorCtx.page;
    enforceHintPage = true;
  }
  try{
    if (!enforceHintPage && anchorIP && anchorIP.isValid){
      var anchorFrame = anchorIP.parentTextFrames && anchorIP.parentTextFrames.length ? anchorIP.parentTextFrames[0] : null;
      if (anchorFrame && anchorFrame.isValid && anchorFrame.parentPage && anchorFrame.parentPage.isValid){
        targetPage = anchorFrame.parentPage;
      }
    }
  }catch(_){}
  if (!targetPage || !targetPage.isValid){
    try{ log("[FRAMEFLOAT][ERROR] missing valid page"); }catch(_){}
    return null;
  }
  var forceSeqBase = seqPageWasApplied;
  function _calcBounds(){
    function _metrics(pageObj){
      var pb = pageObj.bounds || [0,0,0,0];
      var mp = pageObj.marginPreferences || {};
      var pageTop = pb[0], pageLeft = pb[1], pageBottom = pb[2], pageRight = pb[3];
      var marginTop = parseFloat(mp.top)||0;
      var marginBottom = parseFloat(mp.bottom)||0;
      var marginLeft = parseFloat(mp.left)||0;
      var marginRight = parseFloat(mp.right)||0;
      var innerLeft = pageLeft + marginLeft;
      var innerRight = pageRight - marginRight;
      var innerTop = pageTop + marginTop;
      var innerBottom = pageBottom - marginBottom;
      return {
        pageTop: pageTop,
        pageLeft: pageLeft,
        pageBottom: pageBottom,
        pageRight: pageRight,
        marginTop: marginTop,
        marginBottom: marginBottom,
        marginLeft: marginLeft,
        marginRight: marginRight,
        innerLeft: innerLeft,
        innerRight: innerRight,
        innerTop: innerTop,
        innerBottom: innerBottom,
        innerWidth: Math.max(1, innerRight - innerLeft),
        innerHeight: Math.max(1, innerBottom - innerTop),
        pageWidth: Math.max(1, pageRight - pageLeft),
        pageHeight: Math.max(1, pageBottom - pageTop)
      };
    }

    function _computeBase(metrics){
      var base = {};
      var useInnerH = !!{"page":true,"pagearea":true,"pageedge":true,"margin":true,"spread":true}[posHrefRaw];
      var useInnerV = !!{"page":true,"pagearea":true,"pageedge":true,"margin":true,"spread":true}[posVrefRaw] || posVrefRaw==="paragraph";
      base.useInnerH = useInnerH;
      base.useInnerV = useInnerV;
      base.baseX = useInnerH ? metrics.innerLeft : metrics.pageLeft;
      if (posHrefRaw==="column") base.baseX = metrics.pageLeft + metrics.marginLeft;
      base.baseY = useInnerV ? metrics.innerTop : metrics.pageTop;
      base.maxWidth = useInnerH ? metrics.innerWidth : metrics.pageWidth;
      base.maxHeight = useInnerV ? metrics.innerHeight : metrics.pageHeight;
      return base;
    }

    var metrics = _metrics(targetPage);
    var baseVals = _computeBase(metrics);

    var anchorX = (anchorCtx && anchorCtx.anchorX != null) ? anchorCtx.anchorX : null;
    var anchorY = (anchorCtx && anchorCtx.anchorY != null) ? anchorCtx.anchorY : null;
    try{
      if ((anchorX === null || anchorX === undefined) && anchorIP && anchorIP.isValid){
        anchorX = anchorIP.horizontalOffset;
      }
      if ((anchorY === null || anchorY === undefined) && anchorIP && anchorIP.isValid){
        anchorY = anchorIP.baseline;
      }
    }catch(_){}
    var anchorSourcePage = null;
    try{
      if (anchorCtx && anchorCtx.page && anchorCtx.page.isValid){
        anchorSourcePage = anchorCtx.page;
      } else if (anchorIP && anchorIP.isValid){
        var anchorFrame = (anchorIP.parentTextFrames && anchorIP.parentTextFrames.length) ? anchorIP.parentTextFrames[0] : null;
        if (anchorFrame && anchorFrame.isValid && anchorFrame.parentPage && anchorFrame.parentPage.isValid){
          anchorSourcePage = anchorFrame.parentPage;
        }
      }
    }catch(_){}
    var anchorPageBounds = null;
    if (anchorSourcePage && anchorSourcePage.isValid){
      try{ anchorPageBounds = anchorSourcePage.bounds; }catch(_){}
    }
    var anchorPageTop = (anchorPageBounds && anchorPageBounds.length>=2) ? anchorPageBounds[0] : null;
    var anchorPageLeft = (anchorPageBounds && anchorPageBounds.length>=2) ? anchorPageBounds[1] : null;
    var anchorYOffset = (anchorY != null && anchorPageTop != null) ? (anchorY - anchorPageTop) : null;
    var anchorXOffset = (anchorX != null && anchorPageLeft != null) ? (anchorX - anchorPageLeft) : null;
    function _anchorXForMetrics(m){
      if (anchorXOffset != null) return m.pageLeft + anchorXOffset;
      return anchorX;
    }
    function _anchorYForMetrics(m){
      if (anchorYOffset != null) return m.pageTop + anchorYOffset;
      return anchorY;
    }

    var wPt=_toPtLocal(spec && spec.w);
    var hPt=_toPtLocal(spec && spec.h);
    var offXP=_toPtLocal(spec && spec.offX) || 0;
    var offYP=_toPtLocal(spec && spec.offY) || 0;
    var pageRefMap = {"page":true,"pagearea":true,"pageedge":true,"margin":true,"spread":true};
    var posHrefCalc = posHrefRaw;
    var posVrefCalc = posVrefRaw;
    if (forceSeqBase){
      if (!pageRefMap[posHrefCalc]) posHrefCalc = "page";
      if (!pageRefMap[posVrefCalc]) posVrefCalc = "page";
    }
    var srcPageWidth = _toPtLocal(spec && spec.wordPageWidth);
    var srcPageHeight = _toPtLocal(spec && spec.wordPageHeight);
    function _destWidthFor(ref, m){
      if (ref === "margin" || ref === "column") return m.innerWidth;
      return m.pageWidth;
    }
    function _destHeightFor(ref, m){
      if (ref === "margin" || ref === "column") return m.innerHeight;
      return m.pageHeight;
    }
    var destWidth = _destWidthFor(posHrefCalc, metrics);
    var destHeight = _destHeightFor(posVrefCalc, metrics);
    if (srcPageWidth && srcPageWidth > 0 && destWidth){
      offXP = offXP * (destWidth / srcPageWidth);
    }
    if (srcPageHeight && srcPageHeight > 0 && destHeight){
      offYP = offYP * (destHeight / srcPageHeight);
    }
    var anchorXEff = _anchorXForMetrics(metrics);
    var anchorYEff = _anchorYForMetrics(metrics);

    function _areaBounds(ref, m){
      if (ref === "margin" || ref === "column"){
        return {left: m.innerLeft, right: m.innerRight, top: m.innerTop, bottom: m.innerBottom};
      }
      return {left: m.pageLeft, right: m.pageRight, top: m.pageTop, bottom: m.pageBottom};
    }

    function _recomputeBase(){
      baseVals = _computeBase(metrics);
    }

    var horizArea = _areaBounds(posHrefCalc, metrics);
    var vertArea = _areaBounds(posVrefCalc, metrics);

    var baseX = horizArea.left;
    var baseY = vertArea.top;
    if (posVrefCalc === "paragraph" && anchorYEff !== null){
      baseY = anchorYEff;
    }
    if (posHrefCalc === "paragraph" && anchorXEff !== null){
      baseX = anchorXEff;
    }

    var maxWidth = Math.max(1, horizArea.right - horizArea.left);
    var maxHeight = Math.max(1, vertArea.bottom - vertArea.top);
    var targetW = wPt>0 ? Math.min(wPt, maxWidth) : maxWidth;
    var targetH = hPt>0 ? Math.min(hPt, maxHeight) : maxHeight;
    if (targetW <= 0) targetW = maxWidth;
    if (targetH <= 0) targetH = maxHeight;

    var left = baseX + offXP;
    var top = baseY + offYP;
    var pageHeight = metrics.pageHeight;
    var relativeTop = (baseVals.baseY + offYP) - metrics.pageTop;
    var pageShift = 0;
    if (relativeTop >= pageHeight - 0.5){
      pageShift = Math.floor(relativeTop / pageHeight);
    }
    if (pageShift > 0){
      var shiftedPage = _shiftPage(targetPage, pageShift);
      if (shiftedPage && shiftedPage.isValid){
        targetPage = shiftedPage;
        metrics = _metrics(targetPage);
        horizArea = _areaBounds(posHrefCalc, metrics);
        vertArea = _areaBounds(posVrefCalc, metrics);
        baseVals = _computeBase(metrics);
        pageHeight = metrics.pageHeight;
        offYP = offYP - pageShift * pageHeight;
        baseX = horizArea.left;
        baseY = vertArea.top;
        anchorXEff = _anchorXForMetrics(metrics);
        anchorYEff = _anchorYForMetrics(metrics);
        if (posVrefCalc === "paragraph" && anchorYEff !== null){
          baseY = anchorYEff;
        }
        if (posHrefCalc === "paragraph" && anchorXEff !== null){
          baseX = anchorXEff;
        }
        maxWidth = Math.max(1, horizArea.right - horizArea.left);
        maxHeight = Math.max(1, vertArea.bottom - vertArea.top);
        targetW = wPt>0 ? Math.min(wPt, maxWidth) : maxWidth;
        targetH = hPt>0 ? Math.min(hPt, maxHeight) : maxHeight;
        left = baseX + offXP;
        top = baseY + offYP;
      }
    }

    var maxRight = horizArea.right;
    var maxBottom = vertArea.bottom;

    if (left < metrics.pageLeft) left = metrics.pageLeft;
    if (top < metrics.pageTop) top = metrics.pageTop;
    if (left + targetW > maxRight) {
      if ((maxRight - targetW) >= metrics.pageLeft) left = maxRight - targetW;
      if (left < metrics.pageLeft) left = Math.max(metrics.pageLeft, maxRight - targetW);
    }
    if (top + targetH > maxBottom) {
      if ((maxBottom - targetH) >= metrics.pageTop) top = maxBottom - targetH;
      if (top < metrics.pageTop) top = Math.max(metrics.pageTop, maxBottom - targetH);
    }
    var right = left + targetW;
    var bottom = top + targetH;
    return [top, left, bottom, right];
  }
  var gbFrame = _calcBounds();
  if (!gbFrame) return null;
  try{
    log("[FRAMEFLOAT][DBG] page=" + (targetPage?targetPage.name:"NA") +
        " bounds=" + gbFrame.join(",") + " off=(" + (spec && spec.offX) + "," + (spec && spec.offY) + ")");
  }catch(_){}
  var frame = targetPage.textFrames.add();
  frame.geometricBounds = gbFrame;
  try{
    var contentText = spec && spec.text ? String(spec.text) : "";
    if (typeof smartWrapStr === "function") contentText = smartWrapStr(contentText);
    frame.contents = contentText;
  }catch(_setFrame){
    try{
      if (frame.insertionPoints && frame.insertionPoints.length){
        frame.insertionPoints[-1].contents = spec.text || "";
      }
    }catch(__){}
  }
  try{
    var insetTop = _toPtLocal(spec && spec.bodyInsetT);
    var insetLeft = _toPtLocal(spec && spec.bodyInsetL);
    var insetBottom = _toPtLocal(spec && spec.bodyInsetB);
    var insetRight = _toPtLocal(spec && spec.bodyInsetR);
    frame.textFramePreferences.insetSpacing = [
      isFinite(insetTop)?insetTop:0,
      isFinite(insetLeft)?insetLeft:0,
      isFinite(insetBottom)?insetBottom:0,
      isFinite(insetRight)?insetRight:0
    ];
  }catch(_){}
  try{
    __recordWordSeqPage(wordSeq, targetPage);
  }catch(_){}
  function _applyFrameStyles(frameObj){
    var appliedParagraph = false;
    if (!frameObj || !frameObj.isValid) return;
    try{
      var ps = app.activeDocument.paragraphStyles.itemByName("SidebarLabel");
      if (ps && ps.isValid){
        frameObj.parentStory.paragraphs.everyItem().appliedParagraphStyle = ps;
        appliedParagraph = true;
      }
    }catch(_){}
    if (appliedParagraph){
      try{
        var cs = app.activeDocument.characterStyles.itemByName("SidebarLabel-Char");
        if (cs && cs.isValid){
          frameObj.parentStory.characters.everyItem().appliedCharacterStyle = cs;
        }
      }catch(_){}
    } else {
      try{
        frameObj.parentStory.paragraphs.everyItem().pointSize = 8;
        frameObj.parentStory.paragraphs.everyItem().leading = 10;
      }catch(_defaultSize){}
    }
    try{
      var tfp = frameObj.textFramePreferences;
      tfp.autoSizingReferencePoint = AutoSizingReferenceEnum.TOP_LEFT_POINT;
      tfp.autoSizingType = AutoSizingTypeEnum.HEIGHT_ONLY;
    }catch(_autoSize){}
  }
  try{ _applyFrameStyles(frame); }catch(_styleErr){}
  try{
    var wrapKey = _lowerFlag((spec && spec.wrap) || spec.bodyWrap);
    var wrapMode = TextWrapModes.NONE;
    if (wrapKey === "wrapsquare" || wrapKey === "square"){
      wrapMode = TextWrapModes.BOUNDING_BOX_TEXT_WRAP;
    } else if (wrapKey === "wraptight" || wrapKey === "tight" || wrapKey === "wrapthrough"){
      wrapMode = TextWrapModes.OBJECT_SHAPE_TEXT_WRAP;
    } else if (wrapKey === "wraptopbottom" || wrapKey === "topbottom"){
      wrapMode = TextWrapModes.JUMP_OBJECT_TEXT_WRAP;
    } else if (wrapKey === "wrapbehind"){
      wrapMode = TextWrapModes.NONE;
    }
    frame.textWrapPreferences.textWrapMode = wrapMode;
    if (wrapMode !== TextWrapModes.NONE){
      var distT = _toPtLocal(spec && spec.distT) || 0;
      var distB = _toPtLocal(spec && spec.distB) || 0;
      var distL = _toPtLocal(spec && spec.distL);
      var distR = _toPtLocal(spec && spec.distR);
      if (!distL && distL !== 0) distL = 12;
      if (!distR && distR !== 0) distR = 12;
      frame.textWrapPreferences.textWrapOffset = [distT, distL, distB, distR];
    }
  }catch(_wrapErr){}
  try{ frame.label = spec && spec.id ? String(spec.id) : ""; }catch(_){}
  return frame;
  }catch(eFrame){
    try{ log("[FRAMEFLOAT][EXC] " + eFrame); }catch(_){}
    throw eFrame;
  }
}
function _alignFloatingRect(rect, holder, innerW, alignMode){
  if (!rect || !holder || !holder.isValid || innerW <= 0) return null;
  var gb = rect.geometricBounds;
  if (!gb || gb.length !== 4) return null;
  var targetW = gb[3] - gb[1];
  if (targetW <= 0) return null;
  var inset = holder.textFramePreferences && holder.textFramePreferences.insetSpacing;
  var leftBase = (holder.geometricBounds && holder.geometricBounds.length === 4) ? holder.geometricBounds[1] : 0;
  if (inset && inset.length >= 2) leftBase += inset[1];
  var space = Math.max(0, innerW - targetW);
  var offset = 0;
  if (alignMode === "right") offset = space;
  else if (alignMode === "center") offset = space / 2;
  var newLeft = leftBase + offset;
  rect.geometricBounds = [gb[0], newLeft, gb[2], newLeft + targetW];
  return {holder: holder, innerW: innerW, align: alignMode, offset: offset};
}

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


    function findParaStyleCI(doc, name){
        var lower = String(name).toLowerCase();
        var ps = doc.paragraphStyles;
        for (var i=0;i<ps.length;i++){
            try{
                if (String(ps[i].name).toLowerCase() === lower) return ps[i];
            }catch(_){}
        }
        return null;
    }
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

    // === 行内样式应用（保持你原逻辑，只保留下划线） ===
    // 递归搜索字符样式（支持样式组），大小写与空格/下划线不敏感
    function findCharStyleCI(doc, name){
      function norm(n){ return String(n||"").toLowerCase().replace(/\s+/g,"").replace(/[_-]/g,""); }
      var target = norm(name);

      // 先扫顶层
      var cs = doc.characterStyles;
      for (var i=0;i<cs.length;i++){
        try{ if (norm(cs[i].name) === target) return cs[i]; }catch(_){}
      }

      // 再扫样式组（递归）
      function scanGroup(g){
        try{
          var arr = g.characterStyles;
          for (var i=0;i<arr.length;i++){ try{ if (norm(arr[i].name)===target) return arr[i]; }catch(_){ } }
          var subs = g.characterStyleGroups;
          for (var j=0;j<subs.length;j++){ var hit = scanGroup(subs[j]); if (hit) return hit; }
        }catch(_){}
        return null;
      }
      try{
        var groups = doc.characterStyleGroups;
        for (var k=0;k<groups.length;k++){ var hit = scanGroup(groups[k]); if (hit) return hit; }
      }catch(_){}

      return null;
    }

    // 懒加载 + 缓存，避免在还没打开文档时访问 activeDocument
    function getCachedCharStyleByList(names){
        try{
            if (app.documents.length === 0) return null; // 还没打开任何文档就别取
            var doc = app.activeDocument;
            if (!doc || !doc.isValid) return null;
            if (!app._csCache) app._csCache = {};
            for (var k=0;k<names.length;k++){
                var key = String(names[k]).toLowerCase();
                var cs = app._csCache[key];
                if (cs && cs.isValid) return cs;
                cs = findCharStyleCI(doc, names[k]);
                if (cs && cs.isValid) { app._csCache[key] = cs; return cs; }
            }
        }catch(e){}
        return null;
    }

    function _fontInfo(r){
      var fam="NA", sty="NA";
      try{ fam = String(r.appliedFont.name || r.appliedFont.family || r.appliedFont); }catch(_){}
      try{ sty = String(r.fontStyle); }catch(_){}
      var tI="NA", tB="NA";
      try{ tI = String(!!r.trueItalic); }catch(_){}
      try{ tB = String(!!r.trueBold); }catch(_){}
      return "font="+fam+" ; style="+sty+" ; trueItalic="+tI+" ; trueBold="+tB;
    }

    function setItalicSafe(r){
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

    function setBoldSafe(r){
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

    function applyInlineFormattingOnRange(story, startCharIndex, endCharIndex, st) {
      try {
        if (endCharIndex <= startCharIndex) return;
        var r = story.characters.itemByRange(startCharIndex, endCharIndex - 1);
        var txt=""; try{ txt = String(r.contents).substr(0,50); }catch(_){}
        log("[I/B/U] range="+startCharIndex+"-"+endCharIndex+" ; flags="+_s(st)+" ; txt=\""+txt+"\"");

        try { r.underline = !!st.u; log("[U] set="+ (!!st.u)); } catch(eu){ log("[U][ERR] "+eu); }

        if (st.i) {
          try { var howI = setItalicSafe(r); log("[I] via " + howI + " ; " + _fontInfo(r)); } catch(ei){ log("[I][ERR] "+ei); }
        }
        if (st.b) {
          try { var howB = setBoldSafe(r);   log("[B] via " + howB + " ; " + _fontInfo(r)); } catch(eb){ log("[B][ERR] "+eb); }
        }
      } catch(e) {
        log("[IBU][ERR] "+e);
      }
    }

    function _innerFrameWidth(frame){
        if (!frame || !frame.isValid) return 0;
        try{
            var tfp = frame.textFramePreferences;
            if (tfp){
                var cw = parseFloat(tfp.textColumnFixedWidth);
                if (isFinite(cw) && cw > 0) return cw;
            }
        }catch(_){}
        var gb = null;
        try{ gb = frame.geometricBounds; }catch(_){}
        var w = (gb && gb.length>=4) ? (gb[3]-gb[1]) : 0;
        var inset = null;
        try{ inset = frame.textFramePreferences.insetSpacing; }catch(_){}
        var insetWidth = (inset && inset.length>=4) ? ( (parseFloat(inset[1])||0) + (parseFloat(inset[3])||0) ) : 0;
        return w - insetWidth;
    }
    function _mapAlign(h){ if(h=="center") return Justification.CENTER_ALIGN; if(h=="right") return Justification.RIGHT_ALIGN; return Justification.LEFT_ALIGN; }
    function _mapVAlign(v){ if(v=="bottom") return VerticalJustification.BOTTOM_ALIGN; if(v=="center"||v=="middle") return VerticalJustification.CENTER_ALIGN; return VerticalJustification.TOP_ALIGN; }

    function applyTableBorders(tbl, opts){
        try{
            opts = opts || {};
            var outerOn  = (opts.outerOn  !== false);
            var innerHOn = (opts.innerHOn !== false);
            var innerVOn = (opts.innerVOn !== false);

            if (typeof opts.inner === "number" && typeof opts.innerWeight !== "number") opts.innerWeight = opts.inner;
            if (typeof opts.outer === "number" && typeof opts.outerWeight !== "number") opts.outerWeight = opts.outer;

            var outerWeight = (typeof opts.outerWeight === "number") ? opts.outerWeight : 0.75;
            var innerWeight = (typeof opts.innerWeight === "number") ? opts.innerWeight : 0.5;
            var colorHex    = (opts.color || "#000000");
            var cellInset   = (typeof opts.cellInset === "number") ? opts.cellInset : null;

            function getColorByHex(hex){
                try{
                    if(!/^#([0-9A-Fa-f]{6})$/.test(hex)) return app.activeDocument.swatches.item("Black");
                    var name = "Stroke_"+hex.substr(1);
                    var col = app.activeDocument.colors.itemByName(name);
                    if(!col.isValid){
                        col = app.activeDocument.colors.add({
                            name:name, model:ColorModel.PROCESS, space:ColorSpace.RGB,
                            colorValue:[
                                parseInt(hex.substr(1,2),16),
                                parseInt(hex.substr(3,2),16),
                                parseInt(hex.substr(5,2),16)
                            ]
                        });
                    }
                    return col;
                }catch(e){ return app.activeDocument.swatches.item("Black"); }
            }
            var strokeCol = getColorByHex(colorHex);

            var rows = tbl.rows.length, cols = tbl.columns.length;
            var allCells = tbl.cells.everyItem();
            try{
                allCells.strokeWeight = innerWeight;
                allCells.strokeColor  = strokeCol;

                allCells.topStrokeWeight    = innerWeight;
                allCells.bottomStrokeWeight = innerWeight;
                allCells.leftStrokeWeight   = innerWeight;
                allCells.rightStrokeWeight  = innerWeight;

                allCells.topStrokeColor     = strokeCol;
                allCells.bottomStrokeColor  = strokeCol;
                allCells.leftStrokeColor    = strokeCol;
                allCells.rightStrokeColor   = strokeCol;

                allCells.topEdgeStrokeWeight    = innerWeight;
                allCells.bottomEdgeStrokeWeight = innerWeight;
                allCells.leftEdgeStrokeWeight   = innerWeight;
                allCells.rightEdgeStrokeWeight  = innerWeight;

                allCells.topEdgeStrokeColor     = strokeCol;
                allCells.bottomEdgeStrokeColor  = strokeCol;
                allCells.leftEdgeStrokeColor    = strokeCol;
                allCells.rightEdgeStrokeColor   = strokeCol;

                if (cellInset !== null){
                    allCells.topInset    = cellInset;
                    allCells.bottomInset = cellInset;
                    allCells.leftInset   = cellInset;
                    allCells.rightInset  = cellInset;
                }
            }catch(_){}

            if(!innerHOn){
                for(var r=0; r<rows-1; r++){
                    try{
                        var cells = tbl.rows[r].cells.everyItem();
                        cells.bottomStrokeWeight = 0;
                        cells.bottomEdgeStrokeWeight = 0;
                    }catch(_){}
                }
            }
            if(!innerVOn){
                for(var c=0; c<cols-1; c++){
                    try{
                        var cc = tbl.columns[c].cells.everyItem();
                        cc.rightStrokeWeight = 0;
                        cc.rightEdgeStrokeWeight = 0;
                    }catch(_){}
                }
            }

            var topRow    = tbl.rows[0];
            var bottomRow = tbl.rows[rows-1];
            var leftCol   = tbl.columns[0];
            var rightCol  = tbl.columns[cols-1];

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
                    for(var rr=0; rr<Math.min(tbl.headerRowCount, rows); rr++){
                        var row = tbl.rows[rr];
                        var cells = row.cells.everyItem();
                        cells.bottomStrokeWeight = w;
                        cells.bottomEdgeStrokeWeight = w;
                        cells.bottomStrokeColor  = strokeCol;
                        cells.bottomEdgeStrokeColor  = strokeCol;
                    }
                }catch(_){}
            }
        }catch(e){ try{ log("[DBG] applyTableBorders: "+e); }catch(__){} }
    }

    function _normalizeTableWidth(tbl){
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
        }catch(e){ try{ log(__tableWarnTag + " _normalizeTableWidth: "+e); }catch(__){} }
    }


                        function createEndnoteAt(ip, content, idForDisplay){
        if(!ip || !ip.isValid) return null;
        var doc = app.activeDocument, story = ip.parentStory;
        var en = null, ok = false;
        try { if (ip.createEndnote) { en = ip.createEndnote(); ok = (en && en.isValid); } } catch(e){ }
        if (!ok) { try { en = story.endnotes.add(ip); ok = (en && en.isValid); } catch(e){ } }
        if (!ok) { try { en = doc.endnotes.add(ip);   ok = (en && en.isValid); } catch(e){ } }
        if (!ok) {
            try{ log("[NOTE][EN][ERR] unable to create endnote"); }catch(_){}
            return null;
        }
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

    // —— 段落插入：扩展识别 [[IMG ...]] / [[TABLE {...}]] ——
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
        // ★ 正则扩展：新增 IMG/TABLE（修复 I/B/U 与 IMG/TABLE 的匹配）
        var re = /\[{2,}FNI:(\d+)\]{2,}|\[{2,}(FN|EN):(.*?)\]{2,}|\[\[(\/?)(I|B|U)\]\]|\[\[IMG\s+([^\]]+)\]\]|\[\[TABLE\s+(\{[\s\S]*?\})\]\]/g;
        var last = 0, m;
        var st = {i:0, b:0, u:0};
        var PENDING_NOTE_ID = null;
        function on(x){ return x>0; }

        while ((m = re.exec(text)) !== null) {
            var chunk = text.substring(last, m.index);
            if (chunk.length) {
                var startIdx = story.characters.length;
                story.insertionPoints[-1].contents = chunk;
                var endIdx   = story.characters.length;
                applyInlineFormattingOnRange(story, startIdx, endIdx, {i:on(st.i), b:on(st.b), u:on(st.u)});
            }
            try { story.insertionPoints[-1].appliedCharacterStyle = app.activeDocument.characterStyles.itemByName("[None]"); } catch(_){ try { story.insertionPoints[-1].appliedCharacterStyle = app.activeDocument.characterStyles[0]; } catch(__){} }


            if (m[1]) {
                PENDING_NOTE_ID = parseInt(m[1], 10);
            } else if (m[2]) {
                var noteType = m[2];
                var noteContent = m[3];
                var ip = story.insertionPoints[-1];
                try {
                      log("[NOTE] create " + noteType + " id=" + PENDING_NOTE_ID + " len=" + (noteContent||"").length);
                      if (noteType === "FN") createFootnoteAt(ip, noteContent, PENDING_NOTE_ID);
                      else createEndnoteAt(ip, noteContent, PENDING_NOTE_ID);
                } catch(e){ log("[NOTE][ERR] " + e); }
                PENDING_NOTE_ID = null;

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
                // 调紧默认前后距，便于两图紧凑排布；可被 XML 显式覆盖
                if (spec.spaceBefore == null) spec.spaceBefore = 0;
                if (spec.spaceAfter  == null) spec.spaceAfter  = 2;
                if (!spec.wrap) spec.wrap = "none"; // ← 默认不绕排，避免把后文推到文末

                // 关键修正 A：确保插入点在“当前末尾文本框”——先疏通 overset，再取就地安全 IP
                // —— 诊断日志：放图前记录“末尾插入点所在文本框/页 & overset”信息
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
                  // 先尝试疏通（保持原有策略）
                  if (typeof flushOverflow === "function" && typeof tf !== "undefined" && tf && tf.isValid) {
                    var _rs = flushOverflow(story, page, tf);
                    if (_rs && _rs.frame && _rs.page) { page = _rs.page; tf = _rs.frame; story = tf.parentStory; curTextFrame = tf; }
                  }
                  // 再以“story 末尾”的父文本框为准强制刷新 tf/curTextFrame（避免仍指向上一个框）
                  // 再以“story 末尾”作为锚点候选，记录一次定位信息
                  try{
                      var _ipEnd = story.insertionPoints[-1];
                      var _holder = (_ipEnd && _ipEnd.isValid && _ipEnd.parentTextFrames && _ipEnd.parentTextFrames.length)
                                      ? _ipEnd.parentTextFrames[0] : null;
                      if (_holder && _holder.isValid) {
                        tf = _holder;                     // ← 强制把“当前活动文本框”切到 story 实际末尾的文本框
                        curTextFrame = _holder;           // ← 同步全局引用，后续 _safeIP/列宽计算都用这个
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
                // 若当前不在段首（上一字符不是回车），补一个段落结束，保证每张图独占一段
                try {
                  var lastChar = (story.characters.length>0) ? String(story.characters[-1].contents||"") : "";
                  if (lastChar !== "\r") story.insertionPoints[-1].contents = "\r";
                } catch(__){}

                // 插入点：就用上面刷新后的 tf 的末尾；兜底再回退 story 尾（仅加日志）
                var ipNow = (tf && tf.isValid) ? tf.insertionPoints[-1] : story.insertionPoints[-1];
                try{
                  var __h = (ipNow && ipNow.isValid && ipNow.parentTextFrames && ipNow.parentTextFrames.length) ? ipNow.parentTextFrames[0] : null;
                  var __pg = (__h && __h.isValid) ? __h.parentPage : null;
                  log("[IMG-LOC][ipNow] frame=" + (__h?__h.id:"NA") + " page=" + (__pg?__pg.name:"NA")
                      + " ; ip.index=" + (ipNow&&ipNow.isValid?ipNow.index:"NA"));
                }catch(_){}

                // 规范与校验路径（失败只记一行，不抛）
                var fsrc = _normPath(spec.src);
                if (fsrc && fsrc.exists) {
                  spec.src = fsrc.fsName;
                  // 入口调用加一层必要 try，避免整套流程被图片单点中断
                  try {
                    // 规范与校验路径（失败只记一行，不抛）
                    var fsrc = _normPath(spec.src);
                    if (fsrc && fsrc.exists) {
                      spec.src = fsrc.fsName;

                      // △ 根据 XML：inline="1" → 内联锚定；inline="0" → 浮动定位
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
                          // 浮动：使用刚加入的 addFloatingImage（遵循 posH/posV/offX/offY/wrap/dist*）
                          var rect = addFloatingImage(tf, story, page, spec);
                          if (rect && rect.isValid) log("[IMG] ok (float): " + spec.src);
                        } else {
                          // 内联：仍走你原先的稳妥链路（addImageAtV2）
                          var rect = addImageAtV2(ipNow, spec);
                          if (rect && rect.isValid) log("[IMG] ok (inline): " + spec.src);
                        }
                      } catch(e) {
                        log("[ERR] addImage dispatch failed: " + e);
                      }
                    } else {
                      log("[IMG] missing: " + spec.src);
                    }
                    // 可选：成功才轻量记一行
                    if (rect && rect.isValid) log("[IMG] ok: " + spec.src);
                  } catch(e) {
                    log("[ERR] addImageAt failed: " + e);
                  }
                } else {
                  log("[IMG] missing: " + spec.src);
                }

                try { story.insertionPoints[-1].appliedCharacterStyle = app.activeDocument.characterStyles.itemByName("[None]"); } catch(_){ try { story.insertionPoints[-1].appliedCharacterStyle = app.activeDocument.characterStyles[0]; } catch(__){} }
                // 吃掉 [[IMG ...]]，继续
                last = re.lastIndex;
                continue;
            } else if (m[7]) {
                try {
                    var obj = JSON.parse(m[7]);
                    // 使用高保真表格构造：按 colWidthsPt 设置列宽、处理合并/覆盖格
                    addTableHiFi(obj);
                } catch(e){
                    try { var obj2 = eval("("+m[7]+")"); addTableHiFi(obj2); } catch(__){}
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
            applyInlineFormattingOnRange(story, sIdx, eIdx, {i:on(st.i), b:on(st.b), u:on(st.u)});
        }
        try { story.insertionPoints[-1].appliedCharacterStyle = app.activeDocument.characterStyles.itemByName("[None]"); } catch(_){ try { story.insertionPoints[-1].appliedCharacterStyle = app.activeDocument.characterStyles[0]; } catch(__){} }


        story.insertionPoints[-1].contents = "\r";
        story.paragraphs[-1].appliedParagraphStyle = s;
        try {
            story.recompose(); app.activeDocument.recompose();
        } catch(_){}
        // 避免长段堆积造成中途 overset：每写 N 段疏通一次
        try {
            if (typeof __paraCounter === "undefined") __paraCounter = 0;
            __paraCounter++;
            if ((__paraCounter % 50) === 0) {
                var st = flushOverflow(story, page, tf);
                page  = st.page;
                tf    = st.frame;
                story = tf.parentStory;
                curTextFrame = tf;              // ★ 新增：切到新框后更新全局指针
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

    // 打开模板、清空页面框等（保持你原逻辑）
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

    // 清空页面与母版文本框，保留第一页
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


    // 简易样式兜底（保持你原逻辑）
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

    function frameBoundsForPage2(page, doc) {
        var pb = page.bounds, mp = page.marginPreferences;
        return [pb[0] + mp.top, pb[1] + mp.left, pb[2] - mp.bottom, pb[3] - mp.right];
    }

    function _innerFrameHeight(frame){
        if (!frame || !frame.isValid) return 0;
        var gb = null;
        try{ gb = frame.geometricBounds; }catch(_){}
        var h = (gb && gb.length>=4) ? (gb[2]-gb[0]) : 0;
        var inset = null;
        try{ inset = frame.textFramePreferences.insetSpacing; }catch(_){}
        var insetHeight = (inset && inset.length>=4) ? ( (parseFloat(inset[0])||0) + (parseFloat(inset[2])||0) ) : 0;
        return h - insetHeight;
    }

    function __calcInnerWidthForLayout(layout){
        if (!layout) return null;
        var w = parseFloat(layout.pageWidthPt);
        var margins = layout.pageMarginsPt || {};
        if (isFinite(w)){
            var left = parseFloat(margins.left) || 0;
            var right = parseFloat(margins.right) || 0;
            return w - left - right;
        }
        return null;
    }

    function __applyFrameLayout(frame, layoutState){
        try{
            if (!frame || !frame.isValid) return;
            var tfp = frame.textFramePreferences;
            if (!tfp) return;
            try{ tfp.textColumnCount = 1; }catch(_){}
            try{ tfp.textColumnGutter = 0; }catch(_){}
            try{ tfp.useFixedColumnWidth = true; }catch(_){}
            try{ tfp.textColumnFlexibleWidth = false; }catch(_){}
            var leftInset = 0, rightInset = 0;
            try{
                var inset = tfp.insetSpacing;
                if (inset && inset.length >= 4){
                    leftInset = parseFloat(inset[1]) || 0;
                    rightInset = parseFloat(inset[3]) || 0;
                }
            }catch(_){}
            var gb = null;
            try{ gb = frame.geometricBounds; }catch(_){}
            var innerWidth = null;
            if (gb && gb.length >= 4){
                innerWidth = (gb[3] - gb[1]) - leftInset - rightInset;
            }
            if (!isFinite(innerWidth) || innerWidth <= 0){
                var pageWidth = layoutState && layoutState.pageWidthPt;
                var margins = layoutState && layoutState.pageMarginsPt;
                if (isFinite(pageWidth)){
                    innerWidth = pageWidth;
                    if (margins){
                        innerWidth -= (parseFloat(margins.left) || 0);
                        innerWidth -= (parseFloat(margins.right) || 0);
                    }
                }
            }
            if (!isFinite(innerWidth) || innerWidth <= 0){
                innerWidth = 400;
            }
            try{ tfp.textColumnFixedWidth = innerWidth; }catch(_){}
            try{ log("[LAYOUT] apply frame id=" + frame.id + " innerWidth=" + innerWidth + " orient=" + (layoutState && layoutState.pageOrientation)); }catch(_log){}
        }catch(_){}
    }

    function createTextFrameOnPage(page, layoutState) {
        var gb = frameBoundsForPage2(page, doc);
        var tf = page.textFrames.add();
        tf.geometricBounds = gb;
        __applyFrameLayout(tf, layoutState || __CURRENT_LAYOUT);
        try{ tf.textFramePreferences.verticalJustification = VerticalJustification.TOP_ALIGN; }catch(_){}
        return tf;
    }

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
    try{ __progressFinalize(); }catch(_){}
                  // ★ 新增：切到新框后更新全局指针

        // （可选）导出 IDML
        if (%AUTO_EXPORT%) {
            try {
                var outFile = File("%OUT_IDML%");
                doc.exportFile(ExportFormat.INDESIGN_MARKUP, outFile, false);
            } catch(ex) { alert("导出 IDML 失败：" + ex); }
        }
        try{
            if (__origScriptUnit !== null) app.scriptPreferences.measurementUnit = __origScriptUnit;
        }catch(_){}
        try{
            if (__origViewH !== null) app.viewPreferences.horizontalMeasurementUnits = __origViewH;
            if (__origViewV !== null) app.viewPreferences.verticalMeasurementUnits = __origViewV;
        }catch(_){}
    })();

function fixAllTables(){
    try{
        var doc = app.activeDocument;
        var stories = doc.stories;
        for (var si=0; si<stories.length; si++){
            var tbs = stories[si].tables;
            for (var ti=0; ti<tbs.length; ti++){
                var T = tbs[ti];
                try { T.rows.everyItem().autoGrow = true; } catch(_){}
                try { T.rows.everyItem().height = RowAutoHeight.AUTO; } catch(_){}
                try { T.rows.everyItem().heightType = RowHeightType.AT_LEAST; } catch(_){}
                try { T.rows.everyItem().minimumHeight = 0; } catch(_){}
                try { T.rows.everyItem().maximumHeight = 1000000; } catch(_){}
                try { T.rows.everyItem().keepWithNext = false; } catch(_){}
                try { T.rows.everyItem().keepTogether = false; } catch(_){}
                try {
                    var paras = T.cells.everyItem().texts[0].paragraphs.everyItem();
                    paras.keepOptions.keepLinesTogether = false;
                    paras.keepOptions.keepWithNext = false;
                    try { paras.composer = ComposerTypes.ADOBE_PARAGRAPH_COMPOSER; } catch(_){}
                    try { paras.spaceBefore = Math.min( paras.spaceBefore, 6 ); } catch(_){}
                    try { paras.spaceAfter  = Math.min( paras.spaceAfter,  6 ); } catch(_){}
                } catch(_){}
                try { T.recompose(); } catch(_){}
            }
        }
        try { log("[LOG] fixAllTables done"); } catch(_){}
    }catch(e){ try{ log("[DBG] fixAllTables: "+e); }catch(__){} }
}
