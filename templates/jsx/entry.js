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
    var LOG_WRITE  = (CONFIG && CONFIG.flags && typeof CONFIG.flags.logWrite === "boolean")
                     ? CONFIG.flags.logWrite : %LOG_WRITE%;   // true=记录 debug；false=仅保留 warn/error/info
    var __EVENT_LINES = [];

    // 每次执行先清空旧事件日志，避免多次运行叠加
    try{
      if (EVENT_FILE){
        EVENT_FILE.encoding = "UTF-8";
        EVENT_FILE.open("w");
        EVENT_FILE.writeln(""); // 写一空行确保文件被截断创建
        EVENT_FILE.close();
      }
    }catch(_){}

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
                var fsrc = __imgNormPath(spec.src);
                if (fsrc && fsrc.exists) {
                  spec.src = fsrc.fsName;
                  // 入口调用加一层必要 try，避免整套流程被图片单点中断
                  try {
                    // 规范与校验路径（失败只记一行，不抛）
                    var fsrc = __imgNormPath(spec.src);
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
                          // 浮动：使用刚加入的 __imgAddFloatingImage（遵循 posH/posV/offX/offY/wrap/dist*）
                          var rect = __imgAddFloatingImage(tf, story, page, spec);
                          if (rect && rect.isValid) log("[IMG] ok (float): " + spec.src);
                        } else {
                          // 内联：仍走你原先的稳妥链路（__imgAddImageAtV2）
                          var rect = __imgAddImageAtV2(ipNow, spec);
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

    // ?????? IDML
    var __AUTO_EXPORT = (CONFIG && CONFIG.flags && typeof CONFIG.flags.autoExportIdml === "boolean")
                        ? CONFIG.flags.autoExportIdml : %AUTO_EXPORT%;
    if (__AUTO_EXPORT) {
        try {
            var outFile = File("%OUT_IDML%");
            doc.exportFile(ExportFormat.INDESIGN_MARKUP, outFile, false);
        } catch(ex) { alert("?? IDML ???" + ex); }
    }
    try{
        if (__origScriptUnit !== null) app.scriptPreferences.measurementUnit = __origScriptUnit;
    }catch(_){ }
    try{
        if (__origViewH !== null) app.viewPreferences.horizontalMeasurementUnits = __origViewH;
        if (__origViewV !== null) app.viewPreferences.verticalMeasurementUnits = __origViewV;
    }catch(_){ }
