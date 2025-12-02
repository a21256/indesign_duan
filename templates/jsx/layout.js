// Layout state (shared inside template IIFE scope)
var __DEFAULT_LAYOUT = null;
var __CURRENT_LAYOUT = null;
var __DEFAULT_INNER_WIDTH = null;
var __DEFAULT_INNER_HEIGHT = null;
var __ENABLE_TRAILING_TRIM = false;
// unified layout log prefix helper
function __layoutTag(type, msg){
  var t = "[LAYOUT]";
  if (type === "warn") t = "[WARN]";
  else if (type === "err" || type === "error") t = "[ERR]";
  return t + (msg ? " " + msg : "");
}
// basic self-check to ensure required helpers exist
function __layoutSelfCheck(){
  try{
    var required = ["__cloneLayoutState","__ensureLayout","__createLayoutFrame","frameBoundsForPage2"];
    for (var i=0;i<required.length;i++){
      var n = required[i];
      if (typeof eval(n) !== "function") throw ("missing function: " + n);
    }
  }catch(e){
    try{ log("[ERR] LAYOUT selfcheck failed: " + e); }catch(_){}
    throw e;
  }
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
function __layoutNormalizeTarget(target){
  if (!target) return target;
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
  return target;
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
    // fill defaults
    target = __layoutNormalizeTarget(target);
    // resolve base page/spread
    var basePage = (opts.afterPage && opts.afterPage.isValid) ? opts.afterPage : (page && page.isValid ? page : doc.pages[doc.pages.length-1]);
    var newPage = null;
    try{
      try{ doc.allowPageShuffle = true; }catch(_docShuf){}
      if (basePage && basePage.parent && basePage.parent.isValid){
        try{ basePage.parent.allowPageShuffle = true; }catch(_spShuf){}
      }
    }catch(_prep){}
    function __layoutAddPage(basePageLocal, forceNewSpread){
      var newPageLocal = null;
      var forceSpread = !!forceNewSpread;
      if (forceSpread){
        try{
          var targetSpread = null;
          try{
            var baseSpread = (basePageLocal && basePageLocal.parent && basePageLocal.parent.isValid) ? basePageLocal.parent : null;
            if (baseSpread){
            try{ log(__layoutTag("info","base spread pages=" + baseSpread.pages.length + " name=" + baseSpread.name)); }catch(__logBase){}
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
              newPageLocal = targetSpread.pages[0];
            } else {
              newPageLocal = targetSpread.pages.add();
            }
          }
          if (!newPageLocal || !newPageLocal.isValid){
            newPageLocal = doc.pages.add(LocationOptions.AT_END);
          }
        }catch(eAddForce){
          try{ newPageLocal = doc.pages.add(LocationOptions.AT_END); }catch(eAddForce2){ newPageLocal = doc.pages.add(); }
        }
      } else {
        try{
          if (basePageLocal && basePageLocal.isValid){
            newPageLocal = doc.pages.add(LocationOptions.AFTER, basePageLocal);
          } else {
            newPageLocal = doc.pages.add(LocationOptions.AT_END);
          }
        }catch(eAdd){
          try{ newPageLocal = doc.pages.add(LocationOptions.AT_END); }catch(eAdd2){ newPageLocal = doc.pages.add(); }
        }
      }
      return newPageLocal;
    }
    var newPage = __layoutAddPage(basePage, !!(opts && opts.forceNewSpread));
    if (newPage && newPage.isValid){
      try{ newPage.appliedMaster = NothingEnum.NOTHING; }catch(_master){}
      try{
        var pn = newPage.name;
        try{ log(__layoutTag("info","add page name=" + pn)); }catch(_){}
      }catch(_pn){}
    }
    if (newPage && newPage.isValid){
      try{
        var w = parseFloat(target.pageWidthPt), h = parseFloat(target.pageHeightPt);
        if (isFinite(w) && isFinite(h) && w>0 && h>0){
          if (w>h) target.pageOrientation = "landscape";
          else if (w<h) target.pageOrientation = "portrait";
          if (target.pageOrientation === "landscape"){
            newPage.side = PageSideOptions.SINGLE_SIDED;
          }
          newPage.resize(
            CoordinateSpaces.PASTEBOARD_COORDINATES,
            AnchorPoint.TOP_LEFT_ANCHOR,
            ResizeMethods.REPLACING_CURRENT_DIMENSIONS_WITH,
            [w, h]
          );
        }
      }catch(eResize){ try{ log(__layoutTag("warn","layout page resize failed: " + eResize)); }catch(_){ } }
      try{
        var mp = newPage.marginPreferences;
        var margins = target.pageMarginsPt || {};
        if (mp){
          if (isFinite(margins.top)) mp.top = margins.top;
          if (isFinite(margins.bottom)) mp.bottom = margins.bottom;
          if (isFinite(margins.left)) mp.left = margins.left;
          if (isFinite(margins.right)) mp.right = margins.right;
        }
      }catch(eMargin){ try{ log(__layoutTag("warn","layout margin apply failed: " + eMargin)); }catch(_){ } }
      var newFrame = createTextFrameOnPage(newPage, target);
      try{
        if (newFrame && newFrame.isValid){
          log(__layoutTag("info","new frame id=" + newFrame.id + " orient=" + (target.pageOrientation||"") + " page=" + (newPage && newPage.name)));
        }
      }catch(_){}
      if (newFrame && newFrame.isValid && linkFromFrame && linkFromFrame.isValid){
        try{ linkFromFrame.nextTextFrame = newFrame; }catch(eLink){ try{ log(__layoutTag("warn","layout frame link failed: " + eLink)); }catch(_){ } }
      }
      return { page: newPage, frame: newFrame };
    }
  }catch(e){ try{ log(__layoutTag("warn","create layout frame failed: " + e)); }catch(_){ } }
  return null;
}

function __ensureLayout(targetState){
  try{ log(__layoutTag("info","ensure request orient=" + (targetState && targetState.pageOrientation) + " width=" + (targetState && targetState.pageWidthPt) + " height=" + (targetState && targetState.pageHeightPt))); }catch(_){}
  var target = targetState ? __cloneLayoutState(targetState) : __cloneLayoutState(__DEFAULT_LAYOUT);
  target = __layoutNormalizeTarget(target);
  var prevOrientation = __CURRENT_LAYOUT ? __CURRENT_LAYOUT.pageOrientation : null;
  var needNewSpread = !!(target.pageOrientation && prevOrientation && target.pageOrientation !== prevOrientation);
  if (__layoutSkipIfSame(target)) return;
  var prevFrame = (typeof tf !== "undefined" && tf && tf.isValid) ? tf : null;
  var pkt = __createLayoutFrame(target, prevFrame, {forceNewSpread: needNewSpread});
  __layoutApplyPacket(pkt, target);
}

function __ensureLayoutDefault(){
  var target = null;
  try{
    var dp = doc.documentPreferences;
    var mpSource = null;
    try{ if (doc.pages.length > 0){ mpSource = doc.pages[0].marginPreferences; } }catch(_){ }
    if (!mpSource){ try{ mpSource = doc.marginPreferences; }catch(_){ } }
    target = {
      pageOrientation: (dp && dp.pageOrientation === PageOrientation.LANDSCAPE) ? "landscape" : "portrait",
      pageWidthPt: dp ? parseFloat(dp.pageWidth) : null,
      pageHeightPt: dp ? parseFloat(dp.pageHeight) : null,
      pageMarginsPt: mpSource ? {
        top: parseFloat(mpSource.top),
        bottom: parseFloat(mpSource.bottom),
        left: parseFloat(mpSource.left),
        right: parseFloat(mpSource.right)
      } : null
    };
  }catch(_){ }
  __ensureLayout(target);
}

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
        try{ log(__layoutTag("info","apply frame id=" + frame.id + " innerWidth=" + innerWidth + " orient=" + (layoutState && layoutState.pageOrientation))); }catch(_log){}
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

    

function flushOverflow(currentStory, lastPage, lastFrame, maxPagesPerCall) {
        var MAX_PAGES = 20;
        try{
            if (maxPagesPerCall !== undefined && maxPagesPerCall !== null){
                var mp = parseInt(maxPagesPerCall, 10);
                if (isFinite(mp) && mp > 0) MAX_PAGES = mp;
            }
        }catch(_max){}
        var STALL_LIMIT = 3;
        var stallFrameId = null;
        var stallCount = 0;
        var stallCharCount = 0;
        var lastCharLen = null;
        function __logFlushWarn(msg){
            var pgName = "NA";
            var frameId = "NA";
            try{ if (lastPage && lastPage.isValid && lastPage.name) pgName = lastPage.name; }catch(_){}
            try{ if (lastFrame && lastFrame.isValid && lastFrame.id != null) frameId = lastFrame.id; }catch(_){}
            var payload = "flushOverflow " + msg + " page=" + pgName + " frame=" + frameId;
            try{
                if (typeof warn === "function") { warn(payload); return; }
            }catch(_warnFail){}
            try{ log("[WARN] " + payload); }catch(_logFail){}
        }
        for (var guard = 0; currentStory && currentStory.overflows && guard < MAX_PAGES; guard++) {
            var docRef = app && app.activeDocument;
            try{
                if (__SAFE_PAGE_LIMIT && docRef && docRef.pages && docRef.pages.length >= __SAFE_PAGE_LIMIT){
                    __logFlushWarn("flushOverflow page limit hit (" + __SAFE_PAGE_LIMIT + ")");
                    break;
                }
            }catch(_limit){}
            var pkt = __createLayoutFrame(__CURRENT_LAYOUT, lastFrame, {afterPage:lastPage, forceBreak:false});
            if (!pkt || !pkt.frame || !pkt.page) {
                __logFlushWarn("flushOverflow failed to allocate new frame");
                break;
            }
            lastPage  = pkt.page;
            lastFrame = pkt.frame;

            try { currentStory.recompose(); } catch(_) {}
            try { app.activeDocument.recompose(); } catch(_) {}
            $.sleep(10);

            var tailFrameId = null;
            var curLen = null;
            try{ curLen = currentStory && currentStory.characters ? currentStory.characters.length : null; }catch(_cl){}
            try{
                var tailIp = currentStory && currentStory.isValid ? currentStory.insertionPoints[-1] : null;
                if (tailIp && tailIp.isValid && tailIp.parentTextFrames && tailIp.parentTextFrames.length){
                    var tailFrame = tailIp.parentTextFrames[0];
                    if (tailFrame && tailFrame.isValid) tailFrameId = tailFrame.id;
                }
            }catch(_tail){}
            if (tailFrameId !== null){
                if (stallFrameId !== null && tailFrameId === stallFrameId){
                    stallCount++;
                }else{
                    stallCount = 0;
                    stallFrameId = tailFrameId;
                }
                if (stallCount >= STALL_LIMIT){
                    __logFlushWarn("flushOverflow guard hit; no progress resolving overset");
                    break;
                }
            }
            if (curLen !== null && tailFrameId !== null){
                if (lastCharLen !== null && curLen === lastCharLen && tailFrameId === stallFrameId){
                    stallCharCount++;
                }else{
                    stallCharCount = 0;
                    lastCharLen = curLen;
                }
                if (stallCharCount >= STALL_LIMIT){
                    __logFlushWarn("flushOverflow guard hit; story length not advancing");
                    break;
                }
            }
        }
        if (currentStory && currentStory.overflows) {
            __logFlushWarn("flushOverflow guard hit; overset still true");
        }
        return { page: lastPage, frame: lastFrame, overset: (currentStory && currentStory.overflows) };
    }

    

function __trimTrailingEmptyFrames(story){
        if (!__ENABLE_TRAILING_TRIM) return;
        try{
            if (!story || !story.isValid) return;
            var tcs = story.textContainers;
            if (!tcs || !tcs.length) return;
            for (var idx = tcs.length - 1; idx >= 0; idx--){
                var frame = tcs[idx];
                if (!frame || !frame.isValid) continue;
                var hasTable = false;
                try{ hasTable = (frame.tables && frame.tables.length>0); }catch(_){}
                if (!hasTable){
                    try{
                        var tfTexts = frame.texts;
                        if (tfTexts && tfTexts.length){
                            for (var ti=0; ti<tfTexts.length; ti++){
                                var txtObj = tfTexts[ti];
                                try{
                                    if (txtObj.tables && txtObj.tables.length){
                                        hasTable = true;
                                        break;
                                    }
                                }catch(_){}
                            }
                        }
                    }catch(_){}
                }
                if (hasTable) break;
                var txt = "";
                try{ txt = String(frame.contents || ""); }catch(_){}
                if (txt.replace(/[\s\u0000-\u001f\u2028\u2029\uFFFC\uF8FF]+/g, "") !== ""){
                    break;
                }
                try{
                    var prevFrame = null;
                    try{ prevFrame = frame.previousTextFrame; }catch(_){}
                    if (prevFrame && prevFrame.isValid){
                        var prevOverflow = false;
                        try{ prevOverflow = prevFrame.overflows; }catch(_){}
                        if (prevOverflow){
                            break;
                        }
                    }
                }catch(_){}
                try{
                    var prev = null;
                    try{ prev = frame.previousTextFrame; }catch(_){}
                    if (prev && prev.isValid){
                        try{ prev.nextTextFrame = null; }catch(_){}
                    }
                }catch(_){}
                try{ frame.remove(); }catch(_){}
            }
        }catch(eTrim){
            try{ log("[DBG] trim trailing frames failed: " + eTrim); }catch(_){}
        }
    }

    

function __trimTrailingEmptyPages(doc){
        if (!__ENABLE_TRAILING_TRIM) return;
        try{
            for (var idx = doc.pages.length - 1; idx >= 1; idx--){
                var pg = doc.pages[idx];
                if (!pg || !pg.isValid) continue;
                var items = 0;
                try{ items = pg.pageItems.length; }catch(_){}
                if (items && items > 0) break;
                try{ pg.remove(); }catch(_){}
            }
        }catch(ePg){
            try{ log("[DBG] trim trailing pages failed: " + ePg); }catch(_){}
        }
    }

    

function startNewChapter(currentStory, currentPage, currentFrame) {
        if (currentStory) {
            var st = flushOverflow(story, page, tf);
            page  = st.page;
            tf    = st.frame;
            story = tf.parentStory;
            curTextFrame = tf;     
        }
        var np  = doc.pages.add(LocationOptions.AFTER, currentPage);
        var nft = createTextFrameOnPage(np, __CURRENT_LAYOUT);
        try{ __LAST_IMG_ANCHOR_IDX = -1; }catch(_){}
        return { story: nft.parentStory, page: np, frame: nft };
    }

    
function __layoutSkipIfSame(target){
  if (!__layoutsEqual(__CURRENT_LAYOUT, target)) return false;
  try{
    log(__layoutTag("info","ensure skip orient=" + (target.pageOrientation||"") + " width=" + target.pageWidthPt + " height=" + target.pageHeightPt));
  }catch(_){}
  try{
    if (target.pageOrientation && __CURRENT_LAYOUT && __CURRENT_LAYOUT.pageOrientation !== target.pageOrientation){
      var __skipPayload = __cloneLayoutState(target);
      __skipPayload.origin = "skip";
      log(__layoutTag("info","still skipping due to same state; page=" + (page && page.name) + " spec=" + __jsonStringifySafe(__skipPayload)));
    }
  }catch(__skipLog){}
  return true;
}
function __layoutApplyPacket(pkt, target){
  if (pkt && pkt.frame && pkt.frame.isValid){
    try{ log(__layoutTag("info","ensure apply orient=" + (target.pageOrientation||"") + " page=" + (pkt.page && pkt.page.name) + " frame=" + pkt.frame.id)); }catch(_){}
    page = pkt.page;
    tf = pkt.frame;
    story = tf.parentStory;
    curTextFrame = tf;
    __CURRENT_LAYOUT = __cloneLayoutState(target);
    try{ story.recompose(); }catch(_){}
    try{ app.activeDocument.recompose(); }catch(_){}
    return true;
  }else{
    try{ log(__layoutTag("warn","ensure failed - cannot allocate frame")); }catch(_){}
    return false;
  }
}
