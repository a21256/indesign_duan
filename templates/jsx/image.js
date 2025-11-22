// ==== path helpers migrated from entry ====
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

function addImageAtV2(ip, spec) {
      var doc = app.activeDocument;
      try{
        log("[IMG] begin addImageAtV2 src=" + (spec&&spec.src)
            + " w=" + (spec&&spec.w) + " h=" + (spec&&spec.h)
            + " align=" + (spec&&spec.align) + " sb=" + (spec&&spec.spaceBefore) + " sa=" + (spec&&spec.spaceAfter));
      }catch(_){}

      function _toPtLocal(v){
        var s = String(v==null?"":v).replace(/^\s+|\s+$/g,"");
        if (/mm$/i.test(s)) return parseFloat(s)*2.83464567;
        if (/pt$/i.test(s)) return parseFloat(s);
        if (/px$/i.test(s)) return parseFloat(s)*0.75;
        if (s==="") return 0;
        var n = parseFloat(s); if (isNaN(n)) return 0; return n*0.75;
      }

      // 1) 校验文件
      var f = File(spec && spec.src);
      if (!f || !f.exists) { log("[ERR] addImageAtV2: file missing: " + (spec && spec.src)); return null; }

      // 2) story / 安全插入点
      var st = null;
      try {
        st = (ip && ip.isValid && ip.parentStory && ip.parentStory.isValid) ? ip.parentStory
           : (typeof curTextFrame!=="undefined" && curTextFrame && curTextFrame.isValid && curTextFrame.parentStory && curTextFrame.parentStory.isValid) ? curTextFrame.parentStory
           : (doc.stories.length ? doc.stories[0] : null);
      } catch(_){}
      if (!st) { log("[ERR] addImageAtV2: no valid story"); return null; }
      try { st.recompose(); } catch(_){}

      var inlineFlag = String((spec && spec.inline)||"").toLowerCase();
      var isInline = !(inlineFlag==="0" || inlineFlag==="false");
      if (spec && spec.forceBlock) isInline = false;

      // 关键：默认用“当前可写文本框 tf 的末尾插入点”，避免落到上一页的 story 尾框
      var ip2 = (ip && ip.isValid) ? ip
               : ((typeof tf!=="undefined" && tf && tf.isValid && tf.insertionPoints && tf.insertionPoints.length)
                    ? tf.insertionPoints[-1]
                    : st.insertionPoints[-1]);

      // --- FIX: 连续图片落在同一 IP 时，先推进一段，避免叠放 ---
      function _logAnchorContext(tag, ipCandidate){
        try{
          var holder = (ipCandidate && ipCandidate.isValid && ipCandidate.parentTextFrames.length)
                       ? ipCandidate.parentTextFrames[0] : null;
          var page   = (holder && holder.isValid) ? holder.parentPage : null;
          log('[IMGDBG] ' + tag + ' holderTf=' + (holder?holder.id:'NA')
              + ' page=' + (page?page.name:'NA')
              + ' ipIdx=' + (ipCandidate?ipCandidate.index:'NA')
              + ' lastIdx=' + (typeof __LAST_IMG_ANCHOR_IDX==='number'?__LAST_IMG_ANCHOR_IDX:'NA'));
        }catch(_){}}

      function _dedupeAnchor(ipCandidate){
        if (!ipCandidate || !ipCandidate.isValid) return ipCandidate;
        try{
          var lastIdx = (typeof __LAST_IMG_ANCHOR_IDX==='number') ? (__LAST_IMG_ANCHOR_IDX|0) : -1;
          try { log("[IMG-STACK][pre] ip.index=" + ipCandidate.index + " lastIdx=" + lastIdx); } catch(__){}
          if (ipCandidate.index === lastIdx) {
            try { ipCandidate.contents = "\r"; } catch(_){ }
            try { st.recompose(); } catch(_){ }
            try {
              if (typeof tf!=="undefined" && tf && tf.isValid) ipCandidate = tf.insertionPoints[-1];
              else ipCandidate = st.insertionPoints[-1];
            } catch(__){}
            try { log("[IMG-STACK][shift] new ip.index=" + (ipCandidate&&ipCandidate.isValid?ipCandidate.index:"NA")); } catch(__){}
          }
        }catch(_){ }
        return ipCandidate;
      }

      if (!isInline) {
        // --- 保障：每次放图前都新起一段，避免与上一张叠在同一锚点 ---
        try{
          var ipChk = (typeof tf!=="undefined" && tf && tf.isValid) ? tf.insertionPoints[-1] : st.insertionPoints[-1];
          var prev = (ipChk && ipChk.isValid && ipChk.index>0) ? st.insertionPoints[ipChk.index-1] : null;
          var prevIsCR = false; try{ prevIsCR = (prev && prev.isValid && String(prev.contents)=="\r"); }catch(__){}
          if (!prevIsCR) {
            try { ipChk.contents = "\r"; } catch(__){}
            try { st.recompose(); } catch(__){}
            try { ip2 = (typeof tf!=="undefined" && tf && tf.isValid) ? tf.insertionPoints[-1] : st.insertionPoints[-1]; } catch(__){}
            try { log("[IMG-STACK][prebreak] force new para; ip.index=" + (ip2&&ip2.isValid?ip2.index:"NA")); } catch(__){}
          }
        }catch(__){}

        // ---- 关键修正：确保插入点确实在“当前末尾文本框 tf 内”，而不是上一页的尾框 ----
        try{
          if ((!ip || !ip.isValid) && typeof tf!=="undefined" && tf && tf.isValid) {
            var guard = 0;
            while (guard++ < 3) {
              var holder = (ip2 && ip2.isValid && ip2.parentTextFrames && ip2.parentTextFrames.length)
                           ? ip2.parentTextFrames[0] : null;
              var ok = (holder && holder.isValid && tf && tf.isValid && holder.id === tf.id);
              if (ok) break;
              try { tf.insertionPoints[-1].contents = "\r"; } catch(_){ }
              try { st.recompose(); } catch(_){ }
              try { ip2 = tf.insertionPoints[-1]; } catch(_){ }
            }
            try{
              var _h = (ip2 && ip2.isValid && ip2.parentTextFrames && ip2.parentTextFrames.length)
                       ? ip2.parentTextFrames[0] : null;
              var _p = (_h && _h.isValid) ? _h.parentPage : null;
              log("[IMG] ip2.adjust  tf=" + (_h?_h.id:"NA") + " page=" + (_p?_p.name:"NA"));
            }catch(__){}
          }
        }catch(__){}

        // ---- 关键修正②：如果 ip2 处的“段落起点”不在当前文本框 tf（即本框是该段续行），
        try{
          if (ip2 && ip2.isValid && typeof tf!=="undefined" && tf && tf.isValid) {
            var para = ip2.paragraphs[0];
            var p0   = (para && para.isValid) ? para.insertionPoints[0] : null;
            var h0   = (p0 && p0.isValid && p0.parentTextFrames && p0.parentTextFrames.length)
                       ? p0.parentTextFrames[0] : null;
            if (h0 && h0.isValid && h0.id !== tf.id) {
              try { ip2.contents = "\r"; } catch(_){ }
              try { st.recompose(); } catch(_){ }
              try { ip2 = tf.insertionPoints[-1]; } catch(_){ }
              try{
                var _h2 = (ip2 && ip2.isValid && ip2.parentTextFrames && ip2.parentTextFrames.length)
                          ? ip2.parentTextFrames[0] : null;
                var _p2 = (_h2 && _h2.isValid) ? _h2.parentPage : null;
                log("[IMG] ip2.breakPara  tf=" + (_h2?_h2.id:"NA") + " page=" + (_p2?_p2.name:"NA"));
              }catch(__){}
            }
              try{
                log('[IMGDBG] breakPara ipIdx=' + (ip2?ip2.index:'NA'));
              }catch(_){ }
          }
        }catch(__){}
      } // end !isInline guard for prebreak/breakPara adjustments

      try{
        var _tf0 = (ip2 && ip2.isValid && ip2.parentTextFrames && ip2.parentTextFrames.length)? ip2.parentTextFrames[0] : null;
        var _pg0 = (_tf0 && _tf0.isValid)? _tf0.parentPage : null;
        log("[IMG] anchor.pre  tf=" + (_tf0?_tf0.id:"NA") + " page=" + (_pg0?_pg0.name:"NA")
            + " storyLen=" + (st?st.characters.length:"NA"));
      }catch(_){}
      if (!ip2 || !ip2.isValid) { log("[ERR] addImageAtV2: invalid insertion point"); return null; }

      // 3) place
      var placed = null;
      try { placed = ip2.place(f); } catch(ePlace){ log("[ERR] addImageAtV2: place failed: " + ePlace); return null; }
      if (!placed || !placed.length || !(placed[0] && placed[0].isValid)) { log("[ERR] addImageAtV2: place returned invalid"); return null; }

      // 4) 取矩形
      var item = placed[0], rect=null, cname="";
      try { cname = String(item.constructor.name); } catch(_){}
      if (cname==="Rectangle") rect = item;
      else {
        try { if (item && item.parent && item.parent.isValid && String(item.parent.constructor.name)==="Rectangle") rect=item.parent; } catch(_){}
      }
      if (!rect || !rect.isValid) { log("[ERR] addImageAtV2: no rectangle after place"); return null; }

      // 记录最近一次图片锚点，用于下一次“同位放图”检测
      try{
        var aNow = rect.storyOffset;
        if (aNow && aNow.isValid) __LAST_IMG_ANCHOR_IDX = aNow.index;
        // [日志] 本次已放置图片的锚点 index
        try { log("[IMG-STACK][placed] anchor.index=" + aNow.index); } catch(__){}
      }catch(_){}

      try{
        var _tf1 = (rect.storyOffset && rect.storyOffset.isValid && rect.storyOffset.parentTextFrames && rect.storyOffset.parentTextFrames.length)
                    ? rect.storyOffset.parentTextFrames[0] : null;
        var _pg1 = (_tf1 && _tf1.isValid)? _tf1.parentPage : null;
        log("[IMG] placed.rect  holderTf=" + (_tf1?_tf1.id:"NA") + " page=" + (_pg1?_pg1.name:"NA"));
        try{
          var _aNow = rect.storyOffset;
          log('[IMGDBG] anchor.idx=' + (_aNow&&_aNow.isValid?_aNow.index:'NA'));
        }catch(_){}

      }catch(_){}

      // 5) 锚定：Above-Line（块级，最稳），不启用文绕图
      try {
        var _anch = rect.anchoredObjectSettings;
        if (_anch && _anch.isValid !== false) {
          _anch.anchoredPosition = isInline ? AnchorPosition.INLINE_POSITION : AnchorPosition.ABOVE_LINE;
          var _anchorPoint = AnchorPoint.TOP_LEFT_ANCHOR;
          var _alignKey = String((spec&&spec.align)||"left").toLowerCase();
          if (!isInline) {
            if (_alignKey === "center") {
              _anchorPoint = AnchorPoint.TOP_CENTER_ANCHOR;
            } else if (_alignKey === "right") {
              _anchorPoint = AnchorPoint.TOP_RIGHT_ANCHOR;
            }
          }
          _anch.anchorPoint = _anchorPoint;
        }
      } catch(_){}

        // 6) 尺寸：优先使用 XML 的 w/h；w 受列内宽 innerW 限制；无 w/h 时按列宽
        try { rect.fit(FitOptions.PROPORTIONALLY); rect.fit(FitOptions.CENTER_CONTENT); } catch(_){}
        try {
          try { rect.fittingOptions.autoFit=false; } catch(__){}
          var wPt = _toPtLocal(spec && spec.w);
          var hPt = _toPtLocal(spec && spec.h);

          var gb  = rect.geometricBounds;
          var curW = Math.max(1e-6, gb[3]-gb[1]), curH = Math.max(1e-6, gb[2]-gb[0]);
          var ratio = curW / curH;

          // 以“矩形锚点所在文本框”为准求列内宽（同原逻辑）
          var innerW = 0, holder = null;
          try {
            var aIP = rect.storyOffset;
            if (aIP && aIP.isValid && aIP.parentTextFrames && aIP.parentTextFrames.length)
              holder = aIP.parentTextFrames[0];
            if ((!holder || !holder.isValid) && rect.parentTextFrames && rect.parentTextFrames.length)
              holder = rect.parentTextFrames[0];
            if (!holder || !holder.isValid) {
              if (typeof curTextFrame!=="undefined" && curTextFrame && curTextFrame.isValid) holder = curTextFrame;
              else if (typeof tf!=="undefined" && tf && tf.isValid) holder = tf;
            }
            if (holder && holder.isValid){
              innerW = _innerFrameWidth(holder);
              try{
                log("[IMGDBG] widthCtx holderTf=" + holder.id + " innerW=" + innerW.toFixed ? innerW.toFixed(2) : innerW);
              }catch(__){}
            }
          }catch(__){}

          // 目标宽高：直接用绝对几何边界设定，避免“按初始值缩放”引入倍数偏差
          var widthLimit = innerW>0 ? innerW : curW;
          var targetW = (wPt>0 ? Math.min(wPt, widthLimit) : widthLimit);
          var targetH = (hPt>0 ? hPt : (targetW / Math.max(ratio, 1e-6)));

          try{ rect.absoluteHorizontalScale=100; rect.absoluteVerticalScale=100; }catch(_){ }
          rect.geometricBounds = [gb[0], gb[1], gb[0] + targetH, gb[1] + targetW];

          // 再自适应一次，让内容紧贴新框
          try { rect.fit(FitOptions.PROPORTIONALLY); rect.fit(FitOptions.CENTER_CONTENT); } catch(__){}

          // 记录关键数值，便于定位
          try {
            log("[IMG] size targetW=" + (targetW||0).toFixed(2)
                + " innerW=" + (innerW||0).toFixed(2)
                + " wPt=" + (wPt||0) + " hPt=" + (hPt||0)
                + " ratio=" + (ratio||0).toFixed(4));
          } catch(__){}
        } catch(_){}

      // 7) 锚点所在段：根据 align 控制段落对齐；块级图再设置段前段后
      try{
        var p = rect.storyOffset.paragraphs[0];
        if (p && p.isValid){
          var a = String((spec&&spec.align)||"center").toLowerCase();
          p.justification = (a==="right") ? Justification.RIGHT_ALIGN : (a==="center" ? Justification.CENTER_ALIGN : Justification.LEFT_ALIGN);
          if (!isInline) {
            try { p.spaceBefore = _toPtLocal(spec&&spec.spaceBefore) || 0; } catch(_){}
            try { p.spaceAfter  = _toPtLocal(spec&&spec.spaceAfter)  || 2; } catch(_){}
          }
          p.keepOptions.keepWithNext = false;
          p.keepOptions.keepLinesTogether = false;
        }
      }catch(_){}

      // 8) 块级图片在锚点后补「段落结束 + 零宽空格」，保证下一步接在新段
      if (!isInline) {
        try{
          var aIP = rect.storyOffset;
          if (aIP && aIP.isValid){
            // 8.1 先在锚点后补一个段落结束
            var aft1 = aIP.parentStory.insertionPoints[aIP.index+1];
            if (aft1 && aft1.isValid){ aft1.contents = "\r"; }
            // 8.2 再补一个零宽空格，保证 storyEnd 真正来到“新段”末尾
            var aft2 = aIP.parentStory.insertionPoints[aIP.index+2];
            if (aft2 && aft2.isValid){ aft2.contents = "\u200B"; }
            try{ aIP.parentStory.recompose(); }catch(__){}
            // 8.3 用新段的插入点反查父文本框，强制把 tf/curTextFrame/story 切到“下一段所在的框”
            try{
              var holderNext = (aft2 && aft2.isValid && aft2.parentTextFrames && aft2.parentTextFrames.length)
                                 ? aft2.parentTextFrames[0] : null;
              if (holderNext && holderNext.isValid){
                tf = holderNext;
                curTextFrame = holderNext;
                story = holderNext.parentStory;
              }
            }catch(__){}
          }
        }catch(_){}

      }
      // 9) 立即回排并疏通 overset，避免正文被甩到文末；并把 “当前活动文本框” 切到这张图所在的框
      if (!isInline) {
        try {
          if (st && st.isValid) st.recompose();
          if (rect && rect.isValid) { try { rect.recompose(); } catch(__){} }
          var __pg = (rect && rect.parentPage) ? rect.parentPage : (typeof page!=="undefined"?page:null);
          // 用矩形锚点反查真正所在的文本框，作为下一个动作的基准
          var __tf = null;
          try{
            var _a = rect.storyOffset;
            if (_a && _a.isValid && _a.parentTextFrames && _a.parentTextFrames.length)
              __tf = _a.parentTextFrames[0];
          }catch(_){}
          // 优先使用 8.3 中刚切换过来的 tf，其次才兜底
          if (!__tf && typeof tf!=="undefined") __tf = tf;
          if (!__tf && typeof curTextFrame!=="undefined") __tf = curTextFrame;
          if (__pg && __tf && typeof flushOverflow === "function") {
            var fl = flushOverflow(st, __pg, __tf);
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
          // 再兜底一次：若 flush 没返回新框，也把 curTextFrame 切到图所在框
          try{
            if ((!curTextFrame || !curTextFrame.isValid) && rect && rect.isValid){
              var a2 = rect.storyOffset;
              if (a2 && a2.isValid && a2.parentTextFrames && a2.parentTextFrames.length)
                curTextFrame = a2.parentTextFrames[0];
            }
          }catch(_){}
        } catch(eFlush){ log("[WARN] flush after image: " + eFlush); }

      }
      return rect;
    }

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
