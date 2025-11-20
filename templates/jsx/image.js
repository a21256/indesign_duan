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

