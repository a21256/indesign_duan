function __tblAddTableHiFi(obj){
      try{
        var rows = obj.rows|0, cols = obj.cols|0;
        if (rows<=0 || cols<=0) return;
        var __cfgStyles = (typeof CONFIG !== "undefined" && CONFIG.styles) ? CONFIG.styles : {};
        var __tblDiag = [];
        var __phaseName = "init";
        function __diag(label, getter){
          try{
            var v = getter();
            __tblDiag.push(label + "=" + v);
          }catch(_){ __tblDiag.push(label + "=ERR"); }
        }
        function __phase(name){
          __phaseName = name;
          try{ __diag("phase", function(){ return name; }); }catch(__){}
        }
        function __logErr(phase, err){
          try{
            var lineInfo = "";
            try{ lineInfo = " line=" + err.line; }catch(__){}
            var fileInfo = "";
            try{ fileInfo = " file=" + err.fileName; }catch(__){}
            log(__tableErrTag + " phase=" + phase + " err=" + err + lineInfo + fileInfo + " diag=" + __tblDiag.join("|"));
          }catch(__){}
        }
        __phase("start");
        try{ __diag("rows", function(){ return rows; }); }catch(__){}
        try{ __diag("cols", function(){ return cols; }); }catch(__){}
        var __styleCfg = {
          primary:  __cfgStyles.tableBody || %TABLE_BODY_STYLE%,
          fallback: __cfgStyles.tableBodyFallback || %TABLE_BODY_STYLE_FALLBACK%,
          base:     __cfgStyles.tableBodyBase || %TABLE_BODY_STYLE_BASE%,
          auto:     __cfgStyles.tableBodyAuto || %TABLE_BODY_STYLE_AUTO%
        };
        function __tblSelfCheck(){
          try{
            var placeholders = [%TABLE_BODY_STYLE%, %TABLE_BODY_STYLE_FALLBACK%, %TABLE_BODY_STYLE_BASE%, %TABLE_BODY_STYLE_AUTO%];
            if (placeholders.join(",").indexOf("%") >= 0) throw "TABLE style placeholders not replaced";
            var required = ["__ensureLayout","__createLayoutFrame","__applyFrameLayout"];
            for (var i=0;i<required.length;i++){
              var n = required[i];
              if (typeof eval(n) !== "function") throw ("missing function: " + n);
            }
          }catch(e){
            try{ log("[ERR] TABLE selfcheck failed: " + e); }catch(_){ }
            throw e;
          }
        }
        function __sanitizeStyleName(name){
          if (!name) return "[None]";
          if (typeof name === "string" && name.length && name.charAt(0) === "%") return "[None]";
          return name;
        }
        __styleCfg.primary  = __sanitizeStyleName(__styleCfg.primary);
        __styleCfg.fallback = __sanitizeStyleName(__styleCfg.fallback);
        __styleCfg.base     = __sanitizeStyleName(__styleCfg.base);
        __styleCfg.auto     = __sanitizeStyleName(__styleCfg.auto);
        var __tableCtx = (obj && obj.logContext) ? obj.logContext : null;
        function __tblTags(){
          var tag = "[TABLE]", warnTag = "[WARN]", errTag = "[ERROR]";
          try{
            if (__tableCtx && __tableCtx.id){
              tag = "[TABLE][" + __tableCtx.id + "]";
              warnTag = "[WARN][TABLE " + __tableCtx.id + "]";
              errTag = "[ERROR][TABLE " + __tableCtx.id + "]";
            }
          }catch(_){}
          return {tag: tag, warn: warnTag, err: errTag};
        }
        var __tableTagObj = __tblTags();
        var __tableTag = __tableTagObj.tag;
        var __tableWarnTag = __tableTagObj.warn;
        var __tableErrTag = __tableTagObj.err;
        __tblSelfCheck();
        function __resolveTableParaStyle(styleName){
          if (!styleName || styleName === "null" || styleName === "undefined") return null;
          try{
            var st = app.activeDocument.paragraphStyles.itemByName(styleName);
            if (st && st.isValid) return st;
          }catch(_){}
          return null;
        }
        function __ensureAutoTableStyle(styleName, baseName){
          if (!styleName || styleName === "null" || styleName === "undefined") return null;
          var existing = __resolveTableParaStyle(styleName);
          if (existing) return existing;
          try{
            var doc = app.activeDocument;
            var base = __resolveTableParaStyle(baseName);
            var spec = { name: styleName };
            var baseSize = 10;
            var baseLeading = 12;
            if (base && base.isValid){
              try{ spec.basedOn = base; }catch(_){}
              var ps = parseFloat(base.pointSize);
              if (isFinite(ps) && ps > 0) baseSize = ps;
              var ld = base.leading;
              var ldVal = parseFloat(ld);
              if (isFinite(ldVal) && ldVal > 0){ baseLeading = ldVal; }
              else { baseLeading = baseSize + 2; }
              spec.pointSize = Math.max(5, baseSize * 0.9);
              spec.leading = Math.max(spec.pointSize + 1, baseLeading * 0.9);
              try{ spec.appliedFont = base.appliedFont; }catch(_){}
            }else{
              spec.pointSize = 9;
              spec.leading = 11;
            }
            var created = doc.paragraphStyles.add(spec);
            if (created && created.isValid) return created;
          }catch(_){}
          return __resolveTableParaStyle(baseName);
        }
        if (__tableCtx){
          try{
            var __tblPrev = __tableCtx.preview ? String(__tableCtx.preview) : "";
            if (__tblPrev.length > 80) __tblPrev = __tblPrev.substring(0,80) + "...";
            var __tblSummary = ' para=' + (__tableCtx.paraIndex||"?") + ' style=' + (__tableCtx.style||"");
            if (__tblPrev) __tblSummary += ' text="' + __tblPrev + '"';
            log(__tableTag + " ctx" + __tblSummary);
          }catch(__ctxLog){}
        }
        // resolve layout spec (orientation/margins) from obj
        var layoutSpec = null;
        try{
          if (obj){
            if (obj.pageOrientation){
              layoutSpec = { pageOrientation: obj.pageOrientation };
            } else if (obj.pageWidthPt && obj.pageHeightPt){
              var w = parseFloat(obj.pageWidthPt), h = parseFloat(obj.pageHeightPt);
              if (isFinite(w) && isFinite(h)){
                layoutSpec = { pageOrientation: (w > h ? "landscape" : "portrait") };
              }
            }
            if (layoutSpec && layoutSpec.pageOrientation && __DEFAULT_LAYOUT){
              var baseW = parseFloat(__DEFAULT_LAYOUT.pageWidthPt);
              var baseH = parseFloat(__DEFAULT_LAYOUT.pageHeightPt);
              if (isFinite(baseW) && isFinite(baseH)){
                if (layoutSpec.pageOrientation === "landscape"){
                  layoutSpec.pageWidthPt = baseH;
                  layoutSpec.pageHeightPt = baseW;
                } else {
                  layoutSpec.pageWidthPt = baseW;
                  layoutSpec.pageHeightPt = baseH;
                }
              }
              if (__DEFAULT_LAYOUT.pageMarginsPt){
                layoutSpec.pageMarginsPt = __cloneLayoutState({pageMarginsPt: __DEFAULT_LAYOUT.pageMarginsPt}).pageMarginsPt;
              }
            } else if (layoutSpec && obj.pageMarginsPt){
              layoutSpec.pageMarginsPt = obj.pageMarginsPt;
            }
          }
        }catch(_){ layoutSpec = null; }
        function __tblApplyLayout(spec){
          var appliedSwitch = false;
          try{
            if (!spec){
              if (__CURRENT_LAYOUT && __DEFAULT_LAYOUT && !__layoutsEqual(__CURRENT_LAYOUT, __DEFAULT_LAYOUT)){
                __ensureLayoutDefault();
              }
              return appliedSwitch;
            }
            var curOrientation = __CURRENT_LAYOUT ? __CURRENT_LAYOUT.pageOrientation : null;
            var defaultOrientation = (__DEFAULT_LAYOUT && __DEFAULT_LAYOUT.pageOrientation) ? __DEFAULT_LAYOUT.pageOrientation : null;
            var targetOrientation = spec.pageOrientation || null;
            var needSwitch = false;
            try{
              var currentOri = curOrientation || defaultOrientation;
              needSwitch = !!(targetOrientation && targetOrientation !== currentOri);
            }catch(__){}
            try{
              log(__tableTag + " layout pre cur=" + (curOrientation||"NA") + " default=" + (defaultOrientation||"NA") + " target=" + (targetOrientation||"NA") + " needSwitch=" + needSwitch + " page=" + (page&&page.isValid?page.name:"NA"));
            }catch(__preLog){}
            // 若当前尚未记录布局，但目标朝向与默认不一致，也视为需要切换（避免直接把现页改成横版）
            // ??????????????????????????????????????????
            if (needSwitch){
              // ?????????????????????????????
              try{
                var spreadPages = null; try{ spreadPages = (page && page.parent && page.parent.pages) ? page.parent.pages : null; }catch(_sp){}
                var spreadLenDbg = (spreadPages && spreadPages.length) ? spreadPages.length : "NA";
                try{ log(__tableTag + " layout pre-switch spreadLen=" + spreadLenDbg + " page=" + (page&&page.isValid?page.name:"NA")); }catch(__){}
                try{ log(__tableTag + " layout pre-break to avoid mixed spread page=" + (page&&page.isValid?page.name:"NA")); }catch(__){}
                story.insertionPoints[-1].contents = SpecialCharacters.PAGE_BREAK;
                story.recompose();
                try{
                  var ipAlign = story.insertionPoints[-1];
                  var tfAlign = (ipAlign && ipAlign.isValid && ipAlign.parentTextFrames && ipAlign.parentTextFrames.length) ? ipAlign.parentTextFrames[0] : null;
                  if (tfAlign && tfAlign.isValid && tfAlign.parentPage && tfAlign.parentPage.isValid){
                    page = tfAlign.parentPage;
                    tf = tfAlign;
                    curTextFrame = tfAlign;
                    story = tfAlign.parentStory;
                    __ensureLayoutDefault();
                    curOrientation = __CURRENT_LAYOUT ? __CURRENT_LAYOUT.pageOrientation : curOrientation;
                  }
                }catch(_align){}
              }catch(_preSwitch){}
              // ?????????????
            try{
              story.insertionPoints[-1].contents = SpecialCharacters.PAGE_BREAK;
              story.recompose();
            }catch(__preBreakErr){ try{ log(__tableWarnTag + " page break before layout failed: " + __preBreakErr); }catch(_){ } }
            try{
              var ipAfter = null; try{ ipAfter = story.insertionPoints[-1]; }catch(_){ }
              var tfAfter = null; try{ if (ipAfter && ipAfter.isValid) tfAfter = ipAfter.parentTextFrames[0]; }catch(_){ }
              if (tfAfter && tfAfter.isValid && tfAfter.parentPage && tfAfter.parentPage.isValid){
                  // 创建新的跨页以承载目标方向的表格，避免与上一页混排
                  var pktSwitch = null;
                  try{
                    pktSwitch = __createLayoutFrame(spec, tfAfter, {afterPage: tfAfter.parentPage, forceNewSpread:true});
                  }catch(_pktErr){}
                  if (pktSwitch && pktSwitch.frame && pktSwitch.frame.isValid){
                    try{ tfAfter.nextTextFrame = pktSwitch.frame; }catch(_lnk){}
                    tf = pktSwitch.frame;
                    page = pktSwitch.page;
                    story = tf.parentStory;
                    curTextFrame = tf;
                    try{ __applyFrameLayout(tf, spec); }catch(_apSwitch){}
                    try{ __CURRENT_LAYOUT = __cloneLayoutState(spec); }catch(_updSwitch){}
                    appliedSwitch = true;
                    try{
                      var spreadAfter = null; try{ spreadAfter = (page && page.parent && page.parent.pages) ? page.parent.pages : null; }catch(_sa){}
                      var spreadAfterLen = (spreadAfter && spreadAfter.length) ? spreadAfter.length : "NA";
                      log(__tableTag + " layout switch -> " + (spec.pageOrientation||"") + " page=" + (page&&page.isValid?page.name:"NA") + " spreadLen=" + spreadAfterLen);
                    }catch(_logSwitch){
                      try{ log(__tableTag + " layout switch -> " + (spec.pageOrientation||"") + " page=" + (page&&page.isValid?page.name:"NA")); }catch(__){}
                    }
                    return appliedSwitch;
                  }
              }
            }catch(__switchErr){ try{ log(__tableWarnTag + " switch layout apply failed: " + __switchErr); }catch(_){ } }
          }
            __ensureLayout(spec);
            var newOrientation = __CURRENT_LAYOUT ? __CURRENT_LAYOUT.pageOrientation : curOrientation;
            if (spec.pageOrientation && newOrientation !== curOrientation){
              appliedSwitch = true;
            }
            try{ log(__tableTag + " layout ensured cur=" + (newOrientation||"NA") + " target=" + (spec.pageOrientation||"NA")); }catch(__postLog){}
            log(__tableTag + " layout request orient=" + (spec.pageOrientation||""));
          }catch(__layoutErr){
            try{ log(__tableWarnTag + " ensure layout failed: " + __layoutErr); }catch(__layoutLog){}
          }
          return appliedSwitch;
        }
        var layoutSwitchApplied = __tblApplyLayout(layoutSpec);
        try{ log(__tableTag + " begin rows="+rows+" cols="+cols); }catch(__){}
        var doc = app.activeDocument;
        __phase("layout-applied");

        // 如果切换到横版后前一页是空白，移除遗留空页，避免标题与表格之间出现空白
        try{
          if (layoutSwitchApplied && page && page.isValid && doc && doc.pages && doc.pages.length > 1){
            var prevPg = null;
            try{ prevPg = page.previousPage; }catch(_pp){}
            if (prevPg && prevPg.isValid){
              var items = 0;
              try{ items = prevPg.pageItems.length; }catch(_pi){}
              var hasText = false;
              try{
                if (prevPg.textFrames && prevPg.textFrames.length){
                  for (var tfi=0; tfi<prevPg.textFrames.length; tfi++){
                    var tft = prevPg.textFrames[tfi];
                    if (tft && tft.isValid){
                      var txt = "";
                      try{ txt = String(tft.contents||""); }catch(_tc){}
                      if (txt.replace(/[\s\u0000-\u001f\u2028\u2029\uFFFC\uF8FF]+/g,"") !== ""){
                        hasText = true; break;
                      }
                    }
                  }
                }
              }catch(_chk){}
              if (!hasText && items === 0){
                try{ prevPg.remove(); }catch(_rm){}
              }
            }
          }
        }catch(_cleanPrev){}

        function __tblResolveStoryRef(){
          var s = null;
          try{ if (story && story.isValid) s = story; }catch(_){ }
          if (!s){
            try{
              if (curTextFrame && curTextFrame.isValid && curTextFrame.parentStory && curTextFrame.parentStory.isValid){
                s = curTextFrame.parentStory;
              }
            }catch(_){ }
          }
          if (!s){
            try{
              if (typeof tf!=="undefined" && tf && tf.isValid && tf.parentStory && tf.parentStory.isValid){
                s = tf.parentStory;
              }
            }catch(_){ }
          }
          if (!s){
            try{
              if (doc && doc.stories && doc.stories.length>0){
                s = doc.stories[0];
              }
            }catch(_){ }
          }
          return s;
        }
        var storyRef = __tblResolveStoryRef();
        if (!storyRef || !storyRef.isValid){
          try{ log("[ERR] __tblAddTableHiFi: no valid story"); }catch(__){}
          return;
        }
        __phase("story-ok");
        story = storyRef;
        try { story.recompose(); } catch(_){ }
        __diag("pre.storyLen", function(){ return story.characters.length; });
        __diag("pre.pageCount", function(){ return app.activeDocument.pages.length; });
        __diag("pre.tf", function(){ return (typeof tf!=="undefined" && tf && tf.isValid) ? tf.id : "NA"; });
        try{ __diag("hasData", function(){ return obj && obj.data ? ("Y"+obj.data.length) : "N"; }); }catch(__){}

        // ensure a clean paragraph and flush overflow before inserting table
        function __tblPrepareStory(){
          try { story.insertionPoints[-1].contents = "\r"; } catch(_){ }
          try { story.recompose(); } catch(_){ }
          try{
            if (typeof flushOverflow === "function" && typeof tf !== "undefined" && tf && tf.isValid){
              var __pre = flushOverflow(story, page, tf);
              if (__pre && __pre.frame && __pre.page){
                page = __pre.page;
                tf   = __pre.frame;
                story = tf.parentStory;
                curTextFrame = tf;
              }
            }
          }catch(_){ }
        }
        __tblPrepareStory();

        function _ensureWritableFrameLocal(storyArg){
            // pick last non-overflow frame; fallback to tf; create new frame if needed
            var frameCandidate = null;
            try{
                var tcs = storyArg.textContainers;
                if (tcs && tcs.length){
                    for (var i=tcs.length-1; i>=0; i--){
                        try{
                            if (tcs[i].isValid && !tcs[i].overflows){
                                frameCandidate = tcs[i];
                                break;
                            }
                        }catch(_){ }
                    }
                    if (!frameCandidate){
                        frameCandidate = tcs[tcs.length-1];
                    }
                }
            }catch(_){ }

            if (!frameCandidate || !frameCandidate.isValid){
                try{
                    if (typeof tf!=="undefined" && tf && tf.isValid){
                        frameCandidate = tf;
                    }
                }catch(_){ }
            }
            if (frameCandidate && frameCandidate.isValid && !frameCandidate.overflows){
                __applyFrameLayout(frameCandidate, __CURRENT_LAYOUT);
                return frameCandidate;
            }

            var baseFrame = (frameCandidate && frameCandidate.isValid) ? frameCandidate : ((typeof tf!=="undefined" && tf && tf.isValid) ? tf : null);
            var basePage = null;
            try{
                if (baseFrame && baseFrame.isValid && baseFrame.parentPage && baseFrame.parentPage.isValid){
                    basePage = baseFrame.parentPage;
                }
            }catch(_){ }
            if (!basePage || !basePage.isValid){
                basePage = page;
            }
            try{
                var pktLocal = __createLayoutFrame(__CURRENT_LAYOUT, baseFrame, {afterPage: basePage});
                if (pktLocal && pktLocal.frame && pktLocal.frame.isValid){
                    try{ storyArg.recompose(); }catch(__){ }
                    page = pktLocal.page;
                    tf = pktLocal.frame;
                    story = tf.parentStory;
                    curTextFrame = pktLocal.frame;
                    return pktLocal.frame;
                }
            }catch(__){ }

            return frameCandidate;
        }

        function _prepareTableInsertion(storyArg){
            try{
                if (!storyArg || !storyArg.isValid) return;
                var ip = storyArg.insertionPoints[-1];
                if (!ip || !ip.isValid) return;
                var needParaBreak = true;
                try{
                    if (storyArg.characters.length > 0){
                        var lastChar = storyArg.characters[-1].contents;
                        if (lastChar === "\r") needParaBreak = false;
                    }
                }catch(_){ }
                if (needParaBreak){
                    try{ ip.contents = "\r"; }catch(__){ }
                }
                try{ storyArg.recompose(); }catch(__){ }
            }catch(__){ }
        }

        function _roughTableHeight(rowsCount, objSpec){
            // explicit row heights take priority
            var explicitSum = 0;
            try{
                if (objSpec && objSpec.rowHeightsPt && objSpec.rowHeightsPt.length){
                    for (var ri=0; ri<objSpec.rowHeightsPt.length; ri++){
                        var hv = parseFloat(objSpec.rowHeightsPt[ri]);
                        if (isFinite(hv) && hv>0) explicitSum += hv;
                    }
                }
            }catch(_){}
            if (explicitSum > 0) return explicitSum + 24;

            // otherwise estimate per-line height from defaults
            var approxLine = 14;
            try{
                var defaults = app.activeDocument.textDefaults;
                var ps = parseFloat(defaults.pointSize);
                if (!isFinite(ps) || ps<=0) ps = 12;
                var ld = defaults.leading;
                if (typeof ld === "number" && ld>0){
                    approxLine = ld;
                } else if (defaults.leading === Leading.AUTO){
                    approxLine = ps * 1.2;
                } else {
                    approxLine = ps * 1.2;
                }
            }catch(_){}
            if (!isFinite(approxLine) || approxLine<=0) approxLine = 14;
            var total = rowsCount * approxLine;
            try{
                var hdr = parseInt(objSpec && objSpec.headerRows || 0, 10);
                if (hdr>0) total += approxLine * 0.75;
            }catch(_){}
            return total + 24;
        }

        function _maybePreBreakForTable(storyArg, frameArg, rowsCount, objSpec){
            var result = { frame: frameArg, page: (frameArg && frameArg.isValid && frameArg.parentPage && frameArg.parentPage.isValid) ? frameArg.parentPage : null, didBreak: false };
            try{
                if (!storyArg || !storyArg.isValid) return result;
                var ipCheck = storyArg.insertionPoints[-1];
                if (!ipCheck || !ipCheck.isValid) return result;
                // decide holder frame
                var holder = null;
                try{
                    if (ipCheck.parentTextFrames && ipCheck.parentTextFrames.length){
                        holder = ipCheck.parentTextFrames[0];
                    }
                }catch(_){}
                if (!holder || !holder.isValid) holder = frameArg;
                if (!holder || !holder.isValid) return result;

                // measure available height from baseline to frame bottom
                var gbHold = holder.geometricBounds;
                var frameBottom = gbHold[2];
                var frameTop    = gbHold[0];
                var baseline = null;
                try{ baseline = ipCheck.baseline; }catch(_){}
                if (baseline == null || !isFinite(baseline)){
                    try{
                        if (ipCheck.index > 0){
                            var prevIP = storyArg.insertionPoints[ipCheck.index-1];
                            if (prevIP && prevIP.isValid) baseline = prevIP.baseline;
                        }
                    }catch(_){}
                }
                if (baseline == null || !isFinite(baseline)) baseline = frameTop;
                var available = frameBottom - baseline;
                if (!isFinite(available)) available = 0;
                if (available < 0) available = 0;

                var approxNeed = _roughTableHeight(rowsCount, objSpec);

                if (layoutSwitchApplied){
                    // 刚切换版向的第一个文本框，避免预先插入 FRAME_BREAK 而留下一个空页
                    available = approxNeed;
                }

                if (approxNeed > available && available >= 0){
                    try{ log(__tableTag + " pre-break forcing approx=" + approxNeed + " avail=" + available + " rows=" + rowsCount); }catch(__log0){}
                    try{
                        ipCheck.contents = SpecialCharacters.FRAME_BREAK;
                    }catch(_){
                        try{ ipCheck.contents = SpecialCharacters.COLUMN_BREAK; }catch(__){}
                    }
                    try{ storyArg.recompose(); }catch(_){}
                    try{
                        if (typeof flushOverflow === "function" && holder && holder.isValid){
                            var __fl = flushOverflow(storyArg, page, holder);
                            if (__fl && __fl.frame && __fl.frame.isValid){
                                result.frame = __fl.frame;
                                result.page  = __fl.page;
                            }
                        }
                    }catch(_){}
                    try{
                        var tailIP = storyArg.insertionPoints[-1];
                        if (tailIP && tailIP.isValid && tailIP.parentTextFrames && tailIP.parentTextFrames.length){
                            var tfAfter = tailIP.parentTextFrames[0];
                            if (tfAfter && tfAfter.isValid){
                                result.frame = tfAfter;
                                try{ result.page = tfAfter.parentPage; }catch(_){}
                            }
                        }
                    }catch(_){}
                    result.didBreak = true;
                    try{
                        log(__tableTag + " pre-break result frame=" + (result.frame && result.frame.isValid ? result.frame.id : "NA")
                            + " page=" + (result.page && result.page.isValid ? result.page.name : (result.frame && result.frame.parentPage ? result.frame.parentPage.name : "NA")));
                    }catch(__log1){}
                } else {
                    try{ log(__tableTag + " pre-break skip approx=" + approxNeed + " avail=" + available + " rows=" + rowsCount); }catch(__log2){}
                }
            }catch(e){
                try{ log("[WARN] table pre-break failed: " + e); }catch(__){}
            }
            return result;
        }

        var baseFrame = _ensureWritableFrameLocal(story);
        if (!baseFrame || !baseFrame.isValid){
          try{ log("[ERR] __tblAddTableHiFi: no writable frame"); }catch(__){}
          return;
        }
        try{ curTextFrame = baseFrame; }catch(_){ }
        try{ tf = baseFrame; }catch(_){ }
        _prepareTableInsertion(story);
        var __preBreakInfo = _maybePreBreakForTable(story, baseFrame, rows, obj);
        if (__preBreakInfo){
          if (__preBreakInfo.frame && __preBreakInfo.frame.isValid){
            baseFrame = __preBreakInfo.frame;
            try{ tf = baseFrame; }catch(__tf0){}
            try{ curTextFrame = baseFrame; }catch(__tf1){}
          }
          if (__preBreakInfo.page && __preBreakInfo.page.isValid){
            page = __preBreakInfo.page;
          } else {
            try{
              if (baseFrame && baseFrame.isValid && baseFrame.parentPage && baseFrame.parentPage.isValid){
                page = baseFrame.parentPage;
              }
            }catch(__pg0){}
          }
          if (__preBreakInfo.didBreak){
            _prepareTableInsertion(story);
          }
        }

        var anchorParagraph = null;
        try{ anchorParagraph = story.paragraphs[-1]; }catch(_){ }
        if (!anchorParagraph || !anchorParagraph.isValid){
          try{ story.insertionPoints[-1].contents = "\r"; anchorParagraph = story.paragraphs[-1]; }catch(__){ }
        }
        if (!anchorParagraph || !anchorParagraph.isValid){
          try{ log("[ERR] __tblAddTableHiFi: cannot resolve anchor paragraph"); }catch(__){}
          return;
        }
        var anchorIP = null;
        try{ anchorIP = anchorParagraph.insertionPoints[0]; }catch(_){ }
        if (!anchorIP || !anchorIP.isValid){
          try{ log("[ERR] __tblAddTableHiFi: invalid anchor insertion point"); }catch(__){}
          return;
        }
        __phase("anchor-ok");

        var __storyLenBefore = 0;
        try{ __storyLenBefore = story.characters.length; }catch(_){}

        var tableStory = story;
        var activeFrame = baseFrame;
        var layoutInnerWidth = null;
        if (layoutSpec && isFinite(layoutSpec.pageWidthPt)){
          var leftMargin = 0, rightMargin = 0;
          if (layoutSpec.pageMarginsPt){
            leftMargin = parseFloat(layoutSpec.pageMarginsPt.left) || 0;
            rightMargin = parseFloat(layoutSpec.pageMarginsPt.right) || 0;
          }
          layoutInnerWidth = layoutSpec.pageWidthPt - leftMargin - rightMargin;
        }
        var innerWidth = layoutInnerWidth || _innerFrameWidth(activeFrame);
        var insertIP = anchorIP;
        if (!insertIP || !insertIP.isValid){
          insertIP = (typeof _safeIP==='function') ? _safeIP(baseFrame) : baseFrame.insertionPoints[-1];
        }
        if (!insertIP || !insertIP.isValid){
          try{ log('[ERR] __tblAddTableHiFi: cannot resolve inline insertion point'); }catch(__){}
          return;
        }
        try{
          var __baseFrameId = (baseFrame && baseFrame.isValid) ? baseFrame.id : "NA";
          var __basePageName = (baseFrame && baseFrame.isValid && baseFrame.parentPage && baseFrame.parentPage.isValid)
                                ? baseFrame.parentPage.name : "NA";
          var __anchorIdxDbg = (insertIP && insertIP.isValid) ? insertIP.index : "NA";
          log(__tableTag + " anchor pick storyLen=" + __storyLenBefore
              + " frame=" + __baseFrameId + " page=" + __basePageName
              + " ipIdx=" + __anchorIdxDbg);
        }catch(__dbgAnchor){}
        var tbl = null;
        try {
          tbl = insertIP.tables.add({ bodyRowCount: rows, columnCount: cols });
        } catch(eAdd) {
          try{ __diag("err.create", function(){ return eAdd; }); }catch(_){}
          __logErr("create", eAdd);
          return;
        }
        __phase("table-added");
        try{
          var __colLenInit = 0;
          try{ __colLenInit = tbl.columns.length; }catch(__colErr){}
          log(__tableTag + " init columns expected=" + cols + " actual=" + __colLenInit);
        }catch(__colLog){}
        try{ tableStory.recompose(); }catch(_){ }
        try {
          var hr = parseInt(obj.headerRows || 0, 10);
          if (hr > 0) tbl.headerRowCount = Math.min(hr, rows);
          try { tbl.rows.everyItem().autoGrow      = true; } catch(_){ }
          try { tbl.rows.everyItem().height        = RowAutoHeight.AUTO; } catch(_){ }
          try { tbl.rows.everyItem().heightType    = RowHeightType.AT_LEAST; } catch(_){ }
          try { tbl.rows.everyItem().minimumHeight = 0; } catch(_){ }
          try { tbl.rows.everyItem().maximumHeight = 1000000; } catch(_){ }
          try { tbl.rows.everyItem().keepWithNext = false; } catch(_){ }
          try { tbl.rows.everyItem().keepTogether = false; } catch(_){ }
          try{
            var allParas = tbl.cells.everyItem().texts[0].paragraphs.everyItem();
            allParas.keepOptions.keepLinesTogether = false;
            allParas.keepOptions.keepWithNext = false;
            try { allParas.composer = ComposerTypes.ADOBE_PARAGRAPH_COMPOSER; } catch(__){ }
          }catch(_){ }
        }catch(_){ }

        var MAX_ROWSPAN_INLINE = 25;
        var merges = [];
        var cellPlan = [];
        var cellMeta = [];
        var skipPos = [];
        var degradeNotice = false;
        function _cloneCellSpec(base, rsOverride, csOverride){
          var clone = {};
          for (var key in base){
            if (!base.hasOwnProperty(key)) continue;
            if (key === "rowspan" || key === "colspan") continue;
            clone[key] = base[key];
          }
          var rsSrc = base.rowspan==null ? 1 : parseInt(base.rowspan,10);
          var csSrc = base.colspan==null ? 1 : parseInt(base.colspan,10);
          if (!isFinite(rsSrc)) rsSrc = 1;
          if (!isFinite(csSrc)) csSrc = 1;
          clone.rowspan = (rsOverride!=null) ? rsOverride : rsSrc;
          clone.colspan = (csOverride!=null) ? csOverride : csSrc;
          return clone;
        }
        function _markCovered(r, c, spanRows, spanCols){
          for (var rr=r; rr<Math.min(rows, r+spanRows); rr++){
            for (var cc=c; cc<Math.min(cols, c+spanCols); cc++){
              if (rr===r && cc===c) continue;
              skipPos[rr][cc] = true;
            }
          }
        }
        function _flattenRowspan(r, c, rawSpec, spanRows, spanCols){
          degradeNotice = true;
          try{ log(__tableErrTag + " degrade rowspan rows=" + spanRows + " at r=" + r + " c=" + c); }catch(__warnLog){}
          var maxR = Math.min(rows, r + spanRows);
          for (var rr=r; rr<maxR; rr++){
            var clone = _cloneCellSpec(rawSpec, 1, spanCols);
            if (rr !== r){
              try{ clone.text = ""; }catch(__txt){}
              skipPos[rr][c] = true;
            }
            cellPlan[rr][c] = clone;
            cellMeta[rr][c] = { align: clone.align||"left", valign: clone.valign||"top" };
            for (var cc=c+1; cc<c+spanCols && cc<cols; cc++){
              skipPos[rr][cc] = true;
            }
          }
        }

        var data = (obj && obj.data) ? obj.data : [];
        try{ __diag("dataLen", function(){ return data.length; }); }catch(__){}
        for (var r=0; r<rows; r++){
          cellPlan[r] = [];
          cellMeta[r] = [];
          skipPos[r] = [];
          for (var c=0; c<cols; c++){
            cellPlan[r][c] = null;
            cellMeta[r][c] = null;
            skipPos[r][c] = false;
          }
        }

        __phase("plan-init");
        for (var r=0; r<rows; r++){
          var rowEntries = data[r] || [];
          var cPtr = 0;
          for (var e=0; e<rowEntries.length; e++){
            var rawSpec = rowEntries[e];
            if (rawSpec == null) rawSpec = {text:""};
            if (typeof rawSpec === "string") rawSpec = {text: rawSpec};
            var tmpRS = rawSpec.rowspan==null ? 1 : parseInt(rawSpec.rowspan,10);
            var tmpCS = rawSpec.colspan==null ? 1 : parseInt(rawSpec.colspan,10);
            if (!isFinite(tmpRS)) tmpRS = 1;
            if (!isFinite(tmpCS)) tmpCS = 1;

            if (tmpRS === 0 || tmpCS === 0){
              var __adv = (tmpCS <= 0) ? 1 : tmpCS;
              cPtr += __adv;
              continue;
            }

            while (cPtr < cols && skipPos[r][cPtr]) cPtr++;
            if (cPtr >= cols) break;

            var spanRows = Math.max(1, tmpRS);
            var spanCols = Math.max(1, tmpCS);

            if (spanRows > MAX_ROWSPAN_INLINE){
              _flattenRowspan(r, cPtr, rawSpec, spanRows, spanCols);
              cPtr += spanCols;
              continue;
            }

            cellPlan[r][cPtr] = _cloneCellSpec(rawSpec, spanRows, spanCols);
            cellMeta[r][cPtr] = { align: rawSpec.align||"left", valign: rawSpec.valign||"top" };
            if (spanRows>1 || spanCols>1){
              merges.push({r:r, c:cPtr, rs:spanRows, cs:spanCols});
              _markCovered(r, cPtr, spanRows, spanCols);
            }
            cPtr += spanCols;
          }
        }

        __phase("populate-cells");
        for (var r=0; r<rows; r++){
          for (var c=0; c<cols; c++){
            if (skipPos[r][c] && !cellPlan[r][c]) continue;
            var cellSpec2 = cellPlan[r][c];
            if (!cellSpec2){
              var src = (data[r] && data[r][c]) ? data[r][c] : {text:""};
              if (typeof src === "string") src = {text: src};
              cellSpec2 = _cloneCellSpec(src, 1, 1);
              cellPlan[r][c] = cellSpec2;
              cellMeta[r][c] = { align: cellSpec2.align||"left", valign: cellSpec2.valign||"top" };
            }

            var txt = smartWrapStr(String(cellSpec2.text||"").replace(/\r?\n/g, "\r"));
            try { tbl.rows[r].cells[c].texts[0].contents = txt; }
            catch(_){ try { tbl.cells[r*cols+c].contents = txt; } catch(__){} }

            var alignVal = cellSpec2.align || "left";
            var valignVal = cellSpec2.valign || "top";
            try { tbl.rows[r].cells[c].texts[0].paragraphs.everyItem().justification = _mapAlign(alignVal); } catch(_){ }
            try { tbl.rows[r].cells[c].verticalJustification = _mapVAlign(valignVal); } catch(_){ }
            cellMeta[r][c] = { align: alignVal, valign: valignVal };

            try{
              if (cellSpec2.shading && /^#([0-9a-fA-F]{6})$/.test(cellSpec2.shading)){
                var cname="CellFill_"+cellSpec2.shading.substr(1);
                var col; try{ col=app.activeDocument.colors.itemByName(cname); }catch(__){}
                try{
                  if(!col || !col.isValid){
                    col = app.activeDocument.colors.add({
                      name:cname, model:ColorModel.PROCESS, space:ColorSpace.RGB,
                      colorValue:[
                        parseInt(cellSpec2.shading.substr(1,2),16),
                        parseInt(cellSpec2.shading.substr(3,2),16),
                        parseInt(cellSpec2.shading.substr(5,2),16)
                      ]
                    });
                  }
                  tbl.rows[r].cells[c].fillTint  = 100;
                  tbl.rows[r].cells[c].fillColor = col;
                }catch(__){}
              }
            }catch(_){ }
          }
        }

        merges.sort(function(a,b){
          if (a.c !== b.c) return b.c - a.c;
          var areaDiff = (b.rs*b.cs) - (a.rs*a.cs);
          if (areaDiff !== 0) return areaDiff;
          return a.r - b.r;
        });
        for (var i=0; i<merges.length; i++){
          var m  = merges[i];
          var r1 = m.r, c1 = m.c, r2 = Math.min(rows-1, r1+m.rs-1), c2 = Math.min(cols-1, c1+m.cs-1);
          try{
            var range = tbl.cells.itemByRange(tbl.rows[r1].cells[c1], tbl.rows[r2].cells[c2]);
            range.merge();
          }catch(e){ }
        }

        if (merges.length){
          for (var mi=0; mi<merges.length; mi++){
            var mr = merges[mi];
            var meta = cellMeta[mr.r][mr.c];
            if (!meta) continue;
            try{
              var cell = tbl.rows[mr.r].cells[mr.c];
              cell.verticalJustification = _mapVAlign(meta.valign||"top");
              try{
                cell.texts[0].paragraphs.everyItem().justification = _mapAlign(meta.align||"left");
              }catch(_){}
            }catch(_){}
          }
        }

        try{
          var hr2 = parseInt(obj.headerRows||0,10);
          if (hr2>0){
            for (var rr=0; rr<Math.min(hr2, rows); rr++){
              tbl.rows[rr].cells.everyItem().texts[0].paragraphs.everyItem().justification = Justification.CENTER_ALIGN;
            }
          }
        }catch(e){ }

        try{
          var defaultBorders = {
            outerOn: true,
            innerHOn: true,
            innerVOn: true,
            outerWeight: 0.75,
            innerWeight: 0.5,
            color: "#000000",
            cellInset: (typeof obj.cellInset === "number" ? obj.cellInset : 1.5),
            headerBoldBorder: false
          };
          var borderOpts = (obj.borders && typeof obj.borders === "object") ? obj.borders : {};
          for (var key in defaultBorders){ if (!(key in borderOpts)) borderOpts[key] = defaultBorders[key]; }
          applyTableBorders(tbl, borderOpts);
        }catch(_){ }

        var usedExplicitWidths = false;
        var canAdjust = (tbl && tbl.isValid);
        try{
          var policy = String(obj.widthPolicy || "fit").toLowerCase();
          var innerW = innerWidth;
          var widths = null, totalBefore = 0;

          if (policy === "fit" && obj.colWidthFrac && obj.colWidthFrac.length === cols && isFinite(innerW) && innerW>0){
            widths = [];
            var sumFrac = 0;
            for (var f=0; f<cols; f++){ sumFrac += Math.max(0, parseFloat(obj.colWidthFrac[f])||0); }
            if (sumFrac <= 0) sumFrac = 1;
            for (var f2=0; f2<cols; f2++){
              var frac = Math.max(0, parseFloat(obj.colWidthFrac[f2])||0) / sumFrac;
              widths.push(innerW * frac);
            }
            totalBefore = innerW;
            usedExplicitWidths = true;
          }

          if (!widths && obj.colWidthsPt && obj.colWidthsPt.length === cols){
            widths = [];
            totalBefore = 0;
            for (var k=0; k<cols; k++){
              var v = parseFloat(obj.colWidthsPt[k]);
              if (!isFinite(v) || v<=0){
                v = (isFinite(innerW) && innerW>0) ? innerW/Math.max(cols,1) : 1;
              }
              widths.push(v);
              totalBefore += v;
            }
            if (policy === "fit" && isFinite(innerW) && innerW>0 && totalBefore>0){
              var s = innerW / totalBefore;
              for (var j=0; j<cols; j++) widths[j] = widths[j]*s;
              totalBefore = innerW;
            }
            usedExplicitWidths = true;
          }

          if (!widths){
            var base = (isFinite(innerW) && innerW>0) ? innerW : (cols*60);
            widths = [];
            for (var z=0; z<cols; z++) widths.push(base/Math.max(1,cols));
            totalBefore = base;
          }

          var totalBefore = 0;
          for (var sumIdx=0; sumIdx<widths.length; sumIdx++){
            totalBefore += widths[sumIdx];
          }
          var avgWidth = (isFinite(innerW) && innerW>0) ? (innerW/Math.max(cols,1)) : (totalBefore/Math.max(cols,1));
          var tinyThreshold = Math.max(6, avgWidth * 0.08);
          var tinyMask = [];
          var delta = 0;
          var adjustable = 0;
          for (var clampIdx=0; clampIdx<cols; clampIdx++){
            var needClamp = (widths[clampIdx] < tinyThreshold);
            if (needClamp){
              delta += (tinyThreshold - widths[clampIdx]);
              widths[clampIdx] = tinyThreshold;
              tinyMask[clampIdx] = true;
            }else{
              adjustable += widths[clampIdx];
              tinyMask[clampIdx] = false;
            }
          }
          if (delta > 0 && adjustable > 0){
            var nonTinyCount = 0;
            for (var cntIdx=0; cntIdx<cols; cntIdx++){
              if (!tinyMask[cntIdx]) nonTinyCount++;
            }
            var scale = (adjustable - delta) / adjustable;
            if (scale > 0){
              for (var shrinkIdx=0; shrinkIdx<cols; shrinkIdx++){
                if (!tinyMask[shrinkIdx]){
                  widths[shrinkIdx] = widths[shrinkIdx] * scale;
                }
              }
            }else if (nonTinyCount > 0){
              var even = adjustable / nonTinyCount;
              for (var evenIdx=0; evenIdx<cols; evenIdx++){
                if (!tinyMask[evenIdx]){
                  widths[evenIdx] = even;
                }
              }
            }
          }
          totalBefore = 0;
          for (var sumIdx2=0; sumIdx2<widths.length; sumIdx2++){
            totalBefore += widths[sumIdx2];
          }
          var enforcedWidth = innerW;
          if (!isFinite(enforcedWidth) || enforcedWidth <= 0){
            enforcedWidth = totalBefore;
          }
          if (enforcedWidth > 0){
            try{ tbl.preferredWidth = enforcedWidth; }catch(_){}
            try{ tbl.width = enforcedWidth; }catch(_){}
          }

          if (canAdjust){
            try { tbl.width = enforcedWidth; } catch(_){ canAdjust = false; }
          }
          if (canAdjust){
            try{
              var colLen = 0;
              try{ colLen = tbl.columns.length; }catch(__){}
              if (colLen !== cols){
                canAdjust = false;
                try{ log("[WARN] column count mismatch expected=" + cols + " actual=" + colLen); }catch(__){}
              } else {
                for (var tci=0; tci<cols; tci++){
                  var colObj = null;
                  try{ colObj = tbl.columns.item(tci); }catch(__){}
                  if (!colObj || !colObj.isValid){
                    canAdjust = false;
                    try{ log("[WARN] column object invalid idx=" + tci); }catch(__){}
                    break;
                  }
                  var targetWidth = widths[tci];
                  if (!isFinite(targetWidth) || targetWidth <= 0){
                    targetWidth = 1;
                  }else if (targetWidth < 1){
                    targetWidth = 1;
                  }
                  var assigned = false;
                  try{
                    assigned = _assignColumnWidth(colObj, targetWidth, tci);
                  }catch(_){}
                  if (!assigned){
                    canAdjust = false;
                    try{ log("[WARN] column width assign failed idx=" + tci); }catch(__){}
                    break;
                  }
                }
              }
            }catch(__){ canAdjust = false; }
            if (canAdjust){
              try { tbl.recompose(); } catch(__){}
            }
          }
        }catch(eWidth){
          canAdjust = false;
          try{ log("[WARN] col width resolve failed: " + eWidth); }catch(__){}
        }
        if (!usedExplicitWidths){
          try { _normalizeTableWidth(tbl); } catch(__){}
        }

        try{ tbl.recompose(); }catch(__){}
        try{
          var __resolvedTblStyle = __resolveTableParaStyle(__styleCfg.primary)
                                   || __resolveTableParaStyle(__styleCfg.fallback);
          if (!__resolvedTblStyle){
            __resolvedTblStyle = __ensureAutoTableStyle(__styleCfg.auto, __styleCfg.base);
          }
          if (!__resolvedTblStyle){
            __resolvedTblStyle = __resolveTableParaStyle(__styleCfg.base)
                                 || __resolveTableParaStyle("Body");
          }
          if (__resolvedTblStyle && __resolvedTblStyle.isValid){
            try{
              tbl.cells.everyItem().texts[0].paragraphs.everyItem().appliedParagraphStyle = __resolvedTblStyle;
            }catch(__applyTblStyle){}
          }
        }catch(__tableStyleErr){}
        try{
          var tfTbl = (tbl && tbl.isValid) ? tbl.parent : null;
          if (tfTbl && tfTbl.isValid && typeof tfTbl.fit === "function"){
            tfTbl.fit(FitOptions.FRAME_TO_CONTENT);
          } else {
            __diag("fit.parentMissing", function(){ return (tfTbl&&tfTbl.isValid)?tfTbl.id:"NA"; });
          }
        }catch(__fitErr){
          try{ __diag("fit.err", function(){ return __fitErr; }); }catch(_){}
        }
        try{
          var __gbFit = geometricBounds;
          var __hFit = (__gbFit && __gbFit.length>=3) ? (__gbFit[2]-__gbFit[0]) : "NA";
          log(__tableTag + " frame fit height=" + __hFit);
        }catch(_){}
        try{
          var __offsetIdx = (tbl && tbl.isValid && tbl.storyOffset && tbl.storyOffset.isValid) ? tbl.storyOffset.index : "NA";
          var __storyLenAfter = 0;
          try{ __storyLenAfter = story.characters.length; }catch(__lenErr){}
          log(__tableTag + " placed idx=" + __offsetIdx + " storyLenNow=" + __storyLenAfter);
        }catch(__placedDbg){}

        var __postTableIP = null;
        try{
          if (tbl && tbl.isValid){
            __postTableIP = tbl.storyOffset;
            if (__postTableIP && __postTableIP.isValid){
              var __postStory = __postTableIP.parentStory;
              if (__postStory && __postStory.isValid){
                var __idx = __postTableIP.index;
                try{ __postTableIP = __postStory.insertionPoints[__idx+1]; }
                catch(__idxErr){
                  try{ __postTableIP = __postStory.insertionPoints[-1]; }catch(__idxErr2){}
                }
                if (__postTableIP && __postTableIP.isValid){
                  story = __postStory;
                  try{
                    var __holderTf = (__postTableIP.parentTextFrames && __postTableIP.parentTextFrames.length)
                                      ? __postTableIP.parentTextFrames[0] : null;
                    if (__holderTf && __holderTf.isValid){
                      tf = __holderTf;
                      curTextFrame = __holderTf;
                      try{ page = __holderTf.parentPage; }catch(__pageErr){}
                    }
                  }catch(__holderErr){}
                }
              }
            }
          }
        }catch(__postErr){}
        if (!__postTableIP || !__postTableIP.isValid){
          try{ __postTableIP = story.insertionPoints[-1]; }catch(__fallbackErr){}
        }
        try{
          if (__postTableIP && __postTableIP.isValid){
            var __needCR = true;
            try{
              if (__postTableIP.index > 0){
                var __prevChar = story.characters[__postTableIP.index-1];
                if (__prevChar && __prevChar.isValid){
                  var __prevVal = String(__prevChar.contents||"");
                  if (__prevVal === "\r") __needCR = false;
                }
              }
            }catch(__prevErr){}
            if (__needCR){
              try{ __postTableIP.contents = "\r"; }catch(__insertErr){}
            }
            try{
              var __postIdxDbg = __postTableIP.index;
              var __tfDbg = (__postTableIP.parentTextFrames && __postTableIP.parentTextFrames.length)
                              ? __postTableIP.parentTextFrames[0] : null;
              var __tfIdDbg = (__tfDbg && __tfDbg.isValid) ? __tfDbg.id : "NA";
              var __pgDbg = (__tfDbg && __tfDbg.isValid && __tfDbg.parentPage && __tfDbg.parentPage.isValid)
                            ? __tfDbg.parentPage.name : "NA";
              log(__tableTag + " post-ip idx=" + __postIdxDbg + " frame=" + __tfIdDbg + " page=" + __pgDbg);
            }catch(__postDbg){}
          }
        }catch(__ipErr){}
        if (degradeNotice){
          try{
            log(__tableErrTag + " rowspan>=" + MAX_ROWSPAN_INLINE + " flattened; manual review required");
          }catch(__noticeWarn){}
        }
        // keep current layout until after post-table flush; default restore happens later
        try{ __LAST_IMG_ANCHOR_IDX = -1; }catch(__resetErr){}

        var __postFlush = {overset:false};
        try{
          story.recompose();
          if (typeof flushOverflow==="function" && tf && tf.isValid){
            __postFlush = flushOverflow(story, page, tf);
            page = __postFlush.page; tf = __postFlush.frame; story = tf.parentStory; curTextFrame = tf;
          }
        }catch(e){ log("[WARN] flush after table failed: " + e); }
        var __tableStillOverset = (__postFlush && __postFlush.overset);
        try{
          var __dbgTf = (tf && tf.isValid) ? tf.id : "NA";
          var __dbgPg = (tf && tf.isValid && tf.parentPage && tf.parentPage.isValid) ? tf.parentPage.name : "NA";
          var __dbgOri = (__CURRENT_LAYOUT && __CURRENT_LAYOUT.pageOrientation) ? __CURRENT_LAYOUT.pageOrientation : "NA";
          log(__tableTag + " post-table overset=" + __tableStillOverset + " layoutSwitch=" + layoutSwitchApplied + " tf=" + __dbgTf + " page=" + __dbgPg + " curLayoutOri=" + __dbgOri);
        }catch(__dbgPost){}
        if (layoutSwitchApplied && !__tableStillOverset){
          try{
            story.insertionPoints[-1].contents = SpecialCharacters.PAGE_BREAK;
            story.recompose();
          }catch(__restoreBreak){ try{ log("[WARN] page break before restore failed: " + __restoreBreak); }catch(_){ } }
          try{
            var restoreTarget = __DEFAULT_LAYOUT ? __cloneLayoutState(__DEFAULT_LAYOUT) : null;
            try{
              var __curOri = (__CURRENT_LAYOUT && __CURRENT_LAYOUT.pageOrientation) ? __CURRENT_LAYOUT.pageOrientation : "";
              var __pgName = (page && page.isValid && page.name) ? page.name : "NA";
              var __spreadLenDbg = (page && page.isValid && page.parent && page.parent.pages) ? page.parent.pages.length : "NA";
              var __tfIdDbg = (tf && tf.isValid && tf.id!=null) ? tf.id : "NA";
              log(__tableTag + " restore-layout pre curOri=" + __curOri + " target=portrait spreadLen=" + __spreadLenDbg + " page=" + __pgName + " tf=" + __tfIdDbg);
              }catch(__logRestore){}
              if (restoreTarget && typeof __createLayoutFrame === "function"){
                var pktRestore = __createLayoutFrame(restoreTarget, null, {afterPage: page, forceNewSpread:true});
                if (pktRestore && pktRestore.frame && pktRestore.frame.isValid){
                  try{ if (tf && tf.isValid) tf.nextTextFrame = pktRestore.frame; }catch(_lnkRestore){}
                  page = pktRestore.page;
                  tf = pktRestore.frame;
                  story = tf.parentStory;
                  curTextFrame = tf;
                  try{ __applyFrameLayout(tf, restoreTarget); }catch(_apRestore){}
                  try{ __CURRENT_LAYOUT = restoreTarget; }catch(_cln){ }
                  try{
                    var __spAfter = null; try{ __spAfter = (page && page.parent && page.parent.pages) ? page.parent.pages : null; }catch(_spa){}
                    var __spAfterLen = (__spAfter && __spAfter.length) ? __spAfter.length : "NA";
                    log(__tableTag + " restore-layout applied page=" + (page&&page.isValid?page.name:"NA") + " spreadLen=" + __spAfterLen + " tf=" + (tf&&tf.isValid?tf.id:"NA"));
                  }catch(_logRestoreApplied){}
                }
              }else{
                __ensureLayoutDefault();
              }
              story.recompose();
            if (typeof flushOverflow==="function" && tf && tf.isValid){
              var __restoreFlush = flushOverflow(story, page, tf);
              if (__restoreFlush && __restoreFlush.frame && __restoreFlush.page){
                page = __restoreFlush.page;
                tf = __restoreFlush.frame;
                story = tf.parentStory;
                curTextFrame = tf;
              }
            }
          }catch(__restoreErr){
            try{ log("[WARN] restore default layout failed: " + __restoreErr); }catch(_){ }
          }
        }
      }catch(e){
        __logErr(__phaseName || "outer", e);
      }
      try{
        var __tblDetail = (__tableCtx && __tableCtx.id) ? ("id=" + __tableCtx.id) : ("rows=" + rows + " cols=" + cols);
        __progressBump("TABLE", __tblDetail);
      }catch(_){}
    }
    
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
    
