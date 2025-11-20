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

    

function flushOverflow(currentStory, lastPage, lastFrame) {
        // 仅在达到 MAX_PAGES 或总页数限制时退出，保持顺序造页，避免“无进展”误判。
        var MAX_PAGES = 20;
        var STALL_LIMIT = 3;
        var stallFrameId = null;
        var stallCount = 0;
        function __logFlushWarn(msg){
            try{
                var pgName = (lastPage && lastPage.isValid && lastPage.name) ? lastPage.name : "NA";
                var frameId = (lastFrame && lastFrame.isValid && lastFrame.id != null) ? lastFrame.id : "NA";
                log("[ERROR] " + msg + " page=" + pgName + " frame=" + frameId);
            }catch(_){}
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
        }
        if (currentStory && currentStory.overflows) {
            __logFlushWarn("flushOverflow guard hit; overset still true");
        }
        return { page: lastPage, frame: lastFrame };
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
            curTextFrame = tf;              // ★ 新增：切到新框后更新全局指针
        }
        var np  = doc.pages.add(LocationOptions.AFTER, currentPage);
        var nft = createTextFrameOnPage(np, __CURRENT_LAYOUT);
        try{ __LAST_IMG_ANCHOR_IDX = -1; }catch(_){}
        return { story: nft.parentStory, page: np, frame: nft };
    }

    