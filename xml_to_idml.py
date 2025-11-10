# -*- coding: utf-8 -*-
import os, sys, subprocess, re
import argparse
import xml.etree.ElementTree as ET
from docx_to_xml_outline_notes_v13 import DOCXOutlineExporter
import json
import time

# ========== 路径与配置 ==========
OUT_DIR = os.path.abspath(os.path.dirname(__file__))
XML_PATH = os.path.join(OUT_DIR, "formatted_output.xml")
TEMPLATE_PATH = os.path.join(OUT_DIR, "template.idml")
IDML_OUT_PATH = os.path.join(OUT_DIR, "output.idml")  # 可选导出
LOG_PATH = os.path.join(OUT_DIR, "inline_style_debug.log")  # 行内样式&脚注日志

# 始终覆盖同名脚本（不再生成带时间戳）
JSX_PATH = os.path.join(OUT_DIR, "indesign_autoflow_map_levels.jsx")

AUTO_RUN_WINDOWS = True
AUTO_RUN_MACOS = True
AUTO_EXPORT_IDML = True  # 如需脚本结束自动导出 output.idml，改 True
# 是否把运行日志写入文件：开发=True，商用=False，也可用环境变量 INDESIGN_LOG=0/1 覆盖
LOG_WRITE = True

WIN_PROGIDS = [
    "InDesign.Application.2020",
    "InDesign.Application.CC.2020",
    "InDesign.Application.2019",
    "InDesign.Application.CC.2019",
    "InDesign.Application",
]
MAC_APP_NAME = "Adobe InDesign 2020"

# 仅作为“模板缺失样式时”的兜底（不会覆盖模板样式）
BODY_PT = 11
BODY_LEADING = 14
HEADING_BASE_PT = 18
HEADING_STEP_PT = 2
HEADING_MIN_PT = 8
HEADING_EXTRA_LEAD = 3
SPACE_BEFORE_HEAD = 8
SPACE_AFTER_HEAD = 6

# 脚注标记（正文里的小号上标）仅用于“标记”本身
FN_MARK_PT = max(7, BODY_PT - 2)

# 脚注正文段落样式找不到时的兜底字号/行距（只影响脚注内容，不影响正文）
FN_FALLBACK_PT = max(8, BODY_PT - 2)
FN_FALLBACK_LEAD = FN_FALLBACK_PT + 2


# ========== XML 解析（无限层级 + 引用式脚注/尾注；忽略 <meta>/<prop>/<footnotes>/<endnotes>内容） ==========
def _strip_ns(tag):
    return tag.split('}', 1)[-1].lower()


_hn_re = re.compile(r'^h(\d+)$', re.I)
_leveln_re = re.compile(r'^level(\d+)$', re.I)


def _collect_all_text(elem):
    parts = []
    if elem.text: parts.append(elem.text)
    for c in elem:
        parts.append(_collect_all_text(c))
        if c.tail: parts.append(c.tail)
    return "".join(parts)


def _index_notes(root):
    foot_map, end_map = {}, {}
    stack = [root]
    while stack:
        n = stack.pop()
        tag = _strip_ns(n.tag)
        if tag == "footnotes":
            for ch in list(n):
                if _strip_ns(ch.tag) == "footnote":
                    fid = ch.attrib.get("id") or ch.attrib.get("rid") or ch.attrib.get("ref")
                    if fid:
                        foot_map[str(fid)] = _collect_all_text(ch).strip().replace("]]", "】】")
            continue
        if tag == "endnotes":
            for ch in list(n):
                if _strip_ns(ch.tag) == "endnote":
                    eid = ch.attrib.get("id") or ch.attrib.get("rid") or ch.attrib.get("ref")
                    if eid:
                        end_map[str(eid)] = _collect_all_text(ch).strip().replace("]]", "】】")
            continue
        stack.extend(list(n))
    return foot_map, end_map


def _collect_inline_with_notes(elem, foot_map, end_map):
    parts = []
    if elem.text:
        parts.append(elem.text)
    for c in elem:
        tag = _strip_ns(c.tag)

        if tag in ("meta", "prop", "footnotes", "endnotes"):
            if c.tail: parts.append(c.tail)
            continue

        # inline styles
        if tag in ("i", "em"):
            parts.append("[[I]]")
            parts.append(_collect_inline_with_notes(c, foot_map, end_map))
            parts.append("[[/I]]")
            if c.tail: parts.append(c.tail)
            continue
        if tag in ("b", "strong"):
            parts.append("[[B]]")
            parts.append(_collect_inline_with_notes(c, foot_map, end_map))
            parts.append("[[/B]]")
            if c.tail: parts.append(c.tail)
            continue
        if tag == "u":
            parts.append("[[U]]")
            parts.append(_collect_inline_with_notes(c, foot_map, end_map))
            parts.append("[[/U]]")
            if c.tail: parts.append(c.tail)
            continue

        # inline notes
        if tag in ("footnote", "fn"):
            note = _collect_all_text(c).strip().replace("]]", "】】")
            parts.append(f"[[FN:{note}]]")
            if c.tail: parts.append(c.tail)
            continue
        if tag in ("endnote", "en"):
            note = _collect_all_text(c).strip().replace("]]", "】】")
            parts.append(f"[[EN:{note}]]")
            if c.tail: parts.append(c.tail)
            continue

        # referenced notes
        if tag == "fnref":
            rid = c.attrib.get("id") or c.attrib.get("rid") or c.attrib.get("ref")
            parts.append(f"[[FNI:{str(rid)}]]")
            note = foot_map.get(str(rid), "")
            parts.append(f"[[FN:{note}]]" if note else "[*]")
            if c.tail: parts.append(c.tail)
            continue
        if tag == "enref":
            rid = c.attrib.get("id") or c.attrib.get("rid") or c.attrib.get("ref")
            note = end_map.get(str(rid), "")
            parts.append(f"[[EN:{note}]]" if note else "[*]")
            if c.tail: parts.append(c.tail)
            continue

        # --- inline images -> 转成 [[IMG ...]] 标记交给 JSX ---
        if tag in ("img", "image", "graphic", "figureimage", "inlinegraphic"):
            # 尽量兼容多种属性命名
            src = c.attrib.get("src") or c.attrib.get("href") or c.attrib.get("xlink:href") or ""
            # 宽/高可能是 w/width/mm/px，也可能放在 style 里；这里只做最小映射，样式里不解析也不影响排版
            w = c.attrib.get("w") or c.attrib.get("width") or ""
            h = c.attrib.get("h") or c.attrib.get("height") or ""
            align = c.attrib.get("align") or c.attrib.get("placement") or ""
            # 生成 [[IMG ...]]；缺省对齐由 JSX 端处理（默认为 center）
            if src:
                parts.append(f'[[IMG src="{src}" w="{w}" h="{h}" align="{align}"]]')
            if c.tail:
                parts.append(c.tail)
            continue

        parts.append(_collect_inline_with_notes(c, foot_map, end_map))
        if c.tail:
            parts.append(c.tail)
    return "".join(parts)


def extract_paragraphs_with_levels(xml_path):
    tree = ET.parse(xml_path)
    root = tree.getroot()
    foot_map, end_map = _index_notes(root)

    out = []

    def walk(elem, current_level):
        tag = _strip_ns(elem.tag)

        if tag in ("meta", "prop", "footnotes", "endnotes"):
            if elem.tail and elem.tail.strip():
                out.append(("Body", elem.tail.strip()))
            return

        m = _hn_re.match(tag)
        if m:
            lvl = int(m.group(1))
            txt = _collect_inline_with_notes(elem, foot_map, end_map).strip()
            if txt:
                out.append((f"Level{lvl}", txt))
            if elem.tail and elem.tail.strip():
                out.append(("Body", elem.tail.strip()))
            return

        if tag == "p":
            txt = _collect_inline_with_notes(elem, foot_map, end_map).strip()
            if txt:
                out.append(("Body", txt))
            if elem.tail and elem.tail.strip():
                out.append(("Body", elem.tail.strip()))
            return

        if tag == "title":
            lvl = current_level if current_level >= 1 else 1
            txt = _collect_inline_with_notes(elem, foot_map, end_map).strip()
            if txt:
                out.append((f"Level{lvl}", txt))
            if elem.tail and elem.tail.strip():
                out.append(("Body", elem.tail.strip()))
            return

        # 容器：chapter/section/subsection/levelN
        new_level = current_level
        if tag == "chapter":
            new_level = 1
        elif tag == "section":
            new_level = 2
        elif tag == "subsection":
            new_level = 3
        else:
            m2 = _leveln_re.match(tag)
            if m2:
                new_level = int(m2.group(1))

        if elem.text and elem.text.strip() and tag not in ("document", "root"):
            out.append(("Body", elem.text.strip()))

        for c in elem:
            walk(c, new_level)

        if elem.tail and elem.tail.strip():
            out.append(("Body", elem.tail.strip()))

    walk(root, 0)
    return out


# ========== 为 JSX 注入的字符串转义 ==========
def escape_js(s: str) -> str:
    s = s.replace("\\", "\\\\").replace('"', '\\"')
    s = s.replace("\\r\\n", " ").replace("\\r", " ").replace("\\n", " ")
    s = s.replace("\r\n", " ").replace("\r", " ").replace("\n", " ")
    s = re.sub(r"\s+", " ", s)
    return s


# ========== JSX 模板（新增 IMG/TABLE 处理函数与正则） ==========
JSX_TEMPLATE = r"""function smartWrapStr(s){
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


// ExtendScript 没有 Date#toISOString，自己拼一个
function iso() {
  var d = new Date();
  function pad(n){ return (n < 10 ? "0" : "") + n; }
  return d.getUTCFullYear() + "-" +
         pad(d.getUTCMonth() + 1) + "-" +
         pad(d.getUTCDate()) + "T" +
         pad(d.getUTCHours()) + ":" +
         pad(d.getUTCMinutes()) + ":" +
         pad(d.getUTCSeconds()) + "Z";
}
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

    // ====== 日志 ======
    var LOG_FILE   = File("%LOG_PATH%");
    var LOG_WRITE  = %LOG_WRITE%;   // ← Python 注入的总开关：true/false

    function warn(m){ if (LOG_WRITE) log("[WARN] " + m); }
    function err(m){  if (LOG_WRITE) log("[ERR] "  + m); }
    function log(m){
      if (!LOG_WRITE) return;                     // ← 关掉写盘
      var stamp = iso()+" "+m;
      // 1) 尝试写到工程目录日志文件
      try{
        if (LOG_FILE.parent && !LOG_FILE.parent.exists) LOG_FILE.parent.create();
        LOG_FILE.encoding = "UTF-8";
        LOG_FILE.open("a");
        LOG_FILE.writeln(stamp);
        LOG_FILE.close();
      }catch(_){}
    }
    function __logUnitValueFail(msg, err){
      if (__UNITVALUE_FAIL_ONCE) return;
      __UNITVALUE_FAIL_ONCE = true;
      try{ log("[WARN] UnitValue unavailable: " + msg + " err=" + err); }catch(_){}
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
      try{ log("[WARN] width apply failed idx=" + idx + " val=" + widthPt + " trace=" + logs.join("|")); }catch(_){}
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
    var __LAST_IMG_ANCHOR_IDX = -1;     // 用于 addImageAtV2 的“同锚点”检测
    var __DEFAULT_LAYOUT = null;
    var __CURRENT_LAYOUT = null;
    var __DEFAULT_INNER_WIDTH = null;
    var __DEFAULT_INNER_HEIGHT = null;
    var __ENABLE_TRAILING_TRIM = false;
    var __UNITVALUE_FAIL_ONCE = false;
    var __ALLOW_IMG_EXT_FALLBACK = (typeof $.global.__ALLOW_IMG_EXT_FALLBACK !== "undefined")
                                   ? !!$.global.__ALLOW_IMG_EXT_FALLBACK : true;

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
        var docAllowPrev = null;
        var spreadAllowPrev = null;
        try{
          try{ docAllowPrev = doc.allowPageShuffle; doc.allowPageShuffle = true; }catch(_docShuf){}
          if (basePage && basePage.parent && basePage.parent.isValid){
            try{
              spreadAllowPrev = { id: basePage.parent.id, value: basePage.parent.allowPageShuffle };
              basePage.parent.allowPageShuffle = true;
            }catch(_spShuf){}
          }
        }catch(_prep){}
        try{
          if (basePage && basePage.isValid){
            newPage = doc.pages.add(LocationOptions.AFTER, basePage);
          } else {
            newPage = doc.pages.add(LocationOptions.AT_END);
          }
        }catch(eAdd){
          try{ newPage = doc.pages.add(LocationOptions.AT_END); }catch(eAdd2){ newPage = doc.pages.add(); }
        }
        try{
          if (spreadAllowPrev && spreadAllowPrev.id != null){
            try{
              var sp = basePage && basePage.parent && basePage.parent.isValid ? basePage.parent : null;
              if (sp && sp.isValid && sp.id === spreadAllowPrev.id){
                sp.allowPageShuffle = spreadAllowPrev.value;
              }
            }catch(_restoreSp){}
          }
          if (docAllowPrev !== null){
            try{ doc.allowPageShuffle = docAllowPrev; }catch(_restoreDoc){}
          }
          if (docAllowPrev === false && newPage && newPage.parent && newPage.parent.isValid){
            try{ newPage.parent.allowPageShuffle = false; }catch(_restoreNew){}
          }
        }catch(_restore){}
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
      if (__layoutsEqual(__CURRENT_LAYOUT, target)){
        try{ log("[LAYOUT] ensure skip orient=" + (target.pageOrientation||"") + " width=" + target.pageWidthPt + " height=" + target.pageHeightPt); }catch(_){}
        return;
      }
      var prevFrame = (typeof tf !== "undefined" && tf && tf.isValid) ? tf : null;
      var pkt = __createLayoutFrame(target, prevFrame, {});
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
      try { rect.textWrapPreferences.textWrapMode = TextWrapModes.NONE; } catch(_){}

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
  try{
    if (!tf || !tf.isValid) { log("[IMGFLOAT6][ERR] tf invalid"); return null; }
    var f = _normPath(spec && spec.src);
    log("[IMGFLOAT6] resolved file="+(f?f.fsName:"NA"));
    if(!f || !f.exists){ log("[IMGFLOAT6][ERR] file missing: "+(spec&&spec.src)); return null; }

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
      var ip = tf.insertionPoints[-1];
      try{
        var ipIdx = "NA";
        if (ip && ip.isValid) ipIdx = ip.index;
        log("[IMGFLOAT6][DBG] dispatch ip.index=" + ipIdx);
      }catch(_){}
    if (!ip || !ip.isValid) { log("[IMGFLOAT6][ERR] invalid ip"); return null; }
    var anchorIndex = ip.index;

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
        _aos.anchoredPosition = AnchorPosition.ANCHORED;
        _aos.anchorPoint      = AnchorPoint.TOP_LEFT_ANCHOR;
        try{ _aos.lockPosition = false; }catch(_){}
      }
    } catch(_){}
    try { rect.textWrapPreferences.textWrapMode = TextWrapModes.NONE; } catch(_){}
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
    log('[IMGFLOAT6][WARN] gb invalid, use fallback sizing');
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
            log("[IMGFLOAT6][ERR] setBounds failed: " + _boundsErr);
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
    } catch(eSz){ log("[IMGFLOAT6][WARN] size "+eSz); }

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
    return rect;
  }catch(e){
    log("[IMGFLOAT6][ERR] "+e);
    return null;
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
        }catch(e){ try{ log("[WARN] applyTableBorders: "+e); }catch(__){} }
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
        }catch(e){ try{ log("[WARN] _normalizeTableWidth: "+e); }catch(__){} }
    }


                        function addTableHiFi(obj){
      try{
        var rows = obj.rows|0, cols = obj.cols|0;
        if (rows<=0 || cols<=0) return;
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
        var layoutSwitchApplied = false;
        try{
          if (layoutSpec){
            var prevOrientation = __CURRENT_LAYOUT ? __CURRENT_LAYOUT.pageOrientation : null;
            __ensureLayout(layoutSpec);
            var newOrientation = __CURRENT_LAYOUT ? __CURRENT_LAYOUT.pageOrientation : prevOrientation;
            if (layoutSpec.pageOrientation && newOrientation !== prevOrientation){
              layoutSwitchApplied = true;
            }
            log("[TABLE] layout request orient=" + (layoutSpec.pageOrientation||""));
          }else if (__CURRENT_LAYOUT && __DEFAULT_LAYOUT && !__layoutsEqual(__CURRENT_LAYOUT, __DEFAULT_LAYOUT)){
            __ensureLayoutDefault();
          }
        }catch(__layoutErr){
          try{ log("[WARN] ensure layout failed: " + __layoutErr); }catch(__layoutLog){}
        }
        if (layoutSwitchApplied){
          try{
            story.insertionPoints[-1].contents = SpecialCharacters.FRAME_BREAK;
            story.recompose();
          }catch(__frameBreakErr){
            try{ log("[WARN] frame break after layout failed: " + __frameBreakErr); }catch(_){}
          }
        }
        try{ log("[TABLE] begin rows="+rows+" cols="+cols); }catch(__){}
        var doc = app.activeDocument;

        var storyRef = null;
        try{ if (story && story.isValid) storyRef = story; }catch(_){ }
        if (!storyRef){
          try{
            if (curTextFrame && curTextFrame.isValid && curTextFrame.parentStory && curTextFrame.parentStory.isValid){
              storyRef = curTextFrame.parentStory;
            }
          }catch(_){ }
        }
        if (!storyRef){
          try{
            if (typeof tf!=="undefined" && tf && tf.isValid && tf.parentStory && tf.parentStory.isValid){
              storyRef = tf.parentStory;
            }
          }catch(_){ }
        }
        if (!storyRef){
          try{
            if (doc && doc.stories && doc.stories.length>0){
              storyRef = doc.stories[0];
            }
          }catch(_){ }
        }
        if (!storyRef || !storyRef.isValid){
          try{ log("[ERR] addTableHiFi: no valid story"); }catch(__){}
          return;
        }
        story = storyRef;
        try { story.recompose(); } catch(_){ }

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

        function _ensureWritableFrameLocal(storyArg){
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
                var holder = null;
                try{
                    if (ipCheck.parentTextFrames && ipCheck.parentTextFrames.length){
                        holder = ipCheck.parentTextFrames[0];
                    }
                }catch(_){}
                if (!holder || !holder.isValid) holder = frameArg;
                if (!holder || !holder.isValid) return result;

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

                if (approxNeed > available && available >= 0){
                    try{
                        log("[TABLE] pre-break forcing approx=" + approxNeed + " avail=" + available + " rows=" + rowsCount);
                    }catch(__log0){}
                    try{
                        ipCheck.contents = SpecialCharacters.FORCED_FRAME_BREAK;
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
                        log("[TABLE] pre-break result frame=" + (result.frame && result.frame.isValid ? result.frame.id : "NA")
                            + " page=" + (result.page && result.page.isValid ? result.page.name : (result.frame && result.frame.parentPage ? result.frame.parentPage.name : "NA")));
                    }catch(__log1){}
                } else {
                    try{ log("[TABLE] pre-break skip approx=" + approxNeed + " avail=" + available + " rows=" + rowsCount); }catch(__log2){}
                }
            }catch(e){
                try{ log("[WARN] table pre-break failed: " + e); }catch(__){}
            }
            return result;
        }

        var baseFrame = _ensureWritableFrameLocal(story);
        if (!baseFrame || !baseFrame.isValid){
          try{ log("[ERR] addTableHiFi: no writable frame"); }catch(__){}
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
          try{ log("[ERR] addTableHiFi: cannot resolve anchor paragraph"); }catch(__){}
          return;
        }
        var anchorIP = null;
        try{ anchorIP = anchorParagraph.insertionPoints[0]; }catch(_){ }
        if (!anchorIP || !anchorIP.isValid){
          try{ log("[ERR] addTableHiFi: invalid anchor insertion point"); }catch(__){}
          return;
        }

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
          try{ log('[ERR] addTableHiFi: cannot resolve inline insertion point'); }catch(__){}
          return;
        }
        try{
          var __baseFrameId = (baseFrame && baseFrame.isValid) ? baseFrame.id : "NA";
          var __basePageName = (baseFrame && baseFrame.isValid && baseFrame.parentPage && baseFrame.parentPage.isValid)
                                ? baseFrame.parentPage.name : "NA";
          var __anchorIdxDbg = (insertIP && insertIP.isValid) ? insertIP.index : "NA";
          log("[TABLE] anchor pick storyLen=" + __storyLenBefore
              + " frame=" + __baseFrameId + " page=" + __basePageName
              + " ipIdx=" + __anchorIdxDbg);
        }catch(__dbgAnchor){}
        var tbl = null;
        try {
          tbl = insertIP.tables.add({ bodyRowCount: rows, columnCount: cols });
        } catch(eAdd) {
          try{ log('[ERR] addTableHiFi: table create failed ' + eAdd); }catch(__){}
          return;
        }
        try{
          var __colLenInit = 0;
          try{ __colLenInit = tbl.columns.length; }catch(__colErr){}
          log("[TABLE] init columns expected=" + cols + " actual=" + __colLenInit);
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
          try{ log("[WARN] degrade rowspan rows=" + spanRows + " at r=" + r + " c=" + c); }catch(__warnLog){}
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

        for (var r=0; r<rows; r++){
          var rowEntries = obj.data[r] || [];
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

        for (var r=0; r<rows; r++){
          for (var c=0; c<cols; c++){
            if (skipPos[r][c] && !cellPlan[r][c]) continue;
            var cellSpec2 = cellPlan[r][c];
            if (!cellSpec2){
              var src = (obj.data[r] && obj.data[r][c]) ? obj.data[r][c] : {text:""};
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
        try{ fit(FitOptions.FRAME_TO_CONTENT); }catch(__){}
        try{
          var __gbFit = geometricBounds;
          var __hFit = (__gbFit && __gbFit.length>=3) ? (__gbFit[2]-__gbFit[0]) : "NA";
          log("[TABLE] frame fit height=" + __hFit);
        }catch(_){}
        try{
          var __offsetIdx = (tbl && tbl.isValid && tbl.storyOffset && tbl.storyOffset.isValid) ? tbl.storyOffset.index : "NA";
          var __storyLenAfter = 0;
          try{ __storyLenAfter = story.characters.length; }catch(__lenErr){}
          log("[TABLE] placed idx=" + __offsetIdx + " storyLenNow=" + __storyLenAfter);
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
              log("[TABLE] post-ip idx=" + __postIdxDbg + " frame=" + __tfIdDbg + " page=" + __pgDbg);
            }catch(__postDbg){}
          }
        }catch(__ipErr){}
        if (degradeNotice){
          try{
            var __noticeIP = null;
            if (__postTableIP && __postTableIP.isValid && story && story.isValid){
              try{ __noticeIP = story.insertionPoints[__postTableIP.index]; }catch(__noticeIdxErr){}
            }
            if ((!__noticeIP || !__noticeIP.isValid) && tbl && tbl.isValid){
              try{
                var __anchorIdx = (tbl.storyOffset && tbl.storyOffset.isValid) ? tbl.storyOffset.index : null;
                if (__anchorIdx != null){
                  try{ __noticeIP = story.insertionPoints[__anchorIdx+1]; }catch(__noticeIdxErr2){}
                }
              }catch(__noticeIdxEval){}
            }
            if ((!__noticeIP || !__noticeIP.isValid) && story && story.isValid){
              try{ __noticeIP = story.insertionPoints[-1]; }catch(__noticeFallback){}
            }
            if (__noticeIP && __noticeIP.isValid){
              var __noticeMsg = "\u3010\u8868\u683c\u63d0\u793a\u3011\u8be5\u8868\u5305\u542b\u8d85\u8fc7 "
                                + MAX_ROWSPAN_INLINE
                                + " \u884c\u7684\u7eb5\u5411\u5408\u5e76\uff0c\u7cfb\u7edf\u5df2\u62c6\u5206\u5bfc\u5165\uff0c\u8bf7\u6838\u5bf9\u539f\u7a3f\u5e76\u624b\u52a8\u8865\u9f50\u9057\u6f0f\u5185\u5bb9\u3002";
              try{ __noticeIP.contents = __noticeMsg + "\r"; }catch(__noticeInsert){}
              try{ log("[TABLE] degrade notice inserted idx=" + __noticeIP.index); }catch(__noticeLog){}
            }
          }catch(__noticeBlockErr){
            try{ log("[WARN] degrade notice insert failed: " + __noticeBlockErr); }catch(__noticeWarn){}
          }
        }
        try{ __LAST_IMG_ANCHOR_IDX = -1; }catch(__resetErr){}

        try{
          story.recompose();
          if (typeof flushOverflow==="function" && tf && tf.isValid){
            var st = flushOverflow(story, page, tf);
            page = st.page; tf = st.frame; story = tf.parentStory; curTextFrame = tf;
          }
        }catch(e){ log("[WARN] flush after table failed: " + e); }
      }catch(e){
        log("[ERR] addTableHiFi " + e);
      }
    }
    function createFootnoteAt(ip, content, idForDisplay){
        if(!ip || !ip.isValid) return null;
        var doc = app.activeDocument, story = ip.parentStory;
        var fn = null, ok = false;
        try { fn = story.footnotes.add(LocationOptions.AFTER, ip); ok = (fn && fn.isValid); } catch(e){}
        if (!ok) { try { fn = story.footnotes.add(ip); ok = (fn && fn.isValid); } catch(e){} }
        if (!ok) { try { fn = doc.footnotes.add(ip);   ok = (fn && fn.isValid); } catch(e){} }
        if (!ok) { return null; }
        try { fn.texts[0].contents = content; }
        catch(_){ try { fn.contents = content; } catch(__){ try { fn.insertionPoints[-1].contents = content; } catch(___) {} } }
        if (idForDisplay != null) {
            try { fn.texts[0].insertionPoints[0].contents = String(idForDisplay) + " "; } catch(_){}
        }
        try { if (!FOOTNOTE_PS || !FOOTNOTE_PS.isValid) FOOTNOTE_PS = ensureFootnoteParaStyle(doc);
              fn.texts[0].paragraphs.everyItem().appliedParagraphStyle = FOOTNOTE_PS; } catch(_){}
        return fn;
    }

    function createEndnoteAt(ip, content){
        if(!ip || !ip.isValid) return null;
        var doc = app.activeDocument, story = ip.parentStory;
        var en = null, ok = false;
        try { if (ip.createEndnote) { en = ip.createEndnote(); ok = (en && en.isValid); } } catch(e){ }
        if (!ok) { try { en = story.endnotes.add(ip); ok = (en && en.isValid); } catch(e){ } }
        if (!ok) { try { en = doc.endnotes.add(ip);   ok = (en && en.isValid); } catch(e){ } }
        if (!ok) { return null; }
        try { en.endnoteText.contents = content; }
        catch(_){ try { en.texts[0].contents = content; } catch(__){ try { en.contents = content; } catch(___) {} } }
        try { if (!ENDNOTE_PS || !ENDNOTE_PS.isValid) ENDNOTE_PS = ensureEndnoteParaStyle(app.activeDocument);
              (en.endnoteText || en.texts[0] || en).paragraphs.everyItem().appliedParagraphStyle = ENDNOTE_PS; } catch(_){}
        return en;
    }

    // —— 段落插入：扩展识别 [[IMG ...]] / [[TABLE {...}]] ——
    function addParaWithNotes(story, styleName, raw) {
        var s = app.activeDocument.paragraphStyles.itemByName(styleName);
        try { log("[PARA] style=" + styleName + " len=" + String(raw||"").length); } catch(_){}
        if (!s.isValid) { s = app.activeDocument.paragraphStyles.add({name:styleName}); }

        var text = String(raw).replace(/^\s+|\s+$/g, "");
        if (text.length === 0) return;

        // ★ 正则扩展：新增 IMG/TABLE（修复 I/B/U 与 IMG/TABLE 的匹配）
        var re = /\[\[FNI:(\d+)\]\]|\[\[(FN|EN):(.*?)\]\]|\[\[(\/?)(I|B|U)\]\]|\[\[IMG\s+([^\]]+)\]\]|\[\[TABLE\s+(\{[\s\S]*?\})\]\]/g;
        var last = 0, m;
        var st = {i:0, b:0, u:0};
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
                PENDING_FN_ID = parseInt(m[1], 10);
            } else if (m[2]) {
                var noteType = m[2];
                var noteContent = m[3];
                var ip = story.insertionPoints[-1];
                try { if (noteType === "FN") createFootnoteAt(ip, noteContent, PENDING_FN_ID);
                      else createEndnoteAt(ip, noteContent); } catch(_){}
                PENDING_FN_ID = null;

            } else if (m[6]) {
                try{ log("[IMGDBG] enter [[IMG]] attrs=" + m[6]); }catch(_){}
                var kv = m[6], spec = {};
                try{
                  log('[IMGDBG] enter [[IMG]] lastIdx='
                      + (typeof __LAST_IMG_ANCHOR_IDX==='number'?__LAST_IMG_ANCHOR_IDX:'NA'));
                }catch(_){}
                kv.replace(/(\w+)=['"“”]([^'"”]*)['"”]/g, function(_,k,v){ spec[k]=v; return _; });

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
                      log("[IMG-DISPATCH] src="+spec.src+" inline="+inl+" posH="+(spec.posH||"")+" posV="+(spec.posV||""));
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
    }

    // 打开模板、清空页面框等（保持你原逻辑）
    var templateFile = File("%TEMPLATE_PATH%");
    if (!templateFile.exists) { alert("未找到模板文件 template.idml"); return; }
    var doc = app.open(templateFile);

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

    function flushOverflow(currentStory, lastPage, lastFrame) {
        // 说明：原先用 story.characters.length 判断“是否前进”，会误判为卡住（字符总数不随分页变化）。
        // 最小修复：移除早停判定；只要 still overset 就继续加页并接链，直到不 overset 或达到 MAX_PAGES。
        var MAX_PAGES = 20;
        for (var guard = 0; currentStory && currentStory.overflows && guard < MAX_PAGES; guard++) {
            var pkt = __createLayoutFrame(__CURRENT_LAYOUT, lastFrame, {afterPage:lastPage, forceBreak:false});
            if (!pkt || !pkt.frame || !pkt.page) { break; }
            lastPage  = pkt.page;
            lastFrame = pkt.frame;

            try { currentStory.recompose(); } catch(_) {}
            try { app.activeDocument.recompose(); } catch(_) {}
            $.sleep(10);
        }
        if (currentStory && currentStory.overflows) {
            try { log("[WARN] flushOverflow guard hit; overset still true"); } catch(_){}
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
            try{ log("[WARN] trim trailing frames failed: " + eTrim); }catch(_){}
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
            try{ log("[WARN] trim trailing pages failed: " + ePg); }catch(_){}
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

    var page  = doc.pages[0];
    try{ log("[LOG] script boot ok; page="+doc.pages.length); }catch(_){}

    var tf    = createTextFrameOnPage(page, __DEFAULT_LAYOUT);
    if (__DEFAULT_INNER_WIDTH === null) __DEFAULT_INNER_WIDTH = _innerFrameWidth(tf);
    if (__DEFAULT_INNER_HEIGHT === null) __DEFAULT_INNER_HEIGHT = _innerFrameHeight(tf);
    try{ log("[LAYOUT] default inner width=" + __DEFAULT_INNER_WIDTH + " height=" + __DEFAULT_INNER_HEIGHT); }catch(_defaultLog){}
    var story = tf.parentStory;
    curTextFrame = tf; 

    var firstChapterSeen = false;

    __ADD_LINES__
    var tail = flushOverflow(story, page, tf);
    page  = tail.page;
    tf    = tail.frame;
    story = tf.parentStory;
    curTextFrame = tf;
    __trimTrailingEmptyFrames(story);
    __trimTrailingEmptyPages(doc);
    try { fixAllTables(); } catch(_) {}
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
    }catch(e){ try{ log("[WARN] fixAllTables: "+e); }catch(__){} }
}
"""


def build_style_lines(levels_used):
    lines = []
    lines.append(f"ensureStyle('Body', {BODY_PT}, {BODY_LEADING}, 0, 0);")
    for lvl in sorted(levels_used):
        size = max(HEADING_MIN_PT, HEADING_BASE_PT - (lvl - 1) * HEADING_STEP_PT)
        lead = size + HEADING_EXTRA_LEAD
        lines.append(
            f"ensureStyle('Level{lvl}', {size}, {lead}, {SPACE_BEFORE_HEAD}, {SPACE_AFTER_HEAD});"
        )
    return "\n        ".join(lines)


def build_toc_entries(levels_used):
    # （保留你之前的 TOC 部分，这里不再额外插入）
    return ""


def write_jsx(jsx_path, paragraphs, skip_images=False):
    add_lines = []
    levels_used = set()

    add_lines.append(
        "function onNewLevel1(){ var pkt = startNewChapter(story, page, tf); story=pkt.story; page=pkt.page; tf=pkt.frame; }")
    add_lines.append("firstChapterSeen = false;")

    img_pattern = re.compile(r'\[\[IMG\s+[^\]]+\]\]', re.I)
    if skip_images:
        paragraphs = [
            (style, img_pattern.sub(" ", text))
            for style, text in paragraphs
        ]
    expanded_paragraphs = []
    for style, text in paragraphs:
        start = 0
        for match in img_pattern.finditer(text):
            pre = text[start:match.start()]
            if pre:
                expanded_paragraphs.append((style, pre))
            expanded_paragraphs.append((style, match.group(0)))
            start = match.end()
        tail = text[start:]
        if tail:
            expanded_paragraphs.append((style, tail))
    if expanded_paragraphs:
        paragraphs = expanded_paragraphs

    for style, text in paragraphs:
        sty = style
        if sty.lower().startswith("level"):
            try:
                n = int(sty[5:])
                levels_used.add(n)
                sty = f"Level{n}"
            except:
                pass
        elif sty.lower() == "body":
            sty = "Body"

        esc = escape_js(text)

        # === 新增：当整段就是 TABLE/IMG 或 原生 <table>/<img> 时，直落 ===
        m_tbl = re.match(r'^\s*\[\[TABLE\s+(\{[\s\S]*\})\s*\]\]\s*$', text)
        m_img = re.match(r'^\s*\[\[IMG\s+([^\]]+)\]\]\s*$', text)
        m_xmlt = re.match(r'^\s*<table\b[\s\S]*</table>\s*$', text, flags=re.I)
        m_xmli = re.match(r'^\s*<img\b[^>]*>\s*$', text, flags=re.I)

        if sty == "Level1":
            add_lines.append(
                "if (firstChapterSeen) { var __fl = flushOverflow(story, page, tf); story = __fl.frame.parentStory; page = __fl.page; tf = __fl.frame; onNewLevel1(); } else { firstChapterSeen = true; }")

        only_img = False
        m_img = re.match(r'^\s*(\[\[IMG\s+.+?\]\])\s*$', text, flags=re.I)
        if not m_img:
            m_img = re.search(r'\[\[IMG\s+(.+?)\]\]', text, re.I)
        else:
            only_img = True

        if m_tbl:
            try:
                obj = json.loads(m_tbl.group(1))
            except Exception:
                obj = eval("(" + m_tbl.group(1) + ")")
            rows = int(obj.get("rows", 0));
            cols = int(obj.get("cols", 0));
            data = obj.get("data", [])
            add_lines.append('addTableHiFi(%s);' % (json.dumps(obj, ensure_ascii=False)))
            continue
        elif m_img and not skip_images:
            # 解析 [[IMG ...]] 的属性为 kv
            add_lines.append("__ensureLayoutDefault();")
            kv = dict(re.findall(r'(\w+)=["\'“”]([^"\'”]*)["\'”]', m_img.group(1)))

            def _esc(s: str) -> str:
                return (s or "").replace("\\", "\\\\").replace('"', '\\"')

            src = _esc(kv.get("src", ""))
            w = kv.get("w", "") or ""
            h = kv.get("h", "") or ""
            align = kv.get("align", "center")
            inline = kv.get("inline", "") or ""
            wrap = kv.get("wrap", "") or ""
            posH = kv.get("posH", "") or ""
            posV = kv.get("posV", "") or ""
            offX = kv.get("offX", "") or ""
            offY = kv.get("offY", "") or ""
            distT = kv.get("distT", "") or ""
            distB = kv.get("distB", "") or ""
            distL = kv.get("distL", "") or ""
            distR = kv.get("distR", "") or ""
            sb = kv.get("spaceBefore", "6")
            sa = kv.get("spaceAfter", "6")
            cap = _esc(kv.get("caption", "") or "")

            add_lines.append(f'''(function(){{
              log("[PY][m_img] {src} inline={inline}");
              try {{
                // 0) 环境检查
                log("[DBG] typeof addFloatingImage=" + (typeof addFloatingImage)
                    + " typeof addImageAtV2=" + (typeof addImageAtV2)
                    + " typeof _normPath=" + (typeof _normPath));
                log("[DBG] tf=" + (tf&&tf.isValid) + " story=" + (story&&story.isValid) + " page=" + (page&&page.isValid));

                // 1) 溢出兜底
                try{{ if(typeof flushOverflow==="function"){{ var _rs=flushOverflow(story,page,tf);
                  if(_rs&&_rs.frame&&_rs.page){{ page=_rs.page; tf=_rs.frame; story=tf.parentStory; curTextFrame=tf; }} }} }}catch(_){{
                }}

                // 2) 锚点
                var ip=(tf&&tf.isValid)?_safeIP(tf):story.insertionPoints[-1];
                // 3) 路径
                var f=_normPath("{src}");
                log("[DBG] _normPath ok=" + (!!f) + " exists=" + (f&&f.exists ? "Y":"N") + " fsName=" + (f?f.fsName:"NA"));

                if(f&&f.exists){{
                  var spec={{src:f.fsName,w:"{w}",h:"{h}",align:"{align}",spaceBefore:{sb},spaceAfter:{sa},caption:"{cap}",
                            inline:"{inline}",wrap:"{wrap}",posH:"{posH}",posV:"{posV}",offX:"{offX}",offY:"{offY}",
                            distT:"{distT}",distB:"{distB}",distL:"{distL}",distR:"{distR}",forceBlock:{str(only_img).lower()} }};
                  var inl=_trim(spec.inline); // \u517c\u5bb9 InDesign 2020
                  log("[IMG-DISPATCH] src="+spec.src+" inline="+inl+" posH="+(spec.posH||"")+" posV="+(spec.posV||""));

                  if(inl==="0"||/^false$/i.test(inl)){{
                    log("[DBG] dispatch -> addFloatingImage");
                    var rect=addFloatingImage(tf,story,page,spec);
                    if(rect&&rect.isValid) log("[IMG] ok (float): " + spec.src);
                  }} else {{
                    log("[DBG] dispatch -> addImageAtV2");
                    var rect=addImageAtV2(ip,spec);
                    if(rect&&rect.isValid) log("[IMG] ok (inline): " + spec.src);
                  }}
                }} else {{
                  log("[IMG] missing: {src}");
                }}
              }} catch(e) {{
                log("[IMG][EXC] " + e);
              }}
            }})();''')
            continue
        elif m_xmlt and not skip_images:
            try:
                root = ET.fromstring(text)
                rows_data = []
                for tr in root.findall('.//tr'):
                    row = []
                    for td in tr.findall('.//td'):
                        parts = []
                        if td.text and td.text.strip(): parts.append(td.text.strip())
                        for ch in list(td):
                            tag = _strip_ns(ch.tag)
                            if tag == "p":
                                parts.append(''.join(ch.itertext()).strip())
                            elif tag == "img":
                                s = ch.get("src", "") or "";
                                w = ch.get("w", "") or "";
                                h = ch.get("h", "") or ""
                                parts.append('[[IMG src="%s" w="%s" h="%s"]]' % (s, w, h))
                            if ch.tail and ch.tail.strip(): parts.append(ch.tail.strip())
                        row.append("\n".join([x for x in parts if x]))
                    rows_data.append(row)
                cols = max([len(r) for r in rows_data]) if rows_data else 0
                add_lines.append('addTableHiFi(%s);' % (json.dumps(obj, ensure_ascii=False)))
                continue
            except Exception:
                pass
        elif m_xmli and not skip_images:
            add_lines.append("__ensureLayoutDefault();")
            # 处理整段是 <img ...> 的情况（原生 XML/HTML 片段）
            import xml.etree.ElementTree as ET

            def _esc(s: str) -> str:
                return (s or "").replace("\\", "\\\\").replace('"', '\\"')

            try:
                root = ET.fromstring(text)
                # 兼容 src/href/xlink:href
                src = _esc(
                    root.get("src", "") or root.get("href", "") or root.get("{http://www.w3.org/1999/xlink}href", ""))

                # 尺寸与排版属性（都允空字符串，JS 端自行解释）
                w = root.get("w", "") or root.get("width", "") or ""
                h = root.get("h", "") or root.get("height", "") or ""
                align = root.get("align", "center")
                inline = root.get("inline", "") or ""
                wrap = root.get("wrap", "") or ""
                posH = root.get("posH", "") or ""
                posV = root.get("posV", "") or ""
                offX = root.get("offX", "") or ""
                offY = root.get("offY", "") or ""
                distT = root.get("distT", "") or ""
                distB = root.get("distB", "") or ""
                distL = root.get("distL", "") or ""
                distR = root.get("distR", "") or ""
                sb = root.get("spaceBefore", "6")
                sa = root.get("spaceAfter", "6")
                cap = _esc(root.get("caption", "") or "")
            except Exception:
                # 解析失败则回退为普通段落处理
                continue

            add_lines.append(f'''(function(){{
        log("[PY][m_xmli] {src}");
        try{{ if(typeof flushOverflow==="function"){{ var _rs=flushOverflow(story,page,tf);
        if(_rs&&_rs.frame&&_rs.page){{ page=_rs.page; tf=_rs.frame; story=tf.parentStory; curTextFrame=tf; }} }} }}catch(_){{
        }}
        var ip=(tf&&tf.isValid)?_safeIP(tf):story.insertionPoints[-1];
        try{{
          var para=ip.paragraphs[0]; var p0=(para&&para.isValid)?para.insertionPoints[0]:null;
          var h0=(p0&&p0.isValid&&p0.parentTextFrames&&p0.parentTextFrames.length)?p0.parentTextFrames[0]:null;
          if(h0&&h0.isValid&&tf&&tf.isValid&&h0.id!==tf.id){{ ip.contents="\\r"; try{{story.recompose();}}catch(__){{}} ip=tf.insertionPoints[-1]; }}
        }}catch(__){{}}
        var f=_normPath("{src}");
        if(f&&f.exists){{
          var spec={{src:f.fsName,w:"{w}",h:"{h}",align:"{align}",spaceBefore:{sb},spaceAfter:{sa},caption:"{cap}",
                    inline:"{inline}",wrap:"{wrap}",posH:"{posH}",posV:"{posV}",offX:"{offX}",offY:"{offY}",
                    distT:"{distT}",distB:"{distB}",distL:"{distL}",distR:"{distR}"}};
          var inl=_trim(spec.inline);
          if(inl==="0"||/^false$/i.test(inl)){{
            var rect=addFloatingImage(tf,story,page,spec);
            if(rect&&rect.isValid) log("[IMG] ok (float): "+spec.src);
          }} else {{
            var rect=addImageAtV2(ip,spec);
            if(rect&&rect.isValid) log("[IMG] ok (inline): "+spec.src);
          }}
        }} else {{
          log("[IMG] missing: {src}");
        }}
        }})();''')
            continue

        # 默认：仍走 addParaWithNotes（它现在也能识别行内 IMG/TABLE）
        add_lines.append("__ensureLayoutDefault();")
        add_lines.append(f'addParaWithNotes(story, "{sty}", "{esc}");')

    style_lines = build_style_lines(levels_used)

    # 构造图片检索目录（新增）
    img_dirs = [
        OUT_DIR,
        os.path.join(OUT_DIR, "assets"),
        os.path.dirname(XML_PATH) or OUT_DIR,
        os.path.join(os.path.dirname(XML_PATH) or OUT_DIR, "assets"),
    ]
    # 去重 & 规范化
    _seen = set();
    _norm = []
    for d in img_dirs:
        if not d: continue
        dd = os.path.abspath(d)
        if dd not in _seen:
            _seen.add(dd);
            _norm.append(dd)

    jsx = JSX_TEMPLATE
    jsx = jsx.replace("%TEMPLATE_PATH%", TEMPLATE_PATH.replace("\\", "\\\\"))
    jsx = jsx.replace("%OUT_IDML%", IDML_OUT_PATH.replace("\\", "\\\\"))
    jsx = jsx.replace("%AUTO_EXPORT%", "true" if AUTO_EXPORT_IDML else "false")
    jsx = jsx.replace("%BODY_PT%", str(BODY_PT))
    jsx = jsx.replace("%BODY_LEADING%", str(BODY_LEADING))
    jsx = jsx.replace("%FN_MARK_PT%", str(FN_MARK_PT))
    jsx = jsx.replace("%FN_FALLBACK_PT%", str(FN_FALLBACK_PT))
    jsx = jsx.replace("%FN_FALLBACK_LEAD%", str(FN_FALLBACK_LEAD))
    jsx = jsx.replace("%LOG_PATH%", LOG_PATH.replace("\\", "/"))
    jsx = jsx.replace("%LOG_WRITE%", "true" if LOG_WRITE else "false")  # ← 新增
    jsx = jsx.replace("__STYLE_LINES__", style_lines)
    jsx = jsx.replace("__ADD_LINES__", "\n    ".join(add_lines))
    jsx = jsx.replace("%IMG_DIRS_JSON%", json.dumps(_norm).replace("\\", "\\\\"))

    with open(jsx_path, "w", encoding="utf-8") as f:
        f.write(jsx)
    print("[OK] JSX 写入:", jsx_path)
    if LOG_WRITE:
        print("[INFO] 日志写入:", LOG_PATH)
    # 在 write_jsx() 末尾、写完 add_lines 之后临时加一行：
    print("[DEBUG] JSX 是否包含 addImageAtV2：", any("addImageAtV2(" in ln for ln in add_lines))


# ========== 调用 InDesign ==========
def run_indesign_windows(jsx_path):
    try:
        import win32com.client  # pip install pywin32
    except Exception as e:
        print("[WARN] 未安装 pywin32：", e)
        return False

    app = None
    for pid in WIN_PROGIDS:
        try:
            app = win32com.client.Dispatch(pid)
            print(f"[OK] 连接 InDesign: {pid}")
            break
        except Exception:
            app = None

    if not app:
        print("[ERR] 未找到 InDesign COM 接口")
        return False

    try:
        app.DoScript(jsx_path, 1246973031)  # 1246973031 = JavaScript
        print("[OK] 已执行 JSX")
        return True
    except Exception as e:
        print("[ERR] DoScript 执行失败：", e)
        return False


def run_indesign_macos(jsx_path):
    jsx_abs = os.path.abspath(jsx_path)
    jsx_escaped = jsx_abs.replace('"', '\\"')

    env_name = os.environ.get("MAC_APP_NAME", "").strip()
    candidates = []
    if env_name:
        candidates.append(env_name)
    try:
        candidates.append(MAC_APP_NAME)
    except NameError:
        pass
    candidates += [
        "Adobe InDesign 2025",
        "Adobe InDesign 2024",
        "Adobe InDesign 2023",
        "Adobe InDesign 2022",
        "Adobe InDesign 2021",
        "Adobe InDesign 2020",
        "Adobe InDesign CC 2019",
        "Adobe InDesign CC 2018",
        "Adobe InDesign CC 2017",
        "Adobe InDesign CC"
    ]

    tried = []
    for app_name in candidates:
        if not app_name or app_name in tried:
            continue
        tried.append(app_name)
        osa = f'''tell application "{app_name}"
            activate
            do script (POSIX file "{jsx_escaped}") language javascript
        end tell'''
        print(f"[macOS] 尝试调用 InDesign: {app_name}")
        try:
            p = subprocess.Popen(["osascript", "-e", osa], stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
            out, err = p.communicate()
            if p.returncode == 0:
                print("[OK] osascript 返回 0；已向 InDesign 发送脚本。")
                if out.strip():
                    print("[osascript out]", out.strip())
                return True
            else:
                print("[ERR] osascript 返回码:", p.returncode)
                if err.strip():
                    print("[osascript err]", err.strip())
        except Exception as e:
            print(f"[ERR] 调用 {app_name} 失败：{e}")

    print("[ERR] 无法调用任何已知的 InDesign 应用名。可设置环境变量 MAC_APP_NAME=你的应用名 再试。")
    return False


def main():
    parser = argparse.ArgumentParser(
        description="DOCX → XML → JSX → InDesign 自动排版工具"
    )
    parser.add_argument(
        "docx",
        nargs="?",
        help="待转换的 DOCX 文件（默认使用脚本目录下的 1.docx）",
    )
    parser.add_argument(
        "--mode",
        choices=("heading", "regex", "hybrid"),
        default="heading",
        help="DOCXOutlineExporter 的解析模式（默认 heading）",
    )
    parser.add_argument(
        "--skip-docx",
        action="store_true",
        help="跳过 DOCX→XML 流程，直接使用现有 XML（需保证 XML 存在）",
    )
    parser.add_argument(
        "--xml-path",
        help="显式指定 XML 输入/输出路径（默认 formatted_output.xml）",
    )
    parser.add_argument(
        "--no-run",
        action="store_true",
        help="只生成 XML/JSX，不调用 InDesign",
    )
    parser.add_argument(
        "--no-images",
        action="store_true",
        help="skip inserting images when generating JSX",
    )
    args = parser.parse_args()

    global XML_PATH
    if args.xml_path:
        XML_PATH = os.path.abspath(args.xml_path)

    if args.skip_docx:
        if not os.path.exists(XML_PATH):
            print(f"[ERR] --skip-docx 指定但未找到 XML：{XML_PATH}")
            return
        print(f"[INFO] 跳过 DOCX → XML，直接使用：{XML_PATH}")
    else:
        input_file = os.path.abspath(args.docx or "1.docx")
        if not os.path.exists(input_file):
            print(f"[ERR] 找不到 DOCX：{input_file}")
            return
        exporter = DOCXOutlineExporter(input_file, mode=args.mode)
        exporter.process(XML_PATH)

    paragraphs = extract_paragraphs_with_levels(XML_PATH)
    print(f"[INFO] 解析了 {len(paragraphs)} 段；示例： {paragraphs[:3]}")

    write_jsx(JSX_PATH, paragraphs, skip_images=args.no_images)

    ran = False
    if not args.no_run:
        if AUTO_RUN_WINDOWS and sys.platform.startswith("win"):
            ran = run_indesign_windows(JSX_PATH)
        elif AUTO_RUN_MACOS and sys.platform == "darwin":
            ran = run_indesign_macos(JSX_PATH)

    print("\n=== 完成 ===")
    print("XML: ", XML_PATH)
    print("JSX: ", JSX_PATH)
    print("LOG: ", LOG_PATH)
    print("IDML:", IDML_OUT_PATH)
    if ran:
        print("InDesign 已执行 JSX。若设置 AUTO_EXPORT_IDML=True，将在脚本目录生成 output.idml。")



if __name__ == "__main__":
    main()
