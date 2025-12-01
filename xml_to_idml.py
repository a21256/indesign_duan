# -*- coding: utf-8 -*-
import os, sys, subprocess, re
import argparse
import xml.etree.ElementTree as ET
from docx_to_xml_outline_notes_v13 import DOCXOutlineExporter
import json
import time
import threading
from dataclasses import dataclass, field
from typing import Dict, Optional
from pipeline_logger import PipelineLogger

def _runtime_base_dir() -> str:
    """
    Resolve a stable base dir for outputs/assets.
    Order: NUITKA_ONEFILE_PARENT (real exe dir) > argv[0] dir > sys.executable dir > this file dir.
    This avoids writing to the onefile temp extraction dir that gets cleaned up.
    """
    env_parent = os.environ.get("NUITKA_ONEFILE_PARENT")
    if env_parent and os.path.isdir(env_parent):
        return os.path.abspath(env_parent)

    argv0 = sys.argv[0] if sys.argv else None
    if argv0 and os.path.exists(argv0):
        return os.path.abspath(os.path.dirname(argv0))

    exe_path = getattr(sys, "executable", None)
    if exe_path and os.path.exists(exe_path):
        return os.path.abspath(os.path.dirname(exe_path))

    return os.path.abspath(os.path.dirname(__file__))

# ========== 路径与配置 ==========
OUT_DIR = _runtime_base_dir()
XML_PATH = os.path.join(OUT_DIR, "formatted_output.xml")
TEMPLATE_PATH = os.path.join(OUT_DIR, "template.idml")
IDML_OUT_PATH = os.path.join(OUT_DIR, "output.idml") 
LOG_PATH = os.path.join(OUT_DIR, "inline_style_debug.log") 

JSX_PATH = os.path.join(OUT_DIR, "indesign_autoflow_map_levels.jsx")
JSX_TEMPLATE_PATH = os.path.join(OUT_DIR, "templates", "indesign_autoflow_map_levels.tpl.jsx")  # optional external JSX template
JSX_FRAGMENT_DIR = os.path.join(OUT_DIR, "templates", "jsx")
JSX_FRAGMENTS = {
    "UTIL": "util.js",
    "LAYOUT": "layout.js",
    "TABLE": "table.js",
    "IMAGE": "image.js",
    "ENTRY": "entry.js",
}

AUTO_RUN_WINDOWS = True
AUTO_RUN_MACOS = True
AUTO_EXPORT_IDML = True 
LOG_WRITE = False

TABLE_BODY_PAR_STYLE = "TableBody"
TABLE_BODY_PAR_STYLE_FALLBACK = "DocxTable"
TABLE_BODY_PAR_STYLE_BASE = "Body"
TABLE_BODY_PAR_STYLE_AUTO = "__DocxTableAuto"

PROGRESS_HEARTBEAT_MS = 15000 
PROGRESS_CONSOLE_ACTIVE = False
PROGRESS_CONSOLE_LEN = 0

def _write_progress_console_line(text: str, *, final: bool = False):
    global PROGRESS_CONSOLE_ACTIVE, PROGRESS_CONSOLE_LEN
    try:
        pad = max(0, PROGRESS_CONSOLE_LEN - len(text))
        sys.stdout.write("\r" + text + (" " * pad))
        if final:
            sys.stdout.write("\n")
            PROGRESS_CONSOLE_LEN = 0
            PROGRESS_CONSOLE_ACTIVE = False
        else:
            PROGRESS_CONSOLE_LEN = len(text)
            PROGRESS_CONSOLE_ACTIVE = True
        sys.stdout.flush()
    except Exception:
        print(text)
        if final:
            try:
                sys.stdout.flush()
            except Exception:
                pass


def _emit_progress_console(line: str):
    text = (line or "").strip()
    if not text:
        return
    parts = text.split("\t", 2)
    message = ""
    stamp = ""
    if len(parts) == 3:
        _, stamp, message = parts
    elif len(parts) == 2:
        stamp, message = parts
    else:
        message = text
    body = message.strip()
    body_upper = body.upper()
    if "[PROGRESS]" not in body_upper:
        return
    stamp = stamp.strip()
    if stamp:
        console_line = f"[PROGRESS] {stamp} {body}".strip()
    else:
        console_line = f"[PROGRESS] {body}".strip()
    final = "[PROGRESS][COMPLETE]" in body_upper
    _write_progress_console_line(console_line, final=final)

def _watch_jsx_progress(log_path: str, stop_event: "threading.Event"):
    fh = None
    try:
        while not stop_event.is_set():
            try:
                if fh is None:
                    if not os.path.exists(log_path):
                        if stop_event.wait(0.5):
                            break
                        continue
                    fh = open(log_path, "r", encoding="utf-8", errors="ignore")
                    fh.seek(0, os.SEEK_END)
                line = fh.readline()
                if not line:
                    try:
                        size = os.path.getsize(log_path)
                        if fh.tell() > size:
                            fh.seek(0)
                    except Exception:
                        pass
                    if stop_event.wait(0.5):
                        break
                    continue
                _emit_progress_console(line)
            except FileNotFoundError:
                if fh is not None:
                    try:
                        fh.close()
                    except Exception:
                        pass
                    fh = None
                if stop_event.wait(0.5):
                    break
            except Exception:
                if fh is not None:
                    try:
                        fh.close()
                    except Exception:
                        pass
                    fh = None
                if stop_event.wait(0.5):
                    break
    finally:
        if fh is not None:
            try:
                fh.close()
            except Exception:
                pass


def _start_progress_monitor():
    path = LOG_PATH
    if not path:
        return None
    stop_event = threading.Event()
    thread = threading.Thread(
        target=_watch_jsx_progress, args=(path, stop_event), daemon=True
    )
    thread.start()
    return (stop_event, thread)


def _stop_progress_monitor(token):
    if not token:
        return
    stop_event, thread = token
    try:
        stop_event.set()
    except Exception:
        pass
    if thread:
        thread.join(timeout=2.0)


WIN_PROGIDS = [
    "InDesign.Application.2020",
    "InDesign.Application.CC.2020",
    "InDesign.Application.2019",
    "InDesign.Application.CC.2019",
    "InDesign.Application",
]
MAC_APP_NAME = "Adobe InDesign 2020"

BODY_PT = 11
BODY_LEADING = 14
HEADING_BASE_PT = 18
HEADING_STEP_PT = 2
HEADING_MIN_PT = 8
HEADING_EXTRA_LEAD = 3
SPACE_BEFORE_HEAD = 8
SPACE_AFTER_HEAD = 6

FN_MARK_PT = max(7, BODY_PT - 2)

FN_FALLBACK_PT = max(8, BODY_PT - 2)
FN_FALLBACK_LEAD = FN_FALLBACK_PT + 2

PIPELINE_LOGGER: Optional[PipelineLogger] = None

def _user_log(message: str):
    if PIPELINE_LOGGER:
        PIPELINE_LOGGER.user(message)
    else:
        print(message)

def _debug_log(message: str):
    if PIPELINE_LOGGER:
        PIPELINE_LOGGER.debug(message)

def _log_snippet(text: str, limit: int = 120) -> str:
    if not text:
        return ""
    stripped = text.strip()
    if len(stripped) > limit:
        return stripped[:limit] + "..."
    return stripped


def _make_chunk_context(kind: str, seq: int, para_idx: int, style: str, text: str):
    preview = _log_snippet(text or "", limit=120).replace('"', "'")
    return {
        "id": f"{kind}-{seq:03d}",
        "paraIndex": para_idx,
        "style": style,
        "preview": preview,
    }


def _ctx_label(ctx: Optional[Dict[str, str]]) -> str:
    if not ctx:
        return ""
    parts = [ctx.get("id", "ctx")]
    if ctx.get("paraIndex"):
        parts.append(f"para={ctx['paraIndex']}")
    if ctx.get("style"):
        parts.append(f"style={ctx['style']}")
    preview = ctx.get("preview")
    if preview:
        parts.append(f"text={preview}")
    return "[" + " ".join(parts) + "]"


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
                        note_text = _collect_all_text(ch).strip().replace("]]", "】】")
                        foot_map[str(fid)] = note_text
                        _debug_log(f"[NOTE-FOOT] id={fid} len={len(note_text)} snippet={_log_snippet(note_text)}")
            continue
        if tag == "endnotes":
            for ch in list(n):
                if _strip_ns(ch.tag) == "endnote":
                    eid = ch.attrib.get("id") or ch.attrib.get("rid") or ch.attrib.get("ref")
                    if eid:
                        note_text = _collect_all_text(ch).strip().replace("]]", "】】")
                        end_map[str(eid)] = note_text
                        _debug_log(f"[NOTE-END] id={eid} len={len(note_text)} snippet={_log_snippet(note_text)}")
            continue
        stack.extend(list(n))
    _debug_log(f"[NOTES] indexed footnotes={len(foot_map)} endnotes={len(end_map)}")
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
            _debug_log(
                f"[FNREF] id={rid} has_note={bool(note)} noteSnippet={_log_snippet(note)} tailSnippet={_log_snippet(c.tail)}"
            )
            if c.tail: parts.append(c.tail)
            continue
        if tag == "enref":
            rid = c.attrib.get("id") or c.attrib.get("rid") or c.attrib.get("ref")
            parts.append(f"[[FNI:{str(rid)}]]")
            note = end_map.get(str(rid), "")
            parts.append(f"[[EN:{note}]]" if note else "[*]")
            _debug_log(
                f"[ENREF] id={rid} has_note={bool(note)} noteSnippet={_log_snippet(note)} tailSnippet={_log_snippet(c.tail)}"
            )
            if c.tail: parts.append(c.tail)
            continue

        if tag in ("img", "image", "graphic", "figureimage", "inlinegraphic"):
            src = c.attrib.get("src") or c.attrib.get("href") or c.attrib.get("xlink:href") or ""
            w = c.attrib.get("w") or c.attrib.get("width") or ""
            h = c.attrib.get("h") or c.attrib.get("height") or ""
            align = c.attrib.get("align") or c.attrib.get("placement") or ""
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
    _debug_log(f"[XML] parsing paragraphs from {xml_path}")
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
    _debug_log(f"[XML] extracted paragraphs={len(out)} from {xml_path}")
    return out


def escape_js(s: str) -> str:
    s = s.replace("\\", "\\\\").replace('"', '\\"')
    s = s.replace("\\r\\n", " ").replace("\\r", " ").replace("\\n", " ")
    s = s.replace("\r\n", " ").replace("\r", " ").replace("\n", " ")
    s = re.sub(r"\s+", " ", s)
    return s


# External JSX template is required; inline template removed.

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
    return ""



def _load_jsx_template():
    """Load external JSX template; inline template has been removed."""
    tpl_path = os.environ.get("JSX_TEMPLATE_PATH", JSX_TEMPLATE_PATH)
    tpl_abs = os.path.abspath(tpl_path) if tpl_path else None
    if not tpl_abs:
        raise FileNotFoundError("JSX template path is not set. Set JSX_TEMPLATE_PATH or run --dump-jsx-template to see the default path.")
    try:
        with open(tpl_abs, "r", encoding="utf-8") as fh:
            base_text = fh.read()
    except FileNotFoundError:
        raise FileNotFoundError(f"JSX template not found: {tpl_abs}")
    except Exception as exc:
        raise RuntimeError(f"Failed to read JSX template: {tpl_abs} err={exc}")

    frag_dir = os.environ.get("JSX_FRAGMENT_DIR", JSX_FRAGMENT_DIR)
    frag_dir = os.path.abspath(frag_dir)
    fragments = {}
    for key, fname in JSX_FRAGMENTS.items():
        frag_path = os.path.join(frag_dir, fname)
        try:
            with open(frag_path, "r", encoding="utf-8") as fh:
                fragments[key] = fh.read()
        except Exception as exc:
            raise FileNotFoundError(f"Missing JSX fragment {key}: {frag_path} ({exc})")

    composed = base_text
    for key, content in fragments.items():
        composed = composed.replace(f'{{{{{key}}}}}', content)
    return composed, tpl_abs

def _dump_jsx_template(path: str):
    tpl_abs = os.path.abspath(path)
    if os.path.exists(tpl_abs):
        print("[OK] JSX template already exists:", tpl_abs)
        return True
    print("[ERR] JSX template missing and inline template has been removed. Please place templates/indesign_autoflow_map_levels.tpl.jsx at:", tpl_abs)
    return False



IMG_PLACEHOLDER_FULL_RE = re.compile(r'^\s*(\[\[IMG\s+.+?\]\])\s*$', re.I)
IMG_PLACEHOLDER_ANY_RE = re.compile(r'\[\[IMG\s+(.+?)\]\]', re.I)
IMG_KV_PATTERN = r'(\w+)=["\'\u201c\u201d]([^"\'\u201c\u201d]*)["\'\u201c\u201d]'
FRAME_OPEN_RE = re.compile(r'\[\[FRAME\s+([^\]]+)\]\]', re.I)
FRAME_CLOSE_TOKEN = "[[/FRAME]]"


def _js_escape_simple(val: str) -> str:
    return (
        (val or "")
        .replace("\\", "\\\\")
        .replace('"', '\\"')
        .replace("\r\n", "\\n")
        .replace("\r", "\\n")
        .replace("\n", "\\n")
    )


@dataclass
class ImageSpec:
    attrs: Dict[str, str] = field(default_factory=dict)
    force_block: bool = False
    log_context: Optional[Dict[str, str]] = None

    @classmethod
    def from_mapping(cls, mapping: Dict[str, str], *, force_block: bool = False):
        clean = {k: (v or "") for k, v in mapping.items()}
        return cls(attrs=clean, force_block=force_block)

    def get(self, key: str, default: str = "") -> str:
        value = self.attrs.get(key, "")
        return value if value not in (None, "") else default

    def to_js_block(self) -> str:
        src_for_log = self.get("src")
        inline_for_log = self.get("inline")
        spec_js = self._build_spec_js_literal()
        return f'''(function(){{
              try {{
                var spec={spec_js};
                var __imgCtx = spec.logContext || null;
                var __imgTag = "[IMG]";
                if (__imgCtx && __imgCtx.id) __imgTag = "[IMG][" + __imgCtx.id + "]";
                var __imgWarnTag = "[ERROR]";
                if (__imgCtx && __imgCtx.id) __imgWarnTag = "[ERROR][IMG " + __imgCtx.id + "]";
                if (__imgCtx){{
                  var __imgPrev = __imgCtx.preview ? String(__imgCtx.preview) : "";
                  if (__imgPrev.length > 80) __imgPrev = __imgPrev.substring(0,80) + "...";
                  var __imgSummary = ' para=' + (__imgCtx.paraIndex||"?") + ' style=' + (__imgCtx.style||"");
                  if (__imgPrev) __imgSummary += ' text="' + __imgPrev + '"';
                  log(__imgTag + " ctx" + __imgSummary);
                }}
                log(__imgTag + " pyMeta src={src_for_log} inline={inline_for_log}");
                // 0) 环境检查
                log("[DBG] typeof __imgAddFloatingImage=" + (typeof __imgAddFloatingImage)
                    + " typeof __imgAddImageAtV2=" + (typeof __imgAddImageAtV2)
                    + " typeof __imgNormPath=" + (typeof __imgNormPath));
                log("[DBG] tf=" + (tf&&tf.isValid) + " story=" + (story&&story.isValid) + " page=" + (page&&page.isValid));

                // 1) 排版溢出
                try{{ if(typeof flushOverflow==="function"){{ var _rs=flushOverflow(story,page,tf);
                  if(_rs&&_rs.frame&&_rs.page){{ page=_rs.page; tf=_rs.frame; story=tf.parentStory; curTextFrame=tf; }} }} }}catch(_){{
                }}

                // 2) 锚点
                var ip=(tf&&tf.isValid)?_safeIP(tf):story.insertionPoints[-1];
                // 3) 路径
                var f=__imgNormPath("{_js_escape_simple(src_for_log)}");
                log("[DBG] __imgNormPath ok=" + (!!f) + " exists=" + (f&&f.exists ? "Y":"N") + " fsName=" + (f?f.fsName:"NA"));

                if(f&&f.exists){{
                  var inl=_trim(spec.inline); // 兼容 InDesign 2020
                  log(__imgTag + " dispatch src="+spec.src+" inline="+inl+" posH="+(spec.posH||"")+" posV="+(spec.posV||""));

                  if(inl==="0"||/^false$/i.test(inl)){{
                  log("[DBG] dispatch -> __imgAddFloatingImage");
                    var rect=__imgAddFloatingImage(tf,story,page,spec);
                    if(rect&&rect.isValid) log(__imgTag + " ok (float): " + spec.src);
                    try{{
                      if (__FLOAT_CTX && __FLOAT_CTX.lastTf && __FLOAT_CTX.lastTf.isValid){{
                        tf = __FLOAT_CTX.lastTf;
                        story = tf.parentStory;
                        if(__FLOAT_CTX.lastPage && __FLOAT_CTX.lastPage.isValid){{
                          page = __FLOAT_CTX.lastPage;
                        }} else {{
                          page = tf.parentPage;
                        }}
                        try{{
                          if(typeof _safeIP==="function"){{
                            ip = _safeIP(tf);
                          }}
                        }}catch(_){{
                        }}
                        if((!ip || !ip.isValid) && tf && tf.isValid && tf.insertionPoints && tf.insertionPoints.length){{
                          ip = tf.insertionPoints[-1];
                        }}
                      }}
                    }}catch(_){{
                    }}
                  }} else {{
                  log("[DBG] dispatch -> __imgAddImageAtV2");
                    var rect=__imgAddImageAtV2(ip,spec);
                    if(rect&&rect.isValid) log(__imgTag + " ok (inline): " + spec.src);
                  }}
                }} else {{
                  log(__imgWarnTag + " missing: {src_for_log}");
                }}
                try{{
                  var __imgDetail = (__imgCtx && __imgCtx.id) ? ("id=" + __imgCtx.id) : ("src=" + (spec && spec.src ? spec.src : ""));
                  __progressBump("IMG", __imgDetail);
                }}catch(_){{
                }}
              }} catch(e) {{
                log(__imgWarnTag + " exception " + e);
              }}
            }})();'''

    def _build_spec_js_literal(self) -> str:
        align_val = self.attrs.get("align")
        if align_val is None:
            align_val = "center"

        ordered_keys = [
            ("src", self.get("src")),
            ("w", self.get("w")),
            ("h", self.get("h")),
            ("pxw", self.get("pxw")),
            ("pxh", self.get("pxh")),
            ("align", align_val),
            ("inline", self.get("inline")),
            ("wrap", self.get("wrap")),
            ("posH", self.get("posH")),
            ("posHref", self.get("posHref")),
            ("posV", self.get("posV")),
            ("posVref", self.get("posVref")),
            ("rotation", self.get("rotation")),
            ("flipH", self.get("flipH")),
            ("flipV", self.get("flipV")),
            ("offX", self.get("offX")),
            ("offY", self.get("offY")),
            ("distT", self.get("distT")),
            ("distB", self.get("distB")),
            ("distL", self.get("distL")),
            ("distR", self.get("distR")),
            ("cropT", self.get("cropT")),
            ("cropB", self.get("cropB")),
            ("cropL", self.get("cropL")),
            ("cropR", self.get("cropR")),
            ("spaceBefore", self.get("spaceBefore", "6")),
            ("spaceAfter", self.get("spaceAfter", "6")),
            ("caption", self.get("caption")),
            ("docPrId", self.get("docPrId")),
            ("docPrName", self.get("docPrName")),
            ("anchorId", self.get("anchorId")),
            ("anchorEditId", self.get("anchorEditId")),
            ("wordPageWidth", self.get("wordPageWidth")),
            ("wordPageHeight", self.get("wordPageHeight")),
            ("wordPageSeq", self.get("wordPageSeq")),
        ]
        parts = [f'{k}:"{_js_escape_simple(v)}"' for k, v in ordered_keys]
        parts.append(f"forceBlock:{str(self.force_block).lower()}")
        if self.log_context:
            parts.append(f"logContext:{json.dumps(self.log_context, ensure_ascii=False)}")
        return "{%s}" % ",".join(parts)


@dataclass
class FrameSpec:
    attrs: Dict[str, str] = field(default_factory=dict)
    text: str = ""

    @classmethod
    def from_mapping(cls, mapping: Dict[str, str], text: str):
        clean = {k: (v or "") for k, v in mapping.items()}
        return cls(attrs=clean, text=text or "")

    def _build_spec_js_literal(self) -> str:
        keys = [
            "id",
            "wrap",
            "wrapSide",
            "wrapText",
            "posH",
            "posHref",
            "posV",
            "posVref",
            "offX",
            "offY",
            "w",
            "h",
            "distT",
            "distB",
            "distL",
            "distR",
            "relativeHeight",
            "behindDoc",
            "allowOverlap",
            "layoutInCell",
            "hidden",
            "locked",
            "simplePosX",
            "simplePosY",
            "effectL",
            "effectT",
            "effectR",
            "effectB",
            "sizeRelH",
            "sizeRelHref",
            "sizeRelV",
            "sizeRelVref",
            "docPrId",
            "docPrName",
            "anchorId",
            "anchorEditId",
            "bodyInsetL",
            "bodyInsetT",
            "bodyInsetR",
            "bodyInsetB",
            "bodyWrap",
            "bodyRtlCol",
            "pageHint",
            "wordPageWidth",
            "wordPageHeight",
            "wordPageSeq",
        ]
        parts = [f'{k}:"{_js_escape_simple(self.attrs.get(k, ""))}"' for k in keys]
        parts.append(f'text:"{_js_escape_simple(self.text)}"')
        return "{%s}" % ",".join(parts)

    def to_js_block(self) -> str:
        spec_js = self._build_spec_js_literal()
        frame_id = self.attrs.get("id", "")
        return f'''(function(){{
              log("[PY][frame] id={frame_id} len={len(self.text)}");
              try {{
                var spec={spec_js};
                if (typeof __imgAddFloatingFrame === "function") {{
                  __imgAddFloatingFrame(tf, story, page, spec);
                }} else {{
                  log("[FRAME][WARN] addFloatingFrame missing; fallback insert text only (typeof=" + (typeof __imgAddFloatingFrame) + ")");
                  try {{
                    var __ip = (typeof _safeIP==="function") ? _safeIP(tf) : null;
                    if (!__ip || !__ip.isValid) {{
                      if (tf && tf.isValid && tf.insertionPoints && tf.insertionPoints.length) __ip = tf.insertionPoints[-1];
                      else if (story && story.isValid) __ip = story.insertionPoints[-1];
                    }}
                    if (__ip && __ip.isValid) {{
                      var txt = spec.text || "";
                      if (typeof smartWrapStr === "function") txt = smartWrapStr(txt);
                      __ip.contents = txt + "\\r";
                    }}
                  }} catch(__fb) {{
                    log("[FRAME][WARN] fallback insert failed: " + __fb);
                  }}
                }}
              }} catch(e) {{
                log("[FRAME][EXC] " + e);
              }}
            }})();'''


def _prepare_paragraphs_for_jsx(paragraphs, img_pattern):
    """Normalize paragraphs list: split texts around image markers."""
    expanded = []
    for idx, (style, text) in enumerate(paragraphs, 1):
        chunks = _split_media_chunks(style, text)
        # group consecutive ImageSpec chunks for multi-image placement
        merged = []
        img_group = []
        for sty, chunk in chunks:
            if isinstance(chunk, ImageSpec):
                img_group.append(chunk)
                continue
            if img_group:
                if len(img_group) == 1:
                    merged.append((sty, img_group[0]))
                else:
                    _debug_log(f"[PARA-GROUP] idx={idx} style={style} imgs={len(img_group)}")
                    merged.append((sty, img_group.copy()))
                img_group = []
            merged.append((sty, chunk))
        if img_group:
            if len(img_group) == 1:
                merged.append((style, img_group[0]))
            else:
                _debug_log(f"[PARA-GROUP] idx={idx} style={style} imgs={len(img_group)}")
                merged.append((style, img_group.copy()))
        if PIPELINE_LOGGER:
            kinds = [_classify_chunk_value(chunk) for _, chunk in merged]
            _debug_log(f"[PARA-SPLIT idx={idx} style={style}] origLen={len(text or '')} chunks={len(merged)} kinds={kinds}")
        expanded.extend(merged)
    return expanded or paragraphs


def _preflight_snippet(text: str, limit: int = 120) -> str:
    snippet = (text or "").strip().replace("\r", " ").replace("\n", " ")
    if len(snippet) > limit:
        snippet = snippet[:limit] + "..."
    return snippet


def _report_preflight_issue(style: str, text: str, reason: str):
    snippet = _preflight_snippet(text)
    label = style or "Body"
    _user_log(f"[ERROR][PRECHECK] style={label} reason={reason}; snippet={snippet}")
    _debug_log(f"[PRECHECK] style={label} reason={reason} text={text}")


def _preflight_reason(style: str, text: str, max_chars: int = 600000) -> Optional[str]:
    """Return reason string if the paragraph should be skipped before handing to JSX."""
    if not text:
        return None
    length = len(text)
    if length > max_chars:
        reason = f"text too long ({length} chars > {max_chars})"
        _report_preflight_issue(style, text, reason)
        return reason
    opens = text.count("[[")
    closes = text.count("]]")
    if opens != closes:
        reason = f"unbalanced markers ([[ count {opens} vs ]] count {closes})"
        _report_preflight_issue(style, text, reason)
        return reason
    return None


def _split_media_chunks(style, text):
    if not text:
        return [(style, text)]

    parts = []
    stripped_all = text.strip()
    idx = 0
    length = len(text)
    token_pattern = re.compile(r'\[\[(IMG|FRAME|TABLE)\b', re.I)

    def _extract_table_block(start_idx: int):
        json_start = text.find("{", start_idx)
        if json_start == -1:
            return None, start_idx
        brace = 0
        in_string = False
        escape = False
        pos = json_start
        json_end = None
        while pos < length:
            ch = text[pos]
            if escape:
                escape = False
            elif ch == "\\":
                escape = True
            elif ch == '"':
                in_string = not in_string
            elif not in_string:
                if ch == "{":
                    brace += 1
                elif ch == "}":
                    brace -= 1
                    if brace == 0:
                        json_end = pos + 1
                        break
            pos += 1
        if json_end is None:
            return None, start_idx
        close_idx = text.find("]]", json_end)
        if close_idx == -1:
            return None, start_idx
        return text[start_idx:close_idx + 2], close_idx + 2

    search_pos = 0
    while True:
        marker = token_pattern.search(text, search_pos)
        if not marker:
            break
        marker_start = marker.start()
        if marker_start > idx:
            parts.append((style, text[idx:marker_start]))
        token = marker.group(1).lower()
        if token == "img":
            img_match = IMG_PLACEHOLDER_ANY_RE.match(text, marker_start)
            if not img_match:
                search_pos = marker_start + 2
                continue
            attr_section = img_match.group(1)
            only_img = stripped_all == img_match.group(0).strip()
            spec = _image_spec_from_attrs(attr_section, force_block=only_img)
            parts.append((style, spec))
            idx = img_match.end()
            search_pos = idx
            continue
        if token == "frame":
            open_match = FRAME_OPEN_RE.match(text, marker_start)
            if not open_match:
                search_pos = marker_start + 2
                continue
            close_idx = text.find(FRAME_CLOSE_TOKEN, open_match.end())
            if close_idx == -1:
                search_pos = marker_start + 2
                continue
            inner_text = text[open_match.end():close_idx]
            spec = _frame_spec_from_attrs(open_match.group(1), inner_text)
            parts.append((style, spec))
            idx = close_idx + len(FRAME_CLOSE_TOKEN)
            search_pos = idx
            continue
        if token == "table":
            table_chunk, next_idx = _extract_table_block(marker_start)
            if not table_chunk:
                search_pos = marker_start + 2
                continue
            parts.append((style, table_chunk))
            idx = next_idx
            search_pos = next_idx
            continue
        search_pos = marker_start + 2

    if idx < length:
        parts.append((style, text[idx:]))

    return [(style, chunk) for style, chunk in parts if chunk not in ("", None)]


def _classify_chunk_value(chunk):
    if isinstance(chunk, str):
        trimmed = chunk.strip()
        if not trimmed:
            return "text-empty"
        upper = trimmed.upper()
        if upper.startswith("[[IMG"):
            return "IMG_MARKER"
        if upper.startswith("[[TABLE"):
            return "TABLE_MARKER"
        if upper.startswith("[[FRAME"):
            return "FRAME_MARKER"
        return f"text(len={len(trimmed)})"
    name = getattr(chunk, "__class__", type(chunk)).__name__
    if isinstance(chunk, dict):
        if "rows" in chunk and "cols" in chunk:
            return "TABLE_DICT"
        return "dict"
    return name


def _normalize_style_name(style, levels_used):
    sty = style
    lower = sty.lower()
    if lower.startswith("level"):
        try:
            n = int(sty[5:])
            levels_used.add(n)
            sty = f"Level{n}"
        except Exception:
            pass
    elif lower == "body":
        sty = "Body"
    return sty


def _match_img_marker(text):
    attr_match = IMG_PLACEHOLDER_ANY_RE.search(text)
    if not attr_match:
        return None, False
    only_img = bool(IMG_PLACEHOLDER_FULL_RE.match(text))
    return attr_match.group(1), only_img


def _image_spec_from_attrs(attr_text, force_block=False):
    kv = dict(re.findall(IMG_KV_PATTERN, attr_text))
    inline_flag = (kv.get("inline", "") or "").strip().lower()
    pos_href = (kv.get("posHref", "") or kv.get("posH", "") or "").strip().lower()
    pos_vref = (kv.get("posVref", "") or "").strip().lower()
    pos_v = (kv.get("posV", "") or "").strip().lower()
    page_refs = {"page", "pagearea", "pageedge", "margin", "spread"}
    auto_force = force_block
    if not auto_force:
        if inline_flag in ("0", "false", "off", ""):
            if pos_href in page_refs and (pos_vref in page_refs or pos_v in page_refs):
                auto_force = True
    return ImageSpec.from_mapping(kv, force_block=auto_force)


def _frame_spec_from_attrs(attr_text, inner_text):
    kv = dict(re.findall(IMG_KV_PATTERN, attr_text))
    return FrameSpec.from_mapping(kv, text=inner_text.strip())


def _handle_table_marker(text, add_lines, ctx=None):
    m_tbl = re.match(r'^\s*\[\[TABLE\s+(\{[\s\S]*\})\s*\]\]\s*$', text)
    if not m_tbl:
        return False
    payload = m_tbl.group(1)
    parse_source = "json"
    try:
        obj = json.loads(payload)
    except Exception as exc:
        parse_source = "eval"
        _debug_log(f"[TABLE] json decode failed; fallback eval err={exc}")
        obj = eval("(" + payload + ")")
    rows = int(obj.get("rows", 0))
    cols = int(obj.get("cols", 0))
    data = obj.get("data") or []
    ctx_label = _ctx_label(ctx)
    _debug_log(
        f"[TABLE]{ctx_label} marker rows={rows} cols={cols} dataRows={len(data)} source={parse_source}"
    )
    if ctx:
        obj["logContext"] = {
            "id": ctx.get("id"),
            "paraIndex": ctx.get("paraIndex"),
            "style": ctx.get("style"),
            "preview": ctx.get("preview"),
        }
    # rows/cols/data kept here for debugging
    add_lines.append('__tblAddTableHiFi(%s);\ntry{if(__DEFAULT_LAYOUT && __DEFAULT_LAYOUT.pageOrientation=="landscape"){log("[TABLE][restore] skip: default landscape (py-gen)");}else{__tableRestoreLayout(); __ensureLayoutDefault();}}catch(__tblRest){}' % (json.dumps(obj, ensure_ascii=False)))
    return True


def _handle_img_marker(text, add_lines, ctx=None):
    # support multiple IMG markers in one chunk
    all_matches = re.findall(IMG_PLACEHOLDER_ANY_RE, text)
    if not all_matches:
        return False
    if len(all_matches) > 1:
        specs = []
        for m in all_matches:
            spec = _image_spec_from_attrs(m, force_block=False)
            spec.log_context = {
                "id": ctx.get("id") if ctx else None,
                "paraIndex": ctx.get("paraIndex") if ctx else None,
                "style": ctx.get("style") if ctx else None,
                "preview": ctx.get("preview") if ctx else None,
            } if ctx else None
            specs.append(spec)
        ctx_label = _ctx_label(ctx)
        _debug_log(f"[IMG-GROUP]{ctx_label} count={len(specs)}")
        specs_js = "[" + ",".join([s._build_spec_js_literal() for s in specs]) + "]"
        add_lines.append("__ensureLayoutDefault();")
        add_lines.append("(function(){ try{ __imgPlaceImageGroup(tf, story, page, %s); }catch(__e){ try{ log('[IMG-GROUP][ERR] '+__e); }catch(_){ } } })();" % specs_js)
        return True

    match, only_img = _match_img_marker(text)
    if not match:
        return False
    spec = _image_spec_from_attrs(match, force_block=only_img)
    spec.log_context = {
        "id": ctx.get("id") if ctx else None,
        "paraIndex": ctx.get("paraIndex") if ctx else None,
        "style": ctx.get("style") if ctx else None,
        "preview": ctx.get("preview") if ctx else None,
    } if ctx else None
    ctx_label = _ctx_label(ctx)
    _debug_log(
        f"[IMG]{ctx_label} marker src={spec.get('src')} inline={spec.get('inline')} force_block={spec.force_block}"
    )
    add_lines.append("__ensureLayoutDefault();")
    add_lines.append(spec.to_js_block())
    return True


def _handle_html_table(text, add_lines, ctx=None):
    if not re.match(r'^\s*<table\b[\s\S]*</table>\s*$', text, flags=re.I):
        return False
    try:
        root = ET.fromstring(text)
    except Exception:
        return False
    rows_data = []
    for tr in root.findall('.//tr'):
        row = []
        for td in tr.findall('.//td'):
            parts = []
            if td.text and td.text.strip():
                parts.append(td.text.strip())
            for ch in list(td):
                tag = _strip_ns(ch.tag)
                if tag == "p":
                    parts.append(''.join(ch.itertext()).strip())
                elif tag == "img":
                    s = ch.get("src", "") or ""
                    w = ch.get("w", "") or ""
                    h = ch.get("h", "") or ""
                    parts.append('[[IMG src="%s" w="%s" h="%s"]]' % (s, w, h))
                if ch.tail and ch.tail.strip():
                    parts.append(ch.tail.strip())
            row.append("\n".join([x for x in parts if x]))
        rows_data.append(row)
    cols = max([len(r) for r in rows_data]) if rows_data else 0
    obj = {"rows": len(rows_data), "cols": cols, "data": rows_data}
    ctx_label = _ctx_label(ctx)
    _debug_log(f"[HTML-TABLE]{ctx_label} rows={obj['rows']} cols={cols} rowsWithCells={len(rows_data)}")
    if ctx:
        obj["logContext"] = {
            "id": ctx.get("id"),
            "paraIndex": ctx.get("paraIndex"),
            "style": ctx.get("style"),
            "preview": ctx.get("preview"),
        }
    add_lines.append('__tblAddTableHiFi(%s);\ntry{if(__DEFAULT_LAYOUT && __DEFAULT_LAYOUT.pageOrientation=="landscape"){log("[TABLE][restore] skip: default landscape (py-gen)");}else{__tableRestoreLayout(); __ensureLayoutDefault();}}catch(__tblRest){}' % (json.dumps(obj, ensure_ascii=False)))
    return True


def _build_html_image_spec(text):
    if not re.match(r'^\s*<img\b[^>]*>\s*$', text, flags=re.I):
        return None
    try:
        root = ET.fromstring(text)
    except Exception:
        return None

    src = (
        root.get("src", "")
        or root.get("href", "")
        or root.get("{http://www.w3.org/1999/xlink}href", "")
    )
    attrs = {
        "src": src,
        "w": root.get("w", "") or root.get("width", "") or "",
        "h": root.get("h", "") or root.get("height", "") or "",
        "align": root.get("align", "center"),
        "inline": root.get("inline", "") or "",
        "wrap": root.get("wrap", "") or "",
        "posH": root.get("posH", "") or "",
        "posV": root.get("posV", "") or "",
        "posHref": root.get("posHref", "") or "",
        "posVref": root.get("posVref", "") or "",
        "offX": root.get("offX", "") or "",
        "offY": root.get("offY", "") or "",
        "distT": root.get("distT", "") or "",
        "distB": root.get("distB", "") or "",
        "distL": root.get("distL", "") or "",
        "distR": root.get("distR", "") or "",
        "spaceBefore": root.get("spaceBefore", "6"),
        "spaceAfter": root.get("spaceAfter", "6"),
        "caption": root.get("caption", "") or "",
    }
    return ImageSpec.from_mapping(attrs)

def _handle_html_image(text, add_lines, ctx=None):
    spec = _build_html_image_spec(text)
    if not spec:
        return False
    spec.log_context = {
        "id": ctx.get("id") if ctx else None,
        "paraIndex": ctx.get("paraIndex") if ctx else None,
        "style": ctx.get("style") if ctx else None,
        "preview": ctx.get("preview") if ctx else None,
    } if ctx else None
    ctx_label = _ctx_label(ctx)
    _debug_log(
        f"[HTML-IMG]{ctx_label} src={spec.get('src')} inline={spec.get('inline')} force_block={spec.force_block}"
    )
    add_lines.append("__ensureLayoutDefault();")
    add_lines.append(spec.to_js_block())
    return True


def _append_default_paragraph(add_lines, sty, esc):
    add_lines.append("__ensureLayoutDefault();")
    add_lines.append(f'addParaWithNotes(story, "{sty}", "{esc}");')


def write_jsx(jsx_path, paragraphs):
    add_lines = []
    levels_used = set()
    table_seq = 0
    image_seq = 0
    para_chunks = 0

    add_lines.append("function onNewLevel1(){ var pkt = startNewChapter(story, page, tf); story=pkt.story; page=pkt.page; tf=pkt.frame; }")
    add_lines.append("firstChapterSeen = false;")

    img_pattern = re.compile(r'\[\[IMG\s+[^\]]+\]\]', re.I)
    _debug_log(f"[WRITE-JSX] totalParas={len(paragraphs)}")
    for idx, (style, text) in enumerate(paragraphs, 1):
        sty = _normalize_style_name(style, levels_used)
        normalized_text = text or ""
        preview = normalized_text[:40].replace("\n", " ").strip()
        _debug_log(f"[WRITE-JSX idx={idx}] inStyle={style} normalized={sty} origLen={len(normalized_text)} preview={preview!r}")
        reason = _preflight_reason(style, normalized_text)
        if reason:
            preview = escape_js(_preflight_snippet(normalized_text))
            js_reason = escape_js(f"preflight: {reason}")
            add_lines.append(f'__logSkipParagraph(__nextParaSeq(), "{sty}", "{js_reason}", "{preview}")')
            print(f"[ERROR] 段落 {idx+1} ({style}) 预检查失败：{reason}，已跳过")
            continue

        expanded = _prepare_paragraphs_for_jsx([(sty, normalized_text)], img_pattern)
        if not expanded:
            expanded = [(sty, normalized_text)]

        level1_pending = (sty == "Level1")
        for sub_style, chunk in expanded:
            chunk_desc = _classify_chunk_value(chunk)
            _debug_log(f"[WRITE-JSX chunk idx={idx}] type={chunk_desc}")
            if level1_pending:
                add_lines.append("if (firstChapterSeen) { var __fl = flushOverflow(story, page, tf); story = __fl.frame.parentStory; page = __fl.page; tf = __fl.frame; onNewLevel1(); } else { firstChapterSeen = true; }")
                level1_pending = False

            if isinstance(chunk, ImageSpec):
                add_lines.append("__ensureLayoutDefault();")
                add_lines.append(chunk.to_js_block())
                continue
            if isinstance(chunk, list) and chunk and all(isinstance(x, ImageSpec) for x in chunk):
                specs_js = "[" + ",".join([x._build_spec_js_literal() for x in chunk]) + "]"
                _debug_log(f"[WRITE-JSX][IMG-GROUP] idx={idx} count={len(chunk)} style={sub_style}")
                add_lines.append("__ensureLayoutDefault();")
                add_lines.append("(function(){try{__imgPlaceImageGroup(tf, story, page, %s);}catch(__e){try{log('[IMG-GROUP][ERR] '+__e);}catch(_ee){}}})();" % specs_js)
                continue
            if isinstance(chunk, FrameSpec):
                add_lines.append("__ensureLayoutDefault();")
                add_lines.append(chunk.to_js_block())
                continue

            text_chunk = chunk or ""
            sty_chunk = _normalize_style_name(sub_style, levels_used)
            esc = escape_js(text_chunk)

            table_ctx = _make_chunk_context("tbl", table_seq + 1, idx, sty_chunk, text_chunk)
            if _handle_table_marker(text_chunk, add_lines, ctx=table_ctx):
                table_seq += 1
                continue
            img_ctx = _make_chunk_context("img", image_seq + 1, idx, sty_chunk, text_chunk)
            if _handle_img_marker(text_chunk, add_lines, ctx=img_ctx):
                image_seq += 1
                continue
            if _handle_html_table(text_chunk, add_lines, ctx=table_ctx):
                table_seq += 1
                continue
            if _handle_html_image(text_chunk, add_lines, ctx=img_ctx):
                image_seq += 1
                continue

            _append_default_paragraph(add_lines, sty_chunk, esc)
            para_chunks += 1

    progress_total = para_chunks + table_seq + image_seq
    if progress_total <= 0:
        progress_total = len(paragraphs)
    _debug_log(
        f"[WRITE-JSX] progress units para={para_chunks} table={table_seq} img={image_seq} total={progress_total}"
    )

    style_lines = build_style_lines(levels_used)

    img_dirs = [
        OUT_DIR,
        os.path.join(OUT_DIR, "assets"),
        os.path.dirname(XML_PATH) or OUT_DIR,
        os.path.join(os.path.dirname(XML_PATH) or OUT_DIR, "assets"),
    ]
    _seen = set();
    _norm = []
    for d in img_dirs:
        if not d: continue
        dd = os.path.abspath(d)
        if dd not in _seen:
            _seen.add(dd);
            _norm.append(dd)

    jsx_config = {
        "styles": {
            "tableBody": TABLE_BODY_PAR_STYLE,
            "tableBodyFallback": TABLE_BODY_PAR_STYLE_FALLBACK,
            "tableBodyBase": TABLE_BODY_PAR_STYLE_BASE,
            "tableBodyAuto": TABLE_BODY_PAR_STYLE_AUTO,
        },
        "flags": {
            "autoExportIdml": AUTO_EXPORT_IDML,
            "logWrite": LOG_WRITE,
            "allowImgExtFallback": True,
            "safePageLimit": 2000,
        },
        "progress": {
            "heartbeatMs": PROGRESS_HEARTBEAT_MS,
        },
        "imgDirs": _norm,
    }

    jsx, tpl_used = _load_jsx_template()
    jsx = jsx.replace("%TEMPLATE_PATH%", TEMPLATE_PATH.replace("\\", "\\\\"))
    jsx = jsx.replace("%OUT_IDML%", IDML_OUT_PATH.replace("\\", "\\\\"))
    jsx = jsx.replace("%AUTO_EXPORT%", "true" if AUTO_EXPORT_IDML else "false")
    jsx = jsx.replace("%BODY_PT%", str(BODY_PT))
    jsx = jsx.replace("%BODY_LEADING%", str(BODY_LEADING))
    jsx = jsx.replace("%FN_MARK_PT%", str(FN_MARK_PT))
    jsx = jsx.replace("%FN_FALLBACK_PT%", str(FN_FALLBACK_PT))
    jsx = jsx.replace("%FN_FALLBACK_LEAD%", str(FN_FALLBACK_LEAD))
    jsx = jsx.replace("%EVENT_LOG_PATH%", LOG_PATH.replace("\\", "/"))
    jsx = jsx.replace("%LOG_WRITE%", "true" if LOG_WRITE else "false")  # ← 新增
    jsx = jsx.replace("%PROGRESS_TOTAL%", str(max(progress_total, 0)))
    jsx = jsx.replace("%PROGRESS_HEARTBEAT%", str(PROGRESS_HEARTBEAT_MS))
    jsx = jsx.replace("%JSX_CONFIG%", json.dumps(jsx_config, ensure_ascii=False))
    jsx = jsx.replace("%TABLE_BODY_STYLE%", json.dumps(TABLE_BODY_PAR_STYLE))
    jsx = jsx.replace("%TABLE_BODY_STYLE_FALLBACK%", json.dumps(TABLE_BODY_PAR_STYLE_FALLBACK))
    jsx = jsx.replace("%TABLE_BODY_STYLE_BASE%", json.dumps(TABLE_BODY_PAR_STYLE_BASE))
    jsx = jsx.replace("%TABLE_BODY_STYLE_AUTO%", json.dumps(TABLE_BODY_PAR_STYLE_AUTO))
    jsx = jsx.replace("__STYLE_LINES__", style_lines)
    jsx = jsx.replace("__ADD_LINES__", "\n    ".join(add_lines))
    jsx = jsx.replace("%IMG_DIRS_JSON%", json.dumps(_norm).replace("\\", "\\\\"))

    leftovers = sorted(set(m.group(0) for m in re.finditer(r"%[A-Z_][A-Z0-9_]*%", jsx)))
    if leftovers:
        raise RuntimeError(f"JSX placeholder not replaced: {leftovers}")

    with open(jsx_path, "w", encoding="utf-8") as f:
        f.write(jsx)
    # print("[OK] JSX 写入:", jsx_path)
    if tpl_used:
        pass
        # print("[INFO] JSX 模板来源:", tpl_used)
    # print(f"[INFO] JSX 事件日志: {LOG_PATH}")
    print("[DEBUG] JSX 是否包含 addImageAtV2：", any("__imgAddImageAtV2(" in ln for ln in add_lines))


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

    monitor = _start_progress_monitor()
    try:
        app.DoScript(jsx_path, 1246973031)  # 1246973031 = JavaScript
        print("[OK] 已执行 JSX")
        return True
    except Exception as e:
        print("[ERR] DoScript 执行失败：", e)
        return False
    finally:
        _stop_progress_monitor(monitor)


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
        monitor = _start_progress_monitor()
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

        finally:
            _stop_progress_monitor(monitor)
    print("[ERR] 无法调用任何已知的 InDesign 应用名。可设置环境变量 MAC_APP_NAME=你的应用名 再试。")
    return False


def _relay_jsx_events(
    logger: PipelineLogger,
    log_path: str,
    warn_missing: bool = True,
    cleanup: bool = False,
):
    stats = {"info": 0, "warn": 0, "error": 0, "debug": 0}
    if logger is None:
        return stats
    if not os.path.exists(log_path):
        if warn_missing:
            logger.warn(f"未找到 JSX 事件日志: {log_path}")
        return stats
    try:
        with open(log_path, 'r', encoding='utf-8', errors='ignore') as fh:
            for raw in fh:
                entry = raw.rstrip('\n')
                if not entry:
                    continue
                parts = entry.split('\t', 2)
                if len(parts) == 3:
                    level, stamp, message = parts
                else:
                    level, stamp, message = "debug", "", entry
                level = (level or "debug").strip().lower()
                stamp = stamp.strip()
                message = message.strip()
                upper_msg = message.upper()
                if level == "debug":
                    if "[WARN" in upper_msg or upper_msg.startswith("WARN "):
                        level = "warn"
                    elif "[ERR" in upper_msg or upper_msg.startswith("ERROR "):
                        level = "error"
                    elif "[INFO" in upper_msg or upper_msg.startswith("INFO "):
                        level = "info"
                formatted = f"{stamp} {message}".strip()
                module = "JSX"
                logger.debug(formatted, module=module)
                if level == "warn":
                    stats["warn"] += 1
                    logger.warn(formatted, module=module)
                elif level == "error":
                    stats["error"] += 1
                    logger.error(formatted, module=module)
                elif level == "info":
                    stats["info"] += 1
                    logger.user(formatted, module=module)
                else:
                    stats["debug"] += 1
    except Exception as exc:
        logger.warn(f"读取 JSX 事件日志失败: {exc}")
    finally:
        if cleanup:
            try:
                os.remove(log_path)
            except OSError:
                pass
    return stats


def main():
    parser = argparse.ArgumentParser(
        description="DOCX -> XML -> JSX -> InDesign 自动排版"
    )
    parser.add_argument(
        "docx",
        nargs="?",
        help="要转换的 DOCX 文件，未提供则默认脚本目录的 1.docx",
    )
    parser.add_argument(
        "--mode",
        choices=("heading", "regex", "hybrid"),
        default="heading",
        help="DOCXOutlineExporter 的解析模式，默认 heading",
    )
    parser.add_argument(
        "--regex-config",
        help="指定 regex_rules.json（不再支持 .py），用于 --mode=regex 或 hybrid 时自定义正则规则",
    )
    parser.add_argument(
        "--regex-max-depth",
        type=int,
        default=None,
        help="正则分级最大层级，0 表示不限制（默认 200）",
    )
    parser.add_argument(
        "--skip-docx",
        action="store_true",
        help="跳过 DOCX->XML，直接使用已有 XML",
    )
    parser.add_argument(
        "--xml-path",
        help="手动指定 XML 输入/输出路径，默认 formatted_output.xml",
    )
    parser.add_argument(
        "--no-run",
        action="store_true",
        help="只生成 XML/JSX，不实际调用 InDesign",
    )
    parser.add_argument(
        "--dump-jsx-template",
        action="store_true",
        help="Export embedded JSX template to templates/indesign_autoflow_map_levels.tpl.jsx and exit",
    )
    parser.add_argument(
        "--log-dir",
        help="Specify log root directory (default: ./logs)",
    )
    parser.add_argument(
        "--debug-log",
        action="store_true",
        help="Enable debug logging",
    )
    args = parser.parse_args()
    if args.dump_jsx_template:
        target_tpl = os.environ.get("JSX_TEMPLATE_PATH", JSX_TEMPLATE_PATH)
        _dump_jsx_template(target_tpl)
        return


    global XML_PATH, LOG_PATH, LOG_WRITE, PIPELINE_LOGGER
    if args.xml_path:
        XML_PATH = os.path.abspath(args.xml_path)
    docx_input = os.path.abspath(args.docx or "1.docx")
    log_source = docx_input if docx_input else XML_PATH
    PIPELINE_LOGGER = PipelineLogger(
        log_source,
        log_root=args.log_dir,
        enable_debug=args.debug_log,
        console_echo=False,
    )
    LOG_PATH = str(PIPELINE_LOGGER.jsx_event_log_path)
    LOG_WRITE = args.debug_log
    PIPELINE_LOGGER.describe_paths()
    print(f"[LOG] 用户日志: {PIPELINE_LOGGER.user_log_path}")
    if args.debug_log:
        print(f"[LOG] 调试日志: {PIPELINE_LOGGER.debug_log_path}")

    if args.skip_docx:
        if not os.path.exists(XML_PATH):
            msg = f"[ERR] --skip-docx 指定但未找到 XML：{XML_PATH}"
            print(msg)
            PIPELINE_LOGGER.error(msg)
            return
        msg = f"[INFO] 跳过 DOCX → XML，直接使用：{XML_PATH}"
        print(msg)
        PIPELINE_LOGGER.user(msg)
    else:
        if not os.path.exists(docx_input):
            msg = f"[ERR] 找不到 DOCX：{docx_input}"
            print(msg)
            PIPELINE_LOGGER.error(msg)
            return
        exporter = DOCXOutlineExporter(
            docx_input,
            mode=args.mode,
            regex_config_path=args.regex_config,
            regex_max_depth=args.regex_max_depth,
        )
        if args.mode in ("regex", "hybrid"):
            rules_path = getattr(exporter, "regex_rules_path", None)
            if rules_path:
                msg = f"[INFO] regex 规则文件: {rules_path}"
            else:
                msg = "[INFO] regex 使用默认规则"
            print(msg)
            PIPELINE_LOGGER.user(msg)
        summary = exporter.process(XML_PATH)
        _debug_log(f"[DOCX] summary raw={summary}")
        report = (
            f"[REPORT] DOCX 解析完毕: paragraphs={summary.get('word_paragraphs')} "
            f"tables={summary.get('word_tables')} headings={summary.get('headings_detected')} "
            f"footnotes={summary.get('footnotes')} endnotes={summary.get('endnotes')}"
        )
        print(report)
        PIPELINE_LOGGER.user(report)

    paragraphs = extract_paragraphs_with_levels(XML_PATH)
    _debug_log(f"[XML] paragraphs_ready={len(paragraphs)} mode={args.mode}")
    para_msg = f"[INFO] 解析到 {len(paragraphs)} 段；示例： {paragraphs[:3]}"
    print(para_msg)
    PIPELINE_LOGGER.user(para_msg)

    write_jsx(JSX_PATH, paragraphs)
    PIPELINE_LOGGER.user(f"[JSX] 已生成 {JSX_PATH}")

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
    PIPELINE_LOGGER.user(f"[OUTPUT] XML: {XML_PATH}")
    PIPELINE_LOGGER.user(f"[OUTPUT] JSX: {JSX_PATH}")
    PIPELINE_LOGGER.user(f"[OUTPUT] IDML: {IDML_OUT_PATH}")

    stats = _relay_jsx_events(
        PIPELINE_LOGGER, LOG_PATH, warn_missing=not args.no_run, cleanup=False
    )
    summary_line = (
        f"[REPORT] JSX 事件统计 info={stats.get('info', 0)} "
        f"warn={stats.get('warn', 0)} error={stats.get('error', 0)} "
        f"debug={stats.get('debug', 0)}"
    )
    print(summary_line)
    PIPELINE_LOGGER.user(summary_line)

    if ran:
        print("InDesign 已执行 JSX。若设置 AUTO_EXPORT_IDML=True，将在脚本目录生成 output.idml。")



if __name__ == "__main__":
    main()



