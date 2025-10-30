# -*- coding: utf-8 -*-
"""
DOCX -> XML exporter (v13, style-switchable build)
Fixed structural/indentation errors and added stable table/image export.

Original features preserved:
- Heading/regex/hybrid modes
- Footnotes/endnotes extraction
- List numbering reconstruction
- Inline/paragraph style switches

Additions/repairs in this fix:
- Proper block traversal (paragraphs + tables) in document order
- Inline image extraction from runs (saved to assets/, inserted as [[IMG ...]] placeholder)
- Table serialization to [[TABLE {...}]] placeholder (cell text uses existing run-to-text logic)
- Fixed multiple indentation breaks and missing imports
"""
# ===================== CONFIG SWITCHES =====================
STYLE_FLAGS = {
    # Inline run styles
    "italic": True,
    "bold": False,
    "underline": False,
    "color": False,
    "superscript": False,
    "subscript": False,
    "tracking": False,      # letter spacing (tracking), in pt
    "font": False,          # font family
    "fontsize": False,      # font size in pt

    # Paragraph styles
    "paragraph": False,     # set to False to disable all paragraph-level attributes
}

# ===================== CODE =====================
import sys, os
import re
import json
import zipfile
import argparse
from typing import Dict, List, Optional, Tuple

from docx import Document
from lxml import etree

import logging
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")
logger = logging.getLogger(__name__)

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NSMAP = {"w": W_NS}

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
WP_NS= "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
NSMAP_ALL = {"w": W_NS, "a": A_NS, "r": R_NS, "wp": WP_NS}
NS = NSMAP_ALL
XP_RUN_BLIPS = etree.XPath(".//w:drawing//a:blip", namespaces=NSMAP_ALL)
XP_RUN_EXTENT = etree.XPath(
    ".//w:drawing//wp:inline/wp:extent | .//w:drawing//wp:anchor/wp:extent",
    namespaces=NSMAP_ALL
)
# 版式与定位
XP_RUN_ANCHOR = etree.XPath(".//w:drawing//wp:anchor", namespaces=NSMAP_ALL)
XP_RUN_INLINE = etree.XPath(".//w:drawing//wp:inline", namespaces=NSMAP_ALL)
XP_WRAP_ANY  = etree.XPath(".//wp:anchor/*[starts-with(local-name(), 'wrap')]", namespaces=NSMAP_ALL)
XP_POS_H     = etree.XPath(".//wp:anchor/wp:positionH", namespaces=NSMAP_ALL)
XP_POS_V     = etree.XPath(".//wp:anchor/wp:positionV", namespaces=NSMAP_ALL)

# Precompiled XPaths
XP_P_OUTLINE = etree.XPath("./w:pPr/w:outlineLvl", namespaces=NSMAP)
XP_RUN_TEXTS = etree.XPath(".//w:t", namespaces=NSMAP)
XP_RUN_TABS  = etree.XPath(".//w:tab", namespaces=NSMAP)
XP_RUN_FOOTREF = etree.XPath(".//w:footnoteReference", namespaces=NSMAP)
XP_RUN_ENDREF  = etree.XPath(".//w:endnoteReference", namespaces=NSMAP)

# --------- Regex patterns (order = priority) ---------
REGEX_ORDER: List[Tuple[str, Optional[str]]] = [
    ("fixed", r'^第[\d一二三四五六七八九十百]+章[ 　\t]*'),
    ("fixed", r'^第[\d一二三四五六七八九十百]+节[ 　\t]*'),
    ("fixed", r'^第[\d一二三四五六七八九十百]+条[ 　\t]*'),
    ("numeric_dotted", None),
    ("fixed", r'^[（(]\s*\d+\s*[)）]'),
    ("fixed", r'^\d+\s*[)）]'),
    ("fixed", r'^[一二三四五六七八九十百]+、\s*'),
]
COMPILED_FIXED: List[Optional[re.Pattern]] = []
NUMERIC_DOTTED = re.compile(r'^(\d+(?:[\.．]\d+)*)[\.．]?\s*')

def _emu_to_pt(val_emu: Optional[str]) -> Optional[float]:
    try:
        return float(val_emu) / 12700.0
    except Exception:
        return None

def _twip_to_pt(val: Optional[str]) -> Optional[float]:
    try:
        return float(val) / 20.0
    except Exception:
        return None

def _compile_patterns():
    global COMPILED_FIXED
    COMPILED_FIXED = []
    for kind, pat in REGEX_ORDER:
        if kind == "fixed" and pat:
            COMPILED_FIXED.append(re.compile(pat))
        else:
            COMPILED_FIXED.append(None)
_compile_patterns()

def _int_to_roman(n: int, upper: bool=True) -> str:
    if n <= 0:
        return str(n)
    vals = [
        (1000,'M'),(900,'CM'),(500,'D'),(400,'CD'),
        (100,'C'),(90,'XC'),(50,'L'),(40,'XL'),
        (10,'X'),(9,'IX'),(5,'V'),(4,'IV'),(1,'I')
    ]
    res = []
    for v,s in vals:
        while n >= v:
            res.append(s); n -= v
    out = ''.join(res)
    return out if upper else out.lower()

def _int_to_alpha(n: int, upper: bool=False) -> str:
    if n <= 0:
        return str(n)
    s = ""
    while n > 0:
        n, r = divmod(n-1, 26)
        s = chr(ord('A' if upper else 'a') + r) + s
    return s

def _strip_pstyle_marker(text: str) -> str:
    return re.sub(r'\s*\[\[PSTYLE\b[^\]]*\]\]\s*$', '', text or '')

def _parse_pstyle_marker(text: str) -> Tuple[str, Dict[str, str]]:
    """Extract [[PSTYLE ...]] from tail and return (text_wo_marker, attrs_dict)."""
    attrs: Dict[str, str] = {}
    m = re.search(r'\[\[PSTYLE\s+([^\]]+)\]\]\s*$', text or '')
    if not m:
        return text, attrs
    attr_s = m.group(1)
    # parse key="value" pairs
    for k, v in re.findall(r'([a-zA-Z\-]+)="([^"]*)"', attr_s):
        attrs[k] = v
    text_wo = text[:m.start()].rstrip()
    return text_wo, attrs

class MyDOCNode(object):
    LEVEL_TAGS = {1: "chapter", 2: "section", 3: "subsection"}

    def __init__(self, name: str, level: int, index: int = 1, parent: Optional["MyDOCNode"] = None,
                 element_type: str = "heading", properties: Optional[dict] = None):
        self.name = name or ""
        self.level = level or 0
        self.index = index or 0
        self.parent = parent
        self.children: List["MyDOCNode"] = []
        self.element_type = element_type
        self.properties = properties or {}
        self.body_paragraphs: List[str] = []  # paragraphs stored as text + optional [[PSTYLE ...]] marker
        self._pending_for_children: List[str] = []

    def add_child(self, child: "MyDOCNode"):
        self.children.append(child)

    @staticmethod
    def container_tag(level: int) -> str:
        return MyDOCNode.LEVEL_TAGS.get(level, f"level{level}")

    @staticmethod
    def heading_tag(level: int) -> str:
        return f"h{level}"

    @staticmethod
    def _escape_xml(text: str) -> str:
        return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

    @staticmethod
    def _escape_attr(val: str) -> str:
        return (val.replace("&", "&amp;")
                   .replace('"', "&quot;")
                   .replace("<", "&lt;")
                   .replace(">", "&gt;"))

    @staticmethod
    def _convert_refs_to_xml(text: str) -> str:
        # helpers
        def repl_fn(m): return f'<fnref id="{m.group(1)}"/>'
        def repl_en(m): return f'<enref id="{m.group(1)}"/>'
        def repl_i(m):  return f'<i>{m.group(1)}</i>'
        def repl_b(m):  return f'<b>{m.group(1)}</b>'
        def repl_u(m):  return f'<u>{m.group(1)}</u>'
        def repl_sup(m): return f'<sup>{m.group(1)}</sup>'
        def repl_sub(m): return f'<sub>{m.group(1)}</sub>'
        def repl_span(m):
            attrs = m.group(1) or ""
            content = m.group(2)
            out_attrs = []
            for key, data_key in (
                ("font", "data-font"),
                ("size", "data-size"),
                ("color", "data-color"),
                ("tracking", "data-tracking"),
            ):
                mm = re.search(rf'{key}="([^"]*)"', attrs)
                if mm:
                    out_attrs.append(f'{data_key}="{MyDOCNode._escape_attr(mm.group(1))}"')
            attr_out = (" " + " ".join(out_attrs)) if out_attrs else ""
            return f'<span{attr_out}>{content}</span>'

        # apply replacements
        text = re.sub(r"\[\[FNREF:(-?\d+)\]\]", repl_fn, text)
        text = re.sub(r"\[\[ENREF:(-?\d+)\]\]", repl_en, text)
        text = re.sub(r"\[\[I\]\](.*?)\[\[/I\]\]", repl_i, text, flags=re.DOTALL)
        text = re.sub(r"\[\[B\]\](.*?)\[\[/B\]\]", repl_b, text, flags=re.DOTALL)
        text = re.sub(r"\[\[U\]\](.*?)\[\[/U\]\]", repl_u, text, flags=re.DOTALL)
        text = re.sub(r"\[\[SUP\]\](.*?)\[\[/SUP\]\]", repl_sup, text, flags=re.DOTALL)
        text = re.sub(r"\[\[SUB\]\](.*?)\[\[/SUB\]\]", repl_sub, text, flags=re.DOTALL)
        text = re.sub(r'\[\[SPAN(.*?)\]\](.*?)\[\[/SPAN\]\]', repl_span, text, flags=re.DOTALL)
        return text

    @staticmethod
    def _escape_attr_dict(d: Dict[str, str]) -> Dict[str, str]:
        return {k: MyDOCNode._escape_attr(v) for k, v in d.items() if v is not None}

    def to_xml_string(self, notes: dict = None, indent: int = 0) -> str:
        pad = "  " * indent
        parts = [f'{pad}<{self.container_tag(self.level)}>']
        heading_text = self._escape_xml(self.name)
        heading_text = self._convert_refs_to_xml(heading_text)
        parts.append(f'{pad}  <{self.heading_tag(self.level)}>{heading_text}</{self.heading_tag(self.level)}>')
        if self.properties:
            parts.append(f'{pad}  <meta>')
            for k, v in self.properties.items():
                v_str = self._escape_xml(str(v))
                parts.append(f'{pad}    <prop name="{k}">{v_str}</prop>')
            parts.append(f'{pad}  </meta>')
        for para in self.body_paragraphs:
            text, pattrs = _parse_pstyle_marker(para) if STYLE_FLAGS.get("paragraph", True) else (para, {})
            escaped = self._escape_xml(text)
            escaped = self._convert_refs_to_xml(escaped)
            # build <p ...attrs>
            attrs_str = ""
            if STYLE_FLAGS.get("paragraph", True) and pattrs:
                attrs_pairs = [f'{k}="{self._escape_attr(v)}"' for k, v in pattrs.items() if v is not None and v != ""]
                if attrs_pairs:
                    attrs_str = " " + " ".join(attrs_pairs)
            parts.append(f'{pad}  <p{attrs_str}>{escaped}</p>')
        for c in self.children:
            parts.append(c.to_xml_string(notes, indent + 1))
        parts.append(f'{pad}</{self.container_tag(self.level)}>')
        return "\n".join(parts)

class Splitter:
    def matches(self, line: str) -> bool:
        raise NotImplementedError

class FixedRegexSplitter(Splitter):
    def __init__(self, pat: re.Pattern):
        self.pat = pat
    def matches(self, line: str) -> bool:
        return bool(self.pat.match(line))

class NumericDepthSplitter(Splitter):
    def __init__(self, depth: int):
        self.depth = depth
    def matches(self, line: str) -> bool:
        m = NUMERIC_DOTTED.match(line)
        if not m:
            return False
        segs = re.split(r'[\.．]', m.group(1))
        return len([s for s in segs if s]) == self.depth

class DOCXOutlineExporter:
    def __init__(self, input_path: str, mode: str = "heading"):
        assert mode in ("heading", "regex", "hybrid"), "mode must be 'heading', 'regex', or 'hybrid'"
        self.mode = mode
        self.input_path = input_path
        self.doc = Document(input_path)
        self.footnotes: Dict[str, str] = {}
        self.endnotes: Dict[str, str] = {}
        self.root = MyDOCNode("root", level=0, index=0, parent=None, element_type="root")

        # Numbering (lists)
        self.num_map_abstract: Dict[int, int] = {}
        self.abstract_lvls: Dict[int, Dict[int, Dict[str, Optional[str]]]] = {}
        self.num_overrides: Dict[int, Dict[int, int]] = {}
        self.num_counters: Dict[int, List[int]] = {}
        self.assets_dir = None  # set in process()

    # ---------- Notes extraction ----------
    @staticmethod
    def _collect_text_from_p(par_el) -> str:
        chunks = []
        for node in par_el.iter():
            if node.tag == f"{{{W_NS}}}t":
                chunks.append(node.text or "")
            elif node.tag == f"{{{W_NS}}}tab":
                chunks.append("\t")
        return "".join(chunks)

    @staticmethod
    def _parse_notes_xml(xml_bytes: bytes) -> Dict[str, str]:
        result = {}
        if not xml_bytes:
            return result
        root = etree.fromstring(xml_bytes)
        for n in root.findall("w:footnote", NSMAP) + root.findall("w:endnote", NSMAP):
            nid = n.get(f"{{{W_NS}}}id")
            if nid is None:
                continue
            try:
                if int(nid) < 1:
                    continue
            except Exception:
                pass
            paras = n.findall(".//w:p", NSMAP)
            texts = [DOCXOutlineExporter._collect_text_from_p(p) for p in paras]
            result[nid] = "\n".join([t for t in texts if t.strip()])
        return result

    def extract_notes(self):
        with zipfile.ZipFile(self.input_path, "r") as z:
            foot_xml = z.read("word/footnotes.xml") if "word/footnotes.xml" in z.namelist() else None
            end_xml = z.read("word/endnotes.xml") if "word/endnotes.xml" in z.namelist() else None
            num_xml = z.read("word/numbering.xml") if "word/numbering.xml" in z.namelist() else None
        if foot_xml:
            self.footnotes = self._parse_notes_xml(foot_xml)
            logger.info(f"Footnotes parsed: {len(self.footnotes)}")
        else:
            logger.info("No footnotes.xml found.")
        if end_xml:
            self.endnotes = self._parse_notes_xml(end_xml)
            logger.info(f"Endnotes parsed: {len(self.endnotes)}")
        else:
            logger.info("No endnotes.xml found.")
        if num_xml:
            self._parse_numbering_xml(num_xml)
            logger.info("numbering.xml parsed.")
        else:
            logger.info("No numbering.xml found (list numbers may not be reconstructed).")

    # ---------- Numbering (lists) ----------
    def _parse_numbering_xml(self, xml_bytes: bytes):
        try:
            root = etree.fromstring(xml_bytes)
        except Exception as e:
            logger.warning(f"Failed to parse numbering.xml: {e}")
            return
        for abs_num in root.findall("w:abstractNum", NSMAP):
            anid_s = abs_num.get(f"{{{W_NS}}}abstractNumId")
            if anid_s is None:
                continue
            try:
                anid = int(anid_s)
            except Exception:
                continue
            lvls: Dict[int, Dict[str, Optional[str]]] = {}
            for lvl in abs_num.findall("w:lvl", NSMAP):
                ilvl_s = lvl.get(f"{{{W_NS}}}ilvl")
                if ilvl_s is None:
                    continue
                try:
                    ilvl = int(ilvl_s)
                except Exception:
                    continue
                start = 1
                start_el = lvl.find("w:start", NSMAP)
                if start_el is not None and start_el.get(f"{{{W_NS}}}val") is not None:
                    try:
                        start = int(start_el.get(f"{{{W_NS}}}val"))
                    except Exception:
                        start = 1
                numFmt = None
                numFmt_el = lvl.find("w:numFmt", NSMAP)
                if numFmt_el is not None:
                    numFmt = numFmt_el.get(f"{{{W_NS}}}val")
                lvlText = None
                lvlText_el = lvl.find("w:lvlText", NSMAP)
                if lvlText_el is not None:
                    lvlText = lvlText_el.get(f"{{{W_NS}}}val")
                lvls[ilvl] = {"start": start, "numFmt": numFmt, "lvlText": lvlText}
            self.abstract_lvls[anid] = lvls
        for num in root.findall("w:num", NSMAP):
            numId_s = num.get(f"{{{W_NS}}}numId")
            if numId_s is None:
                continue
            try:
                numId = int(numId_s)
            except Exception:
                continue
            an_el = num.find("w:abstractNumId", NSMAP)
            if an_el is not None and an_el.get(f"{{{W_NS}}}val") is not None:
                try:
                    anid = int(an_el.get(f"{{{W_NS}}}val"))
                    self.num_map_abstract[numId] = anid
                except Exception:
                    pass
            for ov in root.findall("w:lvlOverride", NSMAP):
                ilvl_el = ov.find("w:ilvl", NSMAP)
                if ilvl_el is None or ilvl_el.get(f"{{{W_NS}}}val") is None:
                    continue
                try:
                    ilvl = int(ilvl_el.get(f"{{{W_NS}}}val"))
                except Exception:
                    continue
                start_el = ov.find(".//w:startOverride", NSMAP)
                if start_el is not None and start_el.get(f"{{{W_NS}}}val") is not None:
                    try:
                        start_val = int(start_el.get(f"{{{W_NS}}}val"))
                    except Exception:
                        start_val = None
                else:
                    start_val = None
                if start_val is not None:
                    self.num_overrides.setdefault(numId, {})[ilvl] = start_val

    def _ensure_counters(self, numId: int):
        if numId not in self.num_counters:
            counters = [0]*9
            anid = self.num_map_abstract.get(numId)
            for il in range(9):
                start = 1
                if numId in self.num_overrides and il in self.num_overrides[numId]:
                    start = self.num_overrides[numId][il]
                elif anid in self.abstract_lvls and il in self.abstract_lvls[anid]:
                    try:
                        start = int(self.abstract_lvls[anid][il].get("start") or 1)
                    except Exception:
                        start = 1
                counters[il] = start - 1
            self.num_counters[numId] = counters

    def _format_level_num(self, anid: Optional[int], ilvl: int, val: int) -> str:
        numFmt = None
        if anid in self.abstract_lvls and ilvl in self.abstract_lvls[anid]:
            numFmt = self.abstract_lvls[anid][ilvl].get("numFmt")
        if numFmt in (None, "decimal"):
            return str(val)
        if numFmt == "lowerLetter":
            return _int_to_alpha(val, upper=False)
        if numFmt == "upperLetter":
            return _int_to_alpha(val, upper=True)
        if numFmt == "lowerRoman":
            return _int_to_roman(val, upper=False)
        if numFmt == "upperRoman":
            return _int_to_roman(val, upper=True)
        if numFmt == "bullet":
            return "•"
        return str(val)

    def _build_label_from_lvlText(self, anid: Optional[int], ilvl: int, counters: List[int]) -> str:
        lvlText = None
        if anid in self.abstract_lvls and ilvl in self.abstract_lvls[anid]:
            lvlText = self.abstract_lvls[anid][ilvl].get("lvlText")
        if not lvlText:
            cur_val = counters[ilvl]
            return self._format_level_num(anid, ilvl, cur_val) + "."
        out = lvlText
        for i in range(1, 10):
            if f"%{i}" in out:
                level_idx = i-1
                val = counters[level_idx]
                out = out.replace(f"%{i}", self._format_level_num(anid, level_idx, val))
        return out

    def list_label_for_paragraph(self, paragraph) -> str:
        p_el = paragraph._element
        numPr = p_el.find("./w:pPr/w:numPr", NSMAP)
        if numPr is None:
            return ""
        numId_el = numPr.find("./w:numId", NSMAP)
        if numId_el is None or numId_el.get(f"{{{W_NS}}}val") is None:
            return ""
        ilvl_el = numPr.find("./w:ilvl", NSMAP)
        try:
            numId = int(numId_el.get(f"{{{W_NS}}}val"))
        except Exception:
            return ""
        ilvl = 0
        if ilvl_el is not None and ilvl_el.get(f"{{{W_NS}}}val") is not None:
            try:
                ilvl = int(ilvl_el.get(f"{{{W_NS}}}val"))
            except Exception:
                ilvl = 0
        self._ensure_counters(numId)
        counters = self.num_counters[numId]
        counters[ilvl] += 1
        anid = self.num_map_abstract.get(numId)
        for j in range(ilvl+1, len(counters)):
            start = 1
            if numId in self.num_overrides and j in self.num_overrides[numId]:
                start = self.num_overrides[numId][j]
            elif anid in self.abstract_lvls and j in self.abstract_lvls[anid]:
                try:
                    start = int(self.abstract_lvls[anid][j].get("start") or 1)
                except Exception:
                    start = 1
            counters[j] = start - 1
        label = self._build_label_from_lvlText(anid, ilvl, counters)
        if label and not re.search(r'\s$', label):
            label = label + " "
        return label

    # ---------- Paragraph style helpers ----------
    @staticmethod
    def _paragraph_style_attrs(p) -> Dict[str, str]:
        if not STYLE_FLAGS.get("paragraph", True):
            return {}
        attrs: Dict[str, str] = {}
        # style name
        try:
            sname = getattr(getattr(p, "style", None), "name", None)
            if sname:
                attrs["style-name"] = str(sname)
        except Exception:
            pass
        p_el = p._element
        ppr = p_el.find("./w:pPr", NSMAP)
        # alignment
        try:
            jc = ppr.find("./w:jc", NSMAP) if ppr is not None else None
            align = jc.get(f"{{{W_NS}}}val") if jc is not None else None
            if align:
                attrs["align"] = align
        except Exception:
            pass
        # indents
        try:
            ind = ppr.find("./w:ind", NSMAP) if ppr is not None else None
            if ind is not None:
                def _pt(attr):
                    v = ind.get(f"{{{W_NS}}}{attr}")
                    if v is None:
                        return None
                    try:
                        return float(v) / 20.0
                    except Exception:
                        return None
                left = _pt("left"); right = _pt("right"); first = _pt("firstLine"); hanging = _pt("hanging")
                if left is not None: attrs["indent-left"] = f"{int(left)}pt" if float(left).is_integer() else f"{left:.1f}pt"
                if right is not None: attrs["indent-right"] = f"{int(right)}pt" if float(right).is_integer() else f"{right:.1f}pt"
                if first is not None: attrs["indent-first"] = f"{int(first)}pt" if float(first).is_integer() else f"{first:.1f}pt"
                if hanging is not None: attrs["indent-hanging"] = f"{int(hanging)}pt" if float(hanging).is_integer() else f"{hanging:.1f}pt"
        except Exception:
            pass
        # spacing
        try:
            sp = ppr.find("./w:spacing", NSMAP) if ppr is not None else None
            if sp is not None:
                def _pt(v):
                    try:
                        vv = float(v) / 20.0
                        return f"{int(vv)}pt" if float(vv).is_integer() else f"{vv:.1f}pt"
                    except Exception:
                        return None
                before = sp.get(f"{{{W_NS}}}before")
                after = sp.get(f"{{{W_NS}}}after")
                if before: attrs["space-before"] = _pt(before) or ""
                if after:  attrs["space-after"]  = _pt(after) or ""
                line = sp.get(f"{{{W_NS}}}line")
                rule = sp.get(f"{{{W_NS}}}lineRule")
                if line:
                    try:
                        line_val = float(line)
                        if rule == "auto" or rule is None:
                            # 240 = single line
                            mult = line_val / 240.0
                            attrs["line-height"] = f"{mult:.2f}x"
                            attrs["line-rule"] = "auto"
                        else:
                            # exact / atLeast: twentieths of a point
                            pt = line_val / 20.0
                            attrs["line-height"] = f"{int(pt)}pt" if float(pt).is_integer() else f"{pt:.1f}pt"
                            attrs["line-rule"] = rule
                    except Exception:
                        pass
        except Exception:
            pass
        return {k:v for k,v in attrs.items() if v}

    # ---------- Text with inline refs + numbering + styles ----------
    def _paragraph_text_with_refs(self, paragraph, include_pstyle: bool = True) -> str:
        chunks = []
        # list numbering
        try:
            label = self.list_label_for_paragraph(paragraph)
        except Exception as e:
            logger.debug(f"list label error: {e}")
            label = ""
        if label:
            chunks.append(label)

        # Traverse runs: inline images and text/notes/tabs
        for run in paragraph.runs:
            r_el = run._element

            # ---------- IMAGES (high-fidelity [[IMG ...]] placeholder) ----------
            try:
                blips = XP_RUN_BLIPS(r_el) or []
            except Exception:
                blips = []
            if blips:
                import os
                os.makedirs(self.assets_dir or ".", exist_ok=True)
                for bl in blips:
                    rid = bl.get(f"{{{R_NS}}}embed")
                    if not rid:
                        continue
                    part = getattr(run, "part", None)
                    try:
                        image_part = part.related_parts.get(rid) if part else None
                    except Exception:
                        image_part = None
                    if image_part is None:
                        continue

                    # export blob -> assets/...
                    ext = os.path.splitext(getattr(image_part, "partname", ""))[1] or ".png"
                    if not hasattr(self, "_img_seq"):
                        self._img_seq = 0
                    self._img_seq += 1
                    fname = f"image_{self._img_seq:04d}{ext}"
                    out_path = os.path.join(self.assets_dir or ".", fname)
                    try:
                        with open(out_path, "wb") as f:
                            f.write(image_part.blob)
                    except Exception:
                        pass

                    # size from wp:extent (EMU -> pt)
                    wpt = hpt = ""
                    try:
                        exts = XP_RUN_EXTENT(r_el) or []
                        if exts:
                            cx = exts[0].get("cx")
                            cy = exts[0].get("cy")
                            EMU_PER_PT = 12700.0
                            if cx:
                                wpt = f"{float(cx) / EMU_PER_PT:.2f}"
                            if cy:
                                hpt = f"{float(cy) / EMU_PER_PT:.2f}"
                    except Exception:
                        pass

                    if wpt: wpt = f"{wpt}pt"
                    if hpt: hpt = f"{hpt}pt"

                    # inline / anchor, wrap, posH/posV, offsets, distances
                    inline_el = XP_RUN_INLINE(r_el) or []
                    anchor_el = XP_RUN_ANCHOR(r_el) or []
                    is_inline = bool(inline_el) and not bool(anchor_el)

                    wrap = ""
                    posH = posV = ""
                    offX = offY = ""
                    distT = distB = distL = distR = ""
                    try:
                        if anchor_el:
                            # wrap: square / tight / through / none...
                            wnodes = XP_WRAP_ANY(r_el) or []
                            if wnodes:
                                wrap = wnodes[0].tag.split('}')[-1]  # e.g. "wrapSquare"

                            # positionH / positionV: align or posOffset
                            phs = XP_POS_H(r_el) or []
                            pvs = XP_POS_V(r_el) or []
                            if phs:
                                al = phs[0].find("wp:align", NSMAP_ALL)
                                of = phs[0].find("wp:posOffset", NSMAP_ALL)
                                if al is not None and al.text:
                                    posH = al.text
                                if of is not None and of.text:
                                    offX = f"{float(of.text) / 12700.0:.2f}pt"
                            if pvs:
                                al = pvs[0].find("wp:align", NSMAP_ALL)
                                of = pvs[0].find("wp:posOffset", NSMAP_ALL)
                                if al is not None and al.text:
                                    posV = al.text
                                if of is not None and of.text:
                                    offY = f"{float(of.text) / 12700.0:.2f}pt"

                            # anchor distances to text/page (if present)
                            a0 = anchor_el[0]
                            for key_xml, var_name in (("distT", "distT"), ("distB", "distB"),
                                                      ("distL", "distL"), ("distR", "distR")):
                                v = a0.get(f"{{{WP_NS}}}{key_xml}")
                                if v is not None:
                                    try:
                                        locals()[var_name] = f"{float(v) / 12700.0:.2f}pt"
                                    except Exception:
                                        pass
                    except Exception:
                        pass

                    # original pixel size (if available)
                    # original pixel size (true pixels, not EMU)
                    pxw = pxh = ""
                    try:
                        im = getattr(image_part, "image", None)
                        if im is not None and getattr(im, "px_width", None) and getattr(im, "px_height", None):
                            pxw = str(int(im.px_width))
                            pxh = str(int(im.px_height))
                        # fallback: if px not available, leave blank (IDML端按 w/h pt 处理即可)
                    except Exception:
                        pass

                    if (not posV) and offY:
                        posV = "paragraph"  # 让 offY 有参照
                    out_path = out_path.replace("\\", "/")  # 统一为正斜杠

                    # 若距离为空就写 0pt，避免下游判空
                    for k in ("distT", "distB", "distL", "distR"):
                        if locals()[k] == "":
                            locals()[k] = "0pt"

                    inline_flag = "1" if is_inline else "0"

                    chunks.append(
                        '[[IMG src="{src}" w="{w}" h="{h}" pxw="{pxw}" pxh="{pxh}" '
                        'inline="{inline_}" wrap="{wrap}" posH="{posH}" posV="{posV}" '
                        'offX="{offX}" offY="{offY}" distT="{distT}" distB="{distB}" '
                        'distL="{distL}" distR="{distR}"]]'
                            .format(
                            src=out_path, w=wpt, h=hpt, pxw=pxw, pxh=pxh, inline_=inline_flag,
                            wrap=wrap, posH=posH, posV=posV, offX=offX, offY=offY,
                            distT=distT, distB=distB, distL=distL, distR=distR
                        )
                    )

            # ---------- text + notes + tabs ----------
            parts_run = []
            for t in XP_RUN_TEXTS(r_el):
                if t.text:
                    parts_run.append(t.text)
            for ref in XP_RUN_FOOTREF(r_el):
                rid = ref.get(f"{{{W_NS}}}id")
                parts_run.append(f"[[FNREF:{rid}]]")
            for ref in XP_RUN_ENDREF(r_el):
                rid = ref.get(f"{{{W_NS}}}id")
                parts_run.append(f"[[ENREF:{rid}]]")
            if XP_RUN_TABS(r_el):
                parts_run.append("\t")
            run_text = "".join(parts_run)
            if run_text:
                                # --- INLINE STYLE WRAP (from STYLE_FLAGS) ---
                            try:
                                # Detect run styles robustly
                                ital = getattr(run, "italic", None)
                                bold = getattr(run, "bold", None)
                                under = getattr(run, "underline", None)
                                fobj = getattr(run, "font", None)
                                if ital is None and fobj is not None: ital = getattr(fobj, "italic", None)
                                if bold is None and fobj is not None: bold = getattr(fobj, "bold", None)
                                if under is None and fobj is not None: under = getattr(fobj, "underline", None)

                                # Fallback to raw rPr
                                try:
                                    rpr = r_el.find(".//w:rPr", NSMAP)
                                except Exception:
                                    rpr = None

                                def rpr_has(tag):
                                    if rpr is None: return False
                                    el = rpr.find(tag, NSMAP)
                                    if el is None: return False
                                    val = el.get(f"{{{W_NS}}}val")
                                    return (val in (None, "", "1", "true", "True")) or (tag=='w:u' and val not in ("none","0","false"))

                                # superscript / subscript
                                supers = getattr(fobj, "superscript", False) if fobj is not None else False
                                sub    = getattr(fobj, "subscript", False)    if fobj is not None else False
                                if not supers and not sub and rpr is not None:
                                    va = rpr.find("w:vertAlign", NSMAP)
                                    if va is not None:
                                        v = va.get(f"{{{W_NS}}}val")
                                        supers = (v == "superscript")
                                        sub    = (v == "subscript")

                                # font family / size / color
                                font_name = None
                                font_size = None
                                color_hex = None
                                tracking  = None

                                try:
                                    if fobj is not None and getattr(fobj, "name", None):
                                        font_name = str(fobj.name)
                                    elif rpr is not None:
                                        rf = rpr.find("w:rFonts", NSMAP)
                                        if rf is not None:
                                            font_name = rf.get(f"{{{W_NS}}}ascii") or rf.get(f"{{{W_NS}}}hAnsi")
                                except Exception: pass

                                try:
                                    if fobj is not None and getattr(fobj, "size", None):
                                        try:
                                            font_size = float(fobj.size.pt)
                                        except Exception:
                                            font_size = None
                                    elif rpr is not None:
                                        sz = rpr.find("w:sz", NSMAP)
                                        if sz is not None:
                                            val = sz.get(f"{{{W_NS}}}val")
                                            if val:
                                                font_size = float(val)/2.0
                                except Exception: pass

                                try:
                                    if rpr is not None:
                                        c = rpr.find("w:color", NSMAP)
                                        if c is not None:
                                            v = c.get(f"{{{W_NS}}}val")
                                            if v and v.lower() not in ("auto",):
                                                if len(v) in (6,3):
                                                    color_hex = "#" + v if not v.startswith("#") else v
                                except Exception: pass

                                # Apply wrappers based on STYLE_FLAGS
                                span_attrs = []
                                if STYLE_FLAGS.get("font", False) and font_name:
                                    span_attrs.append(f'font="{font_name}"')
                                if STYLE_FLAGS.get("fontsize", False) and font_size:
                                    span_attrs.append(f'size="{int(font_size) if float(font_size).is_integer() else font_size}"')
                                if STYLE_FLAGS.get("color", False) and color_hex:
                                    span_attrs.append(f'color="{color_hex}"')
                                if STYLE_FLAGS.get("tracking", False) and tracking:
                                    span_attrs.append(f'tracking="{tracking}"')
                                if span_attrs:
                                    run_text = f"[[SPAN {' '.join(span_attrs)}]]{run_text}[[/SPAN]]"

                                if STYLE_FLAGS.get("superscript", False) and supers:
                                    run_text = f"[[SUP]]{run_text}[[/SUP]]"
                                elif STYLE_FLAGS.get("subscript", False) and sub:
                                    run_text = f"[[SUB]]{run_text}[[/SUB]]"

                                if STYLE_FLAGS.get("underline", False):
                                    is_u = (under is True) or rpr_has("w:u")
                                    if is_u:
                                        run_text = f"[[U]]{run_text}[[/U]]"

                                if STYLE_FLAGS.get("bold", False):
                                    is_b = (bold is True) or rpr_has("w:b")
                                    if is_b:
                                        run_text = f"[[B]]{run_text}[[/B]]"

                                if STYLE_FLAGS.get("italic", True):
                                    is_i = (ital is True) or rpr_has("w:i")
                                    if is_i:
                                        run_text = f"[[I]]{run_text}[[/I]]"
                            except Exception:
                                pass
            chunks.append(run_text)

        out = "".join(chunks).strip()

        # Paragraph style marker (only for body lines, not headings)
        if include_pstyle and STYLE_FLAGS.get("paragraph", True):
            pattrs = MyDOCNode._escape_attr_dict(self._paragraph_style_attrs(paragraph))
            if pattrs:
                kv = " ".join([f'{k}="{v}"' for k, v in pattrs.items()])
                out += f' [[PSTYLE {kv}]]'
        return out

    # ---------- Heading detection (heading mode) ----------
    @staticmethod
    def _style_based_heading_level(style_name: Optional[str]) -> Optional[int]:
        if not style_name:
            return None
        m = re.match(r"Heading\s+(\d+)$", str(style_name).strip(), flags=re.IGNORECASE)
        if m:
            try:
                return int(m.group(1))
            except Exception:
                return None
        return None

    @staticmethod
    def _outline_level_from_p(paragraph) -> Optional[int]:
        el = paragraph._element
        outlvl_nodes = XP_P_OUTLINE(el)
        if outlvl_nodes:
            try:
                val = outlvl_nodes[0].get(f"{{{W_NS}}}val")
                if val is not None:
                    return int(val) + 1
            except Exception:
                return None
        return None

    # ---------- Regex mode helpers ----------
    def _regex_choose_splitter(self, lines: List[str]) -> Optional[object]:
        for line in lines:
            s = line.strip()
            if not s:
                continue
            for idx, (kind, pat) in enumerate(REGEX_ORDER):
                if kind == "fixed":
                    cp = COMPILED_FIXED[idx]
                    if cp is not None and cp.match(s):
                        return FixedRegexSplitter(cp)
                elif kind == "numeric_dotted":
                    m = NUMERIC_DOTTED.match(s)
                    if m:
                        depth = len([seg for seg in re.split(r'[\.．]', m.group(1)) if seg])
                        return NumericDepthSplitter(depth)
        return None

    def _regex_split_into_nodes(self, parent: MyDOCNode, level: int, lines: List[str]) -> List[MyDOCNode]:
        nodes: List[MyDOCNode] = []
        splitter = self._regex_choose_splitter(lines)
        found_first = False
        current: Optional[MyDOCNode] = None

        if splitter is None:
            parent.body_paragraphs.extend([ln for ln in lines if ln and ln.strip()])
            return nodes

        for ln in lines:
            s = ln.strip()
            if splitter.matches(s):
                index = len(nodes) + 1
                props = {"mode": "regex"}
                title = _strip_pstyle_marker(s)
                node = MyDOCNode(name=title, level=level, index=index, parent=parent,
                                 element_type="heading", properties=props)
                parent.add_child(node)
                nodes.append(node)
                current = node
                found_first = True
            else:
                if not found_first:
                    if s:
                        parent.body_paragraphs.append(ln)
                else:
                    current._pending_for_children.append(ln)
        return nodes

    def _regex_build_recursive(self, parent: MyDOCNode, level: int, lines: List[str], max_depth: int = 200):
        if level > max_depth:
            parent.body_paragraphs.extend([ln for ln in lines if ln and ln.strip()])
            return
        children = self._regex_split_into_nodes(parent, level, lines)
        if not children:
            return
        for child in children:
            if child._pending_for_children:
                child_lines = child._pending_for_children
                child._pending_for_children = []
                self._regex_build_recursive(child, level + 1, child_lines, max_depth=max_depth)

    def _regex_finalize_pending_to_body(self, node: MyDOCNode):
        if node._pending_for_children:
            node.body_paragraphs.extend([ln for ln in node._pending_for_children if ln and ln.strip()])
            node._pending_for_children = []
        for ch in node.children:
            self._regex_finalize_pending_to_body(ch)

    # ---------- Build tree ----------
    def build_tree(self):
        if self.mode == "heading":
            self._build_tree_heading_mode()
        elif self.mode == "regex":
            self._build_tree_regex_mode()
        else:
            logger.info("Building hierarchy (hybrid): heading first, then regex refine bodies...")
            self._build_tree_heading_mode()
            self._refine_bodies_with_regex(self.root)
            logger.info("Hybrid build complete.")

    def _iter_block_items(self):
        """Yield ('p', paragraph) or ('tbl', table) in document order."""
        body = self.doc._element.body
        p_idx = 0
        t_idx = 0
        for child in list(body):
            if child.tag.endswith('}p'):
                yield ('p', self.doc.paragraphs[p_idx])
                p_idx += 1
            elif child.tag.endswith('}tbl'):
                yield ('tbl', self.doc.tables[t_idx])
                t_idx += 1

    def _build_tree_heading_mode(self):
        logger.info("Building hierarchy (heading mode)...")
        current = self.root
        stack: List[MyDOCNode] = [self.root]

        def _tbl_width_pt(tbl_el):
            # 读取 w:tblW；若为 pct，用 480pt 近似 100%（避免缺页面宽时归一失败）
            tw = tbl_el.find("./w:tblPr/w:tblW", NSMAP)
            if tw is not None:
                t = tw.get(f"{{{W_NS}}}type")
                v = tw.get(f"{{{W_NS}}}w")
                if v and t in (None, "dxa"):
                    pt = _twip_to_pt(v)
                    if pt:
                        return max(pt, 1.0)
                if v and t == "pct":
                    try:
                        return max((float(v) / 50.0) * 4.8, 1.0)  # 100%≈480pt
                    except Exception:
                        pass
            return 480.0

        def _expanded_cols(row_cells):
            # 计算“考虑 colspan 展开”的列数
            n = 0
            for cell in row_cells:
                try:
                    cs = int(cell.get("colspan", 1)) or 1
                except Exception:
                    cs = 1
                n += max(1, cs)
            return n

        for kind, obj in self._iter_block_items():
            if kind == "p":
                p = obj
                raw_text = (p.text or "").strip()

                level = self._outline_level_from_p(p)
                if level is None:
                    try:
                        style_name = getattr(p.style, "name", None)
                    except Exception:
                        style_name = None
                    lvl2 = self._style_based_heading_level(style_name)
                    level = lvl2 if lvl2 is not None else 0

                if level > 0 and (raw_text or True):
                    # 调整栈
                    while len(stack) > level:
                        stack.pop()
                    while len(stack) < level:
                        dummy = MyDOCNode(name="", level=len(stack), index=1, parent=stack[-1], element_type="heading")
                        stack[-1].add_child(dummy)
                        stack.append(dummy)

                    parent = stack[level - 1]
                    index = sum(1 for c in parent.children if c.level == level) + 1
                    props = {"style": getattr(p.style, "name", None), "outline_level": level, "mode": "heading"}
                    heading_text = self._paragraph_text_with_refs(p, include_pstyle=False) or raw_text
                    node = MyDOCNode(name=heading_text, level=level, index=index,
                                     parent=parent, element_type="heading", properties=props)
                    parent.add_child(node)
                    if len(stack) == level:
                        stack.append(node)
                    else:
                        stack[level] = node
                        stack = stack[:level + 1]
                    current = node
                else:
                    text_with_refs = self._paragraph_text_with_refs(p, include_pstyle=True)
                    if text_with_refs or raw_text:
                        target = current if current is not None else self.root
                        if text_with_refs.strip():
                            target.body_paragraphs.append(text_with_refs)

            elif kind == "tbl":
                # === 表格导出：保证 cols 与 colWidthsPt 准确 ===
                tbl_el = obj._element  # w:tbl
                tableWidthPt = _tbl_width_pt(tbl_el)

                # 1) 表属性
                headerRows = 1 if tbl_el.find("./w:tblPr/w:tblHeader", NSMAP) is not None else 0
                jc = tbl_el.find("./w:tblPr/w:jc", NSMAP)
                tableAlign = (jc.get(f"{{{W_NS}}}val") if jc is not None else None) or "left"

                borders = {"inner": 0.5, "outer": 0.75}
                tblBorders = tbl_el.find("./w:tblPr/w:tblBorders", NSMAP)

                def _edge_weight(ed):
                    if ed is None:
                        return None
                    sz = ed.get(f"{{{W_NS}}}sz")
                    try:
                        return (float(sz) or 4.0) / 8.0  # Word sz=4≈0.5pt
                    except Exception:
                        return None

                if tblBorders is not None:
                    ins = [_edge_weight(tblBorders.find(f"./w:{n}", NSMAP)) for n in ("insideH", "insideV")]
                    outs = [_edge_weight(tblBorders.find(f"./w:{n}", NSMAP)) for n in
                            ("top", "bottom", "left", "right")]
                    borders["inner"] = next((v for v in ins if v), borders["inner"])
                    borders["outer"] = next((v for v in outs if v), borders["outer"])

                cellPadding = None
                tcMar = tbl_el.find("./w:tblPr/w:tblCellMar", NSMAP)
                if tcMar is not None:
                    def _pad(which):
                        el = tcMar.find(f"./w:{which}", NSMAP)
                        return round(_twip_to_pt(el.get(f"{{{W_NS}}}w")) or 3.0, 2) if el is not None else None

                    cellPadding = {"t": _pad("top"), "l": _pad("left"), "b": _pad("bottom"), "r": _pad("right")}

                # 2) 读取表格内容（含合并/对齐/底纹），先得 rows_data
                rows_data: List[List[dict]] = []
                MAX_ROWS, MAX_COLS, MAX_SPAN = 500, 200, 50
                all_tr = tbl_el.findall("./w:tr", NSMAP)
                for r_idx, tr in enumerate(all_tr):
                    if r_idx >= MAX_ROWS:
                        break
                    row_cells = []
                    tcs = tr.findall("./w:tc", NSMAP)
                    c_vis = 0
                    for tc in tcs:
                        if c_vis >= MAX_COLS:
                            break
                        tcPr = tc.find("./w:tcPr", NSMAP)

                        # 水平/垂直对齐
                        align = "left"
                        p_first = tc.find("./w:p/w:pPr/w:jc", NSMAP)
                        if p_first is not None and p_first.get(f"{{{W_NS}}}val"):
                            align = p_first.get(f"{{{W_NS}}}val")
                        valign = "top"
                        vAli = tcPr.find("./w:vAlign", NSMAP) if tcPr is not None else None
                        if vAli is not None and vAli.get(f"{{{W_NS}}}val"):
                            valign = vAli.get(f"{{{W_NS}}}val")

                        # 底纹
                        shading = None
                        sh = tcPr.find("./w:shd", NSMAP) if tcPr is not None else None
                        if sh is not None and sh.get(f"{{{W_NS}}}fill"):
                            fill = sh.get(f"{{{W_NS}}}fill")
                            if fill and fill != "auto":
                                shading = f"#{fill}" if not fill.startswith("#") else fill

                        # gridSpan/colspan
                        gridSpan = tcPr.find("./w:gridSpan", NSMAP) if tcPr is not None else None
                        colspan = 1
                        if gridSpan is not None and gridSpan.get(f"{{{W_NS}}}val"):
                            try:
                                colspan = int(gridSpan.get(f"{{{W_NS}}}val"))
                            except Exception:
                                colspan = 1
                        colspan = max(1, min(MAX_SPAN, colspan))

                        # vMerge：restart / continue
                        vmerge = tcPr.find("./w:vMerge", NSMAP) if tcPr is not None else None
                        vval = vmerge.get(f"{{{W_NS}}}val") if vmerge is not None else None
                        is_continue = (vmerge is not None and (vval in (None, "", "continue", "cont", "1")))
                        is_restart = (vmerge is not None and (vval in ("restart", "rest", "0")))

                        # 文本（含换行/脚注尾注）
                        texts = []
                        for p_el in tc.findall(".//w:p", NSMAP):
                            parts = []
                            for t in p_el.findall(".//w:t", NSMAP):
                                if t.text:
                                    parts.append(t.text)
                            if p_el.findall(".//w:tab", NSMAP):
                                parts.append("\t")
                            for fr in p_el.findall(".//w:footnoteReference", NSMAP):
                                rid = fr.get(f"{{{W_NS}}}id")
                                parts.append(f"[[FNREF:{rid}]]")
                            for er in p_el.findall(".//w:endnoteReference", NSMAP):
                                rid = er.get(f"{{{W_NS}}}id")
                                parts.append(f"[[ENREF:{rid}]]")
                            txt = "".join(parts).strip()
                            if txt:
                                texts.append(txt)
                        cell_text = "\n".join(texts)

                        if is_continue:
                            row_cells.append({"text": "", "colspan": 1, "rowspan": 0, "align": align, "valign": valign})
                            c_vis += 1
                            continue

                        # 计算 rowspan：向下统计 continue
                        rowspan = 1
                        if is_restart:
                            down = r_idx + 1
                            while down < len(all_tr):
                                tlist = all_tr[down].findall("./w:tc", NSMAP)
                                if len(tlist) == 0: break
                                found_continue = False
                                for n_tc in tlist:
                                    n_pr = n_tc.find("./w:tcPr", NSMAP)
                                    n_vm = n_pr.find("./w:vMerge", NSMAP) if n_pr is not None else None
                                    if n_vm is None:
                                        found_continue = False;
                                        break
                                    nv = n_vm.get(f"{{{W_NS}}}val")
                                    if nv in (None, "", "continue", "cont", "1"):
                                        found_continue = True;
                                        break
                                    else:
                                        found_continue = False;
                                        break
                                if found_continue:
                                    rowspan += 1;
                                    down += 1
                                else:
                                    break

                        row_cells.append({
                            "text": cell_text,
                            "colspan": colspan,
                            "rowspan": max(1, rowspan),
                            "align": align,
                            "valign": valign,
                            "shading": shading
                        })
                        c_vis += 1

                    rows_data.append(row_cells)

                # 3) 真实列数（考虑 colspan 展开）
                rows = len(rows_data)
                cols = 0
                for r in rows_data:
                    cols = max(cols, _expanded_cols(r))

                # 4) 生成 colWidthsPt
                colWidthsPt: List[float] = []
                grid_cols = tbl_el.findall("./w:tblGrid/w:gridCol", NSMAP)
                for gc in grid_cols:
                    w = gc.get(f"{{{W_NS}}}w")
                    pt = _twip_to_pt(w)
                    if pt is not None:
                        colWidthsPt.append(round(pt, 2))

                if (not colWidthsPt) or (len(colWidthsPt) != cols):
                    # 4.1 尝试用首行 tcW（dxa/pct），按 colspan 平均分配到网格列
                    first_tr = all_tr[0] if all_tr else None
                    widths_acc = [0.0] * max(cols, len(colWidthsPt) or 0 or cols)
                    used_tcW = False
                    if first_tr is not None:
                        tcs = first_tr.findall("./w:tc", NSMAP)
                        cur_col = 0
                        for tc in tcs:
                            tcPr = tc.find("./w:tcPr", NSMAP)
                            # 宽度
                            wpt = None
                            if tcPr is not None:
                                tcW = tcPr.find("./w:tcW", NSMAP)
                                if tcW is not None and tcW.get(f"{{{W_NS}}}w"):
                                    wtype = tcW.get(f"{{{W_NS}}}type")
                                    wval = tcW.get(f"{{{W_NS}}}w")
                                    if wtype in (None, "dxa"):
                                        wpt = _twip_to_pt(wval)
                                    elif wtype == "pct":
                                        try:
                                            pct = float(wval) / 50.0  # 100% = 5000
                                            wpt = pct * 4.8  # 近似到 pt
                                        except Exception:
                                            wpt = None
                            # 跨列
                            gridSpan = tcPr.find("./w:gridSpan", NSMAP) if tcPr is not None else None
                            cs = 1
                            if gridSpan is not None and gridSpan.get(f"{{{W_NS}}}val"):
                                try:
                                    cs = max(1, min(50, int(gridSpan.get(f"{{{W_NS}}}val"))))
                                except Exception:
                                    cs = 1
                            if wpt is None:
                                wpt = tableWidthPt * (cs / float(cols or 1))
                            share = float(wpt) / float(cs or 1)
                            for k in range(cs):
                                if cur_col + k < len(widths_acc):
                                    widths_acc[cur_col + k] += share
                            cur_col += cs
                            used_tcW = True
                    if used_tcW:
                        colWidthsPt = [round(max(1.0, v), 2) for v in widths_acc[:cols]]
                    else:
                        # 4.2 完全拿不到：用表宽等分
                        each = max(1.0, tableWidthPt / float(cols or 1))
                        colWidthsPt = [round(each, 2) for _ in range(cols)]

                # 4.3 截断/补齐，并按表宽归一
                if len(colWidthsPt) > cols:
                    colWidthsPt = colWidthsPt[:cols]
                elif len(colWidthsPt) < cols:
                    missing = cols - len(colWidthsPt)
                    each = max(1.0, tableWidthPt / float(cols or 1))
                    colWidthsPt += [round(each, 2) for _ in range(missing)]
                s = sum(colWidthsPt) or 1.0
                scale = tableWidthPt / s
                colWidthsPt = [round(max(1.0, w * scale), 2) for w in colWidthsPt]

                # 列比例（供 InDesign 端按框宽缩放）
                _sum = sum(colWidthsPt) or 1.0
                colWidthFrac = [round(w / _sum, 6) for w in colWidthsPt]
                table_obj = {
                    "rows": rows,
                    "cols": cols,
                    "data": rows_data,
                    "colWidthsPt": colWidthsPt,
                    "colWidthFrac": colWidthFrac,
                    "tableWidthPt": round(tableWidthPt, 2),
                    "headerRows": headerRows,
                    "tableAlign": tableAlign,
                }
                if cellPadding:
                    table_obj["cellPadding"] = cellPadding
                if borders:
                    table_obj["borders"] = borders

                ph = "[[TABLE " + json.dumps(table_obj, ensure_ascii=False) + "]]"
                target = stack[-1] if stack else self.root
                target.body_paragraphs.append(ph)
                continue

        logger.info("Heading tree build complete.")

    def _build_tree_regex_mode(self):
        logger.info("Building hierarchy (regex mode, hierarchical segmentation with dynamic numeric depth)...")
        lines: List[str] = []
        for p in self.doc.paragraphs:
            line = self._paragraph_text_with_refs(p, include_pstyle=True)
            lines.append(line)
        self._regex_build_recursive(self.root, level=1, lines=lines, max_depth=200)
        self._regex_finalize_pending_to_body(self.root)
        logger.info("Regex tree build complete.")

    # ---------- Hybrid refinement ----------
    def _refine_bodies_with_regex(self, node: MyDOCNode):
        if node.level >= 1 and node.body_paragraphs:
            lines = node.body_paragraphs[:]
            node.body_paragraphs = []
            before = len(node.children)
            self._regex_build_recursive(node, level=node.level + 1, lines=lines, max_depth=200)
            self._regex_finalize_pending_to_body(node)
            after = len(node.children)
            if after > before:
                newkids = node.children[before:]
                oldkids = node.children[:before]
                node.children = newkids + oldkids
        for ch in list(node.children):
            self._refine_bodies_with_regex(ch)

    # ---------- XML export ----------
    @staticmethod
    def _escape(text: str) -> str:
        return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

    def _notes_to_xml(self) -> str:
        parts = []
        if self.footnotes:
            parts.append("  <footnotes>")
            for nid, body in sorted(self.footnotes.items(), key=lambda kv: int(kv[0])):
                parts.append(f'    <footnote id="{nid}">{self._escape(body)}</footnote>')
            parts.append("  </footnotes>")
        if self.endnotes:
            parts.append("  <endnotes>")
            for nid, body in sorted(self.endnotes.items(), key=lambda kv: int(kv[0])):
                parts.append(f'    <endnote id="{nid}">{self._escape(body)}</endnote>')
            parts.append("  </endnotes>")
        return "\n".join(parts)

    def to_xml(self, output_path: str):
        logger.info(f"Writing XML -> {output_path}")
        parts = ['<?xml version="1.0" encoding="UTF-8"?>', "<document>"]
        if self.root.body_paragraphs:
            parts.append("  <body>")
            for para in self.root.body_paragraphs:
                text, pattrs = _parse_pstyle_marker(para) if STYLE_FLAGS.get("paragraph", True) else (para, {})
                esc = MyDOCNode._escape_xml(text)
                esc = MyDOCNode._convert_refs_to_xml(esc)
                attrs_str = ""
                if STYLE_FLAGS.get("paragraph", True) and pattrs:
                    kvs = [f'{k}="{MyDOCNode._escape_attr(v)}"' for k,v in pattrs.items() if v]
                    if kvs:
                        attrs_str = " " + " ".join(kvs)
                parts.append(f"    <p{attrs_str}>{esc}</p>")
            parts.append("  </body>")
        for child in self.root.children:
            parts.append(child.to_xml_string(notes={"footnotes": self.footnotes, "endnotes": self.endnotes}, indent=1))
        notes_xml = self._notes_to_xml()
        if notes_xml:
            parts.append(notes_xml)
        parts.append("</document>")

        with open(output_path, "w", encoding="utf-8") as f:
            f.write("\n".join(parts))
        logger.info("Done.")

    def process(self, output_path: str):
        # set assets dir next to output
        outdir = os.path.dirname(os.path.abspath(output_path)) or "."
        self.assets_dir = os.path.join(outdir, "assets")
        self.extract_notes()
        self.build_tree()
        self.to_xml(output_path)

def main(argv=None):
    parser = argparse.ArgumentParser(description="DOCX -> XML exporter (heading/regex/hybrid) with style switches")
    parser.add_argument("--mode", choices=["heading", "regex", "hybrid"], default="heading", help="Detection mode")
    parser.add_argument("input", help="Input .docx path")
    parser.add_argument("output", help="Output .xml path")
    args = parser.parse_args(argv)

    exporter = DOCXOutlineExporter(args.input, mode=args.mode)
    exporter.process(args.output)
    # print(f"[OK] mode={args.mode} XML saved -> {args.output}")

if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
