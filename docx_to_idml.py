# -*- coding: utf-8 -*-
"""
docx_to_idml.py（跨平台增强版：Windows + macOS）
- 维持原有：内置口令 + 文件口令（PBKDF2），启动必须验证口令；支持 --set-password 修改本地口令
- 新增：在“输入密码之前”校验 input 与 TEMPLATE_PATH（实际生效）是否存在；--out 的父目录若不存在则直接报错
- 新增：--template/-t 覆盖 xml_to_idml.TEMPLATE_PATH；--out/-o 覆盖 xml_to_idml.IDML_OUT_PATH
- 密码文件位置改为跨平台更安全的用户目录：
  * Windows: %LOCALAPPDATA%\\Docx2IDML\\docx_to_idml_pass.json
  * macOS : ~/Library/Application Support/Docx2IDML/docx_to_idml_pass.json
  * 其他   : ~/.config/Docx2IDML/docx_to_idml_pass.json
- 仍按平台自动调用 InDesign：Windows 走 COM，macOS 走 AppleScript（xml_to_idml 中已修正 language=javascript）
"""
from __future__ import annotations

import os
import sys
import json
import argparse
import getpass
import base64
import hashlib
import time
import zipfile
import logging
from xml.etree import ElementTree as ET
from typing import Optional

def _runtime_base_dir() -> str:
    """
    Resolve a stable base dir (prefer real exe dir, avoid onefile temp).
    Order: NUITKA_ONEFILE_PARENT > argv[0] dir (if exists) > sys.executable dir > this file dir.
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

BASE_DIR = _runtime_base_dir()
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

import xml_to_idml as X
from docx_to_xml_outline_notes_v13 import DOCXOutlineExporter
from xml_to_idml import XML_PATH, extract_paragraphs_with_levels, write_jsx, JSX_PATH
from xml_to_idml import AUTO_RUN_MACOS, AUTO_RUN_WINDOWS, run_indesign_windows, run_indesign_macos
from xml_to_idml import LOG_PATH  
from pipeline_logger import PipelineLogger

PIPELINE_LOGGER: Optional[PipelineLogger] = None

def _default_pass_dir() -> str:
    if sys.platform == "darwin":
        return os.path.join(os.path.expanduser("~"), "Library", "Application Support", "Docx2IDML")
    if os.name == "nt":
        base = os.environ.get("LOCALAPPDATA") or os.path.join(os.path.expanduser("~"), "AppData", "Local")
        return os.path.join(base, "Docx2IDML")
    return os.path.join(os.path.expanduser("~"), ".config", "Docx2IDML")

PASS_DIR  = _default_pass_dir()
os.makedirs(PASS_DIR, exist_ok=True)
PASS_FILE = os.path.join(PASS_DIR, "docx_to_idml_pass.json")

PBKDF2_ITER = 200_000

EMBEDDED_RECORD = {
    "iterations": PBKDF2_ITER,
    "salt_b64": "ZiUILFcbXqKBcLTE5kgCjw==",
    "hash_b64": "e7kws+WYofc+ChQIYGqeMCmH0P3+y7NGkOJR0drGWXU="
}


def _embedded_record_parsed():
    return {
        "iterations": int(EMBEDDED_RECORD["iterations"]),
        "salt": base64.b64decode(EMBEDDED_RECORD["salt_b64"]),
        "hash": base64.b64decode(EMBEDDED_RECORD["hash_b64"]),
    }


def _pbkdf2_hash(password: str, salt: bytes, iterations: int = PBKDF2_ITER) -> bytes:
    return hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, iterations, dklen=32)


def _load_record():
    try:
        if os.path.exists(PASS_FILE):
            with open(PASS_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            salt = base64.b64decode(data["salt"])
            hsh  = base64.b64decode(data["hash"])
            iters = int(data.get("iterations", PBKDF2_ITER))
            return {"salt": salt, "hash": hsh, "iterations": iters, "source": "file"}
    except Exception:
        pass
    # fallback: embedded
    rec = _embedded_record_parsed()
    rec["source"] = "embedded"
    return rec


def _save_record(salt: bytes, hsh: bytes, iterations: int = PBKDF2_ITER):
    data = {
        "iterations": iterations,
        "salt": base64.b64encode(salt).decode("ascii"),
        "hash": base64.b64encode(hsh).decode("ascii"),
    }
    with open(PASS_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)




def _log_user(message: str):
    print(message)
    if PIPELINE_LOGGER:
        PIPELINE_LOGGER.user(message)


def _log_warn(message: str):
    print(message)
    if PIPELINE_LOGGER:
        PIPELINE_LOGGER.warn(message)


def _log_error(message: str):
    print(message)
    if PIPELINE_LOGGER:
        PIPELINE_LOGGER.error(message)

def _debug_log(message: str):
    if PIPELINE_LOGGER:
        PIPELINE_LOGGER.debug(message)


def _read_docx_app_props(docx_path: str) -> dict:
    stats = {"pages": None, "paragraphs": None, "tables": None}
    try:
        with zipfile.ZipFile(docx_path) as zf:
            data = zf.read("docProps/app.xml")
        root = ET.fromstring(data)
        for node in root.iter():
            tag = node.tag.split("}", 1)[-1]
            if tag == "Pages":
                stats["pages"] = int(node.text or 0)
            elif tag == "Paragraphs":
                stats["paragraphs"] = int(node.text or 0)
            elif tag == "Tables":
                stats["tables"] = int(node.text or 0)
    except Exception:
        pass
    return stats


def _count_idml_pages(idml_path: str) -> Optional[int]:
    if not idml_path or not os.path.exists(idml_path):
        return None
    try:
        total = 0
        with zipfile.ZipFile(idml_path) as zf:
            for name in zf.namelist():
                if not name.lower().startswith("spreads/") or not name.lower().endswith(".xml"):
                    continue
                data = zf.read(name)
                try:
                    root = ET.fromstring(data)
                    total += sum(1 for _ in root.iter() if _.tag.split("}", 1)[-1] == "Page")
                except Exception:
                    continue
        return total if total > 0 else None
    except Exception:
        return None


def _prompt_hidden(prompt_text: str) -> str:
    try:
        return getpass.getpass(prompt_text)
    except Exception:
        return input(prompt_text)


def _verify_flow(cli_password: str | None) -> bool:
    rec = _load_record()
    pwd = cli_password if cli_password is not None else _prompt_hidden("请输入密码：")
    h = _pbkdf2_hash(pwd, rec["salt"], rec["iterations"])
    if h == rec["hash"]:
        return True
    _log_error("密码错误。")
    return False


def _set_password_flow():
    rec = _load_record()
    old = _prompt_hidden("请输入当前密码：")
    if _pbkdf2_hash(old, rec["salt"], rec["iterations"]) != rec["hash"]:
        _log_error("当前密码不正确，无法修改。")
        return 2
    while True:
        p1 = _prompt_hidden("请输入新密码：")
        p2 = _prompt_hidden("请再次输入确认：")
        if p1 != p2:
            _log_warn("两次输入不一致，请重试。")
            continue
        if len(p1) < 6:
            _log_warn("建议使用至少 6 位密码，请重试。")
            continue
        salt = os.urandom(16)
        hsh  = _pbkdf2_hash(p1, salt, PBKDF2_ITER)
        _save_record(salt, hsh, PBKDF2_ITER)
        _log_user("[安全] 密码已修改（写入本地文件）。注意：若删除密码文件，将回退到内置默认密码保护。")
        return 0


# ====== 模板与输出覆盖 ======
def _effective_template_path(cli_template: str | None) -> str:
    if cli_template and str(cli_template).strip():
        return os.path.abspath(cli_template)
    try:
        return os.path.abspath(X.TEMPLATE_PATH)
    except Exception:
        return os.path.abspath("template.idml")


def _apply_overrides(cli_template: str | None, cli_out: str | None) -> None:
    if cli_template and str(cli_template).strip():
        X.TEMPLATE_PATH = os.path.abspath(cli_template)
    if cli_out and str(cli_out).strip():
        X.IDML_OUT_PATH = os.path.abspath(cli_out)


def main(argv=None):
    start_ts = time.time()
    parser = argparse.ArgumentParser(description="DOCX -> XML exporter (heading/regex) with style switches + Password Protection + Cross-platform paths")
    parser.add_argument(
        "--mode",
        choices=["heading", "regex"],
        metavar="{heading,regex}",
        default="heading",
        help="Detection mode：heading 或 regex",
    )
    parser.add_argument(
        "--regex-config",
        help="指定 regex_rules.json，用于 --mode=regex 时自定义正则规则",
    )
    parser.add_argument(
        "--regex-max-depth",
        type=int,
        default=None,
        help="正则分级最大层级，0 表示不限制（默认 200）",
    )
    parser.add_argument("--set-password", action="store_true", help="修改程序密码（需验证当前密码）")
    parser.add_argument("--password", default=None, help="以非交互形式提供密码（可选，留空则提示输入）")
    parser.add_argument("--template", "-t", dest="template", default=None, help="覆盖 TEMPLATE_PATH，模板 .idml 的路径")
    parser.add_argument("--out", "-o", dest="out", default=None, help="覆盖 IDML_OUT_PATH，输出 .idml 路径")
    parser.add_argument("input", nargs="?", help="Input .docx path")
    parser.add_argument("--no-images", action="store_true", help="不处理图片")
    parser.add_argument("--no-tables", action="store_true", help="不处理表格")
    parser.add_argument("--no-textboxes", action="store_true", default=True, help=argparse.SUPPRESS)
    parser.add_argument("--log-dir", help="日志文件目录")
    parser.add_argument("--debug-log", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--no-run", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--no-inline-list-labels", dest="inline_list_labels", action="store_false", help="不把列表编号前缀写回正文，仅保留为元数据（默认写回，与 Word 视觉一致）")
    parser.set_defaults(inline_list_labels=True)
    args = parser.parse_args(argv)

    input_path = None
    if not args.set_password:
        if not args.input:
            _log_error("缺少参数：input .docx 路径")
            sys.exit(2)
        input_path = os.path.abspath(args.input)
        if not os.path.exists(input_path):
            _log_error(f"输入文件不存在：{input_path}")
            sys.exit(2)

    eff_template = _effective_template_path(args.template)
    if not os.path.exists(eff_template):
        _log_error(f"模板文件不存在：{eff_template}")
        sys.exit(2)

    if args.out:
        out_abs = os.path.abspath(args.out)
        out_dir = os.path.dirname(out_abs) or os.getcwd()
        if out_dir and not os.path.exists(out_dir):
            _log_error(f"输出目录不存在：{out_dir}")
            sys.exit(2)

    _apply_overrides(args.template, args.out)
    # 未指定 --out 时，IDML 输出名使用输入 DOCX 同名，放在 JSX 所在目录
    if not args.out:
        try:
            base_name = os.path.splitext(os.path.basename(input_path))[0]
            out_dir = os.path.abspath(os.path.dirname(JSX_PATH))
            X.IDML_OUT_PATH = os.path.join(out_dir, base_name + ".idml")
            _debug_log(f"[IDML] set output path: {X.IDML_OUT_PATH}")
        except Exception as e:
            _log_warn(f"[WARN] 设置 IDML 输出名失败，仍使用默认: {e}")
    arg_debug_line = (
        f"[ARGS] mode={args.mode} regex_cfg={args.regex_config or '(default)'} "
        f"regex_max_depth={args.regex_max_depth if args.regex_max_depth is not None else '(default 200)'} "
        f"skip_images={args.no_images} skip_tables={args.no_tables} "
        f"skip_textboxes={args.no_textboxes} template={eff_template} out={getattr(X, 'IDML_OUT_PATH', None)} "
        f"log_dir={args.log_dir or '(default)'} no_run={args.no_run}"
    )

    if args.set_password:
        code = _set_password_flow()
        sys.exit(code if isinstance(code, int) else 0)

    if not _verify_flow(args.password):
        sys.exit(1)

    global PIPELINE_LOGGER, LOG_PATH
    log_source = input_path or XML_PATH
    PIPELINE_LOGGER = PipelineLogger(
        log_source,
        log_root=args.log_dir,
        enable_debug=args.debug_log,
        console_echo=False,
    )
    X.PIPELINE_LOGGER = PIPELINE_LOGGER
    X.LOG_PATH = str(PIPELINE_LOGGER.jsx_event_log_path)
    LOG_PATH = X.LOG_PATH
    X.LOG_WRITE = args.debug_log
    if args.debug_log:
        # reduce console noise; forward all logs to pipeline debug log
        root_logger = logging.getLogger()
        for h in list(root_logger.handlers):
            try:
                h.setLevel(logging.WARNING)
            except Exception:
                pass
        class _PipelineLogHandler(logging.Handler):
            def emit(self_inner, record):
                try:
                    msg = record.getMessage()
                except Exception:
                    msg = str(record)
                if PIPELINE_LOGGER and PIPELINE_LOGGER.enable_debug:
                    PIPELINE_LOGGER.debug(f"[{record.name}] {msg}")
        root_logger.setLevel(logging.DEBUG)
        root_logger.addHandler(_PipelineLogHandler())
    if args.debug_log:
        PIPELINE_LOGGER.describe_paths()
        _log_user(f"[LOG] 用户日志: {PIPELINE_LOGGER.user_log_path}")
        _log_user(f"[LOG] 调试日志: {PIPELINE_LOGGER.debug_log_path}")
    PIPELINE_LOGGER.user(arg_debug_line)
    _debug_log(arg_debug_line)

    # 1) XML
    docx_meta = _read_docx_app_props(input_path)

    exporter = DOCXOutlineExporter(
        input_path,
        mode=args.mode,
        regex_config_path=args.regex_config,
        regex_max_depth=args.regex_max_depth,
        skip_images=args.no_images,
        skip_tables=args.no_tables,
        skip_textboxes=args.no_textboxes,
        inline_list_labels=args.inline_list_labels,
    )
    if args.mode == "regex":
        rules_path = getattr(exporter, "regex_rules_path", None)
        if rules_path:
            _log_user(f"[INFO] regex 规则来源: {rules_path}")
        else:
            _log_user("[INFO] regex 规则使用内置默认")
    export_summary = exporter.process(XML_PATH)
    if not export_summary:
        export_summary = exporter.summary()
    image_stats = export_summary.get("image_fragments", 0) or 0
    if args.debug_log:
        _log_user(f"[OK] mode={args.mode} XML saved -> {XML_PATH}")

    # 2)  XML -> paragraphs
    paragraphs = extract_paragraphs_with_levels(XML_PATH)
    if args.debug_log:
        _log_user(f"[INFO] 解析到 {len(paragraphs)} 段；示例前3段: {paragraphs[:3]}")
    if PIPELINE_LOGGER and args.debug_log:
        PIPELINE_LOGGER.debug(
            f"[DOCX2IDML] skip_flags images={args.no_images} tables={args.no_tables} textboxes={args.no_textboxes}"
        )

    write_jsx(JSX_PATH, paragraphs)
    if args.debug_log:
        _log_user(f"[OK] JSX 写入: {JSX_PATH}")
        _log_user(f"[INFO] JSX 模板来源: {X.TEMPLATE_PATH}")
        _log_user(f"[INFO] JSX 事件日志: {LOG_PATH}")

    ran = False
    if AUTO_RUN_WINDOWS and sys.platform.startswith("win"):
        ran = run_indesign_windows(JSX_PATH)
    elif AUTO_RUN_MACOS and sys.platform == "darwin":
        ran = run_indesign_macos(JSX_PATH)

    # 生成 IDML 文件名与输入 DOCX 同名
    try:
        base_name = os.path.splitext(os.path.basename(input_path))[0]
        out_dir = os.path.abspath(os.path.dirname(JSX_PATH))
        X.IDML_OUT_PATH = os.path.join(out_dir, base_name + ".idml")
        _debug_log(f"[IDML] set output path: {X.IDML_OUT_PATH}")
    except Exception as e:
        if args.debug_log:
            _log_user(f"[WARN] 设置 IDML 输出名失败，仍使用默认: {e}")

    _log_user("\n=== 完成 ===")
    if args.debug_log:
        _log_user(f"XML: {XML_PATH}")
        _log_user(f"JSX: {JSX_PATH}")
        _log_user(f"LOG: {LOG_PATH}")

    _log_user(f"IDML: {getattr(X, 'IDML_OUT_PATH', None)}")
    stats = X._relay_jsx_events(
        PIPELINE_LOGGER, LOG_PATH, warn_missing=not ran, cleanup=not args.debug_log
    )
    summary_line = (
        f"[REPORT] JSX 事件统计 info={stats.get('info', 0)} "
        f"warn={stats.get('warn', 0)} error={stats.get('error', 0)} "
        f"debug={stats.get('debug', 0)}"
    )
    _log_user(summary_line)
    if ran:
        _log_user("InDesign 已执行 JSX；若设置 AUTO_EXPORT_IDML=True，将在脚本目录生成idml文件。")

    converted_tables = 0
    for _, text in paragraphs:
        if not text:
            continue
        converted_tables += text.count("[[TABLE")
    converted_paragraphs = len(paragraphs)
    converted_images = image_stats
    converted_pages = None
    if ran:
        converted_pages = _count_idml_pages(getattr(X, "IDML_OUT_PATH", ""))

    word_pages = docx_meta.get("pages") or export_summary.get("word_pages")
    word_tables = export_summary.get("word_tables") or docx_meta.get("tables")
    word_paragraphs = export_summary.get("word_paragraphs") or docx_meta.get("paragraphs")

    total_elapsed = time.time() - start_ts
    elapsed_msg = f"[TIME] total={total_elapsed:.2f}s"
    _log_user(elapsed_msg)

    summary_report = (
        f"[REPORT][SUMMARY] wordPages={word_pages if word_pages is not None else 'N/A'} "
        f"wordTables={word_tables if word_tables is not None else 'N/A'} "
        f"wordParagraphs={word_paragraphs if word_paragraphs is not None else 'N/A'} "
        f"convertedPages={converted_pages if converted_pages is not None else 'N/A'} "
        f"convertedTables={converted_tables} convertedParagraphs={converted_paragraphs} "
        f"convertedImages={converted_images} "
        f"elapsed={total_elapsed:.2f}s"
    )
    _log_user(summary_report)
    # 清理中间产物：未开启 debug-log 时移除 XML 和 JSX，避免暴露技术文件
    if not args.debug_log:
        for path in (XML_PATH, JSX_PATH):
            try:
                if path and os.path.exists(path):
                    os.remove(path)
                    _debug_log(f"[CLEANUP] removed {path}")
            except Exception as e:
                _log_user(f"[WARN] 清理中间文件失败: {e}")



if __name__ == "__main__":
    main()
