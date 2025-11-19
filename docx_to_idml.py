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
from typing import Optional

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

import xml_to_idml as X
from docx_to_xml_outline_notes_v13 import DOCXOutlineExporter
from xml_to_idml import XML_PATH, extract_paragraphs_with_levels, write_jsx, JSX_PATH
from xml_to_idml import AUTO_RUN_MACOS, AUTO_RUN_WINDOWS, run_indesign_windows, run_indesign_macos
from xml_to_idml import LOG_PATH  # 打印日志用
from pipeline_logger import PipelineLogger

PIPELINE_LOGGER: Optional[PipelineLogger] = None

# ====== 跨平台口令文件路径 ======
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

# ====== 内置（编译时）默认密码（不可被删除） ======
# 默认值：Moyi#2025!Docx2Idml  （可用 --set-password 修改为新口令，写入文件）
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


# ====== 密码工具函数 ======
def _pbkdf2_hash(password: str, salt: bytes, iterations: int = PBKDF2_ITER) -> bytes:
    return hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, iterations, dklen=32)


def _load_record():
    # 优先读外部文件；失败则回退到内置加密密码
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


def _prompt_hidden(prompt_text: str) -> str:
    try:
        return getpass.getpass(prompt_text)
    except Exception:
        # 某些环境（如无 TTY）可能不支持隐藏输入
        return input(prompt_text)


def _verify_flow(cli_password: str | None) -> bool:
    rec = _load_record()
    # 始终要求输入密码（或 --password 传入），不再“首次运行自动设置密码”
    pwd = cli_password if cli_password is not None else _prompt_hidden("请输入密码：")
    h = _pbkdf2_hash(pwd, rec["salt"], rec["iterations"])
    if h == rec["hash"]:
        return True
    _log_error("密码错误。")
    return False


def _set_password_flow():
    # 修改密码：校验当前密码（可能来自文件或内置记录）
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
    """返回实际生效的 TEMPLATE_PATH（优先命令行，其次 xml_to_idml 全局）"""
    if cli_template and str(cli_template).strip():
        return os.path.abspath(cli_template)
    try:
        return os.path.abspath(X.TEMPLATE_PATH)
    except Exception:
        # 若模块未暴露 TEMPLATE_PATH，可根据你的实现改为固定默认
        return os.path.abspath("template.idml")


def _apply_overrides(cli_template: str | None, cli_out: str | None) -> None:
    """把命令行覆盖写回 xml_to_idml 模块的全局变量，供后续 write_jsx / JSX 模板使用。"""
    if cli_template and str(cli_template).strip():
        X.TEMPLATE_PATH = os.path.abspath(cli_template)
    if cli_out and str(cli_out).strip():
        X.IDML_OUT_PATH = os.path.abspath(cli_out)


def main(argv=None):
    # ====== 命令行解析（新增跨平台与模板/输出覆盖） ======
    parser = argparse.ArgumentParser(description="DOCX -> XML exporter (heading/regex/hybrid) with style switches + Password Protection + Cross-platform paths")
    parser.add_argument("--mode", choices=["heading", "regex", "hybrid"], default="heading", help="Detection mode")
    parser.add_argument("--set-password", action="store_true", help="修改程序密码（先验证当前密码）")
    parser.add_argument("--password", default=None, help="以参数形式传入运行密码（可选，否则将提示输入）")
    parser.add_argument("--template", "-t", dest="template", default=None, help="覆盖 TEMPLATE_PATH（模板 .idml 的路径）")
    parser.add_argument("--out", "-o", dest="out", default=None, help="覆盖 IDML_OUT_PATH（导出的 .idml 路径）")
    parser.add_argument("input", nargs="?", help="Input .docx path")
    parser.add_argument("--no-images", action="store_true", help="生成 XML 时跳过 Word 图片")
    parser.add_argument("--no-tables", action="store_true", help="生成 XML 时跳过 Word 表格")
    parser.add_argument("--no-textboxes", action="store_true", help="生成 XML 时跳过文本框/框架")
    parser.add_argument("--log-dir", help="指定日志目录，默认写入脚本目录的 logs")
    parser.add_argument("--debug-log", action="store_true", help="开启 debug 日志记录")
    args = parser.parse_args(argv)

    # —— 在输入密码之前：先校验 input 和 实际生效的 TEMPLATE_PATH ——
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

    # 若指定了 --out，则提前校验输出目录是否存在（避免后续失败）；若不存在，提示并退出
    if args.out:
        out_abs = os.path.abspath(args.out)
        out_dir = os.path.dirname(out_abs) or os.getcwd()
        if out_dir and not os.path.exists(out_dir):
            _log_error(f"输出目录不存在：{out_dir}")
            sys.exit(2)

    # 覆盖 xml_to_idml 的全局变量（模板路径与导出路径），确保后续 write_jsx 生效
    _apply_overrides(args.template, args.out)

    # 仅执行修改密码功能（可在未校验密码之前执行）
    if args.set_password:
        code = _set_password_flow()
        sys.exit(code if isinstance(code, int) else 0)

    # 其余功能必须先通过密码校验
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
    PIPELINE_LOGGER.describe_paths()
    _log_user(f"[LOG] 用户日志: {PIPELINE_LOGGER.user_log_path}")
    if args.debug_log:
        _log_user(f"[LOG] 调试日志: {PIPELINE_LOGGER.debug_log_path}")

    # ====== 原有流程 ======
    # 1) 生成 XML
    exporter = DOCXOutlineExporter(
        input_path,
        mode=args.mode,
        skip_images=args.no_images,
        skip_tables=args.no_tables,
        skip_textboxes=args.no_textboxes,
    )
    exporter.process(XML_PATH)
    _log_user(f"[OK] mode={args.mode} XML saved -> {XML_PATH}")

    # 2) 解析 XML -> 段落
    paragraphs = extract_paragraphs_with_levels(XML_PATH)
    _log_user(f"[INFO] 解析到 {len(paragraphs)} 段；示例前3段: {paragraphs[:3]}")
    if PIPELINE_LOGGER and args.debug_log:
        PIPELINE_LOGGER.debug(
            f"[DOCX2IDML] skip_flags images={args.no_images} tables={args.no_tables} textboxes={args.no_textboxes}"
        )

    # 3) 生成 JSX（模板 + 每 Level1 新 story + TOC + 脚注/尾注 + i/b/u + 日志 + 脚注/尾注正文样式）
    write_jsx(JSX_PATH, paragraphs)

    # 4) 调 InDesign
    ran = False
    if AUTO_RUN_WINDOWS and sys.platform.startswith("win"):
        ran = run_indesign_windows(JSX_PATH)
    elif AUTO_RUN_MACOS and sys.platform == "darwin":
        ran = run_indesign_macos(JSX_PATH)

    _log_user("\n=== 完成 ===")
    _log_user(f"XML: {XML_PATH}")
    _log_user(f"JSX: {JSX_PATH}")
    _log_user(f"LOG: {LOG_PATH}")

    _log_user(f"IDML: {getattr(X, 'IDML_OUT_PATH', None)}")
    stats = X._relay_jsx_events(
        PIPELINE_LOGGER, LOG_PATH, warn_missing=not ran, cleanup=True
    )
    summary_line = (
        f"[REPORT] JSX 事件统计 info={stats.get('info', 0)} "
        f"warn={stats.get('warn', 0)} error={stats.get('error', 0)} "
        f"debug={stats.get('debug', 0)}"
    )
    _log_user(summary_line)
    if ran:
        _log_user("InDesign 已执行 JSX。若设置 AUTO_EXPORT_IDML=True，将在脚本目录生成 output.idml。")


if __name__ == "__main__":
    main()
