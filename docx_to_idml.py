# -*- coding: utf-8 -*-
"""
docx_to_idml.py (cross-platform Windows + macOS)
- Keeps both embedded and file-based passwords (PBKDF2); verification required at startup; --set-password updates the local credential
- Validates input/TEMPLATE_PATH/out directories before prompting; --template/-t and --out/-o override defaults
- Stores the password file in platform-friendly locations:
  * Windows: %LOCALAPPDATA%\\Docx2IDML\\docx_to_idml_pass.json
  * macOS : ~/Library/Application Support/Docx2IDML/docx_to_idml_pass.json
  * Others: ~/.config/Docx2IDML/docx_to_idml_pass.json
- Still auto-runs InDesign (Windows COM / macOS AppleScript) with language=javascript corrections inside xml_to_idml"""
from __future__ import annotations

import os
import sys
import json
import argparse
import getpass
import base64
import hashlib

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

import xml_to_idml as X
from docx_to_xml_outline_notes_v13 import DOCXOutlineExporter
from xml_to_idml import XML_PATH, extract_paragraphs_with_levels, write_jsx, JSX_PATH
from xml_to_idml import AUTO_RUN_MACOS, AUTO_RUN_WINDOWS, run_indesign_windows, run_indesign_macos
from xml_to_idml import LOG_PATH  # log toggle
# ====== Password file paths ======
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

# ====== Embedded default credential ======
# Default password is Moyi#2025!Docx2Idml; run --set-password to customize
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


# ====== PBKDF2 helpers ======
def _pbkdf2_hash(password: str, salt: bytes, iterations: int = PBKDF2_ITER) -> bytes:
    return hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, iterations, dklen=32)


def _load_record():
    # Fall back to the embedded record when the external file is missing
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


def _prompt_hidden(prompt_text: str) -> str:
    try:
        return getpass.getpass(prompt_text)
    except Exception:
        # Some TTYs cannot hide input; fallback to plain input()
        return input(prompt_text)


def _verify_flow(cli_password: str | None) -> bool:
    rec = _load_record()
    # Prompt interactively when --password is absent
    pwd = cli_password if cli_password is not None else _prompt_hidden("Enter password > ")
    h = _pbkdf2_hash(pwd, rec["salt"], rec["iterations"])
    if h == rec["hash"]:
        return True
    print("")
    return False


def _set_password_flow():
    # Verify the current password before allowing updates
    rec = _load_record()
    old = _prompt_hidden("Current password > ")
    if _pbkdf2_hash(old, rec["salt"], rec["iterations"]) != rec["hash"]:
        print("Current password is incorrect; aborting.")
        return 2
    while True:
        p1 = _prompt_hidden("New password > ")
        p2 = _prompt_hidden("Confirm new password > ")
        if p1 != p2:
            print("Passwords do not match. Try again.")
            continue
        if len(p1) < 6:
            print("Please enter at least 6 characters.")
            continue
        salt = os.urandom(16)
        hsh  = _pbkdf2_hash(p1, salt, PBKDF2_ITER)
        _save_record(salt, hsh, PBKDF2_ITER)
        print("[SECURE] Password updated. Delete the local credential file to restore defaults.")
        return 0


# ====== Template & output overrides ======
def _effective_template_path(cli_template: str | None) -> str:
    """Return the TEMPLATE_PATH used by xml_to_idml"""
    if cli_template and str(cli_template).strip():
        return os.path.abspath(cli_template)
    try:
        return os.path.abspath(X.TEMPLATE_PATH)
    except Exception:
        # Fallback to a fixed template when xml_to_idml lacks TEMPLATE_PATH
        return os.path.abspath("template.idml")


def _apply_overrides(cli_template: str | None, cli_out: str | None) -> None:
    """Push CLI overrides back into xml_to_idml so write_jsx/JSX pick up new paths"""
    if cli_template and str(cli_template).strip():
        X.TEMPLATE_PATH = os.path.abspath(cli_template)
    if cli_out and str(cli_out).strip():
        X.IDML_OUT_PATH = os.path.abspath(cli_out)


def main(argv=None):
    # ====== Cross-platform CLI ======
    parser = argparse.ArgumentParser(description="DOCX -> XML exporter (heading/regex/hybrid) with style switches + Password Protection + Cross-platform paths")
    parser.add_argument("--mode", choices=["heading", "regex", "hybrid"], default="heading", help="Detection mode")
    parser.add_argument("--set-password", action="store_true", help="Update the local password (requires current password)")
    parser.add_argument("--password", default=None, help="Provide the password non-interactively")
    parser.add_argument("--template", "-t", dest="template", default=None, help="Override xml_to_idml.TEMPLATE_PATH (.idml)")
    parser.add_argument("--out", "-o", dest="out", default=None, help="Override xml_to_idml.IDML_OUT_PATH (.idml)")
    parser.add_argument("input", nargs="?", help="Input .docx path")
    parser.add_argument("--no-images", action="store_true", help="skip embedding images when generating JSX")
    args = parser.parse_args(argv)

    # Validate the input path and template before prompting for passwords
    if not args.input:
        print("Missing input .docx path")
        sys.exit(2)
    input_path = os.path.abspath(args.input)
    if not os.path.exists(input_path):
        print(f"File not found: {input_path}")
        sys.exit(2)

    eff_template = _effective_template_path(args.template)
    if not os.path.exists(eff_template):
        print(f"Template file not found: {eff_template}")
        sys.exit(2)

    # When --out is provided, validate the parent directory
    if args.out:
        out_abs = os.path.abspath(args.out)
        out_dir = os.path.dirname(out_abs) or os.getcwd()
        if out_dir and not os.path.exists(out_dir):
            print(f"Directory not found: {out_dir}")
            sys.exit(2)

    # Push overrides back into xml_to_idml so write_jsx uses the new paths
    _apply_overrides(args.template, args.out)

    # Exit immediately when only updating the password
    if args.set_password:
        code = _set_password_flow()
        sys.exit(code if isinstance(code, int) else 0)

    # Otherwise verify the password
    if not _verify_flow(args.password):
        sys.exit(1)

    # ====== Main flow ======
    # 1)  XML
    exporter = DOCXOutlineExporter(input_path, mode=args.mode)
    exporter.process(XML_PATH)
    print(f"[OK] mode={args.mode} XML saved -> {XML_PATH}")

    # 2)  XML -> 
    paragraphs = extract_paragraphs_with_levels(XML_PATH)
    print(f"[INFO] Total {len(paragraphs)} paragraphs, first 3: {paragraphs[:3]}")
    print(f"[DEBUG] skip_images flag = {args.no_images}")

    # 3)  Generate JSX: one story per level-1 heading plus TOC/notes/styles
    write_jsx(JSX_PATH, paragraphs, skip_images=args.no_images)

    # 4)  Optionally call InDesign
    ran = False
    if AUTO_RUN_WINDOWS and sys.platform.startswith("win"):
        ran = run_indesign_windows(JSX_PATH)
    elif AUTO_RUN_MACOS and sys.platform == "darwin":
        ran = run_indesign_macos(JSX_PATH)

    print("\n=== Export Paths ===")
    print("XML: ", XML_PATH)
    print("JSX: ", JSX_PATH)
    print("LOG: ", LOG_PATH)
    # Note: --out only affects the current run
    print("IDML:", getattr(X, "IDML_OUT_PATH", None))
    if ran:
        print("InDesign executed the JSX; AUTO_EXPORT_IDML=True will emit output.idml beside the script")


if __name__ == "__main__":
    main()
