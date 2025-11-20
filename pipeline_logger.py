import logging
from pathlib import Path
from datetime import datetime, timedelta, timezone
from typing import Optional

BEIJING_TZ = timezone(timedelta(hours=8), name="UTC+8")


class PipelineLogger:
    """
    Utility that keeps two log channels:
    - user_log: always-on, records high level progress and report items.
    - debug_log: optional, only created/emitted when enable_debug is True.

    Each run stores logs under {log_root}/{doc_basename}_{timestamp}/
    """

    def __init__(
        self,
        source_path: str,
        log_root: Optional[str] = None,
        enable_debug: bool = False,
        console_echo: bool = True,
    ):
        source = Path(source_path)
        doc_name = source.stem or "document"
        timestamp = datetime.now().strftime("%Y%m%dT%H%M%S")
        self.run_dir = Path(log_root or (source.parent / "logs"))
        self.run_dir.mkdir(parents=True, exist_ok=True)

        self.user_log_path = self.run_dir / f"{doc_name}.user.log"
        self.debug_log_path = self.run_dir / f"{doc_name}.debug.log"
        self.jsx_event_log_path = self.run_dir / f"{doc_name}.jsx-events.log"

        self.enable_debug = enable_debug
        self.console_echo = console_echo

        self._user_logger = self._create_logger(
            f"user.{doc_name}.{timestamp}", self.user_log_path, logging.INFO
        )
        self._debug_logger = None
        if enable_debug:
            self._debug_logger = self._create_logger(
                f"debug.{doc_name}.{timestamp}", self.debug_log_path, logging.DEBUG
            )

    def _create_logger(self, name: str, path: Path, level: int) -> logging.Logger:
        logger = logging.getLogger(name)
        logger.setLevel(level)
        logger.propagate = False
        handler = logging.FileHandler(path, mode="w", encoding="utf-8")
        handler.setFormatter(logging.Formatter("%(message)s"))
        logger.handlers.clear()
        logger.addHandler(handler)
        return logger

    def _timestamp(self) -> str:
        return datetime.now(BEIJING_TZ).strftime("%Y-%m-%d %H:%M:%S")

    def _format_line(self, module: str, level: str, message: str) -> str:
        return f"[{self._timestamp()}][{module}][{level}] {message}"

    def _emit_user(self, module: str, level: str, message: str):
        line = self._format_line(module, level, message)
        self._user_logger.info(line)
        if self.console_echo:
            print(line)

    def _emit_debug(self, module: str, level: str, message: str):
        if not self.enable_debug or not self._debug_logger:
            return
        line = self._format_line(module, level, message)
        self._debug_logger.debug(line)

    def user(self, message: str, module: str = "PY"):
        """High-level info that is always recorded."""
        self._emit_user(module, "INFO", message)

    def warn(self, message: str, module: str = "PY"):
        self._emit_user(module, "WARN", message)

    def error(self, message: str, module: str = "PY"):
        self._emit_user(module, "ERROR", message)

    def debug(self, message: str, module: str = "PY"):
        self._emit_debug(module, "DEBUG", message)

    def relay_jsx_event(self, level: str, message: str):
        """
        Route a JSX event (level in {"debug","info","warn","error"}) to logs.
        """
        level = (level or "debug").upper()
        module = "JSX"
        if level == "DEBUG":
            self.debug(message, module=module)
            return
        self._emit_user(module, level, message)
        if self.enable_debug and self._debug_logger:
            self._emit_debug(module, level, message)

    def summarize(self, heading: str, details: str):
        """Helper to write a section header + details into user log."""
        self.user(f"{heading}: {details}")

    def describe_paths(self):
        self.user(f"user={self.user_log_path}")
        if self.enable_debug:
            self.user(f"debug={self.debug_log_path}")
        self.debug(f"JSX raw log will be written to {self.jsx_event_log_path}")
