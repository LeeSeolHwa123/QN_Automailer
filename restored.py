"""Central main logic that orchestrates the existing QN automailer modules.

This module purposefully keeps `qn_full_pipeline`, `QN_Automailer`, and the SAP
export helpers untouched. It only coordinates how values flow between them so
that both the GUI and any CLI/automation entrypoints can share the exact same
behaviour without duplicating implementation details.
"""
from __future__ import annotations

import argparse
import importlib
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Dict, Optional

import QN_Automailer as automailer
import qn_full_pipeline as pipeline
from settings_manager import (
    DEFAULT_SETTINGS_FILE,
    load_settings,
    save_settings,
    validate_settings,
)

ProgressCallback = Callable[[int], None]
MessageCallback = Callable[[int, str], None]


@dataclass(slots=True)
class CallbackBridge:
    """Fan-out layer that mirrors GUI callbacks while keeping console logs."""

    user_progress: Optional[ProgressCallback] = None
    user_message: Optional[MessageCallback] = None

    def handle_progress(self, step_idx: int) -> None:
        if self.user_progress:
            try:
                self.user_progress(step_idx)
            except Exception as exc:
                print(f"[WARN] progress callback failed: {exc}")
        else:
            print(f"[STEP {step_idx}]")

    def handle_message(self, step_idx: int, message: Optional[str]) -> None:
        text = message or pipeline.STEP_MESSAGES.get(step_idx, "")
        print(f"[STEP {step_idx}] {text}")
        if self.user_message:
            try:
                self.user_message(step_idx, text)
            except Exception as exc:
                print(f"[WARN] message callback failed: {exc}")


@dataclass
class AutomailerMain:
    """Main logic that wires settings, callbacks, and the pipeline together."""

    settings_path: Path = Path(DEFAULT_SETTINGS_FILE)
    progress_callback: Optional[ProgressCallback] = None
    message_callback: Optional[MessageCallback] = None

    def __post_init__(self) -> None:
        self._bridge = CallbackBridge(self.progress_callback, self.message_callback)

    def _load_settings(
        self,
        overrides: Optional[Dict[str, Optional[str]]] = None,
        persist: bool = False,
    ) -> Dict[str, object]:
        settings = load_settings(str(self.settings_path))
        if overrides:
            for key, value in overrides.items():
                if value is not None:
                    settings[key] = value

        is_valid, error_msg = validate_settings(settings)
        if not is_valid:
            raise ValueError(f"Invalid settings provided\n{error_msg}")

        if persist:
            save_settings(settings, str(self.settings_path))

        return settings

    def run_full_pipeline(
        self,
        overrides: Optional[Dict[str, Optional[str]]] = None,
        persist: bool = False,
    ) -> Dict[str, object]:
        settings = self._load_settings(overrides, persist)
        automailer.init_settings(
            settings,
            progress_callback=self._bridge.handle_progress,
            message_callback=self._bridge.handle_message,
        )
        automailer.main(settings)
        return settings

    def run_single_export(self, target: str) -> None:
        module_name = {
            "sap-c": "sap_c1ck",
            "sap-q": "sap_q1q9",
            "sap-z": "sap_znot",
        }.get(target)
        if not module_name:
            raise ValueError(f"Unknown export target: {target}")

        module = importlib.import_module(module_name)
        entry = getattr(module, "main", None)
        if not callable(entry):
            raise AttributeError(f"Module {module_name} does not expose main()")
        entry()


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Run the QN automailer pipeline or individual SAP exports."
    )
    parser.add_argument(
        "--settings",
        default=DEFAULT_SETTINGS_FILE,
        help="Path to qn_automailer_settings.json",
    )
    parser.add_argument("--export-dir", dest="export_dir", help="Override export dir")
    parser.add_argument(
        "--mail-recipients",
        dest="mail_recipients",
        help="Override mail recipients (; separated)",
    )
    parser.add_argument("--execution-date", dest="execution_date", help="YYYYMMDD")
    parser.add_argument("--date-start", dest="date_range_start", help="YYYYMMDD")
    parser.add_argument("--date-end", dest="date_range_end", help="YYYYMMDD")
    parser.add_argument(
        "--persist",
        action="store_true",
        help="Persist overrides back into the settings file",
    )

    sub = parser.add_subparsers(dest="command")
    sub.add_parser("run", help="Run the full automailer pipeline (default)")
    sub.add_parser("sap-c", help="Only run the C1-CK SAP export helper")
    sub.add_parser("sap-q", help="Only run the Q1-Q9 SAP export helper")
    sub.add_parser("sap-z", help="Only run the ZNOT SAP export helper")
    return parser


def _collect_overrides(args: argparse.Namespace) -> Dict[str, Optional[str]]:
    return {
        "export_dir": args.export_dir,
        "mail_recipients": args.mail_recipients,
        "execution_date": args.execution_date,
        "date_range_start": args.date_range_start,
        "date_range_end": args.date_range_end,
    }


def main() -> None:
    parser = _build_parser()
    args = parser.parse_args()

    runner = AutomailerMain(Path(args.settings))
    overrides = _collect_overrides(args)

    command = args.command or "run"
    if command == "run":
        runner.run_full_pipeline(overrides, persist=args.persist)
    else:
        runner.run_single_export(command)


if __name__ == "__main__":
    main()
