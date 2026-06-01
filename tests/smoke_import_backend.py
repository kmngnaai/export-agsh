from __future__ import annotations

import importlib.util
import sys
from pathlib import Path


BACKEND_PATH = Path(__file__).resolve().parents[1] / "00.Detaisublog_v26.py"
REQUIRED_SYMBOLS = (
    "Processor",
    "RepairOptions",
    "safe_save_workbook_atomic",
)


def main() -> int:
    try:
        if not BACKEND_PATH.is_file():
            raise FileNotFoundError(f"Backend file not found: {BACKEND_PATH}")

        spec = importlib.util.spec_from_file_location("smoke_import_backend_target", BACKEND_PATH)
        if spec is None or spec.loader is None:
            raise RuntimeError(f"Cannot create import spec for: {BACKEND_PATH}")

        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)

        missing = [name for name in REQUIRED_SYMBOLS if not hasattr(module, name)]
        if missing:
            raise RuntimeError(f"Missing required symbols: {', '.join(missing)}")

        print("PASS: backend imported successfully")
        print(f"PASS: required symbols available: {', '.join(REQUIRED_SYMBOLS)}")
        return 0
    except Exception as exc:
        print(f"FAIL: backend smoke import failed: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
