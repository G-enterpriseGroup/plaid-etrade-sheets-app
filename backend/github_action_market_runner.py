from __future__ import annotations

import argparse
import json
import os
from pathlib import Path

from market_analysis_engine import build_report


def main() -> None:
    parser = argparse.ArgumentParser(description="Run Portfolio Link market analysis inside GitHub Actions.")
    parser.add_argument("--request-path", required=True, help="Repo path to runtime request JSON")
    args = parser.parse_args()

    request_path = Path(args.request_path)
    if not request_path.exists():
        raise FileNotFoundError(f"Request file not found: {request_path}")

    request = json.loads(request_path.read_text(encoding="utf-8"))
    request_id = str(request.get("request_id") or request_path.stem)
    headers = request.get("headers") or []
    rows = request.get("rows") or []

    report = build_report(rows, headers=headers)
    report["request_id"] = request_id
    report["request_path"] = str(request_path)
    report["engine"] = "GitHub Actions Python market_analysis_engine.py"

    out_dir = Path("runtime/market-outputs")
    out_dir.mkdir(parents=True, exist_ok=True)

    out_path = out_dir / f"report_{request_id}.json"
    latest_path = out_dir / "latest-market-report.json"

    text = json.dumps(report, indent=2, ensure_ascii=False)
    out_path.write_text(text, encoding="utf-8")
    latest_path.write_text(text, encoding="utf-8")

    print(f"Wrote {out_path}")
    print(f"Wrote {latest_path}")


if __name__ == "__main__":
    main()
