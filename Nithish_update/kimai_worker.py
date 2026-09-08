"""
Background Kimai sync worker.
Keeps running after the browser tab closes. Stop only via End button (or stop file).
"""
from __future__ import annotations

import json
import os
import sys
import time
from datetime import datetime

from kimai_sync_core import run_sync

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(BASE_DIR, "_kimai_runner_config.json")
STATE_PATH = os.path.join(BASE_DIR, "_kimai_runner_state.json")
STOP_PATH = os.path.join(BASE_DIR, "_kimai_runner_stop")
LOG_PATH = os.path.join(BASE_DIR, "_kimai_runner.log")


def append_log(line: str) -> None:
    with open(LOG_PATH, "a", encoding="utf-8") as f:
        f.write(line.rstrip() + "\n")


def write_state(data: dict) -> None:
    data = dict(data)
    data["updated_at"] = datetime.now().isoformat(timespec="seconds")
    with open(STATE_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)


def load_config() -> dict:
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


def should_stop() -> bool:
    return os.path.isfile(STOP_PATH)


def clear_stop_flag() -> None:
    if os.path.isfile(STOP_PATH):
        try:
            os.remove(STOP_PATH)
        except OSError:
            pass


def main() -> int:
    if not os.path.isfile(CONFIG_PATH):
        append_log(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] Missing config: {CONFIG_PATH}")
        return 1

    config = load_config()
    interval_seconds = int(config.get("interval_seconds") or 600)
    clear_stop_flag()

    write_state(
        {
            "status": "running",
            "pid": os.getpid(),
            "started_at": datetime.now().isoformat(timespec="seconds"),
            "last_result": None,
            "message": "Worker started",
        }
    )
    append_log(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] Worker PID {os.getpid()} started")

    cycle = 0
    while True:
        if should_stop():
            break

        cycle += 1
        write_state(
            {
                "status": "running",
                "pid": os.getpid(),
                "cycle": cycle,
                "message": f"Running sync cycle #{cycle}",
            }
        )
        append_log(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] === Cycle #{cycle} start ===")

        try:
            result = run_sync(config, log=append_log)
        except Exception as e:
            result = {"ok": False, "errors": [str(e)], "messages": []}
            append_log(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] FATAL: {e}")

        write_state(
            {
                "status": "running",
                "pid": os.getpid(),
                "cycle": cycle,
                "last_result": {
                    "ok": result.get("ok"),
                    "uploaded": result.get("uploaded"),
                    "skipped_existing": result.get("skipped_existing"),
                    "skipped_weekend": result.get("skipped_weekend"),
                    "failed": result.get("failed"),
                    "newly_backed_up": result.get("newly_backed_up"),
                    "errors": (result.get("errors") or [])[:5],
                    "missing_projects": result.get("missing_projects") or [],
                },
                "message": (
                    f"Cycle #{cycle} done — "
                    f"uploaded={result.get('uploaded', 0)}, "
                    f"failed={result.get('failed', 0)}"
                ),
                "next_run_at": datetime.fromtimestamp(
                    time.time() + interval_seconds
                ).isoformat(timespec="seconds"),
            }
        )

        # Sleep in small chunks so End can stop quickly
        slept = 0
        while slept < interval_seconds:
            if should_stop():
                break
            time.sleep(min(2, interval_seconds - slept))
            slept += 2

        if should_stop():
            break

        # Reload config each cycle so UI changes apply without restart (optional)
        try:
            config = load_config()
            interval_seconds = int(config.get("interval_seconds") or interval_seconds)
        except Exception:
            pass

    clear_stop_flag()
    write_state(
        {
            "status": "stopped",
            "pid": None,
            "stopped_at": datetime.now().isoformat(timespec="seconds"),
            "message": "Worker stopped by End button",
        }
    )
    append_log(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] Worker stopped")
    return 0


if __name__ == "__main__":
    sys.exit(main())
