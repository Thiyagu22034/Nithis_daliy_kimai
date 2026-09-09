"""
Kimai Timesheet Syncer — Streamlit UI.
Start launches a background worker that keeps running after the browser closes.
End is the only way to stop it.
"""
from __future__ import annotations

import json
import os
import subprocess
import sys
import time
from datetime import datetime

import pandas as pd
import streamlit as st

from kimai_sync_core import (
    fetch_kimai_projects,
    is_http_url,
    read_excel_safely,
    resolve_column,
    resolve_excel_path,
    resolve_project,
    resolve_special_task_routing,
    cell_text,
    work_hours_from_value,
    day_permission_balance_hours,
    BALANCE_PERMISSION_PROJECT,
    BALANCE_PERMISSION_ACTIVITY,
    run_sync,
    PERMISSION_HINT,
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(BASE_DIR, "_kimai_runner_config.json")
STATE_PATH = os.path.join(BASE_DIR, "_kimai_runner_state.json")
STOP_PATH = os.path.join(BASE_DIR, "_kimai_runner_stop")
LOG_PATH = os.path.join(BASE_DIR, "_kimai_runner.log")
WORKER_SCRIPT = os.path.join(BASE_DIR, "kimai_worker.py")

st.set_page_config(page_title="Kimai Timesheet Syncer & Backup", layout="wide")
st.title("Kimai Timesheet Syncer with Backup & Clear")
st.caption(
    "**Start** runs sync in the background (keeps running if you close this tab). "
    "**End** is the only way to stop it."
)


def load_state() -> dict:
    if not os.path.isfile(STATE_PATH):
        return {"status": "stopped"}
    try:
        with open(STATE_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {"status": "stopped"}


def is_pid_running(pid: int | None) -> bool:
    if not pid:
        return False
    try:
        # Windows-friendly process check
        out = subprocess.run(
            ["tasklist", "/FI", f"PID eq {pid}"],
            capture_output=True,
            text=True,
            timeout=10,
        )
        return str(pid) in (out.stdout or "")
    except Exception:
        return False


def is_worker_running() -> tuple[bool, dict]:
    state = load_state()
    pid = state.get("pid")
    if state.get("status") == "running" and is_pid_running(pid):
        return True, state
    if state.get("status") == "running" and not is_pid_running(pid):
        # Stale state
        state = {
            "status": "stopped",
            "pid": None,
            "message": "Worker is not running (process ended).",
            "updated_at": datetime.now().isoformat(timespec="seconds"),
        }
        try:
            with open(STATE_PATH, "w", encoding="utf-8") as f:
                json.dump(state, f, indent=2)
        except Exception:
            pass
        return False, state
    return False, state


def save_config(config: dict) -> None:
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(config, f, indent=2)


def start_worker(config: dict) -> tuple[bool, str]:
    running, state = is_worker_running()
    if running:
        return False, f"Already running (PID {state.get('pid')})."

    save_config(config)
    if os.path.isfile(STOP_PATH):
        try:
            os.remove(STOP_PATH)
        except OSError:
            pass

    creationflags = 0
    if os.name == "nt":
        # Detach so closing the browser/tab does not stop the worker
        creationflags = (
            subprocess.DETACHED_PROCESS
            | subprocess.CREATE_NEW_PROCESS_GROUP
            | getattr(subprocess, "CREATE_NO_WINDOW", 0x08000000)
        )

    log_f = open(LOG_PATH, "a", encoding="utf-8")
    try:
        proc = subprocess.Popen(
            [sys.executable, WORKER_SCRIPT],
            cwd=BASE_DIR,
            stdout=log_f,
            stderr=subprocess.STDOUT,
            stdin=subprocess.DEVNULL,
            creationflags=creationflags,
            close_fds=True,
        )
    except Exception as e:
        log_f.close()
        return False, f"Failed to start worker: {e}"

    # Give worker a moment to write state
    time.sleep(1.2)
    running, state = is_worker_running()
    if running:
        return True, f"Started background worker (PID {state.get('pid') or proc.pid})."
    # Fallback: process may still be starting
    if is_pid_running(proc.pid):
        return True, f"Started background worker (PID {proc.pid})."
    return False, "Worker did not stay running. Check `_kimai_runner.log`."


def stop_worker() -> tuple[bool, str]:
    running, state = is_worker_running()
    pid = state.get("pid")

    # Signal graceful stop
    try:
        with open(STOP_PATH, "w", encoding="utf-8") as f:
            f.write("stop")
    except Exception as e:
        return False, f"Could not write stop flag: {e}"

    # Wait briefly for graceful exit
    for _ in range(15):
        time.sleep(0.4)
        if not is_pid_running(pid):
            break

    if is_pid_running(pid):
        try:
            subprocess.run(["taskkill", "/PID", str(pid), "/F"], capture_output=True, timeout=10)
        except Exception:
            pass

    # Final state
    stopped_state = {
        "status": "stopped",
        "pid": None,
        "stopped_at": datetime.now().isoformat(timespec="seconds"),
        "message": "Stopped by End button",
    }
    try:
        with open(STATE_PATH, "w", encoding="utf-8") as f:
            json.dump(stopped_state, f, indent=2)
    except Exception:
        pass

    if os.path.isfile(STOP_PATH):
        try:
            os.remove(STOP_PATH)
        except OSError:
            pass

    return True, "Worker stopped."


# --- Status banner ---
running, runner_state = is_worker_running()
if running:
    st.success(
        f"Background sync is **RUNNING** (PID {runner_state.get('pid')}). "
        f"{runner_state.get('message', '')} "
        "Closing this browser tab will NOT stop it."
    )
else:
    st.info("Background sync is **STOPPED**. Fill settings below, then click **Start**.")

# --- 1. API CONFIGURATIONS ---
st.header("1. API Configuration")
col_api1, col_api2, col_api3 = st.columns(3)

with col_api1:
    kimai_url = st.text_input(
        "Kimai Base URL",
        value="https://in-timetracking.euroland.com/api",
    )
with col_api2:
    kimai_token = st.text_input(
        "Kimai API Token",
        type="password",
        placeholder="Enter your Kimai API token",
    )
with col_api3:
    gemini_key = st.text_input(
        "Gemini API Key (Optional)",
        type="password",
        placeholder="Enter Gemini key to refine task descriptions",
    )

projects_data = fetch_kimai_projects(kimai_url, kimai_token) if (kimai_url and kimai_token) else []

# --- 2. Excel paths ---
st.header("2. Excel Paths (Input & Backup)")
st.caption(
    "Excel **Tasks** → Kimai **Description**. "
    "**Ticket ID** is not used for Kimai (kept only in Backup). "
    "Full input rows are joined into your **Backup** path after sync. "
    "If Tasks is **leave**, Kimai uses project **General Operations** and activity **Leave**. "
    "If Tasks is **Permission**, Kimai uses project **General Operations** and activity **Permission**. "
    "If a weekday’s **Hours Spent** total is under **8**, the remaining hours are added as "
    "**General Operations** / **Permission**. "
    "Excel **Project** names match the Kimai project list ignoring capital letters and spaces "
    "(e.g. `generaloperations` = `General Operations`)."
)
col_p1, col_p2 = st.columns(2)
with col_p1:
    input_excel_path = st.text_input(
        "Input data Excel path",
        value=r"C:\Users\Thiyagu.Sekar\OneDrive - Eurolandcom AB\nithis_doc\Timesheet.xlsx",
    )
with col_p2:
    backup_excel_path = st.text_input(
        "Backup Excel path",
        value=r"C:\Users\Thiyagu.Sekar\OneDrive - Eurolandcom AB\nithis_doc\Timesheet_Backup.xlsx",
    )

resolved_input_path, resolve_err = resolve_excel_path(input_excel_path, "Input")
resolved_backup_path, backup_err = resolve_excel_path(backup_excel_path, "Backup")

if resolved_input_path and resolved_backup_path:
    if os.path.normcase(resolved_input_path) == os.path.normcase(resolved_backup_path):
        st.error("Input and Backup paths resolve to the same file.")
        resolved_backup_path = None
        backup_err = "Backup path must differ from Input path."

if resolved_backup_path:
    st.caption(f"Backup target: `{resolved_backup_path}`")

if resolved_input_path:
    try:
        raw_df = read_excel_safely(resolved_input_path)
        col_date = resolve_column(raw_df, "Date")
        col_tasks = resolve_column(raw_df, "Tasks", "Task")
        col_project = resolve_column(raw_df, "Project")
        col_hours = resolve_column(raw_df, "Hours Spent", "Hours", "Working Hours")

        if not col_date or not col_tasks:
            st.error(
                f"Input Excel must have **Date** and **Tasks** columns. Found: {list(raw_df.columns)}"
            )
        else:
            valid_df = raw_df.dropna(subset=[col_date, col_tasks]).copy()
            valid_df = valid_df[valid_df[col_date].astype(str).str.lower() != "total"]
            valid_df[col_date] = pd.to_datetime(valid_df[col_date], errors="coerce")
            input_df = valid_df.dropna(subset=[col_date]).copy()
            if col_hours:
                input_df[col_hours] = pd.to_numeric(input_df[col_hours], errors="coerce").fillna(8.0)

            matched_names = []
            matched_ids = []
            matched_activities = []
            missing_project_labels = []
            projects_series = (
                input_df[col_project].tolist() if col_project else [""] * len(input_df)
            )
            tasks_series = input_df[col_tasks].tolist()
            for p, task in zip(projects_series, tasks_series):
                special = resolve_special_task_routing(cell_text(task))
                if special:
                    lookup_name = special["project"]
                    pid, pname = resolve_project(lookup_name, projects_data)
                    matched_activities.append(special["activity"])
                    if pid is None:
                        missing_project_labels.append(lookup_name)
                else:
                    lookup_name = p
                    pid, pname = resolve_project(p, projects_data)
                    matched_activities.append("—")
                    if pid is None and str(p).strip():
                        missing_project_labels.append(str(p).strip())
                matched_ids.append(pid)
                matched_names.append(pname if pname else "⚠ NOT FOUND")

            preview_df = pd.DataFrame(
                {
                    "Date": input_df[col_date],
                    "Project": input_df[col_project] if col_project else "",
                    "Tasks → Kimai Description": input_df[col_tasks],
                    "Hours Spent": input_df[col_hours] if col_hours else 8.0,
                    "Kimai Project": matched_names,
                    "Kimai Activity": matched_activities,
                }
            )
            preview_hours = preview_df["Hours Spent"].map(work_hours_from_value)
            preview_dates = pd.to_datetime(preview_df["Date"], errors="coerce").dt.date
            balance_rows = []
            for day_key, hour_total in preview_hours.groupby(preview_dates).sum().items():
                if pd.isna(day_key):
                    continue
                try:
                    if day_key.weekday() >= 5:
                        continue
                except AttributeError:
                    continue
                balance = day_permission_balance_hours(hour_total)
                if balance <= 0:
                    continue
                pid, pname = resolve_project(BALANCE_PERMISSION_PROJECT, projects_data)
                if pid is None:
                    missing_project_labels.append(BALANCE_PERMISSION_PROJECT)
                day_mask = preview_dates == day_key
                sample_date = preview_df.loc[day_mask, "Date"].iloc[0]
                balance_rows.append(
                    {
                        "Date": sample_date,
                        "Project": "",
                        "Tasks → Kimai Description": f"Permission (auto: +{balance:g}h to reach 8h)",
                        "Hours Spent": balance,
                        "Kimai Project": pname if pname else "⚠ NOT FOUND",
                        "Kimai Activity": BALANCE_PERMISSION_ACTIVITY,
                    }
                )
            if balance_rows:
                preview_df = pd.concat(
                    [preview_df, pd.DataFrame(balance_rows)],
                    ignore_index=True,
                )

            if is_http_url(input_excel_path):
                st.success(f"SharePoint link → `{resolved_input_path}` ({len(input_df)} records).")
            else:
                st.success(f"Input file found: `{resolved_input_path}` — {len(input_df)} records.")
            st.caption(f"Kimai Description is taken from Excel column **{col_tasks}**.")
            st.dataframe(preview_df, height=180)

            missing_projects = sorted(set(missing_project_labels))
            if missing_projects:
                st.error(
                    "Project value not found in Kimai:\n\n"
                    + "\n".join(f"- **{name}**" for name in missing_projects)
                    + "\n\nCreate this project in Kimai, or fix the Excel **Project** name. "
                    "Tasks **leave** / **Permission** always use **General Operations**."
                )
            if projects_data:
                with st.expander("Kimai projects (for matching)"):
                    st.write(
                        pd.DataFrame(
                            [{"ID": p.get("id"), "Name": p.get("name")} for p in projects_data]
                        )
                    )
            elif kimai_token:
                st.warning("No Kimai projects loaded. Check API URL/token.")
    except PermissionError:
        st.error(PERMISSION_HINT)
    except Exception as e:
        st.error(f"Error reading input Excel: {e}")
elif input_excel_path.strip():
    st.warning(resolve_err or f"Waiting for input file at `{input_excel_path}`...")

if backup_err and backup_excel_path.strip():
    st.warning(backup_err)

# --- 3. Settings ---
st.header("3. Execution & Interval Schedule")
col_t1, col_t2, col_t3, col_t4 = st.columns(4)
with col_t1:
    interval_unit = st.selectbox("Interval Unit", ["Minutes", "Hours"])
with col_t2:
    interval_val = st.number_input(
        f"Repeat Interval ({interval_unit.lower()})",
        min_value=1,
        max_value=1440 if interval_unit == "Minutes" else 24,
        value=10 if interval_unit == "Minutes" else 1,
        step=1,
    )
with col_t3:
    work_start_time = st.text_input("Work day start time", value="09:00")
with col_t4:
    activity_name = st.text_input("Kimai Activity name", value="Design")

allow_fallback_project = st.checkbox(
    "Allow fallback project if Excel project name not found",
    value=False,
)
fallback_project_id = st.number_input(
    "Fallback Project ID",
    value=1,
    step=1,
    disabled=not allow_fallback_project,
)

col_b1, col_b2, col_b3 = st.columns(3)
with col_b1:
    include_lunch_break = st.checkbox("Include 1-hour lunch break", value=True)
with col_b2:
    lunch_start_time = st.text_input("Lunch starts at", value="13:00", disabled=not include_lunch_break)
with col_b3:
    lunch_hours = st.number_input(
        "Lunch break (hours)",
        min_value=0.5,
        max_value=2.0,
        value=1.0,
        step=0.5,
        disabled=not include_lunch_break,
    )

use_gemini = st.checkbox("Refine task descriptions with Gemini", value=False)
interval_seconds = interval_val * 60 if interval_unit == "Minutes" else interval_val * 3600

# --- 4. Start / End ---
st.header("4. Start / End (Background Runner)")

config = {
    "kimai_url": kimai_url,
    "kimai_token": kimai_token,
    "gemini_key": gemini_key,
    "input_excel_path": input_excel_path,
    "backup_excel_path": backup_excel_path,
    "work_start_time": work_start_time,
    "activity_name": activity_name,
    "allow_fallback_project": allow_fallback_project,
    "fallback_project_id": fallback_project_id,
    "include_lunch_break": include_lunch_break,
    "lunch_start_time": lunch_start_time,
    "lunch_hours": lunch_hours,
    "use_gemini": use_gemini,
    "interval_seconds": int(interval_seconds),
}

col_start, col_end, col_once = st.columns(3)
with col_start:
    start_clicked = st.button("Start", type="primary", use_container_width=True, disabled=running)
with col_end:
    end_clicked = st.button("End", use_container_width=True, disabled=not running)
with col_once:
    once_clicked = st.button("Run Once Now", use_container_width=True)

if start_clicked:
    if not kimai_url or not kimai_token:
        st.error("Enter Kimai Base URL and API Token before Start.")
    else:
        ok, msg = start_worker(config)
        if ok:
            st.success(msg)
            st.rerun()
        else:
            st.error(msg)

if end_clicked:
    ok, msg = stop_worker()
    if ok:
        st.success(msg)
        st.rerun()
    else:
        st.error(msg)

if once_clicked:
    if not kimai_url or not kimai_token:
        st.error("Enter Kimai Base URL and API Token.")
    else:
        with st.spinner("Running one sync cycle..."):
            result = run_sync(config)
        if result.get("missing_projects"):
            st.error(
                "Project value not found in Kimai:\n\n"
                + "\n".join(f"- {m}" for m in result["missing_projects"])
            )
        if result.get("ok"):
            st.success(
                f"Kimai updated — uploaded={result.get('uploaded', 0)}, "
                f"already={result.get('skipped_existing', 0)}, "
                f"failed={result.get('failed', 0)}, "
                f"backup+={result.get('newly_backed_up', 0)}"
            )
            if result.get("uploaded", 0) > 0:
                st.balloons()
        else:
            st.error("Sync failed.")
        for e in (result.get("errors") or [])[:10]:
            st.code(e)
        for m in (result.get("messages") or [])[:10]:
            st.info(m)

# Live status panel
st.subheader("Runner status")
running2, state2 = is_worker_running()
st.json(
    {
        "status": "RUNNING" if running2 else "STOPPED",
        "pid": state2.get("pid"),
        "cycle": state2.get("cycle"),
        "message": state2.get("message"),
        "next_run_at": state2.get("next_run_at"),
        "last_result": state2.get("last_result"),
        "updated_at": state2.get("updated_at"),
    }
)

if os.path.isfile(LOG_PATH):
    with st.expander("Worker log (last lines)"):
        try:
            with open(LOG_PATH, "r", encoding="utf-8", errors="ignore") as f:
                lines = f.readlines()[-40:]
            st.code("".join(lines) or "(empty)")
        except Exception as e:
            st.warning(str(e))

if running2:
    st.caption("Auto-refresh while running…")
    time.sleep(3)
    st.rerun()
