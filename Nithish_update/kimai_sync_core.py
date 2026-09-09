"""
Shared Kimai timesheet sync logic (no Streamlit).
Used by the UI and by the background worker process.
"""
from __future__ import annotations

import ctypes
import io
import os
import re
import time
from ctypes import wintypes
from datetime import datetime, timedelta
from typing import Any, Callable

import pandas as pd
import requests

PERMISSION_HINT = (
    "Permission denied while reading/writing the Excel file. "
    "Close the workbook in Excel (desktop or browser), wait for OneDrive to finish syncing, then retry."
)

LogFn = Callable[[str], None]


def _log(msg: str, log: LogFn | None = None) -> None:
    line = f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {msg}"
    if log:
        log(line)
    else:
        print(line, flush=True)


def is_http_url(value: str) -> bool:
    v = (value or "").strip().lower()
    return v.startswith("http://") or v.startswith("https://")


def find_onedrive_roots():
    roots = []
    for key in ("OneDrive", "OneDriveCommercial"):
        path = os.environ.get(key)
        if path and os.path.isdir(path):
            roots.append(path)
    home = os.path.expanduser("~")
    try:
        for entry in os.listdir(home):
            if entry.lower().startswith("onedrive"):
                full = os.path.join(home, entry)
                if os.path.isdir(full):
                    roots.append(full)
    except OSError:
        pass
    return list(dict.fromkeys(roots))


def find_local_timesheet():
    names = {"timesheet.xlsx", "timesheet.xls", "timesheet.xlsm"}
    for root in find_onedrive_roots():
        preferred = os.path.join(root, "nithis_doc", "Timesheet.xlsx")
        if os.path.isfile(preferred):
            return preferred
        for dirpath, dirnames, filenames in os.walk(root):
            rel = os.path.relpath(dirpath, root)
            depth = 0 if rel == "." else rel.count(os.sep) + 1
            if depth > 5:
                dirnames.clear()
                continue
            for name in filenames:
                if name.lower() in names:
                    return os.path.join(dirpath, name)
    return None


def _read_file_bytes_shared(path: str) -> bytes:
    GENERIC_READ = 0x80000000
    FILE_SHARE_READ = 0x00000001
    FILE_SHARE_WRITE = 0x00000002
    FILE_SHARE_DELETE = 0x00000004
    OPEN_EXISTING = 3
    FILE_ATTRIBUTE_NORMAL = 0x80

    kernel32 = ctypes.windll.kernel32
    handle = kernel32.CreateFileW(
        path,
        GENERIC_READ,
        FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
        None,
        OPEN_EXISTING,
        FILE_ATTRIBUTE_NORMAL,
        None,
    )
    if handle in (ctypes.c_void_p(-1).value, 0, -1):
        raise PermissionError(f"[Errno 13] Permission denied: '{path}'")

    try:
        chunks = []
        chunk_size = 1024 * 1024
        bytes_read = wintypes.DWORD(0)
        while True:
            chunk = ctypes.create_string_buffer(chunk_size)
            ok = kernel32.ReadFile(handle, chunk, chunk_size, ctypes.byref(bytes_read), None)
            if not ok:
                raise OSError(f"ReadFile failed ({ctypes.GetLastError()}) for '{path}'")
            if bytes_read.value == 0:
                break
            chunks.append(chunk.raw[: bytes_read.value])
        return b"".join(chunks)
    finally:
        kernel32.CloseHandle(handle)


def read_excel_safely(path: str) -> pd.DataFrame:
    try:
        return pd.read_excel(path)
    except PermissionError:
        pass
    except OSError as e:
        if getattr(e, "errno", None) != 13:
            raise

    data = _read_file_bytes_shared(path)
    if not data:
        raise PermissionError(PERMISSION_HINT)
    return pd.read_excel(io.BytesIO(data))


def backup_row_key(df: pd.DataFrame) -> pd.Series:
    date_part = pd.to_datetime(df.get("Date"), errors="coerce").dt.strftime("%Y-%m-%d").fillna("")
    project_part = df.get("Project", pd.Series([""] * len(df))).astype(str).str.strip().str.lower()
    tasks_part = df.get("Tasks", pd.Series([""] * len(df))).astype(str).str.strip().str.lower()
    hours_part = (
        pd.to_numeric(df.get("Hours Spent", 0), errors="coerce").fillna(0).astype(float).round(2).astype(str)
    )
    return date_part + "|" + project_part + "|" + tasks_part + "|" + hours_part


def append_to_backup(existing_df: pd.DataFrame | None, new_df: pd.DataFrame) -> tuple[pd.DataFrame, int]:
    if new_df is None or new_df.empty:
        return (existing_df.copy() if existing_df is not None else pd.DataFrame()), 0

    incoming = new_df.copy()
    if existing_df is None or existing_df.empty:
        return incoming.reset_index(drop=True), len(incoming)

    existing = existing_df.copy()
    existing_keys = set(backup_row_key(existing).tolist())
    incoming["_bk"] = backup_row_key(incoming)
    only_new = incoming.loc[~incoming["_bk"].isin(existing_keys)].drop(columns=["_bk"])
    if only_new.empty:
        return existing.reset_index(drop=True), 0

    combined = pd.concat([existing, only_new], ignore_index=True)
    return combined, len(only_new)


def write_excel_safely(df: pd.DataFrame, path: str) -> None:
    path = os.path.abspath(path)
    parent = os.path.dirname(path) or "."
    os.makedirs(parent, exist_ok=True)
    tmp_path = os.path.join(parent, f".~{os.path.basename(path)}.tmp.xlsx")

    last_err = None
    try:
        df.to_excel(tmp_path, index=False)
        for _ in range(5):
            try:
                os.replace(tmp_path, path)
                return
            except PermissionError as e:
                last_err = e
                time.sleep(0.4)
        try:
            df.to_excel(path, index=False)
            return
        except PermissionError as e:
            last_err = e
    finally:
        if os.path.isfile(tmp_path):
            try:
                os.remove(tmp_path)
            except OSError:
                pass

    raise PermissionError(
        f"Could not write `{path}` (file locked). Close Excel/OneDrive lock and retry."
    ) from last_err


def resolve_excel_path(path_or_url: str, label: str = "Excel"):
    path_or_url = (path_or_url or "").strip()
    if not path_or_url:
        return None, f"No {label} path provided."

    is_backup = label.lower().startswith("backup")

    if is_http_url(path_or_url):
        local_input = find_local_timesheet()
        if not local_input:
            return None, (
                f"{label}: SharePoint link needs a synced local OneDrive file. "
                "Paste the local .xlsx path instead."
            )
        if is_backup:
            return os.path.join(os.path.dirname(local_input), "Timesheet_Backup.xlsx"), None
        return local_input, None

    if is_backup:
        parent = os.path.dirname(os.path.abspath(path_or_url)) or "."
        if os.path.isdir(parent) or parent == ".":
            return os.path.abspath(path_or_url), None
        return None, f"Backup folder does not exist: `{parent}`"

    if os.path.isfile(path_or_url):
        return os.path.abspath(path_or_url), None
    return None, f"{label} file not found at: `{path_or_url}`"


def fetch_kimai_projects(url, token):
    if not token or not url:
        return []
    try:
        res = requests.get(
            f"{url.rstrip('/')}/projects",
            headers={"Authorization": f"Bearer {token}", "Accept": "application/json"},
            timeout=10,
        )
        return res.json() if res.status_code == 200 else []
    except Exception:
        return []


def fetch_kimai_activities(url, token, project_id=None):
    if not token or not url:
        return []
    try:
        params = {}
        if project_id is not None:
            params["project"] = project_id
        res = requests.get(
            f"{url.rstrip('/')}/activities",
            headers={"Authorization": f"Bearer {token}", "Accept": "application/json"},
            params=params,
            timeout=10,
        )
        return res.json() if res.status_code == 200 else []
    except Exception:
        return []


def normalize_project_name(name: str) -> str:
    s = str(name or "").strip().casefold()
    s = s.replace("\u00a0", " ").replace("&", " and ")
    s = re.sub(r"[^a-z0-9]+", " ", s)
    return " ".join(s.split())


def project_match_key(name: str) -> str:
    """Identity for matching: ignore case, spaces, punctuation (General Operations == generaloperations)."""
    if name is None or (isinstance(name, float) and pd.isna(name)):
        return ""
    s = str(name).casefold().replace("\u00a0", "").replace("&", "and")
    if s.strip() in {"nan", "none", "nat"}:
        return ""
    return re.sub(r"[^a-z0-9]", "", s)


def resolve_project(excel_name: str, projects: list):
    raw = str(excel_name or "").strip()
    if raw.lower() in {"", "nan", "none", "nat"}:
        return None, None
    if not projects:
        return None, None

    excel_key = project_match_key(raw)
    if not excel_key:
        return None, None

    compact_hits = []
    for p in projects:
        pname = p.get("name", "")
        if project_match_key(pname) == excel_key:
            compact_hits.append(p)
    if len(compact_hits) == 1:
        return compact_hits[0]["id"], compact_hits[0]["name"]
    if len(compact_hits) > 1:
        folded = " ".join(raw.casefold().split())
        for p in compact_hits:
            if " ".join(str(p.get("name", "")).casefold().split()) == folded:
                return p["id"], p["name"]
        return compact_hits[0]["id"], compact_hits[0]["name"]

    # Last resort: exactly one Kimai name contains (or is contained by) the Excel name
    candidates = []
    for p in projects:
        pk = project_match_key(p.get("name", ""))
        if not pk:
            continue
        if excel_key in pk or pk in excel_key:
            candidates.append(p)
    if len(candidates) == 1:
        return candidates[0]["id"], candidates[0]["name"]

    return None, None


# Excel Tasks → Kimai project/activity (case-insensitive exact cell match).
SPECIAL_TASK_ROUTING = {
    "leave": {"project": "General Operations", "activity": "Leave"},
    "permission": {"project": "General Operations", "activity": "Permission"},
}
STANDARD_DAY_HOURS = 8.0
BALANCE_PERMISSION_PROJECT = "General Operations"
BALANCE_PERMISSION_ACTIVITY = "Permission"


def work_hours_from_value(value) -> float:
    hours = pd.to_numeric(value, errors="coerce")
    if pd.isna(hours) or float(hours) <= 0:
        return STANDARD_DAY_HOURS
    return float(hours)


def day_permission_balance_hours(total_hours: float) -> float:
    """Hours still needed to reach an 8-hour day (0 if already 8+)."""
    try:
        total = float(total_hours)
    except (TypeError, ValueError):
        return 0.0
    balance = STANDARD_DAY_HOURS - total
    if balance <= 0.01:
        return 0.0
    return round(balance, 2)


def resolve_special_task_routing(task_text: str) -> dict | None:
    """If Tasks is 'leave' or 'Permission', force General Operations + matching activity."""
    key = " ".join(cell_text(task_text).lower().split())
    mapped = SPECIAL_TASK_ROUTING.get(key)
    return dict(mapped) if mapped else None


def _activity_id_by_name(acts: list, name: str):
    target = (name or "").strip().lower()
    if not target:
        return None
    for a in acts:
        if str(a.get("name", "")).strip().lower() == target:
            return a.get("id")
    return None


def resolve_activity_id(
    url,
    token,
    project_id,
    preferred_name: str = "Design",
    *,
    strict: bool = False,
):
    acts = fetch_kimai_activities(url, token, project_id) or []
    if not acts:
        acts = fetch_kimai_activities(url, token, None) or []
    if not acts:
        return None

    prefs = []
    pref = (preferred_name or "").strip()
    if pref:
        prefs.append(pref)
    if not strict and "design" not in [p.lower() for p in prefs]:
        prefs.append("Design")

    for name in prefs:
        found = _activity_id_by_name(acts, name)
        if found is not None:
            return found

    # Leave/Permission may be global activities not listed under the project.
    if strict:
        all_acts = fetch_kimai_activities(url, token, None) or []
        for name in prefs:
            found = _activity_id_by_name(all_acts, name)
            if found is not None:
                return found
        return None

    return acts[0]["id"]


def resolve_column(df: pd.DataFrame, *candidates: str, exclude_substrings: tuple[str, ...] = ("ticket",)):
    """
    Find a column by header name.
    Exact match first, then compact equality.
    Never returns Ticket ID / ticket columns when resolving Tasks (and similar).
    """
    if df is None:
        return None
    exclude = tuple(s.lower() for s in exclude_substrings)

    def is_excluded(header: str) -> bool:
        h = str(header).strip().lower().replace("_", " ")
        return any(ex in h for ex in exclude)

    normalized = {
        str(c).strip().lower().replace("_", " "): c
        for c in df.columns
        if not is_excluded(c)
    }
    # 1) Exact header match
    for cand in candidates:
        key = " ".join(cand.strip().lower().replace("_", " ").split())
        if key in normalized:
            return normalized[key]
    # 2) Compact equality only (not substring — avoids Task↔Ticket)
    for cand in candidates:
        compact = re.sub(r"[^a-z0-9]", "", cand.strip().lower())
        if not compact:
            continue
        for key, orig in normalized.items():
            if re.sub(r"[^a-z0-9]", "", key) == compact:
                return orig
    return None


def cell_text(value) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    text = str(value).strip()
    if text.lower() in {"nan", "none", "nat"}:
        return ""
    return text


def polish_task_with_gemini(raw_task: str, gemini_key: str | None, use_gemini: bool) -> str:
    if not use_gemini or not gemini_key or not raw_task.strip():
        return raw_task
    try:
        from google import genai

        client = genai.Client(api_key=gemini_key.strip())
        prompt = (
            "Convert this developer task into a concise, professional timesheet activity "
            f"(1 sentence, no quotes): {raw_task}"
        )
        resp = client.models.generate_content(model="gemini-2.5-flash", contents=prompt)
        return resp.text.strip() if resp.text else raw_task
    except Exception:
        return raw_task


def parse_hhmm(value: str, field_name: str):
    try:
        return datetime.strptime(value.strip(), "%H:%M").time()
    except ValueError as e:
        raise ValueError(f"Invalid {field_name}. Use HH:MM format, e.g. 09:00") from e


def place_work_intervals(cursor: datetime, duration_hours: float, lunch_start=None, lunch_end=None):
    remaining = timedelta(hours=duration_hours)
    intervals = []

    while remaining > timedelta(0):
        if lunch_start and lunch_end and lunch_start <= cursor < lunch_end:
            cursor = lunch_end

        if lunch_start and lunch_end and cursor < lunch_start:
            available = lunch_start - cursor
            chunk = min(available, remaining)
            end = cursor + chunk
            if chunk > timedelta(0):
                intervals.append((cursor, end))
            remaining -= chunk
            cursor = end
            if remaining > timedelta(0) and cursor >= lunch_start:
                cursor = lunch_end
        else:
            end = cursor + remaining
            intervals.append((cursor, end))
            remaining = timedelta(0)
            cursor = end

    return intervals, cursor


def run_sync(config: dict[str, Any], log: LogFn | None = None) -> dict[str, Any]:
    """
    Run one full sync cycle. Returns a result dict (no Streamlit UI).
    """
    result = {
        "ok": False,
        "uploaded": 0,
        "skipped_weekend": 0,
        "skipped_existing": 0,
        "failed": 0,
        "newly_backed_up": 0,
        "messages": [],
        "errors": [],
        "missing_projects": [],
    }

    def note(msg: str):
        result["messages"].append(msg)
        _log(msg, log)

    def err(msg: str):
        result["errors"].append(msg)
        _log(f"ERROR: {msg}", log)

    kimai_url = (config.get("kimai_url") or "").strip()
    kimai_token = (config.get("kimai_token") or "").strip()
    gemini_key = (config.get("gemini_key") or "").strip()
    input_excel_path = (config.get("input_excel_path") or "").strip()
    backup_excel_path = (config.get("backup_excel_path") or "").strip()
    work_start_time = config.get("work_start_time") or "09:00"
    activity_name = config.get("activity_name") or "Design"
    allow_fallback_project = bool(config.get("allow_fallback_project"))
    fallback_project_id = int(config.get("fallback_project_id") or 1)
    include_lunch_break = bool(config.get("include_lunch_break", True))
    lunch_start_time = config.get("lunch_start_time") or "13:00"
    lunch_hours = float(config.get("lunch_hours") or 1.0)
    use_gemini = bool(config.get("use_gemini"))

    if not kimai_url or not kimai_token:
        err("Missing Kimai Base URL or API Token.")
        return result

    resolved_path, resolve_error = resolve_excel_path(input_excel_path, "Input")
    backup_path, backup_error = resolve_excel_path(backup_excel_path, "Backup")
    if not resolved_path:
        err(resolve_error or f"Input file not found: {input_excel_path}")
        return result
    if not backup_path:
        err(backup_error or f"Invalid backup path: {backup_excel_path}")
        return result
    if os.path.normcase(resolved_path) == os.path.normcase(backup_path):
        err("Input and Backup are the same file. Use a different backup path.")
        return result

    try:
        raw_df = read_excel_safely(resolved_path)
    except Exception as e:
        err(f"Failed to read input Excel: {e}")
        return result

    # Map Excel headers → Kimai fields.
    # Ticket ID is intentionally NOT used for Kimai (description/project/hours).
    col_date = resolve_column(raw_df, "Date")
    col_tasks = resolve_column(raw_df, "Tasks", "Task")  # never Ticket ID
    col_project = resolve_column(raw_df, "Project")
    col_hours = resolve_column(raw_df, "Hours Spent", "Hours", "Hour Spent", "Working Hours")

    if not col_date or not col_tasks:
        err(
            f"Input Excel must have Date and Tasks columns. Found: {list(raw_df.columns)}"
        )
        return result

    note(f"Using Excel column '{col_tasks}' as Kimai Description (Ticket ID ignored for Kimai)")
    note(f"Backup will store full rows (including Ticket ID) to: {backup_path}")

    df_work = raw_df.copy()
    # Standardize working column names while keeping originals for backup
    if col_tasks != "Tasks":
        df_work["Tasks"] = df_work[col_tasks]
    if col_date != "Date":
        df_work["Date"] = df_work[col_date]
    if col_project and col_project != "Project":
        df_work["Project"] = df_work[col_project]
    if col_hours and col_hours != "Hours Spent":
        df_work["Hours Spent"] = df_work[col_hours]
    elif "Hours Spent" not in df_work.columns:
        df_work["Hours Spent"] = 8.0

    df_to_upload = df_work.dropna(subset=["Date", "Tasks"]).copy()
    df_to_upload = df_to_upload[df_to_upload["Date"].astype(str).str.lower() != "total"]
    # Also drop rows where Tasks text is empty after cleanup
    df_to_upload["_task_text"] = df_to_upload["Tasks"].map(cell_text)
    df_to_upload = df_to_upload[df_to_upload["_task_text"] != ""].copy()
    df_to_upload["Date"] = pd.to_datetime(df_to_upload["Date"], errors="coerce")
    df_to_upload = df_to_upload.dropna(subset=["Date"])

    if df_to_upload.empty:
        note("Input file contains no new data to upload.")
        result["ok"] = True
        return result

    # Keep full original input rows (all columns/values) for backup & input rewrite
    input_columns = list(raw_df.columns)
    original_full_rows = raw_df.loc[df_to_upload.index].copy()

    df_to_upload["Hours Spent"] = pd.to_numeric(df_to_upload["Hours Spent"], errors="coerce").fillna(8.0)
    projects_data = fetch_kimai_projects(kimai_url, kimai_token)
    headers = {
        "Authorization": f"Bearer {kimai_token}",
        "Content-Type": "application/json",
        "Accept": "application/json",
    }
    url = f"{kimai_url.rstrip('/')}/timesheets"

    uploaded = skipped_weekend = skipped_existing = failed = 0
    fail_details = []
    skip_details = []
    missing_project_msgs = []
    done_indices = []

    def is_already_in_kimai(response) -> bool:
        if response.status_code != 400:
            return False
        return "already have an entry" in (response.text or "").lower()

    def post_intervals(intervals, proj_id, activity_id, description) -> bool:
        """POST timesheet intervals. Returns True if none of the posts failed."""
        nonlocal uploaded, skipped_existing, failed
        row_failed = False
        for begin_dt, end_dt in intervals:
            payload = {
                "begin": begin_dt.strftime("%Y-%m-%dT%H:%M:%S"),
                "end": end_dt.strftime("%Y-%m-%dT%H:%M:%S"),
                "project": proj_id,
                "activity": activity_id,
                "description": description,
            }
            try:
                res = requests.post(url, headers=headers, json=payload, timeout=10)
                if res.status_code in [200, 201]:
                    uploaded += 1
                elif is_already_in_kimai(res):
                    skipped_existing += 1
                    skip_details.append(
                        f"{begin_dt:%Y-%m-%d %H:%M}–{end_dt:%H:%M} already in Kimai (skipped)"
                    )
                else:
                    failed += 1
                    row_failed = True
                    fail_details.append(
                        f"{begin_dt:%Y-%m-%d %H:%M} → HTTP {res.status_code}: {res.text[:200]}"
                    )
            except Exception as e:
                failed += 1
                row_failed = True
                fail_details.append(f"{begin_dt:%Y-%m-%d %H:%M} → {e}")
        return not row_failed

    try:
        start_clock = parse_hhmm(work_start_time, "work day start time")
        lunch_delta = None
        lunch_clock = None
        if include_lunch_break:
            lunch_clock = parse_hhmm(lunch_start_time, "lunch start time")
            lunch_delta = timedelta(hours=float(lunch_hours))
    except ValueError as e:
        err(str(e))
        return result

    grouped = df_to_upload.sort_values(by="Date").groupby(df_to_upload["Date"].dt.date)

    for day, records in grouped:
        if day.weekday() >= 5:
            skipped_weekend += len(records)
            done_indices.extend(records.index.tolist())
            continue

        current_cursor = datetime.combine(day, start_clock)
        lunch_start = lunch_end = None
        if include_lunch_break and lunch_clock and lunch_delta:
            lunch_start = datetime.combine(day, lunch_clock)
            lunch_end = lunch_start + lunch_delta
        lunch_kw = {
            "lunch_start": lunch_start if include_lunch_break else None,
            "lunch_end": lunch_end if include_lunch_break else None,
        }

        day_success_indices = []
        day_hours = 0.0
        day_had_failure = False

        for idx, row in records.iterrows():
            duration = work_hours_from_value(row.get("Hours Spent", STANDARD_DAY_HOURS))

            # Excel Tasks → Kimai Description (exact text from input file)
            raw_task = cell_text(original_full_rows.at[idx, col_tasks])
            if not raw_task:
                raw_task = cell_text(row.get("_task_text") or row.get("Tasks"))
            if not raw_task:
                failed += 1
                day_had_failure = True
                fail_details.append(f"Empty Tasks/description for date {day} — row skipped")
                continue

            special = resolve_special_task_routing(raw_task)
            row_activity_name = activity_name
            require_exact_activity = False
            if special:
                excel_project = special["project"]
                row_activity_name = special["activity"]
                require_exact_activity = True
                note(
                    f"Tasks '{raw_task}' → Kimai project '{excel_project}', "
                    f"activity '{row_activity_name}'"
                )
            else:
                excel_project = str(row.get("Project", "")).strip()

            proj_id, matched_name = resolve_project(excel_project, projects_data)
            if proj_id is None:
                if special:
                    failed += 1
                    day_had_failure = True
                    msg = (
                        f"Kimai project '{excel_project}' not found "
                        f"(required for Tasks '{raw_task}')."
                    )
                    fail_details.append(msg)
                    if msg not in missing_project_msgs:
                        missing_project_msgs.append(msg)
                    continue
                if allow_fallback_project:
                    proj_id = int(fallback_project_id)
                    matched_name = f"FALLBACK#{proj_id}"
                else:
                    failed += 1
                    day_had_failure = True
                    msg = (
                        f"Project value not found in Kimai: '{excel_project}' "
                        f"(date {pd.to_datetime(row.get('Date')).date() if pd.notna(row.get('Date')) else row.get('Date')})."
                    )
                    fail_details.append(msg)
                    if msg not in missing_project_msgs:
                        missing_project_msgs.append(msg)
                    continue

            activity_id = resolve_activity_id(
                kimai_url,
                kimai_token,
                proj_id,
                preferred_name=row_activity_name,
                strict=require_exact_activity,
            )
            if activity_id is None:
                failed += 1
                day_had_failure = True
                fail_details.append(
                    f"No activity '{row_activity_name}' found for project '{matched_name}'."
                )
                continue

            intervals, current_cursor = place_work_intervals(
                current_cursor,
                duration,
                **lunch_kw,
            )

            # Leave/Permission keep Excel text; Gemini only for normal tasks.
            final_task = (
                polish_task_with_gemini(raw_task, gemini_key, True)
                if use_gemini and not special
                else raw_task
            )
            # Safety: never allow Gemini/empty to wipe Tasks
            if not cell_text(final_task):
                final_task = raw_task

            note(f"Kimai description ← Tasks: {final_task[:160]}")

            if post_intervals(intervals, proj_id, activity_id, final_task):
                day_success_indices.append(idx)
                day_hours += duration
            else:
                day_had_failure = True

        # Short weekday: remaining hours → General Operations / Permission
        if not day_had_failure:
            balance = day_permission_balance_hours(day_hours)
            if balance > 0:
                perm_proj_id, perm_proj_name = resolve_project(
                    BALANCE_PERMISSION_PROJECT, projects_data
                )
                if perm_proj_id is None:
                    failed += 1
                    day_had_failure = True
                    msg = (
                        f"Kimai project '{BALANCE_PERMISSION_PROJECT}' not found "
                        f"(needed to fill {balance}h Permission on {day})."
                    )
                    fail_details.append(msg)
                    if msg not in missing_project_msgs:
                        missing_project_msgs.append(msg)
                else:
                    perm_act_id = resolve_activity_id(
                        kimai_url,
                        kimai_token,
                        perm_proj_id,
                        preferred_name=BALANCE_PERMISSION_ACTIVITY,
                        strict=True,
                    )
                    if perm_act_id is None:
                        failed += 1
                        day_had_failure = True
                        fail_details.append(
                            f"No activity '{BALANCE_PERMISSION_ACTIVITY}' found for "
                            f"project '{perm_proj_name}' (needed to fill {balance}h on {day})."
                        )
                    else:
                        perm_intervals, current_cursor = place_work_intervals(
                            current_cursor,
                            balance,
                            **lunch_kw,
                        )
                        note(
                            f"{day}: {day_hours:g}h in Excel, adding {balance:g}h "
                            f"as {BALANCE_PERMISSION_PROJECT} / {BALANCE_PERMISSION_ACTIVITY}"
                        )
                        if not post_intervals(
                            perm_intervals,
                            perm_proj_id,
                            perm_act_id,
                            BALANCE_PERMISSION_ACTIVITY,
                        ):
                            day_had_failure = True

        if not day_had_failure:
            done_indices.extend(day_success_indices)

    # Backup must store FULL input-file data (Email ID, Team, Ticket ID, etc.)
    df_done_full = original_full_rows.loc[original_full_rows.index.isin(done_indices)].copy()
    df_done_full = df_done_full.reindex(columns=input_columns)
    df_remain_full = original_full_rows.loc[~original_full_rows.index.isin(done_indices)].copy()
    df_remain_full = df_remain_full.reindex(columns=input_columns)

    result["uploaded"] = uploaded
    result["skipped_weekend"] = skipped_weekend
    result["skipped_existing"] = skipped_existing
    result["failed"] = failed
    result["missing_projects"] = missing_project_msgs
    result["errors"].extend(fail_details[:20])
    result["messages"].extend(skip_details[:20])

    if failed > 0 and uploaded == 0 and skipped_existing == 0:
        err(
            f"Kimai update failed — 0 saved, {failed} failed, "
            f"{skipped_existing} already existed, {skipped_weekend} weekend skipped."
        )
        return result

    note(
        f"Kimai cycle: {uploaded} new | {skipped_existing} already exists | "
        f"{skipped_weekend} weekend | {failed} failed"
    )

    newly_appended = 0
    if not df_done_full.empty:
        try:
            existing_backup_df = read_excel_safely(backup_path) if os.path.exists(backup_path) else None
            # Align existing backup to full input columns so nothing is dropped
            if existing_backup_df is not None and not existing_backup_df.empty:
                for col in input_columns:
                    if col not in existing_backup_df.columns:
                        existing_backup_df[col] = None
                existing_backup_df = existing_backup_df.reindex(
                    columns=list(dict.fromkeys(list(existing_backup_df.columns) + input_columns))
                )
            combined_backup_df, newly_appended = append_to_backup(existing_backup_df, df_done_full)
            # Final backup always uses input-file column order first
            ordered_cols = input_columns + [c for c in combined_backup_df.columns if c not in input_columns]
            combined_backup_df = combined_backup_df.reindex(columns=ordered_cols)
            write_excel_safely(combined_backup_df, backup_path)
            note(
                f"Backup join (full input columns): +{newly_appended} new row(s), "
                f"total {len(combined_backup_df)} → {backup_path}"
            )
        except Exception as e:
            err(f"Failed to update backup file: {e}")
            return result

    try:
        if df_remain_full.empty:
            write_excel_safely(pd.DataFrame(columns=input_columns), resolved_path)
            note(f"Input cleared: {resolved_path}")
        else:
            write_excel_safely(df_remain_full, resolved_path)
            note(
                f"Input kept {len(df_remain_full)} failed row(s); "
                f"removed {len(df_done_full)} handled."
            )
    except Exception as e:
        err(f"Failed to update input file: {e}")
        return result

    result["newly_backed_up"] = newly_appended
    result["ok"] = True
    return result
