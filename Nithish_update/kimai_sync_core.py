"""
Shared Kimai timesheet sync logic (no Streamlit).
Used by the UI and by the background worker process.
"""
from __future__ import annotations

import base64
import ctypes
import hashlib
import io
import json
import os
import re
import time
from ctypes import wintypes
from datetime import datetime, timedelta
from typing import Any, Callable
from urllib.parse import parse_qs, unquote, urlencode, urlparse, urlunparse

import pandas as pd
import requests

PERMISSION_HINT = (
    "Permission denied while reading/writing the Excel file. "
    "Close the workbook in Excel (desktop or browser), wait for OneDrive to finish syncing, then retry."
)

LogFn = Callable[[str], None]

_CORE_DIR = os.path.dirname(os.path.abspath(__file__))
_SHARE_CACHE_DIR = os.path.join(_CORE_DIR, "_kimai_share_cache")
_SHARE_MAP_PATH = os.path.join(_CORE_DIR, "_kimai_share_map.json")
_SHARE_CACHE_TTL_SECONDS = 45
_BROWSER_UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36"
)
_EXCEL_NAME_RE = re.compile(r"\.(xlsx|xlsm|xls)$", re.I)


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


def clean_path_or_url(value: str) -> str:
    return (value or "").strip().strip('"').strip("'").strip()


def is_excel_filename(name: str | None) -> bool:
    return bool(name) and bool(_EXCEL_NAME_RE.search(name or ""))


def _safe_filename(name: str) -> str:
    name = os.path.basename(name or "").strip() or "workbook.xlsx"
    name = re.sub(r'[<>:"/\\|?*]+', "_", name)
    if not is_excel_filename(name):
        name = f"{name}.xlsx"
    return name


def _with_download_param(url: str) -> str:
    parsed = urlparse(url)
    qs = parse_qs(parsed.query, keep_blank_values=True)
    qs.pop("web", None)
    qs["download"] = ["1"]
    return urlunparse(parsed._replace(query=urlencode(qs, doseq=True)))


def _encode_sharing_url(url: str) -> str:
    raw = base64.b64encode(url.encode("utf-8")).decode("ascii")
    return "u!" + raw.rstrip("=").replace("/", "_").replace("+", "-")


def parse_sharepoint_file_from_url(url: str) -> tuple[str | None, str | None]:
    """Return (OneDrive-relative path, filename) if the URL contains them."""
    if not url:
        return None, None
    parsed = urlparse(url.strip())
    path = unquote(parsed.path or "")
    rel = None
    fname = None

    m = re.search(r"/Documents/(.+\.(?:xlsx|xlsm|xls))$", path, re.I)
    if m:
        rel = m.group(1).replace("/", os.sep)
        fname = os.path.basename(rel)
    else:
        m = re.search(r"/Shared Documents/(.+\.(?:xlsx|xlsm|xls))$", path, re.I)
        if m:
            rel = unquote(m.group(1)).replace("/", os.sep)
            fname = os.path.basename(rel)
        elif is_excel_filename(os.path.basename(path)):
            fname = os.path.basename(path)

    qs = parse_qs(parsed.query)
    for key in ("file", "fileName", "filename"):
        values = qs.get(key) or []
        if values and is_excel_filename(unquote(values[0])):
            fname = fname or unquote(values[0])
            break
    return rel, fname


def share_download_candidates(url: str) -> list[str]:
    url = url.strip()
    parsed = urlparse(url)
    path = unquote(parsed.path or "")
    host = parsed.netloc
    e = (parse_qs(parsed.query).get("e") or [""])[0]
    out = [url, _with_download_param(url)]

    m = re.match(r"/:[a-z]:/g/personal/([^/]+)/([^/]+)/?$", path, re.I)
    if m:
        user, share = m.group(1), m.group(2)
        out.append(f"https://{host}/personal/{user}/_layouts/15/download.aspx?share={share}")
        if e:
            out.append(
                f"https://{host}/personal/{user}/_layouts/15/download.aspx?share={share}&e={e}"
            )
            out.append(
                f"https://{host}/_layouts/15/guestaccess.aspx?share={share}&e={e}&download=1"
            )

    m = re.match(r"/:[a-z]:/s/([^/]+)/([^/]+)/?$", path, re.I)
    if m:
        site, share = m.group(1), m.group(2)
        out.append(
            f"https://{host}/sites/{site}/_layouts/15/guestaccess.aspx?share={share}&download=1"
        )
        if e:
            out.append(
                f"https://{host}/sites/{site}/_layouts/15/guestaccess.aspx?share={share}&e={e}&download=1"
            )

    return list(dict.fromkeys(out))


def _looks_like_excel_bytes(data: bytes) -> bool:
    if not data or len(data) < 4:
        return False
    if data[:2] == b"PK":
        return True
    return data[:8] == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"


def _filename_from_disposition(header: str | None) -> str | None:
    if not header:
        return None
    m = re.search(r"filename\*=(?:UTF-8'')([^;]+)", header, re.I)
    if m:
        name = unquote(m.group(1).strip().strip('"'))
        return name if is_excel_filename(name) else None
    m = re.search(r'filename="?([^";]+)"?', header, re.I)
    if m:
        name = unquote(m.group(1).strip())
        return name if is_excel_filename(name) else None
    return None


def _share_cache_dir(url: str) -> str:
    digest = hashlib.sha256(url.strip().encode("utf-8")).hexdigest()[:16]
    return os.path.join(_SHARE_CACHE_DIR, digest)


def _existing_cache_file(url: str) -> str | None:
    folder = _share_cache_dir(url)
    if not os.path.isdir(folder):
        return None
    for name in os.listdir(folder):
        path = os.path.join(folder, name)
        if os.path.isfile(path) and is_excel_filename(name):
            return path
    return None


def _cache_is_fresh(path: str) -> bool:
    try:
        age = time.time() - os.path.getmtime(path)
    except OSError:
        return False
    return age <= _SHARE_CACHE_TTL_SECONDS


def _load_share_map() -> dict:
    if not os.path.isfile(_SHARE_MAP_PATH):
        return {}
    try:
        with open(_SHARE_MAP_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def _remember_share_mapping(url: str, local_path: str) -> None:
    try:
        mapping = _load_share_map()
        mapping[url.strip()] = os.path.abspath(local_path)
        with open(_SHARE_MAP_PATH, "w", encoding="utf-8") as f:
            json.dump(mapping, f, indent=2)
    except Exception:
        pass


def _mapped_local_path(url: str, allow_missing: bool = False) -> str | None:
    path = _load_share_map().get(url.strip())
    if not path:
        return None
    if os.path.isfile(path):
        return path
    if allow_missing and os.path.isdir(os.path.dirname(path) or "."):
        return path
    return None


def fetch_share_drive_item(url: str) -> dict:
    encoded = _encode_sharing_url(url)
    host = urlparse(url).netloc
    endpoints = [
        f"https://{host}/_api/v2.0/shares/{encoded}/driveItem",
    ]
    headers = {"Accept": "application/json", "User-Agent": _BROWSER_UA}
    for api in endpoints:
        try:
            res = requests.get(api, headers=headers, timeout=8)
        except Exception:
            continue
        if res.status_code != 200:
            continue
        try:
            data = res.json()
        except Exception:
            continue
        item = data.get("driveItem") or data
        if isinstance(item, dict) and (item.get("name") or item.get("webUrl")):
            return item
    return {}


def download_share_excel(url: str) -> tuple[bytes | None, str | None, str | None]:
    headers = {"User-Agent": _BROWSER_UA, "Accept": "*/*"}
    for candidate in share_download_candidates(url):
        try:
            res = requests.get(candidate, headers=headers, allow_redirects=True, timeout=20)
        except Exception:
            continue
        data = res.content or b""
        if res.status_code == 200 and _looks_like_excel_bytes(data):
            name = _filename_from_disposition(res.headers.get("Content-Disposition"))
            if not name:
                _, name = parse_sharepoint_file_from_url(res.url)
            return data, name or "workbook.xlsx", res.url
        _rel, final_name = parse_sharepoint_file_from_url(res.url)
        if _rel or final_name:
            return None, final_name, res.url
    return None, None, None


def find_local_excel_named(filename: str, hint_rel: str | None = None) -> str | None:
    filename_l = (filename or "").strip().lower()
    if not filename_l:
        return None
    hint_folder = os.path.dirname(hint_rel).replace("/", os.sep) if hint_rel else ""
    hits: list[str] = []
    for root in find_onedrive_roots():
        if hint_rel:
            preferred = os.path.join(root, hint_rel)
            if os.path.isfile(preferred):
                return preferred
        if hint_folder:
            preferred_dir = os.path.join(root, hint_folder, os.path.basename(filename))
            if os.path.isfile(preferred_dir):
                return preferred_dir
        for dirpath, dirnames, filenames in os.walk(root):
            rel = os.path.relpath(dirpath, root)
            depth = 0 if rel == "." else rel.count(os.sep) + 1
            if depth > 6:
                dirnames.clear()
                continue
            for name in filenames:
                if name.lower() == filename_l:
                    hits.append(os.path.join(dirpath, name))
    if hint_folder:
        folder_l = hint_folder.lower()
        hinted = [h for h in hits if folder_l in h.lower()]
        if hinted:
            return hinted[0]
    return hits[0] if hits else None


def local_path_from_share_info(rel: str | None, filename: str | None, *, allow_missing: bool):
    if rel:
        for root in find_onedrive_roots():
            candidate = os.path.join(root, rel)
            if os.path.isfile(candidate):
                return candidate
            parent = os.path.dirname(candidate)
            if allow_missing and parent and os.path.isdir(parent):
                return candidate
    if filename:
        found = find_local_excel_named(filename, rel)
        if found:
            return found
        if allow_missing and rel:
            for root in find_onedrive_roots():
                candidate = os.path.join(root, rel)
                parent = os.path.dirname(candidate)
                if parent and os.path.isdir(parent):
                    return candidate
    return None


def _write_share_cache(url: str, data: bytes, filename: str) -> str:
    folder = _share_cache_dir(url)
    os.makedirs(folder, exist_ok=True)
    path = os.path.join(folder, _safe_filename(filename))
    with open(path, "wb") as f:
        f.write(data)
    return os.path.abspath(path)


def resolve_share_or_http_excel(
    url: str, label: str, *, force_refresh: bool = False
) -> tuple[str | None, str | None]:
    is_backup = label.lower().startswith("backup")
    rel, fname = parse_sharepoint_file_from_url(url)
    mapped = _mapped_local_path(url, allow_missing=is_backup)
    local = mapped or local_path_from_share_info(rel, fname, allow_missing=is_backup)

    if local and os.path.isfile(local):
        _remember_share_mapping(url, local)
        return os.path.abspath(local), None
    if is_backup and local:
        parent = os.path.dirname(os.path.abspath(local))
        if os.path.isdir(parent) or parent == ".":
            _remember_share_mapping(url, local)
            return os.path.abspath(local), None

    cached = _existing_cache_file(url)
    if is_backup and cached:
        return cached, None
    if cached and not is_backup and not force_refresh and _cache_is_fresh(cached):
        return cached, None

    item = fetch_share_drive_item(url)
    web_url = str(item.get("webUrl") or "")
    if web_url:
        rel2, fname2 = parse_sharepoint_file_from_url(web_url)
        rel = rel or rel2
        fname = fname or fname2
    name = item.get("name")
    if is_excel_filename(str(name or "")):
        fname = fname or str(name)
        parent = (item.get("parentReference") or {}).get("path") or ""
        if isinstance(parent, str) and ":/" in parent:
            folder = parent.split(":/", 1)[-1].strip("/")
            if folder:
                rel = rel or os.path.join(*folder.split("/"), str(name))

    if not local:
        local = local_path_from_share_info(rel, fname, allow_missing=is_backup)
        if local and os.path.isfile(local):
            _remember_share_mapping(url, local)
            return os.path.abspath(local), None

    data, dl_name, final_url = download_share_excel(url)
    if final_url:
        rel3, fname3 = parse_sharepoint_file_from_url(final_url)
        rel = rel or rel3
        fname = fname or fname3 or dl_name
    elif dl_name:
        fname = fname or dl_name

    if not local:
        local = local_path_from_share_info(rel, fname, allow_missing=is_backup)

    if data and local:
        parent = os.path.dirname(local)
        if parent:
            os.makedirs(parent, exist_ok=True)
        if not os.path.isfile(local):
            with open(local, "wb") as f:
                f.write(data)
        _remember_share_mapping(url, local)
        return os.path.abspath(local), None

    if local and os.path.isfile(local):
        _remember_share_mapping(url, local)
        return os.path.abspath(local), None
    if is_backup and local:
        parent = os.path.dirname(os.path.abspath(local))
        if not os.path.isdir(parent):
            try:
                os.makedirs(parent, exist_ok=True)
            except OSError:
                return None, f"Backup folder does not exist: `{parent}`"
        _remember_share_mapping(url, local)
        return os.path.abspath(local), None

    if data:
        path = _write_share_cache(
            url,
            data,
            fname or dl_name or ("backup.xlsx" if is_backup else "input.xlsx"),
        )
        _remember_share_mapping(url, path)
        return path, None

    if cached:
        return cached, None

    if is_backup:
        folder = _share_cache_dir(url)
        os.makedirs(folder, exist_ok=True)
        path = os.path.join(folder, _safe_filename(fname or "Timesheet_Backup.xlsx"))
        _remember_share_mapping(url, path)
        return os.path.abspath(path), None

    return None, (
        f"{label}: could not open this SharePoint/OneDrive link. "
        "Paste a Copy link to that .xlsx file (anyone-with-the-link, or a file already synced in OneDrive), "
        "or paste the local .xlsx path. Each link is used as its own file."
    )


def resolve_excel_path(
    path_or_url: str, label: str = "Excel", *, force_refresh: bool = False
):
    path_or_url = clean_path_or_url(path_or_url)
    if not path_or_url:
        return None, f"No {label} path provided."

    is_backup = label.lower().startswith("backup")

    if is_http_url(path_or_url):
        return resolve_share_or_http_excel(path_or_url, label, force_refresh=force_refresh)

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


def fetch_kimai_timesheets(url, token, begin: datetime, end: datetime) -> list:
    if not token or not url:
        return []
    headers = {"Authorization": f"Bearer {token}", "Accept": "application/json"}
    endpoint = f"{url.rstrip('/')}/timesheets"
    items: list = []
    page = 1
    while page <= 20:
        try:
            res = requests.get(
                endpoint,
                headers=headers,
                params={
                    "begin": begin.strftime("%Y-%m-%dT%H:%M:%S"),
                    "end": end.strftime("%Y-%m-%dT%H:%M:%S"),
                    "size": 100,
                    "page": page,
                },
                timeout=20,
            )
        except Exception:
            break
        if res.status_code != 200:
            break
        try:
            data = res.json()
        except Exception:
            break
        if not isinstance(data, list) or not data:
            break
        items.extend(data)
        if len(data) < 100:
            break
        page += 1
    return items


def task_match_key(text: str) -> str:
    return " ".join(cell_text(text).casefold().split())


def naive_datetime(value):
    ts = pd.to_datetime(value, errors="coerce")
    if pd.isna(ts):
        return None
    try:
        if getattr(ts, "tzinfo", None) is not None:
            ts = ts.tz_convert(None)
    except (TypeError, ValueError, AttributeError):
        try:
            ts = ts.tz_localize(None)
        except Exception:
            pass
    try:
        return datetime(
            int(ts.year), int(ts.month), int(ts.day),
            int(ts.hour), int(ts.minute), int(ts.second),
        )
    except Exception:
        return None


def timesheet_date_key(begin_val) -> str:
    s = str(begin_val or "")
    if len(s) >= 10 and s[4] == "-" and s[7] == "-":
        return s[:10]
    dt = naive_datetime(begin_val)
    return dt.strftime("%Y-%m-%d") if dt else ""


def index_existing_timesheets(items: list) -> tuple[dict[str, set[str]], dict[str, datetime]]:
    by_day: dict[str, set[str]] = {}
    latest_end: dict[str, datetime] = {}
    for ts in items or []:
        if not isinstance(ts, dict):
            continue
        day_key = timesheet_date_key(ts.get("begin"))
        if not day_key:
            continue
        key = task_match_key(ts.get("description"))
        if key:
            by_day.setdefault(day_key, set()).add(key)
        end_dt = naive_datetime(ts.get("end"))
        if end_dt:
            prev = latest_end.get(day_key)
            if prev is None or end_dt > prev:
                latest_end[day_key] = end_dt
    return by_day, latest_end


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

    resolved_path, resolve_error = resolve_excel_path(
        input_excel_path, "Input", force_refresh=True
    )
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

    existing_by_day: dict[str, set[str]] = {}
    latest_end_by_day: dict[str, datetime] = {}
    try:
        min_day = pd.Timestamp(df_to_upload["Date"].min()).to_pydatetime()
        max_day = pd.Timestamp(df_to_upload["Date"].max()).to_pydatetime()
        fetch_begin = datetime.combine(min_day.date(), datetime.min.time())
        fetch_end = datetime.combine(max_day.date() + timedelta(days=1), datetime.min.time())
        existing_items = fetch_kimai_timesheets(kimai_url, kimai_token, fetch_begin, fetch_end)
        existing_by_day, latest_end_by_day = index_existing_timesheets(existing_items)
        note(f"Loaded {len(existing_items)} existing Kimai timesheet(s) to skip duplicates")
    except Exception as e:
        note(f"Could not list existing Kimai timesheets ({e}); will skip duplicates on POST")

    uploaded = skipped_weekend = skipped_existing = failed = 0
    fail_details = []
    skip_details = []
    missing_project_msgs = []
    done_indices = []

    def is_already_in_kimai(response) -> bool:
        if response.status_code not in (400, 409, 422):
            return False
        text = (response.text or "").lower()
        try:
            body = response.json()
            if isinstance(body, dict):
                text += " " + str(body.get("message") or "").lower()
                text += " " + str(body.get("errors") or "").lower()
        except Exception:
            pass
        needles = (
            "already have an entry",
            "already have a timesheet",
            "already exists",
            "overlapping",
            "overlap",
            "duplicate",
            "period is already",
        )
        return any(n in text for n in needles)

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

        day_key = day.isoformat()
        existing_keys = existing_by_day.setdefault(day_key, set())
        kimai_end = latest_end_by_day.get(day_key)
        used_kimai_cursor = False
        if kimai_end and kimai_end.date() == day and kimai_end > current_cursor:
            current_cursor = kimai_end
            used_kimai_cursor = True
            note(
                f"{day}: existing Kimai work until {kimai_end:%H:%M}; "
                "already-saved Excel rows are skipped and later rows continue after that"
            )

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

            task_key = task_match_key(raw_task)
            if task_key and task_key in existing_keys:
                skipped_existing += 1
                day_success_indices.append(idx)
                day_hours += duration
                skip_details.append(
                    f"{day}: already in Kimai — skipped '{raw_task[:120]}', next row"
                )
                note(f"{day}: Excel row already in Kimai, moving to next row: {raw_task[:120]}")
                if not used_kimai_cursor:
                    _, current_cursor = place_work_intervals(
                        current_cursor,
                        duration,
                        **lunch_kw,
                    )
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
                if task_key:
                    existing_keys.add(task_key)
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

        # Always clear handled rows (new uploads and already-in-Kimai), even if a later row failed
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
