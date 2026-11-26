# %% [Cell 1] Setup -- imports and SQL connection
import pyodbc
import pandas as pd
from datetime import datetime
import calendar
import re as re
import os as os
import sys
import shutil
import atexit
import subprocess
import traceback
import textwrap
import warnings
import hashlib
import json
from pathlib import Path
from dotenv import load_dotenv

# Force UTF-8 on stdout/stderr: piped/redirected output falls back to the system
# codepage (cp1252), crashing on this script's emoji prints -- including inside the
# crash handler itself. _TeeStream.write() also falls back to replace-errors re-encode
# as a second line of defense. stdin reconfigured too: a piped UTF-8 BOM was being
# mis-decoded into garbage under the wrong codepage, uncatchable by stripping the BOM
# afterward since it was already corrupted by then.
try:
    sys.stdin.reconfigure(encoding="utf-8", errors="replace")
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

# Cosmetic-only: pandas warns on a raw pyodbc connection (wants SQLAlchemy/sqlite3).
# Silenced rather than migrating, since the connection itself works fine.
warnings.filterwarnings(
    "ignore", category=UserWarning,
    message="pandas only supports SQLAlchemy connectable.*"
)

# DB credentials + config/TPA paths live in a .env next to this script, not hardcoded
# (this file gets shared/screen-shared). sys.frozen is True only inside a compiled
# .exe, where __file__ points into a temp extraction folder -- use sys.executable's
# real path instead. __file__ can also be undefined in some interactive contexts, so
# fall back to load_dotenv()'s default (cwd) search in that case.
try:
    if getattr(sys, "frozen", False):
        _env_base_dir = os.path.dirname(sys.executable)
    else:
        _env_base_dir = os.path.dirname(os.path.abspath(__file__))
    _ENV_PATH = os.path.join(_env_base_dir, ".env")
except NameError:
    _ENV_PATH = ".env"
load_dotenv(_ENV_PATH, override=True)

_REQUIRED_ENV_VARS = ["DB_SERVER", "DB_DATABASE", "DB_UID", "DB_PWD", "MCO_CONFIG_FILE", "MCO_TPA_FILE"]
_missing_env_vars = [v for v in _REQUIRED_ENV_VARS if not os.environ.get(v)]
if _missing_env_vars:
    raise RuntimeError(
        f"Missing required .env value(s): {', '.join(_missing_env_vars)}. "
        f"Check that a .env file exists next to this script with all required keys set."
    )

# MCO_CONFIG_FILE/MCO_TPA_FILE in .env are a relative tail (e.g. "Data Team - India -
# Documents/MCO Tool and Config/mco_config.xlsx") shared identically by every team
# member -- only the OneDrive sync-root prefix differs per user, resolved below.
_ONEDRIVE_ORG_FOLDER = "ARCHER Systems LLC"

def _onedrive_sync_roots(org_folder=_ONEDRIVE_ORG_FOLDER):
    """Places a synced SharePoint library usually lives for the CURRENT user."""
    roots = [
        Path.home() / f"OneDrive - {org_folder}",   # "Add shortcut to My files"
        Path.home() / org_folder,                   # "Sync" button
    ]
    for var in ("OneDriveCommercial", "OneDrive"):  # OneDrive's own env vars
        val = os.environ.get(var)
        if val:
            roots.append(Path(val))
    seen, unique = set(), []
    for r in roots:
        if r not in seen:
            seen.add(r)
            unique.append(r)
    return unique

def _resolve_synced_path(relative_path):
    """Tries each known sync root, in order, and returns the first one where
    relative_path actually exists; raises a clear, actionable error if none do (e.g. the
    library isn't synced on this machine yet)."""
    for root in _onedrive_sync_roots():
        candidate = root / relative_path
        if candidate.exists():
            return str(candidate)
    raise RuntimeError(
        f"Could not find '{relative_path}' in any OneDrive-synced folder on this "
        f"machine. Make sure the SharePoint library is synced (or added as a shortcut) "
        f"in OneDrive, then try again."
    )

# Every config/TPA-file decision saved this run (across save_config_updates AND the
# standalone TPA-file patcher) accumulates here -- the source for both the centralized
# Audit_Trail sheet (mco_config.xlsx) and this run's own local audit copy. An entry is
# only added here once its real save has actually succeeded (never claim something was
# recorded when it wasn't).
_AUDIT_TRAIL_COLUMNS = ["Timestamp", "SavedBy", "SheetChanged", "Key", "OldValue", "NewValue", "TicketId", "LogFile"]
# LogFile is dropped only from the LOCAL per-run copy -- that file already lives right
# next to the log it would be pointing at, so repeating the path on every row is dead
# weight. The centralized Audit_Trail sheet in mco_config.xlsx keeps the full column set
# (incl. LogFile) since it's shared across every user/run from every folder, where that
# pointer is the only way back to the specific run's own proof.
_AUDIT_TRAIL_LOCAL_COLUMNS = [c for c in _AUDIT_TRAIL_COLUMNS if c != "LogFile"]
_AUDIT_TRAIL_THIS_RUN = []

# =====================================================
# RUN LOG -- full terminal transcript (every print + every input prompt/answer) for
# later review. Buffered in memory until OUTPUT_FILE's folder is known, then flushed
# to a real file there and written live after.
# =====================================================
class _TeeStream:
    def __init__(self, real_stream):
        self._real = real_stream
        self._buffer = []
        self._file = None

    def write(self, s):
        try:
            self._real.write(s)
        except UnicodeEncodeError:
            # Defensive fallback only -- the real fix is reconfiguring sys.stdout to
            # UTF-8 at the top of the file. This exists so that even an unexpected
            # residual encoding issue can never crash print()/the crash handler itself
            # (a real failure mode found testing the compiled .exe).
            _enc = getattr(self._real, "encoding", None) or "utf-8"
            self._real.write(s.encode(_enc, errors="replace").decode(_enc, errors="replace"))
        self.write_log_only(s)

    def write_log_only(self, s):
        """Logs s WITHOUT echoing it to the real console -- used for a user's typed
        input() answer, which the terminal already echoed on its own as they typed it
        (that's normal OS/terminal keystroke echo, not something Python printed). Using
        plain write() for that too would print it a SECOND time right under the prompt."""
        if self._file is not None:
            self._file.write(s)
        else:
            self._buffer.append(s)

    def flush(self):
        self._real.flush()
        if self._file is not None:
            self._file.flush()

    def attach_file(self, path):
        self._file = open(path, "w", encoding="utf-8")
        self._file.write("".join(self._buffer))
        self._file.flush()
        self._buffer = []

    def close(self):
        if self._file is not None:
            self._file.close()
            self._file = None
        elif self._buffer:
            # A crash before OUTPUT_FILE's folder is known (e.g. DB connection or
            # import-file read failing) means attach_file() never ran, so the buffer
            # (incl. the crash traceback) would otherwise be silently lost. Falls back
            # to a dated file next to the script/exe instead.
            try:
                _fallback_path = os.path.join(
                    _env_base_dir, f"crash_log_{datetime.now().strftime('%Y-%m-%d_%H%M%S')}.txt"
                )
                with open(_fallback_path, "w", encoding="utf-8") as _f:
                    _f.write("".join(self._buffer))
            except Exception:
                pass
            self._buffer = []

_RUN_LOG_STREAM = _TeeStream(sys.stdout)
sys.stdout = _RUN_LOG_STREAM
atexit.register(_RUN_LOG_STREAM.close)

_real_input = input

# Control characters Excel can't store: one in any cell (pasted text, a DB value) made
# the final workbook save crash and lose the whole run. Removed where data comes in.
_CONTROL_CHARS_RE = re.compile('[\x00-\x08\x0b\x0c\x0e-\x1f]')

def _strip_control_chars(v):
    return _CONTROL_CHARS_RE.sub("", v) if isinstance(v, str) else v

def _strip_control_chars_df(df):
    """Same DataFrame with control characters removed from every text cell and header."""
    if df is None or not isinstance(df, pd.DataFrame):
        return df
    df = df.rename(columns=lambda c: _strip_control_chars(c) if isinstance(c, str) else c)
    for col in df.columns:
        if df[col].dtype == object or str(df[col].dtype).startswith(("string", "str")):
            df[col] = df[col].map(_strip_control_chars)
    return df

def _logged_input(prompt=""):
    """Drop-in replacement for the builtin input() -- writes the prompt through
    _RUN_LOG_STREAM (so it's captured) instead of letting input() print it via its own
    internal routing, then logs what the user typed too (plain stdout capture alone
    never sees that -- it's read from stdin, not written to stdout). The typed answer
    is logged WITHOUT echoing it back to the real console -- the terminal already
    echoed it live as the user typed, so writing it again would double-print it right
    under the prompt."""
    sys.stdout.write(prompt)
    sys.stdout.flush()
    answer = _strip_control_chars(_real_input())
    _RUN_LOG_STREAM.write_log_only(answer + "\n")
    return answer

import builtins
builtins.input = _logged_input

# =====================================================
# CRASH HANDLING + "RUN AGAIN?" -- so a double-clicked .exe doesn't just vanish on an
# uncaught exception (the console closes immediately otherwise), and the same window
# can process another file without relaunching by hand. '_restart_fresh' relaunches
# this exact script/exe as a brand-new process for a fully clean state. Note: an
# interactive dev console (Spyder/IPython) installs its own sys.excepthook, so this
# only takes effect as a plain script or compiled .exe.
# =====================================================
def _restart_fresh():
    _RUN_LOG_STREAM.close()
    try:
        conn.close()
    except Exception:
        pass
    # sys.argv is unreliable here (depends how this was launched), so __file__ builds
    # the relaunch command instead. os.execv was tried first -- hangs or drops into a
    # broken REPL on Windows. subprocess.Popen (fire-and-forget) + os._exit() spawns
    # correctly but races: the terminal reclaims the console the instant this process
    # exits, while the new child is still starting and about to read its own input()
    # from that same console. subprocess.run() blocks until the new process (and any
    # further restarts it triggers) fully finishes, so the terminal never reclaims the
    # console mid-restart -- verified with a 4-deep restart chain, no hang.
    try:
        if getattr(sys, "frozen", False):
            _relaunch_args = [sys.executable]
        else:
            _relaunch_args = [sys.executable, os.path.abspath(__file__)]
    except NameError:
        print("Could not determine this script's own file path to restart automatically "
              "-- please re-run it manually.")
        return
    subprocess.run(_relaunch_args)
    os._exit(0)

def _offer_rerun(prompt_text):
    answer = input(prompt_text).strip().upper()
    if answer == "R":
        _restart_fresh()

def _pause_on_crash(exc_type, exc_value, exc_tb):
    print("\n" + "="*60)
    print("❌ The automation stopped because of an unexpected error:")
    print("="*60)
    # Real bug found: traceback.print_exception defaults to sys.stderr, which is
    # NEVER wrapped by _TeeStream (only sys.stdout is) -- the actual traceback was
    # silently never reaching log_*.txt/crash_log_*.txt, only this function's own
    # print() banner lines (routed through the wrapped stdout) were. file=sys.stdout
    # routes it through the same wrapped stream so it's both displayed AND captured.
    traceback.print_exception(exc_type, exc_value, exc_tb, file=sys.stdout)

    # Best-effort log backup even on a crash -- a crashed run's log is exactly the one
    # most worth keeping. _run_log_path/CONFIG_FILE/TicketId may not exist yet if the
    # crash happened very early (before the output folder / config path were resolved),
    # so this is guarded via globals() instead of a bare NameError-prone reference.
    try:
        _g = globals()
        if _g.get("_run_log_path") and _g.get("CONFIG_FILE"):
            _backup_log_file(_run_log_path, _g.get("TicketId", "UNKNOWN"), CONFIG_FILE)
    except Exception:
        pass

    _offer_rerun("\nPress 'R' and Enter to try again, or just press Enter to close this window: ")

sys.excepthook = _pause_on_crash

# 3-tier terminal hierarchy: _phase_banner (once per phase) > _new_item (new item
# within a phase, skipped for the phase's own first item) > _gap (same-item steps).
_PHASE_HAD_ITEM = set()

def _phase_banner(phase_key, title):
    # width grows to fit a longer title (e.g. the LA-file variant embeds the actual
    # tab name) instead of breaking the box's alignment.
    width = max(60, len(title) + 4)
    print()
    print("┌" + "─" * (width - 2) + "┐")
    print("│" + title.center(width - 2) + "│")
    print("└" + "─" * (width - 2) + "┘")
    _PHASE_HAD_ITEM.discard(phase_key)

def _new_item(phase_key):
    if phase_key in _PHASE_HAD_ITEM:
        print("\n" + "─" * 60)
    _PHASE_HAD_ITEM.add(phase_key)

def _gap():
    print()

def _divider():
    """Unconditional line -- used only outside the phase/item system, e.g. the
    permission-retry prompt below, which can interrupt at any point regardless of
    phase and isn't itself a "new item" of anything."""
    print("\n" + "─" * 60)

# PLAN IDENTIFICATION's banner is lazy (unlike every other phase's eager one): a run
# where every combo resolves silently (direct/override match, or auto-routed to a
# Pending_* sheet) used to show the banner then total silence, looking like the phase
# never ran. Only called at points guaranteed to lead to an actual input() this run.
_PLAN_BANNER_SHOWN = [False]

def _plan_banner_once():
    if not _PLAN_BANNER_SHOWN[0]:
        _phase_banner("plan", "PLAN IDENTIFICATION")
        _PLAN_BANNER_SHOWN[0] = True

def _print_candidate_table(candidates_df):
    """Prints CollectorId/LienholderId/LienType candidates (an ambiguous-duplicate
    pick list) as a gridline table -- same style family as _print_scope_diag_table --
    instead of a plain '[i] CollectorId=... LienholderId=... LienType=...' text list.
    The '#' column matches the number the caller's own prompt asks the user to type."""
    cols = [("#", 2), ("CollectorId", 11), ("LienholderId", 12), ("LienType", 19)]
    rows = [
        {"#": str(i), "CollectorId": str(r["CollectorId"]), "LienholderId": str(r["LienholderId"]),
         "LienType": str(r["LienType"])}
        for i, (_, r) in enumerate(candidates_df.iterrows(), start=1)
    ]
    widths = {name: max(minw, len(name), max((len(r[name]) for r in rows), default=0)) for name, minw in cols}

    def hline():
        return "+" + "+".join("-" * (widths[n] + 2) for n, _ in cols) + "+"

    def row_line(vals):
        return "|" + "|".join(f" {vals[n]:<{widths[n]}} " for n, _ in cols) + "|"

    print(hline())
    print(row_line({n: n for n, _ in cols}))
    print(hline())
    for r in rows:
        print(row_line(r))
    print(hline())

_HYDRATED_PATHS = set()

def _ensure_file_hydrated(path):
    """Real user-reported crash: a synced SharePoint/OneDrive file can sit on disk as a
    'Files On-Demand' cloud-only placeholder (NTFS ReparsePoint) even though Explorer
    shows it as if it were fully present. openpyxl's own zipfile-based reader can then
    read an incomplete/garbage stub instead of triggering a real download first, raising
    a confusing InvalidFileException ("does not support [blank] file format") that has
    nothing to do with the file's actual format. A plain buffered read forces Windows to
    fully download it before any Excel-aware code ever touches it. Cached per path (via
    _HYDRATED_PATHS) so a file already proven fully-local this run isn't re-read on
    every single access to it -- cheap, but not free, and OneDrive doesn't dehydrate a
    file back to a placeholder mid-session anyway. Never allowed to block the real read
    that follows -- any failure here (including the file not existing yet) is swallowed;
    whatever reads the file next will raise its own, real error if something's
    genuinely wrong."""
    if path in _HYDRATED_PATHS:
        return
    try:
        with open(path, "rb") as f:
            while f.read(1024 * 1024):
                pass
        _HYDRATED_PATHS.add(path)
    except Exception:
        pass

def _retry_on_permission_error(operation, description):
    """Runs `operation` (a zero-arg callable wrapping a file open/read/save) and
    retries it if it raises PermissionError -- e.g. the file is open elsewhere, such
    as in Excel. Without this, that error crashes the whole run outright, losing
    every interactive answer and resolved plan already made so far, forcing the user
    to redo all of it from scratch on a fresh run. Loops on 'press Enter to retry'
    until it either succeeds or the user explicitly cancels with 'C', so a merely
    temporary lock (close the file, retry) never has to cost the whole run."""
    while True:
        try:
            return operation()
        except PermissionError:
            _divider()
            answer = input(
                f"\n⚠️  Could not access {description} -- it appears to be open "
                f"elsewhere (e.g. in Excel). Close it and press Enter to retry, or "
                f"type 'C' to cancel this run: "
            ).strip().upper()
            if answer == "C":
                raise

CONFIG_BACKUP_RETENTION_DAYS = 30  # how many days of dated config backups to keep -- change just this number
LOG_BACKUP_RETENTION_DAYS = 365  # how many days of backed-up run logs to keep -- separate policy from config backups

def _file_content_hash(path):
    """SHA-256 of the file's raw bytes -- a precise fingerprint of its actual
    content. Read once per check; these are small config/TPA Excel files (a few
    hundred KB at most), so this is cheap. Unlike file size, two genuinely
    different saves can never coincidentally share the same hash -- closing a real
    (if rare) gap size alone left: two edits landing on the exact same byte COUNT
    would otherwise look identical to the conflict check below."""
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()

def _file_fingerprint(path):
    """(mtime, content hash) together -- more robust than mtime alone for detecting
    whether a shared file genuinely changed since it was loaded. These files live in
    an OneDrive-synced SharePoint folder, and OneDrive's sync client can touch a
    file's metadata (a background re-hash/conflict-check pass) without any real
    content change -- mtime alone would then falsely trigger the "someone else
    changed this" concurrent-save check elsewhere, deferring a perfectly safe save
    for no reason. A pure metadata touch never changes a single byte, so the hash
    stays identical; a real save always changes at least one byte, so the hash
    always changes too -- requiring both mtime AND hash to differ (see
    _fingerprint_really_changed) closes that gap."""
    return (os.path.getmtime(path), _file_content_hash(path))

def _fingerprint_really_changed(old_fp, new_fp):
    """True only when BOTH mtime and content hash differ from the loaded fingerprint
    -- a pure metadata touch (mtime moves, content/hash unchanged) must NOT count as
    a real external change. A plain tuple != would wrongly trigger on either one
    alone differing, which defeats the whole point of _file_fingerprint."""
    return old_fp[0] != new_fp[0] and old_fp[1] != new_fp[1]

def _check_file_not_locked(path):
    """A '~$filename' sibling is Excel's own lock marker -- created the instant someone
    opens this file for editing, cleared the instant they close it. Far faster and more
    certain than waiting for a PermissionError on save, and catches the single most
    common real-world cause of a blocked save before any work is wasted building it.
    Returns False only if the user explicitly cancels (loops/waits otherwise) -- the
    caller must abandon just this save, not crash the whole run."""
    lock_path = os.path.join(os.path.dirname(path), "~$" + os.path.basename(path))
    while os.path.exists(lock_path):
        answer = input(
            f"\n⚠️  '{os.path.basename(path)}' appears to be open in Excel right now "
            f"(someone may be editing it). Close it there first, then press Enter to "
            f"retry, or type 'C' to cancel without saving.\n"
            f"    (If no one actually has it open, Excel may have left this marker "
            f"behind from an earlier crash -- delete '{os.path.basename(lock_path)}' "
            f"next to the file and retry.)\n> "
        ).strip().upper()
        if answer == "C":
            return False
    return True

def _backup_config_file(config_path):
    """Dated snapshot of the config file, taken right before any save touches it.
    Lives in a 'Backups\\Config_Backups' subfolder derived from config_path itself
    (never a separate hardcoded location) -- so if CONFIG_FILE's own path ever changes
    (e.g. a future move to a SharePoint-synced folder), the backup location follows it
    automatically with zero code change. Same-day reruns overwrite today's backup
    (filename has no time component) instead of piling up, and anything older than
    CONFIG_BACKUP_RETENTION_DAYS is deleted here too, so this folder never grows
    unbounded. Never allowed to block the real save -- any failure here is swallowed
    with a warning."""
    try:
        backup_dir = os.path.join(os.path.dirname(config_path), "Backups", "Config_Backups")
        os.makedirs(backup_dir, exist_ok=True)
        stem = os.path.splitext(os.path.basename(config_path))[0]
        ext = os.path.splitext(config_path)[1]
        today = datetime.now()
        backup_path = os.path.join(backup_dir, f"{stem}_{today.strftime('%Y-%m-%d')}{ext}")
        shutil.copy2(config_path, backup_path)

        _prefix = f"{stem}_"
        for fname in os.listdir(backup_dir):
            if not (fname.startswith(_prefix) and fname.endswith(ext)):
                continue
            date_part = fname[len(_prefix):-len(ext)]
            try:
                file_date = datetime.strptime(date_part, "%Y-%m-%d")
            except ValueError:
                continue
            if (today - file_date).days > CONFIG_BACKUP_RETENTION_DAYS:
                os.remove(os.path.join(backup_dir, fname))
    except Exception as e:
        print(f"⚠️  Could not back up config file before saving (continuing anyway): {e}")

def _backup_log_file(log_path, ticket_id, config_path):
    """Copies this run's finished log file into a shared 'Backups\\Log_Backups'
    subfolder (next to Config_Backups, derived from config_path -- same self-locating
    pattern as _backup_config_file, so both follow CONFIG_FILE automatically). The
    personal/local log file (log_path, in the user's own Output folder) is left exactly
    as-is; the backup copy's name is that same filename with the ticket number appended,
    so a later search by ticket number finds it. Called once a real TicketId is known
    (end of a successful run, or from the crash handler) -- never mid-run, since a
    resume run's TicketId isn't resolved until partway through. Never allowed to block
    anything else -- any failure here is swallowed with a warning, same as config's
    own backup.
    Real bug found: _TeeStream only flushes its log file on an input() call (see
    _logged_input) -- a stretch of print()-only output (e.g. everything from the last
    prompt through the final "COMPLETED SUCCESSFULLY" banner) can sit unflushed in
    Python's own buffer, invisible to shutil.copy2 below (which reads straight off
    disk). The personal/local log file still ends up complete because its own final
    close() flushes everything, but a shared-backup copy taken before that close()
    would silently miss whatever hadn't been flushed yet -- hence the flush here first."""
    try:
        _RUN_LOG_STREAM.flush()
        backup_dir = os.path.join(os.path.dirname(config_path), "Backups", "Log_Backups")
        os.makedirs(backup_dir, exist_ok=True)
        stem, ext = os.path.splitext(os.path.basename(log_path))
        safe_ticket = re.sub(r'[\\/:*?"<>|]', "_", str(ticket_id)).strip() or "NO_TICKET"
        backup_path = os.path.join(backup_dir, f"{stem}_{safe_ticket}{ext}")
        shutil.copy2(log_path, backup_path)

        today = datetime.now()
        _date_re = re.compile(r"(\d{4}-\d{2}-\d{2})")
        for fname in os.listdir(backup_dir):
            m = _date_re.search(fname)
            if not m:
                continue
            try:
                file_date = datetime.strptime(m.group(1), "%Y-%m-%d")
            except ValueError:
                continue
            if (today - file_date).days > LOG_BACKUP_RETENTION_DAYS:
                os.remove(os.path.join(backup_dir, fname))
    except Exception as e:
        print(f"⚠️  Could not back up this run's log file (continuing anyway): {e}")

def _audit_row(sheet_changed, key, old_value, new_value, ticket_id):
    """One accountability record for a single config/TPA-file decision -- who made it,
    when, exactly what changed, which ticket it belongs to, and which run-log file has
    the full word-for-word prompt/answer proof behind it. SheetChanged names where the
    actual change landed (one of the 5 mco_config.xlsx sheets, or 'TPA_FILE_PATCH' for
    the standalone TPA-file patcher). old_value is the literal string '(new record)'
    when nothing existed before at all, or '(was blank)' when a real existing row's
    field was simply empty -- these are deliberately NOT the same as each other, and
    neither is left as a plain blank cell, so a later reader never has to guess which
    case produced an empty-looking OldValue."""
    return {
        "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "SavedBy": os.environ.get("USERNAME", "UNKNOWN"),
        "SheetChanged": sheet_changed,
        "Key": key,
        "OldValue": old_value,
        "NewValue": new_value,
        "TicketId": ticket_id,
        "LogFile": _run_log_path,
    }

def _ensure_audit_trail_header(ws):
    """The user adds the 'Audit_Trail' sheet by hand (just the tab, no header row) --
    this writes the real header the first time anything is ever appended to it, so the
    exact column names code depends on don't rely on someone having typed them in
    correctly by hand."""
    if ws.max_row < 1 or all(c.value is None for c in ws[1]):
        for i, col_name in enumerate(_AUDIT_TRAIL_COLUMNS, start=1):
            ws.cell(row=1, column=i, value=col_name)

def _audit_trail_ws(wb):
    """The config's Audit_Trail sheet, created (with its header) if the tab is missing."""
    ws = wb["Audit_Trail"] if "Audit_Trail" in wb.sheetnames else wb.create_sheet("Audit_Trail")
    _ensure_audit_trail_header(ws)
    return ws

def _sheet_header(ws):
    """A config sheet's header names with stray spaces trimmed (a typed "ScopeText " used
    to crash the save)."""
    return [str(c.value).strip() if c.value is not None else None for c in ws[1]]

def _config_sheet(wb, name, required_cols):
    """(sheet, header) of a config sheet about to be saved to: the tab is created if it is
    missing, and a column the save needs that the header lacks is added at the end
    (existing columns are never renamed or moved)."""
    ws = wb[name] if name in wb.sheetnames else wb.create_sheet(name)
    header = _sheet_header(ws)
    while header and header[-1] is None:
        header.pop()
    for col in required_cols:
        if col not in header:
            ws.cell(row=1, column=len(header) + 1, value=col)
            header.append(col)
    return ws, header

def _append_audit_trail_rows(entries):
    """For callers that don't already have mco_config.xlsx open (e.g. the standalone
    TPA-file patcher, which only ever opens TPA_FILE) -- opens it fresh just for this,
    same backup-before-write + permission-retry safety as every other config write.
    save_config_updates itself does NOT use this -- it appends to the same already-open
    workbook as its other 5 sheets, in the same save cycle, to avoid a second redundant
    open/lock/save cycle on the same file within one run."""
    if not entries:
        return
    if not _check_file_not_locked(CONFIG_FILE):
        print(f"Audit_Trail save cancelled -- {len(entries)} entry(ies) NOT recorded.")
        return
    _backup_config_file(CONFIG_FILE)
    _ensure_file_hydrated(CONFIG_FILE)
    try:
        audit_wb = _retry_on_permission_error(lambda: _opxl.load_workbook(CONFIG_FILE), CONFIG_FILE)
    except PermissionError:
        print(f"⚠️  Could not open mco_config.xlsx to record {len(entries)} Audit_Trail entry(ies) -- add manually.")
        return
    _loaded_fingerprint = _file_fingerprint(CONFIG_FILE)
    audit_ws = _audit_trail_ws(audit_wb)
    for entry in entries:
        new_row = audit_ws.max_row + 1
        for i, col_name in enumerate(_AUDIT_TRAIL_COLUMNS, start=1):
            audit_ws.cell(row=new_row, column=i, value=entry.get(col_name))
    if _fingerprint_really_changed(_loaded_fingerprint, _file_fingerprint(CONFIG_FILE)):
        print(
            f"⚠️  mco_config.xlsx was changed by someone else just now (likely synced "
            f"in from another user's save) -- {len(entries)} Audit_Trail entry(ies) NOT "
            f"recorded, to avoid overwriting their change. They'll be re-queued if this "
            f"row gets resolved again on a future run."
        )
        return
    try:
        _retry_on_permission_error(lambda: audit_wb.save(CONFIG_FILE), CONFIG_FILE)
        _AUDIT_TRAIL_THIS_RUN.extend(entries)
    except PermissionError:
        print(f"⚠️  Could not save {len(entries)} Audit_Trail entry(ies) -- mco_config.xlsx appears locked.")

def _ensure_writable(path):
    """A `with pd.ExcelWriter(path, ...) as writer: ...` block can span many sheet
    writes before the file is actually saved to disk -- retrying the WHOLE block on a
    PermissionError there would mean re-running all of that write logic. Checking
    upfront, before any of it starts, catches a file already open elsewhere (e.g. in
    Excel) cheaply and immediately, via the exact same retry-or-cancel prompt."""
    def _try_open():
        if os.path.exists(path):
            with open(path, "a"):
                pass
    _retry_on_permission_error(_try_open, path)

def _long_path(path):
    """A Windows path of 250+ characters gets the long-path prefix, so the write still works
    on a machine without long-path support."""
    if os.name == "nt" and len(path) >= 250 and not path.startswith("\\\\?\\"):
        return "\\\\?\\" + os.path.abspath(path)
    return path

Connection_String = (
    "DRIVER={ODBC Driver 17 for SQL Server};"
    f"SERVER={os.environ['DB_SERVER']};"
    f"DATABASE={os.environ['DB_DATABASE']};"
    f"UID={os.environ['DB_UID']};"
    f"PWD={os.environ['DB_PWD']};"
)

conn = pyodbc.connect(Connection_String)
print("✅ SQL connection established successfully")

# SQLSTATEs / texts of a dropped connection (e.g. after a long pause at a question).
_DB_CONN_LOST_MARKERS = ("08S01", "08001", "08003", "08004", "08007", "HYT00", "HYT01",
                         "communication link", "connection")

def _read_sql(query, params=None):
    """pd.read_sql on the run's connection. If the connection was dropped (the user can
    spend a long time answering questions), it reconnects once and tries again."""
    global conn
    def _run():
        return pd.read_sql(query, conn) if params is None else pd.read_sql(query, conn, params=params)
    try:
        return _run()
    except Exception as e:
        text = (" ".join(str(a) for a in getattr(e, "args", ())) or str(e)).lower()
        if not any(m.lower() in text for m in _DB_CONN_LOST_MARKERS):
            raise
        print("⚠️  The database connection was lost -- reconnecting and trying again...")
        try:
            conn.close()
        except Exception:
            pass
        conn = pyodbc.connect(Connection_String)
        return _run()

# %% [Cell 2] Main pipeline
# =====================================================
# USER INPUTS
# =====================================================

_divider()
while True:
    MCO_INPUT_FILE = input("Path of Import File: ").lstrip("﻿").strip().strip('"').strip("'").strip()
    if MCO_INPUT_FILE and os.path.isfile(MCO_INPUT_FILE) and MCO_INPUT_FILE.lower().endswith((".xlsx", ".xlsm")):
        break
    if not MCO_INPUT_FILE:
        print("Please enter the full path of the Excel file.")
    elif not os.path.isfile(MCO_INPUT_FILE):
        print(f"File not found: {MCO_INPUT_FILE}. Please check the path and try again.")
    else:
        print("Please give an Excel .xlsx file (save an .xls or .csv file as .xlsx first).")

# Resume detection must happen FIRST, before any other prompt -- a resume file's rows
# each carry their own originally-stored TicketId/date/response-type, so none of the
# usual upfront questions apply.
import openpyxl as _opxl
_ensure_file_hydrated(MCO_INPUT_FILE)
_input_wb = _retry_on_permission_error(
    lambda: _opxl.load_workbook(MCO_INPUT_FILE, read_only=True), "the import file"
)
_input_tabs = _input_wb.sheetnames
_input_wb.close()

KNOWN_PENDING_SHEETS = ["Pending_Scope_Review", "Pending_TPA_Review", "Pending_TPA_Blank_Fields",
                        "Pending_Lien_Match_Review", "Pending_Plan_Text_Review", "Pending_Error_Review"]
RESUME_SHEETS_PRESENT = [s for s in KNOWN_PENDING_SHEETS if s in _input_tabs]
IS_RESUME_RUN = bool(RESUME_SHEETS_PRESENT)

# Real crash (user report): both of these were originally defined only inside the
# non-resume `else:` branch below, alongside the upfront prompts that are skipped on
# a resume run -- but Step 3's TPA resolution (ask_numbered_choice, e.g. the D-flow
# "provide LienHolder details" prompt) and every resume-processing block's "Save
# this decision for future runs?" question (ask_yes_no) both run regardless of
# IS_RESUME_RUN, and crashed with NameError the first time a resume run actually
# reached either of them. Defined here, unconditionally, before the branch split.
def ask_numbered_choice(options, header, prompt_text):
    _gap()
    print(header)
    for i, opt in enumerate(options, start=1):
        print(f"  [{i}] {opt}")
    while True:
        answer = input(prompt_text).strip()
        if answer.isdigit() and 1 <= int(answer) <= len(options):
            return options[int(answer) - 1]
        print(f"Please enter a number between 1 and {len(options)}.")

def ask_yes_no(prompt_text):
    """Accepts only 'Y' or 'N' (case-insensitive) -- anything else re-prompts
    instead of being silently treated as a decline."""
    _gap()
    while True:
        answer = input(prompt_text).strip().upper()
        if answer in ("Y", "N"):
            return answer == "Y"
        print("Please enter 'Y' or 'N'.")

def _ask_ticket_id():
    """Ticket number: digits only (a leading '#' and spaces are dropped). Blank or text
    re-prompts -- it goes into file names and every note."""
    while True:
        answer = input("TicketId: ").strip().lstrip("#").strip()
        if re.fullmatch(r"[0-9]+", answer):
            return answer
        print("Please enter the ticket number using digits only (e.g. 123456).")

def _ask_typed_id(prompt_text, allow_skip=True):
    """A CollectorId/LienholderId typed at a prompt: digits only (spaces and Excel's '.0'
    dropped). Returns the number as text, or "S" when skipping is allowed and chosen."""
    while True:
        answer = input(prompt_text).strip()
        if allow_skip and answer.upper() == "S":
            return "S"
        value = re.sub(r"\.0+$", "", answer.replace(" ", ""))
        if re.fullmatch(r"[0-9]+", value):
            return value
        print("Please enter a number" + (", or 'S' to skip." if allow_skip else "."))

# Only ever exactly one of these two values in real use -- picked by number (same
# ask_numbered_choice pattern already used for Response type) everywhere a user
# would otherwise free-type LienType, so a typo can never silently create a
# LienType value that doesn't match either real one.
LIEN_TYPE_OPTIONS = ["Medicaid Lien - MT", "Private Lien - PLRP"]

def _canonical_lien_type(v):
    """A hand-typed LienType (review file) -> its exact LIEN_TYPE_OPTIONS value, ignoring
    case, spaces and dashes; None if it is neither."""
    if v is None or (not isinstance(v, str) and pd.isna(v)):
        return None
    key = re.sub(r"[\s\-]+", "", str(v)).lower()
    for opt in LIEN_TYPE_OPTIONS:
        if re.sub(r"[\s\-]+", "", opt).lower() == key:
            return opt
    return None

def _unknown_collector_ids(collector_ids):
    """Numeric CollectorIds not found in LPM's collectors table. Best-effort: a failed
    lookup returns an empty set instead of blocking the run."""
    ids = sorted({str(c) for c in collector_ids if re.fullmatch(r"\d+", str(c))})
    if not ids:
        return set()
    try:
        found = _read_sql(f"select Id from collectors where Id in {make_sql_in(ids)}")
    except Exception:
        return set()
    return set(ids) - {_stringify_id(i) for i in found["Id"]}

# =====================================================
# LPM LIENHOLDERS -- LienholderId for plans the TPA file can't resolve
# =====================================================
# LPM's MCO lienholder names carry a "(XX MCO)" / "(XX CCO)" tag (e.g. "Community Health
# Plan of Washington (WA MCO)"), so one state's MCO lienholders are a small list.
_LPM_TAG_RE = re.compile(r"\(\s*[A-Za-z]{2}\s+(?:MCO|CCO)\s*\)")
_LPM_LIENHOLDER_CACHE = {}
# Review reasons where the TPA file has nothing at all for the plan -- LienholderId comes
# from LPM (match or pick) and MCO_Name becomes the LPM name.
_LPM_REVIEW_REASONS = {"State not found in TPA reference file", "Tort not found in TPA reference file"}

def _strip_lpm_tag(name):
    return re.sub(r"\s+", " ", _LPM_TAG_RE.sub(" ", str(name))).strip()

def _lpm_state_lienholders(state_code):
    """This state's MCO lienholders from LPM (LienHolderId as text, Name), read once per
    run with the old find_lhid query. None if the state code isn't 2 letters or LPM can't
    be read -- the caller then just doesn't pre-fill."""
    st = str(state_code or "").strip().upper()
    if st not in _LPM_LIENHOLDER_CACHE:
        result = None
        if re.fullmatch(r"[A-Z]{2}", st):
            try:
                df = _read_sql("select LienHolderId, Name from LienHolders where Name like ? or Name like ?",
                              params=[f"%{st} MCO%", f"%{st} CCO%"])
                # LIKE '%WA MCO%' also matches a spelled-out "IOWA MCO"
                tag = re.compile(rf"(?<![A-Za-z]){st}\s+(?:MCO|CCO)(?![A-Za-z])", re.IGNORECASE)
                df = df[df["Name"].map(lambda n: pd.notna(n) and bool(tag.search(str(n))))]
                df = df.assign(LienHolderId=df["LienHolderId"].map(_stringify_id),
                               Name=df["Name"].astype(str).map(_strip_control_chars).str.strip())
                result = df.drop_duplicates(subset=["LienHolderId"]).reset_index(drop=True)
            except Exception as e:
                print(f"⚠️  Could not read LPM's lienholder list for {st} ({e}). LienholderId will not be pre-filled.")
        _LPM_LIENHOLDER_CACHE[st] = result
    return _LPM_LIENHOLDER_CACHE[st]

def _lienholder_names_by_id(lienholder_ids):
    """{LienholderId: LPM name} for the numeric IDs given ({} for none). None when LPM
    can't be read -- callers then skip the check instead of blocking the run."""
    ids = sorted({str(i) for i in lienholder_ids if re.fullmatch(r"\d+", str(i))})
    if not ids:
        return {}
    try:
        found = _read_sql(f"select LienHolderId, Name from LienHolders where LienHolderId in {make_sql_in(ids)}")
        return {_stringify_id(r["LienHolderId"]): _strip_control_chars(str(r["Name"])).strip() for _, r in found.iterrows()}
    except Exception:
        return None

def _lpm_initials_match(text, name):
    """A short form typed as initials ("CHPW" = Community Health Plan of Washington, with or
    without small words like "of"). Only orders the pick list -- never fills anything."""
    letters = re.sub(r"[^A-Za-z]", "", str(text)).upper()
    if len(str(text).split()) != 1 or not 2 <= len(letters) <= 8:
        return False
    words = []
    for w in re.findall(r"[A-Za-z]+", _strip_lpm_tag(name)):
        if w.lower().endswith("healthcare") and len(w) > len("healthcare"):
            words += [w[:-len("healthcare")], "Health", "Care"]
        else:
            words.append(w)
    small = {"of", "the", "and", "a", "an", "inc", "llc", "co", "corp", "ltd"}
    return any(r and "".join(w[0] for w in r).upper() == letters
               for r in (words, [w for w in words if w.lower() not in small]))

# Words that never show two names are the same plan (generic, legal, tag, US state words).
_NAME_FILLER_WORDS = set(
    "health healthy healthcare plan plans care services managed group community system systems "
    "national insurance inc llc corp company the and mco cco "
    "alabama alaska arizona arkansas california colorado connecticut delaware florida georgia hawaii "
    "idaho illinois indiana iowa kansas kentucky louisiana maine maryland massachusetts michigan "
    "minnesota mississippi missouri montana nebraska nevada hampshire jersey mexico york carolina "
    "dakota ohio oklahoma oregon pennsylvania rhode island tennessee texas utah vermont virginia "
    "washington wisconsin wyoming north south west new district columbia".split()
)

def _names_related(text, lpm_name):
    """Loose test for the resume y/n: does a file's plan text share any real word with an
    LPM lienholder name, or does one start with the other?"""
    def _words(s):
        return {w for w in re.findall(r"[a-z]+", _strip_lpm_tag(s).lower()) if len(w) >= 3} - _NAME_FILLER_WORDS
    if _words(text) & _words(lpm_name):
        return True
    a = re.sub(r"[^a-z]", "", _strip_lpm_tag(text).lower())
    b = re.sub(r"[^a-z]", "", _strip_lpm_tag(lpm_name).lower())
    return bool(a and b and (a.startswith(b) or b.startswith(a)))

def _pick_lienholder_id(plan_text, ids, lpm_names):
    """Rows of one plan text filled with different LienholderIds -> the user picks one.
    None = skip (those rows stay in review)."""
    _gap()
    print(f"⚠️  Different LienholderIds were filled for {plan_text!r}:")
    for i, lh_id in enumerate(ids, start=1):
        print(f"  [{i}] {lh_id} = {lpm_names.get(lh_id, 'not found in LPM')}")
    while True:
        answer = input("Enter the number of the correct one, or 'S' to skip (these rows stay in review): ").strip().upper()
        if answer == "S":
            return None
        if answer.isdigit() and 1 <= int(answer) <= len(ids):
            return ids[int(answer) - 1]
        print(f"Please enter a number between 1 and {len(ids)}, or 'S'.")

def _confirm_unknown_id(kind, value, plan_label, answers, extra=""):
    """An ID typed on the review file that LPM doesn't know (typo, or created very
    recently): the user decides, asked once per ID (answers caches it)."""
    if (kind, value) not in answers:
        _new_item("resume")
        print(f"\n⚠️  {kind} {value} filled for {plan_label} was not found in LPM.")
        what = "collector" if kind == "CollectorId" else "lienholder"
        answers[(kind, value)] = ask_yes_no(
            f"It may be a typo, or a {what} created very recently. {extra}Use it anyway? (y/n): ")
    return answers[(kind, value)]

def _confirm_name_mismatch(lienholder_id, lpm_name, plan_text, answers):
    """A typed LienholderId whose LPM name shares nothing with the file's plan text (a
    typo can land on another real lienholder): asked once per (text, ID)."""
    key = (_normalize_key(plan_text), lienholder_id)
    if key not in answers:
        _new_item("resume")
        print(f"\n⚠️  LienholderId {lienholder_id} is {lpm_name!r} in LPM, but the plan text in the file is {plan_text!r}.")
        answers[key] = ask_yes_no("Use this LienholderId? (y/n): ")
    return answers[key]

def _prepare_tpa_review_df(pending_tpa_df):
    """Pending_TPA_Review as written: unhelpful columns dropped, the file's own plan text
    next to MCO_Name, the LPM name of any pre-filled LienholderId next to it, each plan's
    rows together, and ONE row per plan + Tort + Scope marked Fill_Here = YES."""
    # The 5 TPA_Match_* columns are only ever filled on the Pending_TPA_Blank_Fields path,
    # and MCOn_Existing doesn't help find LienholderId/CollectorId/LienType.
    _tpa_review_drop_cols = [
        "TPA_Match_TortId", "TPA_Match_StateMCO", "TPA_Match_TortName",
        "TPA_Match_Scope", "TPA_Match_LienholderName",
    ] + [c for c in pending_tpa_df.columns if c.startswith("MCO") and c.endswith("_Existing")]
    df = pending_tpa_df.drop(columns=_tpa_review_drop_cols, errors="ignore").copy()
    if "Original_MCO_Name" not in df.columns:
        df["Original_MCO_Name"] = None
    df["Original_MCO_Name"] = df["Original_MCO_Name"].where(df["Original_MCO_Name"].notna(), df["MCO_Name"])
    if "LPM_LienholderName" not in df.columns:
        df["LPM_LienholderName"] = None
    sort_cols = [c for c in ("MCO_Name", "Project_Name", "Scope_Classification", "ClientId") if c in df.columns]
    df = df.sort_values(sort_cols, kind="stable").reset_index(drop=True)
    group_cols = [c for c in ("State", "Project_Name", "MCO_Name", "Scope_Classification") if c in df.columns]
    # Built cell by cell: a blank (e.g. the Tort on a "Tort is blank" row) is "" here.
    def _txt(v):
        return "" if v is None or (not isinstance(v, str) and pd.isna(v)) else str(v)
    fill_key = df[group_cols].apply(lambda r: "|".join(_txt(v) for v in r), axis=1) if len(df) else pd.Series(dtype=object)
    # Tort blank in LPM: the Tort is per case, so one row to fill per case.
    if "PendingReason" in df.columns and len(df):
        _bt = df["PendingReason"] == BLANK_TORT_REASON
        _case_col = "CaseId" if "CaseId" in df.columns else "ClientId"
        _case = df[_case_col].map(lambda v: "BLANK TORT|" + _txt(v)) if _case_col in df.columns else "BLANK TORT|"
        fill_key = fill_key.where(~_bt, _case)
    df.insert(0, "Fill_Here", (~fill_key.duplicated()).map({True: "YES", False: None}))
    cols = [c for c in df.columns if c not in ("Original_MCO_Name", "LPM_LienholderName", "TortName")]
    cols.insert(cols.index("MCO_Name") + 1, "Original_MCO_Name")
    cols.insert(cols.index("LienholderId") + 1 if "LienholderId" in cols else len(cols), "LPM_LienholderName")
    if "TortName" in df.columns:
        cols.insert(cols.index("TortId") + 1 if "TortId" in cols else len(cols), "TortName")
    return df[cols]

# Colors (agreed 2026-11-24): cells/columns the user fills in a review sheet, the tool's own
# note columns, header-only marks, and tab colors.
FILL_IN_COLOR      = "FFF2CC"   # Butter Yellow: review cells to fill
NOTE_COL_COLOR     = "DDEBF7"   # Powder Blue: Yes_with_MCO note columns
REMARKS_COL_COLOR  = "E4DFEC"   # Lavender: Input_File_Original Remarks
KEY_HEADER_COLOR   = "FCE4D6"   # Peach: Unique / ClientId headers (header only)
UPLOAD_TAB_COLOR   = "A9D08E"   # Sage Green: upload tabs
ALERT_TAB_COLOR    = "FF6961"   # Pastel Red: discrepancy / excluded / error tabs

def _solid_fill(color):
    return _opxl.styles.PatternFill(start_color=color, end_color=color, fill_type="solid")

_REVIEW_LAST_ROW = 1000
_VALIDATION_TEXT = {
    "list": ("Please pick a value from the list.", "Click the arrow at the right of the cell to pick a value."),
    "whole": ("Please enter a whole number (an ID).", "Type the ID as a number."),
    "date": ("Please enter a real date (MM/DD/YYYY).", "Type the date as MM/DD/YYYY."),
}

def _add_review_validations(ws, df, list_cols=None, whole_cols=(), date_cols=()):
    """Strict checks on a review sheet's fill-in columns (rows 2..1000, invisible on empty
    rows): a dropdown for list_cols {column: [options]}, whole numbers for ID columns, real
    dates shown MM/DD/YYYY for date columns. A wrong value is refused with a message; the
    resume checks stay as a second layer."""
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.datavalidation import DataValidation
    cols = list(df.columns)

    def _add(col, dv, kind):
        letter = get_column_letter(cols.index(col) + 1)
        dv.error, dv.prompt = _VALIDATION_TEXT[kind]
        dv.errorTitle, dv.promptTitle = "Not allowed", col
        dv.errorStyle = "stop"
        dv.showErrorMessage = True
        dv.showInputMessage = True
        dv.add(f"{letter}2:{letter}{_REVIEW_LAST_ROW}")
        ws.add_data_validation(dv)

    for col, options in (list_cols or {}).items():
        if col in cols:
            _add(col, DataValidation(type="list", formula1='"' + ",".join(options) + '"', allow_blank=True), "list")
    for col in whole_cols:
        if col in cols:
            _add(col, DataValidation(type="whole", operator="greaterThan", formula1="0", allow_blank=True), "whole")
    for col in date_cols:
        if col in cols:
            _add(col, DataValidation(type="date", operator="between", formula1="DATE(1900,1,1)",
                                     formula2="DATE(2100,12,31)", allow_blank=True), "date")
            ci = cols.index(col) + 1
            for r in range(2, _REVIEW_LAST_ROW + 1):
                ws.cell(row=r, column=ci).number_format = "mm/dd/yyyy"

def _fill_cells(ws, df, cols_to_fill, rows=None):
    """Butter Yellow on the cells to fill (whole column when rows is None)."""
    fill = _solid_fill(FILL_IN_COLOR)
    cols = list(df.columns)
    for c in cols_to_fill:
        if c not in cols:
            continue
        ci = cols.index(c) + 1
        ws.cell(row=1, column=ci).fill = fill
        for r in (rows if rows is not None else range(2, len(df) + 2)):
            ws.cell(row=r, column=ci).fill = fill

def _format_tpa_review_sheet(ws, df):
    """Excel aids on Pending_TPA_Review: the cells to fill on Fill_Here rows highlighted
    (TortId + TortName too on "Tort is blank in LPM" rows), strict checks on the fill-in
    columns, a header filter and a frozen header row."""
    cols = list(df.columns)
    fill = _solid_fill(FILL_IN_COLOR)
    for r, values in enumerate(df.itertuples(index=False), start=2):
        row = dict(zip(cols, values))
        if row.get("Fill_Here") != "YES":
            continue
        to_fill = ["CollectorId", "LienType"]
        if pd.isna(row.get("LienholderId")) or not str(row.get("LienholderId")).strip():
            to_fill.append("LienholderId")
        if row.get("PendingReason") == BLANK_TORT_REASON:
            to_fill += ["TortId", "TortName"]
        for c in to_fill:
            if c in cols:
                ws.cell(row=r, column=cols.index(c) + 1).fill = fill
    _add_review_validations(ws, df, list_cols={"LienType": LIEN_TYPE_OPTIONS},
                            whole_cols=("LienholderId", "CollectorId", "TortId"))
    ws.auto_filter.ref = ws.dimensions
    ws.freeze_panes = "A2"

def _pick_conflicting_fill(plan_label, options):
    """options: the different (CollectorId, LienholderId, LienType) values filled for ONE
    plan. Returns the one the user picks, or None to skip (the plan stays in review)."""
    _gap()
    print(f"⚠️  Different values were filled for the same plan {plan_label}:")
    for i, (c, l, t) in enumerate(options, start=1):
        print(f"  [{i}] CollectorId={c}, LienholderId={l}, LienType={t}")
    while True:
        answer = input("Enter the number of the correct values, or 'S' to skip (the plan stays in review): ").strip().upper()
        if answer == "S":
            return None
        if answer.isdigit() and 1 <= int(answer) <= len(options):
            return options[int(answer) - 1]
        print(f"Please enter a number between 1 and {len(options)}, or 'S'.")

def _ask_save_decisions(labels):
    """One question for several "save for future runs?" decisions: y = all, n = none,
    s = one by one. Returns the indexes to save. A single decision keeps the plain y/n."""
    if len(labels) == 1:
        return [0] if ask_yes_no("Save this decision for future runs? (y/n): ") else []
    _gap()
    while True:
        answer = input(f"Save these {len(labels)} decisions for future runs? "
                       f"(y = all / n = none / s = choose one by one): ").strip().upper()
        if answer in ("Y", "N", "S"):
            break
        print("Please enter 'Y', 'N' or 'S'.")
    if answer == "Y":
        return list(range(len(labels)))
    if answer == "N":
        return []
    return [i for i, label in enumerate(labels) if ask_yes_no(f"Save {label}? (y/n): ")]

# Every "save for future runs?" decision of the run, asked ONCE at the end (see
# _ask_deferred_saves) -- the decision itself already applies to this run when made.
# Entries: (dedup key, label, target save list, item). The item is copied when queued,
# so nothing that runs later can change what gets saved.
_DEFERRED_SAVES = []

def _save_item_key(target_list, item):
    if isinstance(item, dict):
        values = tuple(sorted((k, _normalize_key(v)) for k, v in item.items() if k not in ("Ticket #", "Reason")))
    else:
        values = tuple(_normalize_key(v) for v in item)
    return (id(target_list), values)

def _defer_save(label, target_list, item):
    """Queues one decision for the end-of-run save question. The same decision queued
    twice (same save list, same values -- e.g. one alias reused for several Torts) is
    listed once."""
    key = _save_item_key(target_list, item)
    if any(entry[0] == key for entry in _DEFERRED_SAVES):
        return
    _DEFERRED_SAVES.append((key, label, target_list, dict(item) if isinstance(item, dict) else item))

def _ask_deferred_saves():
    """The run's single save question: lists every queued decision, then y = all /
    n = none / s = one by one (one decision keeps the plain y/n). Chosen items go to
    their save lists. Nothing queued = no question."""
    if not _DEFERRED_SAVES:
        return
    _gap()
    print(f"Decisions made this run ({len(_DEFERRED_SAVES)}):")
    labels = [f"[{i}] {entry[1]}" for i, entry in enumerate(_DEFERRED_SAVES, start=1)]
    for label in labels:
        print(f"  {label}")
    for i in _ask_save_decisions(labels):
        _DEFERRED_SAVES[i][2].append(_DEFERRED_SAVES[i][3])
    _DEFERRED_SAVES.clear()

def _tpa_entry_label(entry):
    return (f"TPA_Overrides: {entry.get('LienholderName')!r} (Tort {entry.get('TortName')!r} / "
            f"Scope {entry.get('Private Lien Case Scope')!r}) -> LienholderId {entry.get('LienholderId')}, "
            f"CollectorId {entry.get('CollectorId')}, {entry.get('LienType')}")

def is_valid_date_mmddyyyy(text):
    try:
        datetime.strptime(text, "%m-%d-%Y")
        return True
    except ValueError:
        return False

# Auto-detect State/Response Type/Rec'd Date from the import filename first (2026-09-25
# convention); only whichever piece isn't found this way still prompts. Tolerant,
# never raises -- a non-conforming filename just falls through to the normal prompt.
def _parse_state_from_filename(name):
    """The 2-letter code immediately before the word 'Mcaid' (any case), e.g.
    '...MI Mcaid...' -> 'MI'. None if that anchor word isn't found at all."""
    try:
        m = re.search(r'\b([A-Za-z]{2})\s*Mcaid\b', name, re.IGNORECASE)
        return m.group(1).upper() if m else None
    except Exception:
        return None

def _parse_response_type_from_filename(name):
    """Whichever of the 3 known response types sits immediately before OR after the
    word 'Response' or 'Request' (any case, either keyword) -- the longer 'EV & Claims'
    alternative is checked before the bare 'EV'/'Claims' ones so it isn't partially
    matched. Every such type found anywhere in the name is collected; if they all agree,
    that's the answer. None (never a guess) if nothing recognizable is found at all, OR
    if two+ DIFFERENT types are found -- e.g. 'Claims Request' and 'EV Response' both
    present -- so the caller falls back to asking the user instead of picking a side."""
    try:
        _TYPE = r'(EV\s*&\s*Claims|EV|Claims)'
        _SEP = r'[\s\-:_]*'
        # No trailing \b on the keyword itself -- "Responses"/"Requests" (plural) must
        # still match, same as the original 'Response'-only version always allowed.
        _KW = r'(?:Response|Request)'
        # One combined, non-overlapping scan (not two separate before/after regexes
        # per keyword) -- a type token sitting between two different keyword
        # occurrences (e.g. "Claims Request EV Response") would otherwise get
        # double-counted as belonging to both, silently canceling out the real
        # disagreement instead of surfacing it.
        _PAT = re.compile(rf'{_TYPE}{_SEP}\b{_KW}|\b{_KW}{_SEP}{_TYPE}', re.IGNORECASE)

        def _normalize_type(raw):
            found = re.sub(r'\s*&\s*', ' & ', raw.strip()).upper()
            for opt in RESPONSE_TYPE_OPTIONS:
                if found == opt.upper():
                    return opt
            return None

        found_types = {_normalize_type(m.group(1) or m.group(2)) for m in _PAT.finditer(name)} - {None}
        return found_types.pop() if len(found_types) == 1 else None
    except Exception:
        return None

def _parse_recd_date_from_filename(name):
    """The date immediately after 'rec'd' (apostrophe optional, straight or curly),
    normalized to MM-DD-YYYY (a 2-digit year is assumed 20XX). None if 'rec'd' isn't
    found, or what follows it isn't a real date."""
    try:
        m = re.search(r"rec[’'`]?d\s+(\d{1,2})[-/](\d{1,2})[-/](\d{2,4})", name, re.IGNORECASE)
        if not m:
            return None
        mm, dd, yy = m.group(1), m.group(2), m.group(3)
        if len(yy) == 2:
            yy = "20" + yy
        candidate = f"{int(mm):02d}-{int(dd):02d}-{yy}"
        return candidate if is_valid_date_mmddyyyy(candidate) else None
    except Exception:
        return None

RESPONSE_TYPE_OPTIONS = ["EV", "Claims", "EV & Claims"]

if IS_RESUME_RUN:
    print(f"\nResume file detected. Found: {', '.join(RESUME_SHEETS_PRESENT)}")
    print("Skipping upfront prompts -- each row will use its own originally stored values.\n")
    TicketId = "RESUME"
    STATE = None
    RESPONSE_DATE = None
    RESPONSE_TYPE = None
else:
    _gap()
    TicketId = _ask_ticket_id()

    _input_filename_only = os.path.basename(MCO_INPUT_FILE)
    STATE         = _parse_state_from_filename(_input_filename_only)
    RESPONSE_DATE = _parse_recd_date_from_filename(_input_filename_only)
    RESPONSE_TYPE = _parse_response_type_from_filename(_input_filename_only)

    _detected_bits = []
    if STATE:
        _detected_bits.append(f"State={STATE}")
    if RESPONSE_DATE:
        _detected_bits.append(f"Rec'd Date={RESPONSE_DATE}")
    if RESPONSE_TYPE:
        _detected_bits.append(f"Response Type={RESPONSE_TYPE}")
    if _detected_bits:
        print(f"✅ Detected from filename: {', '.join(_detected_bits)}")

    if not STATE:
        _gap()
        while True:
            STATE = input("State (2-letter code): ").strip().upper()
            if re.fullmatch(r"[A-Z]{2}", STATE):
                break
            print("Please enter the 2-letter state code (e.g. LA).")

    if not RESPONSE_DATE:
        _gap()
        while True:
            RESPONSE_DATE = input("Rec'd Date (MM-DD-YYYY): ").strip()
            if is_valid_date_mmddyyyy(RESPONSE_DATE):
                break
            print("Please enter a valid date in MM-DD-YYYY format (e.g. 01-15-2026).")

    if not RESPONSE_TYPE:
        RESPONSE_TYPE = ask_numbered_choice(RESPONSE_TYPE_OPTIONS, "Response type:", "Enter the number: ")

input_dir      = os.path.dirname(MCO_INPUT_FILE)
input_filename = os.path.splitext(os.path.basename(MCO_INPUT_FILE))[0]

# A resume run's MCO_INPUT_FILE is itself a Review file the user picked up from a
# previous run's own "Output" folder -- input_dir is ALREADY that Output folder, so
# appending another "Output" here would nest a nested Output/Output. On resume, write
# straight into input_dir instead; a normal/LA run still gets its own "Output" subfolder.
_output_dir = input_dir if IS_RESUME_RUN else os.path.join(input_dir, "Output")

def _next_review_stem(stem):
    """The first Review file for a ticket is named plain '..._Review'. Every later
    resume round feeds that SAME file's name back in as input_filename, so it already
    ends in '_Review'/'_ReviewN' -- incrementing a trailing number here ('_Review' ->
    '_Review1' -> '_Review2' -> ...) instead of always appending another '_Review'
    keeps the name from stacking into '..._Review_Review_Review...'."""
    m = re.search(r'_Review(\d*)$', stem)
    if m is None:
        return stem + "_Review"
    base = stem[:m.start()]
    n = int(m.group(1)) if m.group(1) else 0
    return f"{base}_Review{n + 1}"

# On a resume run, input_filename IS the prior run's own Review/WIP-style name, which
# already starts with "{TicketId} - " -- prepending TicketId again here would either
# duplicate it, or (since the real TicketId isn't recovered from the sheet until later,
# in the BUILD DF section) bake in the literal placeholder "RESUME -". A fresh
# (non-resume) run's raw input file has no such prefix yet, so it still needs one added.
if IS_RESUME_RUN:
    OUTPUT_FILE = os.path.join(_output_dir, f"{input_filename}_WIP.xlsx")
    REVIEW_FILE = os.path.join(_output_dir, f"{_next_review_stem(input_filename)}.xlsx")
else:
    OUTPUT_FILE = os.path.join(_output_dir, f"{TicketId} - {input_filename}_WIP.xlsx")
    REVIEW_FILE = os.path.join(_output_dir, f"{TicketId} - {input_filename}_Review.xlsx")

# REVIEW_FILE is a few characters longer (and "_Review12" longer still), so both are checked.
if max(len(OUTPUT_FILE), len(REVIEW_FILE) + 3) >= 255:
    shortened = input_filename
    shortened = re.sub(r'\s*Responses?\s*', ' ', shortened, flags=re.IGNORECASE).strip()
    shortened = re.sub(r"\s*-\s*\d{2}-\d{2}-\d{2}(?=\s+rec'd)", '', shortened, flags=re.IGNORECASE).strip()
    shortened = re.sub(r"\s+rec'd", " - rec'd", shortened, flags=re.IGNORECASE).strip()
    if IS_RESUME_RUN:
        OUTPUT_FILE = os.path.join(_output_dir, f"{shortened}_WIP.xlsx")
        REVIEW_FILE = os.path.join(_output_dir, f"{_next_review_stem(shortened)}.xlsx")
    else:
        OUTPUT_FILE = os.path.join(_output_dir, f"{TicketId} - {shortened}_WIP.xlsx")
        REVIEW_FILE = os.path.join(_output_dir, f"{TicketId} - {shortened}_Review.xlsx")

# Still too long for Windows after shortening: the long-path prefix keeps the writes working.
OUTPUT_FILE, REVIEW_FILE = _long_path(OUTPUT_FILE), _long_path(REVIEW_FILE)
os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)

# Attach the run log now that the output folder is known -- a down-to-the-second
# timestamp (shared with the audit copy below) keeps a later resume round writing into
# the SAME output folder from ever overwriting an earlier round's log/audit files.
_run_timestamp = datetime.now().strftime('%Y-%m-%d_%H%M%S')
_run_log_path = os.path.join(os.path.dirname(OUTPUT_FILE), f"log_{_run_timestamp}.txt")
_run_audit_path = os.path.join(os.path.dirname(OUTPUT_FILE), f"audit_{_run_timestamp}.xlsx")
_RUN_LOG_STREAM.attach_file(_run_log_path)
print(f"📝 Logging this run's full terminal output to: {_run_log_path}")

CONFIG_FILE = _resolve_synced_path(os.environ["MCO_CONFIG_FILE"])
TPA_FILE    = _resolve_synced_path(os.environ["MCO_TPA_FILE"])


print(f"✅ Config   : {CONFIG_FILE}\n")

# =====================================================
# DETECT LA FILE
# =====================================================
# STATE is known upfront (filename or prompt), so LA-format is simply "State == LA" --
# no more sniffing sheet names for an Entitlement/EV tab pair. On resume, STATE is
# None, but that branch is chosen before IS_LA_FILE is ever consulted.
if IS_RESUME_RUN:
    IS_LA_FILE = False
else:
    IS_LA_FILE = (STATE == "LA")
    print(f"✅ LA File  : {IS_LA_FILE}\n")

# =====================================================
# LOAD CONFIG FROM EXCEL
# =====================================================

def _normalize_key(text):
    if text is None:
        return None
    try:
        if pd.isna(text):
            return None
    except Exception:
        pass
    s = str(text).strip().upper()
    return s if s else None

def _stringify_id(v):
    """Excel can store an ID-like cell (CollectorId/LienholderId/TortId) as a genuine
    number instead of text -- reading it via openpyxl then gives a native int/float,
    and building a DataFrame straight from those raw values lets pandas infer a
    numeric dtype for the whole column, turning every value into a numpy.int64/
    float64. That silently breaks string-based comparisons elsewhere in the
    pipeline, and pyodbc rejects a numpy scalar outright as a SQL parameter
    ("Invalid parameter type"). Converts to a clean string once: NaN/None stays
    None (never the literal string "nan"), and a whole-number float (500.0)
    becomes "500", not "500.0"."""
    if v is None:
        return None
    try:
        if pd.isna(v):
            return None
    except (TypeError, ValueError):
        pass
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)

def _id_key(v):
    """_stringify_id plus text cleanup ("501.0" -> "501", spaces trimmed), for matching an ID
    that went through an Excel review file against a fresh SQL value. Blank -> None."""
    s = _stringify_id(v)
    if s is None:
        return None
    s = s.strip()
    if re.fullmatch(r"\d+\.0+", s):
        s = s.split(".")[0]
    return s if s and s.lower() not in ("nan", "none") else None

def _case_scope_key(case_id):
    """SCOPE_MAP key for a scope decision saved per CaseId (scope blank in LPM). None if no
    CaseId. The "CASEID:" prefix keeps it apart from real scope texts."""
    cid = _id_key(case_id)
    return f"CASEID:{cid}" if cid else None

def _is_true_flag(v):
    """IsMCO-style flag -> bool. SQL NULL/NaN and text like "False"/"0" read as False."""
    if v is None:
        return False
    if isinstance(v, str):
        return v.strip().lower() in ("true", "1", "yes", "y")
    try:
        if pd.isna(v):
            return False
    except (TypeError, ValueError):
        pass
    return bool(v)

_RE_INVISIBLE = re.compile('[\\s ​-\u200F⁠﻿]+')

def _clean_lien_id(v):
    """A LienId as typed in a file, minus spaces/hidden characters and Excel's ".0".
    Digits are never changed (a leading zero stays). "" for blank."""
    if v is None:
        return ""
    try:
        if pd.isna(v):
            return ""
    except (TypeError, ValueError):
        pass
    s = _RE_INVISIBLE.sub('', str(v))
    return re.sub(r'(?<=\d)\.0+$', '', s)

def _drop_empty_info_cols(frame, info_cols):
    """Information columns (never upload or fill-in columns) that are blank on every row
    are left out of the written sheet. "_OpenEnd" counts as blank when it is never True."""
    if frame is None or frame.empty:
        return frame
    drop = []
    for col in info_cols:
        if col not in frame.columns:
            continue
        if col == "_OpenEnd":
            blank = ~frame[col].map(_is_true_flag)
        else:
            blank = frame[col].isna() | frame[col].astype(str).str.strip().isin(["", "nan", "None", "NaT"])
        if blank.all():
            drop.append(col)
    return frame.drop(columns=drop)

YES_WITH_MCO_INFO_COLS = ["LienIds used for updates", "Stage", "Remarks", "Observation", "MCO Count"]
REVIEW_INFO_COLS = ["Observation", "_OpenEnd"]

def _join_observations(values):
    """Unique Observation texts (each may already be " | "-joined), in order; None if none."""
    out = []
    for v in values:
        if v is None or (not isinstance(v, str) and pd.isna(v)):
            continue
        for part in str(v).split(" | "):
            part = part.strip()
            if part and part.lower() not in ("nan", "none") and part not in out:
                out.append(part)
    return " | ".join(out) or None

def _patch_tpa_file_blank_fields(state_mco, tort_id, lienholder_id, collector_id, lien_type,
                                 lienholder_name=None, scope=None, new_lienholder_id=None):
    """Patches ONE blank-field row in the TPA file in place, locating it by
    (TortId, LienholderId) within the state's own sheet (matched by state-code
    prefix, same convention as load_tpa_state_sheet) -- this is the single source
    of truth this case fixes: the row already exists correctly identified in
    TPA, it's just missing CollectorId/LienType. Shared by both the resumed
    Pending_TPA_Blank_Fields batch-patch and the live in-run prompt, so both
    write via the exact same logic instead of two copies that could drift.
    With new_lienholder_id (the row's own LienholderId is blank) the row is found by
    TortId + LienholderName + Scope instead, and that LienholderId is written too.
    Returns True if the row was found and the file saved successfully."""
    if not _check_file_not_locked(TPA_FILE):
        return False
    _ensure_file_hydrated(TPA_FILE)
    try:
        tpa_wb = _retry_on_permission_error(lambda: _opxl.load_workbook(TPA_FILE), "the TPA file")
    except PermissionError:
        return False
    _tpa_loaded_fingerprint = _file_fingerprint(TPA_FILE)
    sheet_to_patch = next(
        (n for n in tpa_wb.sheetnames if n.upper().startswith(str(state_mco).split(" ")[0].upper() + " ")),
        None
    )
    if sheet_to_patch is None:
        return False
    tpa_ws = tpa_wb[sheet_to_patch]
    tpa_header = [str(c.value).strip() if c.value is not None else None for c in tpa_ws[1]]
    tpa_col_idx = {name: i + 1 for i, name in enumerate(tpa_header) if name}
    if not {"TortId", "LienholderId", "LienholderName", "Private Lien Case Scope", "CollectorId",
            "LienType"} <= set(tpa_col_idx):
        return False
    _suffix = re.compile(r"\(\s*" + r"\s+".join(re.escape(w) for w in str(state_mco).split()) + r"\s*\)", re.IGNORECASE)

    def _name_key(v):
        return _normalize_key(_suffix.sub("", str(v))) if v is not None else None

    found_row = None
    for r in range(2, tpa_ws.max_row + 1):
        if _id_key(tpa_ws.cell(row=r, column=tpa_col_idx["TortId"]).value) != _id_key(tort_id):
            continue
        _row_lh = tpa_ws.cell(row=r, column=tpa_col_idx["LienholderId"]).value
        if new_lienholder_id is not None:
            if (_id_key(_row_lh) is None
                    and _name_key(tpa_ws.cell(row=r, column=tpa_col_idx["LienholderName"]).value) == _name_key(lienholder_name)
                    and _normalize_key(tpa_ws.cell(row=r, column=tpa_col_idx["Private Lien Case Scope"]).value)
                    == _normalize_key(scope)):
                found_row = r
                break
        elif _normalize_key(_row_lh) == _normalize_key(lienholder_id):
            found_row = r
            break
    if found_row is None:
        return False
    _old_collector_id = tpa_ws.cell(row=found_row, column=tpa_col_idx["CollectorId"]).value
    _old_lien_type = tpa_ws.cell(row=found_row, column=tpa_col_idx["LienType"]).value
    if new_lienholder_id is not None:
        tpa_ws.cell(row=found_row, column=tpa_col_idx["LienholderId"], value=new_lienholder_id)
        lienholder_id = new_lienholder_id
    if pd.notna(collector_id):
        tpa_ws.cell(row=found_row, column=tpa_col_idx["CollectorId"], value=collector_id)
    if pd.notna(lien_type):
        tpa_ws.cell(row=found_row, column=tpa_col_idx["LienType"], value=lien_type)
    if _fingerprint_really_changed(_tpa_loaded_fingerprint, _file_fingerprint(TPA_FILE)):
        print(
            f"⚠️  The TPA file was changed by someone else just now (likely synced in "
            f"from another user's save) -- this patch was NOT saved, to avoid "
            f"overwriting their change. It'll be retried on a future run/resume."
        )
        return False
    try:
        _retry_on_permission_error(lambda: tpa_wb.save(TPA_FILE), "the TPA file")
    except PermissionError:
        return False
    _append_audit_trail_rows([_audit_row(
        "TPA_FILE_PATCH",
        f"State={state_mco} | TortId={tort_id} | LienholderId={lienholder_id}",
        ("LienholderId=(blank), " if new_lienholder_id is not None else "")
        + f"CollectorId={_old_collector_id if pd.notna(_old_collector_id) else '(blank)'}, "
        f"LienType={_old_lien_type if pd.notna(_old_lien_type) else '(blank)'}",
        (f"LienholderId={new_lienholder_id}, " if new_lienholder_id is not None else "")
        + f"CollectorId={collector_id if pd.notna(collector_id) else _old_collector_id}, "
        f"LienType={lien_type if pd.notna(lien_type) else _old_lien_type}",
        TicketId,
    )])
    return True

TPA_OVERRIDES_COLS = ["TortId", "TortName", "Private Lien Case Scope", "CollectorId",
                       "LienholderId", "LienholderName", "LienType",
                       "State MCO", "Ticket #", "Reason"]

# "Reason" on each new/updated TPA_Overrides row says why it exists (older rows stay blank).
TPA_REASON_USER_LHID   = "Plan not in TPA list - LienholderId entered by user"
TPA_REASON_PICKED_DUP  = "Duplicate rows in TPA - one picked by user"
TPA_REASON_LEGACY_TPA  = "LienholderId from old Hardcoded_LHID list - CollectorId/LienType taken from TPA"
TPA_REASON_LEGACY_USER = "LienholderId from old Hardcoded_LHID list - CollectorId/LienType entered by user"
TPA_REASON_NAME_TPA    = "LienholderId from TPA_Overrides (same plan name) - CollectorId/LienType taken from TPA"
TPA_REASON_NAME_USER   = "LienholderId from TPA_Overrides (same plan name) - CollectorId/LienType entered by user"
TPA_REASON_LEGACY_CONFIRMED = "LienholderId from old Hardcoded_LHID list - CollectorId/LienType suggestion confirmed by user"
TPA_REASON_NAME_CONFIRMED   = "LienholderId from TPA_Overrides (same plan name) - CollectorId/LienType suggestion confirmed by user"
TPA_REASON_DNU         = "TPA row marked DO NOT USE - used by user"

def _tpa_override_reason(pending_reason):
    """Reason for a TPA_Overrides row filled in from Pending_TPA_Review, from its PendingReason."""
    t = str(pending_reason or "").strip().lower()
    for needle, reason in (
        ("legacy", TPA_REASON_LEGACY_USER),
        ("state not found", "State not in TPA"),
        ("tort not found", "Tort not in TPA"),
        ("multiple", "Duplicate rows in TPA"),
        ("conflicts with", "Conflicts with another TPA_Overrides row"),
        ("scope", "Scope not in TPA"),
        ("not found in tpa", "Plan not found in TPA"),
    ):
        if needle in t:
            return reason
    return "Filled in Pending_TPA_Review"

_RE_ANY_SPACE_RUN = re.compile(r"\s+")

def _clean_headers(cols):
    """Header names as trimmed text (a number/date header too). A name that repeats only
    because of spacing ("MCO Plan" / "MCO Plan ") gets " (2)", " (3)" so every column
    stays a single column."""
    out, seen = [], {}
    for c in cols:
        name = "" if c is None or (isinstance(c, float) and pd.isna(c)) else str(c).strip()
        n = seen.get(name, 0)
        seen[name] = n + 1
        out.append(f"{name} ({n + 1})" if n else name)
    return out

def _override_bad_fields(entry):
    """Cleans one TPA_Overrides row in place (IDs "300.0" -> "300", LienType to its exact
    option) and returns {field: value as typed} for values that can't be used: a
    LienholderId/CollectorId that isn't a number, a LienType that is neither option.
    Blank values are not "bad" (a legacy row has no CollectorId/LienType)."""
    bad = {}
    for col in ("LienholderId", "CollectorId"):
        raw = entry.get(col)
        if raw is None or str(raw).strip() == "":
            entry[col] = None
            continue
        clean = _id_key(raw)
        if clean is not None and re.fullmatch(r"\d+", clean):
            entry[col] = clean
        else:
            bad[col] = raw
    raw_lt = entry.get("LienType")
    if raw_lt is not None and str(raw_lt).strip():
        canon = _canonical_lien_type(raw_lt)
        if canon:
            entry["LienType"] = canon
        else:
            bad["LienType"] = raw_lt
    tort_id = _id_key(entry.get("TortId"))
    entry["TortId"] = tort_id
    return bad

def _override_row_label(entry):
    """'TPA_Overrides row 14' for a config row, or 'a TPA_Overrides entry made this run'."""
    return f"TPA_Overrides row {entry['_row']}" if entry.get("_row") else "a TPA_Overrides entry made this run"

def load_config(config_path):
    # One workbook parse serves all 5 sheets below -- pd.read_excel(config_path, ...)
    # re-parses the ENTIRE file from scratch on every call when given a bare path, so
    # calling it 5x (once per sheet) was silently doing 5 full workbook reads instead
    # of 1. Passing this shared ExcelFile object instead reuses the one parse.
    _ensure_file_hydrated(config_path)
    xls = _retry_on_permission_error(lambda: pd.ExcelFile(config_path), "mco_config.xlsx")

    def _read_cfg_sheet(sheet_name):
        """A config sheet as text; a missing or empty sheet is an empty table (a missing one
        is created when a decision is first saved to it), headers trimmed."""
        if sheet_name not in xls.sheet_names:
            print(f"⚠️  mco_config.xlsx has no '{sheet_name}' sheet -- treated as empty "
                  f"(it is created when a decision is saved to it).")
            return pd.DataFrame()
        df = _retry_on_permission_error(
            lambda: pd.read_excel(xls, sheet_name=sheet_name, dtype=str), "mco_config.xlsx"
        )
        df.columns = _clean_headers(df.columns)
        return df

    def _load_sheet(sheet_name, col_keys):
        df = _read_cfg_sheet(sheet_name)
        result = {}
        for _, row in df.iterrows():
            # Spaces collapsed: "MOLINA  HEALTHCARE" and "MOLINA HEALTHCARE" are one key.
            vals = [_RE_ANY_SPACE_RUN.sub(" ", str(row.get(k)).strip().upper()) for k in col_keys]
            if any(v in ("NAN", "NONE") for v in vals):
                continue
            state = vals[0]
            if len(state) != 2:
                continue
            result.setdefault(state, {})[vals[1]] = vals[2]
        return result

    def _load_scope_map(sheet_name):
        df = _read_cfg_sheet(sheet_name)
        result = {}
        for _, row in df.iterrows():
            key = _normalize_key(row.get("ScopeText"))
            if key is None:
                # A blank-scope decision saved per case (ScopeText blank, CaseId filled).
                key = _case_scope_key(row.get("CaseId"))
                if key is None:
                    continue
            result[key] = _normalize_key(row.get("Classification"))
        return result

    def _load_tpa_overrides(sheet_name):
        df = _read_cfg_sheet(sheet_name)
        rows = []
        for _i, row in df.iterrows():
            if pd.isna(row.get("LienholderName")) and pd.isna(row.get("LienholderId")):
                continue
            entry = {c: (row.get(c) if pd.notna(row.get(c)) else None) for c in TPA_OVERRIDES_COLS}
            # Excel row number (header is row 1), so a message can name the row to fix.
            entry["_row"] = _i + 2
            entry["_bad_fields"] = _override_bad_fields(entry)
            rows.append(entry)
        return rows

    def _load_lien_equivalence_map(sheet_name):
        """Confirmed (once, by a human, via a resolved Pending_Lien_Match_Review row)
        facts about the relationship between a TPA-resolved plan (State MCO +
        CollectorId + LienholderId) and an EXISTING lien's own on-file name -- e.g.
        TPA's "HealthShare/CareOregon" is confirmed to appear on file as "CareOregon
        Advantage" (Decision="Yes"), or confirmed NOT to be the same real plan as
        some other existing lien name it coincidentally ties against
        (Decision="No"). These are facts about the PLAN/BRAND, not about any one
        client, so a decision made once applies to every future client with the
        same resolved plan -- no ClientId in this sheet.

        Returns (equivalence_map, excluded_map), both keyed by
        (State MCO, CollectorId, LienholderId) -> list of names:
        - equivalence_map ("Yes" rows): extra names classify_step4_scenario tries,
          in addition to the TPA-resolved name itself.
        - excluded_map ("No" rows): existing-lien names classify_step4_scenario
          excludes from consideration entirely for that plan, before matching.
        """
        df = _read_cfg_sheet(sheet_name)
        equivalence_map, excluded_map = {}, {}
        # One decision per (plan, name); a later row wins, so a newer YES replaces an
        # older NO (and back) instead of the NO silently winning forever.
        decisions = {}
        for _, row in df.iterrows():
            state_mco     = row.get("State MCO")
            collector_id  = row.get("CollectorId")
            lienholder_id = row.get("LienHolderId")
            equiv_name    = row.get("EquivalentName")
            decision      = row.get("Decision")
            if (pd.isna(state_mco) or pd.isna(collector_id) or pd.isna(lienholder_id)
                    or pd.isna(equiv_name) or pd.isna(decision)):
                continue
            decision_key = str(decision).strip().upper()
            if decision_key not in ("YES", "NO"):
                continue
            # Keyed by (State, CollectorId, LienholderId) -- NOT just (State,
            # LienholderId). A LienholderId number can in principle be reused by TPA
            # across more than one Collector relationship (different real plans that
            # happen to share a number) -- scoping to the exact Collector this
            # confirmation was made under means a coincidental same-text name clash
            # under a DIFFERENT Collector is never silently treated as the same plan.
            key = (str(state_mco).strip(), str(collector_id).strip(), str(lienholder_id).strip())
            decisions[(key, _normalize_key(equiv_name))] = (decision_key, str(equiv_name).strip())
        for (key, _name_key), (decision_key, name) in decisions.items():
            target = equivalence_map if decision_key == "YES" else excluded_map
            target.setdefault(key, [])
            if name not in target[key]:
                target[key].append(name)
        return equivalence_map, excluded_map

    def _load_column_map(sheet_name):
        """Confirmed (once, by a human) column-name mappings for a State's input-file
        format: which column holds LienId, and which column(s) hold the free-text MCO
        plan info. Keyed by State (2-letter, from STATE -- known upfront now, before
        column identification even runs), value is a LIST of (lienid_col, mco_cols)
        entries -- a state can accumulate more than one over time if its vendor's file
        format changes, so every saved entry for a state is tried (matched by column
        NAME actually present in the file, never by position -- see
        _try_column_map_lookup), not just the latest one."""
        df = _read_cfg_sheet(sheet_name)
        result = {}
        for _, row in df.iterrows():
            state = row.get("State")
            lienid_col = row.get("LienId_Column")
            mco_cols_raw = row.get("MCO_Columns")
            if pd.isna(state) or pd.isna(lienid_col) or pd.isna(mco_cols_raw):
                continue
            mco_cols = [c.strip() for c in str(mco_cols_raw).split(",") if c.strip()]
            if not mco_cols:
                continue
            state_key = str(state).strip().upper()
            result.setdefault(state_key, []).append((str(lienid_col).strip(), mco_cols))
        return result

    alias_map           = _load_sheet("Alias_Map", ["State", "From_Name", "To_Name"])
    scope_map           = _load_scope_map("Scope_Map")
    tpa_overrides        = _load_tpa_overrides("TPA_Overrides")
    lien_equivalence_map, lien_excluded_map = _load_lien_equivalence_map("Lien_Name_Equivalence_Map")
    column_map          = _load_column_map("Column_Map")
    xls.close()

    return alias_map, scope_map, tpa_overrides, lien_equivalence_map, lien_excluded_map, column_map

ALIAS_MAP, SCOPE_MAP, TPA_OVERRIDES, LIEN_NAME_EQUIVALENCE_MAP, LIEN_NAME_EXCLUDED_MAP, COLUMN_MAP = load_config(CONFIG_FILE)
print("✅ Config loaded successfully\n")

def _config_value_warnings(tpa_overrides):
    """One line per TPA_Overrides value that can't be used as it is (see _override_bad_fields)."""
    lines = []
    for e in tpa_overrides:
        for col, raw in (e.get("_bad_fields") or {}).items():
            what = (f"is neither {' nor '.join(repr(o) for o in LIEN_TYPE_OPTIONS)}" if col == "LienType"
                    else "is not a number")
            lines.append(f"{_override_row_label(e)} for {e.get('LienholderName')!r}: {col} {raw!r} {what}. "
                         f"Fix it in mco_config.xlsx.")
    return lines

_CONFIG_WARNINGS = _config_value_warnings(TPA_OVERRIDES)
if _CONFIG_WARNINGS:
    print("⚠️  Some TPA_Overrides values can't be used as they are. If a plan needs one of these rows,")
    print("   the tool will ask for the correct value (for this run only):")
    for _cw in _CONFIG_WARNINGS:
        print(f"   {_cw}")
    print()

# Name-level YES/NO decisions from a resumed Pending_Lien_Match_Review, persisted to
# mco_config.xlsx (Lien_Name_Equivalence_Map) only if the user says so at the end.
_RESUME_LIEN_EQUIV_SAVES = []

# The same review answers for THIS run: they apply to that one client's plan only, so
# a "n" at the save question never touches other clients. Keyed by (ClientId,
# CollectorId, LienholderId) -> {("ID", LienId) | ("NAME", normalized lien name)}.
_LIEN_MATCH_PICKS = {}
_LIEN_MATCH_CLIENT_EXCLUDED = {}

# Review reasons about one client's look-alike liens: the lien name can't tell them
# apart, so their answer is never saved as a name decision for future runs.
_LIEN_REVIEW_NO_SAVE_REASONS = {"Duplicate liens", "Duplicate placeholders"}

# Newly-confirmed (State, LienId_Column, MCO_Columns) mappings this run, queued for
# the same end-of-run batched config save as everything else -- see
# _try_column_map_lookup for how COLUMN_MAP is consulted before ever asking.
_COLUMN_MAP_SAVES = []

def _try_column_map_lookup(state, headers):
    """Returns (lienid_col, mco_cols) from a saved COLUMN_MAP entry for this state
    whose column names are ALL actually present in `headers` (this file's own real
    column list) -- matched by name, never by position, so column re-ordering in a
    later file of the same format doesn't matter. Tried in save-order; the first
    entry whose columns all exist wins. None if no saved entry for this state fits
    (a brand-new format, or no state at all), so the normal prompts still fire."""
    if not state:
        return None
    header_set = set(headers)
    for lienid_col, mco_cols in COLUMN_MAP.get(str(state).strip().upper(), []):
        if lienid_col in header_set and all(c in header_set for c in mco_cols):
            return lienid_col, mco_cols
    return None

# =====================================================
# PRE-COMPILE ALL REGEX
# =====================================================

_RE_NY_PREFIX      = re.compile(r'^[NY]/')
_RE_AMP            = re.compile(r'&')
_RE_AND_WORD       = re.compile(r'\bAND\b')
_RE_LONG_NUM       = re.compile(r'\b\d[\d\-]{6,}\b(?!/)')
_RE_NON_DATE_SLASH = re.compile(r'(?<!\d)/(?!\d)')
_RE_DASH           = re.compile(r'-')
_RE_MULTI_SPACE    = re.compile(r'\s+')
_RE_HAS_ALPHA      = re.compile(r'[A-Z]')
_RE_STRIP_SPACE    = re.compile(r'\s+')
_RE_ALPHA_IN_PLAN  = re.compile(r'[A-Z]')

# Plan-text date handling (normal files). Every date shape is rewritten to M/D/Y first
# so one tokenizer reads them all; a date-shaped value that still can't be read becomes
# an Observation, never part of a plan name and never "Bad Date".
_RE_LONG_DASH   = re.compile('[‐-―−]')
_RE_ISO_DATE    = re.compile(r'(?<![\d/])(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T]\d{1,2}:\d{2}(?::\d{2}(?:\.\d+)?)?)?(?![\d/])')
_RE_DASH_DATE   = re.compile(r'(?<![\d/])(\d{1,2})-(\d{1,2})-(\d{4}|\d{2})(?![\d/])')
_RE_DOT_DATE    = re.compile(r'(?<![\d.])(\d{1,2})\.(\d{1,2})\.(\d{4})(?![\d.])')
_RE_YEAR_RANGE  = re.compile(r'\b((?:19|20)\d{2})\s*-\s*((?:19|20)\d{2})\b')
_RE_RANGE_WORD  = re.compile(r'\b(?:THRU|THROUGH|TILL|UNTIL)\b')
_OPEN_END_WORDS = ("CURRENT", "PRESENT", "ONGOING", "NOW", "OPEN")
# Full date, month-only (MM/YYYY), no-year (MM/DD) and a bare year -- the last two can
# never be read, but are still recognised so they never end up inside a plan name.
_DATE_TOKEN_RX  = r'\d{1,2}/\d{1,2}/\d+|\d{1,2}/\d{4}|\d{1,2}/\d{1,2}|(?<!\d)(?:19|20)\d{2}(?!\d)'
_DATE_LINK_RX   = r'(?:\s*(?:-|TO|\*)\s*|\s+)'
# Separators: , ; new line, 2+ spaces, a standalone "." (not "ST."), list numbering "2." / "2)".
_RE_PLAN_TOKEN  = re.compile(
    rf'(?P<range>(?P<rs>{_DATE_TOKEN_RX})(?:{_DATE_LINK_RX}(?P<re>{_DATE_TOKEN_RX})'
    rf'|{_DATE_LINK_RX}(?P<ro>{"|".join(_OPEN_END_WORDS)})\b|\s*TO\s*(?P<rd>DATE)\b))'
    rf'|(?P<date>{_DATE_TOKEN_RX})'
    r'|(?P<sep>[,;\n]|\s{2,}|(?<![A-Z])\.(?=\s|$)|(?:(?<=\s)|^)\d{1,2}[.)](?=\s*[A-Z]))'
)
# Any date-like shape in a raw cell (has_any_plan lets date-only cells through to review).
_RE_ANY_DATE_SHAPE = re.compile(r'\d{1,2}[/\-.]\d{1,2}[/\-.]\d{2,4}|\d{4}-\d{1,2}-\d{1,2}|\d{1,2}/\d{4}')

# =====================================================
# INVALID SET
# =====================================================

# YES/Y/MCO/TO are never plan names (a "YES" flag or stray connector, not a payer).
INVALID = {
    "N", "NO", "NA", "NONE", "N/A", "A", "CASE", "PD", "PAID", "CLAIMS",
    "RELATED", "NOT", "TMHP", "COLLECTS", "FOR", "THIS", "DOL",
    "YES", "Y", "MCO", "TO"
}

# Words written before a plan name ("CLIENT HAS MOLINA") -- removed from the START only.
PLAN_LEAD_IN_WORDS = {
    "CLIENT", "CLAIMANT", "PATIENT", "PT", "HAS", "HAD", "HAVE", "IS", "WAS", "ARE", "WERE",
    "ENROLLED", "COVERED", "INSURED", "IN", "WITH", "BY", "UNDER", "CURRENTLY", "PREVIOUSLY", "ALSO",
}
# Removed only when what is left is a known plan name ("ACTIVE HEALTH" may be a plan).
PLAN_LEAD_IN_IF_KNOWN = ("MEMBER OF", "ACTIVE", "ELIGIBLE")
# A whole text that says there is no MCO plan. "FFS" is not here: short forms can be plans.
NOT_A_PLAN_PHRASES = {
    "FEE FOR SERVICE", "STRAIGHT MEDICAID", "MEDICAID", "MEDICARE", "NO MCO", "NOT ENROLLED",
    "UNKNOWN", "DONT KNOW", "DON'T KNOW", "NOT SURE",
}
# Any of these words means the text is a non-answer, not a plan ("CLIENT FORGOT").
NON_ANSWER_WORDS = {"DONT", "DON'T", "DOESNT", "DOESN'T", "FORGOT", "UNKNOWN", "UNSURE", "IDK"}

def _not_a_plan_text(name):
    """True when a plan text says there is no plan (NOT_A_PLAN_PHRASES / NON_ANSWER_WORDS)."""
    t = _RE_ANY_SPACE_RUN.sub(" ", str(name).upper()).strip(" .,;:")
    return t in NOT_A_PLAN_PHRASES or any(w in NON_ANSWER_WORDS for w in t.split())

def _strip_lead_in(name, is_known=None):
    """(name without the lead-in words at its start, removed text or None). A name that
    is a known plan as written is never changed; ACTIVE/ELIGIBLE/MEMBER OF go only when
    the rest is a known plan. A text made only of lead-in words comes back as ""."""
    if is_known is not None and is_known(name):
        return name, None
    words = name.split()
    i = 0
    while i < len(words) and words[i].strip(".,:;") in PLAN_LEAD_IN_WORDS:
        i += 1
    rest = " ".join(words[i:])
    if is_known is not None and rest:
        for lead in PLAN_LEAD_IN_IF_KNOWN:
            if rest.startswith(lead + " ") and is_known(rest[len(lead) + 1:].strip()):
                rest = rest[len(lead) + 1:].strip()
                break
    if rest == name:
        return name, None
    removed = name[:len(name) - len(rest)].strip() if rest else name
    return rest, removed

# =====================================================
# SHARED TEXT NORMALIZER
# =====================================================

def _normalize_mco_text(val: str) -> str:
    val = _RE_NY_PREFIX.sub('', val.strip())
    val = _RE_AMP.sub(',', val)
    val = _RE_AND_WORD.sub(',', val)
    val = _RE_LONG_NUM.sub('', val)
    val = _RE_NON_DATE_SLASH.sub(',', val)
    val = _RE_DASH.sub(' ', val)
    val = _RE_MULTI_SPACE.sub(' ', val).strip()
    return val

# =====================================================
# HELPER FUNCTIONS
# =====================================================

def _valid_dates(dates):
    """Real dates only: None/NaN/NaT and any text (e.g. an old "Bad Date") are skipped."""
    out = []
    for d in dates:
        if d is None or isinstance(d, str):
            continue
        try:
            if pd.isna(d):
                continue
        except (TypeError, ValueError):
            pass
        out.append(d)
    return out

# "Bad Date" is never produced: no readable date -> None (blank, so the lien's own
# on-file date is never overwritten).
def safe_min(dates):
    valid = _valid_dates(dates)
    return min(valid) if valid else None

def safe_max(dates):
    valid = _valid_dates(dates)
    return max(valid) if valid else None

def format_date(x):
    if x is None:
        return None
    if isinstance(x, str):
        x = x.strip()
        parsed = _parse_date(x) if x else None
        return parsed.strftime("%m/%d/%Y") if isinstance(parsed, datetime) else None
    try:
        if pd.isna(x):
            return None
    except Exception:
        pass
    return x.strftime("%m/%d/%Y")

def make_sql_in(id_list):
    unique = list(dict.fromkeys(str(i) for i in id_list))
    if not unique:
        # Real bug found: an empty list fell through to f" {tuple([])}" == " ()",
        # which is invalid SQL ("ClientId in  ()") -- the next pd.read_sql() then
        # crashed with a raw, confusing pyodbc error instead of an explained one.
        raise ValueError(
            "make_sql_in() got an empty id list -- nothing to query. Check that the "
            "LienId/ClientId column was identified correctly from the input file."
        )
    if len(unique) == 1:
        return f" ('{unique[0]}')"
    return f" {tuple(unique)}"

def _drop_lien_ids_not_in_lpm(plan_df, lpm_lien_ids, client_col=None):
    """Plan rows whose LienId LPM doesn't have (or, with client_col, that came back with no
    client) are skipped with a warning and Remarks "LienId not found in LPM"."""
    known = {_clean_lien_id(v) for v in lpm_lien_ids}
    keep = plan_df["LienId"].map(_clean_lien_id).isin(known)
    if client_col is not None and client_col in plan_df.columns:
        keep &= plan_df[client_col].notna() & ~plan_df[client_col].astype(str).str.strip().str.lower().isin(
            ("", "nan", "none"))
    missing = sorted({_clean_lien_id(v) for v in plan_df.loc[~keep, "LienId"]}, key=lambda x: (len(x), x))
    if missing:
        _gap()
        print(f"⚠️  {len(missing)} LienId(s) not found in LPM were skipped (see Remarks): {', '.join(missing)}")
        for lid in missing:
            _add_remark(LIEN_REMARKS_EXTRA, lid, "LienId not found in LPM")
    return plan_df[keep].copy()

def _set_lienid_column(frame, lienid_col, mco_cols):
    """Renames the chosen LienId column to "LienId". Another column the file already calls
    "LienId" keeps its data as "LienId (in file)" (two columns with one name crashed the
    run). Returns (frame, mco_cols with any such rename applied)."""
    renames = {}
    for c in frame.columns:
        if c == "LienId" and c != lienid_col:
            renames[c] = "LienId (in file)"
    renames[lienid_col] = "LienId"
    frame = frame.rename(columns=renames)
    return frame, [renames.get(c, c) if c != lienid_col else c for c in mco_cols]

def _drop_rows_without_lienid(plan_df, mco_cols, lienid_col="LienId"):
    """Returns (rows whose LienId is a plain whole number, one warning line per other
    row). A blank or text LienId can't be looked up, and inside the SQL IN-list it
    crashes the whole run (real case: a footer note typed under the MCO column with no
    File ID). Row numbers assume the header sits on Excel row 1. Kept rows carry the
    cleaned LienId (spaces/hidden characters/Excel ".0" removed, digits untouched)."""
    cleaned = plan_df[lienid_col].map(_clean_lien_id)
    valid = cleaned.str.fullmatch(r"\d+").fillna(False).astype(bool)
    warnings = []
    for idx, row in plan_df[~valid].iterrows():
        text = " | ".join(str(row[c]).strip() for c in mco_cols if pd.notna(row[c]) and str(row[c]).strip())
        if len(text) > 80:
            text = text[:77] + "..."
        lien_val = "" if pd.isna(row[lienid_col]) else str(row[lienid_col]).strip()
        warnings.append(f"Excel row {idx + 2}: LienId={lien_val!r}, MCO text={text!r}")
    kept = plan_df[valid].copy()
    kept[lienid_col] = cleaned[valid]
    return kept, warnings

def bring_in_columns(target_df, source_df, key, columns):
    """Simple-lookup bring-in: pull `columns` from source_df onto target_df, keyed on `key`
    (same column name expected in both). Dedupes source_df by taking the first non-null
    value per key. This is the one shared way every straightforward column lookup into
    yes_with_mco_df should be done — add a new column name to the `columns` list (or a
    new call with a different `key`). The key column is cast to str on both sides first,
    since raw SQL-fetched id columns often come back as int64 while the rest of the
    pipeline treats ids as strings — merging mismatched dtypes raises a ValueError."""
    target_df = target_df.copy()
    source_df = source_df.copy()
    target_df[key] = target_df[key].astype(str)
    source_df[key] = source_df[key].astype(str)
    src = (
        source_df.dropna(subset=columns, how="all")
        .groupby(key, as_index=False)[columns]
        .first()
    )
    return target_df.merge(src, on=key, how="left")

# =====================================================
# ZANTAC — PERSONID SIBLING CLIENTID LOOKUP (Open Item #9)
# =====================================================
# Used by build_client_cnt_update below to expand a Zantac client's updated
# RequiredMedicaidMCOLienCount to every sibling ClientId (same PersonId + truncated
# CaseName), since Zantac's own case structure can split one real person across
# several ClientIds.

def _truncate_casename(name):
    """'Zantac - Freese & Goss - All Defendants' -> 'Zantac - Freese & Goss'
    (keep only the first two ' - '-delimited segments)."""
    if pd.isna(name):
        return None
    parts = str(name).split(" - ")
    return " - ".join(parts[:2])

def _zantac_sibling_groups(df1, zantac_client_counts):
    """df1: result of Query 1 (ClientId, PersonId, CaseName) for the original Zantac
    ClientIds. zantac_client_counts: {ClientId: target_count} for those same ClientIds.
    Returns a list of (PersonId, truncated_CaseName, count_to_apply) — one per distinct
    (PersonId, truncated CaseName) pair found, since different original Zantac ClientIds
    can belong to different cases even under the same person."""
    df1 = df1.copy()
    df1["CaseName_truncated"] = df1["CaseName"].apply(_truncate_casename)

    groups = []
    seen_pairs = set()
    for _, row in df1.iterrows():
        person_id = row["PersonId"]
        casename_trunc = row["CaseName_truncated"]
        client_id = str(row["ClientId"])
        if pd.isna(person_id) or casename_trunc is None:
            continue
        pair = (person_id, casename_trunc)
        if pair in seen_pairs:
            continue
        seen_pairs.add(pair)
        count_for_group = zantac_client_counts.get(client_id)
        if count_for_group is None:
            continue
        groups.append((person_id, casename_trunc, count_for_group))
    return groups

def get_zantac_sibling_counts(zantac_client_counts):
    """zantac_client_counts: {ClientId: target_RequiredMedicaidMCOLienCount} for clients
    whose Tort is Zantac. Returns the same dict expanded to include every sibling ClientId
    (same PersonId + truncated CaseName), each inheriting the count from whichever original
    ClientId seeded that (PersonId, CaseName) group.
    Edge case, not a bug: if two original ClientIds ever shared the same (PersonId,
    CaseName) group with different target counts, whichever is encountered first
    would win. This doesn't occur in real data, so it's left as-is rather than
    guessed at further -- revisit only if it's ever actually observed."""
    if not zantac_client_counts:
        return {}

    # DataClientView's own id column is "Id", not "ClientId" -- selecting a bare
    # "ClientId" raises "Invalid column name 'ClientId'". Aliased back to ClientId
    # so the rest of this function doesn't need to change.
    ids_clause = make_sql_in(list(zantac_client_counts.keys()))
    q1 = f"select Id as ClientId, PersonId, CaseName from DataClientView where id in {ids_clause}"
    df1 = _read_sql(q1)
    df1["ClientId"] = df1["ClientId"].astype(str)

    expanded = dict(zantac_client_counts)
    for person_id, casename_trunc, count in _zantac_sibling_groups(df1, zantac_client_counts):
        # Parameterized (not f-string-interpolated like the line above used to be) --
        # casename_trunc comes from a real, free-text DB column, and a case name
        # containing an apostrophe (e.g. a party named "O'Brien") would otherwise break
        # this query's SQL syntax outright and crash the whole run. Behaves identically
        # to the old interpolated version for any name WITHOUT special characters.
        q2 = """select Id as ClientId, PersonId, CaseName from DataClientView
        where PersonId in (?) and casename like ?"""
        df2 = _read_sql(q2, params=[person_id, f"%{casename_trunc}%"])
        for sibling_id in df2["ClientId"].astype(str).unique():
            expanded[sibling_id] = count
    return expanded

# =====================================================
# INTERACTIVE COLUMN SELECTION HELPERS
# =====================================================

def _preview_columns(df, n=5, max_len=40):
    print("\nAvailable columns:")
    for idx, col in enumerate(df.columns, start=1):
        vals = df[col].dropna().astype(str).str.strip()
        vals = vals[vals != ""]
        examples = list(pd.unique(vals))[:n]
        examples = [re.sub(r'\s+', ' ', v) for v in examples]
        examples = [v if len(v) <= max_len else v[:max_len] + "…" for v in examples]
        if examples:
            print(f"  [{idx:>2}] {col!r:<30}  e.g. {' | '.join(examples)}")
        else:
            print(f"  [{idx:>2}] {col!r:<30}  (no non-null values)")
    print()

def _resolve_columns(df, raw_input):
    cols = list(df.columns)
    tokens = [t.strip() for t in raw_input.split(",") if t.strip()]
    resolved = []
    for t in tokens:
        match = None
        if t.isdigit():
            i = int(t)
            if 1 <= i <= len(cols):
                match = cols[i - 1]
            else:
                raise ValueError(f"Column number {i} is out of range (1-{len(cols)}).")
        else:
            for c in cols:
                if str(c).strip().upper() == t.upper():
                    match = c
                    break
            if match is None:
                raise ValueError(f"No column found matching '{t}'.")
        if match not in resolved:
            resolved.append(match)
    return resolved

def _ask_columns(df, prompt_text, allow_empty=False, multi=True, show_list=True):
    if show_list:
        _preview_columns(df)
    while True:
        raw = input(prompt_text).strip()
        if not raw:
            if allow_empty:
                return []
            print("  Please enter at least one column.")
            continue
        try:
            resolved = _resolve_columns(df, raw)
        except ValueError as e:
            print(f"  Invalid input: {e} Please try again.")
            continue
        if not resolved:
            print("  Please enter at least one column.")
            continue
        if not multi and len(resolved) > 1:
            print("  Please enter exactly one column for this field.")
            continue
        return resolved

def _ask_detected_fallback(options, what_missing, prompt_text):
    """Shared shape for every LA-file auto-detect fallback (a tab name, an SSN
    column, the Healthy Louisiana column start, the Entitlement header row): the
    expected pattern wasn't found in this file -- most likely either the wrong/a
    modified file was fed in, or the file's real structure has genuinely
    changed. Prints a caution either way, then lets the user manually pick from
    what's actually in the file, or cancel outright with 'C' rather than force
    a guess against a file that might be the wrong one. Returns (index, value)
    so callers needing a position (e.g. hl_col_start) and callers needing the
    actual value (e.g. a column name) both have what they need."""
    print(
        f"\n⚠️  {what_missing}\n"
        f"    Please double-check you've attached the correct/original file.\n"
        f"    If this IS the right file and its structure has genuinely changed,\n"
        f"    pick the right one below -- or type 'C' to cancel this run.\n"
    )
    for i, opt in enumerate(options, start=1):
        print(f"  [{i}] {opt}")
    while True:
        answer = input(prompt_text).strip()
        if answer.upper() == "C":
            raise SystemExit("Run cancelled by user.")
        if answer.isdigit() and 1 <= int(answer) <= len(options):
            return int(answer) - 1, options[int(answer) - 1]
        print(f"Please enter a number between 1 and {len(options)}, or 'C' to cancel.")

# =====================================================
# EXTRACT PLAN ROWS (Normal files)
# =====================================================

def _clean_plan(raw: str) -> str:
    """Plan-name cleanup -- the one place stray characters are removed: leading AND /
    punctuation / list numbering, a leading YES/Y/N flag word ("YES - X", "Y: X"), a
    leading "MCO:" label, and trailing punctuation. A dash inside a name is kept
    ("WELL-CARE"); "PASSPORT-PASSPORT" (same name twice) becomes "PASSPORT"."""
    p = _RE_MULTI_SPACE.sub(' ', raw).strip()
    p = re.sub(r'^(?:AND\b\s*|[.,;:/*\-\s])+', '', p)
    p = re.sub(r'^\d+[.)]\s*', '', p)
    p = re.sub(r'^(?:YES|Y|N)\b\s*[-:/]?\s*(?=[A-Z])', '', p)
    p = re.sub(r'^MCO\s*[-:]\s*(?=[A-Z])', '', p)
    p = re.sub(r'[\s.,;:/*\-]+$', '', p)
    p = re.sub(r'^[\s.,;:/*\-]+', '', p).strip()
    if '-' in p:
        parts = [x for x in (s.strip() for s in p.split('-')) if x]
        if len(set(parts)) == 1:
            p = parts[0]
    return p

def _is_valid_plan(plan: str) -> bool:
    if not plan:
        return False
    if not _RE_ALPHA_IN_PLAN.search(plan):
        return False
    if set(w for w in re.split(r'[\s\-]+', plan) if w).issubset(INVALID):
        return False
    return True

def _parse_date(s: str):
    try:
        return datetime.strptime(s, "%m/%d/%Y")
    except Exception:
        return "Bad Date"

def _normalize_plan_text(text):
    """Upper-cases a cell and rewrites every date shape to M/D/Y (ISO, dashed, dotted),
    THRU/THROUGH/TILL/UNTIL to TO, long dashes to "-", AND/& to commas; phone-number-like
    runs and stray (non-date) slashes go, as before."""
    t = str(text).upper()
    t = _RE_LONG_DASH.sub('-', t)
    t = _RE_NY_PREFIX.sub('', t.strip())
    t = _RE_ISO_DATE.sub(lambda m: f"{int(m.group(2))}/{int(m.group(3))}/{m.group(1)}", t)
    t = _RE_DASH_DATE.sub(lambda m: f"{m.group(1)}/{m.group(2)}/{m.group(3)}", t)
    t = _RE_DOT_DATE.sub(lambda m: f"{m.group(1)}/{m.group(2)}/{m.group(3)}", t)
    t = _RE_YEAR_RANGE.sub(r'\1 - \2', t)
    t = _RE_RANGE_WORD.sub('TO', t)
    t = _RE_AMP.sub(',', t)
    t = _RE_AND_WORD.sub(',', t)
    t = _RE_LONG_NUM.sub('', t)
    t = _RE_NON_DATE_SLASH.sub(',', t)
    return t

def _read_text_date(s, is_end=False):
    """(datetime or None, Observation or None) for one date token. MM/YYYY is a whole
    month: the 1st as a start, the last day as an end. A 2-digit year is never guessed
    ("24" could be 1924 or 2024): not readable, like any other bad date."""
    s = s.strip()
    unreadable = f"{'End' if is_end else 'Start'} date not readable: '{s}' - not used"
    m = re.fullmatch(r'(\d{1,2})/(\d{4})', s)
    if m:
        month, year = int(m.group(1)), int(m.group(2))
        if 1 <= month <= 12:
            d = datetime(year, month, calendar.monthrange(year, month)[1] if is_end else 1)
            return d, f"Month-only date '{s}' read as {d:%m/%d/%Y}"
        return None, unreadable
    if re.fullmatch(r'\d{1,2}/\d{1,2}/\d{4}', s):
        try:
            return datetime.strptime(s, "%m/%d/%Y"), None
        except ValueError:
            pass
    return None, unreadable

def _tokenize_plan_text(t, is_known=None, not_plans=None):
    """Normalized cell text -> ordered items: {"kind": "P", "name"} for each valid plan
    name, {"kind": "D", ...} for each date or date range; "piece" counts separators
    passed so far (items sharing a piece were written with no separator between);
    "hard" counts only real punctuation separators (not runs of spaces).
    Lead-in words ("CLIENT HAS ...") are removed from a name's start (+ Observation); a
    text saying there is no plan ("FEE FOR SERVICE") is no plan and goes to not_plans.
    is_known(name) -> True for a known plan name (never changed)."""
    items, piece, hard, pos = [], 0, 0, 0

    def _add_text(chunk):
        name = _clean_plan(chunk)
        if not name:
            return
        obs = []
        if not (is_known is not None and is_known(name)):
            stripped, removed = _strip_lead_in(name, is_known)
            if removed:
                if not stripped:
                    return
                obs.append(f"Removed '{removed}' from the plan text")
                name = _clean_plan(stripped)
            if _not_a_plan_text(name):
                if not_plans is not None and name not in not_plans:
                    not_plans.append(name)
                return
        if _is_valid_plan(name):
            items.append({"kind": "P", "name": name, "piece": piece, "hard": hard, "obs": obs})

    for m in _RE_PLAN_TOKEN.finditer(t):
        _add_text(t[pos:m.start()])
        if m.group("sep") is not None:
            piece += 1
            if m.group("sep").strip(" \t"):
                hard += 1
        elif m.group("range") is not None:
            items.append({"kind": "D", "piece": piece, "hard": hard, "start": m.group("rs"), "end": m.group("re"),
                          "open": m.group("ro") or ("TO DATE" if m.group("rd") else None), "single": False})
        else:
            items.append({"kind": "D", "piece": piece, "hard": hard, "start": m.group("date"), "end": None,
                          "open": None, "single": True})
        pos = m.end()
    _add_text(t[pos:])
    for i, it in enumerate(items):
        it["pos"] = i
    return items

def _date_reading(seq, glued, loose, plan_first):
    """One way of owning the loose dates: plan-first (each goes to the plan before it)
    or date-first (the plan after it). {date pos: plan pos}, or None when unsafe.
    The dates the cell writes right next to a plan show its style; a reading against
    that style is never used. Unsafe also: no plan on that side, 2+ plans in a row that
    could share the date, or -- only when the cell shows no style -- the plan on the
    other side left with no date (it could be that plan's)."""
    plans = [it for it in seq if it["kind"] == "P"]
    date_pos = [it["pos"] for it in seq if it["kind"] == "D"]
    plan_first_shown = any(dp > pp for dp, pp in glued.items())
    date_first_shown = any(dp < pp for dp, pp in glued.items())
    if (plan_first and date_first_shown) or (not plan_first and plan_first_shown):
        return None
    styled = plan_first_shown if plan_first else date_first_shown
    assign = dict(glued)
    for d in loose:
        side = [p for p in plans if (p["pos"] < d["pos"] if plan_first else p["pos"] > d["pos"])]
        if not side:
            return None
        assign[d["pos"]] = (side[-1] if plan_first else side[0])["pos"]
    dated = set(assign.values())
    for d in loose:
        if plan_first:
            prev_d = max((x for x in date_pos if x < d["pos"]), default=-1)
            between = [p for p in plans if prev_d < p["pos"] < d["pos"]]
            other = next((p for p in plans if p["pos"] > d["pos"]), None)
        else:
            next_d = min((x for x in date_pos if x > d["pos"]), default=len(seq))
            between = [p for p in plans if d["pos"] < p["pos"] < next_d]
            other = next((p for p in reversed(plans) if p["pos"] < d["pos"]), None)
        if len(between) > 1 or (not styled and other is not None and other["pos"] not in dated):
            return None
    return assign

SPACE_REVIEW_REASON = "Not sure if this is one plan or several (2 or more spaces between words)"

def _soft_plan_groups(seq):
    """Runs of 2+ plan names written with only 2+ spaces between them (no comma, no date)."""
    groups, cur = [], []
    for it in seq:
        if it["kind"] != "P":
            if len(cur) > 1:
                groups.append(cur)
            cur = []
            continue
        if cur and it["piece"] == cur[-1]["piece"] + 1 and it["hard"] == cur[-1]["hard"]:
            cur.append(it)
        else:
            if len(cur) > 1:
                groups.append(cur)
            cur = [it]
    if len(cur) > 1:
        groups.append(cur)
    return groups

def _apply_space_decisions(seq, helper):
    """Re-reads the plan names of one cell with helper's decisions: a run of names written
    with 2+ spaces between them becomes one plan or separate plans (helper.decide), a
    single name saved in Alias_Map as "A | B" becomes separate plans (helper.alias_split).
    Separate plans keep their own place, so a date written right next to one of them is
    that plan's. Returns the new item list, or None when the user chose review."""
    groups = {id(g[0]): g for g in _soft_plan_groups(seq)}
    out, shift, skip = [], 0, set()
    for it in seq:
        if id(it) in skip:
            continue
        if it["kind"] != "P":
            new_it = dict(it)
            new_it["piece"] = it["piece"] + shift
            out.append(new_it)
            continue
        group = groups.get(id(it), [it])
        names = [g["name"] for g in group]
        new_names = helper.decide(names) if len(group) > 1 else helper.alias_split(it["name"])
        if new_names is None:
            return None
        original = "  ".join(names)
        split_from = original if len(new_names) > 1 else None
        for k, nm in enumerate(new_names):
            if len(new_names) == len(group):
                obs = list(group[k].get("obs", []))
            else:
                obs = [o for g in group for o in g.get("obs", [])] if k == 0 else []
            out.append({"kind": "P", "name": nm, "piece": it["piece"] + shift + k, "hard": it["hard"],
                        "obs": obs, "split_from": split_from, "split_n": len(new_names)})
        shift += len(new_names) - len(group)
        skip.update(id(g) for g in group[1:])
    for i, it in enumerate(out):
        it["pos"] = i
    return out

def extract_plan_rows_detailed(text, helper=None):
    """{"plans": [(name, start, end, [observations], open_end)], "review": reason or None,
    "not_plans": [texts saying there is no plan]}.
    Date ownership (agreed with the user 2026-11-19):
      R1 one plan name in the cell -> every date is its;
      R2 a date in the same piece as exactly one plan -> that plan (before or after it);
      R3/R4 any other date -> plan-first and date-first readings; used only when exactly
         one is safe, else the whole cell goes to review (nothing from it is applied).
    Unreadable dates are left blank with an Observation; CURRENT/PRESENT ends are blank.
    helper (the run's plan-name knowledge, see _PlanTextHelper) decides names written with
    2+ spaces between words; without it they stay separate plans."""
    result = {"plans": [], "review": None, "not_plans": []}
    if text is None or (not isinstance(text, str) and pd.isna(text)) or not str(text).strip():
        return result
    seq = _tokenize_plan_text(_normalize_plan_text(text), is_known=helper.is_known if helper else None,
                              not_plans=result["not_plans"])
    if helper is not None:
        seq = _apply_space_decisions(seq, helper)
        if seq is None:
            result["review"] = SPACE_REVIEW_REASON
            return result
    plans = [it for it in seq if it["kind"] == "P"]
    dates = [it for it in seq if it["kind"] == "D"]
    if not plans:
        if dates:
            result["review"] = "Plan text could not be read"
        return result

    owner = {}
    if dates:
        if len({p["name"] for p in plans}) == 1:
            owner = {d["pos"]: plans[0]["pos"] for d in dates}
        else:
            glued, loose = {}, []
            for d in dates:
                same_piece = [p for p in plans if p["piece"] == d["piece"]]
                if len(same_piece) == 1:
                    p = same_piece[0]
                    # "W D1, D2 A D3": D2 sits before a plan that has its own date after it,
                    # with another plan earlier -- it may be that earlier plan's extra
                    # period, so the readings below decide it.
                    two_sided = d["pos"] < p["pos"] and any(
                        o["piece"] == p["piece"] and o["pos"] > p["pos"] for o in dates)
                    if two_sided and any(q["pos"] < d["pos"] for q in plans):
                        loose.append(d)
                    else:
                        glued[d["pos"]] = p["pos"]
                else:
                    loose.append(d)
            owner = glued
            if loose:
                pf = _date_reading(seq, glued, loose, plan_first=True)
                dfr = _date_reading(seq, glued, loose, plan_first=False)
                if (pf is None and dfr is None) or (pf is not None and dfr is not None and pf != dfr):
                    result["review"] = "Dates could belong to more than one plan"
                    return result
                owner = pf if pf is not None else dfr
    # Observation for a plan whose dates all sit apart from it ("AMERIGROUP, 01/01/2024 -
    # 12/31/2024"); a plan's extra period after its own dated entry is normal, no note.
    _plan_pieces = {p["hard"] for p in plans}
    _apart = {d["pos"] for d in dates if d["hard"] not in _plan_pieces}
    _dated_alongside = {seq[owner[d["pos"]]]["name"] for d in dates if d["pos"] not in _apart}
    loose_obs = {pos for pos in _apart if seq[owner[pos]]["name"] not in _dated_alongside}

    by_name = {}
    for p in plans:
        info = by_name.setdefault(p["name"], {"starts": [], "ends": [], "open": False, "obs": [],
                                              "dated": False, "split_from": None, "split_n": 0})
        info["obs"].extend(p.get("obs", []))
        if p.get("split_from"):
            info["split_from"], info["split_n"] = p["split_from"], p.get("split_n", 0)
    for d in dates:
        info = by_name[seq[owner[d["pos"]]]["name"]]
        info["dated"] = True
        start, obs = _read_text_date(d["start"])
        info["obs"].append(obs)
        if start is not None:
            info["starts"].append(start)
        if d["single"] and start is not None:
            info["obs"].append(f"Only one date given ({start:%m/%d/%Y}) - used as start date")
        if d["open"]:
            info["open"] = True
            info["obs"].append(f"End date given as '{d['open']}' - left blank")
        elif d["end"]:
            end, obs = _read_text_date(d["end"], is_end=True)
            info["obs"].append(obs)
            if end is not None:
                info["ends"].append(end)
        if d["pos"] in loose_obs:
            info["obs"].append("Dates were not written next to the plan name - check")

    for name, info in by_name.items():
        sf = info["split_from"]
        if sf:
            _parts = [n for n, i in by_name.items() if i["split_from"] == sf]
            _head = f"Split from '{sf}' into {info['split_n']} plans"
            if not any(by_name[n]["dated"] for n in _parts):
                info["obs"].insert(0, _head)
            elif info["dated"]:
                info["obs"].insert(0, f"{_head} - dates belong to {name} (written right next to it)")
            else:
                info["obs"].insert(0, f"{_head} - no date was written next to {name}")
    for name, info in by_name.items():
        obs = (["Plan name has a '-' - check"] if "-" in name else []) + info["obs"]
        obs = list(dict.fromkeys(o for o in obs if o))
        end = None if info["open"] else (max(info["ends"]) if info["ends"] else None)
        result["plans"].append((name, min(info["starts"]) if info["starts"] else None, end, obs, info["open"]))
    return result

def extract_plan_rows(text):
    """[(plan, start, end)] -- [] when the cell needs review (see extract_plan_rows_detailed)."""
    res = extract_plan_rows_detailed(text)
    if res["review"]:
        return []
    return [(name, start, end) for name, start, end, _obs, _open in res["plans"]]

# =====================================================
# LA FILE PROCESSOR
# =====================================================

def process_la_file(filepath):

    wb  = _retry_on_permission_error(
        lambda: _opxl.load_workbook(filepath, data_only=True, read_only=True), "the import file"
    )
    tabs = wb.sheetnames
    wb.close()

    # A mismatched tab/column/row naming used to silently resolve to None and crash
    # confusingly later (e.g. sheet_name=None means something else to pd.read_excel).
    # Each of the 5 auto-detects below now falls back to _ask_detected_fallback on
    # failure -- a manual pick from the file's actual contents, or 'C' to cancel.
    entitlement_tab = next((s for s in tabs if "ENTITLEMENT" in s.upper()), None)
    # "EV" as its own word only ("REVIEW"/"LEVEL" are not EV tabs), never the Mismatch/Entitlement tab.
    ev_tab          = next((s for s in tabs
                            if "MISMATCH" not in s.upper() and "ENTITLEMENT" not in s.upper()
                            and (re.search(r"(?<![A-Z])EV(?![A-Z])", s.upper())
                                 or "REQUEST" in s.upper() or "RESPONSE" in s.upper())), None)
    mismatch_tab    = next((s for s in tabs if "MISMATCH" in s.upper()), None)

    if entitlement_tab is None:
        _, entitlement_tab = _ask_detected_fallback(
            tabs, 'Could not find a tab containing "ENTITLEMENT" in this file.',
            "Which tab is the real Entitlement tab? Enter the number: "
        )
    if ev_tab is None:
        _, ev_tab = _ask_detected_fallback(
            tabs, 'Could not find a tab containing "EV"/"REQUEST"/"RESPONSE" in this file.',
            "Which tab is the real EV/Request/Response tab? Enter the number: "
        )

    # EV Tab — SSN → Archer ID ("N/A" kept as text, so a warning can show what was there)
    ev_raw = _strip_control_chars_df(_retry_on_permission_error(
        lambda: pd.read_excel(filepath, sheet_name=ev_tab, dtype=str, keep_default_na=False, na_values=[""]),
        "the import file"
    ))
    ev_raw.columns = _clean_headers(ev_raw.columns)

    _phase_banner("column", f"COLUMN IDENTIFICATION (LA File — '{ev_tab}' tab)")
    archer_col = _ask_columns(
        ev_raw,
        "Which column is the LienId (Archer Id)? Enter the column number: ",
        allow_empty=False, multi=False
    )[0]
    ssn_ev_col = next((c for c in ev_raw.columns if "SSN"    in str(c).upper()), None)
    if ssn_ev_col is None:
        _, ssn_ev_col = _ask_detected_fallback(
            list(ev_raw.columns), f"Could not find an SSN column in the {ev_tab!r} tab.",
            "Which column is SSN? Enter the number: "
        )
    # EV tab: SSN -> its LienId(s). SSNs are compared with only dashes/spaces/Excel ".0"
    # removed (no zero padding). One SSN with 2+ LienIds, or one LienId under 2+ SSNs, is
    # a file mistake: those plans are not applied anywhere (-> Pending_Plan_Text_Review).
    # A LienId with a blank SSN can't be tied to any plan and is simply not used.
    ssn_lien_ids, ssn_bad_value, lien_ssns, ssn_shown = {}, {}, {}, {}
    for _ssn_raw, _lid_raw in zip(ev_raw[ssn_ev_col], ev_raw[archer_col]):
        _ssn = _norm_ssn(_ssn_raw)
        if not _ssn:
            continue
        ssn_shown.setdefault(_ssn, str(_ssn_raw).strip())
        _lid = _clean_lien_id(_lid_raw)
        if re.fullmatch(r"\d+", _lid):
            if _lid not in ssn_lien_ids.setdefault(_ssn, []):
                ssn_lien_ids[_ssn].append(_lid)
            lien_ssns.setdefault(_lid, [])
            if _ssn not in lien_ssns[_lid]:
                lien_ssns[_lid].append(_ssn)
        else:
            ssn_bad_value.setdefault(_ssn, "" if pd.isna(_lid_raw) else str(_lid_raw).strip())

    # Mismatch SSNs -- tracked purely as an informational flag now (see
    # mismatch_flagged_lienids below), never to exclude a row from processing.
    mismatch_ssns = set()
    if mismatch_tab:
        mis_df = _retry_on_permission_error(
            lambda: pd.read_excel(filepath, sheet_name=mismatch_tab, dtype=str), "the import file"
        )
        mis_df.columns = _clean_headers(mis_df.columns)
        ssn_mis_col = next((c for c in mis_df.columns if "SSN" in str(c).upper()), None)
        if ssn_mis_col:
            mismatch_ssns = {s for s in (_norm_ssn(v) for v in mis_df[ssn_mis_col]) if s}

    # Entitlement Tab — header row dhundho
    ent_raw = _retry_on_permission_error(
        lambda: pd.read_excel(filepath, sheet_name=entitlement_tab, header=None, dtype=str), "the import file"
    )

    header_row = None
    for i, row in ent_raw.iterrows():
        if any("SSN" in str(v).upper() for v in row if pd.notna(v)):
            header_row = i
            break
    if header_row is None:
        _header_preview = []
        for i in range(min(10, len(ent_raw))):
            _row_vals = [str(v) for v in ent_raw.iloc[i] if pd.notna(v)]
            _header_preview.append(" | ".join(_row_vals) if _row_vals else "(blank row)")
        header_row, _ = _ask_detected_fallback(
            _header_preview,
            f'Could not find a row containing "SSN" in the {entitlement_tab!r} tab to detect the header row.',
            "Which row number is the real header? Enter the number: "
        )

    # Healthy Louisiana col start
    hl_col_start = None
    for i, row in ent_raw.iterrows():
        for j, val in enumerate(row):
            if pd.notna(val) and "HEALTHY" in str(val).upper():
                hl_col_start = j
                break
        if hl_col_start is not None:
            break
    if hl_col_start is None:
        _hl_header_vals = [
            str(v) if pd.notna(v) else f"(blank col {j})"
            for j, v in enumerate(ent_raw.iloc[header_row])
        ]
        hl_col_start, _ = _ask_detected_fallback(
            _hl_header_vals,
            'Could not find a "Healthy Louisiana" marker in the Entitlement tab -- this is normally how '
            "the code knows where the Plan Name/Start Month/End Month columns begin.",
            "Which column does the Healthy Louisiana plan section start at? Enter the number: "
        )

    ent_df = _strip_control_chars_df(_retry_on_permission_error(
        lambda: pd.read_excel(filepath, sheet_name=entitlement_tab, header=header_row, dtype=str), "the import file"
    ))
    ent_df.columns = _clean_headers(ent_df.columns)

    ssn_ent_col = next((c for c in ent_df.columns if "SSN" in str(c).upper()), None)
    if ssn_ent_col is None:
        _, ssn_ent_col = _ask_detected_fallback(
            list(ent_df.columns), f"Could not find an SSN column in the {entitlement_tab!r} tab.",
            "Which column is SSN? Enter the number: "
        )
    # Processing set -- a Mismatch-tab SSN no longer excludes the row from processing
    # (every MCO plan on file should be processed, per policy; the Mismatch flag is
    # tracked separately below purely as an informational note, never as a reason to
    # skip). No Y/N-entitlement filter here either: a row with an actual plan name
    # should still be looked at regardless of that column's value -- same simple rule
    # normal files use (does this column actually have a plan name in it or not).
    ent_df_for_processing = ent_df.copy()

    # Mismatch flag (information only): every LienId of a Mismatch-tab SSN.
    mismatch_flagged_lienids = {lid for ssn in mismatch_ssns for lid in ssn_lien_ids.get(ssn, [])}

    # Sirf HL Plan columns
    all_cols       = list(ent_df_for_processing.columns)
    plan_name_cols = [c for c in all_cols if "PLAN" in str(c).upper()  and "NAME"  in str(c).upper() and all_cols.index(c) >= hl_col_start]
    start_cols     = [c for c in all_cols if "START" in str(c).upper() and "MONTH" in str(c).upper() and all_cols.index(c) >= hl_col_start]
    end_cols       = [c for c in all_cols if "END"   in str(c).upper() and "MONTH" in str(c).upper() and all_cols.index(c) >= hl_col_start]

    def parse_la_date(d, is_end=False):
        """(datetime or None, Observation or None). Start/End Month may be a real date, or a
        bare YYYYMM ("201601", "201601.0" from an Excel number cell). An End Month that is a
        month (YYYYMM, or a date on the 1st) means that whole month -> its last day. An
        unreadable value is left blank with an Observation (never "Bad Date")."""
        if d is None or (not isinstance(d, str) and pd.isna(d)):
            return None, None
        d_str = str(d).strip()
        if not d_str or d_str.upper() in {"NAN", "NAT", "NONE"}:
            return None, None
        parsed, month_only = None, False
        for fmt, text in (("%Y-%m-%d", d_str[:10]), ("%m/%d/%Y", d_str)):
            try:
                parsed = datetime.strptime(text, fmt)
                break
            except ValueError:
                continue
        if parsed is None:
            _digits = re.sub(r"\.0+$", "", d_str)
            if re.fullmatch(r"\d{6}", _digits) and 1 <= int(_digits[4:6]) <= 12:
                parsed, month_only = datetime(int(_digits[:4]), int(_digits[4:6]), 1), True
        if parsed is None:
            return None, f"{'End' if is_end else 'Start'} date not readable: '{d_str}' - not used"
        if is_end and (month_only or parsed.day == 1):
            parsed = parsed.replace(day=calendar.monthrange(parsed.year, parsed.month)[1])
        return parsed, None

    def _ssn_lien_status(ssn):
        """(LienId, None) when the EV tab ties this SSN to exactly one LienId; else
        (None, problem) with problem one of "multi" / "shared" / "bad" / "missing"."""
        ids = ssn_lien_ids.get(ssn, [])
        if len(ids) > 1:
            return None, "multi"
        if len(ids) == 1:
            return (None, "shared") if len(lien_ssns.get(ids[0], [])) > 1 else (ids[0], None)
        return None, ("bad" if ssn in ssn_bad_value else "missing")

    def _la_review(problem, ssn, ssn_raw, row_idx, plan, start, end):
        """One plan of an LA row that can't be tied to one LienId -> Pending_Plan_Text_Review."""
        _raw_text = "" if ssn_raw is None or (not isinstance(ssn_raw, str) and pd.isna(ssn_raw)) else str(ssn_raw).strip()
        shown = ssn_shown.get(ssn) or _raw_text
        if problem == "multi":
            ids = ssn_lien_ids[ssn]
            id_list = ", ".join(ids)
            reason = f"SSN has {len(ids)} LienIds in EV tab ({id_list})"
            warning = (f"SSN {shown} is listed in the EV tab with {len(ids)} different LienIds ({id_list}). "
                       f"Plans were not applied to any of them. Enter the correct LienId in "
                       f"Pending_Plan_Text_Review and re-run with the review file.")
            remark = f"{reason} - plans not applied, see Pending_Plan_Text_Review"
        elif problem == "shared":
            lid = ssn_lien_ids[ssn][0]
            shared = [ssn_shown.get(s, s) for s in lien_ssns[lid]]
            reason = f"LienId {lid} is shared by {len(shared)} SSNs in EV tab"
            warning = (f"LienId {lid} is listed in the EV tab for {len(shared)} different SSNs "
                       f"({', '.join(shared)}). One lien cannot belong to two people, so plans were not "
                       f"applied. In Pending_Plan_Text_Review, enter the correct LienId for each SSN's rows "
                       f"(or NONE).")
            remark = f"{reason} - plans not applied, see Pending_Plan_Text_Review"
        elif problem == "bad":
            reason = "No valid LienId for this SSN in EV tab"
            warning = (f"SSN {shown} has no valid LienId in the EV tab (value: {ssn_bad_value.get(ssn, '')!r}). "
                       f"Its plans were not applied. Enter the LienId in Pending_Plan_Text_Review.")
            remark = f"{reason} - see Pending_Plan_Text_Review"
        elif shown:
            reason = "SSN not found in EV tab"
            warning = (f"SSN {shown} in the Entitlement tab is not in the EV tab - its plans were not "
                       f"applied. Enter the correct LienId in Pending_Plan_Text_Review.")
            remark = f"{reason} - see Pending_Plan_Text_Review"
        else:
            reason = "SSN blank in Entitlement tab"
            warning = (f"An Entitlement row with a blank SSN has plans ({plan}) - its plans were not "
                       f"applied. Enter the correct LienId in Pending_Plan_Text_Review.")
            remark = f"{reason} - see Pending_Plan_Text_Review"
        _queue_plan_text_review(
            None, reason, ssn=shown or None, plan_name=plan, start=format_date(start), end=format_date(end),
            warning=warning, remark=remark, row_key=row_idx,
        )

    # The Entitlement tab may carry its own LienId column. Same as the SSN lookup ->
    # kept as is; different -> the SSN lookup is used (both columns shown); the SSN lookup
    # failing -> the file's own LienId is used; an SSN with 2+ LienIds -> review.
    file_lid_col = next((c for c in ent_df_for_processing.columns
                         if re.sub(r"[\s_]", "", str(c)).lower() == "lienid"), None)
    rows, used_lid, any_diff = [], {}, False
    for row_idx, row in ent_df_for_processing.iterrows():
        ssn_raw = row[ssn_ent_col]
        ssn = _norm_ssn(ssn_raw)
        plans = []
        for plan_col, start_col, end_col in zip(plan_name_cols, start_cols, end_cols):
            plan = row.get(plan_col)
            plan = "" if plan is None or (not isinstance(plan, str) and pd.isna(plan)) else str(plan).strip()
            if not plan or plan.upper() in {"NAN", "N", "NO", "NA", "NONE"}:
                continue
            start, s_obs = parse_la_date(row.get(start_col))
            end, e_obs = parse_la_date(row.get(end_col), is_end=True)
            plans.append((plan.upper(), start, end, [s_obs, e_obs]))
        file_lid = _clean_lien_id(row[file_lid_col]) if file_lid_col else ""
        file_ok = bool(re.fullmatch(r"\d+", file_lid))
        ours, problem = _ssn_lien_status(ssn) if ssn else (None, "missing")
        lien_id, note = ours, None
        if ours is not None and file_ok and file_lid != ours:
            note = f"LienId in file ({file_lid}) differs from the SSN lookup ({ours}) - used {ours}"
        elif ours is None and problem in ("bad", "missing") and file_ok:
            lien_id, problem = file_lid, None
        used_lid[row_idx] = lien_id
        if not plans:
            continue
        if file_lid_col and lien_id is not None and lien_id != file_lid:
            any_diff = True
        if lien_id is None:
            for plan, start, end, _obs in plans:
                _la_review(problem, ssn, ssn_raw, row_idx, plan, start, end)
            continue
        for plan, start, end, obs in plans:
            rows.append({"LienId": lien_id, "MCO_Name": plan, "Start_Date": start, "End_Date": end,
                         "Observation": _join_observations(obs + [note]), "_OpenEnd": False})

    # Display set (Input_File_Original) -- the full Entitlement tab with the LienId each
    # row was processed with, so Remarks attach to the right row. A file LienId column
    # that agrees everywhere is kept as the only one.
    ent_df_display = ent_df.copy()
    lienid_col_out = "LienId"
    if file_lid_col and not any_diff:
        lienid_col_out = file_lid_col
    else:
        _renames, _n = {}, 0
        for _c in ent_df_display.columns:
            if re.sub(r"[\s_]", "", str(_c)).lower() == "lienid":
                _n += 1
                _renames[_c] = "LienId (in file)" if _n == 1 else f"LienId (in file {_n})"
        ent_df_display = ent_df_display.rename(columns=_renames)
        ent_df_display.insert(0, "LienId", pd.Series(used_lid).reindex(ent_df_display.index))

    return pd.DataFrame(rows), ent_df_display, lienid_col_out, mismatch_flagged_lienids, ssn_ent_col

# =====================================================
# BUILD DF — Normal ya LA
# =====================================================

# Populated only on a resume run (see the Pending_Scope_Review handling below) with
# (raw_scope_text, classification) pairs the user filled in directly on that sheet --
# combined with resolve_scope_for_run's own scope_saves later so BOTH sources persist
# to Scope_Map in the same single batched save at the end of the run.
_RESUME_SCOPE_SAVES = []

# Same idea for Pending_TPA_Review: a completed row gets appended to TPA_OVERRIDES so
# THIS run's own lookups can use it, but that alone never told save_config_updates it
# needs to be written to mco_config.xlsx -- combined with resolve_tpa_for_run's own
# tpa_dirty later, same single-batched-save pattern as _RESUME_SCOPE_SAVES above.
_RESUME_TPA_DIRTY = []

# Populated by resolve_tpa_for_run when the user picks a plan from the name_not_found
# list and the raw input text genuinely differs from its real TPA name -- (State,
# From_Name, To_Name) triples for Alias_Map, so this raw text rewrites to the
# canonical name before TPA lookup for any future Tort/Scope, not just this one.
_ALIAS_SAVES = []

# A client's plans can split across runs (some resolve now, some go to a Pending_*
# sheet). Client-MCO Cnt Update must count the client's REAL total, not just this
# run's own df_in, or it undercounts and re-sends a duplicate NewClientNote. Carried
# on each Pending_* row as "_ClientAlreadyCountedThisRun", read back by ClientId and
# combined with this run's own count in build_client_cnt_update. {} on a fresh run.
_ALREADY_COUNTED_CARRY = {}
# The same, as the set of plan/lien keys already counted ("_ClientCountedPlanKeys", JSON)
# -- a resumed client's total is the union, so no plan is ever counted twice. Review
# files from before this column existed fall back to the number above.
_ALREADY_COUNTED_KEYS_CARRY = {}
# Clients whose Client-MCO Count Placeholder row was already given in an earlier round's
# upload files ("_CountPlaceholderGiven") -- never given again on resume.
_PLACEHOLDER_GIVEN_CARRY = set()

# Real bug found: build_step4_outputs' claimed_lien_ids resets every run, so a
# client's OTHER plan that went to Pending_* this run (unresolved) was invisible to
# it -- on resume, the DB still shows that same placeholder lien unclaimed (this tool
# never writes to the DB; the prior round's MCO Update 2 hasn't been applied to Ragic
# yet), so the resumed plan matched it again, pointing two plans at one real lien.
# Carried like _ALREADY_COUNTED_CARRY via "_ClientAlreadyClaimedLienIds" (comma-
# separated) on each Pending_* row, excluded alongside this run's own IDs below.
_ALREADY_CLAIMED_LIEN_IDS_CARRY = {}

# Pending_TPA_Review's Original_MCO_Name (the file's own plan text) across resume rounds:
# (LienId, normalized MCO_Name) -> original text. {} on a fresh run.
_ORIGINAL_MCO_NAME_CARRY = {}

# Tort filled in Pending_TPA_Review for a case whose Tort is blank in LPM (LPM stays
# blank on the re-fetch): CaseId (or ("LIEN", LienId) without a CaseId) -> (TortId, TortName).
_BLANK_TORT_CARRY = {}
BLANK_TORT_REASON = "Tort is blank in LPM"

# Answers about a TPA row marked "DO NOT USE": (State, plan text, TPA name) -> "Y"/"N"/"S".
_DNU_DECISIONS = {}

def _is_do_not_use(name):
    return bool(re.search(r"DO\W*NOT\W*USE", str(name or ""), re.IGNORECASE))

# =====================================================
# PENDING_PLAN_TEXT_REVIEW -- cells the tool must not guess
# =====================================================
# A whole cell goes here (nothing from it applied) when its plan text can't be read or
# its dates could belong to more than one plan; LA rows go here when their SSN can't be
# tied to exactly one LienId. The user fills Plan_Name/Start_Date/End_Date (normal rows)
# or LienId (LA rows) and feeds the review file back in. Nothing is saved to config.
PENDING_PLAN_TEXT_ROWS = []
_PLAN_TEXT_WARNINGS = []
# CaseIds whose scope is blank in LPM this run (listed once at the end of the run).
BLANK_SCOPE_CASES = []
PENDING_PLAN_TEXT_COLS = ["LienId", "SSN", "Source_Column", "Raw_Text", "Reason",
                          "Plan_Name", "Start_Date", "End_Date", "TicketId", "ResponseDate", "ResponseType"]
_PLAN_TEXT_FILL_COLS = {"LienId", "Plan_Name", "Start_Date", "End_Date"}
# Input_File_Original Remarks that aren't a (LienId, plan) outcome: by LienId, and (LA
# rows with no usable LienId) by SSN.
LIEN_REMARKS_EXTRA = {}
SSN_REMARKS_EXTRA = {}
ROW_REMARKS_EXTRA = {}   # LA rows with a blank SSN: by Entitlement row index
_PLAN_TEXT_REMARK = {
    "Dates could belong to more than one plan": "Dates could belong to more than one plan - see Pending_Plan_Text_Review",
    "Plan text could not be read": "Plan text could not be read - see Pending_Plan_Text_Review",
}

def _add_remark(store, key, text):
    if store is LIEN_REMARKS_EXTRA:
        key = _clean_lien_id(key)
    elif store is SSN_REMARKS_EXTRA:
        key = _norm_ssn(key)
    if key not in (None, "") and text and text not in store.setdefault(key, []):
        store[key].append(text)

def _norm_ssn(v):
    """SSN for matching between LA tabs: only dashes, spaces and Excel's ".0" go; the
    digits are compared exactly as read (no zero padding). Text that isn't digits
    ("N/A", "UNKNOWN") is no SSN -> ""."""
    s = _clean_lien_id(v).replace("-", "")
    return s if s.isdigit() else ""

def _short_text(text, limit=120):
    text = " ".join(str(text).split())
    return text if len(text) <= limit else text[:limit - 3] + "..."

def _queue_plan_text_review(lien_id, reason, raw_text=None, source_col=None, ssn=None,
                            plan_name=None, start=None, end=None, warning=None, remark=None, row_key=None):
    """One Pending_Plan_Text_Review row + its terminal warning and Remarks."""
    PENDING_PLAN_TEXT_ROWS.append({
        "LienId": lien_id, "SSN": ssn, "Source_Column": source_col, "Raw_Text": raw_text, "Reason": reason,
        "Plan_Name": plan_name, "Start_Date": start, "End_Date": end,
    })
    if warning is None and reason in _PLAN_TEXT_REMARK:
        if reason == "Dates could belong to more than one plan":
            warning = (f"LienId {lien_id}: the dates in \"{_short_text(raw_text)}\" could belong to more than "
                       f"one plan - nothing from this cell was applied.")
        else:
            warning = (f"LienId {lien_id}: plan text could not be read: \"{_short_text(raw_text)}\" - "
                       f"nothing from this cell was applied.")
    if warning and warning not in _PLAN_TEXT_WARNINGS:
        _PLAN_TEXT_WARNINGS.append(warning)
    remark = remark or _PLAN_TEXT_REMARK.get(reason)
    if _clean_lien_id(lien_id):
        _add_remark(LIEN_REMARKS_EXTRA, lien_id, remark)
    elif _norm_ssn(ssn):
        _add_remark(SSN_REMARKS_EXTRA, ssn, remark)
    elif row_key is not None:
        _add_remark(ROW_REMARKS_EXTRA, row_key, remark)

def _print_plan_text_warnings():
    if not _PLAN_TEXT_WARNINGS:
        return
    _gap()
    print(f"⚠️  {len(PENDING_PLAN_TEXT_ROWS)} item(s) sent to Pending_Plan_Text_Review:")
    for _w in _PLAN_TEXT_WARNINGS:
        print(f"   {_w}")
    print("   Fill them in Pending_Plan_Text_Review and re-run with the review file.")

def _plan_text_review_summary():
    """'3 cells sent to Pending_Plan_Text_Review (2 unclear dates, 1 unreadable text)'."""
    labels = {"Dates could belong to more than one plan": "unclear dates",
              "Plan text could not be read": "unreadable text",
              SPACE_REVIEW_REASON: "one plan or several"}
    counts = {}
    for r in PENDING_PLAN_TEXT_ROWS:
        reason = str(r.get("Reason") or "")
        if reason.startswith("The tool could not read this cell"):
            label = "tool error"
        else:
            label = labels.get(reason, reason.split(" - ")[0] if reason else "other")
        counts[label] = counts.get(label, 0) + 1
    parts = ", ".join(f"{n} {label}" for label, n in counts.items())
    return f"{len(PENDING_PLAN_TEXT_ROWS)} item(s) sent to Pending_Plan_Text_Review ({parts})"

def _plan_text_review_df(ticket_id, response_date, response_type):
    """The Pending_Plan_Text_Review sheet; info columns that are blank on every row are left out."""
    if not PENDING_PLAN_TEXT_ROWS:
        return pd.DataFrame()
    out = pd.DataFrame(PENDING_PLAN_TEXT_ROWS)
    for col, val in (("TicketId", ticket_id), ("ResponseDate", response_date), ("ResponseType", response_type)):
        if col not in out.columns:
            out[col] = val
        else:
            out[col] = out[col].where(out[col].notna(), val)
    out = out.reindex(columns=PENDING_PLAN_TEXT_COLS)
    keep = [c for c in out.columns if c in _PLAN_TEXT_FILL_COLS or out[c].notna().any()]
    return out[keep]

def _parse_review_date(v):
    """A date typed in a review sheet (MM/DD/YYYY, MM-DD-YYYY, or an Excel date cell read
    as '2024-01-31 00:00:00') -> datetime, else None (a 2-digit year is never guessed)."""
    s = str(v).strip().upper()
    s = _RE_ISO_DATE.sub(lambda m: f"{int(m.group(2))}/{int(m.group(3))}/{m.group(1)}", s)
    s = _RE_DASH_DATE.sub(lambda m: f"{m.group(1)}/{m.group(2)}/{m.group(3)}", s)
    s = _RE_DOT_DATE.sub(lambda m: f"{m.group(1)}/{m.group(2)}/{m.group(3)}", s)
    if not re.fullmatch(r'\d{1,2}/\d{1,2}/\d{4}', s):
        return None
    return _read_text_date(s)[0]

def _resolve_plan_text_review(rows_df):
    """Resume of Pending_Plan_Text_Review. Returns plan rows ready to process (dicts with
    LienId/MCO_Name/Start_Date/End_Date/Observation/_OpenEnd). Rows not finished go back
    to PENDING_PLAN_TEXT_ROWS (with a warning when something is wrong); a LienId whose
    rows say NONE is closed with Remarks "Reviewed - no MCO plan"."""
    def _val(row, col):
        v = row.get(col)
        if v is None or (not isinstance(v, str) and pd.isna(v)):
            return ""
        v = str(v).strip()
        return "" if v.lower() in ("nan", "none", "nat") and col != "Plan_Name" else v

    def _keep(row, warning=None, remark=None):
        rec = {c: (row.get(c) if c in row.index and pd.notna(row.get(c)) else None) for c in PENDING_PLAN_TEXT_COLS}
        PENDING_PLAN_TEXT_ROWS.append(rec)
        lien = _clean_lien_id(rec.get("LienId"))
        if warning:
            label = f"LienId {lien}" if lien else (f"SSN {rec.get('SSN')}" if rec.get("SSN") else "A row")
            if f"{label}: {warning}" not in _PLAN_TEXT_WARNINGS:
                _PLAN_TEXT_WARNINGS.append(f"{label}: {warning}")
        text = remark or (f"{warning} - see Pending_Plan_Text_Review" if warning
                          else "Still in Pending_Plan_Text_Review")
        if lien:
            _add_remark(LIEN_REMARKS_EXTRA, lien, text)
        elif rec.get("SSN"):
            _add_remark(SSN_REMARKS_EXTRA, rec.get("SSN"), text)

    ready, none_rows, by_lien = [], {}, {}
    for _, row in rows_df.iterrows():
        lien = _clean_lien_id(row.get("LienId"))
        plan = _val(row, "Plan_Name")
        s_raw, e_raw = _val(row, "Start_Date"), _val(row, "End_Date")
        if not lien:
            # Not done yet (LA rows wait for a LienId); a normal row that lost its LienId says so.
            _keep(row, None if (plan or _val(row, "SSN")) else ("LienId missing" if (s_raw or e_raw) else None))
            continue
        if lien.upper() == "NONE":
            # LA row: NONE as the LienId = this SSN's plans belong to no lien.
            if _val(row, "SSN"):
                _add_remark(SSN_REMARKS_EXTRA, _val(row, "SSN"), "Reviewed - no MCO plan")
            continue
        if not re.fullmatch(r"\d+", lien):
            _keep(row, f"LienId {row.get('LienId')!r} is not a number")
            continue
        if plan.upper() == "NONE":
            none_rows.setdefault(lien, []).append(row)
            continue
        if not plan:
            _keep(row, "Plan_Name missing" if (s_raw or e_raw) else None)
            continue
        start = _parse_review_date(s_raw) if s_raw else None
        end = _parse_review_date(e_raw) if e_raw else None
        bad = [f"{c} not readable: {v!r}" for c, v, d in (("Start_Date", s_raw, start), ("End_Date", e_raw, end)) if v and d is None]
        if bad:
            _keep(row, "; ".join(bad))
            continue
        obs = None
        if start is not None and end is None:
            obs = "Only start date given in review - end left blank"
        elif end is not None and start is None:
            obs = "Only end date given in review - start left blank"
        by_lien.setdefault(lien, []).append((row, {
            "LienId": lien, "MCO_Name": _RE_MULTI_SPACE.sub(" ", plan).strip().upper(),
            "Start_Date": start, "End_Date": end, "Observation": obs, "_OpenEnd": False,
        }))
    for lien, rows in none_rows.items():
        if lien in by_lien:
            msg = "has both NONE and a plan - please keep only one"
            for row in rows + [r for r, _ in by_lien.pop(lien)]:
                _keep(row, msg)
        else:
            _add_remark(LIEN_REMARKS_EXTRA, lien, "Reviewed - no MCO plan")
    for entries in by_lien.values():
        ready.extend(rec for _, rec in entries)
    return ready

def _print_plan_text_action_needed(ptr_df):
    if ptr_df is None or ptr_df.empty:
        return
    print("\n" + "=" * 72)
    print(f" ACTION NEEDED: {_plan_text_review_summary()}")
    print(" 1. Read Raw_Text and Reason on each row.")
    print(" 2. Fill Plan_Name, Start_Date and End_Date (MM/DD/YYYY). For 2+ plans in one")
    print("    cell, copy the row once per plan (same LienId).")
    print(" 3. Write NONE in Plan_Name if that lien has no MCO plan.")
    if "SSN" in ptr_df.columns:
        print(" 4. Rows with an SSN (LA file): the plan is already filled - fill the correct LienId.")
    print("=" * 72)

def _write_plan_text_review_only(review_path, ticket_id, response_date, response_type, ev_frame=None,
                                 remarks_col="Remarks"):
    """Review file holding only Pending_Plan_Text_Review (a run with no plan left to process),
    plus Input_File_Original with its Remarks when given (no WIP file that round)."""
    ptr_df = _plan_text_review_df(ticket_id, response_date, response_type)
    if ptr_df.empty:
        return
    _ensure_writable(review_path)
    with pd.ExcelWriter(review_path, engine="openpyxl") as review_writer:
        ptr_df.to_excel(review_writer, sheet_name="Pending_Plan_Text_Review", index=False)
        _style_plan_text_review_sheet(review_writer.sheets["Pending_Plan_Text_Review"], ptr_df)
        if ev_frame is not None:
            ev_frame.to_excel(review_writer, sheet_name="Input_File_Original", index=False)
            ws = review_writer.sheets["Input_File_Original"]
            if remarks_col in ev_frame.columns:
                ci = list(ev_frame.columns).index(remarks_col) + 1
                for r in range(1, len(ev_frame) + 2):
                    ws.cell(row=r, column=ci).fill = _solid_fill(REMARKS_COL_COLOR)
    print(f"\n📋 Review file (items needing manual attention): {review_path}")
    _print_plan_text_action_needed(ptr_df)

# =====================================================
# PENDING_ERROR_REVIEW -- last guard: one plan hitting an unexpected error never stops
# the whole run. That plan only (nothing of it uploaded) goes to Pending_Error_Review
# with a targeted fix; the rest of the run carries on. Feeding the Review file back in
# processes it again from scratch with fresh LPM data.
# =====================================================
PENDING_ERROR_ROWS = []
PENDING_ERROR_COLS = [
    "ClientId", "LienId", "MCO_Name", "Start_Date", "End_Date", "Project_Name", "Scope_Classification",
    "State", "Step", "Error", "What_To_Do", "Attempt", "Observation", "TicketId", "ResponseDate", "ResponseType",
]
# Cells the user may correct before feeding the file back in (when the input file's own
# plan name or date caused the error).
_ERROR_FILL_COLS = ["MCO_Name", "Start_Date", "End_Date"]
# (LienId, normalized MCO_Name) -> how many rounds this plan already failed (resume).
_ERROR_ATTEMPT_CARRY = {}
_ERROR_TOOL_BUG_TEXT = ("No data problem found in this row - most likely a tool bug. Send this run's log "
                        "file to the tool owner. After the fix, feed this Review file back in.")
_ERROR_RESUME_TEXT = "then feed this Review file back in."

def _error_cell(row, col):
    v = row.get(col) if hasattr(row, "get") else None
    if v is None or (not isinstance(v, str) and pd.isna(v)):
        return None
    s = str(v).strip()
    return s if s and s.lower() not in ("nan", "none", "nat") else None

def _error_what_to_do(row, exc):
    """The fix for THIS error, from the row's own data and the value the error names.
    Nothing found -> the tool-bug text (send the log to the tool owner)."""
    if not _error_cell(row, "ClientId"):
        return f"LienId {_error_cell(row, 'LienId')} has no client in LPM. Check the LienId, {_ERROR_RESUME_TEXT}"
    if not _error_cell(row, "State"):
        return ("State is blank for this lien in LPM (it comes from the start of the LienHolderName). "
                f"Fix it in LPM, {_ERROR_RESUME_TEXT}")
    if not _error_cell(row, "Project_Name"):
        return f"Tort is blank for this client's case in LPM. Fill it in LPM, {_ERROR_RESUME_TEXT}"
    if not _error_cell(row, "MCO_Name"):
        return f"The plan name is blank. Fill MCO_Name in this row, {_ERROR_RESUME_TEXT}"
    msg = str(exc)
    named = re.findall(r"'([^']{1,80})'", msg) + re.findall(r'"([^"]{1,80})"', msg)
    sources = {
        "Start_Date": "Correct Start_Date in this row (MM/DD/YYYY)",
        "End_Date": "Correct End_Date in this row (MM/DD/YYYY)",
        "MCO_Name": "Correct MCO_Name in this row",
        "LienholderId": None, "CollectorId": None, "LienType": None,
        "TortId": "Fix the Tort of this client's case in LPM",
        "PrivateLiensScope": "Fix the case's PrivateLiensScope in LPM",
    }
    for value in named:
        for col, fix in sources.items():
            cell = _error_cell(row, col)
            if cell is None or cell != value.strip():
                continue
            if fix is None:
                where = _error_cell(row, "_TPA_Source") or "the TPA file or TPA_Overrides"
                fix = f"Correct it in {where}"
            return f"{col} '{cell}' could not be used. {fix}, {_ERROR_RESUME_TEXT}"
    return _ERROR_TOOL_BUG_TEXT

def _record_plan_error(rows_df, step, exc):
    """Sends these plan rows to Pending_Error_Review: screen warning (one per plan), the
    full traceback in the log only, and the plan's Remarks."""
    _tb = "".join(traceback.format_exception(type(exc), exc, exc.__traceback__))
    _RUN_LOG_STREAM.write_log_only(f"\n[Pending_Error_Review] {step} - full error:\n{_tb}\n")
    err_text = _short_text(f"{type(exc).__name__}: {exc}", 200)
    fate = globals().get("PLAN_FATE")
    seen = set()
    for _, row in rows_df.iterrows():
        lien, plan = _error_cell(row, "LienId"), _error_cell(row, "MCO_Name")
        attempt = _ERROR_ATTEMPT_CARRY.get((lien, _normalize_key(plan)), 0) + 1
        todo = _error_what_to_do(row, exc)
        if attempt > 1:
            todo = f"Failed again (round {attempt}). " + todo
        rec = {c: (row.get(c) if c in row.index else None) for c in PENDING_ERROR_COLS}
        rec.update({"Step": step, "Error": err_text, "What_To_Do": todo, "Attempt": attempt})
        for carry in ("_ClientAlreadyCountedThisRun", "_ClientAlreadyClaimedLienIds", "_ClientCountedPlanKeys",
                      "_CountPlaceholderGiven", "Original_MCO_Name", "CaseId", "TortId"):
            if carry in row.index:
                rec[carry] = row.get(carry)
        PENDING_ERROR_ROWS.append(rec)
        if fate is not None and lien is not None:
            fate[(row.get("LienId"), row.get("MCO_Name"))] = "Not processed - unexpected error, see Pending_Error_Review"
        key = (lien, plan)
        if key in seen:
            continue
        seen.add(key)
        _gap()
        print(f"⚠️  Unexpected error while processing ClientId {_error_cell(row, 'ClientId')} (LienId {lien}, "
              f"plan {plan!r}, Tort {_error_cell(row, 'Project_Name')!r}, Scope "
              f"{_error_cell(row, 'Scope_Classification')!r}) in {step}: {err_text}")
        print("   Nothing for this plan was added to the upload files. It is listed in the "
              "'Pending_Error_Review' sheet of the Review file.")

def _pending_error_df(ticket_id, response_date, response_type):
    if not PENDING_ERROR_ROWS:
        return pd.DataFrame(columns=PENDING_ERROR_COLS)
    out = pd.DataFrame(PENDING_ERROR_ROWS)
    for col, val in (("TicketId", ticket_id), ("ResponseDate", response_date), ("ResponseType", response_type)):
        out[col] = out[col].where(out[col].notna(), val) if col in out.columns else val
    extra = [c for c in out.columns if c not in PENDING_ERROR_COLS]
    out = out.reindex(columns=PENDING_ERROR_COLS + extra)
    for col in ("Start_Date", "End_Date"):
        out[col] = out[col].map(lambda v: _fmt_review_date(v))
    return out

def _fmt_review_date(v):
    """A date for a review sheet as MM/DD/YYYY text; blank stays blank."""
    if v is None or (not isinstance(v, str) and pd.isna(v)):
        return None
    if isinstance(v, str):
        return v.strip() or None
    try:
        return pd.Timestamp(v).strftime("%m/%d/%Y")
    except (TypeError, ValueError):
        return str(v)

def _write_error_review_sheet(writer, df):
    """Pending_Error_Review: red tab; MCO_Name / Start_Date / End_Date to correct when the
    input file's own value caused the error (dates as real dates, MM/DD/YYYY)."""
    out = df.copy()
    for c in ("Start_Date", "End_Date"):
        if c in out.columns:
            out[c] = out[c].map(lambda v: (_parse_review_date(v) or v)
                                if v is not None and not (isinstance(v, float) and pd.isna(v)) and str(v).strip()
                                else None)
    out.to_excel(writer, sheet_name="Pending_Error_Review", index=False)
    ws = writer.sheets["Pending_Error_Review"]
    ws.sheet_properties.tabColor = ALERT_TAB_COLOR
    _fill_cells(ws, out, _ERROR_FILL_COLS)
    _add_review_validations(ws, out, date_cols=("Start_Date", "End_Date"))
    ws.auto_filter.ref = ws.dimensions
    ws.freeze_panes = "A2"

def _style_plan_text_review_sheet(ws, df):
    """Pending_Plan_Text_Review: the cells to fill in Butter Yellow, LienId a whole number
    or NONE (an LA SSN whose plans belong to no lien), real dates for Start_Date/End_Date."""
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.datavalidation import DataValidation
    _fill_cells(ws, df, ["LienId", "Plan_Name", "Start_Date", "End_Date"])
    _add_review_validations(ws, df, date_cols=("Start_Date", "End_Date"))
    if "LienId" in df.columns:
        letter = get_column_letter(list(df.columns).index("LienId") + 1)
        dv = DataValidation(type="custom", allow_blank=True,
                            formula1=f'OR(AND(ISNUMBER({letter}2),{letter}2>0),UPPER({letter}2)="NONE")')
        dv.error, dv.prompt = "Please enter the LienId as a number (or NONE).", "Type the LienId as a number."
        dv.errorTitle, dv.promptTitle, dv.errorStyle = "Not allowed", "LienId", "stop"
        dv.showErrorMessage = dv.showInputMessage = True
        dv.add(f"{letter}2:{letter}{_REVIEW_LAST_ROW}")
        ws.add_data_validation(dv)

def _print_error_action_needed(err_df):
    if err_df is None or err_df.empty:
        return
    print("\n" + "=" * 72)
    print(f" ACTION NEEDED: Pending_Error_Review -- {len(err_df)} plan row(s) for "
          f"{err_df['ClientId'].nunique()} client(s) hit an unexpected error and were NOT processed.")
    print(' 1. Read "What_To_Do" on each row and make that fix.')
    print(" 2. Feed this Review file back in. These plans are processed again from scratch.")
    print("=" * 72)

def _read_existing_csv(path):
    """An earlier round's CSV as text. One re-saved by Excel (Windows encoding) still reads;
    an emptied file is an empty table."""
    try:
        return pd.read_csv(path, dtype=str)
    except UnicodeDecodeError:
        return pd.read_csv(path, dtype=str, encoding="cp1252")
    except pd.errors.EmptyDataError:
        return pd.DataFrame()

def _tool_remarks_col(ev_frame):
    """Name of the tool's own Remarks column in Input_File_Original: "Tool Remarks" when the
    input file already has its own "Remarks" column (kept as it is), else "Remarks"."""
    return "Tool Remarks" if "Remarks" in list(ev_frame.columns) else "Remarks"

def _input_file_remarks(ev_frame, lienid_col, ssn_col=None, fate_remarks=None, mismatch_ids=()):
    """Input_File_Original's Remarks per row: each plan's outcome plus any other note for
    that LienId, and (LA files) notes kept by SSN for rows with no usable LienId. A
    Mismatch-tab flag is added only next to a real outcome, never on its own."""
    fate_remarks = fate_remarks or {}
    lien_vals = ev_frame[lienid_col] if lienid_col in ev_frame.columns else pd.Series([None] * len(ev_frame))
    ssn_vals = ev_frame[ssn_col] if ssn_col and ssn_col in ev_frame.columns else pd.Series([None] * len(ev_frame))
    out = []
    for row_idx, lien_val, ssn_val in zip(ev_frame.index, lien_vals, ssn_vals):
        key = _clean_lien_id(lien_val)
        parts = list(fate_remarks.get(key, [])) + LIEN_REMARKS_EXTRA.get(key, [])
        if ssn_val is not None:
            parts += SSN_REMARKS_EXTRA.get(_norm_ssn(ssn_val), [])
        if ssn_col:
            parts += ROW_REMARKS_EXTRA.get(row_idx, [])
        parts = list(dict.fromkeys(parts))
        if parts and key and key in mismatch_ids:
            parts.append("SSN listed in Mismatch tab")
        out.append(" | ".join(parts))
    return out

# =====================================================
# PLAN NAMES WITH 2+ SPACES BETWEEN WORDS ("MOLINA  HEALTHCARE" vs "HUMANA  AETNA")
# =====================================================
# Order: a decision saved in Alias_Map -> the plan names this run already knows -> ask
# once per text. Saved in Alias_Map: To_Name with "|" = separate plans, without = one plan.

# A single one of these words is never a plan on its own (it only completes a name).
_GENERIC_NAME_WORDS = {
    "HEALTH", "HEALTHY", "HEALTHCARE", "CARE", "PLAN", "PLANS", "COMMUNITY", "MEDICAL", "MEDICAID",
    "MEDICARE", "SERVICES", "MANAGED", "GROUP", "SYSTEM", "SYSTEMS", "NATIONAL", "INSURANCE", "MCO",
    "CCO", "OF", "THE", "AND", "INC", "LLC", "BLUE", "CROSS", "SHIELD", "FIRST", "TOTAL", "FAMILY",
    "CHOICE", "PARTNERS", "ALLIANCE", "NETWORK", "CENTER", "SELECT", "BETTER", "COUNTY", "STATE",
}

def _name_words(s):
    return [w for w in re.split(r"[^A-Z0-9]+", str(s).upper()) if w]

def _known_plan_names_for_state(state):
    """Plan names the run knows for a state: the TPA state sheet, TPA_Overrides, Alias_Map
    (both sides) and LPM's lienholder names. A source that can't be read is skipped."""
    names = set()
    st = str(state or "").strip().upper()
    if not re.fullmatch(r"[A-Z]{2}", st):
        return names
    try:
        _ensure_file_hydrated(TPA_FILE)
        _wb = _opxl.load_workbook(TPA_FILE, data_only=True, read_only=True)
        try:
            _sheet = next((n for n in _wb.sheetnames if n.upper().startswith(st + " ")), None)
            if _sheet:
                _rows = _wb[_sheet].iter_rows(values_only=True)
                _header = [str(h).strip() if h is not None else "" for h in next(_rows)]
                if "LienholderName" in _header and "TortId" in _header:
                    _ni, _ti = _header.index("LienholderName"), _header.index("TortId")
                    for _row in _rows:
                        if len(_row) <= max(_ni, _ti) or _row[_ti] is None or str(_row[_ti]).strip() == "":
                            break
                        if _row[_ni]:
                            names.add(_strip_lpm_tag(_row[_ni]))
        finally:
            _wb.close()
    except Exception:
        pass
    for _e in TPA_OVERRIDES:
        if _normalize_key(_e.get("State MCO")) == f"{st} MCO" and _e.get("LienholderName"):
            names.add(str(_e["LienholderName"]))
    for _frm, _to in ALIAS_MAP.get(st, {}).items():
        names.add(_frm)
        names.update(p.strip() for p in str(_to).split("|") if p.strip())
    try:
        _lpm = _lpm_state_lienholders(st)
        if _lpm is not None and "Name" in _lpm.columns:
            names.update(_strip_lpm_tag(n) for n in _lpm["Name"].dropna())
    except Exception:
        pass
    return {n for n in names if _name_words(n)}

class _PlanTextHelper:
    """The run's plan-name knowledge while reading plan text (normal files): is a text a
    known plan, and is "A  B" one plan or several (see _apply_space_decisions)."""

    def __init__(self, state):
        self.state = str(state or "").strip().upper()
        self.lien_id = None
        self._known = None
        self._decided = {}
        self._banner = False

    def _known_list(self):
        if self._known is None:
            self._known = []
            for n in _known_plan_names_for_state(self.state):
                words = _name_words(n)
                cuts, total = set(), 0
                for w in words:
                    total += len(w)
                    cuts.add(total)
                self._known.append(("".join(words), cuts))
        return self._known

    def is_known(self, name):
        """The whole text is a known plan name (spaces/punctuation ignored)."""
        compact = "".join(_name_words(name))
        return bool(compact) and any(c == compact for c, _ in self._known_list())

    def known_prefix(self, text):
        """The text is a known plan name or its first word(s); one generic word never counts."""
        words = _name_words(text)
        if not words or (len(words) == 1 and words[0] in _GENERIC_NAME_WORDS):
            return False
        compact = "".join(words)
        return any(c.startswith(compact) and len(compact) in cuts for c, cuts in self._known_list())

    def alias_split(self, name):
        """A name saved in Alias_Map as "A | B" -> its separate plans; else [name]."""
        to = ALIAS_MAP.get(self.state, {}).get(_RE_ANY_SPACE_RUN.sub(" ", str(name).upper()).strip())
        if to and "|" in to:
            parts = [p for p in (_RE_ANY_SPACE_RUN.sub(" ", x).strip() for x in to.split("|")) if p]
            if len(parts) > 1:
                return parts
        return [name]

    def decide(self, names):
        """Names written with 2+ spaces between them -> [one plan] / [separate plans], or
        None (user sent the cell to review). Asked once per text per run."""
        key = " ".join(names)
        if key in self._decided:
            return self._decided[key]
        saved = ALIAS_MAP.get(self.state, {}).get(key)
        if saved:
            result = self.alias_split(key) if "|" in saved else [key]
        elif self.known_prefix(key):
            result = [key]
        elif all(self.known_prefix(n) for n in names):
            result = list(names)
        else:
            result = self._ask(names)
        self._decided[key] = result
        return result

    def _ask(self, names):
        shown, joined = "  ".join(names), " ".join(names)
        if not self._banner:
            _phase_banner("plantext", "PLAN NAMES")
            self._banner = True
        _new_item("plantext")
        where = f" (LienId {self.lien_id})" if self.lien_id else ""
        print(f"\nThe plan text '{shown}'{where} has 2 or more spaces between words.")
        print(f"  [1] One plan: '{joined}'")
        print(f"  [2] {len(names)} separate plans: " + ", ".join(f"'{n}'" for n in names))
        print("  [S] Not sure - send this cell to Pending_Plan_Text_Review")
        while True:
            answer = input("Enter 1, 2 or S: ").strip().upper()
            if answer in ("1", "2", "S"):
                break
            print("Please enter 1, 2 or S.")
        if answer == "S":
            return None
        to_name = joined if answer == "1" else " | ".join(names)
        ALIAS_MAP.setdefault(self.state, {})[joined] = to_name
        label = (f"Plan text '{shown}' (State {self.state}) = one plan '{joined}'" if answer == "1"
                 else f"Plan text '{shown}' (State {self.state}) = {len(names)} separate plans: {to_name}")
        _defer_save(label, _ALIAS_SAVES, (self.state, joined, to_name))
        return [joined] if answer == "1" else list(names)

def _tpa_tort_pairs(state):
    """[(TortId, TortName)] of the state's TPA sheet ([] when it can't be read)."""
    st = str(state or "").strip().upper()
    pairs = []
    if not re.fullmatch(r"[A-Z]{2}", st):
        return pairs
    try:
        _ensure_file_hydrated(TPA_FILE)
        _wb = _opxl.load_workbook(TPA_FILE, data_only=True, read_only=True)
        try:
            _sheet = next((n for n in _wb.sheetnames if n.upper().startswith(st + " ")), None)
            if _sheet:
                _rows = _wb[_sheet].iter_rows(values_only=True)
                _header = [str(h).strip() if h is not None else "" for h in next(_rows)]
                if "TortId" in _header and "TortName" in _header:
                    _ti, _tn = _header.index("TortId"), _header.index("TortName")
                    for _row in _rows:
                        if len(_row) <= max(_ti, _tn) or _row[_ti] is None or str(_row[_ti]).strip() == "":
                            break
                        _pair = (_id_key(_row[_ti]), str(_row[_tn]).strip() if _row[_tn] else None)
                        if _pair[0] and _pair[1] and _pair not in pairs:
                            pairs.append(_pair)
        finally:
            _wb.close()
    except Exception:
        pass
    return pairs

def _resume_blank_tort_rows(rows):
    """Pending_TPA_Review rows whose Tort was blank in LPM: the TortId/TortName filled
    there -> _BLANK_TORT_CARRY (by CaseId; LPM is still blank on the fresh re-fetch).
    Checked against the TPA file: the missing half filled from it, a TortId and name that
    disagree with it -> the user picks; neither in TPA -> used as typed (Tort not in TPA).
    Rows of one case filled with different Torts -> the user picks. Nothing filled -> asked
    again during the run. Returns {row index: (TortId, TortName)}."""
    applied = {}
    groups = {}
    for idx, r in rows.iterrows():
        case = _id_key(r.get("CaseId"))
        groups.setdefault(case or ("LIEN", _clean_lien_id(r.get("LienId"))), []).append((idx, r))
    for key, grows in groups.items():
        label = f"CaseId {key}" if isinstance(key, str) else f"LienId {key[1]}"
        pairs = _tpa_tort_pairs(grows[0][1].get("State"))
        by_id = {i: n for i, n in pairs}
        by_name = {_normalize_key(n): i for i, n in pairs}
        answers = []
        for _, r in grows:
            tid = _id_key(r.get("TortId"))
            tname = r.get("TortName") if "TortName" in r.index else None
            tname = str(tname).strip() if tname is not None and pd.notna(tname) and str(tname).strip() else None
            if not tid and not tname:
                continue
            if tid and not re.fullmatch(r"\d+", tid):
                _new_item("resume")
                print(f"\n⚠️  Pending_TPA_Review ({label}): TortId {r.get('TortId')!r} is not a number. "
                      f"This row was not used.")
                continue
            if tid and tname:
                tpa_name_for_id = by_id.get(tid)
                tpa_id_for_name = by_name.get(_normalize_key(tname))
                if tpa_name_for_id and _normalize_key(tpa_name_for_id) == _normalize_key(tname):
                    pair = (tid, tpa_name_for_id)
                elif tpa_name_for_id is None and tpa_id_for_name is None:
                    pair = (tid, tname)
                else:
                    options = []
                    if tpa_name_for_id:
                        options.append((tid, tpa_name_for_id))
                    if tpa_id_for_name and (tpa_id_for_name, by_id[tpa_id_for_name]) not in options:
                        options.append((tpa_id_for_name, by_id[tpa_id_for_name]))
                    if (tid, tname) not in options:
                        options.append((tid, tname))
                    shown = [f"{n} (TortId {i})" + (" - as typed" if (i, n) == (tid, tname) else " - TPA file")
                             for i, n in options]
                    _new_item("resume")
                    pick = ask_numbered_choice(
                        shown, f"Pending_TPA_Review ({label}): TortId {tid} and Tort {tname!r} do not belong "
                               f"together in the TPA file. Which Tort is right?", "Enter the number: ")
                    pair = options[shown.index(pick)]
            elif tid:
                if by_id.get(tid):
                    pair = (tid, by_id[tid])
                else:
                    _new_item("resume")
                    print(f"\nPending_TPA_Review ({label}): TortId {tid} is not in the TPA file.")
                    while True:
                        _nm = input(f"Tort name for TortId {tid}: ").strip()
                        if _nm:
                            break
                        print("Please enter the Tort name.")
                    pair = (tid, _nm)
            else:
                if by_name.get(_normalize_key(tname)):
                    pair = (by_name[_normalize_key(tname)], by_id[by_name[_normalize_key(tname)]])
                else:
                    _new_item("resume")
                    print(f"\nPending_TPA_Review ({label}): Tort {tname!r} is not in the TPA file.")
                    pair = (_ask_typed_id(f"TortId for {tname!r}: ", allow_skip=False), tname)
            if pair not in answers:
                answers.append(pair)
        if not answers:
            continue
        if len(answers) == 1:
            chosen = answers[0]
        else:
            shown = [f"{n} (TortId {i})" for i, n in answers]
            _new_item("resume")
            chosen = answers[shown.index(ask_numbered_choice(
                shown, f"Pending_TPA_Review rows of {label} have different Torts filled in.",
                "Enter the number of the right one: "))]
        _BLANK_TORT_CARRY[key] = chosen
        for idx, _ in grows:
            applied[idx] = chosen
        _new_item("resume")
        print(f"\nResolved from Pending_TPA_Review: {label} (Tort blank in LPM) -> {chosen[1]!r} (TortId {chosen[0]})")
    return applied

SSN_COL_IN_EVDF = None

if IS_RESUME_RUN:
    _resume_frames = []
    # Columns each review sheet needs to be processed at all.
    _RESUME_REQUIRED_COLS = {
        "Pending_Scope_Review": ["LienId", "MCO_Name", "ClientId"],
        "Pending_TPA_Review": ["LienId", "MCO_Name", "ClientId", "State"],
        "Pending_TPA_Blank_Fields": ["LienId", "MCO_Name", "ClientId"],
        "Pending_Lien_Match_Review": ["LienId", "MCO_Name", "ClientId", "State", "Decision",
                                      "CandidateLienHolderName"],
        "Pending_Plan_Text_Review": ["LienId"],
        "Pending_Error_Review": ["LienId", "MCO_Name"],
    }
    for _sheet in RESUME_SHEETS_PRESENT:
        # Only a truly empty cell is blank: "None"/"N/A" typed (a scope text, a plan
        # answer) is the user's own text.
        _pending_raw = _strip_control_chars_df(_retry_on_permission_error(
            lambda: pd.read_excel(MCO_INPUT_FILE, sheet_name=_sheet, dtype=str,
                                  keep_default_na=False, na_values=[""]), "the import file"
        ))
        _pending_raw.columns = _clean_headers(_pending_raw.columns)
        _missing_cols = [c for c in _RESUME_REQUIRED_COLS.get(_sheet, []) if c not in _pending_raw.columns]
        if _missing_cols:
            print(f"\n❌ The review sheet '{_sheet}' has no column {', '.join(repr(c) for c in _missing_cols)}. "
                  f"Please restore the column header (it may have been renamed or deleted) and run again.")
            _offer_rerun("\nPress 'R' and Enter to start again, or just press Enter to close this window: ")
            sys.exit(0)
        # A fully blank row (e.g. left after deleting cells) is not a review row.
        _pending_raw = _pending_raw.dropna(how="all")
        _pending_raw["_source_sheet"] = _sheet
        _resume_frames.append(_pending_raw)
    _resume_combined = pd.concat(_resume_frames, ignore_index=True)

    if "_ClientAlreadyCountedThisRun" in _resume_combined.columns:
        _carry_raw = pd.to_numeric(_resume_combined["_ClientAlreadyCountedThisRun"], errors="coerce").fillna(0)
        for _cid, _carry_val in zip(_resume_combined["ClientId"], _carry_raw):
            _ALREADY_COUNTED_CARRY[_cid] = max(_ALREADY_COUNTED_CARRY.get(_cid, 0), int(_carry_val))

    if "_ClientCountedPlanKeys" in _resume_combined.columns:
        for _cid, _keys_raw in zip(_resume_combined["ClientId"], _resume_combined["_ClientCountedPlanKeys"]):
            if pd.isna(_keys_raw) or not str(_keys_raw).strip():
                continue
            try:
                _keys = json.loads(str(_keys_raw))
            except ValueError:
                continue
            if isinstance(_keys, list):
                _ALREADY_COUNTED_KEYS_CARRY.setdefault(_cid, set()).update(str(k) for k in _keys)

    if "_CountPlaceholderGiven" in _resume_combined.columns:
        for _cid, _given in zip(_resume_combined["ClientId"], _resume_combined["_CountPlaceholderGiven"]):
            if pd.notna(_given) and str(_given).strip().upper() in ("YES", "TRUE", "1") and _id_key(_cid):
                _PLACEHOLDER_GIVEN_CARRY.add(_id_key(_cid))

    # Pending_Error_Review: how many rounds each plan already failed ("Failed again").
    if "Attempt" in _resume_combined.columns:
        _err_rows = _resume_combined[_resume_combined["_source_sheet"] == "Pending_Error_Review"]
        for _lid, _mname, _att in zip(_err_rows["LienId"], _err_rows["MCO_Name"], _err_rows["Attempt"]):
            _n = pd.to_numeric(_att, errors="coerce")
            if pd.notna(_n):
                _ERROR_ATTEMPT_CARRY[(_clean_lien_id(_lid), _normalize_key(_mname))] = int(_n)

    if "_ClientAlreadyClaimedLienIds" in _resume_combined.columns:
        for _cid, _claimed_raw in zip(_resume_combined["ClientId"], _resume_combined["_ClientAlreadyClaimedLienIds"]):
            if pd.isna(_claimed_raw) or not str(_claimed_raw).strip():
                continue
            _claimed_set = {x.strip() for x in str(_claimed_raw).split(",") if x.strip()}
            _ALREADY_CLAIMED_LIEN_IDS_CARRY.setdefault(_cid, set()).update(_claimed_set)

    def _resume_first_value(col):
        if col in _resume_combined.columns and _resume_combined[col].notna().any():
            return _resume_combined[col].dropna().iloc[0]
        return None

    TicketId = _resume_first_value("TicketId")
    TicketId = str(TicketId).strip() if TicketId is not None and str(TicketId).strip() else None
    if not TicketId:
        _gap()
        print("This review file has no TicketId.")
        TicketId = _ask_ticket_id()
    RESPONSE_DATE = _resume_first_value("ResponseDate")
    RESPONSE_TYPE = _resume_first_value("ResponseType")

    _phase_banner("resume", "RESUME PROCESSING")

    # Pending_Scope_Review: the user fills "Scope_Classification" directly on that sheet.
    # Pre-seed SCOPE_MAP here (before resolve_scope_for_run's classify_scope() pass)
    # so these rows resolve with no prompt. SCOPE_MAP update is unconditional; saving
    # to _RESUME_SCOPE_SAVES is the end-of-run question -- one decision per unique
    # scope text, not per row.
    _scope_review_rows = _resume_combined[_resume_combined["_source_sheet"] == "Pending_Scope_Review"]
    _scope_review_seen = set()
    for _, _srow in _scope_review_rows.iterrows():
        _raw_scope_text = _srow.get("PrivateLiensScope")
        if pd.isna(_raw_scope_text):
            continue
        _scope_key = _normalize_key(_raw_scope_text)
        if _scope_key in _scope_review_seen:
            continue
        _scope_cell = _srow.get("Scope_Classification")
        _user_scope_input = str(_scope_cell).strip().upper() if pd.notna(_scope_cell) else ""
        if _user_scope_input in ("100", "100%", "100% PLRP", "100%PLRP"):
            _resume_classification = "100% PLRP"
        elif _user_scope_input in ("NO", "NO PLRP", "NOPLRP"):
            _resume_classification = "NO PLRP"
        else:
            continue
        SCOPE_MAP[_scope_key] = _resume_classification
        _scope_review_seen.add(_scope_key)
        _new_item("resume")
        print(f"\nResolved from Pending_Scope_Review: {_raw_scope_text!r} -> {_resume_classification}")
        _defer_save(f"Scope {_raw_scope_text!r} = {_resume_classification}",
                    _RESUME_SCOPE_SAVES, (_raw_scope_text, _resume_classification))

    # Blank-scope rows: one answer on any row of a case applies to the whole case (and can
    # be saved per CaseId). Rows of one case filled with different answers -> the user picks.
    if "CaseId" in _scope_review_rows.columns:
        _case_answers = {}
        for _, _srow in _scope_review_rows.iterrows():
            if _normalize_key(_srow.get("PrivateLiensScope")) is not None:
                continue
            _case = _id_key(_srow.get("CaseId"))
            _cell = _srow.get("Scope_Classification")
            _ans = str(_cell).strip().upper() if pd.notna(_cell) else ""
            _cls = ("100% PLRP" if _ans in ("100", "100%", "100% PLRP", "100%PLRP")
                    else "NO PLRP" if _ans in ("NO", "NO PLRP", "NOPLRP") else None)
            if _case and _cls and _cls not in _case_answers.setdefault(_case, []):
                _case_answers[_case].append(_cls)
        for _case, _answers in _case_answers.items():
            if not _answers:
                continue
            _cls = _answers[0] if len(_answers) == 1 else ask_numbered_choice(
                _answers, f"Pending_Scope_Review rows of CaseId {_case} (scope blank in LPM) have different answers.",
                "Enter the number of the right one: ")
            SCOPE_MAP[_case_scope_key(_case)] = _cls
            _new_item("resume")
            print(f"\nResolved from Pending_Scope_Review: CaseId {_case} (scope blank in LPM) -> {_cls}")
            _defer_save(f"Scope for CaseId {_case} (blank in LPM) = {_cls}", _RESUME_SCOPE_SAVES, (None, _cls, _case))

    # Pending_TPA_Review: same shape -- TPA_OVERRIDES (in-memory) update is
    # unconditional so THIS run's resolve_tpa_for_run finds it via _tpa_overrides_find;
    # the end-of-run save question only controls _RESUME_TPA_DIRTY (the persist-to-
    # mco_config.xlsx queue). Filled rows are grouped per unique (State/Tort/Name/Scope)
    # -> one entry per plan however many of its rows were filled; every row is checked
    # first, and rows of one plan filled with different values make the user pick.
    def _clean_filled_id(v):
        # " 300 " / "300.0" typed by hand -> "300"
        return re.sub(r"\.0+$", "", str(_stringify_id(v)).strip())

    # Every filled LienholderId is looked up in LPM's LienHolders table once (None = LPM
    # couldn't be read -> these checks are skipped, never blocking the resume).
    _tpa_mask = _resume_combined["_source_sheet"] == "Pending_TPA_Review"
    _tpa_lh_names = _lienholder_names_by_id(
        [_clean_filled_id(v) for v in _resume_combined.loc[_tpa_mask, "LienholderId"].dropna()]
    ) if "LienholderId" in _resume_combined.columns else {}
    _tpa_skip_rows = set()
    _tpa_id_answers, _tpa_name_answers = {}, {}

    # State/Tort-not-in-TPA plans show the LPM name of their LienholderId as MCO_Name. If
    # the user changed that ID, every row with the same file text takes the new ID's LPM
    # name (so its CollectorId/LienType group follows the right lienholder) and the
    # alias is offered at the end-of-run save question.
    if _tpa_lh_names and "PendingReason" in _resume_combined.columns:
        _realign = _resume_combined[_tpa_mask & _resume_combined["PendingReason"].isin(_LPM_REVIEW_REASONS)].copy()
        _realign["_orig"] = (_realign["Original_MCO_Name"].where(_realign["Original_MCO_Name"].notna(), _realign["MCO_Name"])
                             if "Original_MCO_Name" in _realign.columns else _realign["MCO_Name"])
        _realign["_st"] = _realign["State"].astype(str).str.strip().str.upper()
        _realign["_okey"] = _realign["_orig"].map(_normalize_key)
        for (_st, _), _g in _realign.groupby(["_st", "_okey"], sort=False, dropna=False):
            _ids = list(dict.fromkeys(
                _clean_filled_id(v) for v in _g["LienholderId"].dropna() if re.fullmatch(r"\d+", _clean_filled_id(v))
            ))
            if not _ids:
                continue
            _orig_shown = _g["_orig"].iloc[0]
            _chosen_lh = _ids[0] if len(_ids) == 1 else _pick_lienholder_id(_orig_shown, _ids, _tpa_lh_names)
            if _chosen_lh is None:
                _tpa_skip_rows.update(_g.index)
                continue
            _resume_combined.loc[_g.index, "LienholderId"] = _chosen_lh
            _lh_lpm_name = _tpa_lh_names.get(_chosen_lh)
            if _lh_lpm_name and any(_normalize_key(m) != _normalize_key(_lh_lpm_name) for m in _g["MCO_Name"]):
                _resume_combined.loc[_g.index, "MCO_Name"] = _lh_lpm_name
                if _normalize_key(_orig_shown) != _normalize_key(_lh_lpm_name):
                    _defer_save(f"Plan name {_orig_shown!r} -> LPM lienholder {_lh_lpm_name!r} (State {_st})",
                                _ALIAS_SAVES, (_st, _orig_shown, _lh_lpm_name))
    if "Original_MCO_Name" in _resume_combined.columns:
        for _lid, _mname, _oname in _resume_combined.loc[_tpa_mask, ["LienId", "MCO_Name", "Original_MCO_Name"]].itertuples(index=False):
            if pd.notna(_oname) and str(_oname).strip():
                _ORIGINAL_MCO_NAME_CARRY[(str(_lid).strip(), _normalize_key(_mname))] = _oname

    # Tort blank in LPM: the Tort filled here is carried by case (see _resume_blank_tort_rows),
    # and set on the row itself so a full CollectorId/LienholderId/LienType fill below
    # becomes a TPA_Overrides entry for that Tort.
    if "PendingReason" in _resume_combined.columns:
        _bt_mask = _tpa_mask & (_resume_combined["PendingReason"].astype(str).str.strip() == BLANK_TORT_REASON)
        if _bt_mask.any():
            for _bt_idx, (_bt_id, _bt_name) in _resume_blank_tort_rows(_resume_combined[_bt_mask]).items():
                _resume_combined.at[_bt_idx, "TortId"] = _bt_id
                _resume_combined.at[_bt_idx, "Project_Name"] = _bt_name

    _tpa_review_rows = _resume_combined[_tpa_mask]
    _tpa_fill_groups = {}
    _alias_review_seen = set()
    for _row_idx, _rrow in _tpa_review_rows.iterrows():
        if _row_idx in _tpa_skip_rows:
            continue
        _picked_name = _rrow.get("Picked LienholderName")
        _has_picked_name = pd.notna(_picked_name) and str(_picked_name).strip() != ""
        _has_full_ids = (pd.notna(_rrow.get("TortId")) and pd.notna(_rrow.get("CollectorId")) and
                         pd.notna(_rrow.get("LienholderId")) and pd.notna(_rrow.get("LienType")))
        if not (_has_full_ids or _has_picked_name):
            continue
        if _has_full_ids:
            _state_mco = f"{str(_rrow.get('State')).strip().upper()} MCO"
            _tpa_review_key = (
                _normalize_key(_rrow.get("TortId")), _normalize_key(_rrow.get("MCO_Name")),
                _normalize_key(_state_mco), _normalize_key(_rrow.get("Scope_Classification")),
            )
            _tpa_fill_groups.setdefault(_tpa_review_key, []).append(_rrow)
        else:
            # User resolved this row by writing the real TPA plan name directly
            # (Picked LienholderName) instead of CollectorId/LienholderId/LienType --
            # the picked name is already a real TPA row, so a TPA_Overrides entry would
            # be redundant and would freeze today's CollectorId/LienType even if TPA's
            # data changes later. Alias_Map alone is enough: rewritten IN-MEMORY (not
            # just queued for save) so THIS run's normalize_mco_name pass benefits too,
            # and the row re-enters the normal Tort->State->Name->Scope cascade like
            # any first-time occurrence. If the picked name isn't actually in TPA for
            # this Tort/Scope, that same cascade's own Pending_* routing handles it.
            # pd.notna guard: a blank/NaN MCO_Name is truthy in Python ("nan or ''"
            # evaluates to NaN, not ""), so "x or ''" used to turn a blank cell into
            # the literal text "nan".
            _raw_text = str(_rrow.get("MCO_Name")).strip() if pd.notna(_rrow.get("MCO_Name")) else ""
            _picked_clean = str(_picked_name).strip()
            _resume_state_key = str(_rrow.get("State") or "").strip().upper()
            _alias_key = (_resume_state_key, _normalize_key(_raw_text))
            if _alias_key in _alias_review_seen:
                continue
            _alias_review_seen.add(_alias_key)
            # In-memory rewrite always applies this run; the Alias_Map save is the
            # end-of-run question, like every other decision.
            if _normalize_key(_raw_text) != _normalize_key(_picked_clean):
                ALIAS_MAP.setdefault(_resume_state_key, {})[_raw_text.upper()] = _picked_clean.upper()
                _defer_save(f"Plan name {_raw_text!r} -> TPA plan {_picked_clean!r} (State {_resume_state_key})",
                            _ALIAS_SAVES, (_resume_state_key, _raw_text, _picked_clean))
            _new_item("resume")
            print(
                f"\nAliased from Pending_TPA_Review: {_raw_text!r} -> {_picked_clean!r} "
                f"(State {_resume_state_key!r}) -- will resolve normally via TPA."
            )

    # A plan whose values fail a check (or are skipped) gets no entry, so its rows simply
    # come back in the next review file.
    _tpa_fill_entries = []
    _unknown_cids = _unknown_collector_ids(
        _clean_filled_id(_r.get("CollectorId")) for _rows in _tpa_fill_groups.values() for _r in _rows
    )
    for _group_rows in _tpa_fill_groups.values():
        _first = _group_rows[0]
        _plan_label = (f"{_first.get('MCO_Name')!r} (Tort {_first.get('Project_Name')!r} / "
                       f"Scope {_first.get('Scope_Classification')!r})")
        # A row whose State was cleared would save a "NAN MCO" override: not used.
        if not re.fullmatch(r"[A-Z]{2}", str(_first.get("State") or "").strip().upper()):
            _new_item("resume")
            print(f"\n⚠️  Pending_TPA_Review row for {_plan_label}: State is blank. This row was not used.")
            continue
        _fill_options = []
        for _grow in _group_rows:
            _cid, _lid = _clean_filled_id(_grow.get("CollectorId")), _clean_filled_id(_grow.get("LienholderId"))
            _ltype = _canonical_lien_type(_grow.get("LienType"))
            _problems = []
            if _ltype is None:
                _problems.append(f"LienType {_grow.get('LienType')!r} must be "
                                 f"{' or '.join(repr(o) for o in LIEN_TYPE_OPTIONS)}")
            if not re.fullmatch(r"\d+", str(_cid)):
                _problems.append(f"CollectorId {_grow.get('CollectorId')!r} must be a number")
            if not re.fullmatch(r"\d+", str(_lid)):
                _problems.append(f"LienholderId {_grow.get('LienholderId')!r} must be a number")
            if _problems:
                _new_item("resume")
                print(f"\n⚠️  Pending_TPA_Review row for {_plan_label}, ClientId {_grow.get('ClientId')!r}: "
                      f"{'; '.join(_problems)}. This row was not used.")
                continue
            if (_cid, _lid, _ltype) not in _fill_options:
                _fill_options.append((_cid, _lid, _ltype))
        if not _fill_options:
            continue
        _chosen = _fill_options[0] if len(_fill_options) == 1 else _pick_conflicting_fill(_plan_label, _fill_options)
        if _chosen is None:
            continue
        # An ID LPM doesn't know (typo, or brand new) is the user's call, never a dead end.
        _not_in_lpm = []
        if _chosen[0] in _unknown_cids:
            if not _confirm_unknown_id("CollectorId", _chosen[0], _plan_label, _tpa_id_answers,
                                       "If used, AssignedUserId will be blank on new-lien rows. "):
                continue
            _not_in_lpm.append("CollectorId")
        _lh_lpm_name = _tpa_lh_names.get(_chosen[1]) if _tpa_lh_names is not None else None
        if _tpa_lh_names is not None and _lh_lpm_name is None:
            if not _confirm_unknown_id("LienholderId", _chosen[1], _plan_label, _tpa_id_answers):
                continue
            _not_in_lpm.append("LienholderId")
        elif _lh_lpm_name:
            # A typed ID whose LPM name shares nothing with the plan text may be a typo that
            # landed on another real lienholder. The tool's own pre-fill is never asked about.
            _prefilled = _first.get("LPM_LienholderName")
            _orig_text = _first.get("Original_MCO_Name") if pd.notna(_first.get("Original_MCO_Name")) else _first.get("MCO_Name")
            if (not (pd.notna(_prefilled) and _normalize_key(_prefilled) == _normalize_key(_lh_lpm_name))
                    and _normalize_key(_first.get("MCO_Name")) != _normalize_key(_lh_lpm_name)
                    and not _names_related(_orig_text, _lh_lpm_name)
                    and not _confirm_name_mismatch(_chosen[1], _lh_lpm_name, _orig_text, _tpa_name_answers)):
                continue
        _resume_tpa_entry = {
            "TortId": _first.get("TortId"), "TortName": _first.get("Project_Name"),
            "Private Lien Case Scope": _first.get("Scope_Classification"),
            "CollectorId": _chosen[0], "LienholderId": _chosen[1],
            "LienholderName": _first.get("MCO_Name"),
            "LienType": _chosen[2],
            "State MCO": f"{str(_first.get('State')).strip().upper()} MCO",
            # The row's own originally-stored ticket is more precise than the
            # module-level TicketId (which just holds the FIRST ticket found across
            # this resume file's combined rows -- only correct by coincidence if
            # rows came from different tickets). Falls back to it only for an older
            # review file with no stored TicketId of its own.
            "Ticket #": (str(_first.get("TicketId")) if pd.notna(_first.get("TicketId")) else TicketId),
            "Reason": _tpa_override_reason(_first.get("PendingReason")),
        }
        TPA_OVERRIDES.extend([_resume_tpa_entry] * len(_group_rows))
        _tpa_fill_entries.append((_resume_tpa_entry, _not_in_lpm))
        _new_item("resume")
        print(
            f"\nResolved from Pending_TPA_Review: {_first.get('MCO_Name')!r} (Tort {_first.get('Project_Name')!r}) "
            f"-> LienholderId={_chosen[1]}{f' ({_lh_lpm_name})' if _lh_lpm_name else ''}, "
            f"CollectorId={_chosen[0]}, LienType={_chosen[2]}"
        )
    # TPA_OVERRIDES (above) only covers THIS run's own lookups -- _RESUME_TPA_DIRTY is
    # what gets persisted to mco_config.xlsx, if chosen at the end-of-run question.
    for _e, _missing in _tpa_fill_entries:
        _defer_save(_tpa_entry_label(_e) + (f" ({' and '.join(_missing)} not found in LPM)" if _missing else ""),
                    _RESUME_TPA_DIRTY, _e)

    _blank_field_rows = _resume_combined[_resume_combined["_source_sheet"] == "Pending_TPA_Blank_Fields"]
    for _, _brow in _blank_field_rows.iterrows():
        # A row whose TPA LienholderId was blank is found by name + scope, and the
        # LienholderId the user filled in is written too.
        _lh_was_blank = "(LienholderId)" in str(_brow.get("PendingReason") or "")
        if _lh_was_blank and not re.fullmatch(r"\d+", _clean_filled_id(_brow.get("LienholderId"))):
            continue
        if not _lh_was_blank and pd.isna(_brow.get("CollectorId")) and pd.isna(_brow.get("LienType")):
            continue
        # Checked before anything is written into the TPA file: LienType one of the 2
        # options, CollectorId a number LPM knows. Anything wrong -> warning, the row
        # stays in review (the TPA file is not touched).
        _bf_problems = []
        _bf_lt = _brow.get("LienType")
        _bf_lt_clean = _canonical_lien_type(_bf_lt) if pd.notna(_bf_lt) else None
        if pd.notna(_bf_lt) and _bf_lt_clean is None:
            _bf_problems.append(f"LienType {_bf_lt!r} must be {' or '.join(repr(o) for o in LIEN_TYPE_OPTIONS)}")
        _bf_co = _clean_filled_id(_brow.get("CollectorId")) if pd.notna(_brow.get("CollectorId")) else None
        if _bf_co is not None and not re.fullmatch(r"\d+", _bf_co):
            _bf_problems.append(f"CollectorId {_brow.get('CollectorId')!r} must be a number")
        elif _bf_co is not None and _bf_co in _unknown_collector_ids([_bf_co]):
            _bf_problems.append(f"CollectorId {_bf_co} is not in LPM")
        if _bf_problems:
            _new_item("resume")
            print(f"\n⚠️  Pending_TPA_Blank_Fields row for {_brow.get('MCO_Name')!r} (LienholderId "
                  f"{_brow.get('LienholderId')!r}): {'; '.join(_bf_problems)}. The TPA file was not changed; "
                  f"this plan goes back to review.")
            continue
        if _patch_tpa_file_blank_fields(
            _brow.get("TPA_Match_StateMCO"), _brow.get("TPA_Match_TortId"),
            _brow.get("LienholderId"), _bf_co, _bf_lt_clean,
            lienholder_name=_brow.get("TPA_Match_LienholderName") if _lh_was_blank else None,
            scope=_brow.get("TPA_Match_Scope") if _lh_was_blank else None,
            new_lienholder_id=_clean_filled_id(_brow.get("LienholderId")) if _lh_was_blank else None,
        ):
            print(f"\n✅ Patched the TPA file in place for LienholderId {_brow.get('LienholderId')!r}.")
        else:
            print(f"\n⚠️  Could not patch the TPA file for LienholderId {_brow.get('LienholderId')!r} -- skipped, try resuming again later.")

    # Pending_Lien_Match_Review already carries a resolved LienholderId/CollectorId/
    # LienType (the tie happened inside Step 4); the row re-enters the pipeline
    # normally so Step 3 re-derives the same LienholderId deterministically. Each
    # row's DECISION applies to THAT client's plan only in this run (YES = use this
    # lien, NO = not this lien), via _LIEN_MATCH_PICKS / _LIEN_MATCH_CLIENT_EXCLUDED;
    # Step 4's normal order still runs on fresh data (e.g. a lien that became MCO
    # meanwhile still wins). The name-level decision is offered at the end-of-run
    # save, so every client reuses it in future runs (Lien_Name_Equivalence_Map).
    # Still undecided -> a fresh Pending_Lien_Match_Review, same loop as the others.
    _lien_match_review_rows = _resume_combined[_resume_combined["_source_sheet"] == "Pending_Lien_Match_Review"]

    def _lm_plan_key(row):
        return (_id_key(row.get("ClientId")), _id_key(row.get("CollectorId")), _id_key(row.get("LienholderId")))

    def _lm_mark(row):
        """The candidate lien a review row points at: its LienId, else its name."""
        _cand_id = _id_key(row.get("CandidateLienId"))
        if _cand_id:
            return ("ID", _cand_id)
        _cand_name = _normalize_key(row.get("CandidateLienHolderName"))
        return ("NAME", _cand_name) if _cand_name else None

    def _lm_savable(row):
        _reason = row.get("ReviewReason")
        return pd.isna(_reason) or str(_reason).strip() not in _LIEN_REVIEW_NO_SAVE_REASONS

    # Real gap found: nothing caught a user mistakenly marking 2+ DIFFERENT candidates
    # "YES" for the same tie (ClientId+LienId+MCO_Name) -- exactly one should ever be
    # the real match. Pre-scan every tie-group; one with multiple DISTINCT "YES"
    # candidates is resolved live here: the user picks the one to use (the others
    # are simply not used -- never turned into a "NO"), or 'S' skips the group
    # (a fresh Pending_Lien_Match_Review gets written). Counts distinct candidate
    # liens -- the same candidate duplicated across 2 "YES" rows isn't a conflict.
    _lm_invalid_groups = set()
    _lm_live_resolved_groups = set()
    if not _lien_match_review_rows.empty:
        for _lm_group_key, _lm_group_df in _lien_match_review_rows.groupby(
            ["ClientId", "LienId", "MCO_Name"], dropna=False
        ):
            _lm_yes_rows = _lm_group_df[_lm_group_df["Decision"].astype(str).str.strip().str.upper() == "YES"]
            _lm_yes_rows = _lm_yes_rows.assign(_lm_mark=[_lm_mark(r) for _, r in _lm_yes_rows.iterrows()])
            _lm_yes_rows = _lm_yes_rows[_lm_yes_rows["_lm_mark"].notna()].drop_duplicates(subset=["_lm_mark"])
            if len(_lm_yes_rows) <= 1:
                continue
            _new_item("resume")
            print(
                f"\n⚠️  Multiple candidates marked 'YES' for ClientId {_lm_group_key[0]!r}, "
                f"LienId {_lm_group_key[1]!r}, Plan {_lm_group_key[2]!r} -- exactly ONE "
                f"candidate should be 'YES'."
            )
            _lm_yes_rows = _lm_yes_rows.reset_index(drop=True)
            for _i, _yes_row in _lm_yes_rows.iterrows():
                print(
                    f"  [{_i + 1}] {_yes_row['CandidateLienHolderName']} "
                    f"(LienId: {_yes_row.get('CandidateLienId')}, "
                    f"CollectorId: {_yes_row.get('CandidateCollectorId')}, "
                    f"LienholderId: {_yes_row.get('CandidateLienholderId')})"
                )
            while True:
                _lm_pick = input(
                    "Enter a number to pick the real match, or 'S' to skip "
                    "(fix in the review file and resume again): "
                ).strip().upper()
                if _lm_pick == "S":
                    break
                if _lm_pick.isdigit() and 1 <= int(_lm_pick) <= len(_lm_yes_rows):
                    break
                print(f"Please enter a number between 1 and {len(_lm_yes_rows)}, or 'S'.")
            if _lm_pick == "S":
                _lm_invalid_groups.add(_lm_group_key)
                continue

            _lm_picked_row = _lm_yes_rows.iloc[int(_lm_pick) - 1]
            _lm_plan_key_val = _lm_plan_key(_lm_picked_row)
            if None not in _lm_plan_key_val:
                _LIEN_MATCH_PICKS.setdefault(_lm_plan_key_val, set()).add(_lm_picked_row["_lm_mark"])
            _lm_state = _lm_picked_row.get("State")
            _lm_picked_name = _lm_picked_row.get("CandidateLienHolderName")
            if _lm_savable(_lm_picked_row) and pd.notna(_lm_state) and pd.notna(_lm_picked_name):
                _lm_picked_name = str(_lm_picked_name).strip()
                _defer_save(
                    f"Lien match: plan {_lm_group_key[2]!r} = existing lien {_lm_picked_name!r} "
                    f"(State {_lm_state}, LienholderId {_lm_picked_row.get('LienholderId')})",
                    _RESUME_LIEN_EQUIV_SAVES,
                    (_lm_state, _lm_picked_row.get("CollectorId"), _lm_picked_row.get("LienholderId"),
                     _lm_group_key[2], _lm_picked_name, "YES", TicketId),
                )
            _lm_live_resolved_groups.add(_lm_group_key)

    _lien_match_applied = set()
    _lien_match_seen = set()
    for _, _lm_row in _lien_match_review_rows.iterrows():
        _lm_row_group = (_lm_row.get("ClientId"), _lm_row.get("LienId"), _lm_row.get("MCO_Name"))
        if _lm_row_group in _lm_invalid_groups or _lm_row_group in _lm_live_resolved_groups:
            continue
        _lm_decision = str(_lm_row.get("Decision") or "").strip().upper()
        if _lm_decision not in ("YES", "NO"):
            continue
        _lm_mark_val = _lm_mark(_lm_row)
        _lm_plan_key_val = _lm_plan_key(_lm_row)
        if _lm_mark_val is None or None in _lm_plan_key_val:
            continue
        if (_lm_plan_key_val, _lm_mark_val, _lm_decision) in _lien_match_applied:
            continue
        _lien_match_applied.add((_lm_plan_key_val, _lm_mark_val, _lm_decision))
        _lm_client_map = _LIEN_MATCH_PICKS if _lm_decision == "YES" else _LIEN_MATCH_CLIENT_EXCLUDED
        _lm_client_map.setdefault(_lm_plan_key_val, set()).add(_lm_mark_val)

        _lm_state = _lm_row.get("State")
        _lm_collector_id = _lm_row.get("CollectorId")
        _lm_lienholder_id = _lm_row.get("LienholderId")
        _lm_plan_name = _lm_row.get("MCO_Name")
        _lm_equiv_name = _lm_row.get("CandidateLienHolderName")
        _lm_verb = "is the lien to use" if _lm_decision == "YES" else "is NOT this plan's lien"
        _new_item("resume")
        print(
            f"\nResolved from Pending_Lien_Match_Review (ClientId {_lm_row.get('ClientId')}): {_lm_plan_name!r} -> "
            f"{_lm_equiv_name!r} (lien {_lm_row.get('CandidateLienId')}) {_lm_verb}"
        )
        if not _lm_savable(_lm_row) or pd.isna(_lm_state) or pd.isna(_lm_equiv_name):
            continue
        _lm_dedup_key = (str(_lm_state), str(_lm_collector_id), str(_lm_lienholder_id),
                         str(_lm_equiv_name).strip(), _lm_decision)
        if _lm_dedup_key in _lien_match_seen:
            continue
        _lien_match_seen.add(_lm_dedup_key)
        _defer_save(
            f"Lien match: plan {_lm_plan_name!r} {'=' if _lm_decision == 'YES' else 'is NOT'} existing lien "
            f"{_lm_equiv_name!r} (State {_lm_state}, LienholderId {_lm_lienholder_id})",
            _RESUME_LIEN_EQUIV_SAVES,
            (_lm_state, _lm_collector_id, _lm_lienholder_id, _lm_plan_name, _lm_equiv_name, _lm_decision, TicketId),
        )

    # _resume_combined is read via dtype=str (needed so every other column round-
    # trips as plain text), which also turns a genuine Excel date cell into a raw
    # string -- and format_date returns the literal "Bad Date" for any string input.
    # Re-parsed into real datetimes here (coercing anything unparseable to NaT, same
    # as "no date found") at the source instead of downstream.
    _ptr_mask = _resume_combined["_source_sheet"] == "Pending_Plan_Text_Review"
    _ptr_ready = _resolve_plan_text_review(_resume_combined[_ptr_mask]) if _ptr_mask.any() else []
    _plan_rows_src = _resume_combined[~_ptr_mask]
    df = _plan_rows_src.reindex(columns=["LienId", "MCO_Name"]).copy()

    def _reparse_resume_date(series):
        """Real bug found: plain pd.to_datetime(series, errors="coerce") with no
        format= relied on pandas' per-string inference, which can silently mis-parse
        an ambiguous date (e.g. "03/04/2026") in the wrong day/month order. Every
        date THIS script itself ever writes is always MM/DD/YYYY text
        (format_date/is_valid_date_mmddyyyy), so that fixed format is tried FIRST,
        removing the ambiguity for the normal case. Falls back to pandas' own
        inference only for whatever doesn't match it (e.g. a cell a reviewer
        hand-edited in Excel, which can round-trip as an unambiguous ISO-style
        string instead) -- never less permissive than the original code, just no
        longer ambiguity-prone for its own known format."""
        strict = pd.to_datetime(series, format="%m/%d/%Y", errors="coerce")
        _fallback_mask = strict.isna() & series.notna()
        if _fallback_mask.any():
            _loose = pd.to_datetime(series[_fallback_mask], errors="coerce", utc=True)
            strict.loc[_fallback_mask] = _loose.dt.tz_localize(None)
        return strict

    df["Start_Date"] = (
        _reparse_resume_date(_plan_rows_src["Start_Date"])
        if "Start_Date" in _plan_rows_src.columns else None
    )
    df["End_Date"] = (
        _reparse_resume_date(_plan_rows_src["End_Date"])
        if "End_Date" in _plan_rows_src.columns else None
    )
    for _carry_col in ("Observation", "_OpenEnd"):
        if _carry_col in _plan_rows_src.columns:
            df[_carry_col] = _plan_rows_src[_carry_col]
    if _ptr_ready:
        df = pd.concat([df, pd.DataFrame(_ptr_ready)], ignore_index=True)
    # Same check as a normal file: a blank/text LienId (or a blank plan) can't be processed.
    df["LienId"] = df["LienId"].map(_clean_lien_id)
    _bad_resume_rows = ~df["LienId"].str.fullmatch(r"\d+").fillna(False).astype(bool)
    _bad_resume_rows |= df["MCO_Name"].isna() | (df["MCO_Name"].astype(str).str.strip() == "")
    if _bad_resume_rows.any():
        _bad_list = [f"LienId={l!r}, MCO_Name={'' if pd.isna(m) else str(m)!r}"
                     for l, m in zip(df.loc[_bad_resume_rows, "LienId"], df.loc[_bad_resume_rows, "MCO_Name"])]
        _gap()
        print(f"⚠️  Skipped {len(_bad_list)} resume row(s) with no valid LienId (blank or not a number) "
              f"or a blank MCO_Name: " + "; ".join(_bad_list))
        df = df[~_bad_resume_rows].copy()
    _print_plan_text_warnings()
    ev_df = _resume_combined
    LIENID_COL_IN_EVDF = "LienId"
    MISMATCH_FLAGGED_LIENIDS = set()
    print(f"✅ Resumed rows: {len(df)}")

elif IS_LA_FILE:
    # LA file processing
    df, ev_input_df, LIENID_COL_IN_EVDF, MISMATCH_FLAGGED_LIENIDS, SSN_COL_IN_EVDF = process_la_file(MCO_INPUT_FILE)
    _print_plan_text_warnings()
    ev_df = ev_input_df  # output mein Input_File_Original ke liye -- ab Entitlement tab hai (LienId sahit),
                          # EV/Request/Response tab nahi (jahan asal MCO plan/date columns hain)
    print(f"✅ LA rows extracted: {len(df)}")

else:
    # Normal file processing
    # Real gap found (user report): reading with no sheet_name defaults to
    # pandas' sheet_name=0 (silently just the FIRST tab) -- if the real data
    # sits in a different tab, that would be the wrong one, with no warning at
    # all. Only ask when there's an actual choice to make -- a single-tab file
    # (the overwhelming majority) proceeds exactly as before, no new prompt.
    def _read_sheet_names():
        # Explicit close (not just letting the ExcelFile object fall out of
        # scope) -- relying on garbage collection to release the file handle
        # left it open long enough, on Windows, to occasionally still be
        # locked by the very next read of this SAME file a few lines down.
        with pd.ExcelFile(MCO_INPUT_FILE) as _xls:
            return _xls.sheet_names
    _input_sheet_names = _retry_on_permission_error(_read_sheet_names, "the import file")
    _input_sheet = (
        ask_numbered_choice(_input_sheet_names, "This file has multiple tabs. Which tab should this run use?",
                             "Enter the number: ")
        if len(_input_sheet_names) > 1 else _input_sheet_names[0]
    )
    ev_df = _strip_control_chars_df(_retry_on_permission_error(
        lambda: pd.read_excel(MCO_INPUT_FILE, sheet_name=_input_sheet, dtype=str), "the import file"
    ))
    ev_df.columns = _clean_headers(ev_df.columns)

    _column_map_hit = _try_column_map_lookup(STATE, list(ev_df.columns))
    # An old saved mapping that uses one column for both is ignored (asked again).
    if _column_map_hit and _column_map_hit[0] in _column_map_hit[1]:
        _column_map_hit = None
    if _column_map_hit:
        lienid_col, MCO_COLUMNS = _column_map_hit
        print(
            f"✅ Column mapping reused from a previous {STATE} run: "
            f"LienId={lienid_col!r}, MCO={MCO_COLUMNS!r}"
        )
    else:
        _phase_banner("column", "COLUMN IDENTIFICATION")
        _preview_columns(ev_df)
        lienid_col = _ask_columns(
            ev_df,
            "LienId: Enter the column number: ",
            allow_empty=False, multi=False, show_list=False
        )[0]

        while True:
            _gap()
            MCO_COLUMNS = _ask_columns(
                ev_df,
                "MCO: Enter column number(s), comma-separated: ",
                allow_empty=False, multi=True, show_list=False
            )
            if lienid_col not in MCO_COLUMNS:
                break
            print(f"This column is already chosen as LienId: {lienid_col!r}. Please choose the MCO column(s) again.")

        if STATE:
            _COLUMN_MAP_SAVES.append((STATE, lienid_col, MCO_COLUMNS, TicketId))

    ev_df, MCO_COLUMNS = _set_lienid_column(ev_df, lienid_col, MCO_COLUMNS)
    LIENID_COL_IN_EVDF = "LienId"
    MISMATCH_FLAGGED_LIENIDS = set()

    def has_any_plan(row):
        for c in MCO_COLUMNS:
            raw = row[c]
            if pd.notna(raw) and str(raw).strip():
                # A date with no readable plan still goes on, to Pending_Plan_Text_Review.
                if _RE_ANY_DATE_SHAPE.search(str(raw)):
                    return True
                val = _normalize_mco_text(str(raw).strip().upper())
                val_clean = _RE_STRIP_SPACE.sub('', val)
                if val_clean in INVALID:
                    continue
                val_words = set(val.replace(',', ' ').split())
                if val_words.issubset(INVALID | PLAN_LEAD_IN_WORDS):
                    continue
                # The whole cell says there is no plan ("FEE FOR SERVICE", "CLIENT DOESN'T KNOW").
                if all(_not_a_plan_text(_strip_lead_in(p.strip())[0] or p.strip())
                       for p in val.split(",") if p.strip()):
                    _add_remark(LIEN_REMARKS_EXTRA, row.get("LienId"), f"Not a plan name: '{_short_text(raw, 80)}'")
                    continue
                if _RE_HAS_ALPHA.search(val):
                    return True
        return False

    # A tab with headers but no data rows: apply() would return an empty frame-shaped
    # result, so the mask is built explicitly.
    _plan_mask = (ev_df.apply(has_any_plan, axis=1).astype(bool) if len(ev_df)
                  else pd.Series(False, index=ev_df.index, dtype=bool))
    ev_plan_df = ev_df.loc[_plan_mask].copy()
    ev_plan_df, _skipped_lienid_rows = _drop_rows_without_lienid(ev_plan_df, MCO_COLUMNS)
    if _skipped_lienid_rows:
        print(f"\n⚠️  Skipped {len(_skipped_lienid_rows)} row(s) with MCO text but no valid LienId "
              f"(blank or not a number). These rows were not processed:")
        for _skipped_line in _skipped_lienid_rows:
            print(f"   {_skipped_line}")

    rows = []
    _PLAN_TEXT_HELPER = _PlanTextHelper(STATE)
    for _, row in ev_plan_df.iterrows():
        for c in MCO_COLUMNS:
            _raw = row[c]
            if pd.isna(_raw) or not str(_raw).strip():
                continue
            _raw = str(_raw).strip()
            if _RE_STRIP_SPACE.sub('', _raw.upper()) in INVALID:
                continue
            _PLAN_TEXT_HELPER.lien_id = row["LienId"]
            try:
                _parsed = extract_plan_rows_detailed(_raw, _PLAN_TEXT_HELPER)
            except Exception as _cell_err:
                # Safety net: a cell the reader fails on goes to review whole, with the error.
                _err_text = _short_text(f"{type(_cell_err).__name__}: {_cell_err}", 100)
                _queue_plan_text_review(
                    row["LienId"], f"The tool could not read this cell (error: {_err_text})",
                    raw_text=_raw, source_col=c,
                    warning=(f"LienId {row['LienId']}: the tool could not read \"{_short_text(_raw)}\" "
                             f"({_err_text}) - nothing from this cell was applied."),
                    remark="The tool could not read the plan text - see Pending_Plan_Text_Review")
                continue
            for _np in _parsed.get("not_plans", []):
                _add_remark(LIEN_REMARKS_EXTRA, row["LienId"], f"Not a plan name: '{_np}'")
            if _parsed["review"] == SPACE_REVIEW_REASON:
                _queue_plan_text_review(
                    row["LienId"], SPACE_REVIEW_REASON, raw_text=_raw, source_col=c,
                    # The raw text as written (its 2+ spaces are the point of the question).
                    warning=(f"LienId {row['LienId']}: not sure if \"{_raw[:117] + '...' if len(_raw) > 120 else _raw}\" "
                             f"is one plan or several - nothing from this cell was applied."),
                    remark="One plan or several? - see Pending_Plan_Text_Review")
                continue
            if _parsed["review"]:
                _queue_plan_text_review(row["LienId"], _parsed["review"], raw_text=_raw, source_col=c)
                continue
            for plan, start, end, obs, open_end in _parsed["plans"]:
                rows.append({"LienId": row["LienId"], "MCO_Name": plan, "Start_Date": start, "End_Date": end,
                             "Observation": " | ".join(obs) or None, "_OpenEnd": open_end})
    _print_plan_text_warnings()

    df = pd.DataFrame(rows)

# =====================================================
# SQL — CLIENT ID QUERY (and LienIds LPM doesn't have)
# =====================================================
# Runs before the empty check: a LienId LPM doesn't have is skipped with a warning and
# Remarks "LienId not found in LPM" (it used to reach Step 3 with no client or scope
# and crash), and a file where none is found still ends cleanly.
sql_query, sql_df = None, pd.DataFrame()
if not df.empty:
    df["LienId"] = df["LienId"].map(_clean_lien_id)
    lien_ids = make_sql_in(df["LienId"].unique())

    sql_query = f"""
    SELECT
        Id AS LienId,
        ClientId,
        ClientLastName  AS Last_Name,
        ClientFirstName AS First_Name,
        RequiredMedicaidMCOLienCount,
        Tort,
        LEFT(LienHolderName,2) AS State
    FROM FullProductViews
    WHERE Id IN {lien_ids}
    """

    sql_df = _strip_control_chars_df(_read_sql(sql_query))
    # Plain ID text: one NULL ClientId would make the column float ("12345.0") and break
    # every later "ClientId in (...)" query. A lien with no client is treated as not in LPM.
    sql_df["LienId"] = sql_df["LienId"].map(_id_key)
    sql_df["ClientId"] = sql_df["ClientId"].map(_id_key)
    sql_df = sql_df[sql_df["LienId"].notna() & sql_df["ClientId"].notna()].copy()
    df = _drop_lien_ids_not_in_lpm(df, sql_df["LienId"])

# =====================================================
# EMPTY CHECK
# =====================================================

if df.empty:
    print("⚠️  No MCO plans to process in this file.")
    _empty_run_remarks = _input_file_remarks(ev_df, LIENID_COL_IN_EVDF, SSN_COL_IN_EVDF)
    _empty_remarks_col = _tool_remarks_col(ev_df)
    if any(_empty_run_remarks):
        ev_df[_empty_remarks_col] = _empty_run_remarks
    if PENDING_PLAN_TEXT_ROWS:
        # Nothing to upload, only review: no WIP -- Input_File_Original goes into the Review file.
        print("ℹ️  Nothing to upload this round -- no WIP file was created; the Review file also has "
              "Input_File_Original with its Remarks.")
        _write_plan_text_review_only(REVIEW_FILE, TicketId, RESPONSE_DATE, RESPONSE_TYPE, ev_frame=ev_df,
                                     remarks_col=_empty_remarks_col)
    else:
        _ensure_writable(OUTPUT_FILE)
        with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
            ev_df.to_excel(writer, sheet_name="Input_File_Original", index=False)
        print(f"✅ Output file (input only): {OUTPUT_FILE}")

else:

    # Observation (parse notes) and _OpenEnd (an end given as CURRENT/PRESENT) ride along
    # every grouping: an open end keeps End_Date blank even if another row has an end.
    for _carry_col, _carry_default in (("Observation", None), ("_OpenEnd", False)):
        if _carry_col not in df.columns:
            df[_carry_col] = _carry_default
    df["_OpenEnd"] = df["_OpenEnd"].map(_is_true_flag)
    yes_with_mco_df = (
        df.groupby(["LienId", "MCO_Name"])
        .agg(Start_Date=("Start_Date", safe_min), End_Date=("End_Date", safe_max),
             Observation=("Observation", _join_observations), _OpenEnd=("_OpenEnd", "any"))
        .reset_index()
    )
    yes_with_mco_df.loc[yes_with_mco_df["_OpenEnd"], "End_Date"] = None

    sql_df["RequiredMedicaidMCOLienCount"] = pd.to_numeric(
        sql_df["RequiredMedicaidMCOLienCount"], errors="coerce"
    )
    rmc_disc_df = sql_df[sql_df["RequiredMedicaidMCOLienCount"] != 9].copy()

    # =====================================================
    # SQL — MCOs (EXISTING PLANS)
    # =====================================================

    client_ids = make_sql_in(sql_df["ClientId"].dropna().unique())

    sql_mcos_query = f"""
    Select 
        ClientId,
        f.Id,
        LienHolderName,
        Stage,
        CollectorName,
        LienType,
        c.RequiredMedicaidMCOLienCount,
        f.IsMCO,         
        Case 
            when c.StatusId = 10 then 'Active'    
            when c.StatusId = 11 then 'Inactive'
        end as [Client Status]
    from fullproductviews f 
    left join clientsviews c on f.clientid = c.id        
    Where 
        LienHolderName not like '%Medicaid%' 
        and lientype not like 'Military Lien - MT' 
        and LienholderName not like 'Placeholder'
        and CollectorId not in (
            '11','42','50','453','297','62','89','119','120','129',
            '622','151','144','166','164','167','173','210','176',
            '179','196','213','216','220','378','232','259','260',
            '268','291','294','620','293','295','305','736','624','174'
        )
        and ClientId in {client_ids}
    Order by ClientId
    """

    sql_mcos_df = _strip_control_chars_df(_read_sql(sql_mcos_query))
    sql_mcos_df["ClientId"] = sql_mcos_df["ClientId"].astype(str)

    # =====================================================
    # SQL — LIENS QUERY
    # =====================================================

    sql_liens_query = f"""
    select
        f.ClientSSN, f.clientfirstname, f.clientlastname, f.ClientId,
        f.Id as LienId, f.Stage, f.LienType,
        f.CollectorId, f.LienholderId, f.AssignedUserId,
        f.statusid, f.LienHolderName, f.finalizedstatusname, f.finalizedstatusid,
        f.IsMCO, f.EnrollmentStartDate, f.EnrollmentEndDate, c.PrivateLiensScope, clientadditionalinformation,
        f.Tort, f.TortId, LEFT(f.LienHolderName,2) AS State, f.CaseId
    from datalienview f
    inner join cases c on f.caseid = c.id
    where f.ClientId in {client_ids}
    Order by 5, 2
    """

    sql_liens_df = _strip_control_chars_df(_read_sql(sql_liens_query))
    sql_liens_df["ClientId"] = sql_liens_df["ClientId"].astype(str)
    sql_liens_df["LienId"]   = sql_liens_df["LienId"].astype(str)
    # Plain ID text: a NULL in a numeric column makes pandas read it as float, and
    # "300.0" never equals "300" (every plan would then look like a new lien).
    for _id_col in ("TortId", "CollectorId", "LienholderId", "AssignedUserId"):
        if _id_col in sql_liens_df.columns:
            sql_liens_df[_id_col] = sql_liens_df[_id_col].map(_id_key).astype(object)
    # CaseId keys a scope decision when the case's PrivateLiensScope is blank in LPM.
    sql_liens_df["CaseId"]   = sql_liens_df["CaseId"].map(_id_key) if "CaseId" in sql_liens_df.columns else None
    sql_liens_df["Last_Name"]  = sql_liens_df["clientlastname"]
    sql_liens_df["First_Name"] = sql_liens_df["clientfirstname"]

    # Everything simple-lookup now comes from ONE source (sql_liens_df), ONE call, keyed on LienId.
    # TortId rides along purely for Step 3's Pending_TPA_Review (so a not-found-in-TPA Tort still
    # gets a real, DB-sourced TortId auto-populated on that review row -- the user never has to
    # supply one) -- it naturally falls out of yes_with_mco_df again at FINAL COLUMN ORDER below,
    # same as CollectorName, since nothing after Step 3 needs it.
    yes_with_mco_df = bring_in_columns(
        yes_with_mco_df, sql_liens_df, "LienId",
        ["ClientId", "Tort", "TortId", "State", "Last_Name", "First_Name", "PrivateLiensScope", "clientadditionalinformation",
         "CaseId"]
    ).rename(columns={"Tort": "Project_Name"})
    # Backstop: a LienId SQL-Liens doesn't return (no client) is skipped the same way.
    yes_with_mco_df = _drop_lien_ids_not_in_lpm(yes_with_mco_df, yes_with_mco_df["LienId"], client_col="ClientId")

    # State comes from each lien's own LienHolderName ("KY ..."). Blank there -> the run's
    # state (one per file; on resume, the one state of the other rows) + Observation.
    # Still unknown -> left blank; Step 3 then sends the plan to Pending_Error_Review.
    _blank_state = yes_with_mco_df["State"].map(
        lambda v: not re.fullmatch(r"[A-Z]{2}", str(v).strip().upper()) if pd.notna(v) else True)
    if _blank_state.any():
        _known_states = {str(v).strip().upper() for v in yes_with_mco_df.loc[~_blank_state, "State"]}
        _fill_state = STATE if STATE else (next(iter(_known_states)) if len(_known_states) == 1 else None)
        if _fill_state:
            yes_with_mco_df.loc[_blank_state, "Observation"] = [
                _join_observations([o, f"State was blank in LPM - used the run's state '{_fill_state}'"])
                for o in yes_with_mco_df.loc[_blank_state, "Observation"]]
            yes_with_mco_df.loc[_blank_state, "State"] = _fill_state
        else:
            yes_with_mco_df.loc[_blank_state, "State"] = None
    yes_with_mco_df["State"] = yes_with_mco_df["State"].map(
        lambda v: str(v).strip().upper() if pd.notna(v) and str(v).strip() else None)

    def normalize_mco_name(row):
        state = str(row.get("State", "")).strip().upper()
        plan  = str(row.get("MCO_Name", "")).strip().upper()
        target = ALIAS_MAP.get(state, {}).get(_RE_ANY_SPACE_RUN.sub(" ", plan), plan)
        # "A | B" means separate plans (split while reading the text); never a plan name.
        return plan if "|" in str(target) else target

    yes_with_mco_df["MCO_Name"] = yes_with_mco_df.apply(normalize_mco_name, axis=1)

    grp_keys  = ["LienId", "ClientId", "Project_Name", "TortId", "State", "Last_Name", "First_Name", "MCO_Name", "PrivateLiensScope", "clientadditionalinformation",
                 "CaseId"]
    date_cols = [c for c in ["Start_Date", "End_Date"] if c in yes_with_mco_df.columns]

    # dropna=False on both groupbys below is required -- grp_keys includes columns like
    # clientadditionalinformation/PrivateLiensScope that are routinely blank for a real
    # client. pandas' groupby DROPS a row entirely (silently, no warning) if ANY of its
    # group-key columns is NaN by default -- that would have thrown away every plan for
    # a client with no additional-info notes before it ever reached Scope/TPA/Step 4.
    if date_cols:
        yes_with_mco_df = (
            yes_with_mco_df
            .groupby(grp_keys, dropna=False)
            .agg(Start_Date=("Start_Date", lambda x: safe_min(list(x))),
                 End_Date  =("End_Date",   lambda x: safe_max(list(x))),
                 Observation=("Observation", _join_observations), _OpenEnd=("_OpenEnd", "any"))
            .reset_index()
        )
        yes_with_mco_df.loc[yes_with_mco_df["_OpenEnd"], "End_Date"] = None
        yes_with_mco_df["Start_Date"] = yes_with_mco_df["Start_Date"].apply(format_date)
        yes_with_mco_df["End_Date"]   = yes_with_mco_df["End_Date"].apply(format_date)
    else:
        yes_with_mco_df = (
            yes_with_mco_df.groupby(grp_keys, dropna=False)
            .agg(Observation=("Observation", _join_observations), _OpenEnd=("_OpenEnd", "any"))
            .reset_index()
        )

    if date_cols:
        if not (yes_with_mco_df["Start_Date"].notna().any() or
                yes_with_mco_df["End_Date"].notna().any()):
            yes_with_mco_df = yes_with_mco_df.drop(columns=["Start_Date", "End_Date"])
            date_cols = []

    # =====================================================
    # PLAN FATE TRACKING — accountability for every qualified (LienId, MCO_Name)
    # =====================================================
    # PLAN_FATE[(LienId, MCO_Name)] records what ultimately happened to each MCO plan,
    # shown in Input_File_Original's Remarks column. Defaults to a visibly-wrong
    # placeholder on purpose, so an unrecorded fate surfaces as a gap to fix.
    #
    # Built HERE (after alias-expansion + the post-alias dedup groupby), not right
    # after the raw groupby, because MCO_Name gets rewritten to its canonical form
    # further down -- keying on the raw pre-alias name would strand the placeholder
    # while the real outcome lands under the aliased name.
    PLAN_FATE = {
        (r["LienId"], r["MCO_Name"]): "Not yet determined (tracking gap -- please report this)"
        for _, r in yes_with_mco_df[["LienId", "MCO_Name"]].iterrows()
    }

    def record_fate(df_subset, remark):
        """Mark every (LienId, MCO_Name) pair in df_subset with `remark`. Safe to call
        multiple times for the same key -- later calls simply update the fate (used
        e.g. when Step 4 gives a plan its final, specific outcome)."""
        for _, r in df_subset[["LienId", "MCO_Name"]].drop_duplicates().iterrows():
            PLAN_FATE[(r["LienId"], r["MCO_Name"])] = remark

    # =====================================================
    # EXCLUDE FINALIZED CLIENTS
    # =====================================================
    # Before any scope/TPA question, so nobody answers questions for a finalized client.

    finalized_df = sql_liens_df[sql_liens_df["finalizedstatusid"].isin([2, 3])].copy()
    finalized_clients = [] if finalized_df.empty else finalized_df["ClientId"].unique().tolist()

    finalized_excluded_df = (
        finalized_df[[
            "ClientId", "clientfirstname", "clientlastname",
            "ClientSSN", "finalizedstatusid", "finalizedstatusname"
        ]].drop_duplicates()
    )

    if finalized_clients:
        fc = set(finalized_clients)
        record_fate(
            yes_with_mco_df[yes_with_mco_df["ClientId"].isin(fc)],
            "Client finalized; excluded -- see Finalized_Clients_Excluded tab"
        )
        sql_df          = sql_df[~sql_df["ClientId"].isin(fc)].copy()
        rmc_disc_df     = rmc_disc_df[~rmc_disc_df["ClientId"].isin(fc)].copy()
        sql_mcos_df     = sql_mcos_df[~sql_mcos_df["ClientId"].isin(fc)].copy()
        sql_liens_df    = sql_liens_df[~sql_liens_df["ClientId"].isin(fc)].copy()
        yes_with_mco_df = yes_with_mco_df[~yes_with_mco_df["ClientId"].isin(fc)].copy()

    # =====================================================
    # COLUMN ENRICHMENT — every mapping applied onto yes_with_mco_df, one after another
    # =====================================================

    # MCOn_Existing columns <- sql_mcos_df.LienHolderName, grouped by ClientId
    existing_mco_map = (
        sql_mcos_df.groupby("ClientId")["LienHolderName"].apply(list).to_dict()
    )
    max_mcos = max((len(v) for v in existing_mco_map.values()), default=0)

    for i in range(max_mcos):
        yes_with_mco_df[f"MCO{i+1}_Existing"] = yes_with_mco_df["ClientId"].apply(
            lambda x, _i=i: existing_mco_map[x][_i]
            if x in existing_mco_map and len(existing_mco_map[x]) > _i else None
        )

    # =====================================================
    # NAME-MATCHING HELPERS (shared by Step 3's TPA-based lookup below)
    # =====================================================

    _word_cache = {}

    # Generic connector words -- never enough on their own to justify a match,
    # since many unrelated real plans share them ("Community Health Plan" vs
    # "Community Care Alliance" -- 2 different real payers sharing only
    # "community"). Never removed from tokenization/scoring, only consulted by
    # the corroboration check inside _match_name.
    _GENERIC_CONNECTORS = {
        "health", "healthy", "plan", "plans", "services", "managed", "group",
        "community", "system", "systems", "national", "insurance", "care",
    }
    # Program-type/state tags -- within one state's own sheet, every row shares
    # these, so they carry zero distinguishing power and never count as real
    # corroboration (distinct from _GENERIC_CONNECTORS, excluded from scoring
    # context too, not just this count).
    _STRUCTURAL_TAGS = {
        "mco", "cco",
        "delaware", "louisiana", "kentucky", "illinois", "indiana",
        "oregon", "minnesota", "michigan", "washington", "missouri",
    }

    def _tokenize_plan(plan_norm: str):
        """Words of length >=3 only -- 1-2 letter fragments ("in", "of", "a") carry no
        identifying signal and can coincidentally appear inside unrelated words (see
        _match_name's docstring). A leading grammatical article ("the"/"a"/"an") is
        dropped too, since it's a data-entry habit, not part of a plan's identity --
        _match_name's first-word requirement would otherwise reject a name like "The
        Molina Group" on "the" alone. Articles are the ONLY words stripped this way
        (no general stopword list): "the"/"a"/"an" can never themselves be a plan's
        distinguishing word, unlike "health"/"services"/"managed", which sometimes
        genuinely are."""
        if plan_norm not in _word_cache:
            raw_words = re.findall(r"[a-z]+", plan_norm)
            words = []
            for w in raw_words:
                if w.endswith("healthcare"):
                    base = w.replace("healthcare", "")
                    if base:
                        words.append(base)
                    words.append("healthcare")
                elif w.endswith("health"):
                    # Same reasoning as "healthcare" above, but for TPA names that glue
                    # a brand directly onto "Health" with no "care" (e.g. real TPA row
                    # "MeridianHealth of Michigan..." -- a bare "Meridian" query would
                    # never gate-match "meridianhealth" as one indivisible token).
                    base = w[: -len("health")]
                    if base:
                        words.append(base)
                    words.append("health")
                else:
                    words.append(w)
            words = [w for w in words if w and len(w) >= 3]
            if len(words) > 1 and words[0] in ("the", "a", "an"):
                words = words[1:]
            _word_cache[plan_norm] = words
        return _word_cache[plan_norm]

    def _singularize(w):
        """Naive singular form, for word-MATCHING purposes only (gate/shared/leftover
        comparisons) -- never changes what a token actually displays as anywhere else.
        Strips one trailing 's' when the word is long enough that doing so still
        leaves a real word (>=4 letters, so a short word like "has" is never touched)
        and isn't a double-s ending ("services"/"class" stay as-is). Real bug found
        (user report): "Louisiana Healthcare Connection" (query) failed to match TPA's
        "Louisiana Healthcare Connections /Centene" purely because of singular vs
        plural -- both sides then looked like they had real unaccounted leftover
        content, tripping the corroboration guard on a false mismatch. _GENERIC_
        CONNECTORS already hand-lists both forms of a few words ("plan"/"plans",
        "system"/"systems") for exactly this reason; this covers every OTHER word
        the same way instead of needing every such pair listed by hand."""
        if len(w) >= 4 and w.endswith("s") and not w.endswith("ss"):
            return w[:-1]
        return w

    def _match_name(plan_text, candidates_df, name_col, id_col, return_tie=False):
        """Source-agnostic fuzzy matcher: prefix-match first, falls back to whole-word
        overlap scored as a RATIO of the query's own words, not a bare count. Used by
        Step 3's TPA lookup and Step 4's existing-lien lookup.

        Design rules:
        1. Both sides tokenize into whole words -- a plain substring search would let
           a short word like "in" (from "...(IN MCO)...") match inside unrelated
           words like "Inc." or "Rawlings".
        2. Score = shared words / query's word count, not a raw count -- a name
           explaining only a slice of the query shouldn't pass, but a short name
           should if it explains all of it (e.g. "Managed Health Services" alone).
           Existing-lien names (collection agencies, payer categories) often share
           one ordinary word ("Health"/"Services") with an unrelated plan by
           coincidence, so "1 shared word" alone is too weak. No stopword list
           (tried once, reverted): generic words like "managed"/"health" are
           sometimes a real plan's ONLY identifying words.
        3. The query's first word is normally required -- it's the brand/identity
           token, while trailing words ("Health", "System", "MCO") are generic
           filler. Exception: a "/" outside parens (co-branded name like
           "HealthShare/CareOregon") also accepts the first word after it, since
           either half can be the real identity.
        4. Corroboration guard: passing the gate word isn't enough alone. Rejected
           if, after setting aside _STRUCTURAL_TAGS (mco/cco/state names -- no
           signal within one state's own sheet): (a) every shared word is a
           _GENERIC_CONNECTOR while the query still has other unaccounted content,
           or (b) the candidate's own leftover is "risky" (a real word, or 2+
           generic ones together -- not just one harmless suffix like "Passport
           Group"'s "Group") AND the query also has real unaccounted content. Only
           fires when BOTH sides have real dangling content left (e.g.
           "MCO - United Healthcare" vs "United Healthcare" -- "MCO" could itself be
           a real identity, so refused rather than guessed).
        5. Ties: if the best match is shared by 2+ candidates with DIFFERENT names,
           that's coincidence, not confidence -- returns None. Candidates sharing
           the exact SAME name (different ids) are NOT a tie -- callers disambiguate
           that themselves (see resolve_tpa_lookup's scope tie-break).

        return_tie=True returns (matched_id, tied_ids) instead -- tied_ids lists
        candidates in a genuine tie (empty otherwise). Used by callers (Step 4) that
        must route a real tie to manual review instead of falling through to "no
        match".
        """
        if candidates_df.empty or pd.isna(plan_text):
            return (None, []) if return_tie else None
        plan_norm  = str(plan_text).strip().lower()
        plan_clean = plan_norm.replace(" ", "")
        words = _tokenize_plan(plan_norm)

        # Prefix match -- skip nameless candidates so "".startswith("") can never
        # trivially fire for a blank/NaN existing-lien name. Track the LONGEST
        # matching prefix(es) -- a longer stored name explains more of the query text,
        # so "United Healthcare" rightly wins over a shorter "United" if both qualify;
        # only a genuine tie in length AND a different name is treated as ambiguous.
        best_prefix_len, best_prefix_id, best_prefix_name, prefix_tie = -1, None, None, False
        prefix_tied_ids = []
        for _, r in candidates_df.iterrows():
            name_norm = str(r[name_col]).strip().lower()
            if not (name_norm and name_norm != "nan" and plan_clean.startswith(name_norm.replace(" ", ""))):
                continue
            name_key = _normalize_key(name_norm)
            if len(name_norm) > best_prefix_len:
                best_prefix_len, best_prefix_id, best_prefix_name = len(name_norm), r[id_col], name_key
                prefix_tie = False
                prefix_tied_ids = [r[id_col]]
            elif len(name_norm) == best_prefix_len and name_key != best_prefix_name:
                prefix_tie = True
                prefix_tied_ids.append(r[id_col])
        if best_prefix_id is not None:
            if return_tie:
                return (None, prefix_tied_ids) if prefix_tie else (best_prefix_id, [])
            return None if prefix_tie else best_prefix_id

        if not words:
            return (None, []) if return_tie else None

        # Parenthetical annotations ("(OR CCO)", "(Formerly WOAH)") are administrative
        # tags, excluded from the corroboration guard only (never gate/scoring). The
        # "/" co-branding exception (rule 3) only applies outside such an annotation --
        # a "/" inside one (e.g. "(Wellcare/Centene)") just names a parent company.
        paren_words = set()
        for paren_content in re.findall(r"\(([^)]*)\)", plan_norm):
            paren_words.update(_tokenize_plan(paren_content))
        plan_norm_no_parens = re.sub(r"\([^)]*\)", " ", plan_norm)

        gate_words = [words[0]]
        slash_exempt = set()
        if "/" in plan_norm_no_parens:
            after_slash = plan_norm_no_parens.split("/", 1)[1]
            slash_words = _tokenize_plan(after_slash.strip())
            if slash_words:
                gate_words.append(slash_words[0])
                slash_exempt.add(words[0])
                slash_exempt.add(slash_words[0])

        best_id, best_score, best_name_key, tie = None, 0.0, None, False
        tied_ids = []
        for _, r in candidates_df.iterrows():
            name_norm = str(r[name_col]).strip().lower()
            if not name_norm or name_norm == "nan":
                continue
            cand_words = set(_tokenize_plan(name_norm))
            if not cand_words:
                continue
            cand_words_sing = {_singularize(cw) for cw in cand_words}
            if not any(_singularize(gw) in cand_words_sing for gw in gate_words):
                continue
            shared = [w for w in words if _singularize(w) in cand_words_sing]

            effective_shared = [w for w in shared if w not in _STRUCTURAL_TAGS]
            leftover_query_raw = set(words) - set(shared)
            if not effective_shared:
                continue  # everything shared was purely a state/program-type tag
            # Real bug found: leftover_query_raw was unfiltered, so a pure
            # _STRUCTURAL_TAGS leftover (e.g. a state name, no signal within one
            # state's own sheet) wrongly counted as "real unaccounted content,"
            # rejecting a match the filtered leftover check below would have allowed.
            if all(w in _GENERIC_CONNECTORS for w in effective_shared) and (leftover_query_raw - _STRUCTURAL_TAGS):
                continue  # only generic words shared, and the query has other unexplained content
            # The candidate's own leftover only counts against it when "risky" -- one
            # generic trailing word ("Passport Group"'s "group") is harmless filler,
            # but a real word or 2+ generic ones together means real unexplained
            # identity -- only then, combined with the query ALSO having real
            # unaccounted content, is the match too uncertain to guess.
            # Real bug found: a leftover _STRUCTURAL_TAG (state name, no signal within
            # one state's sheet) was counted as "real unaccounted content" on either
            # side -- e.g. query "WellCare of Kentucky/Centene" left "kentucky"
            # unaccounted, combined with candidate "WellCare Health Plans/Centene"'s
            # own leftover ("health"+"plans") to wrongly refuse a same-plan match.
            # Excluded here like generic connectors already are, on both sides.
            shared_sing = {_singularize(w) for w in shared}
            leftover_candidate = {
                cw for cw in cand_words if _singularize(cw) not in shared_sing and cw not in _STRUCTURAL_TAGS
            }
            leftover_candidate_generic = [w for w in leftover_candidate if w in _GENERIC_CONNECTORS]
            leftover_candidate_risky = (
                len(leftover_candidate_generic) < len(leftover_candidate)
                or len(leftover_candidate_generic) >= 2
            )
            leftover_query = leftover_query_raw - _GENERIC_CONNECTORS - slash_exempt - paren_words - _STRUCTURAL_TAGS
            if leftover_candidate_risky and leftover_query:
                continue

            score = len(shared) / len(words) + 0.5
            name_key = _normalize_key(name_norm)
            if score > best_score:
                best_score, best_id, best_name_key = score, r[id_col], name_key
                tie = False
                tied_ids = [r[id_col]]
            elif score == best_score and best_id is not None and name_key != best_name_key:
                tie = True
                tied_ids.append(r[id_col])
        if best_score >= 0.5 and not tie:
            return (best_id, []) if return_tie else best_id
        if return_tie:
            return (None, tied_ids if tie else [])
        return None

    def _find_overlap_candidates(plan_text, candidates_df, name_col):
        """Looser than _match_name -- no gate-word/prefix/corroboration rules, just
        "do the two names share at least one real (non-generic, non-structural)
        word". Used only by Step 4's classify_step4_scenario to flag a POSSIBLE
        existing-lien match for human review when _match_name itself can't confirm
        one confidently (real cases: "Healthy Blue/Anthem BCBS" missed "Wellpoint,
        Inc. (Anthem BCBS/Elevance)"; "NextLevel HP/Meridian (Wellcare/Centene IL
        MCO)" missed "WellCare Health Plans/Centene") -- never used to auto-resolve
        a scenario on its own, so being looser than _match_name here is safe.

        Requiring a REAL shared word (not just any shared word) is what keeps this
        from flooding review with coincidental collector-mates: a candidate like
        "Humana" sharing nothing with the query is correctly never flagged."""
        if candidates_df.empty or pd.isna(plan_text):
            return candidates_df.iloc[0:0]
        plan_words = {_singularize(w) for w in _tokenize_plan(str(plan_text).strip().lower())}
        real_plan_words = plan_words - _GENERIC_CONNECTORS - _STRUCTURAL_TAGS
        if not real_plan_words:
            return candidates_df.iloc[0:0]

        def _has_real_overlap(name):
            if pd.isna(name):
                return False
            cand_words = {_singularize(w) for w in _tokenize_plan(str(name).strip().lower())}
            return bool(real_plan_words & (cand_words - _GENERIC_CONNECTORS - _STRUCTURAL_TAGS))

        return candidates_df[candidates_df[name_col].apply(_has_real_overlap)]

    # =====================================================
    # SCOPE CLASSIFICATION (Step 2) -- classify_scope reads SCOPE_MAP
    # =====================================================

    def classify_scope(raw_scope_text):
        key = _normalize_key(raw_scope_text)
        if key is None:
            return None
        classification = SCOPE_MAP.get(key)
        if classification in ("100% PLRP", "NO PLRP"):
            return classification
        return None

    # clientadditionalinformation is free-text, so "PLRP" + an opt-in/opt-out phrase
    # must read as ONE phrase ("PLRP opt out"), not just both appearing in the note.
    # Gap between them: letters/whitespace only, max 15 chars -- punctuation or an
    # unrelated clause in between breaks the match on purpose (different mention).
    # If found, OVERRIDES the client's classified scope (opt-out -> No PLRP, opt-in
    # -> 100% PLRP) regardless of PrivateLiensScope; otherwise no override happens.
    _ADDINFO_GAP = r'[a-zA-Z\s\-]{0,15}'

    _RE_ADDINFO_OPT_OUT_NEAR_PLRP = re.compile(
        r'(plrp' + _ADDINFO_GAP + r'opt(?!ion)\w*[\s\-]*out\b)'
        r'|(opt(?!ion)\w*[\s\-]*out\b' + _ADDINFO_GAP + r'plrp)',
        re.IGNORECASE
    )
    _RE_ADDINFO_OPT_IN_NEAR_PLRP = re.compile(
        r'(plrp' + _ADDINFO_GAP + r'opt(?!ion)\w*[\s\-]*in\b)'
        r'|(opt(?!ion)\w*[\s\-]*in\b' + _ADDINFO_GAP + r'plrp)',
        re.IGNORECASE
    )

    def classify_additional_info(text):
        if pd.isna(text) or not str(text).strip():
            return None
        t = str(text)
        if _RE_ADDINFO_OPT_OUT_NEAR_PLRP.search(t):
            return "OPT OUT"
        if _RE_ADDINFO_OPT_IN_NEAR_PLRP.search(t):
            return "OPT IN"
        return None

    _SCOPE_DIAG_QUERY_TEMPLATE = """
    ;WITH CasesToReview AS
    (
        SELECT DISTINCT CaseId
        FROM DataLienView
        WHERE ClientId IN {ids_clause}
    )
    SELECT
        c.Tort,
        ctr.CaseId,

        SUM(CASE WHEN dl.StatusId = 1 THEN 1 ELSE 0 END) AS [Active PLRPs],

        (
            SELECT STRING_AGG(PhaseName, ', ')
            FROM
            (
                SELECT DISTINCT p.Name AS PhaseName
                FROM DataLienView d
                JOIN ac.Stage s
                    ON s.Name = d.Stage
                   AND s.EntityTypeId = 1
                JOIN ac.Phase p
                    ON p.Id = s.PhaseId
                WHERE d.CaseId = ctr.CaseId
                  AND d.LienType = 'Private Lien - PLRP'
                  AND d.StatusId = 1
            ) x
        ) AS [Active PLRP Phases],

        SUM(CASE WHEN dl.StatusId = 2 THEN 1 ELSE 0 END) AS [Inactive PLRPs],

        (
            SELECT STRING_AGG(PhaseName, ', ')
            FROM
            (
                SELECT DISTINCT p.Name AS PhaseName
                FROM DataLienView d
                JOIN ac.Stage s
                    ON s.Name = d.Stage
                   AND s.EntityTypeId = 1
                JOIN ac.Phase p
                    ON p.Id = s.PhaseId
                WHERE d.CaseId = ctr.CaseId
                  AND d.LienType = 'Private Lien - PLRP'
                  AND d.StatusId = 2
            ) x
        ) AS [Inactive PLRP Phases]

    FROM CasesToReview ctr
    JOIN Cases c
        ON c.Id = ctr.CaseId
    LEFT JOIN DataLienView dl
        ON dl.CaseId = ctr.CaseId
       AND dl.LienType = 'Private Lien - PLRP'
    GROUP BY
        c.Tort,
        ctr.CaseId
    ORDER BY
        ctr.CaseId;
    """

    def _print_scope_diag_table(diag_result, wrap_col_width=42, wrap_threshold=30):
        """Renders the scope-diagnostic query as a gridline table instead of a raw
        DataFrame print -- some columns (a comma-joined Phase/Stage list) can run
        past 500 characters, which turns a plain print into unreadable wrapped
        clutter. Which columns wrap is decided by ACTUAL content length, not a
        hardcoded column-name pattern -- this query has already been renamed once,
        so name-matching would silently break on the next rewrite. Any column whose
        longest value exceeds wrap_threshold wraps to its own width, with the row's
        other (short) columns repeated only on its first physical line."""
        if diag_result.empty:
            print(diag_result)
            return
        cols = list(diag_result.columns)
        wrap_cols = {c for c in cols if max(len(str(v)) for v in diag_result[c]) > wrap_threshold}
        widths = {}
        for c in cols:
            if c in wrap_cols:
                widths[c] = wrap_col_width
            else:
                widths[c] = max(len(str(c)), max(len(str(v)) for v in diag_result[c]))

        def hline():
            return "+" + "+".join("-" * (widths[c] + 2) for c in cols) + "+"

        def cell(c, val):
            return f" {val:<{widths[c]}} "

        print(hline())
        print("|" + "|".join(cell(c, c) for c in cols) + "|")
        print(hline())
        for _, r in diag_result.iterrows():
            wrapped = {
                c: (textwrap.wrap(str(r[c]), width=widths[c]) or [""]) if c in wrap_cols else None
                for c in cols
            }
            n_lines = max((len(wrapped[c]) for c in wrap_cols), default=1)
            for i in range(n_lines):
                cells = []
                for c in cols:
                    if c in wrap_cols:
                        lines = wrapped[c]
                        val = lines[i] if i < len(lines) else ""
                    else:
                        val = str(r[c]) if i == 0 else ""
                    cells.append(cell(c, val))
                print("|" + "|".join(cells) + "|")
            print(hline())

    def resolve_scope_for_run(df_in, scope_map, show_diagnostic=True):
        """Classify each row's PrivateLiensScope via scope_map. For unresolved (normalized)
        text values, resolve once per unique text -- diagnostic query + ask 100% PLRP /
        No PLRP / Skip -- applying the decision to every row sharing that exact text.
        Returns (resolved_df, pending_df, scope_saves): scope_saves receives the
        (scope_text, classification) pairs chosen at the end-of-run save question."""
        working_df = df_in.copy()
        # Safety net: a row whose scope can't even be read goes to Pending_Error_Review.
        _scope_errors = {}

        def _safe_classify(idx, v):
            try:
                return classify_scope(v)
            except Exception as _e:
                _scope_errors[idx] = _e
                return None
        working_df["Scope_Classification"] = [
            _safe_classify(i, v) for i, v in zip(working_df.index, working_df["PrivateLiensScope"])]

        unresolved_mask = working_df["Scope_Classification"].isna() & ~working_df.index.isin(list(_scope_errors))
        unresolved_texts = working_df.loc[unresolved_mask, "PrivateLiensScope"].dropna().unique().tolist()

        scope_saves = []
        skip_texts = set()
        skip_rows = set()
        # The phase banner shows only when something is actually asked.
        _banner = [False]

        def _ask_scope(header_line, client_ids):
            if not _banner[0]:
                _phase_banner("scope", "SCOPE CLASSIFICATION")
                _banner[0] = True
            _new_item("scope")
            print(f"\n{header_line}")
            if show_diagnostic:
                try:
                    ids_clause = make_sql_in(client_ids)
                    diag_query = _SCOPE_DIAG_QUERY_TEMPLATE.format(ids_clause=ids_clause)
                    diag_result = _read_sql(diag_query)
                    _print_scope_diag_table(diag_result)
                except Exception as _diag_err:
                    # Only a help for the question below -- never a reason to stop.
                    print(f"(Could not show the diagnostic table: {_short_text(str(_diag_err), 120)})")
            while True:
                answer = input(
                    "Classify this scope -- enter '100' for 100% PLRP, 'No' for No PLRP, "
                    "or 'S' to skip: "
                ).strip().upper()
                if answer in ("100", "100%", "100% PLRP", "100%PLRP"):
                    return "100% PLRP"
                if answer in ("NO", "NO PLRP", "NOPLRP"):
                    return "NO PLRP"
                if answer in ("SKIP", "S"):
                    return None
                print("Please enter '100', 'No', or 'S' to skip.")

        for raw_text in unresolved_texts:
            key = _normalize_key(raw_text)
            if key is None:
                continue
            if scope_map.get(key) in ("100% PLRP", "NO PLRP"):
                working_df.loc[working_df["PrivateLiensScope"] == raw_text, "Scope_Classification"] = scope_map[key]
                continue
            matching_client_ids = working_df.loc[working_df["PrivateLiensScope"] == raw_text, "ClientId"].unique().tolist()
            classification = _ask_scope(f"Scope not recognized: {raw_text!r}", matching_client_ids)
            if classification:
                scope_map[key] = classification
                working_df.loc[working_df["PrivateLiensScope"] == raw_text, "Scope_Classification"] = classification
                _defer_save(f"Scope {raw_text!r} = {classification}", scope_saves, (raw_text, classification))
            else:
                skip_texts.add(raw_text)

        # Blank scope in LPM (no text at all; scope is case-level): the client's own
        # additional-info opt in/out decides first; else a Scope_Map row saved for this
        # CaseId; else asked ONCE per case, saved per CaseId at the end-of-run question.
        blank_mask = working_df["PrivateLiensScope"].map(_normalize_key).isna() & ~working_df.index.isin(list(_scope_errors))
        if blank_mask.any():
            case_ids = (working_df["CaseId"].map(_id_key) if "CaseId" in working_df.columns
                        else pd.Series(None, index=working_df.index, dtype=object))
            for idx in working_df.index[blank_mask]:
                BLANK_SCOPE_CASES.append(case_ids.get(idx) or f"no CaseId (ClientId {working_df.at[idx, 'ClientId']})")
            if "clientadditionalinformation" in working_df.columns:
                for idx in working_df.index[blank_mask]:
                    _ai = classify_additional_info(working_df.at[idx, "clientadditionalinformation"])
                    if _ai:
                        working_df.at[idx, "Scope_Classification"] = "NO PLRP" if _ai == "OPT OUT" else "100% PLRP"
            groups = {}
            for idx in working_df.index[blank_mask & working_df["Scope_Classification"].isna()]:
                groups.setdefault(case_ids.get(idx) or ("NO CASE", working_df.at[idx, "ClientId"]), []).append(idx)
            for gkey, idxs in groups.items():
                case = gkey if isinstance(gkey, str) else None
                saved = scope_map.get(_case_scope_key(case)) if case else None
                if saved in ("100% PLRP", "NO PLRP"):
                    working_df.loc[idxs, "Scope_Classification"] = saved
                    print(f"CaseId {case}: scope blank in LPM, using saved {saved}")
                    continue
                client_ids = list(dict.fromkeys(working_df.loc[idxs, "ClientId"]))
                label = f"CaseId {case}" if case else "a case with no CaseId"
                classification = _ask_scope(
                    f"Scope is blank in LPM for {label} ({len(client_ids)} client{'s' if len(client_ids) != 1 else ''}: "
                    f"{', '.join(str(c) for c in client_ids)})", client_ids)
                if classification:
                    working_df.loc[idxs, "Scope_Classification"] = classification
                    if case:
                        scope_map[_case_scope_key(case)] = classification
                        _defer_save(f"Scope for CaseId {case} (blank in LPM) = {classification}",
                                    scope_saves, (None, classification, case))
                else:
                    skip_rows.update(idxs)

        _error_mask = working_df.index.isin(list(_scope_errors))
        for _idx, _err in _scope_errors.items():
            _record_plan_error(working_df.loc[[_idx]], "Step 2 (scope)", _err)
        pending_mask = (working_df["PrivateLiensScope"].isin(skip_texts) | working_df.index.isin(list(skip_rows))) & ~_error_mask
        pending_df  = working_df[pending_mask].copy()
        resolved_df = working_df[~pending_mask & ~_error_mask].copy()
        return resolved_df, pending_df, scope_saves

    yes_with_mco_df, pending_scope_df, _SCOPE_SAVES = resolve_scope_for_run(yes_with_mco_df, SCOPE_MAP, show_diagnostic=not IS_RESUME_RUN)

    # Additional-info opt-in/opt-out override -- only applies to rows that already have
    # a resolved Scope_Classification (pending/skipped rows have none to override yet).
    _addinfo_result = yes_with_mco_df["clientadditionalinformation"].apply(classify_additional_info)
    yes_with_mco_df.loc[_addinfo_result == "OPT OUT", "Scope_Classification"] = "NO PLRP"
    yes_with_mco_df.loc[_addinfo_result == "OPT IN", "Scope_Classification"] = "100% PLRP"

    # Safety net (should never fire): no row may reach Step 3 without a scope -- a blank
    # one crashed the TPA lookup. Any such row goes to Pending_Scope_Review instead.
    _no_scope = yes_with_mco_df["Scope_Classification"].map(_normalize_key).isna()
    if _no_scope.any():
        _gap()
        print(f"⚠️  {int(_no_scope.sum())} plan row(s) still had no scope after scope classification -- "
              f"sent to Pending_Scope_Review (please report this).")
        pending_scope_df = pd.concat([pending_scope_df, yes_with_mco_df[_no_scope]], ignore_index=True)
        yes_with_mco_df = yes_with_mco_df[~_no_scope].copy()

    if not pending_scope_df.empty:
        pending_scope_df = pending_scope_df.copy()
        pending_scope_df["TicketId"]      = TicketId
        pending_scope_df["ResponseDate"]  = RESPONSE_DATE
        pending_scope_df["ResponseType"]  = RESPONSE_TYPE
        record_fate(pending_scope_df, "Scope not recognized; sent for review -- see Pending_Scope_Review (Review file)")

    # =====================================================
    # TPA STATE-SHEET READER (Step 3) -- dynamic sheet match + junk-row boundary
    # =====================================================
    # load_tpa_state_sheet(state) returns that state's TPA sheet as a DataFrame (or None
    # if the state isn't covered), reading rows contiguously from the top and stopping at
    # the first row with a blank TortId (some sheets have pivot-table/summary junk rows
    # below the real data).

    _tpa_sheet_cache = {}
    _TPA_BLANK_TEXTS = {"", "NULL", "N/A", "NA", "-", "NONE", "#N/A"}

    def _strip_state_suffix_from_name(name, state_mco):
        """Some TPA sheets bake the state suffix redundantly INTO LienholderName (e.g.
        "United Healthcare (WA MCO)") even though State MCO is its own column. This
        breaks matching: it defeats the prefix-match (real incoming text never
        includes "(WA MCO)"), and dilutes word-overlap by adding a generic "mco"
        token unrelated plans can share -- e.g. "United Healthcare (WA MCO)" and
        "Columbia United Providers (WA MCO)" tie on just "united", forcing a real
        match into name_not_found. Strips only an exact match of THIS row's own
        State MCO value, never a blanket parenthetical strip -- a different
        parenthetical like "(MCHA)" is left alone."""
        if pd.isna(name) or pd.isna(state_mco):
            return name
        words = str(state_mco).strip().split()
        if not words:
            return name
        suffix_pattern = r"\(\s*" + r"\s+".join(re.escape(w) for w in words) + r"\s*\)"
        cleaned = re.sub(suffix_pattern, "", str(name), flags=re.IGNORECASE)
        return re.sub(r"\s+", " ", cleaned).strip()

    def _match_state_sheet_name(names, state):
        state_upper = str(state).strip().upper()
        for n in names:
            if n.upper().startswith(state_upper + " "):
                return n
        return None

    def load_tpa_state_sheet(tpa_path, state):
        if state in _tpa_sheet_cache:
            return _tpa_sheet_cache[state]
        # One workbook open serves both the sheet-name lookup and the data read below
        # (used to be two separate full-workbook opens -- real, avoidable overhead on
        # a large TPA file, since this already only runs once per run/state either way).
        _ensure_file_hydrated(tpa_path)
        wb = _retry_on_permission_error(
            lambda: _opxl.load_workbook(tpa_path, data_only=True, read_only=True), "the TPA file"
        )
        sheet_name = _match_state_sheet_name(wb.sheetnames, state)
        if sheet_name is None:
            wb.close()
            _tpa_sheet_cache[state] = None
            return None
        ws = wb[sheet_name]
        rows_iter = ws.iter_rows(values_only=True)
        header = next(rows_iter, None)
        if header is None:
            # An empty state tab: the state is treated as not covered (review), never a crash.
            wb.close()
            print(f"⚠️  The TPA file's {sheet_name!r} tab is empty -- {state} is treated as not in the TPA file.")
            _tpa_sheet_cache[state] = None
            return None
        header = _clean_headers(header)
        data_rows = []
        for row in rows_iter:
            if row[0] is None or str(row[0]).strip() == "":
                break
            data_rows.append(row)
        wb.close()
        tpa_df = pd.DataFrame(data_rows, columns=header)
        _missing_tpa_cols = [c for c in ("TortId", "TortName", "Private Lien Case Scope", "CollectorId",
                                         "LienholderId", "LienholderName", "LienType") if c not in tpa_df.columns]
        if _missing_tpa_cols:
            # A renamed/missing header: the state is treated as not covered (review), with the reason.
            print(f"⚠️  The TPA file's {sheet_name!r} tab has no column {', '.join(repr(c) for c in _missing_tpa_cols)} "
                  f"-- {state} is treated as not in the TPA file. Fix the header in the TPA file.")
            _tpa_sheet_cache[state] = None
            return None
        # Text like "NULL" / "N/A" / "-" in an ID or LienType cell means blank (real MI rows),
        # so the plan takes the normal blank-field path instead of using that text as an ID.
        for _blank_col in ("CollectorId", "LienholderId", "LienType"):
            if _blank_col in tpa_df.columns:
                tpa_df[_blank_col] = tpa_df[_blank_col].map(
                    lambda v: None if v is not None and str(v).strip().upper() in _TPA_BLANK_TEXTS else v)
        for _id_col in ("TortId", "CollectorId", "LienholderId"):
            if _id_col in tpa_df.columns:
                tpa_df[_id_col] = tpa_df[_id_col].apply(_stringify_id)
        # A CollectorId/LienholderId that is not a number ("TBD") is blank too: the plan takes
        # the normal blank-field question instead of sending that text to SQL.
        for _id_col in ("CollectorId", "LienholderId"):
            tpa_df[_id_col] = tpa_df[_id_col].map(
                lambda v: (lambda k: k if k and re.fullmatch(r"\d+", k) else None)(_id_key(v)))
        if "LienholderName" in tpa_df.columns and "State MCO" in tpa_df.columns:
            tpa_df["LienholderName"] = tpa_df.apply(
                lambda r: _strip_state_suffix_from_name(r["LienholderName"], r["State MCO"]), axis=1
            )
        _tpa_sheet_cache[state] = tpa_df
        return tpa_df

    TPA_OVERRIDES_KEY_COLS = ["TortName", "Private Lien Case Scope", "CollectorId", "LienholderId",
                              "LienholderName", "LienType"]

    def _tpa_overrides_find(tpa_overrides, state_mco, tort_name, lienholder_name, scope):
        """Index of a COMPLETE TPA_Overrides row (TortId filled) for exactly this State MCO +
        TortName + LienholderName + Scope (normalized), or None. Rows without a Tort (old
        legacy list) are used by plan name in _resolve_by_overrides, never here."""
        norm_state = _normalize_key(state_mco)
        norm_tort  = _normalize_key(tort_name)
        norm_name  = _normalize_key(lienholder_name)
        norm_scope = _normalize_key(scope)
        complete_idx = None
        for i, r in enumerate(tpa_overrides):
            if _normalize_key(r.get("State MCO")) != norm_state or _normalize_key(r.get("LienholderName")) != norm_name:
                continue
            if (r.get("TortId") is not None and _normalize_key(r.get("TortName")) == norm_tort
                    and _normalize_key(r.get("Private Lien Case Scope")) == norm_scope
                    and not (r.get("_bad_fields") or {})):
                complete_idx = i
        return complete_idx

    def _override_rows_full_key(tpa_overrides, state_mco, tort_name, lienholder_name, scope):
        """Every complete TPA_Overrides row for exactly this State + Tort + plan name + Scope."""
        keys = [_normalize_key(v) for v in (state_mco, lienholder_name, tort_name, scope)]
        return [e for e in tpa_overrides
                if e.get("TortId") is not None
                and [_normalize_key(e.get(c)) for c in ("State MCO", "LienholderName", "TortName",
                                                         "Private Lien Case Scope")] == keys]

    def _override_rows_by_name(tpa_overrides, state_mco, lienholder_name):
        """Every TPA_Overrides row for this State + plan name: any Tort/Scope, legacy rows too."""
        ns, nn = _normalize_key(state_mco), _normalize_key(lienholder_name)
        if ns is None or nn is None:
            return []
        return [e for e in tpa_overrides
                if _normalize_key(e.get("State MCO")) == ns and _normalize_key(e.get("LienholderName")) == nn]

    def _override_layer_applies(tpa_overrides, state_mco, tort_name, lienholder_name, scope):
        """True when TPA_Overrides answers this plan before the TPA file (see _resolve_by_overrides)."""
        if _override_rows_full_key(tpa_overrides, state_mco, tort_name, lienholder_name, scope):
            return True
        return any(_id_key(e.get("LienholderId")) or "LienholderId" in (e.get("_bad_fields") or {})
                   for e in _override_rows_by_name(tpa_overrides, state_mco, lienholder_name))

    def _co_lt_for_lienholder(state_mco, lienholder_id, tort_name, scope, plan_text, tpa_overrides):
        """CollectorId/LienType for a LienholderId taken from TPA_Overrides by plan name.
        ("auto", (CollectorId, LienType)) when the TPA file has a row with this LienholderId
        for this exact Tort + Scope. Otherwise ("ask", suggestions): each suggestion
        (text, CollectorId, LienType) is only offered as Y/N, never used silently --
          * the same LienholderId with this Scope for other Torts (TPA file and
            TPA_Overrides), when all of those rows agree;
          * this Tort + Scope's own TPA row for the plan (it may carry another LienholderId)."""
        st = str(state_mco or "").split(" ")[0]
        tpa_df = load_tpa_state_sheet(TPA_FILE, st) if st else None
        nt, nsc, lh = _normalize_key(tort_name), _normalize_key(scope), _id_key(lienholder_id)
        suggestions = []
        same_scope_vals, other_torts = set(), []
        if tpa_df is not None and "LienholderId" in tpa_df.columns:
            same_lh = tpa_df[tpa_df["LienholderId"].map(_id_key) == lh]
            same_scope = same_lh[same_lh["Private Lien Case Scope"].map(_normalize_key) == nsc]
            exact = same_scope[same_scope["TortName"].map(_normalize_key) == nt]
            vals = {(_id_key(r["CollectorId"]), _canonical_lien_type(r["LienType"])) for _, r in exact.iterrows()}
            vals = {v for v in vals if None not in v}
            if len(vals) == 1:
                return "auto", next(iter(vals))
            for _, r in same_scope.iterrows():
                v = (_id_key(r["CollectorId"]), _canonical_lien_type(r["LienType"]))
                if None not in v:
                    same_scope_vals.add(v)
                    other_torts.append(str(r["TortName"]))
        for e in tpa_overrides:
            if (e.get("TortId") is not None and _id_key(e.get("LienholderId")) == lh
                    and _normalize_key(e.get("State MCO")) == _normalize_key(state_mco)
                    and _normalize_key(e.get("Private Lien Case Scope")) == nsc
                    and not (e.get("_bad_fields") or {})):
                v = (_id_key(e.get("CollectorId")), _canonical_lien_type(e.get("LienType")))
                if None not in v:
                    same_scope_vals.add(v)
                    other_torts.append(str(e.get("TortName")))
        if len(same_scope_vals) == 1:
            co, lt = next(iter(same_scope_vals))
            torts = ", ".join(repr(t) for t in dict.fromkeys(other_torts))
            suggestions.append((f"CollectorId {co} / {lt} -- used by LienholderId {lh} with Scope {scope!r} "
                                f"for Tort {torts}", co, lt))
        if tpa_df is not None and {"TortName", "LienholderName", "LienholderId"} <= set(tpa_df.columns):
            own = resolve_tpa_lookup(st, tort_name, plan_text, scope)
            row = own.get("row") if own.get("status") in ("matched", "do_not_use") else None
            if row is not None:
                v = (_id_key(row.get("CollectorId")), _canonical_lien_type(row.get("LienType")))
                if None not in v and all((s[1], s[2]) != v for s in suggestions):
                    suggestions.append((f"CollectorId {v[0]} / {v[1]} -- from the TPA file's row "
                                        f"{row.get('LienholderName')!r} (LienholderId {_id_key(row.get('LienholderId'))}) "
                                        f"for this Tort and Scope", v[0], v[1]))
        return "ask", suggestions

    def _find_scoped_override_match(tpa_overrides, state_mco, tort_name, scope, matched_row_name):
        """Narrower, safer companion to _find_conflicting_override: narrows to COMPLETE
        entries matching this EXACT (State MCO, Tort, Scope) first, THEN fuzzy-matches
        the TPA-matched row's own name against just that small bucket via the same
        hardened _match_name. Closes a gap the exact-only check can't: an override
        stored as "Passport" would miss a raw TPA row named "Passport Group".
        Scoping to one bucket first is what makes fuzzy matching safe here -- it
        normally holds very few distinct plans, so the coincidental-overlap risk
        that keeps the PRIMARY lookup exact-only is much smaller (not zero, but
        _match_name's own tie-detection returns None rather than guessing).
        Returns the matching entry, or None."""
        norm_state = _normalize_key(state_mco)
        norm_tort  = _normalize_key(tort_name)
        norm_scope = _normalize_key(scope)
        bucket = [
            e for e in tpa_overrides
            if e.get("TortId") is not None and not (e.get("_bad_fields") or {})
            and _normalize_key(e.get("State MCO")) == norm_state
            and _normalize_key(e.get("TortName")) == norm_tort
            and _normalize_key(e.get("Private Lien Case Scope")) == norm_scope
        ]
        if not bucket:
            return None
        bucket_df = pd.DataFrame(bucket)
        bucket_df["_idx"] = range(len(bucket_df))
        matched_pos = _match_name(matched_row_name, bucket_df, "LienholderName", "_idx")
        return bucket[matched_pos] if matched_pos is not None else None

    def _find_conflicting_override(tpa_overrides, state_mco, lienholder_name, wrong_lienholder_id):
        """Broader FALLBACK safety net, used only when _find_scoped_override_match finds
        nothing for this exact (State,Tort,Scope) bucket. The raw TPA file can have a
        genuinely stale/wrong row (e.g. an old LienholderId) that TPA_Overrides has
        already corrected under that SAME exact LienholderName text, just for a
        DIFFERENT Tort/Scope. Checks whether ANY TPA_Overrides entry (any Tort/Scope)
        carries this EXACT LienholderName (the TPA-matched row's own name, not the raw
        query text) with a DIFFERENT LienholderId -- if so, this name's TPA data is
        known-untrustworthy somewhere, even without a confirmed answer for this
        specific combination, so the caller must not silently trust either source.
        Returns that conflicting entry, or None if there's no conflict."""
        norm_state = _normalize_key(state_mco)
        norm_name  = _normalize_key(lienholder_name)
        for entry in tpa_overrides:
            if (_normalize_key(entry.get("State MCO")) == norm_state
                    and _normalize_key(entry.get("LienholderName")) == norm_name
                    and str(entry.get("LienholderId")) != str(wrong_lienholder_id)):
                return entry
        return None

    def _join_tort_names(tort_names):
        """'A' -> "'A'"; ['A','B'] -> "'A' and 'B'"; ['A','B','C'] -> "'A', 'B', and 'C'"."""
        quoted = [repr(t) for t in tort_names]
        if len(quoted) == 1:
            return quoted[0]
        if len(quoted) == 2:
            return f"{quoted[0]} and {quoted[1]}"
        return ", ".join(quoted[:-1]) + f", and {quoted[-1]}"

    def _find_same_signature_combos(combos_df, exclude_keys, tpa_overrides, text_key, signature):
        """Looks ahead at this run's OTHER combos (excluding exclude_keys) sharing this
        EXACT raw plan text, and returns full details -- (Tort, State, MCO_Name, Scope,
        candidates, distinct_names) -- for every one that will ALSO independently reach
        name_not_found with this SAME candidate-list signature. Used to build ONE group
        that gets announced together AND processed back-to-back, immediately, right
        after being announced -- so a Tort named in a heads-up is never left hanging
        while some unrelated question runs in between. Reuses the exact same real
        _tpa_overrides_find / resolve_tpa_lookup the main loop itself calls (no
        separate/duplicated decision logic that could drift out of sync), so this can
        never disagree with what actually happens for each combo."""
        group = []
        for _, c2 in combos_df.iterrows():
            state2, tort2, mco_name2, scope2 = (
                c2["State"], c2["Project_Name"], c2["MCO_Name"], c2["Scope_Classification"]
            )
            key2 = (state2, tort2, mco_name2, scope2)
            if key2 in exclude_keys:
                continue
            if _normalize_key(mco_name2) != text_key:
                continue
            state_mco2 = f"{str(state2).strip().upper()} MCO"
            if _override_layer_applies(tpa_overrides, state_mco2, tort2, mco_name2, scope2):
                continue
            result2 = resolve_tpa_lookup(state2, tort2, mco_name2, scope2)
            if result2["status"] != "name_not_found":
                continue
            # Same placeholder-row filter as the main loop's distinct_names -- a blank
            # LienholderName is never a meaningful option to pick by name.
            distinct_names2 = [
                n for n in result2["candidates"]["LienholderName"].drop_duplicates().tolist()
                if _normalize_key(n) is not None
            ]
            signature2 = (text_key, tuple(sorted(_normalize_key(n) for n in distinct_names2)))
            if signature2 == signature:
                group.append((tort2, state2, mco_name2, scope2, result2["candidates"], distinct_names2))
        return group

    def _plan_candidates_for_prompt(plan_text, tort_df):
        """Narrows the human-facing 'name not found' candidate list down to TPA rows
        that plausibly relate to the raw text, instead of dumping every plan name in
        this State+Tort sheet (the previous behavior -- a state/tort with a large TPA
        sheet buried the handful of relevant options in dozens of unrelated ones, e.g.
        searching 'United' in a state with ~19 CCO plans showed all ~19). Any shared
        word with the query (ignoring _STRUCTURAL_TAGS, which carry no signal within
        one state's own sheet) is enough to surface a row here -- deliberately looser
        than _match_name's own gate/corroboration rules, since this is only a
        shortlist for a HUMAN to pick from, not an auto-match decision. Falls back to
        the full sheet if nothing shares a word with the raw text at all, so a
        genuinely novel plan name still lets the user browse everything rather than
        see an empty list."""
        query_words = set(_tokenize_plan(str(plan_text).strip().lower())) - _STRUCTURAL_TAGS
        if not query_words:
            return tort_df
        mask = tort_df["LienholderName"].apply(
            lambda n: bool((set(_tokenize_plan(str(n).strip().lower())) - _STRUCTURAL_TAGS) & query_words)
        )
        narrowed = tort_df[mask]
        return narrowed if not narrowed.empty else tort_df

    def _name_not_found_pending_reason(g_state, g_tort, g_candidates):
        """Distinguishes, for the PendingReason message only, a real narrowed
        shortlist (word overlap, or a genuine _match_name tie) from the full-sheet
        fallback (nothing shared a word at all) -- without storing the actual
        candidate list anywhere, since even the narrowed list can run to dozens of
        names (see _plan_candidates_for_prompt) and would make Pending_TPA_Review
        unreadable if dumped into every row."""
        _full_tort_df = load_tpa_state_sheet(TPA_FILE, str(g_state).strip().upper())
        if _full_tort_df is not None:
            _full_tort_df = _full_tort_df[_full_tort_df["TortName"].apply(_normalize_key) == _normalize_key(g_tort)]
        _n = len(g_candidates)
        if _full_tort_df is not None and _n == len(_full_tort_df) and _n > 0:
            return "Plan name not found in TPA -- no candidates shared any words with this text; user skipped"
        return f"Plan name not found in TPA -- {_n} candidate(s) shared words with this text; user skipped"

    def resolve_tpa_lookup(state, tort, mco_name, scope):
        """Tort -> State-sheet -> Name -> Scope cascade against the TPA file (no override
        check -- the caller checks TPA_Overrides first). Returns a dict with a "status"
        key describing the outcome. On "name_not_found", "candidates" is narrowed to
        plausibly-related rows, not the entire State+Tort sheet -- see
        _plan_candidates_for_prompt."""
        state_code = str(state).strip().upper()
        tpa_df = load_tpa_state_sheet(TPA_FILE, state_code)
        if tpa_df is None:
            return {"status": "state_not_covered"}
        norm_tort = _normalize_key(tort)
        tort_df = tpa_df[tpa_df["TortName"].apply(_normalize_key) == norm_tort]
        if tort_df.empty:
            return {"status": "tort_not_covered"}
        # Matched on a per-row key, so a row with the plan's name but a blank LienholderId
        # still counts as the match (only its LienholderId is then asked for).
        tort_df = tort_df.assign(_MatchId=[
            str(_id_key(v)) if _id_key(v) is not None else f"__ROW{i}"
            for i, v in zip(tort_df.index, tort_df["LienholderId"])
        ])
        matched_id, tied_ids = _match_name(mco_name, tort_df, "LienholderName", "_MatchId", return_tie=True)
        if matched_id is None:
            if tied_ids:
                # A genuine _match_name tie (2+ real candidates equally explain the
                # query) IS the relevant shortlist already -- no need to fall back to
                # word-overlap filtering, and no risk of the fallback accidentally
                # excluding one of the tied names.
                _tied_id_strs = {str(t) for t in tied_ids}
                candidates = tort_df[tort_df["_MatchId"].isin(_tied_id_strs)]
            else:
                candidates = _plan_candidates_for_prompt(mco_name, tort_df)
            return {"status": "name_not_found", "candidates": candidates.drop(columns=["_MatchId"], errors="ignore")}
        # Group by the matched NAME, not the specific id _match_name happened to return --
        # two rows can share the exact same LienholderName with DIFFERENT LienholderIds
        # (the real "United Healthcare" duplicate case this design targets).
        matched_row = tort_df[tort_df["_MatchId"] == str(matched_id)].iloc[0]
        tort_df = tort_df.drop(columns=["_MatchId"])
        # A TPA row marked "DO NOT USE" is never matched silently: the user is asked once
        # per plan text (Y = use it, N = pick from the list, S = review).
        if _is_do_not_use(matched_row["LienholderName"]):
            _dnu = _DNU_DECISIONS.get((state_code, _normalize_key(mco_name), _normalize_key(matched_row["LienholderName"])))
            if _dnu == "N":
                return {"status": "name_not_found", "candidates": _plan_candidates_for_prompt(mco_name, tort_df)}
            if _dnu == "S":
                return {"status": "do_not_use_skipped", "row": matched_row}
            if _dnu != "Y":
                return {"status": "do_not_use", "row": matched_row}
        matched_name_key = _normalize_key(matched_row["LienholderName"])
        name_matches = tort_df[tort_df["LienholderName"].apply(_normalize_key) == matched_name_key]
        norm_scope = _normalize_key(scope)
        scope_matches = name_matches[name_matches["Private Lien Case Scope"].apply(_normalize_key) == norm_scope]
        if len(scope_matches) == 1:
            return {"status": "matched", "row": scope_matches.iloc[0]}
        elif len(scope_matches) == 0:
            return {"status": "scope_gap", "candidates": name_matches}
        else:
            return {"status": "ambiguous_duplicate", "candidates": scope_matches}

    def _tpa_torts_for_state(state_code):
        """[(TortId, TortName)] of the state's TPA sheet, in sheet order, no repeats."""
        tpa_df = load_tpa_state_sheet(TPA_FILE, state_code) if state_code else None
        if tpa_df is None or "TortId" not in tpa_df.columns or "TortName" not in tpa_df.columns:
            return []
        out = []
        for tid, tname in zip(tpa_df["TortId"], tpa_df["TortName"]):
            pair = (_id_key(tid), str(tname).strip() if pd.notna(tname) and str(tname).strip() else None)
            if pair[0] and pair[1] and pair not in out:
                out.append(pair)
        return out

    def _resolve_blank_torts(working_df, mask, pending_idx):
        """Rows whose Tort or TortId is blank in LPM (Tort is per case). In order: the Tort
        filled in an earlier Review file for this case; the missing half from the TPA file
        (TortId <-> Tort name); else asked once per case: [n] one of the state's TPA Torts,
        [N] typed, [S] -> Pending_TPA_Review "Tort is blank in LPM". Never saved to config."""
        groups = {}
        for idx in working_df.index[mask]:
            case = _id_key(working_df.at[idx, "CaseId"]) if "CaseId" in working_df.columns else None
            client = _id_key(working_df.at[idx, "ClientId"]) if "ClientId" in working_df.columns else str(idx)
            groups.setdefault(case or ("CLIENT", client), []).append(idx)
        for gkey, idxs in groups.items():
            state_code = str(working_df.at[idxs[0], "State"] or "").strip().upper()
            torts = _tpa_torts_for_state(state_code)
            by_id, by_name = {}, {}
            for _tid, _tname in torts:
                by_id.setdefault(_tid, _tname)
                by_name.setdefault(_normalize_key(_tname), _tid)
            name = next((working_df.at[i, "Project_Name"] for i in idxs
                         if _normalize_key(working_df.at[i, "Project_Name"])), None)
            tid = next((_id_key(working_df.at[i, "TortId"]) for i in idxs if _id_key(working_df.at[i, "TortId"])), None)
            carried = _BLANK_TORT_CARRY.get(gkey) if isinstance(gkey, str) else None
            if carried is None and "LienId" in working_df.columns:
                carried = next((_BLANK_TORT_CARRY[("LIEN", _clean_lien_id(working_df.at[i, "LienId"]))] for i in idxs
                                if ("LIEN", _clean_lien_id(working_df.at[i, "LienId"])) in _BLANK_TORT_CARRY), None)
            obs = None
            if carried is not None:
                tid, name = carried
                obs = f"Tort was blank in LPM - used '{name}' (TortId {tid}) from the Review file"
            elif name and not tid and by_name.get(_normalize_key(name)):
                tid = by_name[_normalize_key(name)]
                obs = f"TortId was blank in LPM - filled from the TPA file ({tid})"
            elif tid and not name and by_id.get(tid):
                name = by_id[tid]
                obs = f"Tort name was blank in LPM - filled from the TPA file ('{name}')"
            else:
                _plan_banner_once()
                _new_item("plan")
                clients = (list(dict.fromkeys(str(c) for c in working_df.loc[idxs, "ClientId"]))
                           if "ClientId" in working_df.columns else [])
                plans = list(dict.fromkeys(str(p) for p in working_df.loc[idxs, "MCO_Name"]))
                label = f"CaseId {gkey}" if isinstance(gkey, str) else f"ClientId {gkey[1]} (no CaseId)"
                what = ("Tort is blank in LPM" if not name and not tid
                        else f"Tort {name!r} has no TortId in LPM" if name else f"TortId {tid} has no Tort name in LPM")
                print(f"\n{what} for {label} ({len(clients)} client(s): {', '.join(clients)}; "
                      f"plan(s): {', '.join(repr(p) for p in plans)}).")
                if torts:
                    print(f"Torts in the TPA file for {state_code}:")
                    for _i, (_t_id, _t_name) in enumerate(torts, start=1):
                        print(f"  [{_i}] {_t_name} (TortId {_t_id})")
                print("  [N] Not in this list - type the TortId and Tort name")
                print("  [S] Skip - send to Pending_TPA_Review (fill TortId and TortName there)")
                while True:
                    _a = input("Enter a number, N or S: " if torts else "Enter N or S: ").strip().upper()
                    if _a in ("N", "S") or (_a.isdigit() and 1 <= int(_a) <= len(torts)):
                        break
                    print(f"Please enter a number between 1 and {len(torts)}, N or S." if torts
                          else "Please enter N or S.")
                if _a == "S":
                    working_df.loc[idxs, "PendingReason"] = BLANK_TORT_REASON
                    pending_idx.extend(idxs)
                    continue
                if _a == "N":
                    tid = _ask_typed_id("TortId: ", allow_skip=False)
                    while True:
                        name = input("Tort name: ").strip()
                        if name:
                            break
                        print("Please enter the Tort name.")
                    obs = f"Tort was blank in LPM - user typed '{name}' (TortId {tid})"
                else:
                    tid, name = torts[int(_a) - 1]
                    obs = f"Tort was blank in LPM - user picked '{name}' (TortId {tid})"
            working_df.loc[idxs, "Project_Name"] = name
            working_df.loc[idxs, "TortId"] = tid
            working_df.loc[idxs, "Observation"] = [_join_observations([o, obs])
                                                   for o in working_df.loc[idxs, "Observation"]]

    def resolve_tpa_for_run(df_in, tpa_overrides, ticket_id, alias_saves=None):
        """Resolves LienholderId/CollectorId/LienType for every unique
        (State, Tort, MCO_Name, Scope_Classification) combo in df_in. Returns
        (resolved_df, pending_tpa_df, pending_blank_fields_df, tpa_dirty). A State or
        Tort missing from the TPA file sends its plans to Pending_TPA_Review.
        tpa_dirty / alias_saves receive the saves chosen at the end-of-run question
        (_ask_deferred_saves); alias_saves is optional for callers that don't need it."""
        if alias_saves is None:
            alias_saves = []
        working_df = df_in.copy()
        for col in ["LienholderId", "CollectorId", "LienType"]:
            working_df[col] = None
        # Review-only columns (dropped from resolved rows): the file's own plan text, the LPM
        # name of a pre-filled LienholderId, and the text to look up in LPM for a TPA-state row.
        _lien_ids = working_df["LienId"] if "LienId" in working_df.columns else pd.Series(None, index=working_df.index)
        working_df["Original_MCO_Name"] = [
            _ORIGINAL_MCO_NAME_CARRY.get((str(_l).strip(), _normalize_key(_m)), _m)
            for _l, _m in zip(_lien_ids, working_df["MCO_Name"])
        ]
        working_df["LPM_LienholderName"] = None
        working_df["_LPM_Match_Text"] = None
        # Where this plan's LienholderId/CollectorId/LienType came from -- named in
        # Pending_Error_Review's fix text if one of them breaks Step 4.
        working_df["_TPA_Source"] = None
        if "Observation" not in working_df.columns:
            working_df["Observation"] = None
        _review_only_cols = ["Original_MCO_Name", "LPM_LienholderName", "_LPM_Match_Text"]

        tpa_dirty = []
        pending_idx = []
        blank_field_idx = []
        error_idx = []
        _uncovered_state_idx = {}

        def _is_blank(v):
            if pd.isna(v):
                return True
            # Defensive (not a proven-reachable path today, CollectorId/LienholderId
            # are protected via _stringify_id): a value that already round-tripped
            # through a str()/dtype=str conversion on a real NaN/None can arrive here
            # as the literal text "nan"/"None" instead of a genuine blank -- treat
            # those the same as a real blank, since they mean the same thing.
            _v_str = str(v).strip()
            return _v_str == "" or _v_str.lower() in ("nan", "none")

        def _finalize_tpa_match(row_mask, lienholder_id, collector_id, lien_type,
                                 tort_id=None, tort_name=None, scope_val=None,
                                 lienholder_name=None, state_mco_val=None,
                                 is_first_for_combo=True, source=None):
            """Applies a resolved LienholderId/CollectorId/LienType onto working_df --
            or, if CollectorId/LienType are blank/NaN, offers to fill them in live
            (only those 2 fields -- LienholderId is already known, unlike D-flow)
            and patch the TPA file in place, skipping Pending_TPA_Blank_Fields
            entirely. Only routes there if the user skips. Shared by every path that
            can produce a final match, since a blank CollectorId/LienType is a real,
            recurring data-quality issue -- without this check it would sail through
            as "resolved" and surface later as blank fields downstream, or a crash
            in get_assigned_user_id.
            Returns True if resolved, False if routed to blank_field_idx."""
            if source:
                working_df.loc[row_mask, "_TPA_Source"] = source
            if _is_blank(lienholder_id):
                # A TPA row with the plan's name but a blank LienholderId: still the match,
                # only the missing LienholderId is asked (same as a blank CollectorId below).
                _plan_banner_once()
                if is_first_for_combo:
                    _new_item("plan")
                else:
                    _gap()
                print(f"TPA row matched for {lienholder_name!r} (Tort {tort_name!r} / Scope {scope_val!r}) "
                      f"but its LienholderId is blank in the TPA file.")
                while True:
                    _given_lhid = _ask_typed_id("LienholderId (or 'S' to skip): ")
                    if _given_lhid == "S":
                        break
                    _tpa_state_df = load_tpa_state_sheet(TPA_FILE, str(state_mco_val or "").split(" ")[0])
                    _other_names = [] if _tpa_state_df is None else sorted({
                        str(n) for n in _tpa_state_df.loc[
                            _tpa_state_df["LienholderId"].map(_id_key) == _given_lhid, "LienholderName"].dropna()
                        if _normalize_key(n) != _normalize_key(lienholder_name)})
                    if not _other_names or ask_yes_no(
                            f"LienholderId {_given_lhid} is already used in the TPA file for "
                            f"{', '.join(_other_names)}. Use it for {lienholder_name!r} anyway? (Y/N): "):
                        break
                if _given_lhid == "S":
                    for _col, _val in (("LienholderId", None), ("CollectorId", collector_id), ("LienType", lien_type),
                                       ("TPA_Match_TortId", tort_id), ("TPA_Match_TortName", tort_name),
                                       ("TPA_Match_Scope", scope_val), ("TPA_Match_LienholderName", lienholder_name),
                                       ("TPA_Match_StateMCO", state_mco_val)):
                        working_df.loc[row_mask, _col] = _val
                    working_df.loc[row_mask, "PendingReason"] = (
                        "TPA entry matched but missing required details (LienholderId)"
                    )
                    blank_field_idx.extend(working_df[row_mask].index.tolist())
                    return False
                if _patch_tpa_file_blank_fields(state_mco_val, tort_id, None, None, None,
                                                lienholder_name=lienholder_name, scope=scope_val,
                                                new_lienholder_id=_given_lhid):
                    print(f"✅ Patched the TPA file in place (LienholderId {_given_lhid} for {lienholder_name!r}).")
                else:
                    print("⚠️  Could not patch the TPA file -- this run still uses the LienholderId you just "
                          "gave, but you'll need to update the TPA file manually so future runs don't ask again.")
                lienholder_id = _given_lhid
            if _is_blank(collector_id) or _is_blank(lien_type):
                _plan_banner_once()
                if is_first_for_combo:
                    _new_item("plan")
                else:
                    _gap()
                print(
                    f"LienholderId {lienholder_id!r} matched for {lienholder_name!r} "
                    f"(Tort {tort_name!r} / Scope {scope_val!r}) but CollectorId/LienType are "
                    f"blank in the TPA file."
                )
                _given_collector_id = _ask_typed_id("CollectorId (or 'S' to skip): ")
                if _given_collector_id != "S":
                    collector_id = _given_collector_id
                    lien_type = ask_numbered_choice(LIEN_TYPE_OPTIONS, "LienType:", "Enter the number: ")
                    if _patch_tpa_file_blank_fields(state_mco_val, tort_id, lienholder_id, collector_id, lien_type):
                        print(f"✅ Patched the TPA file in place (LienholderId {lienholder_id!r}).")
                    else:
                        print(
                            f"⚠️  Could not patch the TPA file -- this run still uses the "
                            f"values you just gave, but you'll need to update the TPA file "
                            f"manually so future runs don't ask again."
                        )
                    # Falls through to the normal success path below with the
                    # now-filled-in collector_id/lien_type.
                else:
                    working_df.loc[row_mask, "LienholderId"] = lienholder_id
                    working_df.loc[row_mask, "CollectorId"] = collector_id
                    working_df.loc[row_mask, "LienType"] = lien_type
                    working_df.loc[row_mask, "TPA_Match_TortId"] = tort_id
                    working_df.loc[row_mask, "TPA_Match_TortName"] = tort_name
                    working_df.loc[row_mask, "TPA_Match_Scope"] = scope_val
                    working_df.loc[row_mask, "TPA_Match_LienholderName"] = lienholder_name
                    working_df.loc[row_mask, "TPA_Match_StateMCO"] = state_mco_val
                    working_df.loc[row_mask, "PendingReason"] = (
                        "TPA entry matched but missing required details (CollectorId/LienType)"
                    )
                    blank_field_idx.extend(working_df[row_mask].index.tolist())
                    return False
            working_df.loc[row_mask, "LienholderId"] = lienholder_id
            working_df.loc[row_mask, "CollectorId"] = collector_id
            working_df.loc[row_mask, "LienType"] = lien_type
            # Once a plan resolves to a known name (fresh TPA match, list-pick, or a
            # saved override's name), MCO_Name is rewritten to it -- same mechanism as
            # Alias_Map's normalize_mco_name, applied consistently across every
            # resolution path. Yes_with_MCO's MCO_Name, Remarks, and every LienNote's
            # "Plan Name:" all read this one field, so this rewrite keeps them in sync.
            # PLAN_FATE is keyed by (LienId, MCO_Name), seeded with the pre-resolution
            # value -- the old key must be dropped here too, or Remarks shows a fake
            # "tracking gap" for the old text alongside the real outcome for the new one.
            if not _is_blank(lienholder_name):
                # Must compare the RAW value, not _normalize_key(...) -- PLAN_FATE's keys
                # are literal, un-normalized text, so "ANTHEM"/"Anthem" are two DIFFERENT
                # keys. Real bug: comparing normalized values wrongly concluded "nothing
                # changed" for a pure case difference and skipped the pop, while the
                # assignment below still changed the literal text -- left the OLD-case
                # key stuck at its placeholder while Step 4 recorded the real outcome
                # under the NEW-case key, so Remarks showed both.
                for _, _fmrow in working_df.loc[row_mask, ["LienId", "MCO_Name"]].drop_duplicates().iterrows():
                    if _fmrow["MCO_Name"] != lienholder_name:
                        PLAN_FATE.pop((_fmrow["LienId"], _fmrow["MCO_Name"]), None)
                working_df.loc[row_mask, "MCO_Name"] = lienholder_name
            return True

        # ---------------- TPA_Overrides before the TPA file ----------------
        # A. A complete row for this exact State + Tort + Scope + plan name -> used as it is
        #    (2+ such rows with different values -> asked once per run).
        # B. Else every row for this State + plan name (any Tort/Scope, old legacy rows
        #    too): one LienholderId -> used, never the TPA file's; 2+ -> asked once per plan
        #    per run. CollectorId/LienType: the TPA row with that LienholderId for this Tort
        #    + Scope, else asked (with Y/N suggestions, never silent).
        # C. Plan name not in TPA_Overrides -> the TPA file (caller).
        _dup_key_choice = {}
        _name_lhid_choice = {}
        _bad_value_answers = {}
        _b_skipped = set()

        def _plan_label(mco_name, tort, scope):
            return f"{mco_name!r} (Tort {tort!r} / Scope {scope!r})"

        def _usable_value(entry, field):
            """The entry's value for field, asking for a config value that can't be used
            (this run only, see _override_bad_fields). None = blank, "S" = skipped."""
            bad = (entry.get("_bad_fields") or {})
            if field not in bad:
                return entry.get(field)
            key = (entry.get("_row") or id(entry), field)
            if key not in _bad_value_answers:
                _plan_banner_once()
                _new_item("plan")
                what = (f"is neither {' nor '.join(repr(o) for o in LIEN_TYPE_OPTIONS)}" if field == "LienType"
                        else "is not a number")
                print(f"\n{_override_row_label(entry)} for {entry.get('LienholderName')!r}: {field} "
                      f"{bad[field]!r} {what}. Please fix it in mco_config.xlsx.")
                if field == "LienType":
                    print("LienType to use in this run:")
                    for _i, _o in enumerate(LIEN_TYPE_OPTIONS, start=1):
                        print(f"  [{_i}] {_o}")
                    while True:
                        _a = input("Enter the number, or 'S' to skip: ").strip().upper()
                        if _a == "S" or (_a.isdigit() and 1 <= int(_a) <= len(LIEN_TYPE_OPTIONS)):
                            break
                        print(f"Please enter a number between 1 and {len(LIEN_TYPE_OPTIONS)}, or 'S'.")
                    _bad_value_answers[key] = "S" if _a == "S" else LIEN_TYPE_OPTIONS[int(_a) - 1]
                else:
                    _bad_value_answers[key] = _ask_typed_id(f"Correct {field} for this run (or 'S' to skip): ")
            return _bad_value_answers[key]

        def _send_to_review(row_mask, reason, lienholder_id=None):
            if lienholder_id is not None:
                working_df.loc[row_mask, "LienholderId"] = lienholder_id
            working_df.loc[row_mask, "PendingReason"] = reason
            pending_idx.extend(working_df[row_mask].index.tolist())

        def _pick_duplicate_full_key(rows, mco_name, tort, scope, state_mco):
            """2+ complete rows for one key with different values -> the user picks (once per run)."""
            key = tuple(_normalize_key(v) for v in (state_mco, mco_name, tort, scope))
            if key not in _dup_key_choice:
                _plan_banner_once()
                _new_item("plan")
                print(f"\nTPA_Overrides has {len(rows)} rows for {_plan_label(mco_name, tort, scope)} "
                      f"({state_mco}) with different values:")
                for _i, _e in enumerate(rows, start=1):
                    print(f"  [{_i}] {_override_row_label(_e)}: LienholderId {_e.get('LienholderId')}, "
                          f"CollectorId {_e.get('CollectorId')}, {_e.get('LienType')}")
                print("Keep only the right row in mco_config.xlsx so this is not asked again.")
                while True:
                    _a = input("Enter the number of the row to use in this run, or 'S' to skip: ").strip().upper()
                    if _a == "S" or (_a.isdigit() and 1 <= int(_a) <= len(rows)):
                        break
                    print(f"Please enter a number between 1 and {len(rows)}, or 'S'.")
                _dup_key_choice[key] = None if _a == "S" else rows[int(_a) - 1]
            return _dup_key_choice[key]

        def _pick_lienholder_by_name(cands, mco_name, state_mco):
            """2+ LienholderIds for one plan name in TPA_Overrides -> asked once per plan per run.
            cands: {LienholderId or ("BAD", row id): [rows]}. Returns the chosen key or "S"."""
            key = (_normalize_key(state_mco), _normalize_key(mco_name))
            if key not in _name_lhid_choice:
                _plan_banner_once()
                _new_item("plan")
                print(f"\nTPA_Overrides has different LienholderIds for plan {mco_name!r} ({state_mco}):")
                options = list(cands.items())
                for _i, (_lh, _rows) in enumerate(options, start=1):
                    _where = "; ".join(
                        f"{_override_row_label(_e)} ("
                        + (f"Tort {_e.get('TortName')!r} / Scope {_e.get('Private Lien Case Scope')!r}"
                           if _e.get("TortId") is not None else "no Tort - old list") + ")"
                        for _e in _rows)
                    _shown = (f"{_rows[0]['_bad_fields']['LienholderId']!r} (not a number)"
                              if isinstance(_lh, tuple) else _lh)
                    print(f"  [{_i}] LienholderId {_shown} -- {_where}")
                while True:
                    _a = input(f"Which LienholderId should {mco_name!r} use in this run? Enter the number, "
                               f"or 'S' to skip (send to Pending_TPA_Review): ").strip().upper()
                    if _a == "S" or (_a.isdigit() and 1 <= int(_a) <= len(options)):
                        break
                    print(f"Please enter a number between 1 and {len(options)}, or 'S'.")
                _name_lhid_choice[key] = "S" if _a == "S" else options[int(_a) - 1][0]
            return _name_lhid_choice[key]

        def _ask_co_lt(lienholder_id, mco_name, tort, scope, source_label, suggestions):
            """CollectorId/LienType for a LienholderId from TPA_Overrides: each suggestion as
            Y/N, then typed. Returns (CollectorId, LienType, "suggested"/"typed") or None when skipped."""
            _plan_banner_once()
            _new_item("plan")
            print(f"\nLienholderId {lienholder_id} for plan {_plan_label(mco_name, tort, scope)} comes from "
                  f"{source_label}. The TPA file has no row with this LienholderId for this Tort and Scope, "
                  f"so CollectorId and LienType are needed.")
            for _text, _co, _lt in suggestions:
                if ask_yes_no(f"Suggestion: {_text}. Use it? (Y/N): "):
                    return _co, _lt, "suggested"
            _co = _ask_typed_id("CollectorId (or 'S' to skip): ")
            if _co == "S":
                return None
            return _co, ask_numbered_choice(LIEN_TYPE_OPTIONS, "LienType:", "Enter the number: "), "typed"

        def _resolve_by_overrides(row_mask, state_mco, tort, mco_name, scope, row_tort_id):
            """Layers A and B above. True when this combo was handled (resolved, sent to a
            review sheet), False when TPA_Overrides has nothing for this plan name (-> TPA)."""
            full = _override_rows_full_key(tpa_overrides, state_mco, tort, mco_name, scope)
            by_name = _override_rows_by_name(tpa_overrides, state_mco, mco_name)
            entry = None
            if full:
                distinct = {(_id_key(e.get("LienholderId")), _id_key(e.get("CollectorId")),
                             _canonical_lien_type(e.get("LienType")) or e.get("LienType")) for e in full}
                entry = full[-1] if len(distinct) == 1 else _pick_duplicate_full_key(full, mco_name, tort, scope, state_mco)
                if entry is None:
                    _send_to_review(row_mask, "TPA_Overrides has rows with different values for this plan; user skipped")
                    return True
                values = [_usable_value(entry, f) for f in ("LienholderId", "CollectorId", "LienType")]
                if "S" in values:
                    _send_to_review(row_mask, "TPA_Overrides value can't be used (see the warning); user skipped")
                    return True
                if not _is_blank(values[0]) and not _is_blank(values[1]) and not _is_blank(values[2]):
                    _finalize_tpa_match(row_mask, values[0], values[1], values[2],
                                        tort_id=entry.get("TortId"), tort_name=entry.get("TortName"),
                                        scope_val=entry.get("Private Lien Case Scope"),
                                        lienholder_name=entry.get("LienholderName"), state_mco_val=state_mco,
                                        source=_override_row_label(entry))
                    return True
                # A complete row missing CollectorId/LienType: its LienholderId goes through B,
                # and that same row is filled (in memory; the save fills it in the file).
                by_name = [entry]
            incomplete_full = entry
            cands = {}
            for e in by_name:
                if "LienholderId" in (e.get("_bad_fields") or {}):
                    cands.setdefault(("BAD", e.get("_row") or id(e)), []).append(e)
                elif _id_key(e.get("LienholderId")):
                    cands.setdefault(_id_key(e.get("LienholderId")), []).append(e)
            if not cands:
                return False
            chosen = next(iter(cands)) if len(cands) == 1 else _pick_lienholder_by_name(cands, mco_name, state_mco)
            if chosen == "S":
                _send_to_review(row_mask, "TPA_Overrides has different LienholderIds for this plan name; user skipped")
                return True
            src_rows = cands[chosen]
            lhid = chosen
            if isinstance(chosen, tuple):
                lhid = _usable_value(src_rows[0], "LienholderId")
                if lhid == "S":
                    _send_to_review(row_mask, "TPA_Overrides value can't be used (see the warning); user skipped")
                    return True
            source_label = _override_row_label(src_rows[0])
            skip_key = (_normalize_key(state_mco), _normalize_key(mco_name), _normalize_key(tort),
                        _normalize_key(scope), lhid)
            if skip_key in _b_skipped:
                _send_to_review(row_mask, "LienholderId from TPA_Overrides needs CollectorId/LienType; user skipped",
                                lienholder_id=lhid)
                return True
            how, found = _co_lt_for_lienholder(state_mco, lhid, tort, scope, mco_name, tpa_overrides)
            given = None
            if how == "auto":
                co, lt = found
            else:
                answer = _ask_co_lt(lhid, mco_name, tort, scope, source_label, found)
                if answer is None:
                    _b_skipped.add(skip_key)
                    _send_to_review(row_mask, "LienholderId from TPA_Overrides needs CollectorId/LienType; user skipped",
                                    lienholder_id=lhid)
                    return True
                co, lt, given = answer
            _finalize_tpa_match(row_mask, lhid, co, lt, tort_id=row_tort_id, tort_name=tort, scope_val=scope,
                                lienholder_name=mco_name, state_mco_val=state_mco,
                                source=f"{source_label} (LienholderId)", is_first_for_combo=False)
            legacy = any(e.get("TortId") is None for e in src_rows)
            _reasons = {None: (TPA_REASON_LEGACY_TPA, TPA_REASON_NAME_TPA),
                        "typed": (TPA_REASON_LEGACY_USER, TPA_REASON_NAME_USER),
                        "suggested": (TPA_REASON_LEGACY_CONFIRMED, TPA_REASON_NAME_CONFIRMED)}[given]
            new_entry = {
                "TortId": row_tort_id, "TortName": tort, "Private Lien Case Scope": scope,
                "CollectorId": co, "LienholderId": lhid, "LienholderName": mco_name, "LienType": lt,
                "State MCO": state_mco, "Ticket #": ticket_id,
                "Reason": _reasons[0] if legacy else _reasons[1],
            }
            if incomplete_full is not None:
                incomplete_full.update(CollectorId=co, LienType=lt)
            else:
                tpa_overrides.append(new_entry)
            if row_tort_id is None:
                return True
            if how == "auto" and legacy and not (src_rows[0].get("_bad_fields") or {}):
                # Old legacy rows are completed for their first Tort + Scope (the save fills
                # that row in place when its LienholderId is the same), as before.
                tpa_dirty.append(new_entry)
            elif how != "auto":
                _defer_save(_tpa_entry_label(new_entry), tpa_dirty, new_entry)
            return True

        # ---------------- Rows that can't be looked up as they are ----------------
        # Blank State (no run state to use either) -> Pending_Error_Review with the fix.
        _no_state = working_df["State"].map(lambda v: _normalize_key(v) is None)
        if _no_state.any():
            _record_plan_error(working_df[_no_state], "Step 3 (TPA lookup)", ValueError("State is blank"))
            error_idx.extend(working_df.index[_no_state].tolist())

        # Blank Tort (or TortId) in LPM: the missing half comes from the TPA file, else a
        # Tort filled in an earlier Review file for the same case, else asked once per case.
        _tort_incomplete = (~working_df.index.isin(error_idx)) & working_df["Project_Name"].map(
            lambda v: _normalize_key(v) is None)
        if "TortId" in working_df.columns:
            _tort_incomplete = _tort_incomplete | ((~working_df.index.isin(error_idx))
                                                   & working_df["TortId"].map(lambda v: _id_key(v) is None))
        else:
            working_df["TortId"] = None
        if _tort_incomplete.any():
            _resolve_blank_torts(working_df, _tort_incomplete, pending_idx)

        # When the SAME raw text fails to match TPA for one Tort, the same candidate
        # list almost always shows again for another Tort sharing that text this run.
        # Combos already fully handled (individually, or as part of a name_not_found
        # GROUP below) are tracked here so the outer loop skips them later in `combos`.
        _processed_combos = set()

        # Two different raw texts can end up with the identical narrowed candidate list
        # (e.g. only one plan on file, or both fall back to the full sheet per
        # _plan_candidates_for_prompt). Tracks candidate lists by content (not raw
        # text) already printed this run -- a later identical list gets a short
        # pointer back instead of a full reprint, but still asks its own independent
        # question every time; only the redundant printing changes.
        _seen_candidate_lists = {}

        _done_idx = set(pending_idx) | set(error_idx)
        combos = working_df.loc[~working_df.index.isin(list(_done_idx)),
                                ["State", "Project_Name", "MCO_Name", "Scope_Classification"]].drop_duplicates()

        def _combo_mask(st, tt, mn, sc):
            """Rows of one combo, blank-safe (a blank never equals itself)."""
            m = ~working_df.index.isin(list(_done_idx))
            for col, val in (("State", st), ("Project_Name", tt), ("MCO_Name", mn), ("Scope_Classification", sc)):
                m = m & (working_df[col].isna() if _is_blank(val) else (working_df[col] == val))
            return pd.Series(m, index=working_df.index)

        for _, combo in combos.iterrows():
            state, tort, mco_name, scope = (
                combo["State"], combo["Project_Name"], combo["MCO_Name"], combo["Scope_Classification"]
            )
            combo_key = (state, tort, mco_name, scope)
            if combo_key in _processed_combos:
                continue
            row_mask = _combo_mask(state, tort, mco_name, scope)
            if not row_mask.any():
                continue
            # Safety net: an unexpected error in one combo sends only its plans (and any
            # combo grouped with it) to Pending_Error_Review; nothing it started is kept.
            _snap = (len(tpa_overrides), len(tpa_dirty), len(alias_saves), len(_DEFERRED_SAVES),
                     len(pending_idx), len(blank_field_idx), set(_processed_combos))
            _mco_snap = working_df["MCO_Name"].copy()
            try:
                state_mco = f"{str(state).strip().upper()} MCO"
                # Real, DB-sourced TortId for this combo -- used wherever a TPA_Overrides
                # entry is built/completed below, so every entry ends up "complete" (real
                # TortId, not hardcoded None) and correctly Tort-scoped: a different Tort
                # sharing the same LienholderName can never pick up or clobber this entry.
                row_tort_id = _id_key(working_df.loc[row_mask, "TortId"].iloc[0]) if "TortId" in working_df.columns else None

                if _resolve_by_overrides(row_mask, state_mco, tort, mco_name, scope, row_tort_id):
                    continue

                result = resolve_tpa_lookup(state, tort, mco_name, scope)
                status = result["status"]

                if status == "do_not_use_skipped":
                    _send_to_review(row_mask, "Plan matched a TPA row marked DO NOT USE; user skipped")
                    continue

                if status == "do_not_use":
                    _dnu_row = result["row"]
                    _dnu_key = (str(state).strip().upper(), _normalize_key(mco_name),
                                _normalize_key(_dnu_row["LienholderName"]))
                    _plan_banner_once()
                    _new_item("plan")
                    print(f"\nPlan {_plan_label(mco_name, tort, scope)} matched TPA row "
                          f"{_dnu_row['LienholderName']!r} (LienholderId {_id_key(_dnu_row.get('LienholderId'))}). "
                          f"This row is marked DO NOT USE.")
                    while True:
                        _a = input("Use it anyway? (Y = use it / N = pick another plan from the list / "
                                   "S = skip to review): ").strip().upper()
                        if _a in ("Y", "N", "S"):
                            break
                        print("Please enter Y, N or S.")
                    _DNU_DECISIONS[_dnu_key] = _a
                    if _a == "S":
                        _send_to_review(row_mask, "Plan matched a TPA row marked DO NOT USE; user skipped")
                        continue
                    result = resolve_tpa_lookup(state, tort, mco_name, scope)
                    status = result["status"]
                    if status == "matched" and _a == "Y":
                        _dnu_entry = {
                            "TortId": _id_key(result["row"].get("TortId")) or row_tort_id, "TortName": tort,
                            "Private Lien Case Scope": scope, "CollectorId": result["row"].get("CollectorId"),
                            "LienholderId": result["row"].get("LienholderId"), "LienholderName": mco_name,
                            "LienType": result["row"].get("LienType"), "State MCO": state_mco,
                            "Ticket #": ticket_id, "Reason": TPA_REASON_DNU,
                        }
                        if _dnu_entry["TortId"] and not any(_is_blank(_dnu_entry[c]) for c in
                                                            ("CollectorId", "LienholderId", "LienType")):
                            _defer_save(_tpa_entry_label(_dnu_entry), tpa_dirty, _dnu_entry)

                if status == "state_not_covered":
                    # Same review path as an uncovered Tort: the user fills CollectorId/
                    # LienholderId/LienType once per plan and the resume finishes it end to end.
                    working_df.loc[row_mask, "PendingReason"] = "State not found in TPA reference file"
                    pending_idx.extend(working_df[row_mask].index.tolist())
                    _uncovered_state_idx.setdefault(str(state).strip().upper(), []).extend(working_df[row_mask].index.tolist())
                    continue

                if status == "matched":
                    row = result["row"]

                    # Real bug: the raw TPA row for this name can be known-stale
                    # (TPA_Overrides already corrects it under a different LienholderId) --
                    # if the current run's raw text doesn't exactly match the override's
                    # stored text, _tpa_overrides_find above never sees it, so this fuzzy
                    # "matched" result would be trusted silently. Two layers: a NARROW
                    # scoped-fuzzy check (this exact Tort+Scope+State bucket -- safe since
                    # it's tiny) that can auto-resolve; then a BROADER exact-name fallback
                    # (any Tort/Scope) that only proves the raw TPA data is untrustworthy
                    # somewhere, without a confirmed answer for this exact combination.
                    _scoped_match = _find_scoped_override_match(tpa_overrides, state_mco, tort, scope, row["LienholderName"])
                    if _scoped_match is not None:
                        _scoped_ok = _finalize_tpa_match(
                            row_mask, _scoped_match["LienholderId"], _scoped_match["CollectorId"],
                            _scoped_match["LienType"], tort_id=_scoped_match.get("TortId"),
                            tort_name=_scoped_match.get("TortName"),
                            scope_val=_scoped_match.get("Private Lien Case Scope"),
                            lienholder_name=_scoped_match.get("LienholderName"), state_mco_val=state_mco,
                            source=_override_row_label(_scoped_match),
                        )
                        if not _scoped_ok:
                            print(
                                f"TPA_Overrides entry for {mco_name!r} (Tort {tort!r} / Scope {scope!r}) "
                                f"is missing required details (CollectorId/LienType)."
                            )
                        # Real bug found: this used to alias to the raw TPA row's known-
                        # stale name instead of the override's actually-applied name -- the
                        # next run would rewrite raw text to the stale name, which
                        # _tpa_overrides_find's exact-match can't find (stored under the
                        # override's own name), silently losing the alias's direct-hit shortcut.
                        elif _normalize_key(mco_name) != _normalize_key(_scoped_match.get("LienholderName")):
                            alias_saves.append((state, mco_name, _scoped_match.get("LienholderName")))
                        continue

                    _conflict = _find_conflicting_override(tpa_overrides, state_mco, row["LienholderName"], row["LienholderId"])
                    if _conflict is not None:
                        # We know the raw TPA data for this name is untrustworthy somewhere
                        # (an override corrects it under a DIFFERENT Tort/Scope), but the
                        # scoped check above found nothing for THIS exact combination.
                        # Don't silently guess either source -- send it for manual review.
                        _new_item("plan")
                        print(
                            f"TPA file's entry for {row['LienholderName']!r} (LienholderId "
                            f"{row['LienholderId']!r}) conflicts with an existing TPA_Overrides "
                            f"correction for this same name (LienholderId {_conflict['LienholderId']!r}, "
                            f"saved for Tort {_conflict.get('TortName')!r} / Scope "
                            f"{_conflict.get('Private Lien Case Scope')!r}), but not for the current "
                            f"Tort {tort!r} / Scope {scope!r}. Needs manual review."
                        )
                        working_df.loc[row_mask, "PendingReason"] = (
                            "TPA name matched, but conflicts with a differently-scoped TPA_Overrides correction"
                        )
                        pending_idx.extend(working_df[row_mask].index.tolist())
                        continue

                    # LienholderName isn't checked here: _match_name only matches on real
                    # LienholderName text, so a blank one structurally can't reach "matched"
                    # status -- checking it would be dead code. Only CollectorId/LienType are
                    # independent of the matching cascade and can be blank on an otherwise-
                    # correct match. CollectorName is excluded too -- unused elsewhere.
                    if not _finalize_tpa_match(
                        row_mask, row["LienholderId"], row.get("CollectorId"), row.get("LienType"),
                        tort_id=row.get("TortId"), tort_name=row.get("TortName"),
                        scope_val=row.get("Private Lien Case Scope"), lienholder_name=row.get("LienholderName"),
                        state_mco_val=state_mco, source=f"the TPA file ({state_mco} sheet)",
                    ):
                        print(
                            f"TPA row matched for {mco_name!r} (Tort {tort!r} / Scope {scope!r}) but "
                            f"is missing required fields (CollectorId/LienType)."
                        )
                    continue

                if status == "tort_not_covered":
                    _new_item("plan")
                    print(f"Tort {tort} is not included in the TPA file for {state_mco}.")
                    working_df.loc[row_mask, "PendingReason"] = "Tort not found in TPA reference file"
                    pending_idx.extend(working_df[row_mask].index.tolist())
                    continue

                if status == "scope_gap":
                    _new_item("plan")
                    print(f"Plan {mco_name!r} was found in TPA for Tort {tort!r}, but not for scope {scope!r}.")
                    working_df.loc[row_mask, "PendingReason"] = "Plan found in TPA but not for this client's scope"
                    working_df.loc[row_mask, "_LPM_Match_Text"] = result["candidates"]["LienholderName"].iloc[0]
                    pending_idx.extend(working_df[row_mask].index.tolist())
                    continue

                if status == "name_not_found":
                    # Deliberately NOT fuzzy-matching the raw text against TPA_Overrides
                    # here (only an exact match, at the top of this loop, reuses a past
                    # override). Real data showed this is unsafe: a scoped bucket can hold
                    # several different real plans sharing one brand prefix (e.g. "HSO
                    # CareOR", "HSO Kaiser", "HSO Prov") -- _match_name's first-word gate
                    # passes on that shared prefix alone, silently resolving to a completely
                    # different plan's override whenever only one candidate is in the bucket.
                    _text_key = _normalize_key(mco_name)
                    candidates = result["candidates"]
                    # Placeholder TPA rows (e.g. "PlaceHolder Lien Creation", a real,
                    # intentional data pattern -- see LA MCO's IVC/Oxbryta/Talc rows) can
                    # have a blank LienholderName -- never a meaningful option to pick by
                    # name from this list; still reachable via the 'D' (LienHolder
                    # details / direct LienholderId) path if genuinely needed.
                    distinct_names = [
                        n for n in candidates["LienholderName"].drop_duplicates().tolist()
                        if _normalize_key(n) is not None
                    ]
                    _signature = (
                        _text_key,
                        tuple(sorted(_normalize_key(n) for n in distinct_names)),
                    )

                    # Announcing a group ("...for Tort 'Hernia Mesh' and 'Abuse'") but
                    # reaching 'Abuse' only later, with an unrelated question in between,
                    # would look disjointed. Gather the full group now (via
                    # _find_same_signature_combos, reusing the real _tpa_overrides_find/
                    # resolve_tpa_lookup so it can't drift), announce once, then ask every
                    # member's own question back-to-back -- each stays independent
                    # (CollectorId/LienType can differ by Tort/Scope), just marked
                    # processed so the outer loop skips them later.
                    _other = _find_same_signature_combos(
                        combos, _processed_combos | {combo_key}, tpa_overrides, _text_key, _signature
                    )
                    _group = [(tort, state, mco_name, scope, candidates, distinct_names)] + _other

                    _candidate_key = tuple(sorted(_normalize_key(n) for n in distinct_names))
                    _plan_banner_once()
                    _new_item("plan")
                    if _candidate_key in _seen_candidate_lists:
                        print(f"\nPlan name {mco_name!r} was not found in TPA for Tort "
                              f"{_join_tort_names([g[0] for g in _group])} (same candidate "
                              f"list already shown above for {_seen_candidate_lists[_candidate_key]!r}).")
                    else:
                        _seen_candidate_lists[_candidate_key] = mco_name
                        print(f"\nPlan name {mco_name!r} was not found in TPA for Tort "
                              f"{_join_tort_names([g[0] for g in _group])}.")
                        for i, n in enumerate(distinct_names, start=1):
                            print(f"  [{i}] {n}")

                    def _resolve_picked_plan(g_row_mask, g_tort, g_mco_name, g_scope, g_state, g_state_mco,
                                              g_candidates, picked_name, auto_reused):
                        """Resolves ONE specific already-picked LienholderName against THIS
                        combo's own Tort+Scope-filtered TPA rows, and finalizes/saves it --
                        shared by the live number-pick below AND the auto-reuse-across-
                        Torts path (see _group_picked_name below it), so both stay
                        provably identical instead of risking drift between two copies of
                        the same logic. auto_reused only changes the wording of messages,
                        never the underlying resolution -- CollectorId/LienType are always
                        freshly re-derived from THIS Tort's own TPA rows, never copied
                        from wherever the name was first picked."""
                        sub_matches = g_candidates[g_candidates["LienholderName"] == picked_name]
                        scope_key = _normalize_key(g_scope)
                        sub_scope_matches = sub_matches[sub_matches["Private Lien Case Scope"].apply(_normalize_key) == scope_key]
                        _scope_was_ambiguous = False
                        if len(sub_scope_matches) == 1:
                            picked = sub_scope_matches.iloc[0]
                        elif len(sub_scope_matches) > 1:
                            # Same ambiguity as resolve_tpa_lookup's ambiguous_duplicate
                            # status (Tort+Name+Scope still leaves 2+ TPA rows) -- same
                            # treatment: show candidates, let the user pick one or skip.
                            _gap()
                            print(
                                f"\nMultiple TPA rows match {picked_name!r} for Tort {g_tort!r} / "
                                f"Scope {g_scope!r}:"
                            )
                            _print_candidate_table(sub_scope_matches)
                            while True:
                                answer = input("Enter a number to pick one, or 'S' to skip: ").strip().upper()
                                if answer == "S":
                                    break
                                if answer.isdigit() and 1 <= int(answer) <= len(sub_scope_matches):
                                    break
                                print(f"Please enter a number between 1 and {len(sub_scope_matches)}, or 'S'.")
                            if answer == "S":
                                working_df.loc[g_row_mask, "PendingReason"] = (
                                    "Plan picked from list, but multiple TPA rows matched its scope; user skipped"
                                )
                                pending_idx.extend(working_df[g_row_mask].index.tolist())
                                return
                            picked = sub_scope_matches.iloc[int(answer) - 1]
                            _scope_was_ambiguous = True
                        else:
                            # Scope doesn't match any TPA row for this name+tort -- no
                            # valid option, straight to review (same as resolve_tpa_lookup's
                            # scope_gap case). Must print feedback here too, since no "Save?"
                            # prompt follows and otherwise a picked number gives no explanation.
                            print(
                                f"{picked_name!r} was {'auto-applied' if auto_reused else 'picked'} for "
                                f"{g_mco_name!r} (Tort {g_tort!r} / Scope {g_scope!r}), but its scope "
                                f"didn't narrow to exactly one TPA row -- sent for review."
                            )
                            working_df.loc[g_row_mask, "PendingReason"] = (
                                "Plan picked from list, but scope didn't narrow to exactly one TPA row"
                            )
                            working_df.loc[g_row_mask, "_LPM_Match_Text"] = picked_name
                            pending_idx.extend(working_df[g_row_mask].index.tolist())
                            return
                        # A TPA row reached via a manual list-pick can itself have blank
                        # CollectorId/LienType, same as an auto (fuzzy) "matched" row can --
                        # route to the same Pending_TPA_Blank_Fields mechanism instead of
                        # treating it as fully resolved and crashing downstream in
                        # get_assigned_user_id.
                        if not _finalize_tpa_match(
                            g_row_mask, picked["LienholderId"], picked.get("CollectorId"), picked.get("LienType"),
                            tort_id=picked.get("TortId"), tort_name=picked.get("TortName"),
                            scope_val=picked.get("Private Lien Case Scope"),
                            lienholder_name=picked.get("LienholderName"), state_mco_val=g_state_mco,
                            is_first_for_combo=False, source=f"the TPA file ({g_state_mco} sheet)",
                        ):
                            print(
                                f"TPA row {'auto-applied' if auto_reused else 'picked'} for {g_mco_name!r} "
                                f"(Tort {g_tort!r} / Scope {g_scope!r}) but is missing required fields "
                                f"(CollectorId/LienType)."
                            )
                            return
                        # The picked name is already a real TPA row for this Tort --
                        # persisting a TPA_Overrides entry is redundant (unlike 'D' flow,
                        # where the data can't be found any other way): once Alias_Map
                        # rewrites this raw text, the normal TPA cascade finds it on its
                        # own every future run. Kept in-memory only (tpa_overrides.append,
                        # not tpa_dirty) so a later combo this run skips re-asking, but
                        # nothing persists to mco_config.xlsx -- only the alias does, below.
                        new_entry = {
                            "TortId": picked.get("TortId"), "TortName": g_tort, "Private Lien Case Scope": g_scope,
                            "CollectorId": picked["CollectorId"], "LienholderId": picked["LienholderId"],
                            "LienholderName": g_mco_name,
                            "LienType": picked["LienType"], "State MCO": g_state_mco, "Ticket #": ticket_id,
                            "Reason": None,
                        }
                        tpa_overrides.append(new_entry)
                        # Remembered as a plain name alias so the next raw occurrence of this
                        # text (any Tort/scope/client) rewrites to the canonical name and
                        # resolves via the normal cascade. Queued for the end-of-run save
                        # question; an auto-reused Tort queues the same alias, listed once.
                        _same_text = _normalize_key(g_mco_name) == _normalize_key(picked_name)
                        if not _same_text:
                            _defer_save(f"Plan name {g_mco_name!r} -> TPA plan {picked_name!r} (State {g_state})",
                                        alias_saves, (g_state, g_mco_name, picked_name))
                        if _scope_was_ambiguous or _same_text:
                            # An alias can't help here: the TPA file has 2+ rows for this
                            # Tort+Name+Scope, or raw text equals the picked name (a tie) --
                            # the cascade would hit the same question again. Persisted under
                            # the CANONICAL name (resolve_tpa_lookup's ambiguous_duplicate
                            # convention) so _tpa_overrides_find catches it first.
                            _canon_entry = {**new_entry, "LienholderName": picked_name}
                            _defer_save(_tpa_entry_label(_canon_entry), tpa_dirty, _canon_entry)

                    # Once a LienholderName is picked for the first Tort in this group,
                    # that choice applies to every other Tort sharing this raw text --
                    # unlike 'D' flow's CollectorId/LienType, which can differ by
                    # Tort/Scope and stays independent. CollectorId/LienType are always
                    # re-derived per Tort via _resolve_picked_plan, so reusing just the
                    # name is safe. 'D' and 'S' choices are never reused this way.
                    _group_picked_name = None
                    for (g_tort, g_state, g_mco_name, g_scope, g_candidates, g_distinct_names) in _group:
                        g_row_mask = _combo_mask(g_state, g_tort, g_mco_name, g_scope)
                        _processed_combos.add((g_state, g_tort, g_mco_name, g_scope))
                        if not g_row_mask.any():
                            continue
                        g_state_mco = f"{str(g_state).strip().upper()} MCO"
                        g_row_tort_id = (
                            _id_key(working_df.loc[g_row_mask, "TortId"].iloc[0]) if "TortId" in working_df.columns else None
                        )

                        if _group_picked_name is not None and _group_picked_name in g_distinct_names:
                            print(f"For {g_tort} Tort / Scope {g_scope!r}: automatically using "
                                  f"{_group_picked_name!r} (same plan picked earlier this run for "
                                  f"{mco_name!r}).")
                            _resolve_picked_plan(g_row_mask, g_tort, g_mco_name, g_scope, g_state, g_state_mco,
                                                  g_candidates, _group_picked_name, auto_reused=True)
                            continue

                        # An invalid answer here (not a valid number, 'D', or 'S') would
                        # otherwise fall through to int(answer) below and crash the run --
                        # validate and re-prompt instead, same defensive pattern as
                        # ask_numbered_choice/ask_yes_no elsewhere.
                        _gap()
                        while True:
                            answer = input(
                                f"For {g_tort} Tort / Scope {g_scope!r}, enter a number to pick a plan, "
                                f"'D' to provide LienHolder details, or 'S' to skip: "
                            ).strip().upper()
                            if answer in ("S", "D"):
                                break
                            if answer.isdigit() and 1 <= int(answer) <= len(g_distinct_names):
                                if not _is_do_not_use(g_distinct_names[int(answer) - 1]) or ask_yes_no(
                                        f"{g_distinct_names[int(answer) - 1]!r} is marked DO NOT USE. "
                                        f"Are you sure? (Y/N): "):
                                    break
                                continue
                            print(f"Please enter a number between 1 and {len(g_distinct_names)}, 'D', or 'S'.")

                        if answer == "S":
                            working_df.loc[g_row_mask, "PendingReason"] = _name_not_found_pending_reason(
                                g_state, g_tort, g_candidates
                            )
                            working_df.loc[g_row_mask, "_LPM_Match_Text"] = g_mco_name
                            pending_idx.extend(working_df[g_row_mask].index.tolist())
                            continue
                        if answer == "D":
                            # The numbered list above is only a narrowed shortlist (see
                            # _plan_candidates_for_prompt), not every TPA plan for this
                            # Tort/State -- 'D' is for when the user knows the LienholderId
                            # and it just isn't one of the names shown. Blank/'S' backs out
                            # to the same skip protocol as the top-level prompt.
                            _gap()
                            given_lhid = _ask_typed_id("LienholderId (or 'S' to skip): ")
                            if given_lhid == "S":
                                working_df.loc[g_row_mask, "PendingReason"] = _name_not_found_pending_reason(
                                    g_state, g_tort, g_candidates
                                )
                                working_df.loc[g_row_mask, "_LPM_Match_Text"] = g_mco_name
                                pending_idx.extend(working_df[g_row_mask].index.tolist())
                                continue
                            _gap()
                            collector_id = _ask_typed_id("CollectorId: ", allow_skip=False)
                            lien_type = ask_numbered_choice(LIEN_TYPE_OPTIONS, "LienType:", "Enter the number: ")
                            # Routed through the shared _finalize_tpa_match helper -- a
                            # blank CollectorId here (an empty Enter at the prompt) used
                            # to be assigned directly, bypassing the blank-fields check
                            # every other resolution path already has.
                            if not _finalize_tpa_match(
                                g_row_mask, given_lhid, collector_id, lien_type,
                                tort_id=g_row_tort_id, tort_name=g_tort, scope_val=g_scope,
                                lienholder_name=None, state_mco_val=g_state_mco,
                                is_first_for_combo=False, source="the values typed in this run",
                            ):
                                print(
                                    f"LienholderId {given_lhid!r} provided for {g_mco_name!r} "
                                    f"(Tort {g_tort!r} / Scope {g_scope!r}) but CollectorId is missing."
                                )
                                continue
                            new_entry = {
                                "TortId": g_row_tort_id, "TortName": g_tort, "Private Lien Case Scope": g_scope,
                                "CollectorId": collector_id, "LienholderId": given_lhid,
                                "LienholderName": g_mco_name, "LienType": lien_type, "State MCO": g_state_mco,
                                "Ticket #": ticket_id, "Reason": TPA_REASON_USER_LHID,
                            }
                            tpa_overrides.append(new_entry)
                            _defer_save(_tpa_entry_label(new_entry), tpa_dirty, new_entry)
                            continue
                        picked_name = g_distinct_names[int(answer) - 1]
                        _resolve_picked_plan(g_row_mask, g_tort, g_mco_name, g_scope, g_state, g_state_mco,
                                              g_candidates, picked_name, auto_reused=False)
                        _group_picked_name = picked_name
                    continue

                if status == "ambiguous_duplicate":
                    candidates = result["candidates"]

                    # ambiguous_duplicate always narrows to ONE shared TPA name first, so
                    # the correct key to save/check under is that stable TPA name, not the
                    # raw run-time text that fuzzy-matched it. Checking here before printing
                    # means a later occurrence -- any wording that fuzzy-matches this same
                    # TPA name -- resolves silently, no Alias_Map needed.
                    _canonical_name = candidates.iloc[0]["LienholderName"]
                    _amb_ov_idx = _tpa_overrides_find(tpa_overrides, state_mco, tort, _canonical_name, scope)
                    if _amb_ov_idx is not None:
                        _amb_entry = tpa_overrides[_amb_ov_idx]
                        if not _finalize_tpa_match(
                            row_mask, _amb_entry["LienholderId"], _amb_entry["CollectorId"], _amb_entry["LienType"],
                            tort_id=_amb_entry.get("TortId"), tort_name=_amb_entry.get("TortName"),
                            scope_val=_amb_entry.get("Private Lien Case Scope"),
                            lienholder_name=_amb_entry.get("LienholderName"), state_mco_val=state_mco,
                            source=_override_row_label(_amb_entry),
                        ):
                            print(
                                f"TPA_Overrides entry for {mco_name!r} (Tort {tort!r} / Scope {scope!r}) "
                                f"is missing required details (CollectorId/LienType)."
                            )
                        continue

                    _plan_banner_once()
                    _new_item("plan")
                    print(f"\nMultiple TPA rows match {mco_name!r} for Tort {tort!r} / Scope {scope!r}:")
                    _print_candidate_table(candidates)
                    # An invalid answer here would otherwise fall through to int(answer)
                    # below and crash the run -- validate and re-prompt instead.
                    while True:
                        answer = input("Enter a number to pick one, or 'S' to skip: ").strip().upper()
                        if answer == "S":
                            break
                        if answer.isdigit() and 1 <= int(answer) <= len(candidates):
                            break
                        print(f"Please enter a number between 1 and {len(candidates)}, or 'S'.")
                    if answer == "S":
                        working_df.loc[row_mask, "PendingReason"] = (
                            f"Multiple matching TPA entries for this plan and scope "
                            f"({len(candidates)} TPA rows share this exact name); user skipped"
                        )
                        pending_idx.extend(working_df[row_mask].index.tolist())
                        continue
                    picked = candidates.iloc[int(answer) - 1]
                    # Routed through the shared _finalize_tpa_match helper (matches every
                    # other resolution path's blank-fields check) -- also means MCO_Name
                    # gets rewritten to the canonical name here too, consistently.
                    if not _finalize_tpa_match(
                        row_mask, picked["LienholderId"], picked["CollectorId"], picked["LienType"],
                        tort_id=picked.get("TortId"), tort_name=tort, scope_val=scope,
                        lienholder_name=_canonical_name, state_mco_val=state_mco,
                        is_first_for_combo=False, source=f"the TPA file ({state_mco} sheet)",
                    ):
                        print(
                            f"TPA row picked for {mco_name!r} (Tort {tort!r} / Scope {scope!r}) but "
                            f"is missing required fields (CollectorId/LienType)."
                        )
                        continue
                    new_entry = {
                        "TortId": picked.get("TortId"), "TortName": tort, "Private Lien Case Scope": scope,
                        "CollectorId": picked["CollectorId"], "LienholderId": picked["LienholderId"],
                        # Saved under the CANONICAL TPA name (not mco_name/raw text) -- lets
                        # the check above find it again regardless of wording differences
                        # in future raw input text, with no Alias_Map entry needed.
                        "LienholderName": _canonical_name,
                        "LienType": picked["LienType"], "State MCO": state_mco, "Ticket #": ticket_id,
                        "Reason": TPA_REASON_PICKED_DUP,
                    }
                    tpa_overrides.append(new_entry)
                    _defer_save(_tpa_entry_label(new_entry), tpa_dirty, new_entry)
                    continue
            except Exception as _combo_err:
                _new_keys = (_processed_combos - _snap[6]) | {combo_key}
                del tpa_overrides[_snap[0]:]
                del tpa_dirty[_snap[1]:]
                del alias_saves[_snap[2]:]
                del _DEFERRED_SAVES[_snap[3]:]
                del pending_idx[_snap[4]:]
                del blank_field_idx[_snap[5]:]
                _err_mask = pd.Series(False, index=working_df.index)
                for _k in _new_keys:
                    _k_mask = ~working_df.index.isin(list(_done_idx))
                    for _col, _val in zip(("State", "Project_Name", "MCO_Name", "Scope_Classification"), _k):
                        _series = _mco_snap if _col == "MCO_Name" else working_df[_col]
                        _k_mask = _k_mask & (_series.isna() if _is_blank(_val) else (_series == _val))
                    _err_mask = _err_mask | _k_mask
                _err_idx = working_df.index[_err_mask].tolist()
                working_df.loc[_err_idx, "MCO_Name"] = _mco_snap.loc[_err_idx]
                _record_plan_error(working_df.loc[_err_idx], "Step 3 (TPA lookup)", _combo_err)
                error_idx.extend(_err_idx)
                _done_idx.update(_err_idx)
                _processed_combos.update(_new_keys)

        # --- LienholderId from LPM for review rows (see _lpm_state_lienholders) ---
        # State/Tort not in TPA: a sure name match, else the user picks from the state's LPM
        # list (once per plan text). MCO_Name becomes that LPM name, so every spelling of one
        # plan reads the same, and TPA_Overrides is checked again under it. Other review rows:
        # pre-filled only on a sure match -- no question, MCO_Name unchanged.
        _lpm_found = {}

        def _lpm_remembered_id(state_code, name, ids):
            hits = {
                _stringify_id(e.get("LienholderId")) for e in tpa_overrides
                if e.get("TortId") is not None
                and _normalize_key(e.get("State MCO")) == _normalize_key(f"{state_code} MCO")
                and _normalize_key(e.get("LienholderName")) == _normalize_key(name)
            } & set(ids)
            return next(iter(hits)) if len(hits) == 1 else None

        def _ask_lpm_pick(text, state_code, cands, same_name=None):
            _plan_banner_once()
            _new_item("plan")
            if same_name:
                print(f"\nPlan {text!r} (State {state_code}) matches LPM lienholder {same_name!r}, "
                      f"which has {len(cands)} records in LPM:")
            else:
                print(f"\nPlan {text!r} (State {state_code}) is not in the TPA file and no LPM "
                      f"lienholder matched it by name.")
                print(f"LPM lienholders for {state_code} (closest first):")
            q_words = set(_tokenize_plan(_strip_lpm_tag(text).lower())) - _STRUCTURAL_TAGS
            def _closeness(pair):
                c_words = set(_tokenize_plan(_strip_lpm_tag(pair[1]).lower())) - _STRUCTURAL_TAGS
                return (not _lpm_initials_match(text, pair[1]), -len(q_words & c_words),
                        _strip_lpm_tag(pair[1]).upper(), pair[0])
            ordered = sorted(zip(cands["LienHolderId"], cands["Name"]), key=_closeness)
            for i, (lh_id, name) in enumerate(ordered, start=1):
                print(f"  [{i}] {name} -- LienholderId {lh_id}")
            while True:
                answer = input("Enter a number to pick, or 'S' to skip "
                               "(fill LienholderId in the review file instead): ").strip().upper()
                if answer == "S":
                    return None
                if answer.isdigit() and 1 <= int(answer) <= len(ordered):
                    return (*ordered[int(answer) - 1], True)
                print(f"Please enter a number between 1 and {len(ordered)}, or 'S'.")

        def _lpm_lookup(state_code, text, allow_pick):
            """(LienholderId, LPM name, picked_by_user) or None. A user's pick (or skip) is
            remembered for the rest of the run; a failed silent match is not."""
            key = (state_code, _normalize_key(text))
            if key in _lpm_found:
                return _lpm_found[key]
            lpm_df = _lpm_state_lienholders(state_code)
            found = None
            if lpm_df is not None and not lpm_df.empty and _normalize_key(text):
                cands = lpm_df.assign(_Match=lpm_df["Name"].map(_strip_lpm_tag))
                m_id, _ = _match_name(text, cands, "_Match", "LienHolderId", return_tie=True)
                if m_id is not None:
                    m_name = cands.loc[cands["LienHolderId"] == m_id, "Name"].iloc[0]
                    same = cands[cands["Name"].map(_normalize_key) == _normalize_key(m_name)]
                    known = (same["LienHolderId"].iloc[0] if len(same) == 1
                             else _lpm_remembered_id(state_code, m_name, set(same["LienHolderId"])))
                    if known is not None:
                        found = (known, m_name, False)
                    elif allow_pick:
                        found = _ask_lpm_pick(text, state_code, same, same_name=m_name)
                elif allow_pick:
                    found = _ask_lpm_pick(text, state_code, cands)
            if found is not None or allow_pick:
                _lpm_found[key] = found
            return found

        _lpm_resolved_idx = set()
        _lpm_b_idx = [i for i in pending_idx if working_df.at[i, "PendingReason"] in _LPM_REVIEW_REASONS]
        if _lpm_b_idx:
            for (_b_state, _b_raw), _b_idx in working_df.loc[_lpm_b_idx].groupby(
                    ["State", "MCO_Name"], sort=False, dropna=False).groups.items():
              try:
                _b_st = str(_b_state).strip().upper()
                _b_found = _lpm_lookup(_b_st, _b_raw, allow_pick=True)
                if _b_found is None:
                    continue
                _b_id, _b_name, _b_picked = _b_found
                if _b_picked and _normalize_key(_b_raw) != _normalize_key(_b_name):
                    _defer_save(f"Plan name {_b_raw!r} -> LPM lienholder {_b_name!r} (State {_b_st})",
                                alias_saves, (_b_st, _b_raw, _b_name))
                working_df.loc[_b_idx, "LienholderId"] = _b_id
                working_df.loc[_b_idx, "LPM_LienholderName"] = _b_name
                for (_b_tort, _b_scope), _b_sub in working_df.loc[_b_idx].groupby(
                        ["Project_Name", "Scope_Classification"], dropna=False).groups.items():
                    _b_ov = _tpa_overrides_find(tpa_overrides, f"{_b_st} MCO", _b_tort, _b_name, _b_scope)
                    if _b_ov is None or tpa_overrides[_b_ov].get("TortId") is None:
                        continue
                    _b_entry = tpa_overrides[_b_ov]
                    _finalize_tpa_match(
                        working_df.index.isin(_b_sub), _b_entry["LienholderId"], _b_entry["CollectorId"],
                        _b_entry["LienType"], tort_id=_b_entry.get("TortId"), tort_name=_b_entry.get("TortName"),
                        scope_val=_b_entry.get("Private Lien Case Scope"),
                        lienholder_name=_b_entry.get("LienholderName"), state_mco_val=f"{_b_st} MCO",
                        source=_override_row_label(_b_entry),
                    )
                    _lpm_resolved_idx.update(_b_sub)
                _b_still = [i for i in _b_idx if i not in _lpm_resolved_idx]
                if _b_still:
                    # Same re-keying as _finalize_tpa_match: PLAN_FATE is keyed by the old text.
                    for _b_lid in working_df.loc[_b_still, "LienId"].unique():
                        PLAN_FATE.pop((_b_lid, _b_raw), None)
                    working_df.loc[_b_still, "MCO_Name"] = _b_name
              except Exception as _lpm_err:
                # Only a pre-fill for review rows: they stay in Pending_TPA_Review as they are.
                print(f"⚠️  Could not look up {_b_raw!r} in LPM ({_short_text(str(_lpm_err), 120)}) -- "
                      f"left in Pending_TPA_Review without a LienholderId.")
            pending_idx = [i for i in pending_idx if i not in _lpm_resolved_idx]
            _lpm_ok = [i for i in _lpm_resolved_idx if i not in blank_field_idx]
            working_df.loc[_lpm_ok, "PendingReason"] = None

        _lpm_t_rows = working_df.loc[pending_idx]
        _lpm_t_rows = _lpm_t_rows[_lpm_t_rows["_LPM_Match_Text"].notna() & _lpm_t_rows["LienholderId"].apply(_is_blank)]
        for (_t_state, _t_text), _t_idx in _lpm_t_rows.groupby(["State", "_LPM_Match_Text"], sort=False).groups.items():
            _t_found = _lpm_lookup(str(_t_state).strip().upper(), _t_text, allow_pick=False)
            if _t_found is not None:
                working_df.loc[_t_idx, "LienholderId"] = _t_found[0]
                working_df.loc[_t_idx, "LPM_LienholderName"] = _t_found[1]

        for _uc_state, _uc_idx in _uncovered_state_idx.items():
            _uc_idx = [i for i in _uc_idx if i not in _lpm_resolved_idx]
            if not _uc_idx:
                continue
            _new_item("plan")
            print(
                f"⚠️  State {_uc_state} is not in the TPA file: {len(_uc_idx)} plan row(s) for "
                f"{working_df.loc[_uc_idx, 'ClientId'].nunique()} client(s) sent to Pending_TPA_Review. "
                f"LienholderId is pre-filled from LPM where found; fill CollectorId and LienType "
                f"there, once per plan."
            )

        pending_idx = [i for i in dict.fromkeys(pending_idx) if i not in error_idx]
        blank_field_idx = [i for i in dict.fromkeys(blank_field_idx) if i not in error_idx and i not in pending_idx]
        pending_tpa_df = working_df.loc[pending_idx].drop(columns=["_LPM_Match_Text", "_TPA_Source"]).copy()
        pending_blank_fields_df = working_df.loc[blank_field_idx].drop(columns=_review_only_cols + ["_TPA_Source"]).copy()
        if not pending_tpa_df.empty:
            # Blank column the user can fill in offline with the real TPA plan name
            # (instead of hand-typing LienholderId/CollectorId/LienType) -- resolved on
            # resume by aliasing to that name (see the resume-preseed "Picked
            # LienholderName" handling), same mechanism the live number-pick uses, so
            # the row resolves via the normal TPA cascade, not a special path.
            pending_tpa_df["Picked LienholderName"] = None
            # Tort blank in LPM: the user fills TortId + TortName (one row per case).
            if (pending_tpa_df["PendingReason"] == BLANK_TORT_REASON).any():
                pending_tpa_df["TortName"] = None
        for _pdf in (pending_tpa_df, pending_blank_fields_df):
            if not _pdf.empty:
                _pdf["TicketId"] = ticket_id
                _pdf["ResponseDate"] = RESPONSE_DATE
                _pdf["ResponseType"] = RESPONSE_TYPE
        resolved_df = working_df.drop(index=pending_idx + blank_field_idx + list(dict.fromkeys(error_idx))).drop(
            columns=_review_only_cols).copy()
        return resolved_df, pending_tpa_df, pending_blank_fields_df, tpa_dirty

    yes_with_mco_df, pending_tpa_df, pending_tpa_blank_fields_df, _TPA_DIRTY = resolve_tpa_for_run(
        yes_with_mco_df, TPA_OVERRIDES, TicketId, _ALIAS_SAVES
    )

    for _reason, _grp in pending_tpa_df.groupby("PendingReason") if not pending_tpa_df.empty else []:
        record_fate(_grp, f"{_reason}; sent for review -- see Pending_TPA_Review (Review file)")
    for _reason, _grp in pending_tpa_blank_fields_df.groupby("PendingReason") if not pending_tpa_blank_fields_df.empty else []:
        record_fate(_grp, f"{_reason}; sent for review -- see Pending_TPA_Blank_Fields (Review file)")

    def merge_same_plan_within_client(df_in):
        """A free-text cell can split into multiple raw plan-name fragments that
        separately resolve, via separate name_not_found list-picks, to the SAME real
        TPA plan for the same client+Tort. Without this merge, Step 4 sees 2 rows for
        one plan -- the double-claim guard lets the first correctly claim the
        existing lien, then the second falls through to Scenario C, creating a wrong
        duplicate "new lien". Collapses rows sharing (ClientId, Project_Name,
        Scope_Classification, LienholderId, CollectorId), combining Enrollment dates
        the same way the "multiple episodes" rule does (earliest start, latest end)
        -- MCO_Name becomes a "/"-joined list of every raw text merged, so the
        LienNote's "Plan Name:" stays traceable."""
        # Real bug found: LienId was missing from merge_keys. This merge is only meant
        # for raw-text fragments from the SAME input cell (always sharing one LienId)
        # resolving to the same plan -- without LienId, two genuinely DIFFERENT liens
        # for the same client resolving to the same (Tort, Scope, LienholderId,
        # CollectorId) would be wrongly collapsed, silently dropping one's own outcome.
        merge_keys = ["ClientId", "LienId", "Project_Name", "Scope_Classification", "LienholderId", "CollectorId"]
        if df_in.empty or not all(k in df_in.columns for k in merge_keys):
            return df_in
        mergeable = df_in[df_in["LienholderId"].notna() & df_in["CollectorId"].notna()]
        if mergeable.empty:
            return df_in
        group_sizes = mergeable.groupby(merge_keys, dropna=False).size()
        dup_keys = group_sizes[group_sizes > 1].index.tolist()
        if not dup_keys:
            return df_in

        has_dates = "Start_Date" in df_in.columns and "End_Date" in df_in.columns
        rows_to_drop = []
        merged_rows = []
        for key in dup_keys:
            key = key if isinstance(key, tuple) else (key,)
            key_mask = pd.Series(True, index=mergeable.index)
            for col, val in zip(merge_keys, key):
                key_mask = key_mask & (mergeable[col] == val)
            group_df = mergeable[key_mask]
            base_row = group_df.iloc[0].copy()
            distinct_texts = group_df["MCO_Name"].drop_duplicates().tolist()
            base_row["MCO_Name"] = " / ".join(distinct_texts)
            if has_dates:
                # Real bug found: by this point Start_Date/End_Date are already
                # formatted "MM/DD/YYYY" strings (format_date ran right after the first
                # plan-row grouping), but safe_min/safe_max expect raw datetimes --
                # every value failed isinstance(d, str), always returning "Bad Date"
                # even when both dates were valid. Re-parse each string back to a
                # datetime first. Also normalize blanks to real None -- pandas silently
                # turns a None in an all-string column into float NaN (not NaT), which
                # safe_min's "is not None" check misses, raising TypeError in min/max.
                def _reparse_formatted_date(d):
                    if pd.isna(d):
                        return None
                    if isinstance(d, str) and d != "Bad Date":
                        return _parse_date(d)
                    return d
                _reparsed_starts = [_reparse_formatted_date(d) for d in group_df["Start_Date"].tolist()]
                _reparsed_ends = [_reparse_formatted_date(d) for d in group_df["End_Date"].tolist()]
                base_row["Start_Date"] = format_date(safe_min(_reparsed_starts))
                base_row["End_Date"]   = format_date(safe_max(_reparsed_ends))
            if "Observation" in group_df.columns:
                base_row["Observation"] = _join_observations(group_df["Observation"].tolist())
            if "_OpenEnd" in group_df.columns and group_df["_OpenEnd"].map(_is_true_flag).any():
                base_row["_OpenEnd"] = True
                if has_dates:
                    base_row["End_Date"] = None
            merged_rows.append(base_row)
            rows_to_drop.extend(group_df.index.tolist())

            # PLAN_FATE is seeded with one placeholder per original pre-merge raw
            # MCO_Name, but Step 4 only records a fate against the new merged key --
            # drop the superseded original keys here or Remarks shows fake "tracking
            # gap" warnings alongside the real outcome. Nothing is lost: the merged
            # key's plan-name text already lists every original raw value.
            for _, _orig_row in group_df.iterrows():
                PLAN_FATE.pop((_orig_row["LienId"], _orig_row["MCO_Name"]), None)

        result = df_in.drop(index=rows_to_drop)
        result = pd.concat([result, pd.DataFrame(merged_rows)], ignore_index=True)
        return result

    yes_with_mco_df = merge_same_plan_within_client(yes_with_mco_df)

    def save_config_updates(config_path, scope_saves, tpa_dirty, alias_saves=None, lien_equiv_saves=None,
                             column_map_saves=None):
        """Single combined batched write: one open -> apply every pending Scope_Map AND
        TPA_Overrides AND Alias_Map AND Lien_Name_Equivalence_Map AND Column_Map update
        -> one save. alias_saves/lien_equiv_saves/column_map_saves default to None
        (treated as empty) so existing callers that don't pass them still work."""
        if alias_saves is None:
            alias_saves = []
        if lien_equiv_saves is None:
            lien_equiv_saves = []
        if column_map_saves is None:
            column_map_saves = []
        if not scope_saves and not tpa_dirty and not alias_saves and not lien_equiv_saves and not column_map_saves:
            return
        if not _check_file_not_locked(config_path):
            print("Config save cancelled -- nothing was written.")
            return
        _backup_config_file(config_path)
        _ensure_file_hydrated(config_path)
        _config_save_fallback_msg = (
            f"\n⚠️  Could not save config updates -- '{config_path}' appears to be open "
            f"elsewhere (e.g. in Excel). Close it and add these manually:\n"
            f"  Scope_Map: {scope_saves}\n  TPA_Overrides: {tpa_dirty}\n"
            f"  Alias_Map: {alias_saves}\n  Lien_Name_Equivalence_Map: {lien_equiv_saves}\n"
            f"  Column_Map: {column_map_saves}"
        )
        try:
            cfg_wb = _retry_on_permission_error(lambda: _opxl.load_workbook(config_path), config_path)
        except PermissionError:
            print(_config_save_fallback_msg)
            return
        # Snapshot of the file's on-disk state the moment we loaded it -- re-checked
        # right before the final save below. If it changed in between (e.g. synced in
        # from another user's own save on a different machine), that means our in-memory
        # cfg_wb was built from data that's no longer current -- saving it as-is would
        # silently overwrite whatever they just added.
        _loaded_fingerprint = _file_fingerprint(config_path)

        # Every decision saved below also gets one Audit_Trail row -- who, when, exactly
        # what changed, which ticket, and which run-log has the full prompt/answer proof.
        # Collected here and written at the very end, in the SAME open workbook as the 5
        # sheets above, so it only persists if the real cfg_wb.save() below succeeds.
        _audit_entries = []

        if scope_saves:
            cfg_ws, header = _config_sheet(cfg_wb, "Scope_Map", ["ScopeText", "Classification"])
            # A blank-scope decision is saved per CaseId (ScopeText left blank); the
            # CaseId column is added the first time one is saved.
            if any(len(item) > 2 and item[2] for item in scope_saves) and "CaseId" not in header:
                cfg_ws.cell(row=1, column=len(header) + 1, value="CaseId")
                header.append("CaseId")
            text_col  = header.index("ScopeText") + 1
            class_col = header.index("Classification") + 1
            case_col  = header.index("CaseId") + 1 if "CaseId" in header else None
            existing_rows = {}
            for r in range(2, cfg_ws.max_row + 1):
                key = _normalize_key(cfg_ws.cell(row=r, column=text_col).value)
                if key is None and case_col:
                    key = _case_scope_key(cfg_ws.cell(row=r, column=case_col).value)
                if key:
                    existing_rows[key] = r
            for item in scope_saves:
                scope_text, classification = item[0], item[1]
                case_id = item[2] if len(item) > 2 else None
                key = _normalize_key(scope_text) or _case_scope_key(case_id)
                label = scope_text if _normalize_key(scope_text) else f"CaseId {case_id} (scope blank in LPM)"
                if key in existing_rows:
                    _old_classification = cfg_ws.cell(row=existing_rows[key], column=class_col).value
                    cfg_ws.cell(row=existing_rows[key], column=class_col, value=classification)
                    _audit_entries.append(_audit_row(
                        "Scope_Map", label,
                        _old_classification if _old_classification not in (None, "") else "(was blank)",
                        classification, TicketId,
                    ))
                else:
                    new_row = cfg_ws.max_row + 1
                    if _normalize_key(scope_text):
                        cfg_ws.cell(row=new_row, column=text_col, value=scope_text)
                    else:
                        cfg_ws.cell(row=new_row, column=case_col, value=_id_key(case_id))
                    cfg_ws.cell(row=new_row, column=class_col, value=classification)
                    existing_rows[key] = new_row
                    _audit_entries.append(_audit_row("Scope_Map", label, "(new record)", classification, TicketId))
            print(f"\n✅ Saved {len(scope_saves)} scope classification update(s) to Scope_Map.")

        if tpa_dirty:
            # A column this version writes but an older sheet lacks (e.g. "Reason") is added.
            cfg_ws2, header2 = _config_sheet(cfg_wb, "TPA_Overrides", TPA_OVERRIDES_COLS)
            col_idx = {name: i + 1 for i, name in enumerate(header2) if name}

            existing_rows2 = list(range(2, cfg_ws2.max_row + 1))

            def _find_target_row(entry):
                # Mirrors _tpa_overrides_find: a COMPLETE row (TortId filled) must match
                # State MCO + TortName + LienholderName + Scope, so two different Torts
                # sharing the same LienholderId/Name/State get separate rows instead of
                # colliding. A PENDING row (TortId blank, e.g. a legacy entry) matches on
                # State MCO + LienholderName alone, so first-use completion patches it in
                # place. Keying on LienholderId+Name+State alone would let a second
                # Tort's entry clobber the first Tort's row.
                norm_state = _normalize_key(entry.get("State MCO"))
                norm_tort  = _normalize_key(entry.get("TortName"))
                norm_name  = _normalize_key(entry.get("LienholderName"))
                norm_scope = _normalize_key(entry.get("Private Lien Case Scope"))
                complete_row, pending_row = None, None
                for r in existing_rows2:
                    row_state = _normalize_key(cfg_ws2.cell(row=r, column=col_idx["State MCO"]).value)
                    row_name  = _normalize_key(cfg_ws2.cell(row=r, column=col_idx["LienholderName"]).value)
                    if row_state != norm_state or row_name != norm_name:
                        continue
                    row_tort_id = cfg_ws2.cell(row=r, column=col_idx["TortId"]).value
                    if row_tort_id is None or str(row_tort_id).strip() == "":
                        # An old no-Tort row is filled in place only with its OWN LienholderId
                        # and a real Tort; any other entry gets a new row (never overwrites it).
                        if (pending_row is None and _id_key(entry.get("TortId"))
                                and _id_key(cfg_ws2.cell(row=r, column=col_idx["LienholderId"]).value)
                                == _id_key(entry.get("LienholderId"))):
                            pending_row = r
                        continue
                    row_tort  = _normalize_key(cfg_ws2.cell(row=r, column=col_idx["TortName"]).value)
                    row_scope = _normalize_key(cfg_ws2.cell(row=r, column=col_idx["Private Lien Case Scope"]).value)
                    if row_tort == norm_tort and row_scope == norm_scope:
                        complete_row = r
                return complete_row if complete_row is not None else pending_row

            for entry in tpa_dirty:
                target_row = _find_target_row(entry)
                if target_row is None:
                    target_row = cfg_ws2.max_row + 1
                    existing_rows2.append(target_row)
                    _old_snapshot = "(new record)"
                else:
                    _old_snapshot = ", ".join(
                        f"{c}={cfg_ws2.cell(row=target_row, column=col_idx[c]).value if cfg_ws2.cell(row=target_row, column=col_idx[c]).value not in (None, '') else '(blank)'}"
                        for c in TPA_OVERRIDES_COLS
                    )
                for col_name in TPA_OVERRIDES_COLS:
                    cfg_ws2.cell(row=target_row, column=col_idx[col_name], value=entry.get(col_name))
                _new_snapshot = ", ".join(f"{c}={entry.get(c)}" for c in TPA_OVERRIDES_COLS)
                _tpa_key = (
                    f"State={entry.get('State MCO')} | Tort={entry.get('TortName')} | "
                    f"Plan={entry.get('LienholderName')} | Scope={entry.get('Private Lien Case Scope')}"
                )
                # entry's OWN "Ticket #" (set at the point this decision was actually made,
                # including the real per-row ticket on a resumed Pending_TPA_Review row) is
                # more precise than this run's module-level TicketId -- prefer it.
                _audit_entries.append(_audit_row(
                    "TPA_Overrides", _tpa_key, _old_snapshot, _new_snapshot, entry.get("Ticket #") or TicketId,
                ))
            print(f"\n✅ Saved {len(tpa_dirty)} TPA override update(s) to TPA_Overrides.")

        if alias_saves:
            cfg_ws3, header3 = _config_sheet(cfg_wb, "Alias_Map", ["State", "From_Name", "To_Name"])
            state_col = header3.index("State") + 1
            from_col  = header3.index("From_Name") + 1
            to_col    = header3.index("To_Name") + 1
            existing_alias_rows = {}
            for r in range(2, cfg_ws3.max_row + 1):
                key = (
                    _normalize_key(cfg_ws3.cell(row=r, column=state_col).value),
                    _normalize_key(cfg_ws3.cell(row=r, column=from_col).value),
                )
                if key[0] and key[1]:
                    existing_alias_rows[key] = r
            for state_val, from_name, to_name in alias_saves:
                key = (_normalize_key(state_val), _normalize_key(from_name))
                _alias_key_str = f"State={state_val} | From={from_name}"
                if key in existing_alias_rows:
                    _old_to_name = cfg_ws3.cell(row=existing_alias_rows[key], column=to_col).value
                    cfg_ws3.cell(row=existing_alias_rows[key], column=to_col, value=to_name)
                    _audit_entries.append(_audit_row(
                        "Alias_Map", _alias_key_str,
                        _old_to_name if _old_to_name not in (None, "") else "(was blank)",
                        to_name, TicketId,
                    ))
                else:
                    new_row = cfg_ws3.max_row + 1
                    cfg_ws3.cell(row=new_row, column=state_col, value=state_val)
                    cfg_ws3.cell(row=new_row, column=from_col, value=from_name)
                    cfg_ws3.cell(row=new_row, column=to_col, value=to_name)
                    existing_alias_rows[key] = new_row
                    _audit_entries.append(_audit_row("Alias_Map", _alias_key_str, "(new record)", to_name, TicketId))
            print(f"\n✅ Saved {len(alias_saves)} alias mapping update(s) to Alias_Map.")

        if lien_equiv_saves:
            cfg_ws4, header4 = _config_sheet(cfg_wb, "Lien_Name_Equivalence_Map",
                                             ["State MCO", "CollectorId", "LienHolderId", "LienHolderName",
                                              "EquivalentName", "Decision", "Ticket #"])
            col_idx4 = {name: i + 1 for i, name in enumerate(header4) if name}
            # (plan, name) -> (row, decision); a later row wins, same as the loader.
            existing_equiv_rows = {}
            for r in range(2, cfg_ws4.max_row + 1):
                key = (
                    _normalize_key(cfg_ws4.cell(row=r, column=col_idx4["State MCO"]).value),
                    _normalize_key(cfg_ws4.cell(row=r, column=col_idx4["CollectorId"]).value),
                    _normalize_key(cfg_ws4.cell(row=r, column=col_idx4["LienHolderId"]).value),
                    _normalize_key(cfg_ws4.cell(row=r, column=col_idx4["EquivalentName"]).value),
                )
                if key[0] and key[1] and key[2] and key[3]:
                    existing_equiv_rows[key] = (r, _normalize_key(cfg_ws4.cell(row=r, column=col_idx4["Decision"]).value))
            _new_equiv_rows = 0
            for state_mco, collector_id, lienholder_id, lienholder_name, equiv_name, decision, ticket_id in lien_equiv_saves:
                key = (_normalize_key(state_mco), _normalize_key(collector_id), _normalize_key(lienholder_id),
                       _normalize_key(equiv_name))
                _existing = existing_equiv_rows.get(key)
                if _existing and _existing[1] == _normalize_key(decision):
                    continue
                if _existing:
                    # A newer decision for the same name replaces the old one (e.g. NO -> YES).
                    _old_decision = cfg_ws4.cell(row=_existing[0], column=col_idx4["Decision"]).value
                    cfg_ws4.cell(row=_existing[0], column=col_idx4["Decision"], value=decision)
                    cfg_ws4.cell(row=_existing[0], column=col_idx4["LienHolderName"], value=lienholder_name)
                    cfg_ws4.cell(row=_existing[0], column=col_idx4["Ticket #"], value=ticket_id)
                    existing_equiv_rows[key] = (_existing[0], _normalize_key(decision))
                    _new_equiv_rows += 1
                    _audit_entries.append(_audit_row(
                        "Lien_Name_Equivalence_Map",
                        f"State={state_mco} | CollectorId={collector_id} | LienHolderId={lienholder_id} | Candidate={equiv_name}",
                        f"Decision={_old_decision}", f"Decision={decision}", ticket_id,
                    ))
                    continue
                new_row = cfg_ws4.max_row + 1
                cfg_ws4.cell(row=new_row, column=col_idx4["State MCO"], value=state_mco)
                cfg_ws4.cell(row=new_row, column=col_idx4["CollectorId"], value=collector_id)
                cfg_ws4.cell(row=new_row, column=col_idx4["LienHolderId"], value=lienholder_id)
                cfg_ws4.cell(row=new_row, column=col_idx4["LienHolderName"], value=lienholder_name)
                cfg_ws4.cell(row=new_row, column=col_idx4["EquivalentName"], value=equiv_name)
                cfg_ws4.cell(row=new_row, column=col_idx4["Decision"], value=decision)
                cfg_ws4.cell(row=new_row, column=col_idx4["Ticket #"], value=ticket_id)
                existing_equiv_rows[key] = (new_row, _normalize_key(decision))
                _new_equiv_rows += 1
                # Brand-new (plan, name) pair -- a repeat or a changed decision was
                # handled above, so this audit row is always a new record.
                _audit_entries.append(_audit_row(
                    "Lien_Name_Equivalence_Map",
                    f"State={state_mco} | CollectorId={collector_id} | LienHolderId={lienholder_id} | Candidate={equiv_name}",
                    "(new record)", f"Decision={decision}", ticket_id,
                ))
            print(f"\n✅ Saved {_new_equiv_rows} lien-name equivalence update(s) to Lien_Name_Equivalence_Map.")

        if column_map_saves:
            cfg_ws5, header5 = _config_sheet(cfg_wb, "Column_Map", ["State", "LienId_Column", "MCO_Columns", "Ticket #"])
            col_idx5 = {name: i + 1 for i, name in enumerate(header5) if name}
            existing_col_map_rows = set()
            for r in range(2, cfg_ws5.max_row + 1):
                key = (
                    _normalize_key(cfg_ws5.cell(row=r, column=col_idx5["State"]).value),
                    _normalize_key(cfg_ws5.cell(row=r, column=col_idx5["LienId_Column"]).value),
                    _normalize_key(cfg_ws5.cell(row=r, column=col_idx5["MCO_Columns"]).value),
                )
                if key[0] and key[1]:
                    existing_col_map_rows.add(key)
            _new_col_map_rows = 0
            for state_val, lienid_col, mco_cols, ticket_id in column_map_saves:
                mco_cols_str = ", ".join(str(c) for c in mco_cols)
                key = (_normalize_key(state_val), _normalize_key(lienid_col), _normalize_key(mco_cols_str))
                if key in existing_col_map_rows:
                    continue
                new_row = cfg_ws5.max_row + 1
                cfg_ws5.cell(row=new_row, column=col_idx5["State"], value=state_val)
                cfg_ws5.cell(row=new_row, column=col_idx5["LienId_Column"], value=lienid_col)
                cfg_ws5.cell(row=new_row, column=col_idx5["MCO_Columns"], value=mco_cols_str)
                cfg_ws5.cell(row=new_row, column=col_idx5["Ticket #"], value=ticket_id)
                existing_col_map_rows.add(key)
                _new_col_map_rows += 1
                _audit_entries.append(_audit_row(
                    "Column_Map", f"State={state_val} | LienId_Column={lienid_col}",
                    "(new record)", f"MCO_Columns={mco_cols_str}", ticket_id,
                ))
            print(f"\n✅ Saved {_new_col_map_rows} column mapping update(s) to Column_Map.")

        if _audit_entries:
            audit_ws = _audit_trail_ws(cfg_wb)
            for _entry in _audit_entries:
                _new_audit_row = audit_ws.max_row + 1
                for _i, _col_name in enumerate(_AUDIT_TRAIL_COLUMNS, start=1):
                    audit_ws.cell(row=_new_audit_row, column=_i, value=_entry.get(_col_name))

        if _fingerprint_really_changed(_loaded_fingerprint, _file_fingerprint(config_path)):
            print(
                f"\n⚠️  '{config_path}' was changed by someone else while this run was "
                f"working (likely synced in from another user's save in the meantime). "
                f"Nothing was saved here, to avoid silently overwriting their changes -- "
                f"please re-run this script once their update has synced, so your "
                f"changes apply on top of the latest version instead."
            )
            if _audit_entries:
                print(f"  Audit_Trail: {len(_audit_entries)} entry(ies) also NOT recorded -- they'll be re-queued if you re-run.")
            return

        try:
            _retry_on_permission_error(lambda: cfg_wb.save(config_path), config_path)
            _AUDIT_TRAIL_THIS_RUN.extend(_audit_entries)
        except PermissionError:
            print(_config_save_fallback_msg)
            if _audit_entries:
                print(f"  Audit_Trail: {len(_audit_entries)} entry(ies) also NOT recorded -- retry once the file is closed.")

    # =====================================================
    # FINAL COLUMN ORDER
    # =====================================================

    # Placed here, not in Step 4, so the SAME cache is shared -- Step 4's Scenario C
    # reuses whatever was already looked up here instead of re-querying the same
    # CollectorId.
    _assigned_user_cache = {}

    def get_assigned_user_id(collector_id):
        # A blank/NaN CollectorId (normally caught by the Pending_TPA_Blank_Fields
        # check, guarded here too as a backstop) would otherwise interpolate directly
        # into the SQL text, producing invalid syntax instead of a clean failure. Also
        # parameterized -- the old f-string interpolation was a SQL injection risk
        # regardless.
        if pd.isna(collector_id) or str(collector_id).strip() == "":
            return None
        # Defensive normalization at this SQL boundary: pyodbc rejects a numpy
        # scalar (e.g. numpy.int64, if an upstream ID column ever ends up numeric)
        # outright as a bound parameter ("Invalid parameter type") -- casting to a
        # plain string here makes this boundary safe regardless of what type flows
        # in from upstream. SQL Server compares a string parameter against an int
        # column fine via implicit conversion.
        collector_id = _id_key(collector_id)
        # Text that is not an ID (e.g. "TBD" in the TPA file) is never sent to SQL.
        if not collector_id or not re.fullmatch(r"\d+", collector_id):
            return None
        if collector_id not in _assigned_user_cache:
            q = "select Id, defaultassigneduserid from collectors where Id = ?"
            result = _read_sql(q, params=[collector_id])
            _assigned_user_cache[collector_id] = _id_key(result["defaultassigneduserid"].iloc[0]) if not result.empty else None
        return _assigned_user_cache[collector_id]

    # State was missing from this selection entirely (a real bug: Step 4's NewLienNote/
    # NewClientNote templates read it via r.get("State")/df_in["State"], which silently
    # produced "None" in the note text since the column had already been dropped by the
    # time Step 4 ran) -- kept here so Step 4 still has it, but (per explicit request)
    # dropped again from the DISPLAYED Yes_with_MCO copy at write time, same as LienType.
    base_cols       = ["ClientId", "LienId", "Project_Name", "State", "Last_Name", "First_Name", "MCO_Name", "PrivateLiensScope", "clientadditionalinformation", "Scope_Classification"]
    existing_cols   = [c for c in yes_with_mco_df.columns if c.startswith("MCO") and "Existing" in c]
    date_cols_final = [c for c in ["Start_Date", "End_Date"] if c in yes_with_mco_df.columns]

    # LienType is kept here (unlike CollectorName, which really is unused everywhere and
    # stays dropped) even though it isn't shown on the Yes_with_MCO tab -- Step 4's
    # Scenario C needs the TPA-resolved LienType for this exact plan, to populate Client
    # Scope Upload/MCO Update 1 correctly instead of re-deriving it from client scope
    # (a separate bug -- see build_step4_outputs). It's dropped only from the DISPLAYED
    # copy right before the Yes_with_MCO sheet itself gets written.
    # Observation (parse notes so far) and _OpenEnd carry on into Step 4.
    carry_cols      = [c for c in ["Observation", "_OpenEnd", "_TPA_Source"] if c in yes_with_mco_df.columns]
    yes_with_mco_df = yes_with_mco_df[base_cols + date_cols_final + existing_cols + ["LienholderId", "CollectorId", "LienType"]
                                      + carry_cols]

    # Pre-fetch every distinct CollectorId this run needs in ONE round trip -- without
    # this, get_assigned_user_id's own per-id cache still means one separate SQL round
    # trip per DISTINCT CollectorId the .apply() below hits (real, measured cost: ~4s
    # for just a handful of distinct collectors in one run).
    _distinct_collector_ids = [
        c for c in dict.fromkeys(_id_key(v) for v in yes_with_mco_df["CollectorId"].dropna())
        if c and re.fullmatch(r"\d+", c)
    ]
    if _distinct_collector_ids:
        _bulk_collectors_df = _read_sql(
            f"select Id, defaultassigneduserid from collectors where Id in {make_sql_in(_distinct_collector_ids)}"
        )
        for _, _bc_row in _bulk_collectors_df.iterrows():
            # Plain ID text: a NULL in the column makes pandas read it as float ("77.0").
            _assigned_user_cache[_stringify_id(_bc_row["Id"])] = _id_key(_bc_row["defaultassigneduserid"])
        for _cid in _distinct_collector_ids:
            _assigned_user_cache.setdefault(_cid, None)

    yes_with_mco_df.insert(
        yes_with_mco_df.columns.get_loc("CollectorId") + 1, "AssignedUserId",
        yes_with_mco_df["CollectorId"].apply(get_assigned_user_id)
    )

    yes_with_mco_df.insert(0, "Unique", yes_with_mco_df["ClientId"].astype(str) + "_" + yes_with_mco_df["LienholderId"].astype(str))


    # =====================================================
    # STEP 4 — COMPARE AGAINST SQL-LIENS (Scenario A / B1 / B2 / C)
    # =====================================================

    PLACEHOLDER_PAIRS = {
        ("227", "214"), ("2498", "1608"), ("1866", "2473"), ("1867", "2474"),
        ("1868", "2475"), ("117", "2027"), ("1625", "2651"), ("124", "1608"),
    }

    def _step4_plan_names(state, resolved_collector_id, resolved_lienholder_id, resolved_name, lpm_names):
        """Every name our plan is known by, de-duplicated: the TPA/override name, the
        human-confirmed YES names (LIEN_NAME_EQUIVALENCE_MAP) and LPM's own name(s) for
        our LienholderId. A lien's LPM name comes from its LienholderId, so the LPM name
        always finds a lien carrying our LienholderId even when the TPA name is a short
        form ("CHPW"). Not used for a placeholder pair (its LPM name is generic)."""
        equiv_key = (str(state), str(resolved_collector_id), str(resolved_lienholder_id))
        names = [resolved_name] + LIEN_NAME_EQUIVALENCE_MAP.get(equiv_key, [])
        if (_stringify_id(resolved_collector_id), _stringify_id(resolved_lienholder_id)) not in PLACEHOLDER_PAIRS:
            names += list(lpm_names or [])
        unique_names, seen = [], set()
        for n in names:
            key = _normalize_key(n)
            if key and key not in seen:
                seen.add(key)
                unique_names.append(n)
        return unique_names

    def _lien_marked(df, marks):
        """Rows of df that a review answer points at (("ID", LienId) / ("NAME", name))."""
        if not marks or df.empty:
            return pd.Series(False, index=df.index)
        ids = {m[1] for m in marks if m and m[0] == "ID"}
        names = {m[1] for m in marks if m and m[0] == "NAME"}
        return df["LienId"].apply(_id_key).isin(ids) | df["LienHolderName"].apply(_normalize_key).isin(names)

    def _top_name_matches(names, df):
        """LienIds that _match_name rates best for any of names. Liens carrying the exact
        same name as a winner are added too -- _match_name itself returns only one of them."""
        top = set()
        for name in names:
            matched_id, tied_ids = _match_name(name, df, "LienHolderName", "LienId", return_tie=True)
            ids = [matched_id] if matched_id is not None else list(tied_ids)
            if not ids:
                continue
            won = {_normalize_key(n) for n in df.loc[df["LienId"].isin(ids), "LienHolderName"]} - {None}
            top |= set(ids) | set(df.loc[df["LienHolderName"].apply(_normalize_key).isin(won), "LienId"])
        return top

    def _strict_name_matches(names, df):
        """LienIds whose name passes _match_name on its own for any of names, even if
        another lien scores higher (a "weaker" but still sure match)."""
        hits = set()
        for name in names:
            for i in range(len(df)):
                one = df.iloc[[i]]
                if _match_name(name, one, "LienHolderName", "LienId") is not None:
                    hits.add(one["LienId"].iloc[0])
        return hits

    def _name_share_pct(name_a, name_b):
        """Both-sides word share of two names: common words / all distinct words (state
        names and MCO/CCO ignored). Only breaks ties between equally good matches."""
        def _words(n):
            if n is None or (not isinstance(n, str) and pd.isna(n)):
                return set()
            return {_singularize(w) for w in _tokenize_plan(str(n).strip().lower()) if w not in _STRUCTURAL_TAGS}
        wa, wb = _words(name_a), _words(name_b)
        if not wa or not wb:
            return 0.0
        return len(wa & wb) / len(wa | wb)

    def _mco_mask(df):
        return df["IsMCO"].apply(_is_true_flag) if "IsMCO" in df.columns else pd.Series(False, index=df.index)

    def _not_fne_mask(df):
        return df["Stage"] != "Final No Entitlement" if "Stage" in df.columns else pd.Series(True, index=df.index)

    def _narrow_step4(cands, steps):
        """Applies each step's filter in order; a step keeps only the liens it prefers
        (or all of them when it prefers none). Returns the one lien left as a Series, or
        the DataFrame of liens still tied after every step."""
        remaining = cands
        for step in steps:
            if len(remaining) == 1:
                break
            keep = remaining[step(remaining)]
            if not keep.empty:
                remaining = keep
        return remaining.iloc[0] if len(remaining) == 1 else remaining

    def _classify_step4_detailed(client_liens_df, state, resolved_collector_id, resolved_lienholder_id, resolved_name,
                                 lpm_names=None, picks=None, client_excluded=None):
        """Step 4 for one plan. client_liens_df: this client's active (statusid==1),
        not-yet-claimed liens. Returns a dict:
          scenario   'A' (update a matched lien), 'B1'/'B2' (fill a placeholder,
                     Final No Entitlement or not), 'C' (create a new lien), 'T' (still
                     tied after every rule -> review) or 'R' (only a possible match ->
                     review)
          row        the chosen lien (A/B1/B2)
          candidates the liens for the review sheet (T/R), reason = why
          others     other sure matches / placeholders left untouched (for Observations)
          weaker_mco liens matched less surely but already IsMCO True (Observation)
          names      every name our plan is known by (_step4_plan_names)

        Sure matches come from every plan name, so a lien with our LienholderId always
        joins in. With 2+ of them each rule keeps only the liens it prefers, in this
        order: IsMCO True, our LienholderId, the highest both-sides name share, this
        client's own review answer, a name saved as YES (LIEN_NAME_EQUIVALENCE_MAP),
        not on Final No Entitlement. Still more than one -> review. Placeholders (only
        when nothing matches by name or possibly) go: IsMCO True, review answer, not
        Final No Entitlement, else review. Names saved as NO, and liens this client's
        review answered NO for, are dropped first."""
        names = _step4_plan_names(state, resolved_collector_id, resolved_lienholder_id, resolved_name, lpm_names)
        result = {"scenario": "C", "row": None, "candidates": None, "reason": None,
                  "others": None, "weaker_mco": None, "names": names, "nameless": set()}
        if client_liens_df.empty:
            return result
        # Compared as plain ID text ("300.0" from a NULL-holding SQL column == "300").
        collector_matches = client_liens_df[client_liens_df["CollectorId"].map(_id_key) == _id_key(resolved_collector_id)]
        equiv_key = (str(state), str(resolved_collector_id), str(resolved_lienholder_id))

        # Names confirmed NOT this plan (saved NO) are removed for every client; a NO
        # from this client's own review removes just that lien for this client.
        excluded_names = {_normalize_key(n) for n in LIEN_NAME_EXCLUDED_MAP.get(equiv_key, [])}
        if excluded_names and not collector_matches.empty:
            collector_matches = collector_matches[
                ~collector_matches["LienHolderName"].apply(_normalize_key).isin(excluded_names)
            ]
        if client_excluded and not collector_matches.empty:
            collector_matches = collector_matches[~_lien_marked(collector_matches, client_excluded)]
        if collector_matches.empty:
            return result

        is_placeholder = collector_matches.apply(
            lambda r: (_stringify_id(r["CollectorId"]), _stringify_id(r["LienholderId"])) in PLACEHOLDER_PAIRS, axis=1
        )
        picked = _lien_marked(collector_matches, picks)
        top = _top_name_matches(names, collector_matches)
        top |= set(collector_matches.loc[picked & ~is_placeholder, "LienId"])
        # A lien with no lienholder name in LPM but our own CollectorId + LienholderId is
        # this plan (else a duplicate lien would be created).
        nameless = (collector_matches["LienHolderName"].map(_normalize_key).isna()
                    & (collector_matches["LienholderId"].map(_id_key) == _id_key(resolved_lienholder_id))
                    & ~is_placeholder)
        result["nameless"] = set(collector_matches.loc[nameless, "LienId"])
        top |= result["nameless"]

        if top:
            cands = collector_matches[collector_matches["LienId"].isin(top)]
            ours = _stringify_id(resolved_lienholder_id)
            pct_names = [resolved_name] + list(lpm_names or [])
            saved_yes = {_normalize_key(n) for n in LIEN_NAME_EQUIVALENCE_MAP.get(equiv_key, [])}

            def _best_share(df):
                share = df["LienHolderName"].apply(lambda n: max(_name_share_pct(n, p) for p in pct_names))
                return share == share.max() if share.max() > 0 else pd.Series(False, index=df.index)

            chosen = _narrow_step4(cands, [
                _mco_mask,
                lambda df: df["LienholderId"].apply(_stringify_id) == ours,
                _best_share,
                lambda df: _lien_marked(df, picks),
                lambda df: df["LienHolderName"].apply(_normalize_key).isin(saved_yes),
                _not_fne_mask,
            ])
            if isinstance(chosen, pd.DataFrame):
                same_name = chosen["LienHolderName"].apply(_normalize_key).nunique(dropna=False) == 1
                result.update(scenario="T", candidates=chosen,
                              reason="Duplicate liens" if same_name else "Equal match")
                return result
            weaker = _strict_name_matches(names, collector_matches) - top
            weaker_mco = collector_matches[collector_matches["LienId"].isin(weaker) & _mco_mask(collector_matches)]
            result.update(scenario="A", row=chosen, others=cands[cands["LienId"] != chosen["LienId"]],
                          weaker_mco=weaker_mco)
            return result

        # A YES this client gave to a placeholder (even in a "possible match" review)
        # means "fill this placeholder" -- skip the possible-match check, or it loops.
        if not (picked & is_placeholder).any():
            overlap_candidates = _find_overlap_candidates(resolved_name, collector_matches, "LienHolderName")
            if overlap_candidates.empty and lpm_names and (
                    (_stringify_id(resolved_collector_id), _stringify_id(resolved_lienholder_id)) not in PLACEHOLDER_PAIRS):
                overlap_candidates = pd.concat(
                    [_find_overlap_candidates(n, collector_matches, "LienHolderName") for n in lpm_names]
                ).drop_duplicates(subset=["LienId"])
            if not overlap_candidates.empty:
                result.update(scenario="R", candidates=overlap_candidates, reason="Possible match")
                return result

        placeholders = collector_matches[is_placeholder]
        if placeholders.empty:
            return result
        chosen = _narrow_step4(placeholders, [_mco_mask, lambda df: _lien_marked(df, picks), _not_fne_mask])
        if isinstance(chosen, pd.DataFrame):
            result.update(scenario="T", candidates=chosen, reason="Duplicate placeholders")
            return result
        result.update(scenario="B1" if chosen["Stage"] == "Final No Entitlement" else "B2", row=chosen,
                      others=placeholders[placeholders["LienId"] != chosen["LienId"]])
        return result

    def classify_step4_scenario(client_liens_df, state, resolved_collector_id, resolved_lienholder_id, resolved_name,
                                lpm_names=None, picks=None, client_excluded=None):
        """(scenario, chosen lien Series for A/B1/B2 | candidates DataFrame for T/R |
        None for C) -- see _classify_step4_detailed. C is the plain catch-all: a lien
        under our CollectorId with another real plan's name is left alone and ours is
        created. R ("possible match", e.g. "Healthy Blue/Anthem BCBS" vs on-file
        "Wellpoint, Inc. (Anthem BCBS/Elevance)") is checked before placeholders, so a
        real overlap is never masked by a placeholder."""
        detail = _classify_step4_detailed(client_liens_df, state, resolved_collector_id, resolved_lienholder_id,
                                          resolved_name, lpm_names=lpm_names, picks=picks,
                                          client_excluded=client_excluded)
        if detail["scenario"] in ("A", "B1", "B2"):
            return detail["scenario"], detail["row"]
        return detail["scenario"], detail["candidates"]

    def build_note(template_key, ticket_id, plan_name, state=None, response_type=None, response_date=None, has_dates=False):
        """A/B1/B2 only mention "Enrollment Dates" when has_dates is actually True --
        the note must describe exactly what got updated, not more."""
        dated_templates = {
            "A":  "Updated IsMCO and Enrollment Dates per instruction on ticket #{ticket}. Plan Name: {plan}",
            "B1": "Updated lienholder Information, stage from Final No Entitlement, OnBenefit details and Enrollment Dates per instruction on ticket #{ticket}. Plan Name: {plan}",
            "B2": "Updated LH, IsMCO and Enrollment Dates per instruction on ticket #{ticket}. Plan Name: {plan}",
            # B2_MCO: the placeholder lien was already IsMCO True (IsMCO not re-sent).
            "B2_MCO": "Updated LH and Enrollment Dates per instruction on ticket #{ticket}. Plan Name: {plan}",
            # A name-matched lien on Final No Entitlement (own LienholderId kept); _MCO = IsMCO was already True.
            "A_FNE":     "Updated IsMCO, stage from Final No Entitlement, OnBenefit details and Enrollment Dates per instruction on ticket #{ticket}. Plan Name: {plan}",
            "A_FNE_MCO": "Updated stage from Final No Entitlement, OnBenefit details and Enrollment Dates per instruction on ticket #{ticket}. Plan Name: {plan}",
        }
        undated_templates = {
            "A":  "Updated IsMCO per instruction on ticket #{ticket}. Plan Name: {plan}",
            "B1": "Updated lienholder Information, stage from Final No Entitlement, OnBenefit details per instruction on ticket #{ticket}. Plan Name: {plan}",
            "B2": "Updated LH and IsMCO per instruction on ticket #{ticket}. Plan Name: {plan}",
            "B2_MCO": "Updated LH per instruction on ticket #{ticket}. Plan Name: {plan}",
            "A_FNE":     "Updated IsMCO, stage from Final No Entitlement, OnBenefit details per instruction on ticket #{ticket}. Plan Name: {plan}",
            "A_FNE_MCO": "Updated stage from Final No Entitlement, OnBenefit details per instruction on ticket #{ticket}. Plan Name: {plan}",
        }
        other_templates = {
            "A_dates":  "Updated Enrollment Dates per instruction on ticket #{ticket}. Plan Name: {plan}",
            "C_lien":   "Updated MCO count per {state} Mcaid {resp_type} response rec\'d {resp_date} per ticket #{ticket}. Plan Name: {plan}",
            "C_client": "Updated MCO count per {state} Mcaid {resp_type} response rec\'d {resp_date} per ticket #{ticket}",
        }
        if template_key in other_templates:
            template = other_templates[template_key]
        else:
            template = dated_templates[template_key] if has_dates else undated_templates[template_key]
        return template.format(
            ticket=ticket_id, plan=plan_name, state=state, resp_type=response_type, resp_date=response_date
        )

    A_COLS  = ["Id", "IsMCO", "Enrollmentstartdate", "EnrollmentEndDate", "NewLienNote"]
    B1_COLS = ["Id", "Stage", "Enrollmentstartdate", "EnrollmentEndDate", "LienholderId", "IsMCO",
               "Onbenefits", "OnBenefitsVerified", "LienDate1", "LienAmount1", "AuditedAmount1",
               "FinalDemandDate", "FinalDemandAmount", "NewLienNote"]
    B2_COLS = ["Id", "LienholderId", "IsMCO", "Enrollmentstartdate", "EnrollmentEndDate", "NewLienNote"]
    SCOPE_UPLOAD_COLS = ["ClientId", "LienType", "SourceLienHolderId", "SourceId"]
    # "Unique"/"ClientId" lead MCO Update 1 (ahead of "Id", which starts blank) so the
    # user can look up each new lien's real Id -- once Client Scope Upload's rows are
    # uploaded and the resulting liens exist on file -- via the "Fetch New Lien IDs"
    # query sheet, matching back on "Unique" (ClientId_LienholderId, same formula as
    # Yes_with_MCO's own Unique column) rather than needing a separate manual lookup.
    UPDATE1_COLS = ["Unique", "ClientId", "Id", "AssignedUserId", "Stage", "CollectorId", "LienholderId",
                    "OnBenefits", "OnBenefitsVerified", "Enrollmentstartdate", "EnrollmentEndDate",
                    "IsMCO", "NewLienNote"]
    # Real bug found: "Order by 5,2" was copied from sql_liens_query, where column 5
    # is LienId -- but in THIS template's own column list, column 5 is LienType (Id
    # is column 3), so it was sorting by LienType then ClientId instead of by Id.
    _FETCH_NEW_LIEN_IDS_QUERY_TEMPLATE = """select ClientSSN, ClientId, Id, Stage, LienType, CollectorId, LienholderId, AssignedUserId, statusid, LienHolderName, finalizedstatusname
from fullproductviews
where ClientId in {client_ids}
Order by 3,2"""
    # Ready-to-run queries, same "query text only" pattern as Fetch New Lien IDs
    # above (no dataframe, nothing auto-run) -- for the user's own convenience after
    # THIS run's Client Scope Upload/MCO Update 1/2/Client-MCO Cnt Update rows have
    # been manually uploaded to the system: paste the query into SSMS, run it
    # whenever ready, and paste the results into that tab. Scoped to yes_with_mco_df
    # (this run's own resolved ClientIds only, never a cross-round merge), so each
    # resume round gets its own query reflecting just what IT resolved.
    _SQL_POST_CREATION_QUERY_TEMPLATE = """select Id, ClientId, Lienholdername, Stage, CollectorName, Lientype, RequiredMedicaidMCOLienCount, IsMCO, CreatedOn
from fullproductviews
where clientid in {client_ids}"""
    _SQL_POST_MCO_CNT_QUERY_TEMPLATE = """select Distinct ClientId, RequiredMedicaidMCOLienCount
from fullproductviews
where clientid in {client_ids}"""
    PENDING_LIEN_MATCH_COLS = [
        "ClientId", "LienId", "MCO_Name", "LienholderId", "CollectorId", "LienType",
        "Scope_Classification", "State", "Start_Date", "End_Date",
        "TicketId", "ResponseDate", "ResponseType",
        "CandidateLienId", "CandidateLienHolderName", "CandidateCollectorId", "CandidateLienholderId",
        "CandidateIsMCO", "CandidateStage", "CandidateLienType", "ReviewReason",
        "Decision", "Observation", "_OpenEnd",
    ]

    def _fmt_enroll_date(v):
        """'MM/DD/YYYY' for a DB date/datetime or date text ('MM/DD/YYYY' first, then ISO
        like '2015-01-01'); other non-blank text (e.g. "Bad Date") is returned as-is;
        blank/NaN/NaT, or their text forms, -> None."""
        if v is None:
            return None
        if isinstance(v, str):
            v = v.strip()
            if not v or v.lower() in ("nan", "nat", "none"):
                return None
            parsed = _parse_date(v)
            if isinstance(parsed, str):
                parsed = pd.to_datetime(v, errors="coerce")
                if pd.isna(parsed):
                    return v
            return parsed.strftime("%m/%d/%Y")
        try:
            if pd.isna(v):
                return None
            return pd.Timestamp(v).strftime("%m/%d/%Y")
        except (TypeError, ValueError):
            return str(v)

    def _enrollment_date_check(our_start, our_end, lien_start, lien_end):
        """Returns (write_dates, observation). Dates are written only when one of ours
        exists and differs from the lien's on-file value; replacing a DIFFERENT on-file
        date (not just filling a blank one) also returns an Observation."""
        ours = (_fmt_enroll_date(our_start), _fmt_enroll_date(our_end))
        theirs = (_fmt_enroll_date(lien_start), _fmt_enroll_date(lien_end))
        changed = [o is not None and o != t for o, t in zip(ours, theirs)]
        if not any(changed):
            return False, None
        if not any(c and t is not None for c, t in zip(changed, theirs)):
            return True, None
        def _span(pair):
            return f"{pair[0] or 'blank'} - {pair[1] or 'blank'}"
        # A side we don't send stays as it is on the lien (blank = no change in the upload).
        after = tuple(o if o is not None else t for o, t in zip(ours, theirs))
        return True, f"Enrollment dates changed on lien: {_span(theirs)} -> {_span(after)}"

    def _lien_label(row):
        """'lien 501 (LH 4410)' / 'lien 501 (LH 4410, already MCO)' for Observations."""
        mco = ", already MCO" if _is_true_flag(row.get("IsMCO")) else ""
        return f"lien {row['LienId']} (LH {_stringify_id(row.get('LienholderId'))}{mco})"

    def _step4_observations(detail, our_lienholder_id, expected_lien_type, date_obs=None):
        """Observation texts for a chosen lien (A/B1/B2), in order: which lien was used
        when 2+ matched, the enrollment-date change, LienType differing from ours, and a
        less-sure match already on MCO."""
        row, others, obs = detail["row"], detail["others"], []
        if others is not None and not others.empty:
            if detail["scenario"] == "A":
                parts = []
                if _stringify_id(row.get("LienholderId")) != _stringify_id(our_lienholder_id):
                    parts.append(f"Expected LH {_stringify_id(our_lienholder_id)}, used {_lien_label(row)}.")
                parts.append("Also matches: " + ", ".join(_lien_label(o) for _, o in others.iterrows())
                             + " - possible duplicate")
                obs.append(" ".join(parts))
            else:
                obs.append("Also a placeholder: " + ", ".join(f"lien {o['LienId']}" for _, o in others.iterrows())
                           + " - left unchanged")
        if date_obs:
            obs.append(date_obs)
        theirs = row.get("LienType")
        if _normalize_key(theirs) and _normalize_key(expected_lien_type):
            if (_canonical_lien_type(theirs) or _normalize_key(theirs)) != (
                    _canonical_lien_type(expected_lien_type) or _normalize_key(expected_lien_type)):
                obs.append(f"LienType on lien {row['LienId']} is {theirs}; expected {expected_lien_type}")
        weaker = detail["weaker_mco"]
        if weaker is not None and not weaker.empty:
            weaker = weaker[weaker["LienId"] != row["LienId"]]
            if not weaker.empty:
                obs.append("Also MCO: " + ", ".join(
                    f"lien {w['LienId']} (LH {_stringify_id(w.get('LienholderId'))})" for _, w in weaker.iterrows()
                ) + ", weaker name match - check")
        return obs

    def _other_collector_observation(client_liens, collector_id, names):
        """When a new lien is created: the same plan already sitting under another
        CollectorId for this client (sure name match, incl. our LienholderId's LPM name)."""
        others = client_liens[client_liens["CollectorId"].map(_id_key) != _id_key(collector_id)]
        if others.empty:
            return None
        hits = _strict_name_matches(names, others)
        texts = [
            f"Same plan already on lien {o['LienId']} (LH {_stringify_id(o['LienholderId'])}) under CollectorId "
            f"{_stringify_id(o['CollectorId'])} - new lien created under CollectorId {_stringify_id(collector_id)}"
            for _, o in others[others["LienId"].isin(hits)].iterrows()
        ]
        return " | ".join(texts) or None

    def _as_enroll_datetime(v):
        s = _fmt_enroll_date(v)
        try:
            return datetime.strptime(s, "%m/%d/%Y") if s else None
        except (TypeError, ValueError):
            return None

    def _same_plan_rows(df_in):
        """Same client + same plan (CollectorId + LienholderId) from 2+ input rows: the
        first row handles it once with the earliest start / latest end of all of them;
        returns {other row index: first row index}. Dates are written onto the first row."""
        groups = {}
        for idx, r in df_in.iterrows():
            key = (_id_key(r.get("ClientId")), _id_key(r.get("CollectorId")), _id_key(r.get("LienholderId")))
            if None not in key:
                groups.setdefault(key, []).append(idx)
        secondary = {}
        for idxs in groups.values():
            if len(idxs) < 2:
                continue
            primary = idxs[0]
            for col, pick in (("Start_Date", min), ("End_Date", max)):
                if col not in df_in.columns:
                    continue
                dated = [(d, df_in.at[i, col]) for i in idxs if (d := _as_enroll_datetime(df_in.at[i, col])) is not None]
                if dated:
                    df_in.at[primary, col] = pick(dated, key=lambda t: t[0])[1]
            # An end given as CURRENT on any of the rows keeps the end blank.
            if "_OpenEnd" in df_in.columns and "End_Date" in df_in.columns and any(
                    _is_true_flag(df_in.at[i, "_OpenEnd"]) for i in idxs):
                df_in.at[primary, "End_Date"] = None
            for other in idxs[1:]:
                secondary[other] = primary
        return secondary

    def build_step4_outputs(df_in, liens_df, ticket_id, response_date, response_type, already_claimed_lien_ids=None,
                            lien_match_picks=None, lien_match_excluded=None):
        if already_claimed_lien_ids is None:
            already_claimed_lien_ids = {}
        lien_match_picks = lien_match_picks or {}
        lien_match_excluded = lien_match_excluded or {}
        # Without these SQL-Liens columns the on-file dates are unknown, so dates are
        # written exactly as before (never compared, never added to an IsMCO-True lien).
        _lien_dates_known = {"EnrollmentStartDate", "EnrollmentEndDate"} <= set(liens_df.columns)
        # LPM's own name per LienholderId, from every lien in SQL-Liens (any client/status).
        _lpm_names_by_lh = {}
        if {"LienholderId", "LienHolderName"} <= set(liens_df.columns):
            for _lh, _lh_name in liens_df[["LienholderId", "LienHolderName"]].dropna().drop_duplicates().itertuples(index=False):
                _lpm_names_by_lh.setdefault(_stringify_id(_lh), []).append(_lh_name)
        # ...plus straight from LPM's LienHolders table, so the name is known even when no
        # lien in this run carries that LienholderId (lookup failure -> SQL-Liens names only).
        _lh_table_names = _lienholder_names_by_id(
            _stringify_id(v) for v in df_in["LienholderId"].dropna()
        ) if "LienholderId" in df_in.columns else {}
        for _lh, _lh_name in (_lh_table_names or {}).items():
            if _normalize_key(_lh_name) not in {_normalize_key(n) for n in _lpm_names_by_lh.get(_lh, [])}:
                _lpm_names_by_lh.setdefault(_lh, []).append(_lh_name)
        rows_a, rows_b1, rows_b2 = [], [], []
        rows_scope_upload, rows_update1, rows_pending_lien_match = [], [], []
        updated_client_ids = set()
        # This run's own newly-claimed LienIds, PER CLIENT -- carried forward (see
        # already_claimed_lien_ids above, and _ALREADY_CLAIMED_LIEN_IDS_CARRY at the
        # call site) onto any of this client's OTHER plans that go to Pending_* this
        # run, so a LATER resume round never matches the SAME already-claimed lien
        # again for a different plan.
        claimed_lien_ids_by_client = {}

        # A client can have MULTIPLE MCO plans in one run, and classify_step4_scenario
        # re-derives client_liens fresh from liens_df every iteration -- without
        # tracking LienIds an earlier plan (same loop, same client) already claimed, a
        # second plan could match the same lien row again instead of going to Scenario C.
        claimed_lien_ids = set()

        # 3 new Yes_with_MCO summary columns -- which existing LienId this plan touched
        # this run, its Stage, and a short description of what happened. Populated
        # per-row below in the same loop (df_in is the caller's own yes_with_mco_df,
        # passed by reference, so writes are visible to the caller). New columns land
        # at the END of the existing columns, since FINAL COLUMN ORDER already ran
        # before Step 4.
        df_in["LienIds used for updates"] = None
        df_in["Stage"] = None
        df_in["Remarks"] = None
        # Observations from reading the input text come first, Step 4's own after them.
        _parse_obs = (df_in["Observation"].copy() if "Observation" in df_in.columns
                      else pd.Series(None, index=df_in.index, dtype=object))
        df_in["Observation"] = None

        def _join_update_fields(fields):
            """"Update X" / "Update X & Y" / "Update X, Y & Z" -- canonical field
            order is always [Enrollment dates, FNI, LH, IsMCO], already filtered by
            the caller to just the ones actually applicable+present."""
            if len(fields) == 1:
                return f"Update {fields[0]}"
            return f"Update {', '.join(fields[:-1])} & {fields[-1]}"

        # Same client + same plan in 2+ input rows: only the first row acts (one lien
        # update or creation, never a duplicate); the others copy its outcome below.
        _same_plan_secondary = _same_plan_rows(df_in)
        # Plans that hit an unexpected error: row index -> the error (see the safety net below).
        error_plans = {}

        for idx, r in df_in.iterrows():
            if idx in _same_plan_secondary:
                continue
            # Safety net: an unexpected error on one plan undoes everything this plan
            # started (upload rows, claimed lien, count) and sends only it to
            # Pending_Error_Review; the client's other plans carry on as usual.
            _snap4 = (len(rows_a), len(rows_b1), len(rows_b2), len(rows_scope_upload), len(rows_update1),
                      len(rows_pending_lien_match), set(claimed_lien_ids),
                      {k: set(v) for k, v in claimed_lien_ids_by_client.items()}, set(updated_client_ids))
            try:
                lien_id        = r["LienId"]
                client_id      = r["ClientId"]
                plan_name      = r["MCO_Name"]
                state          = r.get("State")
                lienholder_id  = r["LienholderId"]
                collector_id   = r["CollectorId"]
                start_date     = r.get("Start_Date")
                end_date       = r.get("End_Date")
                has_dates      = pd.notna(start_date) or pd.notna(end_date)
                fate_key       = (lien_id, plan_name)

                # Excludes both this run's own in-loop claims AND any claim already
                # carried forward from an earlier round for this same client (see
                # already_claimed_lien_ids above) -- a plan this client already has
                # pending from a prior round never re-matches a lien that round already
                # claimed for a DIFFERENT plan.
                _already_claimed_for_client = already_claimed_lien_ids.get(client_id, set())
                client_liens = liens_df[(liens_df["ClientId"] == client_id) & (liens_df["statusid"] == 1)]
                client_liens = client_liens[~client_liens["LienId"].isin(claimed_lien_ids | _already_claimed_for_client)]
                _plan_key = (_id_key(client_id), _id_key(collector_id), _id_key(lienholder_id))
                detail = _classify_step4_detailed(
                    client_liens, state, collector_id, lienholder_id, plan_name,
                    lpm_names=_lpm_names_by_lh.get(_stringify_id(lienholder_id)),
                    picks=lien_match_picks.get(_plan_key), client_excluded=lien_match_excluded.get(_plan_key),
                )
                scenario = detail["scenario"]
                matched_row = detail["candidates"] if scenario in ("T", "R") else detail["row"]
                if scenario in ("A", "B1", "B2") and _lien_dates_known:
                    write_dates, date_obs = _enrollment_date_check(
                        start_date, end_date, matched_row.get("EnrollmentStartDate"), matched_row.get("EnrollmentEndDate")
                    )
                else:
                    write_dates, date_obs = has_dates, None
                # This row's Observations, joined with " | " at the end of the loop body.
                observations = []
                if scenario in ("A", "B1", "B2"):
                    observations = _step4_observations(detail, lienholder_id, r.get("LienType"), date_obs)
                if scenario == "A" and matched_row["LienId"] in detail.get("nameless", ()):
                    observations.append(f"LienId {matched_row['LienId']} has no lienholder name in LPM - "
                                        f"matched by LienholderId and CollectorId")

                if scenario == "A":
                    claimed_lien_ids.add(matched_row["LienId"])
                    claimed_lien_ids_by_client.setdefault(client_id, set()).add(matched_row["LienId"])
                    # A client's liens can already be correct on file while
                    # RequiredMedicaidMCOLienCount on the client record is still stuck at
                    # the generic starting value -- so this client must still be counted in
                    # Client-MCO Cnt Update even when nothing needs an MCO Update 2 row.
                    # Added unconditionally for BOTH branches below, not just one.
                    updated_client_ids.add(client_id)
                    df_in.at[idx, "LienIds used for updates"] = matched_row["LienId"]
                    df_in.at[idx, "Stage"] = matched_row["Stage"]
                    # pd.notna guard (future-proofing): a SQL NULL IsMCO can arrive here
                    # as pandas NaN, and bool(float("nan")) is True in Python -- so a
                    # plain "not matched_row['IsMCO']" would wrongly treat an unknown
                    # value the same as an already-True one, silently skipping a client
                    # who actually needs the IsMCO update.
                    _is_mco_true = _is_true_flag(matched_row["IsMCO"])
                    _is_fni = matched_row["Stage"] == "Final No Entitlement"
                    # A lien on Final No Entitlement moves to Final - No Interest only together
                    # with a real update (IsMCO not yet True, or dates differ); otherwise it
                    # stays as it is and only the count is updated.
                    if _is_fni and (not _is_mco_true or write_dates):
                        # Same move as a B1 placeholder (to Final - No Interest, OnBenefits,
                        # amounts 0), in the B1 block. LienholderId blank: the lien keeps its own;
                        # IsMCO blank when it is already True (never re-sent).
                        row = {
                            "Id": matched_row["LienId"], "Stage": "Final - No Interest",
                            "LienholderId": None, "IsMCO": None if _is_mco_true else True,
                            "Onbenefits": "Yes", "OnBenefitsVerified": response_date,
                            "LienDate1": response_date, "LienAmount1": 0, "AuditedAmount1": 0,
                            "FinalDemandDate": response_date, "FinalDemandAmount": 0,
                            "NewLienNote": build_note("A_FNE_MCO" if _is_mco_true else "A_FNE", ticket_id, plan_name,
                                                      has_dates=write_dates),
                        }
                        if write_dates:
                            row["Enrollmentstartdate"], row["EnrollmentEndDate"] = start_date, end_date
                        rows_b1.append(row)
                        PLAN_FATE[fate_key] = ("Matched existing lien (Final No Entitlement) moved to "
                                               "Final - No Interest -- see MCO Update 2")
                        df_in.at[idx, "Remarks"] = _join_update_fields(
                            (["Enrollment dates"] if write_dates else []) + ["FNI"] + ([] if _is_mco_true else ["IsMCO"])
                        )
                    elif _is_fni:
                        PLAN_FATE[fate_key] = "Matched existing lien already correctly on file; no update needed"
                        df_in.at[idx, "Remarks"] = "Already correct- Only update count"
                    elif not _is_mco_true:
                        row = {"Id": matched_row["LienId"], "IsMCO": True,
                               "NewLienNote": build_note("A", ticket_id, plan_name, has_dates=write_dates)}
                        if write_dates:
                            row["Enrollmentstartdate"], row["EnrollmentEndDate"] = start_date, end_date
                        rows_a.append(row)
                        PLAN_FATE[fate_key] = "Matched existing lien on file -- see MCO Update 2"
                        _a_fields = (["Enrollment dates"] if write_dates else []) + ["IsMCO"]
                        df_in.at[idx, "Remarks"] = _join_update_fields(_a_fields)
                    elif _lien_dates_known and write_dates:
                        # IsMCO already True: not re-sent (blank); only the dates change.
                        rows_a.append({
                            "Id": matched_row["LienId"], "IsMCO": None,
                            "Enrollmentstartdate": start_date, "EnrollmentEndDate": end_date,
                            "NewLienNote": build_note("A_dates", ticket_id, plan_name),
                        })
                        PLAN_FATE[fate_key] = "Matched existing lien on file -- see MCO Update 2"
                        df_in.at[idx, "Remarks"] = _join_update_fields(["Enrollment dates"])
                    else:
                        PLAN_FATE[fate_key] = "Matched existing lien already correctly on file; no update needed"
                        df_in.at[idx, "Remarks"] = "Already correct- Only update count"

                elif scenario == "B1":
                    claimed_lien_ids.add(matched_row["LienId"])
                    claimed_lien_ids_by_client.setdefault(client_id, set()).add(matched_row["LienId"])
                    df_in.at[idx, "LienIds used for updates"] = matched_row["LienId"]
                    df_in.at[idx, "Stage"] = matched_row["Stage"]
                    _ph_mco_true = _is_true_flag(matched_row.get("IsMCO"))
                    row = {
                        "Id": matched_row["LienId"], "Stage": "Final - No Interest",
                        "LienholderId": lienholder_id, "IsMCO": None if _ph_mco_true else True,
                        "Onbenefits": "Yes", "OnBenefitsVerified": response_date,
                        "LienDate1": response_date, "LienAmount1": 0, "AuditedAmount1": 0,
                        "FinalDemandDate": response_date, "FinalDemandAmount": 0,
                        "NewLienNote": build_note("B1", ticket_id, plan_name, has_dates=write_dates),
                    }
                    if write_dates:
                        row["Enrollmentstartdate"], row["EnrollmentEndDate"] = start_date, end_date
                    rows_b1.append(row)
                    updated_client_ids.add(client_id)
                    PLAN_FATE[fate_key] = "Placeholder lien (Final No Entitlement) filled in and finalized -- see MCO Update 2"
                    _b1_fields = (["Enrollment dates"] if write_dates else []) + ["FNI", "LH"] + ([] if _ph_mco_true else ["IsMCO"])
                    df_in.at[idx, "Remarks"] = _join_update_fields(_b1_fields)

                elif scenario == "B2":
                    claimed_lien_ids.add(matched_row["LienId"])
                    claimed_lien_ids_by_client.setdefault(client_id, set()).add(matched_row["LienId"])
                    df_in.at[idx, "LienIds used for updates"] = matched_row["LienId"]
                    df_in.at[idx, "Stage"] = matched_row["Stage"]
                    _ph_mco_true = _is_true_flag(matched_row.get("IsMCO"))
                    row = {"Id": matched_row["LienId"], "LienholderId": lienholder_id,
                           "IsMCO": None if _ph_mco_true else True,
                           "NewLienNote": build_note("B2_MCO" if _ph_mco_true else "B2", ticket_id, plan_name,
                                                     has_dates=write_dates)}
                    if write_dates:
                        row["Enrollmentstartdate"], row["EnrollmentEndDate"] = start_date, end_date
                    rows_b2.append(row)
                    updated_client_ids.add(client_id)
                    PLAN_FATE[fate_key] = "Placeholder lien filled in with lienholder details -- see MCO Update 2"
                    _b2_fields = (["Enrollment dates"] if write_dates else []) + ["LH"] + ([] if _ph_mco_true else ["IsMCO"])
                    df_in.at[idx, "Remarks"] = _join_update_fields(_b2_fields)

                elif scenario in ("T", "R"):
                    # T = 2+ of the client's own liens still equal after every selection
                    # rule; R = no sure match, but a lien under this CollectorId shares real
                    # name content with the plan (without this a duplicate lien got created
                    # -- real user-reported bug). Either way: refuse to guess. Nothing is
                    # claimed or counted for THIS plan; one row per candidate goes to
                    # Pending_Lien_Match_Review, with ReviewReason saying why.
                    for _, cand_row in matched_row.iterrows():
                        rows_pending_lien_match.append({
                            "ClientId": client_id, "LienId": lien_id, "MCO_Name": plan_name,
                            "LienholderId": lienholder_id, "CollectorId": collector_id,
                            "LienType": r["LienType"], "Scope_Classification": r.get("Scope_Classification"),
                            "State": state, "Start_Date": start_date, "End_Date": end_date,
                            "TicketId": ticket_id, "ResponseDate": response_date, "ResponseType": response_type,
                            "CandidateLienId": cand_row["LienId"], "CandidateLienHolderName": cand_row["LienHolderName"],
                            "CandidateCollectorId": cand_row["CollectorId"], "CandidateLienholderId": cand_row["LienholderId"],
                            "CandidateIsMCO": cand_row.get("IsMCO"), "CandidateStage": cand_row.get("Stage"),
                            "CandidateLienType": cand_row.get("LienType"), "ReviewReason": detail["reason"],
                            "Decision": None, "Observation": _parse_obs.get(idx), "_OpenEnd": r.get("_OpenEnd"),
                        })
                    if scenario == "T":
                        PLAN_FATE[fate_key] = "Ambiguous match with existing liens on file -- see Pending_Lien_Match_Review"
                    else:
                        PLAN_FATE[fate_key] = "Possible match with an existing lien on file could not be confirmed -- see Pending_Lien_Match_Review"
                    df_in.at[idx, "Remarks"] = "Pending Lien Match Review"

                else:
                    # LienType comes directly from the TPA-resolved value for THIS plan
                    # (r["LienType"], kept on yes_with_mco_df since Step 3), not re-derived
                    # from the client's own Scope_Classification -- TPA's own row already
                    # carries the correct LienType for this exact Tort/Name/Scope, and
                    # re-deriving it from scope can disagree with it.
                    resolved_lien_type = r["LienType"]
                    rows_scope_upload.append({
                        "ClientId": client_id, "LienType": resolved_lien_type,
                        "SourceLienHolderId": lienholder_id, "SourceId": 5,
                    })
                    # Stage must agree with the LienType actually written above -- deriving
                    # it from raw client scope instead could disagree with resolved_lien_type
                    # and send a Medicaid-Lien-MT creation down the "EV" (PLRP) path or vice
                    # versa, which also silently skips setting OnBenefits/OnBenefitsVerified
                    # below for a row that actually needed them.
                    stage = "To Send - EV" if resolved_lien_type == "Private Lien - PLRP" else "To Send - Claims Request"
                    row = {
                        "Unique": f"{client_id}_{lienholder_id}", "ClientId": client_id,
                        "Id": None, "AssignedUserId": get_assigned_user_id(collector_id),
                        "Stage": stage, "CollectorId": collector_id, "LienholderId": lienholder_id,
                        "IsMCO": True, "NewLienNote": build_note("C_lien", ticket_id, plan_name, state=state,
                                                                  response_type=response_type, response_date=response_date),
                    }
                    if stage == "To Send - Claims Request":
                        row["OnBenefits"], row["OnBenefitsVerified"] = "Yes", response_date
                    if has_dates:
                        row["Enrollmentstartdate"], row["EnrollmentEndDate"] = start_date, end_date
                    rows_update1.append(row)
                    updated_client_ids.add(client_id)
                    PLAN_FATE[fate_key] = "No existing lien to update; new lien created -- see Client Scope Upload / MCO Update 1"
                    df_in.at[idx, "Remarks"] = "Create Mcaid" if resolved_lien_type == "Medicaid Lien - MT" else "Create PLRP"
                    observations.append(_other_collector_observation(client_liens, collector_id, detail["names"]))

                df_in.at[idx, "Observation"] = _join_observations([_parse_obs.get(idx)] + observations)
            except Exception as _plan_err:
                del rows_a[_snap4[0]:], rows_b1[_snap4[1]:], rows_b2[_snap4[2]:]
                del rows_scope_upload[_snap4[3]:], rows_update1[_snap4[4]:], rows_pending_lien_match[_snap4[5]:]
                claimed_lien_ids.clear()
                claimed_lien_ids.update(_snap4[6])
                claimed_lien_ids_by_client.clear()
                claimed_lien_ids_by_client.update(_snap4[7])
                updated_client_ids.clear()
                updated_client_ids.update(_snap4[8])
                for _col in ("LienIds used for updates", "Stage", "Remarks"):
                    df_in.at[idx, _col] = None
                df_in.at[idx, "Observation"] = _parse_obs.get(idx)
                error_plans[idx] = _plan_err

        # The other input rows of a same-plan group copy the first row's outcome (or its error).
        for _sec_idx, _prim_idx in _same_plan_secondary.items():
            if _prim_idx in error_plans:
                error_plans[_sec_idx] = error_plans[_prim_idx]
                continue
            for _col in ("LienIds used for updates", "Stage", "Remarks"):
                df_in.at[_sec_idx, _col] = df_in.at[_prim_idx, _col]
            _prim_lien_id = df_in.at[_prim_idx, "LienId"]
            df_in.at[_sec_idx, "Observation"] = _join_observations([
                _parse_obs.get(_sec_idx), f"Same plan as input LienId {_prim_lien_id} - handled once there"])
            PLAN_FATE[(df_in.at[_sec_idx, "LienId"], df_in.at[_sec_idx, "MCO_Name"])] = (
                f"Same plan as LienId {_prim_lien_id} -- handled once there"
            )

        df_a  = pd.DataFrame(rows_a).reindex(columns=A_COLS)
        df_b1 = pd.DataFrame(rows_b1).reindex(columns=B1_COLS)
        df_b2 = pd.DataFrame(rows_b2).reindex(columns=B2_COLS)
        df_scope   = pd.DataFrame(rows_scope_upload).reindex(columns=SCOPE_UPLOAD_COLS)
        df_update1 = pd.DataFrame(rows_update1).reindex(columns=UPDATE1_COLS)
        df_pending_lien_match = pd.DataFrame(rows_pending_lien_match).reindex(columns=PENDING_LIEN_MATCH_COLS)

        for _df in (df_a, df_b1, df_b2, df_update1):
            if "Enrollmentstartdate" in _df.columns and "EnrollmentEndDate" in _df.columns:
                if not (_df["Enrollmentstartdate"].notna().any() or _df["EnrollmentEndDate"].notna().any()):
                    _df.drop(columns=["Enrollmentstartdate", "EnrollmentEndDate"], inplace=True)
        # IsMCO already True is never re-sent, and a Final No Entitlement lien keeps its own
        # LienholderId: such a column blank on every row is left out.
        for _df, _cols in ((df_a, ["IsMCO"]), (df_b1, ["IsMCO", "LienholderId"]), (df_b2, ["IsMCO"])):
            for _c in _cols:
                if not _df.empty and _c in _df.columns and not _df[_c].notna().any():
                    _df.drop(columns=[_c], inplace=True)

        # Real gap found (user report): "OnBenefits"/"OnBenefitsVerified" on df_update1
        # are only ever set when stage == "To Send - Claims Request" -- a run whose
        # Update 1 rows are ALL "To Send - EV" never sets them for a single row, yet
        # the two columns still showed up on the output, entirely blank. Same drop-if-
        # wholly-blank treatment as the Enrollmentstartdate/EnrollmentEndDate pair just
        # above. df_b1's own "Onbenefits"/"OnBenefitsVerified" (note the different
        # casing -- a distinct pair of columns, not the same ones) are unconditionally
        # set for every B1 row by spec, so they never hit this gap.
        if "OnBenefits" in df_update1.columns and "OnBenefitsVerified" in df_update1.columns:
            if not (df_update1["OnBenefits"].notna().any() or df_update1["OnBenefitsVerified"].notna().any()):
                df_update1.drop(columns=["OnBenefits", "OnBenefitsVerified"], inplace=True)

        return {
            "A": df_a, "B1": df_b1, "B2": df_b2,
            "client_scope_upload": df_scope, "mco_update1": df_update1,
            "pending_lien_match": df_pending_lien_match,
            "updated_client_ids": updated_client_ids,
            "claimed_lien_ids": claimed_lien_ids,
            "claimed_lien_ids_by_client": claimed_lien_ids_by_client,
            "error_plans": error_plans,
        }

    COUNT_PLACEHOLDER_SHEET = "Client-MCO Count Placeholder"
    COUNT_PLACEHOLDER_COLS = ["Id", "RequiredMedicaidMCOLienCount", "NewClientNote"]
    COUNT_PLACEHOLDER_NOTE = "Updated MCO count to 9 as placeholder until MCOs can be created"

    def build_count_placeholder(disc_df):
        """One row per client of MCOLienCount_Disc (count not 9, blank/0 included): count set
        to 9 until its MCO liens are created. Uploaded before Client-MCO Cnt Update, which
        then writes the real count. Empty (no tab) when there is no discrepancy."""
        if disc_df is None or disc_df.empty or "ClientId" not in disc_df.columns:
            return pd.DataFrame(columns=COUNT_PLACEHOLDER_COLS)
        ids = [c for c in dict.fromkeys(_id_key(v) for v in disc_df["ClientId"]) if c]
        return pd.DataFrame({"Id": ids, "RequiredMedicaidMCOLienCount": 9,
                             "NewClientNote": COUNT_PLACEHOLDER_NOTE}, columns=COUNT_PLACEHOLDER_COLS)

    def _plan_count_key(r):
        """One count key per real plan: CollectorId + LienholderId, else the plan name."""
        co, lh = _id_key(r.get("CollectorId")), _id_key(r.get("LienholderId"))
        return f"ID:{co}:{lh}" if co and lh else f"NAME:{_normalize_key(r.get('MCO_Name'))}"

    def build_client_cnt_update(updated_client_ids, df_in, liens_df, claimed_lien_ids, ticket_id, response_date,
                                 response_type, already_counted=None, counted_keys=None, keys_out=None):
        """counted_keys: client -> the plan/lien keys an earlier round of this ticket already
        counted (_ClientCountedPlanKeys) -- the total is the union, so a plan counted in an
        earlier round (e.g. one that was in Pending_Lien_Match_Review) is never counted twice.
        already_counted: the older number-only carry (review files from before), added on top.
        keys_out receives each counted client's keys, carried on to its review rows."""
        if already_counted is None:
            already_counted = {}
        if counted_keys is None:
            counted_keys = {}
        if not updated_client_ids:
            return pd.DataFrame(columns=["Id", "RequiredMedicaidMCOLienCount", "NewClientNote"])
        # One count per real plan: rows with the same CollectorId + LienholderId are the
        # same plan even when the input text differs (e.g. the same plan in 2 input rows).
        _cnt_rows = df_in[df_in["ClientId"].isin(updated_client_ids)]
        per_client_keys = {}
        for _, r in _cnt_rows.iterrows():
            per_client_keys.setdefault(r["ClientId"], set()).add(_plan_count_key(r))

        # A client can have OTHER active MCO liens this run's input never mentions --
        # those still count toward the total (as "LIEN:<id>" keys). Liens Step 4 already
        # claimed this run (Scenario A/B1/B2) are excluded -- already counted via the plan.
        #
        # A client carrying an earlier round's count is excluded from this scan entirely --
        # that count already reflects the earlier round; re-scanning risks double-counting
        # once that round's uploads land in the DB. Only genuinely new plans get added on top.
        _carried_client_ids = [cid for cid in updated_client_ids
                               if counted_keys.get(cid) or already_counted.get(cid, 0) > 0]
        other_active_mco = liens_df[
            liens_df["ClientId"].isin(updated_client_ids)
            & ~liens_df["ClientId"].isin(_carried_client_ids)
            & liens_df["IsMCO"].map(_is_true_flag)
            & (liens_df["statusid"] == 1)
            & (~liens_df["LienId"].isin(claimed_lien_ids))
        ]
        for cid, lid in zip(other_active_mco["ClientId"], other_active_mco["LienId"]):
            per_client_keys.setdefault(cid, set()).add(f"LIEN:{_id_key(lid)}")

        per_client_plans = {}
        for cid, keys in per_client_keys.items():
            if counted_keys.get(cid):
                keys = keys | set(counted_keys[cid])
                per_client_plans[cid] = len(keys)
            else:
                per_client_plans[cid] = len(keys) + already_counted.get(cid, 0)
            if keys_out is not None:
                keys_out[cid] = keys

        zantac_client_ids = [
            cid for cid in updated_client_ids
            if _normalize_key(df_in.loc[df_in["ClientId"] == cid, "Project_Name"].iloc[0]) == "ZANTAC"
        ]
        zantac_counts = {cid: per_client_plans[cid] for cid in zantac_client_ids if cid in per_client_plans}
        expanded_zantac = get_zantac_sibling_counts(zantac_counts) if zantac_counts else {}

        final_counts = dict(per_client_plans)
        final_counts.update(expanded_zantac)

        state = df_in["State"].iloc[0] if "State" in df_in.columns and len(df_in) else None
        note = build_note("C_client", ticket_id, None, state=state, response_type=response_type, response_date=response_date)
        rows = [{
            "Id": cid, "RequiredMedicaidMCOLienCount": cnt,
            # A client who already got a Client-MCO Cnt Update row in an earlier run
            # must not get a duplicate NewClientNote -- only the count refreshes. (On a
            # merge/append resume run this blank is corrected downstream in the
            # merge-mode WRITE OUTPUT section -- see _carried_blank_mask -- since
            # _MERGE_MODE isn't decided yet at this point.)
            "NewClientNote": None if (already_counted.get(cid, 0) > 0 or counted_keys.get(cid)) else note,
        } for cid, cnt in final_counts.items()]
        result_df = pd.DataFrame(rows)
        # Same convention already used for Enrollmentstartdate/EnrollmentEndDate
        # (dropped entirely when no row has one) -- an all-blank NewClientNote column
        # (every client in THIS run's tab was a carry-forward-only count refresh, see
        # already_counted above) shows nothing but empty cells, so drop it outright
        # rather than leave a column that never has anything in it. A MIXED batch
        # (some clients get a real note, some don't) still keeps the column -- only
        # the genuinely blank rows stay blank, exactly as today.
        if not result_df.empty and not result_df["NewClientNote"].notna().any():
            result_df = result_df.drop(columns=["NewClientNote"])
        return result_df

    _STEP4_RESULT = build_step4_outputs(
        yes_with_mco_df, sql_liens_df, TicketId, RESPONSE_DATE, RESPONSE_TYPE,
        already_claimed_lien_ids=_ALREADY_CLAIMED_LIEN_IDS_CARRY,
        lien_match_picks=_LIEN_MATCH_PICKS, lien_match_excluded=_LIEN_MATCH_CLIENT_EXCLUDED,
    )
    pending_lien_match_df = _STEP4_RESULT["pending_lien_match"]
    # Plans that hit an unexpected error in Step 4: only they go to Pending_Error_Review
    # (nothing of them was kept), and they are not counted this round.
    if _STEP4_RESULT["error_plans"]:
        for _err_idx, _err in _STEP4_RESULT["error_plans"].items():
            _record_plan_error(yes_with_mco_df.loc[[_err_idx]], "Step 4 (lien update)", _err)
        yes_with_mco_df = yes_with_mco_df.drop(index=list(_STEP4_RESULT["error_plans"])).copy()
    _COUNTED_KEYS_THIS_RUN = {}
    client_cnt_update_df = build_client_cnt_update(
        _STEP4_RESULT["updated_client_ids"], yes_with_mco_df, sql_liens_df, _STEP4_RESULT["claimed_lien_ids"],
        TicketId, RESPONSE_DATE, RESPONSE_TYPE, already_counted=_ALREADY_COUNTED_CARRY,
        counted_keys=_ALREADY_COUNTED_KEYS_CARRY, keys_out=_COUNTED_KEYS_THIS_RUN,
    )
    # A client whose placeholder row was already in an earlier round's upload files never
    # gets it again (its count may not be uploaded yet, so it can still show here).
    count_placeholder_df = build_count_placeholder(
        rmc_disc_df[~rmc_disc_df["ClientId"].map(_id_key).isin(_PLACEHOLDER_GIVEN_CARRY)]
        if not rmc_disc_df.empty and "ClientId" in rmc_disc_df.columns else rmc_disc_df)

    # Nothing to upload at all (every plan went to a review sheet): no WIP file this
    # round -- only the Review file (with Input_File_Original) and the log. The resume
    # round then creates the WIP with the original name.
    # (Only when something IS in review -- a run with nothing at all still writes its WIP,
    # e.g. to show Finalized_Clients_Excluded.)
    _NOTHING_TO_UPLOAD = yes_with_mco_df.empty and all(
        _f.empty for _f in (count_placeholder_df, _STEP4_RESULT["client_scope_upload"], _STEP4_RESULT["mco_update1"],
                            _STEP4_RESULT["A"], _STEP4_RESULT["B1"], _STEP4_RESULT["B2"], client_cnt_update_df)
    ) and bool(not pending_scope_df.empty or not pending_tpa_df.empty or not pending_tpa_blank_fields_df.empty
               or not pending_lien_match_df.empty or PENDING_ERROR_ROWS or PENDING_PLAN_TEXT_ROWS)

    # "MCO Count" on Yes_with_MCO -- the same ultimate per-client count that lands in
    # Client-MCO Cnt Update's RequiredMedicaidMCOLienCount (Zantac sibling expansion
    # included), just surfaced here too for visibility against each plan's own row.
    _client_to_mco_count = dict(zip(client_cnt_update_df["Id"], client_cnt_update_df["RequiredMedicaidMCOLienCount"]))
    yes_with_mco_df.insert(
        yes_with_mco_df.columns.get_loc("AssignedUserId") + 1, "MCO Count",
        yes_with_mco_df["ClientId"].map(_client_to_mco_count)
    )

    # Carry "already counted" state forward onto any STILL-pending row for the SAME
    # client -- see _ALREADY_COUNTED_CARRY's own comment above for the full bug this
    # fixes. A client who got a fresh Client-MCO Cnt Update row THIS run carries that
    # run's own real total forward (already includes whatever was carried in before);
    # a client who didn't (nothing new resolved for them this run) keeps whatever was
    # already carried in, unchanged.
    _new_carry_forward = dict(_ALREADY_COUNTED_CARRY)
    _new_carry_forward.update(_client_to_mco_count)
    _new_keys_carry = {cid: set(ks) for cid, ks in _ALREADY_COUNTED_KEYS_CARRY.items()}
    _new_keys_carry.update(_COUNTED_KEYS_THIS_RUN)
    _placeholder_given = set(_PLACEHOLDER_GIVEN_CARRY)
    if not _NOTHING_TO_UPLOAD and not count_placeholder_df.empty:
        _placeholder_given |= {_id_key(c) for c in count_placeholder_df["Id"]}

    def _counted_keys_str(cid):
        ks = _new_keys_carry.get(cid)
        return json.dumps(sorted(ks)) if ks else None

    pending_error_df = _pending_error_df(TicketId, RESPONSE_DATE, RESPONSE_TYPE)
    for _pdf in (pending_scope_df, pending_tpa_df, pending_tpa_blank_fields_df, pending_lien_match_df,
                 pending_error_df):
        if not _pdf.empty:
            _pdf["_ClientAlreadyCountedThisRun"] = _pdf["ClientId"].map(_new_carry_forward).fillna(0).astype(int)
            _pdf["_ClientCountedPlanKeys"] = _pdf["ClientId"].map(_counted_keys_str)
            _pdf["_CountPlaceholderGiven"] = _pdf["ClientId"].map(
                lambda c: "YES" if _id_key(c) in _placeholder_given else None)

    # Same carry-forward idea as _ClientAlreadyCountedThisRun above, but for WHICH
    # LienIds this client's plans already claimed (this run, plus whatever was
    # already carried in) -- see _ALREADY_CLAIMED_LIEN_IDS_CARRY's own comment for
    # the full bug this fixes. A LATER resume round reads this back and excludes
    # these LienIds too, so a different plan for the same client can never match
    # the SAME already-claimed lien again.
    _new_claimed_carry_forward = {cid: set(ids) for cid, ids in _ALREADY_CLAIMED_LIEN_IDS_CARRY.items()}
    for _cid, _claimed_set in _STEP4_RESULT["claimed_lien_ids_by_client"].items():
        _new_claimed_carry_forward.setdefault(_cid, set()).update(_claimed_set)

    def _claimed_lien_ids_str(cid):
        return ",".join(sorted(_new_claimed_carry_forward.get(cid, set())))

    for _pdf in (pending_scope_df, pending_tpa_df, pending_tpa_blank_fields_df, pending_lien_match_df,
                 pending_error_df):
        if not _pdf.empty:
            _pdf["_ClientAlreadyClaimedLienIds"] = _pdf["ClientId"].apply(_claimed_lien_ids_str)

    # =====================================================
    # REMARKS COLUMN — accountability for every qualified (LienId, MCO_Name)
    # =====================================================
    # PLAN_FATE is only keyed by (LienId, MCO_Name) for plans that made it into
    # yes_with_mco_df -- a LienId with no qualified plan gets no Remarks. An LA-file
    # Mismatch-tab flag is noted too, but only as a supplement to a real outcome, never
    # standalone -- a Mismatch flag with no plan text isn't actionable on its own.
    _lienid_to_remarks = {}
    for (_fate_lien_id, _fate_plan), _fate_remark in PLAN_FATE.items():
        _lienid_to_remarks.setdefault(_clean_lien_id(_fate_lien_id), []).append(f"{_fate_plan}: {_fate_remark}")

    _REMARKS_COL = _tool_remarks_col(ev_df)
    ev_df[_REMARKS_COL] = _input_file_remarks(ev_df, LIENID_COL_IN_EVDF, SSN_COL_IN_EVDF, _lienid_to_remarks,
                                              MISMATCH_FLAGGED_LIENIDS)
    ev_df = _drop_empty_info_cols(ev_df, [_REMARKS_COL])

    def _write_blocks_with_gap(writer, sheet_name, dfs, gap=3):
        """Writes multiple DataFrames left-to-right into ONE sheet, `gap` blank columns
        apart -- used for MCO Update 2's Scenario A/B1/B2 column-blocks."""
        start_col = 0
        for df in dfs:
            df.to_excel(writer, sheet_name=sheet_name, index=False, startcol=start_col)
            start_col += df.shape[1] + gap

    # =====================================================
    # RESUME OUTPUT: append into the original ticket's files, or keep separate
    # =====================================================
    # User request: a resume run used to always create its OWN new WIP/CSV set, which
    # (since the 4 action-tab CSVs are named by plain tab name, no TicketId/round
    # prefix -- see the CSV export section below) silently overwrote an EARLIER
    # round's already-produced CSVs for the same ticket. Ask once, only on a resume
    # run (the very first/original run has nothing yet to append into).
    # (The question itself is asked further down, once the original WIP's path is known:
    # no question when that WIP doesn't exist yet or this round has nothing to upload.)
    _MERGE_MODE = None

    def _resume_round_number(name):
        """Which resume round THIS run is, purely from the review-file naming
        convention (_next_review_stem's own '_Review' -> '_Review1' -> '_Review2' ->
        ... scheme): no trailing digit means the very FIRST resume round (round 1);
        '_Review1' means round 2; and so on."""
        m = re.search(r'_Review(\d*)$', name)
        if m is None:
            return 1
        return (int(m.group(1)) if m.group(1) else 0) + 1

    def _original_wip_stem(name):
        """Strips '_Review'/'_Review1'/'_Review2'/... back off a resume run's own
        input_filename, recovering the ORIGINAL (non-resume) run's own input_filename
        -- the same base every resume round in this ticket's chain would recover, so
        'append' always finds/extends the SAME one file regardless of how many rounds
        deep the chain is."""
        return re.sub(r'_Review\d*$', '', name)

    def _sheet_name_with_round_suffix(base_name, round_n):
        """Excel sheet names are capped at 31 characters -- '-R{n}' is appended,
        trimming the base name from the end just enough to fit if needed (defensive;
        none of today's real tab names are anywhere close to needing it)."""
        suffix = f"-R{round_n}"
        available = 31 - len(suffix)
        return base_name[:available] + suffix

    def _append_df(existing, new):
        if existing.empty:
            return new
        if new.empty:
            return existing
        return pd.concat([existing, new], ignore_index=True)

    def _safe_str_df(df):
        """Stringifies every column (needed so ClientId/LienId etc. compare
        consistently as text against an existing copy re-read from a CSV/sheet, which
        pandas always reads back as str) WITHOUT turning a genuine blank into the
        literal text "None" -- plain .astype(str) does that (real bug, user report:
        NewClientNote showed the word "None" instead of being empty for a client
        whose note is deliberately blank -- see build_client_cnt_update's own
        already-counted rule). Nullness is captured from the ORIGINAL df before
        stringifying; .where(..., None) alone re-upcasts a restored blank to NaN (a
        real float, the same Excel-write risk as elsewhere) rather than a true None,
        so the restore happens on an object-dtype copy instead, which holds a plain
        None with no coercion."""
        if df.empty:
            return df
        _was_null = df.isna()
        result = df.astype(str).astype(object)
        result[_was_null] = None
        return result

    def _upsert_client_cnt_update(existing_df, new_df):
        """Merges Client-MCO Cnt Update rows by Id (the ClientId, named "Id" as this
        tab's own first column) -- the newest run's row
        always wins (its RequiredMedicaidMCOLienCount already reflects the fully
        carried-forward total), EXCEPT its NewClientNote is kept from the OLDER row
        when the new one is blank. Blank is intentional (see build_client_cnt_update):
        a client already counted in an earlier run gets no SECOND note, since that
        would have been a duplicate note on a SEPARATE per-run upload -- but once runs
        are merged into ONE deliverable, the real note from that earlier round must
        not disappear entirely just because it isn't being duplicated. Ticket/date are
        identical across a whole resume chain (carried forward, never re-asked), so
        whichever round's note text survives is never wrong, only which one stays."""
        if existing_df.empty or "Id" not in existing_df.columns:
            return new_df.copy()
        if new_df.empty:
            return existing_df.copy()
        existing_df = existing_df.drop_duplicates(subset=["Id"], keep="last")
        _has_note_both = "NewClientNote" in existing_df.columns and "NewClientNote" in new_df.columns
        merged = existing_df.set_index("Id", drop=False)
        for _, new_row in new_df.iterrows():
            _cid = new_row["Id"]
            _row = new_row.copy()
            if _has_note_both and _cid in merged.index:
                _new_note = _row.get("NewClientNote")
                if pd.isna(_new_note) or str(_new_note).strip() == "":
                    _row["NewClientNote"] = merged.loc[_cid, "NewClientNote"]
            merged.loc[_cid] = _row
        return merged.reset_index(drop=True)

    _MCO_UPDATE2_COL_SETS = {"A": A_COLS, "B1": B1_COLS, "B2": B2_COLS}

    def _append_unique(existing, new):
        """Rows of new added under existing; an exact repeat of a row already there is not
        added again (all values compared as text)."""
        merged = _append_df(existing, new)
        if merged.empty:
            return merged
        # Built cell by cell: a blank stays "" whatever pandas' text dtype does with NaN.
        key = merged.apply(lambda r: "\x1f".join("" if pd.isna(v) else str(v) for v in r), axis=1)
        return merged[~key.duplicated(keep="first")].reset_index(drop=True)

    def _read_mco_update2_blocks(ws):
        """{scenario: DataFrame} of an existing MCO Update 2 sheet. Blocks are runs of
        header cells with blank columns between them: B1 has "Stage", B2 has
        "LienholderId" (no Stage), A is the other one -- whichever columns that round
        happened to write."""
        headers = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]
        blocks, c = {}, 0
        while c < len(headers):
            if headers[c] is None or str(headers[c]).strip() == "":
                c += 1
                continue
            start = c
            while c < len(headers) and headers[c] is not None and str(headers[c]).strip() != "":
                c += 1
            names = _clean_headers(headers[start:c])
            if "Id" not in names:
                # Never dropped silently (the sheet is rewritten): the caller keeps the sheet
                # as it is and puts this round's rows on their own tab.
                raise ValueError(f"a block without an 'Id' column (headers: {', '.join(names)})")
            rows = []
            for r in range(2, ws.max_row + 1):
                vals = [ws.cell(row=r, column=start + 1 + j).value for j in range(len(names))]
                if any(v is not None and str(v).strip() != "" for v in vals):
                    rows.append(vals)
            key = "B1" if "Stage" in names else ("B2" if "LienholderId" in names else "A")
            block = pd.DataFrame(rows, columns=names)
            blocks[key] = block if key not in blocks else pd.concat([blocks[key], block], ignore_index=True)
        return blocks

    def _merge_mco_update2_frames(existing, new_blocks):
        """{scenario: DataFrame} with every round's rows: per block the union of the columns
        both sides wrote (in the template order), one row per lien Id (this round wins)."""
        merged = {}
        for key in ("A", "B1", "B2"):
            old = existing.get(key, pd.DataFrame())
            new = new_blocks.get(key, pd.DataFrame())
            old = old if old is not None else pd.DataFrame()
            new = new if new is not None else pd.DataFrame()
            if old.empty and new.empty:
                continue
            for side in (old, new):
                if not side.empty and "Id" not in side.columns:
                    raise ValueError("no 'Id' column -- rows can't be matched by lien")
            present = set(old.columns) | set(new.columns)
            cols = [c for c in _MCO_UPDATE2_COL_SETS[key] if c in present]
            combined = pd.concat([old.reindex(columns=cols), new.reindex(columns=cols)], ignore_index=True)
            _ids = combined["Id"].map(_id_key)
            combined = combined[~(_ids.duplicated(keep="last") & _ids.notna())].reset_index(drop=True)
            merged[key] = combined
        return merged

    def _rebuild_mco_update2(wb, sheet_name, new_blocks):
        """Append mode: MCO Update 2 is rewritten whole from the file's own blocks plus this
        round's rows (see _merge_mco_update2_frames), blocks 3 blank columns apart. Nothing
        to write and no sheet yet -> no sheet."""
        existing = _read_mco_update2_blocks(wb[sheet_name]) if sheet_name in wb.sheetnames else {}
        merged = _merge_mco_update2_frames(existing, new_blocks)
        if not merged:
            return
        if sheet_name in wb.sheetnames:
            del wb[sheet_name]
        ws = wb.create_sheet(sheet_name)
        start_col = 1
        for key in ("A", "B1", "B2"):
            df = merged.get(key)
            if df is None or df.empty:
                continue
            for j, col_name in enumerate(df.columns):
                ws.cell(row=1, column=start_col + j, value=col_name)
            for i, (_, row) in enumerate(df.iterrows()):
                for j, col_name in enumerate(df.columns):
                    v = row[col_name]
                    ws.cell(row=2 + i, column=start_col + j, value=None if pd.isna(v) else v)
            start_col += len(df.columns) + 3

    def _mark_key_headers(ws, df, col_names):
        """Peach header (header cell only) on the key columns the user matches rows by."""
        fill = _solid_fill(KEY_HEADER_COLOR)
        cols = list(df.columns)
        for c in col_names:
            if c in cols:
                ws.cell(row=1, column=cols.index(c) + 1).fill = fill

    def _reorder_sheets(wb, desired_order):
        """Moves sheets already in wb to the front, in desired_order (only the ones
        actually present); any sheet not named there keeps its current relative order,
        placed after the named ones. Needed because pandas' if_sheet_exists="replace"
        (used to merge an existing WIP file's tabs across resume rounds) deletes and
        re-creates a sheet, which silently moves it to the very END of the workbook --
        this restores the intended reading order (Client Scope Upload -> Fetch New
        Lien IDs -> MCO Update 1 -> MCO Update 2 -> Client-MCO Cnt Update) regardless
        of what order this round's writes happened in."""
        present = [name for name in desired_order if name in wb.sheetnames]
        rest = [name for name in wb.sheetnames if name not in present]
        wb._sheets = [wb[name] for name in present + rest]

    def _snapshot_variant_sort_key(name, base):
        """The un-suffixed original copy sorts first, then -R1, -R2, ... in round
        order -- used to group a snapshot tab's own round-by-round copies together
        in the same left-to-right sequence they were created in."""
        if name == base:
            return -1
        m = re.match(rf"^{re.escape(base)}-R(\d+)$", name)
        return int(m.group(1)) if m else 10 ** 9

    def _build_merged_wip_tab_order(wb):
        """Full desired tab order for a merged WIP workbook: each snapshot tab's own
        round-by-round copies grouped together (Input_File_Original, then the single
        original-run-only SQL-Client Id copy -- it's never re-written on resume, see
        the RESUME OUTPUT design notes above, so there's only ever one -- then
        MCOLienCount_Disc, Finalized_Clients_Excluded, SQL-MCOs, SQL-Liens) -- same
        left-to-right order a normal (non-merge) run's own WIP already uses -- all
        BEFORE Yes_with_MCO, then the actionable tabs. Built fresh from wb.sheetnames
        each time since which round-suffixed variants exist grows every round."""
        order = []
        for base in ("Input_File_Original",):
            variants = [n for n in wb.sheetnames if n == base or re.match(rf"^{re.escape(base)}-R\d+$", n)]
            order.extend(sorted(variants, key=lambda n: _snapshot_variant_sort_key(n, base)))
        if "SQL-Client Id" in wb.sheetnames:
            order.append("SQL-Client Id")
        for base in ("MCOLienCount_Disc", "Finalized_Clients_Excluded", "SQL-MCOs", "SQL-Liens"):
            variants = [n for n in wb.sheetnames if n == base or re.match(rf"^{re.escape(base)}-R\d+$", n)]
            order.extend(sorted(variants, key=lambda n: _snapshot_variant_sort_key(n, base)))
        order.extend([
            "Yes_with_MCO", COUNT_PLACEHOLDER_SHEET, "Client Scope Upload", "Fetch New Lien IDs",
            "MCO Update 1", "MCO Update 2", "Client-MCO Cnt Update",
            "SQL-Post Creation", "SQL-Post MCO Cnt",
        ])
        return order

    # =====================================================
    # WRITE OUTPUT
    # =====================================================

    def _write_with_query(writer, df, sheet_name, query_str):
        df.to_excel(writer, sheet_name=sheet_name, index=False)
        ws = writer.book[sheet_name]
        sc = df.shape[1] + 2
        ws.cell(row=1, column=sc + 1, value="QUERY USED")
        for i, line in enumerate(query_str.strip().split("\n"), start=2):
            ws.cell(row=i, column=sc + 1, value=line)

    def _highlight_note_cols(ws, df, col_names, color=NOTE_COL_COLOR):
        """Notes THIS TOOL adds (summarizing an outcome/reason), not part of the original
        input/query data, so they stand out: Powder Blue on Yes_with_MCO, Lavender for
        Input_File_Original's Remarks."""
        _fill = _solid_fill(color)
        for col_name in col_names:
            if col_name not in df.columns:
                continue
            col_idx = list(df.columns).index(col_name) + 1
            for r in range(1, len(df) + 2):
                cell = ws.cell(row=r, column=col_idx)
                cell.fill = _fill
                if r == 1:
                    cell.font = _opxl.styles.Font(bold=True)

    _merge_wip_path = (
        os.path.join(_output_dir, f"{_original_wip_stem(input_filename)}_WIP.xlsx")
        if IS_RESUME_RUN else None
    )

    # Resume: append into the original ticket's WIP, or keep this round separate. No
    # question when this round has nothing to upload, or when the ticket has no WIP yet
    # (an earlier round had nothing to upload) -- then this round's WIP takes the
    # original name, like a first run.
    if IS_RESUME_RUN and not _NOTHING_TO_UPLOAD:
        if os.path.exists(_merge_wip_path):
            _divider()
            _merge_choice = ask_numbered_choice(
                ["Append into the original ticket's files", "Keep this run's files separate"],
                "How should this run's results be saved?",
                "Enter the number: "
            )
            _MERGE_MODE = "append" if _merge_choice.startswith("Append") else "separate"
        else:
            _MERGE_MODE = "first"
            OUTPUT_FILE = _merge_wip_path

    if _NOTHING_TO_UPLOAD:
        # Nothing to upload this round: no WIP file (see _NOTHING_TO_UPLOAD). The Review
        # file below also carries Input_File_Original with its Remarks.
        pass

    elif IS_RESUME_RUN and _MERGE_MODE == "append" and _merge_wip_path and os.path.exists(_merge_wip_path):
        # Append this round's own results into the ORIGINAL ticket's WIP file instead
        # of creating a brand-new one -- see the RESUME OUTPUT section above for the
        # full design (skip SQL-Client Id, merge the actionable tabs + Yes_with_MCO,
        # keep every other snapshot tab separate per round via a "-R{n}" suffix).
        _round_n = _resume_round_number(input_filename)

        # One workbook parse serves all 4 sheets below -- pd.read_excel(path, ...)
        # re-parses the ENTIRE file from scratch on every call when given a bare path
        # (the same real slowdown already found and fixed once for load_config and
        # once for the TPA file -- see the earlier perf investigation), so 4 separate
        # calls here would silently redo 4 full workbook reads instead of 1.
        _ensure_file_hydrated(_merge_wip_path)
        _merge_xls = _retry_on_permission_error(
            lambda: pd.ExcelFile(_merge_wip_path), "the original ticket's WIP file"
        )

        def _read_existing_sheet(xls, sheet_name):
            try:
                return pd.read_excel(xls, sheet_name=sheet_name, dtype=str)
            except (ValueError, KeyError):
                return pd.DataFrame()

        _existing_yes = _read_existing_sheet(_merge_xls, "Yes_with_MCO")
        _existing_csu = _read_existing_sheet(_merge_xls, "Client Scope Upload")
        _existing_mu1 = _read_existing_sheet(_merge_xls, "MCO Update 1")
        _existing_cnt = _read_existing_sheet(_merge_xls, "Client-MCO Cnt Update")
        _existing_ph = _read_existing_sheet(_merge_xls, COUNT_PLACEHOLDER_SHEET)
        _merge_xls.close()
        # One placeholder row per client across every round.
        _merged_ph = _append_df(_existing_ph if "Id" in _existing_ph.columns else pd.DataFrame(),
                                _safe_str_df(count_placeholder_df))
        if not _merged_ph.empty and "Id" in _merged_ph.columns:
            _merged_ph = _merged_ph.drop_duplicates(subset=["Id"], keep="first").reset_index(drop=True)

        _merged_yes = _append_df(_existing_yes, _safe_str_df(
            yes_with_mco_df.drop(columns=["LienType", "State", "_OpenEnd", "_TPA_Source"], errors="ignore")))
        # Real duplicate found: a Pending_Lien_Match_Review row was already TPA-
        # resolved and written to Yes_with_MCO in an earlier round (Step 3 succeeded
        # then too -- only Step 4's scenario tie was pending). On resume Step 3
        # re-derives the SAME (LienId, MCO_Name) match deterministically, so a plain
        # append duplicates it. De-duping on that pair only removes genuine
        # re-derivations, never two different liens sharing the same plan text.
        if not _merged_yes.empty and {"LienId", "MCO_Name"} <= set(_merged_yes.columns):
            _merged_yes = _merged_yes.drop_duplicates(subset=["LienId", "MCO_Name"], keep="last").reset_index(drop=True)
        # After merging, so a column filled in any round is kept.
        _merged_yes = _drop_empty_info_cols(_merged_yes, YES_WITH_MCO_INFO_COLS)
        # An exact repeat of a row already in the file is never added again.
        _merged_csu = _append_unique(_existing_csu, _safe_str_df(_STEP4_RESULT["client_scope_upload"]))
        _merged_mu1 = _append_unique(_existing_mu1, _safe_str_df(_STEP4_RESULT["mco_update1"]))

        # Real gap found: client_cnt_update_df was built further up with carried-note
        # suppression always on, so a client already counted in an earlier round got a
        # blank NewClientNote regardless of merge/separate mode. But a merged file is
        # the user's ONE single upload for the whole ticket, so every client should
        # show a real note. Only "separate" mode (below, untouched) keeps suppression.
        _carried_ids = ({cid for cid, cnt in _ALREADY_COUNTED_CARRY.items() if cnt > 0}
                        | {cid for cid, ks in _ALREADY_COUNTED_KEYS_CARRY.items() if ks})
        if not client_cnt_update_df.empty and _carried_ids:
            if "NewClientNote" not in client_cnt_update_df.columns:
                client_cnt_update_df["NewClientNote"] = None
            _carried_note = build_note(
                "C_client", TicketId, None,
                state=yes_with_mco_df["State"].iloc[0] if "State" in yes_with_mco_df.columns and len(yes_with_mco_df) else None,
                response_type=RESPONSE_TYPE, response_date=RESPONSE_DATE,
            )
            _carried_blank_mask = (
                client_cnt_update_df["Id"].isin(_carried_ids) & client_cnt_update_df["NewClientNote"].isna()
            )
            client_cnt_update_df.loc[_carried_blank_mask, "NewClientNote"] = _carried_note

        _merged_cnt = _upsert_client_cnt_update(_existing_cnt, _safe_str_df(client_cnt_update_df))

        _ensure_writable(_merge_wip_path)
        with pd.ExcelWriter(_merge_wip_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            if not _merged_yes.empty:
                _merged_yes.to_excel(writer, sheet_name="Yes_with_MCO", index=False)
                _highlight_note_cols(writer.book["Yes_with_MCO"], _merged_yes,
                                      ["LienIds used for updates", "Stage", "Remarks", "Observation"])
                _mark_key_headers(writer.book["Yes_with_MCO"], _merged_yes, ["Unique"])

            if not _merged_csu.empty:
                _merged_csu.to_excel(writer, sheet_name="Client Scope Upload", index=False)
                _fetch_ids = [c for c in _merged_csu.get("ClientId", pd.Series(dtype=object)).dropna().unique()
                              if str(c).strip()]
                if _fetch_ids:
                    _fetch_query = _FETCH_NEW_LIEN_IDS_QUERY_TEMPLATE.format(client_ids=make_sql_in(_fetch_ids))
                    if "Fetch New Lien IDs" in writer.book.sheetnames:
                        del writer.book["Fetch New Lien IDs"]
                    _fetch_ws = writer.book.create_sheet("Fetch New Lien IDs")
                    for _i, _line in enumerate(_fetch_query.split("\n"), start=1):
                        _fetch_ws.cell(row=_i, column=1, value=_line)

            if not _merged_mu1.empty:
                _merged_mu1.to_excel(writer, sheet_name="MCO Update 1", index=False)
                _mark_key_headers(writer.book["MCO Update 1"], _merged_mu1, ["Unique", "ClientId"])

            if not _merged_cnt.empty:
                _merged_cnt.to_excel(writer, sheet_name="Client-MCO Cnt Update", index=False)

            if not _merged_ph.empty:
                _merged_ph.to_excel(writer, sheet_name=COUNT_PLACEHOLDER_SHEET, index=False)

            _post_ids = ([c for c in _merged_yes["ClientId"].dropna().unique() if str(c).strip()]
                         if "ClientId" in _merged_yes.columns else [])
            if not _merged_yes.empty and _post_ids:
                # These queries are a ready-to-run reference for the user after the
                # merged file is uploaded -- since it's the ONE single upload for the
                # whole ticket, the ClientId list must cover every client across every
                # merged round (_merged_yes, not just this round's own df), in one
                # tab, never a round-suffixed snapshot. Same delete-and-recreate
                # pattern as "Fetch New Lien IDs" for the same reason.
                _post_client_ids = make_sql_in(_post_ids)

                _post_creation_query = _SQL_POST_CREATION_QUERY_TEMPLATE.format(client_ids=_post_client_ids)
                if "SQL-Post Creation" in writer.book.sheetnames:
                    del writer.book["SQL-Post Creation"]
                _post_creation_ws = writer.book.create_sheet("SQL-Post Creation")
                for _i, _line in enumerate(_post_creation_query.split("\n"), start=1):
                    _post_creation_ws.cell(row=_i, column=1, value=_line)

                _post_cnt_query = _SQL_POST_MCO_CNT_QUERY_TEMPLATE.format(client_ids=_post_client_ids)
                if "SQL-Post MCO Cnt" in writer.book.sheetnames:
                    del writer.book["SQL-Post MCO Cnt"]
                _post_cnt_ws = writer.book.create_sheet("SQL-Post MCO Cnt")
                for _i, _line in enumerate(_post_cnt_query.split("\n"), start=1):
                    _post_cnt_ws.cell(row=_i, column=1, value=_line)

            try:
                _rebuild_mco_update2(writer.book, "MCO Update 2", {
                    "A": _STEP4_RESULT["A"], "B1": _STEP4_RESULT["B1"], "B2": _STEP4_RESULT["B2"],
                })
            except Exception as _mu2_err:
                # The earlier MCO Update 2 was edited by hand in a way it can't be read back:
                # it is kept as it is and this round's rows go to their own tab.
                _mu2_round = _sheet_name_with_round_suffix("MCO Update 2", _round_n)
                _mu2_blocks = [d for d in (_STEP4_RESULT["A"], _STEP4_RESULT["B1"], _STEP4_RESULT["B2"]) if not d.empty]
                if _mu2_blocks:
                    _write_blocks_with_gap(writer, _mu2_round, _mu2_blocks)
                print(f"⚠️  Could not merge into the existing MCO Update 2 ({_short_text(str(_mu2_err), 100)}) "
                      f"-- this round's rows are on the '{_mu2_round}' tab instead.")

            _ifo_sheet_name = _sheet_name_with_round_suffix("Input_File_Original", _round_n)
            ev_df.to_excel(writer, sheet_name=_ifo_sheet_name, index=False)
            # Same Remarks-column highlight as the normal (non-merge) write below.
            _highlight_note_cols(writer.book[_ifo_sheet_name], ev_df, [_REMARKS_COL], REMARKS_COL_COLOR)

            _write_with_query(
                writer, sql_mcos_df, _sheet_name_with_round_suffix("SQL-MCOs", _round_n), sql_mcos_query
            )
            _write_with_query(
                writer, sql_liens_df.drop(columns=["Last_Name", "First_Name"]),
                _sheet_name_with_round_suffix("SQL-Liens", _round_n), sql_liens_query
            )
            if not rmc_disc_df.empty:
                rmc_disc_df.to_excel(
                    writer, sheet_name=_sheet_name_with_round_suffix("MCOLienCount_Disc", _round_n), index=False
                )
            if not finalized_df.empty:
                finalized_excluded_df.to_excel(
                    writer, sheet_name=_sheet_name_with_round_suffix("Finalized_Clients_Excluded", _round_n),
                    index=False
                )

            _reorder_sheets(writer.book, _build_merged_wip_tab_order(writer.book))

            for _tab_name in (COUNT_PLACEHOLDER_SHEET, "Client Scope Upload", "MCO Update 1", "MCO Update 2",
                              "Client-MCO Cnt Update"):
                if _tab_name in writer.book.sheetnames:
                    writer.book[_tab_name].sheet_properties.tabColor = UPLOAD_TAB_COLOR
            for _tab_name in (
                _sheet_name_with_round_suffix("MCOLienCount_Disc", _round_n),
                _sheet_name_with_round_suffix("Finalized_Clients_Excluded", _round_n),
            ):
                if _tab_name in writer.book.sheetnames:
                    writer.book[_tab_name].sheet_properties.tabColor = ALERT_TAB_COLOR

        OUTPUT_FILE = _merge_wip_path
        print(f"\n✅ Appended into the original ticket's WIP file: {OUTPUT_FILE}")

    else:
        # Non-resume run, "keep separate" chosen, or the original WIP file is missing
        # (moved/deleted -- can't append into something that doesn't exist) -- write
        # this run's own complete, standalone WIP file exactly as before.
        _ensure_writable(OUTPUT_FILE)
        with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:

            ev_df.to_excel(writer, sheet_name="Input_File_Original", index=False)

            # Highlight the Remarks column -- it's an accountability note THIS TOOL adds
            # (why each row was matched/excluded/sent to review), not part of the original
            # input file, so it should visually stand out from the columns that are.
            _highlight_note_cols(writer.book["Input_File_Original"], ev_df, [_REMARKS_COL], REMARKS_COL_COLOR)

            _write_with_query(writer, sql_df,      "SQL-Client Id", sql_query)

            if not rmc_disc_df.empty:
                rmc_disc_df.to_excel(writer, sheet_name="MCOLienCount_Disc", index=False)

            if not finalized_df.empty:
                finalized_excluded_df.to_excel(writer, sheet_name="Finalized_Clients_Excluded", index=False)

            _write_with_query(writer, sql_mcos_df,  "SQL-MCOs",   sql_mcos_query)
            _write_with_query(
                writer, sql_liens_df.drop(columns=["Last_Name", "First_Name"]), "SQL-Liens", sql_liens_query
            )

            # LienType and State are dropped only from this DISPLAYED copy -- both stay on
            # the working yes_with_mco_df (Step 4's Scenario C reads LienType directly, and
            # its NewLienNote/NewClientNote read State) but neither is meant to be a visible
            # Yes_with_MCO column.
            _ywm_display_df = _drop_empty_info_cols(
                yes_with_mco_df.drop(columns=["LienType", "State", "_OpenEnd", "_TPA_Source"], errors="ignore"), YES_WITH_MCO_INFO_COLS)
            _ywm_display_df.to_excel(writer, sheet_name="Yes_with_MCO", index=False)
            _highlight_note_cols(writer.book["Yes_with_MCO"], _ywm_display_df,
                                  ["LienIds used for updates", "Stage", "Remarks", "Observation"])
            _mark_key_headers(writer.book["Yes_with_MCO"], _ywm_display_df, ["Unique"])

            # Tab order: Client-MCO Count Placeholder (first upload) -> Client Scope Upload ->
            # Fetch New Lien IDs -> MCO Update 1 -> MCO Update 2 -> Client-MCO Cnt Update
            # (last) -- the natural reading order of the creation/update flow.
            if not count_placeholder_df.empty:
                count_placeholder_df.to_excel(writer, sheet_name=COUNT_PLACEHOLDER_SHEET, index=False)

            if not _STEP4_RESULT["client_scope_upload"].empty:
                _STEP4_RESULT["client_scope_upload"].to_excel(writer, sheet_name="Client Scope Upload", index=False)

                # A ready-to-run query -- ClientIds already filled in -- for the user to
                # pull the real Id/CollectorId/LienholderId/etc. of each newly-created lien
                # once these rows are uploaded to the system, matched back into MCO Update
                # 1 via "Unique" (see UPDATE1_COLS above).
                _fetch_client_ids = make_sql_in(_STEP4_RESULT["client_scope_upload"]["ClientId"].dropna().unique())
                _fetch_query = _FETCH_NEW_LIEN_IDS_QUERY_TEMPLATE.format(client_ids=_fetch_client_ids)
                _fetch_ws = writer.book.create_sheet("Fetch New Lien IDs")
                for _i, _line in enumerate(_fetch_query.split("\n"), start=1):
                    _fetch_ws.cell(row=_i, column=1, value=_line)

            if not _STEP4_RESULT["mco_update1"].empty:
                _STEP4_RESULT["mco_update1"].to_excel(writer, sheet_name="MCO Update 1", index=False)
                _mark_key_headers(writer.book["MCO Update 1"], _STEP4_RESULT["mco_update1"], ["Unique", "ClientId"])

            _mco_update2_blocks = [df for df in (_STEP4_RESULT["A"], _STEP4_RESULT["B1"], _STEP4_RESULT["B2"]) if not df.empty]
            if _mco_update2_blocks:
                _write_blocks_with_gap(writer, "MCO Update 2", _mco_update2_blocks)

            if not client_cnt_update_df.empty:
                client_cnt_update_df.to_excel(writer, sheet_name="Client-MCO Cnt Update", index=False)

            if not yes_with_mco_df.empty:
                _post_client_ids = make_sql_in(yes_with_mco_df["ClientId"].dropna().unique())

                _post_creation_query = _SQL_POST_CREATION_QUERY_TEMPLATE.format(client_ids=_post_client_ids)
                _post_creation_ws = writer.book.create_sheet("SQL-Post Creation")
                for _i, _line in enumerate(_post_creation_query.split("\n"), start=1):
                    _post_creation_ws.cell(row=_i, column=1, value=_line)

                _post_cnt_query = _SQL_POST_MCO_CNT_QUERY_TEMPLATE.format(client_ids=_post_client_ids)
                _post_cnt_ws = writer.book.create_sheet("SQL-Post MCO Cnt")
                for _i, _line in enumerate(_post_cnt_query.split("\n"), start=1):
                    _post_cnt_ws.cell(row=_i, column=1, value=_line)

            # Give the "action" tabs (the actual creation/update output, as opposed to the
            # audit/diagnostic tabs above) a shared accent color so they visually stand apart.
            for _tab_name in (COUNT_PLACEHOLDER_SHEET, "Client Scope Upload", "MCO Update 1", "MCO Update 2",
                              "Client-MCO Cnt Update"):
                if _tab_name in writer.book.sheetnames:
                    writer.book[_tab_name].sheet_properties.tabColor = UPLOAD_TAB_COLOR

            # The 2 discrepancy/exclusion tabs (rows flagged or filtered OUT of normal
            # processing) get a red tab color whenever they actually appear, so they catch
            # the eye immediately instead of blending in with the plain audit tabs.
            for _tab_name in ("MCOLienCount_Disc", "Finalized_Clients_Excluded"):
                if _tab_name in writer.book.sheetnames:
                    writer.book[_tab_name].sheet_properties.tabColor = ALERT_TAB_COLOR


    # Standalone CSV exports of the "action" tabs -- what actually gets uploaded/imported
    # downstream, handed off separately from the WIP workbook. MCO Update 2's 3 blocks
    # have different column sets (side by side only for a human reader), so each block
    # that has rows gets its own CSV, named by what it really updates: FNI (the Final No
    # Entitlement block), LH, IsMCO, Dates -- e.g. "MCO Update 2 - LH & IsMCO".
    def _mu2_csv_name(block, columns):
        cols = set(columns)
        parts = (["FNI"] if block == "B1" else []) + (["LH"] if "LienholderId" in cols else [])
        parts += (["IsMCO"] if "IsMCO" in cols else [])
        parts += (["Dates"] if {"Enrollmentstartdate", "EnrollmentEndDate"} & cols else [])
        parts = parts or ["Update"]
        what = parts[0] if len(parts) == 1 else ", ".join(parts[:-1]) + " & " + parts[-1]
        return f"MCO Update 2 - {what}"

    def _mu2_csv_block(file_name):
        """The MCO Update 2 block a plain-named CSV holds (any naming, old or new), or None.
        A round-suffixed file ("...-R2.csv") belongs to its own round and is never touched."""
        m = re.fullmatch(r"MCO Update 2 - (.+)\.csv", file_name)
        if not m or re.search(r"-R\d+$", m.group(1)):
            return None
        what = m.group(1)
        if what.startswith("FNI"):
            return "B1"
        if what.startswith("LH"):
            return "B2"
        return "A" if re.match(r"(IsMCO|Dates|Update)\b", what) else None

    def _write_csv(path, frame):
        path = _long_path(path)
        _ensure_writable(path)
        frame.to_csv(path, index=False)

    _csv_exports = {
        COUNT_PLACEHOLDER_SHEET: count_placeholder_df,
        "Client Scope Upload": _STEP4_RESULT["client_scope_upload"],
        "MCO Update 1": _STEP4_RESULT["mco_update1"],
        "Client-MCO Cnt Update": client_cnt_update_df,
    }
    _mu2_exports = {"A": _STEP4_RESULT["A"], "B1": _STEP4_RESULT["B1"], "B2": _STEP4_RESULT["B2"]}
    if any(not df.empty for df in list(_csv_exports.values()) + list(_mu2_exports.values())):
        _csv_dir = os.path.join(_output_dir, "Upload_Update CSVs")
        os.makedirs(_csv_dir, exist_ok=True)
        _round_tag = f"-R{_resume_round_number(input_filename)}" if IS_RESUME_RUN else ""
        for _csv_name, _csv_df in _csv_exports.items():
            if _csv_df.empty:
                continue
            if IS_RESUME_RUN and _MERGE_MODE == "separate":
                # Plain tab-name filenames used to silently overwrite an earlier
                # round's own CSV of the same name (real user report) -- a round
                # suffix keeps every round's CSVs coexisting in the same folder.
                _csv_path = os.path.join(_csv_dir, f"{_csv_name}{_round_tag}.csv")
            else:
                _csv_path = os.path.join(_csv_dir, f"{_csv_name}.csv")
                if IS_RESUME_RUN and _MERGE_MODE == "append" and os.path.exists(_csv_path):
                    _ensure_file_hydrated(_csv_path)
                    try:
                        _existing_csv_df = _retry_on_permission_error(
                            lambda: _read_existing_csv(_csv_path), "the existing CSV file"
                        )
                        if _csv_name == "Client-MCO Cnt Update":
                            _merged_csv = _upsert_client_cnt_update(_existing_csv_df, _safe_str_df(_csv_df))
                        elif _csv_name == COUNT_PLACEHOLDER_SHEET:
                            _merged_csv = _append_df(_existing_csv_df, _safe_str_df(_csv_df))
                            if "Id" in _merged_csv.columns:
                                _merged_csv = _merged_csv.drop_duplicates(subset=["Id"], keep="first")
                        else:
                            _merged_csv = _append_unique(_existing_csv_df, _safe_str_df(_csv_df))
                        _csv_df = _merged_csv
                    except Exception as _csv_err:
                        # A CSV edited by hand (headers changed, re-saved oddly) is never a dead end:
                        # this round's rows go into their own file instead.
                        _csv_path = os.path.join(_csv_dir, f"{_csv_name}{_round_tag}.csv")
                        print(f"⚠️  Could not add to the existing {_csv_name}.csv ({_short_text(str(_csv_err), 100)}) "
                              f"-- this round's rows were written to {os.path.basename(_csv_path)} instead.")
            _write_csv(_csv_path, _csv_df)

        _csv_files_now = os.listdir(_long_path(_csv_dir))
        for _blk, _blk_df in _mu2_exports.items():
            if _blk_df.empty:
                continue
            if IS_RESUME_RUN and _MERGE_MODE == "separate":
                _write_csv(os.path.join(_csv_dir, f"{_mu2_csv_name(_blk, _blk_df.columns)}{_round_tag}.csv"), _blk_df)
                continue
            # This block's earlier plain-named CSV(s), whatever columns they were named after.
            _olds = [f for f in _csv_files_now if _mu2_csv_block(f) == _blk]
            _out_df = _blk_df
            if IS_RESUME_RUN and _MERGE_MODE == "append" and _olds:
                try:
                    _acc = _safe_str_df(_blk_df)
                    for _old in _olds:
                        _old_path = os.path.join(_csv_dir, _old)
                        _ensure_file_hydrated(_old_path)
                        _old_df = _retry_on_permission_error(
                            lambda _p=_old_path: _read_existing_csv(_p), "the existing CSV file")
                        # Union of columns, one row per lien Id; this round's row wins.
                        _acc = _merge_mco_update2_frames({_blk: _old_df}, {_blk: _acc})[_blk]
                    _out_df = _acc
                except Exception as _csv_err:
                    _fallback = f"{_mu2_csv_name(_blk, _blk_df.columns)}{_round_tag}.csv"
                    _write_csv(os.path.join(_csv_dir, _fallback), _blk_df)
                    print(f"⚠️  Could not add to the existing {', '.join(_olds)} ({_short_text(str(_csv_err), 100)}) "
                          f"-- this round's rows were written to {_fallback} instead.")
                    continue
            _new_file = f"{_mu2_csv_name(_blk, _out_df.columns)}.csv"
            _write_csv(os.path.join(_csv_dir, _new_file), _out_df)
            # The block's other plain-named file is replaced by this one: on append its rows
            # are inside it; on a new run it is an earlier run's file (before, the fixed name
            # was simply overwritten). Removed only now that the new file is written.
            for _old in _olds:
                if _old == _new_file:
                    continue
                try:
                    os.remove(_long_path(os.path.join(_csv_dir, _old)))
                    print(f"ℹ️  {_old} replaced by {_new_file}"
                          + (" (its rows are in it)." if IS_RESUME_RUN and _MERGE_MODE == "append" else "."))
                except OSError:
                    print(f"⚠️  Could not remove the older {_old} (open in Excel?) -- "
                          f"{_new_file} is the one to upload; please delete {_old} by hand.")


    # Review items (Pending_Scope_Review / Pending_TPA_Review / Pending_TPA_Blank_Fields /
    # Pending_Lien_Match_Review) go into their OWN separate spreadsheet, not the main
    # processing output -- give this file's path back in as `Path of Import File`
    # whenever ready to resolve them.
    pending_plan_text_df = _plan_text_review_df(TicketId, RESPONSE_DATE, RESPONSE_TYPE)
    _any_pending = (not pending_scope_df.empty or not pending_tpa_df.empty
                    or not pending_tpa_blank_fields_df.empty or not pending_lien_match_df.empty
                    or not pending_plan_text_df.empty or not pending_error_df.empty)
    if _NOTHING_TO_UPLOAD:
        _gap()
        print("ℹ️  Nothing to upload this round -- every plan needs review. No WIP file was created; "
              "the Review file also has Input_File_Original with its Remarks.")
    if _any_pending:
        _ensure_writable(REVIEW_FILE)
        with pd.ExcelWriter(REVIEW_FILE, engine="openpyxl") as review_writer:
            if not pending_error_df.empty:
                _err_df = _drop_empty_info_cols(pending_error_df, REVIEW_INFO_COLS)
                _write_error_review_sheet(review_writer, _err_df)
            if not pending_plan_text_df.empty:
                pending_plan_text_df.to_excel(review_writer, sheet_name="Pending_Plan_Text_Review", index=False)
                _style_plan_text_review_sheet(review_writer.sheets["Pending_Plan_Text_Review"], pending_plan_text_df)
            if not pending_scope_df.empty:
                _scope_review_df = _drop_empty_info_cols(pending_scope_df, REVIEW_INFO_COLS)
                _scope_review_df.to_excel(review_writer, sheet_name="Pending_Scope_Review", index=False)
                _fill_cells(review_writer.sheets["Pending_Scope_Review"], _scope_review_df, ["Scope_Classification"])
                _add_review_validations(review_writer.sheets["Pending_Scope_Review"], _scope_review_df,
                                        list_cols={"Scope_Classification": ["100% PLRP", "No PLRP"]})
            if not pending_tpa_df.empty:
                _tpa_review_df = _drop_empty_info_cols(_prepare_tpa_review_df(pending_tpa_df), REVIEW_INFO_COLS)
                _tpa_review_df.to_excel(review_writer, sheet_name="Pending_TPA_Review", index=False)
                _format_tpa_review_sheet(review_writer.sheets["Pending_TPA_Review"], _tpa_review_df)
            if not pending_tpa_blank_fields_df.empty:
                # Same MCOn_Existing reasoning as above -- a row here already has a
                # confirmed Tort/Name/Scope match (that's why it's on THIS sheet, not
                # Pending_TPA_Review); only CollectorId/LienType are missing, and
                # existing on-file plans don't help supply those either. The
                # TPA_Match_* columns stay on THIS sheet -- they're genuinely populated
                # here, used by the resume/patch logic, and useful context for the user.
                _blank_fields_drop_cols = [
                    c for c in pending_tpa_blank_fields_df.columns if c.startswith("MCO") and c.endswith("_Existing")
                ]
                _bf_df = _drop_empty_info_cols(pending_tpa_blank_fields_df.drop(columns=_blank_fields_drop_cols,
                                                                                errors="ignore"), REVIEW_INFO_COLS)
                _bf_df.to_excel(review_writer, sheet_name="Pending_TPA_Blank_Fields", index=False)
                _bf_ws = review_writer.sheets["Pending_TPA_Blank_Fields"]
                for _r, _reason in enumerate(_bf_df.get("PendingReason", pd.Series(dtype=object)), start=2):
                    _bf_fill = ["LienholderId"] if "(LienholderId)" in str(_reason) else ["CollectorId", "LienType"]
                    _fill_cells(_bf_ws, _bf_df, _bf_fill, rows=[_r])
                _add_review_validations(_bf_ws, _bf_df, list_cols={"LienType": LIEN_TYPE_OPTIONS},
                                        whole_cols=("CollectorId", "LienholderId"))
            if not pending_lien_match_df.empty:
                _lm_df = _drop_empty_info_cols(pending_lien_match_df, REVIEW_INFO_COLS)
                _lm_df.to_excel(review_writer, sheet_name="Pending_Lien_Match_Review", index=False)
                _fill_cells(review_writer.sheets["Pending_Lien_Match_Review"], _lm_df, ["Decision"])
                _add_review_validations(review_writer.sheets["Pending_Lien_Match_Review"], _lm_df,
                                        list_cols={"Decision": ["YES", "NO"]})
            if _NOTHING_TO_UPLOAD:
                ev_df.to_excel(review_writer, sheet_name="Input_File_Original", index=False)
                _highlight_note_cols(review_writer.book["Input_File_Original"], ev_df, [_REMARKS_COL],
                                     REMARKS_COL_COLOR)
        print(f"\n📋 Review file (items needing manual attention): {REVIEW_FILE}")
        _print_error_action_needed(pending_error_df)
        _print_plan_text_action_needed(pending_plan_text_df)
        if not pending_tpa_df.empty:
            print("\n" + "=" * 72)
            print(f" ACTION NEEDED: Pending_TPA_Review -- {len(_tpa_review_df)} row(s), "
                  f"{int((_tpa_review_df['Fill_Here'] == 'YES').sum())} to fill")
            print(' 1. Filter column "Fill_Here" to YES (one row per plan + Tort + Scope).')
            print(" 2. On those rows fill CollectorId and LienType (and LienholderId if blank).")
            print("    LienholderId is pre-filled from LPM where found. If it is wrong,")
            print("    change it on every row of that plan.")
            print(" 3. Leave the other rows as they are; on resume they get the same values.")
            print(" 4. Do not edit MCO_Name, Project_Name or Scope_Classification.")
            if (_tpa_review_df.get("PendingReason", pd.Series(dtype=object)) == BLANK_TORT_REASON).any():
                print(f' 5. Rows with PendingReason "{BLANK_TORT_REASON}": fill TortId and TortName')
                print("    (one row per case). CollectorId/LienholderId/LienType there are optional.")
            print("=" * 72)
        if not pending_lien_match_df.empty:
            print("\n" + "=" * 72)
            print(f" ACTION NEEDED: Pending_Lien_Match_Review -- {len(pending_lien_match_df)} row(s)")
            print(' 1. For each plan, write YES in "Decision" on the one lien to use.')
            print(" 2. Write NO only if a lien is not this plan at all.")
            print(" 3. Leave the other rows blank. ReviewReason says why the tool could not choose.")
            print("=" * 72)

    if BLANK_SCOPE_CASES:
        print(f"\nScope is blank in LPM for these cases: {', '.join(dict.fromkeys(BLANK_SCOPE_CASES))}")

    _divider()
    # The run's one save question; chosen decisions land in the lists below, so they
    # must be combined only AFTER it (equivalence saves only ever come from a resume).
    _ask_deferred_saves()
    _SCOPE_SAVES = _RESUME_SCOPE_SAVES + _SCOPE_SAVES
    _TPA_DIRTY = _RESUME_TPA_DIRTY + _TPA_DIRTY
    _LIEN_EQUIV_SAVES = _RESUME_LIEN_EQUIV_SAVES
    save_config_updates(CONFIG_FILE, _SCOPE_SAVES, _TPA_DIRTY, _ALIAS_SAVES, _LIEN_EQUIV_SAVES, _COLUMN_MAP_SAVES)

    # Local, readable copy of every config/TPA decision recorded this run (same rows
    # as mco_config.xlsx's centralized Audit_Trail) -- a convenience, not a
    # replacement: lets you see "what did THIS run decide" without digging through
    # the full terminal transcript in log_{...}.txt.
    if _AUDIT_TRAIL_THIS_RUN:
        _audit_wb_local = _opxl.Workbook()
        _audit_ws_local = _audit_wb_local.active
        _audit_ws_local.title = "Audit_Trail"
        for _i, _col_name in enumerate(_AUDIT_TRAIL_LOCAL_COLUMNS, start=1):
            _audit_ws_local.cell(row=1, column=_i, value=_col_name)
        for _r, _entry in enumerate(_AUDIT_TRAIL_THIS_RUN, start=2):
            for _i, _col_name in enumerate(_AUDIT_TRAIL_LOCAL_COLUMNS, start=1):
                _audit_ws_local.cell(row=_r, column=_i, value=_entry.get(_col_name))
        # Real bug found: this was the only workbook save in the whole file not
        # wrapped in _retry_on_permission_error -- a locked local audit copy (e.g.
        # already open in Excel) would crash the run right at the very end, after
        # everything else already succeeded, instead of offering a retry like every
        # other write in this script.
        _retry_on_permission_error(lambda: _audit_wb_local.save(_run_audit_path), "this run's local audit copy")
        print(f"📝 This run's config/TPA decisions also recorded in: {_run_audit_path}")

    print("\n✅ ===============================")
    print("✅ AUTOMATION COMPLETED SUCCESSFULLY")
    if _NOTHING_TO_UPLOAD:
        print(f"✅ No WIP file this round (nothing to upload). Review file: {REVIEW_FILE}")
    else:
        print(f"✅ Output file: {OUTPUT_FILE}")
    print("✅ ===============================")

    # By this point TicketId is always fully resolved (fresh run: typed upfront;
    # resume run: recovered from the resume file earlier in BUILD DF) -- so this is the
    # first point in a successful run where the log backup's filename can carry the
    # real ticket number, not a placeholder.
    _backup_log_file(_run_log_path, TicketId, CONFIG_FILE)

_offer_rerun("\nPress 'R' and Enter to process another file, or just press Enter to exit: ")
