# %% [Cell 1] Setup -- imports and SQL connection
import pyodbc
import pandas as pd
from datetime import datetime
import re as re
import os as os
import sys
import shutil
import atexit
import subprocess
import traceback
import textwrap
import warnings
from pathlib import Path
from dotenv import load_dotenv

# Real crash found testing the compiled .exe: when stdout isn't a normal interactive
# console (piped/redirected, or some Windows setups even for a plain double-clicked
# .exe), Python falls back to the system's ANSI codepage (cp1252 on this machine)
# instead of UTF-8 -- and this script prints emoji (✅❌⚠️📝 etc.) constantly. The very
# FIRST such print then raises UnicodeEncodeError, which is bad enough on its own, but
# it ALSO took down the crash handler itself the same way when it tried to report that
# very error -- defeating the whole point of having one. Reconfiguring stdout/stderr
# here forces real UTF-8 regardless of the console's own codepage; _TeeStream.write()
# below ALSO defensively falls back on a replace-errors re-encode, so even an
# unexpected residual encoding issue can never crash the crash handler again.
# stdin reconfigured too -- found right after fixing the above: piped/redirected input
# containing a UTF-8 BOM was being mis-decoded through the wrong codepage into 3
# garbled characters (not the real BOM codepoint), which no amount of stripping the
# real BOM character afterward could ever catch, since by then it was already corrupted.
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

# DB credentials and the shared-drive config/TPA file paths used to be hardcoded
# directly in this file (a real risk: this file gets shared/reviewed/screen-shared,
# and the DB password was sitting in it in plain text). Now read from a .env file
# kept next to this script (never shared/emailed/backed up anywhere).
# sys.frozen is True only inside a PyInstaller-compiled .exe -- __file__ there points
# into a temporary extraction folder, not the real .exe's own location, so .env would
# never be found once compiled. sys.executable is the .exe's real path in that case.
# __file__ isn't always defined in every way this script can be run even outside a
# frozen exe (e.g. some interactive cell-execution contexts), so fall back to
# load_dotenv()'s own default search (current working directory) if that happens too.
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

# MCO_CONFIG_FILE/MCO_TPA_FILE in .env are now a small RELATIVE tail (e.g. "Data Team -
# India - Documents/MCO Tool and Config/mco_config.xlsx") shared identically by every team
# member, not a full absolute path -- the SharePoint library is synced locally via
# OneDrive, and only the sync-root PREFIX before that tail differs per Windows user. No
# username ever needs to be typed/replaced anywhere -- Python resolves the current
# user's own root automatically.
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
# RUN LOG -- full terminal transcript (every print AND every input prompt/answer),
# so a past run's exact sequence of events can be reviewed later without relying on
# anyone's memory or a screenshot. Captured from this point on, before the very first
# input() prompt the user sees -- buffered in memory until OUTPUT_FILE's folder is
# known (a few prompts in), then flushed to a real file there and written live after.
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
            # Real gap: a crash before OUTPUT_FILE's folder is known (e.g. the DB
            # connection or the import-file read itself failing) means attach_file()
            # never ran -- everything printed so far, including the crash traceback
            # itself, was only ever buffered in memory. Without this, close() here
            # would just silently drop it, defeating the whole point of the run log for
            # exactly the failures most likely to need it. Falls back to a dated file
            # next to the script/exe (the one location always known this early).
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
    answer = _real_input()
    _RUN_LOG_STREAM.write_log_only(answer + "\n")
    return answer

import builtins
builtins.input = _logged_input

# =====================================================
# CRASH HANDLING + "RUN AGAIN?" -- so a double-clicked .exe never just vanishes on an
# error (console windows close immediately on an uncaught exception otherwise, before
# anyone can read why), and so the same window can process another file without
# re-launching the exe by hand. '_restart_fresh' relaunches this exact script/exe as a
# brand-new process -- same console window, zero chance of resuming stale/half-crashed
# state, DB connection and everything else reconnects from a clean slate. Note: inside
# an interactive dev console (e.g. Spyder/IPython), that tool installs its OWN
# sys.excepthook for richer tracebacks -- this one only takes effect running as a plain
# script or a compiled .exe, exactly where it's actually needed.
# =====================================================
def _restart_fresh():
    _RUN_LOG_STREAM.close()
    try:
        conn.close()
    except Exception:
        pass
    # sys.argv is NOT reliable here -- how this script was actually launched (a plain
    # terminal, VS Code's "Run Python File", a debugger, etc.) can leave sys.argv not
    # reflecting a real re-runnable command. __file__ is this script's own real,
    # guaranteed-correct path regardless of how it was launched -- building the relaunch
    # command from that instead is robust to it.
    #
    # os.execv was tried first and verified, empirically, to NOT be reliable on Windows
    # for this: it either hung indefinitely or dropped into a broken interactive REPL
    # instead of actually re-running the script (reproduced in a plain terminal, nothing
    # to do with any particular editor/IDE).
    #
    # subprocess.Popen (fire-and-forget) + an immediate os._exit() was tried next, and
    # DOES correctly spawn a fresh process -- but real-world use (user report) showed a
    # console-ownership race: the terminal (e.g. PowerShell) only ever waited on THIS
    # process, so the instant it exits, the terminal reclaims the console and reprints
    # its own prompt, even though the new child process is simultaneously still starting
    # up and about to read its own input() from that SAME console -- both are briefly
    # attached to the same console at once, racing for keystrokes (observed: the shell's
    # own prompt resurfacing mid-restart, with pasted input showing the shell's syntax
    # highlighting instead of plainly going to the new process). subprocess.run() BLOCKS
    # this process until the new one (and whatever further restarts IT triggers) fully
    # finishes, so the terminal never reclaims the console until the entire chain is
    # truly done -- verified empirically (a 4-deep restart chain, each parent correctly
    # waiting for its child before exiting, unwinding cleanly once the innermost one
    # finished, no hang).
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
    traceback.print_exception(exc_type, exc_value, exc_tb)
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

# PLAN IDENTIFICATION's own banner is lazy, unlike every other phase's (called once,
# eagerly, before its loop starts) -- real bug found (user report): a run where every
# combo resolves silently (a direct/override TPA match, or a combo that's auto-routed
# straight to a Pending_* sheet with nothing to ask) still showed the banner followed
# by total silence, looking like the phase never ran at all. Only called at the
# handful of points in the Plan Identification loop that are guaranteed to lead to an
# actual input() this run -- never at a purely informational or auto-pending point.
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

def _file_fingerprint(path):
    """(mtime, size) together -- more robust than mtime alone for detecting whether a
    shared file genuinely changed since it was loaded. These files live in an
    OneDrive-synced SharePoint folder, and OneDrive's sync client can touch a file's
    metadata (a background re-hash/conflict-check pass) without any real content
    change -- mtime alone would then falsely trigger the "someone else changed this"
    concurrent-save check elsewhere, deferring a perfectly safe save for no reason. A
    REAL save always changes size too (even one new row/cell adds bytes); a pure
    metadata touch does not -- requiring both to differ closes that gap."""
    return (os.path.getmtime(path), os.path.getsize(path))

def _fingerprint_really_changed(old_fp, new_fp):
    """True only when BOTH mtime and size differ from the loaded fingerprint -- a pure
    metadata touch (mtime moves, size unchanged) must NOT count as a real external
    change. A plain tuple != would wrongly trigger on either one alone differing,
    which defeats the whole point of _file_fingerprint."""
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
    Lives in a 'Config_Backups' subfolder derived from config_path itself (never a
    separate hardcoded location) -- so if CONFIG_FILE's own path ever changes (e.g. a
    future move to a SharePoint-synced folder), the backup location follows it
    automatically with zero code change. Same-day reruns overwrite today's backup
    (filename has no time component) instead of piling up, and anything older than
    CONFIG_BACKUP_RETENTION_DAYS is deleted here too, so this folder never grows
    unbounded. Never allowed to block the real save -- any failure here is swallowed
    with a warning."""
    try:
        backup_dir = os.path.join(os.path.dirname(config_path), "Config_Backups")
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
    audit_ws = audit_wb["Audit_Trail"]
    _ensure_audit_trail_header(audit_ws)
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

Connection_String = (
    "DRIVER={ODBC Driver 17 for SQL Server};"
    f"SERVER={os.environ['DB_SERVER']};"
    f"DATABASE={os.environ['DB_DATABASE']};"
    f"UID={os.environ['DB_UID']};"
    f"PWD={os.environ['DB_PWD']};"
)

conn = pyodbc.connect(Connection_String)
print("✅ SQL connection established successfully")

# %% [Cell 2] Main pipeline
# =====================================================
# USER INPUTS
# =====================================================

_divider()
MCO_INPUT_FILE = input("Path of Import File: ").lstrip("﻿").strip().strip('"')

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

KNOWN_PENDING_SHEETS = ["Pending_Scope_Review", "Pending_TPA_Review", "Pending_TPA_Blank_Fields", "Pending_Lien_Match_Review"]
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

def is_valid_date_mmddyyyy(text):
    try:
        datetime.strptime(text, "%m-%d-%Y")
        return True
    except ValueError:
        return False

# Auto-detection from the import file's OWN name (2026-09-25): the file's title
# already carries State/Response Type/Rec'd Date in a consistent shape, so these are
# tried FIRST -- only whichever piece can't be found this way still gets its own
# upfront prompt. Deliberately tolerant (never raises) -- a filename that doesn't
# follow the convention just falls through to the normal prompt for that piece,
# exactly as if auto-detection had never been attempted.
def _parse_state_from_filename(name):
    """The 2-letter code immediately before the word 'Mcaid' (any case), e.g.
    '...MI Mcaid...' -> 'MI'. None if that anchor word isn't found at all."""
    try:
        m = re.search(r'\b([A-Za-z]{2})\s*Mcaid\b', name, re.IGNORECASE)
        return m.group(1).upper() if m else None
    except Exception:
        return None

def _parse_response_type_from_filename(name):
    """Whichever of the 3 known response types sits immediately before the word
    'Response' (any case) -- the longer 'EV & Claims' alternative is checked before
    the bare 'EV'/'Claims' ones so it isn't partially matched. None if 'Response'
    isn't found at all, or nothing recognizable sits right before it."""
    try:
        m = re.search(r'\b(EV\s*&\s*Claims|EV|Claims)\s*Response', name, re.IGNORECASE)
        if not m:
            return None
        found = re.sub(r'\s*&\s*', ' & ', m.group(1).strip()).upper()
        for opt in RESPONSE_TYPE_OPTIONS:
            if found == opt.upper():
                return opt
        return None
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
    TicketId = input("TicketId: ").strip()

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
        STATE = input("State (2-letter code): ").strip().upper()

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

if len(OUTPUT_FILE) >= 260:
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
# STATE is now known upfront (parsed from the import file's own name, or asked
# directly if that failed) -- LA-format processing is simply "this run's State is
# Louisiana", no more need to go sniff sheet names for an Entitlement/EV tab pair.
# On resume, STATE is None and this never matters anyway (that branch is chosen
# before IS_LA_FILE is ever consulted), so it's left False exactly as before.
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

def _patch_tpa_file_blank_fields(state_mco, tort_id, lienholder_id, collector_id, lien_type):
    """Patches ONE blank-field row in the TPA file in place, locating it by
    (TortId, LienholderId) within the state's own sheet (matched by state-code
    prefix, same convention as load_tpa_state_sheet) -- this is the single source
    of truth this case fixes: the row already exists correctly identified in
    TPA, it's just missing CollectorId/LienType. Shared by both the resumed
    Pending_TPA_Blank_Fields batch-patch and the live in-run prompt, so both
    write via the exact same logic instead of two copies that could drift.
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
    tpa_header = [c.value for c in tpa_ws[1]]
    tpa_col_idx = {name: i + 1 for i, name in enumerate(tpa_header)}
    found_row = None
    for r in range(2, tpa_ws.max_row + 1):
        if (str(tpa_ws.cell(row=r, column=tpa_col_idx["TortId"]).value) == str(tort_id) and
                _normalize_key(tpa_ws.cell(row=r, column=tpa_col_idx["LienholderId"]).value) == _normalize_key(lienholder_id)):
            found_row = r
            break
    if found_row is None:
        return False
    _old_collector_id = tpa_ws.cell(row=found_row, column=tpa_col_idx["CollectorId"]).value
    _old_lien_type = tpa_ws.cell(row=found_row, column=tpa_col_idx["LienType"]).value
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
        f"CollectorId={_old_collector_id if pd.notna(_old_collector_id) else '(blank)'}, "
        f"LienType={_old_lien_type if pd.notna(_old_lien_type) else '(blank)'}",
        f"CollectorId={collector_id if pd.notna(collector_id) else _old_collector_id}, "
        f"LienType={lien_type if pd.notna(lien_type) else _old_lien_type}",
        TicketId,
    )])
    return True

TPA_OVERRIDES_COLS = ["TortId", "TortName", "Private Lien Case Scope", "CollectorId",
                       "LienholderId", "LienholderName", "LienType",
                       "State MCO", "Ticket #"]

def load_config(config_path):
    # One workbook parse serves all 5 sheets below -- pd.read_excel(config_path, ...)
    # re-parses the ENTIRE file from scratch on every call when given a bare path, so
    # calling it 5x (once per sheet) was silently doing 5 full workbook reads instead
    # of 1. Passing this shared ExcelFile object instead reuses the one parse.
    _ensure_file_hydrated(config_path)
    xls = _retry_on_permission_error(lambda: pd.ExcelFile(config_path), "mco_config.xlsx")

    def _load_sheet(sheet_name, col_keys):
        df = _retry_on_permission_error(
            lambda: pd.read_excel(xls, sheet_name=sheet_name, dtype=str), "mco_config.xlsx"
        )
        df.columns = df.columns.str.strip()
        result = {}
        for _, row in df.iterrows():
            vals = [str(row[k]).strip().upper() for k in col_keys]
            if any(v == "NAN" for v in vals):
                continue
            state = vals[0]
            if len(state) != 2:
                continue
            result.setdefault(state, {})[vals[1]] = vals[2]
        return result

    def _load_scope_map(sheet_name):
        df = _retry_on_permission_error(
            lambda: pd.read_excel(xls, sheet_name=sheet_name, dtype=str), "mco_config.xlsx"
        )
        df.columns = df.columns.str.strip()
        result = {}
        for _, row in df.iterrows():
            key = _normalize_key(row.get("ScopeText"))
            if key is None:
                continue
            result[key] = _normalize_key(row.get("Classification"))
        return result

    def _load_tpa_overrides(sheet_name):
        df = _retry_on_permission_error(
            lambda: pd.read_excel(xls, sheet_name=sheet_name, dtype=str), "mco_config.xlsx"
        )
        df.columns = df.columns.str.strip()
        rows = []
        for _, row in df.iterrows():
            if pd.isna(row.get("LienholderName")) and pd.isna(row.get("LienholderId")):
                continue
            rows.append({c: (row.get(c) if pd.notna(row.get(c)) else None)
                         for c in TPA_OVERRIDES_COLS})
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
        df = _retry_on_permission_error(
            lambda: pd.read_excel(xls, sheet_name=sheet_name, dtype=str), "mco_config.xlsx"
        )
        df.columns = df.columns.str.strip()
        equivalence_map, excluded_map = {}, {}
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
            target = equivalence_map if decision_key == "YES" else excluded_map
            # Keyed by (State, CollectorId, LienholderId) -- NOT just (State,
            # LienholderId). A LienholderId number can in principle be reused by TPA
            # across more than one Collector relationship (different real plans that
            # happen to share a number) -- scoping to the exact Collector this
            # confirmation was made under means a coincidental same-text name clash
            # under a DIFFERENT Collector is never silently treated as the same plan.
            key = (str(state_mco).strip(), str(collector_id).strip(), str(lienholder_id).strip())
            target.setdefault(key, [])
            if str(equiv_name).strip() not in target[key]:
                target[key].append(str(equiv_name).strip())
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
        df = _retry_on_permission_error(
            lambda: pd.read_excel(xls, sheet_name=sheet_name, dtype=str), "mco_config.xlsx"
        )
        df.columns = df.columns.str.strip()
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

# Same shape as _RESUME_TPA_DIRTY -- a "Yes"/"New" decision on a resumed
# Pending_Lien_Match_Review row is applied to LIEN_NAME_EQUIVALENCE_MAP/
# LIEN_NAME_EXCLUDED_MAP unconditionally (needed for THIS run's own
# classify_step4_scenario calls to resolve it); this queue is only what
# additionally gets persisted to mco_config.xlsx if the user says so.
_RESUME_LIEN_EQUIV_SAVES = []

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
# Semicolon is a plan separator, same as comma/newline/double-space. Single source
# of truth for "what separates multiple plans in free text with no dates" -- used
# in both places extract_plan_rows falls back to a plain split (the no-plan_matches-
# at-all path, and the leftover-after-dated-plans path) so they can't drift apart.
_RE_PLAN_SPLIT     = re.compile(r'[,;\n]+|\s{2,}')

_PAT_WITH_PLAN     = re.compile(r'(?:AND\s+)?([A-Z][A-Z\s]*?)\s*[-:]?\s*(\d{1,2}/\d{1,2}/\d{4})\b\s*(?:TO|-|\*)\s*(\d{1,2}/\d{1,2}/\d{4}|CURRENT|PRESENT)\b')
_PAT_DATE_ONLY     = re.compile(r'(\d{1,2}/\d{1,2}/\d{4})\b\s*(?:TO|-|\*)\s*(\d{1,2}/\d{1,2}/\d{4}|CURRENT|PRESENT)\b')
_PAT_SIDE_BY_SIDE  = re.compile(r'(?:AND\s+)?([A-Z][A-Z\s]*?)\s*[-:]?\s*(\d{1,2}/\d{1,2}/\d{4})\b\s+(\d{1,2}/\d{1,2}/\d{4})\b')
_PAT_DATE_SBS      = re.compile(r'(\d{1,2}/\d{1,2}/\d{4})\b\s+(\d{1,2}/\d{1,2}/\d{4})\b')
_PAT_WITH_PLAN_ANY = re.compile(r'(?:AND\s+)?([A-Z][A-Z\s]*?)\s*[-:]?\s*(\d{1,2}/\d{1,2}/\d+)\s*(?:TO|-|\*)\s*(\d{1,2}/\d{1,2}/\d+|CURRENT|PRESENT)')
_PAT_DATE_ANY      = re.compile(r'(\d{1,2}/\d{1,2}/\d+)\s*(?:TO|-|\*)\s*(\d{1,2}/\d{1,2}/\d+|CURRENT|PRESENT)')

# =====================================================
# INVALID SET
# =====================================================

INVALID = {
    "N", "NO", "NA", "NONE", "N/A", "A", "CASE", "PD", "PAID", "CLAIMS",
    "RELATED", "NOT", "TMHP", "COLLECTS", "FOR", "THIS", "DOL"
}

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

def safe_min(dates):
    valid = [d for d in dates if d is not None and not isinstance(d, str)]
    if valid:
        return min(valid)
    if any(isinstance(d, str) for d in dates if d is not None):
        return "Bad Date"
    return None

def safe_max(dates):
    valid = [d for d in dates if d is not None and not isinstance(d, str)]
    if valid:
        return max(valid)
    if any(isinstance(d, str) for d in dates if d is not None):
        return "Bad Date"
    return None

def format_date(x):
    if x is None:
        return None
    if isinstance(x, str):
        return "Bad Date"
    try:
        if pd.isna(x):
            return None
    except Exception:
        pass
    return x.strftime("%m/%d/%Y")

def make_sql_in(id_list):
    unique = list(dict.fromkeys(str(i) for i in id_list))
    if len(unique) == 1:
        return f" ('{unique[0]}')"
    return f" {tuple(unique)}"

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
    df1 = pd.read_sql(q1, conn)
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
        df2 = pd.read_sql(q2, conn, params=[person_id, f"%{casename_trunc}%"])
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
    p = _RE_MULTI_SPACE.sub(' ', raw).strip()
    p = re.sub(r'^AND\s*', '', p).strip()
    p = re.sub(r'^[\.\,\s]+', '', p).strip()
    p = re.sub(r'^\d+[\.\)]\s*', '', p).strip()
    p = re.sub(r'[-:]+$', '', p).strip()
    return p

def _is_valid_plan(plan: str) -> bool:
    if not plan:
        return False
    if not _RE_ALPHA_IN_PLAN.search(plan):
        return False
    if set(plan.split()).issubset(INVALID):
        return False
    return True

def _parse_date(s: str):
    try:
        return datetime.strptime(s, "%m/%d/%Y")
    except Exception:
        return "Bad Date"

def extract_plan_rows(text):
    if pd.isna(text):
        return []

    text = str(text).upper()
    text = _RE_NY_PREFIX.sub('', text.strip())
    text = _RE_AMP.sub(',', text)
    text = _RE_AND_WORD.sub(',', text)
    text = _RE_LONG_NUM.sub('', text)
    text = _RE_NON_DATE_SLASH.sub(',', text)

    all_dates = []
    for m in _PAT_DATE_ONLY.finditer(text):
        all_dates.append((m.start(), m.group(1), m.group(2)))
    for m in _PAT_DATE_SBS.finditer(text):
        if not any(abs(m.start() - pos) < 5 for pos, *_ in all_dates):
            all_dates.append((m.start(), m.group(1), m.group(2)))
    for m in _PAT_DATE_ANY.finditer(text):
        if not any(abs(m.start() - pos) < 5 for pos, *_ in all_dates):
            all_dates.append((m.start(), m.group(1), m.group(2)))
    all_dates.sort(key=lambda x: x[0])

    plan_matches = []
    for m in _PAT_WITH_PLAN.finditer(text):
        plan = _clean_plan(m.group(1))
        if _is_valid_plan(plan):
            plan_matches.append((m.start(), plan, m.group(2), m.group(3)))
    for m in _PAT_SIDE_BY_SIDE.finditer(text):
        plan = _clean_plan(m.group(1))
        if _is_valid_plan(plan):
            if not any(abs(m.start() - pos) < 5 for pos, *_ in plan_matches):
                plan_matches.append((m.start(), plan, m.group(2), m.group(3)))
    for m in _PAT_WITH_PLAN_ANY.finditer(text):
        plan = _clean_plan(m.group(1))
        if _is_valid_plan(plan):
            if not any(abs(m.start() - pos) < 5 for pos, *_ in plan_matches):
                plan_matches.append((m.start(), plan, m.group(2), m.group(3)))
    plan_matches.sort(key=lambda x: x[0])

    if not plan_matches:
        results = []
        seen = set()
        for part in _RE_PLAN_SPLIT.split(text):
            plan = _clean_plan(part)
            if not _is_valid_plan(plan) or '-' in plan or plan in seen:
                continue
            seen.add(plan)
            results.append((plan, None, None))
        return results

    plan_positions = [p[0] for p in plan_matches]
    plan_dict = {}
    for pos, start, end in all_dates:
        assigned = None
        for i, p_pos in enumerate(plan_positions):
            if p_pos <= pos:
                assigned = plan_matches[i][1]
            else:
                break
        if assigned:
            plan_dict.setdefault(assigned, []).append((start, end))

    results = []
    for plan, date_ranges in plan_dict.items():
        starts = [_parse_date(s) for s, _ in date_ranges]
        ends   = [_parse_date(e) for _, e in date_ranges]
        results.append((plan, safe_min(starts), safe_max(ends)))

    cleaned = _PAT_WITH_PLAN.sub('', text)
    cleaned = _PAT_DATE_ONLY.sub('', cleaned)
    seen_plans = {p[0] for p in results}
    for part in _RE_PLAN_SPLIT.split(cleaned):
        plan = _clean_plan(part)
        if not _is_valid_plan(plan) or '-' in plan or plan in seen_plans:
            continue
        seen_plans.add(plan)
        results.append((plan, None, None))
    for _, plan, *_ in plan_matches:
        if plan not in seen_plans:
            seen_plans.add(plan)
            results.append((plan, None, None))

    return results

# =====================================================
# LA FILE PROCESSOR
# =====================================================

def process_la_file(filepath):

    wb  = _retry_on_permission_error(
        lambda: _opxl.load_workbook(filepath, data_only=True, read_only=True), "the import file"
    )
    tabs = wb.sheetnames
    wb.close()

    # Real gap found (user report): if the file's actual tab/column/row naming
    # doesn't match what's expected here (a modified or genuinely different
    # file), these used to silently resolve to None and crash confusingly much
    # later (e.g. sheet_name=None means something else entirely to
    # pd.read_excel). Each of the 5 auto-detects below (2 tabs, 2 SSN columns,
    # the header row, the Healthy Louisiana column start) now falls back to
    # _ask_detected_fallback on failure -- a caution (most likely either the
    # wrong file was fed in, or the structure has genuinely changed) plus a
    # manual pick from what's actually in the file, or 'C' to cancel outright.
    entitlement_tab = next((s for s in tabs if "ENTITLEMENT" in s.upper()), None)
    ev_tab          = next((s for s in tabs if any(w in s.upper() for w in ["EV", "REQUEST", "RESPONSE"])), None)
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

    # EV Tab — SSN → Archer ID
    ev_raw = _retry_on_permission_error(
        lambda: pd.read_excel(filepath, sheet_name=ev_tab, dtype=str), "the import file"
    )
    ev_raw.columns = ev_raw.columns.str.strip()

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
    ev_raw[ssn_ev_col] = ev_raw[ssn_ev_col].astype(str).str.strip()
    ev_raw[archer_col] = ev_raw[archer_col].astype(str).str.strip()
    ssn_to_lienid      = dict(zip(ev_raw[ssn_ev_col], ev_raw[archer_col]))

    # Mismatch SSNs -- tracked purely as an informational flag now (see
    # mismatch_flagged_lienids below), never to exclude a row from processing.
    mismatch_ssns = set()
    if mismatch_tab:
        mis_df = _retry_on_permission_error(
            lambda: pd.read_excel(filepath, sheet_name=mismatch_tab, dtype=str), "the import file"
        )
        mis_df.columns = mis_df.columns.str.strip()
        ssn_mis_col = next((c for c in mis_df.columns if "SSN" in str(c).upper()), None)
        if ssn_mis_col:
            mismatch_ssns = set(mis_df[ssn_mis_col].dropna().astype(str).str.strip().unique())

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

    ent_df = _retry_on_permission_error(
        lambda: pd.read_excel(filepath, sheet_name=entitlement_tab, header=header_row, dtype=str), "the import file"
    )
    ent_df.columns = ent_df.columns.str.strip()

    ssn_ent_col = next((c for c in ent_df.columns if "SSN" in str(c).upper()), None)
    if ssn_ent_col is None:
        _, ssn_ent_col = _ask_detected_fallback(
            list(ent_df.columns), f"Could not find an SSN column in the {entitlement_tab!r} tab.",
            "Which column is SSN? Enter the number: "
        )
    ent_df[ssn_ent_col] = ent_df[ssn_ent_col].astype(str).str.strip()

    # Processing set -- a Mismatch-tab SSN no longer excludes the row from processing
    # (every MCO plan on file should be processed, per policy; the Mismatch flag is
    # tracked separately below purely as an informational note, never as a reason to
    # skip). No Y/N-entitlement filter here either: a row with an actual plan name
    # should still be looked at regardless of that column's value -- same simple rule
    # normal files use (does this column actually have a plan name in it or not).
    ent_df_for_processing = ent_df.copy()

    # Display set (Input_File_Original) -- the FULL raw Entitlement tab, with LienId
    # merged in via SSN so the Remarks column can attach to the right row for every
    # claimant.
    ent_df_display = ent_df.copy()
    ent_df_display.insert(0, "LienId", ent_df_display[ssn_ent_col].map(ssn_to_lienid))
    mismatch_flagged_lienids = {ssn_to_lienid[ssn] for ssn in mismatch_ssns if ssn in ssn_to_lienid}

    # Sirf HL Plan columns
    all_cols       = list(ent_df_for_processing.columns)
    plan_name_cols = [c for c in all_cols if "PLAN" in str(c).upper()  and "NAME"  in str(c).upper() and all_cols.index(c) >= hl_col_start]
    start_cols     = [c for c in all_cols if "START" in str(c).upper() and "MONTH" in str(c).upper() and all_cols.index(c) >= hl_col_start]
    end_cols       = [c for c in all_cols if "END"   in str(c).upper() and "MONTH" in str(c).upper() and all_cols.index(c) >= hl_col_start]

    def parse_la_date(d):
        if not d or str(d).upper() in {"NAN", ""}:
            return None
        d_str = str(d).strip()
        try:
            return datetime.strptime(d_str[:10], "%Y-%m-%d")
        except:
            try:
                return datetime.strptime(d_str, "%m/%d/%Y")
            except:
                # LA "Start/End Month" sometimes arrives as a bare YYYYMM number
                # ("201601") instead of a real date -- Excel can also store it as a
                # genuine number, which pd.read_excel(dtype=str) then renders with a
                # trailing ".0" ("201601.0"). Only reached once both real-date
                # formats above have already failed, so a genuine date is never
                # reinterpreted -- day is assumed to be the 1st since none is given.
                _digits = re.sub(r"\.0+$", "", d_str)
                if re.fullmatch(r"\d{6}", _digits):
                    year, month = int(_digits[:4]), int(_digits[4:6])
                    if 1 <= month <= 12:
                        try:
                            return datetime(year, month, 1)
                        except ValueError:
                            pass
                return "Bad Date"

    rows = []
    for _, row in ent_df_for_processing.iterrows():
        ssn     = str(row[ssn_ent_col]).strip()
        lien_id = ssn_to_lienid.get(ssn)
        if not lien_id or lien_id == "nan":
            continue
        for plan_col, start_col, end_col in zip(plan_name_cols, start_cols, end_cols):
            plan  = str(row.get(plan_col, "")).strip()
            start = str(row.get(start_col, "")).strip()
            end   = str(row.get(end_col,   "")).strip()
            if not plan or plan.upper() in {"NAN", "N", "NO", "NA", "NONE"}:
                continue
            rows.append({
                "LienId"    : lien_id,
                "MCO_Name"  : plan.upper(),
                "Start_Date": parse_la_date(start),
                "End_Date"  : parse_la_date(end)
            })

    return pd.DataFrame(rows), ent_df_display, "LienId", mismatch_flagged_lienids

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

# Populated by resolve_tpa_for_run whenever the user picks a plan from the
# name_not_found candidate list AND the raw input text genuinely differs from that
# plan's real TPA name -- (State, From_Name, To_Name) triples for Alias_Map, so the
# NEXT time this exact raw text appears (for ANY Tort/Scope, not just this one) it gets
# rewritten to the canonical name before TPA lookup even runs, instead of needing its
# own fresh TPA_Overrides resolution every time a different Tort sees the same raw text.
_ALIAS_SAVES = []

# A client's plans can split across runs (some resolve now, some go to a Pending_*
# sheet). Client-MCO Cnt Update must count the client's REAL total, not just this
# run's own df_in, or it undercounts and re-sends a duplicate NewClientNote. Carried
# on each Pending_* row as "_ClientAlreadyCountedThisRun" (0 if none yet), read back
# here by ClientId and combined with this run's own resolved count in
# build_client_cnt_update below. Stays {} on a non-resume run.
_ALREADY_COUNTED_CARRY = {}

# Real bug found (user report): build_step4_outputs' own claimed_lien_ids only ever
# prevented TWO plans in the SAME run from independently matching the SAME
# existing/placeholder lien -- it resets fresh every run. A client's OTHER plan that
# went to Pending_* this run (not yet resolved) never got a chance to be seen by that
# tracking at all. On resume, the real DB still shows that SAME placeholder lien as
# unclaimed (this tool never writes to the DB itself -- the earlier round's MCO
# Update 2 instructions haven't been manually applied to Ragic yet), so the
# resumed plan matched it AGAIN -- two different plans ending up pointed at the one
# real lien. Carried the same way as _ALREADY_COUNTED_CARRY, just per-client LienIds
# instead of a count: "_ClientAlreadyClaimedLienIds" (comma-separated) on each
# Pending_* row, read back here, and excluded alongside this run's own
# claimed_lien_ids in build_step4_outputs below.
_ALREADY_CLAIMED_LIEN_IDS_CARRY = {}

if IS_RESUME_RUN:
    _resume_frames = []
    for _sheet in RESUME_SHEETS_PRESENT:
        _pending_raw = _retry_on_permission_error(
            lambda: pd.read_excel(MCO_INPUT_FILE, sheet_name=_sheet, dtype=str), "the import file"
        )
        _pending_raw["_source_sheet"] = _sheet
        _resume_frames.append(_pending_raw)
    _resume_combined = pd.concat(_resume_frames, ignore_index=True)

    if "_ClientAlreadyCountedThisRun" in _resume_combined.columns:
        _carry_raw = pd.to_numeric(_resume_combined["_ClientAlreadyCountedThisRun"], errors="coerce").fillna(0)
        for _cid, _carry_val in zip(_resume_combined["ClientId"], _carry_raw):
            _ALREADY_COUNTED_CARRY[_cid] = max(_ALREADY_COUNTED_CARRY.get(_cid, 0), int(_carry_val))

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

    TicketId      = str(_resume_first_value("TicketId") or "RESUME")
    RESPONSE_DATE = _resume_first_value("ResponseDate")
    RESPONSE_TYPE = _resume_first_value("ResponseType")

    _phase_banner("resume", "RESUME PROCESSING")

    # Pending_Scope_Review: the user fills the classification directly into that
    # sheet's existing "Scope_Classification" column (always blank there, since a
    # row only lands there when classify_scope couldn't resolve it). Pre-seed
    # SCOPE_MAP here, BEFORE resolve_scope_for_run's own classify_scope() pass, so
    # these rows resolve with no prompt. Applying to SCOPE_MAP is unconditional
    # (needed for THIS run); the "save for future?" prompt only controls whether it
    # also queues into _RESUME_SCOPE_SAVES -- asked ONCE per unique scope text, not
    # once per row.
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
        if ask_yes_no("Save this decision for future runs? (y/n): "):
            _RESUME_SCOPE_SAVES.append((_raw_scope_text, _resume_classification))

    # Pending_TPA_Review: same shape as above -- applying to TPA_OVERRIDES (in-memory)
    # is unconditional, since that's what lets THIS run's own resolve_tpa_for_run find
    # it via _tpa_overrides_find below instead of re-running the whole cascade; the
    # "save for future?" answer only controls _RESUME_TPA_DIRTY (the persist-to-
    # mco_config.xlsx queue). Asked once per unique resolution (State/Tort/Name/Scope),
    # not once per row, so several clients sharing the identical unresolved plan don't
    # each trigger their own copy of the same question.
    _tpa_review_rows = _resume_combined[_resume_combined["_source_sheet"] == "Pending_TPA_Review"]
    _tpa_review_seen = {}
    _alias_review_seen = set()
    for _, _rrow in _tpa_review_rows.iterrows():
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
            if _tpa_review_key in _tpa_review_seen:
                TPA_OVERRIDES.append(_tpa_review_seen[_tpa_review_key])
                continue
            _resume_tpa_entry = {
                "TortId": _rrow.get("TortId"), "TortName": _rrow.get("Project_Name"),
                "Private Lien Case Scope": _rrow.get("Scope_Classification"),
                "CollectorId": _rrow.get("CollectorId"), "LienholderId": _rrow.get("LienholderId"),
                "LienholderName": _rrow.get("MCO_Name"),
                "LienType": _rrow.get("LienType"),
                "State MCO": _state_mco,
                # The row's OWN originally-stored ticket (written when it first went to
                # Pending_TPA_Review) is more precise than the module-level TicketId --
                # that variable just holds the FIRST ticket found anywhere across this
                # resume file's combined rows (see _resume_first_value above), which is
                # only ever correct by coincidence if a resume file ever mixed rows from
                # more than one original ticket. Falls back to it only if this row
                # somehow has no stored TicketId of its own (an older review file).
                "Ticket #": (str(_rrow.get("TicketId")) if pd.notna(_rrow.get("TicketId")) else TicketId),
            }
            TPA_OVERRIDES.append(_resume_tpa_entry)
            _tpa_review_seen[_tpa_review_key] = _resume_tpa_entry
            _new_item("resume")
            print(
                f"\nResolved from Pending_TPA_Review: {_rrow.get('MCO_Name')!r} (Tort {_rrow.get('Project_Name')!r}) "
                f"-> LienholderId={_resume_tpa_entry['LienholderId']}, CollectorId={_resume_tpa_entry['CollectorId']}, "
                f"LienType={_resume_tpa_entry['LienType']}"
            )
            if ask_yes_no("Save this decision for future runs? (y/n): "):
                # Both needed when saving: TPA_OVERRIDES already has it (for THIS run's
                # own lookups, added above unconditionally); _RESUME_TPA_DIRTY is what
                # actually gets written to mco_config.xlsx -- without it, a resolved
                # Pending_TPA_Review row would work for this run but never persist,
                # forcing the next run to re-ask.
                _RESUME_TPA_DIRTY.append(_resume_tpa_entry)
        else:
            # User resolved this row by writing the real TPA plan name directly
            # (Picked LienholderName) instead of hand-typing LienholderId/CollectorId/
            # LienType. Matches the live number-pick flow's own reasoning (see
            # _resolve_picked_plan): the picked name is already a real, correct TPA
            # row, so a TPA_Overrides entry would be redundant -- and would FREEZE
            # today's CollectorId/LienType even if TPA's own data changes later.
            # Alias_Map alone is enough: rewritten here IN-MEMORY (not just queued for
            # file save) so THIS run's own normalize_mco_name pass (further down)
            # benefits too, not just future runs -- the row then re-enters the normal
            # pipeline and resolves via the ordinary Tort->State->Name->Scope cascade,
            # same as any first-time occurrence of that plan name. If the picked name
            # isn't actually in TPA for this Tort/Scope, that SAME normal cascade's own
            # Pending_TPA_Review/Pending_Scope_Review routing handles it like any other
            # unresolved plan -- no special-case handling needed here.
            _raw_text = str(_rrow.get("MCO_Name") or "").strip()
            _picked_clean = str(_picked_name).strip()
            _resume_state_key = str(_rrow.get("State") or "").strip().upper()
            _alias_key = (_resume_state_key, _normalize_key(_raw_text))
            if _alias_key in _alias_review_seen:
                continue
            _alias_review_seen.add(_alias_key)
            if _normalize_key(_raw_text) != _normalize_key(_picked_clean):
                ALIAS_MAP.setdefault(_resume_state_key, {})[_raw_text.upper()] = _picked_clean.upper()
                _ALIAS_SAVES.append((_resume_state_key, _raw_text, _picked_clean))
            _new_item("resume")
            print(
                f"\nAliased from Pending_TPA_Review: {_raw_text!r} -> {_picked_clean!r} "
                f"(State {_resume_state_key!r}) -- will resolve normally via TPA."
            )

    _blank_field_rows = _resume_combined[_resume_combined["_source_sheet"] == "Pending_TPA_Blank_Fields"]
    for _, _brow in _blank_field_rows.iterrows():
        if pd.isna(_brow.get("CollectorId")) and pd.isna(_brow.get("LienType")):
            continue
        if _patch_tpa_file_blank_fields(
            _brow.get("TPA_Match_StateMCO"), _brow.get("TPA_Match_TortId"),
            _brow.get("LienholderId"), _brow.get("CollectorId"), _brow.get("LienType"),
        ):
            print(f"\n✅ Patched the TPA file in place for LienholderId {_brow.get('LienholderId')!r}.")
        else:
            print(f"\n⚠️  Could not patch the TPA file for LienholderId {_brow.get('LienholderId')!r} -- skipped, try resuming again later.")

    # Pending_Lien_Match_Review already carries a resolved LienholderId/CollectorId/
    # LienType (the tie happened INSIDE Step 4) -- the row re-enters the full
    # pipeline via `df`/`ev_df` like any resumed row, so Step 3 re-derives the same
    # LienholderId deterministically. What needs pre-seeding is each row's own
    # DECISION, applied unconditionally to LIEN_NAME_EQUIVALENCE_MAP ("Yes") /
    # LIEN_NAME_EXCLUDED_MAP ("No") -- keyed by (State, LienholderId), not
    # ClientId, since a confirmed name relationship is a fact about the PLAN.
    # classify_step4_scenario naturally re-resolves with this info -- if enough
    # decisions broke the tie it resolves on its own; if not, it re-hits the tie and
    # writes a fresh Pending_Lien_Match_Review, same self-sustaining loop as the
    # other 3 Pending_* sheets.
    _lien_match_review_rows = _resume_combined[_resume_combined["_source_sheet"] == "Pending_Lien_Match_Review"]

    # Real gap found (user report): nothing caught a user mistakenly marking 2+
    # DIFFERENT candidates "YES" for the SAME tie (same ClientId+LienId+MCO_Name)
    # -- exactly ONE candidate should ever be the real match. Pre-scan every
    # tie-group first; a group with more than one DISTINCT candidate marked "YES"
    # is resolved LIVE, right here (almost always a genuine accident, not an
    # intentional choice, so asking immediately is friendlier than forcing a
    # whole extra file round-trip): show every conflicting candidate, let the
    # user pick the real one (the rest are then recorded as confirmed NON-matches
    # -- the user already said one of them is right, so the others are now known
    # wrong, not just unknown), or 'S' to skip entirely -- a skip leaves the
    # group fully unresolved, the underlying tie re-occurs downstream, and a
    # fresh Pending_Lien_Match_Review is written, same self-sustaining loop as
    # any other unresolved row. Counts DISTINCT candidate names, not raw row
    # count -- the SAME candidate accidentally duplicated across 2 rows (both
    # "YES") is not a real conflict, just a duplicate of one decision.
    _lm_invalid_groups = set()
    _lm_live_resolved_groups = set()
    if not _lien_match_review_rows.empty:
        for _lm_group_key, _lm_group_df in _lien_match_review_rows.groupby(
            ["ClientId", "LienId", "MCO_Name"], dropna=False
        ):
            _lm_yes_rows = _lm_group_df[_lm_group_df["Decision"].astype(str).str.strip().str.upper() == "YES"]
            _lm_yes_rows = _lm_yes_rows.drop_duplicates(subset=["CandidateLienHolderName"])
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

            _lm_state = _lm_group_df["State"].iloc[0]
            _lm_collector_id = _lm_group_df["CollectorId"].iloc[0]
            _lm_lienholder_id = _lm_group_df["LienholderId"].iloc[0]
            _lm_plan_name = _lm_group_key[2]
            _lm_map_key = (str(_lm_state), str(_lm_collector_id), str(_lm_lienholder_id))
            _lm_picked_name = str(_lm_yes_rows.loc[int(_lm_pick) - 1, "CandidateLienHolderName"]).strip()

            LIEN_NAME_EQUIVALENCE_MAP.setdefault(_lm_map_key, [])
            if _lm_picked_name not in LIEN_NAME_EQUIVALENCE_MAP[_lm_map_key]:
                LIEN_NAME_EQUIVALENCE_MAP[_lm_map_key].append(_lm_picked_name)
            if ask_yes_no("Save this decision for future runs (any client with this same plan)? (y/n): "):
                _RESUME_LIEN_EQUIV_SAVES.append(
                    (_lm_state, _lm_collector_id, _lm_lienholder_id, _lm_plan_name, _lm_picked_name, "YES", TicketId)
                )
            for _, _other_row in _lm_yes_rows.iterrows():
                _other_name = str(_other_row["CandidateLienHolderName"]).strip()
                if _other_name == _lm_picked_name:
                    continue
                LIEN_NAME_EXCLUDED_MAP.setdefault(_lm_map_key, [])
                if _other_name not in LIEN_NAME_EXCLUDED_MAP[_lm_map_key]:
                    LIEN_NAME_EXCLUDED_MAP[_lm_map_key].append(_other_name)
                if ask_yes_no(f"Also save that {_other_name!r} does NOT match, for future runs? (y/n): "):
                    _RESUME_LIEN_EQUIV_SAVES.append(
                        (_lm_state, _lm_collector_id, _lm_lienholder_id, _lm_plan_name, _other_name, "NO", TicketId)
                    )
            _lm_live_resolved_groups.add(_lm_group_key)

    _lien_match_seen = set()
    for _, _lm_row in _lien_match_review_rows.iterrows():
        _lm_row_group = (_lm_row.get("ClientId"), _lm_row.get("LienId"), _lm_row.get("MCO_Name"))
        if _lm_row_group in _lm_invalid_groups or _lm_row_group in _lm_live_resolved_groups:
            continue
        _lm_decision = str(_lm_row.get("Decision") or "").strip().upper()
        if _lm_decision not in ("YES", "NO"):
            continue
        _lm_state = _lm_row.get("State")
        _lm_collector_id = _lm_row.get("CollectorId")
        _lm_lienholder_id = _lm_row.get("LienholderId")
        _lm_plan_name = _lm_row.get("MCO_Name")
        _lm_equiv_name = _lm_row.get("CandidateLienHolderName")
        if pd.isna(_lm_state) or pd.isna(_lm_collector_id) or pd.isna(_lm_lienholder_id) or pd.isna(_lm_equiv_name):
            continue
        _lm_dedup_key = (str(_lm_state), str(_lm_collector_id), str(_lm_lienholder_id),
                         str(_lm_equiv_name).strip(), _lm_decision)
        if _lm_dedup_key in _lien_match_seen:
            continue
        _lien_match_seen.add(_lm_dedup_key)
        _lm_map_key = (str(_lm_state), str(_lm_collector_id), str(_lm_lienholder_id))
        _lm_target_map = LIEN_NAME_EQUIVALENCE_MAP if _lm_decision == "YES" else LIEN_NAME_EXCLUDED_MAP
        _lm_target_map.setdefault(_lm_map_key, [])
        if str(_lm_equiv_name).strip() not in _lm_target_map[_lm_map_key]:
            _lm_target_map[_lm_map_key].append(str(_lm_equiv_name).strip())
        _lm_verb = "matches" if _lm_decision == "YES" else "does NOT match"
        _new_item("resume")
        print(
            f"\nResolved from Pending_Lien_Match_Review: {_lm_plan_name!r} -> "
            f"{_lm_equiv_name!r} {_lm_verb} LienholderId {_lm_lienholder_id!r} (State {_lm_state!r})"
        )
        if ask_yes_no("Save this decision for future runs (any client with this same plan)? (y/n): "):
            _RESUME_LIEN_EQUIV_SAVES.append(
                (_lm_state, _lm_collector_id, _lm_lienholder_id, _lm_plan_name, _lm_equiv_name, _lm_decision, TicketId)
            )

    # _resume_combined is read via dtype=str (needed so every other column round-
    # trips as plain text), which also turns a genuine Excel date cell into a raw
    # string -- and format_date returns the literal "Bad Date" for any string input.
    # Re-parsed into real datetimes here (coercing anything unparseable to NaT, same
    # as "no date found") at the source instead of downstream.
    df = _resume_combined[["LienId", "MCO_Name"]].copy()
    df["Start_Date"] = (
        pd.to_datetime(_resume_combined["Start_Date"], errors="coerce")
        if "Start_Date" in _resume_combined.columns else None
    )
    df["End_Date"] = (
        pd.to_datetime(_resume_combined["End_Date"], errors="coerce")
        if "End_Date" in _resume_combined.columns else None
    )
    ev_df = _resume_combined
    LIENID_COL_IN_EVDF = "LienId"
    MISMATCH_FLAGGED_LIENIDS = set()
    print(f"✅ Resumed rows: {len(df)}")

elif IS_LA_FILE:
    # LA file processing
    df, ev_input_df, LIENID_COL_IN_EVDF, MISMATCH_FLAGGED_LIENIDS = process_la_file(MCO_INPUT_FILE)
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
    ev_df = _retry_on_permission_error(
        lambda: pd.read_excel(MCO_INPUT_FILE, sheet_name=_input_sheet, dtype=str), "the import file"
    )
    ev_df.columns = ev_df.columns.str.strip()

    _column_map_hit = _try_column_map_lookup(STATE, list(ev_df.columns))
    if _column_map_hit:
        lienid_col, MCO_COLUMNS = _column_map_hit
        print(
            f"✅ Column mapping reused from a previous {STATE} run: "
            f"LienId={lienid_col!r}, MCO={MCO_COLUMNS!r}"
        )
        ev_df.rename(columns={lienid_col: "LienId"}, inplace=True)
        LIENID_COL_IN_EVDF = "LienId"
        MISMATCH_FLAGGED_LIENIDS = set()
    else:
        _phase_banner("column", "COLUMN IDENTIFICATION")
        _preview_columns(ev_df)
        lienid_col = _ask_columns(
            ev_df,
            "LienId: Enter the column number: ",
            allow_empty=False, multi=False, show_list=False
        )[0]

        _gap()
        MCO_COLUMNS = _ask_columns(
            ev_df,
            "MCO: Enter column number(s), comma-separated: ",
            allow_empty=False, multi=True, show_list=False
        )

        if STATE:
            _COLUMN_MAP_SAVES.append((STATE, lienid_col, MCO_COLUMNS, TicketId))

        ev_df.rename(columns={lienid_col: "LienId"}, inplace=True)
        LIENID_COL_IN_EVDF = "LienId"
        MISMATCH_FLAGGED_LIENIDS = set()

    def has_any_plan(row):
        for c in MCO_COLUMNS:
            raw = row[c]
            if pd.notna(raw) and str(raw).strip():
                val = _normalize_mco_text(str(raw).strip().upper())
                val_clean = _RE_STRIP_SPACE.sub('', val)
                if val_clean in INVALID:
                    continue
                val_words = set(val.replace(',', ' ').split())
                if val_words.issubset(INVALID):
                    continue
                if _RE_HAS_ALPHA.search(val):
                    return True
        return False

    ev_plan_df = ev_df[ev_df.apply(has_any_plan, axis=1)].copy()

    rows = [
        {"LienId": str(row["LienId"]), "MCO_Name": plan, "Start_Date": start, "End_Date": end}
        for _, row in ev_plan_df.iterrows()
        for c in MCO_COLUMNS
        if pd.notna(row[c]) and str(row[c]).strip()
        and _RE_STRIP_SPACE.sub('', str(row[c]).strip().upper()) not in INVALID
        for plan, start, end in extract_plan_rows(str(row[c]).strip())
    ]

    df = pd.DataFrame(rows)

# =====================================================
# EMPTY CHECK
# =====================================================

if df.empty:
    print("⚠️ Koi MCO plans nahi mile input file mein!")
    _ensure_writable(OUTPUT_FILE)
    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        ev_df.to_excel(writer, sheet_name="Input_File_Original", index=False)
    print(f"✅ Output file (sirf Input): {OUTPUT_FILE}")

else:

    yes_with_mco_df = (
        df.groupby(["LienId", "MCO_Name"])
        .agg(Start_Date=("Start_Date", safe_min), End_Date=("End_Date", safe_max))
        .reset_index()
    )

    # =====================================================
    # SQL — CLIENT ID QUERY
    # =====================================================

    lien_ids = make_sql_in(yes_with_mco_df["LienId"].unique())

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

    sql_df = pd.read_sql(sql_query, conn)
    sql_df[["LienId", "ClientId"]] = sql_df[["LienId", "ClientId"]].astype(str)

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

    sql_mcos_df = pd.read_sql(sql_mcos_query, conn)
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
        f.IsMCO, c.PrivateLiensScope, clientadditionalinformation,
        f.Tort, f.TortId, LEFT(f.LienHolderName,2) AS State
    from fullproductviews f
    inner join cases c on f.caseid = c.id
    where f.ClientId in {client_ids}
    Order by 5, 2
    """

    sql_liens_df = pd.read_sql(sql_liens_query, conn)
    sql_liens_df["ClientId"] = sql_liens_df["ClientId"].astype(str)
    sql_liens_df["LienId"]   = sql_liens_df["LienId"].astype(str)
    sql_liens_df["TortId"]   = sql_liens_df["TortId"].astype(str)
    sql_liens_df["Last_Name"]  = sql_liens_df["clientlastname"]
    sql_liens_df["First_Name"] = sql_liens_df["clientfirstname"]

    # Everything simple-lookup now comes from ONE source (sql_liens_df), ONE call, keyed on LienId.
    # TortId rides along purely for Step 3's Pending_TPA_Review (so a not-found-in-TPA Tort still
    # gets a real, DB-sourced TortId auto-populated on that review row -- the user never has to
    # supply one) -- it naturally falls out of yes_with_mco_df again at FINAL COLUMN ORDER below,
    # same as CollectorName, since nothing after Step 3 needs it.
    yes_with_mco_df = bring_in_columns(
        yes_with_mco_df, sql_liens_df, "LienId",
        ["ClientId", "Tort", "TortId", "State", "Last_Name", "First_Name", "PrivateLiensScope", "clientadditionalinformation"]
    ).rename(columns={"Tort": "Project_Name"})

    def normalize_mco_name(row):
        state = str(row.get("State", "")).strip().upper()
        plan  = str(row.get("MCO_Name", "")).strip().upper()
        return ALIAS_MAP.get(state, {}).get(plan, plan)

    yes_with_mco_df["MCO_Name"] = yes_with_mco_df.apply(normalize_mco_name, axis=1)

    grp_keys  = ["LienId", "ClientId", "Project_Name", "TortId", "State", "Last_Name", "First_Name", "MCO_Name", "PrivateLiensScope", "clientadditionalinformation"]
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
                 End_Date  =("End_Date",   lambda x: safe_max(list(x))))
            .reset_index()
        )
        yes_with_mco_df["Start_Date"] = yes_with_mco_df["Start_Date"].apply(format_date)
        yes_with_mco_df["End_Date"]   = yes_with_mco_df["End_Date"].apply(format_date)
    else:
        yes_with_mco_df = (
            yes_with_mco_df.groupby(grp_keys, dropna=False).size().reset_index().drop(columns=[0])
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
    # so Input_File_Original's Remarks column can show it. Defaults to a
    # visibly-wrong placeholder on purpose, so an unrecorded fate surfaces as a gap
    # to fix rather than silently looking "processed".
    #
    # Built HERE, after alias-expansion + the post-alias dedup groupby -- not right
    # after the raw groupby -- because MCO_Name gets rewritten to its canonical form
    # further down (e.g. "MHS" -> "MANAGED HEALTH SERVICES (IN MCO)/CENTENE"), and
    # every later stage reports fate against that final name. Keying on the raw
    # pre-alias name would strand it at the placeholder while the real outcome
    # lands under the aliased name -- Remarks would show both a fake gap and the
    # real result for the same plan.
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
    # Program-type/state tags -- "mco"/"cco" and the states currently covered by
    # the TPA file. Within one state's own sheet, every row shares these, so
    # they carry zero distinguishing power and never count as real
    # corroboration (distinct from _GENERIC_CONNECTORS, which are excluded from
    # scoring context too, not just this count).
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

        # Parenthetical annotations ("(OR CCO)", "(Formerly WOAH)", "(AKA HAP)") are
        # administrative tags, not part of a plan's own identity -- excluded from the
        # corroboration guard below only, never from the gate or scoring. The "/"
        # co-branding exception (rule 3) only applies when the "/" is OUTSIDE any
        # such annotation -- a "/" inside one (e.g. "(Wellcare/Centene)") is just
        # naming a parent company, not offering an alternate name for THIS plan.
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
            if all(w in _GENERIC_CONNECTORS for w in effective_shared) and leftover_query_raw:
                continue  # only generic words shared, and the query has other unexplained content
            # The candidate's own leftover only counts against it when it's
            # "risky" -- a single generic trailing word ("Passport Group"'s
            # "group") is normal, harmless filler, but a REAL leftover word, or
            # 2+ generic ones together (real extra structure, not just one
            # filler suffix), means the candidate has its own unexplained
            # identity -- only then, combined with the query ALSO having real
            # unaccounted content, is the match too uncertain to guess.
            # Real bug found (user report): a leftover _STRUCTURAL_TAG (mco/cco/a
            # state name -- "no signal within one state's own sheet," same as the
            # effective_shared check above already treats them) was being counted
            # as "real unaccounted content" here, on EITHER side -- e.g. query
            # "WellCare of Kentucky/Centene" left "kentucky" unaccounted, which
            # combined with the candidate's own ("WellCare Health Plans/Centene")
            # 2-generic-word leftover ("health"+"plans") to wrongly refuse a match
            # that should have succeeded (same real plan, just a state suffix and
            # a differently-worded generic filler phrase). Excluded here exactly
            # like generic connectors already are, on both sides, for consistency.
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

    # clientadditionalinformation is free-text (case notes), so an exact-match table
    # (like SCOPE_MAP) won't work -- "PLRP" and an opt-in/opt-out phrase must read as
    # ONE phrase ("PLRP opt out", "opted out of PLRP"), not just both appearing
    # somewhere in the note. Gap between them: letters/whitespace only, max 15 chars
    # -- punctuation or an unrelated clause in between ("PLRP - rep name - opt out")
    # breaks the match on purpose (different, unrelated mention). If found, this
    # OVERRIDES the client's classified scope (opt-out -> No PLRP, opt-in -> 100%
    # PLRP) regardless of PrivateLiensScope. If neither is found (the common case),
    # no override happens, the original classification stands.
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
        Returns (resolved_df, pending_df, scope_saves): scope_saves is a list of
        (scope_text, classification) pairs the user chose to persist to Scope_Map."""
        working_df = df_in.copy()
        working_df["Scope_Classification"] = working_df["PrivateLiensScope"].apply(classify_scope)

        unresolved_mask = working_df["Scope_Classification"].isna()
        unresolved_texts = working_df.loc[unresolved_mask, "PrivateLiensScope"].dropna().unique().tolist()

        scope_saves = []
        skip_texts = set()

        # Only banner the phase if at least one text will actually need live
        # resolution -- a run where every scope text already resolves silently via
        # scope_map shouldn't announce a phase that never shows anything.
        if any(scope_map.get(_normalize_key(t)) not in ("100% PLRP", "NO PLRP") for t in unresolved_texts):
            _phase_banner("scope", "SCOPE CLASSIFICATION")

        for raw_text in unresolved_texts:
            key = _normalize_key(raw_text)
            if scope_map.get(key) in ("100% PLRP", "NO PLRP"):
                working_df.loc[working_df["PrivateLiensScope"] == raw_text, "Scope_Classification"] = scope_map[key]
                continue

            matching_client_ids = working_df.loc[working_df["PrivateLiensScope"] == raw_text, "ClientId"].unique().tolist()
            _new_item("scope")
            print(f"\nScope not recognized: {raw_text!r}")

            if show_diagnostic:
                ids_clause = make_sql_in(matching_client_ids)
                diag_query = _SCOPE_DIAG_QUERY_TEMPLATE.format(ids_clause=ids_clause)
                diag_result = pd.read_sql(diag_query, conn)
                _print_scope_diag_table(diag_result)

            while True:
                answer = input(
                    "Classify this scope -- enter '100' for 100% PLRP, 'No' for No PLRP, "
                    "or 'S' to skip: "
                ).strip().upper()
                if answer in ("100", "100%", "100% PLRP", "100%PLRP"):
                    classification = "100% PLRP"
                    break
                if answer in ("NO", "NO PLRP", "NOPLRP"):
                    classification = "NO PLRP"
                    break
                if answer in ("SKIP", "S"):
                    classification = None
                    break
                print("Please enter '100', 'No', or 'S' to skip.")

            if classification:
                scope_map[key] = classification
                working_df.loc[working_df["PrivateLiensScope"] == raw_text, "Scope_Classification"] = classification
                if ask_yes_no("Save this decision for future runs? (y/n): "):
                    scope_saves.append((raw_text, classification))
            else:
                skip_texts.add(raw_text)

        pending_mask = working_df["PrivateLiensScope"].isin(skip_texts)
        pending_df  = working_df[pending_mask].copy()
        resolved_df = working_df[~pending_mask].copy()
        return resolved_df, pending_df, scope_saves

    yes_with_mco_df, pending_scope_df, _SCOPE_SAVES = resolve_scope_for_run(yes_with_mco_df, SCOPE_MAP, show_diagnostic=not IS_RESUME_RUN)
    _SCOPE_SAVES = _RESUME_SCOPE_SAVES + _SCOPE_SAVES

    # Additional-info opt-in/opt-out override -- only applies to rows that already have
    # a resolved Scope_Classification (pending/skipped rows have none to override yet).
    _addinfo_result = yes_with_mco_df["clientadditionalinformation"].apply(classify_additional_info)
    yes_with_mco_df.loc[_addinfo_result == "OPT OUT", "Scope_Classification"] = "NO PLRP"
    yes_with_mco_df.loc[_addinfo_result == "OPT IN", "Scope_Classification"] = "100% PLRP"

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
        header = next(rows_iter)
        data_rows = []
        for row in rows_iter:
            if row[0] is None or str(row[0]).strip() == "":
                break
            data_rows.append(row)
        wb.close()
        tpa_df = pd.DataFrame(data_rows, columns=header)
        for _id_col in ("TortId", "CollectorId", "LienholderId"):
            if _id_col in tpa_df.columns:
                tpa_df[_id_col] = tpa_df[_id_col].apply(_stringify_id)
        if "LienholderName" in tpa_df.columns and "State MCO" in tpa_df.columns:
            tpa_df["LienholderName"] = tpa_df.apply(
                lambda r: _strip_state_suffix_from_name(r["LienholderName"], r["State MCO"]), axis=1
            )
        _tpa_sheet_cache[state] = tpa_df
        return tpa_df

    TPA_OVERRIDES_KEY_COLS = ["TortName", "Private Lien Case Scope", "CollectorId", "LienholderId",
                              "LienholderName", "LienType"]

    # Only ever exactly one of these two values in real use -- picked by number (same
    # ask_numbered_choice pattern already used for Response type) everywhere a user
    # would otherwise free-type LienType, so a typo can never silently create a
    # LienType value that doesn't match either real one.
    LIEN_TYPE_OPTIONS = ["Medicaid Lien - MT", "Private Lien - PLRP"]

    def _tpa_overrides_find(tpa_overrides, state_mco, tort_name, lienholder_name, scope):
        """Search TPA_OVERRIDES for a matching entry. A COMPLETE row (TortId not None)
        must match State MCO + TortName + LienholderName + Scope exactly (normalized).
        A PENDING row (TortId is None, e.g. a not-yet-completed legacy entry) matches on
        State MCO + LienholderName alone. A complete match always wins over a pending one."""
        norm_state = _normalize_key(state_mco)
        norm_tort  = _normalize_key(tort_name)
        norm_name  = _normalize_key(lienholder_name)
        norm_scope = _normalize_key(scope)
        complete_idx, pending_idx = None, None
        for i, r in enumerate(tpa_overrides):
            if _normalize_key(r.get("State MCO")) != norm_state or _normalize_key(r.get("LienholderName")) != norm_name:
                continue
            if r.get("TortId") is None:
                if pending_idx is None:
                    pending_idx = i
            elif _normalize_key(r.get("TortName")) == norm_tort and _normalize_key(r.get("Private Lien Case Scope")) == norm_scope:
                complete_idx = i
        return complete_idx if complete_idx is not None else pending_idx

    # (LienholderId, normalized Tort name) pairs the user already chose to skip this
    # run -- in-memory only, never persisted (a skipped legacy entry stays genuinely
    # pending on disk, same as before, so a FUTURE run can still try again). Keyed on
    # the entry's own LienholderId, not its list position -- the position is only
    # ever meaningful within ONE specific tpa_overrides list, while LienholderId is
    # the legacy entry's real, stable identity. Real bug found (user report):
    # skipping left the entry's TortId at None (by design, so a later Tort could
    # still complete it) -- but with no per-Tort memory, every OTHER row sharing this
    # SAME legacy LienholderId under the SAME Tort kept re-asking the identical
    # question, repeatedly, within one run. A resolved (non-skipped) completion never
    # had this problem: it sets TortId in place, so _tpa_overrides_find already finds
    # it as "complete" for that Tort on the next lookup. A DIFFERENT Tort for the
    # same legacy entry still asks its own independent question, exactly like the
    # name_not_found list-pick flow already does per Tort (see _group_picked_name).
    _pending_override_skipped_torts = set()

    def _complete_pending_override(tpa_overrides, idx, state_mco, tort_id, tort_name, scope, ticket_id):
        """First-use completion of a pending (e.g. migrated-legacy) TPA_Overrides entry:
        search the TPA file for this LienholderId to backfill CollectorId/LienType; if it
        isn't in TPA at all, ask the user once -- or let them skip if they don't know it
        right now. Tort/Scope/State always come from the current run's own already-known
        context. Returns the completed entry, or None if the user skipped -- in that case
        the pending TPA_Overrides entry itself is left completely untouched (still fully
        pending), so a later Tort (this run or a future one) can try completing it again."""
        entry = tpa_overrides[idx]
        lienholder_id = entry["LienholderId"]
        _skip_key = (str(lienholder_id), _normalize_key(tort_name))
        if _skip_key in _pending_override_skipped_torts:
            return None
        tpa_df = load_tpa_state_sheet(TPA_FILE, state_mco.split(" ")[0])
        backfill_row = None
        if tpa_df is not None:
            matches = tpa_df[tpa_df["LienholderId"].astype(str) == str(lienholder_id)]
            if not matches.empty:
                backfill_row = matches.iloc[0]
        if backfill_row is not None:
            entry["CollectorId"]   = backfill_row["CollectorId"]
            entry["LienType"]      = backfill_row["LienType"]
        else:
            _plan_banner_once()
            _new_item("plan")
            print(
                f"\nLienholderId {lienholder_id} (\"{entry['LienholderName']}\") was in the "
                f"legacy override list but is not in the TPA file for Tort {tort_name!r} / "
                f"Scope {scope!r}. Please provide the missing details."
            )
            collector_id = input("CollectorId (or 'S' to skip): ").strip()
            if collector_id.upper() == "S":
                _pending_override_skipped_torts.add(_skip_key)
                return None
            entry["CollectorId"]   = collector_id
            entry["LienType"]      = ask_numbered_choice(LIEN_TYPE_OPTIONS, "LienType:", "Enter the number: ")
        entry["TortId"] = tort_id
        entry["TortName"] = tort_name
        entry["Private Lien Case Scope"] = scope
        entry["State MCO"] = state_mco
        entry["Ticket #"] = ticket_id
        return entry

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
            if e.get("TortId") is not None
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
            if _tpa_overrides_find(tpa_overrides, state_mco2, tort2, mco_name2, scope2) is not None:
                continue
            result2 = resolve_tpa_lookup(state2, tort2, mco_name2, scope2)
            if result2["status"] != "name_not_found":
                continue
            distinct_names2 = result2["candidates"]["LienholderName"].drop_duplicates().tolist()
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
        matched_id, tied_ids = _match_name(mco_name, tort_df, "LienholderName", "LienholderId", return_tie=True)
        if matched_id is None:
            if tied_ids:
                # A genuine _match_name tie (2+ real candidates equally explain the
                # query) IS the relevant shortlist already -- no need to fall back to
                # word-overlap filtering, and no risk of the fallback accidentally
                # excluding one of the tied names.
                _tied_id_strs = {str(t) for t in tied_ids}
                candidates = tort_df[tort_df["LienholderId"].astype(str).isin(_tied_id_strs)]
            else:
                candidates = _plan_candidates_for_prompt(mco_name, tort_df)
            return {"status": "name_not_found", "candidates": candidates}
        # Group by the matched NAME, not the specific id _match_name happened to return --
        # two rows can share the exact same LienholderName with DIFFERENT LienholderIds
        # (the real "United Healthcare" duplicate case this design targets).
        matched_row = tort_df[tort_df["LienholderId"].astype(str) == str(matched_id)].iloc[0]
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

    def resolve_tpa_for_run(df_in, tpa_overrides, ticket_id, alias_saves=None):
        """Resolves LienholderId/CollectorId/LienType for every unique
        (State, Tort, MCO_Name, Scope_Classification) combo in df_in. Returns
        (resolved_df, pending_tpa_df, tpa_dirty). Raises RuntimeError immediately if
        any client's state isn't covered by the TPA file -- every client in one
        input file shares the same state, so nothing here could proceed otherwise.
        alias_saves, if given, gets (State, From_Name, To_Name) appended whenever a
        name_not_found list-pick's raw text differs from the real TPA name --
        optional so callers that don't need Alias_Map persistence don't change."""
        if alias_saves is None:
            alias_saves = []
        working_df = df_in.copy()
        for col in ["LienholderId", "CollectorId", "LienType"]:
            working_df[col] = None

        tpa_dirty = []
        pending_idx = []
        blank_field_idx = []

        def _is_blank(v):
            return pd.isna(v) or str(v).strip() == ""

        def _finalize_tpa_match(row_mask, lienholder_id, collector_id, lien_type,
                                 tort_id=None, tort_name=None, scope_val=None,
                                 lienholder_name=None, state_mco_val=None,
                                 is_first_for_combo=True):
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
                _given_collector_id = input("CollectorId (or 'S' to skip): ").strip()
                if _given_collector_id and _given_collector_id.upper() != "S":
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
            # Once a plan is resolved to a known name (a fresh TPA match, a list-pick,
            # or a saved override's own name), MCO_Name is rewritten to it -- the same
            # mechanism Alias_Map already uses upfront (normalize_mco_name), applied
            # consistently across every resolution path instead of only pre-run
            # aliases. Yes_with_MCO's own MCO_Name column, the Remarks column, and
            # every LienNote's "Plan Name:" all read directly from this one field, so
            # this single rewrite keeps all three in agreement.
            # PLAN_FATE is keyed by (LienId, MCO_Name) and was already seeded with the
            # pre-resolution value before this ever runs (same staleness pattern as
            # merge_same_plan_within_client) -- the old key must be dropped here too,
            # or Remarks would show a fake "tracking gap" for the old text alongside
            # the real outcome for the new one.
            if not _is_blank(lienholder_name):
                # Must compare the RAW value here, not _normalize_key(...) -- PLAN_FATE's
                # own keys are the literal, un-normalized MCO_Name text. "ANTHEM" and
                # "Anthem" normalize to the same key, but they are TWO DIFFERENT PLAN_FATE
                # keys -- comparing normalized values wrongly concluded "nothing changed"
                # for a pure case difference, skipped the pop, while the assignment below
                # still ran unconditionally and changed the literal text anyway. Left the
                # OLD-case key stuck at its placeholder forever while Step 4 recorded the
                # real outcome under the NEW-case key -- Remarks showed both.
                for _, _fmrow in working_df.loc[row_mask, ["LienId", "MCO_Name"]].drop_duplicates().iterrows():
                    if _fmrow["MCO_Name"] != lienholder_name:
                        PLAN_FATE.pop((_fmrow["LienId"], _fmrow["MCO_Name"]), None)
                working_df.loc[row_mask, "MCO_Name"] = lienholder_name
            return True

        # When the SAME raw text fails to match TPA for one Tort, the SAME candidate
        # list almost always shows again for another Tort sharing that text in the
        # same run. Combos already fully handled -- individually, or as part of a
        # name_not_found GROUP (see below) -- are tracked here so the outer loop skips
        # them when it later reaches their own row in `combos`, instead of processing
        # them a second time.
        _processed_combos = set()

        # Two different raw texts can still end up with the exact same narrowed
        # candidate list -- e.g. a state/tort with only one plan on file, or two raw
        # texts that both fail to share a word with anything (both fall back to the
        # full sheet, see _plan_candidates_for_prompt) -- even though the list is now
        # filtered toward what each raw text actually looks like rather than being
        # every plan name for the State+Tort. Tracks candidate lists by content (not
        # raw text) already printed this run -- a later, different raw text whose
        # candidates are identical gets a short pointer back instead of a full
        # reprint, but still gets its own fully independent question every time --
        # only the redundant printing changes.
        _seen_candidate_lists = {}

        combos = working_df[["State", "Project_Name", "MCO_Name", "Scope_Classification"]].drop_duplicates()

        for _, combo in combos.iterrows():
            state, tort, mco_name, scope = (
                combo["State"], combo["Project_Name"], combo["MCO_Name"], combo["Scope_Classification"]
            )
            combo_key = (state, tort, mco_name, scope)
            if combo_key in _processed_combos:
                continue
            row_mask = (
                (working_df["State"] == state) & (working_df["Project_Name"] == tort) &
                (working_df["MCO_Name"] == mco_name) & (working_df["Scope_Classification"] == scope)
            )
            state_mco = f"{str(state).strip().upper()} MCO"
            # Real, DB-sourced TortId for THIS combo (from sql_liens_query, carried on
            # working_df this whole time) -- used everywhere below that builds/completes
            # a TPA_Overrides entry, so every entry ends up "complete" (a real TortId,
            # not hardcoded None) and therefore correctly Tort-scoped: a DIFFERENT Tort
            # sharing the same LienholderName can never silently pick up (or clobber)
            # an entry that was only ever meant for THIS Tort.
            row_tort_id = working_df.loc[row_mask, "TortId"].iloc[0] if "TortId" in working_df.columns else None

            ov_idx = _tpa_overrides_find(tpa_overrides, state_mco, tort, mco_name, scope)
            if ov_idx is not None:
                entry = tpa_overrides[ov_idx]
                _needs_dirty_save = False
                if entry.get("TortId") is None:
                    # Must complete with a REAL TortId, not a hardcoded None --
                    # otherwise _tpa_overrides_find's own pending-vs-complete test
                    # (TortId is None) would keep treating this row as pending on every
                    # later run, letting a SECOND Tort using the same LienholderName
                    # find this SAME row again, re-supply CollectorId/LienType, and
                    # silently overwrite the first Tort's already-saved data.
                    entry = _complete_pending_override(tpa_overrides, ov_idx, state_mco, row_tort_id, tort, scope, ticket_id)
                    if entry is None:
                        # User doesn't know CollectorId/LienType right now, and it isn't
                        # in the raw TPA file either -- send to Pending_TPA_Review with
                        # what IS already known (LienholderId, from the legacy override)
                        # pre-filled, same shape as any other Pending_TPA_Review row, so
                        # the existing resume flow picks it up once filled in by hand.
                        # The pending TPA_Overrides entry itself is untouched (still
                        # fully pending) -- _complete_pending_override guarantees this.
                        working_df.loc[row_mask, "LienholderId"] = tpa_overrides[ov_idx]["LienholderId"]
                        working_df.loc[row_mask, "PendingReason"] = (
                            "Legacy TPA_Overrides entry needs CollectorId/LienType "
                            "(not found in TPA file); user skipped"
                        )
                        pending_idx.extend(working_df[row_mask].index.tolist())
                        continue
                    _needs_dirty_save = True
                # tpa_dirty save only happens once the entry is confirmed complete
                # below -- appending it unconditionally right after legacy completion
                # would persist an incomplete (blank CollectorId/LienType) entry to
                # mco_config.xlsx even on a run where this row correctly ends up
                # pending instead.
                if _finalize_tpa_match(
                    row_mask, entry["LienholderId"], entry["CollectorId"], entry["LienType"],
                    tort_id=entry.get("TortId"), tort_name=entry.get("TortName"),
                    scope_val=entry.get("Private Lien Case Scope"),
                    lienholder_name=entry.get("LienholderName"), state_mco_val=state_mco,
                    is_first_for_combo=not _needs_dirty_save,
                ):
                    if _needs_dirty_save:
                        tpa_dirty.append(entry)
                else:
                    print(
                        f"TPA_Overrides entry for {mco_name!r} (Tort {tort!r} / Scope {scope!r}) "
                        f"is missing required details (CollectorId/LienType)."
                    )
                continue

            result = resolve_tpa_lookup(state, tort, mco_name, scope)
            status = result["status"]

            if status == "state_not_covered":
                print(f"State {state} is not included in the TPA file for this run.")
                raise RuntimeError(f"Run aborted: state {state} is not covered by the TPA file.")

            if status == "matched":
                row = result["row"]

                # Real bug: the raw TPA file's OWN row for this exact name can be known-
                # stale (TPA_Overrides already corrects it under this same name, just a
                # different LienholderId) -- if the CURRENT run's raw text doesn't
                # exactly match the override's own stored text, _tpa_overrides_find
                # above never even sees it, so this fuzzy "matched" result would
                # otherwise be trusted silently, with no review at all. Two layers:
                # first a NARROW, scoped-fuzzy check (this exact Tort+Scope+State
                # bucket, fuzzy name match -- safe because the bucket is tiny) that can
                # confidently auto-resolve; then a BROADER exact-name fallback (any
                # Tort/Scope) that only proves the raw TPA data is untrustworthy
                # somewhere, without a confirmed answer for this exact combination.
                _scoped_match = _find_scoped_override_match(tpa_overrides, state_mco, tort, scope, row["LienholderName"])
                if _scoped_match is not None:
                    _scoped_ok = _finalize_tpa_match(
                        row_mask, _scoped_match["LienholderId"], _scoped_match["CollectorId"],
                        _scoped_match["LienType"], tort_id=_scoped_match.get("TortId"),
                        tort_name=_scoped_match.get("TortName"),
                        scope_val=_scoped_match.get("Private Lien Case Scope"),
                        lienholder_name=_scoped_match.get("LienholderName"), state_mco_val=state_mco,
                    )
                    if not _scoped_ok:
                        print(
                            f"TPA_Overrides entry for {mco_name!r} (Tort {tort!r} / Scope {scope!r}) "
                            f"is missing required details (CollectorId/LienType)."
                        )
                    elif _normalize_key(mco_name) != _normalize_key(row["LienholderName"]):
                        alias_saves.append((state, mco_name, row["LienholderName"]))
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

                # LienholderName is deliberately NOT checked here (and never needs to
                # be): _match_name only ever matches on real LienholderName text, and
                # name_matches/scope_matches are filtered by that same matched name key
                # -- a row with a blank LienholderName structurally can't reach
                # "matched" status at all, so checking it here would be dead code. Only
                # CollectorId/LienType are independent of the matching cascade and can
                # be blank on an otherwise-correctly-matched row. CollectorName is
                # deliberately excluded too -- it isn't read anywhere else in the
                # pipeline, so there's nothing to gain by asking the user to supply it.
                if not _finalize_tpa_match(
                    row_mask, row["LienholderId"], row.get("CollectorId"), row.get("LienType"),
                    tort_id=row.get("TortId"), tort_name=row.get("TortName"),
                    scope_val=row.get("Private Lien Case Scope"), lienholder_name=row.get("LienholderName"),
                    state_mco_val=state_mco,
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
                pending_idx.extend(working_df[row_mask].index.tolist())
                continue

            if status == "name_not_found":
                # Deliberately NOT fuzzy-matching the raw text against the
                # TPA_Overrides bucket here (only an EXACT match, at the top of this
                # loop, is used to reuse a past override for this status). Real data
                # showed this is unsafe: a scoped (State+Tort+Scope) bucket can still
                # hold several genuinely different real plans that share one common
                # brand/system prefix word (e.g. "HSO CareOR", "HSO Kaiser", "HSO
                # Prov" all starting with "HSO") -- _match_name's first-word gate
                # passes on that shared prefix alone, silently resolving one plan's
                # raw text to a COMPLETELY DIFFERENT plan's saved override whenever
                # only one candidate happens to be in the bucket (no tie to catch it).
                _text_key = _normalize_key(mco_name)
                candidates = result["candidates"]
                distinct_names = candidates["LienholderName"].drop_duplicates().tolist()
                _signature = (
                    _text_key,
                    tuple(sorted(_normalize_key(n) for n in distinct_names)),
                )

                # Announcing a group ("...for Tort 'Hernia Mesh' and 'Abuse'") but
                # letting the outer loop reach 'Abuse' only later -- with an unrelated
                # question landing in between -- makes it look like part of the group.
                # Gather the FULL group now (via _find_same_signature_combos, which
                # reuses the same real _tpa_overrides_find/resolve_tpa_lookup so it
                # can't drift), announce once, then ask EVERY member's own question
                # back-to-back. Each stays independent -- CollectorId/LienType can
                # differ by Tort/Scope for the same lienholder -- members are just
                # marked processed so the outer loop skips them later.
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
                        # Same underlying ambiguity as resolve_tpa_lookup's own
                        # ambiguous_duplicate status (Tort+Name+Scope still leaves
                        # 2+ TPA rows differing by CollectorId/LienholderId/
                        # LienType) -- give it the SAME treatment: show the
                        # candidates, let the user pick one or skip, instead of
                        # silently sending to review with no explanation of why.
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
                        # Scope genuinely doesn't match any TPA row for this
                        # name+tort -- no valid option to choose from, straight to
                        # review (same reasoning as resolve_tpa_lookup's own
                        # scope_gap case). Every other pending/skip path in this
                        # file explains itself with a print statement -- this
                        # branch must too, or a picked number produces zero
                        # feedback (no "Save?" prompt either, since nothing
                        # actually resolved) with no way to tell why.
                        print(
                            f"{picked_name!r} was {'auto-applied' if auto_reused else 'picked'} for "
                            f"{g_mco_name!r} (Tort {g_tort!r} / Scope {g_scope!r}), but its scope "
                            f"didn't narrow to exactly one TPA row -- sent for review."
                        )
                        working_df.loc[g_row_mask, "PendingReason"] = (
                            "Plan picked from list, but scope didn't narrow to exactly one TPA row"
                        )
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
                        is_first_for_combo=False,
                    ):
                        print(
                            f"TPA row {'auto-applied' if auto_reused else 'picked'} for {g_mco_name!r} "
                            f"(Tort {g_tort!r} / Scope {g_scope!r}) but is missing required fields "
                            f"(CollectorId/LienType)."
                        )
                        return
                    # The picked name is already a real, present TPA row for this Tort
                    # -- persisting a TPA_Overrides entry for it is redundant (unlike
                    # 'D' flow, where the data genuinely can't be found any other way):
                    # once Alias_Map rewrites this raw text to the picked canonical
                    # name, the normal TPA cascade finds it on its own every future
                    # run, no override needed. Kept in-memory only (tpa_overrides.append,
                    # not tpa_dirty) so a LATER combo this same run sharing this exact
                    # raw text still skips re-asking, but nothing persists to
                    # TPA_Overrides in mco_config.xlsx -- only the alias does, below.
                    new_entry = {
                        "TortId": picked.get("TortId"), "TortName": g_tort, "Private Lien Case Scope": g_scope,
                        "CollectorId": picked["CollectorId"], "LienholderId": picked["LienholderId"],
                        "LienholderName": g_mco_name,
                        "LienType": picked["LienType"], "State MCO": g_state_mco, "Ticket #": ticket_id,
                    }
                    tpa_overrides.append(new_entry)
                    if ask_yes_no("Save this decision for future runs? (y/n): "):
                        # Remembered as a plain name alias so the NEXT raw occurrence
                        # of this text (any Tort, any scope, any client) is rewritten
                        # to the canonical name before TPA lookup runs, and resolves
                        # entirely on its own via the normal cascade. Skipped when raw
                        # and picked text are already identical -- nothing to alias.
                        if _normalize_key(g_mco_name) != _normalize_key(picked_name):
                            alias_saves.append((g_state, g_mco_name, picked_name))
                        if _scope_was_ambiguous:
                            # Unlike the plain "1 row matched" case above, aliasing
                            # alone isn't enough here: even after the raw text is
                            # rewritten to picked_name, the normal TPA cascade would
                            # hit this SAME Tort+Name+Scope ambiguity again for any
                            # future client, since the TPA file itself genuinely has
                            # 2+ rows for it. Persisted under the CANONICAL name
                            # (matching resolve_tpa_lookup's own ambiguous_duplicate
                            # save convention) so _tpa_overrides_find catches it on
                            # a future run before the cascade ever re-hits the
                            # ambiguity.
                            tpa_dirty.append({**new_entry, "LienholderName": picked_name})

                # Once a LienholderName is picked from the list for the FIRST Tort in
                # this group, that SAME choice applies to every OTHER Tort sharing
                # this exact raw text -- unlike 'D' flow's raw CollectorId/LienType,
                # which really can differ by Tort/Scope for the same lienholder and
                # stays independent. CollectorId/LienType are always freshly
                # re-derived from each Tort's own TPA rows via _resolve_picked_plan,
                # so reusing just the name carries no accuracy risk. 'D' and 'S'
                # choices are never reused this way -- only a genuine number-pick
                # sets this.
                _group_picked_name = None
                for (g_tort, g_state, g_mco_name, g_scope, g_candidates, g_distinct_names) in _group:
                    g_row_mask = (
                        (working_df["State"] == g_state) & (working_df["Project_Name"] == g_tort) &
                        (working_df["MCO_Name"] == g_mco_name) & (working_df["Scope_Classification"] == g_scope)
                    )
                    g_state_mco = f"{str(g_state).strip().upper()} MCO"
                    g_row_tort_id = (
                        working_df.loc[g_row_mask, "TortId"].iloc[0] if "TortId" in working_df.columns else None
                    )
                    _processed_combos.add((g_state, g_tort, g_mco_name, g_scope))

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
                            break
                        print(f"Please enter a number between 1 and {len(g_distinct_names)}, 'D', or 'S'.")

                    if answer == "S":
                        working_df.loc[g_row_mask, "PendingReason"] = _name_not_found_pending_reason(
                            g_state, g_tort, g_candidates
                        )
                        pending_idx.extend(working_df[g_row_mask].index.tolist())
                        continue
                    if answer == "D":
                        # Ask all 3 directly, in one go -- the numbered list above is
                        # only a narrowed shortlist of plausibly-related names (see
                        # _plan_candidates_for_prompt), not necessarily every plan TPA
                        # has for this Tort/State, so 'D' is for when the user already
                        # knows the LienholderId and it simply isn't one of the names
                        # shown -- not a claim that TPA has nothing else at all. A
                        # blank/'S' answer here backs out to the same skip protocol as
                        # the top-level prompt.
                        _gap()
                        given_lhid = input("LienholderId (or 'S' to skip): ").strip()
                        if not given_lhid or given_lhid.upper() == "S":
                            working_df.loc[g_row_mask, "PendingReason"] = _name_not_found_pending_reason(
                                g_state, g_tort, g_candidates
                            )
                            pending_idx.extend(working_df[g_row_mask].index.tolist())
                            continue
                        _gap()
                        collector_id = input("CollectorId: ").strip()
                        lien_type = ask_numbered_choice(LIEN_TYPE_OPTIONS, "LienType:", "Enter the number: ")
                        # Routed through the shared _finalize_tpa_match helper -- a
                        # blank CollectorId here (an empty Enter at the prompt) used
                        # to be assigned directly, bypassing the blank-fields check
                        # every other resolution path already has.
                        if not _finalize_tpa_match(
                            g_row_mask, given_lhid, collector_id, lien_type,
                            tort_id=g_row_tort_id, tort_name=g_tort, scope_val=g_scope,
                            lienholder_name=None, state_mco_val=g_state_mco,
                            is_first_for_combo=False,
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
                            "Ticket #": ticket_id,
                        }
                        tpa_overrides.append(new_entry)
                        if ask_yes_no("Save this decision for future runs? (y/n): "):
                            tpa_dirty.append(new_entry)
                        continue
                    picked_name = g_distinct_names[int(answer) - 1]
                    _resolve_picked_plan(g_row_mask, g_tort, g_mco_name, g_scope, g_state, g_state_mco,
                                          g_candidates, picked_name, auto_reused=False)
                    _group_picked_name = picked_name
                continue

            if status == "ambiguous_duplicate":
                candidates = result["candidates"]

                # ambiguous_duplicate always narrows down to ONE shared TPA name first
                # (that's why it's "ambiguous" -- multiple rows sharing that exact
                # name), so the correct key to save/check a resolution under is that
                # stable, already-known TPA name, not the raw run-time text that
                # happened to fuzzy-match it. Checking here before ever printing the
                # list means a later occurrence -- same or differently-worded raw
                # text, as long as it fuzzy-matches to this same TPA name -- resolves
                # silently, with no Alias_Map needed: TPA's own cascade already
                # re-derives this same canonical name fresh every time.
                _canonical_name = candidates.iloc[0]["LienholderName"]
                _amb_ov_idx = _tpa_overrides_find(tpa_overrides, state_mco, tort, _canonical_name, scope)
                if _amb_ov_idx is not None:
                    _amb_entry = tpa_overrides[_amb_ov_idx]
                    if not _finalize_tpa_match(
                        row_mask, _amb_entry["LienholderId"], _amb_entry["CollectorId"], _amb_entry["LienType"],
                        tort_id=_amb_entry.get("TortId"), tort_name=_amb_entry.get("TortName"),
                        scope_val=_amb_entry.get("Private Lien Case Scope"),
                        lienholder_name=_amb_entry.get("LienholderName"), state_mco_val=state_mco,
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
                    is_first_for_combo=False,
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
                }
                tpa_overrides.append(new_entry)
                if ask_yes_no("Save this decision for future runs? (y/n): "):
                    tpa_dirty.append(new_entry)
                continue

        pending_tpa_df = working_df.loc[pending_idx].copy()
        pending_blank_fields_df = working_df.loc[blank_field_idx].copy()
        if not pending_tpa_df.empty:
            # Blank column the user can fill in offline with the REAL TPA plan name
            # (instead of hand-typing LienholderId/CollectorId/LienType) whenever they
            # already know which TPA plan a skipped/unresolved raw text refers to --
            # resolved on resume by aliasing the raw text to that name (see the
            # Pending_TPA_Review resume-preseed block's own "Picked LienholderName"
            # handling), the exact same mechanism the live number-pick flow already
            # uses, so the row resolves via the normal TPA cascade, not a special path.
            pending_tpa_df["Picked LienholderName"] = None
        for _pdf in (pending_tpa_df, pending_blank_fields_df):
            if not _pdf.empty:
                _pdf["TicketId"] = ticket_id
                _pdf["ResponseDate"] = RESPONSE_DATE
                _pdf["ResponseType"] = RESPONSE_TYPE
        resolved_df = working_df.drop(index=pending_idx + blank_field_idx).copy()
        return resolved_df, pending_tpa_df, pending_blank_fields_df, tpa_dirty

    yes_with_mco_df, pending_tpa_df, pending_tpa_blank_fields_df, _TPA_DIRTY = resolve_tpa_for_run(
        yes_with_mco_df, TPA_OVERRIDES, TicketId, _ALIAS_SAVES
    )
    _TPA_DIRTY = _RESUME_TPA_DIRTY + _TPA_DIRTY
    # No live prompt generates a fresh equivalence save mid-run (unlike Scope/TPA) --
    # the only source is a resumed Pending_Lien_Match_Review "YES" decision.
    _LIEN_EQUIV_SAVES = _RESUME_LIEN_EQUIV_SAVES

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
        merge_keys = ["ClientId", "Project_Name", "Scope_Classification", "LienholderId", "CollectorId"]
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
                # Real bug found (user report): by this point in the pipeline,
                # Start_Date/End_Date have ALREADY been run through format_date
                # (real datetimes -> "MM/DD/YYYY" display strings, done once, right
                # after the very first plan-row grouping, long before this merge
                # ever runs). safe_min/safe_max expect raw datetimes (or the
                # literal "Bad Date"/None sentinels) -- fed already-formatted
                # strings instead, every value in the group fails their own
                # isinstance(d, str) check, so it always returned "Bad Date" even
                # when BOTH dates being merged were perfectly valid. Re-parse each
                # formatted string back to a real datetime first (then re-format
                # the result the same way everywhere else does); a blank cell is
                # normalized to a real None first too -- pandas silently turns a
                # None sitting in an otherwise all-string column into a plain
                # float NaN (not NaT), which safe_min's own "is not None" check
                # doesn't catch, so a stray NaN reaching Python's min()/max()
                # against a real datetime raises TypeError.
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
            merged_rows.append(base_row)
            rows_to_drop.extend(group_df.index.tolist())

            # PLAN_FATE is seeded (right after normalize_mco_name, long before this
            # merge runs) with one placeholder entry per original, pre-merge raw
            # MCO_Name. Step 4 only ever records a fate against the new merged key, so
            # the original keys would otherwise stay stuck at their placeholder
            # forever, and Remarks would show fake "tracking gap" warnings alongside
            # the one real outcome. Drop the superseded original keys here -- nothing
            # is lost, since the merged key's own plan-name text already lists every
            # original raw value, and it still gets a real fate recorded downstream
            # like any other row.
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
            cfg_ws = cfg_wb["Scope_Map"]
            header = [c.value for c in cfg_ws[1]]
            text_col  = header.index("ScopeText") + 1
            class_col = header.index("Classification") + 1
            existing_rows = {}
            for r in range(2, cfg_ws.max_row + 1):
                key = _normalize_key(cfg_ws.cell(row=r, column=text_col).value)
                if key:
                    existing_rows[key] = r
            for scope_text, classification in scope_saves:
                key = _normalize_key(scope_text)
                if key in existing_rows:
                    _old_classification = cfg_ws.cell(row=existing_rows[key], column=class_col).value
                    cfg_ws.cell(row=existing_rows[key], column=class_col, value=classification)
                    _audit_entries.append(_audit_row(
                        "Scope_Map", scope_text,
                        _old_classification if _old_classification not in (None, "") else "(was blank)",
                        classification, TicketId,
                    ))
                else:
                    new_row = cfg_ws.max_row + 1
                    cfg_ws.cell(row=new_row, column=text_col, value=scope_text)
                    cfg_ws.cell(row=new_row, column=class_col, value=classification)
                    existing_rows[key] = new_row
                    _audit_entries.append(_audit_row("Scope_Map", scope_text, "(new record)", classification, TicketId))
            print(f"\n✅ Saved {len(scope_saves)} scope classification update(s) to Scope_Map.")

        if tpa_dirty:
            cfg_ws2 = cfg_wb["TPA_Overrides"]
            header2 = [c.value for c in cfg_ws2[1]]
            col_idx = {name: i + 1 for i, name in enumerate(header2)}

            existing_rows2 = list(range(2, cfg_ws2.max_row + 1))

            def _find_target_row(entry):
                # Mirrors _tpa_overrides_find's own matching rule -- a COMPLETE existing
                # row (TortId filled in) must match State MCO + TortName + LienholderName
                # + Scope, so two different Torts that happen to share the same
                # LienholderId/LienholderName/State (a real case -- one collector/plan
                # handling several torts) get separate rows instead of colliding and
                # overwriting each other. A PENDING row (TortId still blank, e.g. a
                # not-yet-completed legacy migrated entry) matches on State MCO +
                # LienholderName alone, so first-use completion still patches it in
                # place. Keying only on LienholderId+LienholderName+State (as before)
                # would let a second Tort's entry silently clobber the first Tort's row.
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
                    if row_tort_id is None:
                        if pending_row is None:
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
            cfg_ws3 = cfg_wb["Alias_Map"]
            header3 = [c.value for c in cfg_ws3[1]]
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
            cfg_ws4 = cfg_wb["Lien_Name_Equivalence_Map"]
            header4 = [c.value for c in cfg_ws4[1]]
            col_idx4 = {name: i + 1 for i, name in enumerate(header4)}
            existing_equiv_rows = set()
            for r in range(2, cfg_ws4.max_row + 1):
                key = (
                    _normalize_key(cfg_ws4.cell(row=r, column=col_idx4["State MCO"]).value),
                    _normalize_key(cfg_ws4.cell(row=r, column=col_idx4["CollectorId"]).value),
                    _normalize_key(cfg_ws4.cell(row=r, column=col_idx4["LienHolderId"]).value),
                    _normalize_key(cfg_ws4.cell(row=r, column=col_idx4["EquivalentName"]).value),
                    _normalize_key(cfg_ws4.cell(row=r, column=col_idx4["Decision"]).value),
                )
                if key[0] and key[1] and key[2] and key[3]:
                    existing_equiv_rows.add(key)
            _new_equiv_rows = 0
            for state_mco, collector_id, lienholder_id, lienholder_name, equiv_name, decision, ticket_id in lien_equiv_saves:
                key = (_normalize_key(state_mco), _normalize_key(collector_id), _normalize_key(lienholder_id),
                       _normalize_key(equiv_name), _normalize_key(decision))
                if key in existing_equiv_rows:
                    continue
                new_row = cfg_ws4.max_row + 1
                cfg_ws4.cell(row=new_row, column=col_idx4["State MCO"], value=state_mco)
                cfg_ws4.cell(row=new_row, column=col_idx4["CollectorId"], value=collector_id)
                cfg_ws4.cell(row=new_row, column=col_idx4["LienHolderId"], value=lienholder_id)
                cfg_ws4.cell(row=new_row, column=col_idx4["LienHolderName"], value=lienholder_name)
                cfg_ws4.cell(row=new_row, column=col_idx4["EquivalentName"], value=equiv_name)
                cfg_ws4.cell(row=new_row, column=col_idx4["Decision"], value=decision)
                cfg_ws4.cell(row=new_row, column=col_idx4["Ticket #"], value=ticket_id)
                existing_equiv_rows.add(key)
                _new_equiv_rows += 1
                # Append-only sheet -- a repeat of an already-recorded decision is
                # silently skipped above (the `continue`), never reaches here, so every
                # audit row below really is a brand-new record, never an overwrite.
                _audit_entries.append(_audit_row(
                    "Lien_Name_Equivalence_Map",
                    f"State={state_mco} | CollectorId={collector_id} | LienHolderId={lienholder_id} | Candidate={equiv_name}",
                    "(new record)", f"Decision={decision}", ticket_id,
                ))
            print(f"\n✅ Saved {_new_equiv_rows} lien-name equivalence update(s) to Lien_Name_Equivalence_Map.")

        if column_map_saves:
            cfg_ws5 = cfg_wb["Column_Map"]
            header5 = [c.value for c in cfg_ws5[1]]
            col_idx5 = {name: i + 1 for i, name in enumerate(header5)}
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
                mco_cols_str = ", ".join(mco_cols)
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
            audit_ws = cfg_wb["Audit_Trail"]
            _ensure_audit_trail_header(audit_ws)
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
    # EXCLUDE FINALIZED CLIENTS
    # =====================================================

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
        collector_id = _stringify_id(collector_id)
        if collector_id not in _assigned_user_cache:
            q = "select Id, defaultassigneduserid from collectors where Id = ?"
            result = pd.read_sql(q, conn, params=[collector_id])
            _assigned_user_cache[collector_id] = result["defaultassigneduserid"].iloc[0] if not result.empty else None
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
    yes_with_mco_df = yes_with_mco_df[base_cols + date_cols_final + existing_cols + ["LienholderId", "CollectorId", "LienType"]]

    # Pre-fetch every distinct CollectorId this run needs in ONE round trip -- without
    # this, get_assigned_user_id's own per-id cache still means one separate SQL round
    # trip per DISTINCT CollectorId the .apply() below hits (real, measured cost: ~4s
    # for just a handful of distinct collectors in one run).
    _distinct_collector_ids = [
        _stringify_id(c) for c in yes_with_mco_df["CollectorId"].dropna().unique()
        if str(c).strip() != ""
    ]
    if _distinct_collector_ids:
        _bulk_collectors_df = pd.read_sql(
            f"select Id, defaultassigneduserid from collectors where Id in {make_sql_in(_distinct_collector_ids)}",
            conn
        )
        for _, _bc_row in _bulk_collectors_df.iterrows():
            _assigned_user_cache[_stringify_id(_bc_row["Id"])] = _bc_row["defaultassigneduserid"]
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

    def classify_step4_scenario(client_liens_df, state, resolved_collector_id, resolved_lienholder_id, resolved_name):
        """client_liens_df: sql_liens_df already filtered to this ClientId AND statusid==1.
        Returns (scenario, matched_row_or_None), scenario in 'A'/'B1'/'B2'/'C'/'T'. C is
        the plain catch-all (not "neither matches" -- that wording had a real gap:
        CollectorId matches, name doesn't, LienholderId isn't a Placeholder means a
        genuinely different, already-correct plan on file that must be left alone,
        while OUR target plan still needs creating). 'T' (Tie) is ambiguity between
        2+ of the client's OWN liens under the same CollectorId -- routed to manual
        review (Pending_Lien_Match_Review), matched_row is the tied-candidates DataFrame.

        LIEN_NAME_EQUIVALENCE_MAP / LIEN_NAME_EXCLUDED_MAP (keyed by (State,
        CollectorId, LienholderId)) are human-confirmed facts about the PLAN, not the
        client -- one client's decision applies to every future client with the same
        plan."""
        if client_liens_df.empty:
            return "C", None
        collector_matches = client_liens_df[client_liens_df["CollectorId"].astype(str) == str(resolved_collector_id)]
        if collector_matches.empty:
            return "C", None

        equiv_key = (str(state), str(resolved_collector_id), str(resolved_lienholder_id))

        # Existing-lien names already confirmed NOT the same real plan as this
        # LienholderId (a resumed "New" decision) are removed from consideration
        # entirely, before any matching runs -- this is what actually breaks a tie
        # for every FUTURE client sharing the same coincidental name clash, not
        # just the one where a human first reviewed it.
        excluded_names = {_normalize_key(n) for n in LIEN_NAME_EXCLUDED_MAP.get(equiv_key, [])}
        if excluded_names:
            collector_matches = collector_matches[
                ~collector_matches["LienHolderName"].apply(_normalize_key).isin(excluded_names)
            ]
            if collector_matches.empty:
                return "C", None

        # Try the TPA-resolved name first (the authoritative one), then any name a
        # human has already confirmed as equivalent for this exact plan
        # (LIEN_NAME_EQUIVALENCE_MAP) -- an EXTRA name to try, never a replacement,
        # and never a bypass of tie-detection: an equivalence name that itself ties
        # against the client's own liens is just as genuinely ambiguous as the
        # original name would have been.
        names_to_try = [resolved_name] + LIEN_NAME_EQUIVALENCE_MAP.get(equiv_key, [])
        first_tie_ids = None
        for name in names_to_try:
            matched_id, tied_ids = _match_name(name, collector_matches, "LienHolderName", "LienId", return_tie=True)
            if matched_id is not None:
                row = collector_matches[collector_matches["LienId"] == matched_id].iloc[0]
                return "A", row
            if tied_ids and first_tie_ids is None:
                first_tie_ids = tied_ids

        if first_tie_ids is not None:
            return "T", collector_matches[collector_matches["LienId"].isin(first_tie_ids)]

        for _, row in collector_matches.iterrows():
            pair = (str(row["CollectorId"]), str(row["LienholderId"]))
            if pair in PLACEHOLDER_PAIRS:
                return ("B1" if row["Stage"] == "Final No Entitlement" else "B2"), row
        return "C", None

    def build_note(template_key, ticket_id, plan_name, state=None, response_type=None, response_date=None, has_dates=False):
        """A/B1/B2 only mention "Enrollment Dates" when has_dates is actually True --
        the note must describe exactly what got updated, not more."""
        dated_templates = {
            "A":  "Updated IsMCO and Enrollment Dates per instruction on ticket #{ticket}. Plan Name: {plan}",
            "B1": "Updated lienholder Information, stage from Final No Entitlement, OnBenefit details and Enrollment Dates per instruction on ticket #{ticket}. Plan Name: {plan}",
            "B2": "Updated LH, IsMCO and Enrollment Dates per instruction on ticket #{ticket}. Plan Name: {plan}",
        }
        undated_templates = {
            "A":  "Updated IsMCO per instruction on ticket #{ticket}. Plan Name: {plan}",
            "B1": "Updated lienholder Information, stage from Final No Entitlement, OnBenefit details per instruction on ticket #{ticket}. Plan Name: {plan}",
            "B2": "Updated LH and IsMCO per instruction on ticket #{ticket}. Plan Name: {plan}",
        }
        other_templates = {
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
    _FETCH_NEW_LIEN_IDS_QUERY_TEMPLATE = """select ClientSSN, ClientId, Id, Stage, LienType, CollectorId, LienholderId, AssignedUserId, statusid, LienHolderName, finalizedstatusname
from fullproductviews
where ClientId in {client_ids}
Order by 5,2"""
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
        "Decision",
    ]

    def build_step4_outputs(df_in, liens_df, ticket_id, response_date, response_type, already_claimed_lien_ids=None):
        if already_claimed_lien_ids is None:
            already_claimed_lien_ids = {}
        rows_a, rows_b1, rows_b2 = [], [], []
        rows_scope_upload, rows_update1, rows_pending_lien_match = [], [], []
        updated_client_ids = set()
        # This run's own newly-claimed LienIds, PER CLIENT -- carried forward (see
        # already_claimed_lien_ids above, and _ALREADY_CLAIMED_LIEN_IDS_CARRY at the
        # call site) onto any of this client's OTHER plans that go to Pending_* this
        # run, so a LATER resume round never matches the SAME already-claimed lien
        # again for a different plan.
        claimed_lien_ids_by_client = {}

        # A client can have MULTIPLE MCO plans in the same run, and classify_step4_scenario
        # re-derives client_liens fresh from the untouched liens_df on every iteration --
        # without tracking which LienIds an EARLIER plan (this same loop, same client) has
        # already matched/claimed, a second plan could independently match the exact same
        # existing/placeholder lien row that an earlier plan already claimed, when it should
        # instead find nothing left and go to Scenario C.
        claimed_lien_ids = set()

        # 3 new Yes_with_MCO summary columns -- which existing LienId (if any) this
        # exact plan touched this run, that lien's current Stage, and a short,
        # structured description of what actually happened. Populated per-row
        # below, inside the SAME scenario classification already happening in this
        # loop (no separate pass, no duplicate logic) -- df_in IS the caller's own
        # yes_with_mco_df (passed by reference, not a copy), so writing onto it
        # here is visible to the caller once this function returns. New columns
        # land at the END of yes_with_mco_df's existing columns (a trailing
        # summary), since FINAL COLUMN ORDER's own column selection already ran
        # BEFORE Step 4 -- nothing downstream re-selects columns in a way that
        # would drop them.
        df_in["LienIds used for updates"] = None
        df_in["Stage"] = None
        df_in["Remarks"] = None

        def _join_update_fields(fields):
            """"Update X" / "Update X & Y" / "Update X, Y & Z" -- canonical field
            order is always [Enrollment dates, FNI, LH, IsMCO], already filtered by
            the caller to just the ones actually applicable+present."""
            if len(fields) == 1:
                return f"Update {fields[0]}"
            return f"Update {', '.join(fields[:-1])} & {fields[-1]}"

        for idx, r in df_in.iterrows():
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
            scenario, matched_row = classify_step4_scenario(client_liens, state, collector_id, lienholder_id, plan_name)

            if scenario == "A":
                claimed_lien_ids.add(matched_row["LienId"])
                claimed_lien_ids_by_client.setdefault(client_id, set()).add(matched_row["LienId"])
                # A client's individual liens can already be correct on file while
                # RequiredMedicaidMCOLienCount on the client record is still stuck at
                # the generic starting value or blank, never refreshed to the real
                # total -- so this client must still be counted in Client-MCO Cnt
                # Update even when nothing here needs an MCO Update 2 row. Added
                # unconditionally for BOTH branches below, not just the "needs an
                # update" one.
                updated_client_ids.add(client_id)
                df_in.at[idx, "LienIds used for updates"] = matched_row["LienId"]
                df_in.at[idx, "Stage"] = matched_row["Stage"]
                if not matched_row["IsMCO"] and matched_row["Stage"] != "Final No Entitlement":
                    row = {"Id": matched_row["LienId"], "IsMCO": True,
                           "NewLienNote": build_note("A", ticket_id, plan_name, has_dates=has_dates)}
                    if has_dates:
                        row["Enrollmentstartdate"], row["EnrollmentEndDate"] = start_date, end_date
                    rows_a.append(row)
                    PLAN_FATE[fate_key] = "Matched existing lien on file -- see MCO Update 2"
                    _a_fields = (["Enrollment dates"] if has_dates else []) + ["IsMCO"]
                    df_in.at[idx, "Remarks"] = _join_update_fields(_a_fields)
                else:
                    PLAN_FATE[fate_key] = "Matched existing lien already correctly on file; no update needed"
                    df_in.at[idx, "Remarks"] = "Already correct- Only update count"

            elif scenario == "B1":
                claimed_lien_ids.add(matched_row["LienId"])
                claimed_lien_ids_by_client.setdefault(client_id, set()).add(matched_row["LienId"])
                df_in.at[idx, "LienIds used for updates"] = matched_row["LienId"]
                df_in.at[idx, "Stage"] = matched_row["Stage"]
                row = {
                    "Id": matched_row["LienId"], "Stage": "Final - No Interest",
                    "LienholderId": lienholder_id, "IsMCO": True,
                    "Onbenefits": "Yes", "OnBenefitsVerified": response_date,
                    "LienDate1": response_date, "LienAmount1": 0, "AuditedAmount1": 0,
                    "FinalDemandDate": response_date, "FinalDemandAmount": 0,
                    "NewLienNote": build_note("B1", ticket_id, plan_name, has_dates=has_dates),
                }
                if has_dates:
                    row["Enrollmentstartdate"], row["EnrollmentEndDate"] = start_date, end_date
                rows_b1.append(row)
                updated_client_ids.add(client_id)
                PLAN_FATE[fate_key] = "Placeholder lien (Final No Entitlement) filled in and finalized -- see MCO Update 2"
                _b1_fields = (["Enrollment dates"] if has_dates else []) + ["FNI", "LH", "IsMCO"]
                df_in.at[idx, "Remarks"] = _join_update_fields(_b1_fields)

            elif scenario == "B2":
                claimed_lien_ids.add(matched_row["LienId"])
                claimed_lien_ids_by_client.setdefault(client_id, set()).add(matched_row["LienId"])
                df_in.at[idx, "LienIds used for updates"] = matched_row["LienId"]
                df_in.at[idx, "Stage"] = matched_row["Stage"]
                row = {"Id": matched_row["LienId"], "LienholderId": lienholder_id, "IsMCO": True,
                       "NewLienNote": build_note("B2", ticket_id, plan_name, has_dates=has_dates)}
                if has_dates:
                    row["Enrollmentstartdate"], row["EnrollmentEndDate"] = start_date, end_date
                rows_b2.append(row)
                updated_client_ids.add(client_id)
                PLAN_FATE[fate_key] = "Placeholder lien filled in with lienholder details -- see MCO Update 2"
                _b2_fields = (["Enrollment dates"] if has_dates else []) + ["LH", "IsMCO"]
                df_in.at[idx, "Remarks"] = _join_update_fields(_b2_fields)

            elif scenario == "T":
                # Genuine ambiguity among the client's OWN existing liens (2+ tie on
                # name) -- refuse to guess. Nothing is claimed or counted this run for
                # THIS plan; a row per tied candidate goes to Pending_Lien_Match_Review
                # so the user can pick the right one (or confirm none apply) once, with
                # that decision remembered going forward via
                # LIEN_NAME_EQUIVALENCE_MAP -- see classify_step4_scenario.
                for _, cand_row in matched_row.iterrows():
                    rows_pending_lien_match.append({
                        "ClientId": client_id, "LienId": lien_id, "MCO_Name": plan_name,
                        "LienholderId": lienholder_id, "CollectorId": collector_id,
                        "LienType": r["LienType"], "Scope_Classification": r.get("Scope_Classification"),
                        "State": state, "Start_Date": start_date, "End_Date": end_date,
                        "TicketId": ticket_id, "ResponseDate": response_date, "ResponseType": response_type,
                        "CandidateLienId": cand_row["LienId"], "CandidateLienHolderName": cand_row["LienHolderName"],
                        "CandidateCollectorId": cand_row["CollectorId"], "CandidateLienholderId": cand_row["LienholderId"],
                        "Decision": None,
                    })
                PLAN_FATE[fate_key] = "Ambiguous match with existing liens on file -- see Pending_Lien_Match_Review"
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

        return {
            "A": df_a, "B1": df_b1, "B2": df_b2,
            "client_scope_upload": df_scope, "mco_update1": df_update1,
            "pending_lien_match": df_pending_lien_match,
            "updated_client_ids": updated_client_ids,
            "claimed_lien_ids": claimed_lien_ids,
            "claimed_lien_ids_by_client": claimed_lien_ids_by_client,
        }

    def build_client_cnt_update(updated_client_ids, df_in, liens_df, claimed_lien_ids, ticket_id, response_date,
                                 response_type, already_counted=None):
        if already_counted is None:
            already_counted = {}
        if not updated_client_ids:
            return pd.DataFrame(columns=["Id", "RequiredMedicaidMCOLienCount", "NewClientNote"])
        per_client_plans = (
            df_in[df_in["ClientId"].isin(updated_client_ids)]
            .groupby("ClientId")["MCO_Name"].nunique().to_dict()
        )

        # A client can have OTHER active MCO liens this run's input never mentions --
        # those still count toward the required total, so add them on top of
        # per_client_plans. Liens Step 4 already claimed this run (Scenario A/B1/B2)
        # are excluded here -- already reflected via df_in's own MCO_Name count,
        # counting again would double it.
        #
        # A client carrying a prior-run count (already_counted[cid] > 0, see
        # _ALREADY_COUNTED_CARRY) is excluded from this scan entirely -- that count
        # already reflects everything as of the earlier run; re-scanning risks
        # double-counting once that run's MCO Update 2 lands in the DB. Only
        # genuinely NEW plans this run finds get added on top for such a client.
        _already_counted_client_ids = [cid for cid, cnt in already_counted.items() if cnt > 0]
        other_active_mco = liens_df[
            liens_df["ClientId"].isin(updated_client_ids)
            & ~liens_df["ClientId"].isin(_already_counted_client_ids)
            & liens_df["IsMCO"].fillna(False).astype(bool)
            & (liens_df["statusid"] == 1)
            & (~liens_df["LienId"].isin(claimed_lien_ids))
        ]
        for cid, extra in other_active_mco.groupby("ClientId")["LienId"].nunique().items():
            per_client_plans[cid] = per_client_plans.get(cid, 0) + extra

        for cid in list(per_client_plans.keys()):
            if already_counted.get(cid, 0) > 0:
                per_client_plans[cid] += already_counted[cid]

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
            # A client who already got a Client-MCO Cnt Update row (and its
            # NewClientNote) in an earlier run must not get a second, duplicate note
            # here -- only the count needs refreshing to the new, real total. (On a
            # merge/append resume run, this blank is corrected back to a real note
            # further downstream, in the merge-mode WRITE OUTPUT section -- see
            # _carried_blank_mask there -- since a merged file is the user's ONE
            # single upload for this whole ticket. _MERGE_MODE isn't decided yet at
            # this point in the script, so that correction can't happen here.)
            "NewClientNote": None if already_counted.get(cid, 0) > 0 else note,
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
    )
    pending_lien_match_df = _STEP4_RESULT["pending_lien_match"]
    client_cnt_update_df = build_client_cnt_update(
        _STEP4_RESULT["updated_client_ids"], yes_with_mco_df, sql_liens_df, _STEP4_RESULT["claimed_lien_ids"],
        TicketId, RESPONSE_DATE, RESPONSE_TYPE, already_counted=_ALREADY_COUNTED_CARRY
    )

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
    for _pdf in (pending_scope_df, pending_tpa_df, pending_tpa_blank_fields_df, pending_lien_match_df):
        if not _pdf.empty:
            _pdf["_ClientAlreadyCountedThisRun"] = _pdf["ClientId"].map(_new_carry_forward).fillna(0).astype(int)

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

    for _pdf in (pending_scope_df, pending_tpa_df, pending_tpa_blank_fields_df, pending_lien_match_df):
        if not _pdf.empty:
            _pdf["_ClientAlreadyClaimedLienIds"] = _pdf["ClientId"].apply(_claimed_lien_ids_str)

    # =====================================================
    # REMARKS COLUMN — accountability for every qualified (LienId, MCO_Name)
    # =====================================================
    # PLAN_FATE is only keyed by (LienId, MCO_Name) for plans that actually made it
    # into yes_with_mco_df -- a LienId with no qualified plan at all gets no Remarks.
    # An LA-file LienId whose SSN was in the Mismatch tab gets that noted TOO, but only
    # as a supplement to a real outcome, never standalone: a Mismatch flag with nothing
    # else to report (no plan text at all) isn't actionable on its own, since only MCO
    # plans matter here -- so it stays silent exactly like any other no-plan LienId.
    _lienid_to_remarks = {}
    for (_fate_lien_id, _fate_plan), _fate_remark in PLAN_FATE.items():
        _lienid_to_remarks.setdefault(str(_fate_lien_id), []).append(f"{_fate_plan}: {_fate_remark}")

    def _build_remarks(lien_id_val):
        parts = list(_lienid_to_remarks.get(str(lien_id_val), []))
        if parts and str(lien_id_val) in MISMATCH_FLAGGED_LIENIDS:
            parts.append("SSN listed in Mismatch tab")
        return " | ".join(parts)

    ev_df["Remarks"] = ev_df[LIENID_COL_IN_EVDF].apply(_build_remarks)

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
    if IS_RESUME_RUN:
        _divider()
        _merge_choice = ask_numbered_choice(
            ["Append into the original ticket's files", "Keep this run's files separate"],
            "How should this run's results be saved?",
            "Enter the number: "
        )
        _MERGE_MODE = "append" if _merge_choice.startswith("Append") else "separate"
    else:
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
        if existing_df.empty:
            return new_df.copy()
        if new_df.empty:
            return existing_df.copy()
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
    # Just enough of each scenario's own header (Id + one scenario-specific column) to
    # tell A/B1/B2 apart unambiguously, REGARDLESS of whether the date columns happen
    # to be present that round -- IsMCO vs Stage vs LienholderId never collide.
    _MCO_UPDATE2_SIGNATURES = {"A": ("Id", "IsMCO"), "B1": ("Id", "Stage"), "B2": ("Id", "LienholderId")}

    def _find_existing_mco_update2_blocks(ws):
        """Finds each scenario's block by its short signature, then reads the block's
        REAL header cells as they actually are on the sheet (whatever columns that
        round happened to write -- 3 or 5, doesn't matter), scanning right until the
        first blank header cell. Real bug found and fixed (user report): an earlier
        version matched against a FIXED canonical column list (with or without the 2
        date columns) -- when a later round's own list didn't match the EXACT same
        variant already on the sheet, values landed under the wrong header entirely
        (NewLienNote text appearing under "Enrollmentstartdate", NewLienNote itself
        blank). Reading the header back, whatever it is, instead of assuming a fixed
        shape removes that guesswork entirely. Returns {scenario_key: (start_col,
        end_col, last_used_row, header_names)}, 1-indexed columns (openpyxl
        convention)."""
        found = {}
        max_col = ws.max_column
        headers = {c: ws.cell(row=1, column=c).value for c in range(1, max_col + 1)}
        for key, (sig0, sig1) in _MCO_UPDATE2_SIGNATURES.items():
            for c in range(1, max_col + 1):
                if headers.get(c) == sig0 and headers.get(c + 1) == sig1:
                    header_names = []
                    cc = c
                    while headers.get(cc):
                        header_names.append(headers[cc])
                        cc += 1
                    end_col = cc - 1
                    last_row = 1
                    for r in range(2, ws.max_row + 1):
                        if any(ws.cell(row=r, column=x).value is not None for x in range(c, end_col + 1)):
                            last_row = r
                    found[key] = (c, end_col, last_row, header_names)
                    break
        return found

    def _merge_mco_update2_blocks(wb, sheet_name, new_blocks):
        """new_blocks: {scenario_key: DataFrame} for whichever of A/B1/B2 have new
        rows this round, exactly as build_step4_outputs produced them (2 date columns
        included or not, depending on whether anyone had a date THIS round -- no
        reindexing needed by the caller). An existing block's new rows are aligned to
        WHATEVER columns that block's OWN header already has -- a value this round
        has that the existing header doesn't (or vice versa) is simply not written,
        rather than forcing either side's shape onto the other. A scenario with no
        existing block gets a brand-new block, using THIS round's own columns as-is,
        placed after whatever's rightmost so far, 3 blank columns over -- no fixed
        A-then-B1-then-B2 ordering is enforced, since position doesn't matter, only
        that each block stays self-contained."""
        if sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            existing = _find_existing_mco_update2_blocks(ws)
            rightmost = max((end for _, end, _, _ in existing.values()), default=0)
        else:
            ws = wb.create_sheet(sheet_name)
            existing = {}
            rightmost = 0

        def _cell_val(v):
            return None if pd.isna(v) else v

        for key, df in new_blocks.items():
            if df.empty:
                continue
            if key in existing:
                start_col, _end_col, last_row, header_names = existing[key]
                for i, (_, row) in enumerate(df.iterrows()):
                    for j, col_name in enumerate(header_names):
                        val = row[col_name] if col_name in row.index else None
                        ws.cell(row=last_row + 1 + i, column=start_col + j, value=_cell_val(val))
            else:
                cols = [c for c in _MCO_UPDATE2_COL_SETS[key] if c in df.columns]
                start_col = rightmost + 3 + 1 if rightmost else 1
                for j, col_name in enumerate(cols):
                    ws.cell(row=1, column=start_col + j, value=col_name)
                for i, (_, row) in enumerate(df.iterrows()):
                    for j, col_name in enumerate(cols):
                        ws.cell(row=2 + i, column=start_col + j, value=_cell_val(row[col_name]))
                rightmost = start_col + len(cols) - 1

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
            "Yes_with_MCO", "Client Scope Upload", "Fetch New Lien IDs",
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

    def _highlight_note_cols(ws, df, col_names):
        """Same accountability-note highlight as Input_File_Original's Remarks column --
        these are notes THIS TOOL adds (summarizing an outcome/reason), not part of the
        original input/query data, so they should visually stand out."""
        _fill = _opxl.styles.PatternFill(start_color="FFE699", end_color="FFE699", fill_type="solid")
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

    if IS_RESUME_RUN and _MERGE_MODE == "append" and _merge_wip_path and os.path.exists(_merge_wip_path):
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
        _merge_xls.close()

        _merged_yes = _append_df(_existing_yes, _safe_str_df(yes_with_mco_df.drop(columns=["LienType", "State"])))
        # Real duplicate found (user report): a Pending_Lien_Match_Review row was
        # ALREADY TPA-resolved and already written to Yes_with_MCO in an earlier
        # round (Step 3 succeeded then too -- Step 4 only tied on WHICH existing-lien
        # scenario applied). On resume it re-enters the full pipeline and Step 3
        # re-derives the SAME (LienId, MCO_Name) match deterministically, by design
        # (see the resume-preseed block's own comment on this) -- so a plain append
        # duplicates it. LienId+MCO_Name uniquely identifies one real plan on one
        # real lien, so de-duping on that pair only removes genuine re-derivations,
        # never two DIFFERENT liens that happen to share the same plan text.
        if not _merged_yes.empty and {"LienId", "MCO_Name"} <= set(_merged_yes.columns):
            _merged_yes = _merged_yes.drop_duplicates(subset=["LienId", "MCO_Name"], keep="last").reset_index(drop=True)
        _merged_csu = _append_df(_existing_csu, _safe_str_df(_STEP4_RESULT["client_scope_upload"]))
        _merged_mu1 = _append_df(_existing_mu1, _safe_str_df(_STEP4_RESULT["mco_update1"]))

        # Real gap found (user report): client_cnt_update_df was built further up
        # (before this merge-mode branch -- and before IS_RESUME_RUN/_MERGE_MODE
        # even run their own upfront prompt, so passing suppress_carried_note there
        # isn't possible without reordering prompts) with the DEFAULT carried-note
        # suppression always on -- a client already counted in an earlier round got
        # a blank NewClientNote here regardless of merge/separate mode. But a merged
        # file is the user's ONE single upload for this whole ticket, so every
        # client in it should show a real note, not just a count refresh, no matter
        # which round first resolved them. Only "separate" mode (handled below,
        # untouched) keeps that suppression.
        _carried_ids = {cid for cid, cnt in _ALREADY_COUNTED_CARRY.items() if cnt > 0}
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
                                      ["LienIds used for updates", "Stage", "Remarks"])

            if not _merged_csu.empty:
                _merged_csu.to_excel(writer, sheet_name="Client Scope Upload", index=False)
                _fetch_client_ids = make_sql_in(_merged_csu["ClientId"].dropna().unique())
                _fetch_query = _FETCH_NEW_LIEN_IDS_QUERY_TEMPLATE.format(client_ids=_fetch_client_ids)
                if "Fetch New Lien IDs" in writer.book.sheetnames:
                    del writer.book["Fetch New Lien IDs"]
                _fetch_ws = writer.book.create_sheet("Fetch New Lien IDs")
                for _i, _line in enumerate(_fetch_query.split("\n"), start=1):
                    _fetch_ws.cell(row=_i, column=1, value=_line)

            if not _merged_mu1.empty:
                _merged_mu1.to_excel(writer, sheet_name="MCO Update 1", index=False)

            if not _merged_cnt.empty:
                _merged_cnt.to_excel(writer, sheet_name="Client-MCO Cnt Update", index=False)

            if not _merged_yes.empty:
                # Real gap found (user report): these queries are a ready-to-run
                # reference for the user's OWN convenience after everything in this
                # merged file gets uploaded -- and a merged file is the user's ONE
                # single upload for the whole ticket, so the ClientId list must
                # cover EVERY client across EVERY merged round (_merged_yes, not
                # just this round's own yes_with_mco_df), in ONE single tab --
                # never a round-suffixed snapshot. Same delete-and-recreate pattern
                # "Fetch New Lien IDs" already uses for exactly this reason.
                _post_client_ids = make_sql_in(_merged_yes["ClientId"].dropna().unique())

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

            _merge_mco_update2_blocks(writer.book, "MCO Update 2", {
                "A": _STEP4_RESULT["A"], "B1": _STEP4_RESULT["B1"], "B2": _STEP4_RESULT["B2"],
            })

            _ifo_sheet_name = _sheet_name_with_round_suffix("Input_File_Original", _round_n)
            ev_df.to_excel(writer, sheet_name=_ifo_sheet_name, index=False)
            # Same Remarks-column highlight as the normal (non-merge) write below --
            # missed on the first pass of this feature (real gap, user report).
            _ifo_ws = writer.book[_ifo_sheet_name]
            _remarks_col = list(ev_df.columns).index("Remarks") + 1
            _remarks_fill = _opxl.styles.PatternFill(start_color="FFE699", end_color="FFE699", fill_type="solid")
            for _r in range(1, len(ev_df) + 2):
                _remarks_cell = _ifo_ws.cell(row=_r, column=_remarks_col)
                _remarks_cell.fill = _remarks_fill
                if _r == 1:
                    _remarks_cell.font = _opxl.styles.Font(bold=True)

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

            for _tab_name in ("Client Scope Upload", "MCO Update 1", "MCO Update 2", "Client-MCO Cnt Update"):
                if _tab_name in writer.book.sheetnames:
                    writer.book[_tab_name].sheet_properties.tabColor = "1F9D55"
            for _tab_name in (
                _sheet_name_with_round_suffix("MCOLienCount_Disc", _round_n),
                _sheet_name_with_round_suffix("Finalized_Clients_Excluded", _round_n),
            ):
                if _tab_name in writer.book.sheetnames:
                    writer.book[_tab_name].sheet_properties.tabColor = "C00000"

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
            _ifo_ws = writer.book["Input_File_Original"]
            _remarks_col = list(ev_df.columns).index("Remarks") + 1
            _remarks_fill = _opxl.styles.PatternFill(start_color="FFE699", end_color="FFE699", fill_type="solid")
            for _r in range(1, len(ev_df) + 2):
                _remarks_cell = _ifo_ws.cell(row=_r, column=_remarks_col)
                _remarks_cell.fill = _remarks_fill
                if _r == 1:
                    _remarks_cell.font = _opxl.styles.Font(bold=True)

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
            _ywm_display_df = yes_with_mco_df.drop(columns=["LienType", "State"])
            _ywm_display_df.to_excel(writer, sheet_name="Yes_with_MCO", index=False)
            _highlight_note_cols(writer.book["Yes_with_MCO"], _ywm_display_df,
                                  ["LienIds used for updates", "Stage", "Remarks"])

            # Tab order: Client Scope Upload -> Fetch New Lien IDs -> MCO Update 1 -> MCO
            # Update 2 -> Client-MCO Cnt Update (last) -- the natural reading order of the
            # creation/update flow.
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

            # Give the 4 "action" tabs (the actual creation/update output, as opposed to the
            # audit/diagnostic tabs above) a shared accent color so they visually stand apart.
            for _tab_name in ("Client Scope Upload", "MCO Update 1", "MCO Update 2", "Client-MCO Cnt Update"):
                if _tab_name in writer.book.sheetnames:
                    writer.book[_tab_name].sheet_properties.tabColor = "1F9D55"

            # The 2 discrepancy/exclusion tabs (rows flagged or filtered OUT of normal
            # processing) get a red tab color whenever they actually appear, so they catch
            # the eye immediately instead of blending in with the plain audit tabs.
            for _tab_name in ("MCOLienCount_Disc", "Finalized_Clients_Excluded"):
                if _tab_name in writer.book.sheetnames:
                    writer.book[_tab_name].sheet_properties.tabColor = "C00000"


    # Standalone CSV exports of the 4 "action" tabs -- these are what actually gets
    # uploaded/imported downstream, so they're handed off separately, not just buried
    # as one tab inside the WIP workbook. MCO Update 2 has no single CSV equivalent:
    # its 3 scenario blocks (A/B1/B2) have different, unrelated column sets, laid out
    # side-by-side in the workbook only for a human reading it -- so each scenario
    # that actually produced rows this run gets its own CSV instead, named for what
    # that scenario actually updates (A = just IsMCO+dates, B1 = the Final No
    # Entitlement full rewrite, B2 = LienholderId & IsMCO+dates).
    _csv_exports = {
        "Client Scope Upload": _STEP4_RESULT["client_scope_upload"],
        "MCO Update 1": _STEP4_RESULT["mco_update1"],
        "MCO Update 2 - IsMCO": _STEP4_RESULT["A"],
        "MCO Update 2 - FNI": _STEP4_RESULT["B1"],
        "MCO Update 2 - LH & IsMCO": _STEP4_RESULT["B2"],
        "Client-MCO Cnt Update": client_cnt_update_df,
    }
    if any(not df.empty for df in _csv_exports.values()):
        _csv_dir = os.path.join(_output_dir, "Upload_Update CSVs")
        os.makedirs(_csv_dir, exist_ok=True)
        for _csv_name, _csv_df in _csv_exports.items():
            if _csv_df.empty:
                continue
            if IS_RESUME_RUN and _MERGE_MODE == "separate":
                # Plain tab-name filenames used to silently overwrite an earlier
                # round's own CSV of the same name (real user report) -- a round
                # suffix keeps every round's CSVs coexisting in the same folder.
                _csv_path = os.path.join(_csv_dir, f"{_csv_name}-R{_resume_round_number(input_filename)}.csv")
            else:
                _csv_path = os.path.join(_csv_dir, f"{_csv_name}.csv")
                if IS_RESUME_RUN and _MERGE_MODE == "append" and os.path.exists(_csv_path):
                    _ensure_file_hydrated(_csv_path)
                    _existing_csv_df = _retry_on_permission_error(
                        lambda: pd.read_csv(_csv_path, dtype=str), "the existing CSV file"
                    )
                    if _csv_name == "Client-MCO Cnt Update":
                        _csv_df = _upsert_client_cnt_update(_existing_csv_df, _safe_str_df(_csv_df))
                    else:
                        _csv_df = _append_df(_existing_csv_df, _safe_str_df(_csv_df))
            _ensure_writable(_csv_path)
            _csv_df.to_csv(_csv_path, index=False)


    # Review items (Pending_Scope_Review / Pending_TPA_Review / Pending_TPA_Blank_Fields /
    # Pending_Lien_Match_Review) go into their OWN separate spreadsheet, not the main
    # processing output -- give this file's path back in as `Path of Import File`
    # whenever ready to resolve them.
    _any_pending = (not pending_scope_df.empty or not pending_tpa_df.empty
                    or not pending_tpa_blank_fields_df.empty or not pending_lien_match_df.empty)
    if _any_pending:
        _ensure_writable(REVIEW_FILE)
        with pd.ExcelWriter(REVIEW_FILE, engine="openpyxl") as review_writer:
            if not pending_scope_df.empty:
                pending_scope_df.to_excel(review_writer, sheet_name="Pending_Scope_Review", index=False)
            if not pending_tpa_df.empty:
                # These 5 columns are only ever populated on the Pending_TPA_Blank_Fields
                # path (_finalize_tpa_match's blank-routing branch) -- a Pending_TPA_Review
                # row never reaches that branch (it's here because Tort/Name/Scope
                # matching itself failed), so they're structurally always blank on this
                # sheet, and never read back by anything either. MCOn_Existing columns
                # are dropped too: existing on-file plans don't help identify the
                # correct LienholderId/CollectorId/LienType for a plan that failed to
                # auto-match.
                _tpa_review_drop_cols = [
                    "TPA_Match_TortId", "TPA_Match_StateMCO", "TPA_Match_TortName",
                    "TPA_Match_Scope", "TPA_Match_LienholderName",
                ] + [c for c in pending_tpa_df.columns if c.startswith("MCO") and c.endswith("_Existing")]
                pending_tpa_df.drop(columns=_tpa_review_drop_cols, errors="ignore").to_excel(
                    review_writer, sheet_name="Pending_TPA_Review", index=False
                )
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
                pending_tpa_blank_fields_df.drop(columns=_blank_fields_drop_cols, errors="ignore").to_excel(
                    review_writer, sheet_name="Pending_TPA_Blank_Fields", index=False
                )
            if not pending_lien_match_df.empty:
                pending_lien_match_df.to_excel(review_writer, sheet_name="Pending_Lien_Match_Review", index=False)
        print(f"\n📋 Review file (items needing manual attention): {REVIEW_FILE}")

    _divider()
    save_config_updates(CONFIG_FILE, _SCOPE_SAVES, _TPA_DIRTY, _ALIAS_SAVES, _LIEN_EQUIV_SAVES, _COLUMN_MAP_SAVES)

    # Local, readable copy of every config/TPA decision actually recorded this run (same
    # rows as the centralized Audit_Trail sheet in mco_config.xlsx) -- a convenience next
    # to the run log, not a replacement for it: the centralized sheet is the one place
    # that's always reachable from a suspicious row regardless of which folder/run
    # produced it; this is for reading "what did THIS run decide" without having to
    # dig through the full terminal transcript in log_{...}.txt.
    if _AUDIT_TRAIL_THIS_RUN:
        _audit_wb_local = _opxl.Workbook()
        _audit_ws_local = _audit_wb_local.active
        _audit_ws_local.title = "Audit_Trail"
        for _i, _col_name in enumerate(_AUDIT_TRAIL_LOCAL_COLUMNS, start=1):
            _audit_ws_local.cell(row=1, column=_i, value=_col_name)
        for _r, _entry in enumerate(_AUDIT_TRAIL_THIS_RUN, start=2):
            for _i, _col_name in enumerate(_AUDIT_TRAIL_LOCAL_COLUMNS, start=1):
                _audit_ws_local.cell(row=_r, column=_i, value=_entry.get(_col_name))
        _audit_wb_local.save(_run_audit_path)
        print(f"📝 This run's config/TPA decisions also recorded in: {_run_audit_path}")

    print("\n✅ ===============================")
    print("✅ AUTOMATION COMPLETED SUCCESSFULLY")
    print(f"✅ Output file: {OUTPUT_FILE}")
    print("✅ ===============================")

_offer_rerun("\nPress 'R' and Enter to process another file, or just press Enter to exit: ")