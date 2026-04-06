import subprocess
import time
import os
import sys
import json
import logging
from pathlib import Path
from datetime import datetime

# ── Single-instance guard ─────────────────────────────────────────────────────
if sys.platform == 'win32':
    import ctypes
    _mutex = ctypes.windll.kernel32.CreateMutexW(None, True, "Global\\ActivityXController")
    if ctypes.windll.kernel32.GetLastError() == 183:  # ERROR_ALREADY_EXISTS
        sys.exit(0)
elif sys.platform == 'darwin':
    import fcntl
    _lock_file = open('/tmp/activityx_controller.lock', 'w')
    try:
        fcntl.flock(_lock_file, fcntl.LOCK_EX | fcntl.LOCK_NB)
    except IOError:
        sys.exit(0)

# ── App data directory ────────────────────────────────────────────────────────
def _get_app_dir():
    if sys.platform == 'win32':
        base = Path(os.environ.get('LOCALAPPDATA', Path.home() / 'AppData' / 'Local'))
    elif sys.platform == 'darwin':
        base = Path.home() / 'Library' / 'Application Support'
    else:
        base = Path.home() / '.local' / 'share'
    d = base / 'ActivityX'
    d.mkdir(parents=True, exist_ok=True)
    return d

APP_DIR = _get_app_dir()

# ── Logging ───────────────────────────────────────────────────────────────────
logging.basicConfig(
    filename=str(APP_DIR / 'controller.log'),
    level=logging.INFO,
    format='%(asctime)s %(levelname)s: %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S',
)

# ── Supabase import ──────────────────────────────────────────────────────────
try:
    from supabase import create_client, Client
except ImportError:
    create_client = None

# ── Config ────────────────────────────────────────────────────────────────────
SUPABASE_URL = os.getenv('SUPABASE_URL') or 'https://btkiqffcjvjyqyokccfh.supabase.co'
SUPABASE_KEY = os.getenv('SUPABASE_KEY') or 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImJ0a2lxZmZjanZqeXF5b2tjY2ZoIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NTA0ODc3OTQsImV4cCI6MjA2NjA2Mzc5NH0.9OHJcPdD-GIpBiUZuLl8NySwj5e0W4-JtV1u2o5LA9U'

# Stale threshold: if last_alive.txt is older than this, force-kill tracker
STALE_THRESHOLD_SECONDS = 600  # 10 minutes

# Auto-update config
GITHUB_REPO = "Orceus/activity-tracker"
UPDATE_CHECK_INTERVAL = 21600  # 6 hours
MAX_CRASH_COUNT = 3  # Rollback after this many crashes in 5 minutes


def get_local_version():
    """Read current version from version.txt — check AppData then exe's own folder."""
    for base in [APP_DIR, Path(os.path.dirname(sys.executable if getattr(sys, 'frozen', False) else os.path.abspath(__file__)))]:
        version_path = base / 'version.txt'
        try:
            if version_path.exists():
                return version_path.read_text().strip()
        except Exception:
            pass
    return "v0.0.0"


def _get_ssl_context():
    """Get an SSL context that works in PyInstaller builds."""
    import ssl
    try:
        import certifi
        return ssl.create_default_context(cafile=certifi.where())
    except (ImportError, Exception):
        pass
    if sys.platform == 'win32':
        try:
            ctx = ssl.SSLContext(ssl.PROTOCOL_TLS_CLIENT)
            ctx.load_default_certs(purpose=ssl.Purpose.SERVER_AUTH)
            return ctx
        except Exception:
            pass
    try:
        return ssl.create_default_context()
    except Exception:
        ctx = ssl.create_default_context()
        ctx.check_hostname = False
        ctx.verify_mode = ssl.CERT_NONE
        return ctx


def check_and_update():
    """Check GitHub for new release and auto-update if available."""
    try:
        import urllib.request
        import urllib.error
        import shutil

        local_version = get_local_version()
        logging.info("Checking for updates... current: %s", local_version)

        def _parse_version(v):
            """Parse 'v1.2.3' or '1.2.3' into tuple (1, 2, 3) for proper comparison."""
            try:
                return tuple(int(x) for x in v.lstrip('v').split('.'))
            except Exception:
                return (0, 0, 0)

        ssl_ctx = _get_ssl_context()
        url = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
        req = urllib.request.Request(url, headers={"Accept": "application/vnd.github.v3+json"})
        with urllib.request.urlopen(req, timeout=15, context=ssl_ctx) as resp:
            release = json.loads(resp.read().decode())

        remote_version = release.get("tag_name", "")
        if not remote_version or _parse_version(remote_version) <= _parse_version(local_version):
            logging.info("Already up to date (%s)", local_version)
            return False

        logging.info("New version available: %s → %s", local_version, remote_version)

        # Find the tracker exe in release assets
        tracker_asset = None
        controller_asset = None
        for asset in release.get("assets", []):
            if asset["name"] == "activity_tracker.exe":
                tracker_asset = asset
            elif asset["name"] == "activity_tracker_controller.exe":
                controller_asset = asset

        if not tracker_asset:
            logging.warning("No activity_tracker.exe in release %s", remote_version)
            return False

        # Backup current exe before replacing
        tracker_path = APP_DIR / 'activity_tracker.exe'
        backup_path = APP_DIR / 'activity_tracker.exe.backup'
        if tracker_path.exists():
            try:
                shutil.copy2(str(tracker_path), str(backup_path))
            except Exception:
                pass

        # Kill tracker before replacing
        kill_process('activity_tracker.exe')
        time.sleep(3)

        # Download new tracker
        temp_path = APP_DIR / 'activity_tracker.exe.tmp'
        dl_req = urllib.request.Request(tracker_asset["browser_download_url"])
        with urllib.request.urlopen(dl_req, context=ssl_ctx) as dl_resp:
            with open(str(temp_path), 'wb') as dl_file:
                dl_file.write(dl_resp.read())

        # Verify download
        if not temp_path.exists() or temp_path.stat().st_size < 1_000_000:
            logging.error("Downloaded file missing or too small, aborting")
            if temp_path.exists():
                temp_path.unlink()
            start_activity_tracker()
            return False

        # Replace exe — antivirus-safe: restore from backup if rename fails
        try:
            if tracker_path.exists():
                tracker_path.unlink()
            temp_path.rename(tracker_path)
        except Exception as e:
            logging.error("Failed to replace tracker exe: %s", e)
            if not tracker_path.exists() and backup_path.exists():
                shutil.copy2(str(backup_path), str(tracker_path))
                logging.info("Restored tracker from backup")
            if temp_path.exists():
                temp_path.unlink()
            start_activity_tracker()
            return False

        # Update version file
        version_path = APP_DIR / 'version.txt'
        version_path.write_text(remote_version)

        start_activity_tracker()
        logging.info("Updated tracker to %s", remote_version)

        # Download new controller for next restart
        if controller_asset:
            try:
                controller_new = APP_DIR / 'activity_tracker_controller.exe.new'
                ctrl_req = urllib.request.Request(controller_asset["browser_download_url"])
                with urllib.request.urlopen(ctrl_req, context=ssl_ctx) as ctrl_resp:
                    with open(str(controller_new), 'wb') as ctrl_file:
                        ctrl_file.write(ctrl_resp.read())
                logging.info("Downloaded new controller, will apply on next restart")
            except Exception as e:
                logging.warning("Failed to download new controller: %s", e)

        return True

    except Exception as e:
        logging.error("Update check error: %s", e, exc_info=True)
        return False


def check_crash_and_rollback():
    """If tracker keeps crashing after update, rollback to backup."""
    backup_path = APP_DIR / 'activity_tracker.exe.backup'
    tracker_path = APP_DIR / 'activity_tracker.exe'
    crash_counter_path = APP_DIR / 'crash_count.txt'

    if not backup_path.exists():
        return

    crash_count = 0
    try:
        if crash_counter_path.exists():
            content = crash_counter_path.read_text().strip().split('\n')
            recent = [float(t) for t in content if time.time() - float(t) < 300]
            crash_count = len(recent)
    except Exception:
        pass

    if crash_count >= MAX_CRASH_COUNT:
        logging.critical("Tracker crashed %d times in 5 minutes! Rolling back to backup...", crash_count)
        try:
            kill_process('activity_tracker.exe')
            time.sleep(2)
            import shutil
            shutil.copy2(str(backup_path), str(tracker_path))
            crash_counter_path.unlink(missing_ok=True)
            start_activity_tracker()
            logging.info("Rollback complete")
        except Exception as e:
            logging.error("Rollback failed: %s", e)


def record_crash():
    """Record a crash timestamp for rollback detection."""
    crash_counter_path = APP_DIR / 'crash_count.txt'
    try:
        with open(crash_counter_path, 'a') as f:
            f.write(f"{time.time()}\n")
    except Exception:
        pass


def init_supabase_client():
    if not create_client:
        return None
    try:
        return create_client(SUPABASE_URL, SUPABASE_KEY)
    except Exception as e:
        logging.error("Failed to initialize Supabase client: %s", e)
        return None


def upload_optimized_batches():
    """Upload optimized batch files to Supabase"""
    documents_path = APP_DIR / 'keytrk_data'
    if not documents_path.exists():
        return

    batch_files = sorted(documents_path.glob("optimized_batch_*.json"))
    if not batch_files:
        return

    # Cap at 500 files — delete oldest if over limit
    MAX_OFFLINE_BATCHES = 500
    if len(batch_files) > MAX_OFFLINE_BATCHES:
        excess = batch_files[:-MAX_OFFLINE_BATCHES]
        for old_file in excess:
            try:
                old_file.unlink()
            except Exception:
                pass
        batch_files = batch_files[-MAX_OFFLINE_BATCHES:]
        logging.warning("Deleted %d old offline batches (over %d limit)", len(excess), MAX_OFFLINE_BATCHES)

    supabase_client = init_supabase_client()
    if not supabase_client:
        return

    successful_uploads = []
    for batch_file in sorted(batch_files):
        try:
            success = upload_single_batch(supabase_client, batch_file)
            if success:
                successful_uploads.append(batch_file)
        except Exception as e:
            logging.error("Error uploading %s: %s", batch_file.name, e)

    for file_path in successful_uploads:
        try:
            file_path.unlink()
        except Exception:
            pass

    if successful_uploads:
        logging.info("Uploaded %d optimized batches", len(successful_uploads))


def upload_single_batch(supabase_client, file_path):
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            data = json.load(f)

        date_str = data.get('d')
        start_str = data.get('s')
        end_str = data.get('e')
        if date_str and start_str and 'T' not in str(start_str):
            start_str = f"{date_str}T{start_str}"
        if date_str and end_str and 'T' not in str(end_str):
            end_str = f"{date_str}T{end_str}"

        insert_data = {
            'batch_id': file_path.stem,
            'user_id': data.get('u'),
            'date_tracked': date_str,
            'batch_start_time': start_str,
            'batch_end_time': end_str,
            'total_time_seconds': data.get('tt', 0),
            'active_time_seconds': data.get('at', 0),
            'inactive_time_seconds': data.get('it', 0),
            'ip_address': data.get('ip'),
            'local_ips': data.get('li'),
            'batch_data': data
        }

        response = supabase_client.table("activity_summary").insert(insert_data).execute()
        return bool(response.data)
    except Exception as e:
        logging.error("Upload error for %s: %s", file_path.name, e)
        return False


def is_process_running(process_name):
    try:
        output = subprocess.check_output(
            ['tasklist', '/FI', f'IMAGENAME eq {process_name}'],
            stderr=subprocess.DEVNULL,
            creationflags=subprocess.CREATE_NO_WINDOW
        )
        return process_name.encode() in output
    except Exception:
        return False


def kill_process(process_name):
    """Force-kill a process by name"""
    try:
        subprocess.call(
            ['taskkill', '/F', '/IM', process_name],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            creationflags=subprocess.CREATE_NO_WINDOW
        )
        logging.info("Force-killed %s", process_name)
        return True
    except Exception as e:
        logging.error("Failed to kill %s: %s", process_name, e)
        return False


def start_activity_tracker():
    tracker_path = APP_DIR / 'activity_tracker.exe'
    if not tracker_path.exists():
        # Fallback: check exe's own directory
        own_dir = Path(os.path.dirname(sys.executable if getattr(sys, 'frozen', False) else os.path.abspath(__file__)))
        candidate = own_dir / 'activity_tracker.exe'
        if candidate.exists():
            tracker_path = candidate
        else:
            logging.error("activity_tracker.exe not found at %s or %s", APP_DIR, own_dir)
            return False
    try:
        subprocess.Popen(
            [str(tracker_path)],
            creationflags=subprocess.CREATE_NO_WINDOW | subprocess.DETACHED_PROCESS,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            stdin=subprocess.DEVNULL,
            cwd=str(tracker_path.parent),
        )
        logging.info("Started activity_tracker.exe from %s", tracker_path)
        return True
    except Exception as e:
        logging.error("Failed to start tracker: %s", e)
        return False


def check_last_alive():
    """Check if the tracker is actually producing data (not just running as zombie)."""
    alive_path = APP_DIR / 'last_alive.txt'
    if not alive_path.exists():
        return False

    try:
        content = alive_path.read_text().strip()
        last_alive = datetime.fromisoformat(content)
        age_seconds = (datetime.now() - last_alive).total_seconds()
        if age_seconds > STALE_THRESHOLD_SECONDS:
            logging.warning("last_alive.txt is %.0f seconds old (threshold: %d)", age_seconds, STALE_THRESHOLD_SECONDS)
            return False
        return True
    except Exception as e:
        logging.error("Error reading last_alive.txt: %s", e)
        return False


def _get_pc_name():
    """Get the PC identifier (matches tracker's user_id format)."""
    try:
        import importlib.util
        if getattr(sys, 'frozen', False):
            base_dir = os.path.dirname(sys.executable)
        else:
            base_dir = os.path.dirname(os.path.abspath(__file__))
        spec = importlib.util.spec_from_file_location("config", os.path.join(base_dir, 'config.py'))
        cfg = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(cfg)
        return cfg.get_user_id()
    except Exception:
        import platform
        return f"{os.environ.get('USERNAME', os.environ.get('USER', 'user'))}@{platform.node()}"


def _read_last_lines(file_path, n=100):
    """Read last N lines of a file."""
    try:
        if not file_path.exists():
            return ""
        with open(file_path, 'r', encoding='utf-8', errors='replace') as f:
            lines = f.readlines()
        return "".join(lines[-n:])
    except Exception:
        return ""


def upload_logs_to_supabase():
    """Upload tracker and controller logs to Supabase for remote monitoring."""
    supabase_client = init_supabase_client()
    if not supabase_client:
        return

    pc_name = _get_pc_name()
    tracker_running = is_process_running('activity_tracker.exe')
    current_version = get_local_version()

    # Read last_alive timestamp
    alive_path = APP_DIR / 'last_alive.txt'
    last_alive = None
    try:
        if alive_path.exists():
            last_alive = alive_path.read_text().strip()
    except Exception:
        pass

    # Update tracker_version on employee record
    try:
        supabase_client.table("employees").update({
            "tracker_version": current_version
        }).eq("pc_name", pc_name).execute()
    except Exception:
        pass

    # Upload logs
    for log_type, filename, lines in [('tracker', 'tracker.log', 100), ('controller', 'controller.log', 50)]:
        log_path = APP_DIR / filename
        content = _read_last_lines(log_path, lines)
        if content:
            try:
                supabase_client.table("tracker_logs").insert({
                    'pc_name': pc_name,
                    'log_type': log_type,
                    'log_content': content,
                    'last_alive': last_alive,
                    'tracker_running': tracker_running,
                }).execute()
                logging.info("Uploaded %s log to Supabase", log_type)
                # Truncate log after successful upload
                try:
                    if log_type != 'controller':
                        open(str(log_path), 'w').close()
                except Exception:
                    pass
            except Exception as e:
                logging.error("Failed to upload %s log: %s", log_type, e)
        # Safety cap: if log file exceeds 5MB, keep only last 1MB
        try:
            if log_path.exists() and log_path.stat().st_size > 5 * 1024 * 1024:
                with open(str(log_path), 'r', encoding='utf-8', errors='replace') as f:
                    f.seek(-1024 * 1024, 2)
                    f.readline()
                    tail = f.read()
                with open(str(log_path), 'w', encoding='utf-8') as f:
                    f.write(tail)
        except Exception:
            pass


def _ensure_scheduled_tasks():
    """Always register Windows Scheduled Tasks pointing to current exe."""
    if sys.platform != 'win32':
        return
    try:
        exe_path = sys.executable if getattr(sys, 'frozen', False) else os.path.abspath(__file__)
        subprocess.call(
            ['schtasks', '/Create', '/TN', 'ActivityX Controller',
             '/TR', f'"{exe_path}"', '/SC', 'MINUTE', '/MO', '5', '/F'],
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
            creationflags=subprocess.CREATE_NO_WINDOW
        )
        subprocess.call(
            ['schtasks', '/Create', '/TN', 'ActivityX Controller Startup',
             '/TR', f'"{exe_path}"', '/SC', 'ONLOGON', '/F'],
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
            creationflags=subprocess.CREATE_NO_WINDOW
        )
        logging.info("Registered Windows Scheduled Tasks → %s", exe_path)
    except Exception as e:
        logging.error("Failed to register scheduled tasks: %s", e)


def main():
    # Hide console window
    if sys.platform == 'win32':
        import ctypes
        ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 0)

    logging.info("Controller started")
    _ensure_scheduled_tasks()

    # Check if a new controller was downloaded — swap and let scheduled task restart
    if sys.platform == 'win32' and getattr(sys, 'frozen', False):
        new_controller = APP_DIR / 'activity_tracker_controller.exe.new'
        if new_controller.exists() and new_controller.stat().st_size > 1_000_000:
            try:
                current_exe = Path(sys.executable)
                old_exe = current_exe.with_suffix('.exe.old')
                if old_exe.exists():
                    old_exe.unlink()
                current_exe.rename(old_exe)
                new_controller.rename(current_exe)
                logging.info("Controller updated, exiting. Scheduled task will start new version.")
                sys.exit(0)
            except Exception as e:
                logging.error("Controller self-update failed: %s", e)

    # Report version immediately on startup
    try:
        _sb = init_supabase_client()
        if _sb:
            _sb.table("employees").update({
                "tracker_version": get_local_version()
            }).eq("pc_name", _get_pc_name()).execute()
    except Exception:
        pass

    # Wait 1 minute on first startup to let tracker initialize
    time.sleep(60)

    last_batch_upload = time.time()
    last_log_upload = time.time()
    last_update_check = 0  # Check on first loop
    last_tracker_start = time.time()  # Grace period for stale check
    last_loop_time = time.time()
    batch_upload_interval = 180  # 3 minutes
    log_upload_interval = 1800  # 30 minutes

    while True:
        try:
            current_time = time.time()

            # Detect wake from sleep — if loop gap > 2 minutes, system was sleeping
            if current_time - last_loop_time > 120:
                logging.info("Detected wake from sleep (gap: %.0fs), resetting grace period", current_time - last_loop_time)
                last_tracker_start = current_time
            last_loop_time = current_time

            process_running = is_process_running('activity_tracker.exe')
            # Only check staleness after tracker has had time to sync (10 min grace)
            tracker_healthy = check_last_alive() if (current_time - last_tracker_start > STALE_THRESHOLD_SECONDS) else True

            if not process_running:
                logging.warning("activity_tracker.exe not running, restarting...")
                record_crash()
                check_crash_and_rollback()
                start_activity_tracker()
                last_tracker_start = time.time()
            elif process_running and not tracker_healthy:
                logging.warning("Tracker process alive but stale. Force-killing...")
                kill_process('activity_tracker.exe')
                time.sleep(5)
                start_activity_tracker()
                last_tracker_start = time.time()

            # Auto-update check
            if current_time - last_update_check >= UPDATE_CHECK_INTERVAL:
                try:
                    updated = check_and_update()
                    if updated:
                        last_tracker_start = time.time()
                except Exception:
                    pass
                last_update_check = current_time

            # Upload queued batches every 3 minutes
            if current_time - last_batch_upload >= batch_upload_interval:
                try:
                    upload_optimized_batches()
                except Exception:
                    pass
                last_batch_upload = current_time

            # Upload logs to Supabase every 30 minutes
            if current_time - last_log_upload >= log_upload_interval:
                try:
                    upload_logs_to_supabase()
                except Exception:
                    pass
                last_log_upload = current_time

            time.sleep(30)

        except Exception as e:
            logging.error("Error in controller main loop: %s", e, exc_info=True)
            time.sleep(30)
            continue


if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        logging.critical("Controller fatal error: %s", e, exc_info=True)
        while True:
            time.sleep(60)
