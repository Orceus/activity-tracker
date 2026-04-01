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

    batch_files = list(documents_path.glob("optimized_batch_*.json"))
    if not batch_files:
        return

    supabase_client = init_supabase_client()
    if not supabase_client:
        return

    successful_uploads = []
    for batch_file in sorted(batch_files):
        try:
            success = upload_single_batch(supabase_client, batch_file)
            if success:
                successful_uploads.append(batch_file)
            else:
                break
        except Exception as e:
            logging.error("Error uploading %s: %s", batch_file.name, e)
            break

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

        insert_data = {
            'batch_id': file_path.stem,
            'user_id': data.get('u'),
            'date_tracked': data.get('d'),
            'batch_start_time': data.get('s'),
            'batch_end_time': data.get('e'),
            'total_time_seconds': data.get('tt', 0),
            'active_time_seconds': data.get('at', 0),
            'inactive_time_seconds': data.get('it', 0),
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
        # Fallback: check Documents path (legacy installs)
        legacy_path = Path.home() / 'Documents' / 'ActivityX' / 'activity_tracker.exe'
        if legacy_path.exists():
            tracker_path = legacy_path
        else:
            logging.error("activity_tracker.exe not found at %s or %s", APP_DIR, legacy_path)
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
    """Check if the tracker is actually producing data (not just running as zombie).
    Returns True if healthy, False if stale or missing."""
    alive_path = APP_DIR / 'last_alive.txt'
    if not alive_path.exists():
        # Also check legacy path
        legacy_alive = Path.home() / 'Documents' / 'ActivityX' / 'last_alive.txt'
        if legacy_alive.exists():
            alive_path = legacy_alive
        else:
            return False  # No alive file at all

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


def main():
    # Hide console window
    if sys.platform == 'win32':
        import ctypes
        ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 0)

    logging.info("Controller started")

    # Wait 1 minute on first startup to let tracker initialize
    time.sleep(60)

    last_batch_upload = time.time()
    batch_upload_interval = 180  # 3 minutes

    while True:
        try:
            current_time = time.time()

            process_running = is_process_running('activity_tracker.exe')
            tracker_healthy = check_last_alive()

            if not process_running:
                # Process is dead — restart it
                logging.warning("activity_tracker.exe not running, restarting...")
                start_activity_tracker()
            elif process_running and not tracker_healthy:
                # Process exists but isn't producing data (zombie) — kill and restart
                logging.warning("Tracker process alive but stale (no data in %ds). Force-killing...", STALE_THRESHOLD_SECONDS)
                kill_process('activity_tracker.exe')
                time.sleep(5)  # Wait for process to die
                start_activity_tracker()

            # Upload queued batches every 3 minutes
            if current_time - last_batch_upload >= batch_upload_interval:
                try:
                    upload_optimized_batches()
                except Exception:
                    pass
                last_batch_upload = current_time

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
        # Keep process alive even on fatal error
        while True:
            time.sleep(60)
