#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import sys
import os

# Force UTF-8 encoding for Windows
if sys.platform.startswith('win'):
    # Set environment variables for UTF-8
    os.environ['PYTHONIOENCODING'] = 'utf-8'
    # Force stdout and stderr to use UTF-8
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
    if hasattr(sys.stderr, 'reconfigure'):
        sys.stderr.reconfigure(encoding='utf-8')

"""
KeyTRK Activity Tracker - Advanced User Activity Monitoring
Enhanced version with optimized data collection and Supabase integration
"""

import time
import json
import psutil
import platform
import sqlite3
import shutil
import os
from datetime import datetime, timedelta
from collections import defaultdict
from pathlib import Path
import threading
import urllib.parse
import glob
import subprocess
import re
import tempfile

# Import user configuration
from config import get_user_id

# Supabase import
try:
    from supabase import create_client, Client
except ImportError:
    print("Warning: Supabase client not available. Install with: pip install supabase")
    create_client = None

# Cross-platform imports for window tracking and input monitoring
try:
    if platform.system() == "Windows":
        import win32gui
        import win32process
        from pynput import mouse, keyboard
    elif platform.system() == "Darwin":  # macOS
        from AppKit import NSWorkspace
        import Quartz
        from pynput import mouse, keyboard
    elif platform.system() == "Linux":
        import subprocess
        from pynput import mouse, keyboard
except ImportError as e:
    print(f"Warning: Some platform-specific modules not available: {e}")
    print("Install missing dependencies: pip install pynput")

class RealTimeURLDetector:
    """Real-time URL detection using browser APIs and system integration"""
    
    def __init__(self):
        self.system = platform.system()
        self.last_urls = {}  # Cache last known URLs for each browser
        self.url_cache_time = {}  # Time when each URL was cached
        
    def get_current_browser_url(self, app_name, window_title):
        """Get the current URL from the active browser tab"""
        try:
            if self.system == "Darwin":  # macOS
                return self._get_macos_browser_url(app_name, window_title)
            elif self.system == "Windows":
                return self._get_windows_browser_url(app_name, window_title)
            elif self.system == "Linux":
                return self._get_linux_browser_url(app_name, window_title)
        except Exception as e:
            # Silently fail and return None - this is expected for non-browser apps
            pass
        return None, None
    
    def _get_macos_browser_url(self, app_name, window_title):
        """Get URL from macOS browsers using AppleScript"""
        app_name_lower = app_name.lower()
        
        try:
            if 'chrome' in app_name_lower:
                script = '''
                tell application "Google Chrome"
                    if (count of windows) > 0 then
                        set currentTab to active tab of front window
                        return {URL of currentTab, title of currentTab}
                    end if
                end tell
                '''
                result = subprocess.run(['osascript', '-e', script], 
                                      capture_output=True, text=True, timeout=2)
                if result.returncode == 0 and result.stdout.strip():
                    output = result.stdout.strip().split(', ')
                    if len(output) >= 2:
                        url = output[0]
                        title = ', '.join(output[1:])  # In case title contains commas
                        return url, title
                        
            elif 'safari' in app_name_lower:
                script = '''
                tell application "Safari"
                    if (count of windows) > 0 then
                        set currentTab to current tab of front window
                        return {URL of currentTab, name of currentTab}
                    end if
                end tell
                '''
                result = subprocess.run(['osascript', '-e', script], 
                                      capture_output=True, text=True, timeout=2)
                if result.returncode == 0 and result.stdout.strip():
                    output = result.stdout.strip().split(', ')
                    if len(output) >= 2:
                        url = output[0]
                        title = ', '.join(output[1:])
                        return url, title
                        
            elif 'firefox' in app_name_lower:
                # Firefox is trickier - try to get from window title or use alternative method
                if window_title and ' - ' in window_title:
                    # Firefox usually shows "Page Title - Mozilla Firefox"
                    page_title = window_title.split(' - ')[0]
                    # Try to match this with recent browser history
                    return self._match_firefox_title_to_url(page_title)
                    
        except subprocess.TimeoutExpired:
            # AppleScript timed out - browser might be unresponsive
            pass
        except Exception as e:
            # Other errors - continue silently
            pass
            
        return None, None
    
    def _match_firefox_title_to_url(self, page_title):
        """Try to match Firefox page title to URL from recent history"""
        try:
            # This is a fallback for Firefox - use the browser history approach
            # but only look at very recent entries (last 2 minutes)
            reader = BrowserHistoryReader()
            recent_urls = reader.get_recent_urls(minutes_back=2)
            
            for url, data in recent_urls.items():
                url_title = data.get('title', '')
                if url_title and page_title.lower() in url_title.lower():
                    return url, url_title
        except:
            pass
        return None, None
    
    def _get_windows_browser_url(self, app_name, window_title):
        """Get URL from Windows browsers"""
        # TODO: Implement Windows-specific URL detection
        # For now, fall back to browser history method
        return None, None
    
    def _get_linux_browser_url(self, app_name, window_title):
        """Get URL from Linux browsers"""
        # TODO: Implement Linux-specific URL detection
        # For now, fall back to browser history method
        return None, None

class BrowserHistoryReader:
    """Read browser history to get actual URLs visited (fallback method)"""
    
    def __init__(self):
        self.browser_paths = self._get_browser_paths()
        self.last_check_time = datetime.now()
        self.recent_urls = {}
        self.url_cache = {}
    
    def _get_browser_paths(self):
        """Get browser history file paths for different operating systems"""
        system = platform.system()
        home = Path.home()
        
        paths = {}
        
        if system == "Windows":
            paths.update({
                'chrome': home / 'AppData/Local/Google/Chrome/User Data/Default/History',
                'edge': home / 'AppData/Local/Microsoft/Edge/User Data/Default/History',
                'firefox': home / 'AppData/Roaming/Mozilla/Firefox/Profiles'
            })
        elif system == "Darwin":  # macOS
            paths.update({
                'chrome': home / 'Library/Application Support/Google/Chrome/Default/History',
                'safari': home / 'Library/Safari/History.db',
                'firefox': home / 'Library/Application Support/Firefox/Profiles',
                'edge': home / 'Library/Application Support/Microsoft Edge/Default/History'
            })
        elif system == "Linux":
            paths.update({
                'chrome': home / '.config/google-chrome/Default/History',
                'firefox': home / '.mozilla/firefox',
                'chromium': home / '.config/chromium/Default/History'
            })
        
        return paths
    
    def get_recent_urls(self, minutes_back=2):
        """Get URLs visited in the last few minutes (much shorter window for real-time detection)"""
        current_time = datetime.now()
        cutoff_time = current_time - timedelta(minutes=minutes_back)
        
        all_urls = {}
        
        # Check each browser
        for browser, path in self.browser_paths.items():
            try:
                urls = self._read_browser_history(browser, path, cutoff_time)
                all_urls.update(urls)
            except Exception as e:
                # Silently continue if browser history can't be read
                continue
        
        # Cache the results
        self.url_cache.update(all_urls)
        return all_urls
    
    def get_most_recent_url(self):
        """Get the most recently visited URL with much shorter time windows"""
        # Try very recent times first for real-time detection
        for seconds in [30, 60, 120, 300]:  # 30s, 1m, 2m, 5m
            recent_urls = self.get_recent_urls(minutes_back=seconds/60)
            if recent_urls:
                most_recent = max(recent_urls.items(), key=lambda x: x[1]['visit_time'])
                return most_recent[0], most_recent[1]
        
        return None, None
    
    def get_url_for_browser(self, window_title):
        """Try to extract URL from browser window title or recent history"""
        # First try to get from very recent history
        recent_url, url_data = self.get_most_recent_url()
        if recent_url and url_data:
            # Check if the URL title matches the window title
            url_title = url_data.get('title', '')
            if window_title and url_title:
                # Fuzzy match between window title and URL title
                if (url_title.lower() in window_title.lower() or 
                    window_title.lower() in url_title.lower() or
                    any(word in window_title.lower() for word in url_title.lower().split() if len(word) > 3)):
                    return recent_url, url_data
        
        return None, None
    
    def _read_browser_history(self, browser, db_path, cutoff_time):
        """Read history from a specific browser"""
        urls = {}
        
        if browser == 'firefox':
            return self._read_firefox_history(db_path, cutoff_time)
        elif browser == 'safari':
            return self._read_safari_history(db_path, cutoff_time)
        else:
            return self._read_chromium_history(db_path, cutoff_time)
    
    def _read_chromium_history(self, db_path, cutoff_time):
        """Read Chrome/Edge/Chromium history"""
        urls = {}
        
        if not os.path.exists(db_path):
            return urls
        
        # Copy the database to avoid locking issues
        temp_db = str(db_path) + "_temp"
        try:
            shutil.copy2(db_path, temp_db)
            
            conn = sqlite3.connect(temp_db)
            cursor = conn.cursor()
            
            # Convert cutoff time to Chrome's time format (microseconds since 1601)
            epoch_start = datetime(1601, 1, 1)
            chrome_time = int((cutoff_time - epoch_start).total_seconds() * 1000000)
            
            # Also get from visits table for more recent activity
            query = """
            SELECT u.url, u.title, u.visit_count, v.visit_time
            FROM urls u
            JOIN visits v ON u.id = v.url
            WHERE v.visit_time > ?
            ORDER BY v.visit_time DESC
            LIMIT 50
            """
            
            cursor.execute(query, (chrome_time,))
            results = cursor.fetchall()
            
            # If no visits found, fall back to urls table
            if not results:
                query = """
                SELECT url, title, visit_count, last_visit_time
                FROM urls
                WHERE last_visit_time > ?
                ORDER BY last_visit_time DESC
                LIMIT 20
                """
                
                cursor.execute(query, (chrome_time,))
                results = cursor.fetchall()
            
            for url, title, visits, timestamp in results:
                # Convert Chrome timestamp back to datetime
                visit_time = epoch_start + timedelta(microseconds=timestamp)
                urls[url] = {
                    'title': title or url,
                    'visit_time': visit_time,
                    'visit_count': visits or 0
                }
            
            conn.close()
            os.remove(temp_db)
            
        except Exception:
            # Clean up temp file if it exists
            if os.path.exists(temp_db):
                try:
                    os.remove(temp_db)
                except:
                    pass
        
        return urls
    
    def _read_firefox_history(self, profiles_path, cutoff_time):
        """Read Firefox history"""
        urls = {}
        
        if not os.path.exists(profiles_path):
            return urls
        
        # Find Firefox profile directories
        try:
            for profile_dir in os.listdir(profiles_path):
                if profile_dir.endswith('.default') or profile_dir.endswith('.default-release'):
                    places_db = os.path.join(profiles_path, profile_dir, 'places.sqlite')
                    if os.path.exists(places_db):
                        urls.update(self._read_places_db(places_db, cutoff_time))
        except:
            pass
        
        return urls
    
    def _read_places_db(self, db_path, cutoff_time):
        """Read Firefox places.sqlite database"""
        urls = {}
        temp_db = str(db_path) + "_temp"
        
        try:
            shutil.copy2(db_path, temp_db)
            conn = sqlite3.connect(temp_db)
            cursor = conn.cursor()
            
            # Firefox stores time as microseconds since Unix epoch
            firefox_time = int(cutoff_time.timestamp() * 1000000)
            
            query = """
            SELECT h.url, h.title, h.visit_count, h.last_visit_date
            FROM moz_places h
            WHERE h.last_visit_date > ?
            ORDER BY h.last_visit_date DESC
            LIMIT 20
            """
            
            cursor.execute(query, (firefox_time,))
            results = cursor.fetchall()
            
            for url, title, visits, timestamp in results:
                if timestamp:
                    visit_time = datetime.fromtimestamp(timestamp / 1000000)
                    urls[url] = {
                        'title': title or url,
                        'visit_time': visit_time,
                        'visit_count': visits or 0
                    }
            
            conn.close()
            os.remove(temp_db)
            
        except Exception:
            if os.path.exists(temp_db):
                try:
                    os.remove(temp_db)
                except:
                    pass
        
        return urls
    
    def _read_safari_history(self, db_path, cutoff_time):
        """Read Safari history"""
        urls = {}
        
        if not os.path.exists(db_path):
            return urls
        
        temp_db = str(db_path) + "_temp"
        
        try:
            shutil.copy2(db_path, temp_db)
            conn = sqlite3.connect(temp_db)
            cursor = conn.cursor()
            
            # Safari stores time as seconds since 2001-01-01
            safari_epoch = datetime(2001, 1, 1)
            safari_time = (cutoff_time - safari_epoch).total_seconds()
            
            query = """
            SELECT url, title, visit_count, visit_time
            FROM history_visits hv
            JOIN history_items hi ON hv.history_item = hi.id
            WHERE visit_time > ?
            ORDER BY visit_time DESC
            LIMIT 20
            """
            
            cursor.execute(query, (safari_time,))
            results = cursor.fetchall()
            
            for url, title, visits, timestamp in results:
                visit_time = safari_epoch + timedelta(seconds=timestamp)
                urls[url] = {
                    'title': title or url,
                    'visit_time': visit_time,
                    'visit_count': visits or 0
                }
            
            conn.close()
            os.remove(temp_db)
            
        except Exception:
            if os.path.exists(temp_db):
                try:
                    os.remove(temp_db)
                except:
                    pass
        
        return urls

class InputTracker:
    """Track mouse clicks and keyboard activity"""
    
    def __init__(self):
        self.click_count = 0
        self.key_count = 0
        self.last_activity = time.time()
        self.is_monitoring = False
        
    def start_monitoring(self):
        """Start monitoring mouse and keyboard input"""
        self.is_monitoring = True
        try:
            # Start mouse listener
            self.mouse_listener = mouse.Listener(
                on_click=self._on_click,
                on_scroll=self._on_scroll,
                on_move=self._on_move
            )
            self.mouse_listener.start()
            
            # Start keyboard listener
            self.keyboard_listener = keyboard.Listener(
                on_press=self._on_key_press
            )
            self.keyboard_listener.start()
            
        except Exception as e:
            print(f"Warning: Could not start input monitoring: {e}")
    
    def stop_monitoring(self):
        """Stop monitoring input"""
        self.is_monitoring = False
        try:
            if hasattr(self, 'mouse_listener'):
                self.mouse_listener.stop()
            if hasattr(self, 'keyboard_listener'):
                self.keyboard_listener.stop()
        except:
            pass
    
    def _on_click(self, x, y, button, pressed):
        """Handle mouse click events"""
        if pressed:  # Only count when button is pressed, not released
            self.click_count += 1
            self.last_activity = time.time()
    
    def _on_scroll(self, x, y, dx, dy):
        """Handle mouse scroll events"""
        self.last_activity = time.time()
    
    def _on_move(self, x, y):
        """Handle mouse movement events"""
        self.last_activity = time.time()
    
    def _on_key_press(self, key):
        """Handle keyboard press events"""
        self.key_count += 1
        self.last_activity = time.time()
    
    def get_activity_stats(self):
        """Get current activity statistics"""
        return {
            'clicks': self.click_count,
            'keystrokes': self.key_count,
            'last_activity': self.last_activity
        }
    
    def reset_counters(self):
        """Reset activity counters"""
        self.click_count = 0
        self.key_count = 0

class DataSyncer:
    """Handles continuous data saving and Supabase synchronization"""

    def __init__(self, supabase_url=None, supabase_key=None):
        self.supabase_url = supabase_url
        self.supabase_key = supabase_key
        self.supabase_client = None
        self.data_directory = Path("keytrk_data")
        self.data_directory.mkdir(exist_ok=True)
        self.sync_interval = 120  # 2 minutes in seconds
        self.is_syncing = False
        self.sync_thread = None
        self.last_save_time = time.time()
        self.pending_data = {}  # Data accumulated since last save
        
        # Get user ID for data identification
        self.user_id = get_user_id()
        
        # Initialize Supabase client if credentials provided
        if self.supabase_url and self.supabase_key and create_client:
            try:
                self.supabase_client = create_client(self.supabase_url, self.supabase_key)
                print("SUCCESS: Supabase client initialized successfully")
            except Exception as e:
                print(f"WARNING: Failed to initialize Supabase client: {e}")
                self.supabase_client = None
        else:
            print("WARNING: Supabase credentials not provided or client not available")
    
    def start_syncing(self):
        """Start the continuous data syncing process"""
        self.is_syncing = True
        self.sync_thread = threading.Thread(target=self._sync_loop, daemon=True)
        self.sync_thread.start()
        print("INFO: Data syncing started - saving every 2 minutes")
    
    def stop_syncing(self):
        """Stop the syncing process and save any remaining data"""
        self.is_syncing = False
        # Save any remaining pending data
        if self.pending_data:
            self._save_current_data()
        print("INFO: Data syncing stopped")
    
    def add_data(self, activity_key, activity_data):
        """Add new activity data to pending collection"""
        if activity_key not in self.pending_data:
            self.pending_data[activity_key] = {
                'app_name': activity_data.get('app_name'),
                'window_title': activity_data.get('window_title'),
                'current_url': activity_data.get('current_url', ''),
                'all_urls': list(activity_data.get('urls', set())),
                'category': activity_data.get('category'),
                'total_time_seconds': 0,
                'total_clicks': 0,
                'total_keystrokes': 0,
                'sessions': [],
                'last_active': activity_data.get('last_active')
            }
        
        # Update existing data
        self.pending_data[activity_key]['total_time_seconds'] += activity_data.get('total_time', 0)
        self.pending_data[activity_key]['total_clicks'] += activity_data.get('clicks', 0)
        self.pending_data[activity_key]['total_keystrokes'] += activity_data.get('keystrokes', 0)
        self.pending_data[activity_key]['sessions'].extend(activity_data.get('sessions', []))
        self.pending_data[activity_key]['last_active'] = activity_data.get('last_active')
        
        # Update URLs if new ones exist
        existing_urls = set(self.pending_data[activity_key]['all_urls'])
        new_urls = set(activity_data.get('urls', set()))
        all_urls = existing_urls.union(new_urls)
        self.pending_data[activity_key]['all_urls'] = list(all_urls)
    
    def _sync_loop(self):
        """Main synchronization loop"""
        while self.is_syncing:
            try:
                current_time = time.time()
                
                # Check if it's time to save data (every 2 minutes)
                if current_time - self.last_save_time >= self.sync_interval:
                    if self.pending_data:
                        self._save_current_data()
                    
                    # Try to sync all pending files
                    self._sync_all_pending_files()
                    
                    self.last_save_time = current_time
                
                time.sleep(10)  # Check every 10 seconds
                
            except Exception as e:
                print(f"ERROR: Error in sync loop: {e}")
                time.sleep(30)  # Wait longer on error
    
    def _save_current_data(self):
        """Save current pending data to a new JSON file"""
        if not self.pending_data:
            return
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = self.data_directory / f"keytrk_batch_{timestamp}.json"
        
        # Calculate totals
        total_tracked_time = sum(data['total_time_seconds'] for data in self.pending_data.values())
        total_clicks = sum(data['total_clicks'] for data in self.pending_data.values())
        total_keystrokes = sum(data['total_keystrokes'] for data in self.pending_data.values())
        
        json_data = {
            'batch_info': {
                'batch_id': timestamp,
                'user_id': self.user_id,
                'created_time': datetime.now().isoformat(),
                'system': platform.system(),
                'total_activities': len(self.pending_data),
                'total_tracked_time_seconds': total_tracked_time,
                'total_clicks': total_clicks,
                'total_keystrokes': total_keystrokes
            },
            'activities': self.pending_data.copy()
        }
        
        try:
            with open(filename, 'w', encoding='utf-8') as f:
                json.dump(json_data, f, indent=2, ensure_ascii=False)
            
            print(f"Data batch saved to: {filename.name}")
            print(f"Batch stats: {len(self.pending_data)} activities, {total_tracked_time:.1f}s tracked")
            
            # Clear pending data after saving
            self.pending_data = {}
            
        except Exception as e:
            print(f"Failed to save data batch: {e}")
    
    def _sync_all_pending_files(self):
        """Attempt to sync all unsent JSON files to Supabase"""
        if not self.supabase_client:
            return
        
        # Get all JSON files in the data directory
        json_files = list(self.data_directory.glob("keytrk_batch_*.json"))
        
        if not json_files:
            return
        
        successful_uploads = []
        
        for json_file in sorted(json_files):  # Process in chronological order
            try:
                success = self._upload_file_to_supabase(json_file)
                if success:
                    successful_uploads.append(json_file)
                else:
                    # If one file fails, stop processing to maintain order
                    break
                    
            except Exception as e:
                print(f"Error uploading {json_file.name}: {e}")
                break
        
        # Delete successfully uploaded files
        for file_path in successful_uploads:
            try:
                file_path.unlink()
                print(f"Deleted uploaded file: {file_path.name}")
            except Exception as e:
                print(f"Failed to delete {file_path.name}: {e}")
        
        if successful_uploads:
            print(f"SUCCESS: Successfully synced {len(successful_uploads)} data batches")
        elif json_files:
            print(f"{len(json_files)} data batches pending upload (will retry)")
    
    def _upload_file_to_supabase(self, file_path):
        """Upload a single JSON file to Supabase"""
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            # Structure data correctly for the table columns
            insert_data = {
                'batch_id': data['batch_info']['batch_id'],
                'user_id': data['batch_info']['user_id'],
                'created_time': data['batch_info']['created_time'],
                'batch_info': data['batch_info'],
                'activities': data['activities']
            }
            
            # Upload to Supabase table
            result = self.supabase_client.table('activity_tracking').insert(insert_data).execute()
            
            if result.data:
                print(f"Uploaded: {file_path.name}")
                return True
            else:
                print(f"Upload failed for: {file_path.name}")
                return False
                
        except Exception as e:
            print(f"Upload error for {file_path.name}: {e}")
            return False
    
    def get_pending_files_count(self):
        """Get count of files waiting to be uploaded"""
        json_files = list(self.data_directory.glob("keytrk_batch_*.json"))
        return len(json_files)

class OptimizedDataSyncer:
    """Optimized data syncer for manager reporting - compact JSON with shortened field names"""

    def __init__(self, supabase_url=None, supabase_key=None):
        self.supabase_url = supabase_url
        self.supabase_key = supabase_key
        self.supabase_client = None
        self.data_directory = Path("keytrk_data")
        self.data_directory.mkdir(exist_ok=True)
        self.sync_interval = 300  # 5 minutes in seconds
        self.is_syncing = False
        self.sync_thread = None
        self.last_sync_time = time.time()
        
        # Data storage for current batch
        self.batch_start_time = datetime.now()
        self.activity_timeline = []  # Track activities chronologically
        self.inactive_periods = []  # [{'s': start_time, 'du': duration}]
        self.total_inactive_time = 0.0
        
        # NEW: Track current inactive state for continuous batching
        self.current_inactive_state = {
            'is_inactive': False,
            'inactive_start_time': None,
            'inactive_start_timestamp': None
        }
        
        # Get user ID for identification
        self.user_id = self._get_user_id()
        
        # Initialize Supabase client
        if self.supabase_url and self.supabase_key and create_client:
            try:
                self.supabase_client = create_client(self.supabase_url, self.supabase_key)
                print("SUCCESS: Optimized Supabase client initialized")
            except Exception as e:
                print(f"Failed to initialize Supabase client: {e}")
                self.supabase_client = None
        else:
            print("Supabase credentials not provided")
    
    def _get_user_id(self):
        """Get user ID from config or system"""
        try:
            from config import get_user_id
            return get_user_id()
        except:
            return f"user_{platform.node()}"
    
    def _clean_app_name(self, app_name):
        """Remove file extensions from app names"""
        if not app_name:
            return "Unknown"
        # Remove common extensions
        extensions = ['.exe', '.dmg', '.app']
        clean_name = app_name
        for ext in extensions:
            if clean_name.lower().endswith(ext.lower()):
                clean_name = clean_name[:-len(ext)]
                break
        return clean_name
    
    def _extract_domain(self, url):
        """Extract domain from URL or use the text as-is if not a valid URL"""
        try:
            parsed = urllib.parse.urlparse(url)
            if parsed.netloc:
                return parsed.netloc.lower()
            else:
                # If no domain found, use the original text as the "domain"
                # This handles window titles that aren't URLs
                return url.strip() if url else None
        except:
            return url.strip() if url else None
    

    
    def start_syncing(self):
        """Start the continuous sync process"""
        self.is_syncing = True
        self.sync_thread = threading.Thread(target=self._sync_loop, daemon=True)  
        self.sync_thread.start()
        print("INFO: Optimized data syncing started - every 5 minutes")
    
    def stop_syncing(self):
        """Stop syncing and send final data"""
        self.is_syncing = False
        if self.activity_timeline or self.inactive_periods:
            self._sync_to_supabase()
        print("INFO: Optimized data syncing stopped")
    
    def add_data(self, activity_key, activity_data):
        """Add activity data with proper time tracking"""
        app_name = self._clean_app_name(activity_data.get('app_name', 'Unknown'))
        time_seconds = activity_data.get('total_time', 0)
        current_url = activity_data.get('current_url', '')
        window_title = activity_data.get('window_title', '')
        
        # FIXED: Only skip truly insignificant micro-activities (< 0.1 seconds)
        # This preserves legitimate work sessions while filtering obvious noise
        if time_seconds < 0.1:
            return
        
        # Round to 2 decimal places
        time_seconds = round(time_seconds, 2)
        
        # Add to timeline with timestamp
        current_time = datetime.now()
        
        activity_record = {
            'timestamp': current_time,
            'app_name': app_name,
            'duration': time_seconds,
            'current_url': current_url,
            'window_title': window_title
        }
        
        self.activity_timeline.append(activity_record)
        
        print(f"INFO: Added activity: {app_name} - {time_seconds:.2f}s at {current_time.strftime('%H:%M:%S')}")
    

    def add_inactive_period(self, start_time, duration_seconds):
        """Add an inactive period"""
        # FIXED: Use consistent threshold with other filtering (0.1s instead of 1.0s)
        if duration_seconds < 0.1:
            return
        
        duration_rounded = round(duration_seconds, 2)
        start_time_str = start_time.strftime("%H:%M:%S")
        
        # FIXED: Add special handling for very long periods
        if duration_rounded > 3600:  # Over 1 hour
            print(f"INFO: Very long inactive period detected: {duration_rounded/3600:.1f} hours")
            print(f"   This likely indicates sleep mode or extended time away from computer")
        elif duration_rounded > 600:  # Over 10 minutes
            print(f"INFO: Long inactive period detected: {duration_rounded/60:.1f} minutes")
        
        inactive_period = {
            's': start_time_str,
            'du': duration_rounded
        }
        
        # FIXED: Add timestamp metadata for better debugging
        if duration_rounded > 600:  # Only for periods > 10 minutes
            inactive_period['timestamp'] = start_time.isoformat()
            inactive_period['end_time'] = (start_time + timedelta(seconds=duration_rounded)).isoformat()
        
        self.inactive_periods.append(inactive_period)
        
        self.total_inactive_time += duration_rounded
        self.total_inactive_time = round(self.total_inactive_time, 2)
        
        print(f"INFO: Added inactive period: {duration_rounded}s (total inactive: {self.total_inactive_time}s)")
    
    def _sync_loop(self):
        """Main sync loop - ALWAYS send a batch every 5 minutes (active or inactive)"""
        while self.is_syncing:
            try:
                current_time = time.time()

                if current_time - self.last_sync_time >= self.sync_interval:
                    # FORCE batch creation and sync every 5 minutes regardless of state
                    self._force_sync_batch()

                    # Try to sync all pending files after the main sync
                    self._sync_all_pending_files()

                    self.last_sync_time = current_time

                time.sleep(30)  # Check every 30 seconds

            except Exception as e:
                print(f"ERROR: Error in optimized sync loop: {e}")
                time.sleep(60)

    def _force_sync_batch(self):
        """Force creation and sync of a batch every 5 minutes, regardless of user state"""
        # 5-minute sync triggered
        
        # If user is inactive, ensure we have an inactive batch for this 5-minute window
        if self.current_inactive_state['is_inactive']:
            self._ensure_inactive_batch_for_window()
        
        # Always sync whatever data we have (active activities or inactive periods)
        self._sync_to_supabase_forced()

    def _ensure_inactive_batch_for_window(self):
        """Ensure we have inactive data for the current 5-minute window"""
        current_time = time.time()
        batch_start_timestamp = self.batch_start_time.timestamp()
        batch_end_timestamp = current_time
        batch_duration = batch_end_timestamp - batch_start_timestamp
        batch_duration = min(batch_duration, self.sync_interval)  # Cap at 5 minutes
        
        inactive_start_timestamp = self.current_inactive_state['inactive_start_timestamp']
        
        # Always create a 5-minute inactive period for the current window
        inactive_start_time_str = self.batch_start_time.strftime("%H:%M:%S")
        
        # Create a 5-minute inactive period
        self.inactive_periods = [{
            's': inactive_start_time_str,
            'du': self.sync_interval  # Always use 5 minutes (300 seconds)
        }]
        self.total_inactive_time = self.sync_interval
        
        # Created inactive batch

    def _sync_to_supabase_forced(self):
        """Force sync to Supabase, always creating a batch even if minimal data"""
        # Attempting FORCED sync
        
        # Force creation of data even if empty
        optimized_data = self._prepare_optimized_data_forced()
        
        if not optimized_data:
            # Could not create any batch data - this should not happen with forced sync
            self._reset_batch()
            return
        
        batch_id = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        # Prepare data for database
        db_data = {
            'batch_id': batch_id,
            'user_id': self.user_id,
            'date_tracked': optimized_data['d'],
            'batch_start_time': optimized_data['s'],
            'batch_end_time': optimized_data['e'],
            'total_time_seconds': optimized_data['tt'],
            'active_time_seconds': optimized_data['at'],
            'inactive_time_seconds': optimized_data['it'],
            'batch_data': optimized_data
        }
        
        # Try to upload to Supabase
        if self.supabase_client:
            try:
                result = self.supabase_client.table('activity_summary').insert(db_data).execute()
                
                if result.data:
                    # Show sync status
                    apps_count = len(optimized_data.get('ap', {}))
                    urls_count = len(optimized_data.get('ur', {}))
                    if apps_count == 0 and urls_count == 0:
                        print(f"SUCCESS: Synced INACTIVE batch: {optimized_data['it']:.1f}s inactive - sent immediately")
                    else:
                        print(f"SUCCESS: Synced ACTIVE batch: {optimized_data['at']:.1f}s active, {apps_count} apps, {urls_count} URLs")
                    self._reset_batch()
                    return True
                else:
                    print("Failed to sync to Supabase")
            except Exception as e:
                print(f"Supabase sync error: {e}")
        
        # Save locally if Supabase fails
        self._save_locally(optimized_data, batch_id)
        self._reset_batch()

    def _prepare_optimized_data_forced(self):
        """Force preparation of data even if no activities (for inactive batches)"""
        batch_end_time = datetime.now()
        
        # If we have no activities AND no inactive periods, create a minimal active batch
        if not self.activity_timeline and not self.inactive_periods:
            print("No activity data - creating minimal active batch for this 5-minute window")
            # Create a minimal active batch (user was present but no specific activities tracked)
            self.total_inactive_time = 0.0
            
            # Create minimal batch data
            batch_start_time = self.batch_start_time
            total_batch_time = 300.00  # 5 minutes
            
            start_date = batch_start_time.date()
            end_date = batch_end_time.date()
            spans_midnight = start_date != end_date
            
            # Create minimal optimized data structure
            optimized_data = {
                'u': self.user_id,
                'd': start_date.isoformat(),
                's': batch_start_time.strftime("%H:%M:%S"),
                'e': batch_end_time.strftime("%H:%M:%S"),
                'tt': total_batch_time,
                'at': total_batch_time,  # Assume all time was active (minimal tracking)
                'it': 0.0,  # No inactive time
                'spans_midnight': spans_midnight
            }
            
            if spans_midnight:
                optimized_data['ed'] = end_date.isoformat()
            
            # Add timezone info
            optimized_data['tz'] = batch_start_time.strftime("%z") or "Local"
            
            return optimized_data
        
        # Use the existing _prepare_optimized_data logic for normal cases
        return self._prepare_optimized_data()
    
    def _handle_inactive_batch_creation(self):
        """Create inactive-only batches during long inactive periods"""
        if not self.current_inactive_state['is_inactive']:
            return
            
        current_time = time.time()
        batch_start_timestamp = self.batch_start_time.timestamp()
        batch_window_seconds = self.sync_interval  # 300 seconds
        
        # Calculate how much of this batch window was inactive
        inactive_start = self.current_inactive_state['inactive_start_timestamp']
        
        # If user has been inactive for the entire batch window, create inactive batch
        if inactive_start <= batch_start_timestamp:
            # User was inactive for the entire 5-minute batch window
            batch_start_time_str = self.batch_start_time.strftime("%H:%M:%S")
            
            # Create a 5-minute inactive period for this batch
            inactive_period = {
                's': batch_start_time_str,
                'du': batch_window_seconds  # Full 5 minutes inactive
            }
            
            # Clear any existing data and add the inactive period
            self.activity_timeline = []  # No active activities
            self.inactive_periods = [inactive_period]  # One 5-minute inactive period
            self.total_inactive_time = batch_window_seconds
            
            print(f"Creating inactive-only batch: {batch_window_seconds}s inactive from {batch_start_time_str}")
            
        else:
            # User became inactive partway through this batch window
            inactive_duration_in_batch = current_time - inactive_start
            inactive_duration_in_batch = min(inactive_duration_in_batch, batch_window_seconds)
            
            if inactive_duration_in_batch > 0:
                inactive_start_time = datetime.fromtimestamp(inactive_start)
                inactive_start_time_str = inactive_start_time.strftime("%H:%M:%S")
                
                inactive_period = {
                    's': inactive_start_time_str,
                    'du': round(inactive_duration_in_batch, 2)
                }
                
                self.inactive_periods.append(inactive_period)
                self.total_inactive_time += inactive_duration_in_batch
                
                print(f"Adding partial inactive period to batch: {inactive_duration_in_batch:.1f}s from {inactive_start_time_str}")
    
    def _calculate_time_distribution(self, valid_inactive_time=None):
        """Calculate proper time distribution ensuring total doesn't exceed window"""
        if not self.activity_timeline:
            # No active activities - this is a purely inactive batch
            print("No active activities in this batch (purely inactive)")
            return {}, {}
        
        # Fixed 5-minute window
        total_window_time = 300.0  # 5 minutes
        # Use provided valid_inactive_time or fall back to total (for backward compatibility)
        inactive_time = valid_inactive_time if valid_inactive_time is not None else self.total_inactive_time
        available_active_time = total_window_time - inactive_time
        available_active_time = max(0.0, available_active_time)
        
        # Sum all raw activity times
        total_raw_time = sum(activity['duration'] for activity in self.activity_timeline)
        
        print(f"Time calculation: Window={total_window_time}s, Inactive={inactive_time}s, Available={available_active_time}s, Raw total={total_raw_time}s")
        
        # If raw time exceeds available time, scale proportionally
        scaling_factor = 1.0
        if total_raw_time > available_active_time and available_active_time > 0:
            scaling_factor = available_active_time / total_raw_time
            print(f"Scaling activities by factor: {scaling_factor:.3f}")
        
        # Calculate app times and URL times
        app_times = {}
        url_times = {}
        
        for activity in self.activity_timeline:
            app_name = activity['app_name']
            scaled_time = round(activity['duration'] * scaling_factor, 2)
            
            # FIXED: Remove arbitrary filtering that was losing legitimate work time
            # Only skip activities that become negligible after scaling (< 0.1s)
            if scaled_time < 0.1:
                continue
            
            # Add to app times
            if app_name not in app_times:
                app_times[app_name] = 0.0
            app_times[app_name] += scaled_time
            app_times[app_name] = round(app_times[app_name], 2)
            
            # Add URL data if available
            current_url = activity['current_url']
            if current_url:
                domain = self._extract_domain(current_url)
                if domain:
                    if domain not in url_times:
                        url_times[domain] = {'t': 0.0}
                    
                    url_times[domain]['t'] += scaled_time
                    url_times[domain]['t'] = round(url_times[domain]['t'], 2)
        
        # Final verification
        final_app_total = sum(app_times.values())
        final_url_total = sum(url['t'] for url in url_times.values())
        
        print(f"SUCCESS: Final totals: Apps={final_app_total}s, URLs={final_url_total}s, Available={available_active_time}s")
        
        return app_times, url_times
    
    def _prepare_optimized_data(self):
        """Prepare compact JSON data with correct time accounting and date boundary handling"""
        if not self.activity_timeline and not self.inactive_periods:
            print("No data to prepare")
            return None
        
        # NEW: Handle purely inactive batches (common during continuous inactive tracking)
        is_purely_inactive_batch = len(self.activity_timeline) == 0 and len(self.inactive_periods) > 0
        if is_purely_inactive_batch:
            print(f"Preparing purely inactive batch: {len(self.inactive_periods)} inactive periods")
        
        batch_end_time = datetime.now()
        
        # Fixed 5-minute window times
        total_batch_time = 300.00  # 5 minutes in seconds
        
        # FIXED: Handle long inactive periods that span multiple batches (e.g., sleep mode)
        batch_start_timestamp = self.batch_start_time.timestamp()
        batch_end_timestamp = batch_end_time.timestamp()
        
        valid_inactive_periods = []
        valid_inactive_time = 0.0
        long_inactive_period_detected = False
        
        for period in self.inactive_periods:
            # Parse the period start time (format: "HH:MM:SS")
            period_start_str = period['s']
            period_duration = period['du']
            
            # FIXED: Handle long inactive periods (> 10 minutes) specially
            if period_duration > 600:  # 10+ minutes = likely sleep/away from computer
                # Long inactive period detected
                
                # For long periods, cap the inactive time at the batch window size
                # This prevents mathematical impossibilities while preserving the data
                capped_duration = min(period_duration, total_batch_time)
                
                adjusted_period = {
                    's': period_start_str,
                    'du': round(capped_duration, 2),
                    'original_duration': round(period_duration, 2),  # Preserve original duration
                    'long_period': True  # Mark as long period for analysis
                }
                valid_inactive_periods.append(adjusted_period)
                valid_inactive_time += capped_duration
                long_inactive_period_detected = True
                
                # Long inactive period capped for batch consistency
                continue
            
            # Regular processing for shorter inactive periods
            try:
                # Combine with current batch date
                period_start_time = datetime.combine(
                    self.batch_start_time.date(),
                    datetime.strptime(period_start_str, "%H:%M:%S").time()
                )
                period_start_timestamp = period_start_time.timestamp()
                period_end_timestamp = period_start_timestamp + period_duration
                
                # Check if period overlaps with current batch window
                if (period_start_timestamp < batch_end_timestamp and 
                    period_end_timestamp > batch_start_timestamp):
                    
                    # Calculate overlap duration
                    overlap_start = max(period_start_timestamp, batch_start_timestamp)
                    overlap_end = min(period_end_timestamp, batch_end_timestamp)
                    overlap_duration = max(0.0, overlap_end - overlap_start)
                    
                    if overlap_duration > 0:
                        # Create adjusted period for the overlap
                        overlap_start_time = datetime.fromtimestamp(overlap_start)
                        adjusted_period = {
                            's': overlap_start_time.strftime("%H:%M:%S"),
                            'du': round(overlap_duration, 2)
                        }
                        valid_inactive_periods.append(adjusted_period)
                        valid_inactive_time += overlap_duration
                        
                        # Inactive period overlap processed
                
            except Exception as e:
                # Error processing inactive period
                continue
        
        # Use only the valid inactive time within the batch window
        batch_inactive_time = round(valid_inactive_time, 2)
        active_time = total_batch_time - batch_inactive_time
        active_time = max(0.0, round(active_time, 2))  # Ensure not negative
        
        # Time validation completed
        
        # FIXED: Special handling when long inactive periods are detected
        if long_inactive_period_detected:
            # Long inactive period detected - adjusting active time calculation
            # When we have a long inactive period, the "active time" should be minimal
            # since most of the period was spent away from computer
            active_time = max(0.0, total_batch_time - batch_inactive_time)
            # Adjusted active time for long inactive period
        
        # Validate time constraints
        if batch_inactive_time > total_batch_time:
            # Inactive time exceeds batch window - capping to batch window size
            batch_inactive_time = total_batch_time
            active_time = 0.0
        
        # Calculate properly distributed app and URL times
        app_times, url_times = self._calculate_time_distribution(batch_inactive_time)
        
        # Handle midnight crossings - check if batch spans multiple dates
        start_date = self.batch_start_time.date()
        end_date = batch_end_time.date()
        spans_midnight = start_date != end_date
        
        # Create optimized data structure with proper date handling
        optimized_data = {
            'u': self.user_id,
            'd': start_date.isoformat(),
            's': self.batch_start_time.strftime("%H:%M:%S"),
            'e': batch_end_time.strftime("%H:%M:%S"),
            'tt': total_batch_time,
            'at': active_time,
            'it': batch_inactive_time,  # Use filtered inactive time
            'spans_midnight': spans_midnight
        }
        
        # FIXED: Add metadata for long inactive periods
        if long_inactive_period_detected:
            optimized_data['long_inactive_detected'] = True
            # Calculate total original inactive time for analysis
            total_original_inactive = sum(
                period.get('original_duration', period['du']) 
                for period in valid_inactive_periods
            )
            optimized_data['original_inactive_time'] = round(total_original_inactive, 2)
            print(f"Batch contains long inactive period: {total_original_inactive/60:.1f} minutes")
        
        # Add end date if different from start date
        if spans_midnight:
            optimized_data['ed'] = end_date.isoformat()
            
            # Calculate time distribution across dates
            midnight_time = datetime.combine(end_date, datetime.min.time())
            seconds_until_midnight = (midnight_time - self.batch_start_time).total_seconds()
            seconds_after_midnight = (batch_end_time - midnight_time).total_seconds()
            
            optimized_data['time_before_midnight'] = max(0, round(seconds_until_midnight, 2))
            optimized_data['time_after_midnight'] = max(0, round(seconds_after_midnight, 2))
            
            print(f"Batch spans midnight: {start_date} to {end_date}")
            print(f"Time before midnight: {seconds_until_midnight:.1f}s")
            print(f"Time after midnight: {seconds_after_midnight:.1f}s")
        
        # Add timezone info for better tracking
        optimized_data['tz'] = self.batch_start_time.strftime("%z") or "Local"
        
        # Add apps (only if there are any with significant time)
        if app_times:
            # FIXED: Use much lower threshold to preserve legitimate work sessions
            # 0.5 seconds instead of 1.0 seconds to reduce data loss
            filtered_apps = {k: v for k, v in app_times.items() if v >= 0.5}
            if filtered_apps:
                optimized_data['ap'] = filtered_apps
        
        # Add URLs (only if there are any with significant time)
        if url_times:
            # FIXED: Use consistent lower threshold for URLs as well
            filtered_urls = {k: v for k, v in url_times.items() if v['t'] >= 0.5}
            if filtered_urls:
                optimized_data['ur'] = filtered_urls
        
        # Add only the valid inactive periods (within batch window)
        if valid_inactive_periods:
            optimized_data['ip'] = valid_inactive_periods
        

        
        return optimized_data
    

    
    def _sync_to_supabase(self):
        """Send optimized data to Supabase with proper time accounting"""
        print(f"Attempting to sync data... Activities: {len(self.activity_timeline)}, Inactive periods: {len(self.inactive_periods)}")
        
        optimized_data = self._prepare_optimized_data()
        
        if not optimized_data:
            print("No data to sync (completely empty batch)")
            self._reset_batch()
            return
        
        batch_id = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        # Prepare data for database - only include fields that exist in the schema
        db_data = {
            'batch_id': batch_id,
            'user_id': self.user_id,
            'date_tracked': optimized_data['d'],
            'batch_start_time': optimized_data['s'],
            'batch_end_time': optimized_data['e'],
            'total_time_seconds': optimized_data['tt'],
            'active_time_seconds': optimized_data['at'],
            'inactive_time_seconds': optimized_data['it'],
            'batch_data': optimized_data
        }
        
        # Note: spans_midnight and end_date_tracked are stored in batch_data JSON
        # Remove them from direct database fields to avoid schema issues
        
        # Validate the data before sending
        app_total = sum(optimized_data.get('ap', {}).values())
        print(f"Final validation: Active time={optimized_data['at']}s, App total={app_total:.2f}s")
        
        # Try to upload to Supabase
        if self.supabase_client:
            try:
                result = self.supabase_client.table('activity_summary').insert(db_data).execute()
                
                if result.data:
                    # NEW: Show if this was a purely inactive batch
                    apps_count = len(optimized_data.get('ap', {}))
                    urls_count = len(optimized_data.get('ur', {}))
                    if apps_count == 0 and urls_count == 0:
                        print(f"SUCCESS: Synced INACTIVE batch: {optimized_data['it']:.1f}s inactive (continuous tracking)")
                    else:
                        print(f"SUCCESS: Synced batch: {optimized_data['at']:.1f}s active, {apps_count} apps, {urls_count} URLs")
                    self._reset_batch()
                    return True
                else:
                    print("Failed to sync to Supabase")
            except Exception as e:
                print(f"Supabase sync error: {e}")
        
        # Save locally if Supabase fails
        self._save_locally(optimized_data, batch_id)
        self._reset_batch()
    
    def _save_locally(self, data, batch_id):
        """Save data locally if Supabase upload fails"""
        try:
            filename = self.data_directory / f"optimized_batch_{batch_id}.json"
            with open(filename, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=2)
            print(f"Saved locally: {filename.name}")
        except Exception as e:
            print(f"Failed to save locally: {e}")
    
    def _reset_batch(self):
        """Reset data for next batch"""
        self.batch_start_time = datetime.now()
        self.activity_timeline = []
        self.inactive_periods = []
        self.total_inactive_time = 0.0

    def _sync_all_pending_files(self):
        """Attempt to sync all unsent optimized JSON files to Supabase"""
        if not self.supabase_client:
            return

        # Get all optimized batch JSON files in the data directory
        json_files = list(self.data_directory.glob("optimized_batch_*.json"))

        if not json_files:
            return

        successful_uploads = []

        for json_file in sorted(json_files):  # Process in chronological order
            try:
                success = self._upload_optimized_file_to_supabase(json_file)
                if success:
                    successful_uploads.append(json_file)
                else:
                    # If one file fails, stop processing to maintain order
                    break

            except Exception as e:
                print(f"Error uploading {json_file.name}: {e}")
                break

        # Delete successfully uploaded files
        for file_path in successful_uploads:
            try:
                file_path.unlink()
                print(f"Deleted uploaded optimized file: {file_path.name}")
            except Exception as e:
                print(f"Failed to delete {file_path.name}: {e}")

        if successful_uploads:
            print(f"SUCCESS: Successfully synced {len(successful_uploads)} optimized batches")
        elif json_files:
            print(f"{len(json_files)} optimized batches pending upload (will retry)")

    def _upload_optimized_file_to_supabase(self, file_path):
        """Upload a single optimized JSON file to Supabase"""
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                data = json.load(f)

            # Structure data correctly for the optimized table
            insert_data = {
                'batch_id': data.get('bi'),
                'user_id': data.get('ui'),
                'created_time': data.get('ct'),
                'date_tracked': data.get('dt'),
                'active_time_seconds': data.get('at'),
                'inactive_time_seconds': data.get('it'),
                'batch_data': data
            }

            # Upload to Supabase
            response = self.supabase_client.table("optimized_activity_batches").insert(insert_data).execute()

            if response.data:
                print(f"SUCCESS: Uploaded optimized file {file_path.name}")
                return True
            else:
                print(f"Upload failed for: {file_path.name}")
                return False

        except Exception as e:
            print(f"Upload error for {file_path.name}: {e}")
            return False



    def set_user_inactive(self, inactive_start_time):
        """Called when user becomes inactive"""
        if not self.current_inactive_state['is_inactive']:
            self.current_inactive_state['is_inactive'] = True
            self.current_inactive_state['inactive_start_time'] = inactive_start_time
            self.current_inactive_state['inactive_start_timestamp'] = inactive_start_time.timestamp()
            print(f"User became inactive at {inactive_start_time.strftime('%H:%M:%S')}")

    def set_user_active(self):
        """Called when user becomes active again"""
        if self.current_inactive_state['is_inactive']:
            self.current_inactive_state['is_inactive'] = False
            self.current_inactive_state['inactive_start_time'] = None
            self.current_inactive_state['inactive_start_timestamp'] = None
            print(f"User became active again")

class SmartActivityDetector:
    """Detects when users are in meetings or other 'passive but productive' activities"""
    
    def __init__(self):
        self.meeting_apps = [
            'zoom', 'teams', 'skype', 'discord', 'slack', 'webex', 'gotomeeting',
            'hangouts', 'meet', 'bluejeans', 'whereby', 'jitsi', 'facetime'
        ]
        
        self.meeting_urls = [
            'zoom.us', 'teams.microsoft.com', 'meet.google.com', 'hangouts.google.com',
            'skype.com', 'webex.com', 'gotomeeting.com', 'bluejeans.com',
            'whereby.com', 'jitsi.org', 'discord.com'
        ]
        
        # Track meeting state
        self.current_meeting_app = None
        self.meeting_start_time = None
        self.is_in_meeting = False
        
    def is_meeting_application(self, app_name, window_title='', current_url=''):
        """Check if current activity is a meeting"""
        if not app_name:
            return False, "unknown"
        
        app_name_lower = app_name.lower()
        url_lower = current_url.lower() if current_url else ''
        
        # 1. DEDICATED MEETING APPLICATIONS
        for meeting_app in self.meeting_apps:
            if meeting_app in app_name_lower:
                return True, "meeting_app"
        
        # 2. KNOWN MEETING URLS
        if current_url:
            for meeting_url in self.meeting_urls:
                if meeting_url in url_lower:
                    return True, "meeting_url"
        
        return False, "none"
    
    def is_passive_productive_activity(self, app_name, window_title='', current_url=''):
        """Check if activity is passive but productive"""
        if not app_name:
            return False, "unknown"
        
        app_name_lower = app_name.lower()
        
        # Reading applications
        reading_apps = [
            'pdf', 'reader', 'preview', 'acrobat', 'kindle', 'books',
            'word', 'powerpoint', 'keynote', 'pages', 'numbers'
        ]
        
        for reading_app in reading_apps:
            if reading_app in app_name_lower:
                return True, "reading_app"
        
        return False, "none"
    
    def get_smart_idle_threshold(self, app_name, window_title='', current_url=''):
        """Get appropriate idle threshold based on activity detection"""
        is_meeting, meeting_type = self.is_meeting_application(app_name, window_title, current_url)
        is_passive, passive_type = self.is_passive_productive_activity(app_name, window_title, current_url)
        
        # Default threshold: 3 minutes (180 seconds)
        default_threshold = 180
        
        if is_meeting:
            if meeting_type == "meeting_app":
                return 900, f"meeting_app_{meeting_type}"  # 15 minutes - dedicated apps
            elif meeting_type == "meeting_url":
                return 600, f"meeting_url_{meeting_type}"  # 10 minutes - known platforms
            else:
                return 480, f"meeting_generic_{meeting_type}"  # 8 minutes - fallback
        
        elif is_passive:
            if passive_type == "reading_app":
                return 420, f"reading_app_{passive_type}"  # 7 minutes - PDF readers, etc.
            else:
                return 300, f"passive_generic_{passive_type}"  # 5 minutes - fallback
        
        return default_threshold, "default"
    
    def update_meeting_state(self, app_name, window_title='', current_url=''):
        """Track meeting state changes"""
        is_meeting, meeting_type = self.is_meeting_application(app_name, window_title, current_url)
        
        if is_meeting and not self.is_in_meeting:
            # Entering a meeting
            self.is_in_meeting = True
            self.current_meeting_app = app_name
            self.meeting_start_time = time.time()
            return "meeting_started"
        
        elif not is_meeting and self.is_in_meeting:
            # Leaving a meeting
            self.is_in_meeting = False
            meeting_duration = (time.time() - self.meeting_start_time) if self.meeting_start_time else 0
            self.current_meeting_app = None
            self.meeting_start_time = None
            return "meeting_ended", meeting_duration
        
        return "no_change"
    
    def get_activity_explanation(self, threshold_info):
        """Get human-readable explanation for why threshold was adjusted"""
        threshold_value, threshold_type = threshold_info
        
        explanations = {
            "meeting_app": f" Meeting app detected - extended idle threshold to {threshold_value//60} minutes",
            "meeting_url": f" Video meeting in browser - extended idle threshold to {threshold_value//60} minutes", 
            "reading": f" Reading/document app - extended idle threshold to {threshold_value//60} minutes",
            "default": f" Standard idle threshold: {threshold_value//60} minutes"
        }
        
        for key, explanation in explanations.items():
            if key in threshold_type:
                return explanation
        
        return f" Custom idle threshold: {threshold_value//60} minutes"

class ActivityTracker:
    def __init__(self, silent_mode=True, supabase_url=None, supabase_key=None):
        self.tracking_data = defaultdict(lambda: {
            'total_time': 0,
            'last_active': None,
            'sessions': [],
            'urls': set(),  # For browser activities
            'category': 'neutral',
            'clicks': 0,
            'keystrokes': 0
        })
        self.current_window = None
        self.last_activity_time = time.time()
        self.session_start_time = time.time()
        self.is_tracking = False
        self.idle_threshold = 180  # 3 minutes of inactivity = idle
        self.silent_mode = silent_mode
        
        # Real-time URL detection (primary method)
        self.url_detector = RealTimeURLDetector()
        
        # Browser history reading (fallback method)
        self.browser_reader = BrowserHistoryReader()
        
        self.input_tracker = InputTracker()
        self.current_session_start = datetime.now()
        self.last_browser_url = None  # Cache last browser URL to avoid repeated checks
        self.is_user_inactive = False
        self.inactive_start_time = None
        self.total_inactive_time = 0
        
        # Initialize optimized data syncer with Supabase credentials
        self.data_syncer = OptimizedDataSyncer(supabase_url, supabase_key)
        self.last_sync_time = time.time()
        
        # Initialize SmartActivityDetector
        self.smart_detector = SmartActivityDetector()

    def _calculate_session_duration(self, start_time, end_time=None):
        """
        UNIFIED TIME CALCULATION METHOD
        
        Centralized method to calculate session duration to ensure consistency
        across all time tracking operations.
        
        Args:
            start_time: Start timestamp (float from time.time())
            end_time: End timestamp (float from time.time(), defaults to current time)
            
        Returns:
            float: Session duration in seconds
        """
        if end_time is None:
            end_time = time.time()
        
        duration = end_time - start_time
        return max(0.0, duration)  # Ensure no negative durations
        
    def log(self, message):
        """Log message only if not in silent mode"""
        if not self.silent_mode:
            print(message)
    
    def get_active_window_info(self):
        """Get information about the currently active window"""
        system = platform.system()
        
        try:
            if system == "Windows":
                return self._get_windows_active_window()
            elif system == "Darwin":
                return self._get_macos_active_window()
            elif system == "Linux":
                return self._get_linux_active_window()
        except Exception as e:
            if not self.silent_mode:
                print(f"Error getting active window: {e}")
            return None, None
    
    def _get_windows_active_window(self):
        """Get active window on Windows"""
        try:
            import win32con
            
            hwnd = win32gui.GetForegroundWindow()
            
            # Check if window is minimized
            if win32gui.IsIconic(hwnd):
                return None, None
                
            # Check if window is visible
            if not win32gui.IsWindowVisible(hwnd):
                return None, None
                
            # Get window rectangle to check if it has size
            try:
                rect = win32gui.GetWindowRect(hwnd)
                width = rect[2] - rect[0]
                height = rect[3] - rect[1]
                
                # Skip if window is too small (likely hidden or minimized)
                if width < 10 or height < 10:
                    return None, None
            except:
                pass
            
            window_title = win32gui.GetWindowText(hwnd)
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            process = psutil.Process(pid)
            app_name = process.name()
            return app_name, window_title
        except:
            return None, None
    
    def _get_macos_active_window(self):
        """Get active window on macOS"""
        try:
            workspace = NSWorkspace.sharedWorkspace()
            active_app = workspace.activeApplication()
            app_name = active_app['NSApplicationName']
            
            # Check if the app has any visible windows
            window_list = Quartz.CGWindowListCopyWindowInfo(
                Quartz.kCGWindowListOptionOnScreenOnly | Quartz.kCGWindowListExcludeDesktopElements,
                Quartz.kCGNullWindowID
            )
            
            # Look for visible windows for this app
            has_visible_window = False
            window_title = app_name  # Default to app name
            
            for window in window_list:
                if window.get('kCGWindowOwnerName') == app_name:
                    window_layer = window.get('kCGWindowLayer', 0)
                    window_alpha = window.get('kCGWindowAlpha', 1.0)
                    window_bounds = window.get('kCGWindowBounds', {})
                    
                    # Skip system windows and tiny windows
                    if (window_layer < 100 and  # Normal window layers
                        window_alpha > 0.5 and  # Visible windows
                        window_bounds and
                        window_bounds.get('Height', 0) > 50 and
                        window_bounds.get('Width', 0) > 100):
                        
                        has_visible_window = True
                        # Try to get a meaningful window title
                        win_name = window.get('kCGWindowName', '')
                        if win_name and len(win_name) > 1:
                            window_title = win_name
                        break
            
            # Only return app info if it has visible windows
            if has_visible_window:
                return app_name, window_title
            else:
                # App is active but no visible windows (minimized/hidden)
                return None, None
                
        except Exception as e:
            # If there's an error, fall back to basic detection
            try:
                workspace = NSWorkspace.sharedWorkspace()
                active_app = workspace.activeApplication()
                app_name = active_app['NSApplicationName']
                return app_name, app_name
            except:
                return None, None
    
    def _get_linux_active_window(self):
        """Get active window on Linux"""
        try:
            # Using xdotool and xprop
            window_id = subprocess.check_output(['xdotool', 'getactivewindow']).decode().strip()
            window_title = subprocess.check_output(['xdotool', 'getwindowname', window_id]).decode().strip()
            
            # Get process name
            pid = subprocess.check_output(['xdotool', 'getwindowpid', window_id]).decode().strip()
            process = psutil.Process(int(pid))
            app_name = process.name()
            
            return app_name, window_title
        except:
            return None, None
    
    def get_current_urls(self):
        """Get currently visited URLs from browser history"""
        return self.browser_reader.get_recent_urls(minutes_back=1)
    
    def categorize_activity(self, app_name, window_title, url=None):
        """Categorize the activity as productive, neutral, or unproductive"""
        if not app_name:
            return "unknown"
        
        app_name_lower = app_name.lower()
        title_lower = window_title.lower() if window_title else ""
        
        # Productive applications
        productive_apps = [
            'code', 'visual studio', 'sublime', 'atom', 'notepad++', 'vim', 'cursor',
            'excel', 'word', 'powerpoint', 'outlook', 'teams', 'slack', 'zoom',
            'photoshop', 'illustrator', 'figma', 'sketch', 'canva',
            'terminal', 'cmd', 'powershell', 'iterm', 'console'
        ]
        
        # Unproductive applications
        unproductive_apps = [
            'game', 'steam', 'spotify', 'music', 'video', 'vlc', 'netflix',
            'youtube', 'twitch', 'discord', 'telegram', 'whatsapp',
            'instagram', 'tiktok', 'snapchat'
        ]
        
        # Check if it's a browser and analyze the URL/title
        browsers = ['chrome', 'firefox', 'safari', 'opera']
        edge_executables = ['msedge.exe', 'edge.exe', 'microsoftedge.exe']
        is_edge = any(edge_exe.lower() == app_name_lower for edge_exe in edge_executables)
        is_browser = any(browser in app_name_lower for browser in browsers) or is_edge
        if is_browser:
            if url:
                return self.categorize_url(url)
            else:
                return self.categorize_website(title_lower)
        
        # Check application categories
        if any(prod_app in app_name_lower for prod_app in productive_apps):
            return "productive"
        elif any(unprod_app in app_name_lower for unprod_app in unproductive_apps):
            return "unproductive"
        else:
            return "neutral"
    
    def categorize_url(self, url):
        """Categorize website based on URL"""
        try:
            domain = urllib.parse.urlparse(url).netloc.lower()
            path = urllib.parse.urlparse(url).path.lower()
            
            # Productive domains
            productive_domains = [
                'github.com', 'stackoverflow.com', 'stackexchange.com', 'docs.microsoft.com',
                'developer.mozilla.org', 'w3schools.com', 'codecademy.com', 'coursera.org',
                'udemy.com', 'linkedin.com', 'gmail.com', 'outlook.com', 'office.com',
                'google.com/drive', 'dropbox.com', 'trello.com', 'asana.com', 'jira',
                'confluence', 'notion.so', 'airtable.com', 'monday.com'
            ]
            
            # Unproductive domains
            unproductive_domains = [
                'youtube.com', 'netflix.com', 'facebook.com', 'instagram.com',
                'twitter.com', 'tiktok.com', 'reddit.com', 'twitch.tv',
                'amazon.com', 'ebay.com', 'aliexpress.com', 'shopping',
                'game', 'steam', 'entertainment', 'sports', 'news.com'
            ]
            
            # Check for exact matches or subdomains
            for prod_domain in productive_domains:
                if prod_domain in domain or domain.endswith(prod_domain):
                    return "productive"
            
            for unprod_domain in unproductive_domains:
                if unprod_domain in domain or domain.endswith(unprod_domain):
                    return "unproductive"
            
            # Special cases for search engines
            if 'google.com' in domain and '/search' in path:
                return "neutral"  # Search can be work-related
            
            return "neutral"
            
        except:
            return "neutral"
    
    def categorize_website(self, title):
        """Categorize website based on page title (fallback method)"""
        productive_keywords = [
            'github', 'stackoverflow', 'documentation', 'docs', 'tutorial',
            'learning', 'course', 'education', 'work', 'office', 'gmail',
            'calendar', 'drive', 'dropbox', 'trello', 'asana', 'jira'
        ]
        
        unproductive_keywords = [
            'youtube', 'netflix', 'facebook', 'instagram', 'twitter',
            'reddit', 'tiktok', 'game', 'entertainment', 'news',
            'shopping', 'amazon', 'ebay', 'music', 'video'
        ]
        
        if any(keyword in title for keyword in productive_keywords):
            return "productive"
        elif any(keyword in title for keyword in unproductive_keywords):
            return "unproductive"
        else:
            return "neutral"
    
    def format_time_spent(self, seconds):
        """Format seconds into readable time format"""
        if seconds < 60:
            return f"{int(seconds)}s"
        elif seconds < 3600:
            minutes = int(seconds / 60)
            secs = int(seconds % 60)
            return f"{minutes}m {secs}s"
        else:
            hours = int(seconds / 3600)
            minutes = int((seconds % 3600) / 60)
            return f"{hours}h {minutes}m"
    
    def track_activity(self):
        """Main tracking loop"""
        if not self.silent_mode:
            self.log(" Starting KeyTRK Activity Tracking...")
            self.log(f" System: {platform.system()}")
            self.log(f" Started: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            self.log("=" * 70)
        
        # Start input monitoring
        self.input_tracker.start_monitoring()
        
        while self.is_tracking:
            try:
                app_name, window_title = self.get_active_window_info()
                current_time = time.time()
                
                # Check user input activity first to determine if user is truly inactive
                input_stats = self.input_tracker.get_activity_stats()
                time_since_last_input = current_time - input_stats['last_activity']
                
                # Get smart idle threshold based on current activity
                if app_name:
                    current_url = None
                    # Quick URL check for threshold calculation
                    browsers = ['chrome', 'firefox', 'safari', 'edge', 'opera']
                    if any(browser in app_name.lower() for browser in browsers):
                        try:
                            current_url, _ = self.url_detector.get_current_browser_url(app_name, window_title)
                        except:
                            pass
                        # If no URL found, use window title as URL for threshold calculation
                        if not current_url and window_title:
                            current_url = window_title
                    
                    smart_threshold_info = self.smart_detector.get_smart_idle_threshold(app_name, window_title, current_url)
                    current_idle_threshold, threshold_type = smart_threshold_info
                    
                    # Show threshold explanation when it changes from default
                    if threshold_type != "default" and not self.silent_mode:
                        if not hasattr(self, '_last_threshold_type') or self._last_threshold_type != threshold_type:
                            explanation = self.smart_detector.get_activity_explanation(smart_threshold_info)
                            self.log(explanation)
                            self._last_threshold_type = threshold_type
                else:
                    current_idle_threshold = self.idle_threshold
                    threshold_type = "default"
                
                user_has_recent_input = time_since_last_input <= current_idle_threshold
                
                # Determine the current activity
                current_activity_key = None
                activity_app_name = None
                activity_window_title = None
                is_desktop_activity = False
                
                if not app_name:
                    # No focused window - this becomes desktop activity if user is active
                    if user_has_recent_input:
                        current_activity_key = "Desktop|Desktop Activity"
                        activity_app_name = "Desktop"
                        activity_window_title = "Desktop Activity"
                        is_desktop_activity = True
                    # If no recent input, this will be handled by inactivity logic below
                else:
                    # Check if this app should be treated as desktop activity
                    is_desktop, desktop_type = self.is_desktop_activity(app_name, window_title)
                    
                    if is_desktop and desktop_type != "no_app_focus":
                        # Special handling for lock screen - treat as inactive, not as desktop activity
                        if desktop_type == "lock_screen":
                            # Don't create an activity key for lock screen
                            current_activity_key = None
                            is_desktop_activity = False
                            if not self.silent_mode:
                                self.log(f" Lock screen detected - treating as inactive time")
                            # This will fall through to inactivity logic
                        else:
                            current_activity_key = f"Desktop|{app_name} - {desktop_type.replace('_', ' ').title()}"
                            activity_app_name = "Desktop"
                            activity_window_title = f"{app_name} - {desktop_type.replace('_', ' ').title()}"
                            is_desktop_activity = True
                    else:
                        # Regular app activity - will be handled in the normal app tracking section below
                        pass
                
                # Handle desktop activities in unified way
                if current_activity_key and user_has_recent_input:
                    # Process this activity (desktop or regular) using unified logic
                    self._process_current_activity(
                        current_activity_key, 
                        activity_app_name, 
                        activity_window_title, 
                        current_time, 
                        input_stats, 
                        is_desktop_activity,
                        None,  # current_url - none for desktop activities
                        app_name,  # original_app_name for logging
                        window_title  # original_window_title
                    )
                    time.sleep(2)
                    continue
                
                # Check for user inactivity using smart threshold
                # Note: input_stats and time_since_last_input already calculated above
                
                if time_since_last_input > current_idle_threshold:
                    # User is inactive
                    if not self.is_user_inactive:
                        # Just became inactive - close current session
                        if self.current_window and self.last_activity_time:
                            session_duration = self._calculate_session_duration(self.last_activity_time, current_time)
                            self.tracking_data[self.current_window]['total_time'] += session_duration
                            
                            # Get input stats for the session
                            self.tracking_data[self.current_window]['clicks'] += input_stats['clicks']
                            self.tracking_data[self.current_window]['keystrokes'] += input_stats['keystrokes']
                            
                            self.tracking_data[self.current_window]['sessions'].append({
                                'start_time': datetime.fromtimestamp(self.last_activity_time).isoformat(),
                                'duration': session_duration,
                                'category': self.tracking_data[self.current_window].get('category', 'neutral'),
                                'clicks': input_stats['clicks'],
                                'keystrokes': input_stats['keystrokes']
                            })
                            
                            # Send session data to OptimizedDataSyncer
                            sync_data = {
                                'app_name': self.tracking_data[self.current_window].get('app_name'),
                                'window_title': self.tracking_data[self.current_window].get('window_title'),
                                'current_url': self.tracking_data[self.current_window].get('current_url', ''),
                                'total_time': session_duration
                            }
                            self.data_syncer.add_data(self.current_window, sync_data)
                            
                            # Show time spent on previous activity
                            prev_app = self.tracking_data[self.current_window].get('app_name', 'Unknown')
                            time_spent = self.format_time_spent(session_duration)
                            self.log(f" Spent {time_spent} on {prev_app}")
                            
                            # Reset input counters
                            self.input_tracker.reset_counters()
                            
                            # Clear current window to indicate we're inactive
                            self.current_window = None
                            self.last_activity_time = None
                        
                        # Start tracking inactivity
                        self.is_user_inactive = True
                        # Set inactive start time to when user actually became inactive
                        self.inactive_start_time = input_stats['last_activity']
                        threshold_minutes = current_idle_threshold // 60
                        self.log(f" User inactive (no input for {threshold_minutes}+ minutes) - continuous inactive tracking")
                        
                        # NEW: Notify data syncer about inactive state
                        inactive_start_datetime = datetime.fromtimestamp(self.inactive_start_time)
                        self.data_syncer.set_user_inactive(inactive_start_datetime)
                    
                    # Continue counting inactive time
                    if self.inactive_start_time:
                        current_inactive_duration = self._calculate_session_duration(self.inactive_start_time, current_time)
                        if not self.silent_mode and int(current_inactive_duration) % 60 == 0:  # Every minute
                            inactive_formatted = self.format_time_spent(current_inactive_duration)
                            self.log(f" Inactive for {inactive_formatted}")
                    
                    time.sleep(2)
                    continue
                else:
                    # User is active again
                    if self.is_user_inactive:
                        # Calculate total inactive time and send to syncer
                        if self.inactive_start_time:
                            inactive_duration = self._calculate_session_duration(self.inactive_start_time, current_time)
                            self.total_inactive_time += inactive_duration
                            inactive_formatted = self.format_time_spent(inactive_duration)
                            self.log(f" User active again - was inactive for {inactive_formatted}")
                            
                            # Send inactive period to optimized syncer
                            inactive_start_datetime = datetime.fromtimestamp(self.inactive_start_time)
                            self.data_syncer.add_inactive_period(inactive_start_datetime, inactive_duration)
                        
                        # Reset inactive state
                        self.is_user_inactive = False
                        self.inactive_start_time = None
                        
                        # NEW: Notify data syncer that user is active again
                        self.data_syncer.set_user_active()
                
                if app_name and not current_activity_key:
                    # This is a regular app activity (not already handled as desktop activity above)
                    
                    browsers = ['chrome', 'firefox', 'safari', 'opera']
                    edge_executables = ['msedge.exe', 'edge.exe', 'microsoftedge.exe']
                    is_edge = any(edge_exe.lower() == app_name.lower() for edge_exe in edge_executables)
                    is_browser = any(browser in app_name.lower() for browser in browsers) or is_edge
                    
                    current_url = None
                    url_title = None
                    domain = None
                    
                    # For browsers, use real-time URL detection for immediate tab switch detection
                    if is_browser:
                        # Try real-time URL detection first (primary method)
                        try:
                            current_url, url_title = self.url_detector.get_current_browser_url(app_name, window_title)
                        except Exception as e:
                            if not self.silent_mode:
                                self.log(f" Real-time URL detection failed: {e}")
                        
                        # If real-time detection failed, fall back to browser history
                        if not current_url:
                            recent_url, url_data = self.browser_reader.get_most_recent_url()
                            if recent_url and url_data:
                                current_url = recent_url
                                url_title = url_data.get('title', '')
                        
                        # If still no URL found, use window title as URL
                        if not current_url and window_title:
                            # Clean up Edge's multi-language "and X more pages" format
                            if is_edge:
                                clean_title = None
                                
                                # English: "New tab and 2 more pages"
                                if ' and ' in window_title and ' more ' in window_title:
                                    parts = window_title.split(' and ')
                                    if len(parts) >= 2 and any(char.isdigit() for char in parts[1]):
                                        clean_title = parts[0]
                                
                                # Turkish: "Yeni sekme ve 2 sayfa daha"
                                elif ' ve ' in window_title and ' daha' in window_title:
                                    parts = window_title.split(' ve ')
                                    if len(parts) >= 2 and any(char.isdigit() for char in parts[1]):
                                        clean_title = parts[0]
                                
                                # German: "Neuer Tab und 2 weitere Seiten"
                                elif ' und ' in window_title and ' weitere' in window_title:
                                    parts = window_title.split(' und ')
                                    if len(parts) >= 2 and any(char.isdigit() for char in parts[1]):
                                        clean_title = parts[0]
                                
                                # French: "Nouvel onglet et 2 pages de plus"
                                elif ' et ' in window_title and ' de plus' in window_title:
                                    parts = window_title.split(' et ')
                                    if len(parts) >= 2 and any(char.isdigit() for char in parts[1]):
                                        clean_title = parts[0]
                                
                                # Italian: "Nuova scheda e 2 pagine in più"
                                elif ' e ' in window_title and ' in più' in window_title:
                                    parts = window_title.split(' e ')
                                    if len(parts) >= 2 and any(char.isdigit() for char in parts[1]):
                                        clean_title = parts[0]
                                
                                # Russian: "Новая вкладка и 2 страницы еще"
                                elif ' и ' in window_title and ' еще' in window_title:
                                    parts = window_title.split(' и ')
                                    if len(parts) >= 2 and any(char.isdigit() for char in parts[1]):
                                        clean_title = parts[0]
                                
                                # Portuguese: "Nova aba e 2 páginas mais"
                                elif ' e ' in window_title and ' mais' in window_title:
                                    parts = window_title.split(' e ')
                                    if len(parts) >= 2 and any(char.isdigit() for char in parts[1]):
                                        clean_title = parts[0]
                                
                                # Chinese: "新标签页和 2 个页面"
                                elif '和' in window_title and '个页面' in window_title:
                                    parts = window_title.split('和')
                                    if len(parts) >= 2 and any(char.isdigit() for char in parts[1]):
                                        clean_title = parts[0]
                                
                                # Japanese: "新しいタブと 2 ページ"
                                elif 'と' in window_title and 'ページ' in window_title:
                                    parts = window_title.split('と')
                                    if len(parts) >= 2 and any(char.isdigit() for char in parts[1]):
                                        clean_title = parts[0]
                                
                                # Spanish: "Nueva pestaña y 2 páginas más"
                                elif ' y ' in window_title and ' más' in window_title:
                                    parts = window_title.split(' y ')
                                    if len(parts) >= 2 and any(char.isdigit() for char in parts[1]):
                                        clean_title = parts[0]
                                
                                # Apply the cleaned title if we found a pattern
                                if clean_title:
                                    # Use the cleaned title (removes "and X more pages" but keeps the actual page name)
                                    window_title = clean_title
                                    url_title = clean_title
                                else:
                                    current_url = window_title
                                    url_title = window_title
                            else:
                                current_url = window_title
                                url_title = window_title
                        
                        # Process the URL if we found one
                        if current_url:
                            domain = urllib.parse.urlparse(current_url).netloc
                            
                            # Detect tab switches by comparing URLs
                            url_changed = (
                                not hasattr(self, 'last_browser_url') or 
                                self.last_browser_url != current_url
                            )
                            
                            if url_changed and not self.silent_mode:
                                self.log(f" Browser URL: {domain}")
                                if url_title:
                                    self.log(f" Page: {url_title[:60]}{'...' if len(url_title) > 60 else ''}")
                        
                        # Store URL for next comparison
                        if current_url:
                            self.last_browser_url = current_url
                    
                    # Track meeting state changes
                    meeting_state = self.smart_detector.update_meeting_state(app_name, window_title, current_url)
                    if meeting_state == "meeting_started" and not self.silent_mode:
                        self.log(" Meeting detected - smart activity tracking activated")
                    elif isinstance(meeting_state, tuple) and meeting_state[0] == "meeting_ended" and not self.silent_mode:
                        meeting_duration = meeting_state[1]
                        duration_formatted = self.format_time_spent(meeting_duration)
                        self.log(f" Meeting ended - total meeting time: {duration_formatted}")
                    
                    # Create activity key (include URL if available)
                    if current_url:
                        activity_key = f"{app_name}|{current_url}"
                        display_title = url_title or window_title
                    else:
                        activity_key = f"{app_name}|{window_title}"
                        display_title = window_title
                    
                    category = self.categorize_activity(app_name, window_title, current_url)
                    
                    # Add meeting context information for logging
                    is_meeting, meeting_type = self.smart_detector.is_meeting_application(app_name, window_title, current_url)
                    is_passive, passive_type = self.smart_detector.is_passive_productive_activity(app_name, window_title, current_url)
                    
                    # Process this regular app activity using unified method
                    self._process_current_activity(
                        activity_key, 
                        app_name, 
                        window_title, 
                        current_time, 
                        input_stats, 
                        False,  # is_desktop = False for regular apps
                        current_url,
                        app_name,  # original_app_name
                        window_title  # original_window_title
                    )
                    
                    # Update additional metadata for regular apps
                    if activity_key in self.tracking_data:
                        self.tracking_data[activity_key]['category'] = category
                        self.tracking_data[activity_key]['is_meeting'] = is_meeting
                        if is_meeting:
                            self.tracking_data[activity_key]['meeting_type'] = meeting_type
                        
                        self.tracking_data[activity_key]['is_passive_productive'] = is_passive
                        if is_passive:
                            self.tracking_data[activity_key]['passive_type'] = passive_type
                        
                        if current_url:
                            self.tracking_data[activity_key]['urls'].add(current_url)
                            self.tracking_data[activity_key]['current_url'] = current_url
                
                # Note: OptimizedDataSyncer handles its own 5-minute sync timing
                # We just need to send data to it when sessions end
                
                time.sleep(2)  # Check every 2 seconds
            except KeyboardInterrupt:
                break
            except Exception as e:
                if not self.silent_mode:
                    print(f"Error in tracking loop: {e}")
                time.sleep(5)
    
    def start_tracking(self):
        """Start the tracking process"""
        self.is_tracking = True
        
        # Start data syncing first
        self.data_syncer.start_syncing()
        
        self.tracking_thread = threading.Thread(target=self.track_activity)
        self.tracking_thread.daemon = True
        self.tracking_thread.start()
    
    def stop_tracking(self):
        """Stop tracking and save final session"""
        self.is_tracking = False
        
        # Handle any active inactive period
        if self.is_user_inactive and self.inactive_start_time:
            current_time = time.time()
            inactive_duration = self._calculate_session_duration(self.inactive_start_time, current_time)
            inactive_start_datetime = datetime.fromtimestamp(self.inactive_start_time)
            self.data_syncer.add_inactive_period(inactive_start_datetime, inactive_duration)
        
        # Save the last session
        if self.current_window:
            session_duration = self._calculate_session_duration(self.last_activity_time)
            self.tracking_data[self.current_window]['total_time'] += session_duration
            
            # Get final input stats
            input_stats = self.input_tracker.get_activity_stats()
            self.tracking_data[self.current_window]['clicks'] += input_stats['clicks']
            self.tracking_data[self.current_window]['keystrokes'] += input_stats['keystrokes']
            
            self.tracking_data[self.current_window]['sessions'].append({
                'start_time': datetime.fromtimestamp(self.last_activity_time).isoformat(),
                'duration': session_duration,
                'category': self.tracking_data[self.current_window].get('category', 'neutral'),
                'clicks': input_stats['clicks'],
                'keystrokes': input_stats['keystrokes']
            })
            
            # Send final data to syncer
            self.data_syncer.add_data(self.current_window, self.tracking_data[self.current_window])
        
        # Stop input monitoring and data syncing
        self.input_tracker.stop_monitoring()
        self.data_syncer.stop_syncing()
    
    def generate_report(self):
        """Generate a summary report of the tracking session"""
        if not self.tracking_data:
            self.log("No tracking data available.")
            return
        
        self.log("\n" + "="*70)
        self.log(" KEYTRK DETAILED ACTIVITY REPORT")
        self.log("="*70)
        
        # Calculate totals by category and meeting types
        category_totals = defaultdict(float)
        meeting_totals = defaultdict(float)
        passive_totals = defaultdict(float)
        total_time = 0
        total_clicks = 0
        total_keystrokes = 0
        total_meeting_time = 0
        total_passive_time = 0
        
        # Sort activities by total time
        sorted_activities = sorted(
            self.tracking_data.items(),
            key=lambda x: x[1]['total_time'],
            reverse=True
        )
        
        self.log("\n TOP ACTIVITIES:")
        self.log("-" * 70)
        
        for i, (activity_key, data) in enumerate(sorted_activities[:15]):
            duration_minutes = data['total_time'] / 60
            category = data.get('category', 'neutral')
            category_totals[category] += data['total_time']
            total_time += data['total_time']
            total_clicks += data.get('clicks', 0)
            total_keystrokes += data.get('keystrokes', 0)
            
            # Track meeting and passive time
            if data.get('is_meeting', False):
                meeting_type = data.get('meeting_type', 'unknown')
                meeting_totals[meeting_type] += data['total_time']
                total_meeting_time += data['total_time']
            
            if data.get('is_passive_productive', False):
                passive_type = data.get('passive_type', 'unknown')
                passive_totals[passive_type] += data['total_time']
                total_passive_time += data['total_time']
            
            app_name = data.get('app_name', 'Unknown')
            window_title = data.get('window_title', 'Unknown')
            current_url = data.get('current_url', '')
            clicks = data.get('clicks', 0)
            keystrokes = data.get('keystrokes', 0)
            
            # Add emoji for meeting/passive activities
            activity_emoji = ""
            if data.get('is_meeting', False):
                activity_emoji = ""
            elif data.get('is_passive_productive', False):
                activity_emoji = ""
            
            self.log(f"{i+1:2d}. {activity_emoji} {app_name}")
            
            # Add context information
            if data.get('is_meeting', False):
                meeting_type = data.get('meeting_type', 'unknown')
                self.log(f"     Meeting ({meeting_type})")
            elif data.get('is_passive_productive', False):
                passive_type = data.get('passive_type', 'unknown')
                self.log(f"     Passive productive ({passive_type})")
            
            if current_url:
                domain = urllib.parse.urlparse(current_url).netloc
                self.log(f"     {domain}")
                self.log(f"     {current_url}")
                self.log(f"     {window_title[:50]}{'...' if len(window_title) > 50 else ''}")
            else:
                self.log(f"     {window_title[:50]}{'...' if len(window_title) > 50 else ''}")
            
            time_formatted = self.format_time_spent(data['total_time'])
            self.log(f"      {time_formatted} |   {clicks} clicks |   {keystrokes} keys")
            self.log("")
        
        self.log(f"\n Total Tracked Time: {self.format_time_spent(total_time)}")
        if self.total_inactive_time > 0:
            self.log(f" Total Inactive Time: {self.format_time_spent(self.total_inactive_time)}")
        self.log(f"  Total Mouse Clicks: {total_clicks:,}")
        self.log(f"  Total Keystrokes: {total_keystrokes:,}")
        
        # Calculate desktop activity statistics
        desktop_totals = defaultdict(float)
        total_desktop_time = 0
        
        for activity_key, data in self.tracking_data.items():
            if data.get('is_desktop', False):
                desktop_type = data.get('desktop_type', 'unknown')
                desktop_totals[desktop_type] += data['total_time']
                total_desktop_time += data['total_time']
        
        # Show meeting, passive, and desktop activity statistics
        if total_meeting_time > 0 or total_passive_time > 0 or total_desktop_time > 0:
            self.log("\n SMART ACTIVITY BREAKDOWN:")
            self.log("-" * 70)
            
            if total_meeting_time > 0:
                meeting_percentage = (total_meeting_time / total_time) * 100
                self.log(f" Total Meeting Time: {self.format_time_spent(total_meeting_time)} ({meeting_percentage:.1f}%)")
                
                for meeting_type, time_spent in meeting_totals.items():
                    if time_spent > 0:
                        type_percentage = (time_spent / total_meeting_time) * 100
                        self.log(f"{meeting_type.replace('_', ' ').title()}: {self.format_time_spent(time_spent)} ({type_percentage:.1f}%)")
            
            if total_passive_time > 0:
                passive_percentage = (total_passive_time / total_time) * 100
                self.log(f"\n Total Passive Productive Time: {self.format_time_spent(total_passive_time)} ({passive_percentage:.1f}%)")
                
                for passive_type, time_spent in passive_totals.items():
                    if time_spent > 0:
                        type_percentage = (time_spent / total_passive_time) * 100
                        self.log(f"{passive_type.replace('_', ' ').title()}: {self.format_time_spent(time_spent)} ({type_percentage:.1f}%)")
            
            if total_desktop_time > 0:
                desktop_percentage = (total_desktop_time / total_time) * 100
                self.log(f"\n Total Desktop/File Management Time: {self.format_time_spent(total_desktop_time)} ({desktop_percentage:.1f}%)")
                
                for desktop_type, time_spent in desktop_totals.items():
                    if time_spent > 0:
                        type_percentage = (time_spent / total_desktop_time) * 100
                        self.log(f"{desktop_type.replace('_', ' ').title()}: {self.format_time_spent(time_spent)} ({type_percentage:.1f}%)")
            
            # Calculate truly interactive time
            interactive_time = total_time - total_meeting_time - total_passive_time - total_desktop_time
            if interactive_time > 0:
                interactive_percentage = (interactive_time / total_time) * 100
                self.log(f"\n Interactive Work Time: {self.format_time_spent(interactive_time)} ({interactive_percentage:.1f}%)")
        
        # Save detailed data to JSON
        self.save_data_to_file()
    
    def save_data_to_file(self):
        """Save tracking data to JSON file"""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"keytrk_data_{timestamp}.json"
        
        # Convert data for JSON serialization
        total_tracked_time = sum(data['total_time'] for data in self.tracking_data.values())
        json_data = {
            'tracking_session': {
                'start_time': self.current_session_start.isoformat(),
                'end_time': datetime.now().isoformat(),
                'system': platform.system(),
                'total_activities': len(self.tracking_data),
                'total_tracked_time_seconds': total_tracked_time,
                'total_inactive_time_seconds': self.total_inactive_time
            },
            'activities': {},
            'inactive_periods': []  # Add inactive periods to JSON export
        }
        
        for activity_key, data in self.tracking_data.items():
            json_data['activities'][activity_key] = {
                'app_name': data.get('app_name'),
                'window_title': data.get('window_title'),
                'current_url': data.get('current_url', ''),
                'all_urls': list(data.get('urls', set())),
                'category': data.get('category'),
                'total_time_seconds': data['total_time'],
                'total_time_minutes': data['total_time'] / 60,
                'total_clicks': data.get('clicks', 0),
                'total_keystrokes': data.get('keystrokes', 0),
                'sessions': data['sessions'],
                'last_active': data['last_active'],
                'is_meeting': data.get('is_meeting', False),
                'meeting_type': data.get('meeting_type'),
                'is_passive_productive': data.get('is_passive_productive', False),
                'passive_type': data.get('passive_type'),
                'is_desktop': data.get('is_desktop', False),
                'desktop_type': data.get('desktop_type')
            }
        
        # Add inactive periods data if available from OptimizedDataSyncer
        if hasattr(self, 'data_syncer') and hasattr(self.data_syncer, 'inactive_periods'):
            json_data['inactive_periods'] = [
                {
                    'start_time': period['s'],
                    'duration_seconds': period['du']
                }
                for period in self.data_syncer.inactive_periods
            ]
        
        with open(filename, 'w', encoding='utf-8') as f:
            json.dump(json_data, f, indent=2, ensure_ascii=False)
        
        self.log(f" Data saved to: {filename}")

    def is_desktop_activity(self, app_name, window_title=''):
        """Check if the current activity is desktop or file management related"""
        if not app_name:
            return True, "no_app_focus"
        
        app_name_lower = app_name.lower()
        title_lower = window_title.lower() if window_title else ''
        
        # Check for LockApp (Windows lock screen) and Screen Saver - treat as inactive, not an app
        # Most conservative - only exact matches, no partial string searches
        if (app_name_lower in ['lockapp', 'scrnsave'] or 
            title_lower in ['lock screen', 'screen saver']):
            return True, "lock_screen"
        
        # File managers and desktop applications
        desktop_apps = [
            'finder', 'explorer', 'nautilus', 'dolphin', 'thunar', 'pcmanfm',
            'file manager', 'files', 'file browser', 'folder', 'directory',
            'desktop', 'wallpaper', 'screen saver', 'screensaver',
            'system preferences', 'control panel', 'settings',
            'activity monitor', 'task manager', 'system monitor',
            'terminal', 'console', 'cmd', 'powershell', 'iterm'
        ]
        
        # Check app name
        for desktop_app in desktop_apps:
            if desktop_app in app_name_lower:
                return True, f"desktop_app_{desktop_app.replace(' ', '_')}"
        
        # Check for desktop-related window titles
        desktop_keywords = [
            'desktop', 'folder', 'file', 'directory', 'downloads', 'documents',
            'pictures', 'music', 'videos', 'applications', 'system',
            'home', 'root', 'trash', 'recycle bin'
        ]
        
        for keyword in desktop_keywords:
            if keyword in title_lower:
                return True, f"desktop_context_{keyword.replace(' ', '_')}"
        
        return False, "regular_app"

    def _process_current_activity(self, activity_key, app_name, window_title, current_time, input_stats, 
                                  is_desktop, current_url=None, original_app_name=None, original_window_title=None):
        """Unified method to process current activity (desktop or regular app)"""
        
        # If this is a new activity or resuming from inactive state
        if self.current_window != activity_key:
            # Close previous activity session if we had one
            if self.current_window and self.last_activity_time:
                session_duration = self._calculate_session_duration(self.last_activity_time, current_time)
                self.tracking_data[self.current_window]['total_time'] += session_duration
                
                # Get input stats for the session
                self.tracking_data[self.current_window]['clicks'] += input_stats['clicks']
                self.tracking_data[self.current_window]['keystrokes'] += input_stats['keystrokes']
                
                self.tracking_data[self.current_window]['sessions'].append({
                    'start_time': datetime.fromtimestamp(self.last_activity_time).isoformat(),
                    'duration': session_duration,
                    'category': self.tracking_data[self.current_window].get('category', 'neutral'),
                    'clicks': input_stats['clicks'],
                    'keystrokes': input_stats['keystrokes']
                })
                
                # Send session data to OptimizedDataSyncer
                sync_data = {
                    'app_name': self.tracking_data[self.current_window].get('app_name'),
                    'window_title': self.tracking_data[self.current_window].get('window_title'),
                    'current_url': self.tracking_data[self.current_window].get('current_url', ''),
                    'total_time': session_duration
                }
                self.data_syncer.add_data(self.current_window, sync_data)
                
                # Show time spent on previous activity
                prev_app = self.tracking_data[self.current_window].get('app_name', 'Unknown')
                time_spent = self.format_time_spent(session_duration)
                self.log(f" Spent {time_spent} on {prev_app}")
                
                # Reset input counters
                self.input_tracker.reset_counters()
            
            # Start tracking new activity
            self.current_window = activity_key
            self.last_activity_time = current_time
            self.session_start_time = current_time
            
            # Initialize activity data if it doesn't exist
            if activity_key not in self.tracking_data:
                self.tracking_data[activity_key] = {
                    'total_time': 0,
                    'last_active': None,
                    'sessions': [],
                    'urls': set(),
                    'category': 'neutral',
                    'clicks': 0,
                    'keystrokes': 0,
                    'app_name': app_name,
                    'window_title': window_title,
                    'current_url': current_url or '',
                    'is_desktop': is_desktop
                }
                
                if is_desktop:
                    self.tracking_data[activity_key]['desktop_type'] = 'desktop_activity'
            
            self.tracking_data[activity_key]['last_active'] = datetime.now().isoformat()
            
            # Log activity start
            if not self.silent_mode:
                if is_desktop:
                    self.log(f" {window_title}")
                    if original_window_title:
                        self.log(f" {original_window_title[:60]}{'...' if len(original_window_title) > 60 else ''}")
                else:
                    self.log(f" {original_app_name or app_name}")
                    if original_window_title:
                        self.log(f" {original_window_title[:60]}{'...' if len(original_window_title) > 60 else ''}")
        
        # Show periodic updates for current activity
        elif current_time - self.session_start_time > 30 and int(current_time - self.session_start_time) % 30 == 0:
            session_time = self._calculate_session_duration(self.session_start_time, current_time)
            time_formatted = self.format_time_spent(session_time)
            
            if not self.silent_mode:
                activity_type = "Desktop" if is_desktop else "App"
                self.log(f" {activity_type} session: {time_formatted} | Clicks: {input_stats['clicks']} | Keys: {input_stats['keystrokes']}")

def ensure_single_instance():
    """Ensure only one instance of the tracker is running"""
    import os
    import sys

    # Create a lock file in temp directory
    lock_file = os.path.join(tempfile.gettempdir(), 'activity_tracker.lock')

    if os.path.exists(lock_file):
        try:
            # Check if the process in the lock file is still running
            with open(lock_file, 'r') as f:
                pid = int(f.read().strip())

            # Check if process is still running
            if platform.system() == "Windows":
                import subprocess
                try:
                    subprocess.check_output(['tasklist', '/FI', f'PID eq {pid}'], stderr=subprocess.DEVNULL)
                    sys.exit(1)
                except subprocess.CalledProcessError:
                    # Process not running, remove stale lock file
                    os.remove(lock_file)
            else:
                # Unix-like systems
                try:
                    os.kill(pid, 0)  # Check if process exists
                    sys.exit(1)
                except OSError:
                    # Process not running, remove stale lock file
                    os.remove(lock_file)
        except (ValueError, FileNotFoundError):
            # Invalid lock file, remove it
            try:
                os.remove(lock_file)
            except FileNotFoundError:
                pass

    # Create new lock file with current PID
    with open(lock_file, 'w') as f:
        f.write(str(os.getpid()))

    # Clean up lock file on exit
    import atexit
    atexit.register(lambda: os.remove(lock_file) if os.path.exists(lock_file) else None)

def main():
    """Main function to run the activity tracker"""
    # Ensure only one instance is running
    ensure_single_instance()

    # Configuration - Try to import from config.py, fall back to environment variables
    SUPABASE_URL = None
    SUPABASE_KEY = None
    SILENT_MODE = True  # Default to silent mode for production
    
    try:
        import config
        SUPABASE_URL = config.SUPABASE_URL
        SUPABASE_KEY = config.SUPABASE_KEY
        SILENT_MODE = getattr(config, 'SILENT_MODE', False)
        # Configuration loaded successfully
    except ImportError:
        # Fall back to environment variables
        SUPABASE_URL = os.getenv('SUPABASE_URL')  # e.g., 'https://your-project.supabase.co'
        SUPABASE_KEY = os.getenv('SUPABASE_KEY')  # Your Supabase anon key
        if SUPABASE_URL and SUPABASE_KEY:
            print("SUCCESS: Configuration loaded from environment variables")
        else:
            print("No configuration found. Copy config_example.py to config.py and update it.")
    
    # For employee deployment, this should run in silent mode
    # For testing, set silent_mode=False in config.py
    tracker = ActivityTracker(
        silent_mode=SILENT_MODE,
        supabase_url=SUPABASE_URL,
        supabase_key=SUPABASE_KEY
    )
    
    try:
        tracker.start_tracking()
        if not tracker.silent_mode:
            print("Tracking in progress... Press Enter to stop and generate report")
            print("Data is being saved every 5 minutes and synced to Supabase")
            input()  # Wait for user input
        else:
            # In production, this runs continuously
            print("Starting continuous tracking mode...")
            while True:
                time.sleep(60)  # Check every minute
                # The tracker will continue running and saving data every 5 minutes
        
    except KeyboardInterrupt:
        print("Stopping tracker...")
    
    finally:
        tracker.stop_tracking()
        tracker.generate_report()

if __name__ == "__main__":
    main() 
