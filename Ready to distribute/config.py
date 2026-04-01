#!/usr/bin/env python3
"""
KeyTRK Configuration - Production Deployment
"""

# Supabase Configuration - REAL CREDENTIALS
SUPABASE_URL = "https://btkiqffcjvjyqyokccfh.supabase.co"
SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImJ0a2lxZmZjanZqeXF5b2tjY2ZoIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NTA0ODc3OTQsImV4cCI6MjA2NjA2Mzc5NH0.9OHJcPdD-GIpBiUZuLl8NySwj5e0W4-JtV1u2o5LA9U"

# Tracking Configuration - Production Optimized
SILENT_MODE = True  # Production deployment
SYNC_INTERVAL = 300  # Data sync interval in seconds (5 minutes)

# Optimized Tracking Configuration
TRACK_CATEGORIES = False        # No productivity categorization
TRACK_URLS = True              # YES - Keep URL tracking for browsers only
TRACK_KEYSTROKES = False       # No keystroke counting  
TRACK_CLICKS = False           # No click counting
TRACK_DETAILED_SESSIONS = False # No detailed session data
TITLE_MAX_LENGTH = 60          # Truncate page titles to 60 characters
MIN_TIME_THRESHOLD = 1.00      # Ignore activities less than 1 second

# User identification function (required)
def get_user_id():
    """Return a unique user identifier for tracking"""
    import os
    import platform
    return f"{os.getenv('USERNAME', 'user')}@{platform.node()}"

# Database Table Configuration
# Make sure you create this table in your Supabase database:
"""
CREATE TABLE activity_summary (
    id SERIAL PRIMARY KEY,
    batch_id VARCHAR(50) NOT NULL,
    user_id VARCHAR(100) NOT NULL,
    date_tracked DATE NOT NULL,
    batch_start_time TIME NOT NULL,
    batch_end_time TIME NOT NULL,
    total_time_seconds DECIMAL(10,2) NOT NULL,
    active_time_seconds DECIMAL(10,2) NOT NULL,
    inactive_time_seconds DECIMAL(10,2) NOT NULL,
    batch_data JSONB NOT NULL,
    created_at TIMESTAMP WITH TIME ZONE DEFAULT NOW(),
    uploaded_at TIMESTAMP WITH TIME ZONE DEFAULT NOW()
);

-- Indexes for performance
CREATE INDEX idx_activity_summary_user_date ON activity_summary(user_id, date_tracked);
CREATE INDEX idx_activity_summary_apps ON activity_summary USING GIN ((batch_data->'ap'));
CREATE INDEX idx_activity_summary_urls ON activity_summary USING GIN ((batch_data->'ur'));
""" 
