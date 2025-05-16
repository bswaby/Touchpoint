#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Touchpoint User Activity Analysis Tool

This script provides a comprehensive analysis of user activity in Touchpoint,
including tracking changes, views, estimated working time, password issues,
account creation, stale accounts, and more. 

The tool allows filtering by person or viewing overall system activity trends,
and provides various reports to help administrators understand how users are
interacting with the system.

--Upload Instructions Start--
To upload code to Touchpoint, use the following steps:
1. Click Admin ~ Advanced ~ Special Content ~ Python
2. Click New Python Script File
3. Name the Python and paste all this code
4. Test and optionally add to menu
--Upload Instructions End--

Written By: Ben Swaby
Email: bswaby@fbchtn.org
"""

import datetime
import time
import re
import traceback
import json
from datetime import timedelta

# ==========================================
# Configuration Options (Modify as needed)
# ==========================================

# Time gap in minutes that indicates a user has stopped working
INACTIVITY_THRESHOLD_MINUTES = 10

# Number of days to check for stale accounts
STALE_ACCOUNT_DAYS = 90

# Number of recent activities to show in single person view
RECENT_ACTIVITY_COUNT = 50

# Minimum events to count as a "session"
MIN_EVENTS_FOR_SESSION = 3

# Office IP addresses - used to determine if activity is from office or remote
OFFICE_IP_ADDRESSES = ['12.23.240.162', '173.166.241.105']

# ==========================================
# Helper Classes
# ==========================================

class ActivityAnalyzer:
    """
    Analyzes user activity data from Touchpoint.
    """
    def __init__(self, model, query):
        self.model = model
        self.query = query
        
    def get_user_info(self, user_id):
        """Get basic user information."""
        try:
            # Debug output
            print "<div style='display:none'>get_user_info received user_id: '{0}' type: {1}</div>".format(
                str(user_id), type(user_id).__name__)
            
            # Force user_id to be an integer
            if not isinstance(user_id, int):
                # Try to convert to string first, then to int to handle various types
                user_id = int(str(user_id))
                
            sql = """
                SELECT u.UserId, u.PeopleId, u.Username, u.Name, u.EmailAddress, 
                    u.LastLoginDate, u.LastActivityDate, u.CreationDate, u.IsLockedOut,
                    u.FailedPasswordAttemptCount, u.MustChangePassword
                FROM Users u
                WHERE u.UserId = {0}
            """.format(user_id)
            
            # Debug the SQL
            print "<div style='display:none'>SQL query: {0}</div>".format(sql)
            
            return self.query.QuerySqlTop1(sql)
        except Exception as e:
            print "<div style='display:none'>Error in get_user_info: {0}</div>".format(str(e))
            return None
    
    def get_users_list(self):
        """Get a list of all users with recent activity."""
        try:
            sql = """
                SELECT TOP 100 u.UserId, u.PeopleId, u.Username, u.Name, u.LastLoginDate, 
                    u.LastActivityDate, p.EmailAddress
                FROM Users u
                JOIN People p ON p.PeopleId = u.PeopleId
                ORDER BY u.LastActivityDate DESC
            """
            return self.query.QuerySql(sql)
        except Exception as e:
            # Log the error
            print "<div style='display:none'>Error in get_users_list: {0}</div>".format(str(e))
            # Return an empty list if there's an error
            return []
    
    def get_stale_accounts(self, days=None):
        """Get accounts that haven't been active for a specified number of days."""
        try:
            if days is None:
                days = STALE_ACCOUNT_DAYS
                
            sql = """
                SELECT u.UserId, u.PeopleId, u.Username, u.Name, u.LastLoginDate, 
                    u.LastActivityDate, u.CreationDate, p.EmailAddress
                FROM Users u
                JOIN People p ON p.PeopleId = u.PeopleId
                WHERE DATEDIFF(day, u.LastActivityDate, GETDATE()) > {0}
                    AND u.IsLockedOut = 0
                ORDER BY u.LastActivityDate
            """.format(days)
            
            return self.query.QuerySql(sql)
        except Exception as e:
            print "<div style='display:none'>Error in get_stale_accounts: {0}</div>".format(str(e))
            return []
    
    def get_locked_accounts(self):
        """Get accounts that are currently locked."""
        try:
            sql = """
                SELECT u.UserId, u.PeopleId, u.Username, u.Name, u.LastLoginDate, 
                    u.LastActivityDate, u.LastLockedOutDate, u.FailedPasswordAttemptCount,
                    p.EmailAddress
                FROM Users u
                JOIN People p ON p.PeopleId = u.PeopleId
                WHERE u.IsLockedOut = 1
                ORDER BY u.LastLockedOutDate DESC
            """
            return self.query.QuerySql(sql)
        except Exception as e:
            print "<div style='display:none'>Error in get_locked_accounts: {0}</div>".format(str(e))
            return []
    
    def get_recent_password_resets(self, days=7):
        """Get accounts with recent password resets."""
        try:
            sql = """
                SELECT u.UserId, u.PeopleId, u.Username, u.Name, u.LastPasswordChangedDate,
                    p.EmailAddress
                FROM Users u
                JOIN People p ON p.PeopleId = u.PeopleId
                WHERE DATEDIFF(day, u.LastPasswordChangedDate, GETDATE()) <= {0}
                ORDER BY u.LastPasswordChangedDate DESC
            """.format(days)
            
            return self.query.QuerySql(sql)
        except Exception as e:
            print "<div style='display:none'>Error in get_recent_password_resets: {0}</div>".format(str(e))
            return []
    
    def get_user_activity(self, user_id, limit=None):
        """Get activity log for a specific user."""
        try:
            if limit is None:
                limit = RECENT_ACTIVITY_COUNT
            
            # Ensure user_id is an integer
            user_id = int(user_id)
            
            # Debug output
            print "<div style='display:none'>get_user_activity called with user_id: {0}, type: {1}</div>".format(
                user_id, type(user_id).__name__)
                
            sql = """
                SELECT TOP {0} ActivityDate, Activity, PageUrl, OrgId, PeopleId, Mobile, ClientIp
                FROM ActivityLog
                WHERE UserId = {1}
                ORDER BY ActivityDate DESC
            """.format(limit, user_id)
            
            # Debug SQL
            print "<div style='display:none'>Activity SQL: {0}</div>".format(sql)
            
            return self.query.QuerySql(sql)
        except Exception as e:
            # Log error and return empty list
            print "<div style='display:none'>Error in get_user_activity: {0}</div>".format(str(e))
            return []
    
    def get_user_activity_count(self, user_id, days=30):
        """Get count of activities for a user in the last X days."""
        try:
            sql = """
                SELECT COUNT(*) AS Count
                FROM ActivityLog
                WHERE UserId = {0}
                    AND DATEDIFF(day, ActivityDate, GETDATE()) <= {1}
            """.format(user_id, days)  # Use string formatting for days parameter
            
            result = self.query.QuerySqlTop1(sql)
            return result.Count if result else 0
        except:
            # Return a default if query fails
            return 0
    
    def get_activity_stats_by_period(self, days=30, period_type='daily'):
        """Get activity statistics grouped by time period."""
        if period_type == 'daily':
            # Use a format that will show date only (MM/DD/YYYY)
            date_part = "CONVERT(varchar(10), ActivityDate, 101)"
        elif period_type == 'weekly':
            date_part = "CONVERT(varchar(10), DATEADD(day, -DATEPART(weekday, ActivityDate)+1, CONVERT(date, ActivityDate)), 101)"
        elif period_type == 'monthly':
            date_part = "CONVERT(varchar(7), ActivityDate, 120)"
        else:
            date_part = "CONVERT(varchar(10), ActivityDate, 101)"
            
        sql = """
            SELECT 
                {0} AS Period,
                COUNT(*) AS ActivityCount,
                COUNT(DISTINCT UserId) AS UserCount
            FROM ActivityLog
            WHERE DATEDIFF(day, ActivityDate, GETDATE()) <= {1}
            GROUP BY {0}
            ORDER BY {0} DESC
        """.format(date_part, days)
        
        try:
            # Try to execute the query with the specified parameters
            return self.query.QuerySql(sql)
        except:
            # If there's an error, try a simpler fallback query
            fallback_sql = """
                SELECT 
                    CONVERT(varchar(10), ActivityDate, 101) AS Period,
                    COUNT(*) AS ActivityCount,
                    COUNT(DISTINCT UserId) AS UserCount
                FROM ActivityLog
                WHERE DATEDIFF(day, ActivityDate, GETDATE()) <= {0}
                GROUP BY CONVERT(varchar(10), ActivityDate, 101)
                ORDER BY CONVERT(varchar(10), ActivityDate, 101) DESC
            """.format(days)
            return self.query.QuerySql(fallback_sql)
    
    def get_most_active_users(self, days=30):
        """Get the most active users in the last X days."""
        try:
            sql = """
                SELECT 
                    al.UserId,
                    u.Username,
                    u.Name,
                    COUNT(*) AS ActivityCount,
                    COUNT(DISTINCT CONVERT(date, al.ActivityDate)) AS DaysActive,
                    MIN(al.ActivityDate) AS FirstActivity,
                    MAX(al.ActivityDate) AS LastActivity
                FROM ActivityLog al
                JOIN Users u ON u.UserId = al.UserId
                WHERE DATEDIFF(day, al.ActivityDate, GETDATE()) <= {0}
                GROUP BY al.UserId, u.Username, u.Name
                HAVING COUNT(*) > 0
                ORDER BY ActivityCount DESC
            """.format(days)
            
            return self.query.QuerySql(sql)
        except Exception as e:
            print "<div style='display:none'>Error in get_most_active_users: {0}</div>".format(str(e))
            return []
    
    def categorize_activities(self, activities):
        """Categorize activities by type."""
        categories = {}
        
        for activity in activities:
            if not hasattr(activity, 'Activity'):
                continue
                
            # Split activity by colon to get the activity type
            activity_parts = activity.Activity.split(':', 1)
            activity_type = activity_parts[0].strip() if len(activity_parts) > 0 else 'Other'
            
            # Count this activity type
            if activity_type in categories:
                categories[activity_type] += 1
            else:
                categories[activity_type] = 1
                
        return categories
    
    def analyze_user_location_stats(self, user_id, days=30):
        """Analyze a user's activity locations (office, remote, mobile)."""
        try:
            # Ensure user_id is an integer
            user_id = int(user_id)
            
            sql = """
                SELECT ActivityDate, Mobile, ClientIp
                FROM ActivityLog
                WHERE UserId = {0}
                    AND DATEDIFF(day, ActivityDate, GETDATE()) <= {1}
                ORDER BY ActivityDate
            """.format(user_id, days)
            
            activities = self.query.QuerySql(sql)
            
            stats = {
                'office': 0,
                'remote': 0,
                'mobile': 0,
                'total': 0
            }
            
            for activity in activities:
                stats['total'] += 1
                
                # Check if mobile
                if hasattr(activity, 'Mobile') and activity.Mobile:
                    stats['mobile'] += 1
                    continue
                    
                # Check if office IP
                if hasattr(activity, 'ClientIp'):
                    client_ip = str(activity.ClientIp) if activity.ClientIp else ''
                    if client_ip in OFFICE_IP_ADDRESSES:
                        stats['office'] += 1
                    else:
                        stats['remote'] += 1
                else:
                    stats['remote'] += 1  # Default to remote if IP not available
                    
            # Calculate percentages
            stats['office_pct'] = (stats['office'] * 100.0 / stats['total']) if stats['total'] > 0 else 0
            stats['remote_pct'] = (stats['remote'] * 100.0 / stats['total']) if stats['total'] > 0 else 0
            stats['mobile_pct'] = (stats['mobile'] * 100.0 / stats['total']) if stats['total'] > 0 else 0
            
            return stats
        except Exception as e:
            print "<div style='display:none'>Error in analyze_user_location_stats: {0}</div>".format(str(e))
            # Return default stats
            return {
                'office': 0,
                'remote': 0,
                'mobile': 0,
                'total': 0,
                'office_pct': 0,
                'remote_pct': 0,
                'mobile_pct': 0
            }
    
    def analyze_user_sessions(self, user_id, days=30):
        """
        Analyze a user's sessions to estimate working time.
        A session ends when there's a gap of INACTIVITY_THRESHOLD_MINUTES.
        """
        try:
            # Ensure user_id is an integer
            user_id = int(user_id)
            
            sql = """
                SELECT ActivityDate, Activity, PageUrl, OrgId, PeopleId, Mobile, ClientIp
                FROM ActivityLog
                WHERE UserId = {0}
                    AND DATEDIFF(day, ActivityDate, GETDATE()) <= {1}
                ORDER BY ActivityDate
            """.format(user_id, days)
            
            activities = self.query.QuerySql(sql)
            
            if not activities:
                return {
                    'total_sessions': 0,
                    'total_duration': timedelta(0),
                    'total_duration_hours': 0,
                    'location_stats': {'office': timedelta(0), 'remote': timedelta(0), 'mobile': timedelta(0)},
                    'location_hours': {'office': 0, 'remote': 0, 'mobile': 0},
                    'sessions': [],
                    'activity_categories': {}
                }
            
            # Process into sessions - if activities exist but sessions aren't being created,
            # force at least one session with a reasonable duration
            sessions = []
            current_session = []
            previous_time = None
            
            for activity in activities:
                activity_time = activity.ActivityDate
                
                # If this is the first activity or there's a gap, start a new session
                if (previous_time is None or 
                    (activity_time - previous_time).total_seconds() > (INACTIVITY_THRESHOLD_MINUTES * 60)):
                    
                    # Save the previous session if it has enough events
                    if len(current_session) >= MIN_EVENTS_FOR_SESSION:
                        sessions.append(current_session)
                        
                    # Start a new session
                    current_session = [activity]
                else:
                    # Continue the current session
                    current_session.append(activity)
                    
                previous_time = activity_time
            
            # Add the last session if it has enough events
            if len(current_session) >= MIN_EVENTS_FOR_SESSION:
                sessions.append(current_session)
            
            # If we have activities but no sessions, create a session from all activities
            # This ensures we have at least one session if there are activities
            if len(activities) > 0 and len(sessions) == 0:
                # Force all activities into one session regardless of time gaps
                sessions.append(list(activities))
                print "<p>Forced all {0} activities into a single session</p>".format(len(activities))
                
            # Calculate session stats
            session_stats = []
            total_duration = timedelta(0)
            location_stats = {
                'office': timedelta(0),
                'remote': timedelta(0),
                'mobile': timedelta(0)
            }
            
            # Categorize activities by type - do this even if there are no proper sessions
            activity_categories = {}
            for activity in activities:
                if not hasattr(activity, 'Activity'):
                    continue
                    
                # Split activity by colon to get the activity type
                activity_parts = activity.Activity.split(':', 1)
                activity_type = activity_parts[0].strip() if len(activity_parts) > 0 else 'Other'
                
                # Count this activity type
                if activity_type in activity_categories:
                    activity_categories[activity_type] += 1
                else:
                    activity_categories[activity_type] = 1
            
            for session in sessions:
                if len(session) < 2:  # Need at least 2 activities for a valid session
                    continue
                    
                start_time = session[0].ActivityDate
                end_time = session[-1].ActivityDate
                
                # If start and end time are the same, add a reasonable duration (15 minutes)
                if start_time == end_time:
                    end_time = start_time + timedelta(minutes=15)
                    
                duration = end_time - start_time
                event_count = len(session)
                
                # Determine location for this session
                session_location = 'remote'  # Default
                
                # Check if mobile
                if hasattr(session[0], 'Mobile') and session[0].Mobile:
                    session_location = 'mobile'
                # Check if office IP
                elif hasattr(session[0], 'ClientIp'):
                    client_ip = str(session[0].ClientIp) if session[0].ClientIp else ''
                    if client_ip in OFFICE_IP_ADDRESSES:
                        session_location = 'office'
                
                # Only count sessions with a reasonable duration
                if duration.total_seconds() > 0:
                    total_duration += duration
                    location_stats[session_location] += duration
                    
                    session_stats.append({
                        'start_time': start_time,
                        'end_time': end_time,
                        'duration': duration,
                        'duration_minutes': duration.total_seconds() / 60,
                        'event_count': event_count,
                        'location': session_location
                    })
            
            # If we still have no valid sessions but have activities, 
            # estimate total work time based on activity count
            if len(session_stats) == 0 and len(activities) > 0:
                # Assume an average of 5 minutes per activity as a fallback
                total_estimated_minutes = len(activities) * 5
                total_duration = timedelta(minutes=total_estimated_minutes)
                
                # Estimate location breakdown based on the activities
                office_count = 0
                remote_count = 0
                mobile_count = 0
                
                for activity in activities:
                    if hasattr(activity, 'Mobile') and activity.Mobile:
                        mobile_count += 1
                    elif hasattr(activity, 'ClientIp'):
                        client_ip = str(activity.ClientIp) if activity.ClientIp else ''
                        if client_ip in OFFICE_IP_ADDRESSES:
                            office_count += 1
                        else:
                            remote_count += 1
                    else:
                        remote_count += 1
                        
                total_count = float(office_count + remote_count + mobile_count)
                if total_count > 0:
                    office_pct = office_count / total_count
                    remote_pct = remote_count / total_count
                    mobile_pct = mobile_count / total_count
                    
                    location_stats['office'] = timedelta(minutes=total_estimated_minutes * office_pct)
                    location_stats['remote'] = timedelta(minutes=total_estimated_minutes * remote_pct)
                    location_stats['mobile'] = timedelta(minutes=total_estimated_minutes * mobile_pct)
                    
                # Create a single estimated session
                session_stats.append({
                    'start_time': activities[0].ActivityDate,
                    'end_time': activities[-1].ActivityDate,
                    'duration': total_duration,
                    'duration_minutes': total_estimated_minutes,
                    'event_count': len(activities),
                    'location': 'estimated'
                })
                
                print "<p>Created estimated work time based on {0} activities</p>".format(len(activities))
            
            return {
                'total_sessions': len(session_stats),
                'total_duration': total_duration,
                'total_duration_hours': total_duration.total_seconds() / 3600,
                'location_stats': location_stats,
                'location_hours': {
                    'office': location_stats['office'].total_seconds() / 3600,
                    'remote': location_stats['remote'].total_seconds() / 3600,
                    'mobile': location_stats['mobile'].total_seconds() / 3600
                },
                'sessions': session_stats,
                'activity_categories': activity_categories
            }
        except Exception as e:
            print "<div style='display:none'>Error in analyze_user_sessions: {0}</div>".format(str(e))
            # Return empty data structure
            return {
                'total_sessions': 0,
                'total_duration': timedelta(0),
                'total_duration_hours': 0,
                'location_stats': {'office': timedelta(0), 'remote': timedelta(0), 'mobile': timedelta(0)},
                'location_hours': {'office': 0, 'remote': 0, 'mobile': 0},
                'sessions': [],
                'activity_categories': {}
            }


class FormHandler:
    """
    Handles form input and validation.
    """
    def __init__(self, model):
        self.model = model
        self.errors = []
        
    def get_param(self, param_name, default=None):
        """Safely get a parameter from form data."""
        try:
            # First try to get it from model.Data
            if hasattr(self.model.Data, param_name):
                value = getattr(self.model.Data, param_name)
                return str(value) if value is not None else default
                
            # If not found, try to get it from Request.QueryString if available
            if hasattr(self.model, 'Request') and hasattr(self.model.Request, 'QueryString'):
                qs = self.model.Request.QueryString
                if qs and param_name in qs:
                    return str(qs[param_name])
            
            # If we're here, parameter was not found
            return default
        except:
            return default
    
    def get_int_param(self, param_name, default=None):
        """Get an integer parameter from form data."""
        try:
            value = self.get_param(param_name)
            
            # Debug output
            print "<div style='display:none'>Parameter '{0}' value: '{1}' type: {2}</div>".format(
                param_name, str(value), type(value).__name__)
            
            # If it's already an integer, return it
            if isinstance(value, int):
                return value
                
            # If we got a string, try to convert to integer
            if isinstance(value, basestring):
                # Remove any non-numeric characters
                clean_value = ''.join(c for c in value if c.isdigit())
                if clean_value:
                    return int(clean_value)
                    
            return default
        except Exception as e:
            print "<div style='display:none'>Error in get_int_param for {0}: {1}</div>".format(param_name, str(e))
            return default
            
    def get_date_param(self, param_name, default=None):
        """Get a date parameter from form data."""
        try:
            value = self.get_param(param_name)
            return self.model.ParseDate(value) if value else default
        except:
            return default


class ReportRenderer:
    """
    Renders various reports and UI components.
    """
    def __init__(self, model):
        self.model = model
        
    def render_page_header(self, title, subtitle=None):
        """Render a consistent page header."""
        html = """
        <div class="panel panel-primary">
            <div class="panel-heading">
                <h1 class="panel-title">{0}<svg xmlns="http://www.w3.org/2000/svg" viewBox="85 75 230 130" style="width: 60px; height: 30px; margin-left: -4px; vertical-align: middle;">
                    <!-- Text portion - TP -->
                    <text x="100" y="120" font-family="Arial, sans-serif" font-weight="bold" font-size="60" fill="#333333">TP</text>
                    
                    <!-- Circular element -->
                    <g transform="translate(190, 107)">
                      <!-- Outer circle -->
                      <circle cx="0" cy="0" r="13.5" fill="#0099FF"/>
                      
                      <!-- White middle circle -->
                      <circle cx="0" cy="0" r="10.5" fill="white"/>
                      
                      <!-- Inner circle -->
                      <circle cx="0" cy="0" r="7.5" fill="#0099FF"/>
                      
                      <!-- X crossing through the circles -->
                      <path d="M-9 -9 L9 9 M-9 9 L9 -9" stroke="white" stroke-width="1.8" stroke-linecap="round"/>
                    </g>
                    
                    <!-- Single "i" letter to the right -->
                    <text x="206" y="105" font-family="Arial, sans-serif" font-weight="bold" font-size="14" fill="#0099FF">si</text>
                  </svg></h1>
                {1}
            </div>
        </div>
        """.format(
            title,
            "<h4>{0}</h4>".format(subtitle) if subtitle else ""
        )
        return html
    
    def render_navigation(self, current_view):
        """Render the navigation menu."""
        views = [
            ('overview', 'Overview'),
            ('user_list', 'User List'),
            ('stale_accounts', 'Stale Accounts'),
            ('locked_accounts', 'Locked Accounts'),
            ('password_resets', 'Recent Password Resets'),
            ('activity_trends', 'Activity Trends')
        ]
        
        # Simple navigation that doesn't rely on JavaScript
        html = """<ul class="nav nav-tabs" style="margin-bottom: 15px;">"""
        
        for view_id, view_name in views:
            active = ' class="active"' if view_id == current_view else ''
            # Use direct links instead of JavaScript
            html += """
            <li{1}><a href="?view={0}">{2}</a></li>
            """.format(view_id, active, view_name)
            
        html += "</ul>"
        return html
    
    def render_filter_form(self, days=30, period_type='daily'):
        """Render a filter form for reports."""
        # Get the current view from Data, defaulting to overview
        current_view = 'overview'
        if hasattr(self.model.Data, 'view') and self.model.Data.view:
            current_view = str(self.model.Data.view)
        
        html = """
        <form method="get" class="form-inline" style="margin-bottom: 20px;">
            <input type="hidden" name="view" value="{0}">
            <div class="form-group">
                <label for="days">Time Period:</label>
                <select name="days" id="days" class="form-control">
        """.format(current_view)
        
        for d in [7, 14, 30, 60, 90, 180, 365]:
            selected = ' selected' if d == days else ''
            html += '<option value="{0}"{1}>Last {0} Days</option>'.format(d, selected)
            
        html += """
                </select>
            </div>
        """
        
        if current_view == 'activity_trends':
            html += """
            <div class="form-group" style="margin-left: 10px;">
                <label for="period_type">Group By:</label>
                <select name="period_type" id="period_type" class="form-control">
            """
            
            for p, label in [('daily', 'Daily'), ('weekly', 'Weekly'), ('monthly', 'Monthly')]:
                selected = ' selected' if p == period_type else ''
                html += '<option value="{0}"{1}>{2}</option>'.format(p, selected, label)
                
            html += """
                </select>
            </div>
            """
            
        html += """
            <button type="submit" class="btn btn-primary" style="margin-left: 10px;">Apply Filters</button>
        </form>
        """
        
        return html
    
    def render_user_search_form(self):
        """Render a user search form."""
        html = """
        <form method="get" class="form-inline" style="margin-bottom: 20px;">
            <input type="hidden" name="view" value="user_list">
            <div class="form-group">
                <label for="search_name">Search Users:</label>
                <input type="text" name="search_name" id="search_name" class="form-control" 
                       placeholder="Name or Username" style="width: 250px;">
            </div>
            <button type="submit" class="btn btn-primary" style="margin-left: 10px;">Search</button>
        </form>
        """
        return html
    
    def render_time_duration(self, seconds):
        """Render a time duration in a human-readable format."""
        if seconds < 60:
            return "{0} seconds".format(int(seconds))
        elif seconds < 3600:
            return "{0:.1f} minutes".format(seconds / 60)
        else:
            return "{0:.1f} hours".format(seconds / 3600)
    
    def render_date(self, date):
        """Render a date in a standard format."""
        if date is None:
            return "Never"
        
        # Handle different date formats safely
        try:
            # First try using strftime if available
            if hasattr(date, 'strftime'):
                return date.strftime("%b %d, %Y %I:%M %p")
            # If it's a string that looks like a date, return it formatted
            elif isinstance(date, basestring):
                # Try to parse the string into a datetime object
                try:
                    parsed_date = self.model.ParseDate(date)
                    if parsed_date:
                        return parsed_date.strftime("%b %d, %Y %I:%M %p")
                except:
                    # If parsing fails, just return the string
                    pass
                return date
            # For any other case, convert to string
            return str(date)
        except:
            # Fallback to string representation if all else fails
            try:
                return str(date)
            except:
                return "Unknown Date"
    
    def render_days_ago(self, date):
        """Render days ago from a date."""
        if date is None:
            return "N/A"
            
        delta = datetime.datetime.now() - date
        return "{0} days ago".format(delta.days)
    
    def render_activity_tooltip(self, activity):
        """Render a tooltip for an activity."""
        tooltip = activity.Activity
        
        if activity.PageUrl:
            tooltip += "\nURL: " + activity.PageUrl
            
        if activity.OrgId:
            tooltip += "\nOrg ID: " + str(activity.OrgId)
            
        if activity.PeopleId:
            tooltip += "\nPeople ID: " + str(activity.PeopleId)
            
        return tooltip.replace('"', '&quot;')
    
    def render_activity_type_icon(self, activity):
        """Render an icon for an activity type."""
        act_text = activity.Activity.lower()
        
        if 'view' in act_text or 'display' in act_text:
            return '<i class="fa fa-eye" title="View"></i>'
        elif 'edit' in act_text or 'update' in act_text or 'save' in act_text:
            return '<i class="fa fa-pencil" title="Edit"></i>'
        elif 'add' in act_text:
            return '<i class="fa fa-plus" title="Add"></i>'
        elif 'delete' in act_text or 'remove' in act_text:
            return '<i class="fa fa-trash" title="Delete"></i>'
        elif 'search' in act_text or 'find' in act_text or 'query' in act_text:
            return '<i class="fa fa-search" title="Search"></i>'
        elif 'report' in act_text or 'export' in act_text:
            return '<i class="fa fa-file-text" title="Report"></i>'
        elif 'login' in act_text or 'logon' in act_text:
            return '<i class="fa fa-sign-in" title="Login"></i>'
        else:
            return '<i class="fa fa-circle" title="Other"></i>'
    
def render_user_detail(self, user, analyzer):
    """Render detailed user information and activity."""
    try:
        # Initialize html variable at the beginning
        html = ""
        
        # Get user ID first thing and ensure it's an integer
        user_id = getattr(user, 'UserId', 0)
        print "<p>Original user_id from user object: '{0}', type: {1}</p>".format(
            str(user_id), type(user_id).__name__)
        
        # Convert to integer explicitly
        user_id = int(str(user_id))
        print "<p>Converted user_id: {0}, type: {1}</p>".format(
            user_id, type(user_id).__name__)
        
        # Initialize required data
        print "<p>Getting user activity...</p>"
        activities = analyzer.get_user_activity(user_id)
        
        # Convert activities to a list and count them directly
        try:
            activity_list = list(activities) if activities else []
            activity_count = len(activity_list)
            print "<p>Direct activity count: {0}</p>".format(activity_count)
        except Exception as e:
            print "<p>Error counting activities: {0}</p>".format(str(e))
            activity_list = []
            activity_count = 0
        
        print "<p>Analyzing user sessions...</p>"
        try:
            session_data = analyzer.analyze_user_sessions(user_id)
            print "<p>Session data retrieved successfully</p>"
        except Exception as e:
            print "<p>Error in analyze_user_sessions: {0}</p>".format(str(e))
            session_data = {
                'total_sessions': 0,
                'total_duration': timedelta(0),
                'total_duration_hours': 0,
                'location_stats': {'office': timedelta(0), 'remote': timedelta(0), 'mobile': timedelta(0)},
                'location_hours': {'office': 0, 'remote': 0, 'mobile': 0},
                'sessions': [],
                'activity_categories': {}
            }
        
        print "<p>Analyzing location stats...</p>"
        try:
            location_stats = analyzer.analyze_user_location_stats(user_id)
            print "<p>Location stats retrieved successfully</p>"
        except Exception as e:
            print "<p>Error in analyze_user_location_stats: {0}</p>".format(str(e))
            location_stats = {
                'office': 0,
                'remote': 0,
                'mobile': 0,
                'total': 0,
                'office_pct': 0,
                'remote_pct': 0,
                'mobile_pct': 0
            }
        
        # Basic user info
        html = """
        <div class="row">
            <div class="col-md-6">
                <div class="panel panel-default">
                    <div class="panel-heading">
                        <h3 class="panel-title">User Information</h3>
                    </div>
                    <div class="panel-body">
                        <table class="table table-striped">
                            <tr>
                                <th>User ID:</th>
                                <td>{0}</td>
                            </tr>
                            <tr>
                                <th>Username:</th>
                                <td>{1}</td>
                            </tr>
                            <tr>
                                <th>Name:</th>
                                <td>{2}</td>
                            </tr>
                            <tr>
                                <th>Email:</th>
                                <td>{3}</td>
                            </tr>
                            <tr>
                                <th>Person ID:</th>
                                <td><a href="/Person2/{4}" target="_blank">{4}</a></td>
                            </tr>
                            <tr>
                                <th>Account Created:</th>
                                <td>{5}</td>
                            </tr>
                            <tr>
                                <th>Last Login:</th>
                                <td>{6}</td>
                            </tr>
                            <tr>
                                <th>Last Activity:</th>
                                <td>{7}</td>
                            </tr>
                            <tr>
                                <th>Account Status:</th>
                                <td>{8}</td>
                            </tr>
                        </table>
                    </div>
                </div>
            </div>
        """.format(
            user_id,  # Use our converted user_id
            getattr(user, 'Username', 'N/A'),
            getattr(user, 'Name', 'N/A'),
            getattr(user, 'EmailAddress', 'N/A'),
            getattr(user, 'PeopleId', 'N/A'),
            self.render_date(getattr(user, 'CreationDate', None)),
            self.render_date(getattr(user, 'LastLoginDate', None)),
            self.render_date(getattr(user, 'LastActivityDate', None)),
            "LOCKED" if getattr(user, 'IsLockedOut', 0) == 1 else "Active"
        )
        
        # Override metrics with directly calculated values from activities
        total_activities = activity_count  # Use direct count
        
        # Estimate session count based on activity count
        estimated_sessions = max(1, activity_count // 10)  # Assume ~10 activities per session
        
        # Estimate total work hours based on activity count (5 minutes per activity)
        estimated_work_hours = (activity_count * 5) / 60.0
        
        # Estimate average session duration
        avg_session_minutes = 0
        if estimated_sessions > 0:
            avg_session_minutes = (estimated_work_hours * 60) / estimated_sessions
        
        # Activity statistics panel with our estimated values
        html += """
            <div class="col-md-6">
                <div class="panel panel-default">
                    <div class="panel-heading">
                        <h3 class="panel-title">Activity Summary (Last 30 Days)</h3>
                    </div>
                    <div class="panel-body">
                        <div class="row">
                            <div class="col-md-6">
                                <div class="panel panel-info">
                                    <div class="panel-heading text-center">
                                        <h4>{0}</h4>
                                        <p>Total Activities</p>
                                    </div>
                                </div>
                            </div>
                            <div class="col-md-6">
                                <div class="panel panel-info">
                                    <div class="panel-heading text-center">
                                        <h4>{1}</h4>
                                        <p>Estimated Sessions</p>
                                    </div>
                                </div>
                            </div>
                        </div>
                        <div class="row">
                            <div class="col-md-6">
                                <div class="panel panel-success">
                                    <div class="panel-heading text-center">
                                        <h4>{2:.1f} hrs</h4>
                                        <p>Estimated Work Time</p>
                                    </div>
                                </div>
                            </div>
                            <div class="col-md-6">
                                <div class="panel panel-success">
                                    <div class="panel-heading text-center">
                                        <h4>{3:.1f} mins</h4>
                                        <p>Avg Session Duration</p>
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
        """.format(
            total_activities,
            estimated_sessions,
            estimated_work_hours,
            avg_session_minutes
        )
        
        # Override location percentages if needed
        if total_activities > 0:
            # Count activities by location
            office_count = 0
            remote_count = 0
            mobile_count = 0
            
            for activity in activity_list:
                try:
                    if hasattr(activity, 'Mobile') and activity.Mobile:
                        mobile_count += 1
                    elif hasattr(activity, 'ClientIp'):
                        client_ip = str(activity.ClientIp) if activity.ClientIp else ''
                        if client_ip in OFFICE_IP_ADDRESSES:
                            office_count += 1
                        else:
                            remote_count += 1
                    else:
                        remote_count += 1
                except Exception as e:
                    print "<p>Error categorizing activity location: {0}</p>".format(str(e))
                    remote_count += 1  # Default to remote on error
                    
            total_location_count = float(office_count + remote_count + mobile_count)
            
            if total_location_count > 0:
                office_pct = (office_count / total_location_count) * 100
                remote_pct = (remote_count / total_location_count) * 100
                mobile_pct = (mobile_count / total_location_count) * 100
                
                # Calculate hours per location
                office_hours = (estimated_work_hours * office_count) / total_location_count
                remote_hours = (estimated_work_hours * remote_count) / total_location_count
                mobile_hours = (estimated_work_hours * mobile_count) / total_location_count
            else:
                # Default values if no locations could be determined
                office_pct = remote_pct = 0
                mobile_pct = 0
                office_hours = remote_hours = mobile_hours = 0
                
                # Use percentages from location_stats if available
                if location_stats:
                    office_pct = location_stats.get('office_pct', 0)
                    remote_pct = location_stats.get('remote_pct', 0)
                    mobile_pct = location_stats.get('mobile_pct', 0)
        else:
            # No activities, use default or existing values
            office_pct = location_stats.get('office_pct', 0)
            remote_pct = location_stats.get('remote_pct', 0)
            mobile_pct = location_stats.get('mobile_pct', 0)
            office_hours = remote_hours = mobile_hours = 0
            
        # Work Location Summary panel
        html += """
        <div class="row">
            <div class="col-md-12">
                <div class="panel panel-default">
                    <div class="panel-heading">
                        <h3 class="panel-title">Work Location Summary</h3>
                    </div>
                    <div class="panel-body">
                        <div class="row">
                            <div class="col-md-4">
                                <div class="panel panel-primary">
                                    <div class="panel-heading text-center">
                                        <h4>{0:.1f} hrs ({1:.1f}%)</h4>
                                        <p>Office Work</p>
                                    </div>
                                </div>
                            </div>
                            <div class="col-md-4">
                                <div class="panel panel-warning">
                                    <div class="panel-heading text-center">
                                        <h4>{2:.1f} hrs ({3:.1f}%)</h4>
                                        <p>Remote Work</p>
                                    </div>
                                </div>
                            </div>
                            <div class="col-md-4">
                                <div class="panel panel-info">
                                    <div class="panel-heading text-center">
                                        <h4>{4:.1f} hrs ({5:.1f}%)</h4>
                                        <p>Mobile Work</p>
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
        """.format(
            office_hours, office_pct,
            remote_hours, remote_pct,
            mobile_hours, mobile_pct
        )
        
        # Activity Categories
        activity_categories = {}
        
        # Categorize activities manually
        for activity in activity_list:
            try:
                if hasattr(activity, 'Activity'):
                    activity_text = str(activity.Activity)
                    # Split by colon to get the category
                    parts = activity_text.split(':', 1)
                    category = parts[0].strip() if parts else "Other"
                    
                    if category in activity_categories:
                        activity_categories[category] += 1
                    else:
                        activity_categories[category] = 1
            except Exception as e:
                print "<p>Error categorizing activity: {0}</p>".format(str(e))
        
        # Sort activity categories
        try:
            sorted_categories = sorted(activity_categories.items(), key=lambda x: x[1], reverse=True)
            total_activities_categorized = sum(activity_categories.values()) if activity_categories else 0
        except Exception as e:
            print "<p>Error sorting activity categories: {0}</p>".format(str(e))
            sorted_categories = []
            total_activities_categorized = 0
        
        # Activity Categories panel
        html += """
        <div class="row">
            <div class="col-md-12">
                <div class="panel panel-default">
                    <div class="panel-heading">
                        <h3 class="panel-title">Activity Categories</h3>
                    </div>
                    <div class="panel-body">
                        <table class="table table-striped">
                            <thead>
                                <tr>
                                    <th>Activity Type</th>
                                    <th>Count</th>
                                    <th>Percentage</th>
                                </tr>
                            </thead>
                            <tbody>
        """
        
        if sorted_categories:
            for category, count in sorted_categories:
                try:
                    percentage = (float(count) * 100.0 / total_activities_categorized) if total_activities_categorized > 0 else 0
                    html += """
                    <tr>
                        <td>{0}</td>
                        <td>{1}</td>
                        <td>{2:.1f}%</td>
                    </tr>
                    """.format(category, count, percentage)
                except Exception as e:
                    print "<p>Error rendering category {0}: {1}</p>".format(category, str(e))
        else:
            html += """
            <tr>
                <td colspan="3" class="text-center">No activity categories available</td>
            </tr>
            """
            
        html += """
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>
        </div>
        """
        
        # Recent Activity
        html += """
        <div class="panel panel-default">
            <div class="panel-heading">
                <h3 class="panel-title">Recent Activity</h3>
            </div>
            <div class="panel-body">
                <table class="table table-striped">
                    <thead>
                        <tr>
                            <th>Date & Time</th>
                            <th>Activity Type</th>
                            <th>Details</th>
                            <th>Location</th>
                            <th>Links</th>
                        </tr>
                    </thead>
                    <tbody>
        """
        
        # Display recent activities
        if activity_list:
            for i, activity in enumerate(activity_list[:20]):  # Show most recent 20 activities
                try:
                    # Safely get activity text
                    activity_text = ""
                    if hasattr(activity, 'Activity'):
                        activity_text = str(activity.Activity) if activity.Activity else ""
                    
                    # Split activity by colon to get the main activity type
                    activity_parts = activity_text.split(':', 1)
                    activity_type = activity_parts[0].strip() if len(activity_parts) > 0 else ''
                    activity_details = activity_parts[1].strip() if len(activity_parts) > 1 else activity_text
                    
                    # Determine location
                    location = "Remote"
                    location_badge = "warning"
                    
                    if hasattr(activity, 'Mobile') and activity.Mobile:
                        location = "Mobile"
                        location_badge = "info"
                    elif hasattr(activity, 'ClientIp'):
                        client_ip = str(activity.ClientIp) if activity.ClientIp else ''
                        if client_ip in OFFICE_IP_ADDRESSES:
                            location = "Office"
                            location_badge = "primary"
                    
                    # Format the activity date
                    activity_date = ""
                    if hasattr(activity, 'ActivityDate'):
                        activity_date = self.render_date(activity.ActivityDate)
                    
                    # Limit details length
                    details_text = activity_details[:100]
                    if len(activity_details) > 100:
                        details_text += '...'
                    
                    html += """
                    <tr>
                        <td>{0}</td>
                        <td><strong>{1}</strong></td>
                        <td>{2}</td>
                        <td><span class="label label-{4}">{3}</span></td>
                        <td>
                    """.format(
                        activity_date,
                        activity_type,
                        details_text,
                        location,
                        location_badge
                    )
                    
                    # Add links for related data
                    org_id = None
                    if hasattr(activity, 'OrgId'):
                        if isinstance(activity.OrgId, (int, basestring)):
                            org_id = activity.OrgId
                        elif isinstance(activity.OrgId, slice):
                            print "<p>WARNING: activity.OrgId is a slice: {0}</p>".format(activity.OrgId)
                    
                    if org_id:
                        html += '<a href="/Org/{0}" target="_blank" class="btn btn-xs btn-info">Org</a> '.format(org_id)
                    
                    people_id = None
                    if hasattr(activity, 'PeopleId'):
                        if isinstance(activity.PeopleId, (int, basestring)):
                            people_id = activity.PeopleId
                        elif isinstance(activity.PeopleId, slice):
                            print "<p>WARNING: activity.PeopleId is a slice: {0}</p>".format(activity.PeopleId)
                    
                    if people_id:
                        html += '<a href="/Person2/{0}" target="_blank" class="btn btn-xs btn-primary">Person</a> '.format(people_id)
                    
                    page_url = None
                    if hasattr(activity, 'PageUrl'):
                        if activity.PageUrl:
                            page_url = str(activity.PageUrl)
                    
                    if page_url:
                        html += '<a href="{0}" target="_blank" class="btn btn-xs btn-default">URL</a>'.format(page_url)
                    
                    html += """
                        </td>
                    </tr>
                    """
                except Exception as e:
                    print "<p>Error rendering activity {0}: {1}</p>".format(i+1, str(e))
        else:
            html += """
            <tr>
                <td colspan="5" class="text-center">No recent activities found</td>
            </tr>
            """
            
        html += """
                    </tbody>
                </table>
            </div>
        </div>
        """
        
        return html
    except Exception as e:
        import traceback
        error_trace = traceback.format_exc()
        return """
        <div class='alert alert-danger'>
            Error rendering user detail: {0}
            <br>
            <pre>{1}</pre>
        </div>
        """.format(str(e), error_trace)
    
    def render_users_table(self, users, include_activity_count=False, analyzer=None):
        """Render a table of users."""
        html = """
        <table class="table table-striped table-hover">
            <thead>
                <tr>
                    <th>ID</th>
                    <th>Username</th>
                    <th>Name</th>
                    <th>Email</th>
                    <th>Last Login</th>
                    <th>Last Activity</th>
        """
        
        if include_activity_count:
            html += "<th>Activity Count</th>"
            
        html += """
                    <th>Actions</th>
                </tr>
            </thead>
            <tbody>
        """
        
        for user in users:
            html += """
            <tr>
                <td>{0}</td>
                <td>{1}</td>
                <td>{2}</td>
                <td>{3}</td>
                <td>{4}</td>
                <td>{5}</td>
            """.format(
                getattr(user, 'UserId', 'N/A'),
                getattr(user, 'Username', 'N/A'),
                getattr(user, 'Name', 'N/A'),
                getattr(user, 'EmailAddress', 'N/A'),
                self.render_date(getattr(user, 'LastLoginDate', None)),
                self.render_date(getattr(user, 'LastActivityDate', None))
            )
            
            if include_activity_count and analyzer:
                # For activity count, simply get a placeholder value if we can't calculate it
                try:
                    count = analyzer.get_user_activity_count(user.UserId, 30)
                except:
                    count = 0
                html += "<td>{0}</td>".format(count)
                
            html += """
                <td>
                    <a href="?view=user_detail&user_id={0}" class="btn btn-xs btn-primary">View Activity</a>
                    <a href="/Person2/{1}" target="_blank" class="btn btn-xs btn-info">View Person</a>
                </td>
            </tr>
            """.format(getattr(user, 'UserId', 0), getattr(user, 'PeopleId', 0))
            
        html += """
            </tbody>
        </table>
        """
        
        return html
    
    def render_locked_accounts_table(self, accounts):
        """Render a table of locked accounts."""
        html = """
        <table class="table table-striped table-hover">
            <thead>
                <tr>
                    <th>Username</th>
                    <th>Name</th>
                    <th>Email</th>
                    <th>Last Login</th>
                    <th>Locked Since</th>
                    <th>Failed Attempts</th>
                    <th>Actions</th>
                </tr>
            </thead>
            <tbody>
        """
        
        for account in accounts:
            html += """
            <tr>
                <td>{0}</td>
                <td>{1}</td>
                <td>{2}</td>
                <td>{3}</td>
                <td>{4}</td>
                <td>{5}</td>
                <td>
                    <a href="?view=user_detail&user_id={6}" class="btn btn-xs btn-primary">View Activity</a>
                    <a href="/Person2/{7}" target="_blank" class="btn btn-xs btn-info">View Person</a>
                </td>
            </tr>
            """.format(
                getattr(account, 'Username', 'Unknown'),
                getattr(account, 'Name', 'Unknown'),
                getattr(account, 'EmailAddress', 'N/A'),
                self.render_date(getattr(account, 'LastLoginDate', None)),
                self.render_date(getattr(account, 'LastLockedOutDate', None)),
                getattr(account, 'FailedPasswordAttemptCount', 0),
                getattr(account, 'UserId', 0),
                getattr(account, 'PeopleId', 0)
            )
            
        html += """
            </tbody>
        </table>
        """
        
        return html
    
    def render_activity_stats_chart(self, stats, period_type):
        """Render a chart of activity statistics."""
        # Prepare data for chart
        labels = []
        activity_counts = []
        user_counts = []
        
        for stat in stats:
            # Format the period label based on the period type
            if period_type == 'daily':
                # Handle datetime object in a way that works with Python 2.7
                if hasattr(stat.Period, 'strftime'):
                    date_label = stat.Period.strftime('%b %d')
                else:
                    # Handle string date if that's what we got
                    date_label = str(stat.Period)
                labels.append(date_label)
            elif period_type == 'weekly':
                # Safely handle date operations
                try:
                    end_date = stat.Period + timedelta(days=6)
                    labels.append('{0} - {1}'.format(
                        stat.Period.strftime('%b %d'),
                        end_date.strftime('%b %d')
                    ))
                except:
                    # Fallback if date operations fail
                    labels.append(str(stat.Period))
            else:  # monthly
                labels.append(str(stat.Period))
                
            activity_counts.append(stat.ActivityCount)
            user_counts.append(stat.UserCount)
        
        # Reverse the lists to show oldest first
        labels.reverse()
        activity_counts.reverse()
        user_counts.reverse()
        
        # Convert to JSON strings for JavaScript
        labels_json = json.dumps(labels)
        activity_json = json.dumps(activity_counts)
        users_json = json.dumps(user_counts)
        
        # Add chart
        html = """
        <div class="panel panel-default">
            <div class="panel-heading">
                <h3 class="panel-title">Activity Trends</h3>
            </div>
            <div class="panel-body">
                <canvas id="activityChart" width="800" height="400"></canvas>
                <script>
                    var ctx = document.getElementById('activityChart').getContext('2d');
                    var chart = new Chart(ctx, {{
                        type: 'line',
                        data: {{
                            labels: {0},
                            datasets: [
                                {{
                                    label: 'Activity Count',
                                    data: {1},
                                    borderColor: 'rgba(54, 162, 235, 1)',
                                    backgroundColor: 'rgba(54, 162, 235, 0.2)',
                                    borderWidth: 2,
                                    yAxisID: 'y-axis-1'
                                }},
                                {{
                                    label: 'Active Users',
                                    data: {2},
                                    borderColor: 'rgba(255, 99, 132, 1)',
                                    backgroundColor: 'rgba(255, 99, 132, 0.2)',
                                    borderWidth: 2,
                                    yAxisID: 'y-axis-2'
                                }}
                            ]
                        }},
                        options: {{
                            responsive: true,
                            hoverMode: 'index',
                            stacked: false,
                            scales: {{
                                yAxes: [
                                    {{
                                        type: 'linear',
                                        display: true,
                                        position: 'left',
                                        id: 'y-axis-1',
                                        scaleLabel: {{
                                            display: true,
                                            labelString: 'Activity Count'
                                        }}
                                    }},
                                    {{
                                        type: 'linear',
                                        display: true,
                                        position: 'right',
                                        id: 'y-axis-2',
                                        gridLines: {{
                                            drawOnChartArea: false
                                        }},
                                        scaleLabel: {{
                                            display: true,
                                            labelString: 'Active Users'
                                        }}
                                    }}
                                ]
                            }}
                        }}
                    }});
                </script>
            </div>
        </div>
        """.format(labels_json, activity_json, users_json)
        
        return html
    
    def render_most_active_users_chart(self, users):
        """Render a chart of most active users."""
        # Prepare data for chart
        labels = []
        activity_counts = []
        
        # Make sure we have users before trying to slice
        user_list = list(users) if users else []
        
        # Use min to avoid index errors if we have fewer than 10 users
        display_count = min(10, len(user_list))
        
        for i in range(display_count):
            user = user_list[i]
            labels.append(user.Name)
            # Ensure we have a proper integer value
            try:
                activity_count = int(user.ActivityCount)
            except (ValueError, TypeError):
                activity_count = 0
            activity_counts.append(activity_count)
        
        # Convert to JSON strings for JavaScript
        labels_json = json.dumps(labels)
        activity_json = json.dumps(activity_counts)
        
        # Add chart
        html = """
        <div class="panel panel-default">
            <div class="panel-heading">
                <h3 class="panel-title">Most Active Users</h3>
            </div>
            <div class="panel-body">
                <canvas id="usersChart" width="800" height="400"></canvas>
                <script>
                    var ctx = document.getElementById('usersChart').getContext('2d');
                    var chart = new Chart(ctx, {{
                        type: 'bar',
                        data: {{
                            labels: {0},
                            datasets: [
                                {{
                                    label: 'Activity Count',
                                    data: {1},
                                    backgroundColor: 'rgba(75, 192, 192, 0.2)',
                                    borderColor: 'rgba(75, 192, 192, 1)',
                                    borderWidth: 1
                                }}
                            ]
                        }},
                        options: {{
                            scales: {{
                                yAxes: [{{
                                    ticks: {{
                                        beginAtZero: true
                                    }}
                                }}]
                            }}
                        }}
                    }});
                </script>
            </div>
        </div>
        """.format(labels_json, activity_json)
        
        return html


# ==========================================
# Page Rendering Functions
# ==========================================

def render_overview_page(form_handler, analyzer, renderer):
    """Render the overview page."""
    try:
        days = form_handler.get_int_param('days', 30)
        

        
        # Get summary metrics
        recent_activity = analyzer.get_activity_stats_by_period(days, 'daily')
        most_active_users = analyzer.get_most_active_users(days)
        locked_accounts = analyzer.get_locked_accounts()
        stale_accounts = analyzer.get_stale_accounts()
        
        # Calculate summary statistics - make these calculations safer
        total_activity = 0
        active_users = 0
        
        # Handle recent_activity safely
        if recent_activity:
            for stat in recent_activity:
                try:
                    total_activity += int(stat.ActivityCount)
                except (ValueError, TypeError, AttributeError):
                    pass  # Skip if we can't parse the activity count
        
        # Handle most_active_users safely
        if most_active_users:
            # Use a set to collect unique user IDs
            unique_users = set()
            for user in most_active_users:
                try:
                    unique_users.add(user.UserId)
                except (AttributeError, TypeError):
                    pass  # Skip if UserId doesn't exist or isn't hashable
            # Calculate active users 
            active_users = len(list(most_active_users)) if most_active_users else 0
        
        html = renderer.render_page_header("User Activity Analysis", "Overview Dashboard")
        html += renderer.render_navigation('overview')
        html += renderer.render_filter_form(days)
        
        # Summary metrics
        html += """
        <div class="row">
            <div class="col-md-3">
                <div class="panel panel-primary">
                    <div class="panel-heading text-center">
                        <h3>{0:,}</h3>
                        <p>Total Activities</p>
                    </div>
                </div>
            </div>
            <div class="col-md-3">
                <div class="panel panel-info">
                    <div class="panel-heading text-center">
                        <h3>{1}</h3>
                        <p>Active Users</p>
                    </div>
                </div>
            </div>
            <div class="col-md-3">
                <div class="panel panel-warning">
                    <div class="panel-heading text-center">
                        <h3>{2}</h3>
                        <p>Locked Accounts</p>
                    </div>
                </div>
            </div>
            <div class="col-md-3">
                <div class="panel panel-danger">
                    <div class="panel-heading text-center">
                        <h3>{3}</h3>
                        <p>Stale Accounts</p>
                    </div>
                </div>
            </div>
        </div>
        """.format(
            total_activity,
            active_users,
            len(locked_accounts) if locked_accounts else 0,
            len(stale_accounts) if stale_accounts else 0
        )
        
        # Add activity chart - implemented directly instead of calling render_activity_stats_chart
        # Prepare data for chart
        labels = []
        activity_counts = []
        user_counts = []
        
        for stat in recent_activity:
            # Format date for label
            if hasattr(stat.Period, 'strftime'):
                date_label = stat.Period.strftime('%b %d')
            else:
                date_label = str(stat.Period)
            labels.append(date_label)
            
            activity_counts.append(stat.ActivityCount)
            user_counts.append(stat.UserCount)
        
        # Reverse the lists to show oldest first
        labels.reverse()
        activity_counts.reverse()
        user_counts.reverse()
        
        # Convert to JSON strings for JavaScript
        labels_json = json.dumps(labels)
        activity_json = json.dumps(activity_counts)
        users_json = json.dumps(user_counts)
        
        html += """
        <div class="panel panel-default">
            <div class="panel-heading">
                <h3 class="panel-title">Activity Trends</h3>
            </div>
            <div class="panel-body">
                <canvas id="activityChart" width="800" height="400"></canvas>
                <script>
                    var ctx = document.getElementById('activityChart').getContext('2d');
                    var chart = new Chart(ctx, {
                        type: 'line',
                        data: {
                            labels: %s,
                            datasets: [
                                {
                                    label: 'Activity Count',
                                    data: %s,
                                    borderColor: 'rgba(54, 162, 235, 1)',
                                    backgroundColor: 'rgba(54, 162, 235, 0.2)',
                                    borderWidth: 2,
                                    yAxisID: 'y-axis-1'
                                },
                                {
                                    label: 'Active Users',
                                    data: %s,
                                    borderColor: 'rgba(255, 99, 132, 1)',
                                    backgroundColor: 'rgba(255, 99, 132, 0.2)',
                                    borderWidth: 2,
                                    yAxisID: 'y-axis-2'
                                }
                            ]
                        },
                        options: {
                            responsive: true,
                            hoverMode: 'index',
                            stacked: false,
                            scales: {
                                yAxes: [
                                    {
                                        type: 'linear',
                                        display: true,
                                        position: 'left',
                                        id: 'y-axis-1',
                                        scaleLabel: {
                                            display: true,
                                            labelString: 'Activity Count'
                                        }
                                    },
                                    {
                                        type: 'linear',
                                        display: true,
                                        position: 'right',
                                        id: 'y-axis-2',
                                        gridLines: {
                                            drawOnChartArea: false
                                        },
                                        scaleLabel: {
                                            display: true,
                                            labelString: 'Active Users'
                                        }
                                    }
                                ]
                            }
                        }
                    });
                </script>
            </div>
        </div>
        """ % (labels_json, activity_json, users_json)
        
        # Add most active users chart - implemented directly instead of calling render_most_active_users_chart
        # Prepare data for chart
        user_labels = []
        user_activity_counts = []
        
        # Make sure we have users before trying to slice
        user_list = list(most_active_users) if most_active_users else []
        
        # Use min to avoid index errors if we have fewer than 10 users
        display_count = min(10, len(user_list))
        
        for i in range(display_count):
            user = user_list[i]
            user_labels.append(user.Name)
            # Ensure we have a proper integer value
            try:
                activity_count = int(user.ActivityCount)
            except (ValueError, TypeError):
                activity_count = 0
            user_activity_counts.append(activity_count)
        
        # Convert to JSON strings for JavaScript
        user_labels_json = json.dumps(user_labels)
        user_activity_json = json.dumps(user_activity_counts)
        
        html += """
        <div class="panel panel-default">
            <div class="panel-heading">
                <h3 class="panel-title">Most Active Users</h3>
            </div>
            <div class="panel-body">
                <canvas id="usersChart" width="800" height="400"></canvas>
                <script>
                    var ctx = document.getElementById('usersChart').getContext('2d');
                    var chart = new Chart(ctx, {
                        type: 'bar',
                        data: {
                            labels: %s,
                            datasets: [
                                {
                                    label: 'Activity Count',
                                    data: %s,
                                    backgroundColor: 'rgba(75, 192, 192, 0.2)',
                                    borderColor: 'rgba(75, 192, 192, 1)',
                                    borderWidth: 1
                                }
                            ]
                        },
                        options: {
                            scales: {
                                yAxes: [{
                                    ticks: {
                                        beginAtZero: true
                                    }
                                }]
                            }
                        }
                    });
                </script>
            </div>
        </div>
        """ % (user_labels_json, user_activity_json)
        
        # Top users table with safer handling
        html += """
        <div class="panel panel-default">
            <div class="panel-heading">
                <h3 class="panel-title">Most Active Users (Last {0} Days)</h3>
            </div>
            <div class="panel-body">
                <table class="table table-striped table-hover">
                    <thead>
                        <tr>
                            <th>Name</th>
                            <th>Username</th>
                            <th>Activities</th>
                            <th>Days Active</th>
                            <th>First Activity</th>
                            <th>Last Activity</th>
                            <th>Actions</th>
                        </tr>
                    </thead>
                    <tbody>
        """.format(days)
        
        # Safely handle most_active_users display
        users_to_display = []
        if most_active_users:
            users_list = list(most_active_users)
            display_count = min(10, len(users_list))
            users_to_display = users_list[:display_count]
        
        for user in users_to_display:
            html += """
            <tr>
                <td>{0}</td>
                <td>{1}</td>
                <td>{2:,}</td>
                <td>{3}</td>
                <td>{4}</td>
                <td>{5}</td>
                <td>
                    <a href="?view=user_detail&user_id={6}" class="btn btn-xs btn-primary">View Activity</a>
                </td>
            </tr>
            """.format(
                getattr(user, 'Name', 'Unknown'),
                getattr(user, 'Username', 'Unknown'),
                getattr(user, 'ActivityCount', 0),
                getattr(user, 'DaysActive', 0),
                renderer.render_date(getattr(user, 'FirstActivity', None)),
                renderer.render_date(getattr(user, 'LastActivity', None)),
                getattr(user, 'UserId', 0)
            )
        
        html += """
                    </tbody>
                </table>
                <a href="?view=user_list" class="btn btn-default">View All Users</a>
            </div>
        </div>
        """
        
        return html
    except Exception as e:
        import traceback
        return """
        <div class='alert alert-danger'>
            Error rendering overview: {0}
            <br>
            <pre>{1}</pre>
        </div>
        """.format(str(e), traceback.format_exc())

def render_user_list_page(form_handler, analyzer, renderer):
    """Render the user list page."""
    try:
        search_name = form_handler.get_param('search_name', '')
        
        html = renderer.render_page_header("User Activity Analysis", "User List")
        html += renderer.render_navigation('user_list')
        
        # Create search form
        search_form = """
        <form method="get" class="form-inline" style="margin-bottom: 20px;">
            <input type="hidden" name="view" value="user_list">
            <div class="form-group">
                <label for="search_name">Search Users:</label>
                <input type="text" name="search_name" id="search_name" class="form-control" 
                       placeholder="Name or Username" style="width: 250px;" value="{0}">
            </div>
            <button type="submit" class="btn btn-primary" style="margin-left: 10px;">
                <i class="fa fa-search"></i> Search
            </button>
        </form>
        """.format(search_name)
        
        html += search_form
        
        # Add loading indicator for searches
        html += """
        <div id="loading" style="display: none;">
            <p><i class="fa fa-spinner fa-spin"></i> Searching users...</p>
        </div>
        
        <script>
            document.addEventListener('DOMContentLoaded', function() {
                var form = document.querySelector('form');
                var loading = document.getElementById('loading');
                
                form.addEventListener('submit', function() {
                    loading.style.display = 'block';
                });
            });
        </script>
        """
        
        # Get users - using direct SQL query with search term if provided
        try:
            if search_name:
                # Use SQL LIKE for case-insensitive search
                sql = """
                    SELECT TOP 100 u.UserId, u.PeopleId, u.Username, u.Name, u.LastLoginDate, 
                        u.LastActivityDate, p.EmailAddress
                    FROM Users u
                    JOIN People p ON p.PeopleId = u.PeopleId
                    WHERE u.Name LIKE '%{0}%' OR u.Username LIKE '%{0}%'
                    ORDER BY u.LastActivityDate DESC
                """.format(search_name.replace("'", "''"))  # Escape single quotes for SQL
                
                users = analyzer.query.QuerySql(sql)
                
                if users and len(list(users)) > 0:
                    html += "<div class='alert alert-info'>Showing users matching '{0}'</div>".format(search_name)
                else:
                    # Try a secondary search if no results found
                    sql = """
                        SELECT TOP 100 u.UserId, u.PeopleId, u.Username, u.Name, u.LastLoginDate, 
                            u.LastActivityDate, p.EmailAddress
                        FROM Users u
                        JOIN People p ON p.PeopleId = u.PeopleId
                        WHERE p.FirstName LIKE '%{0}%' OR p.LastName LIKE '%{0}%' OR p.EmailAddress LIKE '%{0}%'
                        ORDER BY u.LastActivityDate DESC
                    """.format(search_name.replace("'", "''"))
                    
                    users = analyzer.query.QuerySql(sql)
                    
                    if users and len(list(users)) > 0:
                        html += "<div class='alert alert-info'>Showing users with matching People records for '{0}'</div>".format(search_name)
                    else:
                        html += "<div class='alert alert-warning'>No users found matching '{0}'</div>".format(search_name)
            else:
                # Default user list
                sql = """
                    SELECT TOP 100 u.UserId, u.PeopleId, u.Username, u.Name, u.LastLoginDate, 
                        u.LastActivityDate, p.EmailAddress
                    FROM Users u
                    JOIN People p ON p.PeopleId = u.PeopleId
                    WHERE u.LastLoginDate IS NOT NULL
                    ORDER BY u.LastActivityDate DESC
                """
                users = analyzer.query.QuerySql(sql)
        except Exception as e:
            html += "<div class='alert alert-danger'>Error retrieving user list: {0}</div>".format(str(e))
            return html
            
        # Render the users table
        if users and len(list(users)) > 0:
            html += """
            <table class="table table-striped table-hover">
                <thead>
                    <tr>
                        <th>ID</th>
                        <th>Username</th>
                        <th>Name</th>
                        <th>Email</th>
                        <th>Last Login</th>
                        <th>Last Activity</th>
                        <th>Actions</th>
                    </tr>
                </thead>
                <tbody>
            """
            
            for user in users:
                try:
                    html += """
                    <tr>
                        <td>{0}</td>
                        <td>{1}</td>
                        <td>{2}</td>
                        <td>{3}</td>
                        <td>{4}</td>
                        <td>{5}</td>
                        <td>
                            <a href="?view=user_detail&user_id={6}" class="btn btn-xs btn-primary">View Activity</a>
                            <a href="/Person2/{7}" target="_blank" class="btn btn-xs btn-info">View Person</a>
                        </td>
                    </tr>
                    """.format(
                        getattr(user, 'UserId', 'N/A'),
                        getattr(user, 'Username', 'N/A'),
                        getattr(user, 'Name', 'N/A'),
                        getattr(user, 'EmailAddress', 'N/A'),
                        renderer.render_date(getattr(user, 'LastLoginDate', None)),
                        renderer.render_date(getattr(user, 'LastActivityDate', None)),
                        getattr(user, 'UserId', 0),
                        getattr(user, 'PeopleId', 0)
                    )
                except Exception as row_error:
                    print "<div style='display:none'>Error rendering user row: {0}</div>".format(str(row_error))
                    
            html += """
                </tbody>
            </table>
            """
        elif not search_name:
            html += "<div class='alert alert-warning'>No users found in the system.</div>"
        
        return html
    except Exception as e:
        import traceback
        return """
        <div class='alert alert-danger'>
            Error rendering user list: {0}
            <br>
            <pre>{1}</pre>
        </div>
        """.format(str(e), traceback.format_exc())

def render_user_detail_page(form_handler, analyzer, renderer):
    """Render the user detail page."""
    try:
        # Get user ID
        user_id_raw = form_handler.get_param('user_id')
        days = form_handler.get_int_param('days', 7)  # Allow customizing time period
        
        try:
            if user_id_raw:
                user_id = int(user_id_raw)
            else:
                user_id = None
        except:
            user_id = None
            
        if not user_id:
            return "<div class='alert alert-danger'>No user ID specified</div>"
        
        user = analyzer.get_user_info(user_id)
        
        if not user:
            return "<div class='alert alert-danger'>User not found</div>"
        
        html = renderer.render_page_header("User Activity Analysis", 
                                         "User Details for {0}".format(getattr(user, 'Name', 'Unknown')))
        html += renderer.render_navigation('user_detail')
        
        # Add time period selector
        html += """
        <form method="get" class="form-inline" style="margin-bottom: 20px;">
            <input type="hidden" name="view" value="user_detail">
            <input type="hidden" name="user_id" value="{0}">
            <div class="form-group">
                <label for="days">Time Period:</label>
                <select name="days" id="days" class="form-control">
        """.format(user_id)
        
        for d in [7, 14, 30, 60, 90]:
            selected = ' selected' if d == days else ''
            html += '<option value="{0}"{1}>Last {0} Days</option>'.format(d, selected)
            
        html += """
                </select>
            </div>
            <button type="submit" class="btn btn-primary" style="margin-left: 10px;">Apply</button>
        </form>
        
        <p>
            <a href="?view=user_list" class="btn btn-default">
                <i class="fa fa-arrow-left"></i> Back to User List
            </a>
        </p>
        """
        
        # Get all activities for the specified time period
        sql = """
            SELECT ActivityDate, Activity, PageUrl, OrgId, PeopleId, Mobile, ClientIp
            FROM ActivityLog
            WHERE UserId = {0}
                AND DATEDIFF(day, ActivityDate, GETDATE()) <= {1}
            ORDER BY ActivityDate DESC
        """.format(user_id, days)
        
        activities = analyzer.query.QuerySql(sql)
        
        # Convert activities to a list and count them
        try:
            activity_list = list(activities) if activities else []
            activity_count = len(activity_list)
            
            # Sort activities by date (newest first)
            activity_list.sort(key=lambda a: a.ActivityDate, reverse=True)
        except Exception as e:
            print "<div style='display:none'>Error processing activities: {0}</div>".format(str(e))
            activity_list = []
            activity_count = 0
            
        # Basic user info
        html += """
        <div class="row">
            <div class="col-md-6">
                <div class="panel panel-default">
                    <div class="panel-heading">
                        <h3 class="panel-title">User Information</h3>
                    </div>
                    <div class="panel-body">
                        <table class="table table-striped">
                            <tr>
                                <th>User ID:</th>
                                <td>{0}</td>
                            </tr>
                            <tr>
                                <th>Username:</th>
                                <td>{1}</td>
                            </tr>
                            <tr>
                                <th>Name:</th>
                                <td>{2}</td>
                            </tr>
                            <tr>
                                <th>Email:</th>
                                <td>{3}</td>
                            </tr>
                            <tr>
                                <th>Person ID:</th>
                                <td><a href="/Person2/{4}" target="_blank">{4}</a></td>
                            </tr>
                            <tr>
                                <th>Account Created:</th>
                                <td>{5}</td>
                            </tr>
                            <tr>
                                <th>Last Login:</th>
                                <td>{6}</td>
                            </tr>
                            <tr>
                                <th>Last Activity:</th>
                                <td>{7}</td>
                            </tr>
                            <tr>
                                <th>Account Status:</th>
                                <td>{8}</td>
                            </tr>
                        </table>
                    </div>
                </div>
            </div>
        """.format(
            user_id,
            getattr(user, 'Username', 'N/A'),
            getattr(user, 'Name', 'N/A'),
            getattr(user, 'EmailAddress', 'N/A'),
            getattr(user, 'PeopleId', 'N/A'),
            renderer.render_date(getattr(user, 'CreationDate', None)),
            renderer.render_date(getattr(user, 'LastLoginDate', None)),
            renderer.render_date(getattr(user, 'LastActivityDate', None)),
            "LOCKED" if getattr(user, 'IsLockedOut', 0) == 1 else "Active"
        )
        
        # Initialize these variables to be calculated properly later
        total_activities = activity_count
        estimated_sessions = 0
        estimated_work_hours = 0
        avg_session_minutes = 0
        
        # NEW SECTION FOR DAILY ACTIVITY - Using SQL for proper day grouping
        # ==================================================================
        # Get daily activities grouped by day using SQL
        try:
            # Get distinct dates with counts to properly group by day only
            sql_by_day = """
                SELECT 
                    CONVERT(date, ActivityDate) AS ActivityDay,
                    COUNT(*) AS ActivityCount
                FROM ActivityLog
                WHERE UserId = {0}
                    AND DATEDIFF(day, ActivityDate, GETDATE()) <= {1}
                GROUP BY CONVERT(date, ActivityDate)
                ORDER BY ActivityDay DESC
            """.format(user_id, days)
            
            daily_counts = analyzer.query.QuerySql(sql_by_day)
            
            # Now get location stats by day using SQL as well
            office_ips = "'" + "','".join(OFFICE_IP_ADDRESSES) + "'"
            sql_locations = """
                SELECT 
                    CONVERT(date, ActivityDate) AS ActivityDay,
                    SUM(CASE WHEN Mobile = 1 THEN 1 ELSE 0 END) AS MobileCount,
                    SUM(CASE WHEN ClientIp IN ({2}) AND Mobile = 0 THEN 1 ELSE 0 END) AS OfficeCount,
                    SUM(CASE WHEN Mobile = 0 AND (ClientIp NOT IN ({2}) OR ClientIp IS NULL) THEN 1 ELSE 0 END) AS RemoteCount
                FROM ActivityLog
                WHERE UserId = {0}
                    AND DATEDIFF(day, ActivityDate, GETDATE()) <= {1}
                GROUP BY CONVERT(date, ActivityDate)
                ORDER BY ActivityDay DESC
            """.format(user_id, days, office_ips)
            
            daily_locations = analyzer.query.QuerySql(sql_locations)
            
            # Combine the results into a single daily_activity structure
            daily_activity = {}
            
            # First process the counts
            for day_data in daily_counts:
                # Format the day as a string, but only keep the date part
                if hasattr(day_data.ActivityDay, 'strftime'):
                    # Use a date-only format MM/DD/YYYY
                    day_str = day_data.ActivityDay.strftime('%m/%d/%Y')
                else:
                    # If we can't use strftime, try to get just the date part
                    date_parts = str(day_data.ActivityDay).split(' ')[0].split('-')
                    if len(date_parts) >= 3:
                        day_str = "{0}/{1}/{2}".format(date_parts[1], date_parts[2], date_parts[0])
                    else:
                        day_str = str(day_data.ActivityDay).split(' ')[0]
                
                daily_activity[day_str] = {
                    'date': day_str,
                    'count': day_data.ActivityCount,
                    'office': 0,
                    'remote': 0,
                    'mobile': 0
                }
            
            # Now add the location data
            for day_data in daily_locations:
                # Format the day consistently with the counts data
                if hasattr(day_data.ActivityDay, 'strftime'):
                    day_str = day_data.ActivityDay.strftime('%m/%d/%Y')
                else:
                    date_parts = str(day_data.ActivityDay).split(' ')[0].split('-')
                    if len(date_parts) >= 3:
                        day_str = "{0}/{1}/{2}".format(date_parts[1], date_parts[2], date_parts[0])
                    else:
                        day_str = str(day_data.ActivityDay).split(' ')[0]
                
                if day_str in daily_activity:
                    daily_activity[day_str]['office'] = day_data.OfficeCount
                    daily_activity[day_str]['remote'] = day_data.RemoteCount
                    daily_activity[day_str]['mobile'] = day_data.MobileCount
        
        except Exception as e:
            print "<div style='display:none'>Error grouping activities by day using SQL: {0}</div>".format(str(e))
            
            # Fallback method if SQL grouping fails
            daily_activity = {}
            
            # Manually process each activity
            for activity in activity_list:
                date_str = ""
                try:
                    # Try to format the date consistently
                    if hasattr(activity.ActivityDate, 'strftime'):
                        date_str = activity.ActivityDate.strftime('%m/%d/%Y')
                    else:
                        date_parts = str(activity.ActivityDate).split(' ')[0].split('-')
                        if len(date_parts) >= 3:
                            date_str = "{0}/{1}/{2}".format(date_parts[1], date_parts[2], date_parts[0])
                        else:
                            date_str = str(activity.ActivityDate).split(' ')[0]
                except:
                    date_str = "Unknown Date"
                    
                # Initialize if not exists
                if date_str not in daily_activity:
                    daily_activity[date_str] = {
                        'date': date_str,
                        'count': 0,
                        'office': 0,
                        'remote': 0,
                        'mobile': 0
                    }
                    
                # Count this activity
                daily_activity[date_str]['count'] += 1
                
                # Count by location
                try:
                    if hasattr(activity, 'Mobile') and activity.Mobile:
                        daily_activity[date_str]['mobile'] += 1
                    elif hasattr(activity, 'ClientIp'):
                        client_ip = str(activity.ClientIp) if activity.ClientIp else ''
                        if client_ip in OFFICE_IP_ADDRESSES:
                            daily_activity[date_str]['office'] += 1
                        else:
                            daily_activity[date_str]['remote'] += 1
                    else:
                        daily_activity[date_str]['remote'] += 1
                except:
                    # Default to remote if there's an error
                    daily_activity[date_str]['remote'] += 1
        
        # ==================================================================
        # Calculate estimated hours from daily activity first
        # This is important as we'll use this as our base for all other calculations
        
        # Process daily activity to calculate hours and totals by location
        office_total = 0
        remote_total = 0
        mobile_total = 0
        daily_estimated_hours = 0.0
        
        for day_key in daily_activity:
            day_data = daily_activity[day_key]
            
            # Count activities by location
            office_total += day_data['office']
            remote_total += day_data['remote']
            mobile_total += day_data['mobile']
            
            # Calculate realistic hours - max 8 hours per day
            # Estimate 2-5 minutes per activity based on count
            if day_data['count'] <= 50:
                mins_per_activity = 5  # More time for each activity when fewer activities
            elif day_data['count'] <= 100:
                mins_per_activity = 4
            elif day_data['count'] <= 200: 
                mins_per_activity = 3
            else:
                mins_per_activity = 2  # Less time for each when many activities
                
            # Calculate estimated hours with a reasonable cap
            est_mins = day_data['count'] * mins_per_activity
            est_hours = min(8.0, est_mins / 60.0)  # Cap at 8 hours per day
            
            # Store the estimated hours in the day data
            day_data['est_hours'] = est_hours
            
            # Add to total estimated hours
            daily_estimated_hours += est_hours
        
        # Use daily estimated hours for all other calculations for consistency
        estimated_work_hours = daily_estimated_hours
        
        # Calculate sessions based on activity count
        if activity_count > 0:
            # Estimate sessions based on activity count and days
            # Assume 1-3 sessions per day with activity
            active_days = len(daily_activity)
            if active_days > 0:
                estimated_sessions = min(active_days * 3, activity_count // 10)
                estimated_sessions = max(active_days, estimated_sessions)  # At least one session per active day
            else:
                estimated_sessions = max(1, activity_count // 20)
            
            # Calculate average session duration
            if estimated_sessions > 0:
                avg_session_minutes = (estimated_work_hours * 60) / estimated_sessions
            else:
                avg_session_minutes = 0
        
        # Activity statistics panel
        # Convert values to appropriate types to avoid formatting issues
        days_str = str(days)
        total_activities_str = str(total_activities)
        estimated_sessions_str = str(int(estimated_sessions))
        
        # Format the float values with 1 decimal place
        try:
            estimated_work_hours_str = "{:.1f}".format(float(estimated_work_hours))
        except (ValueError, TypeError):
            estimated_work_hours_str = "0.0"  # Default if conversion fails
        
        try:
            avg_session_minutes_str = "{:.1f}".format(float(avg_session_minutes))
        except (ValueError, TypeError):
            avg_session_minutes_str = "0.0"  # Default if conversion fails
        
        # Now build the HTML without using .format()
        html += """
            <div class="col-md-6">
                <div class="panel panel-default">
                    <div class="panel-heading">
                        <h3 class="panel-title">Activity Summary (Last """ + days_str + """ Days)</h3>
                    </div>
                    <div class="panel-body">
                        <div class="row">
                            <div class="col-md-6">
                                <div class="panel panel-info">
                                    <div class="panel-heading text-center">
                                        <h4>""" + total_activities_str + """</h4>
                                        <p>Total Activities</p>
                                    </div>
                                </div>
                            </div>
                            <div class="col-md-6">
                                <div class="panel panel-info">
                                    <div class="panel-heading text-center">
                                        <h4>""" + estimated_sessions_str + """</h4>
                                        <p>Estimated Sessions</p>
                                    </div>
                                </div>
                            </div>
                        </div>
                        <div class="row">
                            <div class="col-md-6">
                                <div class="panel panel-success">
                                    <div class="panel-heading text-center">
                                        <h4>""" + estimated_work_hours_str + """ hrs</h4>
                                        <p>Estimated Work Time</p>
                                    </div>
                                </div>
                            </div>
                            <div class="col-md-6">
                                <div class="panel panel-success">
                                    <div class="panel-heading text-center">
                                        <h4>""" + avg_session_minutes_str + """ mins</h4>
                                        <p>Avg Session Duration</p>
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
        """
        
        # Calculate location percentages and hours based on the daily activity totals
        location_total = float(office_total + remote_total + mobile_total)
        
        if location_total > 0:
            office_pct = (office_total / location_total) * 100
            remote_pct = (remote_total / location_total) * 100
            mobile_pct = (mobile_total / location_total) * 100
            
            # Distribute the total hours by location percentage
            office_hours = (estimated_work_hours * office_total) / location_total
            remote_hours = (estimated_work_hours * remote_total) / location_total
            mobile_hours = (estimated_work_hours * mobile_total) / location_total
        else:
            office_pct = remote_pct = mobile_pct = 0
            office_hours = remote_hours = mobile_hours = 0

        # Convert all values to floats and format outside the string
        office_hours_str = "{:.1f}".format(float(office_hours))
        office_pct_str = "{:.1f}".format(float(office_pct))
        remote_hours_str = "{:.1f}".format(float(remote_hours))
        remote_pct_str = "{:.1f}".format(float(remote_pct))
        mobile_hours_str = "{:.1f}".format(float(mobile_hours))
        mobile_pct_str = "{:.1f}".format(float(mobile_pct))

        # Work Location Summary panel
        # Work Location Summary panel with %s placeholders instead of .format()
        work_location_html = """
        <div class="row">
            <div class="col-md-12">
                <div class="panel panel-default">
                    <div class="panel-heading">
                        <h3 class="panel-title">Work Location Summary</h3>
                    </div>
                    <div class="panel-body">
                        <div class="row">
                            <div class="col-md-4">
                                <div class="panel panel-primary">
                                    <div class="panel-heading text-center">
                                        <h4>%s hrs (%s%%)</h4>
                                        <p>Office Work</p>
                                    </div>
                                </div>
                            </div>
                            <div class="col-md-4">
                                <div class="panel panel-warning">
                                    <div class="panel-heading text-center">
                                        <h4>%s hrs (%s%%)</h4>
                                        <p>Remote Work</p>
                                    </div>
                                </div>
                            </div>
                            <div class="col-md-4">
                                <div class="panel panel-info">
                                    <div class="panel-heading text-center">
                                        <h4>%s hrs (%s%%)</h4>
                                        <p>Mobile Work</p>
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
        """ % (
            office_hours_str, office_pct_str,
            remote_hours_str, remote_pct_str,
            mobile_hours_str, mobile_pct_str
        )
        
        html += work_location_html
        
        # Daily Activity Summary - with proper day grouping from the SQL query
        if daily_activity:
            html += """
            <div class="panel panel-default">
                <div class="panel-heading">
                    <h3 class="panel-title">Daily Activity Summary</h3>
                </div>
                <div class="panel-body">
                    <div style="max-height: 400px; overflow-y: auto;">
                        <table class="table table-striped">
                            <thead>
                                <tr>
                                    <th>Date</th>
                                    <th>Total Activities</th>
                                    <th>Office</th>
                                    <th>Remote</th>
                                    <th>Mobile</th>
                                    <th>Estimated Hours</th>
                                </tr>
                            </thead>
                            <tbody>
            """
            
            # Sort days by date (newest first)
            #sorted_days = list(daily_activity.keys())
            #sorted_days.sort(reverse=True)  # Sort in descending order

            # Sort by date components (MM/DD/YYYY format)
            sorted_days = list(daily_activity.keys())
            sorted_days.sort(key=lambda d: tuple(map(int, d.split('/')[0:3])) if len(d.split('/')) >= 3 else (0,0,0), reverse=True)

            for day_key in sorted_days:
                day_data = daily_activity[day_key]
                
                # Calculate percentages
                day_office_pct = 0
                day_remote_pct = 0 
                day_mobile_pct = 0
                
                if day_data['count'] > 0:
                    day_office_pct = (day_data['office'] * 100.0) / day_data['count']
                    day_remote_pct = (day_data['remote'] * 100.0) / day_data['count']
                    day_mobile_pct = (day_data['mobile'] * 100.0) / day_data['count']
                
                # Format percentages separately
                day_office_pct_str = "{:.1f}".format(float(day_office_pct))
                day_remote_pct_str = "{:.1f}".format(float(day_remote_pct))
                day_mobile_pct_str = "{:.1f}".format(float(day_mobile_pct))
                day_est_hours_str = "{:.1f}".format(float(day_data['est_hours']))
                
                # Use string formatting with %s placeholders
                html += """
                <tr>
                    <td>%s</td>
                    <td>%s</td>
                    <td>%s (%s%%)</td>
                    <td>%s (%s%%)</td>
                    <td>%s (%s%%)</td>
                    <td>%s hrs</td>
                </tr>
                """ % (
                    day_data['date'],
                    day_data['count'],
                    day_data['office'], day_office_pct_str,
                    day_data['remote'], day_remote_pct_str,
                    day_data['mobile'], day_mobile_pct_str,
                    day_est_hours_str
                )
                
            html += """
                        </tbody>
                    </table>
                </div>
            </div>
            """
        
        # Calculate activity categories
        activity_categories = {}
        
        for activity in activity_list[:250]:
            try:
                if hasattr(activity, 'Activity'):
                    activity_text = str(activity.Activity)
                    parts = activity_text.split(':', 1)
                    category = parts[0].strip() if parts else "Other"
                    
                    if category in activity_categories:
                        activity_categories[category] += 1
                    else:
                        activity_categories[category] = 1
            except:
                pass  # Skip on error
        
        # Sort categories
        sorted_categories = sorted(activity_categories.items(), key=lambda x: x[1], reverse=True)
        total_categorized = sum(activity_categories.values()) if activity_categories else 0
        
        # Activity Categories panel - MAKING SCROLLABLE
        html += """
        <div class="row">
            <div class="col-md-12">
                <div class="panel panel-default">
                    <div class="panel-heading">
                        <h3 class="panel-title">Activity Categories</h3>
                    </div>
                    <div class="panel-body">
                        <div style="max-height: 400px; overflow-y: auto;">
                            <table class="table table-striped">
                                <thead>
                                    <tr>
                                        <th>Activity Type</th>
                                        <th>Count</th>
                                        <th>Percentage</th>
                                    </tr>
                                </thead>
                                <tbody>
        """
        
        if sorted_categories:
            for category, count in sorted_categories:
                percentage = (count * 100.0 / total_categorized) if total_categorized > 0 else 0
                percentage_str = "{:.1f}".format(float(percentage))
                
                html += """
                <tr>
                    <td>%s</td>
                    <td>%s</td>
                    <td>%s%%</td>
                </tr>
                """ % (category, count, percentage_str)
        else:
            html += """
            <tr>
                <td colspan="3" class="text-center">No activity categories available</td>
            </tr>
            """
            
        html += """
                                </tbody>
                            </table>
                        </div>
                    </div>
                </div>
            </div>
        </div>
        """
        
        # Recent Activity with scrolling
        html += """
        <div class="panel panel-default">
            <div class="panel-heading">
                <h3 class="panel-title">Recent Activity <small>%s total activities</small></h3>
            </div>
            <div class="panel-body">
                <div style="max-height: 600px; overflow-y: auto;">
                    <table class="table table-striped">
                        <thead>
                            <tr>
                                <th>Date & Time</th>
                                <th>Activity Type</th>
                                <th>Details</th>
                                <th>Location</th>
                                <th>Links</th>
                            </tr>
                        </thead>
                        <tbody>
        """ % activity_count
        
        if activity_list:
            # Use all activities instead of limiting to 20
            for activity in activity_list:
                try:
                    # Safely get activity text
                    activity_text = ""
                    if hasattr(activity, 'Activity'):
                        activity_text = str(activity.Activity) if activity.Activity else ""
                    
                    # Split activity by colon to get the main activity type
                    activity_parts = activity_text.split(':', 1)
                    activity_type = activity_parts[0].strip() if len(activity_parts) > 0 else ''
                    activity_details = activity_parts[1].strip() if len(activity_parts) > 1 else activity_text
                    
                    # Determine location
                    location = "Remote"
                    location_badge = "warning"
                    
                    if hasattr(activity, 'Mobile') and activity.Mobile:
                        location = "Mobile"
                        location_badge = "info"
                    elif hasattr(activity, 'ClientIp'):
                        client_ip = str(activity.ClientIp) if activity.ClientIp else ''
                        if client_ip in OFFICE_IP_ADDRESSES:
                            location = "Office"
                            location_badge = "primary"
                    
                    # Format the activity date
                    activity_date = ""
                    if hasattr(activity, 'ActivityDate'):
                        activity_date = renderer.render_date(activity.ActivityDate)
                    
                    # Limit details length
                    details_text = activity_details[:100]
                    if len(activity_details) > 100:
                        details_text += '...'
                    
                    # Prepare HTML for this activity row
                    activity_row = """
                    <tr>
                        <td>%s</td>
                        <td><strong>%s</strong></td>
                        <td>%s</td>
                        <td><span class="label label-%s">%s</span></td>
                        <td>
                    """ % (
                        activity_date,
                        activity_type,
                        details_text,
                        location_badge,
                        location
                    )
                    
                    # Add links for related data
                    links = []
                    
                    # Add organization link if available
                    org_id = None
                    if hasattr(activity, 'OrgId'):
                        if not isinstance(activity.OrgId, slice):
                            org_id = activity.OrgId
                    
                    if org_id:
                        links.append('<a href="/Org/%s" target="_blank" class="btn btn-xs btn-info">Org</a>' % org_id)
                    
                    # Add person link if available
                    people_id = None
                    if hasattr(activity, 'PeopleId'):
                        if not isinstance(activity.PeopleId, slice):
                            people_id = activity.PeopleId
                    
                    if people_id:
                        links.append('<a href="/Person2/%s" target="_blank" class="btn btn-xs btn-primary">Person</a>' % people_id)
                    
                    # Add URL link if available
                    page_url = None
                    if hasattr(activity, 'PageUrl') and activity.PageUrl:
                        page_url = str(activity.PageUrl)
                    
                    if page_url:
                        links.append('<a href="%s" target="_blank" class="btn btn-xs btn-default">URL</a>' % page_url)
                    
                    # Add links to the row
                    activity_row += ' '.join(links)
                    
                    # Complete the row
                    activity_row += """
                        </td>
                    </tr>
                    """
                    
                    html += activity_row
                except Exception as e:
                    print "<div style='display:none'>Error rendering activity: {0}</div>".format(str(e))
        else:
            html += """
            <tr>
                <td colspan="5" class="text-center">No recent activities found</td>
            </tr>
            """
            
        html += """
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
        """
        
        return html
    except Exception as e:
        import traceback
        error_trace = traceback.format_exc()
        return """
        <div class='alert alert-danger'>
            Error rendering user details: {0}
            <br>
            <pre>{1}</pre>
        </div>
        """.format(str(e), error_trace)

def render_stale_accounts_page(form_handler, analyzer, renderer):
    """Render the stale accounts page."""
    try:
        days = form_handler.get_int_param('days', STALE_ACCOUNT_DAYS)
        
        html = renderer.render_page_header("User Activity Analysis", "Stale Accounts")
        html += renderer.render_navigation('stale_accounts')
        
        # Add filter form
        html += """
        <form method="get" class="form-inline" style="margin-bottom: 20px;">
            <input type="hidden" name="view" value="stale_accounts">
            <div class="form-group">
                <label for="days">Inactive For:</label>
                <select name="days" id="days" class="form-control">
        """
        
        for d in [30, 60, 90, 180, 365]:
            selected = ' selected' if d == days else ''
            html += '<option value="{0}"{1}>{0} Days</option>'.format(d, selected)
            
        html += """
                </select>
            </div>
            <button type="submit" class="btn btn-primary" style="margin-left: 10px;">Apply Filter</button>
        </form>
        """
        
        accounts = analyzer.get_stale_accounts(days)
        
        if not accounts:
            html += "<div class='alert alert-info'>No stale accounts found</div>"
        else:
            html += "<p>Showing {0} accounts inactive for at least {1} days</p>".format(len(accounts), days)
            
            html += """
            <table class="table table-striped table-hover">
                <thead>
                    <tr>
                        <th>Username</th>
                        <th>Name</th>
                        <th>Email</th>
                        <th>Account Created</th>
                        <th>Last Login</th>
                        <th>Last Activity</th>
                        <th>Days Inactive</th>
                        <th>Actions</th>
                    </tr>
                </thead>
                <tbody>
            """
            
            for account in accounts:
                try:
                    days_inactive = (datetime.datetime.now() - account.LastActivityDate).days
                except:
                    days_inactive = "Unknown"
                
                html += """
                <tr>
                    <td>{0}</td>
                    <td>{1}</td>
                    <td>{2}</td>
                    <td>{3}</td>
                    <td>{4}</td>
                    <td>{5}</td>
                    <td>{6}</td>
                    <td>
                        <a href="?view=user_detail&user_id={7}" class="btn btn-xs btn-primary">View Activity</a>
                        <a href="/Person2/{8}" target="_blank" class="btn btn-xs btn-info">View Person</a>
                    </td>
                </tr>
                """.format(
                    getattr(account, 'Username', 'Unknown'),
                    getattr(account, 'Name', 'Unknown'),
                    getattr(account, 'EmailAddress', 'Unknown'),
                    renderer.render_date(getattr(account, 'CreationDate', None)),
                    renderer.render_date(getattr(account, 'LastLoginDate', None)),
                    renderer.render_date(getattr(account, 'LastActivityDate', None)),
                    days_inactive,
                    getattr(account, 'UserId', 0),
                    getattr(account, 'PeopleId', 0)
                )
                
            html += """
                </tbody>
            </table>
            """
            
            for account in accounts:
                try:
                    days_inactive = (datetime.datetime.now() - account.LastActivityDate).days
                except:
                    days_inactive = "Unknown"
                
                html += """
                <tr>
                    <td>{0}</td>
                    <td>{1}</td>
                    <td>{2}</td>
                    <td>{3}</td>
                    <td>{4}</td>
                    <td>{5}</td>
                    <td>{6}</td>
                    <td>
                        <a href="?view=user_detail&user_id={7}" class="btn btn-xs btn-primary">View Activity</a>
                        <a href="/Person2/{8}" target="_blank" class="btn btn-xs btn-info">View Person</a>
                    </td>
                </tr>
                """.format(
                    getattr(account, 'Username', 'Unknown'),
                    getattr(account, 'Name', 'Unknown'),
                    getattr(account, 'EmailAddress', 'Unknown'),
                    renderer.render_date(getattr(account, 'CreationDate', None)),
                    renderer.render_date(getattr(account, 'LastLoginDate', None)),
                    renderer.render_date(getattr(account, 'LastActivityDate', None)),
                    days_inactive,
                    getattr(account, 'UserId', 0),
                    getattr(account, 'PeopleId', 0)
                )
                
            html += """
                </tbody>
            </table>
            """
            
            for account in accounts:
                try:
                    days_inactive = (datetime.datetime.now() - account.LastActivityDate).days
                except:
                    days_inactive = "Unknown"
                
                html += """
                <tr>
                    <td>{0}</td>
                    <td>{1}</td>
                    <td>{2}</td>
                    <td>{3}</td>
                    <td>{4}</td>
                    <td>{5}</td>
                    <td>{6}</td>
                    <td>
                        <a href="?view=user_detail&user_id={7}" class="btn btn-xs btn-primary">View Activity</a>
                        <a href="/Person2/{8}" target="_blank" class="btn btn-xs btn-info">View Person</a>
                    </td>
                </tr>
                """.format(
                    getattr(account, 'Username', 'Unknown'),
                    getattr(account, 'Name', 'Unknown'),
                    getattr(account, 'EmailAddress', 'Unknown'),
                    renderer.render_date(getattr(account, 'CreationDate', None)),
                    renderer.render_date(getattr(account, 'LastLoginDate', None)),
                    renderer.render_date(getattr(account, 'LastActivityDate', None)),
                    days_inactive,
                    getattr(account, 'UserId', 0),
                    getattr(account, 'PeopleId', 0)
                )
                
            html += """
                </tbody>
            </table>
            """
        
        return html
    except Exception as e:
        return "<div class='alert alert-danger'>Error rendering stale accounts: {0}</div>".format(str(e))

def render_locked_accounts_page(analyzer, renderer):
    """Render the locked accounts page."""
    try:
        html = renderer.render_page_header("User Activity Analysis", "Locked Accounts")
        html += renderer.render_navigation('locked_accounts')
        
        accounts = analyzer.get_locked_accounts()
        
        if not accounts:
            html += "<div class='alert alert-info'>No locked accounts found</div>"
        else:
            html += "<p>Showing {0} locked accounts</p>".format(len(accounts))
            html += """
            <table class="table table-striped table-hover">
                <thead>
                    <tr>
                        <th>Username</th>
                        <th>Name</th>
                        <th>Email</th>
                        <th>Last Login</th>
                        <th>Locked Since</th>
                        <th>Failed Attempts</th>
                        <th>Actions</th>
                    </tr>
                </thead>
                <tbody>
            """
            
            for account in accounts:
                html += """
                <tr>
                    <td>{0}</td>
                    <td>{1}</td>
                    <td>{2}</td>
                    <td>{3}</td>
                    <td>{4}</td>
                    <td>{5}</td>
                    <td>
                        <a href="?view=user_detail&user_id={6}" class="btn btn-xs btn-primary">View Activity</a>
                        <a href="/Person2/{7}" target="_blank" class="btn btn-xs btn-info">View Person</a>
                    </td>
                </tr>
                """.format(
                    getattr(account, 'Username', 'Unknown'),
                    getattr(account, 'Name', 'Unknown'),
                    getattr(account, 'EmailAddress', 'N/A'),
                    renderer.render_date(getattr(account, 'LastLoginDate', None)),
                    renderer.render_date(getattr(account, 'LastLockedOutDate', None)),
                    getattr(account, 'FailedPasswordAttemptCount', 0),
                    getattr(account, 'UserId', 0),
                    getattr(account, 'PeopleId', 0)
                )
                
            html += """
                </tbody>
            </table>
            """
        
        return html
    except Exception as e:
        return "<div class='alert alert-danger'>Error rendering locked accounts: {0}</div>".format(str(e))

def render_password_resets_page(form_handler, analyzer, renderer):
    """Render the password resets page."""
    try:
        days = form_handler.get_int_param('days', 7)
        
        html = renderer.render_page_header("User Activity Analysis", "Recent Password Resets")
        html += renderer.render_navigation('password_resets')
        
        # Add filter form
        html += """
        <form method="get" class="form-inline" style="margin-bottom: 20px;">
            <input type="hidden" name="view" value="password_resets">
            <div class="form-group">
                <label for="days">Time Period:</label>
                <select name="days" id="days" class="form-control">
        """
        
        for d in [1, 3, 7, 14, 30]:
            selected = ' selected' if d == days else ''
            html += '<option value="{0}"{1}>Last {0} Days</option>'.format(d, selected)
            
        html += """
                </select>
            </div>
            <button type="submit" class="btn btn-primary" style="margin-left: 10px;">Apply Filter</button>
        </form>
        """
        
        resets = analyzer.get_recent_password_resets(days)
        
        if not resets:
            html += "<div class='alert alert-info'>No password resets found in the last {0} days</div>".format(days)
        else:
            html += "<p>Showing {0} password resets in the last {1} days</p>".format(len(resets), days)
            
            html += """
            <table class="table table-striped table-hover">
                <thead>
                    <tr>
                        <th>Username</th>
                        <th>Name</th>
                        <th>Email</th>
                        <th>Password Changed</th>
                        <th>Actions</th>
                    </tr>
                </thead>
                <tbody>
            """
            
            for reset in resets:
                html += """
                <tr>
                    <td>{0}</td>
                    <td>{1}</td>
                    <td>{2}</td>
                    <td>{3}</td>
                    <td>
                        <a href="/PyScript/UserActivityAnalysis?view=user_detail&user_id={4}" class="btn btn-xs btn-primary">View Activity</a>
                        <a href="/Person2/{5}" target="_blank" class="btn btn-xs btn-info">View Person</a>
                    </td>
                </tr>
                """.format(
                    getattr(reset, 'Username', 'Unknown'),
                    getattr(reset, 'Name', 'Unknown'),
                    getattr(reset, 'EmailAddress', 'Unknown'),
                    renderer.render_date(getattr(reset, 'LastPasswordChangedDate', None)),
                    getattr(reset, 'UserId', 0),
                    getattr(reset, 'PeopleId', 0)
                )
                
            html += """
                </tbody>
            </table>
            """
        
        return html
    except Exception as e:
        return "<div class='alert alert-danger'>Error rendering password resets: {0}</div>".format(str(e))

def render_activity_trends_page(form_handler, analyzer, renderer):
    """Render the activity trends page."""
    try:
        days = form_handler.get_int_param('days', 30)
        period_type = form_handler.get_param('period_type', 'daily')
        
        html = renderer.render_page_header("User Activity Analysis", "Activity Trends")
        html += renderer.render_navigation('activity_trends')
        html += renderer.render_filter_form(days, period_type)
        
        # Get activity stats
        stats = analyzer.get_activity_stats_by_period(days, period_type)
        
        # Prepare data for chart
        labels = []
        activity_counts = []
        user_counts = []
        
        for stat in stats:
            # Format the period label based on the period type
            if period_type == 'daily':
                # Handle datetime object in a way that works with Python 2.7
                if hasattr(stat.Period, 'strftime'):
                    date_label = stat.Period.strftime('%b %d')
                else:
                    # Handle string date if that's what we got
                    date_label = str(stat.Period)
                labels.append(date_label)
            elif period_type == 'weekly':
                # Safely handle date operations
                try:
                    end_date = stat.Period + timedelta(days=6)
                    labels.append('{0} - {1}'.format(
                        stat.Period.strftime('%b %d'),
                        end_date.strftime('%b %d')
                    ))
                except:
                    # Fallback if date operations fail
                    labels.append(str(stat.Period))
            else:  # monthly
                labels.append(str(stat.Period))
                
            activity_counts.append(stat.ActivityCount)
            user_counts.append(stat.UserCount)
        
        # Reverse the lists to show oldest first
        labels.reverse()
        activity_counts.reverse()
        user_counts.reverse()
        
        # Convert to JSON strings for JavaScript
        labels_json = json.dumps(labels)
        activity_json = json.dumps(activity_counts)
        users_json = json.dumps(user_counts)
        
        # Add chart directly instead of using renderer.render_activity_stats_chart
        html += """
        <div class="panel panel-default">
            <div class="panel-heading">
                <h3 class="panel-title">Activity Trends</h3>
            </div>
            <div class="panel-body">
                <canvas id="activityChart" width="800" height="400"></canvas>
                <script>
                    var ctx = document.getElementById('activityChart').getContext('2d');
                    var chart = new Chart(ctx, {
                        type: 'line',
                        data: {
                            labels: %s,
                            datasets: [
                                {
                                    label: 'Activity Count',
                                    data: %s,
                                    borderColor: 'rgba(54, 162, 235, 1)',
                                    backgroundColor: 'rgba(54, 162, 235, 0.2)',
                                    borderWidth: 2,
                                    yAxisID: 'y-axis-1'
                                },
                                {
                                    label: 'Active Users',
                                    data: %s,
                                    borderColor: 'rgba(255, 99, 132, 1)',
                                    backgroundColor: 'rgba(255, 99, 132, 0.2)',
                                    borderWidth: 2,
                                    yAxisID: 'y-axis-2'
                                }
                            ]
                        },
                        options: {
                            responsive: true,
                            hoverMode: 'index',
                            stacked: false,
                            scales: {
                                yAxes: [
                                    {
                                        type: 'linear',
                                        display: true,
                                        position: 'left',
                                        id: 'y-axis-1',
                                        scaleLabel: {
                                            display: true,
                                            labelString: 'Activity Count'
                                        }
                                    },
                                    {
                                        type: 'linear',
                                        display: true,
                                        position: 'right',
                                        id: 'y-axis-2',
                                        gridLines: {
                                            drawOnChartArea: false
                                        },
                                        scaleLabel: {
                                            display: true,
                                            labelString: 'Active Users'
                                        }
                                    }
                                ]
                            }
                        }
                    });
                </script>
            </div>
        </div>
        """ % (labels_json, activity_json, users_json)
        
        # Get most active users
        users = analyzer.get_most_active_users(days)
        
        # Render users chart directly instead of using renderer.render_most_active_users_chart
        # Prepare data for chart
        user_labels = []
        user_activity_counts = []
        
        # Make sure we have users before trying to slice
        user_list = list(users) if users else []
        
        # Use min to avoid index errors if we have fewer than 10 users
        display_count = min(10, len(user_list))
        
        for i in range(display_count):
            user = user_list[i]
            user_labels.append(user.Name)
            # Ensure we have a proper integer value
            try:
                activity_count = int(user.ActivityCount)
            except (ValueError, TypeError):
                activity_count = 0
            user_activity_counts.append(activity_count)
        
        # Convert to JSON strings for JavaScript
        user_labels_json = json.dumps(user_labels)
        user_activity_json = json.dumps(user_activity_counts)
        
        # Add chart
        html += """
        <div class="panel panel-default">
            <div class="panel-heading">
                <h3 class="panel-title">Most Active Users</h3>
            </div>
            <div class="panel-body">
                <canvas id="usersChart" width="800" height="400"></canvas>
                <script>
                    var ctx = document.getElementById('usersChart').getContext('2d');
                    var chart = new Chart(ctx, {
                        type: 'bar',
                        data: {
                            labels: %s,
                            datasets: [
                                {
                                    label: 'Activity Count',
                                    data: %s,
                                    backgroundColor: 'rgba(75, 192, 192, 0.2)',
                                    borderColor: 'rgba(75, 192, 192, 1)',
                                    borderWidth: 1
                                }
                            ]
                        },
                        options: {
                            scales: {
                                yAxes: [{
                                    ticks: {
                                        beginAtZero: true
                                    }
                                }]
                            }
                        }
                    });
                </script>
            </div>
        </div>
        """ % (user_labels_json, user_activity_json)
        
        # Add table of most active users
        html += """
        <div class="panel panel-default">
            <div class="panel-heading">
                <h3 class="panel-title">Most Active Users (Last {0} Days)</h3>
            </div>
            <div class="panel-body">
                <table class="table table-striped table-hover">
                    <thead>
                        <tr>
                            <th>Name</th>
                            <th>Username</th>
                            <th>Activities</th>
                            <th>Days Active</th>
                            <th>First Activity</th>
                            <th>Last Activity</th>
                            <th>Actions</th>
                        </tr>
                    </thead>
                    <tbody>
        """.format(days)
        
        # Safely handle users list
        users_list = list(users) if users else []
        
        for user in users_list:
            html += """
            <tr>
                <td>{0}</td>
                <td>{1}</td>
                <td>{2:,}</td>
                <td>{3}</td>
                <td>{4}</td>
                <td>{5}</td>
                <td>
                    <a href="?view=user_detail&user_id={6}" class="btn btn-xs btn-primary">View Activity</a>
                </td>
            </tr>
            """.format(
                getattr(user, 'Name', 'Unknown'),
                getattr(user, 'Username', 'Unknown'),
                getattr(user, 'ActivityCount', 0),
                getattr(user, 'DaysActive', 0),
                renderer.render_date(getattr(user, 'FirstActivity', None)),
                renderer.render_date(getattr(user, 'LastActivity', None)),
                getattr(user, 'UserId', 0)
            )
        
        html += """
                    </tbody>
                </table>
            </div>
        </div>
        """
        
        return html
    except Exception as e:
        import traceback
        return """
        <div class='alert alert-danger'>
            Error rendering activity trends: {0}
            <br>
            <pre>{1}</pre>
        </div>
        """.format(str(e), traceback.format_exc())

def main():
    """Main function to run the application."""
    try:
        # Output debug info about the script name and URL
        print """
        <div style="display:none;" id="debug-panel">
            <h3>Debug Information</h3>
            <p>Script execution started</p>
            <pre>URL: {0}</pre>
            <pre>Script Name: {1}</pre>
        </div>
        <button onclick="document.getElementById('debug-panel').style.display='block';" 
                class="btn btn-xs btn-danger" style="position:fixed; bottom:10px; right:10px;">Debug</button>
        """.format(
            model.Request.Url if hasattr(model, 'Request') and hasattr(model.Request, 'Url') else "Unknown",
            model.Request.Url.Segments[-1] if hasattr(model, 'Request') and hasattr(model.Request, 'Url') and hasattr(model.Request.Url, 'Segments') else "Unknown"
        )
        
        # Initialize helpers
        form_handler = FormHandler(model)
        analyzer = ActivityAnalyzer(model, q)
        renderer = ReportRenderer(model)
        
        # Debug output to see what's happening with form data
        print "<div style='display:none' id='form-data-debug'>"
        print "<h3>Form Data Debug</h3>"
        print "<p>Raw form data received:</p><ul>"
        for attr_name in dir(model.Data):
            if not attr_name.startswith('__'):
                try:
                    attr_value = getattr(model.Data, attr_name)
                    print "<li>{0}: {1}</li>".format(attr_name, attr_value)
                except:
                    print "<li>{0}: [Error accessing value]</li>".format(attr_name)
        print "</ul>"
        
        # Check for URL query string parameters
        if hasattr(model, 'Request') and hasattr(model.Request, 'QueryString'):
            print "<p>URL Query Parameters:</p><ul>"
            for key in model.Request.QueryString.Keys:
                try:
                    print "<li>{0}: {1}</li>".format(key, model.Request.QueryString[key])
                except:
                    print "<li>{0}: [Error accessing value]</li>".format(key)
            print "</ul>"
        print "</div>"
        
        # Get view parameter directly from URL if it exists
        view = None
        if hasattr(model.Data, 'view'):
            view = str(model.Data.view)
        
        # Fallback to check for view in URL if model.Data.view is not set
        if not view and hasattr(model, 'Request') and hasattr(model.Request, 'QueryString'):
            qs = model.Request.QueryString
            if qs and 'view' in qs:
                view = qs['view']
        
        # If still no view, default to overview
        if not view:
            view = 'overview'
            
        print "<div style='display:none' id='view-debug'>"
        print "<p>Selected view: <strong>{0}</strong></p>".format(view)
        print "</div>"
        
        # Include required CSS and JS in header
        print """
        <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/font-awesome/4.7.0/css/font-awesome.min.css">
        <script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/2.7.3/Chart.min.js"></script>
        <style>
            .panel-heading h3, .panel-heading h4 {
                margin-top: 5px;
                margin-bottom: 5px;
            }
            .text-center {
                text-align: center;
            }
        </style>
        
        <div id="loading" style="position: fixed; top: 0; left: 0; width: 100%; height: 100%; 
             background-color: rgba(255,255,255,0.8); z-index: 9999; display: flex; 
             justify-content: center; align-items: center;">
            <div style="text-align: center;">
                <div class="spinner-border" role="status">
                    <i class="fa fa-spinner fa-spin fa-3x"></i>
                </div>
                <p style="margin-top: 10px;">Loading data, please wait...</p>
            </div>
        </div>
        """
        
        # Render the appropriate view
        if view == 'user_list':
            print render_user_list_page(form_handler, analyzer, renderer)
        elif view == 'user_detail':
            print render_user_detail_page(form_handler, analyzer, renderer)
        elif view == 'stale_accounts':
            print render_stale_accounts_page(form_handler, analyzer, renderer)
        elif view == 'locked_accounts':
            print render_locked_accounts_page(analyzer, renderer)
        elif view == 'password_resets':
            print render_password_resets_page(form_handler, analyzer, renderer)
        elif view == 'activity_trends':
            print render_activity_trends_page(form_handler, analyzer, renderer)
        else:
            # Default to overview for any invalid or missing view parameter
            print render_overview_page(form_handler, analyzer, renderer)
        
        # Hide loading indicator and add debug controls
        print """
        <script>
            document.getElementById('loading').style.display = 'none';
            
            // Add click handlers for buttons that might need loading indicator
            document.addEventListener('click', function(e) {
                if (e.target.tagName === 'BUTTON' || e.target.tagName === 'A') {
                    if (!e.target.classList.contains('no-loading')) {
                        document.getElementById('loading').style.display = 'flex';
                    }
                }
            });
            
            // Add a debug button that shows all debug panels
            document.body.insertAdjacentHTML('beforeend', 
                '<button onclick="showAllDebugPanels()" class="btn btn-warning" ' +
                'style="position:fixed; bottom:10px; left:10px;">Show All Debug</button>');
                
            function showAllDebugPanels() {
                document.getElementById('debug-panel').style.display = 'block';
                document.getElementById('form-data-debug').style.display = 'block';
                document.getElementById('view-debug').style.display = 'block';
            }
        </script>
        """
    
    except Exception as e:
        # Print any errors
        print "<h2>Error</h2>"
        print "<p>An error occurred: " + str(e) + "</p>"
        print "<pre>"
        traceback.print_exc()
        print "</pre>"

# Call the main function
main()
