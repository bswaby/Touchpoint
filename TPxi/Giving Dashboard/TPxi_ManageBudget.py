#Roles=Finance

# TPxi_BudgetManager

# ==========================================
# TOUCHPOINT COMPLETE GIVING DASHBOARD
# with Integrated Budget Manager (separate script)
#
# Created by: Ben Swaby
# Email: bswaby@fbchtn.org
# ==========================================
# OVERVIEW
# This is a two-part system:
#  - TPxi_GivingDashboard | Handles main display and reporting
#  - TPxi_ManageBudget    | Creates and manages the weekly budget
#
# Note: this code was made for our church, and while every effort was made for this to work in other environments,
# this may not work in yours or based on how you are structured.  

# PREREQUISITES
#  - Both scripts installed
#  - Weekly budget configured
#  - Variables set in TPxi_GivingDashboard
#
# INSTALLATION — GIVING DASHBOARD
#  1. Go to: Admin → Advanced → Special Content
#  2. Create a new Python Script named: TPxi_GivingDashboard
#  3. Paste in this Python script
#  4. Update configuration items below as needed
#
# INSTALLATION — BUDGET MANAGER
#  1. Go to: Admin → Advanced → Special Content
#  2. Create a new Python Script named: TPxi_BudgetManager
#  3. Paste in the TPxi_BudgetManager.py code from GitHub
#  4. Run the script
#  5. Add in a budget for the years you want to view.
# ==========================================

import datetime
import json

model.Header = 'Budget Manager for Giving Dashboard'

# Configuration
BUDGET_CONTENT_NAME = 'ChurchBudgetData'
DEFAULT_WEEKLY_BUDGET = 285467
FISCAL_MONTH_OFFSET = 3
FISCAL_YEAR_END_MONTH = 9  # September (fiscal year ends Sept 30)
FISCAL_YEAR_END_DAY = 30

# Week Configuration (0=Monday, 6=Sunday)
WEEK_START_DAY = 0  # Monday
WEEK_END_DAY = 6    # Sunday
# Common patterns:
# Sunday-Saturday: WEEK_START_DAY=6, WEEK_END_DAY=5
# Monday-Sunday: WEEK_START_DAY=0, WEEK_END_DAY=6
# Tuesday-Monday: WEEK_START_DAY=1, WEEK_END_DAY=0

def get_week_display_name():
    """Get human-readable week pattern"""
    days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
    return '{} - {}'.format(days[WEEK_START_DAY], days[WEEK_END_DAY])

def get_split_week_entries(start_date, end_date, weekly_amount):
    """
    Generate budget entries that split at fiscal/calendar year boundaries.
    Returns a list of budget entries with their date ranges and amounts.
    """
    entries = []
    current = start_date
    
    # Go to week start day if not already
    if current.weekday() != WEEK_START_DAY:
        # Calculate days back to the previous week start
        days_back = (current.weekday() - WEEK_START_DAY) % 7
        if days_back == 0:
            days_back = 7  # If we're on the start day but want previous week
        current = current - datetime.timedelta(days=days_back)
    
    while current <= end_date:
        week_end = current + datetime.timedelta(days=6)
        
        # Check for fiscal year boundary (Sept 30)
        fiscal_split = False
        for year in range(current.year, week_end.year + 1):
            fiscal_boundary = datetime.datetime(year, FISCAL_YEAR_END_MONTH, FISCAL_YEAR_END_DAY)
            if current <= fiscal_boundary <= week_end:
                if fiscal_boundary == week_end:
                    # Fiscal boundary is the last day of the week - no split needed
                    entries.append({
                        'start_date': current.strftime('%Y-%m-%d'),
                        'end_date': week_end.strftime('%Y-%m-%d'),
                        'amount': weekly_amount,
                        'days': 7,
                        'type': 'fiscal_end',
                        'fiscal_year': fiscal_boundary.year if fiscal_boundary.month <= 9 else fiscal_boundary.year - 1
                    })
                else:
                    # Split the week at fiscal boundary
                    # Part 1: Days before/on boundary (e.g., Mon Sep 22 - Tue Sep 30)
                    days_in_part1 = (fiscal_boundary - current).days + 1
                    entries.append({
                        'start_date': current.strftime('%Y-%m-%d'),
                        'end_date': fiscal_boundary.strftime('%Y-%m-%d'),
                        'amount': weekly_amount,  # Full amount for partial week
                        'days': days_in_part1,
                        'type': 'fiscal_end',
                        'fiscal_year': fiscal_boundary.year if fiscal_boundary.month <= 9 else fiscal_boundary.year - 1
                    })
                    
                    # Part 2: Days after boundary (e.g., Wed Oct 1 - Sun Oct 5)
                    next_start = fiscal_boundary + datetime.timedelta(days=1)
                    days_in_part2 = (week_end - fiscal_boundary).days
                    entries.append({
                        'start_date': next_start.strftime('%Y-%m-%d'),
                        'end_date': week_end.strftime('%Y-%m-%d'),
                        'amount': weekly_amount,  # Full amount for partial week
                        'days': days_in_part2,
                        'type': 'fiscal_start',
                        'fiscal_year': fiscal_boundary.year + 1 if fiscal_boundary.month <= 9 else fiscal_boundary.year
                    })
                fiscal_split = True
                break
        
        if not fiscal_split:
            # Check for calendar year boundary (Dec 31)
            calendar_split = False
            for year in range(current.year, week_end.year + 1):
                calendar_boundary = datetime.datetime(year, 12, 31)
                if current <= calendar_boundary <= week_end:
                    if calendar_boundary == week_end:
                        # Calendar boundary is the last day of the week - no split needed
                        entries.append({
                            'start_date': current.strftime('%Y-%m-%d'),
                            'end_date': week_end.strftime('%Y-%m-%d'),
                            'amount': weekly_amount,
                            'days': 7,
                            'type': 'calendar_end',
                            'calendar_year': calendar_boundary.year
                        })
                    else:
                        # Split the week at calendar boundary
                        # Part 1: Days before/on boundary
                        days_in_part1 = (calendar_boundary - current).days + 1
                        entries.append({
                            'start_date': current.strftime('%Y-%m-%d'),
                            'end_date': calendar_boundary.strftime('%Y-%m-%d'),
                            'amount': weekly_amount,
                            'days': days_in_part1,
                            'type': 'calendar_end',
                            'calendar_year': calendar_boundary.year
                        })
                        
                        # Part 2: Days after boundary
                        next_start = calendar_boundary + datetime.timedelta(days=1)
                        days_in_part2 = (week_end - calendar_boundary).days
                        entries.append({
                            'start_date': next_start.strftime('%Y-%m-%d'),
                            'end_date': week_end.strftime('%Y-%m-%d'),
                            'amount': weekly_amount,
                            'days': days_in_part2,
                            'type': 'calendar_start',
                            'calendar_year': calendar_boundary.year + 1
                        })
                    calendar_split = True
                    break
            
            if not calendar_split:
                # Normal full week
                entries.append({
                    'start_date': current.strftime('%Y-%m-%d'),
                    'end_date': week_end.strftime('%Y-%m-%d'),
                    'amount': weekly_amount,
                    'days': 7,
                    'type': 'normal',
                    'fiscal_year': current.year if current.month < 10 else current.year + 1
                })
        
        current = current + datetime.timedelta(days=7)
    
    return entries

# Check if this is an AJAX request
if model.HttpMethod == 'post' and hasattr(Data, 'ajax') and Data.ajax == 'true':
    # Handle AJAX request
    action = Data.action
    response = {'success': False, 'message': ''}
    
    # Load current budget data
    try:
        budget_data = model.TextContent(BUDGET_CONTENT_NAME)
        if not budget_data:
            budget_data = '{}'
    except:
        budget_data = '{}'
    
    budgets = json.loads(budget_data) if budget_data else {}
    
    if action == 'add':
        week_date = Data.week_date
        amount = int(Data.amount)
        budgets[week_date] = amount
        model.WriteContentText(BUDGET_CONTENT_NAME, json.dumps(budgets))
        response = {'success': True, 'message': 'Budget entry added for ' + week_date, 'week_date': week_date, 'amount': amount}
    
    elif action == 'bulk_add':
        start_date = datetime.datetime.strptime(Data.start_date, '%Y-%m-%d')
        end_date = datetime.datetime.strptime(Data.end_date, '%Y-%m-%d')
        weekly_amount = int(Data.weekly_amount)
        span_partial = Data.span_partial if hasattr(Data, 'span_partial') else 'off'
        handle_boundaries = Data.handle_boundaries if hasattr(Data, 'handle_boundaries') else 'off'
        special_mode = Data.special_mode if hasattr(Data, 'special_mode') else 'extend'
        notes = Data.notes if hasattr(Data, 'notes') else ''
        
        # Store metadata about budget entries
        try:
            metadata_content = model.TextContent('ChurchBudgetMetadata')
            if not metadata_content:
                metadata_content = '{}'
        except:
            metadata_content = '{}'
        
        metadata = json.loads(metadata_content) if metadata_content else {}
        
        # Check if we're using exact dates mode - NO boundary checking, just use exact dates
        if special_mode == 'exact':
            # Create a single entry for the exact date range WITHOUT any modifications
            # Use the EXACT dates provided - no extending, no splitting
            budget_key = start_date.strftime('%Y-%m-%d')
            budgets[budget_key] = weekly_amount
            
            # Store metadata about the entry
            days_diff = (end_date - start_date).days + 1
            
            # Determine fiscal year based on start date
            fiscal_year = start_date.year
            if FISCAL_MONTH_OFFSET == 3 and start_date.month >= 10:
                fiscal_year += 1
            
            metadata[budget_key] = {
                'start_date': start_date.strftime('%Y-%m-%d'),
                'end_date': end_date.strftime('%Y-%m-%d'),
                'days': days_diff,
                'type': 'exact_range',
                'amount': weekly_amount,
                'notes': notes,
                'fiscal_year': fiscal_year
            }
            
            new_entries = [{
                'date': budget_key,
                'amount': weekly_amount,
                'type': 'exact_range',
                'days': days_diff,
                'notes': notes
            }]
            
            model.WriteContentText(BUDGET_CONTENT_NAME, json.dumps(budgets))
            model.WriteContentText('ChurchBudgetMetadata', json.dumps(metadata))
            
            response = {'success': True, 'message': 'Added exact date range budget entry for {} to {}!'.format(start_date.strftime('%m/%d/%Y'), end_date.strftime('%m/%d/%Y')), 'entries': new_entries}
        
        elif handle_boundaries == 'on':
            # For extend mode with boundary checking, create combined entry
            original_start = start_date
            original_end = end_date
            
            # Only extend to full weeks if NOT using exact dates mode
            if special_mode != 'exact':
                # Extend to full weeks using configured week pattern
                if start_date.weekday() != WEEK_START_DAY:
                    # Calculate days to go back to reach the week start
                    if start_date.weekday() < WEEK_START_DAY:
                        # We're before the week start day in the week cycle
                        days_back = start_date.weekday() + (7 - WEEK_START_DAY)
                    else:
                        # We're after the week start day in the week cycle
                        days_back = start_date.weekday() - WEEK_START_DAY
                    start_date = start_date - datetime.timedelta(days=days_back)
                
                if end_date.weekday() != WEEK_END_DAY:
                    days_ahead = (WEEK_END_DAY - end_date.weekday()) % 7
                    if days_ahead == 0:
                        days_ahead = 7  # If we're on the end day but want next week
                    end_date = end_date + datetime.timedelta(days=days_ahead)
            
            # Create single combined entry with boundary notation
            budget_key = start_date.strftime('%Y-%m-%d')
            budgets[budget_key] = weekly_amount
            
            # Calculate total days and check for boundaries
            total_days = (end_date - start_date).days + 1
            
            # Check if this spans a fiscal year boundary
            fiscal_boundary = datetime.datetime(start_date.year, FISCAL_YEAR_END_MONTH, FISCAL_YEAR_END_DAY)
            spans_fiscal = start_date <= fiscal_boundary <= end_date
            
            # Determine type
            entry_type = 'combined'
            if spans_fiscal:
                entry_type = 'fiscal_boundary'
            
            # Add metadata for the combined entry
            metadata[budget_key] = {
                'start_date': start_date.strftime('%Y-%m-%d'),
                'end_date': end_date.strftime('%Y-%m-%d'),
                'days': total_days,
                'type': entry_type,
                'amount': weekly_amount,
                'notes': notes,
                'original_start': original_start.strftime('%Y-%m-%d'),
                'original_end': original_end.strftime('%Y-%m-%d')
            }
            
            new_entries = [{
                'date': budget_key,
                'amount': weekly_amount,
                'type': entry_type,
                'days': total_days,
                'notes': notes
            }]
            
            model.WriteContentText(BUDGET_CONTENT_NAME, json.dumps(budgets))
            model.WriteContentText('ChurchBudgetMetadata', json.dumps(metadata))
            
            if spans_fiscal:
                response = {'success': True, 'message': 'Added combined budget entry spanning fiscal year boundary ({} to {})!'.format(start_date.strftime('%m/%d/%Y'), end_date.strftime('%m/%d/%Y')), 'entries': new_entries}
            else:
                response = {'success': True, 'message': 'Added combined budget entry for {} to {}!'.format(start_date.strftime('%m/%d/%Y'), end_date.strftime('%m/%d/%Y')), 'entries': new_entries}
        
        else:
            # Regular bulk add without boundary splitting
            original_start = start_date
            original_end = end_date
            
            # Adjust dates if spanning partial weeks
            if span_partial == 'on':
                # Go back to previous week start if not already on week start
                if start_date.weekday() != WEEK_START_DAY:
                    # Calculate days to go back to reach the week start
                    if start_date.weekday() < WEEK_START_DAY:
                        # We're before the week start day in the week cycle
                        days_back = start_date.weekday() + (7 - WEEK_START_DAY)
                    else:
                        # We're after the week start day in the week cycle
                        days_back = start_date.weekday() - WEEK_START_DAY
                    start_date = start_date - datetime.timedelta(days=days_back)
                
                # Extend end_date to next week end if not already on week end
                if end_date.weekday() != WEEK_END_DAY:
                    days_ahead = (WEEK_END_DAY - end_date.weekday()) % 7
                    if days_ahead == 0:
                        days_ahead = 7  # If we're on the end day but want next week
                    end_date = end_date + datetime.timedelta(days=days_ahead)
                
                # Create a single combined entry for the extended range
                budget_key = start_date.strftime('%Y-%m-%d')
                budgets[budget_key] = weekly_amount
                
                # Calculate total days in the range
                total_days = (end_date - start_date).days + 1
                
                # Add metadata for the combined entry
                metadata[budget_key] = {
                    'start_date': start_date.strftime('%Y-%m-%d'),
                    'end_date': end_date.strftime('%Y-%m-%d'),
                    'days': total_days,
                    'type': 'combined',
                    'amount': weekly_amount,
                    'notes': notes,
                    'original_start': original_start.strftime('%Y-%m-%d'),
                    'original_end': original_end.strftime('%Y-%m-%d')
                }
                
                new_entries = [{
                    'date': budget_key,
                    'amount': weekly_amount,
                    'notes': notes,
                    'days': total_days
                }]
                
                model.WriteContentText(BUDGET_CONTENT_NAME, json.dumps(budgets))
                model.WriteContentText('ChurchBudgetMetadata', json.dumps(metadata))
                response = {'success': True, 'message': 'Added combined budget entry for {} to {}!'.format(start_date.strftime('%m/%d/%Y'), end_date.strftime('%m/%d/%Y')), 'entries': new_entries}
            
            else:
                # Original behavior - create individual weekly entries
                current_date = start_date
                # Find week start on or after start_date
                if current_date.weekday() != WEEK_START_DAY:
                    days_ahead = (WEEK_START_DAY - current_date.weekday()) % 7
                    if days_ahead == 0:
                        days_ahead = 7  # Move to next week if already on start day
                    current_date = current_date + datetime.timedelta(days=days_ahead)
                
                weeks_added = 0
                new_entries = []
                while current_date <= end_date:
                    budget_key = current_date.strftime('%Y-%m-%d')
                    budgets[budget_key] = weekly_amount
                    
                    # Add metadata for regular entries too
                    week_end = current_date + datetime.timedelta(days=6)
                    metadata[budget_key] = {
                        'start_date': budget_key,
                        'end_date': week_end.strftime('%Y-%m-%d'),
                        'days': 7,
                        'type': 'normal',
                        'amount': weekly_amount,
                        'notes': notes
                    }
                    
                    new_entries.append({
                        'date': budget_key,
                        'amount': weekly_amount,
                        'notes': notes
                    })
                    current_date = current_date + datetime.timedelta(days=7)
                    weeks_added += 1
                
                model.WriteContentText(BUDGET_CONTENT_NAME, json.dumps(budgets))
                model.WriteContentText('ChurchBudgetMetadata', json.dumps(metadata))
                response = {'success': True, 'message': 'Added {} weekly budget entries!'.format(weeks_added), 'entries': new_entries}
    
    elif action == 'delete':
        week_key = Data.week_key
        if week_key in budgets:
            del budgets[week_key]
            # Also delete metadata if it exists
            try:
                metadata_content = model.TextContent('ChurchBudgetMetadata')
                if not metadata_content:
                    metadata_content = '{}'
            except:
                metadata_content = '{}'
            metadata = json.loads(metadata_content) if metadata_content else {}
            if week_key in metadata:
                del metadata[week_key]
                model.WriteContentText('ChurchBudgetMetadata', json.dumps(metadata))
            model.WriteContentText(BUDGET_CONTENT_NAME, json.dumps(budgets))
            response = {'success': True, 'message': 'Budget entry deleted'}
        else:
            response = {'success': False, 'message': 'Entry not found'}
    
    elif action == 'import_csv':
        # Handle CSV import with comprehensive error handling
        try:
            csv_data = Data.csv_data if hasattr(Data, 'csv_data') else ''
            batch_number = int(Data.batch_number) if hasattr(Data, 'batch_number') else 0
            batch_size = 10  # Process 10 CSV lines at a time
            
            if not csv_data:
                response = {'success': False, 'message': 'No CSV data provided'}
            else:
                # Load existing budget data
                try:
                    budget_content = model.TextContent(BUDGET_CONTENT_NAME)
                    if not budget_content:
                        budget_content = '{}'
                except Exception as e:
                    budget_content = '{}'
                budgets = json.loads(budget_content) if budget_content else {}
                
                # Load existing metadata
                try:
                    metadata_content = model.TextContent('ChurchBudgetMetadata')
                    if not metadata_content:
                        metadata_content = '{}'
                except Exception as e:
                    metadata_content = '{}'
                metadata = json.loads(metadata_content) if metadata_content else {}
                
                imported_count = 0
                error_rows = []
                
                # Process CSV data - handle both with and without header
                lines = csv_data.strip().split('\n')
                
                # Check if first line looks like a header
                first_line = lines[0].lower() if lines else ''
                has_header = 'start' in first_line or 'date' in first_line or 'amount' in first_line
                
                # Calculate batch boundaries
                start_idx = 1 if has_header else 0
                batch_start = start_idx + (batch_number * batch_size)
                batch_end = min(batch_start + batch_size, len(lines))
                total_lines = len(lines) - start_idx
                
                # Check if we have more lines to process
                has_more_batches = batch_end < len(lines)
                
                # Process lines in this batch
                processed_lines = 0
                for line_num, line in enumerate(lines[batch_start:batch_end], start=batch_start+1):
                    try:
                        # Skip empty lines
                        if not line.strip():
                            continue
                        
                        processed_lines += 1
                            
                        # Parse CSV line - simple split approach
                        # Format: start_date, end_date, amount, notes, special_mode, handle_boundaries
                        parts = line.split(',')
                        
                        if len(parts) < 3:
                            continue  # Skip invalid lines
                        
                        # Clean up each part
                        start_date_str = parts[0].strip().strip('"')
                        end_date_str = parts[1].strip().strip('"')
                        
                        # Validate and parse amount
                        try:
                            amount_str = parts[2].strip().strip('"')
                            amount = int(amount_str)
                        except (ValueError, IndexError) as e:
                            error_rows.append({'line': line_num, 'error': 'Invalid amount: {}'.format(str(e))})
                            continue
                        
                        # Handle remaining fields with proper defaults
                        notes = parts[3].strip().strip('"') if len(parts) > 3 else ""
                        special_mode = parts[4].strip().strip('"').lower() if len(parts) > 4 else ""
                        handle_boundaries = False
                        if len(parts) > 5:
                            hb_value = parts[5].strip().strip('"').lower()
                            handle_boundaries = hb_value in ['true', '1', 'yes']
                        
                        # Parse dates - try multiple formats
                        for date_format in ['%Y-%m-%d', '%m/%d/%y', '%m/%d/%Y', '%Y/%m/%d']:
                            try:
                                start_date = datetime.datetime.strptime(start_date_str, date_format)
                                break
                            except ValueError:
                                continue
                        else:
                            raise ValueError("Could not parse start date: {}".format(start_date_str))
                        
                        for date_format in ['%Y-%m-%d', '%m/%d/%y', '%m/%d/%Y', '%Y/%m/%d']:
                            try:
                                end_date = datetime.datetime.strptime(end_date_str, date_format)
                                break
                            except ValueError:
                                continue
                        else:
                            raise ValueError("Could not parse end date: {}".format(end_date_str))
                        
                        # Process based on special_mode
                        if special_mode == 'exact':
                            # Use exact dates as provided - convert to standard format
                            budget_key = start_date.strftime('%Y-%m-%d')
                            
                            # Skip if already exists with same amount
                            if budget_key in budgets and budgets[budget_key] == amount:
                                continue  # Skip duplicate
                                
                            budgets[budget_key] = amount
                            
                            days = (end_date - start_date).days + 1
                            metadata[budget_key] = {
                                'start_date': start_date.strftime('%Y-%m-%d'),
                                'end_date': end_date.strftime('%Y-%m-%d'),
                                'days': days,
                                'type': 'exact_range',
                                'amount': amount,
                                'notes': notes
                            }
                            imported_count += 1
                            
                        elif special_mode == 'extend' or handle_boundaries:
                            # Use the existing function to handle special weeks
                            entries = get_split_week_entries(start_date, end_date, amount)
                            
                            for entry in entries:
                                budget_key = entry['start_date']
                                budgets[budget_key] = entry['amount']
                                
                                # Add notes if this is a boundary entry
                                entry_notes = notes
                                if entry.get('type') in ['fiscal_end', 'fiscal_start', 'calendar_end', 'calendar_start']:
                                    if entry['type'] == 'fiscal_end':
                                        entry_notes = notes + " | Fiscal year-end" if notes else "Fiscal year-end"
                                    elif entry['type'] == 'fiscal_start':
                                        entry_notes = notes + " | Fiscal year start" if notes else "Fiscal year start"
                                    elif entry['type'] == 'calendar_end':
                                        entry_notes = notes + " | Calendar year-end" if notes else "Calendar year-end"
                                    elif entry['type'] == 'calendar_start':
                                        entry_notes = notes + " | Calendar year start" if notes else "Calendar year start"
                                
                                metadata[budget_key] = {
                                    'start_date': entry['start_date'],
                                    'end_date': entry['end_date'],
                                    'days': entry.get('days', 7),
                                    'type': entry.get('type', 'normal'),
                                    'amount': entry['amount'],
                                    'notes': entry_notes
                                }
                                imported_count += 1
                        else:
                            # Standard weekly entries (no special handling)
                            # Calculate weekly entries between dates
                            current = start_date
                            while current <= end_date:
                                week_end = current + datetime.timedelta(days=6)
                                if week_end > end_date:
                                    week_end = end_date
                                
                                budget_key = current.strftime('%Y-%m-%d')
                                
                                # Skip if already exists with same amount
                                if budget_key in budgets and budgets[budget_key] == amount:
                                    current = current + datetime.timedelta(days=7)
                                    continue  # Skip duplicate
                                
                                budgets[budget_key] = amount
                                
                                days = (week_end - current).days + 1
                                metadata[budget_key] = {
                                    'start_date': budget_key,
                                    'end_date': week_end.strftime('%Y-%m-%d'),
                                    'days': days,
                                    'type': 'imported',
                                    'amount': amount,
                                    'notes': notes
                                }
                                imported_count += 1
                                current = current + datetime.timedelta(days=7)
                        
                    except Exception as e:
                        error_rows.append("Line {}: {}".format(line_num, str(e)))
                
                # Save data after each batch
                model.WriteContentText(BUDGET_CONTENT_NAME, json.dumps(budgets))
                model.WriteContentText('ChurchBudgetMetadata', json.dumps(metadata))
                
                # Prepare response with batch information
                if has_more_batches:
                    lines_processed = min(batch_end - batch_start, batch_size)
                    lines_remaining = len(lines) - batch_end
                    response = {
                        'success': True, 
                        'message': 'Batch {} complete. Imported {} entries from {} CSV lines (lines {}-{}). {} lines remaining.'.format(
                            batch_number + 1, imported_count, lines_processed, batch_start + 1, batch_end, lines_remaining
                        ),
                        'has_more': True,
                        'next_batch': batch_number + 1,
                        'total_imported': imported_count,
                        'lines_remaining': lines_remaining,
                        'batch_start': batch_start,
                        'batch_end': batch_end,
                        'total_lines': len(lines)
                    }
                else:
                    if error_rows:
                        response = {
                            'success': True, 
                            'message': 'Import complete! Total entries imported: {}. {} errors occurred.'.format(imported_count, len(error_rows)),
                            'has_more': False,
                            'total_imported': imported_count
                        }
                    else:
                        response = {
                            'success': True, 
                            'message': 'Import complete! Successfully imported {} budget entries.'.format(imported_count),
                            'has_more': False,
                            'total_imported': imported_count
                        }
        except Exception as e:
            # Catch any unexpected errors during import
            response = {'success': False, 'message': 'Import failed: {}'.format(str(e)), 'has_more': False}
    
    elif action == 'delete_multiple':
        # Handle multiple deletes
        week_keys = Data.week_keys.split(',') if hasattr(Data, 'week_keys') else []
        
        # Load metadata
        try:
            metadata_content = model.TextContent('ChurchBudgetMetadata')
            if not metadata_content:
                metadata_content = '{}'
        except:
            metadata_content = '{}'
        metadata = json.loads(metadata_content) if metadata_content else {}
        
        deleted_count = 0
        for week_key in week_keys:
            if week_key in budgets:
                del budgets[week_key]
                if week_key in metadata:
                    del metadata[week_key]
                deleted_count += 1
        
        if deleted_count > 0:
            model.WriteContentText(BUDGET_CONTENT_NAME, json.dumps(budgets))
            model.WriteContentText('ChurchBudgetMetadata', json.dumps(metadata))
            response = {'success': True, 'message': 'Deleted {} entries'.format(deleted_count)}
        else:
            response = {'success': False, 'message': 'No entries found to delete'}
    
    # Return JSON response for AJAX
    print json.dumps(response)
    print "\n<!-- AJAX_COMPLETE -->"

else:
    # Regular page load or form submission
    success_message = ''
    
    # Handle regular form submissions
    if model.HttpMethod == 'post' and hasattr(Data, 'action'):
        action = Data.action
        
        # Load current budget data
        try:
            budget_data = model.TextContent(BUDGET_CONTENT_NAME)
            if not budget_data:
                budget_data = '{}'
        except:
            budget_data = '{}'
        
        budgets = json.loads(budget_data) if budget_data else {}
        
        if action == 'add':
            week_date = Data.week_date
            amount = int(Data.amount)
            budgets[week_date] = amount
            model.WriteContentText(BUDGET_CONTENT_NAME, json.dumps(budgets))
            success_message = 'Budget entry added for ' + week_date + '!'
        
        elif action == 'bulk_add':
            start_date = datetime.datetime.strptime(Data.start_date, '%Y-%m-%d')
            end_date = datetime.datetime.strptime(Data.end_date, '%Y-%m-%d')
            weekly_amount = int(Data.weekly_amount)
            span_partial = Data.span_partial if hasattr(Data, 'span_partial') else 'off'
            handle_boundaries = Data.handle_boundaries if hasattr(Data, 'handle_boundaries') else 'off'
            special_mode = Data.special_mode if hasattr(Data, 'special_mode') else 'extend'
            notes = Data.notes if hasattr(Data, 'notes') else ''
            
            # Store metadata about budget entries
            try:
                metadata_content = model.TextContent('ChurchBudgetMetadata')
                if not metadata_content:
                    metadata_content = '{}'
            except:
                metadata_content = '{}'
            
            metadata = json.loads(metadata_content) if metadata_content else {}
            
            # Debug: Check what we're receiving
            debug_msg = 'Debug: handle_boundaries={}, special_mode={}, start={}, end={}'.format(
                handle_boundaries, special_mode, start_date.strftime('%Y-%m-%d'), end_date.strftime('%Y-%m-%d')
            )
            success_message = debug_msg  # Show debug info in the success message temporarily
            
            # Check if using exact dates mode - just use the exact dates provided
            # When exact mode is selected, we ignore boundary splitting
            if special_mode == 'exact':
                success_message = 'EXACT MODE: Creating single entry for {} to {}'.format(
                    start_date.strftime('%Y-%m-%d'), end_date.strftime('%Y-%m-%d')
                )
                # Create single entry for exact date range
                budget_key = start_date.strftime('%Y-%m-%d')
                budgets[budget_key] = weekly_amount
                
                # Calculate total days
                total_days = (end_date - start_date).days + 1
                
                # Determine fiscal year
                fiscal_year = start_date.year
                if FISCAL_MONTH_OFFSET == 3 and start_date.month >= 10:
                    fiscal_year += 1
                
                # Store metadata
                metadata[budget_key] = {
                    'start_date': start_date.strftime('%Y-%m-%d'),
                    'end_date': end_date.strftime('%Y-%m-%d'),
                    'days': total_days,
                    'type': 'exact_range',
                    'amount': weekly_amount,
                    'notes': notes,
                    'fiscal_year': fiscal_year
                }
                
                model.WriteContentText(BUDGET_CONTENT_NAME, json.dumps(budgets))
                model.WriteContentText('ChurchBudgetMetadata', json.dumps(metadata))
                success_message = 'Added exact date range budget entry for {} to {}!'.format(
                    start_date.strftime('%m/%d/%Y'), 
                    end_date.strftime('%m/%d/%Y')
                )
            
            elif handle_boundaries == 'on':
                # Use the new split week handling
                entries = get_split_week_entries(start_date, end_date, weekly_amount)
                
                entries_added = 0
                split_entries = 0
                
                for entry in entries:
                    # Use start_date as the key for budget entries
                    budget_key = entry['start_date']
                    budgets[budget_key] = entry['amount']
                    
                    # Store metadata about the entry
                    metadata[budget_key] = {
                        'start_date': entry['start_date'],
                        'end_date': entry['end_date'],
                        'days': entry['days'],
                        'type': entry['type'],
                        'amount': entry['amount']
                    }
                    
                    if entry['type'] in ['fiscal_end', 'fiscal_start', 'calendar_end', 'calendar_start']:
                        split_entries += 1
                    
                    entries_added += 1
                
                model.WriteContentText(BUDGET_CONTENT_NAME, json.dumps(budgets))
                model.WriteContentText('ChurchBudgetMetadata', json.dumps(metadata))
                
                if split_entries > 0:
                    success_message = 'Added {} budget entries ({} are partial weeks at year boundaries)!'.format(entries_added, split_entries)
                else:
                    success_message = 'Added {} budget entries!'.format(entries_added)
            
            else:
                # Regular bulk add without boundary splitting
                # Check if this is an exact date range (not extending to full weeks)
                # When span_partial is 'off', use exact dates
                use_exact_dates = (span_partial == 'off')
                weeks_added = 0
                new_entries = []
                
                # If using exact dates and it's not a standard week pattern, create single entry
                if use_exact_dates and (start_date != end_date):
                    # Check if this is not a standard Sunday-Saturday week
                    days_diff = (end_date - start_date).days + 1
                    
                    # Create single entry for the exact date range
                    budgets[start_date.strftime('%Y-%m-%d')] = weekly_amount
                    metadata[start_date.strftime('%Y-%m-%d')] = {
                        'start_date': start_date.strftime('%Y-%m-%d'),
                        'end_date': end_date.strftime('%Y-%m-%d'),
                        'days': days_diff,
                        'type': 'exact_range',
                        'amount': weekly_amount,
                        'notes': notes
                    }
                    new_entries.append({
                        'date': start_date.strftime('%Y-%m-%d'),
                        'amount': weekly_amount
                    })
                    weeks_added = 1
                    
                elif use_exact_dates and start_date == end_date:
                    # Single date entry
                    budgets[start_date.strftime('%Y-%m-%d')] = weekly_amount
                    metadata[start_date.strftime('%Y-%m-%d')] = {
                        'start_date': start_date.strftime('%Y-%m-%d'),
                        'end_date': end_date.strftime('%Y-%m-%d'),
                        'days': 1,
                        'type': 'single',
                        'amount': weekly_amount,
                        'notes': notes
                    }
                    new_entries.append({
                        'date': start_date.strftime('%Y-%m-%d'),
                        'amount': weekly_amount
                    })
                    weeks_added = 1
                    
                elif use_exact_dates and (end_date - start_date).days < 7:
                    # Exact partial week entry
                    budgets[start_date.strftime('%Y-%m-%d')] = weekly_amount
                    metadata[start_date.strftime('%Y-%m-%d')] = {
                        'start_date': start_date.strftime('%Y-%m-%d'),
                        'end_date': end_date.strftime('%Y-%m-%d'),
                        'days': (end_date - start_date).days + 1,
                        'type': 'partial',
                        'amount': weekly_amount,
                        'notes': notes
                    }
                    new_entries.append({
                        'date': start_date.strftime('%Y-%m-%d'),
                        'amount': weekly_amount
                    })
                    weeks_added = 1
                    
                else:
                    # Normal weekly processing
                    current_date = start_date
                    if span_partial == 'on':
                        # Go back to previous week start if not already on week start
                        if current_date.weekday() != WEEK_START_DAY:
                            days_back = (current_date.weekday() - WEEK_START_DAY) % 7
                            if days_back == 0:
                                days_back = 7
                            current_date = current_date - datetime.timedelta(days=days_back)
                    else:
                        # Find week start on or after start_date (original behavior)
                        if current_date.weekday() != WEEK_START_DAY:
                            days_ahead = (WEEK_START_DAY - current_date.weekday()) % 7
                            if days_ahead == 0:
                                days_ahead = 7  # Move to next week start
                            current_date = current_date + datetime.timedelta(days=days_ahead)
                
                # Extend end_date to week end if spanning partial weeks
                if span_partial == 'on' and end_date.weekday() != WEEK_END_DAY:
                    days_ahead = (WEEK_END_DAY - end_date.weekday()) % 7
                    if days_ahead == 0:
                        days_ahead = 7  # If it's already week end, go to next week end
                    end_date = end_date + datetime.timedelta(days=days_ahead)
                
                weeks_added = 0
                while current_date <= end_date:
                    budgets[current_date.strftime('%Y-%m-%d')] = weekly_amount
                    current_date = current_date + datetime.timedelta(days=7)
                    weeks_added += 1
                
                model.WriteContentText(BUDGET_CONTENT_NAME, json.dumps(budgets))
                success_message = 'Added {} weekly budget entries!'.format(weeks_added)
    
    # Load budget data for display
    try:
        budget_data = model.TextContent(BUDGET_CONTENT_NAME)
        if not budget_data:
            budget_data = '{}'
    except:
        budget_data = '{}'
    
    budgets = json.loads(budget_data) if budget_data else {}
    
    # Load metadata for special weeks
    try:
        metadata_content = model.TextContent('ChurchBudgetMetadata')
        if not metadata_content:
            metadata_content = '{}'
    except:
        metadata_content = '{}'
    
    metadata = json.loads(metadata_content) if metadata_content else {}
    
    # Sort budgets by date
    sorted_budgets = []
    for date_str in sorted(budgets.keys(), reverse=True):
        budget_entry = {
            'date': date_str,
            'amount': budgets[date_str]
        }
        
        # Add metadata if this is a special week
        if date_str in metadata:
            budget_entry['is_split'] = metadata[date_str].get('is_split', False)
            budget_entry['split_type'] = metadata[date_str].get('split_type', '')
            budget_entry['split_date'] = metadata[date_str].get('split_date', '')
            budget_entry['days_before'] = metadata[date_str].get('days_before', 7)
            budget_entry['days_after'] = metadata[date_str].get('days_after', 0)
        else:
            budget_entry['is_split'] = False
            
        sorted_budgets.append(budget_entry)
    
    # Build the HTML form
    html = '''
<style>
    .budget-manager { max-width: 1200px; margin: 0 auto; padding: 20px; }
    .card { background: white; border-radius: 10px; padding: 25px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); margin-bottom: 20px; }
    .form-group { margin-bottom: 20px; }
    .form-group label { display: block; margin-bottom: 5px; font-weight: 600; color: #555; }
    .form-group input { width: 100%; padding: 10px; border: 1px solid #ddd; border-radius: 5px; }
    .button { background: #667eea; color: white; padding: 12px 24px; border: none; border-radius: 5px; cursor: pointer; }
    .button:hover { background: #5a67d8; }
    .button.danger { background: #ef4444; }
    .button.danger:hover { background: #dc2626; }
    .table { width: 100%; border-collapse: collapse; }
    .table th { background: #f8f9fa; padding: 12px; text-align: left; }
    .table td { padding: 12px; border-bottom: 1px solid #dee2e6; }
    .alert { padding: 15px; border-radius: 5px; margin-bottom: 20px; }
    .alert.success { background: #d1fae5; color: #065f46; }
    .alert.danger { background: #fee2e2; color: #991b1b; }
    .tabs { display: flex; gap: 10px; margin-bottom: 20px; }
    .tab { padding: 12px 20px; background: none; border: none; cursor: pointer; }
    .tab.active { color: #667eea; border-bottom: 2px solid #667eea; }
    .tab-content { display: none; }
    .tab-content.active { display: block; }
    .loading { opacity: 0.5; pointer-events: none; }
    .checkbox-col { width: 40px; }
    .delete-selected { background: #dc2626; margin-left: 10px; }
    .split-week { background: #fef3c7; }
    .split-indicator { 
        display: inline-block; 
        padding: 2px 6px; 
        border-radius: 3px; 
        font-size: 0.75em; 
        font-weight: bold;
        margin-left: 5px;
    }
    .split-fiscal { background: #ddd6fe; color: #6b21a8; }
    .split-calendar { background: #bfdbfe; color: #1e40af; }
</style>

<div class="budget-manager">
    <div id="alertContainer">'''
    
    if success_message:
        html += '<div class="alert success">' + success_message + '</div>'
    
    html += '''
    </div>
    
    <div class="card">
        <div class="tabs">
            <button class="tab active" onclick="showTab('add')">Add Budget</button>
            <button class="tab" onclick="showTab('import')">Import CSV</button>
            <button class="tab" onclick="showTab('view')">View Budget</button>
        </div>
        
        <!-- Add Budget -->
        <div id="add" class="tab-content active">
            <h2>Add Budget Entries</h2>
            <div style="background: #f0f9ff; border: 1px solid #0284c7; padding: 10px; border-radius: 5px; margin-bottom: 20px;">
                <strong>Week Pattern:</strong> ''' + get_week_display_name() + '''<br>
                <small>Enter the same date for both Start and End to add a single week</small>
            </div>
            <form id="bulkAddForm" onsubmit="return handleBulkAddSubmit(event)">
                <div class="form-group">
                    <label>Start Date</label>
                    <input type="date" name="start_date" id="start_date" required>
                </div>
                <div class="form-group">
                    <label>End Date</label>
                    <input type="date" name="end_date" id="end_date" required>
                </div>
                <div class="form-group">
                    <label>Weekly Amount</label>
                    <input type="number" name="weekly_amount" id="weekly_amount" value="''' + str(DEFAULT_WEEKLY_BUDGET) + '''" required>
                </div>
                <div class="form-group">
                    <label>Notes (Optional)</label>
                    <input type="text" name="notes" id="notes" placeholder="e.g., Fiscal year-end adjustment, Special campaign" style="padding: 10px; border: 1px solid #ddd; border-radius: 5px; width: 100%;">
                </div>
                
                <div style="background: #fef3c7; border: 1px solid #f59e0b; padding: 15px; border-radius: 5px; margin: 20px 0;">
                    <label style="display: flex; align-items: start; cursor: pointer;">
                        <input type="checkbox" name="handle_special" id="handle_special" value="on" style="margin-right: 10px; margin-top: 3px;" onchange="toggleSpecialOptions()"> 
                        <div>
                            <strong>Handle special weeks</strong><br>
                            <small style="color: #666;">
                                Check this to handle weeks that don't follow the standard ''' + get_week_display_name() + ''' pattern.<br>
                                This includes partial weeks at date boundaries and fiscal/calendar year splits.
                            </small>
                        </div>
                    </label>
                    
                    <div id="specialOptions" style="display: none; margin-top: 15px; padding-left: 30px;">
                        <div class="form-group" style="margin-bottom: 10px;">
                            <label>
                                <input type="radio" name="special_mode" value="exact" checked> 
                                <strong>Exact dates</strong> - Use the dates exactly as entered (for partial weeks)
                            </label>
                        </div>
                        <div class="form-group" style="margin-bottom: 10px;">
                            <label>
                                <input type="radio" name="special_mode" value="extend"> 
                                <strong>Extend to full weeks</strong> - Include complete weeks that overlap the date range
                            </label>
                        </div>
                        <div class="form-group">
                            <label>
                                <input type="checkbox" name="handle_boundaries" id="handle_boundaries" value="on"> 
                                Split at year boundaries (Sept 30 for fiscal, Dec 31 for calendar)
                            </label>
                        </div>
                    </div>
                </div>
                <button type="submit" class="button">Add Budget Entries</button>
            </form>
        </div>
        
        <!-- Import CSV -->
        <div id="import" class="tab-content">
            <h2>Import Budget Data from CSV</h2>
            <div style="background: #f0f9ff; border: 1px solid #0284c7; padding: 15px; border-radius: 5px; margin-bottom: 20px;">
                <strong>CSV Format:</strong> start_date, end_date, amount, notes, special_mode, handle_boundaries<br>
                <small style="display: block; margin-top: 10px;">
                <strong>✨ No size limit!</strong> The import will automatically process large files in batches to prevent timeouts.<br><br>
                
                <strong>Columns:</strong><br>
                • <strong>start_date/end_date:</strong> YYYY-MM-DD format<br>
                • <strong>amount:</strong> Budget amount for the period<br>
                • <strong>notes:</strong> Optional notes (leave empty for none)<br>
                • <strong>special_mode:</strong> Leave empty for standard, or use "exact" or "extend"<br>
                • <strong>handle_boundaries:</strong> true/false to split at fiscal/calendar year boundaries<br><br>
                
                <strong>Examples:</strong><br>
                <code style="background: #f3f4f6; padding: 2px 4px;">2023-09-18,2023-09-30,230208,Fiscal year-end,exact,</code> - Exact dates for partial week<br>
                <code style="background: #f3f4f6; padding: 2px 4px;">2023-10-01,2023-10-07,285467,,,</code> - Standard weekly entry<br>
                <code style="background: #f3f4f6; padding: 2px 4px;">2023-09-01,2023-10-31,285467,,extend,true</code> - Extend to full weeks with boundary splits
                </small>
            </div>
            <form id="importCsvForm" onsubmit="return handleImportCsvSubmit(event)">
                <div class="form-group">
                    <label>CSV Data</label>
                    <textarea name="csv_data" id="csv_data" rows="15" style="width: 100%; padding: 10px; border: 1px solid #ddd; border-radius: 5px; font-family: monospace;" placeholder="start_date,end_date,amount,notes,special_mode,handle_boundaries
2023-09-18,2023-09-30,230208,Fiscal year-end,exact,
2023-10-01,2023-10-07,285467,,,
2023-10-08,2023-10-14,285467,,,
2023-12-25,2023-12-31,285467,Calendar year-end,exact," required></textarea>
                </div>
                <button type="submit" class="button">Import Budget Data</button>
            </form>
        </div>
        
        <!-- View Budget -->
        <div id="view" class="tab-content">
            <h2>View Budget Entries</h2>
            
            <!-- Filter Section -->
            <div style="background: #f9fafb; border: 1px solid #e5e7eb; padding: 15px; border-radius: 5px; margin-bottom: 20px;">
                <div style="display: flex; gap: 15px; align-items: end; flex-wrap: wrap;">
                    <div style="flex: 1; min-width: 150px;">
                        <label style="display: block; margin-bottom: 5px; font-size: 0.9em;">Filter by Date Range</label>
                        <input type="date" id="filterStartDate" style="padding: 5px; border: 1px solid #ddd; border-radius: 3px; width: 100%;">
                    </div>
                    <div style="flex: 1; min-width: 150px;">
                        <label style="display: block; margin-bottom: 5px; font-size: 0.9em;">&nbsp;</label>
                        <input type="date" id="filterEndDate" style="padding: 5px; border: 1px solid #ddd; border-radius: 3px; width: 100%;">
                    </div>
                    <div style="flex: 1; min-width: 150px;">
                        <label style="display: block; margin-bottom: 5px; font-size: 0.9em;">Fiscal Year</label>
                        <select id="filterFiscalYear" style="padding: 5px; border: 1px solid #ddd; border-radius: 3px; width: 100%;">
                            <option value="">All Years</option>'''
    
    # Get unique fiscal years from the data
    fiscal_years = set()
    for budget in sorted_budgets:
        # Try multiple date formats to handle legacy data
        date_str = budget['date']
        date_obj = None
        for date_format in ['%Y-%m-%d', '%m/%d/%y', '%m/%d/%Y', '%Y/%m/%d']:
            try:
                date_obj = datetime.datetime.strptime(date_str, date_format)
                break
            except ValueError:
                continue
        
        if not date_obj:
            # Skip entries with unparseable dates
            continue
            
        fiscal_year = date_obj.year
        if FISCAL_MONTH_OFFSET == 3 and date_obj.month >= 10:
            fiscal_year += 1
        fiscal_years.add(fiscal_year)
    
    # Add options for each fiscal year in the data
    for fy in sorted(fiscal_years, reverse=True):
        html += '''
                            <option value="''' + str(fy) + '''">FY''' + str(fy) + '''</option>'''
    
    html += '''
                        </select>
                    </div>
                    <div>
                        <button onclick="applyFilter()" class="button" style="padding: 5px 15px;">Apply Filter</button>
                        <button onclick="clearFilter()" class="button" style="padding: 5px 15px; background: #6b7280;">Clear</button>
                    </div>
                </div>
            </div>'''
    
    if sorted_budgets:
        html += '''
            <div style="margin-bottom: 15px;">
                <button onclick="toggleSelectAll()" class="button" style="padding: 8px 16px; font-size: 0.9em;">Select All</button>
                <button onclick="deleteSelected()" class="button delete-selected" style="padding: 8px 16px; font-size: 0.9em;">Delete Selected</button>
                <span id="filterStatus" style="margin-left: 15px; color: #059669; font-weight: bold;"></span>
            </div>
            <table class="table">
                <thead>
                    <tr>
                        <th class="checkbox-col">
                            <input type="checkbox" id="selectAllCheckbox" onchange="toggleAllCheckboxes()">
                        </th>
                        <th>Week Starting</th>
                        <th>Budget Amount</th>
                        <th>Fiscal Year</th>
                        <th>Notes</th>
                        <th>Actions</th>
                    </tr>
                </thead>
                <tbody id="budgetTableBody">'''
        
        for budget in sorted_budgets:  # Show all budget entries
            # Try multiple date formats to handle legacy data
            date_str = budget['date']
            date_obj = None
            for date_format in ['%Y-%m-%d', '%m/%d/%y', '%m/%d/%Y', '%Y/%m/%d']:
                try:
                    date_obj = datetime.datetime.strptime(date_str, date_format)
                    break
                except ValueError:
                    continue
            
            if not date_obj:
                # Skip entries with unparseable dates
                continue
                
            fiscal_year = date_obj.year
            if FISCAL_MONTH_OFFSET == 3 and date_obj.month >= 10:
                fiscal_year += 1
            
            # Check if this is a partial week entry
            row_class = ''
            notes = ''
            date_display = date_str
            
            # Check metadata for this entry
            if budget['date'] in metadata:
                entry_meta = metadata[budget['date']]
                entry_type = entry_meta.get('type', 'normal')
                end_date = entry_meta.get('end_date', '')
                days = entry_meta.get('days', 7)
                
                # Special formatting for different types
                if entry_type in ['fiscal_end', 'fiscal_start', 'calendar_end', 'calendar_start', 'fiscal_boundary']:
                    row_class = 'split-week'
                elif entry_type in ['exact_range', 'partial', 'combined']:
                    row_class = 'split-week'  # Use same highlighting for partial weeks and combined entries
                
                # Always format date range display if we have an end_date
                if end_date and end_date != budget['date']:
                    # Parse start date with multiple formats
                    start_dt = None
                    for date_format in ['%Y-%m-%d', '%m/%d/%y', '%m/%d/%Y', '%Y/%m/%d']:
                        try:
                            start_dt = datetime.datetime.strptime(budget['date'], date_format)
                            break
                        except ValueError:
                            continue
                    
                    # Parse end date with multiple formats
                    end_dt = None
                    for date_format in ['%Y-%m-%d', '%m/%d/%y', '%m/%d/%Y', '%Y/%m/%d']:
                        try:
                            end_dt = datetime.datetime.strptime(end_date, date_format)
                            break
                        except ValueError:
                            continue
                    
                    if start_dt and end_dt:
                        date_display = '{} - {}'.format(
                            start_dt.strftime('%b %d, %Y'),
                            end_dt.strftime('%b %d, %Y')
                        )
                    else:
                        date_display = budget['date']  # Fallback to raw date
                
                # Add notes about the split type
                if entry_type == 'fiscal_end':
                    notes = '<span class="split-indicator split-fiscal">FY End ({} days)</span>'.format(days)
                elif entry_type == 'fiscal_start':
                    notes = '<span class="split-indicator split-fiscal">FY Start ({} days)</span>'.format(days)
                elif entry_type == 'fiscal_boundary':
                    notes = '<span class="split-indicator split-fiscal">Spans FY Boundary ({} days)</span>'.format(days)
                elif entry_type == 'calendar_end':
                    notes = '<span class="split-indicator split-calendar">CY End ({} days)</span>'.format(days)
                elif entry_type == 'calendar_start':
                    notes = '<span class="split-indicator split-calendar">CY Start ({} days)</span>'.format(days)
                elif entry_type == 'exact_range':
                    notes = '<span class="split-indicator" style="color: #059669;">Exact date range ({} days)</span>'.format(days)
                elif entry_type == 'combined':
                    notes = '<span class="split-indicator" style="color: #4b5563;">Combined period ({} days)</span>'.format(days)
                
                # Add user notes if present
                entry_notes = entry_meta.get('notes', '')
                if entry_notes:
                    if notes:
                        notes = notes + '<br>'
                    notes = notes + '<small style="color: #666;">' + entry_notes + '</small>'
            
            html += '''
                    <tr id="row-''' + budget['date'] + '''" class="''' + row_class + '''">
                        <td class="checkbox-col">
                            <input type="checkbox" class="budget-checkbox" value="''' + budget['date'] + '''">
                        </td>
                        <td>''' + date_display + '''</td>
                        <td>$''' + '{:,}'.format(budget['amount']) + '''</td>
                        <td>FY''' + str(fiscal_year) + '''</td>
                        <td>''' + notes + '''</td>
                        <td>
                            <button onclick="deleteEntry(\'''' + budget['date'] + '''\')" class="button danger" style="padding: 5px 10px; font-size: 0.85em;">Delete</button>
                        </td>
                    </tr>'''
        
        html += '''
                </tbody>
            </table>'''
    else:
        html += '<p>No budget entries found.</p>'
    
    html += '''
        </div>
    </div>
</div>

<script>
function showTab(tabName) {
    var tabs = document.getElementsByClassName('tab-content');
    for (var i = 0; i < tabs.length; i++) {
        tabs[i].classList.remove('active');
    }
    
    var buttons = document.getElementsByClassName('tab');
    for (var i = 0; i < buttons.length; i++) {
        buttons[i].classList.remove('active');
    }
    
    document.getElementById(tabName).classList.add('active');
    
    // If called from a button click, highlight the button
    if (typeof event !== 'undefined' && event && event.target) {
        event.target.classList.add('active');
    } else {
        // If called programmatically, find and highlight the corresponding tab button
        for (var i = 0; i < buttons.length; i++) {
            if (buttons[i].getAttribute('onclick') && buttons[i].getAttribute('onclick').indexOf("'" + tabName + "'") !== -1) {
                buttons[i].classList.add('active');
                break;
            }
        }
    }
}

function showAlert(message, type) {
    var alertDiv = document.createElement('div');
    alertDiv.className = 'alert ' + type;
    alertDiv.textContent = message;
    
    var container = document.getElementById('alertContainer');
    container.innerHTML = '';
    container.appendChild(alertDiv);
    
    setTimeout(function() {
        alertDiv.remove();
    }, 5000);
}

function toggleSelectAll() {
    var checkboxes = document.getElementsByClassName('budget-checkbox');
    var allChecked = true;
    for (var i = 0; i < checkboxes.length; i++) {
        if (!checkboxes[i].checked) {
            allChecked = false;
            break;
        }
    }
    
    for (var i = 0; i < checkboxes.length; i++) {
        checkboxes[i].checked = !allChecked;
    }
    
    document.getElementById('selectAllCheckbox').checked = !allChecked;
}

function toggleAllCheckboxes() {
    var selectAll = document.getElementById('selectAllCheckbox');
    var checkboxes = document.getElementsByClassName('budget-checkbox');
    for (var i = 0; i < checkboxes.length; i++) {
        checkboxes[i].checked = selectAll.checked;
    }
}

function deleteSelected() {
    var checkboxes = document.getElementsByClassName('budget-checkbox');
    var selectedDates = [];
    
    for (var i = 0; i < checkboxes.length; i++) {
        if (checkboxes[i].checked) {
            selectedDates.push(checkboxes[i].value);
        }
    }
    
    if (selectedDates.length === 0) {
        showAlert('No entries selected', 'danger');
        return;
    }
    
    // Show loading state for all selected rows
    for (var i = 0; i < selectedDates.length; i++) {
        var row = document.getElementById('row-' + selectedDates[i]);
        if (row) {
            row.classList.add('loading');
        }
    }
    
    // Build the AJAX URL
    var ajaxUrl = window.location.pathname;
    if (ajaxUrl.indexOf('/PyScript/') !== -1) {
        ajaxUrl = ajaxUrl.replace('/PyScript/', '/PyScriptForm/');
    }
    
    // Create XMLHttpRequest
    var xhr = new XMLHttpRequest();
    xhr.open('POST', ajaxUrl, true);
    xhr.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
    
    xhr.onload = function() {
        if (xhr.status === 200) {
            try {
                var response = JSON.parse(xhr.responseText.split('<!-- AJAX_COMPLETE -->')[0].trim());
                
                if (response.success) {
                    // Remove all selected rows
                    for (var i = 0; i < selectedDates.length; i++) {
                        var row = document.getElementById('row-' + selectedDates[i]);
                        if (row) {
                            row.remove();
                        }
                    }
                    showAlert(response.message, 'success');
                    document.getElementById('selectAllCheckbox').checked = false;
                } else {
                    // Restore rows on error
                    for (var i = 0; i < selectedDates.length; i++) {
                        var row = document.getElementById('row-' + selectedDates[i]);
                        if (row) {
                            row.classList.remove('loading');
                        }
                    }
                    showAlert(response.message || 'Error deleting entries', 'danger');
                }
            } catch (e) {
                showAlert('Error processing response', 'danger');
            }
        } else {
            showAlert('Server error', 'danger');
        }
    };
    
    xhr.onerror = function() {
        showAlert('Network error', 'danger');
    };
    
    // Send AJAX request with selected dates
    var params = 'ajax=true&action=delete_multiple&week_keys=' + encodeURIComponent(selectedDates.join(','));
    xhr.send(params);
}

function deleteEntry(weekDate) {
    // Show loading state immediately (no confirmation)
    var row = document.getElementById('row-' + weekDate);
    if (row) {
        row.classList.add('loading');
    }
    
    // Build the AJAX URL
    var ajaxUrl = window.location.pathname;
    if (ajaxUrl.indexOf('/PyScript/') !== -1) {
        ajaxUrl = ajaxUrl.replace('/PyScript/', '/PyScriptForm/');
    }
    
    // Create XMLHttpRequest
    var xhr = new XMLHttpRequest();
    xhr.open('POST', ajaxUrl, true);
    xhr.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
    
    xhr.onload = function() {
        if (xhr.status === 200) {
            try {
                // Parse JSON response
                var response = JSON.parse(xhr.responseText.split('<!-- AJAX_COMPLETE -->')[0].trim());
                
                if (response.success) {
                    // Remove the row from table
                    if (row) {
                        row.remove();
                    }
                    showAlert('Entry deleted', 'success');
                } else {
                    // Restore row on error
                    if (row) {
                        row.classList.remove('loading');
                    }
                    showAlert(response.message || 'Error deleting entry', 'danger');
                }
            } catch (e) {
                // Fallback for non-JSON response
                if (xhr.responseText.indexOf('deleted') !== -1) {
                    if (row) {
                        row.remove();
                    }
                    showAlert('Entry deleted', 'success');
                } else {
                    if (row) {
                        row.classList.remove('loading');
                    }
                    showAlert('Error deleting entry', 'danger');
                }
            }
        } else {
            if (row) {
                row.classList.remove('loading');
            }
            showAlert('Server error', 'danger');
        }
    };
    
    xhr.onerror = function() {
        if (row) {
            row.classList.remove('loading');
        }
        showAlert('Network error', 'danger');
    };
    
    // Send AJAX request with form data
    var params = 'ajax=true&action=delete&week_key=' + encodeURIComponent(weekDate);
    xhr.send(params);
}

function toggleSpecialOptions() {
    var checkbox = document.getElementById('handle_special');
    var options = document.getElementById('specialOptions');
    
    if (checkbox.checked) {
        options.style.display = 'block';
    } else {
        options.style.display = 'none';
    }
}

function applyFilter() {
    var startDate = document.getElementById('filterStartDate').value;
    var endDate = document.getElementById('filterEndDate').value;
    var fiscalYear = document.getElementById('filterFiscalYear').value;
    
    var rows = document.querySelectorAll('#budgetTableBody tr');
    var visibleCount = 0;
    
    for (var i = 0; i < rows.length; i++) {
        var row = rows[i];
        var dateCell = row.cells[1].textContent;
        var fyCell = row.cells[3].textContent;
        
        var show = true;
        
        // Extract date from cell (handle both single dates and ranges)
        var dateMatch = dateCell.match(/\d{4}-\d{2}-\d{2}/);
        if (dateMatch) {
            var rowDate = dateMatch[0];
            
            // Check date range filter
            if (startDate && rowDate < startDate) show = false;
            if (endDate && rowDate > endDate) show = false;
        }
        
        // Check fiscal year filter
        if (fiscalYear && !fyCell.includes('FY' + fiscalYear)) {
            show = false;
        }
        
        row.style.display = show ? '' : 'none';
        if (show) visibleCount++;
    }
    
    // Update filter status
    var status = document.getElementById('filterStatus');
    if (startDate || endDate || fiscalYear) {
        status.textContent = 'Showing ' + visibleCount + ' of ' + rows.length + ' entries';
    } else {
        status.textContent = '';
    }
}

function clearFilter() {
    document.getElementById('filterStartDate').value = '';
    document.getElementById('filterEndDate').value = '';
    document.getElementById('filterFiscalYear').value = '';
    
    var rows = document.querySelectorAll('#budgetTableBody tr');
    for (var i = 0; i < rows.length; i++) {
        rows[i].style.display = '';
    }
    
    document.getElementById('filterStatus').textContent = '';
}

function handleBulkAddSubmit(event) {
    event.preventDefault();
    
    var startDate = document.getElementById('start_date').value;
    var endDate = document.getElementById('end_date').value;
    var weeklyAmount = document.getElementById('weekly_amount').value;
    var notes = document.getElementById('notes').value;
    
    // Check if special handling is enabled
    var handleSpecial = document.getElementById('handle_special').checked;
    var spanPartial = 'off';
    var handleBoundaries = 'off';
    
    if (handleSpecial) {
        // Check which mode is selected
        var specialMode = document.querySelector('input[name="special_mode"]:checked');
        if (specialMode && specialMode.value === 'extend') {
            spanPartial = 'on';  // Only extend if "extend" mode is selected
        }
        handleBoundaries = document.getElementById('handle_boundaries').checked ? 'on' : 'off';
    }
    
    if (!startDate || !endDate || !weeklyAmount) {
        showAlert('Please fill in all required fields', 'danger');
        return false;
    }
    
    // Build the AJAX URL
    var ajaxUrl = window.location.pathname;
    if (ajaxUrl.indexOf('/PyScript/') !== -1) {
        ajaxUrl = ajaxUrl.replace('/PyScript/', '/PyScriptForm/');
    }
    
    // Show loading state
    event.target.querySelector('button[type="submit"]').disabled = true;
    event.target.querySelector('button[type="submit"]').textContent = 'Processing...';
    
    var xhr = new XMLHttpRequest();
    xhr.open('POST', ajaxUrl, true);
    xhr.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
    
    xhr.onload = function() {
        if (xhr.status === 200) {
            try {
                var response = JSON.parse(xhr.responseText.split('<!-- AJAX_COMPLETE -->')[0].trim());
                
                if (response.success) {
                    showAlert(response.message, 'success');
                    
                    // Reset form
                    document.getElementById('bulkAddForm').reset();
                    document.getElementById('weekly_amount').value = ''' + str(DEFAULT_WEEKLY_BUDGET) + ''';
                    
                    // Refresh the table
                    location.reload(); // Simple reload for now to refresh the table
                } else {
                    showAlert(response.message || 'Error adding entries', 'danger');
                }
            } catch (e) {
                showAlert('Error processing response', 'danger');
            }
        } else {
            showAlert('Server error', 'danger');
        }
        
        // Reset button
        event.target.querySelector('button[type="submit"]').disabled = false;
        event.target.querySelector('button[type="submit"]').textContent = 'Add Budget Entries';
    };
    
    xhr.onerror = function() {
        showAlert('Network error', 'danger');
        event.target.querySelector('button[type="submit"]').disabled = false;
        event.target.querySelector('button[type="submit"]').textContent = 'Add Budget Entries';
    };
    
    var specialMode = 'extend';  // default
    if (handleSpecial) {
        var specialModeInput = document.querySelector('input[name="special_mode"]:checked');
        if (specialModeInput) {
            specialMode = specialModeInput.value;
        }
    }
    
    var params = 'ajax=true&action=bulk_add' +
        '&start_date=' + encodeURIComponent(startDate) +
        '&end_date=' + encodeURIComponent(endDate) +
        '&weekly_amount=' + encodeURIComponent(weeklyAmount) +
        '&notes=' + encodeURIComponent(notes) +
        '&span_partial=' + spanPartial +
        '&handle_boundaries=' + handleBoundaries +
        '&special_mode=' + specialMode;
    xhr.send(params);
    
    return false;
}

function handleImportCsvSubmit(event) {
    event.preventDefault();
    
    var csvData = document.getElementById('csv_data').value;
    
    if (!csvData.trim()) {
        showAlert('Please enter CSV data', 'danger');
        return false;
    }
    
    // Build the AJAX URL
    var ajaxUrl = window.location.pathname;
    if (ajaxUrl.indexOf('/PyScript/') !== -1) {
        ajaxUrl = ajaxUrl.replace('/PyScript/', '/PyScriptForm/');
    }
    
    // Store original CSV data for batch processing
    var originalCsvData = csvData;
    var totalImported = 0;
    var totalBatches = 0;
    
    // Function to process batch
    function processBatch(batchNumber, cumulativeTotal) {
        // Update cumulative total if provided
        if (typeof cumulativeTotal === 'number') {
            totalImported = cumulativeTotal;
        }
        
        totalBatches++;
        
        // Update button text to show progress
        var submitBtn = event.target.querySelector('button[type="submit"]');
        submitBtn.disabled = true;
        submitBtn.textContent = 'Processing batch ' + totalBatches + '... (' + totalImported + ' entries so far)';
        
        var xhr = new XMLHttpRequest();
        xhr.open('POST', ajaxUrl, true);
        xhr.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
        
        xhr.onload = function() {
            if (xhr.status === 200) {
                try {
                    var response = JSON.parse(xhr.responseText.split('<!-- AJAX_COMPLETE -->')[0].trim());
                    
                    if (response.success) {
                        // Add this batch's imports to the total
                        totalImported += response.total_imported || 0;
                        
                        if (response.has_more) {
                            // Show progress message with cumulative total
                            var progressMsg = 'Batch ' + totalBatches + ' complete: ' + response.total_imported + ' entries added. Total so far: ' + totalImported;
                            if (response.lines_remaining) {
                                progressMsg += ' (' + response.lines_remaining + ' CSV lines remaining)';
                            }
                            console.log('Batch ' + totalBatches + ' complete. Next batch: ' + response.next_batch + ', Lines remaining: ' + response.lines_remaining);
                            showAlert(progressMsg, 'success');
                            // Process next batch after a short delay, passing cumulative total
                            setTimeout(function() {
                                processBatch(response.next_batch, totalImported);
                            }, 200);  // Slightly longer delay
                        } else {
                            // All batches complete - use cumulative total
                            showAlert('Import complete! Total entries imported: ' + totalImported + ' in ' + totalBatches + ' batch(es)', 'success');
                            
                            // Clear the CSV data
                            document.getElementById('csv_data').value = '';
                            
                            // Reset button
                            submitBtn.disabled = false;
                            submitBtn.textContent = 'Import Budget Data';
                            
                            // Switch to View Budget tab after successful import
                            setTimeout(function() {
                                showTab('view');
                                location.reload(); // Reload to show the imported data
                            }, 1500);
                        }
                    } else {
                        showAlert(response.message || 'Error importing data', 'danger');
                        submitBtn.disabled = false;
                        submitBtn.textContent = 'Import Budget Data';
                    }
                } catch (e) {
                    showAlert('Error processing response: ' + e.message, 'danger');
                    submitBtn.disabled = false;
                    submitBtn.textContent = 'Import Budget Data';
                }
            } else {
                showAlert('Server error', 'danger');
                submitBtn.disabled = false;
                submitBtn.textContent = 'Import Budget Data';
            }
        };
        
        xhr.onerror = function() {
            showAlert('Network error', 'danger');
            submitBtn.disabled = false;
            submitBtn.textContent = 'Import Budget Data';
        };
        
        var params = 'ajax=true&action=import_csv' +
            '&csv_data=' + encodeURIComponent(originalCsvData) +
            '&batch_number=' + batchNumber;
        xhr.send(params);
    }
    
    // Start processing from batch 0
    processBatch(0);
    
    return false;
}
</script>
'''
    
    # Display the form
    print html
