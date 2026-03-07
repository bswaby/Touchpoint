#roles=Edit
#----------------------------------------------------------------------
# TPxi_CSVPhoneMatcher.py
#
# CSV Phone Matcher for TouchPoint
#
# Takes a CSV list of contact data and matches against TouchPoint records
# by phone number. Phone numbers can be one-to-many (kids share parent
# phone numbers). V1 is match & link only - shows matched people with
# profile links. No record updates.
#
# Created By: Ben Swaby
# Email: bswaby@fbchtn.org
#
# Architecture:
#   Single .py file SPA - Python AJAX handlers for POST, HTML SPA for GET.
#   No persistent state - all data lives in browser session only.
#
# CSS Prefix: cm-
# Root Class: .cm-root
#----------------------------------------------------------------------

import json
import re

model.Header = 'CSV Phone Matcher'

# =====================================================================
# HELPER FUNCTIONS
# =====================================================================

def html_escape(text):
    if not text:
        return ''
    s = str(text)
    s = s.replace('&', '&amp;')
    s = s.replace('<', '&lt;')
    s = s.replace('>', '&gt;')
    s = s.replace('"', '&quot;')
    s = s.replace("'", '&#39;')
    return s

def normalize_phone(phone):
    """Strip to digits only, remove leading country code 1 if 11 digits."""
    if not phone:
        return ''
    digits = re.sub(r'[^0-9]', '', str(phone))
    if len(digits) == 11 and digits[0] == '1':
        digits = digits[1:]
    return digits

# =====================================================================
# AJAX HANDLERS (POST)
# =====================================================================
if model.HttpMethod == "post":
    action = str(Data.action) if hasattr(Data, 'action') and Data.action else ''

    # -----------------------------------------------------------------
    # Match phones - batch lookup
    # -----------------------------------------------------------------
    if action == 'match_phones':
        try:
            phones_json = str(Data.phones_json) if hasattr(Data, 'phones_json') else '[]'
            phones = json.loads(phones_json)

            results = []
            for phone_raw in phones:
                digits = normalize_phone(phone_raw)

                if len(digits) < 7:
                    results.append({'phone': phone_raw, 'matches': [], 'error': 'Too few digits'})
                    continue

                sql = """
                    SELECT p.PeopleId, p.Name2, p.EmailAddress, p.Age,
                        p.CellPhone, p.HomePhone, p.WorkPhone,
                        p.MemberStatusId, ms.Description AS MemberStatus
                    FROM People p
                    LEFT JOIN lookup.MemberStatus ms ON p.MemberStatusId = ms.Id
                    WHERE (p.CellPhone LIKE '%{0}%'
                        OR p.HomePhone LIKE '%{0}%'
                        OR p.WorkPhone LIKE '%{0}%')
                    AND p.DeceasedDate IS NULL
                    AND ISNULL(p.IsDeceased, 0) = 0
                """.format(digits)

                matches = []
                for row in q.QuerySql(sql):
                    matched_fields = []
                    if row.CellPhone and digits in re.sub(r'[^0-9]', '', str(row.CellPhone)):
                        matched_fields.append('Cell')
                    if row.HomePhone and digits in re.sub(r'[^0-9]', '', str(row.HomePhone)):
                        matched_fields.append('Home')
                    if row.WorkPhone and digits in re.sub(r'[^0-9]', '', str(row.WorkPhone)):
                        matched_fields.append('Work')

                    matches.append({
                        'peopleId': row.PeopleId,
                        'name': str(row.Name2 or ''),
                        'email': str(row.EmailAddress or ''),
                        'age': str(row.Age or ''),
                        'memberStatus': str(row.MemberStatus or ''),
                        'cellPhone': model.FmtPhone(row.CellPhone, '') if row.CellPhone else '',
                        'homePhone': model.FmtPhone(row.HomePhone, '') if row.HomePhone else '',
                        'workPhone': model.FmtPhone(row.WorkPhone, '') if row.WorkPhone else '',
                        'matchedFields': matched_fields
                    })

                results.append({'phone': phone_raw, 'matches': matches})

            print json.dumps({'success': True, 'results': results})
        except Exception as e:
            print json.dumps({'success': False, 'message': str(e)})

    else:
        print json.dumps({'success': False, 'message': 'Unknown action: ' + action})

# =====================================================================
# GET - Render the SPA HTML
# =====================================================================
else:
    html = '''<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<style>
/* =====================================================================
   CSS Phone Matcher Styles (cm- prefix)
   ===================================================================== */
.cm-root {
    --cm-primary: #2563eb;
    --cm-primary-light: #dbeafe;
    --cm-success: #059669;
    --cm-success-light: #d1fae5;
    --cm-warning: #d97706;
    --cm-warning-light: #fef3c7;
    --cm-danger: #dc2626;
    --cm-danger-light: #fee2e2;
    --cm-dark: #1e293b;
    --cm-gray: #64748b;
    --cm-light-bg: #f8fafc;
    --cm-border: #e2e8f0;
    --cm-radius: 12px;
    --cm-shadow: 0 1px 3px rgba(0,0,0,0.1), 0 1px 2px rgba(0,0,0,0.06);
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
    color: var(--cm-dark);
    max-width: 1200px;
    margin: 0 auto;
    padding: 12px;
}

/* ---- Common UI ---- */
.cm-btn {
    display: inline-flex; align-items: center; justify-content: center; gap: 8px;
    padding: 10px 20px; border-radius: 8px; border: none; font-size: 15px; font-weight: 600;
    cursor: pointer; transition: all 0.15s; user-select: none;
}
.cm-btn:active { transform: scale(0.97); }
.cm-btn-primary { background: var(--cm-primary); color: white; }
.cm-btn-success { background: var(--cm-success); color: white; }
.cm-btn-outline { background: white; color: var(--cm-dark); border: 2px solid var(--cm-border); }
.cm-btn-sm { padding: 6px 14px; font-size: 13px; }
.cm-btn:disabled { opacity: 0.5; cursor: not-allowed; }

.cm-select {
    width: 100%; padding: 10px 14px; border: 2px solid var(--cm-border); border-radius: 8px;
    font-size: 15px; background: white; box-sizing: border-box;
}

.cm-textarea {
    width: 100%; padding: 10px 14px; border: 2px solid var(--cm-border); border-radius: 8px;
    font-size: 14px; outline: none; box-sizing: border-box; resize: vertical; font-family: monospace;
}
.cm-textarea:focus { border-color: var(--cm-primary); }

.cm-label { display: block; font-size: 13px; font-weight: 600; color: var(--cm-gray); margin-bottom: 6px; text-transform: uppercase; letter-spacing: 0.5px; }

.cm-panel {
    background: white; border-radius: var(--cm-radius); border: 1px solid var(--cm-border);
    box-shadow: var(--cm-shadow); margin-bottom: 16px; overflow: hidden;
}
.cm-panel-header {
    display: flex; align-items: center; justify-content: space-between; padding: 16px 20px;
    background: var(--cm-light-bg); border-bottom: 1px solid var(--cm-border); font-weight: 600; font-size: 16px;
}
.cm-panel-body { padding: 20px; }

.cm-back-btn {
    display: inline-flex; align-items: center; gap: 6px; padding: 8px 16px;
    color: var(--cm-primary); font-size: 16px; font-weight: 600; cursor: pointer;
    border: none; background: none;
}

.cm-step-header { margin-bottom: 20px; }
.cm-step-header h2 { margin: 8px 0 4px; font-size: 24px; }
.cm-text-muted { color: var(--cm-gray); }
.cm-text-sm { font-size: 14px; }

/* ---- Preview Table ---- */
.cm-table { width: 100%; border-collapse: collapse; font-size: 14px; }
.cm-table th, .cm-table td { padding: 8px 12px; border: 1px solid var(--cm-border); text-align: left; }
.cm-table th { background: var(--cm-light-bg); font-weight: 600; font-size: 13px; text-transform: uppercase; letter-spacing: 0.3px; }
.cm-highlight-col { background: var(--cm-primary-light) !important; }

/* ---- Progress ---- */
.cm-progress-label { font-size: 14px; font-weight: 600; margin-bottom: 8px; color: var(--cm-dark); }
.cm-progress-bar { height: 8px; background: #e2e8f0; border-radius: 4px; overflow: hidden; }
.cm-progress-fill { height: 100%; background: var(--cm-primary); border-radius: 4px; transition: width 0.3s; }

/* ---- Summary Bar ---- */
.cm-summary {
    display: flex; gap: 16px; flex-wrap: wrap; margin-bottom: 16px;
    padding: 16px 20px; background: white; border-radius: var(--cm-radius);
    border: 1px solid var(--cm-border); box-shadow: var(--cm-shadow);
}
.cm-summary-item { font-size: 15px; padding: 4px 0; }
.cm-summary-item strong { font-size: 18px; }
.cm-match-single strong { color: var(--cm-success); }
.cm-match-multi strong { color: var(--cm-warning); }
.cm-match-none strong { color: var(--cm-danger); }

/* ---- Filters ---- */
.cm-filters { display: flex; gap: 8px; flex-wrap: wrap; margin-bottom: 16px; align-items: center; }
.cm-filter-btn {
    padding: 6px 16px; border-radius: 20px; border: 2px solid var(--cm-border);
    background: white; font-size: 14px; font-weight: 500; cursor: pointer; transition: all 0.15s;
}
.cm-filter-btn:hover { border-color: var(--cm-primary); }
.cm-filter-active { background: var(--cm-primary); color: white; border-color: var(--cm-primary); }

/* ---- Result Cards ---- */
.cm-results { display: flex; flex-direction: column; gap: 8px; }
.cm-result-card {
    display: flex; gap: 20px; padding: 16px 20px;
    background: white; border-radius: var(--cm-radius); border: 1px solid var(--cm-border);
    box-shadow: var(--cm-shadow); border-left: 4px solid transparent;
}
.cm-row-single { border-left-color: var(--cm-success); }
.cm-row-multi { border-left-color: var(--cm-warning); }
.cm-row-none { border-left-color: var(--cm-danger); }

.cm-result-csv { flex: 1; min-width: 200px; }
.cm-result-csv-phone { font-size: 16px; font-weight: 700; margin-bottom: 4px; color: var(--cm-dark); }
.cm-result-csv-data { display: flex; flex-wrap: wrap; gap: 4px 12px; }
.cm-csv-field { font-size: 13px; color: var(--cm-gray); }
.cm-csv-label { font-weight: 600; }

.cm-result-matches { flex: 1; min-width: 250px; }
.cm-no-match { color: var(--cm-danger); font-weight: 600; font-size: 14px; padding: 8px 0; }

.cm-match-person { padding: 8px 0; border-bottom: 1px solid var(--cm-border); }
.cm-match-person:last-child { border-bottom: none; }
.cm-person-name { font-size: 15px; font-weight: 700; color: var(--cm-primary); text-decoration: none; }
.cm-person-name:hover { text-decoration: underline; }

.cm-person-details { display: flex; flex-wrap: wrap; gap: 4px 10px; margin-top: 4px; }
.cm-detail { font-size: 13px; color: var(--cm-gray); }

.cm-match-badges { display: flex; flex-wrap: wrap; gap: 6px; margin-top: 6px; align-items: center; }
.cm-badge {
    display: inline-block; padding: 2px 8px; border-radius: 4px;
    font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.5px;
}
.cm-badge-cell { background: #dcfce7; color: #166534; }
.cm-badge-home { background: #dbeafe; color: #1e40af; }
.cm-badge-work { background: #fef3c7; color: #92400e; }
.cm-phone-info { font-size: 12px; color: var(--cm-gray); }

/* ---- Toast ---- */
.cm-toast-container { position: fixed; top: 20px; right: 20px; z-index: 10000; display: flex; flex-direction: column; gap: 8px; }
.cm-toast {
    padding: 12px 20px; border-radius: 8px; font-size: 14px; font-weight: 500;
    opacity: 0; transform: translateX(40px); transition: all 0.3s;
    box-shadow: 0 4px 12px rgba(0,0,0,0.15); max-width: 400px;
}
.cm-toast.cm-show { opacity: 1; transform: translateX(0); }
.cm-toast-success { background: #059669; color: white; }
.cm-toast-danger { background: #dc2626; color: white; }
.cm-toast-info { background: #2563eb; color: white; }

/* ---- Loading ---- */
.cm-loading { display: flex; flex-direction: column; align-items: center; justify-content: center; padding: 60px 20px; }
.cm-spinner {
    width: 40px; height: 40px; border: 4px solid #e2e8f0; border-top: 4px solid var(--cm-primary);
    border-radius: 50%; animation: cm-spin 0.8s linear infinite;
}
@keyframes cm-spin { to { transform: rotate(360deg); } }

/* ---- Responsive ---- */
@media (max-width: 768px) {
    .cm-result-card { flex-direction: column; }
    .cm-summary { flex-direction: column; gap: 8px; }
    .cm-filters { flex-direction: column; }
}
</style>
</head>
<body>
<div class="cm-root" id="cmApp">
    <div class="cm-toast-container" id="cmToastContainer"></div>
    <div id="cmContent">
        <div class="cm-loading"><div class="cm-spinner"></div><p>Loading...</p></div>
    </div>
</div>

<script>
(function() {
    "use strict";

    // =====================================================================
    // STATE
    // =====================================================================
    var state = {
        step: 1,
        csvText: '',
        headers: [],
        rows: [],
        phoneColIndex: -1,
        matchedRows: [],
        batchProgress: 0,
        batchTotal: 0,
        matching: false,
        filter: 'all'
    };

    var scriptPath = (function() {
        var p = window.location.pathname;
        if (p.indexOf('/PyScriptForm/') > -1) return p;
        return p.replace('/PyScript/', '/PyScriptForm/');
    })();

    // =====================================================================
    // UTILITIES
    // =====================================================================
    function escHtml(s) {
        if (!s) return '';
        var d = document.createElement('div');
        d.appendChild(document.createTextNode(s));
        return d.innerHTML;
    }

    function extractJson(text) {
        text = (text || '').trim();
        var start = text.indexOf('{');
        var end = text.lastIndexOf('}');
        if (start === -1 || end === -1) {
            start = text.indexOf('[');
            end = text.lastIndexOf(']');
        }
        if (start >= 0 && end > start) {
            return text.substring(start, end + 1);
        }
        return text;
    }

    function ajax(action, params, callback) {
        var data = 'action=' + encodeURIComponent(action);
        if (params) {
            for (var key in params) {
                if (params.hasOwnProperty(key)) {
                    data += '&' + encodeURIComponent(key) + '=' + encodeURIComponent(params[key]);
                }
            }
        }
        var xhr = new XMLHttpRequest();
        xhr.open('POST', scriptPath, true);
        xhr.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
        xhr.onreadystatechange = function() {
            if (xhr.readyState === 4) {
                if (xhr.status === 200) {
                    try {
                        var json = JSON.parse(extractJson(xhr.responseText));
                        callback(null, json);
                    } catch (e) {
                        callback('Invalid response: ' + e.message, null);
                    }
                } else {
                    callback('HTTP ' + xhr.status, null);
                }
            }
        };
        xhr.send(data);
    }

    function showToast(msg, type) {
        type = type || 'info';
        var container = document.getElementById('cmToastContainer');
        var toast = document.createElement('div');
        toast.className = 'cm-toast cm-toast-' + type;
        toast.textContent = msg;
        container.appendChild(toast);
        setTimeout(function() { toast.classList.add('cm-show'); }, 10);
        setTimeout(function() {
            toast.classList.remove('cm-show');
            setTimeout(function() { container.removeChild(toast); }, 300);
        }, 3000);
    }

    // =====================================================================
    // CSV PARSER (client-side)
    // =====================================================================
    function parseCsvLine(line) {
        var fields = [];
        var current = '';
        var inQuotes = false;
        for (var i = 0; i < line.length; i++) {
            var c = line[i];
            if (c === '"') {
                if (inQuotes && i + 1 < line.length && line[i + 1] === '"') {
                    current += '"';
                    i++;
                } else {
                    inQuotes = !inQuotes;
                }
            } else if (c === ',' && !inQuotes) {
                fields.push(current.trim());
                current = '';
            } else {
                current += c;
            }
        }
        fields.push(current.trim());
        return fields;
    }

    function parseCSV(text) {
        var lines = [];
        var current = '';
        var inQuotes = false;
        for (var i = 0; i < text.length; i++) {
            var c = text[i];
            if (c === '"') {
                if (inQuotes && i + 1 < text.length && text[i + 1] === '"') {
                    current += '"';
                    i++;
                } else {
                    inQuotes = !inQuotes;
                }
                current += c;
            } else if ((c === '\\n' || c === '\\r') && !inQuotes) {
                if (c === '\\r' && i + 1 < text.length && text[i + 1] === '\\n') {
                    i++;
                }
                if (current.trim() !== '') {
                    lines.push(current);
                }
                current = '';
            } else {
                current += c;
            }
        }
        if (current.trim() !== '') {
            lines.push(current);
        }

        if (lines.length === 0) return { headers: [], rows: [] };

        var headers = parseCsvLine(lines[0]);
        var rows = [];
        for (var r = 1; r < lines.length; r++) {
            var fields = parseCsvLine(lines[r]);
            var row = {};
            for (var ci = 0; ci < headers.length; ci++) {
                row[headers[ci]] = fields[ci] || '';
            }
            rows.push(row);
        }
        return { headers: headers, rows: rows };
    }

    // =====================================================================
    // RENDER
    // =====================================================================
    function render() {
        var el = document.getElementById('cmContent');
        if (state.step === 1) renderStep1(el);
        else if (state.step === 2) renderStep2(el);
        else if (state.step === 3) renderStep3(el);
    }

    // -----------------------------------------------------------------
    // Step 1: Paste CSV
    // -----------------------------------------------------------------
    function renderStep1(el) {
        var h = '';
        h += '<div class="cm-step-header">';
        h += '<h2>Step 1: Paste Your CSV Data</h2>';
        h += '<p class="cm-text-muted">Paste CSV data below. The first row should contain column headers.</p>';
        h += '</div>';
        h += '<div class="cm-panel">';
        h += '<div class="cm-panel-body">';
        h += '<textarea class="cm-textarea" id="cmCsvInput" rows="12" placeholder="Paste CSV data here...&#10;&#10;Example:&#10;Name,Phone,Email&#10;John Doe,615-555-1234,john@example.com">';
        h += escHtml(state.csvText);
        h += '</textarea>';
        h += '<div style="margin-top:16px;display:flex;gap:12px;align-items:center;">';
        h += '<button class="cm-btn cm-btn-primary" onclick="CMApp.parseCSV()">Parse CSV</button>';
        h += '</div>';
        h += '</div>';
        h += '</div>';
        el.innerHTML = h;
    }

    // -----------------------------------------------------------------
    // Step 2: Select Phone Column
    // -----------------------------------------------------------------
    function renderStep2(el) {
        var h = '';
        h += '<div class="cm-step-header">';
        h += '<button class="cm-back-btn" onclick="CMApp.goStep(1)">&#8592; Back</button>';
        h += '<h2>Step 2: Select Phone Column</h2>';
        h += '<p class="cm-text-muted">Choose which column contains phone numbers. Found ' + state.rows.length + ' data rows.</p>';
        h += '</div>';

        h += '<div class="cm-panel">';
        h += '<div class="cm-panel-body">';
        h += '<label class="cm-label">Phone Number Column</label>';
        h += '<select class="cm-select" id="cmPhoneCol" onchange="CMApp.onPhoneColChange()">';
        for (var i = 0; i < state.headers.length; i++) {
            var sel = i === state.phoneColIndex ? ' selected' : '';
            h += '<option value="' + i + '"' + sel + '>' + escHtml(state.headers[i]) + '</option>';
        }
        h += '</select>';
        h += '</div>';
        h += '</div>';

        // Preview table
        h += '<div class="cm-panel">';
        h += '<div class="cm-panel-header">Preview (first 3 rows)</div>';
        h += '<div class="cm-panel-body" style="overflow-x:auto;">';
        h += '<table class="cm-table">';
        h += '<thead><tr>';
        for (var i = 0; i < state.headers.length; i++) {
            var cls = i === state.phoneColIndex ? ' class="cm-highlight-col"' : '';
            h += '<th' + cls + '>' + escHtml(state.headers[i]) + '</th>';
        }
        h += '</tr></thead>';
        h += '<tbody>';
        var previewCount = Math.min(3, state.rows.length);
        for (var r = 0; r < previewCount; r++) {
            h += '<tr>';
            for (var c = 0; c < state.headers.length; c++) {
                var cls = c === state.phoneColIndex ? ' class="cm-highlight-col"' : '';
                h += '<td' + cls + '>' + escHtml(state.rows[r][state.headers[c]] || '') + '</td>';
            }
            h += '</tr>';
        }
        h += '</tbody></table>';
        h += '</div>';
        h += '</div>';

        h += '<div style="margin-top:16px;">';
        h += '<button class="cm-btn cm-btn-primary" onclick="CMApp.startMatching()">Match Phone Numbers</button>';
        h += '</div>';

        el.innerHTML = h;
    }

    // -----------------------------------------------------------------
    // Step 3: Results
    // -----------------------------------------------------------------
    function renderStep3(el) {
        var h = '';
        var total = state.matchedRows.length;
        var matched = 0, multiple = 0, none = 0;
        for (var i = 0; i < state.matchedRows.length; i++) {
            var mc = state.matchedRows[i].matches.length;
            if (mc === 0) none++;
            else if (mc === 1) matched++;
            else multiple++;
        }

        h += '<div class="cm-step-header">';
        h += '<button class="cm-back-btn" onclick="CMApp.goStep(2)">&#8592; Back</button>';
        h += '<h2>Step 3: Results</h2>';
        h += '</div>';

        // Progress bar (during matching)
        if (state.matching) {
            var pct = state.batchTotal > 0 ? Math.round((state.batchProgress / state.batchTotal) * 100) : 0;
            h += '<div class="cm-panel"><div class="cm-panel-body">';
            h += '<div class="cm-progress-label">Matching phones... ' + state.batchProgress + ' / ' + state.batchTotal + ' (' + pct + '%)</div>';
            h += '<div class="cm-progress-bar"><div class="cm-progress-fill" style="width:' + pct + '%"></div></div>';
            h += '</div></div>';
        }

        // Summary
        h += '<div class="cm-summary">';
        h += '<div class="cm-summary-item">Total: <strong>' + total + '</strong></div>';
        h += '<div class="cm-summary-item cm-match-single">Matched: <strong>' + matched + '</strong></div>';
        h += '<div class="cm-summary-item cm-match-multi">Multiple: <strong>' + multiple + '</strong></div>';
        h += '<div class="cm-summary-item cm-match-none">No Match: <strong>' + none + '</strong></div>';
        h += '</div>';

        // Filter buttons
        h += '<div class="cm-filters">';
        var filters = [['all','All'],['matched','Matched'],['multiple','Multiple'],['none','No Match']];
        for (var f = 0; f < filters.length; f++) {
            var active = state.filter === filters[f][0] ? ' cm-filter-active' : '';
            h += '<button class="cm-filter-btn' + active + '" onclick="CMApp.setFilter(\\'' + filters[f][0] + '\\')">' + filters[f][1] + '</button>';
        }
        h += '<button class="cm-btn cm-btn-outline cm-btn-sm" onclick="CMApp.exportCSV()" style="margin-left:auto;">Export CSV</button>';
        h += '</div>';

        // Results
        h += '<div class="cm-results">';
        for (var i = 0; i < state.matchedRows.length; i++) {
            var mr = state.matchedRows[i];
            var mc = mr.matches.length;

            // Apply filter
            if (state.filter === 'matched' && mc !== 1) continue;
            if (state.filter === 'multiple' && mc <= 1) continue;
            if (state.filter === 'none' && mc !== 0) continue;

            var rowClass = mc === 0 ? 'cm-row-none' : (mc === 1 ? 'cm-row-single' : 'cm-row-multi');

            h += '<div class="cm-result-card ' + rowClass + '">';

            // CSV data (left)
            h += '<div class="cm-result-csv">';
            h += '<div class="cm-result-csv-phone">' + escHtml(mr.phone) + '</div>';
            h += '<div class="cm-result-csv-data">';
            for (var c = 0; c < state.headers.length; c++) {
                if (c === state.phoneColIndex) continue;
                var val = mr.csvRow[state.headers[c]] || '';
                if (val) {
                    h += '<span class="cm-csv-field"><span class="cm-csv-label">' + escHtml(state.headers[c]) + ':</span> ' + escHtml(val) + '</span>';
                }
            }
            h += '</div>';
            h += '</div>';

            // Matches (right)
            h += '<div class="cm-result-matches">';
            if (mc === 0) {
                h += '<div class="cm-no-match">No match found</div>';
            } else {
                for (var m = 0; m < mc; m++) {
                    var match = mr.matches[m];
                    h += '<div class="cm-match-person">';
                    h += '<a href="/Person2/' + match.peopleId + '" target="_blank" class="cm-person-name">' + escHtml(match.name) + '</a>';
                    h += '<div class="cm-person-details">';
                    if (match.age) h += '<span class="cm-detail">Age ' + escHtml(match.age) + '</span>';
                    if (match.memberStatus) h += '<span class="cm-detail">' + escHtml(match.memberStatus) + '</span>';
                    if (match.email) h += '<span class="cm-detail">' + escHtml(match.email) + '</span>';
                    h += '</div>';
                    h += '<div class="cm-match-badges">';
                    for (var b = 0; b < match.matchedFields.length; b++) {
                        h += '<span class="cm-badge cm-badge-' + match.matchedFields[b].toLowerCase() + '">' + match.matchedFields[b] + '</span>';
                    }
                    if (match.cellPhone) h += '<span class="cm-phone-info">Cell: ' + escHtml(match.cellPhone) + '</span>';
                    if (match.homePhone) h += '<span class="cm-phone-info">Home: ' + escHtml(match.homePhone) + '</span>';
                    if (match.workPhone) h += '<span class="cm-phone-info">Work: ' + escHtml(match.workPhone) + '</span>';
                    h += '</div>';
                    h += '</div>';
                }
            }
            h += '</div>';

            h += '</div>';
        }
        h += '</div>';

        el.innerHTML = h;
    }

    // =====================================================================
    // ACTIONS
    // =====================================================================
    function doParse() {
        var textarea = document.getElementById('cmCsvInput');
        state.csvText = textarea.value;
        if (!state.csvText.trim()) {
            showToast('Please paste CSV data first', 'danger');
            return;
        }

        var result = parseCSV(state.csvText);
        if (result.headers.length === 0) {
            showToast('Could not parse CSV headers', 'danger');
            return;
        }
        if (result.rows.length === 0) {
            // Check if first row looks like data (no header row)
            // Heuristic: if any field in "headers" contains digits, it is likely data
            var looksLikeData = false;
            for (var di = 0; di < result.headers.length; di++) {
                if (/\\d/.test(result.headers[di])) { looksLikeData = true; break; }
            }
            if (looksLikeData) {
                // Treat first line as data, generate generic column headers
                var genHeaders = [];
                for (var gi = 0; gi < result.headers.length; gi++) {
                    genHeaders.push('Column ' + (gi + 1));
                }
                var singleRow = {};
                for (var si = 0; si < genHeaders.length; si++) {
                    singleRow[genHeaders[si]] = result.headers[si];
                }
                result.headers = genHeaders;
                result.rows = [singleRow];
            } else {
                showToast('No data rows found. Make sure your CSV has a header row and at least one data row.', 'danger');
                return;
            }
        }

        state.headers = result.headers;
        state.rows = result.rows;

        // Auto-detect phone column
        state.phoneColIndex = 0;
        for (var i = 0; i < state.headers.length; i++) {
            if (state.headers[i].toLowerCase().indexOf('phone') > -1) {
                state.phoneColIndex = i;
                break;
            }
        }

        state.step = 2;
        render();
        showToast('Parsed ' + state.rows.length + ' rows with ' + state.headers.length + ' columns', 'success');
    }

    function startMatching() {
        state.matchedRows = [];
        state.batchProgress = 0;
        state.batchTotal = state.rows.length;
        state.matching = true;
        state.step = 3;
        state.filter = 'all';
        render();

        var batchSize = 10;
        var batches = [];
        for (var i = 0; i < state.rows.length; i += batchSize) {
            var batch = [];
            for (var j = i; j < Math.min(i + batchSize, state.rows.length); j++) {
                batch.push({
                    index: j,
                    phone: state.rows[j][state.headers[state.phoneColIndex]] || '',
                    csvRow: state.rows[j]
                });
            }
            batches.push(batch);
        }

        processBatch(batches, 0);
    }

    function processBatch(batches, batchIndex) {
        if (batchIndex >= batches.length) {
            state.matching = false;
            state.batchProgress = state.batchTotal;
            render();
            showToast('Matching complete!', 'success');
            return;
        }

        var batch = batches[batchIndex];
        var phones = [];
        for (var i = 0; i < batch.length; i++) {
            phones.push(batch[i].phone);
        }

        ajax('match_phones', { phones_json: JSON.stringify(phones) }, function(err, data) {
            if (err || !data || !data.success) {
                showToast('Error matching batch: ' + (data ? data.message : err), 'danger');
                state.matching = false;
                render();
                return;
            }

            for (var i = 0; i < data.results.length; i++) {
                state.matchedRows.push({
                    csvRow: batch[i].csvRow,
                    phone: batch[i].phone,
                    matches: data.results[i].matches || []
                });
            }

            state.batchProgress += batch.length;
            render();

            setTimeout(function() { processBatch(batches, batchIndex + 1); }, 50);
        });
    }

    function exportCSV() {
        var lines = [];
        var exportHeaders = state.headers.slice();
        exportHeaders.push('MatchCount', 'PeopleIds', 'Names', 'ProfileLinks');
        lines.push(exportHeaders.map(function(h) { return '"' + h.replace(/"/g, '""') + '"'; }).join(','));

        for (var i = 0; i < state.matchedRows.length; i++) {
            var mr = state.matchedRows[i];
            var row = [];
            for (var c = 0; c < state.headers.length; c++) {
                var val = mr.csvRow[state.headers[c]] || '';
                row.push('"' + val.replace(/"/g, '""') + '"');
            }
            row.push(mr.matches.length);
            var ids = [], names = [], links = [];
            for (var m = 0; m < mr.matches.length; m++) {
                ids.push(mr.matches[m].peopleId);
                names.push(mr.matches[m].name);
                links.push('/Person2/' + mr.matches[m].peopleId);
            }
            row.push('"' + ids.join('; ') + '"');
            row.push('"' + names.join('; ').replace(/"/g, '""') + '"');
            row.push('"' + links.join('; ') + '"');
            lines.push(row.join(','));
        }

        var csv = lines.join('\\n');
        var blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
        var link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = 'phone_match_results.csv';
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    }

    // =====================================================================
    // PUBLIC API
    // =====================================================================
    window.CMApp = {
        parseCSV: doParse,
        goStep: function(step) { state.step = step; render(); },
        onPhoneColChange: function() {
            var sel = document.getElementById('cmPhoneCol');
            state.phoneColIndex = parseInt(sel.value);
            render();
        },
        startMatching: startMatching,
        setFilter: function(f) { state.filter = f; render(); },
        exportCSV: exportCSV
    };

    // =====================================================================
    // INIT
    // =====================================================================
    render();

})();
</script>
</div>
</body>
</html>'''

    model.Form = html
