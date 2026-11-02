#####################################################################
#### TouchPoint API Explorer 
#
# written by: Ben Swaby
# email: bswaby@fbchtn.org
#
# Warning: code run here changes your live TouchPoint data. Know
#          what you are doing before using it.
#
# Version: 2.0
#
# RELEASE NOTES
# 2.0 (November 2026)
# - Only the Admin or Developer role can use the page or run code.
#   TouchPoint's /PyScriptForm address ignores a #Roles= line and lets
#   any signed-in person in, so the page now checks for itself.
# - Every documented call was reviewed and corrected: 347 entries with
#   correct parameters, every form of each call, and a runnable
#   example. Removed 8 that do not
#   exist (for example model.AddOrganizationMember, model.UserId,
#   q.QuerySqlList), 30 duplicates, and added 32 that were missing.
# - Tabs: several code tabs, double-click to rename, and they come back
#   after a refresh. Use Example, Discover and Quick Test open their own
#   tab instead of replacing your code.
# - History and snippets: the last 50 runs and named snippets, kept in
#   your browser.
# - A warning before running code that changes data or sends email.
# - Output: shows run time, formats JSON, can show printed HTML in a
#   safe frame, and has a Copy button. Each tab keeps its own output.
# - A run now returns only its output instead of rebuilding the page.
# - Discover no longer calls the method it inspects (calling can change
#   data).
# - Autocomplete knows properties from methods (model.UserPeopleId no
#   longer gets "("), works on sub objects such as model.docusign, and
#   shows each call's parameters.
# - The linter checks every call against the docs: calls that do not
#   exist (with a "did you mean"), a property called like a method, too
#   few or too many arguments (including calls split over lines),
#   deprecated calls with their replacement, internal calls, and known
#   traps: hasattr(model.Data, ...) is always True, model.DebugPrint
#   shows nothing, model.UpdateStatusFlag fails, and calls that delete
#   or write across the whole database. Comments and strings are ignored,
#   so "http://" or href="..." no longer raise false alarms, and the wrong
#   "// is not in Python 2.7" warning is gone.
# - Code you run can use Data (same as model.Data), like a real script.
# - Code containing HTML (such as "<div>") now runs. TouchPoint used to
#   reject it as "a potentially dangerous Request.Form value"; the page
#   now encodes the code before sending it.
# - Examples were run on a test site and corrected: they use the signed
#   in person (model.UserPeopleId) instead of a fixed PeopleId, say what
#   they found when there is no data, and handle missing content.
# - Internal members of model and q are grouped with a note on why they
#   are not for scripts and what to use instead.
# - New reference layout for each call: labels (read only, writes data,
#   sends email or text, slow, deprecated, known issue, outside service),
#   a copyable call line, a one line summary with more detail folded,
#   a parameters table with required and default, a "Watch out" box,
#   Run / Open in new tab / Copy on the example, and See also links.
# - 16 clear categories instead of 27. Deprecated calls are struck
#   through in the list. Search covers descriptions, not just names.
# - Categories start collapsed (a search opens the matching ones), the
#   documentation panel starts wider, both side panels resize in the
#   direction you drag, and the widths you set are remembered.
# - A "How it works" section for calls with several steps, starting with
#   model.AddBackgroundCheck (Protect My Ministry only: package lookup,
#   submit types, statuses, and how results come back).
# - Every call was traced through TouchPoint: 143 calls now have a "How
#   it works" section, and many corrections and Watch out items were
#   added. Notable: @UserPeopleId and @CurrentOrgId in q.QuerySql come
#   back blank in scripts (with a fix), GetTasksNotes ignores its person
#   filter, StatusCount matches the flag's saved search name, and
#   UnArchiveRecords needs ArchivedFlag = 1[True] in its query.
#####################################################################

import datetime
import json

model.Header = "TouchPoint API Explorer"

# Who may use this page. TouchPoint's /PyScriptForm address ignores the
# #Roles= line and lets any signed-in person in (a member who only views
# their own giving, too), so the page has to check for itself.
ALLOWED_ROLES = ("Admin", "Developer")

def _user_allowed():
    try:
        for role in ALLOWED_ROLES:
            if model.UserIsInRole(role):
                return True
    except Exception:
        pass
    return False

_allowed = _user_allowed()

# model.Data hands back '' for a field that was not sent (hasattr is always
# True), so a run is a request that actually carries code.
try:
    _enc = model.Data.code_enc
except Exception:
    _enc = None
if _enc:
    # The page sends the code percent-encoded twice. TouchPoint decodes it
    # once; this undoes the second, so text such as "<div" never reaches
    # TouchPoint's "potentially dangerous Request.Form value" check.
    import System
    _code = System.Uri.UnescapeDataString(str(_enc))
else:
    try:
        _code = model.Data.code
    except Exception:
        _code = None
_is_run = bool(_code) and isinstance(_code, basestring) and bool(_code.strip())

if _is_run and not _allowed:
    print "Not allowed: this page is for the Admin or Developer role."
    print "\n<!-- AJAX_COMPLETE -->"

elif _is_run:
    # Run the posted code and send back only what it printed
    import sys
    from StringIO import StringIO
    _old_stdout = sys.stdout
    sys.stdout = StringIO()
    _started = datetime.datetime.now()
    try:
        exec(_code, {'datetime': datetime, 'model': model, 'q': q, 'Data': model.Data})
        _output = sys.stdout.getvalue()
    except Exception as e:
        _output = sys.stdout.getvalue() + "\nError: " + str(e)
    finally:
        sys.stdout = _old_stdout
    _ms = int((datetime.datetime.now() - _started).total_seconds() * 1000)
    print _output
    print "\n<!-- AJAX_COMPLETE -->"
    print "<!-- AJAX_MS:" + str(_ms) + " -->"

elif not _allowed:
    model.Form = '<div class="alert alert-danger">This page is for the Admin or Developer role.</div>'

# COMPLETE TouchPoint API Documentation from https://docs.touchpointsoftware.com
method_docs = {
    "model.Content": {
        "category": "Content & Files",
        "kind": "method",
        "arity": [1, 1],
        "tags": ["read_only"],
        "summary": "Returns the body of a Special Content item by name, whatever its type (HTML, text, SQL or Python).",
        "details": "Use model.TextContent or model.HtmlContent when you know the type.",
        "params": [
            ["name", "string", "Special Content name", True, None]
        ],
        "returns": "string",
        "watch_out": ["Raises an error if no content has that name.", "If items of different types share the name, you get whichever one the database returns first.", "For an Unlayer email or draft, the body is the stored JSON (design plus html), not plain HTML."],
        "example": "body = model.Content(\"CustomReports\")\nprint len(body), \"characters\"",
        "see_also": ["model.TextContent", "model.HtmlContent", "model.SqlContent", "model.PythonContent"]
    },
    "model.ContentForDate": {
        "category": "Content & Files",
        "kind": "method",
        "arity": [2, 2],
        "tags": ["read_only"],
        "summary": "Returns one dated section from an HTML Special Content file.",
        "details": "The file is split on headings written as <h1>m/d/yyyy======</h1>, and the text after the heading that matches the date is returned. Useful for a daily devotional stored in one file.",
        "params": [
            ["contentName", "string", "Name of the HTML Special Content file", True, None],
            ["date", "date", "Date to look for. A DateTime or a date string such as \"11/2/2026\"", True, None]
        ],
        "returns": "string",
        "watch_out": ["Returns \"no date\" if the date cannot be read.", "Returns \"cannot find email content\" if no heading matches.", "Raises an error if the content file does not exist.", "Looks for an HTML item first, then a Text item with the same name.", "Headings may use slashes or dashes and a 2 or 4 digit year, such as <h1>11-2-26===</h1>. An ISO date such as 2026-11-02 in the heading is not recognized.", "The time part of the date you pass is ignored, so a DateTime with a time still matches."],
        "example": "import datetime\ntry:\n    print model.ContentForDate(\"DailyDevotional\", datetime.datetime.now())\nexcept Exception as e:\n    # Raises when the named content does not exist\n    print \"No content named DailyDevotional:\", e",
        "see_also": ["model.HtmlContent", "model.EmailContentWithSubject"]
    },
    "model.DeleteFile": {
        "category": "Content & Files",
        "kind": "method",
        "arity": [1, 1],
        "tags": ["writes_data"],
        "summary": "Deletes a file on the server.",
        "params": [
            ["path", "string", "File path", True, None]
        ],
        "returns": "None",
        "watch_out": ["Works only on a developer's local debug build.", "On a live TouchPoint site it always raises an error.", "The error raised has no message, so it is hard to tell what went wrong."],
        "example": "try:\n    model.DeleteFile(\"debug.txt\")\n    print \"deleted\"\nexcept Exception as e:\n    print \"Not available on this server\"",
        "see_also": ["model.ReadFile", "model.WriteFile", "model.IsDebug"]
    },
    "model.Draft": {
        "category": "Content & Files",
        "kind": "method",
        "arity": [1, 1],
        "tags": ["read_only"],
        "summary": "Returns the stored body of a saved email draft (classic or Unlayer).",
        "params": [
            ["name", "string", "Saved draft name", True, None]
        ],
        "returns": "string: HTML for a classic draft, the stored JSON (design plus html) for an Unlayer draft",
        "watch_out": ["Raises an error if no draft has that name.", "Matches both classic saved drafts and Unlayer saved drafts with that name."],
        "example": "for d in q.QuerySql(\"SELECT TOP 5 Name FROM Content WHERE TypeID IN (3, 6) ORDER BY DateCreated DESC\"):\n    print d.Name\n# body = model.Draft(\"Welcome Email\")",
        "see_also": ["model.DraftTitle", "model.Content"]
    },
    "model.DraftTitle": {
        "category": "Content & Files",
        "kind": "method",
        "arity": [1, 1],
        "tags": ["read_only"],
        "summary": "Returns the subject (title) of a saved email draft (classic or Unlayer).",
        "params": [
            ["name", "string", "Saved draft name", True, None]
        ],
        "returns": "string",
        "watch_out": ["Raises an error if no draft has that name.", "Matches both classic saved drafts and Unlayer saved drafts with that name."],
        "example": "for d in q.QuerySql(\"SELECT TOP 5 Name FROM Content WHERE TypeID IN (3, 6) ORDER BY DateCreated DESC\"):\n    print d.Name\n# title = model.DraftTitle(\"Welcome Email\")",
        "see_also": ["model.Draft", "model.TitleContent"]
    },
    "model.HtmlContent": {
        "category": "Content & Files",
        "kind": "method",
        "arity": [1, 1],
        "tags": ["read_only"],
        "summary": "Gets the body of an HTML Special Content item by name.",
        "details": "If no HTML item has that name it falls back to a Text item with the same name.",
        "params": [
            ["name", "string", "Name of the Special Content item", True, None]
        ],
        "returns": "string",
        "watch_out": ["Raises an error when neither an HTML nor a Text item exists, so wrap it in try/except if the name may be missing.", "For an item created by model.WriteContentText, the fallback returns that Text body as is."],
        "example": "try:\n    html = model.HtmlContent(\"CustomReports\")\n    print len(html), \"characters\"\nexcept Exception as e:\n    print \"Not found:\", e",
        "see_also": ["model.TextContent", "model.TitleContent", "model.WriteContentHtml", "model.Content"]
    },
    "model.LogToContent": {
        "category": "Content & Files",
        "kind": "method",
        "arity": [2, 2],
        "tags": ["writes_data"],
        "summary": "Appends a timestamped line to a text Special Content item, creating it if it does not exist.",
        "details": "Each call adds a line like \"11/2/26 14:05:09 PM text\" and a newline. Useful for a run log that persists between script runs.",
        "flow": ["Looks for a Text Special Content item with that name. Items of other types with the same name are ignored.", "If none exists, creates an empty Text item and saves it.", "Appends a line with the current date and time, your text and a newline, then saves again."],
        "params": [
            ["file", "string", "Name of the text Special Content item", True, None],
            ["text", "string", "Text to append", True, None]
        ],
        "returns": "None",
        "watch_out": ["Writes data.", "The item grows forever unless you clear it.", "The timestamp format is \"M/d/yy HH:mm:ss tt\", so it shows 24 hour time and also AM or PM, such as \"11/2/26 14:05:09 PM\".", "Each call rewrites the whole body, so a very large log gets slower to append to."],
        "example": "model.LogToContent(\"MyScriptLog\", \"Run by PeopleId \" + str(model.UserPeopleId))\nprint model.TextContent(\"MyScriptLog\")[-200:]",
        "see_also": ["model.TextContent", "model.WriteContentText"]
    },
    "model.PythonContent": {
        "category": "Content & Files",
        "kind": "method",
        "arity": [1, 1],
        "tags": ["read_only"],
        "summary": "Returns the text of a Python script in Special Content.",
        "details": "Falls back to a text content of the same name.",
        "params": [
            ["name", "string", "Python script name", True, None]
        ],
        "returns": "string",
        "watch_out": ["Returns an empty string if not found.", "A Python item wins over a Text item when both have the name."],
        "example": "code = model.PythonContent(\"MyOtherScript\")\nprint len(code), \"characters\"",
        "see_also": ["model.Content", "model.WriteContentPython", "model.CallScript"]
    },
    "model.ReadFile": {
        "category": "Content & Files",
        "kind": "method",
        "arity": [1, 1],
        "tags": ["read_only"],
        "summary": "Reads a file from the server.",
        "params": [
            ["name", "string", "File path", True, None]
        ],
        "returns": "string",
        "watch_out": ["Works only on a developer's local debug build.", "On a live TouchPoint site it always raises an error.", "The error raised has no message, so it is hard to tell what went wrong."],
        "example": "try:\n    print model.ReadFile(\"debug.txt\")\nexcept Exception as e:\n    print \"Not available on this server\"",
        "see_also": ["model.WriteFile", "model.DeleteFile", "model.IsDebug"]
    },
    "model.RenderTemplate": {
        "category": "Content & Files",
        "kind": "method",
        "arity": [1, 2],
        "tags": ["read_only"],
        "summary": "Renders a Handlebars template and returns the text.",
        "details": "With one argument the template reads from model.Data; pass a second argument to supply your own data object.",
        "params": [
            ["source", "string", "Handlebars template text", True, None],
            ["data", "any", "Data for the template", False, "model.Data"]
        ],
        "returns": "String",
        "watch_out": ["{{value}} is HTML encoded, so quotes and < > come out as entities. Use {{{value}}} to insert HTML as is."],
        "example": "model.Data.user = model.UserName\nprint model.RenderTemplate(\"Hi {{user}}\")",
        "signatures": ["model.RenderTemplate(source)", "model.RenderTemplate(source, data)"],
        "see_also": ["model.Data", "model.ControlTemplate", "model.TextControlTemplate"]
    },
    "model.SetSetting": {
        "category": "Content & Files",
        "kind": "method",
        "arity": [2, 2],
        "tags": ["writes_data"],
        "summary": "Saves an admin setting to the Setting table, creating it if needed.",
        "details": "The value is converted to text, so True is saved as \"True\".",
        "flow": ["Trims spaces from the setting name and converts the value to text.", "If the setting is marked as a hashed setting in the setting metadata, the value is encrypted before it is stored.", "Updates the Setting row with that name, or inserts a new one, and saves it.", "Updates the in memory settings cache on this server, so model.Setting sees the new value right away."],
        "params": [
            ["name", "string", "Setting name", True, None],
            ["value", "any", "Value to store. Converted to text, so True is saved as \"True\"", True, None]
        ],
        "returns": "None",
        "watch_out": ["Writes data.", "Changes the setting for the whole database.", "Passing None raises an error. Use \"\" to blank a setting.", "Saving an existing setting clears its System flag.", "Settings that are hashed are stored encrypted, so reading them back with model.Setting returns the encrypted text."],
        "example": "print \"Current:\", model.Setting(\"MyCustomSetting\", \"not set\")\n# Writes a database setting. Uncomment to run:\n# model.SetSetting(\"MyCustomSetting\", \"Value\")",
        "see_also": ["model.Setting"]
    },
    "model.Setting": {
        "category": "Content & Files",
        "kind": "method",
        "arity": [1, 2],
        "tags": ["read_only"],
        "summary": "Reads an admin setting from the Setting table.",
        "params": [
            ["name", "string", "Setting name", True, None],
            ["default", "string", "Value returned when the setting is missing", False, "\"\""]
        ],
        "returns": "string",
        "watch_out": ["Returns the default if the setting is missing or blank.", "Also returns the default when the setting exists but is blank.", "Never returns None. With no default it returns \"\".", "The name match ignores case.", "Settings are cached for about 15 seconds, so a change made on the admin screen may take a moment to show up.", "A hashed setting comes back as its encrypted text."],
        "example": "print model.Setting(\"NameOfChurch\")\nprint model.Setting(\"MyCustomSetting\", \"not set\")",
        "see_also": ["model.SetSetting"]
    },
    "model.SqlContent": {
        "category": "Content & Files",
        "kind": "method",
        "arity": [1, 1],
        "tags": ["read_only"],
        "summary": "Gets the body of a SQL Script Special Content item by name.",
        "details": "Handy for running a saved script with q.QuerySql.",
        "params": [
            ["name", "string", "Name of the SQL script", True, None]
        ],
        "returns": "string (\"\" when missing)",
        "watch_out": ["Returns an empty string when the item does not exist.", "Only SQL Script items are found. There is no fallback to a Text item."],
        "example": "sql = model.SqlContent(\"MySavedQuery\")\nif sql:\n    for row in q.QuerySql(sql):\n        print row\nelse:\n    print \"SQL script not found\"",
        "see_also": ["q.QuerySql", "model.WriteContentSql", "model.TextContent", "model.PythonContent"]
    },
    "model.TextContent": {
        "category": "Content & Files",
        "kind": "method",
        "arity": [1, 1],
        "tags": ["read_only"],
        "summary": "Gets the body of a Text Special Content item by name.",
        "details": "Often used to store JSON settings between runs.",
        "params": [
            ["name", "string", "Name of the Text content item", True, None]
        ],
        "returns": "string (\"\" when missing)",
        "watch_out": ["Returns an empty string when the item does not exist.", "Only Text items are found. An HTML or Python item with the same name is ignored and you get \"\"."],
        "example": "text = model.TextContent(\"CustomReports\")\nif text:\n    print text[:500]\nelse:\n    print \"No text content with that name\"",
        "see_also": ["model.WriteContentText", "model.HtmlContent", "model.SqlContent", "model.PythonContent"]
    },
    "model.TitleContent": {
        "category": "Content & Files",
        "kind": "method",
        "arity": [1, 1],
        "tags": ["read_only"],
        "summary": "Gets the Title of an HTML Special Content item by name.",
        "details": "Falls back to a Text item with that name.",
        "params": [
            ["name", "string", "Name of the content item", True, None]
        ],
        "returns": "string",
        "watch_out": ["Raises an error when neither an HTML nor a Text item exists.", "Items created by the model.WriteContent... calls have no title, so this returns None for them."],
        "example": "try:\n    print model.TitleContent(\"CustomReports\")\nexcept Exception as e:\n    print \"Not found:\", e",
        "see_also": ["model.HtmlContent", "model.TextContent", "model.DraftTitle"]
    },
    "model.WriteContent": {
        "category": "Content & Files",
        "kind": "method",
        "arity": [2, 3],
        "tags": ["writes_data"],
        "summary": "Saves text into a Special Content item, replacing the body or creating a new Text item.",
        "details": "If an item with that name already exists, of any type, its body is replaced. If none exists a new Text content item is created. When keyword is given it is set as the item's keyword. model.WriteContentText and model.WriteContentHtml are the clearer choices when you know the content type.",
        "flow": ["Looks for any Special Content item with that name, whatever its type.", "If none exists, creates a new Text item with that name.", "Replaces the body with your text.", "If a keyword is given, sets it as the item's only keyword, removing any others.", "Saves the change."],
        "params": [
            ["name", "string", "Special Content name", True, None],
            ["text", "string", "The new body", True, None],
            ["keyword", "string", "Optional keyword to tag the item with", False, "None"]
        ],
        "returns": "None",
        "watch_out": ["Writes data: an existing item's body is overwritten.", "If several items of different types share the name, whichever the database returns first is overwritten.", "A keyword replaces all other keywords on the item instead of adding to them.", "New items are created without a title.", "There is no role check, so any script that runs can overwrite any Special Content item."],
        "example": "import datetime\nmodel.WriteContent('TEST_LastRun', 'Last run ' + str(datetime.datetime.now()), 'TEST')\nprint model.TextContent('TEST_LastRun')",
        "signatures": ["model.WriteContent(name, text, keyword=None)"],
        "see_also": ["model.WriteContentText", "model.WriteContentHtml", "model.WriteContentSql", "model.WriteContentPython"]
    },
    "model.WriteContentHtml": {
        "category": "Content & Files",
        "kind": "method",
        "arity": [2, 3],
        "tags": ["writes_data"],
        "summary": "Creates or overwrites an HTML Special Content item with the given name and saves it immediately.",
        "details": "If a keyword is given it is attached to the item.",
        "flow": ["Looks for a Special Content item with that name and the HTML type.", "If none exists, creates a new HTML item with that name.", "Replaces the body, and if a keyword is given sets it as the item's only keyword.", "Saves the change right away."],
        "params": [
            ["name", "string", "Name of the content item", True, None],
            ["text", "string", "The full body to store (replaces any existing body)", True, None],
            ["keyword", "string", "Optional keyword to tag the item.", False, "None"]
        ],
        "returns": "None",
        "watch_out": ["Writes data: any existing body is replaced.", "HTML posted from a browser form is blocked by ASP.NET request validation, so encode it in JavaScript (&lt; and &gt;) and decode it in Python before saving.", "Only an item of the HTML type is matched. If an item of another type has the same name, it is left alone and a second item with the same name is created.", "A keyword replaces all other keywords on the item instead of adding to them.", "New items are created without a title.", "There is no role check, so any script that runs can create or overwrite these items."],
        "example": "html = \"<h1>Hello</h1><p>Saved from a script.</p>\"\n# Creates or replaces the HTML item named ApiExplorerTest:\n# model.WriteContentHtml(\"ApiExplorerTest\", html, \"\")\nprint \"Would save\", len(html), \"characters\"",
        "signatures": ["model.WriteContentHtml(name, text)", "model.WriteContentHtml(name, text, keyword)"],
        "see_also": ["model.HtmlContent", "model.WriteContentText", "model.WriteContent", "model.WriteContentPython"]
    },
    "model.WriteContentPython": {
        "category": "Content & Files",
        "kind": "method",
        "arity": [2, 3],
        "tags": ["writes_data"],
        "summary": "Creates or overwrites a Python Special Content item with the given name and saves it immediately.",
        "details": "If a keyword is given it is attached to the item.",
        "flow": ["Looks for a Special Content item with that name and the Python type.", "If none exists, creates a new Python item with that name.", "Replaces the body, and if a keyword is given sets it as the item's only keyword.", "Saves the change right away."],
        "params": [
            ["name", "string", "Name of the content item", True, None],
            ["script", "string", "The full body to store (replaces any existing body)", True, None],
            ["keyword", "string", "Optional keyword to tag the item.", False, "None"]
        ],
        "returns": "None",
        "watch_out": ["Writes data: any existing body is replaced.", "Anyone who can run scripts by name could then run the saved code, so be careful what you store.", "Only an item of the Python type is matched. If an item of another type has the same name, it is left alone and a second item with the same name is created.", "A keyword replaces all other keywords on the item instead of adding to them.", "New items are created without a title.", "There is no role check, so any script that runs can create or overwrite these items."],
        "example": "code = \"print 'Hello from a saved script'\"\n# Creates or replaces the Python script named ApiExplorerTest:\n# model.WriteContentPython(\"ApiExplorerTest\", code, \"\")\nprint code",
        "signatures": ["model.WriteContentPython(name, script)", "model.WriteContentPython(name, script, keyword)"],
        "see_also": ["model.PythonContent", "model.WriteContentSql", "model.WriteContentText", "model.CallScript"]
    },
    "model.WriteContentSql": {
        "category": "Content & Files",
        "kind": "method",
        "arity": [2, 3],
        "tags": ["writes_data"],
        "summary": "Creates or overwrites a SQL Script Special Content item with the given name and saves it immediately.",
        "details": "If a keyword is given it is attached to the item.",
        "flow": ["Looks for a Special Content item with that name and the SQL Script type.", "If none exists, creates a new SQL Script item with that name.", "Replaces the body, and if a keyword is given sets it as the item's only keyword.", "Saves the change right away."],
        "params": [
            ["name", "string", "Name of the content item", True, None],
            ["sql", "string", "The full body to store (replaces any existing body)", True, None],
            ["keyword", "string", "Optional keyword to tag the item.", False, "None"]
        ],
        "returns": "None",
        "watch_out": ["Writes data: any existing body is replaced.", "Only an item of the SQL Script type is matched. If an item of another type has the same name, it is left alone and a second item with the same name is created.", "A keyword replaces all other keywords on the item instead of adding to them.", "New items are created without a title.", "There is no role check, so any script that runs can create or overwrite these items."],
        "example": "sql = \"SELECT TOP 10 PeopleId, Name FROM People ORDER BY PeopleId\"\n# Creates or replaces the SQL script named ApiExplorerTest:\n# model.WriteContentSql(\"ApiExplorerTest\", sql, \"\")\nprint sql",
        "signatures": ["model.WriteContentSql(name, sql)", "model.WriteContentSql(name, sql, keyword)"],
        "see_also": ["model.SqlContent", "model.WriteContentText", "model.WriteContentPython", "model.WriteContent"]
    },
    "model.WriteContentText": {
        "category": "Content & Files",
        "kind": "method",
        "arity": [2, 3],
        "tags": ["writes_data"],
        "summary": "Creates or overwrites a Text Special Content item with the given name and saves it immediately.",
        "details": "If a keyword is given it is attached to the item. Commonly used with TextContent to keep JSON settings between runs.",
        "flow": ["Looks for a Special Content item with that name and the Text type.", "If none exists, creates a new Text item with that name.", "Replaces the body, and if a keyword is given sets it as the item's only keyword.", "Saves the change right away."],
        "params": [
            ["name", "string", "Name of the content item", True, None],
            ["text", "string", "The full body to store (replaces any existing body)", True, None],
            ["keyword", "string", "Optional keyword to tag the item.", False, "None"]
        ],
        "returns": "None",
        "watch_out": ["Writes data: any existing body is replaced.", "Build the JSON with a non-ASCII safe encoder, since json.dumps can fail on names with accents in IronPython.", "Only an item of the Text type is matched. If an item of another type has the same name, it is left alone and a second item with the same name is created.", "A keyword replaces all other keywords on the item instead of adding to them.", "New items are created without a title.", "There is no role check, so any script that runs can create or overwrite these items."],
        "example": "import json\nsettings = {\"lastRun\": model.DateTime.ToString(\"yyyy-MM-dd\")}\n# Creates or replaces the Text item named ApiExplorerTest:\n# model.WriteContentText(\"ApiExplorerTest\", json.dumps(settings), \"\")\nprint json.dumps(settings)",
        "signatures": ["model.WriteContentText(name, text)", "model.WriteContentText(name, text, keyword)"],
        "see_also": ["model.TextContent", "model.WriteContentHtml", "model.WriteContent", "model.JsonSerialize"]
    },
    "model.WriteFile": {
        "category": "Content & Files",
        "kind": "method",
        "arity": [2, 2],
        "tags": ["writes_data"],
        "summary": "Appends text to a file on the server.",
        "params": [
            ["path", "string", "File path", True, None],
            ["text", "string", "Text to append", True, None]
        ],
        "returns": "None",
        "watch_out": ["Works only on a developer's local debug build.", "On a live TouchPoint site it always raises an error.", "The error raised has no message, so it is hard to tell what went wrong."],
        "example": "try:\n    model.WriteFile(\"debug.txt\", \"hello\\n\")\n    print \"written\"\nexcept Exception as e:\n    print \"Not available on this server\"",
        "see_also": ["model.ReadFile", "model.DeleteFile", "model.IsDebug"]
    },
    "model.AddContribution": {
        "category": "Contributions",
        "kind": "method",
        "arity": [6, 7],
        "tags": ["writes_data"],
        "summary": "Builds a new contribution and its bundle detail in memory and returns the BundleDetail.",
        "details": "Add the result to a bundle with bh.BundleDetails.Add(bd), then call model.FinishBundle(bh) to save. The type defaults to tax deductible, or non deductible if the fund is marked that way.",
        "flow": ["Looks up the fund by id. If no fund has that id, it creates one right away with that id, no name, and active status, and saves it.", "Reads the amount text, keeping only digits, the decimal point and the minus sign, so \"$1,250.00\" becomes 1250.00. Text it cannot read leaves the amount empty.", "Picks the type: the contributionTypeId you pass, otherwise tax deductible (1), or non tax deductible (9) when the fund is marked that way. A zero, negative or empty amount defaults to Reversed (7), or to Returned check (6) when you gave a check number.", "Builds a new Contribution with status 0 (posted), the date, fund, check number and description, and the giver when peopleId is above 0. Wraps it in a new BundleDetail.", "Returns the BundleDetail without saving. It is only saved once you add it to a header with bh.BundleDetails.Add(bd) and something saves, normally model.FinishBundle(bh)."],
        "params": [
            ["date", "date", "Contribution date", True, None],
            ["fundId", "integer", "Fund id", True, None],
            ["amount", "string", "Amount as text, such as \"100.00\" or \"$1,250.00\"", True, None],
            ["checkNo", "string", "Check number, or an empty string", True, None],
            ["description", "string", "Contribution description", True, None],
            ["peopleId", "integer", "Giver PeopleId. 0 leaves the gift unassigned", True, None],
            ["contributionTypeId", "integer", "Contribution type id (1 = tax deductible). Omit to use the default", False, None]
        ],
        "returns": "BundleDetail object (unsaved). Its Contribution property holds the new Contribution.",
        "watch_out": ["Nothing is saved by this call until model.FinishBundle runs.", "When you do not pass a type, a zero or negative amount defaults the type to Reversed, or to Returned check when a check number is given.", "If fundId does not exist, a blank fund with that id is created and saved immediately.", "A detail you never add to a header is never saved.", "Passing 0 or a missing fund id creates a nameless fund. Check the id with model.ResolveFundId or model.FirstFundId first.", "A gift with no giver (peopleId 0) keeps model.CloseBundle from closing the bundle."],
        "example": "import datetime\nnow = datetime.datetime.now()\nfund = model.FirstFundId()\nprint \"Default fund:\", fund\n# Writes data. Uncomment to post a $1 test gift for yourself:\n# bh = model.GetBundleHeader(now, now)\n# bd = model.AddContribution(now, fund, \"1.00\", \"\", \"API test\", model.UserPeopleId)\n# bh.BundleDetails.Add(bd)\n# model.FinishBundle(bh)\n# print \"Bundle\", bh.BundleHeaderId",
        "signatures": ["model.AddContribution(date, fundId, amount, checkNo, description, peopleId)", "model.AddContribution(date, fundId, amount, checkNo, description, peopleId, contributionTypeId)"],
        "see_also": ["model.FinishBundle", "model.GetBundleHeader", "model.AddContributionDetail", "model.FindOrCreateBundleHeader"]
    },
    "model.AddContributionDetail": {
        "category": "Contributions",
        "kind": "method",
        "arity": [6, 9],
        "tags": ["writes_data"],
        "summary": "Builds a new contribution and bundle detail in memory for a bank import, storing routing and account encrypted.",
        "details": "When no peopleId is given, it looks up the giver in TouchPoint's saved bank account list by the encrypted routing and account. Add the result to a bundle and call model.FinishBundle to save.",
        "flow": ["Looks up the fund by id and creates a blank active fund with that id if it does not exist, saving it immediately.", "Reads the amount and picks the type the same way as model.AddContribution: the type you pass, otherwise tax deductible or non tax deductible by fund, and Reversed for a zero or negative amount.", "Stores routing and account together, encrypted, on the contribution. An account number that looks like a spreadsheet number (contains \"E+\") is skipped.", "When no peopleId is passed, matches the encrypted account against the saved bank account list (CardIdentifiers) and sets the giver if exactly one person has it. Otherwise the gift is left without a giver.", "Returns the unsaved BundleDetail. Add it to a header and call model.FinishBundle(bh) to save."],
        "params": [
            ["date", "date", "Contribution date", True, None],
            ["fundId", "integer", "Fund id", True, None],
            ["amount", "string", "Amount as text, such as \"100.00\"", True, None],
            ["checkNo", "string", "Check number", True, None],
            ["routing", "string", "Bank routing number", True, None],
            ["account", "string", "Bank account number", True, None],
            ["contributionTypeId", "integer", "Contribution type id. Required if you pass peopleId or description", False, None],
            ["peopleId", "integer", "Giver PeopleId (matched by bank account when omitted)", False, "None"],
            ["description", "string", "Contribution description", False, "None"]
        ],
        "returns": "BundleDetail object (unsaved)",
        "watch_out": ["Nothing is saved until you add the result to a bundle and call model.FinishBundle.", "An unknown fundId creates a blank fund immediately.", "It does not add the account to the saved bank account list, so importing a check this way does not teach TouchPoint who owns the account. Unmatched gifts must be assigned by hand, and the bundle cannot close until they are.", "Account numbers exported from Excel in scientific notation (such as 9.87654E+08) are not stored and cannot be matched.", "Unlike model.AddContribution, a negative amount with a check number still defaults to Reversed, not Returned check."],
        "example": "import datetime\nnow = datetime.datetime.now()\n# Writes data. Uncomment to import one check:\n# bh = model.GetBundleHeader(now, now)\n# bd = model.AddContributionDetail(now, model.FirstFundId(), \"25.00\", \"1001\", \"123456789\", \"987654321\")\n# bh.BundleDetails.Add(bd)\n# model.FinishBundle(bh)\n# print \"Giver matched:\", bd.Contribution.PeopleId\nprint \"Uncomment the lines above to import a test check\"",
        "signatures": ["model.AddContributionDetail(date, fundId, amount, checkNo, routing, account)", "model.AddContributionDetail(date, fundId, amount, checkNo, routing, account, contributionTypeId, peopleId=None, description=None)"],
        "see_also": ["model.AddContribution", "model.FinishBundle", "model.FindOrCreateBundleHeader"]
    },
    "model.CloseBundle": {
        "category": "Contributions",
        "kind": "method",
        "arity": [1, 1],
        "tags": ["writes_data"],
        "summary": "Closes a bundle so it can no longer be edited, and saves.",
        "details": "It only closes when the cash, check and envelope totals are all set, they add up to the sum of the bundle's contributions, and every contribution has a giver. Call model.FinishBundle first to fill in the totals.",
        "flow": ["Checks that TotalCash, TotalChecks and TotalEnvelopes all have values. A new bundle has none until model.FinishBundle sets them.", "Checks that cash plus checks plus envelopes equals the sum of the bundle's gifts.", "Checks that every gift in the bundle has a giver.", "If all checks pass, sets BundleStatusId to Closed (0), saves every pending change, and returns True. If any check fails, it changes nothing and returns False."],
        "params": [
            ["bundleHeader", "any", "BundleHeader object, from model.GetBundleHeader or model.FindBundleHeader", True, None]
        ],
        "returns": "Boolean. True if the bundle was closed.",
        "watch_out": ["When the totals or givers do not check out it changes nothing and returns False.", "A True result also saves any other unsaved changes in the script."],
        "example": "import datetime\nbh = model.FindBundleHeader(datetime.datetime.now())\nif bh:\n    print \"Bundle\", bh.BundleHeaderId, \"status\", bh.BundleStatusId\n    # Writes data. Uncomment to close it:\n    # model.FinishBundle(bh)\n    # print model.CloseBundle(bh)\nelse:\n    print \"No bundle dated today\"",
        "see_also": ["model.FinishBundle", "model.GetBundleHeader", "model.FindBundleHeader"]
    },
    "model.CreateContributionTag": {
        "category": "Contributions",
        "kind": "method",
        "arity": [2, 2],
        "tags": ["writes_data"],
        "summary": "Saves a named set of contribution ids (a contribution tag) so dashboard reports can reuse it quickly.",
        "details": "Criteria are fields on a DynamicData object, such as MinDate, MaxDate, FundIds, CampusId, FromAmount, ToAmount, BundleTypes, Source, NonTaxDed.",
        "flow": ["Turns your DynamicData criteria into JSON and passes it with the tag name to the TagContributions database procedure. This runs and saves right away.", "Unless the criteria include AddToTag = 1, every ContributionTag row with that tag name is deleted first.", "Finds the matching gift ids and inserts one ContributionTag row per gift with that tag name. With AddToTag = 1 it only adds gifts not already in the tag. A Priority value in the criteria is saved on each row.", "Returns the criteria formatted as JSON."],
        "params": [
            ["name", "string", "Tag name. Use a common prefix so you can clear them later", True, None],
            ["criteria", "dict", "DynamicData object from model.DynamicData() holding the search criteria", True, None]
        ],
        "returns": "String. The criteria formatted as JSON.",
        "watch_out": ["Writes to the ContributionTag table.", "Any existing tag with the same name is replaced unless the criteria include AddToTag = 1.", "Tag names longer than 50 characters do not fit the procedure's tag name field."],
        "example": "dd = model.DynamicData()\ndd.MinDate = \"2025-01-01\"\ndd.MaxDate = \"2025-12-31\"\n# Writes data. Uncomment to build the tag:\n# print model.CreateContributionTag(\"MyReport_2025\", dd)\nprint \"Uncomment to create tag MyReport_2025\"",
        "see_also": ["model.CreateContributionTagFromSql", "model.DeleteContributionTags", "model.DynamicData"]
    },
    "model.CreateContributionTagFromSql": {
        "category": "Contributions",
        "kind": "method",
        "arity": [3, 3],
        "tags": ["writes_data"],
        "summary": "Saves a contribution tag using SQL you supply to pick the contributions.",
        "details": "Like CreateContributionTag. The SQL must return a single column of ContributionIds. The criteria object is passed to it as JSON in the @json parameter.",
        "flow": ["Turns the criteria object into JSON and passes it to your SQL as @json.", "Runs your SQL on the read only connection with a 10 minute timeout and collects the ContributionIds it returns.", "Unless the criteria has AddToTag set to 1, deletes every ContributionTag row with this tag name.", "Inserts one ContributionTag row per id, with the Priority value from the criteria if one is set.", "Returns the tag name and the criteria as indented JSON."],
        "params": [
            ["name", "string", "Tag name", True, None],
            ["criteria", "dict", "DynamicData object from model.DynamicData(), available to the SQL as @json", True, None],
            ["sql", "string", "SQL returning one column of ContributionId values", True, None]
        ],
        "returns": "String in the form \"name = {criteria as indented JSON}\".",
        "watch_out": ["Writes to the ContributionTag table.", "Replaces any tag with the same name, unless the criteria includes AddToTag: 1.", "Add AddToTag: 1 to the criteria to keep existing rows and only add new ids.", "Tag names are limited to 50 characters; longer names are cut off.", "The SQL runs read only, under the Finance login only when the current user has the Finance role."],
        "example": "dd = model.DynamicData()\ndd.MinAmount = 1000\nsql = \"SELECT ContributionId FROM Contribution WHERE ContributionAmount >= JSON_VALUE(@json, '$.MinAmount') AND ContributionDate >= '2025-01-01'\"\n# Writes data. Uncomment to build the tag:\n# print model.CreateContributionTagFromSql(\"MyReport_Large\", dd, sql)\nprint \"Uncomment to create tag MyReport_Large\"",
        "see_also": ["model.CreateContributionTag", "model.DeleteContributionTags", "model.DynamicData"]
    },
    "model.CustomStatementsFundIdList": {
        "category": "Contributions",
        "kind": "method",
        "arity": [1, 1],
        "tags": ["read_only"],
        "summary": "Returns the fund ids in a Custom Statements fund set as a comma separated string.",
        "details": "Use it in SQL with dbo.SplitInts to filter contributions to that fund set.",
        "flow": ["Reads the CustomStatements Text Special Content, which is XML.", "Finds the Statement whose description matches the name and reads its Funds list.", "Expands ranges such as 100-110 into the fund ids that exist, and returns them joined with commas."],
        "params": [
            ["name", "string", "Name of the custom statement fund set", True, None]
        ],
        "returns": "string (comma separated fund ids)",
        "watch_out": ["Returns \"\" when no statement has that name.", "The name of the standard set (the StandardFundSetName setting, \"Standard Statements\" by default) returns every fund not used by a custom statement.", "Passing \"all\" raises an error."],
        "example": "funds = model.CustomStatementsFundIdList(\"General Fund Statement\")\nprint funds\nprint q.QuerySqlInt(\"SELECT COUNT(*) FROM dbo.SplitInts('\" + funds + \"')\")",
        "see_also": ["q.QuerySql", "q.ContributionTotals"]
    },
    "model.DeleteContribution": {
        "category": "Contributions",
        "kind": "method",
        "arity": [1, 1],
        "tags": ["writes_data"],
        "summary": "Permanently deletes a contribution, its bundle detail and its contribution tags.",
        "flow": ["Looks up the BundleDetails row for that contribution id. If there is none, it does nothing.", "Deletes that bundle detail, every ContributionTag row for the gift, and the Contribution row, then saves right away. No error is raised when nothing is found."],
        "params": [
            ["contributionId", "integer", "ContributionId to delete", True, None]
        ],
        "returns": "None",
        "watch_out": ["Cannot be undone.", "Does nothing if the contribution is not in a bundle.", "The bundle's totals and count are not updated. Call model.FinishBundle on the bundle afterward if they should match.", "It deletes from closed bundles too. Nothing checks the bundle status.", "Saving here also writes any other unsaved changes in the script, such as a queued bundle header."],
        "example": "c = model.FindContribution(model.UserPeopleId)\nif c:\n    print \"Most recent match:\", c.ContributionId, c.ContributionAmount\n    # Deletes data permanently. Uncomment only if you are sure:\n    # model.DeleteContribution(c.ContributionId)",
        "see_also": ["model.FindContribution", "model.AddContribution"]
    },
    "model.DeleteContributionTags": {
        "category": "Contributions",
        "kind": "method",
        "arity": [1, 1],
        "tags": ["writes_data"],
        "summary": "Deletes every contribution tag whose name matches a SQL LIKE pattern.",
        "details": "Use % as the wildcard. This is how a report clears its old tags before rebuilding them.",
        "params": [
            ["namePattern", "string", "SQL LIKE pattern, such as \"MyReport_%\"", True, None]
        ],
        "returns": "None",
        "watch_out": ["Deletes data for every matching tag name.", "An underscore is also a LIKE wildcard, so \"MyReport_%\" matches any character in that spot. Escape it as [_] for an exact match.", "Returns nothing and raises no error when no tag matches."],
        "example": "# Writes data. Uncomment to clear all tags starting with MyReport_:\n# model.DeleteContributionTags(\"MyReport_%\")\nprint q.QuerySqlInt(\"SELECT COUNT(*) FROM ContributionTag WHERE TagName LIKE 'MyReport_%'\")",
        "see_also": ["model.CreateContributionTag", "model.CreateContributionTagFromSql"]
    },
    "model.DisableRecurringGiving": {
        "category": "Contributions",
        "kind": "method",
        "arity": [2, 2],
        "tags": ["deprecated", "known_issue"],
        "summary": "Was meant to disable recurring giving, but does nothing.",
        "details": "Manage recurring giving in the TouchPoint giving screens.",
        "params": [
            ["peopleId", "integer", "PeopleId", True, None],
            ["fundId", "integer", "Fund id", True, None]
        ],
        "returns": "None",
        "watch_out": ["Deprecated and does nothing. The method body is empty."],
        "example": "# Does nothing; kept only so old scripts do not break.\nprint \"DisableRecurringGiving has no effect\"",
        "see_also": ["model.UpdateContributionOption"]
    },
    "model.FetchOrCreateFund": {
        "category": "Contributions",
        "kind": "method",
        "arity": [1, 1],
        "tags": ["writes_data"],
        "summary": "Finds an active fund by exact name, or creates one if none exists.",
        "details": "A new fund is active, uses that name and gets the next fund id.",
        "flow": ["Looks for an active fund (FundStatusId 1) whose FundName exactly matches the name and returns the first one found.", "If none matches, creates a ContributionFund with the next fund id after the highest existing one, the name as both fund name and description, active status and fund type 1.", "Saves the new fund right away and returns it."],
        "params": [
            ["name", "string", "Fund name", True, None]
        ],
        "returns": "ContributionFund object",
        "watch_out": ["Creates and saves a new fund when no exact name match is found.", "A closed or inactive fund with the same name is not reused. A new active duplicate is created instead.", "Saving the new fund also writes any other unsaved changes in the script, such as a queued bundle header."],
        "example": "fund = model.ResolveFund(\"General Fund\")\nif fund:\n    print fund.FundId, fund.FundName\n# Creates a fund if missing. Uncomment to use:\n# fund = model.FetchOrCreateFund(\"General Fund\")",
        "see_also": ["model.ResolveFund", "model.ResolveFundId", "model.FirstFundId"]
    },
    "model.FindBundleHeader": {
        "category": "Contributions",
        "kind": "method",
        "arity": [1, 4],
        "tags": ["read_only"],
        "summary": "Finds an existing bundle by date and optional bundle type and reference id.",
        "details": "Pass a date (a datetime object, or a date string when it is the only argument) and optionally part of the bundle type description (such as \"Online\"), a reference id and a reference id type. A second form takes only bundleType and referenceId and returns the oldest match. Matches bundles of any status.",
        "flow": ["Date form: reads the date you pass (a date or a date string) and raises an error if it cannot be read as a date. Only the day part is used.", "Looks in BundleHeaders for bundles with that contribution date, narrowed by bundle type (the type description contains your text), reference id and reference id type when you give them.", "Returns the first match, or None. Open and closed bundles both match.", "Two text form: needs both a bundle type and a reference id, or it raises an error. It ignores dates and returns the oldest matching bundle by created date."],
        "params": [
            ["date", "date", "Bundle contribution date", True, None],
            ["bundleType", "string", "Part of the bundle type description, such as \"Online\"", False, "None"],
            ["referenceId", "string", "Bundle reference id", False, "None"],
            ["referenceIdType", "integer", "Reference id type", False, "None"]
        ],
        "returns": "BundleHeader object or None",
        "watch_out": ["Returns None when no bundle matches.", "Calling it with a date string and a bundle type, such as model.FindBundleHeader(\"2026-11-01\", \"Online\"), uses the two text form, so the date is taken as the bundle type and \"Online\" as the reference id. Pass a real date object (datetime) when you also pass a bundle type.", "When several bundles match a date, which one comes back first is not defined. Add a reference id or bundle type to narrow it.", "Closed bundles are returned too. Check bh.BundleStatusId (1 open, 0 closed) before adding gifts."],
        "example": "import datetime\nbh = model.FindBundleHeader(datetime.datetime.now(), \"Online\")\nif bh:\n    print bh.BundleHeaderId, bh.BundleTotal\nelse:\n    print \"No online bundle today\"",
        "signatures": ["model.FindBundleHeader(date, bundleType=None, referenceId=None, referenceIdType=None)", "model.FindBundleHeader(bundleType, referenceId)"],
        "see_also": ["model.FindOrCreateBundleHeader", "model.GetBundleHeader", "model.CloseBundle"]
    },
    "model.FindContribution": {
        "category": "Contributions",
        "kind": "method",
        "arity": [0, 6],
        "tags": ["read_only"],
        "summary": "Finds a contribution matching every criterion you pass.",
        "details": "All arguments are optional, so pass at least one, by keyword if you skip any. Date matches on the day only. Returns the first match.",
        "params": [
            ["peopleId", "integer", "Giver PeopleId", False, "None"],
            ["checkNo", "string", "Check number", False, "None"],
            ["metaInfo", "string", "MetaInfo text", False, "None"],
            ["fundId", "integer", "Fund id", False, "None"],
            ["amount", "decimal", "Exact amount", False, "None"],
            ["date", "date", "Contribution date (day only)", False, "None"]
        ],
        "returns": "Contribution object or None",
        "watch_out": ["Returns None when nothing matches.", "There is no sort order, so with several matches the one returned is not predictable.", "It matches contributions of every status and type, including pledges and reversed or returned items.", "Calling it with no arguments returns some contribution from the whole table."],
        "example": "c = model.FindContribution(peopleId=model.UserPeopleId)\nif c:\n    print c.ContributionId, c.ContributionDate, c.ContributionAmount\nelse:\n    print \"No contribution found\"",
        "see_also": ["model.DeleteContribution", "model.QueryContributionIds"]
    },
    "model.FindOrCreateBundleHeader": {
        "category": "Contributions",
        "kind": "method",
        "arity": [1, 4],
        "tags": ["writes_data"],
        "summary": "Returns the first bundle matching the date and optional type and reference id, or builds a new open one.",
        "details": "Add contributions to a new bundle and call model.FinishBundle to save it.",
        "flow": ["Looks for a bundle on that day, of a type whose description contains bundleType, with the same reference id and type when given.", "If one is found it is returned, open or closed.", "Otherwise builds a new open bundle using the default fund, source and contribution type from the church settings.", "Sets its type to the first bundle type whose description contains bundleType, or Checks and Cash when none matches or none is given.", "The new bundle is queued, not saved. It is written on the next save, usually model.FinishBundle."],
        "params": [
            ["date", "date", "Bundle date", True, None],
            ["bundleType", "string", "Part of the bundle type description, such as \"Online\"", False, "None"],
            ["referenceId", "string", "Reference id stored on the bundle", False, "None"],
            ["referenceIdType", "integer", "Reference id type", False, "None"]
        ],
        "returns": "BundleHeader object",
        "watch_out": ["A new bundle is not saved until you add contributions and call model.FinishBundle.", "Raises \"A valid bundle date is required\" if the date cannot be read.", "It can return a bundle that is already closed.", "Calling it twice before the new bundle is saved creates two bundles.", "bundleType is a partial match, so \"Online\" can match more than one bundle type."],
        "example": "import datetime\n# May create a bundle when you later call FinishBundle.\nbh = model.FindOrCreateBundleHeader(datetime.datetime.now(), \"Online\", \"IMPORT-TEST\")\nprint bh.BundleHeaderId  # 0 means it is new and not saved yet",
        "see_also": ["model.FindBundleHeader", "model.GetBundleHeader", "model.FinishBundle", "model.AddContribution"]
    },
    "model.FinishBundle": {
        "category": "Contributions",
        "kind": "method",
        "arity": [1, 1],
        "tags": ["writes_data"],
        "summary": "Totals a bundle and saves all pending changes. The bundle stays open.",
        "details": "Sets TotalChecks and BundleTotal to the sum of its contributions, BundleCount to the number of items, and cash and envelope totals to 0. This is the call that actually saves contributions added with AddContribution.",
        "flow": ["Adds up the amounts of every gift attached to the bundle and puts that sum in both TotalChecks and BundleTotal.", "Sets BundleCount to the number of gifts, and TotalCash and TotalEnvelopes to 0.", "Saves every pending change in one step: a new header, every gift added with bh.BundleDetails.Add, and any other unsaved edits. New ids such as bh.BundleHeaderId and bd.ContributionId are filled in after this.", "Leaves the bundle Open. Call model.CloseBundle(bh) afterward if you want it locked."],
        "params": [
            ["bundleHeader", "any", "BundleHeader object", True, None]
        ],
        "returns": "None",
        "watch_out": ["Saves every pending change, not only the bundle.", "Cash and envelope totals are set to 0.", "It does not close the bundle.", "Any expected total you passed to model.GetBundleHeader is overwritten with the actual sum."],
        "example": "import datetime\nnow = datetime.datetime.now()\n# Writes data. Uncomment to save a bundle with one $1 gift:\n# bh = model.GetBundleHeader(now, now)\n# bh.BundleDetails.Add(model.AddContribution(now, model.FirstFundId(), \"1.00\", \"\", \"test\", model.UserPeopleId))\n# model.FinishBundle(bh)\n# print bh.BundleHeaderId, bh.BundleTotal\nprint \"Uncomment to run\"",
        "see_also": ["model.AddContribution", "model.GetBundleHeader", "model.CloseBundle", "model.FindOrCreateBundleHeader"]
    },
    "model.FirstFundId": {
        "category": "Contributions",
        "kind": "method",
        "arity": [0, 0],
        "tags": ["read_only"],
        "summary": "Returns the lowest numbered active fund id.",
        "details": "Often used as a default fund.",
        "returns": "Integer fund id",
        "watch_out": ["Raises an error if there are no active funds."],
        "example": "fundId = model.FirstFundId()\nprint \"First active fund:\", fundId",
        "see_also": ["model.ResolveFundId", "model.ResolveFund", "model.FetchOrCreateFund"]
    },
    "model.GetBundleHeader": {
        "category": "Contributions",
        "kind": "method",
        "arity": [2, 5],
        "tags": ["writes_data"],
        "summary": "Creates a new open bundle header and queues it for insert.",
        "details": "Type defaults to Checks and Cash. Fund, source and contribution type come from the church defaults.",
        "flow": ["Builds a new BundleHeaders row with status Open (BundleStatusId 1) and the contribution date you pass. The created date is the now value you pass, and the creator is you.", "Copies the default fund, source and contribution type from the church defaults onto the header.", "Sets the bundle type to the bundleTypeId you pass, or Checks and Cash (2) when you leave it out. Deposit date and expected total are stored as given.", "Queues the header for insert and returns it. Nothing is saved yet, so BundleHeaderId is 0 until something saves.", "Typical script: create the header, add gifts with bh.BundleDetails.Add(model.AddContribution(...)), then call model.FinishBundle(bh) to total and save everything in one step. Call model.CloseBundle(bh) afterward if you want it locked."],
        "params": [
            ["date", "date", "Contribution date for the bundle", True, None],
            ["now", "date", "Created date, usually datetime.datetime.now()", True, None],
            ["bundleTypeId", "integer", "Bundle header type id (Checks and Cash when omitted)", False, "None"],
            ["depositDate", "date", "Deposit date", False, "None"],
            ["bundleTotal", "decimal", "Expected bundle total", False, "None"]
        ],
        "returns": "BundleHeader object (unsaved)",
        "watch_out": ["It is saved by the next call that saves changes, normally model.FinishBundle.", "Anything else that saves along the way also saves this header, even an empty one. For example, model.AddContribution with a fund id that does not exist creates that fund on the spot, and that save writes the queued header too.", "It never looks for an existing bundle. Running the script twice makes two bundles. Use model.FindOrCreateBundleHeader to reuse one."],
        "example": "import datetime\nnow = datetime.datetime.now()\n# Writes data once FinishBundle runs. Uncomment to use:\n# bh = model.GetBundleHeader(now, now)\n# bh.BundleDetails.Add(model.AddContribution(now, model.FirstFundId(), \"1.00\", \"\", \"test\", model.UserPeopleId))\n# model.FinishBundle(bh)\n# print bh.BundleHeaderId\nprint \"Uncomment to create a bundle\"",
        "see_also": ["model.FinishBundle", "model.FindOrCreateBundleHeader", "model.FindBundleHeader", "model.AddContribution"]
    },
    "model.MoveFundIdToExistingFundId": {
        "category": "Contributions",
        "kind": "method",
        "arity": [2, 3],
        "tags": ["writes_data", "known_issue"],
        "summary": "Moves every contribution from one fund to another existing fund.",
        "details": "Uses a direct database update.",
        "flow": ["Looks up both funds by id and raises an error if either one does not exist.", "Runs one direct database update that changes the fund on every Contribution row in the old fund to the new fund, including pledges, returned and reversed gifts, and gifts in closed bundles.", "Leaves the old fund in place, and does not change bundle headers or any other records that point at the old fund."],
        "params": [
            ["fromFundId", "integer", "Fund id to move contributions out of", True, None],
            ["toFundId", "integer", "Existing fund id to move them into", True, None],
            ["name", "string", "Ignored", False, "None"]
        ],
        "returns": "None",
        "watch_out": ["Both funds must exist or it raises an error.", "The name argument is accepted but ignored.", "Cannot be undone easily.", "Needs a role allowed to run scripts that write data.", "No change history is written for the moved gifts, so there is no record of which fund they came from."],
        "example": "print q.QuerySqlInt(\"SELECT COUNT(*) FROM Contribution WHERE FundId = 10\"), \"gifts in fund 10\"\n# Writes data for every gift in the fund. Uncomment only if you are sure:\n# model.MoveFundIdToExistingFundId(10, 20)",
        "see_also": ["model.MoveFundIdToNewFundId", "model.ResolveFundId"]
    },
    "model.MoveFundIdToNewFundId": {
        "category": "Contributions",
        "kind": "method",
        "arity": [2, 3],
        "tags": ["writes_data"],
        "summary": "Creates a new fund with id toFundId and moves every contribution from fromFundId into it.",
        "details": "The new fund is named name, or the old fund's description.",
        "flow": ["Checks that fromFundId exists and that toFundId does not.", "Creates the new fund with id toFundId, active, using name (or the old fund's description) for both its name and description.", "Moves every contribution from the old fund to the new one with one SQL update."],
        "params": [
            ["fromFundId", "integer", "Fund id to move contributions out of", True, None],
            ["toFundId", "integer", "New fund id. Must not exist yet", True, None],
            ["name", "string", "Name for the new fund (the old fund's description when omitted)", False, "None"]
        ],
        "returns": "None",
        "watch_out": ["Raises an error if toFundId already exists or fromFundId does not.", "Moves every contribution in the old fund.", "The old fund is left in place and stays active.", "Every contribution moves, including pledges and reversed or returned items.", "Nothing is logged, and there is no undo."],
        "example": "print q.QuerySqlInt(\"SELECT COUNT(*) FROM ContributionFund WHERE FundId = 999\"), \"funds with id 999\"\n# Writes data for every gift in the fund. Uncomment only if you are sure:\n# model.MoveFundIdToNewFundId(10, 999, \"New Building Fund\")",
        "see_also": ["model.MoveFundIdToExistingFundId", "model.FetchOrCreateFund"]
    },
    "model.QueryContributionIds": {
        "category": "Contributions",
        "kind": "method",
        "arity": [2, 2],
        "tags": ["read_only"],
        "summary": "Runs read only SQL that returns one integer column (usually ContributionId) and returns the values.",
        "details": "The second argument is required. Pass None when the SQL has no parameters.",
        "params": [
            ["sql", "string", "SQL returning a single integer column", True, None],
            ["declarations", "any", "Parameter object for the SQL, or None", True, None]
        ],
        "returns": "List of integers",
        "watch_out": ["Times out after 10 minutes.", "The SQL runs on a read only connection, so INSERT, UPDATE and DELETE fail.", "It uses the Finance read only login only when the current user has the Finance role.", "Every row must be an integer in the first column, or it raises an error."],
        "example": "ids = list(model.QueryContributionIds(\"SELECT TOP 10 ContributionId FROM Contribution WHERE ContributionAmount > 1000 ORDER BY ContributionDate DESC\", None))\nprint len(ids), \"ids:\", ids",
        "see_also": ["model.CreateContributionTagFromSql", "q.QuerySqlInts", "model.FindContribution"]
    },
    "model.ResolveFund": {
        "category": "Contributions",
        "kind": "method",
        "arity": [1, 1],
        "tags": ["read_only"],
        "summary": "Finds a fund whose FundName matches exactly, among funds with a status above 0.",
        "params": [
            ["name", "string", "Exact fund name", True, None]
        ],
        "returns": "ContributionFund object or None",
        "watch_out": ["Returns None if there is no match.", "Raises an error if more than one fund has that name.", "Funds of any status above 0 count, so an inactive fund with the exact name can be returned."],
        "example": "fund = model.ResolveFund(\"General Fund\")\nif fund:\n    print fund.FundId, fund.FundName\nelse:\n    print \"No fund with that name\"",
        "see_also": ["model.ResolveFundId", "model.FetchOrCreateFund", "model.FirstFundId"]
    },
    "model.ResolveFundId": {
        "category": "Contributions",
        "kind": "method",
        "arity": [1, 1],
        "tags": ["read_only"],
        "summary": "Returns a fund id for a fund name, or for a fund id given as text.",
        "params": [
            ["fundName", "string", "Fund name, or a fund id as text such as \"12\"", True, None]
        ],
        "returns": "Integer fund id",
        "watch_out": ["If neither matches, it returns the first active fund id instead, so check the result if a wrong fund matters.", "The name is checked first, so a fund named \"12\" wins over fund id 12.", "The name match includes inactive and closed funds, and the id match accepts any fund status.", "If two funds have exactly the same name, it raises an error.", "The name must match exactly."],
        "example": "print model.ResolveFundId(\"General Fund\")\nprint model.ResolveFundId(\"No Such Fund\")  # falls back to model.FirstFundId()",
        "see_also": ["model.ResolveFund", "model.FirstFundId", "model.FetchOrCreateFund"]
    },
    "model.UpdateContributionOption": {
        "category": "Contributions",
        "kind": "method",
        "arity": [2, 2],
        "tags": ["writes_data"],
        "summary": "Sets the giving statement option for everyone in a query.",
        "details": "Codes: 1 individual, 2 joint, 9 none, 0 clears it. Accepts a PeopleId, a saved search name, or Search Builder code.",
        "flow": ["Runs the query to get the list of people.", "For each person, sets the statement option, where 0 clears it, and records the change in the change log.", "If the code is 1, 2 or 9, the spouse gets the same value.", "If the code is 0, the spouse is cleared only when the spouse is on joint statements.", "Saves each person as it goes."],
        "params": [
            ["query", "query", "People to act on: a PeopleId, a saved search name, or Search Builder code", True, None],
            ["option", "integer", "Statement option code: 1 individual, 2 joint, 9 none, 0 clear", True, None]
        ],
        "returns": "None",
        "watch_out": ["A spouse is updated to match for codes 1, 2 and 9. Clearing with 0 clears the spouse only when the spouse is on joint statements.", "Changes person records.", "Codes other than 0, 1, 2 and 9 are saved on the person only, not the spouse."],
        "example": "print model.GetPerson(model.UserPeopleId).ContributionOptionsId\n# Uncomment to set joint statements for you and your spouse:\n# model.UpdateContributionOption(str(model.UserPeopleId), 2)",
        "see_also": ["model.UpdateElectronicStatement", "model.UpdateEnvelopeOption"]
    },
    "model.UpdateElectronicStatement": {
        "category": "Contributions",
        "kind": "method",
        "arity": [2, 2],
        "tags": ["writes_data"],
        "summary": "Turns electronic giving statements on (True) or off (False) for everyone in a query.",
        "details": "Accepts a PeopleId, a saved search name, or Search Builder code.",
        "flow": ["Runs the query to get the list of people.", "Sets the electronic statement flag on each person.", "If that person is on joint statements, sets the same flag on the spouse.", "Records the change in the change log and saves each person as it goes."],
        "params": [
            ["query", "query", "People to act on: a PeopleId, a saved search name, or Search Builder code", True, None],
            ["tf", "boolean", "True for electronic statements, False for printed", True, None]
        ],
        "returns": "None",
        "watch_out": ["A spouse on joint statements is updated too.", "Changes person records.", "The spouse is updated only when the person in the query is on joint statements; the spouse's own option is not checked."],
        "example": "print model.GetPerson(model.UserPeopleId).ElectronicStatement\n# Uncomment to switch you to electronic statements:\n# model.UpdateElectronicStatement(str(model.UserPeopleId), True)",
        "see_also": ["model.UpdateContributionOption", "model.UpdateEnvelopeOption"]
    },
    "model.UpdateEnvelopeOption": {
        "category": "Contributions",
        "kind": "method",
        "arity": [2, 2],
        "tags": ["writes_data"],
        "summary": "Sets the offering envelope option for everyone in a query.",
        "details": "Codes: 1 individual, 2 joint, 9 none, 0 clears it. Accepts a PeopleId, a saved search name, or Search Builder code.",
        "flow": ["Runs the query to get the list of people.", "For each person, sets the envelope option, where 0 clears it, and records the change in the change log.", "If the code is 1, 2 or 9, the spouse gets the same value.", "If the code is 0, the spouse's envelope option is cleared only when the spouse's statement option is joint.", "Saves each person as it goes."],
        "params": [
            ["query", "query", "People to act on: a PeopleId, a saved search name, or Search Builder code", True, None],
            ["option", "integer", "Envelope option code: 1 individual, 2 joint, 9 none, 0 clear", True, None]
        ],
        "returns": "None",
        "watch_out": ["A spouse is updated to match for codes 1, 2 and 9. Clearing with 0 clears the spouse only when the spouse is on joint statements.", "Changes person records.", "Clearing with 0 checks the spouse's statement option, not the envelope option, to decide whether to clear the spouse."],
        "example": "print model.GetPerson(model.UserPeopleId).EnvelopeOptionsId\n# Uncomment to set individual envelopes for yourself:\n# model.UpdateEnvelopeOption(str(model.UserPeopleId), 1)",
        "see_also": ["model.UpdateContributionOption", "model.UpdateElectronicStatement"]
    },
    "q.ContributionCount": {
        "category": "Contributions",
        "kind": "method",
        "arity": [2, 3],
        "tags": ["read_only"],
        "summary": "Counts recorded contributions between days1 days ago and days2 days ago.",
        "details": "A days2 of 0 means no end date. It uses the same exclusions as q.ContributionTotals. A two argument form counts the last N days. Funds: 0 or \"\" for all, a comma separated list, negative ids to exclude.",
        "params": [
            ["days1", "integer", "Start, in days before today", True, None],
            ["days2", "integer", "End, in days before today (0 for no end)", True, None],
            ["funds", "any", "Fund ids: 0 or \"\" for all, a comma separated list, negative ids to exclude", True, None]
        ],
        "returns": "Integer count",
        "watch_out": ["Only recorded contributions count. Pledges, returned checks (6) and reversals (7) are left out, but event fees (99) and non tax deductible gifts are counted.", "The two argument form compares against the current time, not midnight, so the oldest day is only partly included.", "Passing None for funds raises an error. Use 0 or \"\" for all funds."],
        "example": "print q.ContributionCount(30, 0, 0), \"gifts in the last 30 days\"\nprint q.ContributionCount(7, 0), \"gifts in the last 7 days\"",
        "signatures": ["q.ContributionCount(days1, days2, funds)", "q.ContributionCount(days, funds)"],
        "see_also": ["q.ContributionTotals", "q.DateRangeForContributionTotals"]
    },
    "q.ContributionTotals": {
        "category": "Contributions",
        "kind": "method",
        "arity": [3, 3],
        "tags": ["read_only"],
        "summary": "Sums recorded contributions between days1 days ago and days2 days ago.",
        "details": "A days2 of 0 means no end date. Excludes pledges, returned checks and reversals. Funds: 0 or \"\" for all, a comma separated list, negative ids to exclude.",
        "params": [
            ["days1", "integer", "Start, in days before today", True, None],
            ["days2", "integer", "End, in days before today (0 for no end)", True, None],
            ["funds", "any", "Fund ids: 0 or \"\" for all, a comma separated list, negative ids to exclude", True, None]
        ],
        "returns": "Float total (0 when nothing matches)",
        "watch_out": ["It includes other types such as non tax deductible gifts and event fees (99).", "Only recorded contributions count.", "Returns 0 when nothing matches.", "Passing None for funds raises an error. Use 0 or \"\" for all funds."],
        "example": "total = q.ContributionTotals(30, 0, 0)\nprint \"Last 30 days: ${:,.2f}\".format(total)\nprint q.DateRangeForContributionTotals(30, 0)",
        "see_also": ["q.ContributionCount", "q.DateRangeForContributionTotals"]
    },
    "q.DateRangeForContributionTotals": {
        "category": "Contributions",
        "kind": "method",
        "arity": [2, 2],
        "tags": ["read_only"],
        "summary": "Returns a readable label for the date range q.ContributionTotals uses with the same day offsets.",
        "params": [
            ["days1", "integer", "Start, in days before today", True, None],
            ["days2", "integer", "End, in days before today", True, None]
        ],
        "returns": "String like \"from 10/3/2026 up to 11/2/2026\"",
        "watch_out": ["With days2 of 0 the label ends at today's date, but q.ContributionTotals treats 0 as no end date, so future dated gifts are in the total but not in the label.", "The date format follows the server's short date format."],
        "example": "print \"Contributions \" + q.DateRangeForContributionTotals(30, 0)",
        "see_also": ["q.ContributionTotals", "q.ContributionCount"]
    },
    "DynamicData.AddValue": {
        "category": "Data & JSON",
        "kind": "method",
        "summary": "Adds or replaces a value on a DynamicData object.",
        "details": "Same as setting an attribute, but works with any key text.",
        "params": [
            ["name", "string", "Key to add or update", True, None],
            ["value", "any", "Value to store", True, None]
        ],
        "returns": "None",
        "watch_out": ["Ignored when the key is empty.", "A None key is ignored too."],
        "example": "data = model.DynamicData()\ndata.AddValue(\"Status\", \"Active\")\ndata.AddValue(\"Start Date\", model.DateTime)\nprint data.Status, data.GetValue(\"Start Date\")",
        "see_also": ["DynamicData.SetValue", "DynamicData.GetValue", "DynamicData.Remove", "model.DynamicData"]
    },
    "DynamicData.GetValue": {
        "category": "Data & JSON",
        "kind": "method",
        "tags": ["read_only"],
        "summary": "Gets a value from a DynamicData object by key.",
        "details": "Useful for keys that are not valid Python names.",
        "params": [
            ["key", "string", "Key to look up", True, None]
        ],
        "returns": "any or None",
        "watch_out": ["Returns None when the key is missing, while attribute access returns \"\" instead."],
        "example": "data = model.DynamicData({\"First Name\": \"John\"})\nprint \"First Name:\", data.GetValue(\"First Name\")\nprint \"Missing key gives:\", data.GetValue(\"Missing\")",
        "see_also": ["DynamicData.AddValue", "DynamicData.Keys", "model.DynamicData"]
    },
    "DynamicData.Keys": {
        "category": "Data & JSON",
        "kind": "method",
        "tags": ["read_only"],
        "summary": "Returns the list of keys in a DynamicData object.",
        "details": "If a metadata DynamicData is passed, it returns the metadata's keys instead, skipping any whose type text has hidden, special or readonly as its first word.",
        "params": [
            ["metadata", "DynamicData", "Optional metadata object mapping keys to type text.", False, "None"]
        ],
        "returns": "list of strings",
        "watch_out": ["Keys whose value is another DynamicData are skipped."],
        "example": "data = model.DynamicData({\"Name\": \"John\", \"Age\": 30})\nfor key in data.Keys():\n    print key, \"=\", data.GetValue(key)",
        "signatures": ["data.Keys()", "data.Keys(metadata)"],
        "see_also": ["DynamicData.SpecialKeys", "DynamicData.GetValue", "model.DynamicData"]
    },
    "DynamicData.Remove": {
        "category": "Data & JSON",
        "kind": "method",
        "summary": "Removes a key from a DynamicData object.",
        "details": "Does nothing if the key is not there.",
        "params": [
            ["name", "string", "Key to remove", True, None]
        ],
        "returns": "None",
        "example": "data = model.DynamicData({\"Name\": \"John\", \"Temp\": 1})\ndata.Remove(\"Temp\")\nprint data.Keys()",
        "see_also": ["DynamicData.AddValue", "DynamicData.Keys"]
    },
    "DynamicData.SetValue": {
        "category": "Data & JSON",
        "kind": "method",
        "summary": "Stores a string value on a DynamicData object, converting it to a number when it looks like one.",
        "details": "Text of only digits becomes an integer, decimal text becomes a decimal, anything else stays a string.",
        "params": [
            ["name", "string", "Key to add or update", True, None],
            ["value", "string", "String value to convert and store", True, None]
        ],
        "returns": "None",
        "watch_out": ["Numeric looking text is stored as a number, not a string.", "Digit only text too large for an integer, such as a 10 digit phone number, is stored as 0.", "Leading zeros are lost, so a zip code like \"01234\" becomes 1234.", "Text with commas such as \"1,000\" and negative numbers become decimals.", "Store these with AddValue or attribute assignment to keep them as text."],
        "example": "data = model.DynamicData()\ndata.SetValue(\"Count\", \"123\")\ndata.SetValue(\"Amount\", \"19.95\")\ndata.SetValue(\"Note\", \"abc\")\nprint type(data.Count), type(data.Amount), type(data.Note)",
        "see_also": ["DynamicData.AddValue", "DynamicData.GetValue"]
    },
    "DynamicData.SpecialKeys": {
        "category": "Data & JSON",
        "kind": "method",
        "tags": ["read_only"],
        "summary": "Returns the keys of a metadata DynamicData whose value text starts with \"special \".",
        "details": "For example \"special text\".",
        "params": [
            ["metadata", "DynamicData", "Metadata object mapping keys to type text", True, None]
        ],
        "returns": "list of strings",
        "watch_out": ["The object it is called on is not used.", "The type text must start with \"special\" followed by a space. A value of just \"special\" is not matched.", "A metadata value of None raises an error."],
        "example": "meta = model.DynamicData({\"Name\": \"text\", \"Priority\": \"special int\", \"Notes\": \"hidden text\"})\ndata = model.DynamicData()\nprint data.SpecialKeys(meta)\nprint data.Keys(meta)",
        "see_also": ["DynamicData.Keys"]
    },
    "DynamicData.ToFlatString": {
        "category": "Data & JSON",
        "kind": "method",
        "summary": "Returns the contents as single line JSON with empty values left out and single quotes doubled for SQL.",
        "details": "The doubled quotes make it safe to place inside a SQL string.",
        "returns": "string (JSON)",
        "watch_out": ["It also removes the empty keys from the object itself.", "None values are dropped too, not just empty strings."],
        "example": "data = model.DynamicData({\"Name\": \"O'Brien\", \"Note\": \"\"})\nprint data.ToFlatString()\nprint data.Keys()",
        "see_also": ["DynamicData.ToString", "model.JsonSerialize"]
    },
    "DynamicData.ToString": {
        "category": "Data & JSON",
        "kind": "method",
        "tags": ["read_only"],
        "summary": "Returns the contents of a DynamicData object as indented JSON text.",
        "returns": "string (JSON)",
        "example": "data = model.DynamicData({\"Name\": \"John\", \"Age\": 30})\nprint data.ToString()",
        "see_also": ["DynamicData.ToFlatString", "model.FormatJson", "model.JsonSerialize"]
    },
    "model.AddUpdateJsonRecord": {
        "category": "Data & JSON",
        "kind": "method",
        "arity": [3, 6],
        "tags": ["writes_data"],
        "summary": "Saves a JSON record in the custom.JsonDocumentRecords table, keyed by a section name and up to four ids.",
        "details": "Inserts a new row or replaces the JSON of an existing one. Accepts a JSON string or a DynamicData object. Read records back with SQL.",
        "flow": ["Turns each key into text. Keys you leave out are stored as \"\".", "Removes keys with empty string values from your data, changing your DynamicData object too.", "Converts the data to indented JSON.", "Updates the custom.JsonDocumentRecords row with the same section and four keys, or inserts a new row."],
        "params": [
            ["data", "any", "JSON string or DynamicData object", True, None],
            ["section", "string", "Section name (up to 50 characters)", True, None],
            ["pk1", "any", "First key (stored as text)", True, None],
            ["pk2", "any", "Second key", False, "None"],
            ["pk3", "any", "Third key", False, "None"],
            ["pk4", "any", "Fourth key", False, "None"]
        ],
        "returns": "None",
        "watch_out": ["Empty string values are dropped.", "None values are kept as null in the JSON. Only empty strings are dropped.", "A JSON string must be a single object. A JSON array raises an error.", "pk1 of None raises an error.", "Section and keys are limited to 50 characters each; longer values are cut off, which can make two records collide.", "Keys are compared as text, so 5 and \"5\" are the same record."],
        "example": "# Writes a small demo record keyed by your PeopleId\ndd = model.DynamicData()\ndd.Note = \"Called on Monday\"\nmodel.AddUpdateJsonRecord(dd, \"ApiExplorerDemo\", model.UserPeopleId)\nfor r in q.QuerySql(\"SELECT Id1, Json FROM custom.JsonDocumentRecords WHERE Section = 'ApiExplorerDemo'\"):\n    print r.Id1, r.Json",
        "signatures": ["model.AddUpdateJsonRecord(json, section, pk1, pk2=None, pk3=None, pk4=None)", "model.AddUpdateJsonRecord(dynamicData, section, pk1, pk2=None, pk3=None, pk4=None)"],
        "see_also": ["model.DeleteJsonRecord", "model.DeleteJsonRecordSection", "q.QuerySql", "model.DynamicData"]
    },
    "model.CsvReader": {
        "category": "Data & JSON",
        "kind": "method",
        "arity": [1, 2],
        "tags": ["read_only"],
        "summary": "Creates a CsvHelper reader over CSV text and reads the header row, so fields can be read by column name.",
        "details": "Call Read() to move to each row. The optional second argument sets the delimiter.",
        "params": [
            ["text", "string", "CSV text, header row first", True, None],
            ["delimiter", "string", "Delimiter such as \";\" or \"\\t\"", False, "\",\""]
        ],
        "returns": "CsvHelper.CsvReader",
        "watch_out": ["The header row is already read, so the first Read() moves to the first data row.", "Numbers and dates are read using the server's culture settings."],
        "example": "text = \"Name,Age\\nJohn,40\\nMary,35\"\ncsv = model.CsvReader(text)\nwhile csv.Read():\n    print csv[\"Name\"], csv[\"Age\"]",
        "signatures": ["model.CsvReader(text)", "model.CsvReader(text, delimiter)"],
        "see_also": ["model.CsvReaderNoHeader"]
    },
    "model.CsvReaderNoHeader": {
        "category": "Data & JSON",
        "kind": "method",
        "arity": [1, 1],
        "tags": ["read_only"],
        "summary": "Creates a CsvHelper reader over CSV text that has no header row.",
        "details": "Call Read() to move to each row and read fields by position starting at 0.",
        "params": [
            ["text", "string", "CSV text with no header row", True, None]
        ],
        "returns": "CsvHelper.CsvReader",
        "watch_out": ["There is no delimiter argument, so it always splits on commas. Use model.CsvReader for other delimiters."],
        "example": "csv = model.CsvReaderNoHeader(\"John,40\\nMary,35\")\nwhile csv.Read():\n    print csv[0], csv[1]",
        "see_also": ["model.CsvReader"]
    },
    "model.DeleteJsonRecord": {
        "category": "Data & JSON",
        "kind": "method",
        "arity": [2, 5],
        "tags": ["writes_data"],
        "summary": "Deletes one record from custom.JsonDocumentRecords.",
        "details": "The section and all keys must match what was saved. Unused keys are empty.",
        "params": [
            ["section", "string", "Section name", True, None],
            ["pk1", "any", "First key", True, None],
            ["pk2", "any", "Second key", False, "None"],
            ["pk3", "any", "Third key", False, "None"],
            ["pk4", "any", "Fourth key", False, "None"]
        ],
        "returns": "None",
        "watch_out": ["Raises no error when no record matches.", "pk1 of None raises an error.", "Keys are compared as text, so 5 and \"5\" match the same record."],
        "example": "model.DeleteJsonRecord(\"ApiExplorerDemo\", model.UserPeopleId)\nprint q.QuerySqlInt(\"SELECT COUNT(*) FROM custom.JsonDocumentRecords WHERE Section = 'ApiExplorerDemo'\"), \"left\"",
        "see_also": ["model.AddUpdateJsonRecord", "model.DeleteJsonRecordSection"]
    },
    "model.DeleteJsonRecordSection": {
        "category": "Data & JSON",
        "kind": "method",
        "arity": [1, 1],
        "tags": ["writes_data"],
        "summary": "Deletes every record in a section of custom.JsonDocumentRecords.",
        "params": [
            ["section", "string", "Section name", True, None]
        ],
        "returns": "None",
        "watch_out": ["There is no confirmation.", "It cannot be undone.", "Raises no error when the section is empty."],
        "example": "print q.QuerySqlInt(\"SELECT COUNT(*) FROM custom.JsonDocumentRecords WHERE Section = 'ApiExplorerDemo'\"), \"records\"\n# Deletes data. Uncomment to clear the whole section:\n# model.DeleteJsonRecordSection(\"ApiExplorerDemo\")",
        "see_also": ["model.DeleteJsonRecord", "model.AddUpdateJsonRecord"]
    },
    "model.DynamicData": {
        "category": "Data & JSON",
        "kind": "method",
        "arity": [0, 1],
        "tags": ["read_only"],
        "summary": "Creates a DynamicData object: a dictionary you can use with attribute syntax (data.Name).",
        "details": "Call it with no argument for an empty object, with a Python dict to copy its entries, or with another DynamicData to copy it. A DynamicData can be passed to q.QuerySql as named SQL parameters.",
        "params": [
            ["dict", "dict", "Optional Python dict or DynamicData to copy", False, "None"]
        ],
        "returns": "DynamicData",
        "watch_out": ["Reading a missing attribute returns an empty string, not an error.", "Reading a missing key with brackets, data['Name'], returns None, while data.Name returns an empty string.", "Keys from a Python dict are turned into strings, so a key of 5 becomes '5'."],
        "example": "data = model.DynamicData()\ndata.Name = \"John\"\ndata.Age = 30\nprint data.Name, data.Age\ndata2 = model.DynamicData({\"Name\": \"Jane\", \"Age\": 25})\nprint data2.Name, repr(data2.Missing)",
        "signatures": ["model.DynamicData()", "model.DynamicData(dict)", "model.DynamicData(dynamicData)"],
        "see_also": ["model.DynamicDataFromJson", "DynamicData.AddValue", "q.QuerySql", "DynamicData.Keys"]
    },
    "model.DynamicDataFromJson": {
        "category": "Data & JSON",
        "kind": "method",
        "arity": [1, 1],
        "tags": ["read_only"],
        "summary": "Parses a JSON object string into a DynamicData object.",
        "details": "Read values as attributes.",
        "params": [
            ["json", "string", "JSON object text", True, None]
        ],
        "returns": "DynamicData",
        "watch_out": ["A missing attribute returns an empty string.", "Values are converted as they are read: whole numbers become integers, decimals become decimal numbers, ISO date strings become DateTime, nested objects become DynamicData, and arrays of objects become lists of DynamicData.", "true, false and null stay as JSON value objects, not Python True, False or None. Compare with str(value), for example str(data.Active) == 'True'.", "A whole number larger than 2147483647 raises an error."],
        "example": "data = model.DynamicDataFromJson('{\"name\": \"John\", \"age\": 30}')\nprint data.name, data.age",
        "see_also": ["model.DynamicDataFromJsonArray", "model.JsonDeserialize", "model.DynamicData"]
    },
    "model.DynamicDataFromJsonArray": {
        "category": "Data & JSON",
        "kind": "method",
        "arity": [1, 1],
        "tags": ["read_only"],
        "summary": "Parses a JSON array of objects into a list of DynamicData objects.",
        "params": [
            ["json", "string", "JSON array text", True, None]
        ],
        "returns": "list of DynamicData",
        "watch_out": ["Each item is converted the same way as model.DynamicDataFromJson: ISO date strings become DateTime, nested objects become DynamicData, and true, false and null stay as JSON value objects.", "Raises an error if the text is not a JSON array of objects, or if a whole number is larger than 2147483647."],
        "example": "items = model.DynamicDataFromJsonArray('[{\"id\": 1, \"name\": \"A\"}, {\"id\": 2, \"name\": \"B\"}]')\nfor item in items:\n    print item.id, item.name",
        "see_also": ["model.DynamicDataFromJson", "model.JsonDeserialize2", "model.ElementList"]
    },
    "model.ElementList": {
        "category": "Data & JSON",
        "kind": "method",
        "arity": [2, 2],
        "tags": ["read_only"],
        "summary": "Pulls one named value out of each DynamicData in a list and returns the values as a list of strings.",
        "params": [
            ["array", "list", "List of DynamicData objects, such as from model.DynamicDataFromJsonArray", True, None],
            ["name", "string", "Key to pull from each item", True, None]
        ],
        "returns": "list of strings",
        "watch_out": ["Raises an error if any item is missing that key.", "Every value is turned into text, so numbers come back as strings such as '42'."],
        "example": "items = model.DynamicDataFromJsonArray('[{\"name\": \"A\"}, {\"name\": \"B\"}]')\nnames = model.ElementList(items, \"name\")\nprint \", \".join(names)",
        "see_also": ["model.DynamicDataFromJsonArray"]
    },
    "model.FormatJson": {
        "category": "Data & JSON",
        "kind": "method",
        "arity": [1, 1],
        "tags": ["read_only"],
        "summary": "Returns JSON with indentation, for display.",
        "details": "Accepts a JSON string, a DynamicData object, or a list of DynamicData.",
        "params": [
            ["json", "any", "JSON string, DynamicData, or list of DynamicData", True, None]
        ],
        "returns": "string",
        "watch_out": ["A plain Python dict is not accepted; pass model.JsonSerialize(d) instead.", "Raises an error when the string is not valid JSON."],
        "example": "d = {\"name\": \"John\", \"tags\": [1, 2]}\nprint \"<pre>\" + model.FormatJson(model.JsonSerialize(d)) + \"</pre>\"",
        "signatures": ["model.FormatJson(json)", "model.FormatJson(dynamicData)", "model.FormatJson(listOfDynamicData)"],
        "see_also": ["model.JsonSerialize"]
    },
    "model.FromJson": {
        "category": "Data & JSON",
        "kind": "method",
        "arity": [1, 1],
        "tags": ["read_only", "deprecated"],
        "summary": "Parses a JSON object string into a DynamicData object.",
        "params": [
            ["json", "string", "JSON object text", True, None]
        ],
        "returns": "DynamicData object",
        "watch_out": ["Deprecated. Use model.DynamicDataFromJson instead.", "Unlike DynamicDataFromJson, nested objects and arrays are not converted. They stay as raw JSON objects (read them with obj['key']), and dates stay as text."],
        "example": "# Deprecated. Same result with the replacement:\nd = model.DynamicDataFromJson('{\"name\": \"John\", \"age\": 30}')\nprint d.name, d.age",
        "see_also": ["model.DynamicDataFromJson", "model.DynamicData", "model.JsonDeserialize"]
    },
    "model.JsonDeserialize": {
        "category": "Data & JSON",
        "kind": "method",
        "arity": [1, 1],
        "tags": ["read_only"],
        "summary": "Parses a JSON object string into a Newtonsoft JObject.",
        "details": "Read values with obj[\"key\"].",
        "params": [
            ["s", "string", "JSON object text", True, None]
        ],
        "returns": "JObject",
        "watch_out": ["Raises an error if the text is not a JSON object. Use model.DynamicDataFromJsonArray for arrays.", "Each value is a JSON token, not a plain Python value. Use str(obj['key']) or obj['key'].Value to get text or a number."],
        "example": "obj = model.JsonDeserialize('{\"name\": \"John\", \"age\": 30}')\nprint obj[\"name\"], obj[\"age\"]",
        "see_also": ["model.DynamicDataFromJson", "model.DynamicDataFromJsonArray", "model.JsonDeserialize2", "model.JsonSerialize"]
    },
    "model.JsonDeserialize2": {
        "category": "Data & JSON",
        "kind": "method",
        "arity": [1, 1],
        "tags": ["read_only"],
        "summary": "Parses a JSON array of objects into a sequence of DynamicData objects.",
        "params": [
            ["s", "string", "JSON array of objects", True, None]
        ],
        "returns": "sequence of DynamicData",
        "watch_out": ["The result is a lazy sequence, not a list. Wrap it in list() before using len() or an index such as [0]; each loop over it builds new objects.", "Nested objects and arrays inside each item are not converted to DynamicData. They stay as raw JSON objects. Use model.DynamicDataFromJsonArray for full conversion."],
        "example": "rows = model.JsonDeserialize2('[{\"id\": 1}, {\"id\": 2}]')\nfor r in rows:\n    print r.id",
        "see_also": ["model.DynamicDataFromJsonArray", "model.JsonDeserialize"]
    },
    "model.JsonSerialize": {
        "category": "Data & JSON",
        "kind": "method",
        "arity": [1, 1],
        "tags": ["read_only"],
        "summary": "Converts an object (dict, list, DynamicData or .NET object) to a JSON string with Newtonsoft.",
        "params": [
            ["o", "any", "Object to serialize", True, None]
        ],
        "returns": "string",
        "watch_out": ["Like json.dumps, very large .NET objects such as a person record can be slow or fail.", "Dates are written in ISO form such as \"2026-11-02T09:15:00\", not the US form you get from str()."],
        "example": "print model.JsonSerialize({\"id\": model.UserPeopleId, \"tags\": [\"a\", \"b\"]})",
        "see_also": ["model.JsonDeserialize", "model.FormatJson"]
    },
    "model.AddDays": {
        "category": "Dates",
        "kind": "method",
        "arity": [2, 2],
        "tags": ["read_only"],
        "summary": "Adds a number of days to a date and returns a short date string such as \"11/9/2026\".",
        "details": "Accepts a DateTime or a date string. Use DateAddDays when you need a DateTime back.",
        "params": [
            ["dt1", "any", "A date: a .NET DateTime, a Python datetime, or a string such as \"12/25/2026\"", True, None],
            ["days", "integer", "Days to add. Use a negative number to go back.", True, None]
        ],
        "returns": "string (short date, or \"1/1/1\" when the date is not valid)",
        "watch_out": ["Returns \"1/1/1\" instead of raising an error when the date cannot be read.", "The short date follows the server's US format, so it is \"M/D/YYYY\" with no leading zeros."],
        "example": "print model.AddDays(model.DateTime, 7)\nprint model.AddDays(\"12/25/2026\", -1)",
        "see_also": ["model.DateAddDays", "model.DateAddHours", "model.FormatDate"]
    },
    "model.AgeInMonths": {
        "category": "Dates",
        "kind": "method",
        "arity": [2, 2],
        "tags": ["read_only"],
        "summary": "Calculates whole months between a birth date and an as of date.",
        "params": [
            ["birthdate", "date", "Birth date (may be None)", True, None],
            ["asof", "date", "Date to measure to", True, None]
        ],
        "returns": "Integer number of months, or None",
        "watch_out": ["Returns None when the birth date is missing or later than the as of date.", "Pass .NET DateTime values (System.DateTime), not Python datetime objects."],
        "example": "from System import DateTime\nmonths = model.AgeInMonths(DateTime(2024, 3, 15), DateTime.Now)\nprint \"Age in months:\", months",
        "see_also": ["model.DateDiffDays", "model.ParseDate"]
    },
    "model.DateAddDays": {
        "category": "Dates",
        "kind": "method",
        "arity": [2, 2],
        "tags": ["read_only"],
        "summary": "Adds a number of days to a date and returns a DateTime.",
        "details": "Accepts a DateTime or a date string.",
        "params": [
            ["dt", "any", "A date: a .NET DateTime, a Python datetime, or a string such as \"12/25/2026\"", True, None],
            ["days", "integer", "Days to add. Use a negative number to go back.", True, None]
        ],
        "returns": "DateTime",
        "watch_out": ["Raises \"bad date\" if the date cannot be read."],
        "example": "later = model.DateAddDays(model.DateTime, 30)\nprint later.ToShortDateString()\nprint model.DateAddDays(\"1/1/2026\", -1).ToShortDateString()",
        "see_also": ["model.AddDays", "model.DateAddHours", "model.DateDiffDays"]
    },
    "model.DateAddHours": {
        "category": "Dates",
        "kind": "method",
        "arity": [2, 2],
        "tags": ["read_only"],
        "summary": "Adds a number of hours to a date and time and returns a DateTime.",
        "details": "Accepts a DateTime or a date string.",
        "params": [
            ["dt", "any", "A date: a .NET DateTime, a Python datetime, or a string such as \"12/25/2026\"", True, None],
            ["hours", "integer", "Hours to add. Use a negative number to go back.", True, None]
        ],
        "returns": "DateTime",
        "watch_out": ["Raises \"bad date\" if the date cannot be read.", "A date string with no time starts at midnight."],
        "example": "later = model.DateAddHours(model.DateTime, 2)\nprint later\nprint model.DateAddHours(\"11/2/2026 8:00 AM\", -3)",
        "see_also": ["model.DateAddDays", "model.AddDays"]
    },
    "model.DateDiffDays": {
        "category": "Dates",
        "kind": "method",
        "arity": [2, 2],
        "tags": ["read_only"],
        "summary": "Returns the number of whole days from the first date to the second (dt2 minus dt1).",
        "details": "The result is negative when the second date is earlier. Partial days are dropped.",
        "params": [
            ["dt1", "any", "Start date (DateTime or date string)", True, None],
            ["dt2", "any", "End date (DateTime or date string)", True, None]
        ],
        "returns": "integer",
        "watch_out": ["Raises \"bad date\" if either date cannot be read.", "The time of day counts. 11 PM one day to 1 AM the next is 0 days, not 1."],
        "example": "days = model.DateDiffDays(\"1/1/2026\", model.DateTime)\nprint \"Days since January 1:\", days\nprint model.DateDiffDays(\"12/25/2026\", \"12/1/2026\")",
        "see_also": ["model.DateAddDays", "model.AgeInMonths"]
    },
    "model.DateTime": {
        "category": "Dates",
        "kind": "property",
        "tags": ["read_only"],
        "summary": "The current date and time on the server.",
        "details": "Read only. It honors date simulation (SetToday) when that is turned on for your session.",
        "returns": "DateTime",
        "watch_out": ["str() of the value is US format such as \"11/2/2026 9:15:00 AM\", not ISO."],
        "example": "now = model.DateTime\nprint now\nprint now.Year, now.Month, now.Day\nprint now.ToString(\"yyyy-MM-dd\")",
        "see_also": ["model.SetToday", "model.ResetToday", "model.FormatDate", "model.DayOfWeek"]
    },
    "model.DayOfWeek": {
        "category": "Dates",
        "kind": "property",
        "tags": ["read_only"],
        "summary": "Today's day of the week as a number, where 0 is Sunday and 6 is Saturday.",
        "details": "Read only.",
        "returns": "integer (0 to 6)",
        "watch_out": ["It honors date simulation (SetToday), like model.DateTime."],
        "example": "names = [\"Sunday\", \"Monday\", \"Tuesday\", \"Wednesday\", \"Thursday\", \"Friday\", \"Saturday\"]\nprint model.DayOfWeek, names[model.DayOfWeek]\nif model.DayOfWeek == 0:\n    print \"It is Sunday\"",
        "see_also": ["model.DateTime", "model.SundayForDate"]
    },
    "model.FormatDate": {
        "category": "Dates",
        "kind": "method",
        "arity": [1, 1],
        "tags": ["read_only"],
        "summary": "Formats a .NET DateTime as a short date string such as \"11/2/2026\".",
        "details": "Pass a value from model.DateTime or another model date method.",
        "params": [
            ["dt", "date", "The date to format", True, None]
        ],
        "returns": "string",
        "watch_out": ["The argument must be a DateTime, not a string."],
        "example": "print model.FormatDate(model.DateTime)\nprint model.FormatDate(model.DateAddDays(model.DateTime, 7))",
        "see_also": ["model.DateTime", "model.ParseDate", "model.AddDays"]
    },
    "model.MostRecentAttendedSunday": {
        "category": "Dates",
        "kind": "method",
        "arity": [1, 1],
        "tags": ["read_only"],
        "summary": "Returns the most recent Sunday before now that has a meeting with recorded attendance.",
        "details": "Recorded attendance means MaxCount above 0. Pass a program id to limit the search to that program, or 0 for all programs. The time portion is midnight.",
        "params": [
            ["progid", "integer", "Program id, or 0 for all programs", True, None]
        ],
        "returns": "DateTime",
        "watch_out": ["If no such meeting exists it returns the most recent Sunday on or before today.", "Any meeting counts, including inactive involvements, as long as it was on a Sunday and has a head count above 0. A meeting earlier today counts when today is Sunday."],
        "example": "sunday = model.MostRecentAttendedSunday(0)\nprint \"Last attended Sunday:\", sunday.ToShortDateString()",
        "see_also": ["model.SundayForDate", "q.LastSunday", "q.LastWeekAttendance"]
    },
    "model.ParseDate": {
        "category": "Dates",
        "kind": "method",
        "arity": [1, 1],
        "tags": ["read_only"],
        "summary": "Parses a date string into a DateTime.",
        "details": "Accepts the usual formats such as \"5/30/2026\", \"2026-05-30\" or \"5/30/2026 2:00 PM\", plus 8 digit \"MMddyyyy\". Because it returns None on bad input, it is a safe way to validate user input.",
        "params": [
            ["dt", "string", "The date text to parse", True, None]
        ],
        "returns": "DateTime or None",
        "watch_out": ["Returns None when the text is not a valid date.", "Years before 1900 are pushed forward a century at a time, so \"1/1/0026\" becomes 1/1/1926.", "Passing None returns None."],
        "example": "d = model.ParseDate(\"5/30/2026\")\nif d:\n    print \"Parsed:\", d.ToShortDateString()\nprint \"Not a date gives:\", model.ParseDate(\"not a date\")",
        "see_also": ["model.ParseDateUTC", "model.FormatDate", "model.DateAddDays"]
    },
    "model.ParseDateUTC": {
        "category": "Dates",
        "kind": "method",
        "arity": [1, 1],
        "tags": ["read_only", "known_issue"],
        "summary": "Parses a UTC date string and converts it to the church's local time zone.",
        "params": [
            ["dt", "string", "A UTC date and time, such as \"2026-01-15 15:00:00\"", True, None]
        ],
        "returns": "DateTime or None",
        "watch_out": ["Returns None when the text is not a valid date.", "Pass the time without a trailing \"Z\". A \"Z\" makes .NET convert to server time first, so the value would be shifted twice.", "The daylight saving offset used is the one in effect today, not on the date being converted. A summer time converted in winter (or the reverse) comes out one hour off."],
        "example": "local = model.ParseDateUTC(\"2026-01-15 15:00:00\")\nprint \"Local time:\", local",
        "see_also": ["model.ParseDate", "model.DateTime"]
    },
    "model.ResetToday": {
        "category": "Dates",
        "kind": "method",
        "arity": [0, 0],
        "summary": "Clears a simulated date set with SetToday so dates go back to the real current date.",
        "details": "The simulated date lives in your web session, so call this when you are done testing.",
        "returns": "None",
        "watch_out": ["Changes your web session state.", "It clears both the simulated date and any simulated time. It does not turn date simulation off for your session."],
        "example": "model.ResetToday()\nprint \"Today is\", model.DateTime.ToShortDateString()",
        "see_also": ["model.SetToday", "model.DateTime"]
    },
    "model.ScheduledTime": {
        "category": "Dates",
        "kind": "property",
        "tags": ["read_only"],
        "summary": "The time slot of a ScheduledTasks run in 24 hour \"HHmm\" form, such as \"0600\" or \"1830\".",
        "details": "Read only. It is set when a script runs from the ScheduledTasks job.",
        "returns": "string or None",
        "watch_out": ["It is None when a script runs any other way.", "It is not passed along to scripts started with model.CallScript."],
        "example": "t = model.ScheduledTime\nif t is None:\n    print \"Not running from ScheduledTasks\"\nelse:\n    print \"Scheduled slot:\", t\n    if model.DayOfWeek == 1 and t == \"0600\":\n        print \"Monday 6 AM job would run here\"",
        "see_also": ["model.CallScript", "model.FromMorningBatch", "model.DateTime"]
    },
    "model.SetToday": {
        "category": "Dates",
        "kind": "method",
        "arity": [1, 1],
        "summary": "Simulates a different current date for testing.",
        "details": "The value is stored in your web session.",
        "params": [
            ["dt", "any", "A date: a .NET DateTime, a Python datetime, or a string such as \"12/25/2026\"", True, None]
        ],
        "returns": "None",
        "watch_out": ["It only takes effect when date simulation is turned on for your session and not disabled on the server. Otherwise dates are unchanged.", "It affects other pages until you call ResetToday.", "Raises \"bad date\" if the date cannot be read.", "Only the date is simulated. model.DateTime keeps the real time of day."],
        "example": "model.SetToday(\"12/25/2026\")\nprint \"Now reads as\", model.DateTime\nmodel.ResetToday()\nprint \"Back to\", model.DateTime",
        "see_also": ["model.ResetToday", "model.DateTime"]
    },
    "model.SundayForDate": {
        "category": "Dates",
        "kind": "method",
        "arity": [1, 1],
        "tags": ["read_only"],
        "summary": "Returns the Sunday that starts the week containing the given date.",
        "details": "Returns the same day if it is already a Sunday. The time is set to midnight.",
        "params": [
            ["dt", "any", "A date: a .NET DateTime, a Python datetime, or a string such as \"12/25/2026\"", True, None]
        ],
        "returns": "DateTime",
        "watch_out": ["Raises \"bad date\" if the date cannot be read."],
        "example": "sunday = model.SundayForDate(model.DateTime)\nprint \"This week started\", sunday.ToShortDateString()\nprint model.SundayForDate(\"12/25/2026\").ToShortDateString()",
        "see_also": ["model.SundayForWeek", "model.WeekNumber", "model.WeekOfMonth", "model.MostRecentAttendedSunday"]
    },
    "model.SundayForWeek": {
        "category": "Dates",
        "kind": "method",
        "arity": [2, 2],
        "tags": ["read_only"],
        "summary": "Returns the Sunday for a numbered week of a year.",
        "details": "Week 1 is the first Sunday in January, and each later week adds 7 days.",
        "params": [
            ["year", "integer", "Year, such as 2026", True, None],
            ["week", "integer", "Week number, starting at 1", True, None]
        ],
        "returns": "DateTime",
        "watch_out": ["The week number is not checked. Week 53 or higher rolls into the next year, and week 0 gives the Sunday before the first Sunday of January."],
        "example": "print model.SundayForWeek(2026, 1).ToShortDateString()\nprint model.SundayForWeek(2026, model.WeekNumber(model.DateTime)).ToShortDateString()",
        "see_also": ["model.WeekNumber", "model.SundayForDate"]
    },
    "model.WeekNumber": {
        "category": "Dates",
        "kind": "method",
        "arity": [1, 1],
        "tags": ["read_only"],
        "summary": "Returns the week of the year for a date, with weeks starting on Sunday.",
        "details": "Week 1 is the first full week.",
        "params": [
            ["dt", "any", "A date: a .NET DateTime, a Python datetime, or a string such as \"12/25/2026\"", True, None]
        ],
        "returns": "integer (1 to 53)",
        "watch_out": ["Dates before the first full Sunday week of January return the last week number of the prior year (52 or 53).", "Raises \"bad date\" if the date cannot be read."],
        "example": "print \"Week\", model.WeekNumber(model.DateTime), \"of the year\"\nprint model.WeekNumber(\"1/1/2026\")",
        "see_also": ["model.SundayForWeek", "model.WeekOfMonth", "model.SundayForDate"]
    },
    "model.WeekOfMonth": {
        "category": "Dates",
        "kind": "method",
        "arity": [1, 1],
        "tags": ["read_only"],
        "summary": "Returns which week of the month a date falls in, counted by Sundays.",
        "details": "The date is first moved back to its Sunday, and the week is counted within that Sunday's month, so the first Sunday of the month is week 1. A date early in a month whose week began in the prior month counts in the prior month.",
        "params": [
            ["dt", "any", "A date: a .NET DateTime, a Python datetime, or a string such as \"12/25/2026\"", True, None]
        ],
        "returns": "integer (1 to 5)",
        "watch_out": ["Raises \"bad date\" if the date cannot be read.", "The month is the month of that Sunday, not of the date you pass. Tuesday December 1, 2026 moves back to Sunday November 29 and returns 5 (the 5th week of November)."],
        "example": "wk = model.WeekOfMonth(model.DateTime)\nprint \"Week of month:\", wk\nif wk == 1:\n    print \"First week of the month\"",
        "see_also": ["model.WeekNumber", "model.SundayForDate"]
    },
    "model.docusign": {
        "category": "DocuSign",
        "kind": "property",
        "tags": ["outside_service"],
        "summary": "A helper that creates DocuSign eSign SDK (version 4.5.2) objects so scripts can build and send envelopes.",
        "details": "Read only. Its methods are ApiClient, Configuration, EnvelopesApi, EnvelopeDefinition, Document, Recipients, Signer, SignHere, Text, Tabs, TemplateRole, TemplateTabs, RecipientViewRequest and ListStatusChangesOptions.",
        "flow": ["model.docusign gives one helper object per script run. Each of its methods only creates a new, empty DocuSign SDK object.", "Create the client with docusign.ApiClient(basePath). The default base path is DocuSign production (https://www.docusign.net/restapi); pass the demo address to test in a sandbox account.", "Wrap it with docusign.Configuration(client) and add your own OAuth access token to it using the DocuSign SDK. TouchPoint does not store or refresh tokens.", "Build the envelope from EnvelopeDefinition, Document, Signer, SignHere, Tabs and Recipients, then send it with docusign.EnvelopesApi(config) and the SDK's own create envelope call.", "Nothing is saved in TouchPoint. Envelope ids, status and signed documents stay in DocuSign unless your script stores them, and DocuSign errors are raised in your script."],
        "returns": "ApiDocuSign",
        "watch_out": ["You supply your own DocuSign account and OAuth token.", "Lists on DocuSign objects are .NET List types, so build them with System.Collections.Generic.List rather than a Python list.", "ApiClient points at DocuSign production by default. Pass the demo base path when testing."],
        "example": "from System.Collections.Generic import List\nds = model.docusign\nenv = ds.EnvelopeDefinition()\nenv.EmailSubject = \"Please sign\"\nenv.Status = \"created\"\nsigner = ds.Signer()\nsigner.Name = \"John Doe\"\nsigner.Email = \"john@example.com\"\nsigner.RecipientId = \"1\"\nprint env.EmailSubject, signer.Name",
        "see_also": ["model.docusign.ApiClient", "model.docusign.EnvelopesApi", "model.docusign.EnvelopeDefinition"]
    },
    "model.docusign.ApiClient": {
        "category": "DocuSign",
        "kind": "method",
        "tags": ["outside_service"],
        "summary": "Creates a DocuSign ApiClient (DocuSign.eSign 4.5.2) pointed at basePath.",
        "details": "Use it to get an access token and to build a Configuration. Use https://demo.docusign.net/restapi for a developer sandbox.",
        "params": [
            ["basePath", "string", "REST base url", False, "'https://www.docusign.net/restapi'"],
            ["proxy", "any", "Optional .NET WebProxy", False, "None"]
        ],
        "returns": "DocuSign ApiClient",
        "watch_out": ["You need your own DocuSign account, integration key and token."],
        "example": "client = model.docusign.ApiClient('https://demo.docusign.net/restapi')\nprint client.GetBasePath()",
        "signatures": ["model.docusign.ApiClient(basePath='https://www.docusign.net/restapi', proxy=None)"],
        "see_also": ["model.docusign.Configuration", "model.docusign.EnvelopesApi", "model.docusign"]
    },
    "model.docusign.Configuration": {
        "category": "DocuSign",
        "kind": "method",
        "tags": ["outside_service"],
        "summary": "Creates a DocuSign Configuration that holds the ApiClient and the authorization header.",
        "details": "Pass it to model.docusign.EnvelopesApi.",
        "params": [
            ["apiClient", "any", "ApiClient from model.docusign.ApiClient", False, "None"]
        ],
        "returns": "DocuSign Configuration",
        "example": "client = model.docusign.ApiClient('https://demo.docusign.net/restapi')\nconfig = model.docusign.Configuration(client)\naccess_token = 'YOUR_ACCESS_TOKEN'\nconfig.AddDefaultHeader('Authorization', 'Bearer ' + access_token)\nprint 'Configuration ready'",
        "signatures": ["model.docusign.Configuration(apiClient=None)"],
        "see_also": ["model.docusign.ApiClient", "model.docusign.EnvelopesApi"]
    },
    "model.docusign.Document": {
        "category": "DocuSign",
        "kind": "method",
        "tags": ["outside_service"],
        "summary": "Creates an empty DocuSign Document.",
        "details": "Set DocumentBase64 (the file as base64 text), Name, FileExtension and DocumentId, then add it to the envelope's Documents list.",
        "returns": "DocuSign.eSign.Model.Document",
        "watch_out": ["Lists on DocuSign objects are .NET List types, so build them with System.Collections.Generic.List rather than a Python list."],
        "example": "from System.Collections.Generic import List\ndoc = model.docusign.Document()\ndoc.Name = \"Contract.pdf\"\ndoc.FileExtension = \"pdf\"\ndoc.DocumentId = \"1\"\ndocs = List[type(doc)]()\ndocs.Add(doc)\nprint docs.Count, doc.Name",
        "see_also": ["model.docusign.EnvelopeDefinition", "model.docusign.SignHere"]
    },
    "model.docusign.EnvelopeDefinition": {
        "category": "DocuSign",
        "kind": "method",
        "tags": ["outside_service"],
        "summary": "Creates an empty DocuSign EnvelopeDefinition.",
        "details": "Set EmailSubject, Documents, Recipients and Status (\"sent\" to send, \"created\" for a draft), then pass it to EnvelopesApi.CreateEnvelope.",
        "returns": "DocuSign.eSign.Model.EnvelopeDefinition",
        "watch_out": ["Status \"sent\" sends the envelope for signature when it is created."],
        "example": "env = model.docusign.EnvelopeDefinition()\nenv.EmailSubject = \"Contract for signature\"\nenv.EmailBlurb = \"Please review and sign\"\nenv.Status = \"created\"\nprint env.EmailSubject, env.Status",
        "see_also": ["model.docusign.EnvelopesApi", "model.docusign.Document", "model.docusign.Recipients", "model.docusign.TemplateRole"]
    },
    "model.docusign.EnvelopesApi": {
        "category": "DocuSign",
        "kind": "method",
        "tags": ["outside_service", "sends_message"],
        "summary": "Creates the DocuSign EnvelopesApi object that sends envelopes and reads their status.",
        "details": "Its methods take your DocuSign account id first, for example CreateEnvelope(accountId, envelope), ListStatusChanges(accountId, options) and CreateRecipientView(accountId, envelopeId, viewRequest).",
        "params": [
            ["configuration", "any", "Configuration from model.docusign.Configuration, or a base path string", False, "None"]
        ],
        "returns": "DocuSign EnvelopesApi",
        "watch_out": ["Calls go to DocuSign and can send documents for signature."],
        "example": "client = model.docusign.ApiClient('https://demo.docusign.net/restapi')\nconfig = model.docusign.Configuration(client)\nconfig.AddDefaultHeader('Authorization', 'Bearer YOUR_ACCESS_TOKEN')\napi = model.docusign.EnvelopesApi(config)\n# result = api.CreateEnvelope('YOUR_ACCOUNT_ID', envelope)  # sends for signature\nprint 'EnvelopesApi ready'",
        "signatures": ["model.docusign.EnvelopesApi(configuration=None)", "model.docusign.EnvelopesApi(basePath)"],
        "see_also": ["model.docusign.Configuration", "model.docusign.EnvelopeDefinition", "model.docusign.ListStatusChangesOptions", "model.docusign.RecipientViewRequest"]
    },
    "model.docusign.ListStatusChangesOptions": {
        "category": "DocuSign",
        "kind": "method",
        "tags": ["outside_service"],
        "summary": "Creates the options object for EnvelopesApi.ListStatusChanges.",
        "details": "Used to find envelopes changed since a date or with a given status.",
        "returns": "DocuSign ListStatusChangesOptions",
        "example": "opts = model.docusign.ListStatusChangesOptions()\nopts.fromDate = '2026-10-01'\nopts.status = 'completed'\n# changes = api.ListStatusChanges('YOUR_ACCOUNT_ID', opts)\nprint 'Looking for envelopes completed since ' + opts.fromDate",
        "see_also": ["model.docusign.EnvelopesApi"]
    },
    "model.docusign.Recipients": {
        "category": "DocuSign",
        "kind": "method",
        "tags": ["outside_service"],
        "summary": "Creates an empty DocuSign Recipients container.",
        "details": "Put Signer objects in its Signers list and set it on the envelope's Recipients property.",
        "returns": "DocuSign.eSign.Model.Recipients",
        "watch_out": ["Lists on DocuSign objects are .NET List types, so build them with System.Collections.Generic.List rather than a Python list."],
        "example": "from System.Collections.Generic import List\nds = model.docusign\nsigner = ds.Signer()\nsigner.Name = \"John Doe\"\nrecipients = ds.Recipients()\nrecipients.Signers = List[type(signer)]()\nrecipients.Signers.Add(signer)\nenv = ds.EnvelopeDefinition()\nenv.Recipients = recipients\nprint env.Recipients.Signers.Count, \"signer\"",
        "see_also": ["model.docusign.Signer", "model.docusign.EnvelopeDefinition"]
    },
    "model.docusign.RecipientViewRequest": {
        "category": "DocuSign",
        "kind": "method",
        "tags": ["outside_service"],
        "summary": "Creates a RecipientViewRequest for getting an embedded signing url.",
        "details": "Use it with EnvelopesApi.CreateRecipientView so the person signs inside your page instead of from an email.",
        "returns": "DocuSign RecipientViewRequest",
        "watch_out": ["The recipient fields must match a signer with a ClientUserId on the envelope."],
        "example": "view = model.docusign.RecipientViewRequest()\nview.ReturnUrl = model.CmsHost + '/PyScriptForm/' + (model.ScriptName or 'MyScript')\nview.AuthenticationMethod = 'none'\nview.Email = 'person@example.com'\nview.UserName = 'Jane Smith'\nview.ClientUserId = '{0}'.format(model.UserPeopleId)\nprint view.ReturnUrl",
        "see_also": ["model.docusign.EnvelopesApi", "model.docusign.Signer"]
    },
    "model.docusign.Signer": {
        "category": "DocuSign",
        "kind": "method",
        "tags": ["outside_service"],
        "summary": "Creates an empty DocuSign Signer recipient.",
        "details": "Set Email, Name, RecipientId and RoutingOrder, and attach a Tabs object to show where to sign.",
        "returns": "DocuSign.eSign.Model.Signer",
        "example": "p = model.GetPerson(model.UserPeopleId)\nsigner = model.docusign.Signer()\nsigner.Email = p.EmailAddress\nsigner.Name = p.Name\nsigner.RecipientId = \"1\"\nsigner.RoutingOrder = \"1\"\nprint signer.Name, signer.Email",
        "see_also": ["model.docusign.Recipients", "model.docusign.Tabs", "model.docusign.SignHere"]
    },
    "model.docusign.SignHere": {
        "category": "DocuSign",
        "kind": "method",
        "tags": ["outside_service"],
        "summary": "Creates an empty DocuSign SignHere tab, which marks where a signer signs.",
        "details": "Set DocumentId, PageNumber, XPosition and YPosition (or an anchor string), then add it to Tabs.SignHereTabs.",
        "returns": "DocuSign.eSign.Model.SignHere",
        "watch_out": ["Lists on DocuSign objects are .NET List types, so build them with System.Collections.Generic.List rather than a Python list."],
        "example": "s = model.docusign.SignHere()\ns.DocumentId = \"1\"\ns.PageNumber = \"1\"\ns.XPosition = \"100\"\ns.YPosition = \"150\"\nprint s.PageNumber, s.XPosition, s.YPosition",
        "see_also": ["model.docusign.Tabs", "model.docusign.Text", "model.docusign.Signer"]
    },
    "model.docusign.Tabs": {
        "category": "DocuSign",
        "kind": "method",
        "tags": ["outside_service"],
        "summary": "Creates an empty DocuSign Tabs container.",
        "details": "Put SignHere tabs in SignHereTabs and Text tabs in TextTabs, then set it on a Signer's Tabs property.",
        "returns": "DocuSign.eSign.Model.Tabs",
        "watch_out": ["Lists on DocuSign objects are .NET List types, so build them with System.Collections.Generic.List rather than a Python list."],
        "example": "from System.Collections.Generic import List\nds = model.docusign\ns = ds.SignHere()\ns.DocumentId = \"1\"\ns.PageNumber = \"1\"\ntabs = ds.Tabs()\ntabs.SignHereTabs = List[type(s)]()\ntabs.SignHereTabs.Add(s)\nprint tabs.SignHereTabs.Count, \"sign here tab\"",
        "see_also": ["model.docusign.SignHere", "model.docusign.Text", "model.docusign.Signer"]
    },
    "model.docusign.TemplateRole": {
        "category": "DocuSign",
        "kind": "method",
        "tags": ["outside_service"],
        "summary": "Creates a TemplateRole, which fills a role in a saved DocuSign template with a real person.",
        "details": "Put the roles in envelope.TemplateRoles and set envelope.TemplateId.",
        "returns": "DocuSign TemplateRole",
        "example": "role = model.docusign.TemplateRole()\nrole.RoleName = 'Signer'\nrole.Name = 'Jane Smith'\nrole.Email = 'person@example.com'\nenvelope = model.docusign.EnvelopeDefinition()\nenvelope.TemplateId = 'YOUR_TEMPLATE_ID'\nenvelope.TemplateRoles = [role]\nprint 'Role ' + role.RoleName + ' set'",
        "see_also": ["model.docusign.TemplateTabs", "model.docusign.EnvelopeDefinition"]
    },
    "model.docusign.TemplateTabs": {
        "category": "DocuSign",
        "kind": "method",
        "tags": ["outside_service"],
        "summary": "Creates a TemplateTabs object for prefilling the tabs (fields) of a TemplateRole.",
        "details": "For example, text tabs with a name or date.",
        "returns": "DocuSign TemplateTabs",
        "example": "t = model.docusign.Text()\nt.TabLabel = 'MemberName'\nt.Value = 'Jane Smith'\ntabs = model.docusign.TemplateTabs()\ntabs.TextTabs = [t]\nrole = model.docusign.TemplateRole()\nrole.Tabs = tabs\nprint 'Prefilled ' + t.TabLabel",
        "see_also": ["model.docusign.TemplateRole", "model.docusign.Text"]
    },
    "model.docusign.Text": {
        "category": "DocuSign",
        "kind": "method",
        "tags": ["outside_service"],
        "summary": "Creates an empty DocuSign Text tab, a field the recipient types into.",
        "details": "Set DocumentId, PageNumber, position, TabLabel and Required, then add it to Tabs.TextTabs.",
        "returns": "DocuSign.eSign.Model.Text",
        "example": "t = model.docusign.Text()\nt.DocumentId = \"1\"\nt.PageNumber = \"1\"\nt.XPosition = \"100\"\nt.YPosition = \"250\"\nt.TabLabel = \"Phone Number\"\nt.Required = \"true\"\nprint t.TabLabel",
        "see_also": ["model.docusign.Tabs", "model.docusign.SignHere", "model.docusign.TemplateTabs"]
    },
    "model.Email": {
        "category": "Email & Texting",
        "kind": "method",
        "arity": [6, 8],
        "tags": ["writes_data", "sends_message"],
        "summary": "Queues an email to everyone matched by the query who has an email address and has not turned email off.",
        "details": "The body can use TouchPoint replacement codes such as {first}. Sends right away unless dateWanted is given, in which case the background email job sends it after that time. People who opted out of email from the From address are skipped. Respects model.TestEmail, model.Transactional and model.MaxEmails.",
        "flow": ["Runs the query and keeps people who have a primary email address and have not turned email off. If model.MaxEmails is above 0, only that many are kept. The list is saved in a temporary tag named \".temp email tag\" plus the time.", "Looks up the queuedBy person's email address and raises an error if there is none. In test mode, that address receives the test copies.", "Saves one EmailQueue row (from, subject, body, CC list, SendWhen set to dateWanted, Testing and Transactional flags) and one EmailQueueTo row per person. Anyone who opted out of email from that From address is left off.", "With no dateWanted it sends right away. On the older sender the script waits while each email goes out, with replacement codes filled in per person. On the newer sender the queue is handed to the background sender, and if that handoff fails the regular queue run picks it up later.", "With dateWanted the row waits, and TouchPoint's background email job sends it after that time.", "CC addresses are added to every message's Reply-To, and each CC address also gets one separate copy with no replacement codes filled in.", "When it finishes, EmailQueue.Sent is stamped. If more than one email went out and the send is not transactional, the From address gets a \"sent emails: <subject>\" notice linking to the email's details page. A failed recipient is emailed to the From address as an error notice on the older sender, or saved in the Error column on the newer sender."],
        "params": [
            ["query", "query", "Who to send to: a PeopleId, a saved search name, or Search Builder code", True, None],
            ["queuedBy", "integer", "PeopleId of the person queuing the email. Must have an email address", True, None],
            ["fromAddr", "string", "From email address", True, None],
            ["fromName", "string", "From display name", True, None],
            ["subject", "string", "Subject line", True, None],
            ["body", "string", "HTML body", True, None],
            ["cclist", "string", "Optional comma separated CC addresses", False, "None"],
            ["dateWanted", "date", "Optional date and time to send later (send now when omitted)", False, "None"]
        ],
        "returns": "None",
        "watch_out": ["Sends real email.", "Raises an error if queuedBy has no email address.", "Raises an error if nobody in the query can receive email.", "If nobody in the query can receive the email (no address, email turned off, or opted out of that sender), the call raises an \"Object reference not set\" error instead of doing nothing quietly.", "On the newer email sender, model.TestEmail does not redirect anything: the email goes to the real recipients.", "A scheduled send (dateWanted) does not keep test mode on the older sender, so it goes to the real recipients.", "A dateWanted already more than about a day in the past may never be picked up by the background job.", "On the older sender a large list makes the script wait until every email has gone out, which can make the page slow or time out."],
        "example": "me = model.UserPeopleId\nprint \"Would email PeopleId\", me\n# Sends real email. Uncomment to send to yourself:\n# model.Email(me, me, \"office@mychurch.org\", \"Church Office\", \"Test\", \"<p>Hello {first}</p>\")",
        "signatures": ["model.Email(query, queuedBy, fromAddr, fromName, subject, body, cclist=None, dateWanted=None)"],
        "see_also": ["model.EmailContentWithSubject", "model.TestEmail", "model.Transactional", "model.EmailWithPythonData"]
    },
    "model.Email2": {
        "category": "Email & Texting",
        "kind": "method",
        "arity": [6, 6],
        "tags": ["writes_data", "sends_message", "deprecated"],
        "summary": "Emails everyone in a saved query identified by its GUID.",
        "flow": ["Runs the saved query with that GUID to get the people.", "Keeps only people who have an email address and have not turned email off, and caps the list at model.MaxEmails when that is set.", "Saves the recipients in a special emailer tag and creates an EmailQueue row, with the queuedBy person as the sender of record.", "Sends to the queue right away. Each recipient gets a personalized copy with replacement codes such as {first} filled in."],
        "params": [
            ["qid", "string", "Query GUID", True, None],
            ["queuedBy", "integer", "PeopleId recorded as the sender", True, None],
            ["fromAddr", "string", "From email address", True, None],
            ["fromName", "string", "From display name", True, None],
            ["subject", "string", "Subject", True, None],
            ["body", "string", "HTML body", True, None]
        ],
        "returns": "None",
        "watch_out": ["Deprecated. Use model.Email, which accepts a PeopleId, saved search name or Search Builder code.", "Queues real email.", "Raises \"QueuedBy PeopleId not found in model.Email\" when the queuedBy person has no email address.", "Skips people with no email address and people who turned off both of their email addresses.", "Honors model.MaxEmails (caps the number of recipients), model.TestEmail and model.Transactional set earlier in the script.", "An unknown GUID raises an error instead of sending nothing."],
        "example": "# Deprecated. Use model.Email instead:\n# model.Email(str(model.UserPeopleId), model.UserPeopleId, \"info@church.org\", \"Church Office\", \"Subject\", \"<p>Body</p>\")\nprint \"See model.Email\"",
        "see_also": ["model.Email", "model.EmailContentWithSubject"]
    },
    "model.EmailContent": {
        "category": "Email & Texting",
        "kind": "method",
        "arity": [5, 6],
        "tags": ["writes_data", "sends_message", "deprecated"],
        "summary": "Emails everyone in a query using an HTML Special Content item as the body.",
        "details": "The 5 argument form uses the content's title as the subject. The 6 argument form takes a subject before the content name and works the same as model.EmailContentWithSubject. The content can be HTML Special Content or an email template.",
        "flow": ["Looks up the content by name as HTML Special Content, then as an email template. If neither exists it quietly does nothing.", "Uses the content's title as the subject in the 5 argument form, or the subject you pass in the 6 argument form.", "Runs the query, keeps people with an email address who have not turned email off, and caps the list at model.MaxEmails when set.", "Creates an EmailQueue row from the content's HTML and sends it right away."],
        "params": [
            ["query", "query", "People to act on: a PeopleId, a saved search name, or Search Builder code", True, None],
            ["queuedBy", "integer", "PeopleId recorded as the sender of the email", True, None],
            ["fromAddr", "string", "From email address", True, None],
            ["fromName", "string", "From display name", True, None],
            ["contentName", "string", "Name of the HTML Special Content (or email template) to send", True, None]
        ],
        "returns": "None",
        "watch_out": ["Deprecated. Use model.EmailContentWithSubject instead.", "Queues real email.", "Raises \"QueuedBy PeopleId not found in model.Email\" when the queuedBy person has no email address.", "Skips people with no email address and people who turned off both of their email addresses.", "Honors model.MaxEmails (caps the number of recipients), model.TestEmail and model.Transactional set earlier in the script.", "A content name that is not found sends nothing and raises no error."],
        "example": "# Deprecated. Use model.EmailContentWithSubject instead.\n# This would queue real email to everyone in the query:\n# model.EmailContent(str(model.UserPeopleId), model.UserPeopleId, \"info@church.org\", \"Church Office\", \"WelcomeEmail\")\nprint \"See model.EmailContentWithSubject\"",
        "signatures": ["model.EmailContent(query, queuedBy, fromAddr, fromName, contentName)", "model.EmailContent(query, queuedBy, fromAddr, fromName, subject, contentName)"],
        "see_also": ["model.EmailContentWithSubject", "model.Email", "model.EmailContent2"]
    },
    "model.EmailContent2": {
        "category": "Email & Texting",
        "kind": "method",
        "arity": [5, 6],
        "tags": ["writes_data", "sends_message", "deprecated", "known_issue"],
        "summary": "Emails everyone in a saved query (by GUID) using an HTML Special Content item as the body.",
        "flow": ["Looks up the content by name as HTML Special Content only. Email templates are not found. If it is missing it quietly does nothing.", "Runs the saved query with that GUID and keeps people with an email address who have not turned email off.", "Creates an EmailQueue row from the content body and sends it right away."],
        "params": [
            ["qid", "string", "Query GUID", True, None],
            ["queuedBy", "integer", "PeopleId recorded as the sender", True, None],
            ["fromAddr", "string", "From email address", True, None],
            ["fromName", "string", "From display name", True, None],
            ["contentName", "string", "HTML Special Content name; its title is the subject", True, None]
        ],
        "returns": "None",
        "watch_out": ["Deprecated. Use model.EmailContentWithSubject instead.", "The 6 argument form uses fromAddr as the display name, ignoring fromName.", "Queues real email.", "Raises \"QueuedBy PeopleId not found in model.Email\" when the queuedBy person has no email address.", "Skips people with no email address and people who turned off both of their email addresses.", "Honors model.MaxEmails (caps the number of recipients), model.TestEmail and model.Transactional set earlier in the script.", "A content name that is not found sends nothing and raises no error.", "An unknown GUID raises an error."],
        "example": "# Deprecated. Use model.EmailContentWithSubject instead.\nprint \"See model.EmailContentWithSubject\"",
        "signatures": ["model.EmailContent2(qid, queuedBy, fromAddr, fromName, contentName)", "model.EmailContent2(qid, queuedBy, fromAddr, fromName, subject, contentName)"],
        "see_also": ["model.EmailContentWithSubject", "model.EmailContent", "model.Email"]
    },
    "model.EmailContentWithPythonData": {
        "category": "Email & Texting",
        "kind": "method",
        "arity": [6, 6],
        "tags": ["writes_data", "sends_message"],
        "summary": "Emails recipients using a saved HTML content item or template filled with per person data rows.",
        "details": "Same as model.EmailWithPythonData, but the subject and body come from a saved HTML Special Content item or email template. The content's title is used as the subject.",
        "flow": ["Looks up contentName as an HTML Special Content item (or Text item of the same name), then as a saved email template, and raises an error if neither exists.", "Uses the content's title as the subject and its HTML as the body.", "Then follows model.EmailWithPythonData exactly: queue rows per person, opt-outs by From address, and the per person Handlebars merge from recipientData."],
        "params": [
            ["savedQuery", "query", "Saved search name, search code or query guid that selects the recipients", True, None],
            ["queuedBy", "integer", "PeopleId of the sender who queues the email", True, None],
            ["fromAddr", "string", "From email address", True, None],
            ["fromName", "string", "From display name", True, None],
            ["contentName", "string", "Name of the HTML content or email template with {{Column}} fields", True, None],
            ["recipientData", "list", "Rows with a PeopleId column, such as the result of q.QuerySql", True, None]
        ],
        "returns": "None",
        "watch_out": ["Sends email immediately.", "Raises an error if no content with that name exists.", "On the newer email sender recipientData is ignored and everyone in the query gets the unmerged template.", "Two recipientData rows with the same PeopleId raise an error."],
        "example": "rows = q.QuerySql('SELECT PeopleId, PreferredName AS First FROM People WHERE PeopleId = @UserPeopleId')\nprint 'Would send to %d person' % len(list(rows))\n# Sends a real email using the HTML content named MyGroupReminder:\n# model.EmailContentWithPythonData('peopleids=' + str(model.UserPeopleId), model.UserPeopleId, 'office@church.org', 'Church Office', 'MyGroupReminder', rows)",
        "see_also": ["model.EmailWithPythonData", "model.EmailContentWithSubject", "q.QuerySql"]
    },
    "model.EmailContentWithSubject": {
        "category": "Email & Texting",
        "kind": "method",
        "arity": [6, 8],
        "tags": ["writes_data", "sends_message"],
        "summary": "Queues an email whose body comes from an HTML Special Content file or saved template, with your own subject.",
        "details": "Goes to everyone matched by the query who has an email address. The subject you pass is used instead of the content title. Sends right away unless dateWanted is given.",
        "flow": ["Looks up contentName as an HTML Special Content item (a Text item of the same name also works), then as a saved email template. A designer template uses its finished HTML. If nothing is found it stops without error.", "Uses that HTML as the body with your subject, then follows the same steps as model.Email: keeps people with a usable address, honors model.MaxEmails, and raises an error if queuedBy has no email address.", "Saves an EmailQueue row and one EmailQueueTo row per person, leaving off anyone who opted out of email from that From address.", "Sends now, or after dateWanted through the background email job. Replacement codes in the content are filled in per person.", "When done, the From address gets the \"sent emails\" summary unless the send is transactional or went to only one person."],
        "params": [
            ["query", "query", "Who to send to: a PeopleId, a saved search name, or Search Builder code", True, None],
            ["queuedBy", "integer", "PeopleId of the person queuing the email. Must have an email address", True, None],
            ["fromAddr", "string", "From email address", True, None],
            ["fromName", "string", "From display name", True, None],
            ["subject", "string", "Subject line (used instead of the content title)", True, None],
            ["contentName", "string", "Name of the HTML content or email template", True, None],
            ["cclist", "string", "Optional comma separated CC addresses", False, "None"],
            ["dateWanted", "date", "Optional date and time to send later (send now when omitted)", False, "None"]
        ],
        "returns": "None",
        "watch_out": ["Sends real email.", "Does nothing if the content is not found.", "Raises an error if nobody in the query can receive email.", "On the newer email sender, model.TestEmail does not redirect anything."],
        "example": "me = model.UserPeopleId\nprint \"Would email PeopleId\", me\n# Sends real email. Uncomment to send to yourself:\n# model.EmailContentWithSubject(me, me, \"office@mychurch.org\", \"Church Office\", \"Weekly News\", \"WeeklyNewsTemplate\")",
        "signatures": ["model.EmailContentWithSubject(query, queuedBy, fromAddr, fromName, subject, contentName, cclist=None, dateWanted=None)"],
        "see_also": ["model.Email", "model.EmailContentWithPythonData", "model.EmailContent", "model.HtmlContent"]
    },
    "model.EmailReminders": {
        "category": "Email & Texting",
        "kind": "method",
        "arity": [1, 1],
        "tags": ["writes_data", "sends_message"],
        "summary": "Sends the reminder emails for one involvement.",
        "details": "If the involvement uses volunteer time slots it sends volunteer reminders. Otherwise it sends event reminders.",
        "flow": ["Loads the involvement and its registration settings. The subject and body come from the involvement's Reminder subject and body on the Registration Messages tab.", "Volunteer time slot involvements: finds members (not pending, not inactive) with a committed slot in the next 7 days. Each gets an email listing all their upcoming commitments in {details}.", "Volunteer: each notify person on the involvement then gets a \"Reminder Notices sent for <involvement>\" summary of who is serving this week. The From address is the first notify person's email, or the AdminMail setting if there are none.", "Other involvements: every current member (not pending, not inactive, prospects included) gets the reminder, with {details} filled in from their registration. It is sent from the first notify person's email.", "Each email is its own transactional EmailQueue row, so the sender gets no \"sent emails\" summary. Test mode is applied."],
        "params": [
            ["orgId", "integer", "OrganizationId of the involvement", True, None]
        ],
        "returns": "None",
        "watch_out": ["This sends real email to the involvement's members, and for volunteer involvements also to its notify people.", "For event reminders, a missing Reminder subject or body raises \"no subject or body\", and an involvement with no notify person raises \"no notify person\".", "For volunteer reminders, a missing subject or body is not an error: the email goes out with the words \"no subject\" or \"no body\".", "Event reminders go to every current member with no date window, so running it more than once emails everyone again."],
        "example": "orgId = 30\nprint \"Would send reminders for involvement\", orgId\n# Sends real email. Uncomment to run:\n# model.EmailReminders(orgId)",
        "see_also": ["model.SendAttendanceReminders", "model.Email"]
    },
    "model.EmailReport": {
        "category": "Email & Texting",
        "kind": "method",
        "arity": [6, 8],
        "tags": ["writes_data", "sends_message"],
        "summary": "Runs a Python script from Special Content and emails its printed output to everyone the query matches.",
        "details": "Only people with an email address receive it. The 8 argument form sets model.Data.QueryName and model.Data.QueryDescription first, so one generic report script can be reused for different groups.",
        "flow": ["Runs the query and keeps people who have a primary email address and have not turned email off, saving them in a temporary tag. model.MaxEmails is not applied.", "Loads the report as a Python Special Content script (a Text item of the same name also works) and runs it once, sharing this script's model and model.Data. Everything it prints becomes the email body.", "Saves an EmailQueue row and one EmailQueueTo row per person, leaving off anyone who opted out of email from that From address. Every person gets the same report; replacement codes in the output are still filled in per person.", "Sends right away (older sender: during the script; newer sender: handed to the background sender).", "When done, the From address gets the \"sent emails\" summary unless model.Transactional is True or only one person was emailed."],
        "params": [
            ["query", "query", "Who to send to: a PeopleId, a saved search name, or Search Builder code", True, None],
            ["queuedBy", "integer", "PeopleId of the person queuing the email", True, None],
            ["fromAddr", "string", "From email address", True, None],
            ["fromName", "string", "From display name", True, None],
            ["subject", "string", "Subject line", True, None],
            ["report", "string", "Name of the Python script whose output becomes the body", True, None],
            ["queryName", "string", "Stored in model.Data.QueryName for the report script", False, None],
            ["queryDescription", "string", "Stored in model.Data.QueryDescription for the report script", False, None]
        ],
        "returns": "None",
        "watch_out": ["This sends real email, right away.", "If the script is not found, the email is still queued with an empty body and nobody receives it. On the older sender the From address gets a \"no subject or body\" error notice.", "If the report script raises an error, the error message itself becomes the body and is emailed to everyone.", "model.MaxEmails is ignored, so a test limit does not protect you here.", "queuedBy is not checked for an email address, unlike model.Email.", "Raises an \"Object reference not set\" error if nobody in the query can receive email."],
        "example": "me = model.UserPeopleId\nprint \"Would email report to PeopleId\", me\n# Sends real email. Uncomment to run:\n# model.EmailReport(me, me, \"office@mychurch.org\", \"Church Office\", \"Weekly Stats\", \"StatsReport\")\n# model.EmailReport(str(me), me, \"office@mychurch.org\", \"Church Office\", \"Stats\", \"GenericReport\", \"Age>=65\", \"Senior Adults\")",
        "signatures": ["model.EmailReport(query, queuedBy, fromAddr, fromName, subject, report)", "model.EmailReport(query, queuedBy, fromAddr, fromName, subject, report, queryName, queryDescription)"],
        "see_also": ["model.CallScript", "model.Email", "model.EmailContentWithSubject", "model.TestEmail"]
    },
    "model.EmailStr": {
        "category": "Email & Texting",
        "kind": "method",
        "arity": [1, 3],
        "tags": ["read_only", "deprecated"],
        "summary": "Runs TouchPoint email replacement codes such as {first} or {name} against one person and returns the text.",
        "details": "Defaults to the current user when no PeopleId is given. Nothing is sent.",
        "params": [
            ["body", "string", "Text containing email replacement codes", True, None],
            ["pid", "integer", "Person to fill the codes from (the current user when omitted)", False, "None"]
        ],
        "returns": "String with the replacements applied, or \"no user\" when there is no person",
        "watch_out": ["Deprecated. For Handlebars templates use model.RenderTemplate instead.", "Returns \"no user\" when there is no person.", "Raises an error when the person has no email address or the PeopleId does not exist.", "Codes that depend on an involvement use model.CurrentOrgId.", "In the 3 argument form, queuedBy is the person whose data fills the codes and fromAddr is used for sender codes."],
        "example": "# Deprecated. Prefer model.RenderTemplate for new code.\ntext = model.EmailStr(\"Hello {first}, your email is {emailaddress}\", model.UserPeopleId)\nprint text",
        "signatures": ["model.EmailStr(body, pid=None)", "model.EmailStr(queuedBy, fromAddr, body)"],
        "see_also": ["model.RenderTemplate", "model.Email"]
    },
    "model.EmailWithPythonData": {
        "category": "Email & Texting",
        "kind": "method",
        "arity": [7, 7],
        "tags": ["writes_data", "sends_message"],
        "summary": "Sends an email to the people in a query and merges per-person data from your script into the body.",
        "details": "recipientData is a list of rows, usually from q.QuerySql, and each row must have a PeopleId column. Each recipient's row is applied to the body as a Handlebars template, so {{ColumnName}} inserts that person's value. Normal {replacement} codes also work. If a rendered body contains <!--SUBJECT:text--> that text overrides the subject for that person. It respects model.MaxEmails and model.TestEmail. This per person merge only happens on the older email sender.",
        "flow": ["Runs the query, keeps people with a usable email address, honors model.MaxEmails, and saves them in a temporary tag. Raises an error if queuedBy has no email address.", "Saves an EmailQueue row and one EmailQueueTo row per person, leaving off anyone who opted out of email from that From address.", "Older sender: for each person in the queue it fills in the replacement codes, then applies that person's recipientData row as a Handlebars template. People with no row, or whose result contains <!--SKIP-->, are skipped. A <!--SUBJECT:text--> comment sets that person's subject.", "The CC handling of model.Email does not apply here; this call sends no CC.", "When done, the From address gets a \"sent emails\" summary if more than one person was in the queue."],
        "params": [
            ["savedQuery", "query", "Saved search name, search code or query guid that selects the recipients", True, None],
            ["queuedBy", "integer", "PeopleId of the sender who queues the email (must have an email address)", True, None],
            ["fromAddr", "string", "From email address", True, None],
            ["fromName", "string", "From display name", True, None],
            ["subject", "string", "Default subject line", True, None],
            ["body", "string", "HTML body with {{Column}} Handlebars fields", True, None],
            ["recipientData", "list", "Rows with a PeopleId column, such as the result of q.QuerySql", True, None]
        ],
        "returns": "None",
        "watch_out": ["This sends real email immediately.", "People in the query with no row in recipientData are skipped.", "If a rendered body contains <!--SKIP--> that person is skipped.", "The queuedBy person must have an email address.", "On the newer email sender (JobsServiceRunEmailQueueProcessing on) your recipientData is not used at all. Everyone in the query gets the raw template with the {{Column}} fields unfilled, and <!--SKIP--> and <!--SUBJECT:--> are ignored. Check this setting before using this call.", "Two recipientData rows with the same PeopleId raise an error, and nothing is sent.", "{{Column}} values are HTML encoded. Use {{{Column}}} to insert HTML.", "On the older sender model.Transactional does not stop the \"sent emails\" summary for this call.", "Raises an \"Object reference not set\" error if nobody in the query can receive email."],
        "example": "sql = '''SELECT p.PeopleId, p.PreferredName AS First,\n  (SELECT COUNT(*) FROM OrganizationMembers om WHERE om.PeopleId = p.PeopleId) AS Groups\nFROM People p WHERE p.PeopleId = @UserPeopleId'''\nrows = q.QuerySql(sql)\nbody = '<p>Hi {{First}}, you are in {{Groups}} groups.</p>'\nprint 'Rows: %d' % len(list(rows))\n# Sends a real email to the current user only:\n# model.EmailWithPythonData('peopleids=' + str(model.UserPeopleId), model.UserPeopleId, 'office@church.org', 'Church Office', 'Your groups', body, rows)",
        "see_also": ["model.EmailContentWithPythonData", "model.Email", "q.QuerySql", "model.MaxEmails"]
    },
    "model.MaxEmails": {
        "category": "Email & Texting",
        "kind": "property",
        "summary": "Limits how many people a single email call queues when set above 0.",
        "details": "It applies to model.Email, model.EmailContentWithSubject and model.EmailWithPythonData (and the calls built on them). model.EmailReport ignores it. It takes N people from the query who have an email address and email turned on, in no set order. The default 0 means no limit. Useful for testing a large send on a few people.",
        "returns": "integer (settable, default 0)",
        "watch_out": ["Opt-outs by sender are checked after the limit is applied, so fewer than N may be sent."],
        "example": "model.MaxEmails = 5\nprint 'Each email call in this script now goes to at most %d people' % model.MaxEmails\n# model.Email('IsMember = 1[True]', model.UserPeopleId, 'office@church.org', 'Church Office', 'Test', 'Hello')",
        "see_also": ["model.TestEmail", "model.Email", "model.EmailWithPythonData", "model.Transactional"]
    },
    "model.SendSms": {
        "category": "Email & Texting",
        "kind": "method",
        "arity": [4, 4],
        "tags": ["writes_data", "sends_message"],
        "summary": "Queues a text message to everyone matched by the query, using one of your SMS sending groups.",
        "details": "The title is stored for your records but is not part of the text. The query can also be a person object.",
        "flow": ["Checks the title (150 characters max) and message (1600 characters max) and raises an error if either is too long.", "Saves an SMSList row with the sending group, title, message and you as the sender (or the admin when there is no signed-in user).", "Adds one SMSItems row per person in the query using their cell phone. A person with no cell phone is marked NoNumber. A person with texting turned off, or opted out of that sending group, is marked NoOptIn. Neither is texted.", "Sends right away. From a web page a separate background process does the sending; from the morning batch it sends during the script. Replacement codes such as {first} are filled in per person.", "Sends through the group's provider (PoweredByText or Twilio). Each failure is saved on that person's SMSItems row (ResultStatus error and ErrorMessage). If the group has no phone numbers set up, every item is marked with that error."],
        "params": [
            ["query", "query", "Who to text: a PeopleId, a saved search name, Search Builder code, or a person object", True, None],
            ["iSendGroup", "integer", "Id of the SMS sending group (SMSGroups table)", True, None],
            ["sTitle", "string", "Title for your records, 150 characters max. Not sent", True, None],
            ["sMessage", "string", "Text message body, 1600 characters max", True, None]
        ],
        "returns": "None",
        "watch_out": ["This sends real text messages.", "Raises an error if the title is over 150 characters or the message is over 1600 characters.", "If Twilio is the provider and its credentials are missing, nothing is sent and no error is raised.", "When Twilio reports that someone replied STOP, TouchPoint turns off texting for that person and opts them out of all their sending groups.", "model.TestEmail has no effect on texts.", "With PoweredByText, numbers of 9 digits or fewer are marked as errors and skipped."],
        "example": "for g in q.QuerySql(\"SELECT Id, Name FROM SMSGroups\"):\n    print g.Id, g.Name\n# Sends a real text. Uncomment to text yourself:\n# model.SendSms(model.UserPeopleId, 1, \"Test\", \"Bible study tonight at 7pm\")",
        "see_also": ["model.CreateTinyUrl", "model.Email"]
    },
    "model.TestEmail": {
        "category": "Email & Texting",
        "kind": "property",
        "summary": "When True, later emails go to the queuing person's own address instead of the real recipients.",
        "details": "It can be read and written. On the older email sender it affects model.Email, model.EmailContentWithSubject, model.EmailReport, model.EmailWithPythonData, model.EmailContentWithPythonData and model.EmailReminders: each message goes to the queuing person's address and each name gets \"(test)\" added. Default is False.",
        "returns": "boolean",
        "watch_out": ["Set it before you send, because it only affects emails sent later in the script.", "Only the older email sender honors it. On the newer sender (JobsServiceRunEmailQueueProcessing on) the flag is saved on the EmailQueue row and only stops the \"sent emails\" summary; real recipients still get the email.", "It is not kept for scheduled sends (dateWanted) on the older sender, so those go to the real recipients.", "It does not affect model.SendSms."],
        "example": "model.TestEmail = True\nprint \"Test mode on:\", model.TestEmail\n# Any model.Email(...) call now goes only to the queuedBy person\nmodel.TestEmail = False",
        "see_also": ["model.MaxEmails", "model.Transactional", "model.Email", "model.EmailReport"]
    },
    "model.Transactional": {
        "category": "Email & Texting",
        "kind": "property",
        "summary": "When True, emails queued later in this script are marked transactional.",
        "details": "It can be read and written. The sender then gets no \"sent emails\" summary notice. Default is False.",
        "returns": "boolean",
        "watch_out": ["Transactional emails are left out of the Search Builder count of emails a person has sent (RecentEmailSentCount). The count of emails a person received still includes them.", "On the older email sender model.EmailWithPythonData and model.EmailContentWithPythonData still send the \"sent emails\" summary even when this is True."],
        "example": "model.Transactional = True\nprint \"Transactional:\", model.Transactional\n# Emails queued after this line are transactional",
        "see_also": ["model.TestEmail", "model.MaxEmails", "model.Email"]
    },
    "model.AddExtraValueAttributes": {
        "category": "Extra Values",
        "kind": "method",
        "arity": [3, 3],
        "tags": ["writes_data"],
        "summary": "Adds or updates an Attributes extra value on people.",
        "details": "The text, normally a JSON string, is stored in Data and the field is flagged as attributes so ExtraValueAttributes can read it. The first argument can be a PeopleId or a Search Builder code, and every matching person is updated.",
        "flow": ["Runs the query to get the matching PeopleIds.", "For each person, finds the PeopleExtra row with that field name or creates one.", "Stores the text in the Data column, marks the row as attributes, stamps the time, and saves that person."],
        "params": [
            ["query", "query", "PeopleId, or a Search Builder code or saved search name. Every matching person is updated.", True, None],
            ["name", "string", "Extra value field name", True, None],
            ["text", "string", "JSON text to store", True, None]
        ],
        "returns": "None",
        "watch_out": ["Updates every person the query matches.", "An empty string is silently ignored.", "A query that matches no one does nothing and raises no error.", "Each person is saved one at a time, so a large query can be slow."],
        "example": "import json\npid = model.UserPeopleId\nmodel.AddExtraValueAttributes(pid, \"ApiExplorerTest\", json.dumps({\"color\": \"blue\"}))\nprint model.ExtraValueAttributes(pid, \"ApiExplorerTest\")\nmodel.DeleteExtraValue(pid, \"ApiExplorerTest\")  # clean up",
        "see_also": ["model.ExtraValueAttributes", "model.DeleteExtraValue", "model.AddExtraValueText", "model.AddExtraValueCode"]
    },
    "model.AddExtraValueBool": {
        "category": "Extra Values",
        "kind": "method",
        "arity": [3, 3],
        "tags": ["writes_data"],
        "summary": "Adds or updates a Bit (True/False) extra value on people.",
        "details": "The first argument can be a PeopleId or a Search Builder code, and every matching person is updated.",
        "flow": ["Runs the query to get the matching PeopleIds.", "For each person, finds the PeopleExtra row with that field name or creates one.", "Sets the True/False value, stamps the time, and saves that person."],
        "params": [
            ["query", "query", "PeopleId, or a Search Builder code or saved search name. Every matching person is updated.", True, None],
            ["name", "string", "Extra value field name", True, None],
            ["b", "boolean", "True or False", True, None]
        ],
        "returns": "None",
        "watch_out": ["Updates every person the query matches.", "A query that matches no one does nothing and raises no error.", "Each person is saved one at a time, so a large query can be slow.", "Other values already stored under the same field name, such as a code or text, are left in place."],
        "example": "pid = model.UserPeopleId\nmodel.AddExtraValueBool(pid, \"ApiExplorerTest\", True)\nprint model.ExtraValueBit(pid, \"ApiExplorerTest\")\nmodel.DeleteExtraValue(pid, \"ApiExplorerTest\")  # clean up",
        "see_also": ["model.ExtraValueBit", "model.DeleteExtraValue", "model.AddExtraValueText", "model.AddExtraValueCode"]
    },
    "model.AddExtraValueBoolFamily": {
        "category": "Extra Values",
        "kind": "method",
        "arity": [3, 3],
        "tags": ["writes_data"],
        "summary": "Adds or updates a Bit (True/False) extra value on families.",
        "details": "Family extra values are keyed by person, not FamilyId. Pass the PeopleId of any member and the value is stored on that person's family.",
        "flow": ["Runs the query and keeps one person per family.", "Looks up that person's family and finds the FamilyExtra row with that field name, or creates one.", "Sets the True/False value and saves, one family at a time."],
        "params": [
            ["query", "query", "PeopleId of any family member, or a Search Builder code or saved search name. Each matching person's family is updated once.", True, None],
            ["name", "string", "Extra value field name", True, None],
            ["b", "boolean", "True or False", True, None]
        ],
        "returns": "None",
        "watch_out": ["Pass a PeopleId, not a FamilyId.", "A query that matches no one does nothing and raises no error."],
        "example": "pid = model.UserPeopleId  # your own family\nmodel.AddExtraValueBoolFamily(pid, \"ApiExplorerTest\", True)\nprint model.ExtraValueBitFamily(pid, \"ApiExplorerTest\")\nmodel.DeleteExtraValueFamily(pid, \"ApiExplorerTest\")  # clean up",
        "see_also": ["model.ExtraValueBitFamily", "model.DeleteExtraValueFamily", "model.AddExtraValueTextFamily", "model.AddExtraValueCodeFamily"]
    },
    "model.AddExtraValueBoolOrg": {
        "category": "Extra Values",
        "kind": "method",
        "arity": [3, 3],
        "tags": ["writes_data"],
        "summary": "Adds or updates a Bit (True/False) extra value on an involvement.",
        "flow": ["Loads the involvement by OrganizationId.", "Finds the OrganizationExtra row with that field name, or creates one.", "Sets the True/False value, stamps the time, and saves."],
        "params": [
            ["oid", "integer", "OrganizationId of the involvement", True, None],
            ["name", "string", "Extra value field name", True, None],
            ["b", "boolean", "True or False", True, None]
        ],
        "returns": "None",
        "watch_out": ["Raises an error if the involvement does not exist.", "There is no Python method to delete an involvement extra value.", "A blank field name is silently ignored."],
        "example": "row = q.QuerySqlTop1(\"SELECT TOP 1 OrganizationId, Field, BitValue FROM OrganizationExtra WHERE Type = 'Bit'\")\nif row:\n    print row.OrganizationId, row.Field, row.BitValue\n# Writes to the involvement:\n# model.AddExtraValueBoolOrg(row.OrganizationId, row.Field, True)",
        "see_also": ["model.ExtraValueBitOrg", "model.AddExtraValueTextOrg", "model.AddExtraValueCodeOrg"]
    },
    "model.AddExtraValueCode": {
        "category": "Extra Values",
        "kind": "method",
        "arity": [3, 3],
        "tags": ["writes_data"],
        "summary": "Adds or updates a Code extra value on people.",
        "details": "A Code is a short value from a list, stored in StrValue, 200 characters max. The first argument can be a PeopleId or a Search Builder code, and every matching person is updated.",
        "flow": ["Runs the query (a PeopleId, a saved search name or Search Builder code) to get the list of people.", "For each person, finds the PeopleExtra row with this field name, or creates one if it is missing.", "Sets the code value and the time stamp, then saves that person before moving to the next."],
        "params": [
            ["query", "query", "PeopleId, or a Search Builder code or saved search name. Every matching person is updated.", True, None],
            ["name", "string", "Extra value field name", True, None],
            ["text", "string", "The code value", True, None]
        ],
        "returns": "None",
        "watch_out": ["Updates every person the query matches.", "An empty string is silently ignored.", "None raises an error.", "A code longer than 200 characters raises an error when it saves.", "Field names longer than 150 characters are cut to 150 without warning.", "If the query matches no one, nothing happens and no error is raised.", "Each person is saved separately, so a query that matches thousands of people is slow.", "If the same field name already holds a different kind of value, both are kept on one row and the string readers (ExtraValue, ExtraValueCode) then return text such as \"unknown type: CodeText\"."],
        "example": "pid = model.UserPeopleId\nmodel.AddExtraValueCode(pid, \"ApiExplorerTest\", \"Blue\")\nprint model.ExtraValueCode(pid, \"ApiExplorerTest\")\nmodel.DeleteExtraValue(pid, \"ApiExplorerTest\")  # clean up",
        "see_also": ["model.ExtraValueCode", "model.DeleteExtraValue", "model.AddExtraValueText"]
    },
    "model.AddExtraValueCodeFamily": {
        "category": "Extra Values",
        "kind": "method",
        "arity": [3, 3],
        "tags": ["writes_data"],
        "summary": "Adds or updates a Code extra value on families.",
        "details": "Family extra values are keyed by person, not FamilyId. Pass the PeopleId of any member and the value is stored on that person's family.",
        "flow": ["Runs the query, then keeps one person per family so each family is touched once.", "Looks up that person's FamilyId and finds or creates the FamilyExtra row with this field name.", "Sets the code value and the time stamp, then saves that family."],
        "params": [
            ["query", "query", "PeopleId of any family member, or a Search Builder code or saved search name. Each matching person's family is updated once.", True, None],
            ["name", "string", "Extra value field name", True, None],
            ["text", "string", "The code value", True, None]
        ],
        "returns": "None",
        "watch_out": ["Pass a PeopleId, not a FamilyId.", "An empty string is silently ignored.", "Passing None raises an error.", "A code longer than 200 characters raises an error when it saves.", "Family extra value field names are limited to 50 characters; a longer name raises an error when it saves.", "If the query matches no one, nothing happens and no error is raised.", "If the same field name already holds a different kind of value, both are kept on one row and the string readers (ExtraValue, ExtraValueCode) then return text such as \"unknown type: CodeText\"."],
        "example": "pid = model.UserPeopleId  # your own family\nmodel.AddExtraValueCodeFamily(pid, \"ApiExplorerTest\", \"Gold\")\nprint model.ExtraValueCodeFamily(pid, \"ApiExplorerTest\")\nmodel.DeleteExtraValueFamily(pid, \"ApiExplorerTest\")  # clean up",
        "see_also": ["model.ExtraValueCodeFamily", "model.DeleteExtraValueFamily", "model.AddExtraValueTextFamily"]
    },
    "model.AddExtraValueCodeOrg": {
        "category": "Extra Values",
        "kind": "method",
        "arity": [3, 3],
        "tags": ["writes_data"],
        "summary": "Adds or updates a Code extra value on an involvement.",
        "flow": ["Loads the involvement by OrganizationId.", "Finds the OrganizationExtra row with this field name, or creates one.", "Sets the code value and the time stamp, then saves."],
        "params": [
            ["oid", "integer", "OrganizationId of the involvement", True, None],
            ["name", "string", "Extra value field name", True, None],
            ["text", "string", "The code value", True, None]
        ],
        "returns": "None",
        "watch_out": ["An empty string is silently ignored.", "Raises an error if the involvement does not exist.", "There is no Python method to delete an involvement extra value.", "Passing None raises an error.", "A code longer than 200 characters raises an error when it saves.", "Involvement extra value field names are limited to 50 characters; a longer name raises an error when it saves.", "An empty field name is also silently ignored.", "If the same field name already holds a different kind of value, both are kept on one row and the string readers (ExtraValue, ExtraValueCode) then return text such as \"unknown type: CodeText\"."],
        "example": "row = q.QuerySqlTop1(\"SELECT TOP 1 OrganizationId, Field, StrValue FROM OrganizationExtra WHERE Type = 'Code'\")\nif row:\n    print row.OrganizationId, row.Field, row.StrValue\n# Writes to the involvement:\n# model.AddExtraValueCodeOrg(row.OrganizationId, row.Field, \"Active\")",
        "see_also": ["model.ExtraValueCodeOrg", "model.AddExtraValueTextOrg"]
    },
    "model.AddExtraValueDate": {
        "category": "Extra Values",
        "kind": "method",
        "arity": [3, 3],
        "tags": ["writes_data"],
        "summary": "Adds or updates a Date extra value on people.",
        "details": "Accepts a DateTime or a date string. The first argument can be a PeopleId or a Search Builder code, and every matching person is updated.",
        "flow": ["Runs the query (a PeopleId, a saved search name or Search Builder code) to get the list of people.", "For each person, finds the PeopleExtra row with this field name, or creates one if it is missing.", "Sets the date and the time stamp, then saves that person before moving to the next."],
        "params": [
            ["query", "query", "PeopleId, or a Search Builder code or saved search name. Every matching person is updated.", True, None],
            ["name", "string", "Extra value field name", True, None],
            ["dt", "any", "A date: a .NET DateTime, a Python datetime, or a string such as \"12/25/2026\"", True, None]
        ],
        "returns": "None",
        "watch_out": ["Updates every person the query matches.", "If the date cannot be read nothing is written and no error is raised.", "Field names longer than 150 characters are cut to 150 without warning.", "If the query matches no one, nothing happens and no error is raised.", "Each person is saved separately, so a query that matches thousands of people is slow.", "If the same field name already holds a different kind of value, both are kept on one row and the string readers (ExtraValue, ExtraValueCode) then return text such as \"unknown type: CodeText\"."],
        "example": "pid = model.UserPeopleId\nmodel.AddExtraValueDate(pid, \"ApiExplorerTest\", \"1/15/2020\")\nprint model.ExtraValueDate(pid, \"ApiExplorerTest\").ToShortDateString()\nmodel.DeleteExtraValue(pid, \"ApiExplorerTest\")  # clean up",
        "see_also": ["model.ExtraValueDate", "model.DeleteExtraValue", "model.AddExtraValueText", "model.AddExtraValueCode"]
    },
    "model.AddExtraValueDateFamily": {
        "category": "Extra Values",
        "kind": "method",
        "arity": [3, 3],
        "tags": ["writes_data"],
        "summary": "Adds or updates a Date extra value on families.",
        "details": "Accepts a DateTime or a date string. Family extra values are keyed by person, not FamilyId. Pass the PeopleId of any member and the value is stored on that person's family.",
        "flow": ["Runs the query, then keeps one person per family so each family is touched once.", "Looks up that person's FamilyId and finds or creates the FamilyExtra row with this field name.", "Sets the date and the time stamp, then saves that family."],
        "params": [
            ["query", "query", "PeopleId of any family member, or a Search Builder code or saved search name. Each matching person's family is updated once.", True, None],
            ["name", "string", "Extra value field name", True, None],
            ["dt", "any", "A date: a .NET DateTime, a Python datetime, or a string such as \"12/25/2026\"", True, None]
        ],
        "returns": "None",
        "watch_out": ["Pass a PeopleId, not a FamilyId.", "If the date cannot be read nothing is written and no error is raised.", "Family extra value field names are limited to 50 characters; a longer name raises an error when it saves.", "If the query matches no one, nothing happens and no error is raised.", "If the same field name already holds a different kind of value, both are kept on one row and the string readers (ExtraValue, ExtraValueCode) then return text such as \"unknown type: CodeText\"."],
        "example": "pid = model.UserPeopleId  # your own family\nmodel.AddExtraValueDateFamily(pid, \"ApiExplorerTest\", \"1/15/2026\")\nprint model.ExtraValueDateFamily(pid, \"ApiExplorerTest\").ToShortDateString()\nmodel.DeleteExtraValueFamily(pid, \"ApiExplorerTest\")  # clean up",
        "see_also": ["model.ExtraValueDateFamily", "model.DeleteExtraValueFamily", "model.AddExtraValueTextFamily", "model.AddExtraValueCodeFamily"]
    },
    "model.AddExtraValueDateOrg": {
        "category": "Extra Values",
        "kind": "method",
        "arity": [3, 3],
        "tags": ["writes_data"],
        "summary": "Adds or updates a Date extra value on an involvement.",
        "details": "Accepts a DateTime or a date string.",
        "flow": ["Loads the involvement by OrganizationId.", "Finds the OrganizationExtra row with this field name, or creates one.", "Sets the date and the time stamp, then saves."],
        "params": [
            ["oid", "integer", "OrganizationId of the involvement", True, None],
            ["name", "string", "Extra value field name", True, None],
            ["dt", "any", "A date: a .NET DateTime, a Python datetime, or a string such as \"12/25/2026\"", True, None]
        ],
        "returns": "None",
        "watch_out": ["If the date cannot be read nothing is written and no error is raised.", "Raises an error if the involvement does not exist.", "There is no Python method to delete an involvement extra value.", "Involvement extra value field names are limited to 50 characters; a longer name raises an error when it saves.", "The date is checked first, so an unreadable date returns quietly even when the involvement does not exist.", "If the same field name already holds a different kind of value, both are kept on one row and the string readers (ExtraValue, ExtraValueCode) then return text such as \"unknown type: CodeText\"."],
        "example": "row = q.QuerySqlTop1(\"SELECT TOP 1 OrganizationId, Field, DateValue FROM OrganizationExtra WHERE Type = 'Date'\")\nif row:\n    print row.OrganizationId, row.Field, row.DateValue\n# Writes to the involvement:\n# model.AddExtraValueDateOrg(row.OrganizationId, row.Field, \"1/15/2026\")",
        "see_also": ["model.ExtraValueDateOrg", "model.AddExtraValueTextOrg", "model.AddExtraValueCodeOrg"]
    },
    "model.AddExtraValueInt": {
        "category": "Extra Values",
        "kind": "method",
        "arity": [3, 3],
        "tags": ["writes_data"],
        "summary": "Adds or updates an Int (whole number) extra value on people.",
        "details": "The first argument can be a PeopleId or a Search Builder code, and every matching person is updated.",
        "flow": ["Runs the query (a PeopleId, a saved search name or Search Builder code) to get the list of people.", "For each person, finds the PeopleExtra row with this field name, or creates one if it is missing.", "Sets the number and the time stamp, then saves that person before moving to the next."],
        "params": [
            ["query", "query", "PeopleId, or a Search Builder code or saved search name. Every matching person is updated.", True, None],
            ["name", "string", "Extra value field name", True, None],
            ["n", "integer", "The number to store", True, None]
        ],
        "returns": "None",
        "watch_out": ["Updates every person the query matches.", "Passing None raises an error.", "Field names longer than 150 characters are cut to 150 without warning.", "If the query matches no one, nothing happens and no error is raised.", "Each person is saved separately, so a query that matches thousands of people is slow.", "If the same field name already holds a different kind of value, both are kept on one row and the string readers (ExtraValue, ExtraValueCode) then return text such as \"unknown type: CodeText\"."],
        "example": "pid = model.UserPeopleId\nmodel.AddExtraValueInt(pid, \"ApiExplorerTest\", 5)\nprint model.ExtraValueInt(pid, \"ApiExplorerTest\")\nmodel.DeleteExtraValue(pid, \"ApiExplorerTest\")  # clean up",
        "see_also": ["model.ExtraValueInt", "model.DeleteExtraValue", "model.AddExtraValueText", "model.AddExtraValueCode"]
    },
    "model.AddExtraValueIntFamily": {
        "category": "Extra Values",
        "kind": "method",
        "arity": [3, 3],
        "tags": ["writes_data"],
        "summary": "Adds or updates an Int extra value on families.",
        "details": "Family extra values are keyed by person, not FamilyId. Pass the PeopleId of any member and the value is stored on that person's family.",
        "flow": ["Runs the query, then keeps one person per family so each family is touched once.", "Looks up that person's FamilyId and finds or creates the FamilyExtra row with this field name.", "Sets the number and the time stamp, then saves that family."],
        "params": [
            ["query", "query", "PeopleId of any family member, or a Search Builder code or saved search name. Each matching person's family is updated once.", True, None],
            ["name", "string", "Extra value field name", True, None],
            ["n", "integer", "The number to store", True, None]
        ],
        "returns": "None",
        "watch_out": ["Pass a PeopleId, not a FamilyId.", "Passing None raises an error.", "Family extra value field names are limited to 50 characters; a longer name raises an error when it saves.", "If the query matches no one, nothing happens and no error is raised.", "If the same field name already holds a different kind of value, both are kept on one row and the string readers (ExtraValue, ExtraValueCode) then return text such as \"unknown type: CodeText\"."],
        "example": "pid = model.UserPeopleId  # your own family\nmodel.AddExtraValueIntFamily(pid, \"ApiExplorerTest\", 4)\nprint model.ExtraValueIntFamily(pid, \"ApiExplorerTest\")\nmodel.DeleteExtraValueFamily(pid, \"ApiExplorerTest\")  # clean up",
        "see_also": ["model.ExtraValueIntFamily", "model.DeleteExtraValueFamily", "model.AddExtraValueTextFamily", "model.AddExtraValueCodeFamily"]
    },
    "model.AddExtraValueIntOrg": {
        "category": "Extra Values",
        "kind": "method",
        "arity": [3, 3],
        "tags": ["writes_data"],
        "summary": "Adds or updates an Int extra value on an involvement.",
        "flow": ["Loads the involvement by OrganizationId.", "Finds the OrganizationExtra row with this field name, or creates one.", "Sets the number and the time stamp, then saves."],
        "params": [
            ["oid", "integer", "OrganizationId of the involvement", True, None],
            ["name", "string", "Extra value field name", True, None],
            ["n", "integer", "The number to store", True, None]
        ],
        "returns": "None",
        "watch_out": ["Raises an error if the involvement does not exist.", "There is no Python method to delete an involvement extra value.", "Passing None raises an error.", "Involvement extra value field names are limited to 50 characters; a longer name raises an error when it saves.", "If the same field name already holds a different kind of value, both are kept on one row and the string readers (ExtraValue, ExtraValueCode) then return text such as \"unknown type: CodeText\"."],
        "example": "row = q.QuerySqlTop1(\"SELECT TOP 1 OrganizationId, Field, IntValue FROM OrganizationExtra WHERE Type = 'Int'\")\nif row:\n    print row.OrganizationId, row.Field, row.IntValue\n# Writes to the involvement:\n# model.AddExtraValueIntOrg(row.OrganizationId, row.Field, 100)",
        "see_also": ["model.ExtraValueIntOrg", "model.AddExtraValueTextOrg", "model.AddExtraValueCodeOrg"]
    },
    "model.AddExtraValueMeeting": {
        "category": "Extra Values",
        "kind": "method",
        "arity": [3, 3],
        "tags": ["writes_data"],
        "summary": "Adds or replaces a text extra value on a meeting and saves it.",
        "details": "The value is converted to text before it is stored.",
        "flow": ["Converts the value to text. An empty result is ignored and nothing changes.", "Finds the meeting's extra value row with that field name (spaces trimmed), or creates one.", "Puts the text in the Data column, stamps the time, and saves."],
        "params": [
            ["meetingId", "integer", "MeetingId", True, None],
            ["field", "string", "Extra value name", True, None],
            ["value", "any", "Value to store (converted to text)", True, None]
        ],
        "returns": "None",
        "watch_out": ["An empty string is ignored, so it cannot clear a value.", "None raises an error.", "A datetime is stored in US text form, such as 11/2/2026 12:00:00 AM, not as a date. True is stored as True."],
        "example": "mid = q.QuerySqlInt(\"SELECT TOP 1 MeetingId FROM Meetings WHERE MeetingDate < GETDATE() ORDER BY MeetingDate DESC\")\n# Writes data. Uncomment to set it:\n# model.AddExtraValueMeeting(mid, \"Room\", \"101\")\nprint \"Meeting\", mid, \"Room:\", model.ExtraValueMeeting(mid, \"Room\") or \"(not set)\"",
        "see_also": ["model.ExtraValueMeeting", "model.DeleteExtraValueMeeting", "model.GetMeeting"]
    },
    "model.AddExtraValueText": {
        "category": "Extra Values",
        "kind": "method",
        "arity": [3, 3],
        "tags": ["writes_data"],
        "summary": "Adds or updates a Text extra value on people.",
        "details": "Text is long free text, stored in Data. There is no type argument; use the AddExtraValueCode, Date, Int or Bool methods for other types. The first argument can be a PeopleId or a Search Builder code, and every matching person is updated.",
        "flow": ["Turns the query into a list of people. A value made only of digits is a PeopleId; otherwise it is a saved search name or a Search Builder code.", "Converts text to a string. An empty string skips the person and leaves any existing value alone.", "Finds that person's PeopleExtra row for the field name (cut to 150 characters), or creates one.", "Puts the text in the Data column, stamps TransactionTime, and saves. Each person is saved separately."],
        "params": [
            ["query", "query", "PeopleId, or a Search Builder code or saved search name. Every matching person is updated.", True, None],
            ["name", "string", "Extra value field name", True, None],
            ["text", "string", "The text to store", True, None]
        ],
        "returns": "None",
        "watch_out": ["Updates every person the query matches.", "An empty string is silently ignored.", "None raises an error.", "You cannot clear a value by writing an empty string. Use model.DeleteExtraValue.", "There is one PeopleExtra row per person and field name. Writing text to a field that already holds a code, date, number or yes/no value keeps that value in the same row.", "A number or True is stored as its text, such as 5 or True."],
        "example": "pid = model.UserPeopleId\nmodel.AddExtraValueText(pid, \"ApiExplorerTest\", \"Important info\")\nprint model.ExtraValueText(pid, \"ApiExplorerTest\")\nmodel.DeleteExtraValue(pid, \"ApiExplorerTest\")  # clean up",
        "see_also": ["model.ExtraValueText", "model.DeleteExtraValue", "model.AddExtraValueCode"]
    },
    "model.AddExtraValueTextFamily": {
        "category": "Extra Values",
        "kind": "method",
        "arity": [3, 3],
        "tags": ["writes_data"],
        "summary": "Adds or updates a Text extra value on families.",
        "details": "Family extra values are keyed by person, not FamilyId. Pass the PeopleId of any member and the value is stored on that person's family.",
        "flow": ["Runs the query, then keeps one person per family so each family is touched once.", "Looks up that person's FamilyId and finds or creates the FamilyExtra row with this field name.", "Sets the text and the time stamp, then saves that family."],
        "params": [
            ["query", "query", "PeopleId of any family member, or a Search Builder code or saved search name. Each matching person's family is updated once.", True, None],
            ["name", "string", "Extra value field name", True, None],
            ["text", "string", "The text to store", True, None]
        ],
        "returns": "None",
        "watch_out": ["Pass a PeopleId, not a FamilyId.", "An empty string is silently ignored.", "Passing None raises an error.", "Family extra value field names are limited to 50 characters; a longer name raises an error when it saves.", "If the query matches no one, nothing happens and no error is raised.", "If the same field name already holds a different kind of value, both are kept on one row and the string readers (ExtraValue, ExtraValueCode) then return text such as \"unknown type: CodeText\"."],
        "example": "pid = model.UserPeopleId  # your own family\nmodel.AddExtraValueTextFamily(pid, \"ApiExplorerTest\", \"Allergic to peanuts\")\nprint model.ExtraValueTextFamily(pid, \"ApiExplorerTest\")\nmodel.DeleteExtraValueFamily(pid, \"ApiExplorerTest\")  # clean up",
        "see_also": ["model.ExtraValueTextFamily", "model.DeleteExtraValueFamily", "model.AddExtraValueCodeFamily"]
    },
    "model.AddExtraValueTextOrg": {
        "category": "Extra Values",
        "kind": "method",
        "arity": [3, 3],
        "tags": ["writes_data"],
        "summary": "Adds or updates a Text extra value on an involvement.",
        "flow": ["Loads the involvement by OrganizationId.", "Finds the OrganizationExtra row with this field name, or creates one.", "Sets the text and the time stamp, then saves."],
        "params": [
            ["oid", "integer", "OrganizationId of the involvement", True, None],
            ["name", "string", "Extra value field name", True, None],
            ["text", "string", "The text to store", True, None]
        ],
        "returns": "None",
        "watch_out": ["An empty string is silently ignored.", "Raises an error if the involvement does not exist.", "There is no Python method to delete an involvement extra value.", "Passing None raises an error.", "Involvement extra value field names are limited to 50 characters; a longer name raises an error when it saves.", "If the same field name already holds a different kind of value, both are kept on one row and the string readers (ExtraValue, ExtraValueCode) then return text such as \"unknown type: CodeText\"."],
        "example": "row = q.QuerySqlTop1(\"SELECT TOP 1 OrganizationId, Field, Data FROM OrganizationExtra WHERE Type = 'Text'\")\nif row:\n    print row.OrganizationId, row.Field, row.Data\n# Writes to the involvement:\n# model.AddExtraValueTextOrg(row.OrganizationId, row.Field, \"Meets in room 210\")",
        "see_also": ["model.ExtraValueTextOrg", "model.AddExtraValueCodeOrg"]
    },
    "model.DeleteAllExtraValueLike": {
        "category": "Extra Values",
        "kind": "method",
        "arity": [1, 1],
        "tags": ["writes_data"],
        "summary": "Deletes every person extra value in the database whose field name matches a SQL LIKE pattern.",
        "details": "% is the wildcard. The action is written to the activity log.",
        "flow": ["Writes an activity log entry that shows the pattern.", "Deletes every PeopleExtra row whose field name matches the pattern, for every person and every type, in one statement."],
        "params": [
            ["value", "string", "SQL LIKE pattern for the field name, such as \"Temp%\"", True, None]
        ],
        "returns": "None",
        "watch_out": ["It is not limited to one person; it affects the whole database.", "It cannot be undone.", "The underscore is also a wildcard that matches any one character, so \"My_Field\" also deletes \"MyXField\". Write [_] to match a real underscore.", "There is no role check, so any script that calls it can run it.", "A pattern of \"%\" deletes every person extra value in the database."],
        "example": "pattern = \"ApiExplorerTest%\"\nprint q.QuerySqlInt(\"SELECT COUNT(*) FROM PeopleExtra WHERE Field LIKE '%s'\" % pattern), \"rows would be deleted\"\n# Deletes for EVERY person in the database:\n# model.DeleteAllExtraValueLike(pattern)",
        "see_also": ["model.DeleteExtraValue", "model.DeleteAllExtraValueLikeFamily"]
    },
    "model.DeleteAllExtraValueLikeFamily": {
        "category": "Extra Values",
        "kind": "method",
        "arity": [1, 1],
        "tags": ["writes_data"],
        "summary": "Deletes every family extra value in the database whose field name matches a SQL LIKE pattern.",
        "details": "% is the wildcard. It takes only the pattern. The action is written to the activity log.",
        "flow": ["Writes an activity log entry that shows the pattern.", "Deletes every FamilyExtra row whose field name matches the pattern, for every family, in one statement."],
        "params": [
            ["value", "string", "SQL LIKE pattern for the field name, such as \"BP:%\"", True, None]
        ],
        "returns": "None",
        "watch_out": ["It is not limited to one family; it affects the whole database.", "It cannot be undone.", "The underscore is also a wildcard that matches any one character, so \"My_Field\" also deletes \"MyXField\". Write [_] to match a real underscore.", "There is no role check, so any script that calls it can run it."],
        "example": "pattern = \"ApiExplorerTest%\"\nprint q.QuerySqlInt(\"SELECT COUNT(*) FROM FamilyExtra WHERE Field LIKE '%s'\" % pattern), \"rows would be deleted\"\n# Deletes for EVERY family in the database:\n# model.DeleteAllExtraValueLikeFamily(pattern)",
        "see_also": ["model.DeleteExtraValueFamily", "model.DeleteAllExtraValueLike"]
    },
    "model.DeleteExtraValue": {
        "category": "Extra Values",
        "kind": "method",
        "arity": [2, 2],
        "tags": ["writes_data"],
        "summary": "Deletes an extra value field from people.",
        "details": "The first argument can be a PeopleId or a Search Builder code, and every matching person is updated.",
        "flow": ["Runs the query to get the list of people.", "For each person, finds the PeopleExtra row with this field name and deletes it, saving one person at a time."],
        "params": [
            ["query", "query", "PeopleId, or a Search Builder code or saved search name. Every matching person is updated.", True, None],
            ["name", "string", "Extra value field name to delete", True, None]
        ],
        "returns": "None",
        "watch_out": ["Updates every person the query matches.", "Does nothing for people who do not have the field.", "It removes the whole field, every type of value stored under that name.", "If the query matches no one, nothing happens and no error is raised.", "Each person is saved separately, so a query that matches thousands of people is slow."],
        "example": "pid = model.UserPeopleId\nmodel.AddExtraValueText(pid, \"ApiExplorerTest\", \"temp\")\nmodel.DeleteExtraValue(pid, \"ApiExplorerTest\")\nprint \"Deleted ApiExplorerTest for\", pid",
        "see_also": ["model.DeleteAllExtraValueLike", "model.DeleteExtraValueFamily", "model.AddExtraValueText", "model.ExtraValue"]
    },
    "model.DeleteExtraValueFamily": {
        "category": "Extra Values",
        "kind": "method",
        "arity": [2, 2],
        "tags": ["writes_data"],
        "summary": "Deletes an extra value field from families.",
        "details": "Family extra values are keyed by person, not FamilyId. Pass the PeopleId of any member of the family. Each matching person's family is updated once.",
        "flow": ["Runs the query, then keeps one person per family.", "For each family, finds the FamilyExtra row with this field name and deletes it."],
        "params": [
            ["query", "query", "PeopleId of any family member, or a Search Builder code or saved search name. Each matching person's family is updated once.", True, None],
            ["name", "string", "Extra value field name to delete", True, None]
        ],
        "returns": "None",
        "watch_out": ["Pass a PeopleId, not a FamilyId.", "Does nothing for families that do not have the field.", "If the query matches no one, nothing happens and no error is raised."],
        "example": "pid = model.UserPeopleId  # your own family\nmodel.AddExtraValueTextFamily(pid, \"ApiExplorerTest\", \"temp\")\nmodel.DeleteExtraValueFamily(pid, \"ApiExplorerTest\")\nprint \"Deleted ApiExplorerTest from your family\"",
        "see_also": ["model.DeleteAllExtraValueLikeFamily", "model.DeleteExtraValue", "model.ExtraValueFamily"]
    },
    "model.DeleteExtraValueMeeting": {
        "category": "Extra Values",
        "kind": "method",
        "arity": [2, 2],
        "tags": ["writes_data"],
        "summary": "Deletes the named extra value from a meeting and saves.",
        "details": "Returns the number of extra values deleted.",
        "params": [
            ["meetingId", "integer", "MeetingId", True, None],
            ["field", "string", "Extra value name", True, None]
        ],
        "returns": "Integer. Number of extra values deleted.",
        "watch_out": ["Returns 0, with no error, when the meeting or the field does not exist."],
        "example": "mid = q.QuerySqlInt(\"SELECT TOP 1 MeetingId FROM Meetings WHERE MeetingDate < GETDATE() ORDER BY MeetingDate DESC\")\nprint \"Current:\", model.ExtraValueMeeting(mid, \"Room\")\n# Writes data. Uncomment to delete it:\n# print model.DeleteExtraValueMeeting(mid, \"Room\"), \"deleted\"",
        "see_also": ["model.AddExtraValueMeeting", "model.ExtraValueMeeting"]
    },
    "model.ExtraValue": {
        "category": "Extra Values",
        "kind": "method",
        "arity": [2, 2],
        "tags": ["writes_data", "known_issue"],
        "summary": "Gets a person's extra value as a string, whatever its type.",
        "details": "Code and Text return the value, Date returns a short date, Bit returns \"True\" or \"False\", Int returns the number. A field holding several kinds of values returns them joined with semicolons as code;date;bit;int;text.",
        "flow": ["Looks for the person's PeopleExtra row with that field name.", "If there is no row, it makes a blank one in memory and queues it to be saved. It never treats the field as simply missing.", "Returns the value by type: Code gives the code, Text gives the text, Date gives a short date, Bit gives True or False, and Int gives the number. A row of type Data gives all parts joined with semicolons: code;date;bit;int;text.", "A blank row has no type yet, so a missing field returns unknown type: ."],
        "params": [
            ["pid", "integer", "PeopleId", True, None],
            ["name", "string", "Extra value field name", True, None]
        ],
        "returns": "string",
        "watch_out": ["Returns \"unknown type: \" rather than an empty string when the person does not have the field.", "Reading a field the person does not have can queue an empty extra value row that a later write in the same script saves. Check that the field exists when that matters."],
        "example": "row = q.QuerySqlTop1(\"SELECT TOP 1 PeopleId, Field FROM PeopleExtra WHERE PeopleId = @p1\", model.UserPeopleId)\nif row:\n    print row.Field, \"=\", model.ExtraValue(row.PeopleId, row.Field)\nelse:\n    print \"You have no extra values\"",
        "see_also": ["model.ExtraValueText", "model.ExtraValueFamily", "model.ExtraValueOrg", "model.AddExtraValueText"]
    },
    "model.ExtraValueAttributes": {
        "category": "Extra Values",
        "kind": "method",
        "arity": [2, 2],
        "tags": ["writes_data", "known_issue"],
        "summary": "Gets the JSON text of an Attributes extra value for a person.",
        "params": [
            ["pid", "integer", "PeopleId", True, None],
            ["name", "string", "Extra value field name", True, None]
        ],
        "returns": "string (JSON text, or \"\")",
        "watch_out": ["Returns an empty string if the field is missing or was not saved as attributes, for example a plain Text value.", "Reading a field the person does not have can queue an empty extra value row that a later write in the same script saves. Check that the field exists when that matters."],
        "example": "import json\nrow = q.QuerySqlTop1(\"SELECT TOP 1 PeopleId, Field FROM PeopleExtra WHERE IsAttributes = 1\")\nif row:\n    text = model.ExtraValueAttributes(row.PeopleId, row.Field)\n    print row.Field, \"raw value:\", text\n    try:\n        print \"As JSON:\", json.loads(text)\n    except ValueError:\n        print \"(not JSON)\"\nelse:\n    print \"No attribute extra values found.\"",
        "see_also": ["model.AddExtraValueAttributes", "model.ExtraValueText", "model.DynamicDataFromJson"]
    },
    "model.ExtraValueBit": {
        "category": "Extra Values",
        "kind": "method",
        "arity": [2, 2],
        "tags": ["writes_data", "known_issue"],
        "summary": "Gets a Bit (True/False) extra value for a person.",
        "params": [
            ["pid", "integer", "PeopleId", True, None],
            ["name", "string", "Extra value field name", True, None]
        ],
        "returns": "boolean",
        "watch_out": ["Returns False when the field is missing or empty.", "Reading a field the person does not have can queue an empty extra value row that a later write in the same script saves. Check that the field exists when that matters."],
        "example": "row = q.QuerySqlTop1(\"SELECT TOP 1 PeopleId, Field FROM PeopleExtra WHERE Type = 'Bit'\")\nif row:\n    print row.PeopleId, row.Field, \"=\", model.ExtraValueBit(row.PeopleId, row.Field)",
        "see_also": ["model.AddExtraValueBool", "model.ExtraValueBitFamily", "model.ExtraValueBitOrg"]
    },
    "model.ExtraValueBitFamily": {
        "category": "Extra Values",
        "kind": "method",
        "arity": [2, 2],
        "tags": ["writes_data", "known_issue"],
        "summary": "Gets a Bit (True/False) family extra value.",
        "details": "Family extra values are keyed by person, not FamilyId. Pass the PeopleId of any member of the family.",
        "params": [
            ["pid", "integer", "PeopleId of any member of the family (not the FamilyId)", True, None],
            ["name", "string", "Extra value field name", True, None]
        ],
        "returns": "boolean",
        "watch_out": ["Returns False when missing.", "Pass a PeopleId, not a FamilyId.", "Reading a field the family does not have can queue an empty row that a later write in the same script saves.", "Returns False when the PeopleId does not exist."],
        "example": "row = q.QuerySqlTop1(\"SELECT TOP 1 p.PeopleId, fe.Field FROM FamilyExtra fe JOIN People p ON p.FamilyId = fe.FamilyId WHERE fe.Type = 'Bit'\")\nif row:\n    print row.Field, \"=\", model.ExtraValueBitFamily(row.PeopleId, row.Field)",
        "see_also": ["model.AddExtraValueBoolFamily", "model.ExtraValueBit"]
    },
    "model.ExtraValueBitOrg": {
        "category": "Extra Values",
        "kind": "method",
        "arity": [2, 2],
        "tags": ["writes_data", "known_issue"],
        "summary": "Gets a Bit (True/False) extra value for an involvement.",
        "params": [
            ["oid", "integer", "OrganizationId of the involvement", True, None],
            ["name", "string", "Extra value field name", True, None]
        ],
        "returns": "boolean",
        "watch_out": ["Returns False when missing.", "Reading a field the involvement does not have can queue an empty row that a later write in the same script saves."],
        "example": "row = q.QuerySqlTop1(\"SELECT TOP 1 OrganizationId, Field FROM OrganizationExtra WHERE Type = 'Bit'\")\nif row:\n    print row.OrganizationId, row.Field, \"=\", model.ExtraValueBitOrg(row.OrganizationId, row.Field)",
        "see_also": ["model.AddExtraValueBoolOrg", "model.ExtraValueBit"]
    },
    "model.ExtraValueCode": {
        "category": "Extra Values",
        "kind": "method",
        "arity": [2, 2],
        "tags": ["writes_data", "known_issue"],
        "summary": "Gets a Code extra value for a person.",
        "details": "This is the same as ExtraValue, so it returns any type as a string.",
        "params": [
            ["pid", "integer", "PeopleId", True, None],
            ["name", "string", "Extra value field name", True, None]
        ],
        "returns": "string",
        "watch_out": ["Returns \"unknown type: \" when the person does not have the field.", "Reading a field the person does not have can queue an empty extra value row that a later write in the same script saves. Check that the field exists when that matters.", "If one field name holds more than one kind of value, it returns text such as \"unknown type: CodeText\".", "An Attributes field comes back as \"unknown type: Attr\"; use ExtraValueAttributes for those.", "A Data style field (all values set at once) comes back as code;date;bit;int;text joined with semicolons."],
        "example": "row = q.QuerySqlTop1(\"SELECT TOP 1 PeopleId, Field FROM PeopleExtra WHERE Type = 'Code'\")\nif row:\n    print row.PeopleId, row.Field, \"=\", model.ExtraValueCode(row.PeopleId, row.Field)",
        "see_also": ["model.ExtraValue", "model.AddExtraValueCode", "model.ExtraValueCodeFamily"]
    },
    "model.ExtraValueCodeFamily": {
        "category": "Extra Values",
        "kind": "method",
        "arity": [2, 2],
        "tags": ["writes_data", "known_issue"],
        "summary": "Gets a Code family extra value.",
        "details": "This is the same as ExtraValueFamily, so any type comes back as a string. Family extra values are keyed by person, not FamilyId. Pass the PeopleId of any member of the family.",
        "params": [
            ["pid", "integer", "PeopleId of any member of the family (not the FamilyId)", True, None],
            ["name", "string", "Extra value field name", True, None]
        ],
        "returns": "string",
        "watch_out": ["Returns \"unknown type: \" when the field is missing.", "Pass a PeopleId, not a FamilyId.", "Reading a field the family does not have can queue an empty row that a later write in the same script saves.", "Returns \"\" when the PeopleId does not exist.", "If one field name holds more than one kind of value, it returns text such as \"unknown type: CodeText\"."],
        "example": "row = q.QuerySqlTop1(\"SELECT TOP 1 p.PeopleId, fe.Field FROM FamilyExtra fe JOIN People p ON p.FamilyId = fe.FamilyId WHERE fe.Type = 'Code'\")\nif row:\n    print row.Field, \"=\", model.ExtraValueCodeFamily(row.PeopleId, row.Field)",
        "see_also": ["model.ExtraValueFamily", "model.AddExtraValueCodeFamily", "model.ExtraValueCode"]
    },
    "model.ExtraValueCodeOrg": {
        "category": "Extra Values",
        "kind": "method",
        "arity": [2, 2],
        "tags": ["writes_data", "known_issue"],
        "summary": "Gets a Code extra value for an involvement.",
        "details": "This is the same as ExtraValueOrg, so any type comes back as a string.",
        "params": [
            ["oid", "integer", "OrganizationId of the involvement", True, None],
            ["name", "string", "Extra value field name", True, None]
        ],
        "returns": "string",
        "watch_out": ["Returns \"unknown type: \" when the field is missing.", "Reading a field the involvement does not have can queue an empty row that a later write in the same script saves.", "If one field name holds more than one kind of value, it returns text such as \"unknown type: CodeText\"."],
        "example": "row = q.QuerySqlTop1(\"SELECT TOP 1 OrganizationId, Field FROM OrganizationExtra WHERE Type = 'Code'\")\nif row:\n    print row.OrganizationId, row.Field, \"=\", model.ExtraValueCodeOrg(row.OrganizationId, row.Field)",
        "see_also": ["model.ExtraValueOrg", "model.AddExtraValueCodeOrg"]
    },
    "model.ExtraValueDate": {
        "category": "Extra Values",
        "kind": "method",
        "arity": [2, 2],
        "tags": ["writes_data", "known_issue"],
        "summary": "Gets a Date extra value for a person.",
        "params": [
            ["pid", "integer", "PeopleId", True, None],
            ["name", "string", "Extra value field name", True, None]
        ],
        "returns": "DateTime (1/1/0001 when missing)",
        "watch_out": ["Returns DateTime.MinValue (1/1/0001), not None, when the field is missing or empty, so test the year before using it.", "Reading a field the person does not have can queue an empty extra value row that a later write in the same script saves. Check that the field exists when that matters."],
        "example": "row = q.QuerySqlTop1(\"SELECT TOP 1 PeopleId, Field FROM PeopleExtra WHERE Type = 'Date'\")\nif row:\n    d = model.ExtraValueDate(row.PeopleId, row.Field)\n    if d.Year > 1:\n        print row.Field, d.ToShortDateString()",
        "see_also": ["model.AddExtraValueDate", "model.ExtraValueDateFamily", "model.ExtraValueDateOrgNull"]
    },
    "model.ExtraValueDateFamily": {
        "category": "Extra Values",
        "kind": "method",
        "arity": [2, 2],
        "tags": ["writes_data", "known_issue"],
        "summary": "Gets a Date family extra value.",
        "details": "Family extra values are keyed by person, not FamilyId. Pass the PeopleId of any member of the family.",
        "params": [
            ["pid", "integer", "PeopleId of any member of the family (not the FamilyId)", True, None],
            ["name", "string", "Extra value field name", True, None]
        ],
        "returns": "DateTime (1/1/0001 when missing)",
        "watch_out": ["Returns DateTime.MinValue (1/1/0001), not None, when missing.", "Pass a PeopleId, not a FamilyId.", "Reading a field the family does not have can queue an empty row that a later write in the same script saves.", "Returns 1/1/0001 when the PeopleId does not exist."],
        "example": "row = q.QuerySqlTop1(\"SELECT TOP 1 p.PeopleId, fe.Field FROM FamilyExtra fe JOIN People p ON p.FamilyId = fe.FamilyId WHERE fe.Type = 'Date'\")\nif row:\n    d = model.ExtraValueDateFamily(row.PeopleId, row.Field)\n    print row.Field, d.ToShortDateString() if d.Year > 1 else \"(no date)\"\nelse:\n    print \"No family date extra values found.\"",
        "see_also": ["model.AddExtraValueDateFamily", "model.ExtraValueDate"]
    },
    "model.ExtraValueDateOrg": {
        "category": "Extra Values",
        "kind": "method",
        "arity": [2, 2],
        "tags": ["writes_data", "known_issue"],
        "summary": "Gets a Date extra value for an involvement.",
        "details": "Use ExtraValueDateOrgNull to get None instead.",
        "params": [
            ["oid", "integer", "OrganizationId of the involvement", True, None],
            ["name", "string", "Extra value field name", True, None]
        ],
        "returns": "DateTime (1/1/0001 when missing)",
        "watch_out": ["Returns DateTime.MinValue (1/1/0001) when missing.", "Reading a field the involvement does not have can queue an empty row that a later write in the same script saves."],
        "example": "row = q.QuerySqlTop1(\"SELECT TOP 1 OrganizationId, Field FROM OrganizationExtra WHERE Type = 'Date'\")\nif row:\n    d = model.ExtraValueDateOrg(row.OrganizationId, row.Field)\n    if d.Year > 1:\n        print row.Field, d.ToShortDateString()",
        "see_also": ["model.ExtraValueDateOrgNull", "model.AddExtraValueDateOrg", "model.ExtraValueDate"]
    },
    "model.ExtraValueDateOrgNull": {
        "category": "Extra Values",
        "kind": "method",
        "arity": [2, 2],
        "tags": ["writes_data", "known_issue"],
        "summary": "Gets a Date extra value for an involvement, or None when the field is missing or has no date.",
        "params": [
            ["oid", "integer", "OrganizationId of the involvement", True, None],
            ["name", "string", "Extra value field name", True, None]
        ],
        "returns": "DateTime or None",
        "watch_out": ["Reading a field the involvement does not have can queue an empty row that a later write in the same script saves.", "Returns None when the field is missing or holds no date; it does not raise."],
        "example": "row = q.QuerySqlTop1(\"SELECT TOP 1 OrganizationId, Field FROM OrganizationExtra WHERE Type = 'Date'\")\nif row:\n    d = model.ExtraValueDateOrgNull(row.OrganizationId, row.Field)\n    if d:\n        print row.Field, d.ToShortDateString()",
        "see_also": ["model.ExtraValueDateOrg", "model.AddExtraValueDateOrg"]
    },
    "model.ExtraValueFamily": {
        "category": "Extra Values",
        "kind": "method",
        "arity": [2, 2],
        "tags": ["writes_data", "known_issue"],
        "summary": "Gets a family extra value as a string, whatever its type.",
        "details": "Code and Text return the value, Date returns a short date, Bit returns \"True\" or \"False\", Int returns the number. Family extra values are keyed by person, not FamilyId. Pass the PeopleId of any member of the family.",
        "params": [
            ["pid", "integer", "PeopleId of any member of the family (not the FamilyId)", True, None],
            ["name", "string", "Extra value field name", True, None]
        ],
        "returns": "string",
        "watch_out": ["Returns \"unknown type: \" when the family does not have the field.", "Returns \"\" when the PeopleId does not exist.", "Pass a PeopleId, not a FamilyId.", "Reading a field the family does not have can queue an empty row that a later write in the same script saves.", "If one field name holds more than one kind of value, it returns text such as \"unknown type: CodeText\"."],
        "example": "row = q.QuerySqlTop1(\"SELECT TOP 1 p.PeopleId, fe.Field FROM FamilyExtra fe JOIN People p ON p.FamilyId = fe.FamilyId\")\nif row:\n    print row.Field, \"=\", model.ExtraValueFamily(row.PeopleId, row.Field)",
        "see_also": ["model.ExtraValue", "model.ExtraValueTextFamily", "model.AddExtraValueTextFamily", "model.DeleteExtraValueFamily"]
    },
    "model.ExtraValueInt": {
        "category": "Extra Values",
        "kind": "method",
        "arity": [2, 2],
        "tags": ["writes_data", "known_issue"],
        "summary": "Gets an Int extra value for a person.",
        "params": [
            ["pid", "integer", "PeopleId", True, None],
            ["name", "string", "Extra value field name", True, None]
        ],
        "returns": "integer (0 when missing)",
        "watch_out": ["Returns 0 when the field is missing or empty.", "Reading a field the person does not have can queue an empty extra value row that a later write in the same script saves. Check that the field exists when that matters."],
        "example": "row = q.QuerySqlTop1(\"SELECT TOP 1 PeopleId, Field FROM PeopleExtra WHERE Type = 'Int'\")\nif row:\n    print row.PeopleId, row.Field, \"=\", model.ExtraValueInt(row.PeopleId, row.Field)",
        "see_also": ["model.AddExtraValueInt", "model.ExtraValueIntFamily", "model.ExtraValueIntOrg"]
    },
    "model.ExtraValueIntFamily": {
        "category": "Extra Values",
        "kind": "method",
        "arity": [2, 2],
        "tags": ["writes_data", "known_issue"],
        "summary": "Gets an Int family extra value.",
        "details": "Family extra values are keyed by person, not FamilyId. Pass the PeopleId of any member of the family.",
        "params": [
            ["pid", "integer", "PeopleId of any member of the family (not the FamilyId)", True, None],
            ["name", "string", "Extra value field name", True, None]
        ],
        "returns": "integer (0 when missing)",
        "watch_out": ["Returns 0 when missing.", "Pass a PeopleId, not a FamilyId.", "Reading a field the family does not have can queue an empty row that a later write in the same script saves.", "Returns 0 when the PeopleId does not exist."],
        "example": "row = q.QuerySqlTop1(\"SELECT TOP 1 p.PeopleId, fe.Field FROM FamilyExtra fe JOIN People p ON p.FamilyId = fe.FamilyId WHERE fe.Type = 'Int'\")\nif row:\n    print row.Field, \"=\", model.ExtraValueIntFamily(row.PeopleId, row.Field)",
        "see_also": ["model.AddExtraValueIntFamily", "model.ExtraValueInt"]
    },
    "model.ExtraValueIntOrg": {
        "category": "Extra Values",
        "kind": "method",
        "arity": [2, 2],
        "tags": ["writes_data", "known_issue"],
        "summary": "Gets an Int extra value for an involvement.",
        "params": [
            ["oid", "integer", "OrganizationId of the involvement", True, None],
            ["name", "string", "Extra value field name", True, None]
        ],
        "returns": "integer (0 when missing)",
        "watch_out": ["Returns 0 when missing.", "Reading a field the involvement does not have can queue an empty row that a later write in the same script saves."],
        "example": "row = q.QuerySqlTop1(\"SELECT TOP 1 OrganizationId, Field FROM OrganizationExtra WHERE Type = 'Int'\")\nif row:\n    print row.OrganizationId, row.Field, \"=\", model.ExtraValueIntOrg(row.OrganizationId, row.Field)",
        "see_also": ["model.AddExtraValueIntOrg", "model.ExtraValueInt"]
    },
    "model.ExtraValueMeeting": {
        "category": "Extra Values",
        "kind": "method",
        "arity": [2, 2],
        "tags": ["read_only"],
        "summary": "Returns the value of a meeting extra value.",
        "params": [
            ["meetingId", "integer", "MeetingId", True, None],
            ["field", "string", "Extra value name", True, None]
        ],
        "returns": "String. \"\" when the field is missing, None when the meeting does not exist.",
        "watch_out": ["Returns \"\" if the field does not exist, and None if the meeting does not exist.", "It checks the slots in order Code, Text, Date, Int, Bit and returns the first one that has a value, as a string. A Bit comes back as \"True\" or \"False\"."],
        "example": "mid = q.QuerySqlInt(\"SELECT TOP 1 MeetingId FROM Meetings WHERE MeetingDate < GETDATE() ORDER BY MeetingDate DESC\")\nroom = model.ExtraValueMeeting(mid, \"Room\")\nprint \"Meeting\", mid, \"Room:\", room or \"(not set)\"",
        "see_also": ["model.AddExtraValueMeeting", "model.DeleteExtraValueMeeting", "model.GetMeeting"]
    },
    "model.ExtraValueOrg": {
        "category": "Extra Values",
        "kind": "method",
        "arity": [2, 2],
        "tags": ["writes_data", "known_issue"],
        "summary": "Gets an involvement extra value as a string, whatever its type.",
        "details": "Code and Text return the value, Date returns a short date, Bit returns \"True\" or \"False\", Int returns the number.",
        "params": [
            ["oid", "integer", "OrganizationId of the involvement", True, None],
            ["name", "string", "Extra value field name", True, None]
        ],
        "returns": "string",
        "watch_out": ["Returns \"unknown type: \" rather than an empty string when the involvement does not have the field.", "Reading a field the involvement does not have can queue an empty row that a later write in the same script saves.", "If one field name holds more than one kind of value, it returns text such as \"unknown type: CodeText\"."],
        "example": "row = q.QuerySqlTop1(\"SELECT TOP 1 OrganizationId, Field FROM OrganizationExtra\")\nif row:\n    print row.Field, \"=\", model.ExtraValueOrg(row.OrganizationId, row.Field)",
        "see_also": ["model.ExtraValue", "model.ExtraValueTextOrg", "model.AddExtraValueTextOrg", "model.GetOrganization"]
    },
    "model.ExtraValueText": {
        "category": "Extra Values",
        "kind": "method",
        "arity": [2, 2],
        "tags": ["writes_data", "known_issue"],
        "summary": "Gets a Text extra value for a person.",
        "params": [
            ["pid", "integer", "PeopleId", True, None],
            ["name", "string", "Extra value field name", True, None]
        ],
        "returns": "string (\"\" when missing)",
        "watch_out": ["Returns an empty string when the field is missing.", "It reads only the Text slot, so a Code value comes back empty; use ExtraValue for those.", "Reading a field the person does not have can queue an empty extra value row that a later write in the same script saves. Check that the field exists when that matters."],
        "example": "row = q.QuerySqlTop1(\"SELECT TOP 1 PeopleId, Field FROM PeopleExtra WHERE Type = 'Text'\")\nif row:\n    print row.PeopleId, row.Field, \"=\", model.ExtraValueText(row.PeopleId, row.Field)",
        "see_also": ["model.ExtraValue", "model.AddExtraValueText", "model.ExtraValueTextFamily", "model.ExtraValueTextOrg"]
    },
    "model.ExtraValueTextFamily": {
        "category": "Extra Values",
        "kind": "method",
        "arity": [2, 2],
        "tags": ["writes_data", "known_issue"],
        "summary": "Gets a Text family extra value.",
        "details": "Family extra values are keyed by person, not FamilyId. Pass the PeopleId of any member of the family.",
        "params": [
            ["pid", "integer", "PeopleId of any member of the family (not the FamilyId)", True, None],
            ["name", "string", "Extra value field name", True, None]
        ],
        "returns": "string (\"\" when missing)",
        "watch_out": ["Returns an empty string when missing.", "Pass a PeopleId, not a FamilyId.", "Reading a field the family does not have can queue an empty row that a later write in the same script saves.", "Returns \"\" when the PeopleId does not exist."],
        "example": "row = q.QuerySqlTop1(\"SELECT TOP 1 p.PeopleId, fe.Field FROM FamilyExtra fe JOIN People p ON p.FamilyId = fe.FamilyId WHERE fe.Type = 'Text'\")\nif row:\n    print row.Field, \"=\", model.ExtraValueTextFamily(row.PeopleId, row.Field)",
        "see_also": ["model.AddExtraValueTextFamily", "model.ExtraValueText", "model.ExtraValueFamily"]
    },
    "model.ExtraValueTextOrg": {
        "category": "Extra Values",
        "kind": "method",
        "arity": [2, 2],
        "tags": ["writes_data", "known_issue"],
        "summary": "Gets a Text extra value for an involvement.",
        "params": [
            ["oid", "integer", "OrganizationId of the involvement", True, None],
            ["name", "string", "Extra value field name", True, None]
        ],
        "returns": "string (\"\" when missing)",
        "watch_out": ["Returns an empty string when missing.", "Reading a field the involvement does not have can queue an empty row that a later write in the same script saves."],
        "example": "row = q.QuerySqlTop1(\"SELECT TOP 1 OrganizationId, Field FROM OrganizationExtra WHERE Type = 'Text'\")\nif row:\n    print row.OrganizationId, row.Field, \"=\", model.ExtraValueTextOrg(row.OrganizationId, row.Field)",
        "see_also": ["model.AddExtraValueTextOrg", "model.ExtraValueText", "model.ExtraValueOrg"]
    },
    "model.AddMembersToOrg": {
        "category": "Involvements",
        "kind": "method",
        "arity": [2, 2],
        "tags": ["writes_data"],
        "summary": "Adds everyone matched by a query to an involvement as Member, enrolled now.",
        "details": "People already enrolled stay enrolled. If their current type is Member or Prospect it is set to Member, otherwise it is kept. Logs the change.",
        "flow": ["Runs the query to get the list of people.", "For each person, looks for an existing enrollment in the involvement.", "If they are already enrolled, clears Pending and changes a Member or Prospect type to Member. Other types, such as Leader, are kept.", "If they are not enrolled, adds an OrganizationMembers row as Member with today's date in the church time zone, plus a join entry in enrollment history (EnrollmentTransaction).", "Saves each person. Database triggers then refresh the member and prospect counts, the leader, grade and main fellowship class.", "Writes one activity log entry for the whole call."],
        "params": [
            ["query", "query", "PeopleId, saved search name, or Search Builder code", True, None],
            ["orgId", "integer", "Involvement id", True, None]
        ],
        "returns": "None",
        "watch_out": ["Writes data.", "Raises an error if the involvement does not exist, as soon as the query matches anyone.", "Pending registrants who are already on the roster become full members, because Pending is cleared.", "It does not check registration limits, capacity, gender or age rules.", "If the query matches no one, nothing happens and no error is raised.", "Each person is saved separately, so a query that matches thousands of people is slow."],
        "example": "oid = q.QuerySqlInt(\"SELECT TOP 1 OrganizationId FROM Organizations WHERE OrganizationStatusId = 30 ORDER BY OrganizationId\")\nprint model.PeopleIds(\"PeopleId = {0}\".format(model.UserPeopleId))\n# Writes data. Uncomment to enroll the matches:\n# model.AddMembersToOrg(\"PeopleId = {0}\".format(model.UserPeopleId), oid)",
        "see_also": ["model.AddMemberToOrg", "model.JoinOrg", "model.DropOrgMember"]
    },
    "model.AddMemberToOrg": {
        "category": "Involvements",
        "kind": "method",
        "arity": [2, 2],
        "tags": ["writes_data"],
        "summary": "Adds one person to an involvement as Member, enrolled now.",
        "details": "Same as AddMembersToOrg with a single PeopleId.",
        "flow": ["Turns the PeopleId into a query and runs model.AddMembersToOrg on it.", "For the person found, inserts an OrganizationMembers row as Member (220), enrolled now, with an EnrollmentTransaction join row (type 1). Counts and leader fields refresh as with model.JoinOrg.", "If already enrolled, a Member, Prospect or blank type becomes Member, other types are kept, Pending is cleared, and no history row is written.", "Writes an ActivityLog line."],
        "params": [
            ["peopleId", "integer", "PeopleId", True, None],
            ["orgId", "integer", "Involvement id", True, None]
        ],
        "returns": "None",
        "watch_out": ["Writes data.", "If they are already enrolled with a leader type, the type is kept.", "The PeopleId goes through a people query, which skips deceased and archived people, so they are silently not added. model.JoinOrg does not have this limit.", "Calling it on a pending registrant clears Pending."],
        "example": "oid = q.QuerySqlInt(\"SELECT TOP 1 OrganizationId FROM Organizations WHERE OrganizationStatusId = 30 ORDER BY OrganizationId\")\nprint \"Already in it:\", model.InOrg(model.UserPeopleId, oid)\n# Writes data. Uncomment to add yourself:\n# model.AddMemberToOrg(model.UserPeopleId, oid)",
        "see_also": ["model.AddMembersToOrg", "model.JoinOrg", "model.InOrg", "model.DropOrgMember"]
    },
    "model.AddOrganization": {
        "category": "Involvements",
        "kind": "method",
        "arity": [1, 3],
        "tags": ["writes_data"],
        "summary": "Creates a new active involvement and returns its id.",
        "details": "Form 1 takes a program name and division name, creating either if missing. Form 2 copies type, campus, entry point and divisions from a template involvement, plus its schedules and settings when copySettings is True. Logs the change.",
        "flow": ["Form 1 finds the program by exact name or creates it, then finds a division of that name under that program or creates it and links it to the program.", "Form 1 inserts an Organizations row, Active (30), in that division with a DivOrg link. No type, campus, entry point or schedule is set.", "Form 2 loads the template involvement and copies its entry point, involvement type, campus, main division and all division links.", "Form 2 with copySettings True also copies the meeting schedules and the template's settings: leader type, location, check in, attendance, directory, involvement extra values, and registration settings including fees, limits and scripts.", "Writes an ActivityLog line and returns the new OrganizationId. No members are copied."],
        "params": [
            ["name", "string", "New involvement name", True, None],
            ["program", "string", "Program name (created if missing)", True, None],
            ["division", "string", "Division name (created if missing)", True, None]
        ],
        "returns": "Integer. New OrganizationId.",
        "watch_out": ["Writes data.", "Form 2 without a valid template creates the involvement in division 1 with no division links and no settings.", "Program and division names must match exactly, or new ones are created alongside the old.", "Form 2 with copySettings True adds the template's schedules twice. When the template has a schedule, the second save can fail with a duplicate key error after the involvement row already exists."],
        "example": "# Writes data. Uncomment to create an involvement:\n# newId = model.AddOrganization(\"API Test Group\", \"Adult Education\", \"Test\")\n# print \"Created\", newId\n# Or copy a template involvement and its settings:\n# newId = model.AddOrganization(\"API Test Copy\", 3638, True)\nprint \"Uncomment to create\"",
        "signatures": ["model.AddOrganization(name, program, division)", "model.AddOrganization(name, templateId=None, copySettings=True)"],
        "see_also": ["model.DeleteOrg", "model.GetOrganization", "model.OrganizationIds"]
    },
    "model.AddSubGroup": {
        "category": "Involvements",
        "kind": "method",
        "arity": [3, 3],
        "tags": ["writes_data"],
        "summary": "Puts an enrolled person into a sub group of the involvement, creating the sub group if it does not exist.",
        "flow": ["Finds the person's enrollment in the involvement and raises an error if there is none.", "Trims the sub group name and looks for a sub group (MemberTags) with that name in the involvement. If none exists, creates it.", "Adds the person to that sub group (OrgMemMemTags) if they are not already in it, and saves."],
        "params": [
            ["peopleId", "integer", "PeopleId", True, None],
            ["orgId", "integer", "Involvement id", True, None],
            ["group", "string", "Sub group name", True, None]
        ],
        "returns": "None",
        "watch_out": ["Writes data.", "Raises an error if the person is not enrolled.", "An empty sub group name does nothing and raises no error.", "The name is matched without regard to case, so \"leaders\" joins an existing \"Leaders\" sub group.", "Names are trimmed and cut to 200 characters.", "Adding someone who is already in the sub group does nothing."],
        "example": "oid = q.QuerySqlInt(\"SELECT TOP 1 OrganizationId FROM OrganizationMembers WHERE PeopleId = {0}\".format(model.UserPeopleId))\nprint \"In Team A:\", model.InSubGroup(model.UserPeopleId, oid, \"Team A\")\n# Writes data. Uncomment to add:\n# model.AddSubGroup(model.UserPeopleId, oid, \"Team A\")",
        "see_also": ["model.RemoveSubGroup", "model.InSubGroup", "model.AddSubGroupFromQuery"]
    },
    "model.AddSubGroupFromQuery": {
        "category": "Involvements",
        "kind": "method",
        "arity": [3, 3],
        "tags": ["writes_data"],
        "summary": "Puts everyone matched by the query who is enrolled in the involvement into the named sub group.",
        "details": "Creates the sub group if needed. Logs the change.",
        "flow": ["Writes an activity log entry.", "Runs the query to get the list of people.", "For each person, looks up their enrollment in the involvement and skips them if there is none.", "Adds each enrolled person to the sub group, creating the sub group the first time, and saves."],
        "params": [
            ["query", "query", "PeopleId, saved search name, or Search Builder code", True, None],
            ["orgId", "integer", "Involvement id", True, None],
            ["group", "string", "Sub group name", True, None]
        ],
        "returns": "None",
        "watch_out": ["Writes data.", "People not enrolled are skipped.", "The sub group name is trimmed and matched without regard to case. An empty name does nothing.", "Each person is looked up and saved separately, so large queries are slow."],
        "example": "oid = q.QuerySqlInt(\"SELECT TOP 1 OrganizationId FROM OrganizationMembers WHERE PeopleId = {0}\".format(model.UserPeopleId))\n# Writes data. Uncomment to add matches to a sub group:\n# model.AddSubGroupFromQuery(\"PeopleId = {0}\".format(model.UserPeopleId), oid, \"Team A\")\nprint \"Uncomment to run against\", oid",
        "see_also": ["model.AddSubGroup", "model.RemoveSubGroup", "model.InSubGroup"]
    },
    "model.AddTransaction": {
        "category": "Involvements",
        "kind": "method",
        "arity": [4, 8],
        "tags": ["writes_data"],
        "summary": "Records a payment against a person's involvement balance, as if added in the member dialog.",
        "details": "A positive amount is a payment; a negative amount increases the balance due. Your description is saved as the transaction message, and the transaction description is the involvement name (plus accounting code). The long form (reason, payment, description, amount, adjustFee, pmtDescription) is for advanced use.",
        "flow": ["Finds the person's OrganizationMembers row in that involvement. If there is none, it returns None and changes nothing.", "Finds the member's starting transaction. If the member has none on record, it reuses the latest transaction from the person's earlier registration in that involvement.", "Adds a new Transaction row and saves it at once. It is marked approved, the amount is your payment, your description is saved as its message, and its description is the involvement name with its accounting code in parentheses. The transaction id text reads \"Adjustment (your PeopleId)\".", "Links it to the starting transaction. A payment there lowers the amount due by that much, and a negative amount raises it. When the member had no transaction at all, this one becomes the starting transaction and the balance shows as a credit.", "Writes an activity log entry and returns the new transaction id."],
        "params": [
            ["peopleId", "integer", "PeopleId", True, None],
            ["orgId", "integer", "Involvement id", True, None],
            ["payment", "decimal", "Amount. Negative increases the balance due", True, None],
            ["description", "string", "Transaction description", True, None]
        ],
        "returns": "Integer transaction id, or None if the person is not in the involvement",
        "watch_out": ["Writes data.", "Returns None if the person is not enrolled.", "TouchPoint warns that long form transactions are hard to delete.", "It records a balance change only. No card is charged, no contribution or gift is created, and no receipt is emailed.", "It saves on its own right away, separate from any bundle you are building.", "For someone who never registered or paid, there is no fee to pay against, so the payment shows as a credit balance."],
        "example": "pid = model.UserPeopleId\noid = q.QuerySqlInt(\"SELECT TOP 1 OrganizationId FROM OrganizationMembers WHERE PeopleId = {0}\".format(pid))\nlink = model.GetPayLink(pid, oid) if oid else None\nprint \"Involvement:\", oid, \"Pay link:\", link or \"(none: no transaction on record)\"\n# Writes data. Uncomment to record a $10 payment:\n# print model.AddTransaction(pid, oid, 10.00, \"Cash at the door\")",
        "signatures": ["model.AddTransaction(peopleId, orgId, payment, description)", "model.AddTransaction(peopleId, orgId, reason, payment, description, amount=None, adjustFee=False, pmtDescription=None)"],
        "see_also": ["model.AdjustFee", "model.GetPayLink"]
    },
    "model.AdjustFee": {
        "category": "Involvements",
        "kind": "method",
        "arity": [4, 4],
        "tags": ["writes_data"],
        "summary": "Adjusts a person's fee for an involvement, as the Adjust Fee option in the member dialog does.",
        "details": "A positive amount lowers what they owe (such as a scholarship); a negative amount raises it.",
        "flow": ["Works like model.AddTransaction, but marks the new transaction as a fee adjustment and its id text reads \"AdjustFee (your PeopleId)\".", "Returns None and changes nothing if the person is not in the involvement.", "Saves the transaction at once and links it to the member's starting transaction.", "A fee adjustment changes the fee itself, not the amount paid. A positive amount lowers the fee and the balance due, and a negative amount raises them. Paid totals stay the same.", "Writes an activity log entry and returns the new transaction id."],
        "params": [
            ["peopleId", "integer", "PeopleId", True, None],
            ["orgId", "integer", "Involvement id", True, None],
            ["amount", "decimal", "Adjustment. Negative increases the balance due", True, None],
            ["description", "string", "Reason shown on the transaction", True, None]
        ],
        "returns": "Integer transaction id, or None if not enrolled",
        "watch_out": ["Writes data.", "Returns None if the person is not enrolled.", "For someone who never registered or paid, the adjustment becomes their starting transaction and the balance shows as a credit instead of a reduced fee."],
        "example": "oid = q.QuerySqlInt(\"SELECT TOP 1 OrganizationId FROM OrganizationMembers WHERE PeopleId = {0}\".format(model.UserPeopleId))\n# Writes data. Uncomment to apply a $25 scholarship:\n# print model.AdjustFee(model.UserPeopleId, oid, 25.00, \"Scholarship\")\nprint \"Uncomment to adjust the fee in\", oid",
        "see_also": ["model.AddTransaction", "model.GetPayLink"]
    },
    "model.DeleteOrg": {
        "category": "Involvements",
        "kind": "method",
        "arity": [3, 3],
        "tags": ["writes_data", "known_issue"],
        "summary": "Permanently purges every involvement with this exact name in the given program and division.",
        "details": "Members and history are purged too.",
        "flow": ["Checks for the Developer role. Without it, a denial is logged but the call keeps going.", "Finds the program by exact name. If none is found, it returns with no error.", "Finds the division with that name inside the program.", "Finds every involvement in that division with the exact name.", "Purges each one: members, meetings, attendance, enrollment history, time slots, posts and signature requests are deleted.", "Writes an activity log entry."],
        "params": [
            ["name", "string", "Exact involvement name", True, None],
            ["program", "string", "Program name", True, None],
            ["division", "string", "Division name", True, None]
        ],
        "returns": "None",
        "watch_out": ["Deletes data and cannot be undone.", "The Developer role check only logs a denial and continues, so do not rely on it as a safeguard.", "Raises an error if the division is not found.", "If the program is not found, nothing happens and no error is raised.", "It also raises an error if two divisions in that program share the name.", "An involvement that another involvement names in its registration settings is not deleted, and no error is raised. The same is true if the purge itself fails. The success log entry is still written.", "There is no way to tell from the return value whether anything was deleted; check with SQL afterward."],
        "example": "# Deletes data permanently. Uncomment only if you are sure:\n# model.DeleteOrg(\"API Test Group\", \"Adult Education\", \"Test\")\nprint q.QuerySqlInt(\"SELECT COUNT(*) FROM Organizations WHERE OrganizationName = 'API Test Group'\")",
        "see_also": ["model.AddOrganization", "model.OrganizationIds"]
    },
    "model.DropOrgMember": {
        "category": "Involvements",
        "kind": "method",
        "arity": [2, 2],
        "tags": ["writes_data"],
        "summary": "Drops a person from an involvement, as the Drop button does.",
        "details": "Records the drop transaction and removes the enrollment. Logs the change.",
        "flow": ["Writes an ActivityLog line, then loads the OrganizationMembers row. If there is none it raises an error.", "Inserts an EnrollmentTransaction drop row (type 5) dated now that keeps the member type, enrollment date, pending flag, notes, registration fields, attendance percent and the sub group names.", "Deletes the member's sub group links, member extra values, member documents and OPR registration answers, and marks their new form registrations as Dropped.", "Clears the involvement's main leader if it was this person, deletes the OrganizationMembers row, then recalculates the leader fields.", "Deletes their future attendance and volunteer commitments in that involvement, any sub requests on those, and today's absent rows. Mission trip support amounts tied to them in that trip are marked inactive."],
        "params": [
            ["peopleId", "integer", "PeopleId", True, None],
            ["orgId", "integer", "Involvement id", True, None]
        ],
        "returns": "None",
        "watch_out": ["Writes data.", "Raises an error if they are not enrolled.", "Documents uploaded on the member record are permanently deleted.", "Past attendance stays. Only meetings after now and today's unchecked rows are removed."],
        "example": "oid = q.QuerySqlInt(\"SELECT TOP 1 OrganizationId FROM OrganizationMembers WHERE PeopleId = {0}\".format(model.UserPeopleId))\nif oid and model.InOrg(model.UserPeopleId, oid):\n    print \"{0} is in\".format(model.UserPeopleId), oid\n    # Writes data. Uncomment to drop them:\n    # model.DropOrgMember(model.UserPeopleId, oid)",
        "see_also": ["model.AddMemberToOrg", "model.JoinOrg", "model.MoveToOrg", "model.InOrg"]
    },
    "model.GetOrganization": {
        "category": "Involvements",
        "kind": "method",
        "arity": [1, 1],
        "tags": ["read_only"],
        "summary": "Returns a small summary of an involvement: id, name, location and description.",
        "details": "Property names are lowercase. It is not the full Organization record; use SQL for other columns.",
        "params": [
            ["orgId", "integer", "Involvement id", True, None]
        ],
        "returns": "Object with id, name, location, description, or None",
        "watch_out": ["Returns None if not found.", "Use .name, .location and .description. Names such as .OrganizationName raise an error.", "The object also has extravalue1, extravalue2 and members, but they are always None from this call."],
        "example": "oid = q.QuerySqlInt(\"SELECT TOP 1 OrganizationId FROM Organizations WHERE OrganizationStatusId = 30 ORDER BY OrganizationId\")\norg = model.GetOrganization(oid)\nif org:\n    print org.id, org.name, org.location",
        "see_also": ["model.OrganizationIds", "q.QuerySqlTop1", "model.AddOrganization"]
    },
    "model.GetPayLink": {
        "category": "Involvements",
        "kind": "method",
        "arity": [2, 2],
        "tags": ["read_only"],
        "summary": "Returns the online payment link for a person's balance in an involvement.",
        "details": "The link does not sign the person in.",
        "flow": ["Finds the person's OrganizationMembers row in that involvement. Returns None if they are not in it.", "Returns None if the member has no transaction on record, such as someone added by hand who never registered or paid.", "Otherwise builds a link to /OnlineReg/PayAmtDue on your site with the member's transaction id encrypted in it.", "When opened, the link shows the online payment page for that balance. For a registration made with a new form it sends the person to that registration instead."],
        "params": [
            ["peopleId", "integer", "PeopleId", True, None],
            ["orgId", "integer", "Involvement id", True, None]
        ],
        "returns": "String URL, or None if the person is not in the involvement or has no transaction on record",
        "watch_out": ["Returns None if they are not enrolled.", "Also returns None when the member has no transaction on record, even if you expect a fee.", "A link comes back even when nothing is owed. Check the balance before you send it.", "None does not mean the balance is zero. It means the person is not in the involvement or has no transaction to pay against."],
        "example": "pid = model.UserPeopleId\noid = q.QuerySqlInt(\"SELECT TOP 1 OrganizationId FROM OrganizationMembers WHERE PeopleId = {0}\".format(pid))\nif oid:\n    link = model.GetPayLink(pid, oid)\n    print \"Pay link:\", link or \"(none: no transaction on record in involvement %d)\" % oid\nelse:\n    print \"You are not in any involvement, so there is no pay link to show.\"",
        "see_also": ["model.AddTransaction", "model.AdjustFee"]
    },
    "model.InOrg": {
        "category": "Involvements",
        "kind": "method",
        "arity": [2, 2],
        "tags": ["read_only"],
        "summary": "Returns True if the person is currently enrolled in the involvement.",
        "details": "Any member type counts, including pending and inactive.",
        "params": [
            ["peopleId", "integer", "PeopleId", True, None],
            ["orgId", "integer", "Involvement id", True, None]
        ],
        "returns": "Boolean",
        "watch_out": ["Returns False if either id is not a number."],
        "example": "oid = q.QuerySqlInt(\"SELECT TOP 1 OrganizationId FROM OrganizationMembers WHERE PeopleId = {0}\".format(model.UserPeopleId))\nprint \"In\", oid, \":\", model.InOrg(model.UserPeopleId, oid)",
        "see_also": ["model.InSubGroup", "model.JoinOrg", "model.DropOrgMember"]
    },
    "model.InSubGroup": {
        "category": "Involvements",
        "kind": "method",
        "arity": [3, 3],
        "tags": ["read_only"],
        "summary": "Returns True if the person is enrolled in the involvement and is in the named sub group.",
        "details": "The sub group name must match exactly.",
        "params": [
            ["peopleId", "integer", "PeopleId", True, None],
            ["orgId", "integer", "Involvement id", True, None],
            ["group", "string", "Sub group name (exact)", True, None]
        ],
        "returns": "Boolean",
        "watch_out": ["Here the name is case sensitive and is not trimmed, even though AddSubGroup and RemoveSubGroup ignore case. \"leaders\" returns False for a sub group named \"Leaders\".", "Returns False if the person is not enrolled; it does not raise."],
        "example": "oid = q.QuerySqlInt(\"SELECT TOP 1 OrganizationId FROM OrganizationMembers WHERE PeopleId = {0}\".format(model.UserPeopleId))\nprint model.InSubGroup(model.UserPeopleId, oid, \"Team A\")",
        "see_also": ["model.AddSubGroup", "model.RemoveSubGroup", "model.InOrg"]
    },
    "model.JoinOrg": {
        "category": "Involvements",
        "kind": "method",
        "arity": [2, 2],
        "tags": ["writes_data", "known_issue"],
        "summary": "Enrolls a person in an involvement as Member (type 220), enrolled now.",
        "details": "The person may be an integer PeopleId or a Person object from model.GetPerson. If already enrolled as Member or Prospect they become Member; other types are kept. Logs the change.",
        "flow": ["Accepts an integer PeopleId or a Person object. Anything else returns without doing anything. Writes an ActivityLog line.", "If the person is already in OrganizationMembers for that involvement, a Member, Prospect or blank type becomes Member (220), any other type is kept, and Pending is cleared. No history row is written and the enrollment date is unchanged.", "Otherwise inserts an OrganizationMembers row as Member, enrolled now, and an EnrollmentTransaction join row (type 1).", "The database then refreshes the involvement's member and prospect counts, the leader name when the type is the leader type, the person's last attended date, grade and main fellowship class."],
        "params": [
            ["orgId", "integer", "Involvement id", True, None],
            ["person", "any", "Integer PeopleId or Person object", True, None]
        ],
        "returns": "None",
        "watch_out": ["Writes data.", "Anything other than an integer or Person object (such as a string id) is silently ignored.", "Calling it on a pending registrant clears Pending and makes them a full member.", "No email, sub group or attendance is added.", "A missing involvement id raises an error."],
        "example": "oid = q.QuerySqlInt(\"SELECT TOP 1 OrganizationId FROM Organizations WHERE OrganizationStatusId = 30 ORDER BY OrganizationId\")\nprint \"Already in it:\", model.InOrg(model.UserPeopleId, oid)\n# Writes data. Uncomment to join yourself:\n# model.JoinOrg(oid, model.UserPeopleId)",
        "see_also": ["model.AddMemberToOrg", "model.AddMembersToOrg", "model.DropOrgMember", "model.InOrg"]
    },
    "model.MoveToOrg": {
        "category": "Involvements",
        "kind": "method",
        "arity": [3, 5],
        "tags": ["writes_data"],
        "summary": "Moves a person from one involvement to another.",
        "details": "Enrolls them in the target, keeping their member type unless toMemberTypeId is given. Copies registration data and sub groups when moveRegData is True, then drops them from the source. Logs the change.",
        "flow": ["Stops quietly if the two ids are the same or the person is not in the source involvement. Writes an ActivityLog line.", "Enrolls them in the target with a join EnrollmentTransaction dated now, keeping Pending. If they were already in the target, that existing row is reused and its member type is overwritten.", "Sets the member type (from toMemberTypeId, or copied from the source) and copies the notes field. For a pending member, a PromotingTo member extra value pointing at the source is repointed to the target.", "With moveRegData True, copies registration answers, request, amount, grade, register email, shirt size, tickets and transaction id, and adds them to the same sub groups. With False, a line naming the old involvement id is put at the top of the notes.", "In a main fellowship target, fills the person's grade level from the involvement when theirs is unknown, with a change log entry.", "Marks the new row as Moved, then drops them from the source exactly as model.DropOrgMember does."],
        "params": [
            ["peopleId", "integer", "PeopleId", True, None],
            ["fromOrgId", "integer", "Current involvement id", True, None],
            ["toOrgId", "integer", "Target involvement id", True, None],
            ["moveRegData", "boolean", "Copy registration answers and sub groups", False, "True"],
            ["toMemberTypeId", "integer", "Member type id in the target; -1 keeps the current type", False, "-1"]
        ],
        "returns": "None",
        "watch_out": ["Writes data.", "Does nothing if they are not in the source involvement or the ids are the same.", "Member documents and member extra values from the source are deleted by the drop, not moved.", "Future attendance and volunteer commitments in the source involvement are deleted.", "The target's member count, leader name and the person's last attended date are not refreshed right away, because the move skips that automatic update for the new row."],
        "example": "oid = q.QuerySqlInt(\"SELECT TOP 1 OrganizationId FROM OrganizationMembers WHERE PeopleId = {0}\".format(model.UserPeopleId))\nprint \"{0} is in\".format(model.UserPeopleId), oid\n# Writes data. Uncomment to move them to involvement 3638:\n# model.MoveToOrg(model.UserPeopleId, oid, 3638)",
        "see_also": ["model.JoinOrg", "model.DropOrgMember", "model.SetMemberType"]
    },
    "model.OnlineRegPerson": {
        "category": "Involvements",
        "kind": "method",
        "arity": [1, 1],
        "tags": ["read_only"],
        "summary": "Parses a classic registration's saved answers into an object.",
        "details": "Pass an OrganizationMember or the OnlineRegData XML string. Use .ExtraQuestion[setIndex]['Question'] to read answers.",
        "params": [
            ["regData", "any", "OrganizationMember object or OnlineRegData XML string", True, None]
        ],
        "returns": "OnlineRegPersonModel0 object or None",
        "watch_out": ["Returns None when there is no data.", "New style forms do not store answers here.", "Asking for a question name that is not in the answers raises an error. Check with .ContainsKey(name) first.", "Text box answers are in .Text[set], not .ExtraQuestion. Yes/No answers are in .YesNoQuestion, check boxes in .Checkbox, dropdowns in .option, radio buttons in .choice, menu items in .MenuItem.", ".ExtraQuestion and .Text always have at least one (possibly empty) set, but .YesNoQuestion, .Checkbox, .option, .choice and .MenuItem are None when there were no such answers.", "XML that is not well formed raises an error."],
        "example": "row = q.QuerySqlTop1(\"SELECT TOP 1 OnlineRegData FROM OrganizationMembers WHERE OnlineRegData IS NOT NULL ORDER BY EnrollmentDate DESC\")\nif row:\n    p = model.OnlineRegPerson(row.OnlineRegData)\n    if p:\n        print p.ExtraQuestion.Count, \"extra question sets\"",
        "signatures": ["model.OnlineRegPerson(orgMember)", "model.OnlineRegPerson(xml)"],
        "see_also": ["model.GetExtraQuestionsWithAnswers", "model.GetExtraQuestions"]
    },
    "model.OrganizationIds": {
        "category": "Involvements",
        "kind": "method",
        "arity": [2, 3],
        "tags": ["read_only"],
        "summary": "Returns the ids of involvements in a program and division.",
        "details": "Use 0 for progId or divId to mean any.",
        "params": [
            ["progId", "integer", "Program id, or 0 for any", True, None],
            ["divId", "integer", "Division id, or 0 for any", True, None],
            ["includeInactive", "boolean", "Include inactive involvements", False, "False"]
        ],
        "returns": "List of integers",
        "watch_out": ["Active involvements only unless includeInactive is True.", "The program filter goes through the involvement's divisions, so an involvement counts for a program when any of its divisions belongs to that program.", "Returns an empty list when nothing matches."],
        "example": "progId = q.QuerySqlInt(\"SELECT TOP 1 Id FROM Program ORDER BY Id\")\nids = model.OrganizationIds(progId, 0)\nprint len(ids), \"active involvements in program\", progId",
        "see_also": ["model.GetOrganization", "model.AddOrganization"]
    },
    "model.RemoveSubGroup": {
        "category": "Involvements",
        "kind": "method",
        "arity": [3, 3],
        "tags": ["writes_data"],
        "summary": "Removes an enrolled person from a sub group.",
        "details": "The sub group itself is kept.",
        "flow": ["Finds the person's enrollment in the involvement and raises an error if there is none.", "Finds the sub group by name, ignoring case, and removes the person from it if they are in it, then saves."],
        "params": [
            ["peopleId", "integer", "PeopleId", True, None],
            ["orgId", "integer", "Involvement id", True, None],
            ["group", "string", "Sub group name", True, None]
        ],
        "returns": "None",
        "watch_out": ["Writes data.", "Raises an error if the person is not enrolled.", "If the person is not in the sub group, or the sub group does not exist, nothing happens and no error is raised."],
        "example": "oid = q.QuerySqlInt(\"SELECT TOP 1 OrganizationId FROM OrganizationMembers WHERE PeopleId = {0}\".format(model.UserPeopleId))\nprint \"In Team A:\", model.InSubGroup(model.UserPeopleId, oid, \"Team A\")\n# Writes data. Uncomment to remove:\n# model.RemoveSubGroup(model.UserPeopleId, oid, \"Team A\")",
        "see_also": ["model.AddSubGroup", "model.InSubGroup", "model.AddSubGroupFromQuery"]
    },
    "model.SetMemberType": {
        "category": "Involvements",
        "kind": "method",
        "arity": [3, 3],
        "tags": ["writes_data"],
        "summary": "Changes an enrolled person's member type, given the type's description such as \"Leader\" or \"Member\".",
        "details": "Logs the change.",
        "flow": ["Writes an activity log entry.", "Finds the person's enrollment and raises an error if there is none.", "Looks up the member type by description, ignoring case. If none matches, creates a new member type with an id of 1010 or higher, a code of the first 20 characters, and Member attendance type.", "Sets the enrollment to that member type and saves."],
        "params": [
            ["peopleId", "integer", "PeopleId", True, None],
            ["orgId", "integer", "Involvement id", True, None],
            ["memberType", "string", "Member type description, exactly as in lookup.MemberType", True, None]
        ],
        "returns": "None",
        "watch_out": ["Writes data.", "If no member type has that exact description, a NEW member type is created, so check spelling.", "Raises an error if the person is not enrolled.", "Pass the description, not the id. Passing 220 creates a new member type named \"220\"."],
        "example": "oid = q.QuerySqlInt(\"SELECT TOP 1 OrganizationId FROM OrganizationMembers WHERE PeopleId = {0}\".format(model.UserPeopleId))\nprint q.QuerySqlStr(\"SELECT mt.Description FROM OrganizationMembers om JOIN lookup.MemberType mt ON mt.Id = om.MemberTypeId WHERE om.PeopleId = {0} AND om.OrganizationId = %d\".format(model.UserPeopleId) % oid)\n# Writes data. Uncomment to change it:\n# model.SetMemberType(model.UserPeopleId, oid, \"Leader\")",
        "see_also": ["model.MoveToOrg", "model.JoinOrg"]
    },
    "model.UpdateMainFellowship": {
        "category": "Involvements",
        "kind": "method",
        "arity": [1, 1],
        "tags": ["writes_data"],
        "summary": "Recalculates the main fellowship class of everyone in the involvement and refreshes its member and prospect counts.",
        "flow": ["Recalculates the main fellowship class (BibleFellowshipClassId) of every person enrolled in the involvement.", "Refreshes the involvement's member count and prospect count."],
        "params": [
            ["orgId", "integer", "Involvement id", True, None]
        ],
        "returns": "None",
        "watch_out": ["Writes data.", "Only people currently enrolled in this involvement are recalculated. Someone who was dropped keeps their old value until something else refreshes it."],
        "example": "oid = q.QuerySqlInt(\"SELECT TOP 1 OrganizationId FROM Organizations WHERE OrganizationStatusId = 30 ORDER BY OrganizationId\")\n# Writes data. Uncomment to refresh:\n# model.UpdateMainFellowship(oid)\nprint q.QuerySqlInt(\"SELECT MemberCount FROM Organizations WHERE OrganizationId = %d\" % oid)",
        "see_also": ["model.GetOrganization"]
    },
    "model.EditCommitment": {
        "category": "Meetings & Attendance",
        "kind": "method",
        "arity": [3, 3],
        "tags": ["writes_data"],
        "summary": "Sets a person's commitment for a meeting, as used by volunteer scheduling.",
        "details": "The text is one of Attending, Regrets, Find Sub, Sub Found, Substitute or Uncommitted. Case does not matter. The change is logged.",
        "flow": ["Matches the commitment text, ignoring case, and returns quietly if it is not one of the six choices.", "Finds the meeting and the person's attendance row for it.", "If there is no attendance row, creates one marked not present, with their member type in the involvement, or Visitor if they are not enrolled.", "Sets the commitment. Uncommitted clears it.", "Saves and writes an activity log entry."],
        "params": [
            ["meetingId", "integer", "MeetingId", True, None],
            ["peopleId", "integer", "PeopleId", True, None],
            ["commitment", "string", "Attending, Regrets, Find Sub, Sub Found, Substitute or Uncommitted", True, None]
        ],
        "returns": "None",
        "watch_out": ["Unrecognized text does nothing and raises no error.", "Raises an error if the meeting does not exist.", "It works for people who are not enrolled; they are added to the meeting as a Visitor with the commitment.", "\"Find Sub\" and \"Sub Found\" need the space.", "It does not send any email or notice."],
        "example": "mid = q.QuerySqlInt(\"SELECT TOP 1 MeetingId FROM Meetings WHERE MeetingDate < GETDATE() ORDER BY MeetingDate DESC\")\nprint \"Latest meeting:\", mid\n# Writes data. Uncomment to mark yourself attending:\n# model.EditCommitment(mid, model.UserPeopleId, \"Attending\")",
        "see_also": ["model.EditPersonAttendance", "model.GetMeeting"]
    },
    "model.EditPersonAttendance": {
        "category": "Meetings & Attendance",
        "kind": "method",
        "arity": [3, 3],
        "tags": ["writes_data"],
        "summary": "Marks a person present or absent for a meeting, as if done on the meeting page.",
        "details": "Records attendance the same way the meeting page does, including attendance types and meeting counts, and writes a line to the activity log.",
        "flow": ["Checks that the person and the meeting exist, and returns None if either is missing.", "When marking present, checks whether the person is already marked present in another involvement at the same date and time. If so, and neither involvement allows overlapping attendance, nothing is saved and it returns a message such as Jane Smith(123) already attended elsewhere.", "Adds the person's Attend row for the meeting, or updates the one already there, setting AttendanceFlag to 1 or 0. Marking absent keeps the row with AttendanceFlag 0 rather than deleting it.", "Sets the attendance type from the person's member type on the meeting date. People not in the involvement get New Visitor (60), or Recent Visitor (50) if they visited within the ResetNewVisitorDays setting (180 days by default). Someone visiting from a class they belong to that meets at the same hour is marked Visiting Member (40), and their own class gets an Other Class row.", "Updates the meeting counts and the person's attendance history, or queues that work when the AttendCountUpdatesOffline setting is on. Writes a RecordAttendance via Python API line to the activity log and returns ok."],
        "params": [
            ["meetingId", "integer", "MeetingId", True, None],
            ["peopleId", "integer", "PeopleId", True, None],
            ["attended", "boolean", "True for present, False for absent", True, None]
        ],
        "returns": "The string ok when saved, a message ending in already attended elsewhere when blocked, or None if the person or meeting does not exist",
        "watch_out": ["Returns None if the person or meeting does not exist.", "Check the return value. Anything other than ok means attendance was not saved.", "Marking a non-member absent still adds an Attend row for them as a visitor with AttendanceFlag 0."],
        "example": "mid = q.QuerySqlInt(\"SELECT TOP 1 MeetingId FROM Meetings WHERE MeetingDate < GETDATE() ORDER BY MeetingDate DESC\")\nprint \"Latest meeting:\", mid\n# Writes data. Uncomment to mark yourself present:\n# print model.EditPersonAttendance(mid, model.UserPeopleId, True)",
        "see_also": ["model.EditCommitment", "model.GetMeetingIdByDateTime", "model.GetMeeting"]
    },
    "model.GetMeeting": {
        "category": "Meetings & Attendance",
        "kind": "method",
        "arity": [1, 1],
        "tags": ["read_only"],
        "summary": "Returns the meeting with this id.",
        "params": [
            ["meetingId", "integer", "MeetingId", True, None]
        ],
        "returns": "Meeting object (MeetingId, OrganizationId, MeetingDate, MeetingEnd, MaxCount, HeadCount, NumPresent, Location, Description, DidNotMeet) or None",
        "watch_out": ["Returns None if there is no meeting with that id."],
        "example": "mid = q.QuerySqlInt(\"SELECT TOP 1 MeetingId FROM Meetings WHERE MeetingDate < GETDATE() ORDER BY MeetingDate DESC\")\nm = model.GetMeeting(mid)\nif m:\n    print m.MeetingId, m.OrganizationId, m.MeetingDate, m.MaxCount",
        "see_also": ["model.GetMeetingIdByDateTime", "model.GetMeetingsByDateTime", "model.ExtraValueMeeting"]
    },
    "model.GetMeetingIdByDateTime": {
        "category": "Meetings & Attendance",
        "kind": "method",
        "arity": [2, 4],
        "tags": ["writes_data"],
        "summary": "Returns the id of the involvement's meeting at a date and time, creating it by default if none is found.",
        "details": "The time is church local time and can be a datetime, a date string, or a Unix timestamp in seconds. By default it returns a meeting in progress at that time and creates one if none is found. Pass createIfNotExists=False to only look it up. With exactStart=True the meeting must start at that time. The lookup only sees meetings that have an end time; the create step then reuses a meeting that starts at exactly that time.",
        "flow": ["Turns meetingDate into a date and time. A datetime is used as is. Anything that reads as a whole number is treated as Unix seconds and converted to the church time zone (the LocalTimeZone setting, Central if unset). Anything else is parsed as a date string, and an unreadable value returns None.", "Looks for a meeting of that involvement that is in progress at that moment: it started at or before the time and its end time is after it. With exactStart=True the start must match within one second. Either way the meeting needs an end time to be found here.", "If nothing matched and createIfNotExists is True, it looks for a meeting that starts at exactly that time and returns it.", "If there is still no meeting it creates one. Location comes from the involvement. End time, attendance credit and capacity come from a schedule on the same weekday and time, if there is one. The auto absents flag follows the church setting, and the meeting is open.", "Returns the MeetingId, or None."],
        "params": [
            ["orgId", "integer", "Involvement id", True, None],
            ["meetingDate", "date", "datetime, date string, or Unix timestamp (seconds)", True, None],
            ["createIfNotExists", "boolean", "Create the meeting if missing", False, "True"],
            ["exactStart", "boolean", "Require an exact start time match", False, "False"]
        ],
        "returns": "Integer MeetingId, or None if not found and not created",
        "watch_out": ["By default it creates the meeting if none is found.", "Returns None if not found and not created.", "Meetings with no end time are never found by the in progress lookup. Most meetings have no end time unless the schedule sets one, so with createIfNotExists=False it usually returns None even when a meeting starts at exactly that time.", "With the default createIfNotExists=True, a time in the middle of a meeting that has no end time creates a second meeting at that time instead of returning the first one. Pass the exact start time.", "A date string made only of digits, such as 20261102, is read as Unix seconds and lands in 1970.", "A date with no time means midnight."],
        "example": "import datetime\noid = q.QuerySqlInt(\"SELECT TOP 1 OrganizationId FROM Organizations WHERE OrganizationStatusId = 30 ORDER BY OrganizationId\")\nmid = model.GetMeetingIdByDateTime(oid, datetime.datetime.now(), False)\nprint \"Meeting now in\", oid, \"=\", mid",
        "see_also": ["model.GetMeetingIdsByDateTime", "model.GetMeeting", "model.EditPersonAttendance", "model.GetMeetingsByDateTime"]
    },
    "model.GetMeetingIdsByDateTime": {
        "category": "Meetings & Attendance",
        "kind": "method",
        "arity": [1, 3],
        "tags": ["read_only"],
        "summary": "Returns the ids of meetings in progress at a moment, for one involvement or all of them.",
        "details": "The moment is church local time.",
        "flow": ["Reads the date. A datetime is used as is, a number is a Unix timestamp converted to church time, anything else is parsed as a date string.", "Looks in the Meetings table for meetings that started before the moment and end after it, allowing one second of slack either way.", "With exactStart=True the meeting must instead start within one second of the moment, and it must still end after it.", "Returns the matching MeetingIds."],
        "params": [
            ["meetingDate", "date", "datetime, date string, or Unix timestamp", True, None],
            ["orgId", "integer", "Involvement id; None for all involvements", False, "None"],
            ["exactStart", "boolean", "Only meetings starting at exactly that time", False, "False"]
        ],
        "returns": "List of integer MeetingIds, or None if the date cannot be read",
        "watch_out": ["Only meetings with a MeetingEnd set can match, because the end must be after the given time.", "Returns None if the date cannot be read.", "A value made only of digits, including a digit string such as \"20261102\", is read as a Unix timestamp in seconds and converted from UTC to the church time zone (the LocalTimeZone setting, Central if unset).", "The result is a query you loop over, not a Python list. Wrap it in list() before using len() or indexing.", "exactStart=True still requires a MeetingEnd after the moment, so meetings with no end time never match.", "Passing None as the date raises an error instead of returning None."],
        "example": "import datetime\nids = list(model.GetMeetingIdsByDateTime(datetime.datetime.now()))\nprint len(ids), \"meetings in progress:\", ids",
        "see_also": ["model.GetMeetingsByDateTime", "model.GetMeetingIdByDateTime", "model.UpdateMeetingEnd"]
    },
    "model.GetMeetingsByDateTime": {
        "category": "Meetings & Attendance",
        "kind": "method",
        "arity": [1, 3],
        "tags": ["read_only"],
        "summary": "Returns meetings in progress at a moment for one involvement, a list of involvements, or all involvements.",
        "details": "The moment is church local time.",
        "flow": ["Reads the date the same way as GetMeetingIdsByDateTime: datetime, Unix timestamp, or date string.", "Limits to the involvement or list of involvements you pass, or all involvements when orgIds is None.", "Keeps meetings that start before the moment and end after it, with one second of slack. With exactStart=True the start must match within one second.", "Returns the Meeting objects."],
        "params": [
            ["meetingDate", "date", "datetime, date string, or Unix timestamp", True, None],
            ["orgIds", "list", "One involvement id or a list of ids; None for all", False, "None"],
            ["exactStart", "boolean", "Only meetings starting at exactly that time", False, "False"]
        ],
        "returns": "List of Meeting objects, or None if the date cannot be read",
        "watch_out": ["Only meetings with a MeetingEnd set can match.", "Returns None if the date cannot be read.", "A value made only of digits, including a digit string such as \"20261102\", is read as a Unix timestamp in seconds and converted from UTC to the church time zone (the LocalTimeZone setting, Central if unset).", "The result is a query you loop over, not a Python list. Wrap it in list() before using len() or indexing.", "exactStart=True still requires a MeetingEnd after the moment."],
        "example": "import datetime\nmeetings = list(model.GetMeetingsByDateTime(datetime.datetime.now()))\n# Only meetings that have an end time can match\nprint len(meetings), \"meetings in progress now\"\nfor m in meetings:\n    print m.MeetingId, m.OrganizationId, m.MeetingDate, m.MeetingEnd",
        "signatures": ["model.GetMeetingsByDateTime(meetingDate, orgId, exactStart=False)", "model.GetMeetingsByDateTime(meetingDate, orgIds=None, exactStart=False)"],
        "see_also": ["model.GetMeetingIdsByDateTime", "model.GetMeetingsByDateTimeRange", "model.GetMeeting"]
    },
    "model.GetMeetingsByDateTimeRange": {
        "category": "Meetings & Attendance",
        "kind": "method",
        "arity": [2, 4],
        "tags": ["read_only"],
        "summary": "Returns meetings that overlap a time range for one involvement, a list of involvements, or all of them.",
        "details": "With exactStartEnd=True only meetings that start and end exactly at the range ends match.",
        "flow": ["Reads both dates as a datetime, Unix timestamp, or date string.", "Limits to the involvement or list of involvements you pass, or all of them when orgIds is None.", "Without exactStartEnd, keeps meetings that start before rangeEnd and end after rangeStart, with one second of slack.", "With exactStartEnd=True, keeps only meetings whose start and end each match within one second."],
        "params": [
            ["rangeStart", "date", "Start: datetime, date string, or Unix timestamp", True, None],
            ["rangeEnd", "date", "End: datetime, date string, or Unix timestamp", True, None],
            ["orgIds", "list", "One involvement id or a list of ids; None for all", False, "None"],
            ["exactStartEnd", "boolean", "Require exact start and end", False, "False"]
        ],
        "returns": "List of Meeting objects, or None if a date cannot be read",
        "watch_out": ["Meetings with no MeetingEnd never match.", "Returns None if a date cannot be read.", "A value made only of digits, including a digit string such as \"20261102\", is read as a Unix timestamp in seconds and converted from UTC to the church time zone (the LocalTimeZone setting, Central if unset).", "The result is a query you loop over, not a Python list. Wrap it in list() before using len() or indexing."],
        "example": "import datetime\nend = datetime.datetime.now()\nstart = end + datetime.timedelta(days=-7)\nmeetings = list(model.GetMeetingsByDateTimeRange(start, end))\nprint len(meetings), \"meetings with an end time in the last 7 days\"",
        "signatures": ["model.GetMeetingsByDateTimeRange(rangeStart, rangeEnd, orgId, exactStartEnd=False)", "model.GetMeetingsByDateTimeRange(rangeStart, rangeEnd, orgIds=None, exactStartEnd=False)"],
        "see_also": ["model.GetMeetingsByDateTime", "model.GetMeetingIdsByDateTime", "model.UpdateMeetingDuration"]
    },
    "model.GetNextScheduledMeetingDate": {
        "category": "Meetings & Attendance",
        "kind": "method",
        "arity": [1, 2],
        "tags": ["read_only"],
        "summary": "Returns the next date and time the involvement is scheduled to meet after the reference date.",
        "details": "The reference date defaults to now. It is based on the involvement's weekly schedules and first and last meeting dates.",
        "flow": ["Uses the reference date, or now in church time when none is given.", "Reads the involvement's weekly schedules (OrgSchedule) and their stored next meeting dates, skipping the involvement if its last meeting date has passed or its first meeting date is over a year away.", "Steps stale schedule dates forward a week at a time until one is after the reference date, dropping schedules that pass the last meeting date.", "Returns that date, or None if it is a year or more away."],
        "params": [
            ["orgId", "integer", "Involvement id", True, None],
            ["referenceDate", "date", "Find the next meeting after this", False, "now"]
        ],
        "returns": "Date or None",
        "watch_out": ["Returns None if nothing is scheduled within a year.", "Returns None when the involvement has no weekly schedule.", "When an involvement has more than one weekly schedule and the earliest stored date is already past, it can return a date a week later than another schedule's sooner meeting.", "Unix timestamps are not accepted for referenceDate; pass a datetime or date string.", "It does not look at the Meetings table, so meetings added by hand or cancelled meetings are not considered."],
        "example": "oid = q.QuerySqlInt(\"SELECT TOP 1 OrganizationId FROM OrgSchedule ORDER BY OrganizationId DESC\")\nprint oid, \"next meets\", model.GetNextScheduledMeetingDate(oid)",
        "see_also": ["model.GetMeetingIdByDateTime", "model.GetOrganization"]
    },
    "model.MeetingDidNotMeet": {
        "category": "Meetings & Attendance",
        "kind": "method",
        "arity": [1, 2],
        "tags": ["writes_data"],
        "summary": "Gets or sets the Did Not Meet flag on a meeting.",
        "details": "Call with only meetingId to read it. A new value is saved; only setting it to True is written to the activity log.",
        "flow": ["Looks up the meeting. Returns None if it is not found.", "If you pass a value that differs from the current one, it updates the meeting. True sets Did Not Meet and clears the attendance credit. False clears Did Not Meet and sets the attendance credit to Every Meeting.", "Saves the change. Only setting True writes an entry to the activity log.", "Returns the flag's value after any change."],
        "params": [
            ["meetingId", "integer", "MeetingId", True, None],
            ["didNotMeet", "boolean", "New value; omit to only read", False, "None"]
        ],
        "returns": "Boolean (the current value) or None if the meeting is not found",
        "watch_out": ["Passing True clears the attendance credit. Passing False sets the attendance credit to Every Meeting, not to whatever it was before.", "Returns None if the meeting is not found.", "Passing the value the meeting already has does nothing and saves nothing."],
        "example": "mid = q.QuerySqlInt(\"SELECT TOP 1 MeetingId FROM Meetings WHERE MeetingDate < GETDATE() ORDER BY MeetingDate DESC\")\nprint \"Did not meet:\", model.MeetingDidNotMeet(mid)\n# Writes data. Uncomment to set it:\n# model.MeetingDidNotMeet(mid, True)",
        "see_also": ["model.MeetingIsAllDay", "model.MeetingShowAsBusy", "model.GetMeeting"]
    },
    "model.MeetingIsAllDay": {
        "category": "Meetings & Attendance",
        "kind": "method",
        "arity": [1, 2],
        "tags": ["writes_data"],
        "summary": "Gets or sets the All Day flag on a meeting, used by calendars.",
        "details": "Call with only meetingId to read it. Setting a new value saves and logs it.",
        "params": [
            ["meetingId", "integer", "MeetingId", True, None],
            ["isAllDay", "boolean", "New value; omit to only read", False, "None"]
        ],
        "returns": "Boolean (the current value) or None if the meeting is not found",
        "watch_out": ["Returns None if the meeting is not found.", "Passing the value the meeting already has does nothing and writes no log entry."],
        "example": "mid = q.QuerySqlInt(\"SELECT TOP 1 MeetingId FROM Meetings WHERE MeetingDate < GETDATE() ORDER BY MeetingDate DESC\")\nprint \"All day:\", model.MeetingIsAllDay(mid)\n# Writes data. Uncomment to set it:\n# model.MeetingIsAllDay(mid, True)",
        "see_also": ["model.MeetingShowAsBusy", "model.MeetingDidNotMeet", "model.UpdateMeetingDuration"]
    },
    "model.MeetingShowAsBusy": {
        "category": "Meetings & Attendance",
        "kind": "method",
        "arity": [1, 2],
        "tags": ["writes_data"],
        "summary": "Gets or sets the Show As Busy flag on a meeting, used by calendars.",
        "details": "Call with only meetingId to read it. Setting a new value saves and logs it.",
        "params": [
            ["meetingId", "integer", "MeetingId", True, None],
            ["showAsBusy", "boolean", "New value; omit to only read", False, "None"]
        ],
        "returns": "Boolean (the current value) or None if the meeting is not found",
        "watch_out": ["Returns None if the meeting is not found.", "Passing the value the meeting already has does nothing and writes no log entry."],
        "example": "mid = q.QuerySqlInt(\"SELECT TOP 1 MeetingId FROM Meetings WHERE MeetingDate < GETDATE() ORDER BY MeetingDate DESC\")\nprint \"Show as busy:\", model.MeetingShowAsBusy(mid)\n# Writes data. Uncomment to set it:\n# model.MeetingShowAsBusy(mid, True)",
        "see_also": ["model.MeetingIsAllDay", "model.MeetingDidNotMeet"]
    },
    "model.SendAttendanceReminders": {
        "category": "Meetings & Attendance",
        "kind": "method",
        "arity": [1, 1],
        "tags": ["sends_message"],
        "summary": "Sends attendance links by text or email to the leaders of meetings starting in the 15 minutes after the given time.",
        "details": "Only involvements with Send Attendance Link turned on are included.",
        "flow": ["Raises an error if the date cannot be read, then stops quietly unless the AttendanceReminder setting is on.", "Finds meetings starting after the time and up to 15 minutes later, from both the Meetings table and the weekly schedules of active involvements. Not Weekly involvements are skipped unless the AttendanceReminderForceScheduleForNotWeekly setting is on.", "Skips involvements that do not have Send Attendance Link turned on.", "Leaders with a cell phone who accept texts get a text from the system texting group, using the AttendanceReminderSmsMessage setting or a default message, when texting is configured.", "Every other leader gets a transactional email from the admin email address, using the AttendanceReminderEmail email template in Special Content or a built in default."],
        "params": [
            ["dt", "date", "Start of the 15 minute window: datetime or date string", True, None]
        ],
        "returns": "None",
        "watch_out": ["Sends texts and emails.", "Does nothing unless the AttendanceReminder setting is on.", "Raises an error if the date cannot be read.", "It keeps no record of what it sent, so calling it twice for the same window sends the reminders twice.", "The window excludes a meeting starting exactly at the given time and includes one starting exactly 15 minutes later.", "Leaders whose email is turned off and who cannot get texts receive nothing, with no error."],
        "example": "print \"AttendanceReminder setting:\", model.Setting(\"AttendanceReminder\")\n# Sends texts and emails. Uncomment to run for the next 15 minutes:\n# import datetime\n# model.SendAttendanceReminders(datetime.datetime.now())",
        "see_also": ["model.Setting", "model.EmailReminders", "model.SendSms"]
    },
    "model.UpdateMeetingDate": {
        "category": "Meetings & Attendance",
        "kind": "method",
        "arity": [2, 2],
        "tags": ["writes_data"],
        "summary": "Changes a meeting's start date and time and saves it.",
        "details": "Accepts a datetime, a date string, or a Unix timestamp. Returns the new start.",
        "params": [
            ["meetingId", "integer", "MeetingId", True, None],
            ["dateTimeStart", "date", "New start: datetime, date string, or Unix timestamp", True, None]
        ],
        "returns": "Date (the new start) or None",
        "watch_out": ["Returns None if the meeting is not found or the date cannot be read.", "A value made only of digits, including a digit string such as \"20261102\", is read as a Unix timestamp in seconds and converted from UTC to the church time zone (the LocalTimeZone setting, Central if unset).", "The meeting end time is not moved, so the end can end up before the new start. Use UpdateMeetingDuration or UpdateMeetingEnd afterward.", "The change is not written to the activity log."],
        "example": "mid = q.QuerySqlInt(\"SELECT TOP 1 MeetingId FROM Meetings WHERE MeetingDate < GETDATE() ORDER BY MeetingDate DESC\")\nprint \"Current start:\", model.GetMeeting(mid).MeetingDate\n# Writes data. Uncomment to move it:\n# import datetime\n# print model.UpdateMeetingDate(mid, datetime.datetime(2026, 3, 1, 9, 0))",
        "see_also": ["model.UpdateMeetingEnd", "model.UpdateMeetingDuration", "model.GetMeeting"]
    },
    "model.UpdateMeetingDuration": {
        "category": "Meetings & Attendance",
        "kind": "method",
        "arity": [2, 2],
        "tags": ["writes_data"],
        "summary": "Sets a meeting's end time to its start time plus the given minutes, and saves.",
        "params": [
            ["meetingId", "integer", "MeetingId", True, None],
            ["minutes", "integer", "Length of the meeting in minutes", True, None]
        ],
        "returns": "None",
        "watch_out": ["Negative values are treated as 0.", "Does nothing if the meeting is not found.", "The change is not written to the activity log."],
        "example": "mid = q.QuerySqlInt(\"SELECT TOP 1 MeetingId FROM Meetings WHERE MeetingDate < GETDATE() ORDER BY MeetingDate DESC\")\nm = model.GetMeeting(mid)\nprint m.MeetingDate, \"to\", m.MeetingEnd\n# Writes data. Uncomment to make it 90 minutes long:\n# model.UpdateMeetingDuration(mid, 90)",
        "see_also": ["model.UpdateMeetingEnd", "model.UpdateMeetingDate"]
    },
    "model.UpdateMeetingEnd": {
        "category": "Meetings & Attendance",
        "kind": "method",
        "arity": [2, 2],
        "tags": ["writes_data"],
        "summary": "Changes a meeting's end date and time and saves it.",
        "details": "Accepts a datetime, a date string, or a Unix timestamp. Returns the new end.",
        "params": [
            ["meetingId", "integer", "MeetingId", True, None],
            ["dateTimeEnd", "date", "New end: datetime, date string, or Unix timestamp", True, None]
        ],
        "returns": "Date (the new end) or None",
        "watch_out": ["Returns None if the meeting is not found or the date cannot be read.", "A value made only of digits, including a digit string such as \"20261102\", is read as a Unix timestamp in seconds and converted from UTC to the church time zone (the LocalTimeZone setting, Central if unset).", "It does not check that the end is after the start.", "The change is not written to the activity log."],
        "example": "mid = q.QuerySqlInt(\"SELECT TOP 1 MeetingId FROM Meetings WHERE MeetingDate < GETDATE() ORDER BY MeetingDate DESC\")\nprint \"Current end:\", model.GetMeeting(mid).MeetingEnd\n# Writes data. Uncomment to change it:\n# import datetime\n# print model.UpdateMeetingEnd(mid, datetime.datetime(2026, 3, 1, 10, 30))",
        "see_also": ["model.UpdateMeetingDate", "model.UpdateMeetingDuration", "model.GetMeetingsByDateTime"]
    },
    "q.AttendanceTypeCountDateRange": {
        "category": "Meetings & Attendance",
        "kind": "method",
        "arity": [6, 6],
        "tags": ["read_only"],
        "summary": "Counts the unique people marked present with the given attendance types during the days starting at startDate.",
        "details": "Types can be ids (\"50,60\"), descriptions (\"New Guest, Recent Guest\"), or \"All\". Use 0 for any program, division or involvement.",
        "params": [
            ["progId", "integer", "Program id, or 0", True, None],
            ["divId", "integer", "Division id, or 0", True, None],
            ["orgId", "integer", "Involvement id, or 0", True, None],
            ["attendTypes", "string", "Attendance type ids or descriptions, comma separated, or \"All\"", True, None],
            ["startDate", "date", "First day", True, None],
            ["days", "integer", "Number of days to include", True, None]
        ],
        "returns": "Integer count of people",
        "watch_out": ["Raises an error if startDate cannot be read.", "A description that does not match any attendance type is skipped, so a misspelled list returns 0 with no error.", "\"All\" must be spelled with a capital A. Ids and descriptions cannot be mixed in one list.", "A program or division matches any division the involvement belongs to, not only its main one."],
        "example": "import datetime\nstart = datetime.datetime.now() + datetime.timedelta(days=-30)\nprint q.AttendanceTypeCountDateRange(0, 0, 0, \"New Guest, Recent Guest\", start, 30), \"guests in the last 30 days\"",
        "see_also": ["q.AttendCountAsOf", "q.AttendMemberTypeCountAsOf"]
    },
    "q.AttendCountAsOf": {
        "category": "Meetings & Attendance",
        "kind": "method",
        "arity": [6, 6],
        "tags": ["read_only"],
        "summary": "Counts the unique people marked present between two dates.",
        "details": "A date with no time includes that whole day. Set guestOnly to True to count only Recent Guest and New Guest attendance. Use 0 for any program, division or involvement.",
        "params": [
            ["startDate", "date", "Start date", True, None],
            ["endDate", "date", "End date (whole day included when no time is given)", True, None],
            ["guestOnly", "boolean", "Only attendance types 50 and 60", True, None],
            ["progId", "integer", "Program id, or 0", True, None],
            ["divId", "integer", "Division id, or 0", True, None],
            ["orgId", "integer", "Involvement id, or 0", True, None]
        ],
        "returns": "Integer count of people",
        "watch_out": ["Raises an error if either date cannot be read.", "An end date at exactly midnight, even a datetime you meant as midnight, is moved to the end of that day.", "Guest only means attendance types 50 and 60 exactly, whatever your church calls them."],
        "example": "import datetime\nend = datetime.date.today()\nstart = end + datetime.timedelta(days=-30)\nprint q.AttendCountAsOf(start, end, False, 0, 0, 0), \"people attended in the last 30 days\"",
        "see_also": ["q.AttendanceTypeCountDateRange", "q.AttendMemberTypeCountAsOf", "q.NumPresentDateRange"]
    },
    "q.AttendMemberTypeCountAsOf": {
        "category": "Meetings & Attendance",
        "kind": "method",
        "arity": [7, 7],
        "tags": ["read_only"],
        "summary": "Counts attendance records between two dates for people whose member type at the time matches.",
        "details": "Includes types in memberTypes and leaves out types in notMemberTypes. Types can be ids or descriptions, comma separated; an empty memberTypes means all. A date with no time includes that whole day.",
        "params": [
            ["startDate", "date", "Start date", True, None],
            ["endDate", "date", "End date", True, None],
            ["memberTypes", "string", "Member type ids or descriptions to include; \"\" for all", True, None],
            ["notMemberTypes", "string", "Member type ids or descriptions to exclude; \"\" for none", True, None],
            ["progId", "integer", "Program id, or 0", True, None],
            ["divId", "integer", "Division id, or 0", True, None],
            ["orgId", "integer", "Involvement id, or 0", True, None]
        ],
        "returns": "Integer count of attendance records",
        "watch_out": ["It counts attendance records, not unique people.", "Raises an error if either date cannot be read.", "A member type description that does not match anything leaves the include list empty, which counts every member type. A typo silently counts everyone.", "Ids and descriptions cannot be mixed in one list. \"All\" also works for memberTypes.", "Only records marked present are counted."],
        "example": "import datetime\nend = datetime.date.today()\nstart = end + datetime.timedelta(days=-7)\nprint q.AttendMemberTypeCountAsOf(start, end, \"Member\", \"\", 0, 0, 0), \"member attendances this week\"",
        "see_also": ["q.AttendCountAsOf", "q.AttendanceTypeCountDateRange"]
    },
    "q.LastSunday": {
        "category": "Meetings & Attendance",
        "kind": "property",
        "tags": ["read_only"],
        "summary": "The most recent past Sunday that had a meeting with attendance.",
        "details": "If none is found, it is the most recent Sunday on the calendar. Read it without parentheses.",
        "returns": "Date",
        "watch_out": ["A meeting counts when its headcount or present count is above zero.", "If today is Sunday and no meeting today has attendance yet, the fallback is today, not last week.", "The value is worked out once and reused for the rest of the script."],
        "example": "print \"Last Sunday with attendance:\", q.LastSunday",
        "see_also": ["q.LastWeekAttendance", "model.MostRecentAttendedSunday"]
    },
    "q.LastWeekAttendance": {
        "category": "Meetings & Attendance",
        "kind": "method",
        "arity": [4, 4],
        "tags": ["read_only"],
        "summary": "Sums attendance for meetings in a window measured in hours from midnight at the start of q.LastSunday.",
        "details": "Each meeting counts the larger of headcount or present. For example 0 and 24 is that Sunday; -96 and 72 runs from the prior Wednesday through Wednesday. Use 0 for any program or division.",
        "params": [
            ["progId", "integer", "Program id, or 0", True, None],
            ["divId", "integer", "Division id, or 0", True, None],
            ["startHour", "integer", "Hours from LastSunday midnight where the window starts (can be negative)", True, None],
            ["endHour", "integer", "Hours from LastSunday midnight where the window ends", True, None]
        ],
        "returns": "Integer total attendance",
        "watch_out": ["Meetings marked Did Not Meet are not skipped.", "An involvement linked to more than one division, or a division linked to more than one program, is counted once per link, so totals with progId and divId of 0 can be too high.", "There is no involvement parameter."],
        "example": "print \"Sunday\", q.LastSunday, \"attendance:\", q.LastWeekAttendance(0, 0, 0, 24)",
        "see_also": ["q.LastSunday", "q.NumPresent", "q.NumPresentDateRange"]
    },
    "q.MeetingCount": {
        "category": "Meetings & Attendance",
        "kind": "method",
        "arity": [4, 4],
        "tags": ["read_only"],
        "summary": "Counts meetings in the last N days that were not marked Did Not Meet.",
        "details": "Use 0 for any program, division or involvement.",
        "params": [
            ["days", "integer", "Days back from now", True, None],
            ["progId", "integer", "Program id, or 0", True, None],
            ["divId", "integer", "Division id, or 0", True, None],
            ["orgId", "integer", "Involvement id, or 0", True, None]
        ],
        "returns": "Integer",
        "watch_out": ["There is no upper date limit, so meetings already created for future dates are counted.", "Days are counted back from the current time of day, not from midnight.", "A program or division matches any division the involvement belongs to."],
        "example": "print q.MeetingCount(7, 0, 0, 0), \"meetings in the last 7 days\"",
        "see_also": ["q.MeetingCountDateHours", "q.NumPresent"]
    },
    "q.MeetingCountDateHours": {
        "category": "Meetings & Attendance",
        "kind": "method",
        "arity": [5, 5],
        "tags": ["read_only", "known_issue"],
        "summary": "Meant to count meetings (not Did Not Meet) in a window of hours from a start date.",
        "details": "Use 0 for any program, division or involvement.",
        "params": [
            ["progId", "integer", "Program id, or 0", True, None],
            ["divId", "integer", "Division id, or 0", True, None],
            ["orgId", "integer", "Involvement id, or 0", True, None],
            ["startDate", "date", "Start date", True, None],
            ["hours", "integer", "Window length in hours (currently ignored)", True, None]
        ],
        "returns": "Integer",
        "watch_out": ["The hours argument is currently ignored.", "It counts every meeting on or after startDate, including future ones.", "Raises an error if startDate cannot be read."],
        "example": "print q.MeetingCountDateHours(0, 0, 0, q.LastSunday, 24), \"meetings since last Sunday\"",
        "see_also": ["q.MeetingCount", "q.LastSunday"]
    },
    "q.NumPresent": {
        "category": "Meetings & Attendance",
        "kind": "method",
        "arity": [4, 4],
        "tags": ["read_only"],
        "summary": "Sums attendance for meetings in the last N days that were not marked Did Not Meet.",
        "details": "Each meeting counts the larger of its headcount or people marked present. Use 0 for any program, division or involvement.",
        "params": [
            ["days", "integer", "Days back from now", True, None],
            ["progId", "integer", "Program id, or 0", True, None],
            ["divId", "integer", "Division id, or 0", True, None],
            ["orgId", "integer", "Involvement id, or 0", True, None]
        ],
        "returns": "Integer",
        "watch_out": ["Days are counted back from the current time of day, not from midnight.", "There is no upper date limit, so future meetings are included, though they usually add 0.", "Returns 0 when no meetings match."],
        "example": "print q.NumPresent(7, 0, 0, 0), \"total attendance in the last 7 days\"",
        "see_also": ["q.NumPresentDateRange", "q.MeetingCount", "q.LastWeekAttendance"]
    },
    "q.NumPresentDateRange": {
        "category": "Meetings & Attendance",
        "kind": "method",
        "arity": [5, 5],
        "tags": ["read_only"],
        "summary": "Sums attendance for meetings from startDate for the given number of days, skipping Did Not Meet.",
        "details": "Each meeting counts the larger of its headcount or people marked present. Use 0 for any program, division or involvement.",
        "params": [
            ["progId", "integer", "Program id, or 0", True, None],
            ["divId", "integer", "Division id, or 0", True, None],
            ["orgId", "integer", "Involvement id, or 0", True, None],
            ["startDate", "date", "First day", True, None],
            ["days", "integer", "Number of days", True, None]
        ],
        "returns": "Integer",
        "watch_out": ["Raises an error if startDate cannot be read.", "A date string with no time starts at midnight, so days=1 covers exactly that day."],
        "example": "print q.NumPresentDateRange(0, 0, 0, q.LastSunday, 1), \"attendance on\", q.LastSunday",
        "see_also": ["q.NumPresent", "q.AttendCountAsOf", "q.LastSunday"]
    },
    "model.BuildDisplay": {
        "category": "Page & Script",
        "kind": "method",
        "arity": [2, 5],
        "tags": ["read_only"],
        "summary": "Builds a read only HTML box with a two column Name and Value table from a DynamicData object.",
        "details": "The box is titled with the name split into words. The edit HTML goes below the table. If data is None, the add HTML is shown in place of the rows. Metadata can mark fields \"hidden \", \"special \" or \"money\". TextControlTemplate controls how each value is drawn.",
        "params": [
            ["name", "string", "Title, such as \"ContactInfo\" (shown as \"Contact Info\")", True, None],
            ["data", "dict", "DynamicData object, or None", True, None],
            ["edit", "string", "HTML shown under the table", False, "\"\""],
            ["add", "string", "HTML shown when data is None", False, "\"\""],
            ["metadata", "dict", "DynamicData of field name to type hints; also limits and orders fields", False, "None"]
        ],
        "returns": "String of HTML",
        "watch_out": ["Values are not HTML escaped.", "The \"money\" hint only formats the value when it is exactly \"money\", with nothing else in the text.", "Fields holding another DynamicData or a list are left out.", "With metadata, a \"special \" field gets an id equal to its name split into words, such as \"First Name\", which contains a space."],
        "example": "dd = model.DynamicData()\ndd.FirstName = \"Ben\"\ndd.Gift = \"125.5\"\nmeta = model.DynamicData()\nmeta.FirstName = \"str\"\nmeta.Gift = \"money\"\nprint model.BuildDisplay(\"DonorInfo\", dd, \"\", \"\", meta)",
        "see_also": ["model.BuildDisplayRows", "model.BuildForm", "model.TextControlTemplate", "model.DynamicData"]
    },
    "model.BuildDisplayRows": {
        "category": "Page & Script",
        "kind": "method",
        "arity": [1, 2],
        "tags": ["read_only"],
        "summary": "Builds only the HTML rows used by BuildDisplay, one labeled read only row per field.",
        "details": "With metadata, only the listed fields are shown, in that order.",
        "params": [
            ["data", "dict", "DynamicData object", True, None],
            ["metadata", "dict", "DynamicData of field name to type hints", False, "None"]
        ],
        "returns": "String of HTML, or None",
        "watch_out": ["Returns None if data is None.", "Values are not HTML escaped.", "The \"money\" hint only formats the value when it is exactly \"money\".", "Without metadata no money formatting is done and every field is shown.", "Fields holding another DynamicData or a list are left out."],
        "example": "dd = model.DynamicData()\ndd.City = \"Hendersonville\"\ndd.State = \"TN\"\nprint model.BuildDisplayRows(dd)",
        "see_also": ["model.BuildDisplay", "model.BuildFormRows", "model.TextControlTemplate"]
    },
    "model.BuildForm": {
        "category": "Page & Script",
        "kind": "method",
        "arity": [2, 4],
        "tags": ["read_only"],
        "summary": "Builds an HTML form with one input per field of a DynamicData object, plus a heading and your buttons.",
        "details": "The form id is the name, and your buttons HTML goes at the end. Metadata can set each field to str, int, money, checkbox, textarea, or dropdown, and add \"hidden \" or \"readonly \". ControlTemplate controls how each input is drawn.",
        "params": [
            ["name", "string", "Form id and heading, such as \"EditContact\"", True, None],
            ["data", "dict", "DynamicData object with field values", True, None],
            ["buttons", "string", "HTML appended after the inputs", False, "\"\""],
            ["metadata", "dict", "DynamicData of field name to control type", False, "None"]
        ],
        "returns": "String of HTML",
        "watch_out": ["Values are not HTML escaped.", "str, int and money all draw the same plain text box; there is no number check or money format.", "A dropdown hint must be a DynamicData that contains the word dropdown and a values list of options.", "Hints are matched by the text they contain, checked in the order str, money, int, checkbox, dropdown, textarea. A dropdown whose options contain \"str\" or \"int\", such as \"Winter\", becomes a text box.", "readonly only affects text boxes and textareas. A checkbox is checked only when the value is the text True.", "The form tag has no action or method, and the buttons HTML goes inside the fieldset.", "A value containing a single quote breaks the input's value attribute."],
        "example": "dd = model.DynamicData()\ndd.Name = \"Ben\"\ndd.Notes = \"\"\nmeta = model.DynamicData()\nmeta.Name = \"str\"\nmeta.Notes = \"textarea\"\nhtml = model.BuildForm(\"EditContact\", dd, \"<button type='submit' class='btn btn-primary'>Save</button>\", meta)\n# In a real script, show it on the page:\n# model.Form = html\nprint html",
        "see_also": ["model.BuildFormRows", "model.BuildDisplay", "model.ControlTemplate", "model.DynamicData"]
    },
    "model.BuildFormRows": {
        "category": "Page & Script",
        "kind": "method",
        "arity": [1, 2],
        "tags": ["read_only"],
        "summary": "Builds only the input rows used by BuildForm, one labeled control per field.",
        "details": "Without metadata every field is a text box.",
        "params": [
            ["data", "dict", "DynamicData object", True, None],
            ["metadata", "dict", "DynamicData of field name to control type", False, "None"]
        ],
        "returns": "String of HTML",
        "watch_out": ["Values are not HTML escaped.", "Fields holding another DynamicData or a list are left out.", "A value containing a single quote breaks the input's value attribute."],
        "example": "dd = model.DynamicData()\ndd.Email = \"someone@example.com\"\nprint model.BuildFormRows(dd)",
        "see_also": ["model.BuildForm", "model.BuildDisplayRows", "model.ControlTemplate"]
    },
    "model.CallScript": {
        "category": "Page & Script",
        "kind": "method",
        "arity": [1, 1],
        "summary": "Runs another Python script from Special Content and returns everything it printed.",
        "details": "The called script shares this script's model.Data values, so you can pass values in with model.DictionaryAdd.",
        "flow": ["Loads the named Python script from Special Content (a Text item of the same name also works). A missing script is treated as empty.", "Runs it with a fresh model that shares this script's model.Data, so values either script sets in model.Data are seen by the other. Only model.FromMorningBatch is copied over.", "Returns everything the called script printed. An error in the called script is raised in your script with its Python traceback."],
        "params": [
            ["scriptName", "string", "Name of the Python script in Special Content", True, None]
        ],
        "returns": "string",
        "watch_out": ["Returns an empty result if the script is not found.", "Settings on your model such as model.TestEmail, model.MaxEmails, model.Transactional and model.ScheduledTime are not passed to the called script. Pass what it needs through model.Data."],
        "example": "model.DictionaryAdd(\"PeopleId\", model.UserPeopleId)\noutput = model.CallScript(\"MyOtherScript\")\nprint output",
        "see_also": ["model.DictionaryAdd", "model.Data", "model.PythonContent", "model.EmailReport"]
    },
    "model.CmsHost": {
        "category": "Page & Script",
        "kind": "property",
        "tags": ["read_only"],
        "summary": "The web address of your TouchPoint database, with no trailing slash.",
        "details": "It is read only. An example is https://mychurch.tpsdb.com. Use it to build links.",
        "returns": "string",
        "example": "url = model.CmsHost + \"/Person2/{0}\".format(model.UserPeopleId)\nprint url",
        "see_also": ["model.DatabaseName", "model.CreateTinyUrl", "model.GetAuthenticatedUrl"]
    },
    "model.ControlTemplate": {
        "category": "Page & Script",
        "kind": "property",
        "summary": "The HTML template BuildForm and BuildFormRows use to draw each editable input.",
        "details": "It is a public field. The tokens {name}, {displayname} and {input} are replaced for every field. The default is a Bootstrap form-group with a col-md-4 label and a col-md-4 input. Change it before calling BuildForm to restyle the form.",
        "returns": "string (settable)",
        "example": "model.ControlTemplate = '''<div class=\"form-group\">\n  <label for=\"{name}\">{displayname}</label>\n  {input}\n</div>'''\ndd = model.DynamicData()\ndd.FirstName = 'Ben'\nprint model.BuildForm('ContactInfo', dd)",
        "see_also": ["model.BuildForm", "model.BuildFormRows", "model.TextControlTemplate"]
    },
    "model.CurrentOrgId": {
        "category": "Page & Script",
        "kind": "property",
        "summary": "The involvement id used by email replacement codes that refer to an involvement.",
        "details": "Integer and settable. Defaults to 0. Set it before sending, such as with model.Email. Read it without parentheses.",
        "returns": "Integer",
        "example": "print \"Current:\", model.CurrentOrgId\noid = q.QuerySqlInt(\"SELECT TOP 1 OrganizationId FROM Organizations WHERE OrganizationStatusId = 30 ORDER BY OrganizationId\")\nmodel.CurrentOrgId = oid\nprint \"Now:\", model.CurrentOrgId",
        "see_also": ["model.Email", "model.EmailContentWithSubject", "model.GetOrganization"]
    },
    "model.Data": {
        "category": "Page & Script",
        "kind": "property",
        "summary": "Holds the values passed to the script: query string values, form POST values and anything set by the caller.",
        "details": "It is a read only property. Read and assign fields with dot notation.",
        "flow": ["On a POST to /PyScriptForm, each posted form field is added first. Uploaded files are added as text, and a zip file becomes a DynamicData of its files plus a keyword value.", "On every web run, each query string value is then added, overwriting a posted value with the same name.", "The p1 and p2 parts of the URL are added as p1 and p2, and pyscript is set to the script name.", "When the script text contains @\x42lueToolbarTagId, BlueToolbarGuid is added for the current search.", "Values added later with model.DictionaryAdd or a field assignment are in the same store."],
        "returns": "DynamicData object",
        "watch_out": ["A missing field returns an empty string, so hasattr is always True; test the value instead.", "All values from the web arrive as strings.", "Uploaded text files are read as plain ASCII, so accented characters become question marks."],
        "example": "orgId = model.Data.orgId\nif orgId:\n    print \"orgId =\", orgId\nelse:\n    print \"No orgId passed (try ?orgId=30)\"",
        "see_also": ["model.DataHas", "model.DictionaryAdd", "model.Dictionary", "model.HttpMethod"]
    },
    "model.DatabaseName": {
        "category": "Page & Script",
        "kind": "property",
        "tags": ["read_only"],
        "summary": "The church's host name, such as \"fbchville\".",
        "details": "The SQL database name is usually \"CMS_\" plus the host.",
        "returns": "string",
        "watch_out": ["Despite the name it is the host, not the SQL database name."],
        "example": "print \"Host:\", model.DatabaseName\nprint \"Site:\", \"https://\" + model.DatabaseName + \".tpsdb.com\"",
        "see_also": ["model.CmsHost"]
    },
    "model.DataHas": {
        "category": "Page & Script",
        "kind": "method",
        "arity": [1, 1],
        "tags": ["read_only"],
        "summary": "Returns True if model.Data has a value with this key.",
        "params": [
            ["key", "string", "Key name, case sensitive", True, None]
        ],
        "returns": "boolean",
        "watch_out": ["The key match is case sensitive.", "It returns True even when the value stored is an empty string."],
        "example": "if model.DataHas(\"orgId\"):\n    print \"orgId =\", model.Data.orgId\nelse:\n    print \"No orgId passed in\"",
        "see_also": ["model.Data", "model.Dictionary"]
    },
    "model.DebugPrint": {
        "category": "Page & Script",
        "kind": "method",
        "arity": [1, 1],
        "tags": ["read_only", "known_issue"],
        "summary": "Writes text to the .NET debug output.",
        "details": "Use print for anything you need to see.",
        "params": [
            ["s", "string", "Text to write", True, None]
        ],
        "returns": "None",
        "watch_out": ["On a live TouchPoint site this output goes nowhere and nothing is shown."],
        "example": "model.DebugPrint(\"only visible on a developer's debug build\")\nprint \"Use print to see output on a live site\"",
        "see_also": ["model.IsDebug"]
    },
    "model.Dictionary": {
        "category": "Page & Script",
        "kind": "method",
        "arity": [1, 1],
        "tags": ["read_only"],
        "summary": "Returns a value from the script's data dictionary as text.",
        "details": "This is the same store behind model.Data, so values posted to the script or set with model.DictionaryAdd are visible here.",
        "params": [
            ["s", "string", "Key name, case sensitive", True, None]
        ],
        "returns": "string",
        "watch_out": ["Returns an empty string if the key is missing.", "The key is case sensitive.", "Raises an error if the key exists but its value is None."],
        "example": "model.DictionaryAdd(\"Greeting\", \"Hello\")\nprint model.Dictionary(\"Greeting\")\nprint repr(model.Dictionary(\"Missing\"))",
        "see_also": ["model.DictionaryAdd", "model.Data", "model.DataHas"]
    },
    "model.DictionaryAdd": {
        "category": "Page & Script",
        "kind": "method",
        "arity": [2, 2],
        "summary": "Adds or replaces a value in the script's data dictionary.",
        "details": "The value is then readable as model.Data.key and model.Dictionary(key). It is passed along to scripts run with model.CallScript or model.EmailReport.",
        "params": [
            ["key", "string", "Key name", True, None],
            ["value", "any", "Value to store", True, None]
        ],
        "returns": "None",
        "watch_out": ["The value is stored as given, not converted to text, so model.Data.key returns the same type you stored."],
        "example": "model.DictionaryAdd(\"processed\", 100)\nprint model.Data.processed\nprint model.Dictionary(\"processed\")",
        "signatures": ["model.DictionaryAdd(key, value)"],
        "see_also": ["model.Dictionary", "model.Data", "model.CallScript", "model.EmailReport"]
    },
    "model.Form": {
        "category": "Page & Script",
        "kind": "property",
        "summary": "The HTML for the body of a /PyScriptForm/ScriptName page.",
        "details": "Set it to a string of HTML and TouchPoint renders it inside the standard page box on a GET request. If the text model.Form appears anywhere in a script, opening it at /PyScript/ScriptName redirects to /PyScriptForm/ScriptName. On a POST only the printed output comes back, which is how AJAX handlers return JSON or HTML fragments.",
        "flow": ["On GET /PyScript, if the script text matches model.Form anywhere, even in a comment, the browser is sent to /PyScriptForm instead.", "On GET /PyScriptForm, TouchPoint runs the script and renders model.Form inside the page box, with model.Styles in the head and model.Script at the bottom.", "On POST /PyScriptForm, posted fields go into model.Data, the script runs, and only the printed output is sent back as the response."],
        "returns": "string (settable)",
        "watch_out": ["On a POST to /PyScriptForm, Form, Title and Styles are ignored.", "On a GET to /PyScriptForm, printed output is thrown away, including Python error messages. A script that fails before setting model.Form shows an empty box.", "A #Roles= line at the top of the script only blocks /PyScript. /PyScriptForm does not check it, so test model.UserIsInRole yourself."],
        "example": "if model.HttpMethod == 'post':\n    print 'You clicked the button at ' + str(model.DateTime)\nelse:\n    model.Header = 'Form Demo'\n    model.Form = '''<p>Hello, %s.</p>\n<button class=\"btn btn-primary\" onclick=\"$.post(location.pathname, {}, function(r){ alert(r); })\">Try it</button>''' % model.UserName",
        "see_also": ["model.Header", "model.Script", "model.Styles", "model.HttpMethod"]
    },
    "model.FromMorningBatch": {
        "category": "Page & Script",
        "kind": "property",
        "summary": "True when the script is running from the morning batch rather than a signed in user.",
        "details": "It can be read and set.",
        "returns": "Boolean",
        "watch_out": ["Many methods that need a current user fail in that context.", "Scripts run with model.CallScript inherit the caller's value."],
        "example": "if model.FromMorningBatch:\n    print \"Running from the morning batch\"\nelse:\n    print \"Running interactively\"",
        "see_also": ["model.ScheduledTime", "model.UserPeopleId", "model.HttpMethod"]
    },
    "model.GetCacheVariable": {
        "category": "Page & Script",
        "kind": "method",
        "arity": [1, 1],
        "tags": ["read_only"],
        "summary": "Reads a value saved with model.SetCacheVariable.",
        "details": "Values are shared by every user of your database on that web server and expire 1 minute after being set.",
        "params": [
            ["name", "string", "Cache key", True, None]
        ],
        "returns": "string",
        "watch_out": ["Returns an empty string if missing, expired, or run from the morning batch.", "The cache lives in one web server's memory, so a value set on one server may not be seen on another."],
        "example": "model.SetCacheVariable(\"LastRun\", \"ok\")\nprint model.GetCacheVariable(\"LastRun\")",
        "see_also": ["model.SetCacheVariable"]
    },
    "model.Header": {
        "category": "Page & Script",
        "kind": "property",
        "summary": "The heading shown at the top of a /PyScript or /PyScriptForm page.",
        "details": "It is a settable string. It is also used as the browser tab title unless model.Title is set.",
        "returns": "String",
        "watch_out": ["On /PyScript, when neither Title nor Header is set, the tab title is \"Test Script\"."],
        "example": "model.Header = \"Weekly Attendance\"\nprint \"Header is:\", model.Header",
        "see_also": ["model.Title", "model.Form", "model.Styles"]
    },
    "model.HttpMethod": {
        "category": "Page & Script",
        "kind": "property",
        "tags": ["read_only"],
        "summary": "The current request's HTTP method in lowercase: \"get\" or \"post\".",
        "details": "It is read only. \"get\" means the first page load and \"post\" means a form or AJAX post back. On a POST to /PyScriptForm only printed output is returned.",
        "returns": "String: \"get\" or \"post\"",
        "watch_out": ["Returns \"get\" when there is no web request, such as in a batch run.", "On a POST, a Python error comes back as plain error text with a normal success status, so JavaScript that expects JSON should guard its parse."],
        "example": "if model.HttpMethod == \"post\":\n    print \"Handling a post back\"\nelse:\n    print \"First load, method is\", model.HttpMethod",
        "see_also": ["model.Data", "model.Form", "model.FromMorningBatch"]
    },
    "model.IsDebug": {
        "category": "Page & Script",
        "kind": "property",
        "tags": ["read_only"],
        "summary": "True only when TouchPoint itself is a developer DEBUG build.",
        "details": "It is read only. It controls whether model.ReadFile, model.WriteFile, model.DeleteFile and model.ExecuteSql are allowed.",
        "returns": "boolean",
        "watch_out": ["On hosted TouchPoint sites it is always False.", "When it is False, ReadFile, WriteFile, DeleteFile and ExecuteSql raise an error."],
        "example": "if model.IsDebug:\n    print 'Developer build: file and ExecuteSql helpers are available'\nelse:\n    print 'Production build: ReadFile, WriteFile and ExecuteSql will raise an error'",
        "see_also": ["model.ReadFile", "model.WriteFile", "model.ExecuteSql", "model.DeleteFile"]
    },
    "model.Output": {
        "category": "Page & Script",
        "kind": "property",
        "summary": "Holds the text the script printed, filled in after the script finishes.",
        "details": "TouchPoint renders it as the body of a /PyScript page. Use print instead of setting it. If the printed output starts with REDIRECT=url, a GET to /PyScript or a POST to /PyScriptForm redirects to that url. A GET to /PyScriptForm does not show printed output or redirect.",
        "flow": ["While the script runs, print output is collected.", "When the script ends, Output is set to that text. If the script raised an error, Output is the error message instead.", "On GET /PyScript, if Output starts with REDIRECT= the browser goes to that url; otherwise Output is rendered as the page body.", "On POST /PyScriptForm the printed text is returned directly, also honoring REDIRECT=. On GET /PyScriptForm it is not shown at all."],
        "returns": "string (filled in after the script runs)",
        "watch_out": ["It is empty while the script is running.", "Anything you assign is overwritten when the script ends.", "If the script text contains the phrase Background Process \x43ompleted, /PyScript runs it in the background and its output is saved to a text Special Content log instead of shown."],
        "example": "# Output is filled from print statements once the script ends\nprint '<p>This text becomes model.Output.</p>'\n# To send the user elsewhere, print a REDIRECT line as the only output:\n# print 'REDIRECT=/Person2/' + str(model.UserPeopleId)",
        "see_also": ["model.Form", "model.CallScript"]
    },
    "model.Script": {
        "category": "Page & Script",
        "kind": "property",
        "summary": "JavaScript added to the bottom of a /PyScript or GET /PyScriptForm page.",
        "details": "It is a settable string. Plain code is wrapped in a script tag for you. Text that starts with <script is inserted as is.",
        "returns": "String",
        "watch_out": ["Ignored on POST requests."],
        "example": "html = \"<button id='hi' class='btn btn-default'>Say hi</button>\"\njs = \"document.getElementById('hi').onclick = function() { alert('Hi'); };\"\n# In a real /PyScriptForm script, these three lines build the page:\n# model.Header = \"Script Demo\"\n# model.Form = html\n# model.Script = js\nprint html\nprint \"Script for the page:\", js",
        "see_also": ["model.Form", "model.Styles", "model.Header"]
    },
    "model.ScriptName": {
        "category": "Page & Script",
        "kind": "property",
        "summary": "The Special Content name of the running script.",
        "details": "TouchPoint sets it for /PyScript, /PyScriptForm (GET and POST), model.CallScript, the morning batch, registration event scripts, SMS actions and special reports. It is also recorded as the source when model.AddPerson or model.FindAddPerson creates a person.",
        "returns": "string, or None if not set",
        "watch_out": ["It can be None when a script is run some other way, so provide a fallback."],
        "example": "name = model.ScriptName or 'UnknownScript'\nprint 'This script is ' + name\nprint 'Post AJAX calls to /PyScriptForm/' + name",
        "see_also": ["model.CallScript", "model.FromMorningBatch", "model.AddPerson"]
    },
    "model.SetCacheVariable": {
        "category": "Page & Script",
        "kind": "method",
        "arity": [2, 2],
        "summary": "Saves a text value in the web server's memory for 1 minute.",
        "params": [
            ["name", "string", "Cache key", True, None],
            ["value", "string", "Text to store", True, None]
        ],
        "returns": "None",
        "watch_out": ["The value is shared by every user of your database, so do not store anything private or per user.", "Does nothing when run from the morning batch.", "The cache lives in one web server's memory, so another server may not see the value.", "Expiry is fixed at 1 minute from the last set; reading it does not extend it."],
        "example": "model.SetCacheVariable(\"LastRun\", \"ok\")\nprint model.GetCacheVariable(\"LastRun\")",
        "see_also": ["model.GetCacheVariable"]
    },
    "model.Styles": {
        "category": "Page & Script",
        "kind": "property",
        "summary": "CSS or other markup that TouchPoint places in the page head for /PyScript or /PyScriptForm output.",
        "details": "Include the style tags yourself, because the value is inserted as raw HTML.",
        "returns": "string (settable)",
        "watch_out": ["Ignored on POST requests."],
        "example": "model.Styles = '<style>.big { font-size: 24px; color: #2a6ebb; }</style>'\nprint '<div class=\"big\">Styled output</div>'",
        "see_also": ["model.Script", "model.Form", "model.Header"]
    },
    "model.TextControlTemplate": {
        "category": "Page & Script",
        "kind": "property",
        "summary": "The HTML template BuildDisplay and BuildDisplayRows use to draw each read only value.",
        "details": "It is a public field. The tokens {displayname}, {o} (the value) and {id} are replaced for every field. Change it before calling BuildDisplay to restyle the display.",
        "returns": "string (settable)",
        "watch_out": ["{id} is empty unless the field's metadata marks it \"special \"."],
        "example": "model.TextControlTemplate = '<p><b>{displayname}:</b> <span {id}>{o}</span></p>'\ndd = model.DynamicData()\ndd.FirstName = 'Ben'\ndd.City = 'Hendersonville'\nprint model.BuildDisplay('ContactInfo', dd)",
        "see_also": ["model.BuildDisplay", "model.BuildDisplayRows", "model.ControlTemplate"]
    },
    "model.Title": {
        "category": "Page & Script",
        "kind": "property",
        "summary": "The browser tab title for the page a script renders at /PyScript or /PyScriptForm.",
        "details": "When it is not set, TouchPoint uses model.Header instead.",
        "returns": "string (settable)",
        "watch_out": ["It has no effect on POST requests or on scripts run from the morning batch.", "On /PyScript, when neither Title nor Header is set, the tab title is \"Test Script\"."],
        "example": "model.Title = 'Attendance Check'\nmodel.Header = 'Weekly Attendance Check'\nprint 'The tab title is: ' + model.Title",
        "see_also": ["model.Header", "model.Form", "model.Styles"]
    },
    "model.UserIsInRole": {
        "category": "Page & Script",
        "kind": "method",
        "arity": [1, 1],
        "tags": ["read_only", "known_issue"],
        "summary": "Returns True if the signed in user has the role.",
        "params": [
            ["role", "string", "Role name, such as Admin or Finance", True, None]
        ],
        "returns": "boolean",
        "watch_out": ["When there is no web request it returns True for every role if running from the morning batch, and False for every role otherwise."],
        "example": "for r in [\"Admin\", \"Finance\", \"Developer\"]:\n    print r, model.UserIsInRole(r)",
        "see_also": ["model.UserPeopleId", "model.FromMorningBatch", "model.AddRole"]
    },
    "model.UserName": {
        "category": "Page & Script",
        "kind": "property",
        "tags": ["read_only"],
        "summary": "The sign in username of the user running the script.",
        "details": "It is read only.",
        "returns": "String",
        "example": "print \"Username:\", model.UserName",
        "see_also": ["model.UserPeopleId", "model.UserIsInRole"]
    },
    "model.UserPeopleId": {
        "category": "Page & Script",
        "kind": "property",
        "tags": ["read_only"],
        "summary": "The PeopleId of the signed in user.",
        "details": "Read only.",
        "returns": "integer or None",
        "watch_out": ["It is a property, not a method: model.UserPeopleId() raises \"'int' object is not callable\".", "It is None when there is no signed in user, such as in the morning batch."],
        "example": "pid = model.UserPeopleId\nif pid:\n    print \"You are PeopleId\", pid, model.GetPerson(pid).Name\nelse:\n    print \"No signed in user\"",
        "see_also": ["model.UserName", "model.UserIsInRole", "model.GetPerson"]
    },
    "model.AddBackgroundCheck": {
        "category": "People",
        "kind": "method",
        "arity": [2, 10],
        "tags": ["writes_data", "outside_service"],
        "summary": "Orders a Protect My Ministry background check for a person through TouchPoint's PMM connection.",
        "details": "This is for Protect My Ministry (PMM) only. MinistrySafe checks share TouchPoint's background check list, but scripts have no call to start one. The check is recorded in the BackgroundChecks table under the person, with the signed in user as the one who ordered it, so it shows on the person's background check page like one started there.",
        "flow": ["Finds the package in your PMM package list (Special Content PMMUserPackages) by code, or by name when sPackageName is given. TouchPoint refreshes that list when someone opens the background check screen.", "Saves any SSN or driver license number you pass on the person's record, encrypted, then adds a BackgroundChecks row with status Not submitted.", "Finds or creates the person as a candidate in PMM. This needs an email address, and an SSN on file for the PII submit type.", "Sends the order. PII: TouchPoint sends the person's details and the check runs now. EMAIL: PMM emails the person a link to give consent and enter their own details (status Emailed). RE_SUBMIT: runs again for someone already in PMM.", "If PMM answers right away, the status becomes Complete, the report link is saved, and the person who ordered it gets an email. Otherwise the status is Submitted and PMM sends the result to TouchPoint later.", "A finished check with any problem is flagged for review (IssueCount 1). The label you pick becomes the billing reference."],
        "params": [
            ["peopleId", "integer", "Person to check", True, None],
            ["code", "string", "PMM package code from PMMUserPackages", True, None],
            ["submitType", "string", "\"PII\" (send their details now), \"EMAIL\" (PMM emails them for consent) or \"RE_SUBMIT\" (run again for someone already in PMM)", False, "\"PII\""],
            ["type", "integer", "Report type id", False, "1"],
            ["label", "integer", "Background check label id from lookup.BackgroundCheckLabels; its code becomes the billing reference", False, "0"],
            ["sSSN", "string", "Social Security number to save on the person first", False, "None"],
            ["sDLN", "string", "Driver license number to save on the person first, for a motor vehicle record check", False, "None"],
            ["sPlusCounty", "string", "County to add a county criminal search", False, "\"\""],
            ["sPlusState", "string", "State to add a statewide criminal search", False, "\"\""],
            ["sPackageName", "string", "Package name to match when you do not have the code", False, "None"]
        ],
        "returns": "Integer id of the new BackgroundChecks row. A failed submit still returns an id; check that row's StatusID and ErrorMessages.",
        "watch_out": ["Orders a real, billable check.", "Protect My Ministry only. There is no script call for MinistrySafe.", "Raises an error when no package matches the code or name, for example when PMMUserPackages is empty.", "Problems such as a missing email, missing SSN or unconfigured PMM do not raise. They are saved on the check (StatusID 0 Error or 1 Not submitted, with ErrorMessages).", "An SSN or driver license number you pass overwrites the one on the person's record."],
        "example": "import json\npackages = json.loads(model.TextContent(\"PMMUserPackages\") or \"[]\")\nprint len(packages), \"PMM packages:\"\nfor pk in packages:\n    print \" \", pk.get(\"packageCode\"), \"|\", pk.get(\"packageName\")\nnames = {0: \"Error\", 1: \"Not submitted\", 2: \"Submitted\", 3: \"Complete\", 4: \"Emailed\"}\nfor c in q.QuerySql(\"SELECT TOP 5 Id, PeopleID, StatusID, ServiceCode, Created FROM BackgroundChecks ORDER BY Id DESC\"):\n    print c.Id, c.PeopleID, names.get(c.StatusID, c.StatusID), c.Created\n# Orders a real, billable check. Pick a code from the list above, then uncomment:\n# checkId = model.AddBackgroundCheck(model.UserPeopleId, \"CODE\", \"EMAIL\")\n# row = q.QuerySqlTop1(\"SELECT StatusID, ErrorMessages FROM BackgroundChecks WHERE Id = {0}\".format(checkId))\n# print names.get(row.StatusID), row.ErrorMessages",
        "see_also": ["model.GetPerson", "q.QuerySqlTop1", "model.TextContent"]
    },
    "model.AddPerson": {
        "category": "People",
        "kind": "method",
        "arity": [4, 7],
        "tags": ["writes_data", "known_issue"],
        "summary": "Creates a new person in a new family as the primary adult and returns the new PeopleId.",
        "details": "Gender codes: 0 unknown, 1 male, 2 female. Marital codes: 0 unknown, 10 single, 20 married, 30 separated, 40 divorced, 50 widowed.",
        "flow": ["Builds a new person with the first, last and nickname you give. Names typed all lowercase or all uppercase are title cased, and long names are cut (first name 25 characters, nickname 15).", "Creates a new Families row with no address and puts the person in it as Primary Adult, with member status Just Added.", "Saves the People row, then sets gender, marital status and email address and saves again.", "Returns the new PeopleId. No duplicate check, no field requirement check and no new person notices or tasks."],
        "params": [
            ["firstname", "string", "First name", True, None],
            ["nickname", "string", "Nickname (goes by), or empty string", True, None],
            ["lastname", "string", "Last name", True, None],
            ["email", "string", "Email address", True, None],
            ["marital", "integer", "Marital status code", False, "0"],
            ["gender", "integer", "Gender code", False, "0"],
            ["familymemberid", "integer", "Ignored by the current code", False, "None"]
        ],
        "returns": "Integer PeopleId of the new person",
        "watch_out": ["Writes data.", "It does not check for duplicates, so prefer model.FindAddPeopleId when you might already have the person.", "The familymemberid argument is accepted but ignored.", "No title (Mr., Mrs., Ms.) is filled in, because gender is set after the title is worked out.", "No entry point or origin is set, and the email address is not checked for a valid format.", "Nobody is told: the new people managers get no email and no task is created, unlike model.FindAddPeopleId."],
        "example": "# Creates a real person record. Uncomment to run.\n# newId = model.AddPerson(\"John\", \"\", \"Smith\", \"john.smith@example.com\", 20, 1)\n# print \"Created PeopleId\", newId\nprint \"Use model.FindAddPeopleId to avoid duplicates\"",
        "see_also": ["model.FindAddPeopleId", "model.FindAddPerson", "model.FindPersonId"]
    },
    "model.AddRole": {
        "category": "People",
        "kind": "method",
        "arity": [2, 2],
        "tags": ["writes_data"],
        "summary": "Adds a user role to everyone in a query.",
        "details": "People without a user account get one created, with a random password, so the role can be attached.",
        "flow": ["Refuses the powerful roles in the list without raising an error, and only writes an ActivityLog line. The check ignores case.", "Runs the query. For each person, takes their first user account and adds the role if they do not have it yet.", "A person with no account gets one: username from first initial plus last name (letters only, a number added if taken), a random password they are never told, approved, with no welcome email.", "Saves after each person. When finished, asks the mobile app and API to drop their cached roles so the change takes effect on the next request."],
        "params": [
            ["query", "query", "People to act on: a PeopleId, a saved search name, or Search Builder code", True, None],
            ["role", "string", "Role name to add", True, None]
        ],
        "returns": "None",
        "watch_out": ["Writes data, including new user accounts.", "Powerful roles are refused and only logged: Admin, ManageApplication, BackgroundCheck, Delete, Developer, Finance, FinanceAdmin, Manager, Manager2, Membership, ManageTransactions, MemberDocs.", "The role must already exist. A misspelled role raises \"Role ... does not exist\" partway, after earlier people are already saved.", "Only the person's first user account gets the role when they have more than one.", "The query skips deceased and archived people unless it asks for them (for example IncludeArchived=1[True]), so passing a plain PeopleId for such a person silently does nothing."],
        "example": "# Gives you the Access role. Uncomment to run.\n# model.AddRole(str(model.UserPeopleId), \"Access\")\nprint \"Roles changed on the user record of each person in the query\"",
        "see_also": ["model.RemoveRole", "model.UserIsInRole"]
    },
    "model.ArchiveRecords": {
        "category": "People",
        "kind": "method",
        "arity": [1, 1],
        "tags": ["writes_data"],
        "summary": "Sets the ArchivedFlag on everyone in a query, hiding them from normal searches.",
        "details": "Changes are logged. Reverse it with model.UnArchiveRecords.",
        "flow": ["Runs the query. Archived and deceased people are already left out unless the query asks for them.", "For each person, sets ArchivedFlag to true, writes a Basic Info ChangeLog entry and saves.", "Writes one ActivityLog line with the count."],
        "params": [
            ["query", "query", "People to act on: a PeopleId, a saved search name, or Search Builder code", True, None]
        ],
        "returns": "None",
        "watch_out": ["Writes data.", "model.UnArchiveRecords uses the same kind of query, which leaves archived people out by default, so a plain PeopleId finds no one. Include IncludeArchived=1[True] or an ArchivedFlag condition in the query to undo."],
        "example": "print q.QueryCount(\"peopleids='999999'\"), \"people would be archived\"\n# Uncomment to run:\n# model.ArchiveRecords(\"peopleids='999999'\")",
        "see_also": ["model.UnArchiveRecords", "model.DeletePeople"]
    },
    "model.BirthdayList": {
        "category": "People",
        "kind": "method",
        "arity": [0, 0],
        "tags": ["read_only"],
        "summary": "Returns upcoming birthdays for the signed in user, like the home page Birthdays panel.",
        "details": "The source is their TrackBirthdays saved search, their TrackBirthdays tag, or their main fellowship class, in that order. The tag and class lists are limited to the next 15 days. The saved search list is the next 100 birthdays with no day limit.",
        "flow": ["Returns an empty list when nobody is signed in.", "If the user has a saved search named TrackBirthdays, refreshes a system tag from it once a day and returns up to 100 of those people with a birthday on file, soonest birthday first. This path has no 15 day limit and does not skip deceased people.", "Otherwise uses the user's personal TrackBirthdays tag, creating an empty one if they do not have it.", "If that tag is empty, uses the members of the user's main fellowship class.", "For the tag and class sources, keeps living people whose next birthday is within 15 days, soonest first. Birthday is text such as March 5, and a February 29 birthday shows as February 28."],
        "returns": "List of objects with PeopleId, Name and Birthday (text such as \"March 5\")",
        "watch_out": ["Returns an empty list when nobody is signed in, such as in a batch run.", "Calling it can create an empty TrackBirthdays personal tag for the signed in user.", "With a TrackBirthdays saved search, the list is the next 100 birthdays at any distance, not just the next 15 days, and deceased people are included."],
        "example": "for b in model.BirthdayList():\n    print b.Birthday, b.Name, b.PeopleId",
        "see_also": ["model.UserPeopleId"]
    },
    "model.DeletePeople": {
        "category": "People",
        "kind": "method",
        "arity": [1, 1],
        "tags": ["writes_data", "known_issue"],
        "summary": "Permanently purges everyone in a query from the database.",
        "flow": ["Turns the query into a list of PeopleIds.", "For each person, runs the database purge in its own transaction and writes an ActivityLog line.", "The purge deletes enrollments and enrollment history, attendance (meeting counts are recalculated), extra values, tags they are in and tags they own, tasks and notes about or owned by them, registrations, volunteer and background check records, saved payment methods and recurring giving, and their user account with its roles.", "Contributions are kept but no longer linked to anyone, and ActivityLog rows lose their PeopleId. If they were head of household or spouse those links are cleared.", "When the last person in a family is purged, the family and its related family links are deleted too, along with the person's picture."],
        "params": [
            ["query", "query", "People to act on: a PeopleId, a saved search name, or Search Builder code", True, None]
        ],
        "returns": "None",
        "watch_out": ["Deletes data and cannot be undone.", "TouchPoint only logs a warning when the user lacks the Developer role; it does not stop the delete.", "The query skips deceased and archived people unless it asks for them (for example IncludeArchived=1[True]), so passing a plain PeopleId for such a person silently does nothing.", "PeopleId 1 is never purged and no error says so.", "If one person fails, an error is raised and the loop stops. Everyone before them is already gone.", "Their giving history stays in the totals but is no longer attached to any person."],
        "example": "# PERMANENT. Check the count first and run only on a tight query.\nprint q.QueryCount(\"peopleids='999999'\"), \"people would be deleted\"\n# model.DeletePeople(\"peopleids='999999'\")",
        "see_also": ["model.ArchiveRecords", "q.QueryCount"]
    },
    "model.FindAddPeopleId": {
        "category": "People",
        "kind": "method",
        "arity": [5, 6],
        "tags": ["writes_data", "sends_message"],
        "summary": "Finds an existing person by email, or by name plus email, phone or birth date, or creates one, and returns the PeopleId.",
        "details": "Same as model.FindAddPerson but returns the PeopleId.",
        "flow": ["Runs model.FindAddPerson with the same arguments and returns only the PeopleId.", "Matching order, record creation and new person notices are exactly as described for model.FindAddPerson."],
        "params": [
            ["first", "string", "First name", True, None],
            ["last", "string", "Last name", True, None],
            ["dob", "string", "Birth date such as \"1/15/1980\", or None", True, None],
            ["email", "string", "Email address, or None", True, None],
            ["phone", "string", "Phone number, or None", True, None],
            ["firstLastMatch", "boolean", "Require first and last name to match too", False, "False"]
        ],
        "returns": "Integer PeopleId",
        "watch_out": ["Creating a person writes data.", "Creating a person can send new person notices.", "With firstLastMatch False, an email alone decides the match, so a shared family email returns the existing person who owns it even when the names differ."],
        "example": "# Creates a person if no match is found. Uncomment to run.\n# pid = model.FindAddPeopleId(\"John\", \"Smith\", None, \"john.smith@example.com\", \"6155551234\", True)\n# print \"PeopleId\", pid\nprint \"Existing match (None if nobody):\", model.FindPersonId(\"John\", \"Smith\", None, \"john.smith@example.com\", None)",
        "see_also": ["model.FindAddPerson", "model.FindPersonId", "model.AddPerson"]
    },
    "model.FindAddPerson": {
        "category": "People",
        "kind": "method",
        "arity": [5, 6],
        "tags": ["writes_data", "sends_message"],
        "summary": "Finds an existing person by email, or by name plus email, phone or birth date, or creates a new one when there is no match.",
        "details": "Returns the Person record, not an id; use model.FindAddPeopleId for the id. When multiple people match, the lowest PeopleId wins.",
        "flow": ["With an email and firstLastMatch False, it first looks only at email. The oldest record whose EmailAddress or EmailAddress2 equals it wins, whatever the name.", "Otherwise, with first, last and email, it looks for a name match whose own email, second email or a family member's email matches. If that does not give exactly one person, it looks for a child of that name in a family whose primary adult has that email.", "If still nothing, it tries first and last name with the phone (family home phone, own cell, or a family member's cell), then first and last name with the birth date (year may be off by one). Name matching also accepts nickname, maiden and middle names and starts of names, and skips deceased people.", "If anyone matched, it returns the matched person with the lowest PeopleId and changes nothing else.", "If no one matched, it creates a new family with the person as Primary Adult, member status Just Added, origin Contribution and an entry point named after your script (created if missing). The phone is saved as the home phone, and Comments says \"Added in context of ... because record was not found\".", "On create, if the signed in user has Access, is not OrgLeadersOnly and is not the new people manager, a New Person Data Entry task can be created (when new people tasks are on). The new people managers are emailed \"Just Added Person\" when an admin email is set, even from a batch run."],
        "params": [
            ["first", "string", "First name", True, None],
            ["last", "string", "Last name", True, None],
            ["dob", "string", "Birth date such as \"1/15/1980\", or None", True, None],
            ["email", "string", "Email address, or None", True, None],
            ["phone", "string", "Phone number, or None", True, None],
            ["firstLastMatch", "boolean", "Require first and last name to match too", False, "False"]
        ],
        "returns": "Person object",
        "watch_out": ["Creating a person writes data.", "Creating a person can send new person notices.", "With firstLastMatch False, an email alone decides the match. Passing a spouse's shared email returns the spouse, not a new person, even when the names differ. Set firstLastMatch True when names must agree.", "The email only match does not skip deceased or archived people.", "Every script that creates someone this way adds an entry point named after that script to your entry point list.", "The phone you pass is stored as home phone on the person and the new family, not as cell phone."],
        "example": "# Creates a person if no match is found. Uncomment to run.\n# p = model.FindAddPerson(\"John\", \"Smith\", None, \"john.smith@example.com\", None, True)\n# print p.PeopleId, p.Name\nprint \"Existing match (None if nobody):\", model.FindPersonId(\"John\", \"Smith\", None, \"john.smith@example.com\", None)",
        "see_also": ["model.FindAddPeopleId", "model.FindPersonId", "model.AddPerson", "model.GetPerson"]
    },
    "model.FindPersonId": {
        "category": "People",
        "kind": "method",
        "arity": [4, 5],
        "tags": ["read_only"],
        "summary": "Looks for an existing person by first and last name plus email, phone or birth date, and returns the first match's PeopleId.",
        "details": "Uses TouchPoint's matching rules. A second form takes a full name and splits it into first and last.",
        "flow": ["Extracts the digits from the phone. The full name form first splits the name into first and last.", "With first, last and email, looks for people whose first name (or nickname, or a name it starts with) and last name (or maiden or middle name) match, then scores them on email, phone and birth date. Accents and spaces in names are ignored.", "If that does not give exactly one person, it looks for a child with that name whose parent has that email.", "If still nothing, it tries first and last name with the phone (cell, family home phone, or a family member's cell), then first and last name with the birth date, where the year may be off by one.", "Returns the PeopleId of the first person left, or None. Nothing is saved."],
        "params": [
            ["first", "string", "First name", True, None],
            ["last", "string", "Last name", True, None],
            ["dob", "string", "Birth date such as \"1/15/1980\", or None", True, None],
            ["email", "string", "Email address, or None", True, None],
            ["phone", "string", "Phone number (digits are extracted), or None", True, None]
        ],
        "returns": "Integer PeopleId, or None if no match",
        "watch_out": ["Returns None when no one matches.", "Never creates a record.", "Both first and last name are required. An email alone never matches, unlike model.FindAddPerson.", "A name with no email, phone or birth date always returns None.", "A birth date that cannot be read as a date is ignored."],
        "example": "pid = model.FindPersonId(\"John\", \"Smith\", None, \"john.smith@example.com\", None)\nif pid:\n    print \"Found\", pid, model.GetPerson(pid).Name\nelse:\n    print \"No match\"",
        "signatures": ["model.FindPersonId(first, last, dob, email, phone)", "model.FindPersonId(fullName, dob, email, phone)"],
        "see_also": ["model.FindAddPerson", "model.FindPersonIdExtraValue", "model.GetPerson"]
    },
    "model.FindPersonIdExtraValue": {
        "category": "People",
        "kind": "method",
        "arity": [2, 2],
        "tags": ["read_only"],
        "summary": "Finds the first person whose extra value field matches the given value as text, long text or (for a whole number) integer.",
        "params": [
            ["extraKey", "string", "Extra value field name", True, None],
            ["extraValue", "string", "Value to match", True, None]
        ],
        "returns": "Integer PeopleId, or None",
        "watch_out": ["Returns None when there is no match.", "It matches the field's text value, its long text (Data) value, and, when the value is a whole number, its integer value.", "When several people match, which one comes back first is not defined.", "Deceased and archived people are not skipped."],
        "example": "pid = model.FindPersonIdExtraValue(\"EmployeeId\", \"12345\")\nprint pid if pid else \"No match\"",
        "see_also": ["model.FindPersonIdExtraValueInt", "model.ExtraValueText", "model.FindPersonId"]
    },
    "model.FindPersonIdExtraValueInt": {
        "category": "People",
        "kind": "method",
        "arity": [2, 2],
        "tags": ["read_only", "known_issue"],
        "summary": "Meant to find a person by an integer extra value.",
        "details": "Prefer a SQL query on PeopleExtra with both Field and IntValue.",
        "params": [
            ["extraKey", "string", "Extra value field name (ignored by the current code)", True, None],
            ["extraValue", "integer", "Integer value to match", True, None]
        ],
        "returns": "Integer PeopleId, or None",
        "watch_out": ["The current code ignores extraKey.", "It returns the first person with ANY integer extra value equal to the number, so results can come from an unrelated field.", "Returns None when there is no match.", "A value that is not a whole number is treated as 0, so it matches anyone with an integer extra value of 0 in any field."],
        "example": "pid = q.QuerySqlInt(\"SELECT TOP 1 PeopleId FROM PeopleExtra WHERE Field = 'BadgeNumber' AND IntValue = 100\")\nprint pid if pid else \"No match\"",
        "see_also": ["model.FindPersonIdExtraValue", "model.ExtraValueInt", "q.QuerySqlInt"]
    },
    "model.GetPerson": {
        "category": "People",
        "kind": "method",
        "arity": [1, 1],
        "tags": ["read_only"],
        "summary": "Loads a person record by PeopleId.",
        "details": "Passing an integer returns the full Person record, whose fields (Name, EmailAddress, CellPhone, FamilyId and so on) and related data you can read. Passing a string id returns a smaller summary object.",
        "params": [
            ["peopleId", "integer", "PeopleId of the person", True, None]
        ],
        "returns": "Person object, or None if not found",
        "watch_out": ["Returns None when no one has that id.", "A string id returns a summary object, not the full Person record.", "The summary object (string id) includes SpouseName, family and personal addresses, emergency contact, medical description and Usernames, which the full record does not offer as simple fields."],
        "example": "person = model.GetPerson(model.UserPeopleId)\nif person:\n    print person.Name\n    print person.EmailAddress\nelse:\n    print \"No person with that id\"",
        "signatures": ["model.GetPerson(peopleId)  # integer: full Person record", "model.GetPerson(pid)  # any other value: summary person object"],
        "see_also": ["model.GetSpouse", "model.FindPersonId", "model.UpdatePerson"]
    },
    "model.GetSpouse": {
        "category": "People",
        "kind": "method",
        "arity": [1, 1],
        "tags": ["read_only"],
        "summary": "Returns a summary object for the person's spouse.",
        "details": "The object has PeopleId, Name, FirstName, LastName, EmailAddress, CellPhone and similar fields.",
        "params": [
            ["pid", "integer", "PeopleId of the person", True, None]
        ],
        "returns": "Summary person object, or None",
        "watch_out": ["Returns None when the person is not found or has no spouse.", "It follows the person's stored SpouseId. When that is empty it returns None even for a married couple; model.UpdateAllSpouseId fills in missing values.", "The spouse is returned even if deceased."],
        "example": "spouse = model.GetSpouse(model.UserPeopleId)\nif spouse:\n    print spouse.Name, spouse.EmailAddress\nelse:\n    print \"No spouse on record\"",
        "see_also": ["model.GetPerson", "model.UpdateAllSpouseId"]
    },
    "model.RemoveRole": {
        "category": "People",
        "kind": "method",
        "arity": [2, 2],
        "tags": ["writes_data"],
        "summary": "Removes a user role from everyone in a query who has a user account.",
        "details": "People without that role are skipped. Accepts a PeopleId, a saved search name, or Search Builder code.",
        "flow": ["Writes an Activity Log entry naming the role.", "Turns the query text into people: digits alone become one PeopleId, peopleids='1,2,3' becomes a list, text that exactly matches any saved search name runs that search, and anything else is read as Search Builder code. An empty string matches no one.", "Deceased and archived people are left out unless the code itself mentions them (IncludeDeceased, IncludeArchived, DeceasedDate or ArchivedFlag). A campus restriction or OrgLeadersOnly limit on the signed in user also narrows the list.", "For each person, takes their first user account and removes the role. The role name match ignores case. Accounts without the role are skipped.", "Each change is saved right away, and the changed users' cached roles are cleared so the change takes effect on their next page load."],
        "params": [
            ["query", "query", "People to act on: a PeopleId, a saved search name, or Search Builder code", True, None],
            ["role", "string", "Role name to remove", True, None]
        ],
        "returns": "None",
        "watch_out": ["Changes user roles.", "Only the person's first user account is changed. A second login for the same person keeps the role.", "No change log entry is written for the role change."],
        "example": "# Removes the Access role from you. Uncomment to run.\n# model.RemoveRole(str(model.UserPeopleId), \"Access\")\nprint \"RemoveRole only touches people who have a user account\"",
        "see_also": ["model.AddRole", "model.UserIsInRole"]
    },
    "model.UnArchiveRecords": {
        "category": "People",
        "kind": "method",
        "arity": [1, 1],
        "tags": ["writes_data"],
        "summary": "Clears the ArchivedFlag on everyone in a query so they show in normal searches again.",
        "details": "Changes are logged. Accepts a PeopleId, a saved search name, or Search Builder code.",
        "flow": ["Turns the query text into people: digits alone become one PeopleId, peopleids='1,2,3' becomes a list, text that exactly matches any saved search name runs that search, and anything else is read as Search Builder code. An empty string matches no one.", "Unless the code mentions archived people, archived records are left out of the match, so they are never reached.", "For each person found, sets ArchivedFlag to False, writes a change log entry and saves.", "Writes one Activity Log entry with the number of records processed."],
        "params": [
            ["query", "query", "People to act on: a PeopleId, a saved search name, or Search Builder code", True, None]
        ],
        "returns": "None",
        "watch_out": ["Changes person records.", "A plain PeopleId or peopleids='...' finds no archived people, so nothing changes. Use code such as ArchivedFlag = 1[True] AND PeopleId = 123, or add IncludeArchived = 1[True].", "Each person is loaded and saved one at a time, so a query that matches thousands of people is slow."],
        "example": "print q.QueryCount(\"ArchivedFlag = 1[True] AND PeopleIds = '999999'\"), \"archived people would be unarchived\"\n# Uncomment to run:\n# model.UnArchiveRecords(\"ArchivedFlag = 1[True] AND PeopleIds = '999999'\")",
        "see_also": ["model.ArchiveRecords", "q.QueryCount"]
    },
    "model.UpdateAllSpouseId": {
        "category": "People",
        "kind": "method",
        "arity": [0, 0],
        "tags": ["writes_data", "known_issue"],
        "summary": "Fills in missing SpouseId values for everyone in the database, and missing head of household spouse ids on families.",
        "flow": ["Checks for the Developer role, but only writes a \"denied\" Activity Log entry if it is missing.", "Fills in Families.HeadOfHouseholdSpouseId where it is empty.", "Fills in People.SpouseId where it is empty, then writes an Activity Log entry."],
        "returns": "None",
        "watch_out": ["Writes data across the whole database.", "When the user lacks the Developer role it only logs a warning; it does not stop the update.", "Only empty values are filled. A SpouseId that is already set but wrong is not corrected."],
        "example": "# Touches every person record. Uncomment to run.\n# model.UpdateAllSpouseId()\nprint \"SpouseId recalculation is database wide\"",
        "see_also": ["model.GetSpouse"]
    },
    "model.UpdateCampus": {
        "category": "People",
        "kind": "method",
        "arity": [2, 2],
        "tags": ["writes_data"],
        "summary": "Sets the campus for everyone in a query.",
        "details": "Pass a campus id or a campus name. Changes are logged. Accepts a PeopleId, a saved search name, or Search Builder code.",
        "flow": ["An integer or a string of digits is used as the campus id as is. Any other text is looked up by exact campus description, and a new campus is created when none matches.", "Turns the query text into people: digits alone become one PeopleId, peopleids='1,2,3' becomes a list, text that exactly matches any saved search name runs that search, and anything else is read as Search Builder code. An empty string matches no one.", "Deceased and archived people are left out unless the code itself mentions them (IncludeDeceased, IncludeArchived, DeceasedDate or ArchivedFlag). A campus restriction or OrgLeadersOnly limit on the signed in user also narrows the list.", "For each person, sets CampusId, writes a change log entry and saves."],
        "params": [
            ["query", "query", "People to act on: a PeopleId, a saved search name, or Search Builder code", True, None],
            ["campus", "any", "Campus id (integer) or campus name (string)", True, None]
        ],
        "returns": "None",
        "watch_out": ["A campus name that does not exist yet creates a new campus.", "Changes person records.", "A campus id is not checked against the campus list.", "Campus 0, None or an empty name does nothing.", "Each person is loaded and saved one at a time, so a query that matches thousands of people is slow."],
        "example": "campuses = list(q.QuerySql(\"SELECT Id, Description FROM lookup.Campus\"))\nprint len(campuses), \"campuses\"\nfor c in campuses:\n    print c.Id, c.Description\n# Uncomment to move yourself to campus 1:\n# model.UpdateCampus(str(model.UserPeopleId), 1)",
        "see_also": ["model.UpdateMemberStatus", "model.UpdateNamedField"]
    },
    "model.UpdateField": {
        "category": "People",
        "kind": "method",
        "arity": [3, 3],
        "tags": ["writes_data"],
        "summary": "Sets one field on a person record.",
        "details": "Pass the Person object from model.GetPerson (with an integer id).",
        "flow": ["Loads a fresh copy of the person by the PeopleId on the object you pass.", "For HomePhone, sets the value on the family and writes a family change log entry.", "For any other field, sets the value on the person. Text is converted to the field's type, and an unchanged value is skipped.", "Saves."],
        "params": [
            ["p", "any", "Person object from model.GetPerson(peopleId)", True, None],
            ["field", "string", "People table field name, such as EmailAddress", True, None],
            ["value", "any", "New value", True, None]
        ],
        "returns": "None",
        "watch_out": ["HomePhone is written to the family, not the person.", "Changes the person record.", "Only HomePhone changes get a change log entry. Other fields are saved but not logged in the person's change history.", "The object you passed is not refreshed, so it still shows the old value after the call. Call model.GetPerson again to see the change.", "Text for a date field that cannot be read as a date clears the field."],
        "example": "p = model.GetPerson(model.UserPeopleId)\nprint \"Current work phone:\", p.WorkPhone\n# Writes to your own record. Uncomment to run.\n# model.UpdateField(p, \"WorkPhone\", \"6155551234\")",
        "see_also": ["model.UpdatePerson", "model.UpdateNamedField", "model.GetPerson"]
    },
    "model.UpdateMemberStatus": {
        "category": "People",
        "kind": "method",
        "arity": [2, 2],
        "tags": ["writes_data"],
        "summary": "Sets the member status for everyone in a query.",
        "details": "Pass a status id or a status description. Changes are logged. Accepts a PeopleId, a saved search name, or Search Builder code.",
        "flow": ["An integer or a string of digits is used as the status id as is. Any other text is looked up by exact description, and a new member status is created when none matches.", "Turns the query text into people: digits alone become one PeopleId, peopleids='1,2,3' becomes a list, text that exactly matches any saved search name runs that search, and anything else is read as Search Builder code. An empty string matches no one.", "Deceased and archived people are left out unless the code itself mentions them (IncludeDeceased, IncludeArchived, DeceasedDate or ArchivedFlag). A campus restriction or OrgLeadersOnly limit on the signed in user also narrows the list.", "For each person, sets MemberStatusId, writes a change log entry and saves."],
        "params": [
            ["query", "query", "People to act on: a PeopleId, a saved search name, or Search Builder code", True, None],
            ["status", "any", "Member status id (integer) or description (string)", True, None]
        ],
        "returns": "None",
        "watch_out": ["A description that does not exist yet creates a new member status.", "Changes person records.", "A status id is not checked against the member status list.", "Status 0, None or an empty description does nothing.", "Each person is loaded and saved one at a time, so a query that matches thousands of people is slow."],
        "example": "for s in q.QuerySql(\"SELECT Id, Description FROM lookup.MemberStatus\"):\n    print s.Id, s.Description\n# Uncomment to set yourself to Member (10):\n# model.UpdateMemberStatus(str(model.UserPeopleId), 10)",
        "see_also": ["model.UpdateCampus", "model.UpdateNamedField"]
    },
    "model.UpdateNamedField": {
        "category": "People",
        "kind": "method",
        "arity": [3, 3],
        "tags": ["writes_data"],
        "summary": "Sets one People table field to the same value for everyone in a query.",
        "details": "Changes are logged. Accepts a PeopleId, a saved search name, or Search Builder code.",
        "flow": ["Turns the query text into people: digits alone become one PeopleId, peopleids='1,2,3' becomes a list, text that exactly matches any saved search name runs that search, and anything else is read as Search Builder code. An empty string matches no one.", "Deceased and archived people are left out unless the code itself mentions them (IncludeDeceased, IncludeArchived, DeceasedDate or ArchivedFlag). A campus restriction or OrgLeadersOnly limit on the signed in user also narrows the list.", "For each person, sets the field. Text is converted to the field's type (number, yes or no, date), and an unchanged value is skipped.", "Writes a change log entry for each changed person and saves."],
        "params": [
            ["query", "query", "People to act on: a PeopleId, a saved search name, or Search Builder code", True, None],
            ["field", "string", "People table field name", True, None],
            ["value", "any", "New value", True, None]
        ],
        "returns": "None",
        "watch_out": ["Changes person records.", "Text for a date field that cannot be read as a date clears the field.", "This works on People fields only. HomePhone and address fields live on the family and are not changed here.", "Each person is loaded and saved one at a time, so a query that matches thousands of people is slow."],
        "example": "print q.QueryCount(\"{0}\".format(model.UserPeopleId)), \"people would change\"\n# Uncomment to run:\n# model.UpdateNamedField(str(model.UserPeopleId), \"DoNotMailFlag\", True)",
        "see_also": ["model.UpdateField", "model.UpdatePerson", "q.QueryCount"]
    },
    "model.UpdateNewMemberClassDate": {
        "category": "People",
        "kind": "method",
        "arity": [2, 2],
        "tags": ["writes_data"],
        "summary": "Sets the new member class date for everyone in a query.",
        "details": "Pass a .NET DateTime or a date string such as \"9/14/2026\". Changes are logged. Accepts a PeopleId, a saved search name, or Search Builder code.",
        "flow": ["Turns the query text into people: digits alone become one PeopleId, peopleids='1,2,3' becomes a list, text that exactly matches any saved search name runs that search, and anything else is read as Search Builder code. An empty string matches no one.", "Deceased and archived people are left out unless the code itself mentions them (IncludeDeceased, IncludeArchived, DeceasedDate or ArchivedFlag). A campus restriction or OrgLeadersOnly limit on the signed in user also narrows the list.", "For each person, sets NewMemberClassDate, writes a change log entry and saves."],
        "params": [
            ["query", "query", "People to act on: a PeopleId, a saved search name, or Search Builder code", True, None],
            ["dt", "date", "New member class date", True, None]
        ],
        "returns": "None",
        "watch_out": ["Changes person records.", "Text that cannot be read as a date clears the date.", "Each person is loaded and saved one at a time, so a query that matches thousands of people is slow."],
        "example": "from System import DateTime\nprint model.GetPerson(model.UserPeopleId).NewMemberClassDate\n# Uncomment to run:\n# model.UpdateNewMemberClassDate(str(model.UserPeopleId), DateTime(2026, 9, 14))",
        "see_also": ["model.UpdateNewMemberClassDateIfNullForLastAttended", "model.UpdateNewMemberClassStatus"]
    },
    "model.UpdateNewMemberClassDateIfNullForLastAttended": {
        "category": "People",
        "kind": "method",
        "arity": [2, 2],
        "tags": ["writes_data"],
        "summary": "Fills a missing new member class date with the date each person last attended the given involvement.",
        "details": "Only people in the query who have no new member class date yet are changed. Accepts a PeopleId, a saved search name, or Search Builder code.",
        "flow": ["Turns the query text into people: digits alone become one PeopleId, peopleids='1,2,3' becomes a list, text that exactly matches any saved search name runs that search, and anything else is read as Search Builder code. An empty string matches no one.", "Deceased and archived people are left out unless the code itself mentions them (IncludeDeceased, IncludeArchived, DeceasedDate or ArchivedFlag). A campus restriction or OrgLeadersOnly limit on the signed in user also narrows the list.", "Skips anyone who already has a new member class date.", "Finds the person's most recent meeting in the involvement where they were marked present.", "Sets NewMemberClassDate to that meeting date, writes a change log entry and saves."],
        "params": [
            ["query", "query", "People to act on: a PeopleId, a saved search name, or Search Builder code", True, None],
            ["orgId", "integer", "OrganizationId of the new member class", True, None]
        ],
        "returns": "None",
        "watch_out": ["People who never attended the involvement are skipped.", "Changes person records.", "Only attendance marked present counts. Being on the roll is not enough.", "The saved date includes the meeting time.", "Each person is loaded and saved one at a time, so a query that matches thousands of people is slow."],
        "example": "orgId = 30  # your new member class involvement\nprint q.QueryCount(\"{0}\".format(model.UserPeopleId)), \"people checked\"\n# Uncomment to run:\n# model.UpdateNewMemberClassDateIfNullForLastAttended(str(model.UserPeopleId), orgId)",
        "see_also": ["model.UpdateNewMemberClassDate", "model.UpdateNewMemberClassStatus"]
    },
    "model.UpdateNewMemberClassStatus": {
        "category": "People",
        "kind": "method",
        "arity": [2, 2],
        "tags": ["writes_data"],
        "summary": "Sets the new member class status for everyone in a query.",
        "details": "Pass a status id or a description. Changes are logged. Accepts a PeopleId, a saved search name, or Search Builder code.",
        "flow": ["An integer or a string of digits is used as the status id as is. Any other text is looked up by exact description, and a new status is created when none matches.", "Turns the query text into people: digits alone become one PeopleId, peopleids='1,2,3' becomes a list, text that exactly matches any saved search name runs that search, and anything else is read as Search Builder code. An empty string matches no one.", "Deceased and archived people are left out unless the code itself mentions them (IncludeDeceased, IncludeArchived, DeceasedDate or ArchivedFlag). A campus restriction or OrgLeadersOnly limit on the signed in user also narrows the list.", "For each person, sets NewMemberClassStatusId, writes a change log entry and saves."],
        "params": [
            ["query", "query", "People to act on: a PeopleId, a saved search name, or Search Builder code", True, None],
            ["status", "any", "Status id (integer) or description (string)", True, None]
        ],
        "returns": "None",
        "watch_out": ["A description that does not exist yet creates a new status.", "Changes person records.", "A status id is not checked against the status list.", "Status 0 does nothing.", "Each person is loaded and saved one at a time, so a query that matches thousands of people is slow."],
        "example": "for s in q.QuerySql(\"SELECT Id, Description FROM lookup.NewMemberClassStatus\"):\n    print s.Id, s.Description\n# Uncomment to run:\n# model.UpdateNewMemberClassStatus(str(model.UserPeopleId), \"Attended\")",
        "see_also": ["model.UpdateNewMemberClassDate", "model.UpdateMemberStatus"]
    },
    "model.UpdatePerson": {
        "category": "People",
        "kind": "method",
        "arity": [2, 2],
        "tags": ["writes_data"],
        "summary": "Updates several fields on one person in a single call.",
        "details": "Keys are People table field names such as EmailAddress or CellPhone. Changes are logged.",
        "flow": ["Loads the person by PeopleId and walks your dictionary. Field names are matched without regard to case.", "HomePhone, AddressLineOne, AddressLineTwo, CityName, StateCode, ZipCode and CountryName go to the person's Families row. Every other key goes to the People row.", "A value that equals what is already there is skipped. Text values are converted to the field's type.", "Writes one ChangeLog entry for the person (Basic Info) with a detail row per changed field, plus one Family entry when a family field changed, and an ActivityLog line.", "Saves everything in one save at the end."],
        "params": [
            ["peopleId", "integer", "PeopleId of the person to update", True, None],
            ["data", "dict", "Dictionary of field name to new value", True, None]
        ],
        "returns": "None",
        "watch_out": ["Family address fields (HomePhone, AddressLineOne, AddressLineTwo, CityName, StateCode, ZipCode, CountryName) are written to the person's family.", "Changes the person record.", "A misspelled field name with a text value is silently not saved but still shows in the change log. With a number or other non text value it raises an error.", "Values are not checked or reformatted. Phone numbers keep any dashes or spaces you pass, so strip them to digits first.", "Address keys always go to the family, so this call cannot set a person's own alternate address.", "An unknown PeopleId raises an error."],
        "example": "pid = model.UserPeopleId\nprint \"Current cell:\", model.GetPerson(pid).CellPhone\n# Writes to your own record. Uncomment to run.\n# model.UpdatePerson(pid, {\"CellPhone\": \"6155551234\"})",
        "see_also": ["model.UpdateField", "model.UpdateNamedField", "model.GetPerson"]
    },
    "model.CreateCustomView": {
        "category": "SQL Queries",
        "kind": "method",
        "arity": [2, 2],
        "tags": ["writes_data"],
        "summary": "Creates or replaces a SQL view named custom.<view> in the database.",
        "details": "Creates the custom schema if needed. The name must be a single alphanumeric word.",
        "flow": ["Returns without doing anything when run from the morning batch.", "Raises \"must be developer and admin\" unless the user has both roles, and raises an error if the name is not a single word starting with a letter.", "Creates the custom schema if it does not exist.", "Drops any existing custom.<view>, then creates it from your SELECT."],
        "params": [
            ["view", "string", "View name, letters and digits only", True, None],
            ["sql", "string", "SELECT statement that defines the view", True, None]
        ],
        "returns": "None",
        "watch_out": ["This changes the database schema.", "Needs both the Developer and Admin roles.", "Does nothing when run from the morning batch.", "The old view is dropped before the new one is created, so a SELECT with an error leaves you with no view at all.", "Changes take effect for every report and script that reads the view."],
        "example": "print \"Developer:\", model.UserIsInRole(\"Developer\"), \"Admin:\", model.UserIsInRole(\"Admin\")\n# Creates custom.ActiveAdults. Uncomment to run:\n# model.CreateCustomView(\"ActiveAdults\", \"SELECT PeopleId, Name FROM People WHERE Age >= 18 AND IsDeceased = 0\")",
        "see_also": ["model.UserIsInRole", "q.QuerySql"]
    },
    "model.ExecuteSql": {
        "category": "SQL Queries",
        "kind": "method",
        "arity": [1, 1],
        "tags": ["writes_data"],
        "summary": "Runs any SQL statement, including writes.",
        "params": [
            ["sql", "string", "SQL statement", True, None]
        ],
        "returns": "None",
        "watch_out": ["Works only on a developer's local debug build.", "On a live TouchPoint site it always raises an error.", "The error raised on a live site has no helpful message, only a generic exception."],
        "example": "try:\n    model.ExecuteSql(\"SELECT 1\")\n    print \"ran\"\nexcept Exception as e:\n    print \"Not available on this server\"",
        "see_also": ["q.QuerySql", "model.IsDebug"]
    },
    "model.SqlGrid": {
        "category": "SQL Queries",
        "kind": "method",
        "arity": [1, 1],
        "tags": ["read_only"],
        "summary": "Runs a read only SQL query and returns the results as an HTML table in TouchPoint's report styling.",
        "details": "If the SQL uses @\x42lueToolbarTagId and the script was run from the Blue Toolbar, that parameter is filled with the selected people.",
        "flow": ["If the SQL mentions @\x42lueToolbarTagId and the script was run from the Blue Toolbar, saves the selected people in a temporary tag and passes its id.", "Runs the SQL on the read only connection with a 300 second timeout. Only the first result set is used.", "Builds the table: money and decimal columns show as currency, columns whose names start or end with pct or percent show as percentages, PeopleId and SpouseId link to the person, OrganizationId and columns ending in OrgId link to the involvement, and whole numbers get thousands separators except id, year and number columns.", "A column named linkfornext turns the next cell into a link to its value. Cells reading Total, Total: or Grand Total are bolded. A footer shows the row count."],
        "params": [
            ["sql", "string", "SELECT statement", True, None]
        ],
        "returns": "string (HTML)",
        "watch_out": ["Timeout is 300 seconds.", "If the SQL uses @\x42lueToolbarTagId but the script was not run from the Blue Toolbar, the parameter is never set and SQL Server raises a \"Must declare the scalar variable\" error.", "A decimal column that is not named like pct or percent is shown as dollars, so cast averages to float if they are not money.", "Cell values are not HTML escaped, so text containing < or & shows as HTML."],
        "example": "sql = \"\"\"\nSELECT TOP 10 p.PeopleId, p.Name, p.Age\nFROM People p\nWHERE p.IsDeceased = 0\nORDER BY p.Name\n\"\"\"\nprint model.SqlGrid(sql)",
        "see_also": ["model.SqlList", "q.QuerySql"]
    },
    "model.SqlList": {
        "category": "SQL Queries",
        "kind": "method",
        "arity": [1, 1],
        "tags": ["read_only"],
        "summary": "Runs a read only SQL query and returns all rows as a list.",
        "details": "Read columns as attributes. Same rows as q.QuerySql, but as a list so you can use len() and indexes.",
        "params": [
            ["sql", "string", "SELECT statement", True, None]
        ],
        "returns": "list of rows",
        "watch_out": ["Like q.QuerySql it fills in @UserPeopleId, @CurrentOrgId and @\x42lueToolbarTagId, and allows up to 5 minutes.", "Every row is loaded into memory at once, so use TOP on large tables.", "Users with the Finance role run on a different read only login than other users, so a query on giving tables can work for one user and fail for another."],
        "example": "rows = model.SqlList(\"SELECT TOP 5 PeopleId, Name FROM People ORDER BY PeopleId\")\nprint len(rows), \"rows\"\nfor r in rows:\n    print r.PeopleId, r.Name",
        "see_also": ["q.QuerySql", "model.SqlListDynamicData", "model.SqlGrid"]
    },
    "model.SqlListDynamicData": {
        "category": "SQL Queries",
        "kind": "method",
        "arity": [1, 2],
        "tags": ["read_only"],
        "summary": "Runs a read only SQL query and returns each row as a DynamicData object.",
        "details": "If a column name contains \"Json\", the first such column is parsed into a nested DynamicData. The optional metadata maps JSON keys to types for conversion.",
        "flow": ["Runs the SQL like q.QuerySql, filling in @UserPeopleId, @CurrentOrgId and @\x42lueToolbarTagId.", "Turns each row into a DynamicData.", "Finds the first column whose name contains \"json\" (any case) and parses its text as JSON into a nested DynamicData.", "When metadata is given, converts the listed JSON keys: \"money\" to a decimal and \"int\" to an integer. Other type names are ignored."],
        "params": [
            ["sql", "string", "SELECT statement", True, None],
            ["metadata", "DynamicData", "Type hints for the JSON column", False, "None"]
        ],
        "returns": "list of DynamicData",
        "watch_out": ["Text in the Json column that is not valid JSON raises an error.", "Every row is loaded into memory at once."],
        "example": "rows = model.SqlListDynamicData(\"SELECT TOP 5 PeopleId, Name FROM People ORDER BY PeopleId\")\nfor r in rows:\n    print r.PeopleId, r.Name",
        "see_also": ["model.SqlTop1DynamicData", "model.SqlList", "model.DynamicDataFromJson"]
    },
    "model.SqlTop1DynamicData": {
        "category": "SQL Queries",
        "kind": "method",
        "arity": [1, 2],
        "tags": ["read_only"],
        "summary": "Runs a read only SQL query and returns the first row as a DynamicData object.",
        "details": "A column whose name contains \"Json\" is parsed into a nested DynamicData.",
        "params": [
            ["sql", "string", "SELECT statement", True, None],
            ["metadata", "DynamicData", "Type hints for the JSON column", False, "None"]
        ],
        "returns": "DynamicData or None",
        "watch_out": ["Returns None if there are no rows.", "The full result is read before the first row is taken, so add TOP 1 to the SQL.", "Text in a column whose name contains \"json\" must be valid JSON or the call raises an error.", "It fills in @UserPeopleId, @CurrentOrgId and @\x42lueToolbarTagId like q.QuerySql."],
        "example": "p = model.SqlTop1DynamicData(\"SELECT PeopleId, Name, Age FROM People WHERE PeopleId = {0}\".format(model.UserPeopleId))\nif p:\n    print p.Name, p.Age",
        "see_also": ["model.SqlListDynamicData", "q.QuerySqlTop1"]
    },
    "q.QuerySql": {
        "category": "SQL Queries",
        "kind": "method",
        "arity": [1, 3],
        "tags": ["read_only"],
        "summary": "Runs a SQL query on a read-only connection and returns the rows.",
        "details": "Each row exposes its columns as attributes (row.Name). Pass a single value as the second argument to fill @p1, or a dict or DynamicData to fill named parameters. TouchPoint tries to fill in @UserPeopleId, @CurrentOrgId and @\x42lueToolbarTagId when the SQL uses them, but in a script the first two usually come back blank (see Watch out).",
        "flow": ["Opens a read-only database connection. People with the Finance role get the finance read-only login; everyone else gets the standard read-only login.", "With a second argument that is a dict or DynamicData, adds each key as a named parameter. Any other value becomes @p1. In the three argument form, a line starting with declare @p1 is commented out so the same SQL can be tested in SQL Server Management Studio.", "If the SQL text contains @UserPeopleId it is filled from Data.PeopleId, falling back to the signed in user. @CurrentOrgId is filled the same way from Data.OrgId, falling back to the current involvement.", "If the SQL contains @\x42lueToolbarTagId and the script was run from the Blue Toolbar, it puts up to the first 1000 selected people in a new temporary tag and passes that tag's id.", "Runs the query with a 300 second timeout and returns every row. Each row exposes its columns as attributes."],
        "params": [
            ["sql", "string", "The SQL to run", True, None],
            ["p1", "any", "Value for @p1", False, "None"],
            ["declarations", "dict", "Dict or DynamicData of named parameters (without the @)", False, "None"]
        ],
        "returns": "list of rows",
        "watch_out": ["In a script, @UserPeopleId and @CurrentOrgId come back blank, not you and your current involvement. Set model.Data.PeopleId = model.UserPeopleId (and model.Data.OrgId) before the query, or pass your own value: q.QuerySql(sql, None, {'me': model.UserPeopleId}).", "Timeout is 300 seconds.", "The automatic parameters are matched by exact spelling and case. @userpeopleid is not filled and the query fails with Must declare the scalar variable.", "Data returns an empty string, not None, for a value that was not passed. When Data has no PeopleId, @UserPeopleId gets that empty value instead of the signed in user, and it matches nobody. If the page was opened with ?PeopleId=123, @UserPeopleId becomes 123. The same applies to @CurrentOrgId and Data.OrgId. Pass model.UserPeopleId yourself as a named parameter when it matters.", "@\x42lueToolbarTagId only sees the first 1000 selected people.", "Run outside the Blue Toolbar, @\x42lueToolbarTagId is not filled and the query fails.", "Only the declare @p1 line, and only in the three argument form, is commented out for you. A DECLARE line for any other parameter you pass makes the query fail. Remove it before saving the script.", "Write statements fail because the connection is read-only. Temporary tables (#name) still work."],
        "example": "rows = q.QuerySql(\"SELECT TOP 5 PeopleId, Name FROM People WHERE LastName = @p1\", \"Smith\")\nfor r in rows:\n    print r.PeopleId, r.Name\nrows = q.QuerySql(\"SELECT Name FROM People WHERE PeopleId = @pid\", {\"pid\": model.UserPeopleId})\nfor r in rows:\n    print r.Name",
        "signatures": ["q.QuerySql(sql)", "q.QuerySql(sql, p1_or_declarations)", "q.QuerySql(sql, p1, declarations)"],
        "see_also": ["q.QuerySqlTop1", "model.SqlList", "q.QuerySqlInt"]
    },
    "q.QuerySqlInt": {
        "category": "SQL Queries",
        "kind": "method",
        "arity": [1, 1],
        "tags": ["read_only"],
        "summary": "Runs a SQL query on a read-only connection and returns the first column of the first row as an integer.",
        "params": [
            ["sql", "string", "SQL that returns a single number", True, None]
        ],
        "returns": "integer",
        "watch_out": ["Returns 0 when the result is NULL.", "It takes no parameters, and @UserPeopleId and the other automatic parameters are not filled in.", "Returns 0 when the query returns no rows, not only when the value is NULL.", "Uses the standard 30 second SQL timeout, unlike q.QuerySql which allows 5 minutes."],
        "example": "count = q.QuerySqlInt(\"SELECT COUNT(*) FROM People WHERE IsDeceased = 0 AND ArchivedFlag = 0\")\nprint \"Living people:\", count",
        "see_also": ["q.QuerySqlInts", "q.QuerySqlStr", "q.QuerySql"]
    },
    "q.QuerySqlInts": {
        "category": "SQL Queries",
        "kind": "method",
        "arity": [1, 1],
        "tags": ["read_only"],
        "summary": "Runs read only SQL and returns the first column as a list of integers.",
        "params": [
            ["sql", "string", "SQL returning one integer column", True, None]
        ],
        "returns": "List of integers",
        "watch_out": ["@UserPeopleId, @CurrentOrgId and @\x42lueToolbarTagId are not filled in.", "Uses the standard 30 second SQL timeout, unlike q.QuerySql which allows 5 minutes."],
        "example": "ids = list(q.QuerySqlInts(\"SELECT TOP 10 OrganizationId FROM Organizations WHERE OrganizationStatusId = 30\"))\nprint ids",
        "see_also": ["q.QuerySqlPeopleIds", "q.QuerySqlInt", "q.QuerySql"]
    },
    "q.QuerySqlJsonArray": {
        "category": "SQL Queries",
        "kind": "method",
        "arity": [1, 1],
        "tags": ["read_only"],
        "summary": "Runs a read-only SQL query and returns the rows as a JSON array of arrays, one inner array per row in column order.",
        "details": "The output is indented. Dates come out as ISO strings. Good for feeding a chart or a JavaScript table directly.",
        "params": [
            ["sql", "string", "SQL to run", True, None]
        ],
        "returns": "string (JSON)",
        "watch_out": ["Column names are not included.", "Parameters such as @UserPeopleId are not filled in.", "Returns \"[]\" when there are no rows. NULL values come out as null.", "Uses the standard 30 second SQL timeout, unlike q.QuerySql which allows 5 minutes."],
        "example": "data = q.QuerySqlJsonArray('SELECT TOP 5 PeopleId, Name FROM People WHERE IsDeceased = 0 ORDER BY PeopleId DESC')\nprint '<pre>' + data + '</pre>'",
        "see_also": ["q.SqlDateValueArray", "q.SqlNameCountArray", "q.QuerySql"]
    },
    "q.QuerySqlPeopleIds": {
        "category": "SQL Queries",
        "kind": "method",
        "arity": [1, 1],
        "tags": ["read_only"],
        "summary": "Runs read only SQL and returns the first column as a list of integers.",
        "details": "Same as q.QuerySqlInts.",
        "params": [
            ["sql", "string", "SQL returning one integer column of PeopleIds", True, None]
        ],
        "returns": "List of integers",
        "watch_out": ["@UserPeopleId, @CurrentOrgId and @\x42lueToolbarTagId are not filled in.", "Uses the standard 30 second SQL timeout, unlike q.QuerySql which allows 5 minutes."],
        "example": "ids = list(q.QuerySqlPeopleIds(\"SELECT TOP 10 PeopleId FROM People WHERE Age >= 90\"))\nprint ids",
        "see_also": ["q.QuerySqlInts", "q.SqlPeopleIdsToQuery", "q.TagSqlPeopleIds"]
    },
    "q.QuerySqlScalar": {
        "category": "SQL Queries",
        "kind": "method",
        "arity": [1, 1],
        "tags": ["read_only"],
        "summary": "Runs read only SQL and returns the first column of the first row as a string.",
        "params": [
            ["sql", "string", "SQL returning one value", True, None]
        ],
        "returns": "String",
        "watch_out": ["Raises an error when the SQL returns no rows. A NULL value comes back as an empty string.", "A NULL value returns an empty string, not an error.", "@UserPeopleId, @CurrentOrgId and @\x42lueToolbarTagId are not filled in.", "Uses the standard 30 second SQL timeout, unlike q.QuerySql which allows 5 minutes."],
        "example": "name = q.QuerySqlScalar(\"SELECT Name FROM People WHERE PeopleId = {0}\".format(model.UserPeopleId))\nprint name",
        "see_also": ["q.QuerySqlStr", "q.QuerySqlInt", "q.QuerySqlTop1"]
    },
    "q.QuerySqlStr": {
        "category": "SQL Queries",
        "kind": "method",
        "arity": [1, 1],
        "tags": ["read_only"],
        "summary": "Runs a read-only SQL query and returns the first column of the first row as a string.",
        "details": "q.QuerySqlScalar is similar but fails when the result is empty.",
        "params": [
            ["sql", "string", "SQL to run", True, None]
        ],
        "returns": "string or None",
        "watch_out": ["Returns None when there are no rows.", "It does not fill in @UserPeopleId, @CurrentOrgId or other parameters, so put values directly in the SQL.", "Returns None for a NULL value as well as for no rows.", "Uses the standard 30 second SQL timeout, unlike q.QuerySql which allows 5 minutes."],
        "example": "name = q.QuerySqlStr('SELECT Name FROM People WHERE PeopleId = {0}'.format(model.UserPeopleId))\nprint name or 'Not found'",
        "see_also": ["q.QuerySqlScalar", "q.QuerySqlInt", "q.QuerySqlTop1"]
    },
    "q.QuerySqlTop1": {
        "category": "SQL Queries",
        "kind": "method",
        "arity": [1, 3],
        "tags": ["read_only"],
        "summary": "Runs a SQL query like q.QuerySql and returns only the first row.",
        "details": "Supports the same @p1 and named parameters.",
        "params": [
            ["sql", "string", "The SQL to run", True, None],
            ["p1", "any", "Value for @p1", False, "None"],
            ["declarations", "dict", "Dict or DynamicData of named parameters", False, "None"]
        ],
        "returns": "row or None",
        "watch_out": ["Returns None when there are no rows.", "The full result is read before the first row is taken, so add TOP 1 to the SQL for speed.", "Passing None as p1 does not define @p1, so SQL that uses @p1 then fails.", "A line starting with \"declare @p1\" is commented out for you, so SQL tested in SQL Server Management Studio runs unchanged.", "It fills in @UserPeopleId, @CurrentOrgId and @\x42lueToolbarTagId, and allows up to 5 minutes."],
        "example": "p = q.QuerySqlTop1(\"SELECT Name, EmailAddress FROM People WHERE PeopleId = @p1\", model.UserPeopleId)\nif p:\n    print p.Name, p.EmailAddress",
        "signatures": ["q.QuerySqlTop1(sql)", "q.QuerySqlTop1(sql, p1)", "q.QuerySqlTop1(sql, p1, declarations)"],
        "see_also": ["q.QuerySql", "model.SqlTop1DynamicData", "q.QuerySqlStr"]
    },
    "q.SqlDateValueArray": {
        "category": "SQL Queries",
        "kind": "method",
        "arity": [3, 3],
        "tags": ["read_only"],
        "summary": "Builds a Google Charts data table from SQL with a Date column followed by 1 to 10 value columns named Val1, Val2 and so on.",
        "details": "header is the header row without the outer brackets. Paste the result into script code and pass it to google.visualization.arrayToDataTable. @\x71tagid in the SQL is replaced with the user's last search.",
        "flow": ["If the SQL contains @\x71tagid, tags everyone in the user's most recent search and adds DECLARE @\x71tagid at the top of the SQL with that tag's id.", "Runs the SQL on the read only connection.", "Writes each row as [new Date(year, month, day), Val1, ... ValN], with the month counted from 0 the way JavaScript expects.", "Puts your header row first and returns the whole array as text."],
        "params": [
            ["header", "string", "Header row, for example \"{label: 'Date', type: 'date'}, 'Meetings'\"", True, None],
            ["numvalcols", "integer", "How many Val columns follow the Date column (1 to 10)", True, None],
            ["sql", "string", "SQL returning Date, Val1 ... ValN", True, None]
        ],
        "returns": "string (JavaScript array literal)",
        "watch_out": ["The result contains JavaScript new Date(...) calls, so it is not JSON.", "When the query returns no rows a sample data set is returned instead.", "The date column must be named Date and the values Val1, Val2 and so on.", "A NULL value prints as nothing, which breaks the JavaScript. Use ISNULL in the SQL.", "Using @\x71tagid saves tag rows for the user's last search, so it is not purely read only.", "A numvalcols outside 1 to 10 writes only Val1."],
        "example": "sql = '''SELECT CAST(MeetingDate AS DATE) AS Date, COUNT(*) AS Val1, SUM(NumPresent) AS Val2\nFROM Meetings WHERE MeetingDate >= DATEADD(day, -30, GETDATE())\nGROUP BY CAST(MeetingDate AS DATE) ORDER BY 1'''\nhdr = \"{label: 'Date', type: 'date'}, 'Meetings', 'Present'\"\ndata = q.SqlDateValueArray(hdr, 2, sql)\nprint '<pre>' + data[:300] + '</pre>'",
        "see_also": ["q.SqlNameCountArray", "q.QuerySqlJsonArray"]
    },
    "q.SqlFirstColumnRowKey": {
        "category": "SQL Queries",
        "kind": "method",
        "arity": [2, 2],
        "tags": ["read_only"],
        "summary": "Runs a read only SQL query and returns a DynamicData keyed by the first column's value.",
        "details": "Each key holds another DynamicData with that row's columns. Datetime columns come back as formatted strings. Read rows by key with data.GetValue(\"Key\") or data.Key.",
        "params": [
            ["sql", "string", "SELECT statement; first column is the key", True, None],
            ["declarations", "dict", "Parameter values as a dict (names without @), or None. Required", True, None]
        ],
        "returns": "DynamicData",
        "watch_out": ["The first column must be text (use CAST for numbers) and should be unique.", "Stops after 100 rows.", "The declarations argument is required; pass None if there are no parameters.", "data.Keys() does not list the rows, because Keys() skips values that are rows.", "When two rows have the same first column value, the later row replaces the earlier one.", "Only columns of type datetime are formatted. date and datetime2 columns come back as raw values.", "NULL values in other columns come back as DBNull, not None.", "@UserPeopleId, @CurrentOrgId and @\x42lueToolbarTagId are not filled in."],
        "example": "sql = \"\"\"\nSELECT CAST(ms.Description AS VARCHAR(50)) AS Status, COUNT(*) AS Cnt\nFROM People p JOIN lookup.MemberStatus ms ON ms.Id = p.MemberStatusId\nGROUP BY ms.Description\n\"\"\"\ndata = q.SqlFirstColumnRowKey(sql, None)\n# Look rows up by their first-column value. data.Keys() skips row values, so it lists nothing here.\nrow = data.GetValue(\"Member\")\nprint \"Member:\", row.Cnt if row else \"(no such status)\"",
        "see_also": ["q.SqlNameValues", "DynamicData.GetValue", "DynamicData.Keys"]
    },
    "q.SqlNameCountArray": {
        "category": "SQL Queries",
        "kind": "method",
        "arity": [2, 2],
        "tags": ["read_only"],
        "summary": "Runs read only SQL with Name and Cnt columns and formats the rows as a Google Charts data array string.",
        "details": "If the SQL uses @\x71tagid, it is filled with the people from your last search.",
        "flow": ["If the SQL contains @\x71tagid, tags everyone in the user's most recent search and declares @\x71tagid with that tag's id.", "Runs the SQL on the read only connection and reads the Name and Cnt columns.", "Returns a JavaScript array with your title row followed by one ['Name', Cnt] row each."],
        "params": [
            ["title", "string", "Header for the name column", True, None],
            ["sql", "string", "SQL returning Name and Cnt columns", True, None]
        ],
        "returns": "String JavaScript array",
        "watch_out": ["Returns placeholder data when there are no rows.", "Cnt must be a whole number and must not be NULL.", "Names are not escaped, so a name with an apostrophe breaks the JavaScript.", "Using @\x71tagid saves tag rows for the user's last search."],
        "example": "sql = \"\"\"SELECT ms.Description AS Name, COUNT(*) AS Cnt\nFROM People p JOIN lookup.MemberStatus ms ON ms.Id = p.MemberStatusId\nGROUP BY ms.Description\"\"\"\nprint q.SqlNameCountArray(\"Member Status\", sql)",
        "see_also": ["q.SqlDateValueArray", "q.QuerySqlJsonArray"]
    },
    "q.SqlNameValues": {
        "category": "SQL Queries",
        "kind": "method",
        "arity": [3, 3],
        "tags": ["read_only"],
        "summary": "Runs read only SQL and builds a DynamicData object where each row's name column becomes a property holding the value column.",
        "params": [
            ["sql", "string", "SQL query", True, None],
            ["NameCol", "string", "Column holding property names", True, None],
            ["ValueCol", "string", "Column holding values", True, None]
        ],
        "returns": "DynamicData object",
        "watch_out": ["Names must be unique.", "A repeated name raises an error rather than overwriting.", "A NULL value comes back as DBNull, not None.", "@UserPeopleId, @CurrentOrgId and @\x42lueToolbarTagId are not filled in."],
        "example": "sql = \"SELECT ms.Description AS Name, COUNT(*) AS Cnt FROM People p JOIN lookup.MemberStatus ms ON ms.Id = p.MemberStatusId GROUP BY ms.Description\"\nd = q.SqlNameValues(sql, \"Name\", \"Cnt\")\nprint d.Member",
        "see_also": ["q.SqlFirstColumnRowKey", "model.DynamicData"]
    },
    "model.AddTag": {
        "category": "Searches & Tags",
        "kind": "method",
        "arity": [3, 4],
        "tags": ["writes_data"],
        "summary": "Adds everyone in a query to a personal tag owned by ownerId, creating the tag if needed.",
        "details": "Set clear to True to empty the tag first.",
        "flow": ["Finds the personal tag with that name and owner, or creates it.", "When clear is True, removes everyone already in the tag.", "Turns the query text into people: digits alone become one PeopleId, peopleids='1,2,3' becomes a list, text that exactly matches any saved search name runs that search, and anything else is read as Search Builder code. An empty string matches no one. Deceased and archived people are left out unless the code itself mentions them (IncludeDeceased, IncludeArchived, DeceasedDate or ArchivedFlag). A campus restriction or OrgLeadersOnly limit on the signed in user also narrows the list.", "Adds each person not already in the tag, then saves."],
        "params": [
            ["query", "query", "People to act on: a PeopleId, a saved search name, or Search Builder code", True, None],
            ["tagName", "string", "Name of the tag", True, None],
            ["ownerId", "integer", "PeopleId who owns the tag", True, None],
            ["clear", "boolean", "Remove everyone already in the tag first", False, "False"]
        ],
        "returns": "None",
        "watch_out": ["Writes data.", "Each person is checked and added one at a time, so a large query is slow.", "Avoid ! in the tag name. It is saved as *, so the next lookup by the original name does not find it and a new tag is created."],
        "example": "me = model.UserPeopleId\nmodel.AddTag(\"Age >= 65\", \"Seniors Test\", me, True)\nprint \"Tagged\", q.QueryCount(\"Age >= 65\"), \"people in your tag 'Seniors Test'\"",
        "see_also": ["model.ClearTag", "model.CreateQueryTag", "q.TagCount"]
    },
    "model.ClearTag": {
        "category": "Searches & Tags",
        "kind": "method",
        "arity": [2, 2],
        "tags": ["writes_data"],
        "summary": "Removes everyone from a personal tag owned by ownerId.",
        "details": "The tag itself remains.",
        "params": [
            ["tagName", "string", "Name of the tag", True, None],
            ["ownerId", "integer", "PeopleId who owns the tag", True, None]
        ],
        "returns": "None",
        "watch_out": ["Writes data.", "Does nothing if the tag does not exist.", "Only personal tags are searched, so a Query Tag with that name is not cleared."],
        "example": "me = model.UserPeopleId\nmodel.ClearTag(\"Seniors Test\", me)\nprint \"Cleared your tag 'Seniors Test'\"",
        "see_also": ["model.AddTag", "model.DeleteQueryTags"]
    },
    "model.CreateQueryTag": {
        "category": "Searches & Tags",
        "kind": "method",
        "arity": [2, 2],
        "tags": ["writes_data"],
        "summary": "Saves everyone matched by a query into a named Query Tag and returns how many people are in it.",
        "details": "The tag's previous members are replaced.",
        "flow": ["Turns the query text into people: digits alone become one PeopleId, peopleids='1,2,3' becomes a list, text that exactly matches any saved search name runs that search, and anything else is read as Search Builder code. An empty string matches no one. Deceased and archived people are left out unless the code itself mentions them (IncludeDeceased, IncludeArchived, DeceasedDate or ArchivedFlag). A campus restriction or OrgLeadersOnly limit on the signed in user also narrows the list.", "If model.QueryTagLimit is above 0, keeps only that many of the matches.", "Finds or creates a Query Tag with this name owned by the signed in user, removes its current members and updates its created date.", "Adds the matched people and returns the number of people now in the tag."],
        "params": [
            ["name", "string", "Query Tag name", True, None],
            ["code", "query", "Search Builder code, saved search name, or PeopleId", True, None]
        ],
        "returns": "integer (number of people in the tag)",
        "watch_out": ["Writes data (tag rows).", "If model.QueryTagLimit is set, at most that many people are tagged.", "Which people are kept under QueryTagLimit is not defined, since no sort order is applied.", "The tag belongs to the user running the script, so two users running the same script get separate tags of the same name."],
        "example": "count = model.CreateQueryTag(\"Test_Members\", \"MemberStatusId = 10[Member]\")\nprint count, \"people tagged\"\nmodel.DeleteQueryTags(\"Test_%\")",
        "see_also": ["model.DeleteQueryTags", "model.QueryTagLimit", "model.TagLastQuery", "model.AddTag"]
    },
    "model.DeleteQueryTags": {
        "category": "Searches & Tags",
        "kind": "method",
        "arity": [1, 1],
        "tags": ["writes_data"],
        "summary": "Deletes every Query Tag whose name matches a SQL LIKE pattern, along with its people.",
        "details": "Use % as the wildcard.",
        "flow": ["Deletes the people in every Query Tag whose name matches the pattern, then deletes those tags.", "Resets the user's current tag to UnNamed."],
        "params": [
            ["namelike", "string", "Name pattern with % wildcards", True, None]
        ],
        "returns": "None",
        "watch_out": ["Deletes data.", "It matches Query Tags owned by anyone, not only your own.", "Personal tags and other tag types are never touched."],
        "example": "model.CreateQueryTag(\"Test_Temp\", \"PeopleId = {0}\".format(model.UserPeopleId))\nmodel.DeleteQueryTags(\"Test_%\")\nprint \"Deleted Query Tags named Test_...\"",
        "see_also": ["model.CreateQueryTag"]
    },
    "model.OrgMembersQuery": {
        "category": "Searches & Tags",
        "kind": "method",
        "arity": [4, 4],
        "tags": ["writes_data", "deprecated"],
        "summary": "Builds a scratchpad search for members of a program, division or involvement and returns its query GUID.",
        "details": "memberTypes is a comma separated list of member type descriptions, not codes.",
        "flow": ["Picks one of the user's scratchpad searches: a new one when the user has fewer than 5, otherwise the oldest one, which is cleared.", "Looks up the member types whose descriptions are in your list.", "Adds one member type condition for the program, division and involvement given, and saves the search.", "Returns the search's id."],
        "params": [
            ["progid", "integer", "Program id, or 0", True, None],
            ["divid", "integer", "Division id, or 0", True, None],
            ["orgid", "integer", "OrganizationId, or 0", True, None],
            ["memberTypes", "string", "Comma separated member type descriptions, such as \"Member,Leader\"", True, None]
        ],
        "returns": "Query GUID",
        "watch_out": ["Deprecated. Use Search Builder code or q.QuerySql instead.", "Saves a scratchpad search.", "A description that is not in the member type list is silently skipped.", "Spaces after commas are kept, so \"Member, Leader\" does not find Leader."],
        "example": "# Deprecated. A direct SQL replacement:\noid = q.QuerySqlInt(\"SELECT TOP 1 OrganizationId FROM Organizations WHERE OrganizationStatusId = 30\")\nn = q.QuerySqlInt(\"SELECT COUNT(*) FROM OrganizationMembers WHERE OrganizationId = {0} AND MemberTypeId = 220\".format(oid))\nprint n, \"members in involvement\", oid",
        "see_also": ["q.QuerySql", "q.QueryList", "q.QueryCount"]
    },
    "model.PeopleIds": {
        "category": "Searches & Tags",
        "kind": "method",
        "arity": [1, 1],
        "tags": ["read_only"],
        "summary": "Returns the PeopleIds matched by a query.",
        "details": "The query can be a PeopleId, a saved search name, or Search Builder code.",
        "flow": ["Turns the query text into people: digits alone become one PeopleId, peopleids='1,2,3' becomes a list, text that exactly matches any saved search name runs that search, and anything else is read as Search Builder code. An empty string matches no one.", "Deceased and archived people are left out unless the code itself mentions them (IncludeDeceased, IncludeArchived, DeceasedDate or ArchivedFlag). A campus restriction or OrgLeadersOnly limit on the signed in user also narrows the list.", "Returns the matching PeopleIds as a list."],
        "params": [
            ["query", "query", "PeopleId, saved search name, or Search Builder code", True, None]
        ],
        "returns": "List of integers",
        "watch_out": ["A PeopleId of an archived or deceased person is not returned unless the code mentions those people."],
        "example": "pids = model.PeopleIds(\"PeopleId = {0}\".format(model.UserPeopleId))\nprint len(pids), \"matched:\", pids",
        "see_also": ["q.QueryPeopleIds", "q.QueryList", "q.QueryCount"]
    },
    "model.QueryTagLimit": {
        "category": "Searches & Tags",
        "kind": "property",
        "summary": "Caps how many people model.CreateQueryTag puts in each tag.",
        "details": "Leave it as None (the default) or 0 for no limit. It can be set.",
        "returns": "integer or None (settable)",
        "watch_out": ["It affects only CreateQueryTag calls made later in the same script.", "Which people are kept under the limit is not defined, since no sort order is applied."],
        "example": "model.QueryTagLimit = 50\ncount = model.CreateQueryTag('TEST_Members50', 'IsMember = 1[True]')\nprint 'Tagged %d people (limit 50)' % count\nmodel.DeleteQueryTags('TEST_Members50')",
        "see_also": ["model.CreateQueryTag", "model.DeleteQueryTags"]
    },
    "model.ReplaceQueryFromCode": {
        "category": "Searches & Tags",
        "kind": "method",
        "arity": [2, 2],
        "tags": ["writes_data"],
        "summary": "Overwrites a saved search with new Search Builder code.",
        "flow": ["Loads the saved search by its id.", "Parses the new Search Builder code and saves it as the search's text. The name and owner stay the same."],
        "params": [
            ["encodedguid", "string", "Saved search id, such as the Id from model.StatusFlagList", True, None],
            ["code", "string", "New Search Builder code", True, None]
        ],
        "returns": "None",
        "watch_out": ["Writes data.", "Changes the saved search for everyone who uses it, including any status flag built on it.", "An id that does not match a saved search raises an error, and so does code that cannot be parsed.", "Status flag membership is not refreshed until the flags are next updated, for example by model.UpdateStatusFlags."],
        "example": "for f in model.StatusFlagList():\n    print f.Flag, f.Id\n# Overwrites a saved search. Uncomment to run:\n# model.ReplaceQueryFromCode(\"<saved search id>\", \"Age >= 50\")",
        "see_also": ["model.StatusFlagList", "model.UpdateStatusFlags"]
    },
    "model.StatusFlagDictionary": {
        "category": "Searches & Tags",
        "kind": "method",
        "arity": [0, 1],
        "tags": ["read_only"],
        "summary": "Returns status flags as a dictionary keyed by flag code (such as F01).",
        "details": "Pass a comma separated list of flag codes to return only those. Each value has Flag, Desc, Id and Code.",
        "params": [
            ["flags", "string", "Comma separated flag codes such as \"F01,F02\"", False, "None"]
        ],
        "returns": "Dictionary of flag code to StatusFlag",
        "watch_out": ["Spaces in the flags list are kept, so pass \"F01,F02\" not \"F01, F02\".", "Two saved searches with the same flag code raise an error."],
        "example": "flags = model.StatusFlagDictionary()\nfor key in flags.Keys:\n    print key, flags[key].Desc",
        "see_also": ["model.StatusFlagList", "model.UpdateStatusFlags"]
    },
    "model.StatusFlagList": {
        "category": "Searches & Tags",
        "kind": "method",
        "arity": [0, 0],
        "tags": ["read_only"],
        "summary": "Returns every status flag, meaning saved searches named like \"F01:Description\".",
        "details": "Each item has Flag (such as F01), Desc, Id (the saved search id) and Code (the Search Builder code).",
        "returns": "list of StatusFlag (Flag, Desc, Id, Code)",
        "watch_out": ["Code lines are indented with a tab.", "Sorting is by text, so F10 comes before F2."],
        "example": "for f in model.StatusFlagList():\n    print f.Flag, f.Desc",
        "see_also": ["model.StatusFlagDictionary", "model.UpdateStatusFlags", "model.ReplaceQueryFromCode"]
    },
    "model.TagLastQuery": {
        "category": "Searches & Tags",
        "kind": "method",
        "arity": [1, 1],
        "tags": ["writes_data"],
        "summary": "Saves people into a query tag and returns the tag id.",
        "details": "Interactively it tags the results of the user's most recent Search Builder search. From the morning batch there is no last search, so it runs defaultcode instead.",
        "flow": ["From the morning batch, runs defaultcode as the search.", "Otherwise, uses the signed in user's most recently run search.", "Finds or creates a tag named for the user's session, removes its current members, then adds everyone the search matches.", "Returns the tag id."],
        "params": [
            ["defaultcode", "string", "Search Builder code used only when run from the morning batch", True, None]
        ],
        "returns": "integer (TagId)",
        "watch_out": ["Writes data (tag rows).", "Calling it again in the same session overwrites the same tag.", "The last search is whichever search the user ran most recently, which may be a saved search and not the scratchpad."],
        "example": "tagId = model.TagLastQuery(\"PeopleId = {0}\".format(model.UserPeopleId))\nprint \"TagId:\", tagId\nprint q.QuerySqlInt(\"SELECT COUNT(*) FROM TagPerson WHERE Id = %d\" % tagId), \"people\"",
        "see_also": ["model.CreateQueryTag", "model.AddTag", "q.TagCount"]
    },
    "model.UpdateStatusFlag": {
        "category": "Searches & Tags",
        "kind": "method",
        "arity": [2, 2],
        "tags": ["writes_data", "known_issue"],
        "summary": "Meant to refresh one status flag from a saved search.",
        "details": "Use model.UpdateStatusFlags() to refresh all flags instead.",
        "params": [
            ["flagid", "string", "Status flag id", True, None],
            ["encodedguid", "string", "Saved search id (the Id from model.StatusFlagList)", True, None]
        ],
        "returns": "None",
        "watch_out": ["It calls dbo.UpdateStatusFlag with 2 arguments, but the stored procedure now needs 3 (@sfid, @name, @tagid).", "Because of that it raises a SQL error."],
        "example": "for f in model.StatusFlagList():\n    print f.Flag, f.Desc\n# Broken in current TouchPoint. Refresh all flags instead:\n# model.UpdateStatusFlags()",
        "see_also": ["model.UpdateStatusFlags", "model.StatusFlagList"]
    },
    "model.UpdateStatusFlags": {
        "category": "Searches & Tags",
        "kind": "method",
        "arity": [0, 0],
        "tags": ["writes_data", "slow"],
        "summary": "Recalculates every status flag.",
        "details": "It runs each status flag query, rewrites the flag tags, removes tags for flags that no longer exist and refreshes the StatusFlagColumns view. TouchPoint normally does it in the nightly batch, so call it only when flags must be current right now.",
        "flow": ["Lists every saved search marked as a status flag, plus searches used as conditions in published process steps.", "For each one, runs the search into a temporary tag.", "Finds or creates the flag's tag (type 100), adds people who now match and removes people who no longer match.", "Deletes flag tags whose search is no longer a status flag, then rebuilds the StatusFlagColumns view."],
        "returns": "None",
        "watch_out": ["Writes data.", "Can take a long time on a large database.", "The single flag version, model.UpdateStatusFlag, raises a SQL error."],
        "example": "# Recalculates ALL status flags. Slow on large databases, so it is commented out.\n# model.UpdateStatusFlags()\nflags = model.StatusFlagList()\nprint 'There are %d status flags' % len(flags)",
        "see_also": ["model.UpdateStatusFlag", "model.StatusFlagList", "model.StatusFlagDictionary"]
    },
    "q.BlueToolbarCount": {
        "category": "Searches & Tags",
        "kind": "method",
        "arity": [0, 1],
        "tags": ["read_only"],
        "summary": "Counts the people the user selected when running the script from the Blue Toolbar.",
        "details": "Unlike @\x42lueToolbarTagId, which is capped at 1000 people, this counts the whole selection. Pass a query guid to count that query instead.",
        "params": [
            ["guidStr", "string", "Query guid to count instead of the Blue Toolbar selection", False, "None"]
        ],
        "returns": "integer",
        "watch_out": ["From the Blue Toolbar menu, TouchPoint refuses to run a script that does not contain the text BlueToolbarReport or @\x42lueToolbarTagId, so include one of them (a comment is enough).", "Returns 0 when there is no selection.", "When run by URL instead of from the Blue Toolbar menu, a script containing @\x42lueToolbarTagId receives the user's most recent search as the selection."],
        "example": "# @\x42lueToolbarTagId  (this text must appear in the script)\nn = q.BlueToolbarCount()\nif n == 0:\n    print 'Run this from the Blue Toolbar after a search.'\nelse:\n    print '%d people selected' % n",
        "signatures": ["q.BlueToolbarCount(guidStr=None)"],
        "see_also": ["q.BlueToolbarReport", "q.QueryCount"]
    },
    "q.BlueToolbarReport": {
        "category": "Searches & Tags",
        "kind": "method",
        "arity": [0, 4],
        "tags": ["read_only"],
        "summary": "Returns the people selected in the Blue Toolbar when the script is run from it.",
        "details": "Sort options match q.QueryList.",
        "flow": ["Uses the guid you pass, or else the selection TouchPoint handed the script.", "Runs that search. Deceased and archived people are left out unless the search includes them.", "Sorts when sort is given, then skips offset rows and returns up to limit people (1000 by default)."],
        "params": [
            ["sort", "string", "Sort order such as \"name\", or None for unsorted", False, "None"],
            ["limit", "integer", "Maximum rows", False, "1000"],
            ["offset", "integer", "Rows to skip", False, "0"],
            ["guidStr", "string", "Query id to use instead of the Blue Toolbar selection", False, "None"]
        ],
        "returns": "List of Person objects",
        "watch_out": ["Returns an empty list when there is no Blue Toolbar selection.", "Sort names match q.QueryList and ignore case. An unknown sort name leaves the list unsorted.", "From the Blue Toolbar menu, a script must contain the text BlueToolbarReport or @\x42lueToolbarTagId, so calling this method is enough."],
        "example": "people = list(q.BlueToolbarReport(\"name\"))\nif not people:\n    print \"Run this from the Blue Toolbar\"\nfor p in people:\n    print p.PeopleId, p.Name",
        "see_also": ["q.BlueToolbarCount", "q.QueryList"]
    },
    "q.ElapsedTime": {
        "category": "Searches & Tags",
        "kind": "property",
        "tags": ["read_only"],
        "summary": "The number of seconds (rounded) the most recent q.QueryCount call took.",
        "details": "Handy for spotting slow search conditions.",
        "returns": "integer",
        "watch_out": ["It is 0 until QueryCount has run.", "It is rounded to whole seconds, so most counts show 0."],
        "example": "n = q.QueryCount('IsMember = 1[True]')\nprint '%d members counted in %d seconds' % (n, q.ElapsedTime)",
        "see_also": ["q.QueryCount"]
    },
    "q.GetWhereClause": {
        "category": "Searches & Tags",
        "kind": "method",
        "arity": [1, 1],
        "tags": ["read_only"],
        "summary": "Translates Search Builder code or a saved search name into the SQL statement TouchPoint runs for it.",
        "details": "Handy for learning or debugging a search.",
        "flow": ["Turns the query text into people: digits alone become one PeopleId, peopleids='1,2,3' becomes a list, text that exactly matches any saved search name runs that search, and anything else is read as Search Builder code. An empty string matches no one.", "Deceased and archived people are left out unless the code itself mentions them (IncludeDeceased, IncludeArchived, DeceasedDate or ArchivedFlag). A campus restriction or OrgLeadersOnly limit on the signed in user also narrows the list.", "Builds the SQL for that search and writes the values into it in place of parameters."],
        "params": [
            ["code", "string", "Search Builder code or saved search name", True, None]
        ],
        "returns": "String: the full SQL SELECT statement for the search, with values filled in",
        "watch_out": ["A search that uses several steps returns a message saying a WHERE clause cannot be created, instead of SQL."],
        "example": "print q.GetWhereClause(\"Age >= 65\")",
        "see_also": ["q.QueryCount", "q.QueryList"]
    },
    "q.QueryCount": {
        "category": "Searches & Tags",
        "kind": "method",
        "arity": [1, 1],
        "tags": ["read_only"],
        "summary": "Counts the people matching a query.",
        "details": "Accepts a PeopleId string, a saved search name, or Search Builder code.",
        "flow": ["Turns the query text into people: digits alone become one PeopleId, peopleids='1,2,3' becomes a list, text that exactly matches any saved search name runs that search, and anything else is read as Search Builder code. An empty string matches no one.", "Deceased and archived people are left out unless the code itself mentions them (IncludeDeceased, IncludeArchived, DeceasedDate or ArchivedFlag). A campus restriction or OrgLeadersOnly limit on the signed in user also narrows the list.", "Counts the matches and saves the time taken in q.ElapsedTime."],
        "params": [
            ["query", "string", "PeopleId, saved search name, or Search Builder code", True, None]
        ],
        "returns": "Integer count",
        "watch_out": ["Returns 0 when nothing matches.", "A PeopleId of an archived or deceased person counts as 0 unless the code mentions those people."],
        "example": "print q.QueryCount(\"Age >= 65\"), \"people 65 and over\"",
        "see_also": ["q.QueryList", "q.QueryPeopleIds", "q.ElapsedTime"]
    },
    "q.QueryCountDate": {
        "category": "Searches & Tags",
        "kind": "method",
        "arity": [2, 2],
        "tags": ["read_only", "deprecated"],
        "summary": "Counts people matching a search, with every date condition's start and end set to one date.",
        "details": "Works with a saved search or search code.",
        "params": [
            ["s", "query", "Saved search name or search code", True, None],
            ["startdt", "date", "Date to use as both start and end date", True, None]
        ],
        "returns": "integer",
        "watch_out": ["Deprecated. Write the dates into the search condition and use q.QueryCount instead.", "A date TouchPoint cannot read raises a \"bad date\" error.", "The override only reaches conditions that use a start or end date. Conditions without dates are counted as written."],
        "example": "n = q.QueryCountDate('Age >= 65', '10/5/2026')\nprint n",
        "see_also": ["q.QueryCount", "q.QueryCountDateRange", "q.QueryCountDivDate"]
    },
    "q.QueryCountDateRange": {
        "category": "Searches & Tags",
        "kind": "method",
        "arity": [3, 3],
        "tags": ["read_only", "deprecated"],
        "summary": "Counts people matching a search, with every date condition set to run from startdt to startdt plus days.",
        "details": "Works with a saved search or search code.",
        "params": [
            ["s", "query", "Saved search name or search code", True, None],
            ["startdt", "date", "Start date", True, None],
            ["days", "integer", "Number of days after the start date for the end date", True, None]
        ],
        "returns": "integer",
        "watch_out": ["Deprecated. Write the dates into the search condition and use q.QueryCount instead.", "A date TouchPoint cannot read raises a \"bad date\" error.", "The override only reaches conditions that use a start or end date. Conditions without dates are counted as written.", "The end date is startdt plus days, so 30 days from 10/1 ends on 10/31."],
        "example": "n = q.QueryCountDateRange('Age >= 65', '10/1/2026', 30)\nprint n",
        "see_also": ["q.QueryCount", "q.QueryCountDate", "q.QueryCountDivDateRange"]
    },
    "q.QueryCountDivDate": {
        "category": "Searches & Tags",
        "kind": "method",
        "arity": [3, 4],
        "tags": ["read_only", "deprecated"],
        "summary": "Like q.QueryCountDate, but also overrides the division in every condition with a named division.",
        "details": "The division can optionally be limited to a program id.",
        "flow": ["Looks up the division by exact name in the Division table, limited to progid when you give one.", "Sets the start date, end date and division on every condition in the search for this one call.", "Runs the search, counts the people, then clears the overrides."],
        "params": [
            ["s", "query", "Saved search name or search code", True, None],
            ["progid", "integer", "Optional program id used to find the division", False, "None"],
            ["division", "string", "Division name", True, None],
            ["startdt", "date", "Date to use as both start and end date", True, None]
        ],
        "returns": "integer",
        "watch_out": ["Deprecated. Write the dates into the search condition and use q.QueryCount instead.", "If the name matches no division, the division becomes 0, which drops the division filter instead of keeping the original division.", "A date TouchPoint cannot read raises a \"bad date\" error.", "If two divisions share the name and no program id narrows it, the call raises an error.", "The division override applies to every condition in the search, even ones that had no division."],
        "example": "n = q.QueryCountDivDate('Age >= 65', 'First Division', '10/5/2026')\nprint n",
        "signatures": ["q.QueryCountDivDate(s, division, startdt)", "q.QueryCountDivDate(s, progid, division, startdt)"],
        "see_also": ["q.QueryCount", "q.QueryCountDate", "q.QueryCountDivDateRange"]
    },
    "q.QueryCountDivDateRange": {
        "category": "Searches & Tags",
        "kind": "method",
        "arity": [4, 4],
        "tags": ["read_only", "deprecated"],
        "summary": "Like q.QueryCountDateRange, but also overrides the division in every condition with a named division.",
        "flow": ["Looks up the division by exact name in the Division table, limited to progid when you give one.", "Sets the start date, end date and division on every condition in the search for this one call.", "Runs the search, counts the people, then clears the overrides."],
        "params": [
            ["s", "query", "Saved search name or search code", True, None],
            ["division", "string", "Division name", True, None],
            ["startdt", "date", "Start date", True, None],
            ["days", "integer", "Days after the start date for the end date", True, None]
        ],
        "returns": "integer",
        "watch_out": ["Deprecated. Write the dates into the search condition and use q.QueryCount instead.", "A date TouchPoint cannot read raises a \"bad date\" error.", "If two divisions share the name and no program id narrows it, the call raises an error.", "The division override applies to every condition in the search, even ones that had no division.", "If the name matches no division, the division becomes 0, which drops the division filter."],
        "example": "n = q.QueryCountDivDateRange('Age >= 65', 'First Division', '10/1/2026', 30)\nprint n",
        "see_also": ["q.QueryCount", "q.QueryCountDateRange", "q.QueryCountProgIdDivDateRange"]
    },
    "q.QueryCountProgIdDivDateRange": {
        "category": "Searches & Tags",
        "kind": "method",
        "arity": [5, 5],
        "tags": ["read_only", "deprecated"],
        "summary": "Like q.QueryCountDivDateRange, with a program id to pick the right division when several share a name.",
        "flow": ["Looks up the division by exact name in the Division table, limited to progid when you give one.", "Sets the start date, end date and division on every condition in the search for this one call.", "Runs the search, counts the people, then clears the overrides."],
        "params": [
            ["s", "query", "Saved search name or search code", True, None],
            ["progid", "integer", "Program id, or None", True, None],
            ["division", "string", "Division name", True, None],
            ["startdt", "date", "Start date", True, None],
            ["days", "integer", "Days after the start date for the end date", True, None]
        ],
        "returns": "integer",
        "watch_out": ["Deprecated. Write the dates into the search condition and use q.QueryCount instead.", "A date TouchPoint cannot read raises a \"bad date\" error.", "If the name matches no division in that program, the division becomes 0, which drops the division filter.", "Pass None for progid to search all programs; then two divisions with the same name raise an error."],
        "example": "n = q.QueryCountProgIdDivDateRange('Age >= 65', None, 'First Division', '10/1/2026', 30)\nprint n",
        "see_also": ["q.QueryCount", "q.QueryCountDivDateRange"]
    },
    "q.QueryList": {
        "category": "Searches & Tags",
        "kind": "method",
        "arity": [1, 4],
        "tags": ["read_only"],
        "summary": "Returns the people matching a query as Person records, sorted and paged.",
        "details": "Sort options: name, age, birthday, peopleid, familyid, rand, each optionally followed by \" desc\". Accepts a PeopleId, a saved search name, or Search Builder code.",
        "params": [
            ["query", "query", "People to act on: a PeopleId, a saved search name, or Search Builder code", True, None],
            ["sort", "string", "Sort order", False, "\"name\""],
            ["limit", "integer", "Maximum rows", False, "1000"],
            ["offset", "integer", "Rows to skip", False, "0"]
        ],
        "returns": "List of Person objects",
        "watch_out": ["Returns at most 1000 people unless you pass a larger limit.", "An unknown sort value returns the people in no set order, so paging with offset can repeat or skip people.", "Passing None for sort raises an error."],
        "example": "for p in q.QueryList(\"Age >= 65\", \"age desc\", 10):\n    print p.PeopleId, p.Name, p.Age",
        "see_also": ["q.QueryCount", "q.QueryPeopleIds", "q.BlueToolbarReport"]
    },
    "q.QueryList2": {
        "category": "Searches & Tags",
        "kind": "method",
        "arity": [3, 3],
        "tags": ["read_only", "deprecated"],
        "summary": "Returns up to 1000 people matching a saved search or search code, sorted by the given field.",
        "params": [
            ["query", "query", "Saved search name or search code", True, None],
            ["sort", "string", "Sort field: name, age, birthday and the other q.QueryList sorts", True, None],
            ["ascending", "boolean", "False to sort descending", True, None]
        ],
        "returns": "list of Person",
        "watch_out": ["Deprecated. Use q.QueryList(query, sort, limit), passing a sort such as 'name desc' for descending order.", "Returns at most 1000 people.", "With ascending False, only name, age, birthday, peopleid and familyid sort; rand comes back unsorted."],
        "example": "for p in q.QueryList2('IsMember = 1[True]', 'age', False):\n    print p.Name, p.Age\n    break",
        "see_also": ["q.QueryList", "q.QueryCount"]
    },
    "q.QueryPeopleIds": {
        "category": "Searches & Tags",
        "kind": "method",
        "arity": [1, 1],
        "tags": ["read_only"],
        "summary": "Returns the PeopleIds matching a query.",
        "details": "Accepts a PeopleId string, a saved search name, or Search Builder code.",
        "params": [
            ["query", "string", "PeopleId, saved search name, or Search Builder code", True, None]
        ],
        "returns": "List of integers",
        "watch_out": ["There is no 1000 row limit here, unlike q.QueryList.", "A query such as peopleids='1,2,3' is accepted directly."],
        "example": "ids = list(q.QueryPeopleIds(\"Age >= 65\"))\nprint len(ids), \"people\", ids[:10]",
        "see_also": ["q.QueryList", "q.QueryCount", "q.SqlPeopleIdsToQuery"]
    },
    "q.SqlPeopleIdsToQuery": {
        "category": "Searches & Tags",
        "kind": "method",
        "arity": [1, 1],
        "tags": ["read_only"],
        "summary": "Runs read only SQL that returns PeopleIds and turns the result into a query string like peopleids='1,2,3'.",
        "details": "The result is usable anywhere a query is accepted.",
        "params": [
            ["sql", "string", "SQL whose first column is PeopleId", True, None]
        ],
        "returns": "String like peopleids='1001,1002', or \"\"",
        "watch_out": ["Returns an empty string when the SQL returns nothing.", "The SQL runs on the read only connection, so it cannot change data.", "The first column must be a whole number. Duplicate ids are kept.", "A very large result makes a very long query string."],
        "example": "query = q.SqlPeopleIdsToQuery(\"SELECT TOP 5 PeopleId FROM People WHERE Age >= 90\")\nprint query\nprint q.QueryCount(query) if query else 0",
        "see_also": ["q.TagSqlPeopleIds", "q.QuerySqlPeopleIds", "q.QueryCount"]
    },
    "q.StatusCount": {
        "category": "Searches & Tags",
        "kind": "method",
        "arity": [1, 1],
        "tags": ["read_only"],
        "summary": "Counts people who carry every one of the given status flags, matched by the flag's name.",
        "details": "Pass status flag names separated by commas, such as \"Active Attender,In a Connect Group\". \"F00\" counts everyone in the People table.",
        "params": [
            ["flags", "string", "Comma separated status flag names", True, None]
        ],
        "returns": "Integer count",
        "watch_out": ["It matches tags by name across all tags, so a personal tag with the same name as a flag is counted too.", "A flag name that contains a comma cannot be used, because the list is split on commas.", "Counts every row in People, including deceased and archived records."],
        "example": "print q.StatusCount(\"Active Attender\"), \"people with that flag\"",
        "see_also": ["model.StatusFlagDictionary", "q.QueryCount"]
    },
    "q.TagCount": {
        "category": "Searches & Tags",
        "kind": "method",
        "arity": [1, 1],
        "tags": ["read_only"],
        "summary": "Counts the people in a tag by tag id.",
        "params": [
            ["tagid", "integer", "Tag id", True, None]
        ],
        "returns": "Integer count",
        "watch_out": ["Works for any tag id, including other users' tags. An unknown id returns 0."],
        "example": "tagId = q.TagQueryList(\"Age >= 65\")\nprint q.TagCount(tagId), \"people in tag\", tagId",
        "see_also": ["q.TagQueryList", "q.TagSqlPeopleIds"]
    },
    "q.TagQueryList": {
        "category": "Searches & Tags",
        "kind": "method",
        "arity": [1, 1],
        "tags": ["writes_data"],
        "summary": "Puts the people matching a query into a temporary tag and returns the tag id.",
        "details": "Useful for SQL that joins TagPerson on that id. Accepts a PeopleId, a saved search name, or Search Builder code.",
        "flow": ["Creates a new temporary tag owned by the current user, named by the session id.", "Runs the search and inserts the matching PeopleIds into TagPerson in one database statement.", "Returns the new tag id for use in SQL joins on TagPerson.Id."],
        "params": [
            ["query", "query", "People to act on: a PeopleId, a saved search name, or Search Builder code", True, None]
        ],
        "returns": "Integer tag id",
        "watch_out": ["Creates a temporary tag.", "Each call makes a separate new tag. These tags stay in the Tag table after the script ends."],
        "example": "tagId = q.TagQueryList(\"Age >= 65\")\nrows = q.QuerySql(\"SELECT COUNT(*) AS n FROM TagPerson WHERE Id = %d\" % tagId)\nprint rows[0].n",
        "see_also": ["q.TagSqlPeopleIds", "q.TagCount", "model.CreateQueryTag"]
    },
    "q.TagSqlPeopleIds": {
        "category": "Searches & Tags",
        "kind": "method",
        "arity": [1, 1],
        "tags": ["writes_data"],
        "summary": "Runs SQL that returns PeopleIds, puts them in a temporary tag and returns the tag id.",
        "flow": ["Runs your SQL and collects the PeopleIds from the first column.", "Creates a new temporary tag owned by the current user, named by the session id.", "Adds those PeopleIds to TagPerson and returns the new tag id."],
        "params": [
            ["sql", "string", "SQL returning one PeopleId column", True, None]
        ],
        "returns": "Integer tag id",
        "watch_out": ["Creates a temporary tag.", "Each call makes a separate new tag. These tags stay in the Tag table after the script ends.", "Unlike q.QuerySql, this SQL runs on the regular connection, not the read only one. Keep it to SELECT statements."],
        "example": "tagId = q.TagSqlPeopleIds(\"SELECT PeopleId FROM People WHERE Age >= 90\")\nprint tagId, q.TagCount(tagId)",
        "see_also": ["q.TagQueryList", "q.TagCount", "q.SqlPeopleIdsToQuery"]
    },
    "model.CreateTask": {
        "category": "Tasks & Notes",
        "kind": "method",
        "arity": [3, 6],
        "tags": ["writes_data", "sends_message", "deprecated"],
        "summary": "Creates a task in the legacy Task table and emails the owner.",
        "flow": ["Finds or creates the owner's InBox task list.", "Adds a row to the legacy Task table with status Active, about the given person.", "Unless sendEmail is False, emails the owner from the church admin address. The subject starts with the TaskEmailSubjectPrefix setting, TASK: by default."],
        "params": [
            ["ministerId", "integer", "PeopleId of the task owner", True, None],
            ["aboutId", "integer", "PeopleId the task is about", True, None],
            ["description", "string", "Task description", True, None],
            ["notes", "string", "Notes", False, "None"],
            ["sendEmail", "boolean", "Email the owner", False, "True"]
        ],
        "returns": "None",
        "watch_out": ["Deprecated. Use model.CreateTaskNote instead.", "The legacy Task table is not shown on the current Tasks and Notes screens.", "Emails the owner unless sendEmail is False.", "If aboutId matches no one, the task is saved and then the email step raises an error.", "A second form takes a delegate PeopleId after aboutId; that person becomes co-owner and is emailed too."],
        "example": "# Deprecated. Use model.CreateTaskNote instead.\n# model.CreateTask(model.UserPeopleId, model.UserPeopleId, \"Please contact about small group\")\nprint \"See model.CreateTaskNote\"",
        "signatures": ["model.CreateTask(ministerId, aboutId, description, notes=None, sendEmail=True)", "model.CreateTask(ministerId, aboutId, delegateId, description, notes=None, sendEmail=True)"],
        "see_also": ["model.CreateTaskNote", "model.GetTaskNote", "model.GetTasksNotes"]
    },
    "model.CreateTaskNote": {
        "category": "Tasks & Notes",
        "kind": "method",
        "arity": [2, 10],
        "tags": ["writes_data", "sends_message"],
        "summary": "Creates a task or a note and returns its TaskNoteId.",
        "details": "A task with an assignee starts as Pending; without one it starts as Accepted.",
        "flow": ["Builds a new TaskNote row. CreatedBy is the signed in user, or the database admin person when nobody is signed in, and the created date is in the church time zone.", "Sets the status: Note (5) when isNote is True, Pending (2) when there is an assignee, otherwise Accepted (3). It saves the row right away.", "Attaches each keyword id as a TaskNoteKeyword row. Each id must be a number.", "For a task (never a note) with sendEmails left True, it builds the email from the System Task Assigned text content and sends it to the assignee, or to the owner when there is no assignee. The email comes from the owner's email address and is sent as a transactional email.", "Returns the new TaskNoteId."],
        "params": [
            ["ownerId", "integer", "PeopleId of the owner", True, None],
            ["aboutPersonId", "integer", "PeopleId the task or note is about", True, None],
            ["assigneeId", "integer", "PeopleId to assign to", False, "None"],
            ["roleId", "integer", "Limit visibility to this role id", False, "None"],
            ["isNote", "boolean", "True for a note, False for a task", False, "False"],
            ["instructions", "string", "Instructions for the assignee", False, "None"],
            ["notes", "string", "Notes text", False, "None"],
            ["dueDate", "date", "Due date as System.DateTime", False, "None"],
            ["keywordIdList", "list", "Keyword ids to attach", False, "None"],
            ["sendEmails", "boolean", "Send the usual notification emails", False, "True"]
        ],
        "returns": "Integer TaskNoteId",
        "watch_out": ["Tasks send notification emails unless sendEmails is False.", "It does not check that the owner, about person or assignee exist.", "With no assignee, the owner gets the new task email about their own task.", "The task is saved before keywords and email. A bad keyword id or a missing System Task Assigned text content raises an error, but the task already exists.", "Notes never send email, whatever sendEmails says."],
        "example": "from System import DateTime\nme = model.UserPeopleId\n# Creates a real task assigned to you. Uncomment to run.\n# tid = model.CreateTaskNote(me, model.UserPeopleId, me, None, False, \"Call about the visit\", None, DateTime.Now.AddDays(7), [], False)\n# print \"Created task\", tid\nprint \"Owner would be\", me",
        "see_also": ["model.GetTaskNote", "model.TaskNoteEdit", "model.SetTaskNoteStatus"]
    },
    "model.GetExtraQuestions": {
        "category": "Tasks & Notes",
        "kind": "method",
        "arity": [1, 1],
        "tags": ["read_only"],
        "summary": "Returns the keyword extra questions that apply to a task or note, based on its keywords.",
        "params": [
            ["taskNoteId", "integer", "TaskNoteId", True, None]
        ],
        "returns": "List of question objects (name, dataType, response and more)",
        "watch_out": ["Returns an empty list if the task is not found or has no keywords.", "Questions limited to a role are still returned; check roleId yourself.", "Answers are not filled in. Use model.GetExtraQuestionsWithAnswers for those."],
        "example": "tid = q.QuerySqlInt(\"SELECT TOP 1 TaskNoteId FROM TaskNote WHERE IsNote = 0 ORDER BY TaskNoteId DESC\")\nquestions = list(model.GetExtraQuestions(tid))\nprint len(questions), \"extra questions on task\", tid\nfor qn in questions:\n    print qn.name, qn.dataType",
        "see_also": ["model.GetExtraQuestionsWithAnswers", "model.GetExtraQuestionsByKeywords", "model.TaskNoteComplete"]
    },
    "model.GetExtraQuestionsByKeywords": {
        "category": "Tasks & Notes",
        "kind": "method",
        "arity": [1, 1],
        "tags": ["read_only"],
        "summary": "Returns the extra questions attached to the given keyword ids.",
        "params": [
            ["keywords", "list", "Keyword ids", True, None]
        ],
        "returns": "List of question objects",
        "watch_out": ["An empty list or None returns an empty list.", "A keyword id that does not exist raises an error.", "Repeating a keyword id repeats its questions."],
        "example": "ids = list(q.QuerySqlInts(\"SELECT TOP 5 KeywordId FROM Keyword\"))\nfor qn in model.GetExtraQuestionsByKeywords(ids):\n    print qn.name",
        "see_also": ["model.GetExtraQuestions", "model.GetTaskNoteKeywordIds"]
    },
    "model.GetExtraQuestionsWithAnswers": {
        "category": "Tasks & Notes",
        "kind": "method",
        "arity": [1, 1],
        "tags": ["read_only"],
        "summary": "Returns the extra questions for a task or note along with the answers already recorded.",
        "params": [
            ["taskNoteId", "integer", "TaskNoteId", True, None]
        ],
        "returns": "List of question objects with response filled in",
        "watch_out": ["If nothing has been answered yet, it returns the blank questions.", "A dropdown response is the option id as text; dropdownResponse holds the id and text. Checkbox answers are in checkboxesList."],
        "example": "tid = q.QuerySqlInt(\"SELECT TOP 1 TaskNoteId FROM TaskNote WHERE IsNote = 0 ORDER BY TaskNoteId DESC\")\nanswers = list(model.GetExtraQuestionsWithAnswers(tid))\nprint len(answers), \"answered questions on task\", tid\nfor qn in answers:\n    print qn.name, \"=\", qn.response",
        "see_also": ["model.GetExtraQuestions", "model.TaskNoteComplete"]
    },
    "model.GetTaskNote": {
        "category": "Tasks & Notes",
        "kind": "method",
        "arity": [1, 1],
        "tags": ["read_only"],
        "summary": "Loads one task or note with its details.",
        "details": "Details include Instructions, Notes, StatusId, DueDate, AboutPersonId, OwnerId, AssigneeId and more.",
        "params": [
            ["taskNoteId", "integer", "TaskNoteId", True, None]
        ],
        "returns": "Task/note object, or None",
        "watch_out": ["Returns None if not found.", "Also returns None for archived items and for notes attached to an involvement.", "Needs a signed in user; without one, such as in a scheduled run, it raises an error.", "It does not check whether the current user's roles allow them to see the task."],
        "example": "tid = q.QuerySqlInt(\"SELECT TOP 1 TaskNoteId FROM TaskNote WHERE IsNote = 0 ORDER BY TaskNoteId DESC\")\nt = model.GetTaskNote(tid)\nif t:\n    print t.TaskNoteId, t.StatusId, t.Instructions",
        "see_also": ["model.TaskNoteEdit", "model.GetTasksNotes", "model.CreateTaskNote"]
    },
    "model.GetTaskNoteKeywordIds": {
        "category": "Tasks & Notes",
        "kind": "method",
        "arity": [1, 1],
        "tags": ["read_only"],
        "summary": "Returns the keyword ids attached to a task or note.",
        "params": [
            ["taskNoteId", "integer", "TaskNoteId", True, None]
        ],
        "returns": "List of integer keyword ids, or None",
        "watch_out": ["Returns None if the task is not found.", "An item with no keywords returns an empty list, not None."],
        "example": "tid = q.QuerySqlInt(\"SELECT TOP 1 TaskNoteId FROM TaskNote WHERE IsNote = 0 ORDER BY TaskNoteId DESC\")\nprint list(model.GetTaskNoteKeywordIds(tid) or [])",
        "see_also": ["model.SetTaskNoteKeywordIds", "model.TaskNoteMassAddKeywords"]
    },
    "model.GetTasksNotes": {
        "category": "Tasks & Notes",
        "kind": "method",
        "arity": [3, 6],
        "tags": ["read_only", "slow"],
        "summary": "Returns non archived tasks and notes, optionally limited by status or to one involvement's notes.",
        "details": "The first three arguments are required but can be None. Status ids: 1 complete, 2 pending, 3 accepted, 4 declined, 5 note, 6 archived.",
        "params": [
            ["numOfTaskNotes", "integer", "Maximum number to return, or None", True, None],
            ["notesOnly", "boolean", "True with an orgId for that involvement's notes, or None", True, None],
            ["orgId", "integer", "OrganizationId for notesOnly, or None", True, None],
            ["statuses", "list", "Status ids to include, as a .NET int array: 1 complete, 2 pending, 3 accepted, 4 declined, 5 note", False, "None"],
            ["peopleId", "integer", "Ignored", False, "None"],
            ["filter", "string", "Ignored", False, "\"\""]
        ],
        "returns": "List of task/note objects",
        "watch_out": ["It loads every matching task before keeping the first numOfTaskNotes, so it can take a minute or more on a large database.", "peopleId and filter are accepted but ignored.", "Without statuses it returns non archived items for everyone, not only the current user's. With statuses it returns only items owned by or assigned to the current user.", "orgId only applies when notesOnly is True; then it returns that involvement's notes. Otherwise only items not tied to an involvement come back.", "Archived items are always left out, so status 6 returns nothing.", "Results are in no set order, so the first numOfTaskNotes are not the newest.", "Needs a signed in user; without one it raises an error."],
        "example": "from System import Array\n# Your own pending, accepted and declined tasks\nfor t in model.GetTasksNotes(10, False, None, Array[int]([2, 3, 4])):\n    print t.TaskNoteId, t.StatusId, t.Instructions",
        "see_also": ["model.GetTaskNote", "model.CreateTaskNote"]
    },
    "model.SetTaskNoteKeywordIds": {
        "category": "Tasks & Notes",
        "kind": "method",
        "arity": [1, 2],
        "tags": ["writes_data"],
        "summary": "Replaces the keywords on a task or note with exactly the given list.",
        "params": [
            ["taskNoteId", "integer", "TaskNoteId", True, None],
            ["keywordIds", "list", "Keyword ids to keep", False, "None"]
        ],
        "returns": "None",
        "watch_out": ["Passing an empty list or None removes all keywords.", "An unknown task id does nothing and raises no error.", "Each id must be a number or numeric text; anything else raises an error.", "Sends no emails and does not add or remove extra question answers."],
        "example": "tid = q.QuerySqlInt(\"SELECT TOP 1 TaskNoteId FROM TaskNote WHERE IsNote = 0 ORDER BY TaskNoteId DESC\")\nprint \"Current:\", list(model.GetTaskNoteKeywordIds(tid) or [])\n# Uncomment to replace them:\n# model.SetTaskNoteKeywordIds(tid, [1, 2])",
        "see_also": ["model.GetTaskNoteKeywordIds", "model.TaskNoteMassAddKeywords"]
    },
    "model.SetTaskNoteStatus": {
        "category": "Tasks & Notes",
        "kind": "method",
        "arity": [2, 2],
        "tags": ["writes_data"],
        "summary": "Changes a task's status, or turns a note into a task and back.",
        "details": "Accepts a name (Complete, Pending, Accepted, Declined, Note, Archived) or its id 1 to 6. Sends no emails.",
        "params": [
            ["taskNoteId", "integer", "TaskNoteId", True, None],
            ["status", "any", "Status name or id 1 to 6", True, None]
        ],
        "returns": "None",
        "watch_out": ["Unknown values are ignored.", "An unknown task id does nothing and raises no error.", "Complete does not set the completed date or completed by fields. Use model.TaskNoteComplete for that.", "Archived hides the item from model.GetTaskNote and model.GetTasksNotes. Any other status unarchives it."],
        "example": "tid = q.QuerySqlInt(\"SELECT TOP 1 TaskNoteId FROM TaskNote WHERE IsNote = 0 ORDER BY TaskNoteId DESC\")\nprint \"Status now\", model.GetTaskNote(tid).StatusId\n# Uncomment to mark it accepted:\n# model.SetTaskNoteStatus(tid, \"Accepted\")",
        "see_also": ["model.TaskNoteComplete", "model.TaskNoteAccept", "model.TaskNoteDecline"]
    },
    "model.TaskNoteAccept": {
        "category": "Tasks & Notes",
        "kind": "method",
        "arity": [1, 1],
        "tags": ["writes_data", "sends_message"],
        "summary": "Marks a task as Accepted by the current user and emails the owner.",
        "flow": ["Loads the task by id. There is no check, so a wrong id raises an error.", "Sets the status to Accepted and records the signed in user and church time as ModifiedBy and ModifiedDate, then saves.", "Emails the owner using the System Task Accepted text content, or System Task Assigned when that is missing. The subject names the assignee, or the signed in user when the task has no assignee.", "Returns the same TaskNoteId."],
        "params": [
            ["taskNoteId", "integer", "TaskNoteId", True, None]
        ],
        "returns": "Integer TaskNoteId",
        "watch_out": ["Sends an email to the owner.", "Needs a signed in user, so it fails in the morning batch.", "Accepting does not make the signed in user the assignee. An unassigned task stays unassigned.", "It does not check the current status, so calling it on a completed or declined task sets it back to Accepted."],
        "example": "tid = q.QuerySqlInt(\"SELECT TOP 1 TaskNoteId FROM TaskNote WHERE IsNote = 0 ORDER BY TaskNoteId DESC\")\n# Sends an email. Uncomment to run.\n# model.TaskNoteAccept(tid)\nprint \"Would accept task\", tid",
        "see_also": ["model.TaskNoteDecline", "model.TaskNoteComplete", "model.SetTaskNoteStatus"]
    },
    "model.TaskNoteComplete": {
        "category": "Tasks & Notes",
        "kind": "method",
        "arity": [3, 3],
        "tags": ["writes_data", "sends_message"],
        "summary": "Completes a task as the current user and saves the completion note and extra question answers.",
        "details": "It also triggers any Process Builder step waiting on the task.",
        "flow": ["Loads the task by id. There is no check, so a wrong id raises an error.", "Sets the status to Complete, replaces the task's Notes with your note, and records the signed in user and the current church time as CompletedBy and CompletedDate.", "Saves the extra question answers in TaskNoteExtraValue, updating an answer that already exists or adding a new one. Instruction items in the list are skipped.", "Emails the owner using the System Task Completed text content, or System Task Assigned when that is missing. The subject names the assignee, or the signed in user when the task has no assignee.", "Queues a Process Builder check so any step waiting on this task for the about person can move on."],
        "params": [
            ["taskNoteId", "integer", "TaskNoteId", True, None],
            ["note", "string", "Completion note", True, None],
            ["extraQuestions", "list", "Answers from model.GetExtraQuestions, or an empty list", True, None]
        ],
        "returns": "Integer TaskNoteId",
        "watch_out": ["Sends an email to the owner.", "Needs a signed in user.", "Your note replaces whatever was in the task's Notes, including notes given when the task was created.", "There is no way to skip the email to the owner, even if the task was created with sendEmails=False.", "The email says completed by the assignee even when someone else completed it.", "Passing None for extraQuestions also works and saves no answers."],
        "example": "tid = q.QuerySqlInt(\"SELECT TOP 1 TaskNoteId FROM TaskNote WHERE IsNote = 0 ORDER BY TaskNoteId DESC\")\nanswers = model.GetExtraQuestions(tid)\n# Sends an email. Uncomment to run.\n# model.TaskNoteComplete(tid, \"Called and left a message\", answers)\nprint \"Would complete task\", tid",
        "see_also": ["model.GetExtraQuestions", "model.TaskNoteMassComplete", "model.SetTaskNoteStatus"]
    },
    "model.TaskNoteDecline": {
        "category": "Tasks & Notes",
        "kind": "method",
        "arity": [2, 2],
        "tags": ["writes_data", "sends_message"],
        "summary": "Declines a task as the current user and stores the reason.",
        "flow": ["Sets the status to Declined, saves the reason, and records the current user as the last editor.", "Writes an activity log entry \"decline task\" with the reason, on the assignee's record.", "Emails the owner, using the System Task Declined Special Content if you have the System Task Assigned template set up, or a plain message otherwise."],
        "params": [
            ["taskNoteId", "integer", "TaskNoteId", True, None],
            ["reason", "string", "Reason for declining", True, None]
        ],
        "returns": "Integer TaskNoteId",
        "watch_out": ["Sends an email to the owner.", "Needs a signed in user.", "An unknown task id raises an error.", "It does not check that the current user is the assignee."],
        "example": "tid = q.QuerySqlInt(\"SELECT TOP 1 TaskNoteId FROM TaskNote WHERE IsNote = 0 ORDER BY TaskNoteId DESC\")\n# Sends an email. Uncomment to run.\n# model.TaskNoteDecline(tid, \"Out of town that week\")\nprint \"Would decline task\", tid",
        "see_also": ["model.TaskNoteAccept", "model.SetTaskNoteStatus"]
    },
    "model.TaskNoteEdit": {
        "category": "Tasks & Notes",
        "kind": "method",
        "arity": [1, 1],
        "tags": ["writes_data", "sends_message"],
        "summary": "Saves changes to a task or note using the object from model.GetTaskNote.",
        "details": "Edit Instructions, Notes, Owner, Assignee, Role or DueDateString, then pass it back.",
        "flow": ["Loads the task by TaskNoteId and copies Instructions, Notes, Owner, AboutPerson and Role from the object.", "If Assignee changed or was cleared, sets the status to Pending and emails the new assignee, or the owner when cleared.", "Sets the due date from DueDateString, at noon in your time zone.", "Deletes all keywords on the task and adds back the ones in IncomingKeywords.", "Saves extra question answers from extraQuestions when a new keyword has questions."],
        "params": [
            ["viewModel", "any", "Task/note object from model.GetTaskNote, with changes applied", True, None]
        ],
        "returns": "Integer 1",
        "watch_out": ["Reassigning emails the new assignee.", "The due date is read from DueDateString, not DueDate. An object from model.GetTaskNote has it empty, so the due date is cleared unless you set it.", "Keywords come from IncomingKeywords, not Keywords. If you leave it empty, every keyword is removed.", "Changing StatusId, OwnerId, AssigneeId or AboutPersonId has no effect. Change Owner, Assignee or AboutPerson, or use model.SetTaskNoteStatus.", "Always returns 1. An unknown task id raises an error."],
        "example": "tid = q.QuerySqlInt(\"SELECT TOP 1 TaskNoteId FROM TaskNote WHERE IsNote = 0 ORDER BY TaskNoteId DESC\")\nt = model.GetTaskNote(tid)\nif t:\n    print t.Instructions\n    # The edit reads DueDateString, so copy the date or it is cleared\n    # if t.DueDate: t.DueDateString = t.DueDate.ToString(\"M/d/yyyy\")\n    # t.Instructions = \"Updated instructions\"\n    # model.TaskNoteEdit(t)  # also removes keywords unless IncomingKeywords is set",
        "see_also": ["model.GetTaskNote", "model.TaskNoteMassAssign"]
    },
    "model.TaskNoteMassAddKeywords": {
        "category": "Tasks & Notes",
        "kind": "method",
        "arity": [2, 2],
        "tags": ["writes_data"],
        "summary": "Adds keywords to several tasks or notes at once.",
        "details": "Keywords they already have are skipped.",
        "params": [
            ["taskNoteIdList", "list", "TaskNoteIds", True, None],
            ["keywordsArray", "list", "Keyword ids, as System.Collections.Generic.List[int]", True, None]
        ],
        "returns": "Integer 1",
        "watch_out": ["The keyword list must be a .NET List[int].", "Ids of 0 or less are skipped. Keyword extra questions are not added or answered.", "No emails are sent."],
        "example": "from System.Collections.Generic import List\nkw = List[int]([1, 2])\n# Uncomment to run:\n# model.TaskNoteMassAddKeywords([123, 124], kw)\nprint \"Would add\", list(kw)",
        "see_also": ["model.SetTaskNoteKeywordIds", "model.GetTaskNoteKeywordIds"]
    },
    "model.TaskNoteMassArchive": {
        "category": "Tasks & Notes",
        "kind": "method",
        "arity": [1, 1],
        "tags": ["writes_data"],
        "summary": "Archives several tasks or notes at once.",
        "params": [
            ["taskNoteIdList", "list", "TaskNoteIds", True, None]
        ],
        "returns": "Integer 1",
        "watch_out": ["An unknown task id raises an error and nothing in the list is saved.", "The status stays as it was; only the archived flag changes. Archived items drop out of model.GetTasksNotes."],
        "example": "# Uncomment to run:\n# model.TaskNoteMassArchive([123, 124, 125])\nprint \"Reverse with model.TaskNoteMassUnarchive\"",
        "see_also": ["model.TaskNoteMassUnarchive", "model.TaskNoteMassDelete"]
    },
    "model.TaskNoteMassAssign": {
        "category": "Tasks & Notes",
        "kind": "method",
        "arity": [2, 2],
        "tags": ["writes_data", "sends_message"],
        "summary": "Assigns several tasks to one person.",
        "flow": ["For each id, sets the assignee on the TaskNote row.", "Emails the new assignee about each one, from the owner's email address, and saves as it goes.", "Returns 1 when done."],
        "params": [
            ["taskNoteIdList", "list", "TaskNoteIds", True, None],
            ["assigneeId", "integer", "PeopleId of the new assignee", True, None]
        ],
        "returns": "Integer 1",
        "watch_out": ["Emails the assignee about each task.", "The status does not change, so a declined or completed task keeps that status.", "An unknown task id raises an error partway through; earlier tasks are already saved and emailed."],
        "example": "# Sends one email per task. Uncomment to run.\n# model.TaskNoteMassAssign([123, 124], model.UserPeopleId)\nprint \"Would assign to\", model.UserPeopleId",
        "see_also": ["model.TaskNoteEdit", "model.CreateTaskNote"]
    },
    "model.TaskNoteMassComplete": {
        "category": "Tasks & Notes",
        "kind": "method",
        "arity": [3, 3],
        "tags": ["writes_data", "sends_message"],
        "summary": "Completes several tasks as the current user with one completion date and note.",
        "details": "Pass a .NET DateTime for the date.",
        "flow": ["For each id, sets the status to Complete, the completed date, completed by (the current user) and replaces Notes with your note.", "Emails each task's owner that it was completed.", "Checks with TouchPoint's process steps for any step waiting on that task, one web call per task.", "Saves all the changes at the end and returns 1."],
        "params": [
            ["taskNoteIdList", "list", "TaskNoteIds", True, None],
            ["completionDate", "date", "Completion date (System.DateTime)", True, None],
            ["note", "string", "Completion note", True, None]
        ],
        "returns": "Integer 1",
        "watch_out": ["Emails each task's owner.", "Needs a signed in user.", "Passing None for the note raises an error. Pass \"\" for no note.", "The note replaces any notes already on the task.", "The process step check runs one task at a time, so a long list can be slow."],
        "example": "from System import DateTime\n# Sends email for each task. Uncomment to run.\n# model.TaskNoteMassComplete([123, 124], DateTime.Now, \"Handled at staff meeting\")\nprint \"Would complete as of\", DateTime.Now",
        "see_also": ["model.TaskNoteComplete", "model.TaskNoteMassArchive"]
    },
    "model.TaskNoteMassDelete": {
        "category": "Tasks & Notes",
        "kind": "method",
        "arity": [1, 1],
        "tags": ["writes_data"],
        "summary": "Permanently deletes several tasks or notes, skipping any that have a follow up task.",
        "details": "Returns how many were skipped because a follow up task was made from them.",
        "flow": ["If the user has ViewPrivateTouchpoints but not Admin, Developer or ManageTouchpoints, drops tasks they do not own or are not assigned to, tasks limited to a role they lack, and tasks about themselves.", "Skips any task that has a follow up task made from it, and counts those.", "For the rest, writes an activity log entry, then deletes their extra question answers, keywords, linked mobile prayer requests and the TaskNote rows.", "Returns the number skipped because they have a follow up task."],
        "params": [
            ["taskNoteIdList", "list", "TaskNoteIds", True, None]
        ],
        "returns": "Integer count of tasks skipped because they have a follow up task",
        "watch_out": ["Deletion cannot be undone.", "For users with ViewPrivateTouchpoints but not Admin, Developer or ManageTouchpoints, tasks they do not own or are not assigned to are silently skipped.", "Users without ViewPrivateTouchpoints get no role check at all; every id passed is deleted.", "Tasks dropped by the role check are not counted in the returned number.", "An id that does not exist raises an error."],
        "example": "# PERMANENT. Uncomment to run.\n# rejected = model.TaskNoteMassDelete([123, 124])\n# print rejected, \"could not be deleted\"\nprint \"Consider model.TaskNoteMassArchive instead\"",
        "see_also": ["model.TaskNoteMassArchive", "model.TaskNoteMassUnarchive"]
    },
    "model.TaskNoteMassUnarchive": {
        "category": "Tasks & Notes",
        "kind": "method",
        "arity": [1, 1],
        "tags": ["writes_data"],
        "summary": "Unarchives several tasks or notes at once.",
        "params": [
            ["taskNoteIdList", "list", "TaskNoteIds", True, None]
        ],
        "returns": "Integer 1",
        "watch_out": ["An unknown task id raises an error and nothing in the list is saved.", "The status is not changed."],
        "example": "# Uncomment to run:\n# model.TaskNoteMassUnarchive([123, 124, 125])\nprint \"Restores archived tasks to the active list\"",
        "see_also": ["model.TaskNoteMassArchive"]
    },
    "model.AppendIfBoth": {
        "category": "Text & Formatting",
        "kind": "method",
        "arity": [3, 3],
        "tags": ["read_only"],
        "summary": "Joins two strings with the join text only when both have a value.",
        "details": "If only one has a value, that one is returned alone.",
        "params": [
            ["s1", "string", "First string", True, None],
            ["join", "string", "Text placed between them", True, None],
            ["s2", "string", "Second string", True, None]
        ],
        "returns": "string",
        "watch_out": ["Text that is only spaces counts as empty.", "If neither has a value, s2 comes back as given, which may be None."],
        "example": "print model.AppendIfBoth(\"Nashville\", \", \", \"TN\")\nprint model.AppendIfBoth(\"Nashville\", \", \", \"\")",
        "see_also": ["model.Trim", "model.Replace"]
    },
    "model.FmtPhone": {
        "category": "Text & Formatting",
        "kind": "method",
        "arity": [1, 2],
        "tags": ["read_only"],
        "summary": "Formats a phone number from its digits, such as (615) 555-1234.",
        "details": "10 digits become (615) 555-1234, 7 digits become 555-1234, and 11 digits become +1 (615) 555-1234. If a prefix is given it is added in front with a space.",
        "params": [
            ["s", "string", "Phone number in any format", True, None],
            ["prefix", "string", "Label put in front, such as \"c\" or \"Cell:\"", False, "None"]
        ],
        "returns": "string",
        "watch_out": ["Returns an empty string for a blank number.", "Formats 7 to 13 digits; 11 digits becomes +1 (615) 555-1234. Any other length comes back exactly as passed in.", "The prefix is joined with a space, such as \"C (615) 555-1234\"."],
        "example": "print model.FmtPhone(\"6155551234\")\nprint model.FmtPhone(\"615.555.1234\", \"Cell:\")",
        "see_also": ["model.FmtZip"]
    },
    "model.FmtZip": {
        "category": "Text & Formatting",
        "kind": "method",
        "arity": [1, 1],
        "tags": ["read_only"],
        "summary": "Formats a zip code, turning 9 digits into 12345-6789.",
        "details": "5 digits stay as is. Anything else is returned unchanged.",
        "params": [
            ["s", "string", "Zip code", True, None]
        ],
        "returns": "string",
        "watch_out": ["Returns an empty string for a blank value.", "Anything that is not 5 or 9 digits comes back unchanged."],
        "example": "print model.FmtZip(\"370751234\")\nprint model.FmtZip(\"37075\")",
        "see_also": ["model.FmtPhone"]
    },
    "model.Markdown": {
        "category": "Text & Formatting",
        "kind": "method",
        "arity": [1, 1],
        "tags": ["read_only"],
        "summary": "Converts Markdown text to HTML.",
        "params": [
            ["text", "string", "Markdown text", True, None]
        ],
        "returns": "string (HTML)",
        "watch_out": ["Returns an empty string for None.", "Leading and trailing whitespace is trimmed before converting."],
        "example": "html = model.Markdown(\"# Heading\\n\\n**Bold** text\")\nprint html"
    },
    "model.Md5Hash": {
        "category": "Text & Formatting",
        "kind": "method",
        "arity": [1, 1],
        "tags": ["read_only"],
        "summary": "Returns the MD5 hash of the text (UTF-8) as 32 lowercase hex characters.",
        "params": [
            ["s", "string", "Text to hash", True, None]
        ],
        "returns": "string",
        "watch_out": ["Do not use it for passwords.", "Raises an error for None."],
        "example": "print model.Md5Hash(\"hello\")"
    },
    "model.RegexMatch": {
        "category": "Text & Formatting",
        "kind": "method",
        "arity": [2, 2],
        "tags": ["read_only"],
        "summary": "Returns the first match of a regular expression in the text, ignoring case.",
        "params": [
            ["s", "string", "Text to search", True, None],
            ["regex", "string", ".NET regular expression", True, None]
        ],
        "returns": "string",
        "watch_out": ["Returns an empty string when there is no match.", "A dot also matches line breaks.", "Returns only the whole match, not groups.", "Raises an error for None."],
        "example": "print model.RegexMatch(\"Room 214B\", r\"\\d+\")\nprint repr(model.RegexMatch(\"none here\", r\"\\d+\"))",
        "see_also": ["model.Replace"]
    },
    "model.Replace": {
        "category": "Text & Formatting",
        "kind": "method",
        "arity": [3, 3],
        "tags": ["read_only"],
        "summary": "Replaces every match of a regular expression with the replacement text.",
        "details": "The replacement can use $1 for groups.",
        "params": [
            ["text", "string", "Text to change", True, None],
            ["pattern", "string", ".NET regular expression", True, None],
            ["replacement", "string", "Replacement text; $1 inserts group 1", True, None]
        ],
        "returns": "string",
        "watch_out": ["The pattern is a .NET regex, so characters such as . ( ) $ need escaping.", "Matching is case sensitive, and a dot also matches line breaks.", "Use $1 in the replacement for a captured group.", "Raises an error for None text."],
        "example": "print model.Replace(\"Hello World\", \"World\", \"Church\")\nprint model.Replace(\"615-555-1234\", r\"(\\d{3})-(\\d{3})-(\\d{4})\", \"($1) $2-$3\")",
        "see_also": ["model.RegexMatch", "model.ReplaceCodeStr"]
    },
    "model.ReplaceCodeStr": {
        "category": "Text & Formatting",
        "kind": "method",
        "arity": [2, 2],
        "tags": ["read_only"],
        "summary": "Does plain text find and replace for several pairs written as \"find=replace,find2=replace2\".",
        "details": "Pairs are applied in order.",
        "params": [
            ["text", "string", "Text to change", True, None],
            ["codes", "string", "Pairs as \"find=replace\" separated by commas", True, None]
        ],
        "returns": "string",
        "watch_out": ["All whitespace is removed from the codes string first, so neither side can contain spaces.", "A pair without an equals sign raises an error.", "Finds are case sensitive. A comma cannot appear in either side.", "Text starting with // up to the next space or symbol is treated as a comment and dropped."],
        "example": "print model.ReplaceCodeStr(\"Hello {name}, your {item} is ready\", \"{name}=John,{item}=order\")",
        "see_also": ["model.Replace"]
    },
    "model.SpaceCamelCase": {
        "category": "Text & Formatting",
        "kind": "method",
        "arity": [1, 1],
        "tags": ["read_only"],
        "summary": "Adds a space before each capital letter that starts a new word, so MyVariableName becomes My Variable Name.",
        "params": [
            ["s", "string", "CamelCase text", True, None]
        ],
        "returns": "string",
        "watch_out": ["Runs of capitals stay together, so HTMLPage becomes HTML Page.", "Raises an error for None."],
        "example": "print model.SpaceCamelCase(\"MemberStatusId\")"
    },
    "model.Trim": {
        "category": "Text & Formatting",
        "kind": "method",
        "arity": [1, 1],
        "tags": ["read_only"],
        "summary": "Removes leading and trailing whitespace.",
        "params": [
            ["s", "string", "Text to trim", True, None]
        ],
        "returns": "string",
        "watch_out": ["Raises an error for None."],
        "example": "print \"[\" + model.Trim(\"  text  \") + \"]\""
    },
    "model.UrlEncode": {
        "category": "Text & Formatting",
        "kind": "method",
        "arity": [1, 1],
        "tags": ["read_only"],
        "summary": "URL encodes text for use in a query string.",
        "details": "Spaces become +.",
        "params": [
            ["s", "string", "Text to encode", True, None]
        ],
        "returns": "string",
        "watch_out": ["Spaces become + rather than %20.", "Returns None for None."],
        "example": "print model.UrlEncode(\"Jane Smith & Co\")",
        "see_also": ["model.CreateTinyUrl"]
    },
    "model.CreateTinyUrl": {
        "category": "Web & Integrations",
        "kind": "method",
        "arity": [1, 2],
        "tags": ["outside_service"],
        "summary": "Shortens a URL with the URL shortener service configured for your database.",
        "flow": ["Reads the shortener service address and token from your database settings, or from the server defaults if either is blank.", "Posts the token and URL to that service.", "Adds https:// if the answer does not start with http, and returns it."],
        "params": [
            ["url", "string", "URL to shorten", True, None]
        ],
        "returns": "string",
        "watch_out": ["Returns the original URL unchanged if no service is configured or the request fails.", "Only an empty answer falls back to the original URL. If the service answers with an error message, that text comes back as the link."],
        "example": "short = model.CreateTinyUrl(model.CmsHost + \"/Person2/{0}\".format(model.UserPeopleId))\nprint short",
        "see_also": ["model.GetAuthenticatedUrl", "model.SendSms"]
    },
    "model.GetAuthenticatedUrl": {
        "category": "Web & Integrations",
        "kind": "method",
        "arity": [2, 3],
        "tags": ["read_only"],
        "summary": "Builds a link that signs the person in automatically and then opens the given page.",
        "details": "The link can be shortened. The first argument can be a PeopleId or a person object.",
        "flow": ["Finds the person's user account, preferring one with the Access role.", "Saves a sign in token in the OneTimeLinks table, good for 7 days and usable more than once.", "Builds a /Logon link with that token that returns to your page, then shortens it with model.CreateTinyUrl when shorten is True."],
        "params": [
            ["peopleId", "integer", "PeopleId the link signs in as", True, None],
            ["url", "string", "Page to open after sign in, such as \"/Person2/Current\"", True, None],
            ["shorten", "boolean", "Return a shortened link", False, "True"]
        ],
        "returns": "String URL",
        "watch_out": ["Anyone holding the link can act as that person, so only send it to that person.", "If the person has no user account, the URL comes back unchanged and will not sign them in.", "The link works repeatedly for 7 days, not just once.", "An unknown PeopleId raises an error."],
        "example": "link = model.GetAuthenticatedUrl(model.UserPeopleId, \"/Person2/Current\", True)\nprint link",
        "signatures": ["model.GetAuthenticatedUrl(peopleId, url, shorten=True)", "model.GetAuthenticatedUrl(person, url, shorten=True)"],
        "see_also": ["model.CreateTinyUrl", "model.GetPerson"]
    },
    "model.RestDelete": {
        "category": "Web & Integrations",
        "kind": "method",
        "arity": [2, 4],
        "tags": ["outside_service"],
        "summary": "Sends an HTTP DELETE request and returns the response body as text.",
        "details": "Optional user and password add Basic authentication.",
        "params": [
            ["url", "string", "Full URL", True, None],
            ["headers", "dict", "Request headers. Required; use {} for none", True, None],
            ["user", "string", "Basic auth user name", False, "None"],
            ["password", "string", "Basic auth password", False, "None"]
        ],
        "returns": "string",
        "watch_out": ["The headers dictionary is required; pass {} for none.", "Basic authentication is only used when both user and password are given.", "Returns the response body even for an error status such as 404 or 500. A failed connection returns an empty result instead of raising."],
        "example": "body = model.RestDelete(\"https://httpbin.org/delete\", {})\nprint body",
        "signatures": ["model.RestDelete(url, headers, user=None, password=None)"],
        "see_also": ["model.RestGet", "model.RestPost", "model.RestPostJson"]
    },
    "model.RestGet": {
        "category": "Web & Integrations",
        "kind": "method",
        "arity": [2, 4],
        "tags": ["outside_service"],
        "summary": "Sends an HTTP GET request and returns the response body as text.",
        "details": "Optional user and password add Basic authentication.",
        "flow": ["Adds Basic authentication when both user and password are given.", "Adds each header from the dictionary. Every value must already be a string.", "Sends the GET request and returns the response body as a string."],
        "params": [
            ["url", "string", "Full URL", True, None],
            ["headers", "dict", "Request headers. Required; use {} for none", True, None],
            ["user", "string", "Basic auth user name", False, "None"],
            ["password", "string", "Basic auth password", False, "None"]
        ],
        "returns": "string",
        "watch_out": ["The headers dictionary is required; pass {} for none.", "HTTP errors such as 404 or 500 do not raise. You get the error page's body back, so check the text before trusting it.", "A connection failure does not raise either. It returns an empty string.", "A header value that is a number raises an error. Convert it with str() first."],
        "example": "body = model.RestGet(\"https://httpbin.org/get\", {\"Accept\": \"application/json\"})\nprint body",
        "signatures": ["model.RestGet(url, headers, user=None, password=None)"],
        "see_also": ["model.RestPost", "model.RestPostJson", "model.RestDelete"]
    },
    "model.RestPost": {
        "category": "Web & Integrations",
        "kind": "method",
        "arity": [1, 5],
        "tags": ["outside_service"],
        "summary": "Sends an HTTP POST request and returns the response body as text.",
        "details": "The body is sent with the Content-Type from headers. Optional user and password add Basic authentication.",
        "flow": ["Adds Basic authentication when both user and password are given.", "Adds each header. Every header value must already be a string.", "If there is a body, reads the Content-Type header (spelled exactly Content-Type). Without one it uses application/x-www-form-urlencoded, and a dict body is turned into name=value pairs joined with &.", "Sends the body as text with that content type and returns the response body as a string."],
        "params": [
            ["url", "string", "Full URL", True, None],
            ["headers", "dict", "Request headers", False, "None"],
            ["body", "any", "Body: a string, or a dict for a form post", False, "None"],
            ["user", "string", "Basic auth user name", False, "None"],
            ["password", "string", "Basic auth password", False, "None"]
        ],
        "returns": "string",
        "watch_out": ["With no Content-Type the body is sent as a form, and a dict body is turned into name=value pairs.", "Use model.RestPostJson for JSON.", "HTTP errors such as 404 or 500 do not raise. You get the error page's body back, so check the text before trusting it.", "A connection failure does not raise either. It returns an empty string.", "A dict body with Content-Type application/json is not turned into JSON; it is sent as Python's text form of the dict. Use json.dumps or model.RestPostJson.", "The Content-Type header name is case sensitive here. content-type is not seen, so the body is labeled as a form post.", "A header value that is a number raises an error. Convert it with str() first."],
        "example": "body = model.RestPost(\"https://httpbin.org/post\", {}, {\"name\": \"Test\", \"id\": \"{0}\".format(model.UserPeopleId)})\nprint body",
        "signatures": ["model.RestPost(url, headers=None, body=None, user=None, password=None)"],
        "see_also": ["model.RestPostJson", "model.RestPostXml", "model.RestGet"]
    },
    "model.RestPostJson": {
        "category": "Web & Integrations",
        "kind": "method",
        "arity": [3, 5],
        "tags": ["outside_service"],
        "summary": "Sends an HTTP POST with the object serialized as a JSON body and returns the response body as text.",
        "details": "Optional user and password add Basic authentication.",
        "params": [
            ["url", "string", "Full URL", True, None],
            ["headers", "dict", "Request headers. Required; use {} for none", True, None],
            ["obj", "any", "Object to send as JSON, such as a dict", True, None],
            ["user", "string", "Basic auth user name", False, "None"],
            ["password", "string", "Basic auth password", False, "None"]
        ],
        "returns": "string",
        "watch_out": ["The headers dictionary is required; pass {} for none.", "Basic authentication is only used when both user and password are given.", "Returns the response body even for an error status such as 404 or 500. A failed connection returns an empty result instead of raising."],
        "example": "body = model.RestPostJson(\"https://httpbin.org/post\", {}, {\"name\": \"Test\", \"id\": model.UserPeopleId})\nprint body",
        "signatures": ["model.RestPostJson(url, headers, obj, user=None, password=None)"],
        "see_also": ["model.RestPost", "model.RestPostXml", "model.JsonSerialize"]
    },
    "model.RestPostXml": {
        "category": "Web & Integrations",
        "kind": "method",
        "arity": [3, 5],
        "tags": ["outside_service"],
        "summary": "Sends an HTTP POST with an XML body (Content-Type text/xml) and returns the response body as text.",
        "details": "Optional user and password add Basic authentication.",
        "params": [
            ["url", "string", "Full URL", True, None],
            ["headers", "dict", "Request headers. Required; use {} for none", True, None],
            ["body", "string", "XML text", True, None],
            ["user", "string", "Basic auth user name", False, "None"],
            ["password", "string", "Basic auth password", False, "None"]
        ],
        "returns": "string",
        "watch_out": ["The headers dictionary is required; pass {} for none.", "Basic authentication is only used when both user and password are given.", "Returns the response body even for an error status such as 404 or 500. A failed connection returns an empty result instead of raising."],
        "example": "xml = \"<person><id>{0}</id></person>\".format(model.UserPeopleId)\nbody = model.RestPostXml(\"https://httpbin.org/post\", {}, xml)\nprint body",
        "signatures": ["model.RestPostXml(url, headers, body, user=None, password=None)"],
        "see_also": ["model.RestPost", "model.RestPostJson"]
    },
    "model.UploadExcelFromSqlToDropBox": {
        "category": "Web & Integrations",
        "kind": "method",
        "arity": [3, 4],
        "tags": ["outside_service", "known_issue"],
        "summary": "Runs a SQL Script from Special Content, builds an Excel file from the result and uploads it to Dropbox.",
        "details": "The script is named, not raw SQL. Needs the DropBoxAccessToken setting. If the script uses @\x71tagid, the people from the savedQuery argument are tagged and passed in. Values added with model.DictionaryAdd are passed as SQL parameters.",
        "flow": ["Reads the DropBoxAccessToken setting and the named SQL Script from Special Content. A missing script raises \"no sql script found\".", "Passes every model.DictionaryAdd value as a SQL parameter. If the script uses @\x71tagid, tags the people from savedQuery and passes that tag id.", "Runs the SQL and builds an Excel file from the result.", "Uploads the file to targetpath/filename in Dropbox, overwriting any file there."],
        "params": [
            ["savedQuery", "string", "Search Builder code or saved search name used to fill @\x71tagid. Omit this argument when the script does not use @\x71tagid.", False, None],
            ["sqlscript", "string", "Name of a SQL Script in Special Content", True, None],
            ["targetpath", "string", "Dropbox folder, such as \"/Reports\"", True, None],
            ["filename", "string", "File name, such as \"members.xlsx\"", True, None]
        ],
        "returns": "None",
        "watch_out": ["Overwrites any Dropbox file with the same name.", "Sends data off site.", "It calls Dropbox's retired \"2-beta-2\" upload endpoint, so it likely fails today.", "The SQL runs on the regular connection, not the read only one.", "An upload failure raises an error."],
        "example": "# Sends data to Dropbox; needs the DropBoxAccessToken setting:\n# model.UploadExcelFromSqlToDropBox(\"MemberStatusId = 10\", \"MyMemberReport\", \"/Reports\", \"members.xlsx\")\n# model.UploadExcelFromSqlToDropBox(\"MyMemberReport\", \"/Reports\", \"members.xlsx\")\nprint model.Setting(\"DropBoxAccessToken\", \"\") != \"\"",
        "signatures": ["model.UploadExcelFromSqlToDropBox(savedQuery, sqlscript, targetpath, filename)", "model.UploadExcelFromSqlToDropBox(sqlscript, targetpath, filename)"],
        "see_also": ["model.UploadExcelFromSqlToFtp", "model.DictionaryAdd"]
    },
    "model.UploadExcelFromSqlToFtp": {
        "category": "Web & Integrations",
        "kind": "method",
        "arity": [5, 5],
        "tags": ["outside_service"],
        "summary": "Builds an Excel file from a SQL query and uploads it to an FTP server.",
        "details": "The first argument can be raw SQL or, if it is 75 characters or fewer, the name of a SQL Script in Special Content. Returns the uploaded file's URL with ftp:// changed to http://.",
        "flow": ["If the first argument is 75 characters or fewer, looks for a SQL Script by that name in Special Content, and uses it as SQL only if it contains \"select\".", "Removes any line containing LinkForNext, runs the SQL and builds an Excel file.", "Uploads it to targetpath plus filename with your FTP user name and password.", "Returns that address with ftp:// changed to http://."],
        "params": [
            ["sqlscript", "string", "SQL text, or the name of a SQL Script in Special Content", True, None],
            ["username", "string", "FTP user name", True, None],
            ["password", "string", "FTP password", True, None],
            ["targetpath", "string", "FTP folder URL, such as \"ftp://ftp.example.com/uploads\"", True, None],
            ["filename", "string", "File name, such as \"people.xlsx\"", True, None]
        ],
        "returns": "string (URL)",
        "watch_out": ["Sends data off site.", "targetpath must be a full ftp:// URL.", "The SQL runs on the regular connection, not the read only one.", "The returned http address is only a guess; it works only if your server also serves that folder over the web.", "An upload failure raises an error."],
        "example": "sql = \"SELECT TOP 100 PeopleId, Name FROM People\"\n# Sends data to an outside server:\n# url = model.UploadExcelFromSqlToFtp(sql, \"ftpuser\", \"secret\", \"ftp://ftp.example.com/uploads\", \"people.xlsx\")\n# print url\nprint \"Would upload\", len(list(q.QuerySql(sql))), \"rows\"",
        "see_also": ["model.UploadExcelFromSqlToDropBox"]
    }
}

# Members of model and q that exist but are not meant for scripts, and why
INTERNAL_NOTES = {
    'model.Birthday': 'A field on the rows model.BirthdayList returns, not something to call on model.',
    'model.Name': 'A field on the rows model.BirthdayList returns, not something to call on model.',
    'model.PeopleId': 'A field on the rows model.BirthdayList returns. For the signed in person use model.UserPeopleId.',
    'model.DebugWriteLine': 'Writes to a developer debugger that is not visible on hosted TouchPoint. Use print instead.',
    'model.Encode': 'Used inside the template engine. Not meant for scripts.',
    'model.ExecutePython': 'Used by TouchPoint to start scripts. To run another script, use model.CallScript.',
    'model.ExecutePythonFilex': 'Used by TouchPoint to start scripts from the server. To run another script, use model.CallScript.',
    'model.ExecutePythonFromFileWithClass': 'Used by registration event scripts. To run another script, use model.CallScript.',
    'model.RunScript': 'Used by TouchPoint to start scripts. To run another script, use model.CallScript.',
    'model.RunScriptFile': 'Runs a file from the server, which hosted TouchPoint does not allow. Use model.CallScript.',
    'model.HtmlTable': 'Used inside model.SqlGrid. It needs an object a script cannot create. Use model.SqlGrid.',
    'model.Table': 'Used inside model.SqlGrid. It needs an object a script cannot create. Use model.SqlGrid.',
    'model.PageBreakTables': 'Used inside SQL reports. It needs objects a script cannot create.',
    'model.NewBundleDetail': 'Used inside the contribution calls. It needs an object a script cannot create. Use model.AddContribution.',
    'model.RegisterHelpers': 'Used inside the template engine. Use model.RenderTemplate.',
    'model.Replacements': 'Not used by TouchPoint; always empty.',
    'model.SmtpDebug': 'A developer switch. On a hosted site it stops email from being delivered, so do not set it.',
    'model.StartsEndsWith': "A formatting helper for model.SqlGrid. Use Python's startswith and endswith.",
    'model.instance': 'Used by registration event scripts. Not meant for general scripts.',
    'q.AddP1Parameter': 'Used inside q.QuerySql when you pass one value. Call q.QuerySql instead.',
    'q.OldVitalStats': 'An old reporting helper that needs an object a script cannot create.',
    'q.RemoveDeclaration': 'Exists only in developer builds, not on hosted TouchPoint.',
}

if _allowed and not _is_run:
    try:
        _apx_user = str(int(model.UserPeopleId or 0))
    except Exception:
        _apx_user = '0'

    # TouchPoint opens any script whose text mentions model.Form at
    # /PyScriptForm, and a page load there shows only model.Form (printed
    # output is dropped). So the page is printed into a buffer and handed
    # to model.Form at the end of this file.
    import sys
    from StringIO import StringIO
    _page_buf = StringIO()
    _page_stdout = sys.stdout
    sys.stdout = _page_buf
    # Continue with the HTML interface...
    print """<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>TouchPoint API Explorer</title>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/codemirror/5.65.2/codemirror.min.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/codemirror/5.65.2/theme/monokai.min.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/codemirror/5.65.2/addon/dialog/dialog.min.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/codemirror/5.65.2/addon/hint/show-hint.min.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/codemirror/5.65.2/addon/lint/lint.min.css">
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        
        html, body { 
            margin: 0;
            padding-bottom: 50px;
        }
        
        body { 
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Arial, sans-serif;
            display: flex;
            flex-direction: column;
        }
        
        .header {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 10px 20px;
            flex-shrink: 0;
            height: 50px;
        }
        
        .header h1 { font-size: 16px; margin: 0; }
        .header p { font-size: 11px; opacity: 0.9; margin: 2px 0 0 0; }
        
        .main-container {
            display: flex;
            flex: 1;
            min-height: 0;
            width: 100%;
            height: calc(100vh - 50px);
            overflow: hidden;
        }
        
        .sidebar {
            width: 280px;
            min-width: 50px;
            background: #f8f9fa;
            border-right: 1px solid #dee2e6;
            display: flex;
            flex-direction: column;
            overflow: hidden;
            position: relative;
            flex-shrink: 0;
        }
        
        
        .sidebar-resize-handle {
            position: absolute;
            right: -3px;
            top: 0;
            bottom: 0;
            width: 6px;
            background: transparent;
            cursor: col-resize;
            z-index: 10;
        }
        
        .sidebar-resize-handle:hover {
            background: #007bff;
            opacity: 0.5;
        }
        
        .search-box {
            padding: 10px;
            border-bottom: 1px solid #dee2e6;
            flex-shrink: 0;
        }
        
        .search-box input {
            width: 100%;
            padding: 5px;
            border: 1px solid #ced4da;
            border-radius: 3px;
            font-size: 12px;
        }
        
        .method-list {
            flex: 1;
            overflow-y: auto;
            overflow-x: hidden;
            padding: 10px;
        }
        
        .category-section {
            margin-bottom: 15px;
        }
        
        .category-header {
            font-size: 11px;
            text-transform: uppercase;
            color: #6c757d;
            margin-bottom: 5px;
            font-weight: 600;
            padding: 5px;
            background: #e9ecef;
            border-radius: 3px;
            cursor: pointer;
            user-select: none;
            position: relative;
            padding-left: 20px;
        }
        
        .category-header::before {
            content: '▼';
            position: absolute;
            left: 5px;
            transition: transform 0.2s;
        }
        
        .category-header.collapsed::before {
            transform: rotate(-90deg);
        }
        
        .category-content {
            display: block;
            overflow: hidden;
            transition: max-height 0.3s ease-out;
            max-height: 1000px;
        }
        
        .category-content.collapsed {
            max-height: 0;
        }
        
        .category-content.collapsed {
            display: none;
        }
        
        .method-item {
            padding: 4px 8px;
            cursor: pointer;
            font-size: 11px;
            font-family: monospace;
            border-radius: 3px;
            margin-bottom: 2px;
            border-left: 2px solid transparent;
        }
        
        .method-item:hover {
            background: #e9ecef;
            border-left-color: #667eea;
        }
        
        .method-item.active {
            background: #667eea;
            color: white;
            border-left-color: #764ba2;
        }
        
        .method-item.documented {
            font-weight: bold;
        }
        
        .method-item.undocumented {
            opacity: 0.8;
            font-style: italic;
        }
        
        .doc-panel {
            width: 560px;
            min-width: 50px;
            padding: 15px;
            background: white;
            border-right: 1px solid #dee2e6;
            overflow-y: auto;
            flex-shrink: 0;
            position: relative;
        }
        
        
        .doc-resize-handle {
            position: absolute;
            right: -3px;
            top: 0;
            bottom: 0;
            width: 6px;
            background: transparent;
            cursor: col-resize;
            z-index: 10;
        }
        
        .doc-resize-handle:hover {
            background: #007bff;
            opacity: 0.5;
        }
        
        .doc-title {
            font-size: 14px;
            font-weight: bold;
            margin-bottom: 5px;
            font-family: monospace;
            color: #2c3e50;
        }
        
        .doc-category {
            font-size: 10px;
            color: #667eea;
            text-transform: uppercase;
            margin-bottom: 5px;
        }
        
        .doc-desc {
            font-size: 12px;
            color: #6c757d;
            margin-bottom: 15px;
        }
        
        .param {
            background: #f8f9fa;
            padding: 6px;
            margin-bottom: 4px;
            border-left: 2px solid #667eea;
            font-size: 11px;
        }
        
        .param-name {
            font-weight: bold;
            font-family: monospace;
            color: #2c3e50;
        }
        
        .param-type {
            color: #667eea;
            font-size: 10px;
            margin-left: 5px;
        }
        
        .example-box {
            background: #2d2d30;
            color: #d4d4d4;
            padding: 10px;
            border-radius: 3px;
            font-family: monospace;
            font-size: 11px;
            white-space: pre;
            overflow-x: auto;
        }
        
        .editor-panel {
            flex: 1 1 auto;
            display: flex;
            flex-direction: column;
            min-width: 300px;
            width: 100%;
            min-height: 0;  /* Critical for flexbox children */
            overflow: hidden;
        }
        
        .editor-toolbar {
            padding: 10px;
            background: #f8f9fa;
            border-bottom: 1px solid #dee2e6;
            display: flex;
            gap: 8px;
            flex-wrap: wrap;
            flex-shrink: 0;
        }
        
        .apx-btn {
            padding: 5px 12px;
            border: none;
            border-radius: 3px;
            font-size: 12px;
            cursor: pointer;
        }
        
        .apx-btn-primary {
            background: #667eea;
            color: white;
        }
        
        .apx-btn-secondary {
            background: #6c757d;
            color: white;
        }
        
        .apx-spacer { flex: 1; }
        .apx-badges { display: flex; flex-wrap: wrap; gap: 4px; margin: 6px 0 8px; }
        .apx-badge { font-size: 10px; font-weight: 600; text-transform: uppercase; letter-spacing: 0.03em; padding: 2px 7px; border-radius: 10px; border: 1px solid transparent; }
        .apx-badge.kind { background: #eef0ff; color: #4c51bf; }
        .apx-badge.read_only { background: #e6fcf5; color: #087f5b; }
        .apx-badge.writes_data { background: #fff4e6; color: #c2410c; }
        .apx-badge.sends_message { background: #ffe3e3; color: #c92a2a; }
        .apx-badge.slow { background: #fff9db; color: #946200; }
        .apx-badge.deprecated { background: #f1f3f5; color: #495057; text-decoration: line-through; }
        .apx-badge.known_issue { background: #ffe3e3; color: #a61e4d; }
        .apx-badge.outside_service { background: #e7f5ff; color: #1864ab; }
        .apx-sig { display: flex; align-items: center; gap: 8px; background: #f8f9fa; border: 1px solid #e9ecef; border-radius: 4px; padding: 6px 8px; margin-bottom: 8px; }
        .apx-sig code { font-family: monospace; font-size: 12px; color: #212529; white-space: pre-wrap; word-break: break-word; flex: 1; }
        .apx-details { margin: 6px 0 10px; font-size: 13px; color: #495057; }
        .apx-details summary { cursor: pointer; color: #667eea; font-size: 12px; }
        .apx-details div { margin-top: 6px; line-height: 1.45; }
        .apx-params { width: 100%; border-collapse: collapse; font-size: 12px; margin-bottom: 6px; }
        .apx-params th { text-align: left; font-weight: 600; color: #6c757d; border-bottom: 1px solid #dee2e6; padding: 4px 6px; }
        .apx-params td { border-bottom: 1px solid #f1f3f5; padding: 5px 6px; vertical-align: top; }
        .apx-params td.n { font-family: monospace; font-weight: 600; white-space: nowrap; }
        .apx-params td.t { color: #667eea; white-space: nowrap; }
        .apx-req { color: #c92a2a; font-weight: 600; }
        .apx-opt { color: #868e96; }
        .apx-watch { background: #fff9db; border: 1px solid #ffe066; border-radius: 4px; padding: 8px 10px 6px; margin: 10px 0; }
        .apx-watch h4, .apx-seealso h4 { font-size: 12px; margin: 0 0 6px; }
        .apx-watch ul { margin: 0; padding-left: 18px; font-size: 12px; line-height: 1.45; }
        .apx-seealso { margin: 10px 0; font-size: 12px; }
        .apx-flow { margin: 10px 0; }
        .apx-flow h4 { font-size: 12px; margin: 0 0 6px; }
        .apx-flow ol { margin: 0; padding-left: 20px; font-size: 12px; line-height: 1.5; color: #343a40; }
        .apx-flow li { margin-bottom: 4px; }
        .apx-seealso a { display: inline-block; margin: 0 10px 4px 0; color: #667eea; cursor: pointer; font-family: monospace; }
        .apx-ex-actions { float: right; font-weight: normal; }
        .method-item.deprecated { text-decoration: line-through; opacity: 0.7; }
        .apx-out-meta { font-size: 12px; color: #6c757d; align-self: center; }
        .apx-out-mode { font-size: 12px; padding: 3px 6px; border: 1px solid #ced4da; border-radius: 3px; background: #fff; }
        .apx-tabs { display: flex; align-items: flex-end; gap: 2px; padding: 6px 8px 0; background: #f8f9fa; border-bottom: 1px solid #dee2e6; overflow-x: auto; flex-shrink: 0; }
        .apx-tab { display: flex; align-items: center; gap: 6px; padding: 5px 8px 5px 12px; font-size: 12px; color: #495057; background: #e9ecef; border: 1px solid #dee2e6; border-bottom: none; border-radius: 5px 5px 0 0; cursor: pointer; white-space: nowrap; max-width: 220px; margin-bottom: -1px; user-select: none; }
        .apx-tab:hover { background: #dee2e6; }
        .apx-tab.active { background: #272822; color: #f8f8f2; border-color: #272822; }
        .apx-tab-name { overflow: hidden; text-overflow: ellipsis; }
        .apx-tab-x { opacity: 0.6; padding: 0 2px; }
        .apx-tab-x:hover { opacity: 1; color: #dc3545; }
        .apx-tab-new { padding: 4px 10px; font-size: 16px; color: #6c757d; cursor: pointer; align-self: center; }
        .apx-tab-new:hover { color: #667eea; }
        .apx-warn { white-space: normal; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Arial, sans-serif; font-size: 13px; color: #1e1e1e; background: #fff3cd; border: 1px solid #ffc107; border-radius: 4px; padding: 10px 12px; }
        .apx-warn code { background: #ffe69c; padding: 0 3px; border-radius: 2px; }
        .apx-warn button { margin: 8px 8px 0 0; }
        .apx-html-frame { width: 100%; height: 100%; min-height: 100%; border: 0; background: #fff; }
        .apx-modal { display: none; position: fixed; inset: 0; background: rgba(0, 0, 0, 0.45); z-index: 2000; }
        .apx-modal-box { position: absolute; top: 50%; left: 50%; transform: translate(-50%, -50%); width: min(640px, calc(100vw - 32px)); max-height: 80vh; overflow-y: auto; background: #fff; border-radius: 8px; padding: 16px 18px; box-shadow: 0 20px 40px rgba(0, 0, 0, 0.25); font-size: 13px; }
        .apx-modal-head { display: flex; justify-content: space-between; align-items: center; margin-bottom: 10px; font-size: 15px; }
        .apx-link { color: #667eea; cursor: pointer; font-size: 12px; font-weight: normal; margin-left: 8px; }
        .apx-search { width: 100%; padding: 6px 8px; border: 1px solid #ced4da; border-radius: 4px; margin-bottom: 8px; }
        .apx-lib-title { font-weight: 600; margin: 10px 0 4px; color: #343a40; }
        .apx-lib-item { padding: 6px 8px; border-bottom: 1px solid #f1f3f5; cursor: pointer; }
        .apx-lib-item:hover { background: #eef0ff; }
        .apx-lib-code { font-family: monospace; font-size: 12px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; color: #212529; }
        .apx-lib-meta { font-size: 11px; color: #6c757d; }
        .apx-lib-err { color: #c92a2a; }
        .apx-note { font-size: 11px; color: #6c757d; margin-top: 10px; }
        .apx-sigs { font-family: monospace; font-size: 11px; background: #f1f3f5; padding: 6px 8px; border-radius: 3px; margin-top: 6px; white-space: pre-wrap; }

        .editor-split {
            flex: 1 1 auto;
            display: flex;
            flex-direction: column;
            min-height: 0;  /* Critical for proper flex sizing */
            overflow: hidden;
        }
        
        .output-panel {
            flex: 0 0 200px;
            min-height: 100px;
            max-height: 400px;
            background: #1e1e1e;
            color: #d4d4d4;
            padding: 10px;
            font-family: monospace;
            font-size: 11px;
            overflow-y: auto;
            overflow-x: auto;  /* Allow horizontal scroll */
            white-space: pre;   /* Don't wrap output text */
        }

        .resize-handle {
            height: 8px;
            background: linear-gradient(to bottom, #495057 0%, #6c757d 50%, #495057 100%);
            cursor: ns-resize;
            flex-shrink: 0;
            z-index: 10;
            transition: all 0.2s;
            position: relative;
            display: flex;
            align-items: center;
            justify-content: center;
        }

        .resize-handle::before {
            content: '⋯';
            color: #adb5bd;
            font-size: 16px;
            letter-spacing: 2px;
            transform: rotate(90deg);
        }

        .resize-handle:hover {
            background: linear-gradient(to bottom, #667eea 0%, #764ba2 50%, #667eea 100%);
            height: 10px;
        }

        .resize-handle:hover::before {
            color: #fff;
        }
        
        .code-container {
            flex: 1 1 auto;
            min-height: 0;  /* Critical - allows container to shrink */
            position: relative;
            overflow: hidden;  /* Let CodeMirror handle scrolling */
        }
        
        .CodeMirror {
            height: 100% !important;
            position: absolute;
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;
        }
        
        .CodeMirror-scroll {
            height: 100% !important;
            overflow: auto !important;
        }
        
        .CodeMirror-sizer {
            padding-bottom: 100px !important;  /* Extra space at bottom */
        }
        
        .CodeMirror-vscrollbar, .CodeMirror-hscrollbar {
            display: block !important;
        }
        
        .stats {
            padding: 10px;
            background: #f8f9fa;
            font-size: 11px;
            color: #6c757d;
            border-top: 1px solid #dee2e6;
            flex-shrink: 0;
        }
        
        /* Modal Styles */
        .modal {
            display: none;
            position: fixed;
            z-index: 1000;
            left: 0;
            top: 0;
            width: 100%;
            height: 100%;
            background-color: rgba(0,0,0,0.5);
        }
        
        .modal-content {
            background-color: #fefefe;
            margin: 10% auto;
            padding: 0;
            border-radius: 8px;
            width: 600px;
            max-width: 90%;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        }
        
        .modal-header {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 15px 20px;
            border-radius: 8px 8px 0 0;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }
        
        .modal-title {
            font-size: 18px;
            font-weight: 600;
        }
        
        .close {
            color: white;
            font-size: 28px;
            font-weight: bold;
            border: none;
            background: none;
            cursor: pointer;
            padding: 0;
            width: 30px;
            height: 30px;
            display: flex;
            align-items: center;
            justify-content: center;
        }
        
        .close:hover,
        .close:focus {
            opacity: 0.8;
        }
        
        .shortcut-list {
            display: grid;
            gap: 10px;
            padding: 20px;
        }
        
        .shortcut-item {
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding: 8px 0;
            border-bottom: 1px solid #f0f0f0;
            font-size: 14px;
        }
        
        .shortcut-key {
            font-family: monospace;
            background: #667eea;
            color: white;
            padding: 2px 6px;
            border-radius: 3px;
            font-size: 12px;
            margin: 0 2px;
        }
        
        /* Custom linting styles for dark theme */
        .CodeMirror-lint-tooltip {
            background: #2d2d30;
            color: #f0f0f0;
            border: 1px solid #454545;
            border-radius: 4px;
            font-size: 12px;
            padding: 4px 6px;
            z-index: 10000;
        }
        
        .CodeMirror-lint-mark-error {
            border-bottom: 2px solid #ff5555;
        }
        
        .CodeMirror-lint-mark-warning {
            border-bottom: 2px solid #f1fa8c;
        }
        
        .CodeMirror-lint-mark-info {
            border-bottom: 1px dotted #50fa7b;
        }
        
        .CodeMirror-lint-marker-error {
            color: #ff5555;
        }
        
        .CodeMirror-lint-marker-warning {
            color: #f1fa8c;
        }
        
        .CodeMirror-lint-marker-info {
            color: #50fa7b;
        }
        
        .CodeMirror-gutters {
            background-color: #282a36;
            border-right: 1px solid #44475a;
        }
        
        /* Enhanced autocomplete styles */
        .CodeMirror-hints {
            background: #2d2d30;
            border: 1px solid #454545;
            border-radius: 4px;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.3);
            font-family: 'Consolas', 'Monaco', monospace;
            font-size: 13px;
            max-height: 300px;
            overflow-y: auto;
        }
        
        .CodeMirror-hint {
            color: #f0f0f0;
            padding: 4px 8px;
            line-height: 1.4;
            cursor: pointer;
        }
        
        .CodeMirror-hint-active {
            background: #094771;
            color: #ffffff;
        }
        
        /* Custom hint styles */
        .hint-method {
            display: flex;
            justify-content: space-between;
            align-items: center;
        }
        
        .hint-name {
            font-weight: bold;
            color: #4EC9B0;
        }
        
        .hint-category {
            font-size: 10px;
            background: #667eea;
            color: white;
            padding: 2px 6px;
            border-radius: 3px;
            margin-left: 10px;
        }
        
        .hint-desc {
            font-size: 11px;
            color: #a0a0a0;
            margin-top: 2px;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }
        
        .CodeMirror-hint.touchpoint-hint {
            padding: 6px 8px;
        }
        
        .CodeMirror-hint.object-hint .hint-name {
            color: #569CD6;
        }
        
        .CodeMirror-hint.keyword-hint .hint-name {
            color: #C586C0;
        }
        
        .CodeMirror-hint.method-hint .hint-name {
            color: #DCDCAA;
        }
        
        /* Tooltip for parameter info */
        .parameter-hint {
            position: absolute;
            background: #2d2d30;
            border: 1px solid #454545;
            padding: 8px;
            border-radius: 4px;
            font-size: 12px;
            color: #f0f0f0;
            z-index: 10001;
            max-width: 400px;
        }
        
        .parameter-hint .param-name {
            color: #9CDCFE;
            font-weight: bold;
        }
        
        .parameter-hint .param-type {
            color: #4EC9B0;
        }
    </style>
</head>
<body>
    <!-- History and snippets -->
    <div id="apxLibrary" class="apx-modal" onclick="if (event.target === this) apxCloseLibrary()">
        <div class="apx-modal-box">
            <div class="apx-modal-head">
                <b>History &amp; snippets</b>
                <a class="apx-link" onclick="apxCloseLibrary()">close</a>
            </div>
            <input class="apx-search" id="apxLibSearch" placeholder="Search..." oninput="apxRenderLibrary()">
            <div class="apx-lib-title">Saved snippets</div>
            <div id="apxSnippetList"></div>
            <div class="apx-lib-title">Recent runs <a class="apx-link" onclick="apxClearHistory(this)">clear</a></div>
            <div id="apxHistoryList"></div>
            <div class="apx-note">Kept in this browser only. Click one to open it in a new tab.</div>
        </div>
    </div>

    <!-- Keyboard Shortcuts Modal -->
    <div id="shortcutsModal" class="modal">
        <div class="modal-content">
            <div class="modal-header">
                <div class="modal-title">⌨️ Keyboard Shortcuts</div>
                <button class="close" onclick="closeShortcutsModal()">&times;</button>
            </div>
            <div class="shortcut-list">
                <div class="shortcut-item">
                    <span>Run all code</span>
                    <span><span class="shortcut-key">Ctrl+Enter</span> or <span class="shortcut-key">Cmd+Enter</span></span>
                </div>
                <div class="shortcut-item">
                    <span>Run selected code (or all if none selected)</span>
                    <span><span class="shortcut-key">Shift+Enter</span></span>
                </div>
                <div class="shortcut-item">
                    <span>Clear editor</span>
                    <span><span class="shortcut-key">Ctrl+L</span> or <span class="shortcut-key">Cmd+L</span></span>
                </div>
                <div class="shortcut-item">
                    <span>Indent / Insert tab (4 spaces)</span>
                    <span><span class="shortcut-key">Tab</span></span>
                </div>
                <div class="shortcut-item">
                    <span>Unindent</span>
                    <span><span class="shortcut-key">Shift+Tab</span></span>
                </div>
                <div class="shortcut-item">
                    <span>Toggle comment</span>
                    <span><span class="shortcut-key">Ctrl+/</span> or <span class="shortcut-key">Cmd+/</span></span>
                </div>
                <div class="shortcut-item">
                    <span>Autocomplete</span>
                    <span><span class="shortcut-key">Ctrl+Space</span></span>
                </div>
                <div class="shortcut-item">
                    <span>Find</span>
                    <span><span class="shortcut-key">Ctrl+F</span> or <span class="shortcut-key">Cmd+F</span></span>
                </div>
                <div class="shortcut-item">
                    <span>Replace</span>
                    <span><span class="shortcut-key">Ctrl+H</span> or <span class="shortcut-key">Cmd+H</span></span>
                </div>
                <div class="shortcut-item">
                    <span>Undo</span>
                    <span><span class="shortcut-key">Ctrl+Z</span> or <span class="shortcut-key">Cmd+Z</span></span>
                </div>
                <div class="shortcut-item">
                    <span>Redo</span>
                    <span><span class="shortcut-key">Ctrl+Y</span> or <span class="shortcut-key">Cmd+Shift+Z</span></span>
                </div>
            </div>
        </div>
    </div>
    
    <div class="header">
        <h1>🚀 TouchPoint API Explorer</h1>
        <p>Swaby's reference with """ + str(len(method_docs)) + """ documented methods</p>
    </div>
    
    <div class="main-container">
        <!-- Sidebar -->
        <div class="sidebar" id="sidebar">
            <div class="sidebar-resize-handle" id="sidebarResizeHandle"></div>
            <div class="search-box">
                <input type="text" placeholder="Search names and descriptions..." oninput="filterMethods(this.value)">
            </div>
            
            <div class="method-list">
"""

    # Group methods by category
    categories = {}
    for method_name, doc in method_docs.items():
        category = doc.get('category', 'Other')
        if category not in categories:
            categories[category] = []
        categories[category].append(method_name)

    # Display methods by category
    CATEGORY_ORDER = ['People', 'Extra Values', 'Involvements', 'Meetings & Attendance', 'Contributions',
                      'Tasks & Notes', 'Email & Texting', 'Searches & Tags', 'SQL Queries', 'Data & JSON',
                      'Dates', 'Text & Formatting', 'Content & Files', 'Page & Script', 'Web & Integrations', 'DocuSign']
    ordered = [c for c in CATEGORY_ORDER if c in categories] + sorted(c for c in categories if c not in CATEGORY_ORDER)
    for category in ordered:
        print '                <div class="category-section">'
        print '                    <div class="category-header" onclick="toggleCategory(this)">' + category + ' (' + str(len(categories[category])) + ')</div>'
        print '                    <div class="category-content">'
        for method in sorted(categories[category], key=lambda m: m.lower()):
            extra = ' deprecated' if 'deprecated' in method_docs[method].get('tags', []) else ''
            print '                        <div class="method-item documented' + extra + '" data-key="' + method + '" onclick="showMethod(\'' + method + '\')">' + method + '</div>'
        print '                    </div>'
        print '                </div>'

    # Add undocumented methods section
    # Count undocumented methods first
    undoc_model_count = 0
    undoc_q_count = 0

    # Members every .NET object has; not TouchPoint calls
    NET_BASICS = set(['Equals', 'GetHashCode', 'GetType', 'ToString', 'MemberwiseClone', 'ReferenceEquals', 'Finalize'])

    def _undocumented(obj, prefix):
        try:
            documented = set(k[len(prefix):] for k in method_docs.keys() if k.startswith(prefix) and '.' not in k[len(prefix):])
            return sorted(m for m in dir(obj) if not m.startswith('_') and m not in NET_BASICS and m not in documented)
        except Exception:
            return []

    undoc = ['model.' + m for m in _undocumented(model, 'model.')] + ['q.' + m for m in _undocumented(q, 'q.')]
    internal = [m for m in undoc if m in INTERNAL_NOTES]
    unknown = [m for m in undoc if m not in INTERNAL_NOTES]

    print '                <div class="category-section">'
    print '                    <div class="category-header" onclick="toggleCategory(this)">Internal, not for scripts (' + str(len(internal)) + ')</div>'
    print '                    <div class="category-content collapsed">'
    for method in internal:
        print '                        <div class="method-item undocumented" onclick="showMethod(\'' + method + '\')">' + method + '</div>'
    if not internal:
        print '                        <div style="padding: 10px; font-size: 12px; color: #6c757d; font-style: italic;">None found</div>'
    print '                    </div>'
    print '                </div>'

    if unknown:
        print '                <div class="category-section">'
        print '                    <div class="category-header" onclick="toggleCategory(this)">Not documented yet (' + str(len(unknown)) + ')</div>'
        print '                    <div class="category-content">'
        for method in unknown:
            print '                        <div class="method-item undocumented" onclick="showMethod(\'' + method + '\')">' + method + '</div>'
        print '                    </div>'
        print '                </div>'

    print """            </div>
            
            <div class="stats">
                Documented: """ + str(len(method_docs)) + """ methods<br>
                Categories: """ + str(len(categories)) + """
            </div>
        </div>
        
        <!-- Documentation Panel -->
        <div class="doc-panel" id="docPanel">
            <div class="doc-resize-handle" id="docResizeHandle"></div>
            <div class="doc-category" id="methodCategory">WELCOME</div>
            <div class="doc-title" id="methodName">TouchPoint API Explorer</div>
            <div class="apx-badges" id="apxBadges"></div>
            <div class="apx-sig" id="apxSig" style="display: none;"><code id="apxSigText"></code><a class="apx-link" onclick="apxCopySig()">copy</a></div>
            <div class="doc-desc" id="methodDesc">
                <strong>About This Tool:</strong><br>
                The TouchPoint API Explorer is an interactive development environment for testing and exploring TouchPoint's Python API. It provides real-time code execution with access to the complete TouchPoint model and query objects, along with comprehensive documentation for all available methods.<br><br>
                
                <div style="background: #fff3cd; border: 1px solid #ffc107; padding: 10px; border-radius: 4px; margin: 10px 0;">
                    <strong style="color: #856404;">⚠️ WARNING:</strong><br>
                    <span style="color: #856404;">This tool executes code directly against your TouchPoint database. Improper use can cause irreversible data changes, deletions, or system issues. Only use this tool if you understand the implications of your code. Always test queries and operations carefully before executing them in production.</span>
                </div>
                
                <strong>Getting Started:</strong><br>
                • Browse methods in the sidebar by category<br>
                • Click any method to view its documentation and example<br>
                • Write your code in the editor and click "Run Code" or press Ctrl/Cmd+Enter<br>
                • Use "Quick Test" to verify your environment setup<br>
                • Use "Discover" to explore available objects and methods
            </div>
            
            <details class="apx-details" id="apxDetails" style="display: none;"><summary>More detail</summary><div id="apxDetailsText"></div></details>

            <div id="apxFlowSection" class="apx-flow" style="display: none;">
                <h4>How it works</h4>
                <ol id="apxFlowList"></ol>
            </div>

            <div id="paramsSection" style="display: none;">
                <h4 style="font-size: 12px; margin: 10px 0;">Parameters</h4>
                <div id="paramsList"></div>
            </div>
            
            <div id="returnsSection" style="display: none;">
                <h4 style="font-size: 12px; margin: 10px 0;">Returns</h4>
                <div class="doc-desc" id="returnsInfo"></div>
            </div>
            
            <div id="apxWatchSection" class="apx-watch" style="display: none;">
                <h4>Watch out</h4>
                <ul id="apxWatchList"></ul>
            </div>

            <div id="exampleSection" style="display: none;">
                <h4 style="font-size: 12px; margin: 10px 0;">Example
                    <span class="apx-ex-actions">
                        <a class="apx-link" onclick="apxRunExample()">Run</a>
                        <a class="apx-link" onclick="useExample()">Open in new tab</a>
                        <a class="apx-link" onclick="apxCopyExample(this)">Copy</a>
                    </span>
                </h4>
                <div class="example-box" id="exampleCode"></div>
            </div>

            <div id="apxSeeAlsoSection" class="apx-seealso" style="display: none;">
                <h4>See also</h4>
                <div id="apxSeeAlso"></div>
            </div>
        </div>
        
        <!-- Editor Panel -->
        <div class="editor-panel">
            <div class="editor-toolbar">
                <button class="apx-btn apx-btn-primary" onclick="executeCode()">▶ Run Code</button>
                <button class="apx-btn apx-btn-secondary" onclick="clearEditor()">Clear</button>
                <button class="apx-btn apx-btn-secondary" onclick="useExample()">Use Example</button>
                <button class="apx-btn apx-btn-secondary" onclick="discover()">🔍 Discover</button>
                <button class="apx-btn apx-btn-secondary" onclick="quickTest()">🧪 Quick Test</button>
                <button class="apx-btn apx-btn-secondary" onclick="showShortcuts()">⌨ Shortcuts</button>
                <button class="apx-btn apx-btn-secondary" onclick="apxOpenLibrary()">History &amp; snippets</button>
                <button class="apx-btn apx-btn-secondary" onclick="apxSaveSnippet()">Save snippet</button>
                <span class="apx-spacer"></span>
                <span class="apx-out-meta" id="apxOutMeta"></span>
                <select class="apx-out-mode" id="apxOutMode" onchange="apxSetOutMode(this.value)" title="How to show the output">
                    <option value="text">Text output</option>
                    <option value="html">HTML output</option>
                </select>
                <button class="apx-btn apx-btn-secondary" onclick="apxCopyOutput()">Copy output</button>
            </div>
            <div class="apx-tabs" id="apxTabs"></div>
            
            <div class="editor-split">
                <div class="output-panel" id="output">Ready to execute code...</div>
                <div class="resize-handle" id="outputResizeHandle"></div>
                <div class="code-container">
                    <textarea id="codeEditor" style="display: none;"></textarea>
                </div>
            </div>
        </div>
    </div>
    
    <script src="https://cdnjs.cloudflare.com/ajax/libs/codemirror/5.65.2/codemirror.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/codemirror/5.65.2/mode/python/python.min.js"></script>
    
    <!-- Search functionality -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/codemirror/5.65.2/addon/search/searchcursor.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/codemirror/5.65.2/addon/search/search.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/codemirror/5.65.2/addon/search/replace.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/codemirror/5.65.2/addon/dialog/dialog.min.js"></script>
    
    <!-- Comment functionality -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/codemirror/5.65.2/addon/comment/comment.min.js"></script>
    
    <!-- Autocomplete functionality -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/codemirror/5.65.2/addon/hint/show-hint.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/codemirror/5.65.2/addon/hint/python-hint.min.js"></script>
    
    <!-- Edit functionality -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/codemirror/5.65.2/addon/edit/closebrackets.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/codemirror/5.65.2/addon/edit/matchbrackets.min.js"></script>
    
    <!-- Linting functionality -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/codemirror/5.65.2/addon/lint/lint.min.js"></script>
    
    <script>
        var methodDocs = """ + json.dumps(method_docs).replace('<', '\\u003c') + """;
        var APX_USER = '""" + _apx_user + """';
        var internalNotes = """ + json.dumps(INTERNAL_NOTES).replace('<', '\\u003c') + """;
        
        var currentMethod = null;
        var codeEditor = null;
        
        // Function to properly size CodeMirror
        function resizeCodeMirror() {
            if (codeEditor) {
                var container = document.querySelector('.code-container');
                if (container) {
                    var rect = container.getBoundingClientRect();
                    // Set explicit height to match container
                    codeEditor.setSize(null, rect.height);
                }
                codeEditor.refresh();
            }
        }
        
        // Helper function for calculating edit distance (for typo suggestions)
        function levenshteinDistance(a, b) {
            if (a.length === 0) return b.length;
            if (b.length === 0) return a.length;
            
            var matrix = [];
            for (var i = 0; i <= b.length; i++) {
                matrix[i] = [i];
            }
            for (var j = 0; j <= a.length; j++) {
                matrix[0][j] = j;
            }
            
            for (var i = 1; i <= b.length; i++) {
                for (var j = 1; j <= a.length; j++) {
                    if (b.charAt(i - 1) === a.charAt(j - 1)) {
                        matrix[i][j] = matrix[i - 1][j - 1];
                    } else {
                        matrix[i][j] = Math.min(
                            matrix[i - 1][j - 1] + 1,
                            matrix[i][j - 1] + 1,
                            matrix[i - 1][j] + 1
                        );
                    }
                }
            }
            
            return matrix[b.length][a.length];
        }
        
        // Custom TouchPoint Python Linter
        // ---- Linter: checks each call against the docs, plus Python 2.7 and TouchPoint traps ----
        var APX_NOTES = {
            'model.DebugPrint': ['info', 'model.DebugPrint shows nothing on a live TouchPoint site. Use print.'],
            'model.UpdateStatusFlag': ['warning', 'model.UpdateStatusFlag fails on current TouchPoint. Use model.UpdateStatusFlags() instead.'],
            'model.ExecuteSql': ['warning', 'model.ExecuteSql only works on a developer build. On a live site it raises an error.'],
            'model.DeletePeople': ['warning', 'model.DeletePeople permanently purges people. It runs even without the Developer role.'],
            'model.DeleteOrg': ['warning', 'model.DeleteOrg permanently purges involvements. It runs even without the Developer role.'],
            'model.UpdateAllSpouseId': ['warning', 'model.UpdateAllSpouseId writes to every person in the database.'],
            'model.DeleteAllExtraValueLike': ['warning', 'Deletes every matching extra value in the whole database, not just one person.'],
            'model.DeleteAllExtraValueLikeFamily': ['warning', 'Deletes every matching family extra value in the whole database.']
        };

        // Blank out comments and the inside of strings, keeping every position, so the checks
        // below only look at code.
        function apxMask(text) {
            var out = text.split(''), i = 0, n = text.length;
            while (i < n) {
                var c = text[i];
                if (c === '#') {
                    while (i < n && text[i] !== '\\n') { out[i] = ' '; i++; }
                    continue;
                }
                if (c === '"' || c === "'") {
                    var triple = text.substr(i, 3) === c + c + c;
                    var quote = triple ? c + c + c : c;
                    var j = i + quote.length, closed = false;
                    while (j < n) {
                        if (text[j] === '\\\\') { j += 2; continue; }
                        if (!triple && text[j] === '\\n') break;
                        if (text.substr(j, quote.length) === quote) { closed = true; break; }
                        j++;
                    }
                    for (var k = i + quote.length; k < Math.min(j, n); k++) if (text[k] !== '\\n') out[k] = ' ';
                    i = closed ? j + quote.length : j;
                    continue;
                }
                i++;
            }
            return out.join('');
        }

        function apxPos(text, idx) {
            var before = text.substring(0, idx);
            var line = before.split('\\n').length - 1;
            return CodeMirror.Pos(line, idx - before.lastIndexOf('\\n') - 1);
        }

        // Arguments in the call whose "(" is at open. null when it can't tell (*args, unclosed).
        function apxCountArgs(m, open) {
            var depth = 0, segs = [], cur = '';
            for (var i = open; i < m.length; i++) {
                var ch = m[i];
                if (ch === '(' || ch === '[' || ch === '{') { depth++; if (depth === 1) continue; }
                else if (ch === ')' || ch === ']' || ch === '}') {
                    depth--;
                    if (depth === 0) {
                        segs.push(cur);
                        segs = segs.map(function(x) { return x.trim(); }).filter(function(x) { return x; });
                        for (var k = 0; k < segs.length; k++) if (segs[k][0] === '*') return null;
                        return segs.length;
                    }
                }
                if (depth === 1 && ch === ',') { segs.push(cur); cur = ''; continue; }
                cur += ch;
            }
            return null;
        }

        function apxNearest(obj, name) {
            var out = [];
            Object.keys(methodDocs).forEach(function(k) {
                if (k.indexOf(obj + '.') !== 0) return;
                var rest = k.substring(obj.length + 1);
                if (rest.indexOf('.') >= 0) return;
                if (levenshteinDistance(rest.toLowerCase(), name.toLowerCase()) <= 2) out.push(k);
            });
            return out.slice(0, 3);
        }

        function apxHasMembers(path) {
            for (var k in methodDocs) if (k.indexOf(path + '.') === 0) return true;
            return false;
        }

        function touchpointPythonLinter(text, options) {
            var errors = [];
            var m = apxMask(text);
            var add = function(from, to, msg, sev) {
                errors.push({ from: apxPos(text, from), to: apxPos(text, to), message: msg, severity: sev });
            };
            var x;

            var re = /\\b(model|q)((?:\\.[A-Za-z_]\\w*)+)/g;
            while ((x = re.exec(m)) !== null) {
                if (x.index > 0 && m[x.index - 1] === '.') continue;
                var full = x[1] + x[2];
                var start = x.index, end = start + full.length;
                var parts = full.split('.');
                var key = null;
                for (var n = parts.length; n >= 2; n--) {
                    var k = parts.slice(0, n).join('.');
                    if (methodDocs[k] || internalNotes[k]) { key = k; break; }
                }
                var after = m.substring(end).match(/^\\s*\\(/);
                if (!key) {
                    var first = parts.slice(0, 2).join('.');
                    var near = apxNearest(parts[0], parts[1]);
                    add(start, start + first.length, first + ' does not exist' + (near.length ? '. Did you mean ' + near.join(', ') + '?' : '.'), 'warning');
                    continue;
                }
                if (internalNotes[key]) {
                    add(start, start + key.length, key + ' is internal: ' + internalNotes[key], 'warning');
                    continue;
                }
                if (key !== full) {
                    // e.g. model.docusign.Foo: the object has documented members, but not this one
                    var child = parts.slice(0, key.split('.').length + 1).join('.');
                    if (apxHasMembers(key) && !methodDocs[child]) {
                        add(start, start + child.length, child + ' does not exist.', 'warning');
                    }
                    continue;
                }
                var doc = methodDocs[key];
                if (doc.kind === 'property' && after) {
                    add(start, end + after[0].length, key + ' is a property, not a method. Remove the ().', 'error');
                    continue;
                }
                if (after && doc.arity) {
                    var got = apxCountArgs(m, end + after[0].length - 1);
                    var lo = doc.arity[0], hi = doc.arity[1];
                    if (got !== null && (got < lo || (hi !== null && got > hi))) {
                        var want = hi === null ? 'at least ' + lo : (lo === hi ? String(lo) : lo + ' to ' + hi);
                        add(start, end, key + ' takes ' + want + ' argument' + (want === '1' ? '' : 's') + ', not ' + got + '.' +
                            (doc.signatures ? ' Forms: ' + doc.signatures.join('   ') : ''), 'error');
                    }
                }
                if (doc.category === 'Deprecated') {
                    var use = /Use .*?\\.(?=\\s|$)/.exec((doc.watch_out || []).join(' ') + ' ' + (doc.details || '') + ' ' + (doc.summary || ''));
                    add(start, end, key + ' is deprecated.' + (use ? ' ' + use[0] : ''), 'warning');
                }
                if (APX_NOTES[key]) add(start, end, APX_NOTES[key][1], APX_NOTES[key][0]);
            }

            // hasattr on request data is always True in TouchPoint
            re = /\\bhasattr\\s*\\(\\s*(model\\.)?Data\\s*,/g;
            while ((x = re.exec(m)) !== null) {
                add(x.index, x.index + x[0].length, "model.Data returns '' for a field that was not sent, so hasattr is always True. Test the value instead: if model.Data.name:", 'warning');
            }

            // f-strings do not exist in Python 2.7
            re = /(^|[^\\w])([fF][rR]?|[rR][fF])(["'])/g;
            while ((x = re.exec(m)) !== null) {
                var at = x.index + x[1].length;
                add(at, at + x[2].length + 1, 'f-strings are not available in Python 2.7. Use "{0}".format(x) or % formatting.', 'error');
            }

            // print(a, b) prints a tuple in Python 2
            re = /\\bprint\\s*\\(([^()\\n]*,[^()\\n]*)\\)/g;
            while ((x = re.exec(m)) !== null) {
                add(x.index, x.index + 5, 'In Python 2.7, print(a, b) prints a tuple. Write print a, b instead.', 'info');
            }

            // Values glued into SQL text
            re = /\\bq\\.QuerySql\\w*\\s*\\(\\s*(["'])\\s*\\1?\\s*(\\+|%)/g;
            while ((x = re.exec(m)) !== null) {
                add(x.index, x.index + x[0].length, "Building SQL by joining in values lets a value change the query. Pass values as parameters instead: q.QuerySql(sql, {'Name': value}).", 'info');
            }
            return errors;
        }

        CodeMirror.registerHelper("lint", "python", touchpointPythonLinter);
        
        // Custom TouchPoint autocomplete hint provider
        // ---- Autocomplete: members of model, q and their sub objects, from the docs ----
        function apxObjectPath(line, ch) {
            var m = /([A-Za-z_][\\w.]*)\\.(\\w*)$/.exec(line.substring(0, ch));
            return m ? { path: m[1], word: m[2], start: ch - m[2].length } : null;
        }

        function apxRenderHint(el, self, data) {
            var doc = methodDocs[data.key] || {};
            el.innerHTML = '<div class="hint-method"><span class="hint-name">' + apxEsc(data.displayText) + '</span>' +
                '<span class="hint-category">' + apxEsc(doc.kind === 'property' ? 'property' : (doc.category || '')) + '</span></div>' +
                '<div class="hint-desc">' + apxEsc((doc.summary || doc.desc || '').substring(0, 110)) + '</div>';
        }

        function touchpointHint(editor, options) {
            var cur = editor.getCursor();
            var line = editor.getLine(cur.line);
            var ctx = apxObjectPath(line, cur.ch);
            var list = [];
            if (ctx && apxHasMembers(ctx.path)) {
                var prefix = ctx.path + '.';
                Object.keys(methodDocs).forEach(function(key) {
                    if (key.indexOf(prefix) !== 0) return;
                    var name = key.substring(prefix.length);
                    if (name.indexOf('.') >= 0) return;
                    if (ctx.word && name.toLowerCase().indexOf(ctx.word.toLowerCase()) !== 0) return;
                    var doc = methodDocs[key];
                    var isProp = doc.kind === 'property';
                    var shown = isProp ? name : name + '(' + (doc.params || []).map(function(p) { return p[0]; }).join(', ') + ')';
                    list.push({ text: isProp ? name : name + '(', displayText: shown, key: key, render: apxRenderHint });
                });
                list.sort(function(a, b) { return a.displayText.localeCompare(b.displayText); });
                return { list: list, from: CodeMirror.Pos(cur.line, ctx.start), to: CodeMirror.Pos(cur.line, cur.ch) };
            }
            var wm = /([A-Za-z_]\\w*)$/.exec(line.substring(0, cur.ch));
            var word = wm ? wm[1] : '';
            if (!word) return { list: [], from: cur, to: cur };
            var base = [
                ['model', 'TouchPoint API object'], ['q', 'Query object'],
                ['Data', 'Values sent to the page (same as model.Data)'], ['datetime', 'Python date and time module'],
                ['print ', 'Print output (Python 2.7 statement)'], ['import ', 'Import a module'],
                ['try:', 'Start error handling'], ['except Exception as e:', 'Catch an error'],
                ['for ', 'Loop'], ['if ', 'Condition'], ['def ', 'Define a function']
            ];
            base.forEach(function(b) {
                if (b[0].toLowerCase().indexOf(word.toLowerCase()) === 0) list.push({ text: b[0], displayText: b[0].trim() + '  ' + b[1] });
            });
            return { list: list, from: CodeMirror.Pos(cur.line, cur.ch - word.length), to: cur };
        }

        // Override Python hint with our custom TouchPoint hint
        CodeMirror.registerHelper('hint', 'python', touchpointHint);
        
        // Initialize CodeMirror
        document.addEventListener('DOMContentLoaded', function() {
            codeEditor = CodeMirror.fromTextArea(document.getElementById('codeEditor'), {
                mode: 'python',
                theme: 'monokai',
                lineNumbers: true,
                lineWrapping: false,  /* Disable line wrapping */
                indentUnit: 4,
                tabSize: 4,
                matchBrackets: true,
                autoCloseBrackets: true,
                gutters: ["CodeMirror-lint-markers"],
                lint: true,
                viewportMargin: 10,  /* Render 10 lines above and below viewport */
                scrollbarStyle: "native",  /* Use native scrollbars */
                lintOnChange: true,
                hintOptions: {
                    completeSingle: false
                },
                extraKeys: {
                    'Ctrl-Enter': function() { executeCode(); },
                    'Cmd-Enter': function() { executeCode(); },
                    'Shift-Enter': function(cm) {
                        var selected = cm.getSelection();
                        executeCode(selected && selected.trim() ? selected : null);
                    },
                    'Tab': function(cm) {
                        if (cm.somethingSelected()) {
                            cm.indentSelection("add");
                        } else {
                            cm.replaceSelection('    ', 'end');
                        }
                    },
                    'Shift-Tab': function(cm) {
                        cm.indentSelection("subtract");
                    },
                    'Ctrl-L': function() { clearEditor(); },
                    'Cmd-L': function() { clearEditor(); },
                    'F5': function() { executeCode(); },
                    'Ctrl-/': 'toggleComment',
                    'Cmd-/': 'toggleComment',
                    'Ctrl-Space': function(cm) { cm.showHint(); },
                    'Ctrl-Z': 'undo',
                    'Cmd-Z': 'undo',
                    'Ctrl-Y': 'redo',
                    'Cmd-Y': 'redo',
                    'Ctrl-Shift-Z': 'redo',
                    'Cmd-Shift-Z': 'redo'
                }
            });
            codeEditor.setSize('100%', '100%');
            apxInitTabs();
            codeEditor.on('change', apxSchedulePersist);
            
            // Auto-trigger autocomplete on dot after model or q
            codeEditor.on('inputRead', function(cm, change) {
                if (change.text && change.text.length === 1 && change.text[0] === '.') {
                    var cursor = cm.getCursor();
                    var token = cm.getTokenAt(CodeMirror.Pos(cursor.line, cursor.ch - 1));
                    var dotCtx = apxObjectPath(cm.getLine(cursor.line), cursor.ch);
                    if (dotCtx && apxHasMembers(dotCtx.path)) {
                        setTimeout(function() { cm.showHint(); }, 100);
                    }
                }
            });
            
            // Initialize resizable panels
            initResizablePanels();
            
            // Force refresh CodeMirror after a short delay to ensure proper sizing
            setTimeout(function() {
                resizeCodeMirror();
                // Additional delayed refresh to catch any layout shifts
                setTimeout(function() {
                    if (codeEditor) {
                        codeEditor.refresh();
                        // Force recalculation of scroll area
                        var totalLines = codeEditor.lineCount();
                        if (totalLines > 0) {
                            codeEditor.scrollTo(0, 0);
                            var lastLine = {line: totalLines - 1, ch: 0};
                            var coords = codeEditor.charCoords(lastLine, "local");
                            // Ensure the editor knows about all content
                            codeEditor.setSize(null, null);
                        }
                    }
                }, 300);
            }, 100);
            
            // Refresh CodeMirror on window resize
            window.addEventListener('resize', function() {
                resizeCodeMirror();
            });
            
            // Also refresh when output panel is resized
            var outputResizeObserver = new ResizeObserver(function() {
                resizeCodeMirror();
            });
            var outputPanel = document.getElementById('output');
            if (outputPanel) {
                outputResizeObserver.observe(outputPanel);
            }
        });
        
        // Helper function to update output
        function updateOutput(text) {
            var outputPanel = document.getElementById('output');
            outputPanel.textContent = text;
        }

        // ---- Tabs, output modes, history and snippets ----
        var apxTabs = [];
        var apxCurrentId = null;
        var apxCounter = 0;
        var apxOutMode = 'text';
        var apxPendingCode = null;
        var apxPersistTimer = null;
        var apxRestoring = false;

        function apxKey(name) { return 'apx_' + name + '_' + (typeof APX_USER !== 'undefined' ? APX_USER : '0'); }
        function apxRead(name, fallback) {
            try { var v = JSON.parse(localStorage.getItem(apxKey(name)) || 'null'); return v === null ? fallback : v; }
            catch (e) { return fallback; }
        }
        function apxWrite(name, value) {
            try { localStorage.setItem(apxKey(name), JSON.stringify(value)); } catch (e) { }
        }
        function apxEsc(t) {
            return String(t).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
        }
        function apxCurrentTab() {
            for (var i = 0; i < apxTabs.length; i++) if (apxTabs[i].id === apxCurrentId) return apxTabs[i];
            return null;
        }

        function apxNewTab(name, code, select) {
            apxCounter++;
            var tab = { id: apxCounter, name: name || ('Code ' + apxCounter), doc: CodeMirror.Doc(code || '', 'python'), output: null };
            apxTabs.push(tab);
            if (select !== false) apxSwitchTab(tab.id); else apxRenderTabs();
            apxSchedulePersist();
            return tab;
        }

        function apxSwitchTab(id) {
            var tab = null;
            for (var i = 0; i < apxTabs.length; i++) if (apxTabs[i].id === id) tab = apxTabs[i];
            if (!tab || !codeEditor) return;
            apxCurrentId = id;
            codeEditor.swapDoc(tab.doc);
            apxShowOutput(tab.output);
            apxRenderTabs();
            apxSchedulePersist();
            setTimeout(function() { codeEditor.refresh(); codeEditor.focus(); }, 1);
        }

        function apxCloseTab(id, event) {
            if (event) event.stopPropagation();
            if (apxTabs.length === 1) {
                apxTabs[0].doc.setValue('');
                apxTabs[0].output = null;
                apxShowOutput(null);
                apxSchedulePersist();
                return;
            }
            var idx = -1;
            for (var i = 0; i < apxTabs.length; i++) if (apxTabs[i].id === id) idx = i;
            if (idx < 0) return;
            apxTabs.splice(idx, 1);
            if (apxCurrentId === id) apxSwitchTab(apxTabs[Math.max(0, idx - 1)].id);
            else apxRenderTabs();
            apxSchedulePersist();
        }

        function apxRenameTab(id) {
            for (var i = 0; i < apxTabs.length; i++) {
                if (apxTabs[i].id === id) {
                    var name = prompt('Tab name:', apxTabs[i].name);
                    if (name && name.trim()) { apxTabs[i].name = name.trim(); apxRenderTabs(); apxSchedulePersist(); }
                }
            }
        }

        function apxRenderTabs() {
            var bar = document.getElementById('apxTabs');
            if (!bar) return;
            var h = '';
            apxTabs.forEach(function(t) {
                h += '<div class="apx-tab' + (t.id === apxCurrentId ? ' active' : '') + '" onclick="apxSwitchTab(' + t.id + ')" ondblclick="apxRenameTab(' + t.id + ')" title="Double-click to rename">' +
                    '<span class="apx-tab-name">' + apxEsc(t.name) + '</span>' +
                    '<span class="apx-tab-x" onclick="apxCloseTab(' + t.id + ', event)" title="Close tab">&times;</span></div>';
            });
            h += '<span class="apx-tab-new" onclick="apxNewTab()" title="New tab">+</span>';
            bar.innerHTML = h;
        }

        function apxSchedulePersist() {
            if (apxRestoring) return;
            clearTimeout(apxPersistTimer);
            apxPersistTimer = setTimeout(apxPersist, 400);
        }

        function apxPersist() {
            if (apxRestoring) return;
            var idx = 0;
            var list = apxTabs.map(function(t, i) {
                if (t.id === apxCurrentId) idx = i;
                return { name: t.name, code: t.doc.getValue() };
            });
            apxWrite('tabs', { v: 1, current: idx, tabs: list });
        }

        function apxInitTabs() {
            var saved = apxRead('tabs', null);
            apxRestoring = true;
            try {
                if (saved && saved.tabs && saved.tabs.length) {
                    saved.tabs.forEach(function(t) { apxNewTab(t.name, t.code, false); });
                    apxSwitchTab(apxTabs[Math.min(saved.current || 0, apxTabs.length - 1)].id);
                } else {
                    apxNewTab('Code 1', '# Enter Python code here\\n# Available: model, q\\n\\nprint "Hello TouchPoint!"', true);
                }
            } finally {
                apxRestoring = false;
            }
            apxOutMode = apxRead('outmode', 'text');
            var sel = document.getElementById('apxOutMode');
            if (sel) sel.value = apxOutMode;
            window.addEventListener('beforeunload', apxPersist);
        }

        function apxSetOutMode(mode) {
            apxOutMode = mode;
            apxWrite('outmode', mode);
            var tab = apxCurrentTab();
            apxShowOutput(tab ? tab.output : null);
        }

        function apxShowOutput(result) {
            var out = document.getElementById('output');
            var meta = document.getElementById('apxOutMeta');
            if (!result) {
                out.textContent = 'Ready to execute code...';
                if (meta) meta.textContent = '';
                return;
            }
            var text = result.text || '';
            var note = '';
            if (apxOutMode === 'html') {
                // Shown in a frame with scripts turned off, so printed HTML cannot touch this page
                out.innerHTML = '<iframe class="apx-html-frame" sandbox="" srcdoc="' + apxEsc(text) + '"></iframe>';
            } else {
                var trimmed = text.trim();
                if (trimmed && (trimmed[0] === '{' || trimmed[0] === '[')) {
                    try { text = JSON.stringify(JSON.parse(trimmed), null, 2); note = ' (JSON, formatted)'; } catch (e) { }
                }
                out.textContent = text || '(no output)';
            }
            if (meta) meta.textContent = (result.ms !== null && result.ms !== undefined ? result.ms + ' ms' : '') + note;
        }

        function apxCopyOutput() {
            var tab = apxCurrentTab();
            var text = tab && tab.output ? tab.output.text : '';
            if (!text) return;
            var done = function() {
                var meta = document.getElementById('apxOutMeta');
                if (meta) { var old = meta.textContent; meta.textContent = 'Copied'; setTimeout(function() { meta.textContent = old; }, 1200); }
            };
            if (navigator.clipboard && window.isSecureContext) {
                navigator.clipboard.writeText(text).then(done);
            } else {
                var ta = document.createElement('textarea');
                ta.value = text; ta.style.position = 'fixed'; ta.style.opacity = '0';
                document.body.appendChild(ta); ta.select();
                try { document.execCommand('copy'); done(); } catch (e) { }
                document.body.removeChild(ta);
            }
        }

        // Calls that change data or send email. Comments and strings are ignored.
        var APX_WRITE_RE = /^(Add|Delete|Remove|Update|Set|Join|Drop|Move|Edit|Adjust|Record|Enroll|Create|Close|Finish|Clear|Email|Send|Write|Execute|Merge|Change|Copy|Tag|Untag|Make|Save|Mark|Log|Insert|Upload|Post|Import|Process|Cancel|Schedule|Rest(Post|Put|Delete)|Pay|Refund|Void|Resend|Register|Unregister|Unsubscribe|Subscribe|Opt|Assign|Archive|Restore|Purge|Import|Batch)/;
        var APX_WRITE_ALSO = /(OrCreate|FindAdd|AddPerson|Email|Sms)/;
        function apxWriteCalls(code) {
            var stripped = code
                .replace(/(\x22\x22\x22|\x27\x27\x27)[\\s\\S]*?\\1/g, '""')
                .replace(/"(?:[^"\\\\\\n]|\\\\.)*"/g, '""')
                .replace(/\x27(?:[^\x27\\\\\\n]|\\\\.)*\x27/g, '""')
                .replace(/#[^\\n]*/g, '');
            var found = [], seen = {}, m;
            var re = /\\b(model|q)\\.([A-Za-z_]\\w*)\\s*\\(/g;
            while ((m = re.exec(stripped)) !== null) {
                var name = m[2];
                if ((APX_WRITE_RE.test(name) || APX_WRITE_ALSO.test(name)) && !/^(Get|Find(?!Add|OrCreate)|Query|Sql|Fmt)/.test(name)) {
                    var full = m[1] + '.' + name;
                    if (!seen[full]) { seen[full] = 1; found.push(full); }
                }
            }
            return found;
        }

        function apxRunPending() {
            var code = apxPendingCode;
            apxPendingCode = null;
            if (code) executeCode(code, true);
        }

        function apxCancelPending() {
            apxPendingCode = null;
            var tab = apxCurrentTab();
            apxShowOutput(tab ? tab.output : null);
        }

        function apxAddHistory(entry) {
            var h = apxRead('history', []).filter(function(x) { return x && x.code !== entry.code; });
            h.unshift(entry);
            apxWrite('history', h.slice(0, 50));
        }

        function apxSaveSnippet() {
            var tab = apxCurrentTab();
            if (!tab) return;
            var code = tab.doc.getValue();
            if (!code.trim()) return;
            var name = prompt('Snippet name:', tab.name);
            if (!name || !name.trim()) return;
            name = name.trim();
            var list = apxRead('snippets', []).filter(function(x) { return x.name !== name; });
            list.push({ name: name, code: code, saved: Date.now() });
            list.sort(function(a, b) { return a.name.toLowerCase() < b.name.toLowerCase() ? -1 : 1; });
            apxWrite('snippets', list);
            var meta = document.getElementById('apxOutMeta');
            if (meta) meta.textContent = 'Saved snippet "' + name + '"';
        }

        function apxOpenLibrary() {
            document.getElementById('apxLibrary').style.display = 'block';
            apxRenderLibrary();
            setTimeout(function() { document.getElementById('apxLibSearch').focus(); }, 1);
        }

        function apxCloseLibrary() {
            document.getElementById('apxLibrary').style.display = 'none';
        }

        function apxRenderLibrary() {
            var term = (document.getElementById('apxLibSearch').value || '').toLowerCase();
            var snippets = apxRead('snippets', []);
            var history = apxRead('history', []);
            var sh = '';
            snippets.forEach(function(x, i) {
                if (term && (x.name + ' ' + x.code).toLowerCase().indexOf(term) < 0) return;
                sh += '<div class="apx-lib-item" onclick="apxOpenSnippet(' + i + ')"><div><b>' + apxEsc(x.name) + '</b>' +
                    '<a class="apx-link" style="float:right" onclick="apxDeleteSnippet(' + i + ', this, event)">delete</a></div>' +
                    '<div class="apx-lib-code">' + apxEsc(x.code.replace(/\\s+/g, ' ').substring(0, 160)) + '</div></div>';
            });
            document.getElementById('apxSnippetList').innerHTML = sh || '<div class="apx-note">No snippets yet. Use Save snippet to keep the current tab.</div>';
            var hh = '';
            history.forEach(function(x, i) {
                if (!x || typeof x.code !== 'string') return;
                if (term && x.code.toLowerCase().indexOf(term) < 0) return;
                var when = new Date(x.at).toLocaleString(undefined, { month: 'short', day: 'numeric', hour: 'numeric', minute: '2-digit' });
                hh += '<div class="apx-lib-item" onclick="apxOpenHistory(' + i + ')">' +
                    '<div class="apx-lib-code">' + apxEsc(x.code.replace(/\\s+/g, ' ').substring(0, 160)) + '</div>' +
                    '<div class="apx-lib-meta">' + apxEsc(when) + (x.ms !== null && x.ms !== undefined ? ' &middot; ' + x.ms + ' ms' : '') +
                    (x.ok ? '' : ' &middot; <span class="apx-lib-err">error</span>') + '</div></div>';
            });
            document.getElementById('apxHistoryList').innerHTML = hh || '<div class="apx-note">Code you run will show here.</div>';
        }

        function apxOpenSnippet(i) {
            var x = apxRead('snippets', [])[i];
            if (!x) return;
            apxCloseLibrary();
            apxNewTab(x.name, x.code, true);
        }

        function apxDeleteSnippet(i, a, event) {
            event.stopPropagation();
            if (a.dataset.armed !== '1') {
                a.dataset.armed = '1';
                a.textContent = 'click again to delete';
                setTimeout(function() { a.dataset.armed = ''; a.textContent = 'delete'; }, 3000);
                return;
            }
            var list = apxRead('snippets', []);
            list.splice(i, 1);
            apxWrite('snippets', list);
            apxRenderLibrary();
        }

        function apxOpenHistory(i) {
            var x = apxRead('history', [])[i];
            if (!x) return;
            apxCloseLibrary();
            apxNewTab('From history', x.code, true);
        }

        function apxClearHistory(a) {
            if (a.dataset.armed !== '1') {
                a.dataset.armed = '1';
                a.textContent = 'click again to clear';
                setTimeout(function() { a.dataset.armed = ''; a.textContent = 'clear'; }, 3000);
                return;
            }
            apxWrite('history', []);
            a.dataset.armed = '';
            a.textContent = 'clear';
            apxRenderLibrary();
        }
        
        function toggleCategory(header) {
            var content = header.nextElementSibling;
            content.classList.toggle('collapsed');
        }
        
        var APX_TAG_LABELS = {
            read_only: 'Read only', writes_data: 'Writes data', sends_message: 'Sends email or text',
            slow: 'Slow', deprecated: 'Deprecated', known_issue: 'Known issue', outside_service: 'Outside service'
        };

        function apxSignature(name, doc) {
            if (doc.kind === 'property') return name;
            var parts = (doc.params || []).map(function(p) {
                var opt = p.length > 3 ? !p[3] : /optional/i.test(p[2] || '');
                var def = p.length > 4 && p[4] !== null && p[4] !== undefined ? '=' + p[4] : (opt ? '=...' : '');
                return p[0] + def;
            });
            return name + '(' + parts.join(', ') + ')' + (doc.returns ? '  ->  ' + doc.returns.split('.')[0] : '');
        }

        function apxCopySig() {
            var text = document.getElementById('apxSigText').textContent;
            if (navigator.clipboard && window.isSecureContext) navigator.clipboard.writeText(text);
        }

        function apxCopyExample(a) {
            var doc = methodDocs[currentMethod];
            if (!doc || !doc.example) return;
            if (navigator.clipboard && window.isSecureContext) navigator.clipboard.writeText(doc.example);
            var old = a.textContent; a.textContent = 'Copied'; setTimeout(function() { a.textContent = old; }, 1200);
        }

        function apxRunExample() {
            var doc = methodDocs[currentMethod];
            if (!doc || !doc.example) return;
            apxNewTab(currentMethod, doc.example, true);
            executeCode();
        }

        function apxShow(id, on) { document.getElementById(id).style.display = on ? 'block' : 'none'; }

        function showMethod(methodName) {
            document.querySelectorAll('.method-item').forEach(function(item) {
                item.classList.toggle('active', item.getAttribute('data-key') === methodName || item.textContent === methodName);
            });
            currentMethod = methodName;
            var doc = methodDocs[methodName];
            document.getElementById('methodName').textContent = methodName;
            var badges = document.getElementById('apxBadges');
            badges.innerHTML = '';
            apxShow('apxSig', false); apxShow('apxDetails', false); apxShow('apxWatchSection', false); apxShow('apxSeeAlsoSection', false); apxShow('apxFlowSection', false);

            if (!doc) {
                if (internalNotes[methodName]) {
                    document.getElementById('methodCategory').textContent = 'INTERNAL';
                    document.getElementById('methodDesc').textContent = internalNotes[methodName];
                } else {
                    document.getElementById('methodCategory').textContent = 'UNDOCUMENTED';
                    document.getElementById('methodDesc').textContent = 'No documentation yet. Discover shows its type and arguments without calling it.';
                }
                apxShow('paramsSection', false); apxShow('returnsSection', false); apxShow('exampleSection', false);
                return;
            }

            document.getElementById('methodCategory').textContent = doc.category || 'OTHER';
            var tags = [doc.kind === 'property' ? 'property' : 'method'].concat(doc.tags || []);
            badges.innerHTML = tags.map(function(t, i) {
                return '<span class="apx-badge ' + (i === 0 ? 'kind' : apxEsc(t)) + '">' + apxEsc(i === 0 ? t : (APX_TAG_LABELS[t] || t)) + '</span>';
            }).join('');

            document.getElementById('apxSigText').textContent = apxSignature(methodName, doc);
            apxShow('apxSig', true);

            document.getElementById('methodDesc').textContent = doc.summary || doc.desc || '';
            if (doc.summary && doc.details) {
                document.getElementById('apxDetailsText').textContent = doc.details;
                apxShow('apxDetails', true);
                document.getElementById('apxDetails').open = false;
            }

            if (doc.flow && doc.flow.length) {
                document.getElementById('apxFlowList').innerHTML = doc.flow.map(function(f) { return '<li>' + apxEsc(f) + '</li>'; }).join('');
                apxShow('apxFlowSection', true);
            }

            var params = doc.params || [];
            if (params.length || (doc.signatures && doc.signatures.length)) {
                var h = '';
                if (params.length) {
                    h += '<table class="apx-params"><tr><th>Name</th><th>Type</th><th>Required</th><th>Notes</th></tr>';
                    params.forEach(function(p) {
                        var required = p.length > 3 ? !!p[3] : !/optional/i.test(p[2] || '');
                        var def = p.length > 4 && p[4] !== null && p[4] !== undefined ? 'default ' + p[4] : '';
                        h += '<tr><td class="n">' + apxEsc(p[0]) + '</td><td class="t">' + apxEsc(p[1] || '') + '</td><td>' +
                            (required ? '<span class="apx-req">yes</span>' : '<span class="apx-opt">no' + (def ? ', ' + apxEsc(def) : '') + '</span>') +
                            '</td><td>' + apxEsc(p[2] || '') + '</td></tr>';
                    });
                    h += '</table>';
                }
                if (doc.signatures && doc.signatures.length > 1) {
                    h += '<div class="apx-sigs">' + apxEsc(doc.signatures.join('\\n')) + '</div>';
                }
                document.getElementById('paramsList').innerHTML = h;
                apxShow('paramsSection', true);
            } else {
                apxShow('paramsSection', false);
            }

            if (doc.returns) {
                document.getElementById('returnsInfo').textContent = doc.returns;
                apxShow('returnsSection', true);
            } else {
                apxShow('returnsSection', false);
            }

            var watch = doc.watch_out || [];
            if (watch.length) {
                document.getElementById('apxWatchList').innerHTML = watch.map(function(w) { return '<li>' + apxEsc(w) + '</li>'; }).join('');
                apxShow('apxWatchSection', true);
            }

            if (doc.example) {
                document.getElementById('exampleCode').textContent = doc.example;
                apxShow('exampleSection', true);
            } else {
                apxShow('exampleSection', false);
            }

            var see = (doc.see_also || []).filter(function(k) { return methodDocs[k]; });
            if (see.length) {
                document.getElementById('apxSeeAlso').innerHTML = see.map(function(k) {
                    return '<a data-key="' + apxEsc(k) + '" onclick="showMethod(this.dataset.key)">' + apxEsc(k) + '</a>';
                }).join('');
                apxShow('apxSeeAlsoSection', true);
            }
            var panel = document.getElementById('docPanel');
            if (panel) panel.scrollTop = 0;
        }
        
        function discover() {
            if (!currentMethod) {
                alert('Please select a method first');
                return;
            }
            
            var parts = currentMethod.split('.');
            var objName = parts[0];
            var methodName = parts[1];
            
            var code = '# Discovering ' + currentMethod + '\\n' +
'import inspect\\n' +
'\\n' +
'try:\\n' +
'    method = ' + objName + '.' + methodName + '\\n' +
'    print "=" * 60\\n' +
'    print "Method: ' + currentMethod + '"\\n' +
'    print "=" * 60\\n' +
'    \\n' +
'    # Get type\\n' +
'    print "\\\\nType:", type(method)\\n' +
'    \\n' +
'    # Try to get docstring\\n' +
'    try:\\n' +
'        doc = method.__doc__\\n' +
'        if doc:\\n' +
'            print "\\\\nDocumentation:"\\n' +
'            print doc[:500]\\n' +
'    except:\\n' +
'        pass\\n' +
'    \\n' +
'    # Try to get signature\\n' +
'    try:\\n' +
'        if hasattr(inspect, \\'getargspec\\'):\\n' +
'            spec = inspect.getargspec(method)\\n' +
'            print "\\\\nArguments:", spec.args\\n' +
'            if spec.defaults:\\n' +
'                print "Defaults:", spec.defaults\\n' +
'    except:\\n' +
'        pass\\n' +
'    \\n' +
'    # Discover only looks. It never calls the method, because calling\\n' +
'    # can change data (for example model.UpdateStatusFlags).\\n' +
'    \\n' +
'except Exception as e:\\n' +
'    print "Error:", str(e)';
            
            apxNewTab('Discover ' + methodName, code, true);
            executeCode();
        }
        
        function quickTest() {
            var code = '# Quick System Test\\n' +
'print "=" * 50\\n' +
'print "TouchPoint API Explorer - System Test"\\n' +
'print "=" * 50\\n' +
'\\n' +
'# Check user\\n' +
'if hasattr(model, "PeopleId"):\\n' +
'    print "\\\\nUser ID:", model.PeopleId\\n' +
'if hasattr(model, "UserName"):\\n' +
'    print "Username:", model.UserName\\n' +
'\\n' +
'# Count available methods\\n' +
'if "model" in globals():\\n' +
'    model_methods = [m for m in dir(model) if not m.startswith("_")]\\n' +
'    print "\\\\nModel methods available:", len(model_methods)\\n' +
'if "q" in globals():\\n' +
'    q_methods = [m for m in dir(q) if not m.startswith("_")]\\n' +
'    print "Query methods available:", len(q_methods)';
    
            apxNewTab('Quick test', code, true);
            executeCode();
        }
        
        function executeCode(codeOverride, force) {
            var tab = apxCurrentTab();
            var code = codeOverride || (codeEditor ? codeEditor.getValue() : document.getElementById('codeEditor').value);
            if (!code.trim()) {
                updateOutput('Please enter code to execute');
                return;
            }
            if (!force) {
                var writes = apxWriteCalls(code);
                if (writes.length) {
                    apxPendingCode = code;
                    var out = document.getElementById('output');
                    out.innerHTML = '<div class="apx-warn"><b>This code can change data or send email.</b> It calls ' +
                        writes.map(function(w) { return '<code>' + apxEsc(w) + '</code>'; }).join(', ') +
                        '. Check it before running.<br>' +
                        '<button class="apx-btn apx-btn-primary" onclick="apxRunPending()">Run anyway</button>' +
                        '<button class="apx-btn apx-btn-secondary" onclick="apxCancelPending()">Cancel</button></div>';
                    document.getElementById('apxOutMeta').textContent = '';
                    return;
                }
            }

            updateOutput('Executing...');
            document.getElementById('apxOutMeta').textContent = '';
            var formPath = window.location.pathname.replace('/PyScript/', '/PyScriptForm/');
            var xhr = new XMLHttpRequest();
            xhr.open('POST', formPath, true);
            xhr.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
            xhr.onreadystatechange = function() {
                if (xhr.readyState !== 4) return;
                var result = { text: '', ms: null, ok: false };
                if (xhr.status === 200) {
                    var response = xhr.responseText;
                    var msMatch = /AJAX_MS:(\\d+)/.exec(response);
                    var marker = '<!-- AJAX_COMPLETE -->';
                    if (response.indexOf(marker) > -1) response = response.substring(0, response.indexOf(marker));
                    result.text = response.replace(/\\s+$/, '');
                    result.ms = msMatch ? parseInt(msMatch[1], 10) : null;
                    result.ok = !/(^|\\n)Error: /.test(result.text) && result.text.indexOf('Not allowed:') !== 0;
                } else {
                    result.text = 'Error: Request failed (status ' + xhr.status + ')';
                }
                if (tab) tab.output = result;
                apxAddHistory({ code: code, at: Date.now(), ms: result.ms, ok: result.ok });
                if (!tab || tab === apxCurrentTab()) apxShowOutput(result);
            };
            xhr.send('code_enc=' + encodeURIComponent(encodeURIComponent(code)));
        }

        function clearEditor() {
            if (codeEditor) {
                codeEditor.setValue('');
            }
        }
        
        function useExample() {
            if (currentMethod && methodDocs[currentMethod] && methodDocs[currentMethod].example) {
                apxNewTab(currentMethod, methodDocs[currentMethod].example, true);
            }
        }
        
        function filterMethods(term) {
            var words = term.toLowerCase().split(/\\s+/).filter(function(w) { return w; });
            document.querySelectorAll('.method-item').forEach(function(item) {
                var key = item.getAttribute('data-key') || item.textContent;
                var doc = methodDocs[key] || {};
                var hay = (key + ' ' + (doc.summary || '') + ' ' + (doc.details || '') + ' ' + (doc.watch_out || []).join(' ') + ' ' + (doc.desc || '') + ' ' + (doc.category || '') + ' ' + (internalNotes[key] || '')).toLowerCase();
                var hit = words.every(function(w) { return hay.indexOf(w) >= 0; });
                item.style.display = hit ? 'block' : 'none';
            });
            document.querySelectorAll('.category-section').forEach(function(section) {
                var hasVisible = false;
                section.querySelectorAll('.method-item').forEach(function(item) {
                    if (item.style.display !== 'none') hasVisible = true;
                });
                section.style.display = hasVisible ? 'block' : 'none';
                if (words.length && hasVisible) {
                    var content = section.querySelector('.category-content');
                    if (content) content.classList.remove('collapsed');
                    var header = section.querySelector('.category-header');
                    if (header) header.classList.remove('collapsed');
                }
            });
            if (!words.length) apxSetAllCategories(true);
        }
        
        function showShortcuts() {
            document.getElementById('shortcutsModal').style.display = 'block';
        }
        
        function closeShortcutsModal() {
            document.getElementById('shortcutsModal').style.display = 'none';
        }
        
        // Close modal when clicking outside
        window.onclick = function(event) {
            var modal = document.getElementById('shortcutsModal');
            if (event.target == modal) {
                modal.style.display = 'none';
            }
        }
        
        // Add collapsible functionality to category headers
        function toggleCategory(header) {
            header.classList.toggle('collapsed');
            var content = header.nextElementSibling;
            if (content && content.classList.contains('category-content')) {
                content.classList.toggle('collapsed');
            }
        }
        
        // Initialize resizable panels
        function initResizablePanels() {
            // === OUTPUT PANEL RESIZE (vertical) ===
            var outputPanel = document.getElementById('output');
            var outputResizeHandle = document.getElementById('outputResizeHandle');

            // Prevent multiple event listeners
            if (outputResizeHandle && !outputResizeHandle.hasAttribute('data-resize-initialized')) {
                outputResizeHandle.setAttribute('data-resize-initialized', 'true');

                var isResizingOutput = false;
                var startY = 0;
                var startHeight = 0;

                var handleOutputMouseDown = function(e) {
                    isResizingOutput = true;
                    startY = e.pageY;
                    startHeight = outputPanel.offsetHeight;
                    document.body.style.cursor = 'ns-resize';
                    document.body.style.userSelect = 'none';
                    e.preventDefault();
                    e.stopPropagation();
                };

                var handleOutputMouseMove = function(e) {
                    if (!isResizingOutput) return;

                    var deltaY = e.pageY - startY;
                    var newHeight = startHeight + deltaY;

                    // Limit the height
                    newHeight = Math.max(100, Math.min(600, newHeight));

                    // Use height instead of flex for smoother resizing
                    outputPanel.style.height = newHeight + 'px';
                    outputPanel.style.flex = 'none';

                    e.preventDefault();
                };

                var handleOutputMouseUp = function() {
                    if (isResizingOutput) {
                        isResizingOutput = false;
                        document.body.style.cursor = '';
                        document.body.style.userSelect = '';

                        // Refresh CodeMirror after resize
                        if (codeEditor) {
                            setTimeout(function() {
                                codeEditor.refresh();
                            }, 0);
                        }
                    }
                };

                outputResizeHandle.addEventListener('mousedown', handleOutputMouseDown);
                document.addEventListener('mousemove', handleOutputMouseMove);
                document.addEventListener('mouseup', handleOutputMouseUp);
            }

            // === SIDEBAR RESIZE (horizontal) ===
            var sidebarHandle = document.getElementById('sidebarResizeHandle');
            var sidebar = document.getElementById('sidebar');
            var isResizingSidebar = false;
            var sideStartX = 0;
            var sideStartWidth = 0;

            if (sidebarHandle && sidebar) {
                sidebarHandle.addEventListener('mousedown', function(e) {
                    isResizingSidebar = true;
                    sideStartX = e.clientX;
                    sideStartWidth = sidebar.getBoundingClientRect().width;
                    document.body.style.cursor = 'col-resize';
                    e.preventDefault();
                });

                document.addEventListener('mousemove', function(e) {
                    if (!isResizingSidebar) return;
                    var newWidth = sideStartWidth + (e.clientX - sideStartX);
                    if (newWidth > 150 && newWidth < 500) {
                        sidebar.style.width = newWidth + 'px';
                        sidebar.style.flexShrink = '0';
                    }
                });

                document.addEventListener('mouseup', function() {
                    if (isResizingSidebar) {
                        isResizingSidebar = false;
                        document.body.style.cursor = 'default';
                        apxSaveWidths();
                        if (typeof resizeCodeMirror === 'function') {
                            resizeCodeMirror();
                        }
                    }
                });
            }

            // === DOC PANEL RESIZE (horizontal) ===
            var docHandle = document.getElementById('docResizeHandle');
            var docPanel = document.getElementById('docPanel');
            var isResizingDoc = false;
            var docStartX = 0;
            var docStartWidth = 0;

            if (docHandle && docPanel) {
                docHandle.addEventListener('mousedown', function(e) {
                    isResizingDoc = true;
                    docStartX = e.clientX;
                    docStartWidth = parseInt(window.getComputedStyle(docPanel).width, 10);
                    document.body.style.cursor = 'col-resize';
                    e.preventDefault();
                });

                document.addEventListener('mousemove', function(e) {
                    if (!isResizingDoc) return;
                    var newWidth = docStartWidth + (e.clientX - docStartX);
                    if (newWidth > 250 && newWidth < 1100) {
                        docPanel.style.width = newWidth + 'px';
                        docPanel.style.flexShrink = '0';
                    }
                });

                document.addEventListener('mouseup', function() {
                    if (isResizingDoc) {
                        isResizingDoc = false;
                        document.body.style.cursor = 'default';
                        apxSaveWidths();
                        if (typeof resizeCodeMirror === 'function') {
                            resizeCodeMirror();
                        }
                    }
                });
            }
        }
        
        // Panel widths you drag to are remembered in this browser
        function apxSaveWidths() {
            var sb = document.getElementById('sidebar'), dp = document.getElementById('docPanel');
            apxWrite('widths', { sidebar: sb ? Math.round(sb.getBoundingClientRect().width) : null,
                                 doc: dp ? Math.round(dp.getBoundingClientRect().width) : null });
        }

        function apxRestoreWidths() {
            var w = apxRead('widths', null);
            if (!w) return;
            var sb = document.getElementById('sidebar'), dp = document.getElementById('docPanel');
            if (sb && w.sidebar > 150 && w.sidebar < 500) { sb.style.width = w.sidebar + 'px'; sb.style.flexShrink = '0'; }
            if (dp && w.doc > 250 && w.doc < 1100) { dp.style.width = w.doc + 'px'; dp.style.flexShrink = '0'; }
            if (typeof resizeCodeMirror === 'function') resizeCodeMirror();
        }

        function apxSetAllCategories(collapsed) {
            document.querySelectorAll('.category-header').forEach(function(header) {
                header.classList.toggle('collapsed', collapsed);
                var content = header.nextElementSibling;
                if (content) content.classList.toggle('collapsed', collapsed);
            });
            var btn = document.getElementById('toggleAllBtn');
            if (btn) btn.textContent = collapsed ? 'Expand All' : 'Collapse All';
        }

        // Initialize collapsible categories on page load
        window.addEventListener('load', function() {
            // Initialize resizable panels
            initResizablePanels();
            apxRestoreWidths();
            // Make category headers collapsible
            document.querySelectorAll('.category-header').forEach(function(header) {
                header.onclick = function() {
                    toggleCategory(this);
                };
            });
            
            // Add a "Collapse All" / "Expand All" button to the sidebar
            var searchBox = document.querySelector('.search-box');
            if (searchBox && !document.getElementById('toggleAllBtn')) {
                var toggleBtn = document.createElement('button');
                toggleBtn.id = 'toggleAllBtn';
                toggleBtn.textContent = 'Collapse All';
                toggleBtn.style.cssText = 'width: 100%; margin-top: 5px; padding: 5px; font-size: 11px; background: #007bff; color: white; border: none; border-radius: 3px; cursor: pointer;';
                toggleBtn.onclick = function() {
                    var headers = document.querySelectorAll('.category-header');
                    var allCollapsed = Array.from(headers).every(h => h.classList.contains('collapsed'));
                    
                    headers.forEach(function(header) {
                        var content = header.nextElementSibling;
                        if (allCollapsed) {
                            header.classList.remove('collapsed');
                            if (content) content.classList.remove('collapsed');
                        } else {
                            header.classList.add('collapsed');
                            if (content) content.classList.add('collapsed');
                        }
                    });
                    
                    toggleBtn.textContent = allCollapsed ? 'Collapse All' : 'Expand All';
                };
                searchBox.appendChild(toggleBtn);
            }
            // Start with every category closed
            apxSetAllCategories(true);
        });
    </script>
</body>
</html>"""

if _allowed and not _is_run:
    sys.stdout = _page_stdout
    model.Form = _page_buf.getvalue()
